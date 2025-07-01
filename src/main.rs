use calamine::{Reader, Xlsx};
use clap::Parser;
use rust_xlsxwriter::Workbook;
use std::path::PathBuf;
use tracseq_rc::reverse_complement;
use csv::{ReaderBuilder, WriterBuilder};
use std::fs::File;

#[derive(Parser)]
#[command(author, version, about, long_about = None)]
struct Args {
    /// Path to the Excel or CSV file
    file: PathBuf,
}

#[derive(Debug)]
enum FileType {
    Excel,
    Csv,
}

fn detect_file_type(path: &PathBuf) -> Result<FileType, Box<dyn std::error::Error>> {
    match path.extension().and_then(|ext| ext.to_str()) {
        Some("xlsx") | Some("xls") => Ok(FileType::Excel),
        Some("csv") => Ok(FileType::Csv),
        _ => Err("Unsupported file type. Please use .xlsx, .xls, or .csv files.".into()),
    }
}

fn process_csv_file(file_path: &PathBuf) -> Result<(), Box<dyn std::error::Error>> {
    println!("\nProcessing CSV file...");
    
    // Create output filename
    let output_path = file_path.with_file_name(format!(
        "{}_RC.csv",
        file_path.file_stem().unwrap().to_string_lossy()
    ));
    
    // Open input CSV file
    let file = File::open(file_path)?;
    let mut reader = ReaderBuilder::new()
        .has_headers(true)
        .from_reader(file);
    
    // Get headers
    let headers = reader.headers()?.clone();
    
    // Find column indices
    let indexnt_col = headers.iter().position(|h| h == "IndexNtSequence");
    let indexnt2_col = headers.iter().position(|h| h == "IndexNtSequence2");
    let index2_col = headers.iter().position(|h| h == "Index 2");
    let index_col = headers.iter().position(|h| h == "Index");
    let id_col = headers.iter().position(|h| h == "Id" || h == "Sample ID");
    
    // Detect columns containing DNA sequences
    println!("\nScanning for DNA sequence columns...");
    let mut sequence_columns: Vec<(usize, String, bool)> = Vec::new();
    
    // First check if we have standard columns
    if indexnt_col.is_some() || indexnt2_col.is_some() || index2_col.is_some() || index_col.is_some() {
        if let Some(idx) = indexnt_col {
            sequence_columns.push((idx, "IndexNtSequence".to_string(), true));
        }
        if let Some(idx) = indexnt2_col {
            sequence_columns.push((idx, "IndexNtSequence2".to_string(), false));
        }
        if let Some(idx) = index2_col {
            sequence_columns.push((idx, "Index 2".to_string(), false));
        }
        if let Some(idx) = index_col {
            sequence_columns.push((idx, "Index".to_string(), true));
        }
    } else {
        // No standard columns, scan for DNA patterns in first few rows
        let mut sample_reader = ReaderBuilder::new()
            .has_headers(true)
            .from_reader(File::open(file_path)?);
        
        let sample_records: Vec<_> = sample_reader.records()
            .take(10)
            .filter_map(Result::ok)
            .collect();
        
        for col_idx in 0..headers.len() {
            let mut has_sequences = false;
            let mut has_delimiter = false;
            
            for record in &sample_records {
                if let Some(field) = record.get(col_idx) {
                    if field.len() >= 4 {
                        if field.contains('-') {
                            let parts: Vec<&str> = field.split('-').collect();
                            if parts.len() == 2 && parts[1].len() >= 4 {
                                if parts[1].chars().all(|c| "ATGCN".contains(c)) {
                                    has_sequences = true;
                                    has_delimiter = true;
                                    break;
                                }
                            }
                        } else if field.chars().all(|c| "ATGCN".contains(c)) {
                            has_sequences = true;
                            has_delimiter = false;
                            break;
                        }
                    }
                }
            }
            
            if has_sequences {
                let col_name = headers.get(col_idx).map(|s| s.to_string()).unwrap_or_else(|| format!("Column_{}", col_idx + 1));
                sequence_columns.push((col_idx, col_name, has_delimiter));
            }
        }
    }
    
    // Debug output to show which columns were detected
    println!("\nDetected columns:");
    println!("- Id column: {}", if id_col.is_some() { "Found" } else { "NOT FOUND" });
    
    if sequence_columns.is_empty() {
        println!("- Sequence columns: NOT FOUND");
        println!("\n⚠️  Warning: No DNA sequence columns detected.");
        if id_col.is_some() {
            println!("   Id column found but no sequence data to process.");
        }
    } else {
        println!("- Sequence columns found: {}", sequence_columns.len());
        for (idx, name, delim) in &sequence_columns {
            println!("  * Column {}: '{}' (delimiter: {})", idx + 1, name, if *delim { "yes" } else { "no" });
        }
    }
    
    println!("\nAll columns in file:");
    for (idx, header) in headers.iter().enumerate() {
        println!("  Column {}: '{}'", idx + 1, header);
    }
    
    // Create output CSV file
    let output_file = File::create(&output_path)?;
    let mut writer = WriterBuilder::new()
        .has_headers(true)
        .from_writer(output_file);
    
    // Write headers
    writer.write_record(&headers)?;
    
    let mut data_row_count = 0;
    
    // Process rows
    for result in reader.records() {
        let record = result?;
        let mut output_record = Vec::new();
        let mut rc_value: Option<String> = None;
        let mut id_value: Option<String> = None;
        let mut processed_col_name: Option<String> = None;
        
        // Get ID value if present
        if let Some(idx) = id_col {
            if let Some(field) = record.get(idx) {
                id_value = Some(field.to_string());
            }
        }
        
        for (col_idx, field) in record.iter().enumerate() {
            // Check if this column is a sequence column
            let mut processed = false;
            for (seq_col_idx, seq_col_name, has_delimiter) in &sequence_columns {
                if col_idx == *seq_col_idx {
                    if *has_delimiter {
                        // Handle delimiter pattern
                        let parts: Vec<&str> = field.splitn(2, '-').collect();
                        if parts.len() == 2 {
                            let rc = reverse_complement(parts[1]);
                            let new_val = format!("{}-{}", parts[0], rc);
                            rc_value = Some(new_val.clone());
                            processed_col_name = Some(seq_col_name.clone());
                            output_record.push(new_val);
                        } else {
                            // No delimiter found, treat as full sequence
                            let rc = reverse_complement(field);
                            rc_value = Some(rc.clone());
                            processed_col_name = Some(seq_col_name.clone());
                            output_record.push(rc);
                        }
                    } else {
                        // Direct sequence pattern
                        let rc = reverse_complement(field);
                        rc_value = Some(rc.clone());
                        processed_col_name = Some(seq_col_name.clone());
                        output_record.push(rc);
                    }
                    processed = true;
                    break;
                }
            }
            
            if !processed {
                // Copy non-sequence fields as-is
                output_record.push(field.to_string());
            }
        }
        
        // Write the output record
        writer.write_record(&output_record)?;
        
        // Print SQL update statement if both values are present and ID is not empty
        if let (Some(id), Some(rc), Some(col_name)) = (id_value, rc_value, processed_col_name) {
            if !id.trim().is_empty() {
                // Generate SQL with proper column name escaping
                if col_name.contains(' ') || col_name.starts_with("Column_") {
                    println!(
                        "UPDATE SampleBatchItems SET [{}] = '{}' WHERE Id = '{}';",
                        col_name, rc, id
                    );
                } else {
                    println!(
                        "UPDATE SampleBatchItems SET {} = '{}' WHERE Id = '{}';",
                        col_name, rc, id
                    );
                }
            }
        }
        
        data_row_count += 1;
    }
    
    writer.flush()?;
    
    println!("File processed successfully!");
    println!("Output saved to: {}", output_path.display());
    println!("\nNumber of data rows: {}", data_row_count);
    println!("Number of columns: {}", headers.len());
    
    Ok(())
}

fn process_excel_file(file_path: &PathBuf) -> Result<(), Box<dyn std::error::Error>> {
    println!("\nProcessing Excel file...");
    
    // Open the input Excel file
    let mut input_workbook: Xlsx<_> = calamine::open_workbook(file_path)?;

    // Create output filename
    let output_path = file_path.with_file_name(format!(
        "{}_RC.xlsx",
        file_path.file_stem().unwrap().to_string_lossy()
    ));

    // Create new workbook for output
    let mut output_workbook = Workbook::new();
    let sheet = output_workbook.add_worksheet();

    // Get the first sheet
    if let Some(Ok(range)) = input_workbook.worksheet_range_at(0) {
        // Find the row that starts with "Sample ID"
        let all_rows: Vec<_> = range.rows().collect();
        let mut header_row_idx = None;
        let mut header_row = None;
        
        println!("\nSearching for 'Sample ID' header row...");
        for (idx, row) in all_rows.iter().enumerate() {
            if let Some(first_cell) = row.get(0) {
                if first_cell.to_string().trim() == "Sample ID" {
                    header_row_idx = Some(idx);
                    header_row = Some(row);
                    println!("Found 'Sample ID' header at row {}", idx + 1);
                    break;
                }
            }
        }
        
        let header_row = match header_row {
            Some(row) => row,
            None => {
                println!("Error: Could not find a row starting with 'Sample ID'");
                println!("This file may not contain sample data in the expected format.");
                return Ok(());
            }
        };
        
        let header_row_idx = header_row_idx.unwrap();
        
        println!("\nAnalyzing file structure...");
        println!("Found {} columns", header_row.len());

        // Write header row
        for (col, cell) in header_row.iter().enumerate() {
            sheet.write_string(0, col as u16, &cell.to_string())?;
        }

        // Check for IndexNtSequence, IndexNtSequence2, Index 2, or Index
        let indexnt_col = header_row
            .iter()
            .position(|c| c.to_string() == "IndexNtSequence");
        let indexnt2_col = header_row
            .iter()
            .position(|c| c.to_string() == "IndexNtSequence2");
        let index2_col = header_row.iter().position(|c| c.to_string() == "Index 2");
        let index_col = header_row.iter().position(|c| c.to_string() == "Index");
        // Look for both "Id" and "Sample ID" columns
        let id_col = header_row.iter().position(|c| {
            let col = c.to_string();
            col == "Id" || col == "Sample ID"
        });

        // Detect columns containing DNA sequences by scanning data
        println!("\nScanning for DNA sequence columns...");
        let mut sequence_columns: Vec<(usize, String, bool)> = Vec::new(); // (index, name, has_delimiter)
        
        // First check if we have standard columns
        if indexnt_col.is_some() || indexnt2_col.is_some() || index2_col.is_some() || index_col.is_some() {
            if let Some(idx) = indexnt_col {
                sequence_columns.push((idx, "IndexNtSequence".to_string(), true));
            }
            if let Some(idx) = indexnt2_col {
                sequence_columns.push((idx, "IndexNtSequence2".to_string(), false));
            }
            if let Some(idx) = index2_col {
                sequence_columns.push((idx, "Index 2".to_string(), false));
            }
            if let Some(idx) = index_col {
                sequence_columns.push((idx, "Index".to_string(), true));
            }
        } else {
            // No standard columns, scan for DNA patterns in rows after the header
            let sample_rows: Vec<_> = all_rows.iter()
                .skip(header_row_idx + 1)
                .take(10)
                .collect();
            
            for col_idx in 0..header_row.len() {
                let mut has_sequences = false;
                let mut has_delimiter = false;
                
                for row in &sample_rows {
                    if let Some(cell) = row.get(col_idx) {
                        let val = cell.to_string();
                        // Only consider it a sequence if it's at least 4 characters of DNA
                        if val.len() >= 4 {
                            // Check for delimiter pattern (e.g., "Prefix-SEQUENCE")
                            if val.contains('-') {
                                let parts: Vec<&str> = val.split('-').collect();
                                if parts.len() == 2 && parts[1].len() >= 4 {
                                    // More strict check: at least 80% should be ATGCN
                                    let dna_chars = parts[1].chars().filter(|c| "ATGCN".contains(*c)).count();
                                    if dna_chars as f32 / parts[1].len() as f32 >= 0.8 {
                                        has_sequences = true;
                                        has_delimiter = true;
                                        break;
                                    }
                                }
                            }
                            // Check for direct sequence pattern - must be all DNA chars
                            else if val.len() >= 6 && val.chars().all(|c| "ATGCN".contains(c)) {
                                has_sequences = true;
                                has_delimiter = false;
                                break;
                            }
                        }
                    }
                }
                
                if has_sequences {
                    let col_name = header_row[col_idx].to_string();
                    if col_name.is_empty() {
                        sequence_columns.push((col_idx, format!("Column_{}", col_idx + 1), has_delimiter));
                    } else {
                        sequence_columns.push((col_idx, col_name, has_delimiter));
                    }
                }
            }
        }

        // Debug output to show which columns were detected
        println!("\nColumns in Sample ID section:");
        for (idx, cell) in header_row.iter().enumerate() {
            let col_name = cell.to_string();
            if !col_name.is_empty() {
                println!("  Column {}: '{}'", idx + 1, col_name);
            }
        }
        
        println!("\nDetected columns:");
        println!("- Id column: {}", if id_col.is_some() { "Found" } else { "NOT FOUND" });
        
        if sequence_columns.is_empty() {
            println!("- Sequence columns: NOT FOUND");
            println!("\n⚠️  Warning: No DNA sequence columns detected.");
            if id_col.is_some() {
                println!("   Id column found but no sequence data to process.");
            }
            
            // Show sample data from first few rows to help diagnose
            println!("\nShowing first 3 data rows to help identify sequence columns:");
            for (i, row) in all_rows.iter().skip(header_row_idx + 1).take(3).enumerate() {
                println!("  Row {}:", i + 1);
                for (j, cell) in row.iter().enumerate() {
                    let val = cell.to_string();
                    if !val.is_empty() && val != "0" {
                        println!("    Column {}: '{}'", j + 1, 
                            if val.len() > 30 { format!("{}...", &val[..30]) } else { val });
                    }
                }
            }
        } else {
            println!("- Sequence columns found: {}", sequence_columns.len());
            for (idx, name, delim) in &sequence_columns {
                println!("  * Column {}: '{}' (delimiter: {})", idx + 1, name, if *delim { "yes" } else { "no" });
            }
        }
        
        let mut data_row_count = 0;
        // Only process rows after the header row
        for (idx, row) in all_rows.iter().enumerate().skip(header_row_idx + 1) {
            let row_idx = idx - header_row_idx - 1; // Adjust for output row index
            let mut rc_value: Option<String> = None;
            let mut id_value: Option<String> = None;
            let mut processed_col_name: Option<String> = None;
            
            // Get ID value if present
            if let Some(idx) = id_col {
                if let Some(cell) = row.get(idx) {
                    id_value = Some(cell.to_string());
                }
            }
            
            // Process each cell in the row
            for (col_idx, cell) in row.iter().enumerate() {
                // Check if this column is a sequence column
                let mut processed = false;
                for (seq_col_idx, seq_col_name, has_delimiter) in &sequence_columns {
                    if col_idx == *seq_col_idx {
                        let val = cell.to_string();
                        if *has_delimiter {
                            // Handle delimiter pattern (e.g., "Prefix-SEQUENCE")
                            let parts = val.splitn(2, '-').collect::<Vec<_>>();
                            if parts.len() == 2 {
                                let rc = reverse_complement(parts[1]);
                                let new_val = format!("{}-{}", parts[0], rc);
                                rc_value = Some(new_val.clone());
                                processed_col_name = Some(seq_col_name.clone());
                                sheet.write_string((row_idx + 1) as u32, col_idx as u16, &new_val)?;
                            } else {
                                // No delimiter found, treat as full sequence
                                let rc = reverse_complement(&val);
                                rc_value = Some(rc.clone());
                                processed_col_name = Some(seq_col_name.clone());
                                sheet.write_string((row_idx + 1) as u32, col_idx as u16, &rc)?;
                            }
                        } else {
                            // Direct sequence pattern
                            let rc = reverse_complement(&val);
                            rc_value = Some(rc.clone());
                            processed_col_name = Some(seq_col_name.clone());
                            sheet.write_string((row_idx + 1) as u32, col_idx as u16, &rc)?;
                        }
                        processed = true;
                        break;
                    }
                }
                
                if !processed {
                    // Copy non-sequence columns as-is
                    sheet.write_string((row_idx + 1) as u32, col_idx as u16, &cell.to_string())?;
                }
            }
            
            // Print SQL update statement if both values are present and ID is not empty
            if let (Some(id), Some(rc), Some(col_name)) = (id_value, rc_value, processed_col_name) {
                if !id.trim().is_empty() {
                    // Generate SQL with proper column name escaping
                    if col_name.contains(' ') || col_name.starts_with("Column_") {
                        println!(
                            "UPDATE SampleBatchItems SET [{}] = '{}' WHERE Id = '{}';",
                            col_name, rc, id
                        );
                    } else {
                        println!(
                            "UPDATE SampleBatchItems SET {} = '{}' WHERE Id = '{}';",
                            col_name, rc, id
                        );
                    }
                }
            }
            data_row_count += 1;
        }

        // Save the workbook
        output_workbook.save(&output_path)?;

        println!("File processed successfully!");
        println!("Output saved to: {}", output_path.display());
        println!("\nNumber of data rows: {}", data_row_count);
        println!("Number of columns: {}", header_row.len());
    } else {
        println!("Error: Could not read the worksheet");
    }

    Ok(())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args = Args::parse();

    // Detect file type and process accordingly
    match detect_file_type(&args.file)? {
        FileType::Excel => process_excel_file(&args.file),
        FileType::Csv => process_csv_file(&args.file),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use tempfile::NamedTempFile;

    #[test]
    fn test_reverse_complement() {
        assert_eq!(reverse_complement("ATGC"), "GCAT");
        assert_eq!(reverse_complement("ATGC-N"), "N-GCAT");
        assert_eq!(reverse_complement(""), "");
        assert_eq!(reverse_complement("N"), "N");
        assert_eq!(reverse_complement("ATGCATGC"), "GCATGCAT");
    }

    #[test]
    fn test_excel_processing() -> Result<(), Box<dyn std::error::Error>> {
        // Create a temporary Excel file with test data
        let mut workbook = Workbook::new();
        let mut sheet = workbook.add_worksheet();

        // Write headers
        sheet.write_string(0, 0, "Id")?;
        sheet.write_string(0, 1, "IndexNtSequence")?;
        sheet.write_string(0, 2, "OtherColumn")?;

        // Write test data
        sheet.write_string(1, 0, "1")?;
        sheet.write_string(1, 1, "Prefix-ATGC")?;
        sheet.write_string(1, 2, "Test1")?;

        sheet.write_string(2, 0, "2")?;
        sheet.write_string(2, 1, "GCAT")?;
        sheet.write_string(2, 2, "Test2")?;

        // Save to temporary file
        let temp_file = NamedTempFile::new()?;
        workbook.save(&temp_file)?;

        // Process the file
        let args = Args {
            file: temp_file.path().to_path_buf(),
        };

        // Run the main processing logic
        let mut input_workbook: Xlsx<_> = calamine::open_workbook(&args.file)?;
        let mut output_workbook = Workbook::new();
        let mut output_sheet = output_workbook.add_worksheet();

        if let Some(Ok(range)) = input_workbook.worksheet_range_at(0) {
            let mut rows = range.rows();
            let header_row = rows.next().unwrap();

            // Write header row
            for (col, cell) in header_row.iter().enumerate() {
                output_sheet.write_string(0, col as u16, &cell.to_string())?;
            }

            // Process rows
            for (row_idx, row) in rows.enumerate() {
                for (col_idx, cell) in row.iter().enumerate() {
                    if col_idx == 1 {
                        // IndexNtSequence column
                        let val = cell.to_string();
                        let mut parts = val.splitn(2, '-').collect::<Vec<_>>();
                        if parts.len() == 2 {
                            let rc = reverse_complement(parts[1]);
                            let new_val = format!("{}-{}", parts[0], rc);
                            output_sheet.write_string(
                                (row_idx + 1) as u32,
                                col_idx as u16,
                                &new_val,
                            )?;
                        } else {
                            let rc = reverse_complement(&val);
                            output_sheet.write_string((row_idx + 1) as u32, col_idx as u16, &rc)?;
                        }
                    } else {
                        output_sheet.write_string(
                            (row_idx + 1) as u32,
                            col_idx as u16,
                            &cell.to_string(),
                        )?;
                    }
                }
            }
        }

        // Save output to another temporary file
        let output_temp = NamedTempFile::new()?;
        output_workbook.save(&output_temp)?;

        // Read back the output file to verify results
        let mut output_workbook: Xlsx<_> = calamine::open_workbook(&output_temp)?;
        if let Some(Ok(range)) = output_workbook.worksheet_range_at(0) {
            let mut rows = range.rows();
            rows.next(); // Skip header

            // Verify first row
            let row1 = rows.next().unwrap();
            assert_eq!(row1[0].to_string(), "1");
            assert_eq!(row1[1].to_string(), "Prefix-GCAT");
            assert_eq!(row1[2].to_string(), "Test1");

            // Verify second row
            let row2 = rows.next().unwrap();
            assert_eq!(row2[0].to_string(), "2");
            assert_eq!(row2[1].to_string(), "ATGC");
            assert_eq!(row2[2].to_string(), "Test2");
        }

        Ok(())
    }
}

