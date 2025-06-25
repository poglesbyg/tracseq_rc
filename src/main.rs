use calamine::{Reader, Xlsx};
use clap::Parser;
use rust_xlsxwriter::Workbook;
use std::path::PathBuf;
use tracseq_rc::reverse_complement;

#[derive(Parser)]
#[command(author, version, about, long_about = None)]
struct Args {
    /// Path to the Excel file
    file: PathBuf,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args = Args::parse();

    // Open the input Excel file
    let mut input_workbook: Xlsx<_> = calamine::open_workbook(&args.file)?;

    // Create output filename
    let output_path = args.file.with_file_name(format!(
        "{}_RC{}",
        args.file.file_stem().unwrap().to_string_lossy(),
        args.file
            .extension()
            .map(|ext| format!(".{}", ext.to_string_lossy()))
            .unwrap_or_default()
    ));

    // Create new workbook for output
    let mut output_workbook = Workbook::new();
    let sheet = output_workbook.add_worksheet();

    // Get the first sheet
    if let Some(Ok(range)) = input_workbook.worksheet_range_at(0) {
        println!("\nProcessing file...");

        // Find the header row (first row)
        let mut rows = range.rows();
        let header_row = match rows.next() {
            Some(row) => row,
            None => {
                println!("No data found in the sheet.");
                return Ok(());
            }
        };

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
        let id_col = header_row.iter().position(|c| c.to_string() == "Id");

        let mut data_row_count = 0;
        for (row_idx, row) in rows.enumerate() {
            let mut rc_value: Option<String> = None;
            let mut id_value: Option<String> = None;
            for (col_idx, cell) in row.iter().enumerate() {
                if let Some(idx) = indexnt_col {
                    // Handle IndexNtSequence logic
                    if col_idx == idx {
                        let val = cell.to_string();
                        let parts = val.splitn(2, '-').collect::<Vec<_>>();
                        if parts.len() == 2 {
                            let rc = reverse_complement(parts[1]);
                            let new_val = format!("{}-{}", parts[0], rc);
                            rc_value = Some(new_val.clone());
                            sheet.write_string((row_idx + 1) as u32, col_idx as u16, &new_val)?;
                        } else {
                            rc_value = Some(val.clone());
                            sheet.write_string((row_idx + 1) as u32, col_idx as u16, &val)?;
                        }
                    } else {
                        sheet.write_string(
                            (row_idx + 1) as u32,
                            col_idx as u16,
                            &cell.to_string(),
                        )?;
                    }
                } else if let Some(idx) = indexnt2_col {
                    // Handle IndexNtSequence2 logic - reverse complement entire value
                    if col_idx == idx {
                        let rc = reverse_complement(&cell.to_string());
                        rc_value = Some(rc.clone());
                        sheet.write_string((row_idx + 1) as u32, col_idx as u16, &rc)?;
                    } else {
                        sheet.write_string(
                            (row_idx + 1) as u32,
                            col_idx as u16,
                            &cell.to_string(),
                        )?;
                    }
                } else if let Some(idx) = index2_col {
                    // Handle Index 2 logic
                    if col_idx == idx {
                        let rc = reverse_complement(&cell.to_string());
                        rc_value = Some(rc.clone());
                        sheet.write_string((row_idx + 1) as u32, col_idx as u16, &rc)?;
                    } else {
                        sheet.write_string(
                            (row_idx + 1) as u32,
                            col_idx as u16,
                            &cell.to_string(),
                        )?;
                    }
                } else if let Some(idx) = index_col {
                    // Handle Index logic - reverse complement second part after '-'
                    if col_idx == idx {
                        let val = cell.to_string();
                        let parts = val.splitn(2, '-').collect::<Vec<_>>();
                        if parts.len() == 2 {
                            let rc = reverse_complement(parts[0]);
                            let new_val = format!("{}-{}", parts[0], rc);
                            rc_value = Some(new_val.clone());
                            sheet.write_string((row_idx + 1) as u32, col_idx as u16, &new_val)?;
                        } else {
                            rc_value = Some(val.clone());
                            sheet.write_string((row_idx + 1) as u32, col_idx as u16, &val)?;
                        }
                    } else {
                        sheet.write_string(
                            (row_idx + 1) as u32,
                            col_idx as u16,
                            &cell.to_string(),
                        )?;
                    }
                } else {
                    // No special columns, just copy
                    sheet.write_string((row_idx + 1) as u32, col_idx as u16, &cell.to_string())?;
                }
                if let Some(idx) = id_col {
                    if col_idx == idx {
                        id_value = Some(cell.to_string());
                    }
                }
            }
            // Print SQL update statement if both values are present
            if let (Some(id), Some(rc)) = (id_value.clone(), rc_value.clone()) {
                if let Some(idx) = indexnt_col {
                    if rc_value.is_some() && id_value.is_some() {
                        println!(
                            "UPDATE SampleBatchItems SET IndexNtSequence = '{}' WHERE Id = {};",
                            rc, id
                        );
                    }
                } else if let Some(idx) = index_col {
                    if rc_value.is_some() && id_value.is_some() {
                        println!(
                            "UPDATE SampleBatchItems SET [Index] = '{}' WHERE Id = {};",
                            rc, id
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
