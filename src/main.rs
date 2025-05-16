use calamine::{Reader, Xlsx};
use clap::Parser;
use rust_xlsxwriter::{Format, Workbook};
use std::path::PathBuf;

#[derive(Parser)]
#[command(author, version, about, long_about = None)]
struct Args {
    /// Path to the Excel file
    file: PathBuf,
}

fn reverse_complement(dna: &str) -> String {
    dna.chars()
        .rev()
        .map(|c| match c {
            'A' => 'T',
            'T' => 'A',
            'G' => 'C',
            'C' => 'G',
            'N' => 'N',
            _ => c,
        })
        .collect()
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
    let mut sheet = output_workbook.add_worksheet();

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

        // Check for IndexNtSequence or Index 2
        let indexnt_col = header_row
            .iter()
            .position(|c| c.to_string() == "IndexNtSequence");
        let index2_col = header_row.iter().position(|c| c.to_string() == "Index 2");
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
                        let mut parts = val.splitn(2, '-').collect::<Vec<_>>();
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
            if let (Some(id), Some(rc)) = (id_value, rc_value) {
                println!(
                    "UPDATE SampleBatchItems SET IndexNtSequence = '{}' WHERE Id = {};",
                    rc, id
                );
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
