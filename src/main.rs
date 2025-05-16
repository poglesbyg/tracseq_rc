use calamine::{Reader, Xlsx};
use clap::Parser;
use std::path::PathBuf;
use xlsxwriter::Workbook;

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

    // Open the Excel file
    let mut workbook: Xlsx<_> = calamine::open_workbook(&args.file)?;

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
    let workbook = Workbook::new(&output_path)?;
    let mut sheet = workbook.add_worksheet(None)?;

    // Get the first sheet
    if let Some(Ok(range)) = workbook.worksheet_range_at(0) {
        println!("\nProcessing file...");

        // Find the row that starts with "Sample Name"
        let mut found_sample_name = false;
        let mut data_rows = Vec::new();
        let mut header_row = None;
        let mut index2_col = None;

        for row in range.rows() {
            if !found_sample_name {
                // Check if this row starts with "Sample Name"
                if let Some(first_cell) = row.first() {
                    if first_cell.to_string() == "Sample Name" {
                        found_sample_name = true;
                        // Find the Index 2 column
                        for (i, cell) in row.iter().enumerate() {
                            if cell.to_string() == "Index 2" {
                                index2_col = Some(i);
                                break;
                            }
                        }
                        header_row = Some(row.to_vec());
                    }
                }
            } else {
                // After finding "Sample Name", collect all subsequent rows
                data_rows.push(row.to_vec());
            }
        }

        // Write header row if found
        if let Some(header) = header_row {
            for (col, cell) in header.iter().enumerate() {
                sheet.write_string(0, col as u16, &cell.to_string(), None)?;
            }
        }

        // Write all data rows with reverse complement of Index 2
        for (row_idx, row) in data_rows.iter().enumerate() {
            for (col_idx, cell) in row.iter().enumerate() {
                if Some(col_idx) == index2_col {
                    // Write reverse complement of Index 2
                    sheet.write_string(
                        (row_idx + 1) as u32,
                        col_idx as u16,
                        &reverse_complement(&cell.to_string()),
                        None,
                    )?;
                } else {
                    sheet.write_string(
                        (row_idx + 1) as u32,
                        col_idx as u16,
                        &cell.to_string(),
                        None,
                    )?;
                }
            }
        }

        // Close the workbook to save the file
        workbook.close()?;

        println!("File processed successfully!");
        println!("Output saved to: {}", output_path.display());
        println!("\nNumber of data rows: {}", data_rows.len());
        if !data_rows.is_empty() {
            println!("Number of columns: {}", data_rows[0].len());
        }
    } else {
        println!("Error: Could not read the worksheet");
    }

    Ok(())
}
