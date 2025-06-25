use calamine::{Reader, Xlsx};
use rust_xlsxwriter::Workbook;
use tempfile::NamedTempFile;
use tracseq_rc::reverse_complement;

fn assert_reverse_complement(input: &str, expected: &str) {
    let result = reverse_complement(input);
    assert_eq!(
        result, expected,
        "Reverse complement of '{}' should be '{}', but got '{}'",
        input, expected, result
    );
}

#[test]
fn test_reverse_complement() {
    assert_reverse_complement("ATGC", "GCAT");
    assert_reverse_complement("ATGC-N", "N-GCAT");
    assert_reverse_complement("", "");
    assert_reverse_complement("N", "N");
    assert_reverse_complement("ATGCATGC", "GCATGCAT");
}

#[test]
fn test_excel_processing() -> Result<(), Box<dyn std::error::Error>> {
    // Create a temporary Excel file with test data
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();

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
    let mut input_workbook: Xlsx<_> = calamine::open_workbook(&temp_file)?;
    let mut output_workbook = Workbook::new();
    let output_sheet = output_workbook.add_worksheet();

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
                    let parts = val.splitn(2, '-').collect::<Vec<_>>();
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

#[test]
fn test_specific_index_reverse_complement() {
    // Test a specific index from the Excel file
    let input = "GCAT";
    let expected = "ATGC";
    assert_reverse_complement(input, expected);

    // Test another index with a prefix
    let input = "Prefix-ATGC";
    let expected = "Prefix-GCAT";
    let parts = input.splitn(2, '-').collect::<Vec<_>>();
    let rc = reverse_complement(parts[1]);
    let result = format!("{}-{}", parts[0], rc);
    assert_eq!(
        result, expected,
        "Reverse complement of '{}' should be '{}', but got '{}'",
        input, expected, result
    );
}
