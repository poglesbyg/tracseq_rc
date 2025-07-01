# TracSeq RC - DNA Sequence Reverse Complement Tool

A command-line tool for processing Excel and CSV files containing DNA sequence data and generating reverse complements of nucleotide sequences. This tool is specifically designed for bioinformatics workflows that require DNA sequence reverse complementation.

## Features

- **Excel and CSV File Processing**: Reads both Excel files (.xlsx, .xls) and CSV files (.csv) containing DNA sequence data
- **Flexible Sequence Detection**: Automatically detects columns containing DNA sequences, even if they don't match standard naming conventions
- **Report Format Support**: Handles complex Excel reports by finding data sections starting with "Sample ID" 
- **Reverse Complement Generation**: Converts DNA sequences to their reverse complements
- **Multiple Column Support**: Handles various column naming conventions:
  - `IndexNtSequence` - Processes sequences after a hyphen delimiter
  - `IndexNtSequence2` - Processes entire sequence values
  - `Index 2` - Processes entire sequence values
  - `Index` - Processes sequences after a hyphen delimiter
  - Any column containing DNA sequences (automatic detection)
- **SQL Statement Generation**: Outputs SQL UPDATE statements to terminal for database updates
- **Output File Creation**: Creates a new Excel or CSV file with processed data (matching input format)

## Installation

### Prerequisites

- Rust (latest stable version)
- Cargo package manager

### Build from Source

```bash
git clone <repository-url>
cd tracseq_rc
cargo build --release
```

The executable will be available at `target/release/tracseq_rc.exe` (Windows) or `target/release/tracseq_rc` (Unix-like systems).

## Usage

### Basic Usage

```bash
tracseq_rc input_file.xlsx  # For Excel files
tracseq_rc input_file.csv   # For CSV files
```

### Examples

#### Excel File Example

```bash
tracseq_rc sample_sequences.xlsx
```

This will:
1. Process `sample_sequences.xlsx`
2. Generate reverse complements for DNA sequences in supported columns
3. Create an output file named `sample_sequences_RC.xlsx`
4. Print SQL UPDATE statements to the console

#### CSV File Example

```bash
tracseq_rc sample_sequences.csv
```

Output:
```
Processing CSV file...
UPDATE SampleBatchItems SET IndexNtSequence = 'Prefix-GCAT' WHERE Id = 1001;
UPDATE SampleBatchItems SET IndexNtSequence = 'Prefix-TAGC' WHERE Id = 1002;
UPDATE SampleBatchItems SET IndexNtSequence = 'TTAA' WHERE Id = 1003;
UPDATE SampleBatchItems SET IndexNtSequence = 'Prefix-GGCC' WHERE Id = 1004;
File processed successfully!
Output saved to: sample_sequences_RC.csv

Number of data rows: 4
Number of columns: 3
```

### Input File Format

The tool accepts both Excel (.xlsx, .xls) and CSV (.csv) files with:
- A header row containing column names
- One or more of the following columns:
  - `IndexNtSequence`: Sequences in format "Prefix-SEQUENCE"
  - `IndexNtSequence2`: Full sequences
  - `Index 2`: Full sequences
  - `Index`: Sequences in format "Prefix-SEQUENCE"
- Optional `Id` column for SQL statement generation

### Output

The tool generates:
1. **Output File**: 
   - Excel files: Named `{original_filename}_RC.xlsx` with processed sequences
   - CSV files: Named `{original_filename}_RC.csv` with processed sequences
2. **Console Output**: 
   - Processing status
   - SQL UPDATE statements (printed to terminal if `Id` column is present)
   - Summary statistics

## DNA Reverse Complement Logic

The tool converts DNA bases according to standard Watson-Crick base pairing:
- A ↔ T
- G ↔ C
- N → N (ambiguous bases remain unchanged)

The sequence is also reversed (3' to 5' direction becomes 5' to 3').

### Examples

- `ATGC` → `GCAT`
- `AAATTTGGGCCC` → `GGGCCCAAATTT`
- `Prefix-ATGC` → `Prefix-GCAT` (for delimited formats)

## Development

### Running Tests

```bash
# Run all tests
cargo test

# Run integration tests only
cargo test --test integration_tests

# Run with verbose output
cargo test -- --nocapture
```

### Project Structure

```
tracseq_rc/
├── src/
│   ├── main.rs          # Main application logic
│   └── lib.rs           # Reverse complement function
├── tests/
│   └── integration_tests.rs  # Integration tests
├── Cargo.toml           # Project configuration
└── README.md            # This file
```

### Dependencies

- `calamine` - Excel file reading
- `clap` - Command-line argument parsing
- `rust_xlsxwriter` - Excel file writing
- `tempfile` - Temporary file handling (dev dependency)

## Error Handling

The tool handles common error scenarios:
- Invalid or missing input files
- Unsupported file formats
- Empty worksheets
- Missing expected columns

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Run the test suite
6. Submit a pull request

### Version 0.2.0
- Added flexible sequence detection that automatically finds DNA sequence columns
- Added support for complex Excel reports with "Sample ID" sections
- Improved DNA pattern detection with stricter validation
- Fixed SQL generation to quote ID values and skip empty IDs
- Added support for both "Id" and "Sample ID" column names

### Version 0.1.1
- Added CSV file support (.csv)
- Fixed SQL statement generation for all supported column types
- Fixed reverse complement logic for Index column
- SQL statements now print correctly for IndexNtSequence2 and Index 2 columns

### Version 0.1.0
- Initial release
- Basic Excel file processing
- DNA reverse complement functionality
- SQL statement generation
- Support for multiple column naming conventions
