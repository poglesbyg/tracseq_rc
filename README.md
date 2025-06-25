# TracSeq RC - DNA Sequence Reverse Complement Tool

A command-line tool for processing Excel files containing DNA sequence data and generating reverse complements of nucleotide sequences. This tool is specifically designed for bioinformatics workflows that require DNA sequence reverse complementation.

## Features

- **Excel File Processing**: Reads Excel files (.xlsx) and processes DNA sequence data
- **Reverse Complement Generation**: Converts DNA sequences to their reverse complements
- **Multiple Column Support**: Handles various column naming conventions:
  - `IndexNtSequence` - Processes sequences after a hyphen delimiter
  - `IndexNtSequence2` - Processes entire sequence values
  - `Index 2` - Processes entire sequence values
  - `Index` - Processes sequences after a hyphen delimiter
- **SQL Statement Generation**: Outputs SQL UPDATE statements for database updates
- **Output File Creation**: Creates a new Excel file with processed data

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
tracseq_rc input_file.xlsx
```

### Example

```bash
tracseq_rc sample_sequences.xlsx
```

This will:
1. Process `sample_sequences.xlsx`
2. Generate reverse complements for DNA sequences in supported columns
3. Create an output file named `sample_sequences_RC.xlsx`
4. Print SQL UPDATE statements to the console

### Input File Format

The tool expects Excel files with:
- A header row containing column names
- One or more of the following columns:
  - `IndexNtSequence`: Sequences in format "Prefix-SEQUENCE"
  - `IndexNtSequence2`: Full sequences
  - `Index 2`: Full sequences
  - `Index`: Sequences in format "Prefix-SEQUENCE"
- Optional `Id` column for SQL statement generation

### Output

The tool generates:
1. **Output Excel File**: Named `{original_filename}_RC.xlsx` with processed sequences
2. **Console Output**: 
   - Processing status
   - SQL UPDATE statements (if `Id` column is present)
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

## License

[Add your license information here]

## Changelog

### Version 0.1.0
- Initial release
- Basic Excel file processing
- DNA reverse complement functionality
- SQL statement generation
- Support for multiple column naming conventions
