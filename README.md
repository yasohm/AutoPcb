# Excel Data Converter

A powerful tool for processing Excel files and generating output with automatic data lookup and calculations.

## Quick Start

```bash
# 1. Place files in input/ directory
# 2. Run converter
./run.sh
```

## Project Structure

```
the_converter/
├── output.xlsx                  # Generated output
├── modif.c                      # Main C program
├── file_utils.py                # File preprocessing utilities
├── import_xlsx_to_sqlite.py     # Import ABC/FB/PCB into SQLite
├── export_sqlite_to_xlsx.py     # Export abc/fb/pcb tables back to input/*.xlsx
├── run.sh                       # Execution script
├── run_with_error_handling.sh   # Execution with error handling
├── Makefile                     # Compilation rules
└── README.md                    # This file
```

## Features

### Excel File Processing
- Excel file processing (PCB.xlsx, ABC.xlsx, FB.xlsx)
- Automatic file format detection (.xls/.xlsx)
- Column name handling (French to English)
- Error handling and validation

### Data Processing
- WLOM values: From ABC.xlsx WKQCO column
- FB values: From FB.xlsx REF column (automatic week detection)
- MAX calculation: Maximum of FB and WCMJ values
- Automatic week detection for FB data

## Installation

### Dependencies
```bash
# Install required libraries
make install-deps

# Or manually:
sudo apt-get install libxlsxio-dev libxlsxwriter-dev
```

### Python Dependencies
```bash
pip3 install -r requirements.txt
```

## Usage

### Basic Usage
```bash
# Place files in input/ directory
mkdir -p input
cp PCB.xlsx ABC.xlsx FB.xlsx input/

# Run converter
./run.sh
```

### With Error Handling
```bash
# Run with comprehensive error handling
./run_with_error_handling.sh
```

### Import .xlsx Files Into SQLite
```bash
# Import input/ABC.xlsx, input/FB.xlsx, input/PCB.xlsx into data.db
make db-import

# Or directly
python3 import_xlsx_to_sqlite.py
```
- Creates `data.db` in the project root
- Tables: `abc`, `fb`, `pcb`
- Reads with pandas/openpyxl and replaces tables on each run

### Export Tables Back to input/*.xlsx
```bash
# Export tables abc, fb, pcb from data.db to input/ABC.xlsx, input/FB.xlsx, input/PCB.xlsx
make db-export

# Or directly
python3 export_sqlite_to_xlsx.py
```
- Exports each table to a single-sheet `.xlsx` file (sheet named after table)
- Skips tables not present in the DB

## Development

### Compile Program
```bash
# Compile the program
make modif

# Clean build
make clean
```

## Input File Requirements

### PCB File
- Format: .xls or .xlsx
- Columns: WIDF, WCMJ, WLOM, FB, MAX
- Location: `input/PCB.xlsx`

### ABC File
- Format: .xlsx
- Columns: WKIDF (col 0), WKQCO (col 4)
- Location: `input/ABC.xlsx`

### FB File
- Format: .xlsx
- Columns: REF (col 0), week columns (34-52, 1-8)
- Location: `input/FB.xlsx`

## Output

### Generated Files
- output.xlsx: Generated output file

### Output Columns
1. WIDF: Product identifier
2. WCMJ: WCMJ value from PCB
3. WLOM: WLOM value from ABC (WKQCO)
4. FB: FB value from FB (automatic week detection)
5. MAX: Maximum of FB and WCMJ

## Troubleshooting

### Common Issues
1. Compilation errors: Install dependencies with `make install-deps`
2. Python errors: Install dependencies with `pip3 install -r requirements.txt`
3. Input files missing: Ensure files are in `input/` directory
4. File format issues: Use `./run_with_error_handling.sh` for automatic conversion

### Error Handling
- Automatic file format detection
- Column name normalization
- Data type validation
- Comprehensive error messages

## Support

For issues or questions:
1. Check the error messages
2. Verify input files are in the correct location
3. Use `./run_with_error_handling.sh` for better error reporting
# AutoPcb
