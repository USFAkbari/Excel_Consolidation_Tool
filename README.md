# Excel Data Processing Tool

A Python script that processes multiple Excel files, aggregates data by name and phone number, and generates a consolidated summary report.

## Overview

This tool reads Excel files from a resource folder, concatenates all data, groups records by name and phone number, sums the count column, and creates a "Poyesh" column that counts the number of duplicate entries for each unique name+phone combination.

## Features

- **Batch Processing**: Automatically processes all Excel files (.xlsx, .xls) in the resource folder
- **Data Aggregation**: Groups data by name and phone number
- **Count Summation**: Sums the 'count' column for each unique combination
- **Duplicate Detection**: Creates a "Poyesh" column showing how many times each name+phone appears across all files
- **Sorted Output**: Results are sorted by Poyesh count in descending order

## Requirements

- Python 3.6 or higher
- pandas
- openpyxl (for Excel file reading/writing)

## Quick Start

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Create the resource folder and add your Excel files:**
   ```bash
   mkdir resorce
   # Copy your Excel files (.xlsx or .xls) into the resorce folder
   ```

3. **Run the script:**
   ```bash
   python process_excel.py
   ```

4. **Find your results:**
   - Output file: `output/All_poyesh.xlsx`

## Installation

1. Clone or download this repository

2. Install required dependencies:

```bash
pip install pandas openpyxl
```

Or using requirements.txt:

```bash
pip install -r requirements.txt
```

## Project Structure

```
.
├── process_excel.py    # Main processing script
├── resorce/            # Input folder (place your Excel files here)
│   └── *.xlsx         # Excel files to process
├── output/             # Output folder (generated automatically)
│   └── All_poyesh.xlsx # Consolidated output file
└── README.md           # This file
```

## How to Use

### Step 1: Prepare Your Data

1. Create a folder named `resorce` in the project root (if it doesn't exist)
2. Place all your Excel files (.xlsx or .xls) in the `resorce` folder
3. Ensure your Excel files have the following columns:
   - `name`: Person's name
   - `phone`: Phone number
   - `count`: Numeric count value
   - Additional date columns are optional and will be ignored

### Step 2: Run the Script

Execute the processing script:

```bash
python process_excel.py
```

Or on Linux/Mac:

```bash
python3 process_excel.py
```

### Step 3: Check Results

The script will:
- Display progress information in the console
- Create an `output` folder automatically
- Generate `All_poyesh.xlsx` in the `output` folder

### Output File Structure

The generated `All_poyesh.xlsx` contains:

| Column | Description |
|--------|-------------|
| `name` | Person's name |
| `phone` | Phone number |
| `count` | Sum of all count values for this name+phone combination |
| `Poyesh` | Number of times this name+phone combination appears across all input files |

The data is sorted by `Poyesh` in descending order (highest duplicates first).

## Example Output

```
Found 19 Excel files to process
Read file1.xlsx: 827 rows
Read file2.xlsx: 1214 rows
...
Total rows after concatenation: 77974

Saved All_poyesh.xlsx with 23131 rows
Output saved to: output/All_poyesh.xlsx
```

## Error Handling

- The script will continue processing even if some files fail to read
- Error messages will be displayed for problematic files
- The script will create the output folder if it doesn't exist

## Notes

- The script processes all Excel files in the `resorce` folder
- Files are processed in alphabetical order
- Duplicate name+phone combinations across files are aggregated
- The output file overwrites any existing `All_poyesh.xlsx` file

## Troubleshooting

### Common Issues

1. **ModuleNotFoundError**: Install missing dependencies with `pip install pandas openpyxl`
2. **FileNotFoundError**: Ensure the `resorce` folder exists and contains Excel files
3. **Empty Output**: Check that your Excel files have the required columns (`name`, `phone`, `count`)
4. **Encoding Issues**: If you encounter encoding errors with non-ASCII characters, ensure your Excel files are properly encoded

## License

This project is provided as-is for data processing purposes.

## Contributing

Feel free to modify the script to suit your specific needs. Common modifications:
- Change output file name or location
- Add additional aggregation columns
- Modify sorting criteria
- Add filtering options

