# Universal File Handler

This script is designed to provide a unified approach to handle CSV and Excel files for various operations including extracting specific columns, searching the entire file, and searching within a specific column.

## Features

- **Extract Columns**: Extract specified columns from a CSV or Excel file and write them to a new file.
- **Search CSV**: Search the entire CSV or Excel file using a string or regex pattern.
- **Search Column**: Search a specific column in a CSV or Excel file using a string or regex pattern.

## Requirements

- Python 3.x
- `openpyxl` library for handling .xlsx files.
- `xlrd` library for handling .xls files.
- `csv` library for handling .csv files (comes with standard Python).

## Installation

Ensure you have the required Python version and libraries installed:

```bash
pip install openpyxl xlrd
```

## Usage
Run the script from the command line, providing the necessary arguments:

```bash
python ufhpy --filename 'path/to/your/file.csv'
Command-line Arguments
filename: The file to process. (Required)
--extract: Columns to extract. Provide column names separated by spaces.
--newfile: Name of the new file when extracting columns or saving search results.
--search: String or regex to search in the entire file.
--searchcol: Column to perform search on.
--pattern: String or regex pattern to search for.
--output: Output choice for search results (print or csv). Default is print.
```

## Example
Extract columns from a file:

```bash
python ufh.py --filename 'data.csv' --extract 'Column1' 'Column2' --newfile 'extracted.csv'
```

Search entire file:
```bash
python universal_file_handler.py --filename 'data.csv' --search 'pattern'
```

Search specific column:

```bash
python universal_file_handler.py --filename 'data.csv' --searchcol 'Column1' --pattern 'pattern' --output 'csv'
```

## Disclaimer
This script assumes a simple, flat table structure for Excel files without considering merged cells, formulas, or other complexities. For real-world applications, you might need to expand or modify the code to handle such scenarios.

## Contributing
Contributions, issues, and feature requests are welcome.

License
This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.




