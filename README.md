# Excel File Modifier

This Python script is designed to process two files: an Excel file and a text file. The script reads the text file, extracts data in the form of key-value pairs (with optional units), and then modifies the corresponding rows in the Excel file based on the extracted data.

## Installation

1. Clone this repository to your local machine.

```
git clone https://github.com/tlond/excel-file-modifier.git
```
2. Navigate to the project directory.

```
cd excel-file-modifier
```

3. Install the required dependencies using pip.

```
pip install openpyxl
```
## Dependencies

- `re`: Regular Expression module for pattern matching
- `os`: Operating System module for handling file paths
- `openpyxl`: Library to interact with Excel files
- `warnings`: Module to handle warnings
- `logging`: Module for logging messages
- `datetime`: Module for working with dates and times
- `argparse`: Module for parsing command-line arguments

## Global Settings

The script contains some global settings that can be customized:

- `key_name_coulmn_index`: The column index (0-based) of the key names in the Excel file.
- `unit_coulmn_index`: The column index (0-based) of the units in the Excel file.
- `order_coulmn_index`: The column index (0-based) of the order in the Excel file.
- `sheet_name`: The name of the sheet in the Excel file to work with.

## Logging

The script sets up logging to record debug and info messages. Log messages are printed to the console and stored in a log file named "log.txt" with UTF-8 encoding.

## Functions

1. `split_string_by_integer(s)`: Splits a string into three parts: text before the first integer, the integer (value), and the text after the integer (unit).
2. `fix_unit(input_str)`: Replaces an empty or "банка" unit with "шт".
3. `get_downloads_folder()`: Gets the path to the user's Downloads folder.
4. `convert_units(value, input_unit, desired_unit)`: Converts the value from the input unit to the desired unit.
5. `class ExcelModifier`: A class to open, modify, and save Excel files.
   - `__init__(self, filename, sheetname)`: Initializes the ExcelModifier instance with the filename and optional sheetname.
   - `find_row_by_key(self, column_index, key)`: Finds the row index in the Excel sheet that matches the given key.
   - `set_row_key_value(self, key, value, unit)`: Modifies the row in the Excel sheet based on the key, value, and unit.
   - `save(self, output_filename)`: Saves the modified Excel file to a new file with the given output_filename.
   - `close(self)`: Closes the Excel file.

## Usage

The script can be executed from the command line, taking two mandatory arguments:

1. `excel`: The path to the Excel file to be used as a tamplate.
2. `text`: The path to the text file containing key-value pairs (with optional units).

Example usage:

```
python script.py path/to/excel_file.xlsx path/to/text_file.txt
```

The script will read the text file, extract key-value pairs, and modify the corresponding rows in the Excel file based on the extracted data. The modified Excel file will be saved at the current directory with a timestamped filename in the same directory as the original Excel file.