# VNOVIME - Excel Rows Adder

![VNOVIME](https://img.shields.io/badge/VNOVIME-Excel%20Merger-blue.svg)
![Python](https://img.shields.io/badge/Python-3.6%2B-green.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

A powerful Python script designed to merge Excel files by identifying and adding missing line items from a smaller file to a larger file, with advanced formatting and highlighting features.

## Features

- **Automatic File Detection**: Automatically detects Excel files in the current directory
- **Smart File Identification**: Identifies big and small files based on file size
- **Missing Row Detection**: Finds rows present in the small file but missing in the big file using 'Sr No' column
- **Column Alignment**: Ensures new rows match the exact column structure of the big file
- **Advanced Sorting**: Sorts merged data by Channel Name, Program Date, and Clip Start Time
- **Professional Formatting**: 
  - Centered headers with gray background
  - Auto-adjusted column widths
  - Yellow highlighting for newly added rows
  - Duration column formatting (hh:mm:ss)
- **Error Handling**: Comprehensive error handling for file operations

## Requirements

Install the required packages using:

```bash
pip install pandas openpyxl xlsxwriter
```

## Usage

1. Place the script (`vnovime.py`) in a directory containing at least 2 Excel files
2. Ensure the Excel files have a worksheet named 'expressreport'
3. Run the script:

```bash
python vnovime.py
```

## How It Works

1. **File Detection**: The script automatically detects all `.xlsx` files in the current directory (excluding the output file)
2. **Size-Based Identification**: Files are sorted by size, with the largest becoming the "big file" and the second largest becoming the "small file"
3. **Data Loading**: Both files are loaded from the 'expressreport' worksheet
4. **Missing Row Identification**: Compares 'Sr No' columns to find rows in the small file that don't exist in the big file
5. **Column Alignment**: Ensures new rows have the same column structure as the big file
6. **Merging**: Combines the original data with new rows
7. **Sorting**: Sorts the combined data by Channel Name, Program Date, and Clip Start Time (if available)
8. **Formatting**: Applies professional formatting including highlighting new rows in yellow
9. **Output**: Creates a new Excel file with "Line Items Added" suffix

## Output File

The output file will be named based on the big file:
- Original: `Report.xlsx`
- Output: `Report Line Items Added.xlsx`

## File Structure Requirements

- Excel files must contain a worksheet named 'expressreport'
- Both files must have a 'Sr No' column for row comparison
- Optional columns for sorting: 'Channel Name', 'Program Date', 'Clip Start Time'
- Duration column (if present) will be automatically formatted as hh:mm:ss

## Error Handling

The script handles various error conditions:
- Missing worksheet 'expressreport'
- Insufficient Excel files (minimum 2 required)
- Permission issues when writing output
- Memory limitations
- File access problems

## Example Output

```
╔══════════════════════════════════════════════════════════════════╗
║                                                                  ║
║   ██╗   ██╗███╗   ██╗ ██████╗ ██╗   ██╗██╗███╗   ███╗███████╗    ║
║   ██║   ██║████╗  ██║██╔═══██╗██║   ██║██║████╗ ████║██╔════╝    ║
║   ██║   ██║██╔██╗ ██║██║   ██║██║   ██║██║██╔████╔██║█████╗      ║
║   ██║   ██║██║╚██╗██║██║   ██║╚██╗ ██╔╝██║██║╚██╔╝██║██╔══╝      ║
║   ╚██████╔╝██║ ╚████║╚██████╔╝ ╚████╔╝ ██║██║ ╚═╝ ██║███████╗    ║
║    ╚═════╝ ╚═╝  ╚═══╝ ╚═════╝   ╚═══╝  ╚═╝╚═╝     ╚═╝╚══════╝    ║
║                                                                  ║
║                            V N O V I M E                         ║
║                                                                  ║
╚══════════════════════════════════════════════════════════════════╝ 

Big file: WEEK-47 BENGALI STACKED REPORT.xlsx (2.45 MB)
Small file: 20251127 (3).xlsx (0.89 MB)
Loading files...
  Successfully read 1234 rows from big file.
  Successfully read 567 rows from small file.
Found 45 new rows to add.
Starting merge process...
Successfully merged 1279 total rows.
Sorting by columns: ['Channel Name', 'Program Date', 'Clip Start Time']
Attempting to save and format the file as 'WEEK-47 BENGALI STACKED REPORT Line Items Added.xlsx'...
Highlighting new rows...
Applying duration formatting...
Found Duration column at index 8: Duration

Checking if file was created...
Success! Output file created: WEEK-47 BENGALI STACKED REPORT Line Items Added.xlsx
File size: 2.67 MB
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## Support


For support and questions, please open an issue in the repository.
