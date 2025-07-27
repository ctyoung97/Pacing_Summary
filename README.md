# Q3 Pacing Analysis Tool

This Python script automatically processes Q3 pacing data from Excel files and generates comprehensive analysis reports.

## Overview

The tool processes two Excel files (current week and prior week) and generates two analysis reports:

1. **Q3 Pacing Setup with Core Revenue** - Analyzes Q3 bookings, pacing data, and core revenue growth
2. **Q3 Billings Prior Year** - Compares current week vs prior week core billing data

## File Structure

```
Pacing_Summary/
├── inputs/                     # Place your Excel files here
├── output/                     # Generated reports appear here
├── pacing_analyzer.py          # Main analysis script
├── run_analysis.bat           # Easy-to-use batch file
├── prompts.txt                # Analysis specifications
└── README.md                  # This file
```

## How to Use

### Method 1: Using the Batch File (Recommended)
1. Place your Excel files in the `inputs/` folder
2. Double-click `run_analysis.bat`
3. The analysis will run automatically
4. Check the `output/` folder for `Q3_Pacing_Analysis_{prior_date}_{current_date}.xlsx`

### Method 2: Using Command Line
1. Open command prompt in this directory
2. Activate virtual environment: `.venv\Scripts\activate`
3. Run the script: `python pacing_analyzer.py`

## Input Requirements

- **File Location**: Place Excel files in the `inputs/` folder
- **File Naming**: Files must contain date patterns like "07.21.25" and "07.14.25"
- **File Format**: Excel files (.xlsx) with station worksheets before "MTH Summary" sheet

## Output

The tool generates `Q3_Pacing_Analysis_{prior_date}_{current_date}.xlsx` with two sheets:

### Sheet 1: Q3 Pacing Setup
- Station names
- Total Q3 Bookings (cell C40)
- New Core Billing in Prior Year Same Week (difference between current and prior week core bookings)
- Local/National/Digital Q3 Pace (cells K34, K35, K37)
- Total Q3 Pace (cell K40)
- Core Rev Q3 2025 (D40 - D36)
- Core Rev Q3 2024 (F40 - F36)
- Core Rev Growth % (formatted as percentage)

### Sheet 2: Q3 Billings Prior Year
- Station names
- Core Bookings Current Week (F40 - F36 from current file)
- Core Bookings Prior Week (F40 - F36 from prior file)
- New Core Billing in Prior Year Same Week (difference)

## Features

- **Automatic Station Detection**: Only processes worksheets before "MTH Summary"
- **Smart Filtering**: Excludes summary/template sheets (QTD Summ by Station, YTD Summ by Station, Station Summary, Instructions, ->, AMB Corp, <-)
- **Intelligent Organization**: Moves QTR Summary to bottom of report
- **Professional Formatting**: 
  - Dollar columns formatted as $#,##0 (Total Q3 Bookings, New Core Billing, Core Rev columns)
  - Percentage columns formatted as 0% (Local/National/Digital/Total Q3 Pace, Core Rev Growth %)
  - Auto-adjusts column widths
- **Error Handling**: Safely handles missing cells or worksheets
- **Detailed Logging**: Shows processing status and exclusions for each station

## Technical Details

- **Language**: Python 3.13+
- **Dependencies**: pandas, openpyxl
- **Virtual Environment**: Uses `.venv` for isolated package management

## Troubleshooting

- **Missing Files**: Ensure Excel files are in `inputs/` folder with correct date patterns
- **Missing Stations**: Check that station worksheets appear before "MTH Summary" sheet
- **Cell Errors**: Script handles missing/invalid cells by using 0 as default value

## Updating Input Files

Simply replace the files in the `inputs/` folder and run the analysis again. The script automatically detects the correct files based on date patterns in filenames.
