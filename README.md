# Mutual Fund Allocation Changes Analyzer

## Overview
This Python script helps analyze changes in mutual fund allocations over time by comparing Excel files containing fund allocation data.

## Features
- Load and clean mutual fund allocation data from Excel files
- Compare allocations between different time periods
- Calculate changes in:
  - Quantity of investments
  - Market value 
  - Percentage to Net Asset Value (NAV)
- Visualize top fund allocation changes
- Export detailed changes to Excel

## Prerequisites
- Python 3.x
- Required libraries:
  - pandas
  - matplotlib
  

## Installation
1. Clone the repository
2. Install required dependencies:
```bash
pip install pandas matplotlib 
```

## Usage
1. Prepare Excel files with mutual fund allocation data
   - Files should be named with fund identifier (e.g., ZN250)
   - Include fund name in filename
   - Files sorted chronologically

2. Run the script:
```bash
python excel.py
```

3. When prompted:
   - Enter fund name (e.g., ZN250)
   - Specify date range in months

## Output
- Generates Excel file with detailed allocation changes
- Creates a horizontal bar chart visualizing top allocation changes

## Example
```
Enter fund name (e.g., ZN250): ZN250
Enter date range in months (e.g., 5 for last 5 months): 5
```

## Error Handling
- Validates required Excel file columns
- Handles missing or incomplete data
- Provides informative error messages


