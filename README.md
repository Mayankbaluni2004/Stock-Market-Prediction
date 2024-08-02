# Correlation Analysis of Financial Datasets

This script analyzes the correlation between datasets in `Reliance Dataset.xlsx` and `Adani Dataset.xlsx` over several weeks.

## Prerequisites

Install the required libraries:
```sh
pip install openpyxl numpy matplotlib
```

## Files

Ensure the following files are in the same directory:
- `Reliance Dataset.xlsx`
- `Adani Dataset.xlsx`

## How to Run

1. Save the script as `correlation_analysis.py`.
2. Place the required Excel files in the same directory.
3. Run the script:
   ```sh
   python correlation_analysis.py
   ```

## Output

1. **Printed Correlation Coefficients:** For each week.
2. **Plot:** Displays the correlation coefficients over weeks.

## Code Overview

### Functions

- `extract_values_from_sheet(sheet, k, column)`: Extracts values from a specified column and range.
- `calculate_correlation(sheet1, sheet2, k, column)`: Calculates the correlation coefficient between two columns.
- `process_correlation(columns)`: Processes and prints correlation coefficients for specified columns over weeks.

### Main Execution

- Defines columns to process (`['B', 'C', 'D', 'E']`).
- Calls `process_correlation(columns_to_process)`.
- Prints and plots the weekly correlation coefficients.

---
