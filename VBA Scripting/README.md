# VBA-Challenge
# Quarterly Stock Data Summary

## Table of Contents
- [Description](#description)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Example](#example)
- [Author](#author)
- [License](#license)
- [Contributing](#contributing)
- [Contact](#contact)

## Description

**Quarterly Stock Data Summary** is an Excel VBA macro designed to analyze and summarize stock performance data across multiple quarterly worksheets. This tool automates the process of calculating key metrics such as quarterly change, percentage change, and total volume for each stock ticker. Additionally, it identifies the tickers with the greatest percentage increase, decrease, and total volume, providing a comprehensive overview of stock performance for each quarter.

## Features

- **Automated Data Processing:** Iterates through worksheets named with a "Q" prefix (e.g., Q1, Q2) to process quarterly data.
- **Key Metrics Calculation:** Computes quarterly change, percentage change, and total volume for each stock ticker.
- **Highlighting Changes:** Applies color coding to indicate positive (light green) and negative (light red) changes in quarterly performance.
- **Top Performers Identification:** Identifies and displays the tickers with the greatest percentage increase, decrease, and total volume.
- **Performance Optimization:** Disables screen updating and automatic calculations during execution to enhance performance.

## Installation

1. **Prerequisites:**
   - Microsoft Excel (compatible with VBA macros)
   - Basic understanding of Excel and VBA

2. **Setup:**
   - Download the `QuarterlyStockDataSummary` VBA macro file.
   - Open your Excel workbook containing the quarterly stock data.
   - Press `ALT + F11` to open the VBA editor.
   - In the VBA editor, insert a new module:
     - Right-click on your workbook in the Project pane.
     - Select `Insert` > `Module`.
   - Copy and paste the VBA macro code into the new module.
   - Save your workbook as a macro-enabled file (`.xlsm`).

## Usage

1. **Prepare Your Data:**
   - Ensure each quarterly worksheet is named with a "Q" prefix (e.g., Q1, Q2, Q3, Q4).
   - Organize your data with the following columns:
     - **Column A:** Ticker Symbol
     - **Column C:** Opening Price
     - **Column F:** Closing Price
     - **Column G:** Volume

2. **Run the Macro:**
   - Open the Excel workbook containing the macro.
   - Press `ALT + F8` to open the Macro dialog box.
   - Select `QuarterlyStockDataSummary` from the list.
   - Click `Run`.

3. **Review the Summary:**
   - The macro will add a summary section to each quarterly worksheet, displaying calculated metrics and highlighting significant changes.
   - It will also identify the tickers with the greatest percentage increase, decrease, and total volume.

## Example

*After running the macro, your quarterly worksheet (e.g., Q1) will include a summary section similar to the following:*

| Ticker | Quarterly Change | Percentage Change | Total Volume |
|--------|-------------------|---------------------|--------------|
| AAPL   | 15.00             | 5.00%               | 1,500,000    |
| MSFT   | -10.00            | -3.33%              | 2,000,000    |
| GOOGL  | 20.00             | 10.00%              | 1,200,000    |

**Greatest % Increase**
- **Ticker:** GOOGL
- **Value:** 10.00%

**Greatest % Decrease**
- **Ticker:** MSFT
- **Value:** -3.33%

**Greatest Total Volume**
- **Ticker:** MSFT
- **Value:** 2,000,000

*Positive quarterly changes are highlighted in light green, while negative changes are highlighted in light red.*

## Author

**[Kendall Burkett]**  
[https://github.com/KendallBurkett?tab=repositories]
[kbz1987@icloud.com]



## Additional Notes

- **Performance Optimization:** The macro temporarily disables screen updating and automatic calculations to speed up the processing. These settings are restored upon completion.
  
- **Error Handling:** The macro includes basic error handling to manage duplicate tickers gracefully.

- **Customization:** You can modify the color codes or add additional metrics as needed to suit your specific requirements.