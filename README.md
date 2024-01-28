# Compile Excels Together into a single report using nothing but excel macros

## Overview

This Excel macro, named `Compilar`, is designed to compile data obtained from automatic web scraping of investing.com into a workbook of your choice. The compiled data provides a list of stocks with key financial metrics, facilitating the gathering of stock lists for higher-level analysis.

## Usage

1. Open the Excel workbook containing the macro.
2. Run the `Compilar` macro.

## Compilation Types

The macro supports three compilation types:

### Overview Compilation

- Creates a summary with columns: Name, Symbol, Last, Ch%, Market Cap, Vol.

### Fundamental Compilation

- Creates a summary with columns: Name, Market Cap, P/E Ratio, Revenue, Average Vol. (3m), EPS, Beta, Dividend, Dividend Yield (%).

### Technical and Performance Compilations

- Placeholder cases for future expansion.

## Instructions

1. Enter the desired compilation type in cell F9.
2. Set the destination path for the compiled summary in cell E12.
3. Specify the final summary name in cell F10.

## Code Comments

- **Lines 2-9:** Declare and initialize variables for file paths and compilation details.
- **Lines 11-29:** Create a new workbook and set headers based on the compilation type.
- **Lines 31-64:** Allow the user to select files from a folder (obtained through web scraping on investing.com) and store their names in an array.
- **Lines 66-88:** Open each selected file, copy relevant data, and paste it into the specified workbook.
- **Lines 90-103:** Display messages in case of missing files or successful compilation.

## Notes

- Ensure that the folder contains files obtained from investing.com through automatic web scraping; otherwise, an error message will be displayed.
- The compiled summary will be saved at the specified path with the designated name.
- The resulting compilation provides a comprehensive list of stocks for higher-level analysis, encompassing financial metrics crucial for informed decision-making.

--- 
