# Spreadsheet-Comparison-Report-Builder
Generate reports comparing similarities and highlighting differences in spreadsheets. Can read CSV or XLSX files. 

## Requirements
### Pandas
 - pip install pandas
___
### openpyxl
 - pip install openpyxl
___
## Feature Flags
### Multicore Processing
Runs tasks concurrently where possible
The following tasks are run as groups. Step 1 runs all together and so on. 
 - Step 1. Read table files
 - Step 2. Collect table files to arrays (raw and reordered)
 - Step 3. Perform row comparisons (raw and reordered)

Generating the discrepancies table and HTML document are run in series on one core. 

<b> Enabled by default. Use flag -n or --multioff to disable </b>
___
### Auto Number Conversion
Converts numeric strings to numbers for purpose of comparisons.
Values "1.0100" and "1.01" would be considered equal

Reads Null, NaN, or None values as 0. "NULL" and "0.0" would be considered equal
<b> Enabled be default. Use flag -nc or --noconvert to disable. </b>
___
### Case Insensitivity
Converts all cells to lower case for purpose of comparisons. 
Values "Hello World" and "HELLO world" would be considered equal

<b> Disabled by default. Use flag -c or --case to enable </b>
