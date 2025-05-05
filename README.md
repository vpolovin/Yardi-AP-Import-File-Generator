# Yardi-AP-Import-File-Generator
Macro-enabled spreadsheet that takes invoice input and generates file with expenses distributed accounts, entities, and columns for backend input for Yardi.

# Overview

This Excel workbook automates the process for inputting accounts payables invoices into Yardi. Normally this is done one-by-one via Yardi's interface. However, for situations where there is a large volume of invoices (especially recurring) this spreadsheet removes much of the redundant/unnecessary data entry into Yardi's GUI. Once all the pertinent invoice data is entered, the code is designed to split up each invoice between the various properties, property splits, vendors, etc. There are no rounding errors because it uses Yardi's banker's rounding and the columns are placed according to Yardi's input column designation format. Once the file is exported, it can be imported into Yardi as an unposted batch for approval. I built this for one of my prior employers and it cut AP processing time by up to 50%. 

## Files Included
- `Batch Import Generator.xlsm`: The macro-enabled Excel file.
- `Module1.bas`: Exported VBA module with core logic.
- `Module1.bas`: Helper class for calculations.

## Instructions
1. Download and open the `.xlsm` file.
2. Enable macros when prompted.
3. Click the "Run" button to execute the macro.
