# Trial Balance Consolidator

  A Python desktop application that consolidates multiple Trial Balance Excel files into a single file with entity names, account categories, and mapping lookups.

  ## Features

  - **Batch Processing**: Select a folder of Trial Balance Excel files to consolidate
  - **Entity Name Extraction**: Automatically reads company name from cell A1
  - **Category Classification**: Detects account types from cell formatting:
    - Group Accounts (background fill)
    - Sub-GL Accounts (italic)
    - Control Accounts (bold, no fill)
    - GL Accounts (no formatting)
  - **Mapping File Support**: Optional mapping file for:
    - Entity Code lookup (Company_Code sheet, matched via Company Name-ERP)
    - Company GL Code + FSLI details (FSLI_Code sheet, matched via GL Description)
  - **Balance Calculations**: Opening Balance and Closing Balance computed automatically
  - **Control Sheet**: Summary sheet with closing balance totals per company (GL + Control Accounts)
  - **Currency Column**: Default "INR" for all rows

  ## Output Columns

  | Column | Description |
  |--------|-------------|
  | Period | Blank (for manual entry) |
  | Company GL Code | From FSLI mapping |
  | Company GL Description | Account head from TB |
  | Company Name | Entity name from TB header |
  | Company Code | From Company_Code mapping |
  | Opening Balance | Opening Debit - Opening Credit |
  | Debit Amount | Period Debit |
  | Credit Amount | Period Credit |
  | Closing Balance | Closing Debit - Closing Credit |
  | Currency | Default "INR" |
  | Category | Auto-detected from formatting |
  | FS Header | From FSLI mapping |
  | FS Account Type | From FSLI mapping |
  | FS Account Sub-Type | From FSLI mapping |
  | FSLI | From FSLI mapping |

  ## Requirements

  - Python 3.8+
  - openpyxl

  ## Installation

  ```bash
  pip install -r requirements.txt
  ```

  ## Usage

  ```bash
  python tb_consolidator.py
  ```

  1. Click **Browse** next to "Source Folder" to select the folder with your Trial Balance Excel files
  2. (Optional) Click **Browse** next to "Mapping File" to select your mapping Excel file
  3. Click **Consolidate & Save** and choose where to save the output
  