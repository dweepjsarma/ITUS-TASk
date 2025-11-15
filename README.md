# Excel UDF Setup Guide for Database Integration

This guide will help you set up Python User-Defined Functions (UDFs) in Excel to fetch data from your SQLite database.

## Prerequisites

1. **Python 3.7+** installed on your system
2. **Microsoft Excel** (Windows or Mac with xlwings support)
3. **SQLite database** file (ebitda_margins.db)

## Installation Steps

### Step 1: Install Required Python Packages

Open Command Prompt (Windows) or Terminal (Mac/Linux) and run:

```bash
pip install xlwings openpyxl
```

### Step 2: Configure xlwings

Run the following command to set up xlwings:

```bash
xlwings addin install
```

This installs the xlwings add-in for Excel.

### Step 3: File Organization

Place all files in the same directory:

```
your_project_folder/
├── ebitda_margins_data_udf_updated.py    # Python script with UDF functions
├── config.ini                             # Configuration file
├── ebitda_margins.db                      # Your SQLite database
└── example.xlsx                           # Your Excel workbook
```

### Step 4: Update config.ini

Edit `config.ini` to match your database setup:

```ini
[database]
path = ebitda_margins.db        # Update if database has different name
table = financial_data          # Update if table has different name
date_format = %Y-%m-%d          # Update if using different date format
```

### Step 5: Test Python Script

Before using in Excel, test the script standalone:

```bash
python ebitda_margins_data_udf_updated.py
```

This will show sample data from your database and verify connectivity.

## Using UDFs in Excel

### Method 1: Direct xlwings Integration (Recommended)

1. **Open Excel** and your workbook (example.xlsx)

2. **Enable the xlwings add-in**:
   - Go to Excel ribbon → xlwings tab
   - If you don't see xlwings tab, go to File → Options → Add-ins → Manage Excel Add-ins → Browse
   - Select xlwings.xlam and enable it

3. **Import Python module**:
   - Click on xlwings ribbon → "Import Python UDFs"
   - Navigate to `ebitda_margins_data_udf_updated.py` and select it
   - OR add this VBA code to your workbook (Alt+F11 to open VBA editor)

4. **Add VBA Module** (Alternative method):

   Press Alt+F11 to open VBA editor, then Insert → Module, and paste:

```vba
Sub ImportPythonUDFs()
    ' Import Python UDFs from the script
    Dim xw As Object
    Set xw = Application.COMAddIns("xlwings").Object

    ' Set the path to your Python script
    Dim scriptPath As String
    scriptPath = ThisWorkbook.Path & "\ebitda_margins_data_udf_updated.py"

    ' Import the module
    xw.ImportPythonModule scriptPath
End Sub
```

### Method 2: Using xlwings Quickstart

1. Navigate to your project folder in terminal
2. Run: `xlwings quickstart myproject`
3. Copy the Python UDF functions into the generated Python file
4. Open the generated Excel file

## Available Functions

Once set up, you can use these formulas in Excel:

### 1. Get Quarterly Data (Single Value)
```excel
=xw_get_quarterly_data(100186, "quarterly_ebitda_margins", "2025-06-30")
```
Returns a single value for the specified company, field, and date.

### 2. Get Time Series Data
```excel
=xw_get_series(100186, "quarterly_ebitda_margins", "2022-03-31", "2025-09-30")
```
Returns a 2-column array (date, value) for the specified date range.

### 3. Get Quarterly Matrix (All Companies)
```excel
=xw_get_quarterly_matrix("2025-06-30", "quarterly_ebitda_margins")
```
Returns data for all companies on a specific date (multi-row array).

### 4. Get All Historical Data
```excel
=xw_get_all_ebitda_margins(100186, "quarterly_ebitda_margins")
```
Returns all historical data for a specific company.

## Function Parameters

- **accord_code**: Integer company identifier (e.g., 100186)
- **field**: String column name from database (e.g., "quarterly_ebitda_margins")
- **date_text**: String date in YYYY-MM-DD format (e.g., "2025-06-30")
- **start_date**: String start date for time series
- **end_date**: String end date for time series

## Troubleshooting

### Error: "#ERROR: Database file not found"
- Check that `ebitda_margins.db` exists in the correct directory
- Verify the path in `config.ini`

### Error: "#ERROR: no such table: financial_data"
- Check your database table name
- Update the `table` setting in `config.ini`

### Error: "Import failed" or functions not recognized
- Ensure xlwings add-in is properly installed
- Try running: `xlwings addin install` again
- Restart Excel after installation

### Functions return #NAME? error
- The Python module is not loaded
- Click xlwings → "Import Python UDFs" in Excel
- Make sure Python is in your system PATH

### Performance Issues
- Functions use caching (lru_cache) for better performance
- First query might be slow, subsequent queries are faster
- Check `query_log.txt` for timing information

## Database Structure

Your database should have the following structure:

```sql
CREATE TABLE ebitda_margins (
    accord_code INTEGER,
    company_name TEXT,
    sector TEXT,
    mcap_category TEXT,
    date TEXT,
    quarterly_ebitda_margins REAL,
    -- other fields as needed
    PRIMARY KEY (accord_code, date)
);
```

## Logging

All function calls are logged to `query_log.txt` in the same directory. Check this file to:
- Monitor query performance
- Debug errors
- Track database access patterns

## Advanced Configuration

### Custom Database Location

Edit `config.ini`:
```ini
[database]
path = C:\Users\YourName\Documents\databases\ebitda_margins.db
```

### Different Date Format

If your database uses different date format (e.g., DD/MM/YYYY):
```ini
[database]
date_format = %d/%m/%Y
```

## Security Notes

- Keep your database file secure
- Do not share config.ini if it contains sensitive paths
- Log files may contain query details - review before sharing

## Support

If you encounter issues:
1. Check `query_log.txt` for detailed error messages
2. Run the Python script directly to test database connectivity
3. Verify all file paths in `config.ini`
4. Ensure xlwings version is up to date: `pip install --upgrade xlwings`

---

## Quick Start Checklist

- [ ] Install Python packages: `pip install xlwings openpyxl`
- [ ] Install xlwings add-in: `xlwings addin install`
- [ ] Place all files in same directory
- [ ] Update `config.ini` with correct database path
- [ ] Test Python script: `python ebitda_margins_data_udf_updated.py`
- [ ] Open Excel and import Python UDFs
- [ ] Test a simple formula like: `=xw_get_quarterly_data(100186, "quarterly_ebitda_margins", "2025-06-30")`

Good luck with your database integration!
