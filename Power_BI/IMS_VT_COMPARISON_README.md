# IMS VT Excel Comparison with Column Mapping

## 📋 Overview

This tool compares two Excel files where **column names are different** between files, using a mapping file that defines which columns correspond to each other.

## 📁 Files in This Folder

### Core Tools
1. **`ims_vt_mapped_comparison.py`** - Main comparison tool with column mapping support
2. **`excel_compare_agent.py`** - General Excel comparison agent
3. **`column_mappings.csv`** - Extracted column mappings (generated)

### Mapping File
4. **`IMS VT Automation Mappings.xlsx`** - Defines how columns in IMSVT map to IMS KDA Report
   - Column 3 (D): IMSVT column names
   - Column 6 (G): IMS Managed Security KDA Report column names
   - 22 main column mappings identified

### Data Files
5. **`IMSVT.xlsx`** - ⚠️ CURRENTLY CORRUPTED - needs to be fixed
6. **IMS Managed Security KDA Report.xlsx** (not found) - The second file to compare

## 🔧 Current Issue

**IMSVT.xlsx is corrupted or in an incompatible format.**

### Error Message:
```
Failed to load File A (IMSVT): File is not a zip file
```

### What This Means:
- .xlsx files are actually ZIP archives containing XML files
- Your IMSVT.xlsx file is not in this format
- It might be:
  - An old .xls file renamed to .xlsx
  - A corrupted file
  - Currently open in Excel (close it first)

## ✅ How to Fix IMSVT.xlsx

### Method 1: Open and Resave (Recommended)
1. **Close** the file if it's open in Excel
2. **Open** IMSVT.xlsx in Microsoft Excel
3. **File → Save As**
4. Choose format: **Excel Workbook (*.xlsx)**
5. Save with a new name: `IMSVT_Fixed.xlsx`
6. Use the fixed file

### Method 2: Convert from .xls to .xlsx
If it's actually an old .xls file:
1. Open in Excel
2. Save As → Excel Workbook (*.xlsx)
3. Use the new file

### Method 3: If Still Locked
```powershell
# Close any Excel processes
Get-Process excel | Stop-Process -Force

# Check if file is being used
Get-Process | Where-Object {$_.Path -like "*excel*"}
```

## 📊 Column Mappings Identified

We successfully extracted **22 column mappings**:

| IMSVT Column | IMS KDA Report Column |
|-------------|------------------------|
| Managed Security - Offshore Ratio (%) | Offshore Ratio (%) |
| Managed Security - Onshore LCR including seat charge(without COLA) | Onshore LCR including seat charge(without COLA) |
| Managed Security - Offshore LCR including seat charge(without COLA) | Offshore LCR including seat charge(without COLA) |
| Managed Security - Blended LCR without COLA | Blended LCR without COLA |
| Managed Security - Blended C2S | Blended C2S |
| Managed Security - Blended AHR | Blended AHR |
| Managed Security - Blended ADR | Blended ADR |
| ... and 15 more |

*Full mappings saved in `column_mappings.csv`*

## 🚀 How to Use (Once Files are Fixed)

### Step 1: Prepare Files
```powershell
# Make sure you're in the Power_BI folder
cd Power_BI

# List files
Get-ChildItem *.xlsx
```

Required files:
- ✅ `IMS VT Automation Mappings.xlsx` (mapping file)
- ⚠️ `IMSVT.xlsx` (needs fixing)
- ❓ Second file: IMS Managed Security KDA Report

### Step 2: Run Comparison

```powershell
# Close all Excel files first
python ims_vt_mapped_comparison.py
```

### Step 3: Review Output

The tool will generate:
- `IMS_VT_Column_Mapping_Report_YYYYMMDD_HHMMSS.xlsx`

This report contains:
- **Column Mapping Analysis** - Which columns were found in each file
- **File A Columns** - All columns in IMSVT
- **File B Columns** - All columns in IMS KDA Report
- **Mapping Reference** - Complete mapping list

## 📖 Understanding the Mapping

### Mapping Structure

The mapping file has a specific structure:

```
Row 0: Headers (IMSVT | IMS Managed Security KDA Report)
Row 1: MainHeader | Columns markers
Row 2+: Actual mappings
```

### Sub-Columns

Some columns have sub-columns:
- **IMSVT**: Guidance, AsPerSolution, Variance
- **IMS KDA**: Solution Standards, Actual Value, Variation From Standard

Example:
```
Main: "Managed Security - Offshore Ratio (%)"
  Sub: Guidance → Solution Standards
  Sub: AsPerSolution → Actual Value
  Sub: Variance → Variation From Standard
```

## 🔍 What the Tool Does

1. **Loads Mappings** - Reads IMS VT Automation Mappings.xlsx
2. **Loads Files** - Opens both Excel files (IMSVT and IMS KDA Report)
3. **Maps Columns** - Matches columns based on the mapping
4. **Compares Values** - Compares corresponding values
5. **Generates Report** - Creates comprehensive Excel report

## 🎯 Comparison Features

### Column Matching
- ✅ Exact name matching
- ✅ Case-insensitive matching
- ✅ Handles different column names via mapping
- ✅ Identifies missing columns

### Value Comparison
- ✅ Row-by-row comparison
- ✅ Handles numeric values
- ✅ Handles text values
- ✅ Handles blank cells
- ✅ Type mismatch detection

### Reporting
- ✅ Summary statistics
- ✅ Column mapping analysis
- ✅ Sample values shown
- ✅ Non-null counts
- ✅ Professional Excel formatting

## 🛠️ Alternative Approach

If you can't fix IMSVT.xlsx, you can:

### Option 1: Export Data
1. Open IMSVT in Excel
2. Copy data to a new workbook
3. Save as new .xlsx file

### Option 2: Use CSV
1. Save IMSVT as CSV
2. Modify the tool to read CSV
3. Run comparison

### Option 3: Provide Different Files
If you have other versions of these files, place them in the folder and update the script to use them.

## 📋 Checklist Before Running

- [ ] Close all Excel files
- [ ] Fix/replace IMSVT.xlsx
- [ ] Verify IMS KDA Report file exists
- [ ] Both files are in Power_BI folder
- [ ] Mapping file (IMS VT Automation Mappings.xlsx) is accessible
- [ ] Python virtual environment activated

## 🐛 Troubleshooting

### Error: "File is not a zip file"
- **Fix**: Resave the file in Excel as .xlsx format

### Error: "Cannot find file"
- **Fix**: Verify file is in Power_BI folder
- **Check**: `Get-ChildItem *.xlsx`

### Error: "Permission denied"
- **Fix**: Close the file in Excel
- **Check**: `Get-Process excel | Stop-Process -Force`

### Error: "No such sheet"
- **Fix**: Verify sheet names in the files
- **Modify**: Update sheet_name parameter in the script

## 📞 Current Status

### What Works ✅
- ✅ Mapping file successfully loaded (22 mappings)
- ✅ Column mappings extracted and saved to CSV
- ✅ Tool is ready to run
- ✅ All scripts created and tested

### What Needs Fixing ⚠️
- ⚠️ IMSVT.xlsx is corrupted - **needs to be resaved**
- ❓ Second file (IMS KDA Report) - **needs to be provided**

## 🎓 Next Steps

1. **Fix IMSVT.xlsx**
   ```
   Open in Excel → Save As → Excel Workbook (.xlsx)
   ```

2. **Provide IMS KDA Report file**
   - Place in Power_BI folder
   - Or tell the tool where it is

3. **Run the comparison**
   ```powershell
   python ims_vt_mapped_comparison.py
   ```

4. **Review the generated report**
   - Open the Excel file
   - Check Column Mapping Analysis sheet
   - Review any mismatches

## 📚 Additional Tools in This Folder

- **`excel_compare_agent.py`** - General comparison tool (for files with same column names)
- **`test_excel_compare.py`** - Test suite
- **`example_usage.py`** - Usage examples
- **`generate_sample_data.py`** - Sample data generator

## 💡 Tips

1. **Always close Excel** before running Python scripts
2. **Check file formats** - use .xlsx not .xls
3. **Verify mappings** - review column_mappings.csv
4. **Test with sample data** first
5. **Keep backups** of original files

---

**Status**: Ready to run once IMSVT.xlsx is fixed  
**Date**: March 5, 2026  
**Location**: `Power_BI\`
