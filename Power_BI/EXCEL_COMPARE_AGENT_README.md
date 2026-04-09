# Excel Compare Agent

A comprehensive Python tool to compare values between two or more Excel workbooks based on user-provided requirements.

## Features

- **Natural Language Requirement Parsing**: Describe what you want to compare in plain English
- **Intelligent Column Matching**: Match rows using single or composite keys
- **Flexible Comparison Rules**:
  - Text: exact, case-insensitive, trimmed
  - Numeric: absolute tolerance, percentage tolerance
  - Date: format and timezone handling
  - Blank value handling (treat as 0, NA, or distinct)
- **Comprehensive Output**:
  - Detailed Excel report with 5 sheets
  - Summary statistics
  - Row-level comparison results
  - Unmatched records
  - Configuration tracking
- **Professional Formatting**: Color-coded differences, auto-sized columns, frozen headers

## Installation

Ensure you have the required dependencies:

```bash
pip install pandas openpyxl
```

All required packages should already be in your `requirements.txt`.

## Quick Start

### Method 1: Natural Language (Recommended)

```python
from excel_compare_agent import ExcelCompareAgent

requirement = """
Compare Mar 28th_Release.xlsx vs 21st Feb_Release.xlsx. 
Use sheet "Release" in both. 
Match rows using US_ID. 
Compare fields: PT Status, In sprint test case count, If insprint YES - % of completion. 
Treat text as case-insensitive and trimmed. 
Treat numeric values with tolerance 0.0 (exact). 
Output full report.
"""

agent = ExcelCompareAgent(requirement_text=requirement)
output_file = agent.run()
print(f"Report saved to: {output_file}")
```

### Method 2: Programmatic Configuration

```python
from excel_compare_agent import ExcelCompareAgent, ComparisonConfig, ComparisonRule

rules = ComparisonRule(
    text_mode="case_insensitive_trimmed",
    numeric_tolerance=0.0
)

config = ComparisonConfig(
    fileA="baseline.xlsx",
    fileB="updated.xlsx",
    sheetA="Sheet1",
    sheetB="Sheet1",
    keyColumns=["US_ID"],
    compareColumns=["PT Status", "Test Count", "Completion %"],
    rules=rules
)

agent = ExcelCompareAgent(config=config)
output_file = agent.run()
```

### Method 3: Command Line Interface

```bash
# Interactive mode (paste requirement text)
python excel_compare_agent.py --interactive

# Direct requirement
python excel_compare_agent.py --requirement "Compare file1.xlsx vs file2.xlsx. Match using ID..."

# Using saved config
python excel_compare_agent.py --config comparison_config.json --output my_report.xlsx
```

## Understanding the Requirement Format

The agent can parse natural language requirements. Include these elements:

### Required Elements:

1. **File Names**: `Compare <file1.xlsx> vs <file2.xlsx>`
2. **Key Columns**: `Match rows using <column_name>` or `Match using <col1>, <col2>`
3. **Compare Fields**: `Compare fields: <field1>, <field2>, <field3>`

### Optional Elements:

- **Sheet Names**: `Use sheet "SheetName"` (defaults to first sheet)
- **Text Comparison**: `case-insensitive` or `exact match` (default: case-insensitive trimmed)
- **Numeric Tolerance**: `tolerance 0.5` or `tolerance 0.0` for exact (default: 0.0)
- **Blank Handling**: `treat blank as zero` or `treat blank as NA`

### Example Requirements:

```text
Example 1 - Release Tracking:
Compare Mar 28th_Release.xlsx vs 21st Feb_Release.xlsx. 
Use sheet "Release" in both. 
Match rows using US_ID. 
Compare fields: PT Status, In sprint test case count, If insprint YES - % of completion. 
Treat text as case-insensitive and trimmed. 
Treat numeric values with tolerance 0.0 (exact).

Example 2 - Financial Data:
Compare Q1_Report.xlsx with Q2_Report.xlsx. 
Sheet: "Financial Summary". 
Match using Account_ID and Department. 
Compare: Revenue, Expenses, Profit_Margin. 
Numeric tolerance: 0.01 (allow rounding differences).

Example 3 - Inventory:
Compare current_inventory.xlsx vs last_week_inventory.xlsx. 
Match rows using SKU. 
Compare fields: Quantity, Unit_Price, Status. 
Treat text as exact match.
```

## Output Report Structure

The generated Excel workbook contains 5 sheets:

### 1. Summary Sheet
- Overall statistics (total rows, matched, missing, duplicates)
- Field-wise comparison metrics
- Overall match rate percentage

### 2. Compared_Data Sheet
Contains all comparison results with columns:
- `Key`: Composite key value
- `Match_Status`: Matched | Missing_in_A | Missing_in_B | Duplicate_Key_A | Duplicate_Key_B
- For each compared field:
  - `<Field>_A`: Value from File A
  - `<Field>_B`: Value from File B
  - `<Field>_Result`: Match | Mismatch | Both_Blank | Type_Mismatch | Tolerance_Match
  - `<Field>_Notes`: Additional details (e.g., numeric difference)
- `Row_Result`: All_Match | Has_Mismatch | Unmatched
- `Mismatch_Count`: Number of mismatched fields

### 3. Unmatched_In_A Sheet
Records present in File B but not in File A

### 4. Unmatched_In_B Sheet
Records present in File A but not in File B

### 5. Config Sheet
Configuration used for this comparison:
- File paths
- Sheet names
- Key and compare columns
- Comparison rules
- Timestamp

## Comparison Rules

### Text Comparison Modes:
- `exact`: Character-by-character match (case-sensitive)
- `case_insensitive`: Ignore case differences
- `case_insensitive_trimmed`: Ignore case and whitespace (default)

### Numeric Comparison:
- `numeric_tolerance`: Absolute difference allowed (e.g., 0.01)
- `numeric_tolerance_percent`: Percentage difference allowed (e.g., 5.0 for 5%)

### Date Comparison:
- `date_format`: Expected date format
- `date_only`: Compare only date part, ignore time

### Blank Handling:
- `treat_blank_as_zero`: Empty cells treated as 0 for numeric comparisons
- `treat_blank_as_na`: Empty cells treated as "NA" string

## Advanced Usage

### Auto-Detection of Columns

If key columns aren't specified, the agent will:
1. Look for columns containing "ID", "Key", "US_ID", etc.
2. Use the first column if no ID-like column found

If compare columns aren't specified, the agent will:
1. Compare all columns except the key columns

```python
# Minimal requirement - agent will auto-detect
requirement = "Compare file1.xlsx with file2.xlsx using sheet Data"
agent = ExcelCompareAgent(requirement_text=requirement)
output = agent.run()
```

### Saving and Reusing Configuration

```python
from excel_compare_agent import RequirementParser

# Parse and save
config = RequirementParser.parse(requirement_text)
RequirementParser.save_config(config, "my_comparison.json")

# Load and reuse
python excel_compare_agent.py --config my_comparison.json
```

### Composite Keys (Multiple Key Columns)

```python
config = ComparisonConfig(
    fileA="data.xlsx",
    fileB="data_updated.xlsx",
    keyColumns=["Department", "Employee_ID"],  # Composite key
    compareColumns=["Salary", "Position"]
)
```

## Testing

Run the test suite to verify installation:

```bash
python test_excel_compare.py
```

This will:
1. Check for required Excel files
2. Generate sample data if needed
3. Run comparison tests
4. Generate sample reports

## Error Handling

The agent provides clear error messages for common issues:

- **Missing Files**: Reports which file is not found
- **Missing Columns**: Lists available columns and what's missing
- **Duplicate Keys**: Warns about duplicate key values
- **Type Mismatches**: Identifies fields with incompatible data types

## Performance

The agent is optimized for:
- Files with up to 100,000 rows
- Up to 50 comparison columns
- Handles merged cells and multi-row headers
- Processes empty rows and columns gracefully

## Limitations

- **Excel Format Only**: Currently supports .xlsx files only (not .xls or .csv)
- **In-Memory Processing**: Very large files (>500MB) may require additional memory
- **Formula Comparison**: Compares values, not formulas (by design)

## Troubleshooting

### Issue: "Column not found"
**Solution**: Check column names for exact spelling. Use auto-detection or explicitly list available columns.

### Issue: "Many duplicates detected"
**Solution**: Verify you're using the correct key columns. Consider using a composite key.

### Issue: "Memory error with large files"
**Solution**: 
- Filter data in Excel before comparison
- Use only required columns
- Process in batches if needed

### Issue: "All values showing as mismatch"
**Solution**: Check comparison rules - you may need case-insensitive or trimmed mode.

## Examples in This Workspace

### Compare Release Trackers

```python
requirement = """
Compare Mar 28th_Release.xlsx vs 21st Feb_Release.xlsx. 
Use sheet "Release". 
Match using US_ID. 
Compare: PT Status, In sprint test case count, If insprint YES - % of completion.
Case-insensitive text, exact numeric.
"""

agent = ExcelCompareAgent(requirement_text=requirement)
agent.run()
```

### Compare Bug Reports

```python
requirement = """
Compare Bug_summary.xlsx from two different dates.
Match using Bug_ID.
Compare: Status, Priority, Assignee, Resolution_Date.
Treat text as case-insensitive.
"""
```

## API Reference

### Classes

#### `ExcelCompareAgent(requirement_text=None, config=None)`
Main agent class.
- **requirement_text**: Natural language requirement string
- **config**: ComparisonConfig object

Methods:
- `run()`: Execute comparison and return output file path
- `validate_config()`: Check configuration validity
- `auto_detect_columns()`: Auto-detect key and compare columns

#### `ComparisonConfig`
Configuration dataclass with attributes:
- `fileA`, `fileB`: File paths
- `sheetA`, `sheetB`: Sheet names
- `keyColumns`: List of key column names
- `compareColumns`: List of columns to compare
- `rules`: ComparisonRule object
- `outputPath`: Output file path
- `includeAllRows`: Include matched rows (True) or only mismatches (False)
- `highlightDifferences`: Apply color formatting (True/False)

#### `ComparisonRule`
Rules dataclass with attributes:
- `text_mode`: "exact" | "case_insensitive" | "case_insensitive_trimmed"
- `numeric_tolerance`: Float (absolute tolerance)
- `numeric_tolerance_percent`: Float (percentage tolerance)
- `treat_blank_as_zero`: Boolean
- `treat_blank_as_na`: Boolean

#### `RequirementParser`
Static methods:
- `parse(requirement_text)`: Parse text into ComparisonConfig
- `save_config(config, output_path)`: Save config to JSON

## Support

For issues or questions:
1. Check the examples in `test_excel_compare.py`
2. Review error messages and logs
3. Verify Excel file structure and column names
4. Check the Config sheet in output reports to see what was parsed

## Version History

- **v1.0.0** (March 2026) - Initial release
  - Natural language requirement parsing
  - Comprehensive comparison engine
  - Professional Excel reports
  - CLI interface
  - Auto-detection features

## License

Internal tool for project use.

---

**Author**: Excel Compare Agent System  
**Date**: March 5, 2026  
**Python Version**: 3.8+  
**Dependencies**: pandas, openpyxl
