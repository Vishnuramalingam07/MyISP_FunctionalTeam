import os
import re
import math
import time
import json
import requests
import pandas as pd
from collections import defaultdict

ORG = "accenturecio08"
PROJECT = "AutomationProcess_29697"
QUERY_ID = "0ed0091f-b665-4f51-b553-4c5afdea5e92"

OUTPUT_PATH = r"C:\Users\vishnu.ramalingam\MyISP_Tools\Missing_Filed_Report\Missing_Fields_Report.xlsx"
PT_LEAD_MAPPING_PATH = r"C:\Users\vishnu.ramalingam\MyISP_Tools\Missing_Filed_Report\PT Lead.xlsx"
    
FIELDS_DISPLAY = [
    "Defect Record",
    "mySP Initiative",
    "Sub Initiative",
    "TextVerification",
    "TextVerification1",
    "Category",
    "mySP Bug Link the User Story",
    "Broken",
    "Identified in mySP Release Name",
    "Fixed in mySP Release Name",
    "mySP RCA",
]



FIELD_ALIASES = {
    "mySP Bug Link the User Story": [
        "mySP Bug Link the User Story",
        "mySP Bug Link - User Story",
        "mySP Bug Link – User Story",
        "mySP Bug Link — User Story",
    ],
}

NA_STRINGS = {"na", "n/a", "null", "none", "-", ""}


def normalize_field_name(name: str) -> str:
    if name is None:
        return ""
    name = name.strip().lower()
    name = name.replace("–", "-")
    name = re.sub(r"\s+", " ", name)
    return name


def extract_enterprise_id(created_by) -> str:
    if not created_by:
        return ""
    if isinstance(created_by, dict):
        unique_name = created_by.get("uniqueName") or ""
        display_name = created_by.get("displayName") or ""
    else:
        unique_name = ""
        display_name = str(created_by)

    if "@" in unique_name:
        return unique_name.split("@", 1)[0].lower()
    if "\\" in unique_name:
        return unique_name.split("\\", 1)[-1].lower()

    if "@" in display_name:
        return display_name.split("@", 1)[0].lower()
    if "\\" in display_name:
        return display_name.split("\\", 1)[-1].lower()

    return display_name.strip().lower()


def is_missing(value) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return normalize_field_name(value) in NA_STRINGS
    return False


def get_pat() -> str:
    pat = os.getenv("AZURE_DEVOPS_PAT")
    if not pat:
        raise RuntimeError("Missing AZURE_DEVOPS_PAT environment variable. Set your Azure DevOps PAT and rerun.")
    return pat


def azdo_get(url: str, pat: str, retries: int = 3, backoff: float = 5.0):
    for attempt in range(1, retries + 1):
        response = requests.get(url, auth=("", pat))
        if response.status_code in (429, 500, 502, 503, 504) and attempt < retries:
            wait = backoff * attempt
            print(f"  ⚠ {response.status_code} on attempt {attempt}/{retries}, retrying in {wait}s …")
            time.sleep(wait)
            continue
        response.raise_for_status()
        
        # Check if the response is JSON before parsing
        content_type = response.headers.get('Content-Type', '')
        if 'application/json' not in content_type:
            # Check for authentication issues
            if 'text/html' in content_type and 'Sign In' in response.text:
                print("\n" + "="*70)
                print("ERROR: AUTHENTICATION FAILED")
                print("="*70)
                print("The Azure DevOps API returned a sign-in page instead of data.")
                print("\nPossible causes:")
                print("  1. AZURE_DEVOPS_PAT environment variable is not set")
                print("  2. Your PAT token has expired")
                print("  3. Your PAT token is invalid")
                print("  4. Your PAT token doesn't have required permissions")
                print("\nHow to fix:")
                print("  1. Go to https://dev.azure.com/accenturecio08/_usersSettings/tokens")
                print("  2. Create a new PAT with 'Work Items (Read)' permission")
                print("  3. Set the environment variable:")
                print('     $env:AZURE_DEVOPS_PAT = "your-pat-token-here"')
                print("  4. Re-run this script")
                print("="*70 + "\n")
                raise ValueError("Authentication failed - Please set a valid AZURE_DEVOPS_PAT")
            
            print(f"ERROR: Expected JSON but received Content-Type: {content_type}")
            print(f"URL: {url}")
            print(f"Response status: {response.status_code}")
            print(f"Response preview: {response.text[:500]}")
            raise ValueError(f"API returned non-JSON response (Content-Type: {content_type})")
        
        try:
            return response.json()
        except requests.exceptions.JSONDecodeError as e:
            print(f"ERROR: Failed to parse JSON response")
            print(f"URL: {url}")
            print(f"Response status: {response.status_code}")
            print(f"Response content: {response.text[:500]}")
            raise


def chunked(iterable, size):
    for i in range(0, len(iterable), size):
        yield iterable[i : i + size]


def add_vba_highlighting(xlsx_path: str) -> str:
    """Open the .xlsx with Excel, inject a VBA macro that highlights detail rows
    in light green when a summary hyperlink is clicked, and save as .xlsm."""
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print("pywin32 not installed. Install with: pip install pywin32")
        print("Report saved without macro support.")
        return xlsx_path

    xlsm_path = xlsx_path.rsplit(".", 1)[0] + ".xlsm"

    vba_code = (
        "Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)\n"
        "    Dim subAddr As String\n"
        "    subAddr = Target.SubAddress\n"
        "    If Len(subAddr) = 0 Then Exit Sub\n"
        "\n"
        "    Dim excl As Long\n"
        '    excl = InStr(subAddr, "!")\n'
        "    If excl = 0 Then Exit Sub\n"
        "\n"
        "    Dim sheetName As String\n"
        "    sheetName = Left(subAddr, excl - 1)\n"
        '    sheetName = Replace(sheetName, "' + "'" + '", "")\n'
        "\n"
        "    Dim rangeAddr As String\n"
        "    rangeAddr = Mid(subAddr, excl + 1)\n"
        "\n"
        "    Dim ws As Worksheet\n"
        "    On Error Resume Next\n"
        "    Set ws = ThisWorkbook.Sheets(sheetName)\n"
        "    On Error GoTo 0\n"
        "    If ws Is Nothing Then Exit Sub\n"
        "\n"
        "    Dim lastRow As Long\n"
        "    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row\n"
        "    If lastRow >= 4 Then\n"
        '        ws.Range("A4:H" & lastRow).Interior.ColorIndex = xlNone\n'
        "    End If\n"
        "\n"
        "    On Error Resume Next\n"
        "    ws.Range(rangeAddr).Interior.Color = RGB(198, 239, 206)\n"
        "    On Error GoTo 0\n"
        "End Sub\n"
    )

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))

        # Add the FollowHyperlink handler to every "… Summary" sheet
        for ws in wb.Worksheets:
            if ws.Name.endswith(" Summary"):
                # Find the matching VBComponent for this worksheet
                target_comp = None
                for comp in wb.VBProject.VBComponents:
                    # Type 100 = document module (sheet / workbook)
                    if comp.Type == 100:
                        try:
                            if comp.Properties("Name").Value == ws.CodeName:
                                target_comp = comp
                                break
                        except Exception:
                            pass
                # Fallback: try matching by iterating all doc components
                if target_comp is None:
                    for comp in wb.VBProject.VBComponents:
                        if comp.Type == 100 and comp.Name == ws.CodeName:
                            target_comp = comp
                            break
                if target_comp is None:
                    print(f"  Warning: Could not find VBComponent for sheet '{ws.Name}' (CodeName='{ws.CodeName}')")
                    # Last resort: try direct index
                    try:
                        target_comp = wb.VBProject.VBComponents(ws.CodeName)
                    except Exception:
                        continue
                target_comp.CodeModule.AddFromString(vba_code)

        wb.SaveAs(os.path.abspath(xlsm_path), FileFormat=52)  # xlsm
        wb.Close(False)

        os.remove(xlsx_path)
        print(f"VBA macro added – saved as .xlsm")
        return xlsm_path

    except Exception as e:
        print(f"Warning: Could not add VBA macro: {e}")
        print("Tip: In Excel go to File > Options > Trust Center > Trust Center Settings")
        print('     > Macro Settings > tick "Trust access to the VBA project object model".')
        return xlsx_path
    finally:
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def generate_html_report(missing_df: pd.DataFrame, fields_display: list, output_path: str):
    """Generate an interactive HTML report with summary page and separate detail pages."""
    base_dir = os.path.dirname(output_path)
    base_name = os.path.splitext(os.path.basename(output_path))[0]
    
    # Create a subfolder for detail pages
    details_dir = os.path.join(base_dir, f"{base_name}_details")
    os.makedirs(details_dir, exist_ok=True)
    
    summary = missing_df.groupby(["Field", "PT Lead"]).size().reset_index(name="Missing Count")
    pt_leads = sorted(missing_df["PT Lead"].unique())
    
    # Get current date and time
    report_datetime = pd.Timestamp.now().strftime("%B %d, %Y at %I:%M %p")
    
    # Generate individual detail pages for each group
    detail_files = {}
    for field in fields_display:
        field_data = missing_df[missing_df["Field"] == field]
        if field_data.empty:
            continue
        
        for pt in pt_leads:
            pt_data = field_data[field_data["PT Lead"] == pt]
            if pt_data.empty:
                continue 
            
            group_key = (field, pt)
            safe_field = field.replace(' ', '_').replace('/', '_')
            safe_pt = pt.replace(' ', '_').replace('/', '_')
            detail_filename = f"detail_{safe_field}_{safe_pt}.html"
            detail_path = os.path.join(details_dir, detail_filename)
            detail_files[group_key] = detail_filename
            
            # Generate detail page HTML
            detail_html = []
            detail_html.append(f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{field} - {pt}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 30px;
            min-height: 100vh;
        }}
        .container {{ 
            max-width: 1500px; 
            margin: 0 auto; 
            background: white; 
            padding: 40px; 
            border-radius: 12px; 
            box-shadow: 0 10px 40px rgba(0,0,0,0.3); 
        }}
        .header {{
            border-bottom: 4px solid #0078d4;
            padding-bottom: 20px;
            margin-bottom: 30px;
        }}
        h1 {{ 
            color: #0078d4; 
            margin-bottom: 8px; 
            font-size: 32px;
            font-weight: 700;
        }}
        .subtitle {{ 
            color: #555; 
            font-size: 18px;
            margin-top: 5px;
            display: flex;
            align-items: center;
            gap: 20px;
        }}
        .badge {{
            background: #0078d4;
            color: white;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 14px;
            font-weight: 600;
        }}
        .back-link {{ 
            display: inline-block; 
            margin-bottom: 25px; 
            color: white;
            background: #0078d4;
            padding: 10px 20px;
            border-radius: 6px;
            text-decoration: none; 
            font-weight: 600;
            transition: all 0.3s ease;
        }}
        .back-link:hover {{ 
            background: #005a9e;
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0,120,212,0.3);
        }}
        table {{ 
            width: 100%; 
            border-collapse: collapse; 
            margin-top: 20px; 
            background: white;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            border-radius: 8px;
            overflow: hidden;
        }}
        th {{ 
            background: linear-gradient(135deg, #0078d4 0%, #005a9e 100%);
            color: white; 
            padding: 16px 12px; 
            text-align: left; 
            font-weight: 600;
            font-size: 14px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        td {{ 
            padding: 14px 12px; 
            border-bottom: 1px solid #f0f0f0;
            vertical-align: top;
        }}
        tr:hover {{ 
            background: #C6EFCE;
            transition: background 0.2s ease;
        }}
        tr:last-child td {{
            border-bottom: none;
        }}
        .url-link {{ 
            color: white;
            background: #0563C1;
            padding: 6px 14px;
            border-radius: 4px;
            text-decoration: none; 
            font-weight: 600;
            font-size: 13px;
            display: inline-block;
            transition: all 0.3s ease;
        }}
        .url-link:hover {{ 
            background: #004080;
            transform: scale(1.05);
        }}
        .record-count {{ 
            color: #0078d4; 
            font-weight: 600; 
            margin-bottom: 20px;
            font-size: 16px;
            background: #e3f2fd;
            padding: 10px 16px;
            border-radius: 6px;
            display: inline-block;
        }}
        .work-item-id {{
            font-weight: 600;
            color: #0078d4;
        }}
        .timestamp {{
            color: #999;
            font-size: 13px;
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #e0e0e0;
            text-align: center;
        }}
    </style>
</head>
<body>
    <div class="container">
        <a href="../{base_name}.html" class="back-link">← Back to Summary</a>
        <div class="header">
            <h1>{field}</h1>
            <div class="subtitle">
                <span>PT Lead: <strong>{pt}</strong></span>
                <span class="badge">{len(pt_data)} Records</span>
            </div>
        </div>
        <table>
            <tr>
                <th>Work Item ID</th>
                <th>Title</th>
                <th>State</th>
                <th>Created By</th>
                <th style="text-align: center;">Azure DevOps</th>
                <th>Field Value</th>
            </tr>
""")
            
            for _, row in pt_data.iterrows():
                detail_html.append(f"""            <tr>
                <td class="work-item-id">{row["Work Item ID"]}</td>
                <td>{row["Title"]}</td>
                <td>{row["State"]}</td>
                <td>{row["Created By"]}</td>
                <td style="text-align: center;"><a href="{row["URL"]}" target="_blank" class="url-link">Open</a></td>
                <td>{row["Field Value"] if pd.notna(row["Field Value"]) else "—"}</td>
            </tr>
""")
            
            detail_html.append("""        </table>
        <div class="timestamp">Generated on {0}</div>
        <br>
        <a href="../{1}.html" class="back-link">← Back to Summary</a>
    </div>
</body>
</html>
""".format(report_datetime, base_name))
            
            with open(detail_path, 'w', encoding='utf-8') as f:
                f.write(''.join(detail_html))
    
    # Generate summary page HTML with tab-based design
    html_path = os.path.join(base_dir, f"{base_name}.html")
    html_parts = []
    
    # Prepare data for JavaScript - organize records by field
    records_by_field = {}
    records_by_field['All Fields'] = missing_df.to_dict('records')
    for field in fields_display:
        field_data = missing_df[missing_df["Field"] == field]
        if not field_data.empty:
            records_by_field[field] = field_data.to_dict('records')
    
    html_parts.append("""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Missing Fields - Dashboard Report</title>
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700;800;900&family=Inter:wght@400;500;600;700;800&family=Playfair+Display:wght@700;900&display=swap" rel="stylesheet">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Inter', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%);
            background-attachment: fixed;
            padding: 20px;
            min-height: 100vh;
        }
        .container { 
            max-width: 1400px; 
            margin: 0 auto; 
            background: white; 
            padding: 40px; 
            border-radius: 20px; 
            box-shadow: 0 20px 60px rgba(0,0,0,0.3); 
            border: 3px solid rgba(255,255,255,0.5);
        }
        .header {
            text-align: center;
            margin-bottom: 25px;
            padding-bottom: 15px;
            position: relative;
            min-height: 70px;
            border-bottom: 3px solid #f0f0f0;
        }
        h1 { 
            font-family: 'Poppins', sans-serif;
            color: #1a1a1a; 
            margin-bottom: 15px; 
            font-size: 32px; 
            font-weight: 900;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            letter-spacing: -0.5px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
            text-transform: uppercase;
        }
        .report-info {
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: transparent;
            padding: 0 20px;
            border-radius: 0;
            box-shadow: none;
            margin-top: 5px;
        }
        .info-badge {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 10px 18px;
            font-size: 13px;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 8px;
            border-radius: 25px;
            box-shadow: 0 4px 15px rgba(102,126,234,0.35);
        }
        
        /* Tab Styles */
        .tabs-container {
            margin: 30px 0 20px 0;
            border-bottom: 4px solid transparent;
            background: linear-gradient(to right, #667eea, #764ba2, #f093fb);
            border-radius: 12px 12px 0 0;
            overflow: visible;
            padding: 8px 8px 0 8px;
        }
        .tabs {
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
            padding-bottom: 0;
            justify-content: center;
        }
        .tab-btn {
            font-family: 'Poppins', sans-serif;
            background: white;
            border: none;
            padding: 12px 20px;
            font-size: 13px;
            font-weight: 800;
            color: #555;
            cursor: pointer;
            transition: all 0.3s ease;
            border-radius: 10px 10px 0 0;
            white-space: nowrap;
            position: relative;
            bottom: 0;
            flex-shrink: 0;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            box-shadow: 0 -2px 8px rgba(0,0,0,0.1);
        }
        .tab-btn:hover {
            background: linear-gradient(135deg, #ffd89b 0%, #ffb88c 100%);
            color: #fff;
            transform: translateY(-3px);
            box-shadow: 0 4px 12px rgba(255,184,140,0.4);
        }
        .tab-btn.active {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            transform: translateY(-5px);
            box-shadow: 0 6px 20px rgba(102,126,234,0.5);
        }
        .tab-btn.active:hover {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            transform: translateY(-5px);
            box-shadow: 0 6px 20px rgba(102,126,234,0.5);
        }
        .tab-btn.summary-tab {
            font-weight: 900;
            background: linear-gradient(135deg, #ffd89b 0%, #ff8c42 100%);
            color: white;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.2);
        }
        .tab-btn.summary-tab:hover {
            background: linear-gradient(135deg, #ffb88c 0%, #ff6b35 100%);
            color: white;
            transform: translateY(-3px);
            box-shadow: 0 4px 12px rgba(255,107,53,0.4);
        }
        .tab-btn.summary-tab.active {
            background: linear-gradient(135deg, #ff8c42 0%, #ff5722 100%);
            color: white;
            transform: translateY(-5px);
            box-shadow: 0 6px 20px rgba(255,87,34,0.6);
        }
        
        /* Tab Content */
        .tab-content {
            display: none;
        }
        .tab-content.active {
            display: block;
            animation: fadeIn 0.4s ease-in;
        }
        
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        /* Summary Table */
        .summary-table-container {
            overflow: visible;
            margin-bottom: 25px;
            border-radius: 12px;
            box-shadow: 0 8px 24px rgba(0,0,0,0.15);
            border: 2px solid #e0e0e0;
        }
        .summary-table {
            width: 100%;
            table-layout: fixed;
            border-collapse: collapse;
            font-size: 11px;
            font-family: 'Inter', sans-serif;
        }
        .summary-table thead {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%);
        }
        .summary-table th {
            background: transparent;
            color: white !important;
            padding: 12px 8px;
            text-align: center;
            font-weight: 900;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            border: 1px solid rgba(255,255,255,0.3);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.2);
        }
        .summary-table thead tr:hover th {
            background: transparent !important;
            color: white !important;
        }
        .summary-table th:first-child {
            text-align: left;
            width: 150px;
            background: linear-gradient(135deg, #4a5568 0%, #2d3748 100%);
            font-size: 12px;
            font-weight: 900;
        }
        .summary-table thead tr:hover th:first-child {
            background: linear-gradient(135deg, #4a5568 0%, #2d3748 100%) !important;
            color: white !important;
        }
        .summary-table td {
            padding: 8px 6px;
            border: 1px solid #e0e0e0;
            text-align: center;
            font-size: 11px;
            font-weight: 600;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .summary-table td:first-child {
            text-align: left;
            font-weight: 800;
            color: #2d3748;
            background: linear-gradient(135deg, #f7fafc 0%, #edf2f7 100%);
            border-right: 3px solid #cbd5e0;
            width: 150px;
        }
        .summary-table tbody tr:hover {
            background: linear-gradient(135deg, #e6f7ff 0%, #f0f8ff 100%);
            transform: scale(1.01);
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        .summary-table tbody tr:hover td:first-child {
            background: linear-gradient(135deg, #bee3f8 0%, #90cdf4 100%);
            color: #1a365d;
        }
        .summary-table .count-cell {
            font-weight: 800;
            color: #0078d4;
            background: linear-gradient(135deg, #e6f7ff 0%, #f0f8ff 100%);
            cursor: default;
        }
        .summary-table .count-cell.zero {
            color: #a0aec0;
            font-weight: 500;
            background: #f7fafc;
        }
        .summary-table .grand-total-row {
            background: linear-gradient(135deg, #ffd89b 0%, #ffb88c 100%);
            font-weight: 900;
            border-top: 4px solid #ff8c42;
            border-bottom: 4px solid #ff8c42;
        }
        .summary-table .grand-total-row td {
            font-weight: 900;
            color: #7c2d12;
            border-top: 4px solid #ff8c42;
            padding: 10px 6px;
            text-shadow: 1px 1px 2px rgba(255,255,255,0.5);
        }
        .summary-table .grand-total-row td:first-child {
            background: linear-gradient(135deg, #ffd89b 0%, #ffb88c 100%);
            text-transform: uppercase;
            letter-spacing: 1px;
            font-size: 13px;
        }
        .summary-table .total-count {
            color: #dc2626;
            font-size: 13px;
            font-weight: 900;
            text-shadow: 1px 1px 2px rgba(255,255,255,0.5);
        }
        
        /* Filter Section */
        .filter-section {
            background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
            padding: 16px 20px;
            border-radius: 12px;
            margin-bottom: 20px;
            border-left: 6px solid #0284c7;
            border-right: 6px solid #7c3aed;
            display: flex;
            align-items: center;
            gap: 15px;
            flex-wrap: nowrap;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }
        .filter-group {
            flex: 0 0 auto;
            min-width: 200px;
            max-width: 300px;
        }
        .filter-label {
            font-size: 13px;
            font-weight: 800;
            color: #1e293b;
            margin-bottom: 8px;
            display: block;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .multiselect-wrapper {
            position: relative;
        }
        .multiselect-display {
            width: 100%;
            padding: 10px 14px;
            font-size: 14px;
            font-weight: 600;
            border: 3px solid #cbd5e0;
            border-radius: 10px;
            background: white;
            color: #2d3748;
            cursor: pointer;
            transition: all 0.3s ease;
            display: flex;
            justify-content: space-between;
            align-items: center;
            user-select: none;
        }
        .multiselect-display:hover {
            border-color: #667eea;
            background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(102,126,234,0.2);
        }
        .multiselect-display.active {
            border-color: #667eea;
            background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%);
            box-shadow: 0 0 0 4px rgba(102,126,234,0.15);
        }
        .multiselect-dropdown {
            position: absolute;
            top: 100%;
            left: 0;
            right: 0;
            background: white;
            border: 3px solid #667eea;
            border-radius: 10px;
            margin-top: 6px;
            max-height: 250px;
            overflow-y: auto;
            display: none;
            z-index: 1000;
            box-shadow: 0 8px 24px rgba(0,0,0,0.2);
        }
        .multiselect-dropdown.show {
            display: block;
            animation: dropdownSlide 0.3s ease-out;
        }
        
        @keyframes dropdownSlide {
            from { opacity: 0; transform: translateY(-10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        .multiselect-option {
            padding: 10px 14px;
            cursor: pointer;
            transition: all 0.2s ease;
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 14px;
            font-weight: 600;
        }
        .multiselect-option:hover {
            background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%);
        }
        .multiselect-option input[type="checkbox"] {
            cursor: pointer;
            width: 16px;
            height: 16px;
        }
        .multiselect-option.select-all {
            border-bottom: 3px solid #cbd5e0;
            background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
            font-weight: 800;
        }
        .clear-filter-btn {
            background: linear-gradient(135deg, #fee2e2 0%, #fecaca 100%);
            color: #dc2626;
            border: 2px solid #dc2626;
            padding: 10px 16px;
            font-size: 13px;
            font-weight: 800;
            cursor: pointer;
            transition: all 0.3s ease;
            white-space: nowrap;
            flex-shrink: 0;
            height: fit-content;
            align-self: flex-end;
            margin-bottom: 0;
            margin-left: 10px;
            border-radius: 10px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .clear-filter-btn:hover {
            background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%);
            color: white;
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(220,38,38,0.3);
        }
        
        /* Table Styles */
        .table-container {
            overflow-x: auto;
            margin-bottom: 25px;
            border-radius: 12px;
            box-shadow: 0 8px 24px rgba(0,0,0,0.15);
        }
        table { 
            width: 100%; 
            border-collapse: collapse; 
            box-shadow: 0 4px 16px rgba(0,0,0,0.1);
            border-radius: 12px;
            overflow: hidden;
            font-family: 'Inter', sans-serif;
        }
        th { 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%);
            color: white; 
            padding: 16px 18px; 
            text-align: left; 
            font-weight: 900;
            font-size: 14px;
            text-transform: uppercase;
            letter-spacing: 1px;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.2);
            border-right: 1px solid rgba(255,255,255,0.2);
        }
        td { 
            padding: 14px 18px; 
            border-bottom: 2px solid #e0e0e0;
            font-size: 14px;
            font-weight: 500;
            border-right: 1px solid #f0f0f0;
        }
        tr:hover { 
            background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
            transition: all 0.3s ease;
            transform: scale(1.005);
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        tbody tr {
            transition: all 0.2s ease;
        }
        tbody tr:nth-child(even) {
            background: #f8fafc;
        }
        tbody tr:nth-child(odd) {
            background: white;
        }
        .work-item-id {
            font-weight: 800;
            color: #0ea5e9;
            font-size: 15px;
            text-shadow: 1px 1px 2px rgba(14,165,233,0.1);
        }
        .url-link {
            color: white;
            background: linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%);
            padding: 8px 16px;
            border-radius: 8px;
            text-decoration: none;
            font-weight: 700;
            font-size: 13px;
            display: inline-block;
            transition: all 0.3s ease;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            box-shadow: 0 4px 12px rgba(14,165,233,0.3);
        }
        .url-link:hover {
            background: linear-gradient(135deg, #7c3aed 0%, #6d28d9 100%);
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(124,58,237,0.4);
        }
        
        /* Pagination */
        .pagination-container {
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 12px;
            margin-top: 30px;
            flex-wrap: wrap;
            padding: 20px;
            background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
            border-radius: 12px;
        }
        .pagination-btn {
            font-family: 'Poppins', sans-serif;
            background: white;
            border: 3px solid #667eea;
            color: #667eea;
            padding: 10px 18px;
            border-radius: 10px;
            font-size: 14px;
            font-weight: 800;
            cursor: pointer;
            transition: all 0.3s ease;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            box-shadow: 0 2px 8px rgba(102,126,234,0.2);
        }
        .pagination-btn:hover:not(:disabled) {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(102,126,234,0.4);
        }
        .pagination-btn.active {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            transform: scale(1.1);
            box-shadow: 0 6px 20px rgba(102,126,234,0.5);
        }
        .pagination-btn:disabled {
            opacity: 0.3;
            cursor: not-allowed;
            background: #e2e8f0;
            border-color: #cbd5e0;
            color: #94a3b8;
        }
        .pagination-info {
            font-family: 'Poppins', sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 10px 20px;
            border-radius: 10px;
            font-size: 14px;
            font-weight: 800;
            box-shadow: 0 4px 12px rgba(102,126,234,0.3);
        }
        
        /* Empty State */
        .empty-state {
            text-align: center;
            padding: 80px 20px;
            background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
            border-radius: 12px;
            margin: 20px 0;
        }
        .empty-state-icon {
            font-size: 72px;
            margin-bottom: 20px;
            filter: drop-shadow(2px 4px 6px rgba(0,0,0,0.1));
        }
        .empty-state-text {
            font-family: 'Poppins', sans-serif;
            font-size: 20px;
            font-weight: 700;
            color: #64748b;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .timestamp {
            font-family: 'Inter', sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%);
            color: white;
            font-size: 14px;
            font-weight: 800;
            text-align: center;
            margin-top: 40px;
            padding: 20px 30px;
            border-radius: 12px;
            border: 3px solid rgba(255,255,255,0.3);
            box-shadow: 0 8px 24px rgba(0,0,0,0.2);
            text-transform: uppercase;
            letter-spacing: 0.8px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Missing Field Summary Report</h1>
            <div class="report-info">
                <span class="info-badge">📅 """ + report_datetime + """</span>
                <span class="info-badge">📁 """ + str(len(missing_df)) + """ Total Records</span>
            </div>
        </div>
        
        <!-- Tabs -->
        <div class="tabs-container">
            <div class="tabs" id="tabsContainer">
                <button class="tab-btn summary-tab active" data-tab="Overall PT Lead Summary" onclick="switchTab('Overall PT Lead Summary')">Overall PT Lead Summary</button>
                <button class="tab-btn summary-tab" data-tab="UAT Summary" onclick="switchTab('UAT Summary')">UAT Summary</button>
                <button class="tab-btn" data-tab="All Fields" onclick="switchTab('All Fields')">All Fields</button>
""")
    
    # Add tab buttons for each field
    for field in fields_display:
        field_data = missing_df[missing_df["Field"] == field]
        if not field_data.empty:
            html_parts.append(f'                <button class="tab-btn" data-tab="{field}" onclick="switchTab(\'{field}\')">{field}</button>\n')
    
    html_parts.append("""            </div>
        </div>
        
        <!-- Tab Contents -->
        <div id="tabContents">
            
            <!-- Overall PT Lead Summary Tab -->
            <div class="tab-content active" data-tab-content="Overall PT Lead Summary">
                <div class="summary-table-container">
                    <table class="summary-table" id="overallSummaryTable">
                        <thead>
                            <tr>
                                <th>PT Lead</th>
                                <th>All Fields</th>
""")
    
    # Add column headers for each field
    for field in fields_display:
        html_parts.append(f'                                <th>{field}</th>\n')
    
    html_parts.append("""                            </tr>
                        </thead>
                        <tbody id="summaryTableBody">
                            <!-- Will be populated by JavaScript -->
                        </tbody>
                    </table>
                </div>
            </div>
""")
    
    # Add UAT Summary Tab
    html_parts.append("""
            <!-- UAT Summary Tab -->
            <div class="tab-content" data-tab-content="UAT Summary">
                <div class="summary-table-container">
                    <table class="summary-table" id="uatSummaryTable">
                        <thead>
                            <tr>
                                <th>UAT Lead</th>
                                <th>All Fields</th>
""")
    
    # Add column headers for UAT Summary (exclude mySP RCA)
    for field in fields_display:
        if field != 'mySP RCA':
            html_parts.append(f'                                <th>{field}</th>\n')
    
    html_parts.append("""                            </tr>
                        </thead>
                        <tbody id="uatSummaryTableBody">
                            <!-- Will be populated by JavaScript -->
                        </tbody>
                    </table>
                </div>
            </div>
""")
    
    # Generate tab content for "All Fields"
    html_parts.append("""
            <!-- All Fields Tab -->
            <div class="tab-content" data-tab-content="All Fields">
                <div class="filter-section">
                    <div class="filter-group">
                        <label class="filter-label">👥 Filter by PT Lead:</label>
                        <div class="multiselect-wrapper">
                            <div class="multiselect-display" onclick="toggleMultiselect(event, 'All Fields')">
                                <span>All PT Leads</span>
                                <span style="color: #999;">▼</span>
                            </div>
                            <div class="multiselect-dropdown">
                                <div class="multiselect-option select-all">
                                    <input type="checkbox" id="selectAll_All_Fields" onchange="handleSelectAll('All Fields')">
                                    <label><strong>Select All</strong></label>
                                </div>
""")
    
    for pt in pt_leads:
        html_parts.append(f'                                <div class="multiselect-option"><input type="checkbox" value="{pt}" class="pt-checkbox-All_Fields" onchange="handlePTChange(\'All Fields\')"><label>{pt}</label></div>\n')
    
    html_parts.append("""                            </div>
                        </div>
                    </div>
                    <button class="clear-filter-btn" onclick="clearTabFilters('All Fields')">🔄 Clear Filters</button>
                </div>
                
                <div class="table-container">
                    <table id="table_All_Fields">
                        <thead>
                            <tr>
                                <th>Work Item ID</th>
                                <th>Title</th>
                                <th>Field</th>
                                <th>PT Lead</th>
                                <th>State</th>
                                <th>Created By</th>
                                <th style="text-align: center;">Link</th>
                            </tr>
                        </thead>
                        <tbody id="tbody_All_Fields"></tbody>
                    </table>
                </div>
                
                <div class="pagination-container" id="pagination_All_Fields"></div>
            </div>
""")
    
    # Generate tab content for each field
    for field in fields_display:
        field_data = missing_df[missing_df["Field"] == field]
        if field_data.empty:
            continue
        
        safe_field_id = field.replace(' ', '_').replace('/', '_').replace('-', '_')
        
        html_parts.append(f"""
            <!-- {field} Tab -->
            <div class="tab-content" data-tab-content="{field}">
                <div class="filter-section">
                    <div class="filter-group">
                        <label class="filter-label">👥 Filter by PT Lead:</label>
                        <div class="multiselect-wrapper">
                            <div class="multiselect-display" onclick="toggleMultiselect(event, '{field}')">
                                <span>All PT Leads</span>
                                <span style="color: #999;">▼</span>
                            </div>
                            <div class="multiselect-dropdown">
                                <div class="multiselect-option select-all">
                                    <input type="checkbox" id="selectAll_{safe_field_id}" onchange="handleSelectAll('{field}')">
                                    <label><strong>Select All</strong></label>
                                </div>
""")
        
        for pt in pt_leads:
            html_parts.append(f'                                <div class="multiselect-option"><input type="checkbox" value="{pt}" class="pt-checkbox-{safe_field_id}" onchange="handlePTChange(\'{field}\')"><label>{pt}</label></div>\n')
        
        html_parts.append(f"""                            </div>
                        </div>
                    </div>
                    <button class="clear-filter-btn" onclick="clearTabFilters('{field}')">🔄 Clear Filters</button>
                </div>
                
                <div class="table-container">
                    <table id="table_{safe_field_id}">
                        <thead>
                            <tr>
                                <th>Work Item ID</th>
                                <th>Title</th>
                                <th>PT Lead</th>
                                <th>State</th>
                                <th>Created By</th>
                                <th style="text-align: center;">Link</th>
                            </tr>
                        </thead>
                        <tbody id="tbody_{safe_field_id}"></tbody>
                    </table>
                </div>
                
                <div class="pagination-container" id="pagination_{safe_field_id}"></div>
            </div>
""")
    
    html_parts.append("""        </div>
        
        <div class="timestamp">
            Report generated on """ + report_datetime + """<br>
            Azure DevOps Organization: """ + ORG + """ | Project: """ + PROJECT + """
        </div>
    </div>
    
    <script>
        // Data structure - records organized by field
        const recordsByField = """ + json.dumps(records_by_field, default=str) + """;
        
        // State management
        let currentPage = {};
        let filteredRecords = {};
        const RECORDS_PER_PAGE = 10;
        
        // Initialize pagination state for all tabs
        function initializeTabs() {
            const tabs = ['All Fields'];""")
    
    for field in fields_display:
        if field in records_by_field and records_by_field[field]:
            html_parts.append(f"\n            tabs.push('{field}');")
    
    html_parts.append("""
            
            tabs.forEach(tab => {
                currentPage[tab] = 1;
                filteredRecords[tab] = recordsByField[tab] || [];
            });
            
            // Populate both summary tables
            populateSummaryTable();
            populateUATSummaryTable();
        }
        
        // Populate Overall PT Lead Summary Table
        function populateSummaryTable() {
            const allFields = ['All Fields'];""")
    
    for field in fields_display:
        if field in records_by_field and records_by_field[field]:
            html_parts.append(f"\n            allFields.push('{field}');")
    
    html_parts.append("""
            
            // Get all unique PT Leads
            const allPTLeads = new Set();
            Object.values(recordsByField).forEach(records => {
                records.forEach(record => {
                    if (record['PT Lead']) {
                        allPTLeads.add(record['PT Lead']);
                    }
                });
            });
            
            // Filter out specific PT Leads for Overall PT Lead Summary only
            const excludedPTLeads = ['BA Team', 'Dev Team', 'PO Team', 'akila.krishnamoorth', 'Unmapped'];
            const sortedPTLeads = Array.from(allPTLeads)
                .filter(pt => !excludedPTLeads.includes(pt))
                .sort();
            
            // Calculate counts
            const counts = {};
            sortedPTLeads.forEach(ptLead => {
                counts[ptLead] = {};
                allFields.forEach(field => {
                    const records = recordsByField[field] || [];
                    const count = records.filter(r => r['PT Lead'] === ptLead).length;
                    counts[ptLead][field] = count;
                });
            });
            
            // Calculate grand totals (only for non-excluded PT Leads)
            const grandTotals = {};
            allFields.forEach(field => {
                grandTotals[field] = 0;
                sortedPTLeads.forEach(ptLead => {
                    grandTotals[field] += counts[ptLead][field];
                });
            });
            
            // Build table HTML
            const tbody = document.getElementById('summaryTableBody');
            let html = '';
            
            sortedPTLeads.forEach(ptLead => {
                html += '<tr>';
                html += `<td>${ptLead}</td>`;
                allFields.forEach(field => {
                    const count = counts[ptLead][field];
                    const cellClass = count === 0 ? 'count-cell zero' : 'count-cell';
                    html += `<td class="${cellClass}">${count}</td>`;
                });
                html += '</tr>';
            });
            
            // Add grand total row (based on visible PT Leads only)
            html += '<tr class="grand-total-row">';
            html += '<td>Grand Total</td>';
            allFields.forEach(field => {
                html += `<td class="total-count">${grandTotals[field]}</td>`;
            });
            html += '</tr>';
            
            tbody.innerHTML = html;
        }
        
        // Populate UAT Summary Table (akila.krishnamoorth only)
        function populateUATSummaryTable() {
            const allFields = ['All Fields'];""")
    
    for field in fields_display:
        if field in records_by_field and records_by_field[field] and field != 'mySP RCA':
            html_parts.append(f"\n            allFields.push('{field}');")
    
    html_parts.append("""
            
            const tbody = document.getElementById('uatSummaryTableBody');
            
            // Only show akila.krishnamoorth
            const uatPTLead = 'akila.krishnamoorth';
            
            // Calculate counts
            const counts = {};
            counts[uatPTLead] = {};
            allFields.forEach(field => {
                const records = recordsByField[field] || [];
                let count;
                if (field === 'All Fields') {
                    // For All Fields in UAT Summary, exclude mySP RCA records
                    count = records.filter(r => r['PT Lead'] === uatPTLead && r['Field'] !== 'mySP RCA').length;
                } else {
                    count = records.filter(r => r['PT Lead'] === uatPTLead).length;
                }
                counts[uatPTLead][field] = count;
            });
            
            // Calculate grand totals
            const grandTotals = {};
            allFields.forEach(field => {
                grandTotals[field] = counts[uatPTLead][field];
            });
            
            // Build table HTML
            let html = '';
            html += '<tr>';
            html += `<td>${uatPTLead}</td>`;
            allFields.forEach(field => {
                const count = counts[uatPTLead][field];
                const cellClass = count === 0 ? 'count-cell zero' : 'count-cell';
                html += `<td class="${cellClass}">${count}</td>`;
            });
            html += '</tr>';
            
            // Add grand total row
            html += '<tr class="grand-total-row">';
            html += '<td>Grand Total</td>';
            allFields.forEach(field => {
                html += `<td class="total-count">${grandTotals[field]}</td>`;
            });
            html += '</tr>';
            
            tbody.innerHTML = html;
        }
        
        // Switch between tabs
        function switchTab(tabName) {
            // Update tab buttons
            document.querySelectorAll('.tab-btn').forEach(btn => {
                btn.classList.remove('active');
            });
            document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
            
            // Update tab contents
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
            });
            document.querySelector(`[data-tab-content="${tabName}"]`).classList.add('active');
            
            // Handle different tab types
            if (tabName === 'Overall PT Lead Summary') {
                populateSummaryTable();
            } else if (tabName === 'UAT Summary') {
                populateUATSummaryTable();
            } else {
                // Reset to page 1 when switching tabs
                currentPage[tabName] = 1;
                renderTable(tabName);
            }
        }
        
        // Toggle multiselect dropdown
        function toggleMultiselect(event, tabName) {
            event.stopPropagation();
            const wrapper = event.target.closest('.multiselect-wrapper');
            const display = wrapper.querySelector('.multiselect-display');
            const dropdown = wrapper.querySelector('.multiselect-dropdown');
            
            display.classList.toggle('active');
            dropdown.classList.toggle('show');
        }
        
        // Close dropdown when clicking outside
        document.addEventListener('click', function(event) {
            document.querySelectorAll('.multiselect-wrapper').forEach(wrapper => {
                if (!wrapper.contains(event.target)) {
                    const display = wrapper.querySelector('.multiselect-display');
                    const dropdown = wrapper.querySelector('.multiselect-dropdown');
                    display.classList.remove('active');
                    dropdown.classList.remove('show');
                }
            });
        });
        
        // Handle select all checkbox
        function handleSelectAll(tabName) {
            const safeTabName = tabName.replace(/ /g, '_').replace(/\\//g, '_').replace(/-/g, '_');
            const selectAllId = `selectAll_${safeTabName}`;
            const selectAllCheckbox = document.getElementById(selectAllId);
            const checkboxClass = `pt-checkbox-${safeTabName}`;
            const checkboxes = document.querySelectorAll(`.${checkboxClass}`);
            
            checkboxes.forEach(cb => {
                cb.checked = selectAllCheckbox.checked;
            });
            
            handlePTChange(tabName);
        }
        
        // Handle PT Lead selection change
        function handlePTChange(tabName) {
            applyFilters(tabName);
        }
        
        // Apply PT Lead filters
        function applyFilters(tabName) {
            const safeTabName = tabName.replace(/ /g, '_').replace(/\\//g, '_').replace(/-/g, '_');
            const checkboxClass = `pt-checkbox-${safeTabName}`;
            const selectedPTLeads = Array.from(document.querySelectorAll(`.${checkboxClass}:checked`))
                .map(cb => cb.value);
            
            const allRecords = recordsByField[tabName] || [];
            
            if (selectedPTLeads.length === 0) {
                filteredRecords[tabName] = allRecords;
            } else {
                filteredRecords[tabName] = allRecords.filter(record => 
                    selectedPTLeads.includes(record['PT Lead'])
                );
            }
            
            // Update display text
            updateMultiselectDisplay(tabName, selectedPTLeads.length);
            
            // Reset to page 1 and render
            currentPage[tabName] = 1;
            renderTable(tabName);
        }
        
        // Update multiselect display text
        function updateMultiselectDisplay(tabName, selectedCount) {
            const tabContent = document.querySelector(`[data-tab-content="${tabName}"]`);
            const displaySpan = tabContent.querySelector('.multiselect-display span');
            
            if (selectedCount === 0) {
                displaySpan.textContent = 'All PT Leads';
            } else if (selectedCount === 1) {
                const safeTabName = tabName.replace(/ /g, '_').replace(/\\//g, '_').replace(/-/g, '_');
                const checkboxClass = `pt-checkbox-${safeTabName}`;
                const selected = document.querySelector(`.${checkboxClass}:checked`);
                displaySpan.textContent = selected ? selected.value : 'All PT Leads';
            } else {
                displaySpan.textContent = `${selectedCount} PT Leads selected`;
            }
        }
        
        // Clear filters for a tab
        function clearTabFilters(tabName) {
            const safeTabName = tabName.replace(/ /g, '_').replace(/\\//g, '_').replace(/-/g, '_');
            const selectAllId = `selectAll_${safeTabName}`;
            const checkboxClass = `pt-checkbox-${safeTabName}`;
            
            // Uncheck all
            document.getElementById(selectAllId).checked = false;
            document.querySelectorAll(`.${checkboxClass}`).forEach(cb => cb.checked = false);
            
            // Reset filters
            filteredRecords[tabName] = recordsByField[tabName] || [];
            updateMultiselectDisplay(tabName, 0);
            
            // Reset to page 1 and render
            currentPage[tabName] = 1;
            renderTable(tabName);
        }
        
        // Render table for current tab
        function renderTable(tabName) {
            const safeTabName = tabName.replace(/ /g, '_').replace(/\\//g, '_').replace(/-/g, '_');
            const tbody = document.getElementById(`tbody_${safeTabName}`);
            const records = filteredRecords[tabName] || [];
            
            if (records.length === 0) {
                tbody.innerHTML = `
                    <tr>
                        <td colspan="7" style="text-align: center; padding: 40px;">
                            <div class="empty-state">
                                <div class="empty-state-icon">📭</div>
                                <div class="empty-state-text">No records found</div>
                            </div>
                        </td>
                    </tr>
                `;
                renderPagination(tabName, 0);
                return;
            }
            
            const page = currentPage[tabName] || 1;
            const startIdx = (page - 1) * RECORDS_PER_PAGE;
            const endIdx = startIdx + RECORDS_PER_PAGE;
            const pageRecords = records.slice(startIdx, endIdx);
            
            let html = '';
            pageRecords.forEach(record => {
                const workItemId = record['Work Item ID'] || '';
                const title = record['Title'] || '';
                const field = record['Field'] || '';
                const ptLead = record['PT Lead'] || '';
                const state = record['State'] || '';
                const createdBy = record['Created By'] || '';
                const url = record['URL'] || '';
                
                if (tabName === 'All Fields') {
                    html += `
                        <tr>
                            <td class="work-item-id">${workItemId}</td>
                            <td>${title}</td>
                            <td>${field}</td>
                            <td>${ptLead}</td>
                            <td>${state}</td>
                            <td>${createdBy}</td>
                            <td style="text-align: center;"><a href="${url}" target="_blank" class="url-link">Open</a></td>
                        </tr>
                    `;
                } else {
                    html += `
                        <tr>
                            <td class="work-item-id">${workItemId}</td>
                            <td>${title}</td>
                            <td>${ptLead}</td>
                            <td>${state}</td>
                            <td>${createdBy}</td>
                            <td style="text-align: center;"><a href="${url}" target="_blank" class="url-link">Open</a></td>
                        </tr>
                    `;
                }
            });
            
            tbody.innerHTML = html;
            renderPagination(tabName, records.length);
        }
        
        // Render pagination controls
        function renderPagination(tabName, totalRecords) {
            const safeTabName = tabName.replace(/ /g, '_').replace(/\\//g, '_').replace(/-/g, '_');
            const container = document.getElementById(`pagination_${safeTabName}`);
            
            if (totalRecords <= RECORDS_PER_PAGE) {
                container.innerHTML = '';
                return;
            }
            
            const totalPages = Math.ceil(totalRecords / RECORDS_PER_PAGE);
            const page = currentPage[tabName] || 1;
            
            let html = '';
            
            // Previous button
            html += `<button class="pagination-btn" ${page === 1 ? 'disabled' : ''} 
                     onclick="changePage('${tabName}', ${page - 1})">← Previous</button>`;
            
            // Page numbers
            for (let i = 1; i <= totalPages; i++) {
                if (i === 1 || i === totalPages || (i >= page - 2 && i <= page + 2)) {
                    html += `<button class="pagination-btn ${i === page ? 'active' : ''}" 
                             onclick="changePage('${tabName}', ${i})">${i}</button>`;
                } else if (i === page - 3 || i === page + 3) {
                    html += `<span class="pagination-info">...</span>`;
                }
            }
            
            // Next button
            html += `<button class="pagination-btn" ${page === totalPages ? 'disabled' : ''} 
                     onclick="changePage('${tabName}', ${page + 1})">Next →</button>`;
            
            // Info
            const startIdx = (page - 1) * RECORDS_PER_PAGE + 1;
            const endIdx = Math.min(page * RECORDS_PER_PAGE, totalRecords);
            html += `<span class="pagination-info">Showing ${startIdx}-${endIdx} of ${totalRecords}</span>`;
            
            container.innerHTML = html;
        }
        
        // Change page
        function changePage(tabName, newPage) {
            currentPage[tabName] = newPage;
            renderTable(tabName);
        }
        
        // Initialize on page load
        window.onload = function() {
            initializeTabs();
        };
    </script>
</body>
</html>
""")
    
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(''.join(html_parts))
    
    print(f"HTML report saved to: {html_path}")
    print(f"  {len(detail_files)} detail pages in: {details_dir}")
    return html_path


def main():
    pat = get_pat()

    fields_url = f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/wit/fields?api-version=7.1"
    fields_data = azdo_get(fields_url, pat)

    name_to_ref = {}
    for f in fields_data.get("value", []):
        name_to_ref[normalize_field_name(f.get("name", ""))] = f.get("referenceName")

    field_ref_map = {}
    for display in FIELDS_DISPLAY:
        ref = name_to_ref.get(normalize_field_name(display))
        if not ref and display in FIELD_ALIASES:
            for alias in FIELD_ALIASES[display]:
                ref = name_to_ref.get(normalize_field_name(alias))
                if ref:
                    break
        if not ref:
            # try loose match
            for k, v in name_to_ref.items():
                if normalize_field_name(display) == k or normalize_field_name(display) in k:
                    ref = v
                    break
        if not ref:
            raise RuntimeError(f"Could not resolve field reference name for: {display}")
        field_ref_map[display] = ref

    wiql_url = f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/wit/wiql/{QUERY_ID}?api-version=7.1"
    wiql_data = azdo_get(wiql_url, pat)
    work_items = wiql_data.get("workItems", [])
    ids = [w["id"] for w in work_items]

    if not ids:
        raise RuntimeError("No work items returned by the query.")

    fields = [
        "System.Id",
        "System.Title",
        "System.State",
        "System.CreatedBy",
        "System.WorkItemType",
    ] + list(field_ref_map.values())

    all_items = []
    for batch in chunked(ids, 50):
        ids_str = ",".join(str(i) for i in batch)
        work_url = (
            f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/wit/workitems?ids={ids_str}"
            f"&fields={','.join(fields)}&api-version=7.1"
        )
        data = azdo_get(work_url, pat)
        all_items.extend(data.get("value", []))

    mapping_df = pd.read_excel(PT_LEAD_MAPPING_PATH)
    mapping_df["Enterprise ID"] = mapping_df["Enterprise ID"].astype(str).str.strip().str.lower()
    mapping_df["PT Lead Name"] = mapping_df["PT Lead Name"].astype(str).str.strip()

    enterprise_to_pt = dict(zip(mapping_df["Enterprise ID"], mapping_df["PT Lead Name"]))

    missing_records = []
    for item in all_items:
        fields_data = item.get("fields", {})
        created_by = fields_data.get("System.CreatedBy")
        enterprise_id = extract_enterprise_id(created_by)
        pt_lead = enterprise_to_pt.get(enterprise_id, "Unmapped")

        state = fields_data.get("System.State", "") or ""

        for display, ref in field_ref_map.items():
            if display == "mySP RCA" and str(state).lower() != "closed":
                continue
            
            # Only show Closed state defects for "Fixed in mySP Release Name"
            if display == "Fixed in mySP Release Name" and str(state).lower() != "closed":
                continue

            value = fields_data.get(ref)
            if is_missing(value):
                missing_records.append(
                    {
                        "PT Lead": pt_lead,
                        "Field": display,
                        "Work Item ID": item.get("id"),
                        "Title": fields_data.get("System.Title", ""),
                        "State": state,
                        "Created By": (
                            created_by.get("displayName") if isinstance(created_by, dict) else str(created_by)
                        ),
                        "URL": f"https://dev.azure.com/{ORG}/{PROJECT}/_workitems/edit/{item.get('id')}",
                        "Field Value": value,
                    }
                )

    if not missing_records:
        raise RuntimeError("No missing values detected for the requested fields.")

    missing_df = pd.DataFrame(missing_records)

    summary = (
        missing_df.groupby(["Field", "PT Lead"]).size().reset_index(name="Missing Count")
    )

    pt_leads = sorted(missing_df["PT Lead"].unique())

    def write_report(output_path: str):
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            workbook = writer.book
            title_format = workbook.add_format({"bold": True, "font_size": 14})
            header_format = workbook.add_format({"bold": True, "bg_color": "#D9E1F2"})
            link_format = workbook.add_format({"font_color": "#0563C1", "underline": 1})

            def write_summary_and_details(sheet_prefix: str, fields_display, data_df):
                summary_sheet = f"{sheet_prefix} Summary"
                details_sheet = f"{sheet_prefix} Details"

                summary_ws = workbook.add_worksheet(summary_sheet)
                details_ws = workbook.add_worksheet(details_sheet)
                writer.sheets[summary_sheet] = summary_ws
                writer.sheets[details_sheet] = details_ws

                summary_local = (
                    data_df.groupby(["Field", "PT Lead"]).size().reset_index(name="Missing Count")
                )
                pt_leads_local = sorted(data_df["PT Lead"].unique())

                row = 0
                summary_ws.write(row, 0, "Missing Field Summary", title_format)
                row += 2

                group_rows = {}

                for field in fields_display:
                    field_summary = summary_local[summary_local["Field"] == field]
                    if field_summary.empty:
                        continue

                    summary_ws.write(row, 0, field, header_format)
                    row += 1
                    summary_ws.write(row, 0, "PT Lead", header_format)
                    summary_ws.write(row, 1, field, header_format)
                    row += 1

                    for pt in pt_leads_local:
                        count_series = field_summary[field_summary["PT Lead"] == pt]["Missing Count"]
                        count = int(count_series.iloc[0]) if not count_series.empty else 0
                        summary_ws.write(row, 0, pt)
                        summary_ws.write(row, 1, count)
                        row += 1

                    row += 2

                details_start = 0
                details_ws.write(details_start, 0, "Defect Details (Missing Fields)", title_format)
                details_start += 2

                details_columns = [
                    "PT Lead",
                    "Field",
                    "Work Item ID",
                    "Title",
                    "State",
                    "Created By",
                    "URL",
                    "Field Value",
                ]
                for col, name in enumerate(details_columns):
                    details_ws.write(details_start, col, name, header_format)

                details_row = details_start + 1
                details_df = data_df.sort_values(by=["PT Lead", "Field"], kind="mergesort").reset_index(drop=True)
                
                for _, rec in details_df.iterrows():
                    group_key = (rec["Field"], rec["PT Lead"])
                    group_rows.setdefault(group_key, []).append(details_row)

                    details_ws.write(details_row, 0, rec["PT Lead"])
                    details_ws.write(details_row, 1, rec["Field"])
                    details_ws.write(details_row, 2, rec["Work Item ID"])
                    details_ws.write(details_row, 3, rec["Title"])
                    details_ws.write(details_row, 4, rec["State"])
                    details_ws.write(details_row, 5, rec["Created By"])
                    details_ws.write_url(details_row, 6, rec["URL"], link_format, string=rec["URL"])
                    details_ws.write(details_row, 7, rec["Field Value"])
                    details_row += 1

                row = 2
                for field in fields_display:
                    field_summary = summary_local[summary_local["Field"] == field]
                    if field_summary.empty:
                        continue

                    row += 1
                    row += 1
                    for pt in pt_leads_local:
                        count_series = field_summary[field_summary["PT Lead"] == pt]["Missing Count"]
                        count = int(count_series.iloc[0]) if not count_series.empty else 0
                        if count > 0:
                            group_key = (field, pt)
                            rows = group_rows.get(group_key)
                            if rows:
                                first_row = min(rows) + 1   # 1-based Excel row
                                last_row = max(rows) + 1
                                # Link selects the full range of matching rows
                                summary_ws.write_url(
                                    row,
                                    1,
                                    f"internal:'{details_sheet}'!A{first_row}:H{last_row}",
                                    link_format,
                                    string=str(count),
                                )
                        row += 1
                    row += 2

                summary_ws.set_column(0, 0, 24)
                summary_ws.set_column(1, 1, 18)

                details_ws.set_column(0, 0, 24)
                details_ws.set_column(1, 1, 18)
                details_ws.set_column(2, 2, 14)
                details_ws.set_column(3, 3, 50)
                details_ws.set_column(4, 4, 12)
                details_ws.set_column(5, 5, 22)
                details_ws.set_column(6, 6, 70)
                details_ws.set_column(7, 7, 22)

            write_summary_and_details("All Fields", FIELDS_DISPLAY, missing_df)

    try:
        write_report(OUTPUT_PATH)
        final_output = OUTPUT_PATH
    except PermissionError:
        timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        alt_output = OUTPUT_PATH.replace(
            ".xlsx", f"_{timestamp}.xlsx"
        )
        write_report(alt_output)
        final_output = alt_output

    # Post-process: inject VBA macro and convert to .xlsm
    final_output = add_vba_highlighting(final_output)

    print(f"Excel report saved to: {final_output}")
    
    # Generate HTML report
    html_output = generate_html_report(missing_df, FIELDS_DISPLAY, final_output)
    print(f"\n✓ Reports generated successfully!")
    print(f"  • Excel (with interactive highlighting): {final_output}")
    print(f"  • HTML (web-based view): {html_output}")


if __name__ == "__main__":
    main()
