import pandas as pd
import logging
from datetime import datetime

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)

class IMSVTComparison:
    def __init__(self):
        self.mapping_file = "IMS VT Automation Mappings.xlsx"
        self.imsvt_file = "IMSVT.xlsx"
        self.mappings = []
        
    def load_mappings(self):
        """Load column mappings from the mapping file"""
        df = pd.read_excel(self.mapping_file, header=None)
        
        # The actual sub-column names in the Excel file are: Guidance, AsPerSolution, Variance
        # The mapping file uses descriptive names, so we map them:
        sub_col_map = {
            'Guidance': 'Guidance',
            'AsPerSolution': 'AsPerSolution',
            'Variance': 'Variance',
            'Solution Standards': 'Guidance',  # IMS side also uses 'Guidance'
            'Actual Value': 'AsPerSolution',   # IMS side also uses 'AsPerSolution'
            'Variation From Standard': 'Variance'  # IMS side also uses 'Variance'
        }
        
        # Parse mappings (rows starting from 2, every 3 rows per metric)
        for i in range(2, len(df), 3):
            if pd.notna(df.iloc[i, 3]) and pd.notna(df.iloc[i, 6]):
                imsvt_col = df.iloc[i, 3]
                imsvt_subcol_label = df.iloc[i, 4]  # From mapping file
                ims_col = df.iloc[i, 6]
                ims_subcol_label = df.iloc[i, 7]  # From mapping file
                
                # Map to actual column names
                imsvt_subcol = sub_col_map.get(imsvt_subcol_label, imsvt_subcol_label)
                ims_subcol = sub_col_map.get(ims_subcol_label, ims_subcol_label)
                
                mapping = {
                    'imsvt_main': imsvt_col,
                    'imsvt_sub': imsvt_subcol,
                    'ims_main': ims_col,
                    'ims_sub': ims_subcol,
                    'row_offset': 0
                }
                self.mappings.append(mapping)
                
                # AsPerSolution vs AsPerSolution
                if i+1 < len(df):
                    imsvt_subcol2 = sub_col_map.get(df.iloc[i+1, 4], df.iloc[i+1, 4])
                    ims_subcol2 = sub_col_map.get(df.iloc[i+1, 7], df.iloc[i+1, 7])
                    mapping2 = {
                        'imsvt_main': df.iloc[i+1, 3],
                        'imsvt_sub': imsvt_subcol2,
                        'ims_main': df.iloc[i+1, 6],
                        'ims_sub': ims_subcol2,
                        'row_offset': 1
                    }
                    self.mappings.append(mapping2)
                
                # Variance vs Variance
                if i+2 < len(df):
                    imsvt_subcol3 = sub_col_map.get(df.iloc[i+2, 4], df.iloc[i+2, 4])
                    ims_subcol3 = sub_col_map.get(df.iloc[i+2, 7], df.iloc[i+2, 7])
                    mapping3 = {
                        'imsvt_main': df.iloc[i+2, 3],
                        'imsvt_sub': imsvt_subcol3,
                        'ims_main': df.iloc[i+2, 6],
                        'ims_sub': ims_subcol3,
                        'row_offset': 2
                    }
                    self.mappings.append(mapping3)
        
        logger.info(f"✓ Loaded {len(self.mappings)} comparison mappings")
        return len(self.mappings)
    
    def load_imsvt_with_multiheader(self):
        """Load IMSVT with multi-level headers"""
        # Load with no header first to understand structure
        df_raw = pd.read_excel(self.imsvt_file, header=None)
        
        # Row 2 = main headers, Row 3 = units, Row 4 = sub-column names
        main_headers = df_raw.iloc[2, :].values
        units = df_raw.iloc[3, :].values  
        sub_headers = df_raw.iloc[4, :].values
        
        # Create multi-level column index
        columns = []
        for i in range(len(main_headers)):
            main = main_headers[i] if pd.notna(main_headers[i]) else ''
            sub = sub_headers[i] if pd.notna(sub_headers[i]) else ''
            columns.append((main, sub))
        
        # Load data (starting from row 5)
        df_data = pd.read_excel(self.imsvt_file, header=None, skiprows=5)
        df_data.columns = pd.MultiIndex.from_tuples(columns[:len(df_data.columns)])
        
        logger.info(f"✓ Loaded IMSVT: {df_data.shape[0]} rows, {df_data.shape[1]} columns")
        return df_data, df_raw
    
    def find_column_index(self, df_raw, main_col, sub_col):
        """Find the column index for a main column + sub-column combination"""
        for col_idx in range(df_raw.shape[1]):
            main_header = df_raw.iloc[2, col_idx]
            sub_header = df_raw.iloc[4, col_idx]
            
            if pd.notna(main_header) and str(main_header).strip() == str(main_col).strip():
                if pd.notna(sub_header) and str(sub_header).strip() == str(sub_col).strip():
                    return col_idx
        return None
    
    def compare_values(self):
        """Compare values between mapped columns"""
        logger.info("\n" + "="*80)
        logger.info("LOADING DATA")
        logger.info("="*80)
        
        df_data, df_raw = self.load_imsvt_with_multiheader()
        
        logger.info("\n" + "="*80)
        logger.info("COMPARING VALUES")
        logger.info("="*80)
        
        results = []
        
        for idx, mapping in enumerate(self.mappings):
            imsvt_main = mapping['imsvt_main']
            imsvt_sub = mapping['imsvt_sub']
            ims_main = mapping['ims_main']
            ims_sub = mapping['ims_sub']
            
            # Find column indices
            imsvt_col_idx = self.find_column_index(df_raw, imsvt_main, imsvt_sub)
            ims_col_idx = self.find_column_index(df_raw, ims_main, ims_sub)
            
            if imsvt_col_idx is None:
                logger.warning(f"⚠ IMSVT column not found: {imsvt_main} / {imsvt_sub}")
                continue
            
            if ims_col_idx is None:
                logger.warning(f"⚠ IMS column not found: {ims_main} / {ims_sub}")
                continue
            
            # Get column data (starting from row 5)
            imsvt_values = df_raw.iloc[5:, imsvt_col_idx]
            ims_values = df_raw.iloc[5:, ims_col_idx]
            
            # Compare row by row
            for row_idx in range(min(len(imsvt_values), len(ims_values))):
                imsvt_val = imsvt_values.iloc[row_idx]
                ims_val = ims_values.iloc[row_idx]
                
                # Skip if both are NaN
                if pd.isna(imsvt_val) and pd.isna(ims_val):
                    continue
                
                # Check for match
                match = False
                if pd.isna(imsvt_val) and pd.isna(ims_val):
                    match = True
                elif pd.notna(imsvt_val) and pd.notna(ims_val):
                    try:
                        # Try numeric comparison
                        if float(imsvt_val) == float(ims_val):
                            match = True
                    except:
                        # String comparison
                        if str(imsvt_val).strip() == str(ims_val).strip():
                            match = True
                
                results.append({
                    'Row': row_idx + 5 + 1,  # +5 for skipped rows, +1 for Excel 1-based
                    'IMSVT_Column': f"{imsvt_main}",
                    'IMSVT_SubColumn': imsvt_sub,
                    'IMSVT_Value': imsvt_val,
                    'IMS_Column': f"{ims_main}",
                    'IMS_SubColumn': ims_sub,
                    'IMS_Value': ims_val,
                    'Match': '✓' if match else '✗',
                    'Status': 'MATCH' if match else 'MISMATCH'
                })
            
            logger.info(f"  {idx+1}/{len(self.mappings)}: {imsvt_main}/{imsvt_sub} vs {ims_main}/{ims_sub}")
        
        return pd.DataFrame(results)
    
    def generate_report(self, df_results):
        """Generate Excel report with comparison results"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"IMS_VT_Comparison_Report_{timestamp}.xlsx"
        
        # Handle empty results
        if len(df_results) == 0:
            logger.warning("⚠ No comparison results to report")
            return None
        
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Summary sheet
            total = len(df_results)
            matches = len(df_results[df_results['Status'] == 'MATCH'])
            mismatches = total - matches
            match_pct = (100 * matches / total) if total > 0 else 0
            
            summary_data = {
                'Metric': ['Total Comparisons', 'Matches', 'Mismatches', 'Match %'],
                'Value': [
                    total,
                    matches,
                    mismatches,
                    f"{match_pct:.1f}%"
                ]
            }
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            
            # Detailed comparison 
            df_results.to_excel(writer, sheet_name='Detailed_Comparison', index=False)
            
            # Mismatches only
            df_mismatches = df_results[df_results['Status'] == 'MISMATCH']
            if len(df_mismatches) > 0:
                df_mismatches.to_excel(writer, sheet_name='Mismatches_Only', index=False)
            
            # Format sheets
            workbook = writer.book
            header_format = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white'})
            match_format = workbook.add_format({'bg_color': '#C6EFCE'})
            mismatch_format = workbook.add_format({'bg_color': '#FFC7CE'})
            
            for sheet_name in ['Summary', 'Detailed_Comparison', 'Mismatches_Only']:
                if sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    worksheet.freeze_panes(1, 0)
                    
                    # Auto-fit columns
                    for col_num, col_name in enumerate(df_summary.columns if sheet_name == 'Summary' else df_results.columns):
                        worksheet.set_column(col_num, col_num, 20)
        
        logger.info(f"\n✓ Report saved: {output_file}")
        return output_file

def main():
    print("\n" + "="*80)
    print("IMS VT COMPARISON TOOL - VALUE COMPARISON")
    print("="*80)
    print("\nCompares values within IMSVT.xlsx between:")
    print("  - IMSVT columns (e.g., 'Managed Security - Offshore Ratio (%)')")
    print("  - IMS columns (e.g., 'Offshore Ratio (%)')")
    print("\nUsing mappings from: IMS VT Automation Mappings.xlsx")
    print("="*80)
    
    comparator = IMSVTComparison()
    
    # Load mappings
    logger.info("\n📋 Loading comparison mappings...")
    comparator.load_mappings()
    
    # Perform comparison
    logger.info("\n🔍 Comparing values...")
    df_results = comparator.compare_values()
    
    # Generate report
    logger.info("\n📊 Generating report...")
    output_file = comparator.generate_report(df_results)
    
    if output_file is None:
        print("\n" + "="*80)
        print("NO RESULTS TO REPORT")
        print("="*80)
        print("No matching columns were found.")
        print("Please check:")
        print("  1. Column names in IMSVT.xlsx")
        print("  2. Column names in mapping file")
        print("  3. Sub-column names (Guidance, AsPerSolution, Variance)")
        print("="*80)
        return
    
    # Summary
    matches = len(df_results[df_results['Status'] == 'MATCH'])
    total = len(df_results)
    print("\n" + "="*80)
    print("COMPARISON COMPLETE")
    print("="*80)
    print(f"Total Comparisons: {total}")
    print(f"Matches: {matches}")
    print(f"Mismatches: {total - matches}")
    print(f"Match Rate: {100 * matches / total:.1f}%" if total > 0 else "Match Rate: 0%")
    print(f"\n✓ Report: {output_file}")
    print("="*80)

if __name__ == "__main__":
    main()
