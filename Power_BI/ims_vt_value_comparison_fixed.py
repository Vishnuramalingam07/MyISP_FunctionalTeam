"""
IMS VT Comparison Tool - Fixed Version
Correctly handles IMSVT structure where Guidance/AsPerSolution/Variance are data values, not headers
"""
import pandas as pd
import numpy as np
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)

class IMSVTComparisonFixed:
    def __init__(self):
        self.mapping_file = "IMS VT Automation Mappings.xlsx"
        self.imsvt_file = "IMSVT.xlsx"
        self.mappings = []
        self.imsvt_df = None
        self.imsvt_column_map = {}  # Maps (main_header, label) -> actual_column
        
        # Map IMS label names to actual IMSVT label names
        self.label_map = {
            'Solution Standards': 'Guidance',
            'Actual Value': 'AsPerSolution',
            'Variation From Standard': 'Variance',
            'Guidance': 'Guidance',
            'AsPerSolution': 'AsPerSolution',
            'Variance': 'Variance'
        }
        
    def load_mappings(self):
        """Load column mappings from the mapping file"""
        df = pd.read_excel(self.mapping_file, header=None)
        
        # Parse mappings (rows starting from 2, every 3 rows per metric)
        for i in range(2, len(df), 3):
            if pd.notna(df.iloc[i, 3]) and pd.notna(df.iloc[i, 6]):
                imsvt_main = str(df.iloc[i, 3]).strip()
                ims_main = str(df.iloc[i, 6]).strip()
                
                # Add all 3 sub-column mappings
                for j in range(3):
                    if i+j < len(df):
                        imsvt_label = str(df.iloc[i+j, 4]).strip()  # Guidance/AsPerSolution/Variance
                        ims_label = str(df.iloc[i+j, 7]).strip()
                        
                        mapping = {
                            'imsvt_main': imsvt_main,
                            'imsvt_label': imsvt_label,  # What we're looking for (Guidance/AsPerSolution/Variance)
                            'ims_main': ims_main,
                            'ims_label': ims_label,
                            'row_offset': j
                        }
                        self.mappings.append(mapping)
        
        logger.info(f"✓ Loaded {len(self.mappings)} comparison mappings")
        return len(self.mappings)
    
    def load_imsvt(self):
        """
        Load IMSVT with special handling for its structure
        - Headers are at rows 2-3 (0-indexed)
        - First data row (row 4) contains labels: Guidance, AsPerSolution, Variance
        """
        logger.info("\n" + "="*80)
        logger.info("LOADING IMSVT DATA")
        logger.info("="*80)
        
        # Load with multi-level headers (rows 3 and 4 in Excel = indices 2,3)
        self.imsvt_df = pd.read_excel(self.imsvt_file, header=[2, 3])
        logger.info(f"✓ Loaded IMSVT: {len(self.imsvt_df)} rows, {len(self.imsvt_df.columns)} columns")
        
        # Build a map of (main_header, label) -> actual_column
        # by checking the first data row for Guidance/AsPerSolution/Variance
        logger.info("\n📍 Mapping columns by their data labels...")
        
        # Group columns by main header
        main_headers = {}
        for col in self.imsvt_df.columns:
            if isinstance(col, tuple) and len(col) >= 1:
                main = str(col[0]).strip()
                if 'Unnamed' not in main:
                    if main not in main_headers:
                        main_headers[main] = []
                    main_headers[main].append(col)
        
        # For each main header, find which columns contain Guidance, AsPerSolution, Variance
        for main_header, cols in main_headers.items():
            # Check first data row (index 0) for each column
            for col in cols:
                if len(self.imsvt_df) > 0:
                    first_value = str(self.imsvt_df[col].iloc[0]).strip()
                    if first_value in ['Guidance', 'AsPerSolution', 'Variance']:
                        key = (main_header, first_value)
                        self.imsvt_column_map[key] = col
        
        logger.info(f"✓ Mapped {len(self.imsvt_column_map)} labeled columns")
        
        # Show sample mappings
        sample_count = 0
        for key, col in self.imsvt_column_map.items():
            if 'Offshore Ratio' in key[0] and sample_count < 3:
                logger.info(f"   Example: {key[0]} + {key[1]} -> Column {col}")
                sample_count += 1
        
        return self.imsvt_df
    
    def get_imsvt_value(self, main_header, label, row_index=2):
        """
        Get value from IMSVT for a specific main header and label
        
        Args:
            main_header: Main column header (e.g., "Managed Security - Offshore Ratio (%)")
            label: Data label (e.g., "Guidance", "AsPerSolution", "Variance")
            row_index: Row index to get value from (default 2, where actual data starts)
        """
        key = (main_header, label)
        if key in self.imsvt_column_map:
            col = self.imsvt_column_map[key]
            if row_index < len(self.imsvt_df):
                return self.imsvt_df[col].iloc[row_index]
        return None
    
    def get_ims_value(self, main_header, label, row_index=2):
        """
        Get value from IMS columns within IMSVT file
        
        Args:
            main_header: Main column header (e.g., "Offshore Ratio (%)")
            label: Data label from mapping file (e.g., "Solution Standards", "Actual Value")
            row_index: Row index to get value from (default 2, where actual data starts)
        """
        # IMS columns are also in the IMSVT file
        # Map the IMS label to the actual label used in IMSVT
        actual_label = self.label_map.get(label, label)
        key = (main_header, actual_label)
        if key in self.imsvt_column_map:
            col = self.imsvt_column_map[key]
            if row_index < len(self.imsvt_df):
                return self.imsvt_df[col].iloc[row_index]
        return None
    
    def compare_values(self):
        """Compare values between IMSVT and IMS columns"""
        logger.info("\n" + "="*80)
        logger.info("COMPARING VALUES")
        logger.info("="*80)
        
        if self.imsvt_df is None:
            logger.error("IMSVT data not loaded")
            return None
        
        results = []
        found_count = 0
        not_found_count = 0
        
        for i, mapping in enumerate(self.mappings):
            imsvt_main = mapping['imsvt_main']
            imsvt_label = mapping['imsvt_label']
            ims_main = mapping['ims_main']
            ims_label = mapping['ims_label']
            
            # Get values from both sides
            imsvt_value = self.get_imsvt_value(imsvt_main, imsvt_label, row_index=2)
            ims_value = self.get_ims_value(ims_main, ims_label, row_index=2)
            
            # Check if columns were found
            imsvt_found = (imsvt_main, imsvt_label) in self.imsvt_column_map
            # For IMS, check with the mapped label
            ims_actual_label = self.label_map.get(ims_label, ims_label)
            ims_found = (ims_main, ims_actual_label) in self.imsvt_column_map
            
            if not imsvt_found:
                not_found_count += 1
                logger.warning(f"⚠ IMSVT column not found: {imsvt_main} / {imsvt_label}")
                continue
            
            if not ims_found:
                not_found_count += 1
                logger.warning(f"⚠ IMS column not found: {ims_main} / {ims_label}")
                continue
            
            found_count += 1
            if (i + 1) % 10 == 1:  # Log progress every 10 items
                logger.info(f"  {i+1}/{len(self.mappings)}: Comparing {imsvt_main}/{imsvt_label} vs {ims_main}/{ims_label}")
            
            # Compare values
            status = "MATCH"
            match_exact = False
            match_close = False
            
            if pd.isna(imsvt_value) and pd.isna(ims_value):
                status = "BOTH_EMPTY"
                match_exact = True
            elif pd.isna(imsvt_value) or pd.isna(ims_value):
                status = "ONE_EMPTY"
            else:
                # Try exact match
                if imsvt_value == ims_value:
                    match_exact = True
                # Try numeric comparison with tolerance
                elif isinstance(imsvt_value, (int, float)) and isinstance(ims_value, (int, float)):
                    if abs(imsvt_value - ims_value) < 0.0001:
                        match_close = True
                        status = "MATCH_CLOSE"
                    else:
                        status = "MISMATCH"
                # Try string comparison
                elif str(imsvt_value).strip().lower() == str(ims_value).strip().lower():
                    match_exact = True
                else:
                    status = "MISMATCH"
            
            if match_exact:
                status = "MATCH"
            
            results.append({
                'IMSVT_Column': f"{imsvt_main} / {imsvt_label}",
                'IMS_Column': f"{ims_main} / {ims_label}",
                'IMSVT_Value': imsvt_value,
                'IMS_Value': ims_value,
                'Status': status,
                'Match': '✓' if match_exact or match_close else '✗'
            })
        
        logger.info(f"\n✓ Compared {found_count} column pairs")
        logger.info(f"⚠ {not_found_count} columns not found")
        
        if results:
            return pd.DataFrame(results)
        return None
    
    def generate_report(self, df_results):
        """Generate Excel report with comparison results"""
        if df_results is None or len(df_results) == 0:
            return None
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"IMS_VT_Comparison_Report_Fixed_{timestamp}.xlsx"
        
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Write detailed results
            df_results.to_excel(writer, sheet_name='Comparison Results', index=False)
            
            # Write summary
            matches = len(df_results[df_results['Status'].isin(['MATCH', 'MATCH_CLOSE'])])
            total = len(df_results)
            
            summary_data = {
                'Metric': ['Total Comparisons', 'Matches', 'Mismatches', 'Match %'],
                'Value': [
                    total,
                    matches,
                    total - matches,
                    f"{100 * matches / total:.1f}%" if total > 0 else "0%"
                ]
            }
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            
            # Format worksheets
            workbook = writer.book
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4472C4',
                'font_color': 'white',
                'border': 1
            })
            
            match_format = workbook.add_format({
                'bg_color': '#C6EFCE',
                'font_color': '#006100'
            })
            
            mismatch_format = workbook.add_format({
                'bg_color': '#FFC7CE',
                'font_color': '#9C0006'
            })
            
            # Format results sheet
            worksheet = writer.sheets['Comparison Results']
            worksheet.set_column('A:B', 60)
            worksheet.set_column('C:D', 20)
            worksheet.set_column('E:E', 15)
            worksheet.set_column('F:F', 10)
            
            # Apply conditional formatting
            for row_idx in range(1, len(df_results) + 1):
                status = df_results.iloc[row_idx - 1]['Status']
                if status in ['MATCH', 'MATCH_CLOSE']:
                    worksheet.write(row_idx, 5, '✓', match_format)
                else:
                    worksheet.write(row_idx, 5, '✗', mismatch_format)
        
        logger.info(f"\n✓ Report saved: {output_file}")
        return output_file

def main():
    print("\n" + "="*80)
    print("IMS VT COMPARISON TOOL - FIXED VERSION")
    print("="*80)
    print("\nThis version correctly handles IMSVT structure where:")
    print("  - Guidance/AsPerSolution/Variance are DATA VALUES (not headers)")
    print("  - Headers are at rows 3-4 with units (%, %.1, %.2)")
    print("\nUsing mappings from: IMS VT Automation Mappings.xlsx")
    print("="*80)
    
    comparator = IMSVTComparisonFixed()
    
    # Load mappings
    logger.info("\n📋 Loading comparison mappings...")
    comparator.load_mappings()
    
    # Load IMSVT data
    logger.info("\n📂 Loading IMSVT data...")
    comparator.load_imsvt()
    
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
        print("="*80)
        return
    
    # Summary
    matches = len(df_results[df_results['Status'].isin(['MATCH', 'MATCH_CLOSE'])])
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
