"""
IMS VT Comparison Tool with Column Mapping
===========================================
This tool compares two Excel files using the column mappings defined in 
IMS VT Automation Mappings.xlsx
"""

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json
from datetime import datetime
from pathlib import Path
import logging
import sys

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def load_column_mappings(mapping_file: str = "IMS VT Automation Mappings.xlsx"):
    """
    Load column mappings from the IMS VT Automation Mappings file
    
    Returns:
        List of dictionaries with IMSVT_Column and IMS_Column mappings
    """
    logger.info(f"Loading column mappings from {mapping_file}...")
    
    # Read the raw data
    df = pd.read_excel(mapping_file, header=None)
    
    mappings = []
    sub_column_mappings = {}
    current_main_column_imsvt = None
    current_main_column_ims = None
    
    for i in range(len(df)):
        imsvt_col = df.iloc[i, 3]  # Column 3 has IMSVT column names
        imsvt_sub = df.iloc[i, 4]  # Column 4 has IMSVT sub-columnnames
        ims_col = df.iloc[i, 6]    # Column 6 has IMS column names  
        ims_sub = df.iloc[i, 7]    # Column 7 has IMS sub-column names
        
        # Skip header rows
        if pd.notna(imsvt_col):
            imsvt_str = str(imsvt_col).strip()
            if 'MainHeader' in imsvt_str or 'IMSVT' in imsvt_str or 'Columns' in imsvt_str:
                continue
        
        # Main column mapping
        if pd.notna(imsvt_col) and pd.notna(ims_col):
            imsvt_str = str(imsvt_col).strip()
            ims_str = str(ims_col).strip()
            
            if imsvt_str and ims_str:
                current_main_column_imsvt = imsvt_str
                current_main_column_ims = ims_str
                
                mappings.append({
                    'IMSVT_MainColumn': imsvt_str,
                    'IMS_MainColumn': ims_str,
                    'sub_columns': []
                })
                logger.info(f"  Mapping: '{imsvt_str}' -> '{ims_str}'")
        
        # Sub-column mapping
        if pd.notna(imsvt_sub) and pd.notna(ims_sub) and current_main_column_imsvt:
            imsvt_sub_str = str(imsvt_sub).strip()
            ims_sub_str = str(ims_sub).strip()
            
            if imsvt_sub_str and ims_sub_str and mappings:
                mappings[-1]['sub_columns'].append({
                    'IMSVT_SubColumn': imsvt_sub_str,
                    'IMS_SubColumn': ims_sub_str
                })
    
    logger.info(f"Loaded {len(mappings)} main column mappings")
    return mappings


def analyze_file_structure(file_path: str):
    """
    Analyze the structure of an Excel file
    """
    logger.info(f"\nAnalyzing {file_path}...")
    
    xl = pd.ExcelFile(file_path)
    logger.info(f"  Sheets: {xl.sheet_names}")
    
    for sheet in xl.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet)
        logger.info(f"  Sheet '{sheet}': {df.shape[0]} rows x {df.shape[1]} columns")
        logger.info(f"    Columns: {list(df.columns)[:10]}...")  # First 10 columns
    
    return xl.sheet_names


def compare_with_mapping(file_path: str, mappings: list, sheet_name: str = None):
    """
    Compare file columns against the mapping
    
    Args:
        file_path: Path to Excel file
        mappings: List of column mappings
        sheet_name: Sheet name to use (if None, use first sheet)
    
    Returns:
        DataFrame with comparison results
    """
    logger.info(f"\n{'='*80}")
    logger.info(f"COMPARING FILE: {file_path}")
    logger.info(f"{'='*80}")
    
    # Load file
    if sheet_name:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    else:
        xl = pd.ExcelFile(file_path)
        sheet_name = xl.sheet_names[0]
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    logger.info(f"Using sheet: {sheet_name}")
    logger.info(f"Shape: {df.shape}")
    logger.info(f"Columns: {list(df.columns)}")
    
    # Check which mapped columns exist in the file
    results = []
    
    for mapping in mappings:
        imsvt_col = mapping['IMSVT_MainColumn']
        ims_col = mapping['IMS_MainColumn']
        
        # Check if columns exist in file
        imsvt_exists = imsvt_col in df.columns
        ims_exists = ims_col in df.columns
        
        # Try to match with variations
        if not imsvt_exists:
            for col in df.columns:
                if col and imsvt_col and str(col).strip().lower() == str(imsvt_col).strip().lower():
                    imsvt_exists = True
                    imsvt_col = col
                    break
        
        if not ims_exists:
            for col in df.columns:
                if col and ims_col and str(col).strip().lower() == str(ims_col).strip().lower():
                    ims_exists = True
                    ims_col = col
                    break
        
        result = {
            'IMSVT_Column': mapping['IMSVT_MainColumn'],
            'IMS_Column': mapping['IMS_MainColumn'],
            'IMSVT_Found': imsvt_exists,
            'IMS_Found': ims_exists,
            'Status': 'Both Found' if (imsvt_exists and ims_exists) else 
                     'IMSVT Only' if imsvt_exists else 
                     'IMS Only' if ims_exists else 
                     'Neither Found'
        }
        
        # Get values if columns exist
        if imsvt_exists:
            result['IMSVT_SampleValue'] = df[imsvt_col].iloc[0] if len(df) > 0 else None
        if ims_exists:
            result['IMS_SampleValue'] = df[ims_col].iloc[0] if len(df) > 0 else None
        
        results.append(result)
    
    df_results = pd.DataFrame(results)
    
    # Summary
    logger.info(f"\n{'='*80}")
    logger.info("COLUMN MAPPING SUMMARY")
    logger.info(f"{'='*80}")
    logger.info(f"Total mapped columns: {len(mappings)}")
    logger.info(f"Both columns found: {len(df_results[df_results['Status'] == 'Both Found'])}")
    logger.info(f"Only IMSVT found: {len(df_results[df_results['Status'] == 'IMSVT Only'])}")
    logger.info(f"Only IMS found: {len(df_results[df_results['Status'] == 'IMS Only'])}")
    logger.info(f"Neither found: {len(df_results[df_results['Status'] == 'Neither Found'])}")
    
    return df_results


def main():
    """
    Main execution
    """
    print("\n")
    print("="*80)
    print("IMS VT COMPARISON TOOL WITH COLUMN MAPPING")
    print("="*80)
    print("\n")
    
    # Step 1: Load mappings
    try:
        mappings = load_column_mappings("IMS VT Automation Mappings.xlsx")
    except Exception as e:
        logger.error(f"Failed to load mappings: {e}")
        return
    
    # Step 2: Analyze IMSVT file
    try:
        imsvt_sheets = analyze_file_structure("IMSVT.xlsx")
    except Exception as e:
        logger.error(f"Failed to analyze IMSVT.xlsx: {e}")
        logger.info("\nPlease ensure IMSVT.xlsx exists in the current directory")
        return
    
    # Step 3: Compare file with mappings
    try:
        results = compare_with_mapping("IMSVT.xlsx", mappings)
        
        # Save results
        output_file = f"Column_Mapping_Analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        results.to_excel(output_file, index=False)
        logger.info(f"\n✓ Analysis saved to: {output_file}")
        
        # Show columns that need attention
        print("\n" + "="*80)
        print("COLUMNS NEEDING ATTENTION")
        print("="*80)
        
        needs_attention = results[results['Status'] != 'Both Found']
        if len(needs_attention) > 0:
            print(f"\nFound {len(needs_attention)} columns that don't match perfectly:")
            for _, row in needs_attention.iterrows():
                print(f"\n  {row['Status']}:")
                print(f"    IMSVT Column: {row['IMSVT_Column']}")
                print(f"    IMS Column: {row['IMS_Column']}")
        else:
            print("\n✓ All mapped columns found in the file!")
        
        print("\n" + "="*80)
        print("NEXT STEPS")
        print("="*80)
        print("\nTo perform value comparison:")
        print("1. Identify which columns to compare (all columns with 'Both Found' status)")
        print("2. Use the excel_compare_agent.py with these column mappings")
        print("3. Or provide the second file (IMS Managed Security KDA Report) for comparison")
        
    except Exception as e:
        logger.error(f"Failed to analyze file: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
