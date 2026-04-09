"""
Example Usage: Excel Compare Agent
===================================
This script demonstrates the exact example from the user requirement.
"""

from excel_compare_agent import ExcelCompareAgent

def example_1_natural_language():
    """Example 1: Using natural language requirement (as specified by user)"""
    
    print("="*80)
    print("EXAMPLE 1: Natural Language Requirement")
    print("="*80)
    
    # The exact requirement from the user specification
    requirement = """
    Compare Mar 28th_Release.xlsx vs 21st Feb_Release.xlsx.
    Use sheet "Release" in both.
    Match rows using US_ID.
    Compare fields: PT Status, In sprint test case count, If insprint YES - % of completion.
    Treat text as case-insensitive and trimmed.
    Treat numeric values with tolerance 0.0 (exact).
    Output full report + a new Excel with compared values and results.
    """
    
    # For this example, we need to specify the full path since files are in subdirectory
    requirement = """
    Compare GHC files/Daily status report/Mar 28th_Release.xlsx 
    vs GHC files/Daily status report/21st Feb_Release.xlsx.
    Use sheet "Release" in both.
    Match rows using US_ID.
    Compare fields: PT Status, In sprint test case count, If insprint YES - % of completion.
    Treat text as case-insensitive and trimmed.
    Treat numeric values with tolerance 0.0 (exact).
    Output full report + a new Excel with compared values and results.
    """
    
    # Create and run the agent
    agent = ExcelCompareAgent(requirement_text=requirement)
    output_file = agent.run()
    
    print(f"\n✓ Success! Report saved to: {output_file}")
    print("\nThe report contains 5 sheets:")
    print("  1. Summary - Overall statistics and match rates")
    print("  2. Compared_Data - Row-by-row comparison with results")
    print("  3. Unmatched_In_A - Records only in File B")
    print("  4. Unmatched_In_B - Records only in File A")
    print("  5. Config - Configuration used for this comparison")
    
    return output_file


def example_2_programmatic():
    """Example 2: Using programmatic configuration for more control"""
    
    print("\n")
    print("="*80)
    print("EXAMPLE 2: Programmatic Configuration")
    print("="*80)
    
    from excel_compare_agent import ComparisonConfig, ComparisonRule
    
    # Define comparison rules
    rules = ComparisonRule(
        text_mode="case_insensitive_trimmed",
        numeric_tolerance=0.0,  # exact match for numeric
        treat_blank_as_zero=False
    )
    
    # Define configuration
    config = ComparisonConfig(
        fileA=r"GHC files\Daily status report\21st Feb_Release.xlsx",
        fileB=r"GHC files\Daily status report\Mar 28th_Release.xlsx",
        sheetA="Release",
        sheetB="Release",
        keyColumns=["US_ID"],
        compareColumns=[
            "PT Status",
            "In sprint test case count",
            "If insprint YES - % of completion"
        ],
        rules=rules,
        outputPath="Comparison_Report_Programmatic.xlsx",
        highlightDifferences=True,
        includeAllRows=True
    )
    
    # Create and run the agent
    agent = ExcelCompareAgent(config=config)
    output_file = agent.run()
    
    print(f"\n✓ Success! Report saved to: {output_file}")
    
    return output_file


def example_3_minimal():
    """Example 3: Minimal requirement with auto-detection"""
    
    print("\n")
    print("="*80)
    print("EXAMPLE 3: Minimal Requirement (Auto-Detection)")
    print("="*80)
    
    # Minimal requirement - agent will auto-detect key and compare columns
    requirement = """
    Compare GHC files/Daily status report/21st Feb_Release.xlsx 
    with GHC files/Daily status report/Mar 28th_Release.xlsx.
    Use sheet Release.
    """
    
    agent = ExcelCompareAgent(requirement_text=requirement)
    output_file = agent.run()
    
    print(f"\n✓ Success! Report saved to: {output_file}")
    print("\nNote: Agent auto-detected:")
    print(f"  - Key column: {agent.config.keyColumns}")
    print(f"  - Compare columns: {len(agent.config.compareColumns)} columns")
    
    return output_file


def example_4_with_tolerance():
    """Example 4: Using numeric tolerance for approximate matching"""
    
    print("\n")
    print("="*80)
    print("EXAMPLE 4: With Numeric Tolerance")
    print("="*80)
    
    from excel_compare_agent import ComparisonConfig, ComparisonRule
    
    # Rules with tolerance
    rules = ComparisonRule(
        text_mode="case_insensitive_trimmed",
        numeric_tolerance=0.5,  # Allow difference up to 0.5
        numeric_tolerance_percent=5.0  # Or 5% difference
    )
    
    config = ComparisonConfig(
        fileA=r"GHC files\Daily status report\21st Feb_Release.xlsx",
        fileB=r"GHC files\Daily status report\Mar 28th_Release.xlsx",
        sheetA="Release",
        sheetB="Release",
        keyColumns=["US_ID"],
        compareColumns=["In sprint test case count", "If insprint YES - % of completion"],
        rules=rules,
        outputPath="Comparison_Report_WithTolerance.xlsx"
    )
    
    agent = ExcelCompareAgent(config=config)
    output_file = agent.run()
    
    print(f"\n✓ Success! Report saved to: {output_file}")
    print("\nNote: Numeric differences within tolerance are marked as 'Tolerance_Match'")
    
    return output_file


def main():
    """Run all examples"""
    
    print("\n")
    print("╔" + "═"*78 + "╗")
    print("║" + " "*15 + "EXCEL COMPARE AGENT - USAGE EXAMPLES" + " "*27 + "║")
    print("╚" + "═"*78 + "╝")
    print("\n")
    
    print("This script demonstrates various ways to use the Excel Compare Agent.")
    print("Each example will generate a separate comparison report.\n")
    
    try:
        # Run examples
        output1 = example_1_natural_language()
        
        # Uncomment to run additional examples:
        # output2 = example_2_programmatic()
        # output3 = example_3_minimal()
        # output4 = example_4_with_tolerance()
        
        print("\n")
        print("="*80)
        print("ALL EXAMPLES COMPLETED SUCCESSFULLY!")
        print("="*80)
        print("\nGenerated Reports:")
        print(f"  - {output1}")
        
        print("\nOpen the Excel files to see:")
        print("  ✓ Summary sheet with statistics")
        print("  ✓ Compared_Data sheet with row-by-row comparison")
        print("  ✓ Color-coded differences (Red=Mismatch, Green=Match)")
        print("  ✓ Unmatched records from both files")
        print("  ✓ Configuration details")
        
    except FileNotFoundError as e:
        print(f"\n✗ Error: {e}")
        print("\nMake sure the Excel files exist:")
        print("  - GHC files/Daily status report/21st Feb_Release.xlsx")
        print("  - GHC files/Daily status report/Mar 28th_Release.xlsx")
        print("\nRun 'python generate_sample_data.py' to create sample files.")
        
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
