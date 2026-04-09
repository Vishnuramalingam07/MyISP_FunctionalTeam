"""
Generate Individual Scrum Trend Dashboard
Shows each scrum's quality journey across 3 releases
"""

import json
import os
import subprocess
import sys
from datetime import datetime

print("="*80)
print("GENERATING SCRUM TREND DASHBOARD")
print("="*80)

# Load the trend analysis data
base_path = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics\analysis'
json_file = f'{base_path}\\scrum_trend_data.json'
analysis_script = f'{base_path}\\generate_scrum_trend_analysis.py'

# Always run the prerequisite analysis script to ensure fresh data
print("\n📋 Running trend analysis script to generate fresh data...")
print("="*80)

try:
    result = subprocess.run(
        [sys.executable, analysis_script],
        check=True,
        capture_output=False
    )
    print("="*80)
    print("✅ Trend analysis completed successfully!")
    print("="*80)
except subprocess.CalledProcessError as e:
    print(f"\n❌ ERROR: Failed to run analysis script")
    print(f"   Please check the script: {analysis_script}")
    exit(1)
except FileNotFoundError:
    print(f"\n❌ ERROR: Analysis script not found!")
    print(f"   Expected location: {analysis_script}")
    exit(1)

print("\n✓ Loading trend analysis data...")
with open(json_file, 'r') as f:
    data = json.load(f)

print(f"  - Found data for {data['summary_stats']['Total_Scrums']} scrums")
print(f"  - Generated on: {data['generation_date']}")

# Generate HTML Dashboard
print("\n✓ Generating interactive HTML dashboard...")

html_content = '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Scrum Trend Analysis - Individual Performance Journey</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 15px;
            color: #333;
        }

        .container {
            max-width: 1800px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.4);
            overflow: hidden;
        }

        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 35px;
            text-align: center;
        }

        .header h1 {
            font-size: 2.8em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }

        .header .subtitle {
            font-size: 1.2em;
            opacity: 0.95;
            margin-top: 5px;
        }

        .header .timestamp {
            font-size: 0.9em;
            margin-top: 15px;
            opacity: 0.8;
        }

        .summary-cards {
            display: grid;
            grid-template-columns: repeat(6, 1fr);
            gap: 15px;
            padding: 25px 30px;
            background: #f8f9fa;
        }

        .summary-card {
            background: white;
            padding: 18px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            text-align: center;
            border-left: 5px solid;
            transition: transform 0.3s ease;
        }

        .summary-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.15);
        }

        .summary-card.total { border-color: #667eea; }
        .summary-card.excellent { border-color: #28a745; }
        .summary-card.good { border-color: #17a2b8; }
        .summary-card.moderate { border-color: #ffc107; }
        .summary-card.stable { border-color: #6c757d; }
        .summary-card.decline { border-color: #dc3545; }

        .card-title {
            font-size: 0.75em;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            margin-bottom: 8px;
        }

        .card-value {
            font-size: 2em;
            font-weight: 800;
            margin: 8px 0;
        }

        .card-subtitle {
            font-size: 0.85em;
            color: #999;
        }

        /* Tab-level summary cards (smaller) */
        .tab-summary-cards {
            display: grid;
            grid-template-columns: repeat(6, 1fr);
            gap: 10px;
            padding: 15px 20px;
            background: #f8f9fa;
            margin-bottom: 20px;
            border-radius: 8px;
        }

        .tab-summary-card {
            background: white;
            padding: 12px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            text-align: center;
            border-left: 4px solid;
        }

        .tab-summary-card.total { border-color: #667eea; }
        .tab-summary-card.excellent { border-color: #28a745; }
        .tab-summary-card.good { border-color: #17a2b8; }
        .tab-summary-card.moderate { border-color: #ffc107; }
        .tab-summary-card.stable { border-color: #6c757d; }
        .tab-summary-card.decline { border-color: #dc3545; }

        .tab-card-title {
            font-size: 0.65em;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 5px;
        }

        .tab-card-value {
            font-size: 1.5em;
            font-weight: 800;
            margin: 5px 0;
        }

        .tab-card-subtitle {
            font-size: 0.7em;
            color: #999;
        }

        .section {
            padding: 30px;
        }

        .section-title {
            font-size: 1.8em;
            color: #2c3e50;
            margin-bottom: 25px;
            padding-bottom: 15px;
            border-bottom: 3px solid #667eea;
        }

        .scrums-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(700px, 1fr));
            gap: 25px;
        }

        .scrum-card {
            background: white;
            border-radius: 12px;
            box-shadow: 0 2px 15px rgba(0,0,0,0.1);
            padding: 25px;
            border-left: 6px solid;
            transition: all 0.3s ease;
        }

        .scrum-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.15);
        }

        .scrum-card.excellent { border-color: #28a745; background: linear-gradient(to right, #d4edda 0%, white 15%); }
        .scrum-card.good { border-color: #17a2b8; background: linear-gradient(to right, #d1ecf1 0%, white 15%); }
        .scrum-card.moderate { border-color: #ffc107; background: linear-gradient(to right, #fff3cd 0%, white 15%); }
        .scrum-card.stable { border-color: #6c757d; background: linear-gradient(to right, #e9ecef 0%, white 15%); }
        .scrum-card.slight-decline { border-color: #fd7e14; background: linear-gradient(to right, #ffe8d1 0%, white 15%); }
        .scrum-card.decline { border-color: #dc3545; background: linear-gradient(to right, #f8d7da 0%, white 15%); }

        .scrum-name {
            font-size: 1.4em;
            font-weight: 700;
            color: #2c3e50;
            margin-bottom: 10px;
        }

        .scrum-trend {
            display: inline-block;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: 700;
            margin-bottom: 15px;
        }

        .scrum-trend.excellent { background: #28a745; color: white; }
        .scrum-trend.good { background: #17a2b8; color: white; }
        .scrum-trend.moderate { background: #ffc107; color: #333; }
        .scrum-trend.stable { background: #6c757d; color: white; }
        .scrum-trend.slight-decline { background: #fd7e14; color: white; }
        .scrum-trend.decline { background: #dc3545; color: white; }

        .release-comparison {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 12px;
            margin: 20px 0;
        }

        .release-item {
            background: #f8f9fa;
            padding: 12px;
            border-radius: 8px;
            text-align: center;
        }

        .release-label {
            font-size: 0.85em;
            color: #6c757d;
            font-weight: 600;
            margin-bottom: 5px;
        }

        .release-value {
            font-size: 1.3em;
            font-weight: 800;
            color: #2c3e50;
        }

        .release-subtext {
            font-size: 0.75em;
            color: #999;
            margin-top: 3px;
        }

        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 10px;
            margin-top: 15px;
        }

        .metric-item {
            background: #f8f9fa;
            padding: 10px;
            border-radius: 6px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .metric-label {
            font-size: 0.85em;
            color: #6c757d;
            font-weight: 600;
        }

        .metric-value {
            font-size: 1.1em;
            font-weight: 700;
            color: #2c3e50;
        }

        .metric-value.positive {
            color: #28a745;
        }

        .metric-value.negative {
            color: #dc3545;
        }

        .score-badge {
            display: inline-block;
            padding: 8px 16px;
            border-radius: 25px;
            font-weight: 700;
            font-size: 1.1em;
            margin: 10px 5px;
        }

        .score-badge.excellent { background: #28a745; color: white; }
        .score-badge.good { background: #17a2b8; color: white; }
        .score-badge.moderate { background: #ffc107; color: #333; }
        .score-badge.stable { background: #6c757d; color: white; }
        .score-badge.decline { background: #dc3545; color: white; }

        .poc-info {
            font-size: 0.85em;
            color: #6c757d;
            margin-top: 15px;
            padding-top: 15px;
            border-top: 1px solid #dee2e6;
        }

        .chart-mini {
            height: 120px;
            margin-top: 15px;
        }

        .charts-section {
            margin-top: 20px;
            padding-top: 20px;
            border-top: 2px solid #e9ecef;
        }

        .charts-toggle {
            background: linear-gradient(135deg, #667eea, #764ba2);
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 8px;
            font-weight: 600;
            cursor: pointer;
            width: 100%;
            margin-bottom: 15px;
            transition: all 0.3s ease;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .charts-toggle:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
        }

        .charts-container {
            display: none;
            margin-top: 10px;
        }

        .charts-container.show {
            display: block;
        }

        .comparison-table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border: 1px solid #dee2e6;
        }

        .comparison-table th {
            background: linear-gradient(135deg, #667eea, #764ba2);
            color: white;
            padding: 8px;
            text-align: center;
            font-weight: 700;
            font-size: 0.75em;
            border: 1px solid #9fa8da;
        }

        .comparison-table td {
            padding: 8px;
            text-align: center;
            border: 1px solid #e9ecef;
            font-size: 0.7em;
        }

        .comparison-table tr:last-child td {
            border-bottom: none;
        }

        .comparison-table tr:hover {
            background: #f8f9fa;
        }

        .metric-category {
            font-weight: 700;
            color: #2c3e50;
            text-align: left !important;
            background: #f8f9fa;
        }

        .value-cell {
            font-weight: 600;
            color: #495057;
        }

        .value-improved {
            color: #28a745;
            font-weight: 700;
        }

        .value-declined {
            color: #dc3545;
            font-weight: 700;
        }

        .value-stable {
            color: #6c757d;
            font-weight: 700;
        }

        .filter-section {
            background: #f8f9fa;
            padding: 20px 30px;
            border-bottom: 3px solid #667eea;
        }

        .filter-title {
            font-size: 1.1em;
            font-weight: 700;
            color: #2c3e50;
            margin-bottom: 12px;
        }

        .filters-container {
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 12px;
            align-items: end;
        }

        .filter-group {
            display: flex;
            flex-direction: column;
        }

        .filter-label {
            font-size: 0.8em;
            font-weight: 600;
            color: #495057;
            margin-bottom: 6px;
            display: flex;
            align-items: center;
            gap: 4px;
        }

        .filter-select {
            padding: 8px 12px;
            border: 2px solid #dee2e6;
            border-radius: 6px;
            font-size: 0.85em;
            background: white;
            color: #495057;
            transition: all 0.3s ease;
            cursor: pointer;
        }

        .filter-select:enabled:hover {
            border-color: #667eea;
        }

        .filter-select:enabled:focus {
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }

        .filter-select:disabled {
            background: #e9ecef;
            cursor: not-allowed;
            opacity: 0.6;
        }

        .filter-actions {
            display: flex;
            gap: 8px;
        }

        .reset-btn {
            padding: 8px 16px;
            background: linear-gradient(135deg, #dc3545, #c82333);
            color: white;
            border: none;
            border-radius: 6px;
            font-size: 0.85em;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
            box-shadow: 0 2px 8px rgba(220, 53, 69, 0.3);
            white-space: nowrap;
        }

        .reset-btn:hover {
            background: linear-gradient(135deg, #c82333, #bd2130);
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(220, 53, 69, 0.4);
        }

        .reset-btn:active {
            transform: translateY(0);
        }

        .filter-badge {
            display: inline-block;
            background: #667eea;
            color: white;
            padding: 2px 7px;
            border-radius: 10px;
            font-size: 0.7em;
            font-weight: 600;
            margin-left: 4px;
        }

        .tabs {
            display: flex;
            background: #e9ecef;
            border-bottom: 3px solid #667eea;
            padding: 0 30px;
            flex-wrap: wrap;
        }

        .tab {
            padding: 18px 20px;
            cursor: pointer;
            border: none;
            background: transparent;
            font-size: 0.9em;
            font-weight: 600;
            color: #495057;
            transition: all 0.3s ease;
            border-bottom: 3px solid transparent;
            margin-bottom: -3px;
        }

        .tab:hover {
            background: rgba(102, 126, 234, 0.1);
        }

        .tab.active {
            color: #667eea;
            background: white;
            border-bottom-color: #667eea;
        }

        .tab-content {
            display: none;
            padding: 30px;
        }

        .tab-content.active {
            display: block;
        }

        .performance-table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-radius: 8px;
            overflow: hidden;
            margin-top: 20px;
            border: 1px solid #dee2e6;
        }

        .performance-table th {
            background: linear-gradient(135deg, #667eea, #764ba2);
            color: white;
            padding: 12px 8px;
            text-align: left;
            font-weight: 700;
            font-size: 0.8em;
            position: sticky;
            top: 0;
            z-index: 10;
            border: 1px solid #9fa8da;
        }

        .performance-table td {
            padding: 10px 8px;
            border: 1px solid #e9ecef;
            font-size: 0.75em;
            vertical-align: top;
        }

        .performance-table tr:hover {
            background: #f8f9fa;
        }

        .performance-table tr:last-child td {
            border-bottom: none;
        }

        .scrum-name-col {
            font-weight: 700;
            max-width: 180px;
        }

        .category-col {
            font-weight: 700;
            text-align: center;
        }

        .score-col {
            font-weight: 700;
            text-align: center;
        }

        /* Category-based coloring for table cells */
        .scrum-name-col.excellent,
        .category-col.excellent,
        .score-col.excellent {
            background: #28a745;
            color: white;
        }

        .scrum-name-col.good,
        .category-col.good,
        .score-col.good {
            background: #17a2b8;
            color: white;
        }

        .scrum-name-col.moderate,
        .category-col.moderate,
        .score-col.moderate {
            background: #ffc107;
            color: #333;
        }

        .scrum-name-col.stable,
        .category-col.stable,
        .score-col.stable {
            background: #6c757d;
            color: white;
        }

        .scrum-name-col.slight-decline,
        .category-col.slight-decline,
        .score-col.slight-decline {
            background: #fd7e14;
            color: white;
        }

        .scrum-name-col.decline,
        .category-col.decline,
        .score-col.decline {
            background: #dc3545;
            color: white;
        }

        .category-col {
            font-weight: 600;
            padding: 8px 12px;
            border-radius: 6px;
            text-align: center;
        }

        .score-col {
            font-size: 1.1em;
            font-weight: 700;
            text-align: center;
        }

        .comments-col {
            max-width: 350px;
            line-height: 1.6;
            color: #495057;
        }

        .improvement-col {
            max-width: 300px;
            line-height: 1.6;
            color: #495057;
        }

        .comment-item {
            margin-bottom: 5px;
        }

        .improvement-item {
            margin-bottom: 5px;
            padding-left: 15px;
            position: relative;
        }

        .improvement-item:before {
            content: "💡";
            position: absolute;
            left: 0;
        }

        @media print {
            body { background: white; padding: 0; }
            .container { box-shadow: none; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Individual Scrum Trend Analysis</h1>
            <div class="subtitle">Quality Performance Journey Across 3 Releases (Nov 8, Dec 13, Jan 10)</div>
        </div>

        <!-- Summary Cards -->
        <div class="summary-cards">
            <div class="summary-card total">
                <div class="card-title">📊 Total Scrums</div>
                <div class="card-value">''' + str(len(data['all_scrums'])) + '''</div>
                <div class="card-subtitle">Analyzed</div>
            </div>
            <div class="summary-card excellent">
                <div class="card-title">📈 Excellent Improvement</div>
                <div class="card-value" id="excellent-count">''' + str(data['summary_stats']['Excellent_Improvement']) + '''</div>
                <div class="card-subtitle">Score ≥ 80</div>
            </div>
            <div class="summary-card good">
                <div class="card-title">✅ Good Improvement</div>
                <div class="card-value" id="good-count">''' + str(data['summary_stats']['Good_Improvement']) + '''</div>
                <div class="card-subtitle">Score 70-79</div>
            </div>
            <div class="summary-card moderate">
                <div class="card-title">⚠️ Moderate Improvement</div>
                <div class="card-value" id="moderate-count">''' + str(data['summary_stats']['Moderate_Improvement']) + '''</div>
                <div class="card-subtitle">Score 60-69</div>
            </div>
            <div class="summary-card stable">
                <div class="card-title">➡️ Stable</div>
                <div class="card-value" id="stable-count">''' + str(data['summary_stats']['Stable']) + '''</div>
                <div class="card-subtitle">Score 50-59</div>
            </div>
            <div class="summary-card decline">
                <div class="card-title">📉 Declining</div>
                <div class="card-value" id="decline-count">''' + str(data['summary_stats']['Slight_Decline'] + data['summary_stats']['Significant_Decline']) + '''</div>
                <div class="card-subtitle">Score < 50</div>
            </div>
        </div>

        <!-- Filter Section -->
        <div class="filter-section">
            <div class="filter-title">🔍 Filter Scrums by Hierarchy</div>
            <div class="filters-container">
                <div class="filter-group">
                    <label class="filter-label">
                        👤 AD POC
                        <span class="filter-badge" id="ad-count">0</span>
                    </label>
                    <select id="adPocFilter" class="filter-select" onchange="filterByADPOC()">
                        <option value="">All AD POCs</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label class="filter-label">
                        👥 SM POC
                        <span class="filter-badge" id="sm-count">0</span>
                    </label>
                    <select id="smPocFilter" class="filter-select" onchange="filterBySMPOC()" disabled>
                        <option value="">Select AD POC first</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label class="filter-label">
                        📋 Test Manager (M POC)
                        <span class="filter-badge" id="m-count">0</span>
                    </label>
                    <select id="mPocFilter" class="filter-select" onchange="filterByMPOC()" disabled>
                        <option value="">Select SM POC first</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label class="filter-label">
                        🎯 Node Name
                        <span class="filter-badge" id="node-count">0</span>
                    </label>
                    <select id="nodeFilter" class="filter-select" onchange="filterByNode()" disabled>
                        <option value="">Select M POC first</option>
                    </select>
                </div>
                <div class="filter-group">
                    <label class="filter-label">&nbsp;</label>
                    <button class="reset-btn" onclick="resetFilters()">🔄 Reset</button>
                </div>
            </div>
        </div>

        <!-- Tabs -->
        <div class="tabs">
            <button class="tab active" onclick="showTab('ad-score', event)">AD Score Table</button>
            <button class="tab" onclick="showTab('sm-score', event)">SM Score Table</button>
            <button class="tab" onclick="showTab('m-score', event)">M Score Table</button>
            <button class="tab" onclick="showTab('performance', event)">Performance Table</button>
            <button class="tab" onclick="showTab('all', event)">Detailed performance of all scrums</button>
            <button class="tab" onclick="showTab('scoring', event)">Score Calculation</button>
        </div>

        <!-- Tab Contents -->
        
        <!-- AD Score Table Tab -->
        <div id="ad-score" class="tab-content active">
            <div class="section">
                <div class="section-title">AD POC Performance Rankings</div>
                <p style="text-align: center; color: #666; margin-bottom: 20px;">
                    Aggregated Overall Scores for each AD POC based on all scrums they manage
                </p>
                <div id="ad-score-table-container"></div>
            </div>
        </div>

        <!-- SM Score Table Tab -->
        <div id="sm-score" class="tab-content">
            <div class="section">
                <div class="section-title">SM POC Performance Rankings</div>
                <p style="text-align: center; color: #666; margin-bottom: 20px;">
                    Aggregated Overall Scores for each SM POC based on all scrums they manage
                </p>
                <div id="sm-score-table-container"></div>
            </div>
        </div>

        <!-- M Score Table Tab -->
        <div id="m-score" class="tab-content">
            <div class="section">
                <div class="section-title">Test Manager Performance Rankings</div>
                <p style="text-align: center; color: #666; margin-bottom: 20px;">
                    Aggregated Overall Scores for each Test Manager based on all scrums they manage
                </p>
                <div id="m-score-table-container"></div>
            </div>
        </div>

        <div id="performance" class="tab-content">
            <div class="section">
                <div class="section-title">All Scrums - Performance Table</div>
                
                <!-- Tab Summary Cards -->
                <div class="tab-summary-cards">
                    <div class="tab-summary-card total">
                        <div class="tab-card-title">📊 Total Scrums</div>
                        <div class="tab-card-value" id="tab-total-count">''' + str(len(data['all_scrums'])) + '''</div>
                    </div>
                    <div class="tab-summary-card excellent">
                        <div class="tab-card-title">📈 Excellent</div>
                        <div class="tab-card-value" id="tab-excellent-count">''' + str(data['summary_stats']['Excellent_Improvement']) + '''</div>
                    </div>
                    <div class="tab-summary-card good">
                        <div class="tab-card-title">✅ Good</div>
                        <div class="tab-card-value" id="tab-good-count">''' + str(data['summary_stats']['Good_Improvement']) + '''</div>
                    </div>
                    <div class="tab-summary-card moderate">
                        <div class="tab-card-title">⚠️ Moderate</div>
                        <div class="tab-card-value" id="tab-moderate-count">''' + str(data['summary_stats']['Moderate_Improvement']) + '''</div>
                    </div>
                    <div class="tab-summary-card stable">
                        <div class="tab-card-title">➡️ Stable</div>
                        <div class="tab-card-value" id="tab-stable-count">''' + str(data['summary_stats']['Stable']) + '''</div>
                    </div>
                    <div class="tab-summary-card decline">
                        <div class="tab-card-title">🔻 Declining</div>
                        <div class="tab-card-value" id="tab-decline-count">''' + str(data['summary_stats']['Slight_Decline'] + data['summary_stats']['Significant_Decline']) + '''</div>
                    </div>
                </div>
                
                <div id="performance-table-container"></div>
            </div>
        </div>

        <div id="all" class="tab-content">
            <div class="section">
                <div class="section-title">Detailed Performance of All Scrums</div>
                
                <!-- Tab Summary Cards -->
                <div class="tab-summary-cards">
                    <div class="tab-summary-card total">
                        <div class="tab-card-title">📊 Total Scrums</div>
                        <div class="tab-card-value" id="tab-total-count-all">''' + str(len(data['all_scrums'])) + '''</div>
                    </div>
                    <div class="tab-summary-card excellent">
                        <div class="tab-card-title">📈 Excellent</div>
                        <div class="tab-card-value" id="tab-excellent-count-all">''' + str(data['summary_stats']['Excellent_Improvement']) + '''</div>
                    </div>
                    <div class="tab-summary-card good">
                        <div class="tab-card-title">✅ Good</div>
                        <div class="tab-card-value" id="tab-good-count-all">''' + str(data['summary_stats']['Good_Improvement']) + '''</div>
                    </div>
                    <div class="tab-summary-card moderate">
                        <div class="tab-card-title">⚠️ Moderate</div>
                        <div class="tab-card-value" id="tab-moderate-count-all">''' + str(data['summary_stats']['Moderate_Improvement']) + '''</div>
                    </div>
                    <div class="tab-summary-card stable">
                        <div class="tab-card-title">➡️ Stable</div>
                        <div class="tab-card-value" id="tab-stable-count-all">''' + str(data['summary_stats']['Stable']) + '''</div>
                    </div>
                    <div class="tab-summary-card decline">
                        <div class="tab-card-title">🔻 Declining</div>
                        <div class="tab-card-value" id="tab-decline-count-all">''' + str(data['summary_stats']['Slight_Decline'] + data['summary_stats']['Significant_Decline']) + '''</div>
                    </div>
                </div>
                
                <div class="scrums-grid" id="all-scrums"></div>
            </div>
        </div>

        <div id="scoring" class="tab-content">
            <div class="section">
                <div class="section-title">📊 Score Calculation Methodology</div>
                <div style="background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                    <h2 style="color: #667eea; margin-top: 0;">Overall Score Formula (0-100 scale)</h2>
                    <p style="font-size: 1.1em; background: #f8f9fa; padding: 15px; border-radius: 6px; font-weight: 600;">
                        Overall Score = (DD Score × 30%) + (Severity Score × 10%) + (RCA Score × 20%) + (Re open Score × 10%) + (SLA Score × 10%) + (Agent Augmented Score × 10%) + (Story On Time Delivery Score × 10%)
                    </p>

                    <h3 style="color: #667eea; margin-top: 30px;">1️⃣ Defect Density (DD) Score - 30% Weight</h3>
                    <p>Measures improvement in defect density from <strong>Dec 13 → Jan 10</strong>:</p>
                    <table class="comparison-table" style="margin-top: 15px;">
                        <thead>
                            <tr><th>DD Change</th><th>Score</th><th>Meaning</th></tr>
                        </thead>
                        <tbody>
                            <tr><td>< -20%</td><td>100</td><td>Major improvement</td></tr>
                            <tr><td>-10% to -20%</td><td>90</td><td>Significant improvement</td></tr>
                            <tr><td>-5% to -10%</td><td>80</td><td>Good improvement</td></tr>
                            <tr><td>0% to -5%</td><td>70</td><td>Slight improvement</td></tr>
                            <tr><td>0% to +5%</td><td>60</td><td>Stable</td></tr>
                            <tr><td>+5% to +10%</td><td>40</td><td>Slight decline</td></tr>
                            <tr><td>+10% to +20%</td><td>20</td><td>Significant decline</td></tr>
                            <tr><td>> +20%</td><td>0</td><td>Major decline</td></tr>
                        </tbody>
                    </table>

                    <h3 style="color: #667eea; margin-top: 30px;">2️⃣ Severity Score - 10% Weight</h3>
                    <p>Measures change in Critical+High bug rate:</p>
                    <table class="comparison-table" style="margin-top: 15px;">
                        <thead>
                            <tr><th>Critical+High Rate Change</th><th>Score</th></tr>
                        </thead>
                        <tbody>
                            <tr><td>< -10%</td><td>100</td></tr>
                            <tr><td>-5% to -10%</td><td>80</td></tr>
                            <tr><td>0% to -5%</td><td>70</td></tr>
                            <tr><td>0% to +5%</td><td>60</td></tr>
                            <tr><td>+5% to +10%</td><td>40</td></tr>
                            <tr><td>> +10%</td><td>20</td></tr>
                        </tbody>
                    </table>

                    <h3 style="color: #667eea; margin-top: 30px;">3️⃣ RCA Score - 20% Weight</h3>
                    <p>Measures Dev-owned defect percentage (lower is better):</p>
                    <table class="comparison-table" style="margin-top: 15px;">
                        <thead>
                            <tr><th>Dev RCA %</th><th>Score</th><th>Meaning</th></tr>
                        </thead>
                        <tbody>
                            <tr><td>< 20%</td><td>100</td><td>Excellent - few dev-caused bugs</td></tr>
                            <tr><td>20-40%</td><td>80</td><td>Good</td></tr>
                            <tr><td>40-60%</td><td>60</td><td>Moderate</td></tr>
                            <tr><td>60-80%</td><td>40</td><td>Poor</td></tr>
                            <tr><td>> 80%</td><td>20</td><td>Critical - most bugs from dev</td></tr>
                        </tbody>
                    </table>

                    <h3 style="color: #667eea; margin-top: 30px;">4️⃣ Re open Score - 10% Weight</h3>
                    <p>Based on Reopen Rate (Closure Efficiency = 100 - Reopen Rate):</p>
                    <table class="comparison-table" style="margin-top: 15px;">
                        <thead>
                            <tr><th>Closure Efficiency</th><th>Score</th></tr>
                        </thead>
                        <tbody>
                            <tr><td>≥ 95%</td><td>100</td></tr>
                            <tr><td>90-95%</td><td>90</td></tr>
                            <tr><td>85-90%</td><td>80</td></tr>
                            <tr><td>80-85%</td><td>70</td></tr>
                            <tr><td>70-80%</td><td>50</td></tr>
                            <tr><td>< 70%</td><td>30</td></tr>
                        </tbody>
                    </table>

                    <h3 style="color: #667eea; margin-top: 30px;">5️⃣ SLA Compliance Score - 10% Weight</h3>
                    <p>Percentage of bugs closed within SLA:</p>
                    <table class="comparison-table" style="margin-top: 15px;">
                        <thead>
                            <tr><th>SLA Compliance</th><th>Score</th></tr>
                        </thead>
                        <tbody>
                            <tr><td>≥ 95%</td><td>100</td></tr>
                            <tr><td>90-95%</td><td>90</td></tr>
                            <tr><td>85-90%</td><td>80</td></tr>
                            <tr><td>80-85%</td><td>70</td></tr>
                            <tr><td>70-80%</td><td>60</td></tr>
                            <tr><td>< 70%</td><td>40</td></tr>
                        </tbody>
                    </table>

                    <h3 style="color: #667eea; margin-top: 30px;">6️⃣ Agent Augmented Score - 10% Weight</h3>
                    <p>Measures Agent Augmented usage (lower Agent No % is better):</p>
                    <table class="comparison-table" style="margin-top: 15px;">
                        <thead>
                            <tr><th>Agent No %</th><th>Score</th><th>Meaning</th></tr>
                        </thead>
                        <tbody>
                            <tr><td>< 5%</td><td>100</td><td>Excellent - high agent usage</td></tr>
                            <tr><td>5-10%</td><td>90</td><td>Very Good</td></tr>
                            <tr><td>10-15%</td><td>80</td><td>Good</td></tr>
                            <tr><td>15-20%</td><td>70</td><td>Acceptable</td></tr>
                            <tr><td>20-30%</td><td>60</td><td>Needs Improvement</td></tr>
                            <tr><td>30-40%</td><td>50</td><td>Poor</td></tr>
                            <tr><td>40-50%</td><td>40</td><td>Critical</td></tr>
                            <tr><td>> 50%</td><td>20</td><td>Severe Issues</td></tr>
                        </tbody>
                    </table>

                    <h3 style="color: #667eea; margin-top: 30px;">7️⃣ Story On Time Delivery Score - 10% Weight</h3>
                    <p>Measures on-time story delivery (lower Delayed Yes % is better):</p>
                    <table class="comparison-table" style="margin-top: 15px;">
                        <thead>
                            <tr><th>Delayed Yes %</th><th>Score</th><th>Meaning</th></tr>
                        </thead>
                        <tbody>
                            <tr><td>< 5%</td><td>100</td><td>Excellent - on-time delivery</td></tr>
                            <tr><td>5-10%</td><td>90</td><td>Very Good</td></tr>
                            <tr><td>10-15%</td><td>80</td><td>Good</td></tr>
                            <tr><td>15-20%</td><td>70</td><td>Acceptable</td></tr>
                            <tr><td>20-30%</td><td>60</td><td>Needs Improvement</td></tr>
                            <tr><td>30-40%</td><td>50</td><td>Poor</td></tr>
                            <tr><td>40-50%</td><td>40</td><td>Critical</td></tr>
                            <tr><td>> 50%</td><td>20</td><td>Severe Issues</td></tr>
                        </tbody>
                    </table>

                    <h3 style="color: #667eea; margin-top: 30px;">📋 Final Category Assignment</h3>
                    <table class="comparison-table" style="margin-top: 15px;">
                        <thead>
                            <tr><th>Overall Score</th><th>Category</th></tr>
                        </thead>
                        <tbody>
                            <tr><td>≥ 80</td><td class="excellent">EXCELLENT IMPROVEMENT</td></tr>
                            <tr><td>70-79</td><td class="good">GOOD IMPROVEMENT</td></tr>
                            <tr><td>60-69</td><td class="moderate">MODERATE IMPROVEMENT</td></tr>
                            <tr><td>50-59</td><td class="stable">STABLE</td></tr>
                            <tr><td>40-49</td><td class="slight-decline">SLIGHT DECLINE</td></tr>
                            <tr><td>< 40</td><td class="decline">SIGNIFICANT DECLINE</td></tr>
                        </tbody>
                    </table>

                    <h3 style="color: #667eea; margin-top: 30px;">💡 Example Calculation</h3>
                    <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin-top: 15px;">
                        <p><strong>Component Scores:</strong></p>
                        <ul style="line-height: 2;">
                            <li>DD Score: 90 (DD dropped -15% from Dec 13 to Jan 10)</li>
                            <li>Severity Score: 70 (Critical+High rate dropped -3%)</li>
                            <li>RCA Score: 80 (25% dev RCA)</li>
                            <li>Closure Score: 90 (92% closure efficiency)</li>
                            <li>SLA Score: 80 (87% SLA compliance)</li>
                            <li>Agent Augmented Score: 90 (8% Agent No)</li>
                            <li>Story On Time Delivery Score: 80 (12% Delayed Yes)</li>
                        </ul>
                        <p style="font-weight: 600; margin-top: 15px;">
                            Overall Score = (90×0.30) + (70×0.10) + (80×0.20) + (90×0.10) + (80×0.10) + (90×0.10) + (80×0.10)<br>
                            = 27 + 7 + 16 + 9 + 8 + 9 + 8 = <span style="color: #28a745; font-size: 1.2em;">84</span> → EXCELLENT IMPROVEMENT ✨
                        </p>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        const data = ''' + json.dumps(data) + ''';
        let filteredData = [...data.all_scrums];
        
        // Create a map of scrum data by scrum name for quick lookup
        const scrumDataMap = {};
        data.all_scrums.forEach(scrum => {
            // Use same ID format as createScrumCard function
            const scrumKey = `scrum-${scrum.Scrum.replace(/[^a-zA-Z0-9]/g, '-')}`;
            scrumDataMap[scrumKey] = scrum;
            console.log('Mapped scrum:', scrumKey, 'Data:', scrum);
        });
        
        console.log('Total scrums in map:', Object.keys(scrumDataMap).length);

        // Generate performance insights based on scores
        function generateInsights(scrum) {
            const comments = [];
            const improvements = [];
            
            // DD Score analysis (30% weight)
            if (scrum.DD_Score >= 90) {
                comments.push(`✓ Excellent DD trend (${scrum.DD_Change > 0 ? '+' : ''}${scrum.DD_Change}%)`);
            } else if (scrum.DD_Score >= 70) {
                comments.push(`Good DD improvement (${scrum.DD_Change > 0 ? '+' : ''}${scrum.DD_Change}%)`);
            } else if (scrum.DD_Score < 40) {
                comments.push(`⚠ DD increased significantly (${scrum.DD_Change > 0 ? '+' : ''}${scrum.DD_Change}%)`);
                improvements.push('Focus on defect prevention and code quality reviews');
            }
            
            // Severity Score analysis (20% weight)
            if (scrum.Severity_Score < 60) {
                comments.push(`⚠ High critical/high severity rate`);
                improvements.push('Improve test coverage for critical scenarios');
            }
            
            // RCA Score analysis (20% weight)
            if (scrum.Dev_RCA_Pct >= 60) {
                comments.push(`⚠ High dev-owned defects (${scrum.Dev_RCA_Pct}%)`);
                improvements.push('Strengthen unit testing and code reviews');
            } else if (scrum.Dev_RCA_Pct < 40) {
                comments.push(`✓ Good dev ownership (${scrum.Dev_RCA_Pct}% dev RCA)`);
            }
            
            // Closure Score analysis (20% weight)
            if (scrum.Reopen_Rate >= 20) {
                comments.push(`⚠ High reopen rate (${scrum.Reopen_Rate}%)`);
                improvements.push('Improve root cause analysis before closure');
            } else if (scrum.Reopen_Rate < 10) {
                comments.push(`✓ Excellent closure efficiency (${scrum.Closure_Efficiency}%)`);
            }
            
            // SLA Score analysis (10% weight)
            if (scrum.SLA_Compliance < 60) {
                comments.push(`⚠ Low SLA compliance (${scrum.SLA_Compliance}%)`);
                improvements.push('Prioritize timely bug resolution');
            } else if (scrum.SLA_Compliance >= 90) {
                comments.push(`✓ Excellent SLA compliance (${scrum.SLA_Compliance}%)`);
            }
            
            // Overall trend
            if (scrum.Overall_Score >= 80) {
                comments.push(`🌟 Outstanding overall performance`);
            } else if (scrum.Overall_Score < 40) {
                improvements.push('Comprehensive quality process review needed');
            }
            
            return {
                comments: comments.length > 0 ? comments.join('<br>') : 'Moderate performance across all metrics',
                improvements: improvements.length > 0 ? improvements.map(i => `<div class="improvement-item">${i}</div>`).join('') : '<div class="improvement-item">Continue current practices</div>'
            };
        }

        // Generate Performance Table
        function generatePerformanceTable() {
            const sortedScrums = [...filteredData].sort((a, b) => b.Overall_Score - a.Overall_Score);
            
            let tableHTML = `
                <table class="performance-table">
                    <thead>
                        <tr>
                            <th>Scrum Name</th>
                            <th style="text-align: center;">Category</th>
                            <th style="text-align: center;">Overall Score</th>
                            <th style="text-align: center;">Defect Density Score</th>
                            <th style="text-align: center;">Severity Score</th>
                            <th style="text-align: center;">RCA Score</th>
                            <th style="text-align: center;">Re open Score</th>
                            <th style="text-align: center;">SLA Score</th>
                            <th style="text-align: center;">Agent Augmented Score</th>
                            <th style="text-align: center;">Story On Time Delivery Score</th>
                            <th>AD POC</th>
                            <th>SM POC</th>
                            <th>M POC</th>
                        </tr>
                    </thead>
                    <tbody>
            `;
            
            if (sortedScrums.length === 0) {
                tableHTML += `
                    <tr>
                        <td colspan="13" style="text-align: center; padding: 50px; color: #999;">No scrums match the selected filters</td>
                    </tr>
                `;
            } else {
                sortedScrums.forEach(scrum => {
                    const categoryClass = scrum.Trend_Class;
                    
                    tableHTML += `
                        <tr>
                            <td class="scrum-name-col ${categoryClass}">${scrum.Scrum}</td>
                            <td class="category-col ${categoryClass}">${scrum.Trend}</td>
                            <td class="score-col ${categoryClass}">${scrum.Overall_Score}</td>
                            <td style="text-align: center;">${scrum.DD_Score}</td>
                            <td style="text-align: center;">${scrum.Severity_Score}</td>
                            <td style="text-align: center;">${scrum.RCA_Score}</td>
                            <td style="text-align: center;">${scrum.Closure_Score}</td>
                            <td style="text-align: center;">${scrum.SLA_Score}</td>
                            <td style="text-align: center;">${scrum.Agent_Augmented_Score}</td>
                            <td style="text-align: center;">${scrum.Story_OnTime_Score}</td>
                            <td>${scrum.AD_POC}</td>
                            <td>${scrum.SM_POC}</td>
                            <td>${scrum.Test_Manager}</td>
                        </tr>
                    `;
                });
            }
            
            tableHTML += `
                    </tbody>
                </table>
            `;
            
            document.getElementById('performance-table-container').innerHTML = tableHTML;
        }

        // Generate AD POC Score Table
        function generateADScoreTable() {
            // Group scrums by AD POC and calculate average scores
            const pocMap = {};
            
            data.all_scrums.forEach(scrum => {
                const poc = scrum.AD_POC;
                if (!pocMap[poc]) {
                    pocMap[poc] = {
                        poc: poc,
                        totalScore: 0,
                        totalDD: 0,
                        totalSeverity: 0,
                        totalRCA: 0,
                        totalClosure: 0,
                        totalSLA: 0,
                        totalAgent: 0,
                        totalStoryOnTime: 0,
                        count: 0
                    };
                }
                pocMap[poc].totalScore += scrum.Overall_Score;
                pocMap[poc].totalDD += scrum.DD_Score;
                pocMap[poc].totalSeverity += scrum.Severity_Score;
                pocMap[poc].totalRCA += scrum.RCA_Score;
                pocMap[poc].totalClosure += scrum.Closure_Score;
                pocMap[poc].totalSLA += scrum.SLA_Score;
                pocMap[poc].totalAgent += scrum.Agent_Augmented_Score;
                pocMap[poc].totalStoryOnTime += scrum.Story_OnTime_Score;
                pocMap[poc].count++;
            });
            
            // Calculate averages and sort by score
            const pocScores = Object.values(pocMap).map(item => ({
                poc: item.poc,
                avgScore: Math.round(item.totalScore / item.count),
                avgDD: Math.round(item.totalDD / item.count),
                avgSeverity: Math.round(item.totalSeverity / item.count),
                avgRCA: Math.round(item.totalRCA / item.count),
                avgClosure: Math.round(item.totalClosure / item.count),
                avgSLA: Math.round(item.totalSLA / item.count),
                avgAgent: Math.round(item.totalAgent / item.count),
                avgStoryOnTime: Math.round(item.totalStoryOnTime / item.count),
                scrumCount: item.count
            })).sort((a, b) => b.avgScore - a.avgScore);
            
            // Generate table HTML
            let tableHTML = `
                <table class="performance-table">
                    <thead>
                        <tr>
                            <th>Rank</th>
                            <th>AD POC</th>
                            <th style="text-align: center;">Average Overall Score</th>
                            <th style="text-align: center;">Number of Scrums</th>
                            <th style="text-align: center;">Defect Density Score</th>
                            <th style="text-align: center;">Severity Score</th>
                            <th style="text-align: center;">RCA Score</th>
                            <th style="text-align: center;">Re open Score</th>
                            <th style="text-align: center;">SLA Score</th>
                            <th style="text-align: center;">Agent Augmented Score</th>
                            <th style="text-align: center;">Story On Time Delivery Score</th>
                        </tr>
                    </thead>
                    <tbody>
            `;            
            pocScores.forEach((item, index) => {
                tableHTML += `
                    <tr>
                        <td style="text-align: center; font-weight: bold;">${index + 1}</td>
                        <td style="font-weight: bold;">${item.poc}</td>
                        <td style="text-align: center; font-size: 1.1em; font-weight: bold;">${item.avgScore}</td>
                        <td style="text-align: center;">${item.scrumCount}</td>
                        <td style="text-align: center;">${item.avgDD}</td>
                        <td style="text-align: center;">${item.avgSeverity}</td>
                        <td style="text-align: center;">${item.avgRCA}</td>
                        <td style="text-align: center;">${item.avgClosure}</td>
                        <td style="text-align: center;">${item.avgSLA}</td>
                        <td style="text-align: center;">${item.avgAgent}</td>
                        <td style="text-align: center;">${item.avgStoryOnTime}</td>
                    </tr>
                `;
            });
            
            tableHTML += `
                    </tbody>
                </table>
            `;
            
            document.getElementById('ad-score-table-container').innerHTML = tableHTML;
        }

        // Generate SM POC Score Table
        function generateSMScoreTable() {
            // Group scrums by SM POC and calculate average scores
            const pocMap = {};
            
            data.all_scrums.forEach(scrum => {
                const poc = scrum.SM_POC;
                if (!pocMap[poc]) {
                    pocMap[poc] = {
                        poc: poc,
                        totalScore: 0,
                        totalDD: 0,
                        totalSeverity: 0,
                        totalRCA: 0,
                        totalClosure: 0,
                        totalSLA: 0,
                        totalAgent: 0,
                        totalStoryOnTime: 0,
                        count: 0
                    };
                }
                pocMap[poc].totalScore += scrum.Overall_Score;
                pocMap[poc].totalDD += scrum.DD_Score;
                pocMap[poc].totalSeverity += scrum.Severity_Score;
                pocMap[poc].totalRCA += scrum.RCA_Score;
                pocMap[poc].totalClosure += scrum.Closure_Score;
                pocMap[poc].totalSLA += scrum.SLA_Score;
                pocMap[poc].totalAgent += scrum.Agent_Augmented_Score;
                pocMap[poc].totalStoryOnTime += scrum.Story_OnTime_Score;
                pocMap[poc].count++;
            });
            
            // Calculate averages and sort by score
            const pocScores = Object.values(pocMap).map(item => ({
                poc: item.poc,
                avgScore: Math.round(item.totalScore / item.count),
                avgDD: Math.round(item.totalDD / item.count),
                avgSeverity: Math.round(item.totalSeverity / item.count),
                avgRCA: Math.round(item.totalRCA / item.count),
                avgClosure: Math.round(item.totalClosure / item.count),
                avgSLA: Math.round(item.totalSLA / item.count),
                avgAgent: Math.round(item.totalAgent / item.count),
                avgStoryOnTime: Math.round(item.totalStoryOnTime / item.count),
                scrumCount: item.count
            })).sort((a, b) => b.avgScore - a.avgScore);
            
            // Generate table HTML
            let tableHTML = `
                <table class="performance-table">
                    <thead>
                        <tr>
                            <th>Rank</th>
                            <th>SM POC</th>
                            <th style="text-align: center;">Average Overall Score</th>
                            <th style="text-align: center;">Number of Scrums</th>
                            <th style="text-align: center;">Defect Density Score</th>
                            <th style="text-align: center;">Severity Score</th>
                            <th style="text-align: center;">RCA Score</th>
                            <th style="text-align: center;">Re open Score</th>
                            <th style="text-align: center;">SLA Score</th>
                            <th style="text-align: center;">Agent Augmented Score</th>
                            <th style="text-align: center;">Story On Time Delivery Score</th>
                        </tr>
                    </thead>
                    <tbody>
            `;            
            pocScores.forEach((item, index) => {
                tableHTML += `
                    <tr>
                        <td style="text-align: center; font-weight: bold;">${index + 1}</td>
                        <td style="font-weight: bold;">${item.poc}</td>
                        <td style="text-align: center; font-size: 1.1em; font-weight: bold;">${item.avgScore}</td>
                        <td style="text-align: center;">${item.scrumCount}</td>
                        <td style="text-align: center;">${item.avgDD}</td>
                        <td style="text-align: center;">${item.avgSeverity}</td>
                        <td style="text-align: center;">${item.avgRCA}</td>
                        <td style="text-align: center;">${item.avgClosure}</td>
                        <td style="text-align: center;">${item.avgSLA}</td>
                        <td style="text-align: center;">${item.avgAgent}</td>
                        <td style="text-align: center;">${item.avgStoryOnTime}</td>
                    </tr>
                `;
            });
            
            tableHTML += `
                    </tbody>
                </table>
            `;
            
            document.getElementById('sm-score-table-container').innerHTML = tableHTML;
        }

        // Generate M POC Score Table
        function generateMScoreTable() {
            // Group scrums by Test Manager and calculate average scores
            const pocMap = {};
            
            data.all_scrums.forEach(scrum => {
                const poc = scrum.Test_Manager;
                if (!pocMap[poc]) {
                    pocMap[poc] = {
                        poc: poc,
                        totalScore: 0,
                        totalDD: 0,
                        totalSeverity: 0,
                        totalRCA: 0,
                        totalClosure: 0,
                        totalSLA: 0,
                        totalAgent: 0,
                        totalStoryOnTime: 0,
                        count: 0
                    };
                }
                pocMap[poc].totalScore += scrum.Overall_Score;
                pocMap[poc].totalDD += scrum.DD_Score;
                pocMap[poc].totalSeverity += scrum.Severity_Score;
                pocMap[poc].totalRCA += scrum.RCA_Score;
                pocMap[poc].totalClosure += scrum.Closure_Score;
                pocMap[poc].totalSLA += scrum.SLA_Score;
                pocMap[poc].totalAgent += scrum.Agent_Augmented_Score;
                pocMap[poc].totalStoryOnTime += scrum.Story_OnTime_Score;
                pocMap[poc].count++;
            });
            
            // Calculate averages and sort by score
            const pocScores = Object.values(pocMap).map(item => ({
                poc: item.poc,
                avgScore: Math.round(item.totalScore / item.count),
                avgDD: Math.round(item.totalDD / item.count),
                avgSeverity: Math.round(item.totalSeverity / item.count),
                avgRCA: Math.round(item.totalRCA / item.count),
                avgClosure: Math.round(item.totalClosure / item.count),
                avgSLA: Math.round(item.totalSLA / item.count),
                avgAgent: Math.round(item.totalAgent / item.count),
                avgStoryOnTime: Math.round(item.totalStoryOnTime / item.count),
                scrumCount: item.count
            })).sort((a, b) => b.avgScore - a.avgScore);
            
            // Generate table HTML
            let tableHTML = `
                <table class="performance-table">
                    <thead>
                        <tr>
                            <th>Rank</th>
                            <th>Test Manager</th>
                            <th style="text-align: center;">Average Overall Score</th>
                            <th style="text-align: center;">Number of Scrums</th>
                            <th style="text-align: center;">Defect Density Score</th>
                            <th style="text-align: center;">Severity Score</th>
                            <th style="text-align: center;">RCA Score</th>
                            <th style="text-align: center;">Re open Score</th>
                            <th style="text-align: center;">SLA Score</th>
                            <th style="text-align: center;">Agent Augmented Score</th>
                            <th style="text-align: center;">Story On Time Delivery Score</th>
                        </tr>
                    </thead>
                    <tbody>
            `;            
            pocScores.forEach((item, index) => {
                tableHTML += `
                    <tr>
                        <td style="text-align: center; font-weight: bold;">${index + 1}</td>
                        <td style="font-weight: bold;">${item.poc}</td>
                        <td style="text-align: center; font-size: 1.1em; font-weight: bold;">${item.avgScore}</td>
                        <td style="text-align: center;">${item.scrumCount}</td>
                        <td style="text-align: center;">${item.avgDD}</td>
                        <td style="text-align: center;">${item.avgSeverity}</td>
                        <td style="text-align: center;">${item.avgRCA}</td>
                        <td style="text-align: center;">${item.avgClosure}</td>
                        <td style="text-align: center;">${item.avgSLA}</td>
                        <td style="text-align: center;">${item.avgAgent}</td>
                        <td style="text-align: center;">${item.avgStoryOnTime}</td>
                    </tr>
                `;
            });
            
            tableHTML += `
                    </tbody>
                </table>
            `;
            
            document.getElementById('m-score-table-container').innerHTML = tableHTML;
        }

        // Helper function to get category class based on score
        function getCategoryClass(score) {
            if (score >= 80) return 'excellent-improvement';
            if (score >= 70) return 'good-improvement';
            if (score >= 60) return 'moderate-improvement';
            if (score >= 50) return 'stable';
            if (score >= 40) return 'slight-decline';
            return 'significant-decline';
        }

        // Initialize filters on page load
        function initializeFilters() {
            const adPocs = [...new Set(data.all_scrums.map(s => s.AD_POC))].sort();
            const adSelect = document.getElementById('adPocFilter');
            adSelect.innerHTML = '<option value="">All AD POCs</option>' + 
                adPocs.map(poc => `<option value="${poc}">${poc}</option>`).join('');
            
            updateFilterBadges();
            updateSummaryCards();
        }

        // Filter by AD POC (Level 1)
        function filterByADPOC() {
            const adValue = document.getElementById('adPocFilter').value;
            const smSelect = document.getElementById('smPocFilter');
            const mSelect = document.getElementById('mPocFilter');
            const nodeSelect = document.getElementById('nodeFilter');
            
            if (!adValue) {
                // Reset all cascading filters
                smSelect.disabled = true;
                smSelect.innerHTML = '<option value="">Select AD POC first</option>';
                mSelect.disabled = true;
                mSelect.innerHTML = '<option value="">Select SM POC first</option>';
                nodeSelect.disabled = true;
                nodeSelect.innerHTML = '<option value="">Select M POC first</option>';
                filteredData = [...data.all_scrums];
            } else {
                // Filter by AD POC
                filteredData = data.all_scrums.filter(s => s.AD_POC === adValue);
                
                // Populate SM POC dropdown
                const smPocs = [...new Set(filteredData.map(s => s.SM_POC))].sort();
                smSelect.disabled = false;
                smSelect.innerHTML = '<option value="">All SM POCs</option>' + 
                    smPocs.map(poc => `<option value="${poc}">${poc}</option>`).join('');
                
                // Reset downstream filters
                mSelect.disabled = true;
                mSelect.innerHTML = '<option value="">Select SM POC first</option>';
                nodeSelect.disabled = true;
                nodeSelect.innerHTML = '<option value="">Select M POC first</option>';
            }
            
            updateFilterBadges();
            updateDisplay();
        }

        // Filter by SM POC (Level 2)
        function filterBySMPOC() {
            const adValue = document.getElementById('adPocFilter').value;
            const smValue = document.getElementById('smPocFilter').value;
            const mSelect = document.getElementById('mPocFilter');
            const nodeSelect = document.getElementById('nodeFilter');
            
            if (!smValue) {
                // Reset to AD POC level
                filteredData = data.all_scrums.filter(s => s.AD_POC === adValue);
                mSelect.disabled = true;
                mSelect.innerHTML = '<option value="">Select SM POC first</option>';
                nodeSelect.disabled = true;
                nodeSelect.innerHTML = '<option value="">Select M POC first</option>';
            } else {
                // Filter by SM POC
                filteredData = data.all_scrums.filter(s => 
                    s.AD_POC === adValue && s.SM_POC === smValue
                );
                
                // Populate M POC dropdown
                const mPocs = [...new Set(filteredData.map(s => s.Test_Manager))].sort();
                mSelect.disabled = false;
                mSelect.innerHTML = '<option value="">All Test Managers</option>' + 
                    mPocs.map(poc => `<option value="${poc}">${poc}</option>`).join('');
                
                // Reset downstream filter
                nodeSelect.disabled = true;
                nodeSelect.innerHTML = '<option value="">Select M POC first</option>';
            }
            
            updateFilterBadges();
            updateDisplay();
        }

        // Filter by M POC (Level 3)
        function filterByMPOC() {
            const adValue = document.getElementById('adPocFilter').value;
            const smValue = document.getElementById('smPocFilter').value;
            const mValue = document.getElementById('mPocFilter').value;
            const nodeSelect = document.getElementById('nodeFilter');
            
            if (!mValue) {
                // Reset to SM POC level
                filteredData = data.all_scrums.filter(s => 
                    s.AD_POC === adValue && s.SM_POC === smValue
                );
                nodeSelect.disabled = true;
                nodeSelect.innerHTML = '<option value="">Select M POC first</option>';
            } else {
                // Filter by M POC
                filteredData = data.all_scrums.filter(s => 
                    s.AD_POC === adValue && s.SM_POC === smValue && s.Test_Manager === mValue
                );
                
                // Populate Node dropdown
                const nodes = [...new Set(filteredData.map(s => s.Scrum))].sort();
                nodeSelect.disabled = false;
                nodeSelect.innerHTML = '<option value="">All Nodes</option>' + 
                    nodes.map(node => `<option value="${node}">${node}</option>`).join('');
            }
            
            updateFilterBadges();
            updateDisplay();
        }

        // Filter by Node (Level 4)
        function filterByNode() {
            const adValue = document.getElementById('adPocFilter').value;
            const smValue = document.getElementById('smPocFilter').value;
            const mValue = document.getElementById('mPocFilter').value;
            const nodeValue = document.getElementById('nodeFilter').value;
            
            if (!nodeValue) {
                // Reset to M POC level
                filteredData = data.all_scrums.filter(s => 
                    s.AD_POC === adValue && s.SM_POC === smValue && s.Test_Manager === mValue
                );
            } else {
                // Filter by Node
                filteredData = data.all_scrums.filter(s => 
                    s.AD_POC === adValue && s.SM_POC === smValue && 
                    s.Test_Manager === mValue && s.Scrum === nodeValue
                );
            }
            
            updateFilterBadges();
            updateDisplay();
        }

        // Reset all filters
        function resetFilters() {
            document.getElementById('adPocFilter').value = '';
            document.getElementById('smPocFilter').value = '';
            document.getElementById('smPocFilter').disabled = true;
            document.getElementById('smPocFilter').innerHTML = '<option value="">Select AD POC first</option>';
            document.getElementById('mPocFilter').value = '';
            document.getElementById('mPocFilter').disabled = true;
            document.getElementById('mPocFilter').innerHTML = '<option value="">Select SM POC first</option>';
            document.getElementById('nodeFilter').value = '';
            document.getElementById('nodeFilter').disabled = true;
            document.getElementById('nodeFilter').innerHTML = '<option value="">Select M POC first</option>';
            
            filteredData = [...data.all_scrums];
            updateFilterBadges();
            updateDisplay();
        }

        // Update filter count badges
        function updateFilterBadges() {
            const adValue = document.getElementById('adPocFilter').value;
            const smValue = document.getElementById('smPocFilter').value;
            const mValue = document.getElementById('mPocFilter').value;
            const nodeValue = document.getElementById('nodeFilter').value;
            
            // Count unique values at each level based on current filter
            const adPocs = adValue ? [adValue] : [...new Set(data.all_scrums.map(s => s.AD_POC))];
            const smPocs = smValue ? [smValue] : [...new Set(filteredData.map(s => s.SM_POC))];
            const mPocs = mValue ? [mValue] : [...new Set(filteredData.map(s => s.Test_Manager))];
            const nodes = nodeValue ? [nodeValue] : filteredData.length;
            
            document.getElementById('ad-count').textContent = adPocs.length;
            document.getElementById('sm-count').textContent = smPocs.length;
            document.getElementById('m-count').textContent = mPocs.length;
            document.getElementById('node-count').textContent = nodes;
        }

        // Update summary cards based on filtered data
        function updateSummaryCards() {
            const excellent = filteredData.filter(s => s.Trend === 'EXCELLENT IMPROVEMENT').length;
            const good = filteredData.filter(s => s.Trend === 'GOOD IMPROVEMENT').length;
            const moderate = filteredData.filter(s => s.Trend === 'MODERATE IMPROVEMENT').length;
            const stable = filteredData.filter(s => s.Trend === 'STABLE').length;
            const decline = filteredData.filter(s => 
                s.Trend === 'SLIGHT DECLINE' || s.Trend === 'SIGNIFICANT DECLINE'
            ).length;
            const total = filteredData.length;
            
            // Update tab-level summary cards only (top cards remain static)
            document.getElementById('tab-total-count').textContent = total;
            document.getElementById('tab-excellent-count').textContent = excellent;
            document.getElementById('tab-good-count').textContent = good;
            document.getElementById('tab-moderate-count').textContent = moderate;
            document.getElementById('tab-stable-count').textContent = stable;
            document.getElementById('tab-decline-count').textContent = decline;
            
            document.getElementById('tab-total-count-all').textContent = total;
            document.getElementById('tab-excellent-count-all').textContent = excellent;
            document.getElementById('tab-good-count-all').textContent = good;
            document.getElementById('tab-moderate-count-all').textContent = moderate;
            document.getElementById('tab-stable-count-all').textContent = stable;
            document.getElementById('tab-decline-count-all').textContent = decline;
        }

        // Update display with filtered data
        function updateDisplay() {
            populateScrums();
            updateSummaryCards();
            generatePerformanceTable();
        }

        function showTab(tabName, event) {
            const tabs = document.querySelectorAll('.tab');
            const contents = document.querySelectorAll('.tab-content');
            
            tabs.forEach(t => t.classList.remove('active'));
            contents.forEach(c => c.classList.remove('active'));
            
            if (event && event.target) {
                event.target.classList.add('active');
            } else {
                // Fallback: find and activate the clicked tab by tabName
                document.querySelector(`[onclick*="showTab('${tabName}')"]`).classList.add('active');
            }
            document.getElementById(tabName).classList.add('active');
        }

        function createScrumCard(scrum, index) {
            const releaseData = scrum.Release_Data;
            const first = releaseData[0];
            const latest = releaseData[releaseData.length - 1];
            // Use scrum name for unique ID to avoid index mismatch issues
            const scrumId = `scrum-${scrum.Scrum.replace(/[^a-zA-Z0-9]/g, '-')}`;
            
            return `
                <div class="scrum-card ${scrum.Trend_Class}">
                    <div class="scrum-name">${scrum.Scrum}</div>
                    <span class="scrum-trend ${scrum.Trend_Class}">${scrum.Trend}</span>
                    <span class="score-badge ${scrum.Trend_Class}">Score: ${scrum.Overall_Score}/100</span>
                    
                    <div class="poc-info">
                        <strong>AD POC:</strong> ${scrum.AD_POC} | 
                        <strong>SM POC:</strong> ${scrum.SM_POC} | 
                        <strong>Test Manager:</strong> ${scrum.Test_Manager}
                    </div>
                    
                    <!-- Charts Section -->
                    <div class="charts-section">
                        <h4 style="color: #667eea; margin: 15px 0 10px 0; font-size: 1em;">📊 Detailed Comparison Table</h4>
                        <div class="charts-container show">
                            <table class="comparison-table">
                                <thead>
                                    <tr>
                                        <th>Metric</th>
                                        ${releaseData.map(r => `<th>${r.Release}</th>`).join('')}
                                        <th>Change</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td class="metric-category">DD% (Defect Density)</td>
                                        ${releaseData.map((r, idx) => {
                                            const val = r.Defect_Density;
                                            let cssClass = 'value-cell';
                                            if (idx > 0) {
                                                const prev = releaseData[idx - 1].Defect_Density;
                                                cssClass = val < prev ? 'value-improved' : val > prev ? 'value-declined' : 'value-stable';
                                            }
                                            return `<td class="${cssClass}">${val}%</td>`;
                                        }).join('')}
                                        <td class="${scrum.DD_Change < 0 ? 'value-improved' : scrum.DD_Change > 0 ? 'value-declined' : 'value-stable'}">
                                            ${scrum.DD_Change > 0 ? '+' : ''}${scrum.DD_Change}%
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="metric-category">Story Points</td>
                                        ${releaseData.map((r, idx) => {
                                            const val = r.Story_Points;
                                            let cssClass = 'value-cell';
                                            if (idx > 0) {
                                                const prev = releaseData[idx - 1].Story_Points;
                                                cssClass = val > prev ? 'value-improved' : val < prev ? 'value-declined' : 'value-stable';
                                            }
                                            return `<td class="${cssClass}">${val}</td>`;
                                        }).join('')}
                                        <td class="${latest.Story_Points > first.Story_Points ? 'value-improved' : latest.Story_Points < first.Story_Points ? 'value-declined' : 'value-stable'}">
                                            ${latest.Story_Points > first.Story_Points ? '+' : ''}${latest.Story_Points - first.Story_Points}
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="metric-category">Total Stories</td>
                                        ${releaseData.map((r, idx) => {
                                            const val = r.Total_Stories;
                                            let cssClass = 'value-cell';
                                            if (idx > 0) {
                                                const prev = releaseData[idx - 1].Total_Stories;
                                                cssClass = val > prev ? 'value-improved' : val < prev ? 'value-declined' : 'value-stable';
                                            }
                                            return `<td class="${cssClass}">${val}</td>`;
                                        }).join('')}
                                        <td class="${latest.Total_Stories > first.Total_Stories ? 'value-improved' : latest.Total_Stories < first.Total_Stories ? 'value-declined' : 'value-stable'}">
                                            ${latest.Total_Stories > first.Total_Stories ? '+' : ''}${latest.Total_Stories - first.Total_Stories}
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="metric-category">Valid Defects</td>
                                        ${releaseData.map((r, idx) => {
                                            const val = r.Valid_Bugs;
                                            let cssClass = 'value-cell';
                                            if (idx > 0) {
                                                const prev = releaseData[idx - 1].Valid_Bugs;
                                                cssClass = val < prev ? 'value-improved' : val > prev ? 'value-declined' : 'value-stable';
                                            }
                                            return `<td class="${cssClass}">${val}</td>`;
                                        }).join('')}
                                        <td class="${latest.Valid_Bugs < first.Valid_Bugs ? 'value-improved' : latest.Valid_Bugs > first.Valid_Bugs ? 'value-declined' : 'value-stable'}">
                                            ${latest.Valid_Bugs > first.Valid_Bugs ? '+' : ''}${latest.Valid_Bugs - first.Valid_Bugs}
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="metric-category">Critical+High Bugs</td>
                                        ${releaseData.map((r, idx) => {
                                            const val = r.Critical + r.High;
                                            let cssClass = 'value-cell';
                                            if (idx > 0) {
                                                const prev = releaseData[idx - 1].Critical + releaseData[idx - 1].High;
                                                cssClass = val < prev ? 'value-improved' : val > prev ? 'value-declined' : 'value-stable';
                                            }
                                            return `<td class="${cssClass}">${val}</td>`;
                                        }).join('')}
                                        <td class="${(latest.Critical + latest.High) < (first.Critical + first.High) ? 'value-improved' : (latest.Critical + latest.High) > (first.Critical + first.High) ? 'value-declined' : 'value-stable'}">
                                            ${(latest.Critical + latest.High) > (first.Critical + first.High) ? '+' : ''}${(latest.Critical + latest.High) - (first.Critical + first.High)}
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="metric-category">Critical+High Rate</td>
                                        ${releaseData.map((r, idx) => {
                                            const val = r.Critical_High_Rate;
                                            let cssClass = 'value-cell';
                                            if (idx > 0) {
                                                const prev = releaseData[idx - 1].Critical_High_Rate;
                                                cssClass = val < prev ? 'value-improved' : val > prev ? 'value-declined' : 'value-stable';
                                            }
                                            return `<td class="${cssClass}">${val}%</td>`;
                                        }).join('')}
                                        <td class="${latest.Critical_High_Rate < first.Critical_High_Rate ? 'value-improved' : latest.Critical_High_Rate > first.Critical_High_Rate ? 'value-declined' : 'value-stable'}">
                                            ${latest.Critical_High_Rate > first.Critical_High_Rate ? '+' : ''}${(latest.Critical_High_Rate - first.Critical_High_Rate).toFixed(2)}%
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="metric-category">Dev RCA %</td>
                                        ${releaseData.map((r, idx) => {
                                            const val = r.dev_rca_pct || 0;
                                            let cssClass = 'value-cell';
                                            if (idx > 0) {
                                                const prev = releaseData[idx - 1].dev_rca_pct || 0;
                                                cssClass = val < prev ? 'value-improved' : val > prev ? 'value-declined' : 'value-stable';
                                            }
                                            return `<td class="${cssClass}">${val}%</td>`;
                                        }).join('')}
                                        <td class="${(releaseData[releaseData.length-1].dev_rca_pct || 0) < (releaseData[0].dev_rca_pct || 0) ? 'value-improved' : (releaseData[releaseData.length-1].dev_rca_pct || 0) > (releaseData[0].dev_rca_pct || 0) ? 'value-declined' : 'value-stable'}">
                                            ${(releaseData[releaseData.length-1].dev_rca_pct || 0) > (releaseData[0].dev_rca_pct || 0) ? '+' : ''}${((releaseData[releaseData.length-1].dev_rca_pct || 0) - (releaseData[0].dev_rca_pct || 0)).toFixed(2)}%
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="metric-category">Reopened Bugs</td>
                                        ${releaseData.map((r, idx) => {
                                            const val = r.total_reopened || 0;
                                            let cssClass = 'value-cell';
                                            if (idx > 0) {
                                                const prev = releaseData[idx - 1].total_reopened || 0;
                                                cssClass = val < prev ? 'value-improved' : val > prev ? 'value-declined' : 'value-stable';
                                            }
                                            return `<td class="${cssClass}">${val}</td>`;
                                        }).join('')}
                                        <td class="${(releaseData[releaseData.length-1].total_reopened || 0) < (releaseData[0].total_reopened || 0) ? 'value-improved' : (releaseData[releaseData.length-1].total_reopened || 0) > (releaseData[0].total_reopened || 0) ? 'value-declined' : 'value-stable'}">
                                            ${(releaseData[releaseData.length-1].total_reopened || 0) > (releaseData[0].total_reopened || 0) ? '+' : ''}${(releaseData[releaseData.length-1].total_reopened || 0) - (releaseData[0].total_reopened || 0)}
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="metric-category">Reopen Rate</td>
                                        ${releaseData.map((r, idx) => {
                                            const val = r.reopen_rate || 0;
                                            let cssClass = 'value-cell';
                                            if (idx > 0) {
                                                const prev = releaseData[idx - 1].reopen_rate || 0;
                                                cssClass = val < prev ? 'value-improved' : val > prev ? 'value-declined' : 'value-stable';
                                            }
                                            return `<td class="${cssClass}">${val}%</td>`;
                                        }).join('')}
                                        <td class="${(releaseData[releaseData.length-1].reopen_rate || 0) < (releaseData[0].reopen_rate || 0) ? 'value-improved' : (releaseData[releaseData.length-1].reopen_rate || 0) > (releaseData[0].reopen_rate || 0) ? 'value-declined' : 'value-stable'}">
                                            ${(releaseData[releaseData.length-1].reopen_rate || 0) > (releaseData[0].reopen_rate || 0) ? '+' : ''}${((releaseData[releaseData.length-1].reopen_rate || 0) - (releaseData[0].reopen_rate || 0)).toFixed(2)}%
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="metric-category">SLA Not Met Bugs</td>
                                        ${releaseData.map((r, idx) => {
                                            const val = r.sla_not_met || 0;
                                            let cssClass = 'value-cell';
                                            if (idx > 0) {
                                                const prev = releaseData[idx - 1].sla_not_met || 0;
                                                cssClass = val < prev ? 'value-improved' : val > prev ? 'value-declined' : 'value-stable';
                                            }
                                            return `<td class="${cssClass}">${val}</td>`;
                                        }).join('')}
                                        <td class="${(releaseData[releaseData.length-1].sla_not_met || 0) < (releaseData[0].sla_not_met || 0) ? 'value-improved' : (releaseData[releaseData.length-1].sla_not_met || 0) > (releaseData[0].sla_not_met || 0) ? 'value-declined' : 'value-stable'}">
                                            ${(releaseData[releaseData.length-1].sla_not_met || 0) > (releaseData[0].sla_not_met || 0) ? '+' : ''}${(releaseData[releaseData.length-1].sla_not_met || 0) - (releaseData[0].sla_not_met || 0)}
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="metric-category">SLA Compliance</td>
                                        ${releaseData.map((r, idx) => {
                                            const val = r.sla_compliance || 0;
                                            let cssClass = 'value-cell';
                                            if (idx > 0) {
                                                const prev = releaseData[idx - 1].sla_compliance || 0;
                                                cssClass = val > prev ? 'value-improved' : val < prev ? 'value-declined' : 'value-stable';
                                            }
                                            return `<td class="${cssClass}">${val}%</td>`;
                                        }).join('')}
                                        <td class="${(releaseData[releaseData.length-1].sla_compliance || 0) > (releaseData[0].sla_compliance || 0) ? 'value-improved' : (releaseData[releaseData.length-1].sla_compliance || 0) < (releaseData[0].sla_compliance || 0) ? 'value-declined' : 'value-stable'}">
                                            ${(releaseData[releaseData.length-1].sla_compliance || 0) > (releaseData[0].sla_compliance || 0) ? '+' : ''}${((releaseData[releaseData.length-1].sla_compliance || 0) - (releaseData[0].sla_compliance || 0)).toFixed(2)}%
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="metric-category">Closure Efficiency</td>
                                        ${releaseData.map((r, idx) => {
                                            const val = r.closure_efficiency || 0;
                                            let cssClass = 'value-cell';
                                            if (idx > 0) {
                                                const prev = releaseData[idx - 1].closure_efficiency || 0;
                                                cssClass = val > prev ? 'value-improved' : val < prev ? 'value-declined' : 'value-stable';
                                            }
                                            return `<td class="${cssClass}">${val}%</td>`;
                                        }).join('')}
                                        <td class="${(releaseData[releaseData.length-1].closure_efficiency || 0) > (releaseData[0].closure_efficiency || 0) ? 'value-improved' : (releaseData[releaseData.length-1].closure_efficiency || 0) < (releaseData[0].closure_efficiency || 0) ? 'value-declined' : 'value-stable'}">
                                            ${(releaseData[releaseData.length-1].closure_efficiency || 0) > (releaseData[0].closure_efficiency || 0) ? '+' : ''}${((releaseData[releaseData.length-1].closure_efficiency || 0) - (releaseData[0].closure_efficiency || 0)).toFixed(2)}%
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="metric-category">Agent No %</td>
                                        ${releaseData.map((r, idx) => {
                                            const val = r.agent_no_pct || 0;
                                            let cssClass = 'value-cell';
                                            if (idx > 0) {
                                                const prev = releaseData[idx - 1].agent_no_pct || 0;
                                                cssClass = val < prev ? 'value-improved' : val > prev ? 'value-declined' : 'value-stable';
                                            }
                                            return `<td class="${cssClass}">${val}%</td>`;
                                        }).join('')}
                                        <td class="${(releaseData[releaseData.length-1].agent_no_pct || 0) < (releaseData[0].agent_no_pct || 0) ? 'value-improved' : (releaseData[releaseData.length-1].agent_no_pct || 0) > (releaseData[0].agent_no_pct || 0) ? 'value-declined' : 'value-stable'}">
                                            ${(releaseData[releaseData.length-1].agent_no_pct || 0) > (releaseData[0].agent_no_pct || 0) ? '+' : ''}${((releaseData[releaseData.length-1].agent_no_pct || 0) - (releaseData[0].agent_no_pct || 0)).toFixed(2)}%
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="metric-category">Delayed Delivery %</td>
                                        ${releaseData.map((r, idx) => {
                                            const val = r.delayed_yes_pct || 0;
                                            let cssClass = 'value-cell';
                                            if (idx > 0) {
                                                const prev = releaseData[idx - 1].delayed_yes_pct || 0;
                                                cssClass = val < prev ? 'value-improved' : val > prev ? 'value-declined' : 'value-stable';
                                            }
                                            return `<td class="${cssClass}">${val}%</td>`;
                                        }).join('')}
                                        <td class="${(releaseData[releaseData.length-1].delayed_yes_pct || 0) < (releaseData[0].delayed_yes_pct || 0) ? 'value-improved' : (releaseData[releaseData.length-1].delayed_yes_pct || 0) > (releaseData[0].delayed_yes_pct || 0) ? 'value-declined' : 'value-stable'}">
                                            ${(releaseData[releaseData.length-1].delayed_yes_pct || 0) > (releaseData[0].delayed_yes_pct || 0) ? '+' : ''}${((releaseData[releaseData.length-1].delayed_yes_pct || 0) - (releaseData[0].delayed_yes_pct || 0)).toFixed(2)}%
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                        
                    </div>
                </div>
            `;
        }

        function populateScrums() {
            // All scrums (filtered)
            const allContainer = document.getElementById('all-scrums');
            allContainer.innerHTML = filteredData.length > 0 
                ? filteredData.map((s, idx) => createScrumCard(s, idx)).join('')
                : '<p style="text-align:center;padding:50px;color:#999;">No scrums match the selected filters</p>';
        }

        // Initialize dashboard on page load
        window.addEventListener('DOMContentLoaded', function() {
            initializeFilters();
            generateADScoreTable();
            generateSMScoreTable();
            generateMScoreTable();
            generatePerformanceTable();
            populateScrums();
        });
    </script>
</body>
</html>
'''

# Save HTML file
output_file = f'{base_path}\\Scrum_Trend_Dashboard.html'
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"\n✓ Dashboard generated successfully!")
print(f"\n📄 Output file: {output_file}")
print(f"\n{'='*80}")
print("DASHBOARD GENERATION COMPLETE!")
print(f"{'='*80}")
print("\n✅ Open the HTML file in your browser to view individual scrum trends")
