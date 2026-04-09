"""
Individual Scrum Trend Analysis Across 3 Releases
Compares each scrum's performance with itself across Nov 8, Dec 13, and Jan 10 releases
Focus: Quality improvement/decline within the same scrum team
"""

import pandas as pd
import numpy as np
import json
from datetime import datetime

print("="*80)
print("INDIVIDUAL SCRUM TREND ANALYSIS")
print("Analyzing each scrum's performance across 3 releases")
print("="*80)

# ============= FILE PATHS =============
base_path = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics\analysis'
input_path = r'C:\Users\d.sampathkumar\GHC files\Quality Metrics'

quality_metrics_file = f'{input_path}\\Quality Metrics Complete Input file for stories and bugs.xlsx'
closure_file = f'{input_path}\\quality_metrics_input_closure_reopen.xlsx'
rca_file = f'{input_path}\\quality metrcis input_RCA Analysis.xlsx'
stories_file = f'{input_path}\\Stories score summary.xlsx'

# ============= LOAD DATA =============
print("\n✓ Loading source data...")

# Load all 3 releases
xl = pd.ExcelFile(quality_metrics_file)
releases = ['Nov 8', 'Dec 13', 'Jan 10']
print(f"  - Analyzing releases: {releases}")

release_data = {}
for release in releases:
    release_data[release] = pd.read_excel(xl, release)

# Load closure/reopen data (per release)
closure_xl = pd.ExcelFile(closure_file)
closure_data = {}
for release in releases:
    closure_data[release] = pd.read_excel(closure_xl, release)
print(f"  - Loaded closure/reopen records for {len(releases)} releases")

# Load RCA data (per release)
rca_xl = pd.ExcelFile(rca_file)
rca_data = {}
for release in releases:
    rca_data[release] = pd.read_excel(rca_xl, release)
print(f"  - Loaded RCA records for {len(releases)} releases")

# Load Stories score summary data (per release)
stories_xl = pd.ExcelFile(stories_file)
stories_data = {}
for release in releases:
    stories_data[release] = pd.read_excel(stories_xl, release)
print(f"  - Loaded story delivery records for {len(releases)} releases")

# ============= NORMALIZE NAMES =============
print("\n✓ Normalizing scrum names for matching...")
for release in releases:
    release_data[release]['Node Name Lower'] = release_data[release]['Node Name'].str.lower().str.strip()
    rca_data[release]['Node Name Lower'] = rca_data[release]['Node Name'].str.lower().str.strip()
    closure_data[release]['Node Name Lower'] = closure_data[release]['Node Name'].str.lower().str.strip()
    stories_data[release]['Node Name Lower'] = stories_data[release]['Node Name'].str.lower().str.strip()

# Get unique scrum names from first release
scrums = release_data['Nov 8']['Node Name'].unique()
print(f"  - Found {len(scrums)} unique scrums")

# ============= ANALYZE EACH SCRUM =============
print("\n✓ Analyzing individual scrum trends...")

scrum_trends = []

for scrum in scrums:
    scrum_lower = scrum.lower().strip()
    
    # Collect data for this scrum across all 3 releases
    scrum_release_data = []
    
    for release in releases:
        df = release_data[release]
        scrum_data = df[df['Node Name Lower'] == scrum_lower]
        
        if scrum_data.empty:
            continue
        
        row = scrum_data.iloc[0]
        
        # Extract metrics
        testable_stories = float(row.get('Testable stories', 0))
        testing_na_stories = float(row.get('Testing Not Applicable Stories', 0))
        total_stories = testable_stories + testing_na_stories
        story_points = float(row.get('Total story Points', 0))
        total_bugs = float(row.get('Total Bugs', 0))
        valid_bugs = float(row.get('Valid bugs', 0))
        critical = float(row.get('Total Critical', 0))
        high = float(row.get('Total High', 0))
        medium = float(row.get('Total Medium', 0))
        low = float(row.get('Total Low', 0))
        
        # Calculate defect density
        defect_density = (valid_bugs / story_points * 100) if story_points > 0 else 0
        
        # Severity rates
        critical_rate = (critical / valid_bugs * 100) if valid_bugs > 0 else 0
        high_rate = (high / valid_bugs * 100) if valid_bugs > 0 else 0
        critical_high_rate = ((critical + high) / valid_bugs * 100) if valid_bugs > 0 else 0
        
        # Get per-release RCA data
        release_rca = rca_data[release][rca_data[release]['Node Name Lower'] == scrum_lower]
        release_dev_rca = release_rca[release_rca['RCA Type'] == 'DEV Countable RCA']['Count'].sum() if not release_rca.empty else 0
        release_total_rca = release_rca['Count'].sum() if not release_rca.empty else 0
        release_dev_rca_pct = (release_dev_rca / release_total_rca * 100) if release_total_rca > 0 else 0
        
        # Get per-release closure/reopen data
        release_closure = closure_data[release][closure_data[release]['Node Name Lower'] == scrum_lower]
        release_closed = len(release_closure)
        release_reopened = release_closure['Re open Count'].sum() if not release_closure.empty else 0
        release_reopen_rate = (release_reopened / release_closed * 100) if release_closed > 0 else 0
        
        # Get per-release SLA data
        release_sla_met = len(release_closure[release_closure['SLA'] == 'Met SLA']) if not release_closure.empty else 0
        release_sla_not_met = len(release_closure[release_closure['SLA'] == 'Not Met SLA']) if not release_closure.empty else 0
        release_sla_compliance = (release_sla_met / release_closed * 100) if release_closed > 0 else 0
        release_closure_efficiency = 100 - release_reopen_rate if release_reopen_rate < 100 else 0
        
        # Get per-release Story Delivery data
        release_stories = stories_data[release][stories_data[release]['Node Name Lower'] == scrum_lower]
        release_total_stories = len(release_stories) if not release_stories.empty else 0
        release_agent_no = (release_stories['Agent Augmented delivery_Development'] == 'No').sum() if not release_stories.empty else 0
        release_delayed_yes = (release_stories['Delayed Story Delivery'] == 'Yes').sum() if not release_stories.empty else 0
        release_agent_no_pct = (release_agent_no / release_total_stories * 100) if release_total_stories > 0 else 0
        release_delayed_yes_pct = (release_delayed_yes / release_total_stories * 100) if release_total_stories > 0 else 0
        
        scrum_release_data.append({
            'Release': release,
            'Testable_Stories': int(testable_stories),
            'Total_Stories': int(total_stories),
            'Story_Points': int(story_points),
            'Total_Bugs': int(total_bugs),
            'Valid_Bugs': int(valid_bugs),
            'Critical': int(critical),
            'High': int(high),
            'Medium': int(medium),
            'Low': int(low),
            'Defect_Density': round(defect_density, 2),
            'Critical_Rate': round(critical_rate, 2),
            'High_Rate': round(high_rate, 2),
            'Critical_High_Rate': round(critical_high_rate, 2),
            'dev_rca_bugs': int(release_dev_rca),
            'dev_rca_pct': round(release_dev_rca_pct, 2),
            'total_reopened': int(release_reopened),
            'reopen_rate': round(release_reopen_rate, 2),
            'sla_not_met': int(release_sla_not_met),
            'sla_compliance': round(release_sla_compliance, 2),
            'closure_efficiency': round(release_closure_efficiency, 2),
            'total_stories': int(release_total_stories),
            'agent_no_count': int(release_agent_no),
            'agent_no_pct': round(release_agent_no_pct, 2),
            'delayed_yes_count': int(release_delayed_yes),
            'delayed_yes_pct': round(release_delayed_yes_pct, 2)
        })
    
    # Need at least 2 releases for trend
    if len(scrum_release_data) < 2:
        continue
    
    # ============= CALCULATE AGGREGATE METRICS FROM PER-RELEASE DATA =============
    
    # ============= ANALYZE TRENDS =============
    first = scrum_release_data[0]
    latest = scrum_release_data[-1]
    
    # Use Dec 13 and Jan 10 for DD scoring (if available, otherwise use first and latest)
    if len(scrum_release_data) >= 3:
        # Use Dec 13 (middle) and Jan 10 (last) for DD change calculation
        dec_13_data = scrum_release_data[1]
        jan_10_data = scrum_release_data[2]
        dd_change = jan_10_data['Defect_Density'] - dec_13_data['Defect_Density']
        dd_pct_change = (dd_change / dec_13_data['Defect_Density'] * 100) if dec_13_data['Defect_Density'] > 0 else 0
    else:
        # Fallback to first and latest if less than 3 releases
        dd_change = latest['Defect_Density'] - first['Defect_Density']
        dd_pct_change = (dd_change / first['Defect_Density'] * 100) if first['Defect_Density'] > 0 else 0
    
    # Story volume trend (still use first to latest for overall trend)
    story_volume_change = latest['Story_Points'] - first['Story_Points']
    story_volume_pct = (story_volume_change / first['Story_Points'] * 100) if first['Story_Points'] > 0 else 0
    
    # Severity trend
    critical_high_change = latest['Critical_High_Rate'] - first['Critical_High_Rate']
    
    # ============= RCA ANALYSIS (AGGREGATE FROM PER-RELEASE DATA) =============
    dev_rca = sum(r['dev_rca_bugs'] for r in scrum_release_data)
    # Calculate total RCA from all releases
    total_rca = 0
    for release in releases:
        release_rca = rca_data[release][rca_data[release]['Node Name Lower'] == scrum_lower]
        total_rca += release_rca['Count'].sum() if not release_rca.empty else 0
    dev_rca_pct = (dev_rca / total_rca * 100) if total_rca > 0 else 0
    
    # ============= CLOSURE ANALYSIS (AGGREGATE FROM PER-RELEASE DATA) =============
    total_closed = sum(r.get('total_reopened', 0) for r in scrum_release_data) + sum(len(closure_data[rel][closure_data[rel]['Node Name Lower'] == scrum_lower]) for rel in releases)
    total_reopened = sum(r['total_reopened'] for r in scrum_release_data)
    reopen_rate = (total_reopened / total_closed * 100) if total_closed > 0 else 0
    
    # Aggregate SLA metrics
    total_sla_not_met = sum(r['sla_not_met'] for r in scrum_release_data)
    sla_compliance = sum(r['sla_compliance'] * len(closure_data[releases[i]][closure_data[releases[i]]['Node Name Lower'] == scrum_lower]) for i, r in enumerate(scrum_release_data)) / total_closed if total_closed > 0 else 0
    
    # Aggregate closure by severity from all releases
    critical_closed = sum(len(closure_data[rel][(closure_data[rel]['Node Name Lower'] == scrum_lower) & (closure_data[rel]['Severity'] == '1 - Critical')]) for rel in releases)
    high_closed = sum(len(closure_data[rel][(closure_data[rel]['Node Name Lower'] == scrum_lower) & (closure_data[rel]['Severity'] == '2 - High')]) for rel in releases)
    
    # ============= STORY DELIVERY ANALYSIS (AGGREGATE FROM PER-RELEASE DATA) =============
    total_stories_all = sum(r['total_stories'] for r in scrum_release_data)
    total_agent_no = sum(r['agent_no_count'] for r in scrum_release_data)
    total_delayed_yes = sum(r['delayed_yes_count'] for r in scrum_release_data)
    agent_no_pct = (total_agent_no / total_stories_all * 100) if total_stories_all > 0 else 0
    delayed_yes_pct = (total_delayed_yes / total_stories_all * 100) if total_stories_all > 0 else 0
    
    # ============= QUALITY SCORE CALCULATION =============
    
    # 1. Defect Density Score (0-100, higher is better)
    if dd_change < -20:
        dd_score = 100  # Major improvement
    elif dd_change < -10:
        dd_score = 90
    elif dd_change < -5:
        dd_score = 80
    elif dd_change < 0:
        dd_score = 70
    elif dd_change < 5:
        dd_score = 60
    elif dd_change < 10:
        dd_score = 40
    elif dd_change < 20:
        dd_score = 20
    else:
        dd_score = 0  # Major decline
    
    # 2. Severity Score (0-100, lower critical/high is better)
    if critical_high_change < -10:
        severity_score = 100
    elif critical_high_change < -5:
        severity_score = 80
    elif critical_high_change < 0:
        severity_score = 70
    elif critical_high_change < 5:
        severity_score = 60
    elif critical_high_change < 10:
        severity_score = 40
    else:
        severity_score = 20
    
    # 3. RCA Score (0-100, lower dev ownership is better)
    if dev_rca_pct < 20:
        rca_score = 100
    elif dev_rca_pct < 40:
        rca_score = 80
    elif dev_rca_pct < 60:
        rca_score = 60
    elif dev_rca_pct < 80:
        rca_score = 40
    else:
        rca_score = 20
    
    # 4. Closure Efficiency Score (0-100)
    closure_efficiency = 100 - reopen_rate if reopen_rate < 100 else 0
    if closure_efficiency >= 95:
        closure_score = 100
    elif closure_efficiency >= 90:
        closure_score = 90
    elif closure_efficiency >= 85:
        closure_score = 80
    elif closure_efficiency >= 80:
        closure_score = 70
    elif closure_efficiency >= 70:
        closure_score = 50
    else:
        closure_score = 30
    
    # 5. SLA Compliance Score (0-100)
    if sla_compliance >= 95:
        sla_score = 100
    elif sla_compliance >= 90:
        sla_score = 90
    elif sla_compliance >= 85:
        sla_score = 80
    elif sla_compliance >= 80:
        sla_score = 70
    elif sla_compliance >= 70:
        sla_score = 60
    else:
        sla_score = 40
    
    # 6. Agent Augmented Score (0-100)
    # Lower Agent No % is better
    if agent_no_pct < 5:
        agent_augmented_score = 100
    elif agent_no_pct < 10:
        agent_augmented_score = 90
    elif agent_no_pct < 15:
        agent_augmented_score = 80
    elif agent_no_pct < 20:
        agent_augmented_score = 70
    elif agent_no_pct < 30:
        agent_augmented_score = 60
    elif agent_no_pct < 40:
        agent_augmented_score = 50
    elif agent_no_pct < 50:
        agent_augmented_score = 40
    else:
        agent_augmented_score = 20
    
    # 7. Story Delivery Delay Score (0-100)
    # Lower Delayed Yes % is better
    if delayed_yes_pct < 5:
        story_ontime_score = 100
    elif delayed_yes_pct < 10:
        story_ontime_score = 90
    elif delayed_yes_pct < 15:
        story_ontime_score = 80
    elif delayed_yes_pct < 20:
        story_ontime_score = 70
    elif delayed_yes_pct < 30:
        story_ontime_score = 60
    elif delayed_yes_pct < 40:
        story_ontime_score = 50
    elif delayed_yes_pct < 50:
        story_ontime_score = 40
    else:
        story_ontime_score = 20
    
    # Overall Trend Score (weighted average)
    overall_score = (
        dd_score * 0.30 +                # 30% - Defect Density Trend
        severity_score * 0.10 +          # 10% - Severity Trend
        rca_score * 0.20 +               # 20% - RCA Quality
        closure_score * 0.10 +           # 10% - Closure Efficiency
        sla_score * 0.10 +               # 10% - SLA Compliance
        agent_augmented_score * 0.10 +   # 10% - Agent Augmented Score
        story_ontime_score * 0.10         # 10% - Story On Time Delivery Score
    )
    
    # Determine trend direction
    if overall_score >= 80:
        trend = "EXCELLENT IMPROVEMENT"
        trend_class = "excellent"
    elif overall_score >= 70:
        trend = "GOOD IMPROVEMENT"
        trend_class = "good"
    elif overall_score >= 60:
        trend = "MODERATE IMPROVEMENT"
        trend_class = "moderate"
    elif overall_score >= 50:
        trend = "STABLE"
        trend_class = "stable"
    elif overall_score >= 40:
        trend = "SLIGHT DECLINE"
        trend_class = "slight-decline"
    else:
        trend = "SIGNIFICANT DECLINE"
        trend_class = "decline"
    
    # Get POC information from the first available release
    first_release = releases[0]
    ad_poc = release_data[first_release][release_data[first_release]['Node Name'] == scrum]['AD POC'].iloc[0] if not release_data[first_release][release_data[first_release]['Node Name'] == scrum].empty else 'N/A'
    sm_poc = release_data[first_release][release_data[first_release]['Node Name'] == scrum]['SM POC'].iloc[0] if not release_data[first_release][release_data[first_release]['Node Name'] == scrum].empty else 'N/A'
    test_manager = release_data[first_release][release_data[first_release]['Node Name'] == scrum]['M POC'].iloc[0] if 'M POC' in release_data[first_release].columns and not release_data[first_release][release_data[first_release]['Node Name'] == scrum].empty else 'N/A'
    
    scrum_trends.append({
        'Scrum': scrum,
        'AD_POC': ad_poc,
        'SM_POC': sm_poc,
        'Test_Manager': test_manager,
        
        # Release-wise data
        'Release_Data': scrum_release_data,
        
        # Trend metrics
        'Story_Volume_Change': story_volume_change,
        'Story_Volume_Pct_Change': round(story_volume_pct, 1),
        'DD_Change': round(dd_change, 2),
        'DD_Pct_Change': round(dd_pct_change, 1),
        'Critical_High_Change': round(critical_high_change, 2),
        
        # RCA metrics
        'Total_RCA': int(total_rca),
        'Dev_RCA': int(dev_rca),
        'Dev_RCA_Pct': round(dev_rca_pct, 1),
        
        # Closure metrics
        'Total_Closed': total_closed,
        'Total_Reopened': int(total_reopened),
        'Reopen_Rate': round(reopen_rate, 1),
        'Closure_Efficiency': round(closure_efficiency, 1),
        'SLA_Compliance': round(sla_compliance, 1),
        'Critical_Closed': critical_closed,
        'High_Closed': high_closed,
        
        # Story Delivery metrics
        'Total_Stories': int(total_stories_all),
        'Agent_No_Count': int(total_agent_no),
        'Agent_No_Pct': round(agent_no_pct, 1),
        'Delayed_Yes_Count': int(total_delayed_yes),
        'Delayed_Yes_Pct': round(delayed_yes_pct, 1),
        
        # Scores
        'DD_Score': round(dd_score, 1),
        'Severity_Score': round(severity_score, 1),
        'RCA_Score': round(rca_score, 1),
        'Closure_Score': round(closure_score, 1),
        'SLA_Score': round(sla_score, 1),
        'Agent_Augmented_Score': round(agent_augmented_score, 1),
        'Story_OnTime_Score': round(story_ontime_score, 1),
        'Overall_Score': round(overall_score, 1),
        
        # Classification
        'Trend': trend,
        'Trend_Class': trend_class,
        
        # First and Latest values for quick reference
        'First_DD': first['Defect_Density'],
        'Latest_DD': latest['Defect_Density'],
        'First_Story_Points': first['Story_Points'],
        'Latest_Story_Points': latest['Story_Points'],
        'First_Valid_Bugs': first['Valid_Bugs'],
        'Latest_Valid_Bugs': latest['Valid_Bugs']
    })

# Convert to DataFrame and sort by overall score
trends_df = pd.DataFrame(scrum_trends)
trends_df = trends_df.sort_values('Overall_Score', ascending=False)

# ============= CATEGORIZE SCRUMS =============
excellent = trends_df[trends_df['Trend'] == 'EXCELLENT IMPROVEMENT']
good = trends_df[trends_df['Trend'] == 'GOOD IMPROVEMENT']
moderate = trends_df[trends_df['Trend'] == 'MODERATE IMPROVEMENT']
stable = trends_df[trends_df['Trend'] == 'STABLE']
slight_decline = trends_df[trends_df['Trend'] == 'SLIGHT DECLINE']
decline = trends_df[trends_df['Trend'] == 'SIGNIFICANT DECLINE']

# ============= SAVE TO JSON =============
output_data = {
    'generation_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    'releases_analyzed': releases,
    'summary_stats': {
        'Total_Scrums': len(trends_df),
        'Excellent_Improvement': len(excellent),
        'Good_Improvement': len(good),
        'Moderate_Improvement': len(moderate),
        'Stable': len(stable),
        'Slight_Decline': len(slight_decline),
        'Significant_Decline': len(decline),
        'Avg_Overall_Score': round(trends_df['Overall_Score'].mean(), 1),
        'Avg_DD_Change': round(trends_df['DD_Change'].mean(), 2),
        'Avg_Reopen_Rate': round(trends_df['Reopen_Rate'].mean(), 1),
        'Avg_SLA_Compliance': round(trends_df['SLA_Compliance'].mean(), 1)
    },
    'all_scrums': trends_df.to_dict('records'),
    'excellent_improvement': excellent.to_dict('records'),
    'good_improvement': good.to_dict('records'),
    'moderate_improvement': moderate.to_dict('records'),
    'stable': stable.to_dict('records'),
    'slight_decline': slight_decline.to_dict('records'),
    'significant_decline': decline.to_dict('records')
}

output_json = f'{base_path}\\scrum_trend_data.json'
with open(output_json, 'w') as f:
    json.dump(output_data, f, indent=2)

print(f"\n✓ Analysis data saved to: {output_json}")

# ============= PRINT SUMMARY =============
print(f"\n{'='*80}")
print("SCRUM TREND ANALYSIS SUMMARY")
print(f"{'='*80}")
print(f"\nTotal Scrums Analyzed: {len(trends_df)}")
print(f"Average Overall Score: {trends_df['Overall_Score'].mean():.1f}/100")

print(f"\n{'='*80}")
print("TREND DISTRIBUTION")
print(f"{'='*80}")
print(f"📈 EXCELLENT IMPROVEMENT ({len(excellent)}): Score >= 80")
print(f"✅ GOOD IMPROVEMENT ({len(good)}): Score 70-79")
print(f"⚠️  MODERATE IMPROVEMENT ({len(moderate)}): Score 60-69")
print(f"➡️  STABLE ({len(stable)}): Score 50-59")
print(f"🔴 SLIGHT DECLINE ({len(slight_decline)}): Score 40-49")
print(f"🚨 SIGNIFICANT DECLINE ({len(decline)}): Score < 40")

print(f"\n{'='*80}")
print("TOP 10 MOST IMPROVED SCRUMS")
print(f"{'='*80}")
for idx, row in trends_df.head(10).iterrows():
    print(f"📈 {row['Scrum']:30s} | Score: {row['Overall_Score']:5.1f} | DD Change: {row['DD_Change']:+7.2f}% | {row['Trend']}")

print(f"\n{'='*80}")
print("BOTTOM 10 DECLINING SCRUMS")
print(f"{'='*80}")
for idx, row in trends_df.tail(10).iterrows():
    print(f"📉 {row['Scrum']:30s} | Score: {row['Overall_Score']:5.1f} | DD Change: {row['DD_Change']:+7.2f}% | {row['Trend']}")

print(f"\n{'='*80}")
print("ANALYSIS COMPLETE!")
print(f"{'='*80}")
print("\nNext: Run 'python generate_scrum_trend_dashboard.py' to create the HTML dashboard")
