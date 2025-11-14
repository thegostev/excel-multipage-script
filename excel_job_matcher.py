"""
EXCEL PYTHON IN EXCEL - JOB MATCHER SCRIPT
===========================================

This script matches active Jenkins jobs against a baseline list and marks them in Excel.

COPY-PASTE READY: Just adjust the sheet names below if they're different in your workbook.

USE IN EXCEL:
1. Insert → Python in Excel ribbon
2. Paste this entire code into the Python cell
3. Adjust BASELINE_SHEET and ACTIVE_SHEET names below if needed
4. Run the cell
5. The output will show you the results

"""

import pandas as pd

# ============================================================================
# CONFIGURATION - ADJUST THESE SHEET NAMES IF NEEDED
# ============================================================================

# Sheet with all 731 baseline jobs (the one you want to update)
BASELINE_SHEET = "OldSharedJenkinsJobsThatRunLast6Months_Nov_2025__all_-_baseline__"

# Sheet with 331 active jobs (with timestamps)
ACTIVE_SHEET = "OldSharedJenkinsJobsThatRunLast6Months_Nov_2025__active__-2"

# ============================================================================
# STEP 1: READ ACTIVE JOBS AND CLEAN THE NAMES
# ============================================================================

# Read active jobs sheet starting from row 4 (rows 1-3 are metadata/headers)
# Column A contains: "Job Name - Mon Nov 03 12:07:14 UTC 2025"
active_jobs_range = f"'{ACTIVE_SHEET}'!A4:A340"
active_df = xl(active_jobs_range, headers=False)

# Clean up the job names by removing timestamps
# Split on " - " and take the first part (the job name)
active_df['Job_Clean'] = active_df.iloc[:, 0].astype(str).str.split(' - ', n=1).str[0]

# Remove any empty or NaN values
active_df = active_df[active_df['Job_Clean'].notna() & (active_df['Job_Clean'] != 'nan')]

# Create a set of active job names for fast lookup
active_jobs_set = set(active_df['Job_Clean'].unique())

print(f"✓ Found {len(active_jobs_set)} unique active jobs")

# ============================================================================
# STEP 2: READ BASELINE JOBS
# ============================================================================

# Read baseline sheet (732 rows total: 1 header + 731 data rows)
baseline_range = f"'{BASELINE_SHEET}'!A1:G732"
baseline_df = xl(baseline_range, headers=True)

print(f"✓ Read {len(baseline_df)} baseline jobs")

# ============================================================================
# STEP 3: MATCH JOBS AND MARK AS ACTIVE
# ============================================================================

# Check if each baseline job is in the active jobs set
# Create "Yes" for matches, empty string for non-matches
baseline_df['Active'] = baseline_df['Job'].apply(
    lambda job: 'Yes' if job in active_jobs_set else ''
)

# Count how many matches we found
matches_count = (baseline_df['Active'] == 'Yes').sum()

print(f"✓ Marked {matches_count} jobs as Active")
print(f"✓ {len(baseline_df) - matches_count} jobs remain inactive")

# ============================================================================
# STEP 4: PREPARE OUTPUT FOR EXCEL
# ============================================================================

# Return just the Active column to paste into Column D of your baseline sheet
# This makes it easy to copy the results to the right location
active_column = baseline_df[['Active']]

# Display summary
print("\n" + "="*60)
print("SUMMARY")
print("="*60)
print(f"Total baseline jobs: {len(baseline_df)}")
print(f"Total active jobs: {len(active_jobs_set)}")
print(f"Matched jobs: {matches_count}")
print(f"Match rate: {matches_count/len(baseline_df)*100:.1f}%")
print("="*60)
print("\nOUTPUT: Active column values are shown below")
print("Copy these values to Column D (Active) in your baseline sheet")
print("="*60)

# Return the Active column - Excel will display it
active_column
