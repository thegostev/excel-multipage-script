"""
EXCEL PYTHON IN EXCEL - JOB MATCHER (FULL OUTPUT VERSION)
==========================================================

This is an ALTERNATIVE version that returns the COMPLETE baseline table with the
Active column populated. This might be easier if you want to see all columns together.

DIFFERENCE FROM MAIN VERSION:
- Main version (excel_job_matcher.py): Returns just the Active column
- This version: Returns the complete baseline table with Active column added

USE THIS VERSION IF:
- You want to see all columns together
- You want to create a new table with results
- You prefer to copy the entire result to a new location

COPY-PASTE READY: Just adjust the sheet names below.

"""

import pandas as pd

# ============================================================================
# CONFIGURATION - ADJUST THESE SHEET NAMES IF NEEDED
# ============================================================================

BASELINE_SHEET = "OldSharedJenkinsJobsThatRunLast6Months_Nov_2025__all_-_baseline__"
ACTIVE_SHEET = "OldSharedJenkinsJobsThatRunLast6Months_Nov_2025__active__-2"

# ============================================================================
# STEP 1: READ AND PROCESS ACTIVE JOBS
# ============================================================================

active_jobs_range = f"'{ACTIVE_SHEET}'!A4:A340"
active_df = xl(active_jobs_range, headers=False)

# Clean job names by removing timestamps
active_df['Job_Clean'] = active_df.iloc[:, 0].astype(str).str.split(' - ', n=1).str[0]
active_df = active_df[active_df['Job_Clean'].notna() & (active_df['Job_Clean'] != 'nan')]

# Create set of active jobs
active_jobs_set = set(active_df['Job_Clean'].unique())

print(f"✓ Found {len(active_jobs_set)} unique active jobs")

# ============================================================================
# STEP 2: READ BASELINE AND ADD ACTIVE MARKERS
# ============================================================================

baseline_range = f"'{BASELINE_SHEET}'!A1:G732"
baseline_df = xl(baseline_range, headers=True)

print(f"✓ Read {len(baseline_df)} baseline jobs")

# Mark active jobs
baseline_df['Active'] = baseline_df['Job'].apply(
    lambda job: 'Yes' if job in active_jobs_set else ''
)

matches_count = (baseline_df['Active'] == 'Yes').sum()

# ============================================================================
# SUMMARY AND OUTPUT
# ============================================================================

print(f"✓ Marked {matches_count} jobs as Active")
print("\n" + "="*60)
print("SUMMARY")
print("="*60)
print(f"Total baseline jobs: {len(baseline_df)}")
print(f"Total active jobs: {len(active_jobs_set)}")
print(f"Matched jobs: {matches_count}")
print(f"Match rate: {matches_count/len(baseline_df)*100:.1f}%")
print("="*60)
print("\nOUTPUT: Complete baseline table with Active column")
print("="*60)

# Return the complete baseline table with Active column populated
baseline_df
