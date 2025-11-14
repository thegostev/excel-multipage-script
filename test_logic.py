"""
TEST SCRIPT - Validates the job matching logic
This script tests the core logic without needing Excel
"""

import pandas as pd

# ============================================================================
# CREATE SAMPLE DATA (simulating the Excel sheets)
# ============================================================================

print("Creating sample data...\n")

# Sample ACTIVE jobs (with timestamps like in Excel)
active_jobs_data = [
    "api-ecosystem/platform-gateway - Mon Nov 03 12:07:14 UTC 2025",
    "Ubik/mimic-aws - Mon Sep 29 18:29:17 UTC 2025",
    "custom-codesign-pipeline - Tue Oct 15 08:30:22 UTC 2025",
    "Apple Natives/Harvest/Worker - Wed Oct 30 14:22:11 UTC 2025",
    "data-processing/etl-pipeline - Thu Nov 02 09:15:33 UTC 2025",
]

# Sample BASELINE jobs (without timestamps)
baseline_jobs_data = {
    'Job': [
        'api-ecosystem/platform-gateway',
        'Ubik/mimic-aws',
        'custom-codesign-pipeline',
        'Apple Natives/Harvest/Worker',
        'data-processing/etl-pipeline',
        'old-legacy-job',
        'inactive-service',
        'deprecated-pipeline',
    ],
    'Folder': ['api', 'Ubik', 'custom', 'Apple', 'data', 'old', 'inactive', 'deprecated'],
    'Team': ['Platform', 'Ubik', 'DevOps', 'Apple', 'Data', 'Legacy', 'Old', 'Old'],
    'Active': ['', '', '', '', '', '', '', ''],  # Empty - we'll populate this
    'Moved': ['', '', '', '', '', '', '', ''],
    'Not needed': ['', '', '', '', '', '', '', ''],
    'Last communication': ['', '', '', '', '', '', '', ''],
}

# ============================================================================
# SIMULATE THE SCRIPT LOGIC
# ============================================================================

print("=" * 70)
print("TESTING JOB MATCHER LOGIC")
print("=" * 70)

# Step 1: Process active jobs (remove timestamps)
print("\nSTEP 1: Processing active jobs...")
active_df = pd.DataFrame({'Job': active_jobs_data})
active_df['Job_Clean'] = active_df['Job'].str.split(' - ', n=1).str[0]

print(f"  Original active jobs (with timestamps):")
for job in active_jobs_data:
    print(f"    - {job}")

print(f"\n  Cleaned active jobs (timestamps removed):")
for job in active_df['Job_Clean']:
    print(f"    - {job}")

# Create set of active jobs
active_jobs_set = set(active_df['Job_Clean'])
print(f"\n  ✓ Created set of {len(active_jobs_set)} unique active jobs")

# Step 2: Process baseline jobs
print("\nSTEP 2: Processing baseline jobs...")
baseline_df = pd.DataFrame(baseline_jobs_data)
print(f"  ✓ Loaded {len(baseline_df)} baseline jobs")

# Step 3: Match and mark
print("\nSTEP 3: Matching jobs...")
baseline_df['Active'] = baseline_df['Job'].apply(
    lambda job: 'Yes' if job in active_jobs_set else ''
)

matches_count = (baseline_df['Active'] == 'Yes').sum()
print(f"  ✓ Found {matches_count} matches")

# ============================================================================
# DISPLAY RESULTS
# ============================================================================

print("\n" + "=" * 70)
print("RESULTS")
print("=" * 70)

print("\nBaseline jobs with Active status:")
print(baseline_df[['Job', 'Active']].to_string(index=False))

print("\n" + "=" * 70)
print("SUMMARY")
print("=" * 70)
print(f"Total baseline jobs: {len(baseline_df)}")
print(f"Total active jobs: {len(active_jobs_set)}")
print(f"Matched jobs: {matches_count}")
print(f"Unmatched jobs: {len(baseline_df) - matches_count}")
print(f"Match rate: {matches_count/len(baseline_df)*100:.1f}%")

# ============================================================================
# VALIDATION
# ============================================================================

print("\n" + "=" * 70)
print("VALIDATION")
print("=" * 70)

expected_matches = 5  # We have 5 active jobs that should match
expected_non_matches = 3  # We have 3 baseline jobs that shouldn't match

if matches_count == expected_matches:
    print(f"✅ PASS: Found expected {expected_matches} matches")
else:
    print(f"❌ FAIL: Expected {expected_matches} matches, got {matches_count}")

if (len(baseline_df) - matches_count) == expected_non_matches:
    print(f"✅ PASS: Found expected {expected_non_matches} non-matches")
else:
    print(f"❌ FAIL: Expected {expected_non_matches} non-matches, got {len(baseline_df) - matches_count}")

# Verify specific jobs
print("\nSpot checks:")
test_cases = [
    ('api-ecosystem/platform-gateway', 'Yes', 'Should be active'),
    ('old-legacy-job', '', 'Should be inactive'),
    ('custom-codesign-pipeline', 'Yes', 'Should be active'),
    ('deprecated-pipeline', '', 'Should be inactive'),
]

all_passed = True
for job, expected, reason in test_cases:
    actual = baseline_df[baseline_df['Job'] == job]['Active'].values[0]
    if actual == expected:
        print(f"  ✅ {job}: {reason} - PASS")
    else:
        print(f"  ❌ {job}: {reason} - FAIL (expected '{expected}', got '{actual}')")
        all_passed = False

print("\n" + "=" * 70)
if all_passed and matches_count == expected_matches:
    print("🎉 ALL TESTS PASSED - Logic is working correctly!")
else:
    print("⚠️  SOME TESTS FAILED - Review the logic")
print("=" * 70)
