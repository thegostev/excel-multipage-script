# Example Output - What You'll See

This document shows you exactly what the script output will look like when it runs successfully.

---

## ✅ Successful Run - Complete Output

When you run the script in Excel, you'll see output like this:

```
✓ Found 331 unique active jobs
✓ Read 731 baseline jobs
✓ Marked 331 jobs as Active
✓ 400 jobs remain inactive

============================================================
SUMMARY
============================================================
Total baseline jobs: 731
Total active jobs: 331
Matched jobs: 331
Match rate: 45.3%
============================================================

OUTPUT: Active column values are shown below
Copy these values to Column D (Active) in your baseline sheet
============================================================
```

Below this summary, you'll see a table that looks like this:

```
    Active
0
1   Yes
2   Yes
3
4   Yes
5   Yes
6
7   Yes
...
730 Yes
```

---

## 📊 Understanding the Output

### Summary Statistics

**"✓ Found 331 unique active jobs"**
- Script successfully read the Active sheet
- Found 331 jobs with timestamps
- Removed timestamps and got unique job names

**"✓ Read 731 baseline jobs"**
- Script successfully read the Baseline sheet
- Found all 731 jobs

**"✓ Marked 331 jobs as Active"**
- Found matches between the two lists
- Created "Yes" markers for these jobs

**Match rate: 45.3%**
- 331 out of 731 jobs = 45.3%
- This means 45.3% of your baseline jobs are currently active

---

## 📋 The Active Column

The table output shows the Active column values for all 731 rows:

- **Row 0**: Empty (this is row 2 in Excel - first data row)
- **Row 1**: "Yes" (this is row 3 in Excel - second data row)
- **Row 2**: "Yes" (this is row 4 in Excel)
- **Row 3**: Empty
- etc.

### What the values mean:

- **"Yes"** = This job is active (found in the Active sheet)
- **Empty (blank)** = This job is not active

---

## 🎯 What You Should Do Next

1. **Verify the summary numbers:**
   - Does 331 matches make sense?
   - Is 45.3% match rate reasonable?

2. **Scroll through the Active column:**
   - You should see a mix of "Yes" and empty cells
   - About 331 "Yes" values
   - About 400 empty cells

3. **Copy to your baseline sheet:**
   - Select all the Active column values
   - Copy them (Ctrl+C)
   - Go to baseline sheet
   - Click on cell D2 (first data row, Active column)
   - Paste (Ctrl+V)

4. **Verify in Excel:**
   - Column D should now have "Yes" values
   - Use Excel filter to count "Yes" values (should be ~331)

---

## 🔍 Example: Before and After

### BEFORE (Baseline Sheet - Column D is empty):

| Job | Folder | Team | Active | Moved | Not needed | Last communication |
|-----|--------|------|--------|-------|------------|-------------------|
| api-ecosystem/platform-gateway | api | Platform | | | | |
| custom-codesign-pipeline | custom | DevOps | | | | |
| old-legacy-job | old | Legacy | | | | |

### AFTER (Column D is populated):

| Job | Folder | Team | Active | Moved | Not needed | Last communication |
|-----|--------|------|--------|-------|------------|-------------------|
| api-ecosystem/platform-gateway | api | Platform | **Yes** | | | |
| custom-codesign-pipeline | custom | DevOps | **Yes** | | | |
| old-legacy-job | old | Legacy | | | | |

**Notice:**
- First two jobs are in Active sheet → marked "Yes"
- Third job is NOT in Active sheet → remains empty

---

## 🚨 If Your Output Looks Different

### Problem: Shows 0 matches

```
✓ Found 331 unique active jobs
✓ Read 731 baseline jobs
✓ Marked 0 jobs as Active    ← PROBLEM
✓ 731 jobs remain inactive
```

**What this means:** No job names matched between sheets

**Why it happens:**
- Sheet names are wrong (script didn't read correct data)
- Job name format is different than expected
- Timestamp format is different
- Column positions are wrong

**Solution:** See TROUBLESHOOTING.md → Issue 3

---

### Problem: Shows too many or too few matches

```
✓ Found 331 unique active jobs
✓ Read 731 baseline jobs
✓ Marked 150 jobs as Active    ← Expected ~331
✓ 581 jobs remain inactive
```

**What this means:** Found matches, but fewer than expected

**Why it happens:**
- Some job names don't match exactly (spaces, casing, etc.)
- Active sheet might have fewer jobs than you thought
- Some rows in Active sheet are empty

**Solution:**
1. Check if 150 matches actually makes sense for your data
2. See TROUBLESHOOTING.md → Issue 5
3. Add debug output to see actual job names being compared

---

### Problem: Error message instead of output

```
KeyError: 'Sheet not found: OldSharedJenkinsJobsThatRunLast6Months...'
```

**What this means:** Script can't find the sheet

**Solution:** See TROUBLESHOOTING.md → Issue 2

---

## ✨ Perfect Output Checklist

Your output is correct if:

- [x] Shows "✓ Found XXX unique active jobs" (some number)
- [x] Shows "✓ Read 731 baseline jobs" (or your actual number)
- [x] Shows "✓ Marked XXX jobs as Active" (reasonable number)
- [x] Match rate makes sense (probably 30-60%)
- [x] Active column shows mix of "Yes" and empty
- [x] Number of "Yes" values matches "Marked XXX jobs"

---

## 📊 Sample Data Verification

To verify your results are correct:

### Pick a job you KNOW is active:

1. Find it in the Active sheet (e.g., "api-ecosystem/platform-gateway")
2. Note the name (before the " - timestamp")
3. Find it in the Baseline sheet
4. Check Column D (Active) for that row
5. Should say "Yes" ✓

### Pick a job you KNOW is inactive:

1. Find a job only in Baseline sheet (not in Active)
2. Check Column D (Active) for that row
3. Should be empty/blank ✓

### Count total "Yes" values:

1. In baseline sheet, click Column D header
2. Data → Filter
3. Filter to show only "Yes"
4. Count should match "Marked XXX jobs as Active" ✓

---

## 🎉 Success Criteria

**You're done successfully if:**

1. ✅ Script ran without errors
2. ✅ Summary shows reasonable numbers
3. ✅ Active column has mix of "Yes" and empty
4. ✅ Spot checks verify correctness
5. ✅ Count of "Yes" values matches summary
6. ✅ Data is in Column D of baseline sheet

**Congratulations! Your job matching is complete! 🎊**

---

## 💡 Using the Results

Now that Column D is populated, you can:

- **Filter** to see only active jobs
- **Sort** by Active status
- **Count** active vs inactive jobs
- **Export** the results
- **Create reports** based on active status
- **Make decisions** about which jobs to keep/remove

---

**Need help? See TROUBLESHOOTING.md for detailed solutions to common issues.**
