# Excel Python Job Matcher - User Guide

**For Microsoft 365 Excel Users (Cloud Version)**

This guide will walk you through using Python directly in Excel to automatically mark active jobs in your baseline sheet. **No Python experience needed!**

---

## Prerequisites

✅ **Microsoft 365 Excel** (cloud version with Python support)
✅ **Your workbook** with two sheets:
   - Baseline sheet (all jobs)
   - Active sheet (jobs with timestamps)

---

## Step-by-Step Instructions

### Step 1: Open Your Workbook in Excel Online

1. Open your workbook in **Microsoft 365 Excel** (web version)
2. Make sure you can see both sheets at the bottom:
   - The baseline sheet (all 731 jobs)
   - The active jobs sheet (331 jobs with timestamps)

### Step 2: Check Your Sheet Names

⚠️ **IMPORTANT**: You need to know the exact names of your sheets.

1. Look at the sheet tabs at the bottom of Excel
2. **Right-click** on each sheet tab and note the exact name
3. Common names might be:
   - Baseline: `OldSharedJenkinsJobsThatRunLast6Months_Nov_2025__all_-_baseline__`
   - Active: `OldSharedJenkinsJobsThatRunLast6Months_Nov_2025__active__-2`

**Write down your sheet names - you'll need them in Step 4!**

---

### Step 3: Insert Python Cell in Excel

1. Go to the **Insert** tab in the Excel ribbon
2. Look for the **Python** section (it might say "Python" or "PY")
3. Click **Python**
4. A new cell will appear where you can type Python code

> 💡 **Tip**: If you don't see "Python" in the Insert tab, you might need to enable it:
> - Go to **Formulas** tab → **Insert Python**
> - Or check if your Microsoft 365 subscription includes Python in Excel

---

### Step 4: Paste the Python Code

1. **Open the file**: `excel_job_matcher.py`
2. **Select all the code** (Ctrl+A or Cmd+A)
3. **Copy it** (Ctrl+C or Cmd+C)
4. **Go back to Excel** and click in the Python cell you created
5. **Paste the code** (Ctrl+V or Cmd+V)

---

### Step 5: Adjust Sheet Names (If Needed)

Look at the top of the code you just pasted. You'll see these lines:

```python
# Sheet with all 731 baseline jobs (the one you want to update)
BASELINE_SHEET = "OldSharedJenkinsJobsThatRunLast6Months_Nov_2025__all_-_baseline__"

# Sheet with 331 active jobs (with timestamps)
ACTIVE_SHEET = "OldSharedJenkinsJobsThatRunLast6Months_Nov_2025__active__-2"
```

**If your sheet names are different:**
1. Replace the text inside the quotes with YOUR exact sheet names
2. Keep the quotes around the sheet name
3. Be very precise - sheet names are case-sensitive!

**Example:**
If your baseline sheet is named "AllJobs", change it to:
```python
BASELINE_SHEET = "AllJobs"
```

---

### Step 6: Run the Python Code

1. Make sure the Python cell is selected (click on it)
2. Press **Enter** or click the **Run** button (▶️) if you see one
3. **Wait a few seconds** while Excel processes the data

You'll see output that looks like:
```
✓ Found 331 unique active jobs
✓ Read 731 baseline jobs
✓ Marked 331 jobs as Active
✓ 400 jobs remain inactive

====================================
SUMMARY
====================================
Total baseline jobs: 731
Total active jobs: 331
Matched jobs: 331
Match rate: 45.3%
====================================
```

---

### Step 7: Copy Results to Your Baseline Sheet

After the code runs, you'll see a column of "Yes" values and empty cells below the summary.

**To apply these to your baseline sheet:**

**Option A: Automatic (Preferred)**
1. The Python cell output should automatically populate
2. You can reference this output in your baseline sheet

**Option B: Manual Copy**
1. Click on the Python cell output
2. Select all the "Active" column values
3. Copy them (Ctrl+C)
4. Go to your baseline sheet
5. Click on cell **D2** (first data row in Active column)
6. Paste (Ctrl+V)

---

### Step 8: Verify the Results

Let's make sure everything worked correctly!

1. **Go to your baseline sheet**
2. **Look at Column D** (Active column)
3. You should see:
   - About **331 rows with "Yes"**
   - About **400 rows that are empty or blank**

**Quick verification checks:**

✅ **Count the "Yes" values:**
- Click on a cell in column D
- Use Excel's filter: Data → Filter → Select "Yes"
- Check if the count matches the summary (should be ~331)

✅ **Spot check some jobs:**
- Find a job you know is active in the Active sheet
- Search for the same job name in the Baseline sheet
- Column D should show "Yes" for that job

✅ **Check inactive jobs:**
- Find a job that's NOT in the Active sheet
- It should NOT have "Yes" in column D

---

## What the Script Does (Behind the Scenes)

For those curious about how it works:

1. **Reads active jobs** from your Active sheet (rows 4-340)
2. **Cleans job names** by removing the timestamp part (e.g., "Job - Mon Nov 03..." → "Job")
3. **Creates a list** of all unique active job names (331 jobs)
4. **Reads baseline jobs** from your Baseline sheet (731 jobs)
5. **Compares each baseline job** against the active jobs list
6. **Marks matches** with "Yes" in the Active column
7. **Shows you the results** with a summary

---

## Troubleshooting

### Problem: "Python not available" or can't find Insert → Python

**Solution:**
- Make sure you're using **Microsoft 365 Excel Online** (web version)
- Check your Microsoft 365 subscription includes Python in Excel
- Try: Formulas tab → Insert Python

---

### Problem: Error says "Sheet not found"

**Solution:**
- Double-check your sheet names in Step 5
- Sheet names must be EXACT (case-sensitive)
- Right-click the sheet tab to see the exact name
- Make sure there are no extra spaces in the sheet name

---

### Problem: Script runs but no "Yes" values appear

**Solution:**
- Check that your Active sheet has jobs in column A starting from row 4
- Check that jobs have timestamps in format "Job - Mon Nov 03..."
- Verify that job names in both sheets match exactly

---

### Problem: Wrong number of matches (not 331)

**Solution:**
- This might be normal if your data is slightly different
- Check the summary output - it tells you how many matches were found
- Verify your Active sheet actually has 331 jobs (some might be empty rows)

---

### Problem: Code shows error when running

**Solution:**
1. Make sure you copied the ENTIRE code (all 100+ lines)
2. Check that sheet names are correct
3. Make sure your sheets have data in the expected columns
4. Try running the code again (sometimes Excel needs a refresh)

---

## Tips for Success

💡 **Before running:**
- Make a backup copy of your workbook
- Write down your exact sheet names
- Check that both sheets have data

💡 **While running:**
- Be patient - it might take 5-10 seconds
- Don't click away from Excel while it's processing

💡 **After running:**
- Save your workbook to keep the results
- Verify with spot checks before using the data

---

## Need Help?

If you encounter issues:

1. **Check your sheet names** - this is the #1 cause of errors
2. **Verify your data format** - make sure columns match the expected structure
3. **Try running again** - sometimes Excel just needs a refresh
4. **Check the output messages** - they tell you what went wrong

---

## Advanced: Understanding the Output

When the script runs, you'll see several messages:

```
✓ Found 331 unique active jobs
```
↳ This means the script successfully read and cleaned 331 job names from your Active sheet

```
✓ Read 731 baseline jobs
```
↳ This means the script successfully read all jobs from your Baseline sheet

```
✓ Marked 331 jobs as Active
```
↳ This means 331 jobs in your baseline matched jobs in the active list

```
Match rate: 45.3%
```
↳ This means 45.3% of your baseline jobs are currently active

---

## Summary Checklist

Before you start:
- [ ] Open workbook in Excel Online (Microsoft 365)
- [ ] Note exact sheet names
- [ ] Have both sheets visible

Steps:
- [ ] Insert → Python
- [ ] Paste the code
- [ ] Adjust sheet names if needed
- [ ] Run the code
- [ ] Wait for results
- [ ] Copy to Column D (Active) if needed
- [ ] Verify with spot checks
- [ ] Save your workbook

---

## Questions?

This script is designed to be copy-paste ready with minimal configuration. The only thing you should need to change is the sheet names at the top of the code.

**Common customizations:**

1. **Different sheet names**: Edit `BASELINE_SHEET` and `ACTIVE_SHEET` variables
2. **Different column positions**: The code assumes Column A for job names
3. **Different row ranges**: Adjust the range numbers (A4:A340, etc.)

---

**That's it! You're ready to match your jobs. Good luck! 🚀**
