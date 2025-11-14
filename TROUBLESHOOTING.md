# Troubleshooting Guide - Excel Job Matcher

This guide covers common issues and their solutions.

---

## 🔍 Quick Diagnostic Checklist

Before diving into specific issues, check these basics:

- [ ] Using **Microsoft 365 Excel Online** (web version)?
- [ ] Both sheets exist in the workbook?
- [ ] Sheet names are **exactly correct** (case-sensitive)?
- [ ] Copied the **complete script** (all lines)?
- [ ] Python cell is **selected** when you press Enter?

---

## Common Issues

### Issue 1: "Python" button not visible in Insert tab

**Symptoms:**
- Can't find Python in the Insert tab
- Insert tab doesn't show Python option

**Solutions:**

1. **Check Excel version:**
   - Must be Microsoft 365 Excel **Online** (web version)
   - Desktop Excel might not have Python support
   - Open your file in Excel Online: office.com → Excel

2. **Try Formulas tab:**
   - Go to **Formulas** tab
   - Look for **Insert Python** button
   - Click it to insert a Python cell

3. **Check subscription:**
   - Python in Excel requires Microsoft 365 subscription
   - Not available in all regions yet
   - Check: https://support.microsoft.com/en-us/office/python-in-excel

4. **Enable Python in Excel:**
   - File → Options → Formulas → Python in Excel
   - Make sure it's enabled

---

### Issue 2: Error - "Sheet not found" or "Invalid sheet name"

**Symptoms:**
```
KeyError: 'Sheet not found'
```
or
```
Invalid sheet name: 'OldSharedJenkinsJobsThatRunLast6Months...'
```

**Solutions:**

1. **Get exact sheet name:**
   - Right-click on the sheet tab at the bottom
   - Select "Rename" to see the full name
   - Or hover over the tab to see the complete name
   - Copy it exactly

2. **Update the script:**
   ```python
   BASELINE_SHEET = "YourExactSheetNameHere"
   ACTIVE_SHEET = "YourExactActiveSheetNameHere"
   ```
   - Keep the quotes
   - Match case exactly
   - Include all special characters

3. **Check for special characters:**
   - If sheet name has apostrophes or quotes, be careful
   - Excel might need escaping
   - Try removing special characters from sheet name first

4. **Simplify sheet names (if possible):**
   - Rename sheets to something simple like "Baseline" and "Active"
   - Update script accordingly
   - This avoids special character issues

---

### Issue 3: Script runs but shows 0 matches

**Symptoms:**
```
✓ Found 331 unique active jobs
✓ Read 731 baseline jobs
✓ Marked 0 jobs as Active    ← Problem: Should be ~331
```

**Solutions:**

1. **Check data format in Active sheet:**
   - Open the Active sheet
   - Look at Column A, Row 4
   - Should look like: "Job Name - Mon Nov 03 12:07:14 UTC 2025"
   - If format is different, the script needs adjustment

2. **Check data format in Baseline sheet:**
   - Open the Baseline sheet
   - Look at Column A, Row 2
   - Should have job names without timestamps
   - Should match the part before " - " in Active sheet

3. **Check for leading/trailing spaces:**
   - Job names might have extra spaces
   - Try this fix in the script (add after line 35):
   ```python
   # Add this line to strip spaces
   active_df['Job_Clean'] = active_df['Job_Clean'].str.strip()
   ```

4. **Check case sensitivity:**
   - Job names are case-sensitive
   - "Job" is different from "job"
   - Make sure casing matches between sheets

5. **Verify timestamp format:**
   - If timestamps are in different format, adjust the split:
   ```python
   # Current: splits on " - "
   # If your format is different, adjust this line:
   active_df['Job_Clean'] = active_df.iloc[:, 0].astype(str).str.split(' - ', n=1).str[0]
   ```

---

### Issue 4: Script shows error "xl is not defined"

**Symptoms:**
```
NameError: name 'xl' is not defined
```

**Solutions:**

1. **Wrong environment:**
   - This error means you're NOT running in Python in Excel
   - `xl()` is only available in Excel's Python environment
   - Don't run this script in regular Python
   - Must run in Excel: Insert → Python

2. **Copied to wrong place:**
   - Make sure you pasted in a Python cell in Excel
   - Not in a regular Excel cell
   - Not in a separate Python IDE

---

### Issue 5: Wrong number of matches (not 331)

**Symptoms:**
```
✓ Marked 215 jobs as Active    ← Should be ~331
```
or
```
✓ Marked 450 jobs as Active    ← Should be ~331
```

**Diagnosis steps:**

1. **Check Active sheet row count:**
   - Go to Active sheet
   - Count actual job entries (not empty rows)
   - Might not be exactly 331
   - This is NORMAL if your data is different

2. **Check for empty rows:**
   - Active sheet might have empty rows mixed in
   - This is handled by the script
   - Actual match count might be less than 331

3. **Check for duplicate jobs:**
   - Some jobs might appear multiple times
   - Script uses set() which removes duplicates
   - This is correct behavior

4. **Verify it's actually wrong:**
   - Look at the summary output
   - Check "Match rate"
   - Should be around 45% (331/731)
   - If it's close, the result is probably correct

**If truly wrong:**

Check baseline and active sheets manually:
```python
# Add these debug lines after line 35:
print("Sample active jobs:")
print(active_df['Job_Clean'].head(10))
print("\nSample baseline jobs:")
# (add after reading baseline)
print(baseline_df['Job'].head(10))
```

This shows you the actual job names being compared.

---

### Issue 6: "ModuleNotFoundError: No module named 'pandas'"

**Symptoms:**
```
ModuleNotFoundError: No module named 'pandas'
```

**Solutions:**

1. **Wrong environment (most likely):**
   - This means you're running in regular Python
   - NOT in Excel's Python environment
   - **Solution:** Run in Excel, not in terminal/IDE

2. **If truly in Excel:**
   - Excel Python should have pandas pre-installed
   - Try restarting Excel
   - Check Python in Excel is enabled
   - Contact Microsoft support if issue persists

---

### Issue 7: Script takes very long or seems stuck

**Symptoms:**
- Script runs for more than 30 seconds
- Excel seems frozen
- No output appears

**Solutions:**

1. **Wait a bit longer:**
   - First run might take 10-15 seconds
   - Excel is loading Python environment
   - Be patient

2. **Check data size:**
   - 731 rows should be fast
   - If your sheets are much larger, it might take longer
   - Still shouldn't take more than 30 seconds

3. **Restart:**
   - Cancel the script (if possible)
   - Refresh the Excel page
   - Try again

4. **Check for infinite loop:**
   - Make sure you copied the script correctly
   - No modifications that might cause loops
   - Re-copy the original script

---

### Issue 8: Results look correct but won't copy to Column D

**Symptoms:**
- Script runs successfully
- Summary shows correct numbers
- But can't copy results to baseline sheet

**Solutions:**

1. **Use alternative script:**
   - Try `excel_job_matcher_full_output.py` instead
   - Returns complete table
   - Might be easier to work with

2. **Manual copy:**
   - Select the output Active column
   - Copy (Ctrl+C)
   - Go to baseline sheet
   - Click cell D2
   - Paste (Ctrl+V)

3. **Use formula reference:**
   - Excel might allow you to reference the Python output
   - Try formula: `=PythonCell.Active`
   - (Replace PythonCell with actual cell reference)

---

### Issue 9: Some jobs marked wrong (false positives/negatives)

**Symptoms:**
- Most matches are correct
- But a few are wrong
- Known active jobs not marked
- Or inactive jobs marked as active

**Diagnosis:**

1. **Check specific job names:**
   - Find the problematic job in Active sheet
   - Find same job in Baseline sheet
   - Compare names character-by-character
   - Even one character difference will cause mismatch

2. **Common causes:**
   - Extra spaces
   - Different casing (Job vs job)
   - Special characters (/ vs \)
   - Different folder structure

3. **Check timestamp removal:**
   - Make sure " - timestamp" is being removed correctly
   - Look at the "Sample active jobs" if you added debug output

**Solutions:**

Add fuzzy matching (advanced):
```python
# Replace the exact matching with case-insensitive
baseline_df['Active'] = baseline_df['Job'].str.lower().apply(
    lambda job: 'Yes' if job in {j.lower() for j in active_jobs_set} else ''
)
```

Or trim spaces more aggressively:
```python
# Add after cleaning job names
active_df['Job_Clean'] = active_df['Job_Clean'].str.strip().str.lower()
baseline_df['Job'] = baseline_df['Job'].str.strip().str.lower()
```

---

### Issue 10: "Index out of range" or "KeyError"

**Symptoms:**
```
IndexError: single positional indexer is out-of-bounds
```
or
```
KeyError: 'Job'
```

**Solutions:**

1. **Check column positions:**
   - Script assumes Column A has job names
   - If your columns are different, adjust the script
   - Check header row matches expectations

2. **Check row ranges:**
   - Script reads A4:A340 for active sheet
   - If your data is in different rows, adjust:
   ```python
   # Change this line:
   active_jobs_range = f"'{ACTIVE_SHEET}'!A4:A340"
   # To your actual range, e.g.:
   active_jobs_range = f"'{ACTIVE_SHEET}'!A2:A500"
   ```

3. **Check baseline range:**
   - Script reads A1:G732
   - If you have more/fewer rows, adjust:
   ```python
   # Change this line:
   baseline_range = f"'{BASELINE_SHEET}'!A1:G732"
   # To your actual range
   ```

---

## 🧪 Testing Your Setup

Try this minimal test to check if Python in Excel is working:

### Test 1: Basic Python
```python
# Paste this in a Python cell
print("Python is working!")
42
```
Expected output: Should see "Python is working!" and 42

### Test 2: Pandas
```python
# Test if pandas is available
import pandas as pd
df = pd.DataFrame({'A': [1, 2, 3]})
df
```
Expected output: Should see a table with values 1, 2, 3

### Test 3: Excel function
```python
# Test xl() function
# Change SheetName to your actual sheet
xl("'SheetName'!A1:A5")
```
Expected output: Should see first 5 cells from column A

---

## 📞 Still Having Issues?

If none of these solutions work:

1. **Check the basics again:**
   - Sheet names correct?
   - Complete script copied?
   - Running in Excel Online?

2. **Simplify:**
   - Try renaming sheets to simple names: "Baseline", "Active"
   - Try with smaller row range first (e.g., A1:A10)
   - Build up from there

3. **Get more info:**
   - Add print statements to see what's happening
   - Check each step of the script
   - Share the error message for help

4. **Alternative approach:**
   - If Python in Excel isn't working, consider:
   - Power Query (Data → Get Data → Combine Queries)
   - VLOOKUP formulas (more manual but reliable)
   - Export to CSV and use regular Python

---

## 📋 Debugging Checklist

Copy this checklist when asking for help:

```
[ ] Excel version: Microsoft 365 Online
[ ] Baseline sheet name: _________________
[ ] Active sheet name: _________________
[ ] Number of baseline jobs: _____
[ ] Number of active jobs: _____
[ ] Error message (if any): _________________
[ ] Script runs: Yes / No
[ ] Output appears: Yes / No
[ ] Number of matches found: _____
[ ] Expected number of matches: _____
```

---

## 🎯 Prevention Tips

To avoid issues in the future:

1. **Keep sheet names simple** - avoid special characters
2. **Document your changes** - if you modify the script
3. **Test with small data first** - before running on full dataset
4. **Save often** - Excel Online auto-saves, but be safe
5. **Keep a backup** - of your original workbook

---

**Good luck! You've got this! 💪**
