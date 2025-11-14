# Excel Job Matcher for Python in Excel

**Automatically mark active Jenkins jobs in your baseline Excel sheet using Microsoft 365's built-in Python**

---

## 📁 Files in This Package

### Main Script (USE THIS)
- **`excel_job_matcher.py`** - The main Python script to copy-paste into Excel
  - Returns just the Active column for easy copying to Column D
  - Most straightforward approach

### Alternative Script
- **`excel_job_matcher_full_output.py`** - Alternative version
  - Returns the complete baseline table with Active column
  - Use if you prefer to see all columns together

### Documentation
- **`QUICK_START.md`** - 5-step quick start guide (2 minutes)
- **`USER_GUIDE.md`** - Comprehensive step-by-step guide with troubleshooting
- **`README.md`** - This file

### Testing
- **`test_logic.py`** - Test script to validate the matching logic (optional)

---

## 🚀 Quick Start

1. **Open your Excel workbook** in Microsoft 365 Excel Online
2. **Insert → Python** in the ribbon
3. **Copy all code** from `excel_job_matcher.py`
4. **Paste into Python cell**
5. **Adjust sheet names** if they don't match (see lines 19-22)
6. **Press Enter** to run
7. **Copy results** to Column D in your baseline sheet

**See `QUICK_START.md` for more details**

---

## 📊 What It Does

This script:
1. Reads 331 active jobs from your "Active Jobs" sheet
2. Removes timestamps from job names (e.g., "Job - Mon Nov 03..." → "Job")
3. Reads 731 baseline jobs from your "Baseline" sheet
4. Matches job names between the two sheets
5. Marks matching jobs with "Yes" in the Active column
6. Shows summary statistics

**Expected Result:**
- ~331 jobs marked as "Yes"
- ~400 jobs remain empty (inactive)
- Match rate: ~45%

---

## 🎯 Requirements

- **Microsoft 365 Excel** (web version with Python support)
- **Two sheets in your workbook:**
  - Baseline sheet (731 jobs)
  - Active sheet (331 jobs with timestamps)

---

## 📋 Your Data Structure

### Sheet 1: Baseline (All Jobs)
- Row 1: Headers
- Rows 2-732: Data (731 jobs)
- Column A: Job name
- Column D: Active (currently empty - will be populated)

### Sheet 2: Active Jobs
- Rows 1-3: Metadata/headers
- Rows 4-340: Job data
- Column A: Job name with timestamp format "Job - Mon Nov 03 12:07:14 UTC 2025"

---

## ⚙️ Configuration

The only thing you might need to change is the sheet names at the top of the script:

```python
BASELINE_SHEET = "OldSharedJenkinsJobsThatRunLast6Months_Nov_2025__all_-_baseline__"
ACTIVE_SHEET = "OldSharedJenkinsJobsThatRunLast6Months_Nov_2025__active__-2"
```

Replace these with your actual sheet names (keep the quotes).

---

## 🔧 Troubleshooting

### Can't find Python in Excel?
- Make sure you're using Microsoft 365 Excel Online (web version)
- Check: Insert tab → Python section
- Or try: Formulas tab → Insert Python

### "Sheet not found" error?
- Check your sheet names are exactly correct (case-sensitive)
- Right-click sheet tab to see exact name
- Update the sheet names in the script

### Wrong number of matches?
- Check that your Active sheet has jobs in Column A starting from row 4
- Verify jobs have timestamps in format "Job - timestamp"
- Check that job names match exactly between sheets

**See `USER_GUIDE.md` for complete troubleshooting**

---

## 📝 Usage Notes

- **Zero configuration needed** except potentially adjusting sheet names
- **No Python installation required** - runs directly in Excel
- **Copy-paste ready** - just paste and run
- **Works entirely in the cloud** - uses Microsoft's Python runtime
- **No external files needed** - all data stays in your workbook

---

## 🎓 For First-Time Python in Excel Users

**Don't worry if you've never used Python before!** This script is designed to be copy-paste ready.

1. You don't need to understand Python
2. You don't need to install anything
3. Just copy, paste, and run
4. Follow the step-by-step guide in `USER_GUIDE.md`

---

## ✅ Verification Checklist

After running:
- [ ] Column D has ~331 "Yes" values
- [ ] Remaining rows are empty/blank
- [ ] Spot-check: Active jobs show "Yes"
- [ ] Spot-check: Inactive jobs are blank
- [ ] Summary stats match expectations
- [ ] Save your workbook

---

## 🆘 Need Help?

1. Start with **`QUICK_START.md`** for basic steps
2. Read **`USER_GUIDE.md`** for detailed instructions
3. Check the troubleshooting section
4. Verify your sheet names and data structure

---

## 📄 License & Usage

This script is provided as-is for your use. Feel free to modify it for your needs.

**Common modifications:**
- Change sheet names (variables at top)
- Adjust row ranges if your data is different
- Modify matching logic if needed

---

## 🔍 How It Works (Technical)

For those interested in the details:

1. **Reads active jobs** using Excel's `xl()` function
2. **Cleans data** using pandas string operations (`.str.split()`)
3. **Creates a set** for fast O(1) lookup performance
4. **Matches jobs** using pandas `.apply()` with lambda function
5. **Returns results** as Excel-compatible DataFrame output

The entire process takes just a few seconds even with 731 jobs.

---

## 🎉 Success!

Once complete, you'll have:
- ✅ Active column populated in your baseline sheet
- ✅ 331 jobs marked as active
- ✅ Clear summary statistics
- ✅ Easy verification of results

**Happy matching! 🚀**
