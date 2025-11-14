# Quick Start - 5 Steps to Match Jobs

**Total time: 2 minutes**

---

## 1. Check Your Sheet Names

Look at the sheet tabs at the bottom of Excel and note the names of:
- Your **baseline sheet** (all 731 jobs)
- Your **active jobs sheet** (331 jobs with timestamps)

---

## 2. Insert Python in Excel

1. Click **Insert** tab → **Python**
2. A Python cell will appear

---

## 3. Copy & Paste the Code

1. Open `excel_job_matcher.py`
2. Select all (Ctrl+A) and copy (Ctrl+C)
3. Paste into the Python cell in Excel

---

## 4. Update Sheet Names (If Needed)

Look at the top of the code for these lines:

```python
BASELINE_SHEET = "OldSharedJenkinsJobsThatRunLast6Months_Nov_2025__all_-_baseline__"
ACTIVE_SHEET = "OldSharedJenkinsJobsThatRunLast6Months_Nov_2025__active__-2"
```

Replace with YOUR exact sheet names (keep the quotes).

---

## 5. Run It!

1. Press **Enter** or click **Run**
2. Wait 5-10 seconds
3. You'll see a summary showing how many jobs were matched
4. The Active column values will appear below
5. Copy these to Column D in your baseline sheet

---

## ✅ Verify

- Column D should have about **331 "Yes"** values
- Remaining rows should be empty
- Spot-check a few jobs to confirm accuracy

---

**Done! 🎉**

See `USER_GUIDE.md` for detailed instructions and troubleshooting.
