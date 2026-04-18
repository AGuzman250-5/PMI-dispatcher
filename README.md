# PMI Auto-Dispatcher
**TXANT — San Antonio Hub | Automated PM Assignment Tool**

Built by Andres (UPS)

---

## What This Does

This tool takes your monthly Maximo PM Compliance export and automatically assigns every PM to the right mechanic — no manual sorting, no copy-paste.

- **PMI-04 and PMI-05** are assigned based on the quarterly rotation (who owns which sorter, walk route, and tractor)
- **PMI-06, PMI-06A, PMI-10, PMI-02** are distributed using labor hour balancing, weighted by each mechanic's shift and downtime availability
- The quarter is **auto-detected from the month** — no manual input needed
- Any item that can't be matched is flagged in a **Review Required** tab instead of silently assigned wrong
- Output is a formatted Excel file ready to copy into your dispatch dashboard

---

## How To Use It

**Every month:**

1. Go to the app URL
2. Select the **month and year** you are dispatching for
3. Upload **3 files** (see below)
4. Click **Generate Dispatch**
5. Download the Excel file

> ⚠️ Pull the compliance report between the **20th–25th of the month**. This ensures all WOs have generated without too much bleedover from the next month. The tool automatically filters out any PMs with due dates outside your selected month.

---

## The Three Files You Upload

### ① PM Compliance Export
- **What it is:** The monthly Maximo PM Compliance Detail report
- **How to pull it:** Maximo → PM Compliance Detail → export to Excel
- **Pull dates:** Between the 20th–25th of the month
- **Changes:** Every month

### ② Quarterly Rotation File
- **What it is:** Who owns which sorter, walk route, and tractor this quarter
- **Template:** `quarterly_rotation_template.xlsx` (in this repo)
- **Changes:** Every quarter — see instructions below

### ③ Mechanic Schedule File
- **What it is:** Each mechanic's shift times and which sorts they overlap
- **Template:** `mechanic_schedule_template.xlsx` (in this repo)
- **Changes:** Twice a year after union bid — see instructions below

---

## Updating the Quarterly Rotation

**When:** At the start of each new quarter (Q1=Jan, Q2=Apr, Q3=Jul, Q4=Oct)

**How:**
1. Open `quarterly_rotation_template.xlsx`
2. The rotation shifts down by one row each quarter
3. Move the **last mechanic's row** (Cesar) to the **top** (above Frank)
4. All other rows shift down one position
5. Save the file and upload it next time you run the dispatcher

**Example — going from Q2 to Q3:**

| Before (Q2) | After (Q3) |
|---|---|
| Frank → PS2 | Frank → M5/6 (was Cesar's) |
| Santos → PS1 | Santos → PS2 (was Frank's) |
| Thomas → M1/2 | Thomas → PS1 (was Santos's) |
| ... | ... |
| Cesar → M5/6 | Cesar → M3/4 (was Brian's) |

> The bottom mechanic always wraps to the top. Everyone else shifts down one.

---

## Updating the Mechanic Schedule

**When:** After the union bid (happens twice a year)

**How:**
1. Open `mechanic_schedule_template.xlsx`
2. Update each mechanic's shift times (Sun–Sat columns)
3. Update the **SortOverlap** columns — which sorts does their shift overlap?
   - Valid values: `Preload`, `Day`, `Twilight`, `Night`
   - Leave blank if no overlap
4. Save the file and upload it next time you run the dispatcher

> The sort overlap columns are important — they tell the balancer which mechanics have the most downtime availability for heavy PMIs.

---

## Understanding the Output

The dispatch Excel has three tabs:

### Final Master
All PM assignments grouped by mechanic. Within each mechanic block, PMs are ordered:
1. PMI-04 (sorted by date)
2. PMI-05 (sorted by date)
3. PMI-06 / PMI-02 (sorted by date)
4. PMI-10 (sorted by date)

### Hour Summary
Total estimated hours and PM count per mechanic. Use this to spot if someone is loaded too heavy before you finalize.

### Review Required ⚠️
Items flagged for manual review. Two types:
- **Hours not found in master plan** — PM was assigned but the labor estimate couldn't be matched. Verify the time estimate is reasonable before dispatching.
- **No pattern match** — the description didn't match any ownership rule. These need manual assignment.

> Flagged items are still assigned to a mechanic as a best guess — don't ignore this tab, just verify those specific items.

---

## File Structure

```
pmi_dispatcher/
├── app.py                          # Streamlit web app
├── dispatcher_core.py              # Assignment engine
├── PMI_12MoCal_master.xlsx         # Bundled 12-month labor hours reference
├── quarterly_rotation_template.xlsx # Template — update each quarter
├── mechanic_schedule_template.xlsx  # Template — update after union bid
├── requirements.txt                # Python dependencies
└── .streamlit/
    └── config.toml                 # App theme config
```

---

## Mechanic List

The tool is configured for the following 11 mechanics in this fixed order:

Frank · Santos · Thomas · Ben · Hugo · Robert · Devin · Steven · Rafa · Brian · Cesar

> If the mechanic roster changes, the rotation and schedule templates need to be updated to match. The mechanic names in both files must match exactly (case-sensitive).

---

## Known Limitations

- **Labor hours** come from the bundled 12-month master plan. If Maximo updates the estimated durations, the master plan file inside the app needs to be updated. Contact whoever manages the GitHub repo.
- **The master plan descriptions** sometimes get truncated on export — if you see many items in the Review Required tab flagged for unknown hours, try re-exporting the master plan with wider columns.
- **The tool is not actively maintained.** It was built as a handoff tool. If something breaks after a major Maximo update or roster change, the logic is in `dispatcher_core.py` and is documented throughout.

---

## Technical Notes

Built with Python · Streamlit · openpyxl · pandas

Deployed on Streamlit Community Cloud (free tier)

*Built by Andres Guzman 2026 — For TXANT BaSE Department*
