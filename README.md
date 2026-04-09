# DSPG 2026 Curriculum

Interactive curriculum book for the Data Science for the Public Good program at Virginia Tech, Summer 2026.

**Live site:** `https://pdahal23.github.io/dspg-2026-curriculum/dspg_curriculum_2026.html`

---

## Files

| File | Purpose |
|------|---------|
| `curriculum-2026.xlsx` | Source of truth — edit this to update the curriculum |
| `generate_curriculum.py` | Script that reads the Excel and generates the HTML |
| `dspg_curriculum_2026.html` | Auto-generated output — never edit this directly |

---

## How to update

1. Pull the latest version
   ```bash
   git pull
   ```

2. Edit `curriculum-2026.xlsx` and save

3. Regenerate the HTML
   ```bash
   python generate_curriculum.py
   ```

4. Commit and push
   ```bash
   git add .
   git commit -m "brief description of what changed"
   git push
   ```

The live site updates within about a minute of pushing.

---

## Requirements

- Python 3
- openpyxl (`pip install openpyxl`)
- Git

---

## Excel structure

**Master sheet** — one row per week: `Week · Topics · Important Milestones · Deliverables`

**Week sheets (Weeks 1–3)** — day-by-day layout:
- Col A: date (e.g. `Mon, Jun 8`) to start a new day
- Col A empty: sub-event row — fill cols B (activity), C (time), E (who)
- Col A = `Deliverables:` to add a week-level deliverable
- Days with text like `Memorial Day` or `no activities` are auto-detected as holidays

**Week sheets (Weeks 4–7)** — topic list layout:
- One row per topic: `Topic title · Instructor · Notes · Timing`

**Card colors** are assigned automatically based on keywords:
- Workshop / Training / Lecture → purple
- Deliverable / Deadline / Submit → green
- Everything else → blue

---

## Program links

- [DSPG Program](https://aaec.vt.edu/academics/undergraduate/dspg.html)
- [Virginia Tech](https://www.vt.edu/)
