# Credit Hours Prep — App & Code Guide (No Certificates Yet)

This project prepares per-attendee credit hours from two input spreadsheets (event hours and registration roster). It is designed for **non-technical users** via a simple **GUI app**, but also supports a **CLI** for power users. This README explains **how to use the app** *and* breaks down **how the code works** so you can maintain or extend it confidently.

---

## TL;DR (GUI Quick Start)
1. **Build the app once** (Mac):
   ```bash
   pip install pyinstaller pandas openpyxl
   cd ~/Desktop/PDH
   pyinstaller --windowed --onefile --name "Credit Hours Prep" run_me_nocerts.py
   ```
2. **Open** `~/Desktop/PDH/dist/Credit Hours Prep.app` (right‑click → Open the first time if Gatekeeper warns).
3. **Generate Proposals → Review → Apply Decisions.**
   - Review and edit **`proposed_matches.xlsx`** (use **Pick** column: `1/2/3` or an email). 
   - Click **Apply Decisions** to create **`master_list.xlsx`**.

> Final deliverable for certificates: **`master_list.xlsx`** (Name/Email/Total Hours and optional Category/Subcategory).

---

## Files & Schemas

### File A — Event Hours (Excel)
- **Columns (first 3, in order):**
  1. Full Name
  2. Credit Hours
  3. Event Name

### File B — Registration Roster (Excel)
- **Columns (first 7, in order):**
  1. Category
  2. Subcategory
  3. Full Name
  4. Country
  5. Email
  6. CC Email
  7. First Conference (Yes/No)

### Manual Overrides (Optional, CSV)
Columns (case-insensitive accepted):
- `FullName_A`
- `Override_FullName_B` *(optional exact name from File B)*
- `Override_Email` *(optional direct email; takes precedence)*

---

## Build the App (Mac)
```bash
pip install pyinstaller pandas openpyxl
cd ~/Desktop/PDH
pyinstaller --windowed --onefile --name "Credit Hours Prep" run_me_nocerts.py
```
- Output: `~/Desktop/PDH/dist/Credit Hours Prep.app`
- First launch: right‑click → **Open** → **Open** to bypass Gatekeeper once.
- You can drag the app into **Applications** if you like.

> **Tip:** You can also run the GUI directly without packaging:
> ```bash
> python ~/Desktop/PDH/run_me_nocerts.py --gui
> ```

---

## Using the App (Step‑by‑Step)
Open **Credit Hours Prep.app**. The window has fields and three buttons.

### Step 1 — Select Inputs
- **File A (PDH Data):** pick your event hours Excel.
- **File B (Registration):** pick your roster Excel.
- **Output Folder:** choose where results go (defaults to `~/Desktop/output_nocerts`).
- *(Optional)* **Decisions File:** if you already edited a previous `proposed_matches.xlsx`.
- *(Optional)* **Overrides CSV:** if you maintain a manual mapping file.
- *(Optional)* **Category:** filter final results to a category (e.g., `User`).
- *(Optional)* **Min Match:** fuzzy threshold (default `0.85`).

### Step 2 — Generate Proposals
Click **Generate Proposals**. The app will:
1. Deduplicate exact duplicates in File A (same Name+Hours+Event).
2. Collapse File B to one row per **Email** (logging any email collisions).
3. Compute **Top 3** name matches for each unique name in File A using enhanced matching (details below).
4. Write **`proposed_matches.xlsx`** and open it.

### Step 3 — Review `proposed_matches.xlsx`
Open the file the app created and focus on these columns:
- **FullName_A** — name from File A.
- **Top1_Name_B / Top1_Email / Top1_Score** — best guess.
- **Top2_… / Top3_…** — alternates (hidden when a certain match is already clear).
- **Suggested_Email** — what will be used if you accept without changes (Top1).
- **Certain** — `TRUE` only when **letters‑only equal** (ignores spaces/punctuation **without** reordering) *and* Email present. These rows are **green** and prefilled with **Pick=1** and **Decision=ACCEPT**.
- **Decision** — set to `ACCEPT` or `REJECT`. Blank means undecided.
- **Pick** — enter `1`, `2`, `3`, **or a manual email**. If you enter a value here and leave `Decision` blank, it is treated as **accepted**.
- **Chosen_Email** — optional; you can still use this, but **Pick** is simpler.

**Rules:**
- If **Top2** is correct → set **Pick = 2** and either set **Decision = ACCEPT** or leave Decision blank (Pick implies accept).
- If **Top3** is correct → **Pick = 3**.
- If none is right but you know the email → type the email in **Pick**; the app will pull full details from File B.
- If you need to exclude → **Decision = REJECT**.
- Save and close the Excel when done.

### Step 4 — Apply Decisions & Build Final List
Back in the app, click **Apply Decisions**. The app will:
1. Read your `proposed_matches.xlsx` (or decisions file you chose).
2. Apply **ACCEPT** decisions, resolving the chosen identity via **Pick/Chosen_Email/Suggested_Email**.
3. Enrich from File B (Category/Subcategory/Country/CC Email/First Conference) using the selected Email.
4. Deduplicate by **(Email, EventName)** to prevent over‑counting.
5. Aggregate total hours **by Email**; compute **`master_list.xlsx`** and **`master_list.csv`**.
6. Write diagnostics: `joined_events.xlsx`, `unmatched_needs_email.xlsx`, etc.

> **Workflow Summary:** Generate Proposals → Edit `proposed_matches.xlsx` → Apply Decisions → Use `master_list.xlsx`.

---

## Command Line Usage (Advanced)
```bash
python run_me_nocerts.py \
  --file_a "/path/to/FileA.xlsx" \
  --file_b "/path/to/FileB.xlsx" \
  --out_dir "/path/to/output" \
  --category "User" \
  --min_match 0.85 \
  --overrides_csv "/path/to/manual_overrides.csv" \
  --decisions_path "/path/to/proposed_matches.xlsx"
```
Launch GUI from CLI:
```bash
python run_me_nocerts.py --gui
```

---

## Output Files
- **`proposed_matches.xlsx`** — your review sheet (Decision/Pick/Chosen_Email).
- **`joined_events_pre_overrides.xlsx`** — event rows before decisions/overrides (for audit).
- **`joined_events.xlsx`** — event rows after decisions/overrides, with Confidence & MatchSource.
- **`duplicates_removed_same_email_event.xlsx`** — duplicates dropped for same **Email+Event**.
- **`unmatched_needs_email.xlsx`** — rows without Email (fix via Pick or overrides).
- **`master_list.xlsx` / `master_list.csv`** — **final totals** (DisplayName, Email, TotalCreditHours, Category, Subcategory).
- **`excluded_by_category.xlsx`** — if you used `--category`.

---

## How the Code Works (Architecture)
**Main script:** `run_me_nocerts.py`

### Pipeline Overview
1. **Read inputs** (File A & B) and standardize columns by position.
2. **Dedupe File A** exact duplicates (Name+Hours+Event).
3. **Collapse File B** to a canonical row per **Email** (log duplicates by email).
4. **Proposals:** for each unique name in A, compute **Top 3** matches from B and export `proposed_matches.xlsx`.
5. **Decisions & Overrides:** merge user decisions (Decision/Pick/Chosen_Email) and optional overrides.
6. **Event‑level dedupe:** drop duplicate **(Email, Event)** to prevent double‑counting.
7. **Aggregate** to `master_list` by Email (sum hours, attach roster attributes).

### Matching & Confidence (Key Functions)
- **`normalize_name(name)`**: lowercases, strips punctuation, collapses spaces, maps common nicknames (e.g., `mike`→`michael`).
- **`composite_name_score(a, b)`**: blended score of difflib similarity, token Jaccard (order‑insensitive), letters‑only similarity, and initials.
- **`is_spacing_punct_equal(a, b)`**: letters‑only equality **without** reordering. If true and roster Email exists → row is **Certain** (green) and can be auto‑accepted.
- **`is_absolute_name_match(a, b)`**: broader equality that also accepts **token permutations** up to 4 tokens (handles concatenated names like `Jangwanjae` ↔ `Wan Jae Jang`). These are **not** auto‑green because order was changed; they still require human review.
- **`top_k_matches(name_a, df_b, k)`**: ranks matches and ensures any letters‑only (spacing/punct) equality surfaces as **Top1 with score 1.0**.

### Decisions Merge (Pick/Chosen_Email)
- The app reads `proposed_matches.xlsx` and imports **Top1/2/3** columns, **Decision**, **Pick**, **Chosen_Email**.
- **Resolution order** when accepting:
  1) `Pick`: `1`→Top1, `2`→Top2, `3`→Top3, or **email**
  2) `Chosen_Email`
  3) `Suggested_Email` (Top1)
- If you typed an **email** in `Pick` or `Chosen_Email`, the app looks it up in File B and fills **MatchedName_B** + roster fields. If not found, it still uses the email and flags **`EMAIL_NOT_IN_ROSTER`**.

### Overrides
- `Override_Email`: sets Email directly and bypasses name matching.
- `Override_FullName_B`: uses an exact File B name to fetch the Email + roster attributes.

### Dedupe Logic
- Removes exact duplicate rows in A by **(FullName_A, CreditHours, EventName)**.
- After decisions: sorts duplicates of **(Email, EventName)** by highest confidence and keeps only the best one.

### GUI Architecture
- **Tkinter** app (no extra install on macOS). Buttons run work on a background **thread** so the window stays responsive.
- **File pickers** (open/save dialogs) and an **output opener** that reveals your results.
- Same codepath as CLI: the GUI just calls `run_pipeline(...)` with the fields you supplied.

---

## Extending & Configuration
- **Nickname map**: `_NICK_MAP` in the script; add entries as needed.
- **Min match**: `--min_match` CLI flag or GUI field (default `0.85`).
- **Diacritics**: we can add normalization (e.g., `José`→`Jose`) if your datasets need it.
- **Hide greens**: we can add a toggle to hide `Certain=TRUE` rows in proposals.
- **Certificate export**: ready to wire into `master_list.xlsx` when you’re satisfied with the workflow.

---

## Troubleshooting
- **App closes on open**: rebuild with the current script; double‑click should open the GUI automatically.
- **Unidentified developer**: right‑click the app → **Open** once.
- **Nothing in Top 3 but you know the match**: type the email in **Pick**; the app will pull roster details if the email exists in File B.
- **Too many false positives**: increase **Min Match** (e.g., `0.88`).
- **Too few candidates**: decrease **Min Match** (e.g., `0.80`) and re‑generate proposals.
- **CSV/Excel errors**: ensure File A has at least 3 columns, File B has at least 7 columns in the specified order.

---

## Release Checklist
1. Build app and test GUI with sample files.
2. Generate proposals; verify greens and a few manual picks.
3. Apply decisions; inspect `joined_events.xlsx` and `master_list.xlsx`.
4. (Optional) Filter by Category.
5. Hand off `master_list.xlsx` to the certificate step.

---

**Workflow Recap:** **Generate Proposals** → edit **`proposed_matches.xlsx`** → **Apply Decisions** → use **`master_list.xlsx`**.