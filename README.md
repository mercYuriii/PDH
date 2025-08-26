# PDH
PDH Cert Generator

# Credit Hours Prep (No Certificates Yet)

This program cleans up event attendance records and conference rosters, matches participants by name/email, and produces a **master list of credit hours per participant**. It is designed to be simple enough for non-technical users: you can run it as a Mac app with a graphical interface.

---

## Features
- Deduplicates duplicate rows in **File A** (event hours).
- Collapses **File B** (roster) so each person/email is unique.
- Matches File A names to File B names (fuzzy matching, nicknames, spacing, punctuation, concatenated names).
- Flags uncertain matches for manual review in Excel.
- Applies your review decisions and optional manual overrides.
- Prevents double-counting (same person, same event).
- Produces:
  - `master_list.xlsx` / `master_list.csv` — final credit totals
  - `proposed_matches.xlsx` — names to review
  - diagnostic files (`joined_events.xlsx`, `unmatched_needs_email.xlsx`, etc.)

---

## File Requirements

### File A (PDH Data)
- **Columns (first 3, in order):**
  - Full Name  
  - Credit Hours  
  - Event Name  

### File B (Registration Data)
- **Columns (first 7, in order):**
  - Category  
  - Subcategory  
  - Full Name  
  - Country  
  - Email  
  - CC Email  
  - First Conference  

### Manual Overrides (Optional, CSV)
Headers (case-insensitive):
- `FullName_A`  
- `Override_FullName_B` *(optional, must match File B “Full Name” exactly)*  
- `Override_Email` *(optional, direct email — takes precedence)*  

---

## How to Run (GUI App)

1. **Build the app once:**
   ```bash
   pip install pyinstaller openpyxl pandas
   cd ~/Desktop/PDH
   pyinstaller --windowed --onefile --name "Credit Hours Prep" run_me_nocerts.py