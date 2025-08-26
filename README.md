2. **Open the app:**
   - On Mac, find the app named **Credit Hours Prep** in the `dist` folder inside your PDH directory.
   - Double-click to launch the graphical interface.

3. **How to use the app:**
   - Select your **File A** (PDH Data) and **File B** (Registration Data) Excel or CSV files.
   - (Optional) Load a Manual Overrides CSV file if you have one.
   - Click **Run** to process the data.
   - Review the proposed matches in the generated Excel file by following these steps:
     - Open the `proposed_matches.xlsx` file in Excel.
     - Check the columns `FullName_A`, `Top1_Name_B`, `Top1_Email`, and the confidence score.
     - For green rows (where the `Certain` column is TRUE), no action is needed as these matches are auto-accepted.
     - For other rows, review carefully and set the `Decision` column to either ACCEPT or REJECT.
     - If you find that `Top2` or `Top3` is a better match, copy the corresponding email into the `Chosen_Email` column and mark the `Decision` as ACCEPT.
     - Save the file after completing all your decisions.
   - Apply your review decisions and generate the final master list.

---

## Command Line Usage

If you prefer to run the program via command line without the GUI, use:

```bash
python run_me_nocerts.py --fileA path/to/fileA.xlsx --fileB path/to/fileB.xlsx [--overrides path/to/overrides.csv] [--output path/to/output_folder]
```

- Replace paths with your actual file locations.
- The overrides and output folder arguments are optional.

---

## Output Files

After running the program, you will find the following files in the specified output directory (default is the current folder):

- `master_list.xlsx` / `master_list.csv` — Final credit hours per participant.
- `proposed_matches.xlsx` — List of names flagged for manual review.
- `joined_events.xlsx` — Combined event and roster data for diagnostics.
- `unmatched_needs_email.xlsx` — Participants missing email addresses.
- Additional diagnostic files to assist with troubleshooting.

---

## Troubleshooting

- **No output files generated:**  
  Ensure your input files meet the column requirements and are properly formatted.

- **Matches not found or incorrect:**  
  Review the `proposed_matches.xlsx` file and manually correct any mismatches.

- **App won’t open or crashes:**  
  Confirm you have installed all dependencies (`pyinstaller`, `openpyxl`, `pandas`) and that you are running the correct Python version (3.6+ recommended).

- **Manual overrides not applied:**  
  Verify your overrides CSV headers and values match exactly as specified.

For further help, please contact the project maintainer or check the GitHub repository for updates and issues.