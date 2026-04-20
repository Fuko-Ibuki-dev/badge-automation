================================================================================
  TransformBadgeData — Office Script README
================================================================================

WHAT THIS SCRIPT DOES
----------------------
This script takes the raw badge/training data from your source Excel file and
reformats it into a new sheet called "Transformed". The Transformed sheet uses
the exact same column layout as the master Skills Badge tracker file, so you
can simply copy and paste the rows across without any manual rearranging.


--------------------------------------------------------------------------------
BEFORE YOU START — check your source file has these columns:
--------------------------------------------------------------------------------

  Column A  →  Email Address
  Column D  →  Preferred Name (name to appear on badge)
  Column E  →  Name of Skills Badge  (format: "Area Name (Level)")
                 e.g. "HR Consulting (Basic)"
  Column G  →  Course / Programme Title
  Column H  →  Date of Course Completion
  Column I  →  Training Provider

  Row 1 should be the header row. Data starts from Row 2.


--------------------------------------------------------------------------------
HOW TO RUN THE SCRIPT (step by step)
--------------------------------------------------------------------------------

  1. Open the source Excel file in Excel for the Web (office.com).

  2. Click on the sheet tab at the bottom that contains your badge data,
     so that sheet is active (highlighted).

  3. Click the "Automate" tab in the top ribbon.

  4. Click "TransformBadgeData" in the script list, then click "Run".
     → The script will run and a new sheet tab called "Transformed" will appear.

  5. Click on the "Transformed" sheet tab and check the data looks correct.

  6. Select the data rows you want (not the header — the master file already
     has its own header), copy them, then paste into the master tracker file.


--------------------------------------------------------------------------------
WHAT THE TRANSFORMED SHEET CONTAINS
--------------------------------------------------------------------------------

  Column A  →  Name
  Column B  →  Email
  Column C  →  Training Provider
  Column D  →  Skills Area           (e.g. "HR Consulting")
  Column E  →  Skills Area and Level (e.g. "HR Consulting (Basic)")
  Column F  →  Badge Level           (e.g. "Basic")
  Column G  →  Date of Award         (e.g. "12-Sep-25")  ← centred
  Column H  →  Month                 (e.g. 9)            ← centred
  Column I  →  Year                  (e.g. 2025)         ← centred
  Column J  →  Programme


--------------------------------------------------------------------------------
GOOD TO KNOW
--------------------------------------------------------------------------------

  • Re-running the script is safe — it clears and rewrites the "Transformed"
    sheet each time, so you won't end up with duplicate rows.

  • The Date of Award is saved as readable text (e.g. "12-Sep-25") instead of
    a number, so it will always display correctly regardless of cell formatting.

  • The script automatically skips any blank rows in the source data.

  • DO NOT paste into the master tracker file with formulas intact if the master
    uses validation or protected cells — use "Paste Special → Values Only".


--------------------------------------------------------------------------------
TROUBLESHOOTING
--------------------------------------------------------------------------------

  Problem : The "Transformed" sheet appears but has no data rows.
  Fix     : Make sure you are on the correct data sheet before running.
            The script reads whichever sheet is currently active.

  Problem : Dates still show as numbers (e.g. 45912).
  Fix     : The source column H may be formatted as General/Number instead of
            Date. Try formatting column H as "Date" in the source sheet first,
            then re-run the script.

  Problem : Skills Area / Badge Level columns are empty.
  Fix     : Check that column E in the source follows the exact format:
            "Area Name (Level)" — the level must be inside round brackets ( ).

  Problem : Script not visible in the Automate tab.
  Fix     : Make sure the script was saved in the same workbook. Go to
            Automate → All Scripts to find it.

================================================================================