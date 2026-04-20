// =============================================================================
// SCRIPT NAME : TransformBadgeData
// PURPOSE     : Reads badge/training data from the current sheet and writes it
//               into a new "Transformed" sheet formatted to match the master
//               Skills Badge tracker file — ready to copy-paste in.
//
// HOW TO RUN  : 1. Open the source Excel file (the one with the raw data).
//               2. Make sure you are on the sheet that contains the data
//                  (it will be highlighted/active in the sheet tabs at the bottom).
//               3. Go to Automate tab → click "TransformBadgeData" → Run.
//               4. A sheet called "Transformed" will appear. Review it, then
//                  copy the rows you need into the master tracker file.
//
// SOURCE COLUMNS (what this script reads):
//   A  – Email Address
//   D  – Preferred Name (name on badge)
//   E  – Name of Skills Badge  e.g. "HR Consulting (Basic)"
//   G  – Course / Programme Title
//   H  – Date of Course Completion
//   I  – Training Provider
//
// OUTPUT COLUMNS (what the Transformed sheet contains):
//   A  – Name
//   B  – Email
//   C  – Training Provider
//   D  – Skills Area          (text before the bracket, e.g. "HR Consulting")
//   E  – Skills Area and Level (full badge name, e.g. "HR Consulting (Basic)")
//   F  – Badge Level           (text inside the bracket, e.g. "Basic")
//   G  – Date of Award         (formatted as DD-MMM-YY, e.g. "12-Sep-25")
//   H  – Month                 (number only, e.g. 9)
//   I  – Year                  (4-digit, e.g. 2025)
//   J  – Programme
// =============================================================================

function main(workbook: ExcelScript.Workbook) {

  // ---------------------------------------------------------------------------
  // STEP 1: Get the sheet the user is currently viewing (the source data sheet)
  // ---------------------------------------------------------------------------
  const srcSheet = workbook.getActiveWorksheet();

  // Find out how many rows have data (including the header row)
  const usedRange = srcSheet.getUsedRange();
  if (!usedRange) {
    console.log("The active sheet is empty. Please switch to the sheet with badge data and run again.");
    return;
  }
  const lastRow = usedRange.getRowCount(); // total rows including header

  if (lastRow < 2) {
    console.log("No data rows found — the sheet only has a header or is empty.");
    return;
  }

  // ---------------------------------------------------------------------------
  // STEP 2: Set up the output sheet called "Transformed"
  //         If it already exists, clear it so we start fresh.
  //         If not, create it.
  // ---------------------------------------------------------------------------
  const OUTPUT_SHEET_NAME = "Transformed";
  let outSheet = workbook.getWorksheet(OUTPUT_SHEET_NAME);
  if (outSheet) {
    // Sheet exists from a previous run — wipe it clean before writing new data
    outSheet.getUsedRange()?.clear();
  } else {
    outSheet = workbook.addWorksheet(OUTPUT_SHEET_NAME);
  }

  // ---------------------------------------------------------------------------
  // STEP 3: Write the header row into the Transformed sheet (starting at A1)
  // ---------------------------------------------------------------------------
  const headers = [
    "Name",                  // A
    "Email",                 // B
    "Training Provider",     // C
    "Skills Area",           // D
    "Skills Area and Level", // E
    "Badge Level",           // F
    "Date of Award",         // G
    "Month",                 // H
    "Year",                  // I
    "Programme",             // J
  ];

  for (let c = 0; c < headers.length; c++) {
    outSheet.getCell(0, c).setValue(headers[c]);
  }

  // ---------------------------------------------------------------------------
  // STEP 4: Define a helper function to convert a date value into:
  //           • a readable string  → "DD-MMM-YY"  (e.g. "12-Sep-25")
  //           • the month number   → 9
  //           • the 4-digit year   → 2025
  //
  //         Why is this needed?
  //         Excel stores dates internally as a plain number (called a "serial
  //         number") — e.g. 45912 means 12 Sep 2025. If we just copy that
  //         number across, it shows as 45912 instead of a real date.
  //         This function converts it to a human-readable text string so the
  //         value always looks correct regardless of cell formatting.
  // ---------------------------------------------------------------------------
  const MONTH_ABBR = ["Jan","Feb","Mar","Apr","May","Jun",
                      "Jul","Aug","Sep","Oct","Nov","Dec"];

  function formatDate(raw: ExcelScript.RangeValueType): {
    display: string;
    month: string;
    year: string;
  } {
    let dateObj: Date | null = null;

    if (typeof raw === "number" && raw > 0) {
      // Convert Excel serial number to a JavaScript Date.
      // Excel counts days from 30 Dec 1899 (epoch), so we add the serial
      // number of days (in milliseconds) to that starting point.
      const excelEpoch = new Date(1899, 11, 30); // 30 Dec 1899
      dateObj = new Date(excelEpoch.getTime() + raw * 86400000);

    } else if (typeof raw === "string" && raw.trim() !== "") {
      // If the source cell already has the date stored as text, try parsing it
      const parsed = new Date(raw);
      if (!isNaN(parsed.getTime())) dateObj = parsed;
    }

    // If we couldn't parse a valid date, return empty strings (nothing to show)
    if (!dateObj) return { display: "", month: "", year: "" };

    // Build the "DD-MMM-YY" string
    const dd  = String(dateObj.getDate()).padStart(2, "0"); // e.g. "07" or "12"
    const mmm = MONTH_ABBR[dateObj.getMonth()];             // e.g. "Sep"
    const yy  = String(dateObj.getFullYear()).slice(-2);    // e.g. "25"

    return {
      display: `${dd}-${mmm}-${yy}`,                    // "12-Sep-25"
      month:   String(dateObj.getMonth() + 1),           // "9"
      year:    String(dateObj.getFullYear()),             // "2025"
    };
  }

  // ---------------------------------------------------------------------------
  // STEP 5: Read ALL data rows from the source sheet at once (faster than
  //         reading cell-by-cell) then loop through and write to output sheet.
  // ---------------------------------------------------------------------------

  // Read rows 2 to end (skip header at row 1), columns A to I (9 columns)
  const dataRange = srcSheet.getRangeByIndexes(1, 0, lastRow - 1, 9);
  const values = dataRange.getValues();

  let outRow = 1; // Start writing data at row 2 of the output sheet (index 1)

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    // Pull out each field from the correct source column
    const email            = String(row[0] ?? "").trim(); // Column A
    const preferredName    = String(row[3] ?? "").trim(); // Column D
    const skillsBadge      = String(row[4] ?? "").trim(); // Column E  e.g. "HR Consulting (Basic)"
    const programme        = String(row[6] ?? "").trim(); // Column G
    const dateRaw          = row[7];                      // Column H  (may be a serial number)
    const trainingProvider = String(row[8] ?? "").trim(); // Column I

    // Skip rows that have no meaningful content (e.g. blank rows in the data)
    if (!email && !preferredName && !skillsBadge) continue;

    // ── Parse the Skills Badge name ──────────────────────────────────────────
    // The badge name follows the pattern "Area Name (Level)"
    // We split it into three parts:
    //   skillsArea         → "HR Consulting"          (text before the bracket)
    //   skillsAreaAndLevel → "HR Consulting (Basic)"  (the full original text)
    //   badgeLevel         → "Basic"                  (text inside the bracket)

    let skillsArea         = skillsBadge; // default: use full text if no bracket found
    let skillsAreaAndLevel = skillsBadge; // always the full original text
    let badgeLevel         = "";

    const parenMatch = skillsBadge.match(/^(.*?)\s*\(([^)]+)\)\s*$/);
    if (parenMatch) {
      skillsArea = parenMatch[1].trim(); // everything before " ("
      badgeLevel = parenMatch[2].trim(); // everything between "(" and ")"
    }

    // ── Convert the date ─────────────────────────────────────────────────────
    const { display: dateDisplay, month, year } = formatDate(dateRaw);

    // ── Write this row to the Transformed sheet ───────────────────────────────
    outSheet.getCell(outRow, 0).setValue(preferredName);      // A – Name
    outSheet.getCell(outRow, 1).setValue(email);              // B – Email
    outSheet.getCell(outRow, 2).setValue(trainingProvider);   // C – Training Provider
    outSheet.getCell(outRow, 3).setValue(skillsArea);         // D – Skills Area
    outSheet.getCell(outRow, 4).setValue(skillsAreaAndLevel); // E – Skills Area and Level
    outSheet.getCell(outRow, 5).setValue(badgeLevel);         // F – Badge Level
    outSheet.getCell(outRow, 6).setValue(dateDisplay);        // G – Date of Award  e.g. "12-Sep-25"
    outSheet.getCell(outRow, 7).setValue(month);              // H – Month
    outSheet.getCell(outRow, 8).setValue(year);               // I – Year
    outSheet.getCell(outRow, 9).setValue(programme);          // J – Programme

    // ── Centre-align Date of Award (G), Month (H), and Year (I) ─────────────
    // These are short values that look neater centred in the cell.
    outSheet.getCell(outRow, 6).getFormat().setHorizontalAlignment(
      ExcelScript.HorizontalAlignment.center
    );
    outSheet.getCell(outRow, 7).getFormat().setHorizontalAlignment(
      ExcelScript.HorizontalAlignment.center
    );
    outSheet.getCell(outRow, 8).getFormat().setHorizontalAlignment(
      ExcelScript.HorizontalAlignment.center
    );

    outRow++; // Move to the next row for the next record
  }

  // ---------------------------------------------------------------------------
  // STEP 6: Auto-fit all column widths so text isn't cut off
  // ---------------------------------------------------------------------------
  outSheet.getUsedRange()?.getFormat().autofitColumns();

  // ---------------------------------------------------------------------------
  // STEP 7: Switch the view to the Transformed sheet so the user sees results
  // ---------------------------------------------------------------------------
  outSheet.activate();

  console.log(`Done! ${outRow - 1} row(s) written to the "${OUTPUT_SHEET_NAME}" sheet.`);
}