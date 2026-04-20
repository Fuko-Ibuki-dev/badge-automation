// =============================================================================
// SCRIPT NAME : AppendToStaging
// PURPOSE     : Reads the "Transformed" sheet from the uploaded badge file
//               and writes the data into Test_Masterfile.xlsx in two places:
//
//               1. Sheet 1 ("All Records") — every row ever processed is
//                  appended here cumulatively. This is what the team reviews
//                  weekly before copy-pasting into the real master tracker.
//
//               2. A new sheet per run — named "Attachment_YYYY-MM-DD_HH-MM"
//                  so you can trace exactly what came in from each email.
//
// HOW IT'S CALLED : Power Automate runs this script on Test_Masterfile.xlsx
//                   after TransformBadgeData has already run on the uploaded file.
//
// PARAMETERS  : uploadedFileId — the OneDrive file ID of the uploaded attachment,
//               passed in from Power Automate so this script can read its
//               Transformed sheet directly.
//
// NOTE        : This script is run ON Test_Masterfile.xlsx, but it reads data
//               FROM the uploaded file via the passed workbook reference.
// =============================================================================

function main(workbook: ExcelScript.Workbook) {

  // ---------------------------------------------------------------------------
  // STEP 1: Get the "Transformed" sheet from THIS workbook
  //         (Power Automate runs TransformBadgeData first, then copies the
  //          result here — so we read the Transformed sheet from the same file)
  // ---------------------------------------------------------------------------
  const transformedSheet = workbook.getWorksheet("Transformed");
  if (!transformedSheet) {
    console.log("ERROR: No 'Transformed' sheet found. Make sure TransformBadgeData ran first.");
    return;
  }

  // Read all data from the Transformed sheet (including header row)
  const transformedRange = transformedSheet.getUsedRange();
  if (!transformedRange) {
    console.log("ERROR: The Transformed sheet is empty. Nothing to append.");
    return;
  }

  const allValues = transformedRange.getValues();

  if (allValues.length < 2) {
    console.log("No data rows found in Transformed sheet (only header or empty).");
    return;
  }

  // Separate header row from data rows
  const headerRow = allValues[0];   // Row 1: column headers
  const dataRows  = allValues.slice(1); // Row 2 onwards: actual data

  console.log(`Found ${dataRows.length} data row(s) to append.`);

  // ---------------------------------------------------------------------------
  // STEP 2: Build a timestamp string for naming the new sheet
  //         Format: Attachment_YYYY-MM-DD_HH-MM
  //         e.g.   Attachment_2026-04-08_14-30
  // ---------------------------------------------------------------------------
  const now = new Date();

  // Pad single-digit numbers with a leading zero (e.g. 9 → "09")
  const pad = (n: number) => String(n).padStart(2, "0");

  const yyyy = now.getFullYear();
  const mm   = pad(now.getMonth() + 1); // getMonth() is 0-based
  const dd   = pad(now.getDate());
  const hh   = pad(now.getHours());
  const min  = pad(now.getMinutes());

  const timestamp  = `${yyyy}-${mm}-${dd}_${hh}-${min}`;
  const sheetName  = `Attachment_${timestamp}`; // e.g. "Attachment_2026-04-08_14-30"

  // ---------------------------------------------------------------------------
  // STEP 3: Set up Sheet 1 ("All Records") — create it if it doesn't exist yet
  //         If it's brand new, write the header row first.
  //         If it already exists, just find the next empty row and append.
  // ---------------------------------------------------------------------------
  const ALL_RECORDS_SHEET = "All Records";
  let allRecordsSheet = workbook.getWorksheet(ALL_RECORDS_SHEET);

  if (!allRecordsSheet) {
    // First time this script runs — create the sheet and write the header
    allRecordsSheet = workbook.addWorksheet(ALL_RECORDS_SHEET);

    // Move it to be the first sheet (position 0)
    allRecordsSheet.setPosition(0);

    // Write the header row
    for (let c = 0; c < headerRow.length; c++) {
      allRecordsSheet.getCell(0, c).setValue(headerRow[c]);
      // Bold the header
      allRecordsSheet.getCell(0, c).getFormat().getFont().setBold(true);
    }

    console.log(`Created new sheet: "${ALL_RECORDS_SHEET}" with headers.`);
  }

  // Find the next empty row in "All Records"
  // getUsedRange() returns null if truly empty (just created), so handle both
  const existingRange = allRecordsSheet.getUsedRange();
  const nextRow = existingRange ? existingRange.getRowCount() : 1;
  // nextRow is now the index of the first empty row (0-based), below all existing content

  // Write each data row into "All Records"
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    for (let c = 0; c < row.length; c++) {
      allRecordsSheet.getCell(nextRow + i, c).setValue(row[c]);
    }

    // Centre-align Date of Award (col G=6), Month (col H=7), Year (col I=8)
    // to match the formatting from TransformBadgeData
    allRecordsSheet.getCell(nextRow + i, 6).getFormat().setHorizontalAlignment(
      ExcelScript.HorizontalAlignment.center
    );
    allRecordsSheet.getCell(nextRow + i, 7).getFormat().setHorizontalAlignment(
      ExcelScript.HorizontalAlignment.center
    );
    allRecordsSheet.getCell(nextRow + i, 8).getFormat().setHorizontalAlignment(
      ExcelScript.HorizontalAlignment.center
    );
  }

  // Auto-fit columns so nothing is cut off
  allRecordsSheet.getUsedRange()?.getFormat().autofitColumns();
  console.log(`Appended ${dataRows.length} row(s) to "${ALL_RECORDS_SHEET}" starting at row ${nextRow + 1}.`);

  // ---------------------------------------------------------------------------
  // STEP 4: Create a new sheet named "Attachment_YYYY-MM-DD_HH-MM"
  //         This is a snapshot of just this batch — useful for tracing which
  //         email each row came from.
  //
  //         If by rare chance a sheet with this exact name already exists
  //         (two emails processed in the same minute), add a counter suffix.
  // ---------------------------------------------------------------------------
  let finalSheetName = sheetName;
  let counter = 1;
  while (workbook.getWorksheet(finalSheetName)) {
    // Sheet name already taken — append _2, _3, etc.
    finalSheetName = `${sheetName}_${counter}`;
    counter++;
  }

  const batchSheet = workbook.addWorksheet(finalSheetName);

  // Write header row with bold formatting
  for (let c = 0; c < headerRow.length; c++) {
    batchSheet.getCell(0, c).setValue(headerRow[c]);
    batchSheet.getCell(0, c).getFormat().getFont().setBold(true);
  }

  // Write data rows
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    for (let c = 0; c < row.length; c++) {
      batchSheet.getCell(i + 1, c).setValue(row[c]);
    }

    // Centre-align Date of Award, Month, Year columns
    batchSheet.getCell(i + 1, 6).getFormat().setHorizontalAlignment(
      ExcelScript.HorizontalAlignment.center
    );
    batchSheet.getCell(i + 1, 7).getFormat().setHorizontalAlignment(
      ExcelScript.HorizontalAlignment.center
    );
    batchSheet.getCell(i + 1, 8).getFormat().setHorizontalAlignment(
      ExcelScript.HorizontalAlignment.center
    );
  }

  // Auto-fit columns
  batchSheet.getUsedRange()?.getFormat().autofitColumns();
  console.log(`Created batch sheet: "${finalSheetName}" with ${dataRows.length} row(s).`);

  // ---------------------------------------------------------------------------
  // STEP 5: Activate "All Records" so the user lands on the summary sheet
  // ---------------------------------------------------------------------------
  allRecordsSheet.activate();

  console.log(`Done! Test_Masterfile.xlsx updated successfully at ${timestamp}.`);
}