function moveFilesToFolder() {
  // ID folder tujuan
  const folderId = "your folder id";
  // Nama sheet
  const sheetName = "Link";
  // Kolom nama file
  const filenameCol = 1; // Column A
  // Kolom nama folder tujuan
  const foldernameCol = 8; // Column H

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  let filesMoved = 0; // Track the number of files moved

  // Loop through each row in the sheet starting from row 2
  for (let i = 2; i <= lastRow; i++) {
    const filename = sheet.getRange(i, filenameCol).getValue();
    const foldername = sheet.getRange(i, foldernameCol).getValue();

    // Check if both filename and foldername are not empty
    if (filename && foldername && foldername !== "Invoice") {
      const files = DriveApp.getFilesByName(filename);
      const folders = DriveApp.getFoldersByName(foldername);

      if (files.hasNext() && folders.hasNext()) {
        const file = files.next();
        const folder = folders.next();

        // Move the file to the folder
        folder.createFile(file.getBlob());
        file.setTrashed(true);
        Logger.log(`File "${filename}" moved to folder "${foldername}"`);
        filesMoved++;
      } else {
        Logger.log(`File "${filename}" or folder "${foldername}" not found.`);
      }
    }

    // Check if all files have been moved
    if (filesMoved === lastRow - 1) {
      Logger.log("All files have been moved.");
      return; // End the function execution
    }
  }
}
