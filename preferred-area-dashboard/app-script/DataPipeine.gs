function DataPipeline() {
  ////// TEMPAT UPDATE /////////////////////////////////////////////////////////////////////////////////////////

  const folderId               = "1FPBrmI_lCBNdHPTGtaydDFLNPaqLNHUy";
  const cells                  = ['I4', 'BL4', 'CB4', 'DF4', 'FA4', 'GW4'];
  const preferredAreaSheetName = "PREFERRED AREA";
  const maintenanceSheetName   = "MAINTENANCE";

  //////////////////////////////////////////////////////////////////////////////////////////////////////////////
  var errorCount = 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  errorCount = importExcelAndCopySheet(folderId);
  Logger.log(errorCount);
  errorCount = extractSpecificColumns() + errorCount;
  Logger.log(errorCount);

  // Logging
  if (errorCount === 0) {
    ss.toast("Data Berhasil/Sudah Dimuat", "Sukses", 10);
    const preferredAreaSheet = ss.getSheetByName(preferredAreaSheetName);
    const maintenanceSheet = ss.getSheetByName(maintenanceSheetName);

    // Last Update
    const timestamp = new Date();
    const formatted = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    const message = "Last Updated: " + formatted;

    // Iterate through dashboard headers
    
    cells.forEach(cell => {
      preferredAreaSheet.getRange(cell).setValue(message);
      maintenanceSheet.getRange(cell).setValue(message);
    });
  } else {
    ss.toast("Data Gagal Dimuat", "Error", 10);
    throw new Error("Data Gagal Dimuat");
  }
}
