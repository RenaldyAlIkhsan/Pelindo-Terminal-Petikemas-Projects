function removeDuplicateRows(data) {
  const seen = new Set();
  const uniqueData = [];

  for (let i = 0; i < data.length; i++) {
    const row = JSON.stringify(data[i]);
    if (!seen.has(row)) {
      seen.add(row);
      uniqueData.push(data[i]);
    }
  }

  return uniqueData;
}

function importExcelAndCopySheet(folderId) {
  var errorCount = 0;
  const sheetNameToGet = 'MONITORING KAPASITAS';
  const destinationSheetName = 'MONITORING KAPASITAS Raw';

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);

  let excelFiles = [];
  while (files.hasNext()) {
    excelFiles.push(files.next());
  }

  if (excelFiles.length != 1) {
    Logger.log('File tidak ditemukan, pastikan hanya terdapat 1 file excel saja');
    SpreadsheetApp.getActiveSpreadsheet().toast('File tidak ditemukan, pastikan hanya terdapat 1 file excel saja', 'Warning', 10);
    errorCount++;
    Utilities.sleep(5000);
    throw new Error('File tidak ditemukan, pastikan hanya terdapat 1 file excel saja');
  }

  const file = excelFiles[0];
  Logger.log('File ditemukan ' + file.getName());

  const resource = {
    title: file.getName(),
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: folderId }]
  };

  const converted = Drive.Files.insert(resource, file.getBlob());
  Logger.log('File telah diconvert menjadi google sheet: ' + converted.id);

  try {
    const tempSpreadsheet = SpreadsheetApp.openById(converted.id);
    const sourceSheet = tempSpreadsheet.getSheetByName(sheetNameToGet);

    if (!sourceSheet) {
      Logger.log('Sheet "' + sheetNameToGet + '" tidak ditemukan, pastikan nama sheet adalah MONITORING KAPASITAS!');
      SpreadsheetApp.getActiveSpreadsheet().toast('Sheet "' + sheetNameToGet + '" tidak ditemukan, pastikan nama sheet adalah MONITORING KAPASITAS!', 'Warning', 10);
      errorCount++;
      Utilities.sleep(5000);
      throw new Error('Sheet "' + sheetNameToGet + '" tidak ditemukan, pastikan nama sheet adalah MONITORING KAPASITAS!');
    }

    let data = sourceSheet.getDataRange().getValues();
    data = removeDuplicateRows(data);

    const destSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let destSheet = destSpreadsheet.getSheetByName(destinationSheetName);

    if (!destSheet) {
      destSheet = destSpreadsheet.insertSheet(destinationSheetName);
    } else {
      destSheet.clearContents();
    }

    destSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    Logger.log('Data telah termuat tanpa duplikat');

  } catch (err) {
    Logger.log('Error: ' + err.message);
  } finally {
    DriveApp.getFileById(converted.id).setTrashed(true);
    Logger.log('Membersihkan tembolok (cache)');
  }

  return errorCount;
}

function extractSpecificColumns() {
  var errorCount = 0;
  const sourceSheetName = 'MONITORING KAPASITAS Raw';
  const targetSheetName = 'MONITORING KAPASITAS';
  const desiredHeaders = [
    'CARGO', 'IN TERMINAL', 'IN CARRIER NAME', 'OUT CARRIER NAME',
    'GW (Kg)', 'CONTAINER SIZE', 'CONTAINER TYPE', 'CONTAINER STATUS',
    'BLOK', 'SLOT', 'ROW', 'TIER', 'CLASS', 'POD','OPR','F/M','Comm.'
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  const data = sourceSheet.getDataRange().getValues();

  const headers = data[0];

  const columnIndices = [];
  const missingHeaders = [];

  desiredHeaders.forEach(header => {
    const index = headers.indexOf(header);
    if (index === -1) {
      missingHeaders.push(header);
    } else {
      columnIndices.push(index);
    }
  });

  if (missingHeaders.length > 0) {
    ss.toast("Terdapat kolom yang tidak lengkap, mohon dilengkapi sesuai: " + JSON.stringify(desiredHeaders), "Error", 10);
    errorCount++;
    throw new Error("Kolom tidak ditemukan: " + missingHeaders.join(", "));
  }

  const filteredData = data.map(row => columnIndices.map(i => row[i]));

  const targetSheet = ss.getSheetByName(targetSheetName) || ss.insertSheet(targetSheetName);
  targetSheet.clearContents();
  targetSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);

  // ====== PEWARNAAN ======
  const classIndex = desiredHeaders.indexOf("CLASS");
  const podIndex = desiredHeaders.indexOf("POD");
  const oprIndex = desiredHeaders.indexOf("OPR");
  const lastRow = filteredData.length;

  for (let i = 1; i < lastRow; i++) {
    const classValue = filteredData[i][classIndex];
    const podValue = filteredData[i][podIndex];
    const oprValue = filteredData[i][oprIndex];

    const classCell = targetSheet.getRange(i + 1, classIndex + 1);
    const podCell = targetSheet.getRange(i + 1, podIndex + 1);
    const oprCell = targetSheet.getRange(i + 1, oprIndex + 1);

    // CLASS = IMPORT
    if (String(classValue).toUpperCase() === "IMPORT") {
      classCell.setBackground("#00FFFF");
    }

    // CLASS = EXPORT
    else if (String(classValue).toUpperCase() === "EXPORT") {
      switch (String(podValue).toUpperCase()) {
        case "IDJKT":
          podCell.setFontColor("#0066FF");
          break;
        case "IDSUB":
          podCell.setFontColor("#FF0000");
          break;
        case "IDMAK":
          podCell.setFontColor("#FF00FF");
          break;
        case "IDBIK":
          podCell.setBackground("#C00000").setFontColor("#FFFFFF");
          break;
        case "IDMKW":
          podCell.setFontColor("#000000");
          break;
        case "IDSOQ":
          podCell.setFontColor("#000000");
          break;
        case "IDAMQ":
          podCell.setFontColor("#000000");
          break;
      }
    }

    // OPR COLORING
    switch (String(oprValue).toUpperCase()) {
      case "SPL":
        oprCell.setBackground("#92D050");
        break;
      case "TNT":
        oprCell.setBackground("#0099FF");
        break;
      case "TMS":
        oprCell.setBackground("#F4B084");
        break;
      default:
        oprCell.setBackground("#FFFFFF");
        break;
    }
  }

  Logger.log("Pewarnaan berhasil diterapkan.");
  return errorCount;
}
