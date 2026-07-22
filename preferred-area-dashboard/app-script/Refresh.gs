// =====================================================================================
//  GLOBAL CONFIG
// =====================================================================================

const SHEET_PREFERRED   = "PREFERRED AREA";
const SHEET_MAINTENANCE = "MAINTENANCE";

/**
 * Daftar blok yang mau digenerate ulang saat Refresh.
 * 
 * IMPROVISASI:
 * - Tambah property "reverseRow" untuk blok yang dekat truck lane (1E, 1D, 1B)
 * - reverseRow: true → row numbering terbalik (descending)
 * - reverseRow: false/undefined → row numbering normal (ascending)
 * 
 * Angka slot/row/tier diambil dari sheet MAINTENANCE.
 */
const BLOCK_CONFIGS = [
  {
    orientation: "H",
    blockName:   "1E",
    cellStart:   "S6",
    reverseRow:  true    // ← TERBALIK (dekat truck lane)
  },
  {
    orientation: "V",
    blockName:   "1F",
    cellStart:   "AT6",
    reverseRow:  false
  },
  {
    orientation: "V",
    blockName:   "1G",
    cellStart:   "AT25",
    reverseRow:  false
  },
  {
    orientation: "H",
    blockName:   "1D",
    cellStart:   "S43",
    reverseRow:  true    // ← TERBALIK (dekat truck lane)
  },
  {
    orientation: "H",
    blockName:   "1C",
    cellStart:   "S65",
    reverseRow:  false
  },
  {
    orientation: "H",
    blockName:   "1B",
    cellStart:   "S81",
    reverseRow:  true    // ← TERBALIK (dekat truck lane)
  },
  {
    orientation: "V",
    blockName:   "FLT",
    cellStart:   "B81",
    reverseRow:  false
  },
  {
    orientation: "H",
    blockName:   "1A",
    cellStart:   "S103",
    reverseRow:  false
  },
  {
    orientation: "H",
    blockName:   "1A1",
    cellStart:   "B103",
    reverseRow:  false
  },
  {
    orientation: "V",
    blockName:   "1A2",
    cellStart:   "AQ103",
    reverseRow:  false
  }
];


// =====================================================================================
//  REMOVE ALL CHARTS/DIAGRAMS FROM SHEET
// =====================================================================================

const GAUGE_PROTECTED_AREA = {
  startRow: 10,  // baris 10–22
  endRow:   22,
  startCol: 4,   // kolom D–J
  endCol:   10
};

function removeAllCharts(sheet) {
  if (!sheet) return;
  
  Logger.log(`Removing all charts from sheet: ${sheet.getName()}`);
  
  const charts = sheet.getCharts();
  
  if (charts && charts.length > 0) {
    Logger.log(`  Found ${charts.length} chart(s) to remove...`);
    
    charts.forEach((chart, index) => {
      const info = chart.getContainerInfo();
      const anchorRow = info.getAnchorRow();
      const anchorCol = info.getAnchorColumn();

      const r1 = anchorRow + 1;
      const c1 = anchorCol + 1;

      if (
        sheet.getName() === SHEET_PREFERRED &&
        r1 >= GAUGE_PROTECTED_AREA.startRow &&
        r1 <= GAUGE_PROTECTED_AREA.endRow &&
        c1 >= GAUGE_PROTECTED_AREA.startCol &&
        c1 <= GAUGE_PROTECTED_AREA.endCol
      ) {
        Logger.log(`  ⏩ Skipping protected gauge chart at row ${r1}, col ${c1}`);
        return;
      }

      sheet.removeChart(chart);
      Logger.log(`  ✓ Chart ${index + 1} removed`);
    });
    
    Logger.log(`  ✓ All non-protected chart(s) removed successfully`);
  } else {
    Logger.log(`  No charts found in ${sheet.getName()}`);
  }
}

// =====================================================================================
//  CLEAR DASHBOARD AREA + UNMERGE CELLS
// =====================================================================================

function clearDashboardArea(sheet) {
  if (!sheet) return;

  Logger.log(`Clearing dashboard area in sheet: ${sheet.getName()}`);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const HEADER_PROTECTED_ROWS = 5;

  if (lastRow <= HEADER_PROTECTED_ROWS || lastCol < 1) return;

  // Protected ranges: header, blok instruksi, dan TRUCK LANE labels
  const protectedRangesA1 = [
    "B24:P24",
    "S37:AX37",
    "S56:AX56",
    "S59:AX59",
    "B94:BQ94",
    "B96:BQ96",
    "AT20:BI20",
    "S78:AX78",
    "S100:AX100",
    "S5:AM5",
    "AT5:BE5",
    "AT24:BE24",
    "S42:AT42",
    "S64:AU64",
    "S80:AU80",
    "B80:B80",
    "B102:M102",
    "S102:AK102",
    "AQ102:BM102"
  ];

  const saved = protectedRangesA1.map(a1 => {
    const r = sheet.getRange(a1);
    return {
      a1: a1,
      range: r,
      values: r.getValues(),
      bg: r.getBackgrounds(),
      fonts: r.getFontColors(),
      merged: r.isPartOfMerge()
    };
  });

  const startRowProtected = HEADER_PROTECTED_ROWS + 1;
  const endRowProtected   = 64;
  const protectedLastRow  = Math.min(endRowProtected, lastRow);

  if (lastCol >= 16 && protectedLastRow >= startRowProtected) {
    const rows = protectedLastRow - startRowProtected + 1;
    const cols = lastCol - 16 + 1;
    const range1 = sheet.getRange(startRowProtected, 16, rows, cols);

    Logger.log(`  Unmerging area 1: rows ${startRowProtected}-${protectedLastRow}, cols 16-${lastCol}`);

    try {
      const mergedRanges = range1.getMergedRanges();
      if (mergedRanges && mergedRanges.length > 0) {
        Logger.log(`    Found ${mergedRanges.length} merged ranges`);
        mergedRanges.forEach(mr => mr.breakApart());
      }
    } catch (e) {
      Logger.log(`    Warning: ${e.message}`);
      range1.breakApart();
    }

    range1.clearContent();
    range1.setBackground("#ffffff");
    range1.setFontColor("black");
    Logger.log(`  ✓ Area 1 cleared`);
  }

  if (lastRow > endRowProtected) {
    const startRow2 = endRowProtected + 1;
    const rows2 = lastRow - endRowProtected;
    const range2 = sheet.getRange(startRow2, 1, rows2, lastCol);

    Logger.log(`  Unmerging area 2: rows ${startRow2}-${lastRow}, all columns`);

    try {
      const mergedRanges = range2.getMergedRanges();
      if (mergedRanges && mergedRanges.length > 0) {
        Logger.log(`    Found ${mergedRanges.length} merged ranges`);
        mergedRanges.forEach(mr => mr.breakApart());
      }
    } catch (e) {
      Logger.log(`    Warning: ${e.message}`);
      range2.breakApart();
    }

    range2.clearContent();
    range2.setBackground("#ffffff");
    range2.setFontColor("black");
    Logger.log(`  ✓ Area 2 cleared`);
  }

  saved.forEach(obj => {
    obj.range.setValues(obj.values);
    obj.range.setBackgrounds(obj.bg);
    obj.range.setFontColors(obj.fonts);
    if (obj.merged) {
      obj.range.merge();
    }
  });

  Logger.log(`✓ Dashboard area cleared and unmerged (rows 1–5 preserved)`);
}

// =====================================================================================
//  CLEAR MAINTENANCE AREA + PRESERVE BROKEN DATA
// =====================================================================================

function clearMaintenanceArea(sheet) {
  if (!sheet) return;

  Logger.log(`Clearing maintenance area in sheet: ${sheet.getName()}`);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const HEADER_PROTECTED_ROWS = 5;
  const protectedRowEnd = 20;
  const protectedColEnd = 11;

  if (lastRow <= protectedRowEnd && lastCol <= protectedColEnd) {
    Logger.log(`  Nothing to clear (all within protected area)`);
    return;
  }

  // ========================================
  // SAVE BROKEN DATA DARI SEMUA BLOK
  // ========================================
  const savedBrokenData = [];
  const blockParamsMap = getBlockParamsFromMaintenanceSheet();
  
  BLOCK_CONFIGS.forEach(cfg => {
    try {
      const blockRange = sheet.getRange(cfg.cellStart);
      const startRow = blockRange.getRow();
      const startCol = blockRange.getColumn() + 3;
      
      const params = blockParamsMap[cfg.blockName];
      if (!params) return;
      
      const { slotCount, rowCount } = params;
      
      if (cfg.orientation === "H") {
        // Horizontal: data di baris
        const dataRange = sheet.getRange(startRow + 2, startCol + 2, rowCount, slotCount);
        const values = dataRange.getValues();
        
        savedBrokenData.push({
          blockName: cfg.blockName,
          orientation: cfg.orientation,
          startRow: startRow + 2,
          startCol: startCol + 2,
          values: values
        });
        
        Logger.log(`  ✓ Saved broken data for ${cfg.blockName} (H): ${rowCount}x${slotCount}`);
        
      } else if (cfg.orientation === "V") {
        // Vertical: data di kolom
        const dataRange = sheet.getRange(startRow + 6, startCol + 1, slotCount, rowCount);
        const values = dataRange.getValues();
        
        savedBrokenData.push({
          blockName: cfg.blockName,
          orientation: cfg.orientation,
          startRow: startRow + 6,
          startCol: startCol + 1,
          values: values
        });
        
        Logger.log(`  ✓ Saved broken data for ${cfg.blockName} (V): ${slotCount}x${rowCount}`);
      }
    } catch (err) {
      Logger.log(`  ⚠ Could not save broken data for ${cfg.blockName}: ${err.message}`);
    }
  });

  // Simpan ke PropertiesService
  try {
    PropertiesService.getScriptProperties().setProperty(
      'SAVED_BROKEN_DATA',
      JSON.stringify(savedBrokenData)
    );
    Logger.log(`✓ Saved ${savedBrokenData.length} blocks' broken data to properties`);
  } catch (err) {
    Logger.log(`⚠ Could not save to properties: ${err.message}`);
  }

  // ========================================
  // CLEAR AREAS (EXISTING LOGIC)
  // ========================================
  
  if (lastCol > protectedColEnd && lastRow > HEADER_PROTECTED_ROWS) {
    const startRow1 = HEADER_PROTECTED_ROWS + 1;
    const clearRows = lastRow - HEADER_PROTECTED_ROWS;
    const clearCols = lastCol - protectedColEnd;
    const range1 = sheet.getRange(startRow1, protectedColEnd + 1, clearRows, clearCols);

    Logger.log(`  Unmerging area 1: rows ${startRow1}-${lastRow}, cols ${protectedColEnd + 1}-${lastCol}`);

    try {
      const mergedRanges = range1.getMergedRanges();
      if (mergedRanges && mergedRanges.length > 0) {
        Logger.log(`    Found ${mergedRanges.length} merged ranges`);
        mergedRanges.forEach(mr => mr.breakApart());
      }
    } catch (e) {
      Logger.log(`    Warning: ${e.message}`);
      range1.breakApart();
    }

    range1.clearContent();
    range1.setBackground("#ffffff");
    range1.setFontColor("black");
    Logger.log(`  ✓ Area 1 cleared`);
  }

  if (lastRow > protectedRowEnd) {
    const startRow2 = protectedRowEnd + 1;
    const clearRows2 = lastRow - protectedRowEnd;
    const range2 = sheet.getRange(startRow2, 1, clearRows2, Math.min(protectedColEnd, lastCol));

    Logger.log(`  Unmerging area 2: rows ${startRow2}-${lastRow}, cols 1-${Math.min(protectedColEnd, lastCol)}`);

    try {
      const mergedRanges = range2.getMergedRanges();
      if (mergedRanges && mergedRanges.length > 0) {
        Logger.log(`    Found ${mergedRanges.length} merged ranges`);
        mergedRanges.forEach(mr => mr.breakApart());
      }
    } catch (e) {
      Logger.log(`    Warning: ${e.message}`);
      range2.breakApart();
    }

    range2.clearContent();
    range2.setBackground("#ffffff");
    range2.setFontColor("black");
    Logger.log(`  ✓ Area 2 cleared`);
  }

  Logger.log(`✓ Maintenance area cleared (rows 1–5 and A6:K20 preserved)`);
}


// =====================================================================================
//  RESTORE BROKEN DATA AFTER GENERATE
// =====================================================================================

function restoreBrokenData() {
  Logger.log("Restoring broken data from maintenance...");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MAINTENANCE);
  
  try {
    const savedDataJson = PropertiesService.getScriptProperties().getProperty('SAVED_BROKEN_DATA');
    if (!savedDataJson) {
      Logger.log("  No saved broken data found");
      return;
    }
    
    const savedBrokenData = JSON.parse(savedDataJson);
    Logger.log(`  Found ${savedBrokenData.length} blocks to restore`);
    
    savedBrokenData.forEach(data => {
      try {
        // Validasi bahwa data masih valid
        if (!data.values || data.values.length === 0) {
          Logger.log(`  ⏩ Skipping ${data.blockName}: no data`);
          return;
        }
        
        const range = sheet.getRange(
          data.startRow,
          data.startCol,
          data.values.length,
          data.values[0].length
        );
        
        range.setValues(data.values);
        Logger.log(`  ✓ Restored broken data for ${data.blockName} (${data.orientation})`);
      } catch (err) {
        Logger.log(`  ⚠ Could not restore ${data.blockName}: ${err.message}`);
      }
    });
    
    // Clear saved data setelah restore
    PropertiesService.getScriptProperties().deleteProperty('SAVED_BROKEN_DATA');
    Logger.log("✓ All broken data restored and cleared from properties");
    
  } catch (err) {
    Logger.log(`⚠ Error restoring broken data: ${err.message}`);
  }
}


// =====================================================================================
//  HELPER: BACA PARAMETER SLOT/ROW/TIER DARI SHEET "MAINTENANCE"
// =====================================================================================

function getBlockParamsFromMaintenanceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MAINTENANCE);

  const START_ROW = 10;
  const END_COL   = 6;
  const lastRow   = sheet.getLastRow();

  const numRows = lastRow - START_ROW + 1;
  const values = sheet.getRange(START_ROW, 1, numRows, END_COL).getValues();
  
  const paramMap = {};

  values.forEach((row, idx) => {
    const blockName = row[0];
    const slot      = row[1];
    const rowCount  = row[3];
    const tier      = row[5];

    if (!blockName || blockName === "") return;

    if ([slot, rowCount, tier].some(x => isNaN(Number(x)))) {
      Logger.log(`⚠ Skipping block "${blockName}" - slot/row/tier bukan angka`);
      return;
    }

    paramMap[blockName] = {
      slotCount: Number(slot),
      rowCount:  Number(rowCount),
      tierCount: Number(tier)
    };

    Logger.log(`✓ Loaded params: ${blockName} → Slot=${slot}, Row=${rowCount}, Tier=${tier}`);
  });

  return paramMap;
}


// =====================================================================================
//  REGENERATE SEMUA BLOK DARI CONFIG
// =====================================================================================

function regenerateAllBlocksFromConfig() {
  Logger.log("=== Regenerating all blocks ===");

  const blockParamsMap = getBlockParamsFromMaintenanceSheet();

  BLOCK_CONFIGS.forEach((cfg, index) => {
    Logger.log(`[${index + 1}/${BLOCK_CONFIGS.length}] Generating ${cfg.blockName} (${cfg.orientation})...`);

    const params = blockParamsMap[cfg.blockName];
    if (!params) {
      Logger.log(`  ⚠ No numeric params found for block "${cfg.blockName}" in MAINTENANCE. Block skipped.`);
      return;
    }

    const slotCount = params.slotCount;
    const rowCount  = params.rowCount;
    const tierCount = params.tierCount;
    const reverseRow = cfg.reverseRow || false;

    if (cfg.orientation === "H") {
      generateMonitoringTableHorizontal(
        SHEET_PREFERRED,
        SHEET_MAINTENANCE,
        cfg.blockName,
        slotCount,
        rowCount,
        tierCount,
        cfg.cellStart,
        reverseRow
      );

      generateMaintenanceBlockHorizontal(
        SHEET_MAINTENANCE,
        cfg.blockName,
        slotCount,
        rowCount,
        tierCount,
        cfg.cellStart,
        reverseRow
      );

    } else if (cfg.orientation === "V") {
      generateMonitoringTableVertical(
        SHEET_PREFERRED,
        SHEET_MAINTENANCE,
        cfg.blockName,
        slotCount,
        rowCount,
        tierCount,
        cfg.cellStart,
        reverseRow
      );

      generateMaintenanceBlockVertical(
        SHEET_MAINTENANCE,
        cfg.blockName,
        slotCount,
        rowCount,
        tierCount,
        cfg.cellStart,
        reverseRow
      );
    }

    Logger.log(`  ✓ ${cfg.blockName} completed`);
  });

  Logger.log("=== All blocks regenerated successfully ===");
}


// =====================================================================================
//  DATAPIPELINE
// =====================================================================================

function DataPipeline() {
  const folderId               = "1nUbgPK5RBaSknmOLiMShPOypsZeSQEZp";
  const cells                  = ["I4", "BL4", "CB4", "DF4", "FA4", "GW4"];
  const preferredAreaSheetName = "PREFERRED AREA";
  const maintenanceSheetName   = "MAINTENANCE";

  var errorCount = 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  errorCount = importExcelAndCopySheet(folderId);
  Logger.log(errorCount);
  errorCount = extractSpecificColumns() + errorCount;
  Logger.log(errorCount);

  if (errorCount === 0) {
    ss.toast("Data Berhasil/Sudah Dimuat", "Sukses", 10);
    const preferredAreaSheet = ss.getSheetByName(preferredAreaSheetName);
    const maintenanceSheet   = ss.getSheetByName(maintenanceSheetName);

    const timestamp = new Date();
    const formatted = Utilities.formatDate(
      timestamp,
      Session.getScriptTimeZone(),
      "dd-MM-yyyy HH:mm:ss"
    );
    const message = "Last Updated: " + formatted;

    cells.forEach(cell => {
      preferredAreaSheet.getRange(cell).setValue(message);
      maintenanceSheet.getRange(cell).setValue(message);
    });
  } else {
    ss.toast("Data Gagal Dimuat", "Error", 10);
    throw new Error("Data Gagal Dimuat");
  }
}


// =====================================================================================
//  FUNGSI UTAMA TOMBOL: REFRESHDASHBOARD
// =====================================================================================

function RefreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const preferredAreaSheet = ss.getSheetByName(SHEET_PREFERRED);
  const maintenanceSheet   = ss.getSheetByName(SHEET_MAINTENANCE);

  Logger.log("\n========================================");
  Logger.log("  REFRESH DASHBOARD STARTED");
  Logger.log("========================================\n");

  try {
    // 0. Hapus HANYA CHART (diagram YOR/gauge)
    Logger.log("Step 0: Removing all charts (YOR diagrams)...");
    removeAllCharts(preferredAreaSheet);
    removeAllCharts(maintenanceSheet);
    
    SpreadsheetApp.flush();
    Logger.log("✓ Step 0 completed (charts removed & flushed)\n");

    // 1. Clear area di PREFERRED AREA
    Logger.log("Step 1: Clearing PREFERRED AREA...");
    clearDashboardArea(preferredAreaSheet);
    
    SpreadsheetApp.flush();
    Logger.log("✓ Step 1 completed (area cleared & flushed)\n");

    // 2. Clear area di MAINTENANCE + SAVE BROKEN DATA
    Logger.log("Step 2: Clearing MAINTENANCE & saving broken data...");
    clearMaintenanceArea(maintenanceSheet);
    
    SpreadsheetApp.flush();
    Logger.log("✓ Step 2 completed (maintenance cleared & broken data saved)\n");

    // 3. Jalankan pipeline data
    Logger.log("Step 3: Running DataPipeline...");
    DataPipeline();
    
    SpreadsheetApp.flush();
    Logger.log("✓ Step 3 completed (data loaded & flushed)\n");

    // 4. Generate ulang SEMUA blok
    Logger.log("Step 4: Regenerating all blocks...");
    regenerateAllBlocksFromConfig();
    
    SpreadsheetApp.flush();
    Logger.log("✓ Step 4 completed (blocks generated & flushed)\n");

    // 5. RESTORE BROKEN DATA ← FITUR BARU!
    Logger.log("Step 5: Restoring broken data...");
    restoreBrokenData();
    
    SpreadsheetApp.flush();
    Logger.log("✓ Step 5 completed (broken data restored & flushed)\n");

    Logger.log("========================================");
    Logger.log("  REFRESH DASHBOARD COMPLETED");
    Logger.log("========================================\n");

    ss.toast(
      "Refresh selesai: semua data preserved & blok tabel digenerate ulang!",
      "✓ Success",
      5
    );
    
  } catch (error) {
    Logger.log(`\n✗ ERROR: ${error.message}`);
    Logger.log(error.stack);
    
    ss.toast(
      `Refresh gagal: ${error.message}`,
      "✗ Error",
      10
    );
    
    throw error;
  }
}
