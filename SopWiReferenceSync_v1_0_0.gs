var SopWiReferenceSync = (function() {
  var VERSION = '1.0.0';
  var SOP_WI_SHEET_NAME = 'SOP_WI_Master';
  var SLA_PHASE_SHEET_NAME = 'SLA_Phase';
  var MASTER_SHEET_NAME = 'Master Tracking';
  var ADMIN_SHEET_NAME = 'Admin_Input';
  var PHASE_HISTORY_SHEET_NAME = 'Phase_History';

  function syncFromSla2Source() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sopWiSheet = ss.getSheetByName(SOP_WI_SHEET_NAME) || ss.insertSheet(SOP_WI_SHEET_NAME);
    var slaPhaseSheet = ss.getSheetByName(SLA_PHASE_SHEET_NAME) || ss.insertSheet(SLA_PHASE_SHEET_NAME);

    writeSopWiMaster_(sopWiSheet);
    writeSlaPhase_(slaPhaseSheet);
    var masterResult = migrateMaster_(ss.getSheetByName(MASTER_SHEET_NAME));
    var adminResult = migrateAdminInput_(ss.getSheetByName(ADMIN_SHEET_NAME));
    var historyResult = refreshPhaseHistory_(ss.getSheetByName(PHASE_HISTORY_SHEET_NAME));

    return {
      success: true,
      version: VERSION,
      source: '7 กำหนด SLA (E - Track)1.xlsx / SLA (2)',
      sopWiRows: countWiRows_(),
      phaseRows: MasterTrackingTools.PHASES.length,
      master: masterResult,
      adminInput: adminResult,
      phaseHistory: historyResult,
      totalMainSla: MasterTrackingTools.DEFAULT_MAIN_SLA_DAYS,
    };
  }

  function writeSopWiMaster_(sheet) {
    var headers = ['ผู้รับผิดชอบ', 'Phase', 'SOP', 'ลำดับ WI', 'WI', 'SLA WI (วันทำการ)'];
    var rows = [];
    MasterTrackingTools.PHASES.forEach(function(phase) {
      var wiList = MasterTrackingTools.getWiListByPhase(phase);
      wiList.forEach(function(wi, index) {
        rows.push([
          MasterTrackingTools.getResponsibleUnitByPhase(phase),
          phase,
          MasterTrackingTools.PHASE_DEFINITIONS[phase] || '',
          index + 1,
          wi,
          MasterTrackingTools.getSlaDays(phase, wi),
        ]);
      });
    });
    replaceSheetData_(sheet, headers, rows);
    sheet.getRange(2, 6, Math.max(rows.length, 1), 1).setNumberFormat('0');
  }

  function writeSlaPhase_(sheet) {
    var headers = ['Phase', 'SOP', 'ผู้รับผิดชอบ', 'จำนวน WI', 'SLA รวม Phase (วันทำการ)', 'หมายเหตุ'];
    var rows = MasterTrackingTools.PHASES.map(function(phase) {
      return [
        phase,
        MasterTrackingTools.PHASE_DEFINITIONS[phase] || '',
        MasterTrackingTools.getResponsibleUnitByPhase(phase),
        MasterTrackingTools.getWiListByPhase(phase).length,
        MasterTrackingTools.getPhaseTotalSla(phase),
        phase === 'Phase 5' ? 'รวมจาก SLA ราย WI = 120 วัน (ไฟล์ต้นทางช่องรวมระบุ 90 วัน)' : '',
      ];
    });
    replaceSheetData_(sheet, headers, rows);
  }

  function replaceSheetData_(sheet, headers, rows) {
    sheet.clear();
    var requiredRows = Math.max(rows.length + 1, 2);
    var requiredColumns = headers.length;
    if (sheet.getMaxRows() < requiredRows) {
      sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
    }
    if (sheet.getMaxColumns() < requiredColumns) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1d4ed8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
  }

  function migrateMaster_(sheet) {
    if (!sheet || sheet.getLastRow() < 2) return { updated: 0 };
    var rowCount = sheet.getLastRow() - 1;
    var values = sheet.getRange(2, 1, rowCount, MasterTrackingTools.MASTER_HEADERS.length).getValues();
    var updated = 0;
    values.forEach(function(row, index) {
      var phase = String(row[8] || '');
      if (MasterTrackingTools.PHASES.indexOf(phase) === -1) return;
      var normalizedWi = MasterTrackingTools.normalizeWiStep(phase, row[9]);
      if (!normalizedWi || ['เสร็จสิ้น', 'ยกเลิกรายการ'].indexOf(normalizedWi) !== -1) return;
      var sheetRow = index + 2;
      sheet.getRange(sheetRow, 10).setValue(normalizedWi);
      sheet.getRange(sheetRow, 12).setValue(MasterTrackingTools.getSlaDays(phase, normalizedWi));
      sheet.getRange(sheetRow, 17).setValue(MasterTrackingTools.DEFAULT_MAIN_SLA_DAYS);
      MasterTrackingTools.setWiValidationForRow(sheetRow, phase);
      MasterTrackingTools.setMasterFormulasForRow(sheetRow);
      updated += 1;
    });
    return { updated: updated };
  }

  function migrateAdminInput_(sheet) {
    if (!sheet || sheet.getLastRow() < 2) return { updated: 0 };
    var rowCount = sheet.getLastRow() - 1;
    var values = sheet.getRange(2, 9, rowCount, 2).getValues();
    var updated = 0;
    values.forEach(function(row, index) {
      var phase = String(row[0] || '');
      var wi = String(row[1] || '');
      if (MasterTrackingTools.PHASES.indexOf(phase) === -1) return;
      var normalizedWi = MasterTrackingTools.normalizeWiStep(phase, wi);
      if (normalizedWi && normalizedWi !== wi) {
        sheet.getRange(index + 2, 10).setValue(normalizedWi);
        updated += 1;
      }
    });
    return { updated: updated };
  }

  function refreshPhaseHistory_(sheet) {
    if (!sheet || sheet.getLastRow() < 2) return { updated: 0 };
    var rowCount = sheet.getLastRow() - 1;
    var values = sheet.getRange(2, 1, rowCount, 11).getValues();
    var updated = 0;
    values.forEach(function(row, index) {
      var phase = String(row[1] || '');
      if (MasterTrackingTools.PHASES.indexOf(phase) === -1) return;
      var usedDays = Number(row[5] || 0);
      var totalSla = MasterTrackingTools.getPhaseTotalSla(phase);
      var overdue = Math.max(0, usedDays - totalSla);
      var normalizedFinalWi = MasterTrackingTools.normalizeWiStep(phase, row[8]);
      sheet.getRange(index + 2, 5).setValue(totalSla);
      sheet.getRange(index + 2, 7).setValue(overdue ? -overdue : 0);
      sheet.getRange(index + 2, 8).setValue(overdue ? 'ล่าช้า' : 'ปกติ');
      if (normalizedFinalWi) sheet.getRange(index + 2, 9).setValue(normalizedFinalWi);
      updated += 1;
    });
    return { updated: updated };
  }

  function countWiRows_() {
    return MasterTrackingTools.PHASES.reduce(function(total, phase) {
      return total + MasterTrackingTools.getWiListByPhase(phase).length;
    }, 0);
  }

  return {
    VERSION: VERSION,
    syncFromSla2Source: syncFromSla2Source,
  };
})();

function syncSopWiAndSlaFromSla2Source() {
  return SopWiReferenceSync.syncFromSla2Source();
}
