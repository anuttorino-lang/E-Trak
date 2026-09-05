var PhaseHistoryTools = (function() {
  var VERSION = '1.0.0';
  var SHEET_NAME = 'Phase_History';
  var ADMIN_SHEET_NAME = 'Admin_Input';
  var MASTER_SHEET_NAME = 'Master Tracking';
  var HEADERS = [
    'รหัสโครงการ',
    'SOP / Phase',
    'วันที่เข้า Phase',
    'วันที่ออก Phase',
    'SLA Phase (วัน)',
    'จำนวนวันที่ใช้',
    'จำนวนวันเกินกำหนด',
    'สถานะเมื่อออกจาก Phase',
    'WI สุดท้าย',
    'หมายเหตุ',
    'บันทึกเมื่อ',
  ];

  function inspectWorkbook() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return {
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      sheets: ss.getSheets().map(function(sheet) {
        return {
          name: sheet.getName(),
          sheetId: sheet.getSheetId(),
          lastRow: sheet.getLastRow(),
          lastColumn: sheet.getLastColumn(),
          maxRows: sheet.getMaxRows(),
          maxColumns: sheet.getMaxColumns(),
        };
      }),
    };
  }

  function setupPhaseHistory() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    ensureShape_(sheet);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    formatSheet_(sheet);
    return {
      success: true,
      sheetName: SHEET_NAME,
      headers: HEADERS,
    };
  }

  function recordClosedPhase(project, exitDate, finalWi, note) {
    if (!project || !project.projectCode || !project.phase) return null;
    var phase = String(project.phase || '');
    if (MasterTrackingTools.getPhaseList().indexOf(phase) === -1) return null;

    var entryDate = findCurrentPhaseEntryDate_(project.projectCode, phase) ||
      toDate_(project.phaseEntryDate);
    var closedAt = toDate_(exitDate) || new Date();
    if (!entryDate) entryDate = closedAt;

    var slaDays = Number(project.phaseTotalSla ||
      MasterTrackingTools.getPhaseTotalSla(phase) ||
      MasterTrackingTools.DEFAULT_PHASE_SLA_DAYS || 30) || 30;
    var usedDays = Math.max(1, dayDiffInclusive_(entryDate, closedAt));
    var overdueDays = Math.max(0, usedDays - slaDays);
    var status = overdueDays > 0 ? 'ล่าช้า' : 'ปกติ';

    var sheet = getOrCreateSheet_();
    var existingRow = findExistingRow_(sheet, project.projectCode, phase, entryDate);
    var values = [
      project.projectCode,
      phase,
      entryDate,
      closedAt,
      slaDays,
      usedDays,
      overdueDays ? -overdueDays : 0,
      status,
      finalWi || project.wiStep || '',
      note || '',
      new Date(),
    ];

    if (existingRow) {
      sheet.getRange(existingRow, 1, 1, values.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }
    return {
      projectCode: project.projectCode,
      phase: phase,
      usedDays: usedDays,
      overdueDays: overdueDays ? -overdueDays : 0,
      status: status,
    };
  }

  function backfillPhaseHistory() {
    setupPhaseHistory();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!adminSheet || adminSheet.getLastRow() < 2) {
      return { success: true, inserted: 0, message: 'ไม่มีข้อมูล Admin_Input สำหรับย้อนสร้างประวัติ' };
    }

    var masterMap = readMasterMap_(masterSheet);
    var rows = adminSheet.getRange(2, 1, adminSheet.getLastRow() - 1, 16).getValues();
    var grouped = {};
    rows.forEach(function(row, index) {
      var projectCode = String(row[0] || '');
      var phase = String(row[8] || '');
      if (!projectCode || MasterTrackingTools.getPhaseList().indexOf(phase) === -1) return;
      var recordedAt = toDate_(row[14]) || toDate_(row[10]);
      if (!recordedAt) return;
      grouped[projectCode] = grouped[projectCode] || [];
      grouped[projectCode].push({
        rowNumber: index + 2,
        phase: phase,
        wiStep: String(row[9] || ''),
        phaseEntryDate: toDate_(row[10]) || recordedAt,
        note: String(row[12] || ''),
        recordedAt: recordedAt,
      });
    });

    var inserted = 0;
    Object.keys(grouped).forEach(function(projectCode) {
      var logs = grouped[projectCode].sort(function(a, b) {
        var diff = a.recordedAt.getTime() - b.recordedAt.getTime();
        return diff || a.rowNumber - b.rowNumber;
      });
      var phaseGroups = [];
      logs.forEach(function(log) {
        var current = phaseGroups[phaseGroups.length - 1];
        if (!current || current.phase !== log.phase) {
          phaseGroups.push({
            phase: log.phase,
            entryDate: log.phaseEntryDate || log.recordedAt,
            lastWi: log.wiStep,
            note: log.note,
            lastRecordedAt: log.recordedAt,
          });
        } else {
          current.lastWi = log.wiStep || current.lastWi;
          current.note = log.note || current.note;
          current.lastRecordedAt = log.recordedAt;
        }
      });

      for (var i = 0; i < phaseGroups.length - 1; i += 1) {
        var currentGroup = phaseGroups[i];
        var nextGroup = phaseGroups[i + 1];
        var master = masterMap[projectCode] || {};
        recordClosedPhase({
          projectCode: projectCode,
          phase: currentGroup.phase,
          phaseEntryDate: currentGroup.entryDate,
          phaseTotalSla: MasterTrackingTools.getPhaseTotalSla(currentGroup.phase),
          wiStep: currentGroup.lastWi,
        }, nextGroup.entryDate, currentGroup.lastWi, currentGroup.note);
        inserted += 1;
      }
    });

    return { success: true, inserted: inserted, sheetName: SHEET_NAME };
  }

  function getProjectPhaseHistory(projectCode) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues()
      .filter(function(row) { return String(row[0] || '') === String(projectCode || ''); })
      .map(function(row) {
        return {
          projectCode: String(row[0] || ''),
          phase: String(row[1] || ''),
          phaseEntryDate: formatDateTime_(row[2]),
          phaseExitDate: formatDateTime_(row[3]),
          phaseSla: Number(row[4] || 0),
          usedDays: Number(row[5] || 0),
          overdueDays: Number(row[6] || 0),
          status: String(row[7] || ''),
          finalWi: String(row[8] || ''),
          note: String(row[9] || ''),
        };
      });
  }

  function cleanupUnusedEmptySheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var required = {};
    [MASTER_SHEET_NAME, ADMIN_SHEET_NAME, 'Admin_Users', SHEET_NAME].forEach(function(name) {
      required[name] = true;
    });
    var removed = [];
    ss.getSheets().forEach(function(sheet) {
      if (required[sheet.getName()]) return;
      if (sheet.getLastRow() === 0 && ss.getSheets().length > 1) {
        removed.push(sheet.getName());
        ss.deleteSheet(sheet);
      }
    });
    return { success: true, removed: removed };
  }

  function seedPhaseHistoryTestData() {
    setupPhaseHistory();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    var adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    var historySheet = ss.getSheetByName(SHEET_NAME);
    if (!masterSheet || !adminSheet || !historySheet) {
      throw new Error('ไม่พบชีตที่จำเป็นสำหรับสร้างข้อมูลทดสอบ');
    }

    var now = new Date();
    var tests = [
      {
        projectCode: 'A-TEST-101',
        documentNo: 'TEST-HISTORY/101',
        projectName: '[TEST] Phase เดิมล่าช้าและต้องคงสีแดง',
        currentPhase: 'Phase 3',
        currentWi: '1.ส่งมอบเอกสารให้ผู้รับบริการ (User)',
        receivedDate: addDays_(now, -70),
        currentEntryDate: addDays_(now, -8),
        budgetAmount: 1100000,
        note: '[TEST DATA] Phase 1 ล่าช้า -16 วัน, Phase 2 ปกติ',
        events: [
          ['Phase 1', '1.สำรวจพื้นที่', -70, 9],
          ['Phase 1', '3.แจ้งมติที่ประชุม', -30, 14],
          ['Phase 2', '1.เขียนแบบ', -24, 9],
          ['Phase 2', '6.เสนอคณบดีเห็นชอบราคาประมาณการ', -14, 15],
          ['Phase 3', '1.ส่งมอบเอกสารให้ผู้รับบริการ (User)', -8, 10],
        ],
        histories: [
          ['Phase 1', -70, -25, 46, 0, 'ปกติ', '3.แจ้งมติที่ประชุม'],
          ['Phase 2', -24, -13, 12, 0, 'ปกติ', '6.เสนอคณบดีเห็นชอบราคาประมาณการ'],
        ],
      },
      {
        projectCode: 'A-TEST-102',
        documentNo: 'TEST-HISTORY/102',
        projectName: '[TEST] Phase ปัจจุบันล่าช้า',
        currentPhase: 'Phase 2',
        currentWi: '5.คณะกรรมการพิจารณาราคาประมาณการ',
        receivedDate: addDays_(now, -55),
        currentEntryDate: addDays_(now, -40),
        budgetAmount: 850000,
        note: '[TEST DATA] Phase 1 ปกติ, Phase 2 ปัจจุบันล่าช้า',
        events: [
          ['Phase 1', '1.สำรวจพื้นที่', -55, 9],
          ['Phase 1', '3.แจ้งมติที่ประชุม', -47, 14],
          ['Phase 2', '1.เขียนแบบ', -40, 10],
          ['Phase 2', '5.คณะกรรมการพิจารณาราคาประมาณการ', -5, 15],
        ],
        histories: [
          ['Phase 1', -55, -41, 15, 0, 'ปกติ', '3.แจ้งมติที่ประชุม'],
        ],
      },
    ];

    var results = tests.map(function(test) {
      removeRowsByCode_(masterSheet, test.projectCode, 1);
      removeRowsByCode_(adminSheet, test.projectCode, 1);
      removeRowsByCode_(historySheet, test.projectCode, 1);

      test.events.forEach(function(event, index) {
        var eventDate = dateAt_(addDays_(now, event[2]), event[3], index * 7, index * 111);
        AdminInputImportTools.appendAdminInputLog({
          projectCode: test.projectCode,
          documentNo: test.documentNo,
          ownerUnit: '[TEST] หน่วยงานทดสอบ Phase History',
          receivedDate: test.receivedDate,
          projectName: test.projectName,
          budgetAmount: test.budgetAmount,
          budgetSource: 'งบทดสอบ',
          fiscalYear: '2569',
          phase: event[0],
          wiStep: event[1],
          phaseEntryDate: eventDate,
          responsibleUnit: MasterTrackingTools.getResponsibleUnitByPhase(event[0]),
          note: '[TEST DATA] Timeline จำลอง',
          recordedBy: 'PHASE_HISTORY_TEST',
          recordedAt: eventDate,
          recordStatus: 'นำเข้าแล้ว',
        });
      });

      appendTestMasterRow_(masterSheet, test);

      test.histories.forEach(function(history) {
        historySheet.appendRow([
          test.projectCode,
          history[0],
          addDays_(now, history[1]),
          addDays_(now, history[2]),
          30,
          history[3],
          history[4],
          history[5],
          history[6],
          '[TEST DATA] ประวัติ Phase จำลอง',
          new Date(),
        ]);
      });

      return {
        projectCode: test.projectCode,
        currentPhase: test.currentPhase,
        histories: test.histories.length,
      };
    });

    return {
      success: true,
      created: results,
      message: 'รีเฟรชเฉพาะ A-TEST-101 และ A-TEST-102 โดยไม่แตะข้อมูลจริง',
    };
  }

  function ensureShape_(sheet) {
    if (sheet.getMaxColumns() < HEADERS.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.length - sheet.getMaxColumns());
    }
  }

  function formatSheet_(sheet) {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setBackground('#1d4ed8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.autoResizeColumns(1, HEADERS.length);
    sheet.getRange('C:D').setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sheet.getRange('K:K').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }

  function getOrCreateSheet_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      setupPhaseHistory();
      sheet = ss.getSheetByName(SHEET_NAME);
    }
    return sheet;
  }

  function findExistingRow_(sheet, projectCode, phase, entryDate) {
    if (sheet.getLastRow() < 2) return 0;
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    var entryTime = stripTime_(entryDate).getTime();
    for (var i = 0; i < values.length; i += 1) {
      if (String(values[i][0] || '') !== String(projectCode)) continue;
      if (String(values[i][1] || '') !== String(phase)) continue;
      var existingEntry = toDate_(values[i][2]);
      if (existingEntry && stripTime_(existingEntry).getTime() === entryTime) return i + 2;
    }
    return 0;
  }

  function readMasterMap_(sheet) {
    var result = {};
    if (!sheet || sheet.getLastRow() < 2) return result;
    sheet.getRange(2, 1, sheet.getLastRow() - 1, MasterTrackingTools.MASTER_HEADERS.length)
      .getValues()
      .forEach(function(row) {
        if (!row[0]) return;
        result[String(row[0])] = {
          phaseSla: Number(row[11] || 30),
        };
      });
    return result;
  }

  function findCurrentPhaseEntryDate_(projectCode, phase) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return null;
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 15).getValues();
    var projectLogs = rows
      .map(function(row, index) {
        return {
          rowNumber: index + 2,
          projectCode: String(row[0] || ''),
          phase: String(row[8] || ''),
          entryDate: toDate_(row[10]) || toDate_(row[14]),
          recordedAt: toDate_(row[14]) || toDate_(row[10]),
        };
      })
      .filter(function(item) {
        return item.projectCode === String(projectCode) && item.entryDate;
      })
      .sort(function(a, b) {
        var diff = a.recordedAt.getTime() - b.recordedAt.getTime();
        return diff || a.rowNumber - b.rowNumber;
      });
    if (!projectLogs.length) return null;

    var first = null;
    for (var i = projectLogs.length - 1; i >= 0; i -= 1) {
      if (projectLogs[i].phase !== String(phase)) {
        if (first) break;
        continue;
      }
      first = projectLogs[i].entryDate;
    }
    return first;
  }

  function toDate_(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (!value) return null;
    var parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function stripTime_(date) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  function dayDiffInclusive_(start, end) {
    return Math.floor((stripTime_(end).getTime() - stripTime_(start).getTime()) / 86400000) + 1;
  }

  function formatDateTime_(value) {
    var date = toDate_(value);
    if (!date) return '';
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss.SSS');
  }

  function appendTestMasterRow_(sheet, test) {
    var row = sheet.getLastRow() + 1;
    var values = [
      test.projectCode,
      test.documentNo,
      '[TEST] หน่วยงานทดสอบ Phase History',
      test.receivedDate,
      test.projectName,
      test.budgetAmount,
      'งบทดสอบ',
      '2569',
      test.currentPhase,
      test.currentWi,
      test.currentEntryDate,
      30,
      '',
      '',
      MasterTrackingTools.getResponsibleUnitByPhase(test.currentPhase),
      '',
      180,
      '',
      '',
      '',
      test.note,
      new Date(),
    ];
    sheet.getRange(row, 1, 1, values.length).setValues([values]);
    MasterTrackingTools.setWiValidationForRow(row, test.currentPhase);
    MasterTrackingTools.setMasterFormulasForRow(row);
  }

  function removeRowsByCode_(sheet, projectCode, codeColumn) {
    if (!sheet || sheet.getLastRow() < 2) return;
    var values = sheet.getRange(2, codeColumn, sheet.getLastRow() - 1, 1).getValues();
    for (var index = values.length - 1; index >= 0; index -= 1) {
      if (String(values[index][0] || '') === String(projectCode)) {
        sheet.deleteRow(index + 2);
      }
    }
  }

  function addDays_(date, days) {
    var result = new Date(date.getTime());
    result.setDate(result.getDate() + days);
    return result;
  }

  function dateAt_(date, hour, minute, millisecond) {
    var result = new Date(date.getTime());
    result.setHours(hour, minute, 0, millisecond || 0);
    return result;
  }

  return {
    VERSION: VERSION,
    SHEET_NAME: SHEET_NAME,
    inspectWorkbook: inspectWorkbook,
    setupPhaseHistory: setupPhaseHistory,
    recordClosedPhase: recordClosedPhase,
    backfillPhaseHistory: backfillPhaseHistory,
    getProjectPhaseHistory: getProjectPhaseHistory,
    cleanupUnusedEmptySheets: cleanupUnusedEmptySheets,
    seedPhaseHistoryTestData: seedPhaseHistoryTestData,
  };
})();

function inspectWorkbookForPhaseHistory() {
  return PhaseHistoryTools.inspectWorkbook();
}

function setupAndBackfillPhaseHistory() {
  return {
    setup: PhaseHistoryTools.setupPhaseHistory(),
    backfill: PhaseHistoryTools.backfillPhaseHistory(),
    cleanup: PhaseHistoryTools.cleanupUnusedEmptySheets(),
    workbook: PhaseHistoryTools.inspectWorkbook(),
  };
}

function seedPhaseHistoryTestData() {
  return PhaseHistoryTools.seedPhaseHistoryTestData();
}
