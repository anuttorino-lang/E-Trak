var AdminInputImportTools = (function() {
  var VERSION = '1.0.2';
  var ADMIN_SHEET_NAME = 'Admin_Input';
  var MASTER_SHEET_NAME = 'Master Tracking';
  var PENDING_STATUS = 'รอตรวจสอบ/นำเข้า Master';
  var IMPORTED_STATUS = 'นำเข้าแล้ว';
  var FAILED_STATUS = 'นำเข้าไม่สำเร็จ';
  var FIXED_STATUS = 'แก้ไขแล้ว';
  var REJECTED_STATUS = '❌ ไม่อนุมัติ';

  var ADMIN_HEADERS = [
    'รหัสโครงการ',
    'เลขที่หนังสือ',
    'หน่วยงานเจ้าของเรื่อง / ต้นเรื่อง',
    'วันที่รับเรื่อง',
    'ชื่อโครงการ / รายการ',
    'วงเงินงบประมาณ',
    'แหล่งงบประมาณ',
    'ปีงบประมาณ',
    'SOP / Phase',
    'WI / ขั้นตอนย่อย',
    'วันที่รับเข้า Phase นี้',
    'ฝ่ายรับผิดชอบ',
    'หมายเหตุ',
    'ผู้บันทึก',
    'วันที่บันทึก',
    'สถานะบันทึก',
  ];

  function setupAdminInput() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet_(ss, ADMIN_SHEET_NAME);

    ensureAdminShape_(sheet);
    sheet.getRange(1, 1, 1, ADMIN_HEADERS.length).setValues([ADMIN_HEADERS]);
    formatAdminSheet_(sheet);
    setAdminValidation_(sheet);

    return {
      success: true,
      version: VERSION,
      sheetName: ADMIN_SHEET_NAME,
    };
  }

  function importPendingAdminInputToMasterTracking() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

    if (!adminSheet || !masterSheet) {
      throw new Error('ไม่พบชีต Admin_Input หรือ Master Tracking');
    }

    var lastRow = adminSheet.getLastRow();
    if (lastRow < 2) {
      return { imported: 0, failed: 0 };
    }

    var rows = adminSheet.getRange(2, 1, lastRow - 1, ADMIN_HEADERS.length).getValues();
    var imported = 0;
    var failed = 0;

    rows.forEach(function(row, index) {
      var sheetRow = index + 2;
      var status = row[15];

      if (status !== PENDING_STATUS) {
        return;
      }

      try {
        validateAdminRow_(row);

        if (!row[0]) {
          row[0] = MasterTrackingTools.generateNextProjectCode();
          adminSheet.getRange(sheetRow, 1).setValue(row[0]);
        }

        upsertMasterFromAdminRow_(masterSheet, row);
        adminSheet.getRange(sheetRow, 16).setValue(IMPORTED_STATUS);
        imported += 1;
      } catch (error) {
        adminSheet.getRange(sheetRow, 16).setValue(FAILED_STATUS);
        failed += 1;
      }
    });

    return { imported: imported, failed: failed };
  }

  function appendAdminInputLog(payload) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    if (!adminSheet) {
      setupAdminInput();
      adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    }

    var phase = payload.phase || 'Phase 1';
    var now = new Date();
    var row = [
      payload.projectCode || '',
      payload.documentNo || '',
      payload.ownerUnit || '',
      toDateOrDefault_(payload.receivedDate, now),
      payload.projectName || '',
      valueOrBlank_(payload.budgetAmount),
      payload.budgetSource || '',
      payload.fiscalYear || '',
      phase,
      payload.wiStep || getDefaultWi_(phase),
      toDateOrDefault_(payload.phaseEntryDate, payload.receivedDate || now),
      payload.responsibleUnit || MasterTrackingTools.getResponsibleUnitByPhase(phase),
      payload.note || '',
      payload.recordedBy || getActiveUserEmail_(),
      payload.recordedAt || now,
      payload.recordStatus || PENDING_STATUS,
    ];

    adminSheet.appendRow(row);
    return adminSheet.getLastRow();
  }

  function upsertMasterFromAdminRow_(masterSheet, adminRow) {
    var projectCode = adminRow[0];
    var rowNumber = findMasterRowByProjectCode_(masterSheet, projectCode);
    var targetRow = rowNumber || Math.max(masterSheet.getLastRow() + 1, 2);
    var existing = rowNumber
      ? masterSheet.getRange(rowNumber, 1, 1, MasterTrackingTools.MASTER_HEADERS.length).getValues()[0]
      : [];

    var receivedDate = existing[3] || adminRow[3];
    var phaseEntryDate = adminRow[10] || existing[10] || receivedDate;
    var normalizedWi = MasterTrackingTools.normalizeWiStep(adminRow[8], adminRow[9]);
    var phaseSla = MasterTrackingTools.getSlaDays(adminRow[8], normalizedWi);
    var mainSla = existing[16] || MasterTrackingTools.DEFAULT_MAIN_SLA_DAYS;
    var updatedAt = adminRow[14] || new Date();

    masterSheet.getRange(targetRow, 1).setValue(adminRow[0]);
    masterSheet.getRange(targetRow, 2).setValue(adminRow[1]);
    masterSheet.getRange(targetRow, 3).setValue(adminRow[2]);
    masterSheet.getRange(targetRow, 4).setValue(receivedDate);
    masterSheet.getRange(targetRow, 5).setValue(adminRow[4]);
    masterSheet.getRange(targetRow, 6).setValue(adminRow[5]);
    masterSheet.getRange(targetRow, 7).setValue(adminRow[6]);
    masterSheet.getRange(targetRow, 8).setValue(adminRow[7]);
    masterSheet.getRange(targetRow, 9).setValue(adminRow[8]);
    MasterTrackingTools.setWiValidationForRow(targetRow, adminRow[8]);
    masterSheet.getRange(targetRow, 10).setValue(normalizedWi);
    masterSheet.getRange(targetRow, 11).setValue(phaseEntryDate);
    masterSheet.getRange(targetRow, 12).setValue(phaseSla);
    masterSheet.getRange(targetRow, 17).setValue(mainSla);
    masterSheet.getRange(targetRow, 21).setValue(adminRow[12]);
    masterSheet.getRange(targetRow, 22).setValue(updatedAt);

    MasterTrackingTools.setMasterFormulasForRow(targetRow);
  }


  function moveProjectToConsiderationSafe(payload) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

    if (!adminSheet || !masterSheet) {
      throw new Error('ไม่พบชีต Admin_Input หรือ Master Tracking');
    }

    var projectCode = payload.projectCode;
    if (!projectCode) {
      throw new Error('ไม่พบรหัสโครงการ');
    }

    var rowNumber = findMasterRowByProjectCode_(masterSheet, projectCode);
    if (!rowNumber) {
      throw new Error('ไม่พบโครงการใน Master Tracking');
    }

    var masterRow = masterSheet.getRange(rowNumber, 1, 1, 22).getValues()[0];
    var currentPhase = String(masterRow[8] || '').trim();
    var currentWi = String(masterRow[9] || '').trim();

    if (currentPhase !== 'Phase 1' ||
        MasterTrackingTools.normalizeWiStep(currentPhase, currentWi) !== '2.เสนอคณะทำงานกลั่นกรองพิจารณา') {
      throw new Error('รายการนี้ต้องอยู่ที่ Phase 1 / 2.เสนอคณะทำงานกลั่นกรองพิจารณา ก่อน');
    }

    var now = new Date();
    var nextWi = '3.แจ้งมติที่ประชุม';
    var note = payload.note || 'ส่งต่อเข้าสู่ขั้นตอนพิจารณารายการ';
    var responsibleUnit = 'งานอาคารสถานที่';

    // บันทึก Log ลง Admin_Input โดยไม่เรียก validation เดิม
    adminSheet.appendRow([
      masterRow[0], masterRow[1], masterRow[2], masterRow[3],
      masterRow[4], masterRow[5], masterRow[6], masterRow[7],
      'Phase 1',
      nextWi,
      now,
      responsibleUnit,
      note,
      payload.recordedBy || getActiveUserEmail_(),
      now,
      IMPORTED_STATUS,
    ]);

    // อัปเดต Master Tracking โดยตรง ไม่เรียก transferProjectPhase() เดิม
    masterSheet.getRange(rowNumber, 9).setValue('Phase 1');      // I
    masterSheet.getRange(rowNumber, 10).setValue(nextWi);        // J
    masterSheet.getRange(rowNumber, 11).setValue(now);           // K
    masterSheet.getRange(rowNumber, 12).setValue(MasterTrackingTools.getSlaDays('Phase 1', nextWi)); // L
    masterSheet.getRange(rowNumber, 15).setValue(responsibleUnit); // O
    masterSheet.getRange(rowNumber, 21).setValue(note);          // U
    masterSheet.getRange(rowNumber, 22).setValue(now);           // V

    // ใส่สูตรแบบปลอดภัยเฉพาะแถวนี้
    masterSheet.getRange(rowNumber, 13).setFormula('=IF($K' + rowNumber + '="","",$L' + rowNumber + '-(INT(TODAY())-INT($K' + rowNumber + ')+1))');
    masterSheet.getRange(rowNumber, 14).setFormula('=IF($J' + rowNumber + '="","",IF(REGEXMATCH($J' + rowNumber + '&$U' + rowNumber + ',"ไม่อนุมัติ"),"❌ ไม่อนุมัติ",IF($J' + rowNumber + '="เสร็จสิ้น","🔵 เสร็จสิ้น",IF(REGEXMATCH($J' + rowNumber + ',"ยกเลิก"),"⚫ ยกเลิกรายการ",IF($M' + rowNumber + '<0,"🔴 ล่าช้า",IF($M' + rowNumber + '<=5,"🟡 ใกล้กำหนด","🟢 ปกติ"))))))');
    masterSheet.getRange(rowNumber, 18).setFormula('=IF($D' + rowNumber + '="","",INT(TODAY())-INT($D' + rowNumber + ')+1)');
    masterSheet.getRange(rowNumber, 19).setFormula('=IF($D' + rowNumber + '="","",$Q' + rowNumber + '-$R' + rowNumber + ')');
    masterSheet.getRange(rowNumber, 20).setFormula('=IF($A' + rowNumber + '="","",IF(REGEXMATCH($N' + rowNumber + '&$U' + rowNumber + ',"ไม่อนุมัติ"),"❌ ไม่อนุมัติ",IF(REGEXMATCH($J' + rowNumber + ',"ยกเลิก"),"⚫ ยกเลิกรายการ",IF($J' + rowNumber + '="เสร็จสิ้น","🔵 เสร็จสิ้น",IF($S' + rowNumber + '<0,"🔴 ล่าช้า",IF($S' + rowNumber + '<=10,"🟡 ใกล้กำหนด","🟢 ปกติ"))))))');

    return {
      success: true,
      projectCode: projectCode,
      phase: 'Phase 1',
      wiStep: nextWi,
      responsibleUnit: responsibleUnit
    };
  }

  function rejectProjectFromConsiderationSafe(payload) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

    if (!adminSheet || !masterSheet) {
      throw new Error('ไม่พบชีต Admin_Input หรือ Master Tracking');
    }

    var projectCode = payload.projectCode;
    if (!projectCode) {
      throw new Error('ไม่พบรหัสโครงการ');
    }

    var rowNumber = findMasterRowByProjectCode_(masterSheet, projectCode);
    if (!rowNumber) {
      throw new Error('ไม่พบโครงการใน Master Tracking');
    }

    var masterRow = masterSheet.getRange(rowNumber, 1, 1, 22).getValues()[0];
    var currentPhase = String(masterRow[8] || '').trim();
    var currentWi = String(masterRow[9] || '').trim();

    if (currentPhase !== 'Phase 1' ||
        MasterTrackingTools.normalizeWiStep(currentPhase, currentWi) !== '2.เสนอคณะทำงานกลั่นกรองพิจารณา') {
      throw new Error('รายการนี้ต้องอยู่ที่ Phase 1 / 2.เสนอคณะทำงานกลั่นกรองพิจารณา ก่อน');
    }

    var now = new Date();
    var note = payload.note || 'ไม่อนุมัติรายการ';
    var responsibleUnit = masterRow[14] || 'งานอาคารสถานที่';

    adminSheet.appendRow([
      masterRow[0], masterRow[1], masterRow[2], masterRow[3],
      masterRow[4], masterRow[5], masterRow[6], masterRow[7],
      currentPhase,
      '2.เสนอคณะทำงานกลั่นกรองพิจารณา',
      masterRow[10] || now,
      responsibleUnit,
      note,
      payload.recordedBy || getActiveUserEmail_(),
      now,
      IMPORTED_STATUS,
    ]);

    masterSheet.getRange(rowNumber, 10).setValue('2.เสนอคณะทำงานกลั่นกรองพิจารณา'); // J
    masterSheet.getRange(rowNumber, 14).setValue(REJECTED_STATUS);    // N
    masterSheet.getRange(rowNumber, 21).setValue(note);               // U
    masterSheet.getRange(rowNumber, 22).setValue(now);                // V

    return {
      success: true,
      projectCode: projectCode,
      status: REJECTED_STATUS
    };
  }



  function approveProjectFromConsiderationSafe(payload) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

    if (!adminSheet || !masterSheet) {
      throw new Error('ไม่พบชีต Admin_Input หรือ Master Tracking');
    }

    var projectCode = payload.projectCode;
    if (!projectCode) throw new Error('ไม่พบรหัสโครงการ');

    var rowNumber = findMasterRowByProjectCode_(masterSheet, projectCode);
    if (!rowNumber) throw new Error('ไม่พบโครงการใน Master Tracking');

    var masterRow = masterSheet.getRange(rowNumber, 1, 1, 22).getValues()[0];
    var currentPhase = String(masterRow[8] || '').trim();
    var currentWi = String(masterRow[9] || '').trim();

    if (currentPhase !== 'Phase 1' ||
        MasterTrackingTools.normalizeWiStep(currentPhase, currentWi) !== '2.เสนอคณะทำงานกลั่นกรองพิจารณา') {
      throw new Error('รายการนี้ต้องอยู่ที่ Phase 1 / 2.เสนอคณะทำงานกลั่นกรองพิจารณา ก่อนอนุมัติ');
    }

    var now = new Date();
    var note = payload.note || 'อนุมัติรายการและส่งต่อ Phase 2';
    var nextPhase = 'Phase 2';
    var nextWi = '1.เขียนแบบ';
    var responsibleUnit = 'งานอาคารสถานที่';

    PhaseHistoryTools.recordClosedPhase({
      projectCode: masterRow[0],
      phase: currentPhase,
      phaseEntryDate: masterRow[10],
      phaseSla: masterRow[11],
      wiStep: currentWi,
    }, now, currentWi, note);

    adminSheet.appendRow([
      masterRow[0], masterRow[1], masterRow[2], masterRow[3],
      masterRow[4], masterRow[5], masterRow[6], masterRow[7],
      currentPhase,
      MasterTrackingTools.normalizeWiStep(currentPhase, currentWi),
      masterRow[10] || now,
      masterRow[14] || responsibleUnit,
      note,
      payload.recordedBy || getActiveUserEmail_(),
      now,
      IMPORTED_STATUS,
    ]);

    adminSheet.appendRow([
      masterRow[0], masterRow[1], masterRow[2], masterRow[3],
      masterRow[4], masterRow[5], masterRow[6], masterRow[7],
      nextPhase,
      nextWi,
      now,
      responsibleUnit,
      'ส่งต่อขั้นตอนถัดไป',
      payload.recordedBy || getActiveUserEmail_(),
      now,
      IMPORTED_STATUS,
    ]);

    masterSheet.getRange(rowNumber, 9).setValue(nextPhase);        // I
    masterSheet.getRange(rowNumber, 10).setValue(nextWi);          // J
    masterSheet.getRange(rowNumber, 11).setValue(now);             // K
    masterSheet.getRange(rowNumber, 12).setValue(MasterTrackingTools.getSlaDays(nextPhase, nextWi)); // L
    masterSheet.getRange(rowNumber, 15).setValue(responsibleUnit); // O
    masterSheet.getRange(rowNumber, 21).setValue(note);            // U
    masterSheet.getRange(rowNumber, 22).setValue(now);             // V

    masterSheet.getRange(rowNumber, 13).setFormula('=IF($K' + rowNumber + '="","",$L' + rowNumber + '-(INT(TODAY())-INT($K' + rowNumber + ')+1))');
    masterSheet.getRange(rowNumber, 14).setFormula('=IF($J' + rowNumber + '="","",IF(REGEXMATCH($U' + rowNumber + '&$J' + rowNumber + ',"ไม่อนุมัติ"),"❌ ไม่อนุมัติ",IF($J' + rowNumber + '="เสร็จสิ้น","🔵 เสร็จสิ้น",IF(REGEXMATCH($J' + rowNumber + ',"ยกเลิก"),"⚫ ยกเลิกรายการ",IF($M' + rowNumber + '<0,"🔴 ล่าช้า",IF($M' + rowNumber + '<=5,"🟡 ใกล้กำหนด","🟢 ปกติ"))))))');
    masterSheet.getRange(rowNumber, 18).setFormula('=IF($D' + rowNumber + '="","",INT(TODAY())-INT($D' + rowNumber + ')+1)');
    masterSheet.getRange(rowNumber, 19).setFormula('=IF($D' + rowNumber + '="","",$Q' + rowNumber + '-$R' + rowNumber + ')');
    masterSheet.getRange(rowNumber, 20).setFormula('=IF($A' + rowNumber + '="","",IF(REGEXMATCH($N' + rowNumber + '&$U' + rowNumber + ',"ไม่อนุมัติ"),"❌ ไม่อนุมัติ",IF(REGEXMATCH($J' + rowNumber + ',"ยกเลิก"),"⚫ ยกเลิกรายการ",IF($J' + rowNumber + '="เสร็จสิ้น","🔵 เสร็จสิ้น",IF($S' + rowNumber + '<0,"🔴 ล่าช้า",IF($S' + rowNumber + '<=10,"🟡 ใกล้กำหนด","🟢 ปกติ"))))))');

    return {
      success: true,
      projectCode: projectCode,
      phase: nextPhase,
      wiStep: nextWi,
      responsibleUnit: responsibleUnit
    };
  }

  function rejectProjectFromConsiderationKeepVisibleSafe(payload) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

    if (!adminSheet || !masterSheet) {
      throw new Error('ไม่พบชีต Admin_Input หรือ Master Tracking');
    }

    var projectCode = payload.projectCode;
    if (!projectCode) throw new Error('ไม่พบรหัสโครงการ');

    var rowNumber = findMasterRowByProjectCode_(masterSheet, projectCode);
    if (!rowNumber) throw new Error('ไม่พบโครงการใน Master Tracking');

    var masterRow = masterSheet.getRange(rowNumber, 1, 1, 22).getValues()[0];
    var currentPhase = String(masterRow[8] || '').trim();
    var currentWi = String(masterRow[9] || '').trim();

    if (currentPhase !== 'Phase 1' ||
        MasterTrackingTools.normalizeWiStep(currentPhase, currentWi) !== '2.เสนอคณะทำงานกลั่นกรองพิจารณา') {
      throw new Error('รายการนี้ต้องอยู่ที่ Phase 1 / 2.เสนอคณะทำงานกลั่นกรองพิจารณา ก่อนบันทึกไม่อนุมัติ');
    }

    var now = new Date();
    var note = payload.note || 'ไม่อนุมัติรายการ';
    var responsibleUnit = masterRow[14] || 'งานอาคารสถานที่';

    adminSheet.appendRow([
      masterRow[0], masterRow[1], masterRow[2], masterRow[3],
      masterRow[4], masterRow[5], masterRow[6], masterRow[7],
      currentPhase,
      '2.เสนอคณะทำงานกลั่นกรองพิจารณา',
      masterRow[10] || now,
      responsibleUnit,
      note,
      payload.recordedBy || getActiveUserEmail_(),
      now,
      IMPORTED_STATUS,
    ]);

    // สำคัญ: ไม่ลบ ไม่ย้าย ไม่ปิดรายการ
    // คง Phase/WI เดิมเพื่อให้อยู่ในรายการทั้งหมด และให้ dashboard นับรวมในกลุ่มไม่อนุมัติ
    masterSheet.getRange(rowNumber, 9).setValue('Phase 1');              // I
    masterSheet.getRange(rowNumber, 10).setValue('2.เสนอคณะทำงานกลั่นกรองพิจารณา'); // J
    masterSheet.getRange(rowNumber, 14).setValue(REJECTED_STATUS);       // N
    masterSheet.getRange(rowNumber, 20).setValue(REJECTED_STATUS);       // T
    masterSheet.getRange(rowNumber, 21).setValue(note);                  // U
    masterSheet.getRange(rowNumber, 22).setValue(now);                   // V

    return {
      success: true,
      projectCode: projectCode,
      status: REJECTED_STATUS,
      keepInAllProjects: true
    };
  }


  function validateAdminRow_(row) {
    if (!row[3]) throw new Error('วันที่รับเรื่องว่าง');
    if (!row[4]) throw new Error('ชื่อโครงการ / รายการว่าง');
    if (MasterTrackingTools.getPhaseList().indexOf(row[8]) === -1) {
      throw new Error('SOP / Phase ไม่ถูกต้อง');
    }
    if (!isAllowedWiForImport_(row[8], row[9])) {
      throw new Error('WI ไม่ตรงกับ Phase');
    }
  }

  function isAllowedWiForImport_(phase, wiStep) {
    var wi = String(wiStep || '');
    if (wi === 'เสร็จสิ้น' || wi === 'ยกเลิกรายการ') return true;
    return MasterTrackingTools.getWiListByPhase(phase)
      .indexOf(MasterTrackingTools.normalizeWiStep(phase, wiStep)) !== -1;
  }

  function findMasterRowByProjectCode_(sheet, projectCode) {
    if (!projectCode || sheet.getLastRow() < 2) {
      return null;
    }

    var codes = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < codes.length; i += 1) {
      if (String(codes[i][0]) === String(projectCode)) {
        return i + 2;
      }
    }
    return null;
  }

  function ensureAdminShape_(sheet) {
    if (sheet.getMaxColumns() < ADMIN_HEADERS.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), ADMIN_HEADERS.length - sheet.getMaxColumns());
    }
    if (sheet.getMaxColumns() > ADMIN_HEADERS.length) {
      sheet.deleteColumns(ADMIN_HEADERS.length + 1, sheet.getMaxColumns() - ADMIN_HEADERS.length);
    }
    if (sheet.getMaxRows() < 500) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 500 - sheet.getMaxRows());
    }
  }

  function formatAdminSheet_(sheet) {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, ADMIN_HEADERS.length)
      .setFontWeight('bold')
      .setFontColor('#ffffff')
      .setBackground('#075f6c')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, ADMIN_HEADERS.length)
      .setVerticalAlignment('middle')
      .setWrap(true);
    sheet.getRange(2, 4, sheet.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(2, 6, sheet.getMaxRows() - 1, 1).setNumberFormat('#,##0.00');
    sheet.getRange(2, 11, sheet.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(2, 15, sheet.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  }

  function setAdminValidation_(sheet) {
    var maxRows = sheet.getMaxRows() - 1;
    var phaseRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(MasterTrackingTools.getPhaseList(), true)
      .setAllowInvalid(false)
      .build();
    var allWiList = MasterTrackingTools.getAllWiList().slice();
    if (allWiList.indexOf('เสร็จสิ้น') === -1) allWiList.push('เสร็จสิ้น');
    if (allWiList.indexOf('ยกเลิกรายการ') === -1) allWiList.push('ยกเลิกรายการ');
    var wiRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(allWiList, true)
      .setAllowInvalid(true)
      .build();
    var unitRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['งานอาคารสถานที่', 'งานนโยบายและแผน', 'งานพัสดุและยานพาหนะ', '-'], true)
      .setAllowInvalid(false)
      .build();
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([PENDING_STATUS, IMPORTED_STATUS, FIXED_STATUS, FAILED_STATUS], true)
      .setAllowInvalid(false)
      .build();
    var dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .build();

    sheet.getRange(2, 4, maxRows, 1).setDataValidation(dateRule);
    sheet.getRange(2, 9, maxRows, 1).setDataValidation(phaseRule);
    sheet.getRange(2, 10, maxRows, 1).setDataValidation(wiRule);
    sheet.getRange(2, 11, maxRows, 1).setDataValidation(dateRule);
    sheet.getRange(2, 12, maxRows, 1).setDataValidation(unitRule);
    sheet.getRange(2, 16, maxRows, 1).setDataValidation(statusRule);
  }

  function getDefaultWi_(phase) {
    return MasterTrackingTools.getWiListByPhase(phase)[0] || '';
  }

  function getOrCreateSheet_(ss, sheetName) {
    return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  }

  function toDateOrDefault_(value, fallback) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
      var parts = value.split('-').map(Number);
      return new Date(parts[0], parts[1] - 1, parts[2]);
    }
    if (value) {
      var parsed = new Date(value);
      if (!Number.isNaN(parsed.getTime())) return parsed;
    }
    if (typeof fallback === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(fallback)) {
      var fallbackParts = fallback.split('-').map(Number);
      return new Date(fallbackParts[0], fallbackParts[1] - 1, fallbackParts[2]);
    }
    return fallback instanceof Date ? fallback : new Date(fallback);
  }

  function valueOrBlank_(value) {
    return value === null || value === undefined ? '' : value;
  }

  function getActiveUserEmail_() {
    return Session.getActiveUser().getEmail() || '';
  }

  return {
    VERSION: VERSION,
    PENDING_STATUS: PENDING_STATUS,
    IMPORTED_STATUS: IMPORTED_STATUS,
    FAILED_STATUS: FAILED_STATUS,
    ADMIN_HEADERS: ADMIN_HEADERS,
    setupAdminInput: setupAdminInput,
    importPendingAdminInputToMasterTracking: importPendingAdminInputToMasterTracking,
    appendAdminInputLog: appendAdminInputLog,
    approveProjectFromConsiderationSafe: approveProjectFromConsiderationSafe,
    rejectProjectFromConsiderationKeepVisibleSafe: rejectProjectFromConsiderationKeepVisibleSafe,
    moveProjectToConsiderationSafe: moveProjectToConsiderationSafe,
    rejectProjectFromConsiderationSafe: rejectProjectFromConsiderationSafe,
  };
})();

function setupAdminInputTools() {
  return AdminInputImportTools.setupAdminInput();
}

function importPendingAdminInputToMasterTracking() {
  return AdminInputImportTools.importPendingAdminInputToMasterTracking();
}


function moveProjectToConsiderationSafe(payload) {
  return AdminInputImportTools.moveProjectToConsiderationSafe(payload);
}

function rejectProjectFromConsiderationSafe(payload) {
  return AdminInputImportTools.rejectProjectFromConsiderationSafe(payload);
}


function approveProjectFromConsiderationSafe(payload) {
  return AdminInputImportTools.approveProjectFromConsiderationSafe(payload);
}

function rejectProjectFromConsiderationKeepVisibleSafe(payload) {
  return AdminInputImportTools.rejectProjectFromConsiderationKeepVisibleSafe(payload);
}
