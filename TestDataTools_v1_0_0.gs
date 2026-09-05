var TestDataTools = (function() {
  var VERSION = '1.1.1';
  var MASTER_SHEET_NAME = 'Master Tracking';
  var ADMIN_SHEET_NAME = 'Admin_Input';
  var TEST_PREFIX = 'A-TEST-';
  var IMPORTED_STATUS = 'นำเข้าแล้ว';

  function seedSafeDashboardTestData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    var adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    if (!masterSheet || !adminSheet) {
      throw new Error('ไม่พบชีต Master Tracking หรือ Admin_Input');
    }

    var existingRows = readExistingRows_(masterSheet);
    var now = new Date();
    var definitions = buildTestDefinitions_(now);
    var created = [];
    var refreshed = [];

    definitions.forEach(function(definition) {
      removeAdminLogsForCode_(adminSheet, definition.projectCode);
      definition.timeline.forEach(function(event) {
        AdminInputImportTools.appendAdminInputLog({
          projectCode: definition.projectCode,
          documentNo: definition.documentNo,
          ownerUnit: definition.ownerUnit,
          receivedDate: definition.receivedDate,
          projectName: definition.projectName,
          budgetAmount: definition.budgetAmount,
          budgetSource: definition.budgetSource,
          fiscalYear: definition.fiscalYear,
          phase: event.phase,
          wiStep: event.wiStep,
          phaseEntryDate: event.phaseEntryDate,
          responsibleUnit: MasterTrackingTools.getResponsibleUnitByPhase(event.phase),
          note: event.note,
          recordedBy: 'TEST_DATA_TOOL',
          recordedAt: event.recordedAt,
          recordStatus: IMPORTED_STATUS,
        });
      });

      if (existingRows[definition.projectCode]) {
        writeMasterRow_(masterSheet, existingRows[definition.projectCode], definition);
        refreshed.push(definition.projectCode);
      } else {
        var newRow = masterSheet.getLastRow() + 1;
        writeMasterRow_(masterSheet, newRow, definition);
        existingRows[definition.projectCode] = newRow;
        created.push(definition.projectCode);
      }
    });

    return {
      success: true,
      version: VERSION,
      created: created,
      refreshed: refreshed,
      message: 'สร้างเฉพาะข้อมูลที่ขึ้นต้นด้วย ' + TEST_PREFIX + ' และไม่แก้ไขแถวเดิม',
    };
  }

  /**
   * Refreshes the previously generated A-0001..A-0020 sample rows and adds
   * A-0021..A-0030 without clearing non-sample operational data.
   */
  function refreshSampleTimesAndAddTen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    var adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    if (!masterSheet || !adminSheet) {
      throw new Error('ไม่พบชีต Master Tracking หรือ Admin_Input');
    }

    var now = new Date();
    var existingRows = readExistingRows_(masterSheet);
    var definitions = buildFreshSampleDefinitions_(masterSheet, existingRows, now);
    var created = [];
    var refreshed = [];

    definitions.forEach(function(definition) {
      removeAdminLogsForCode_(adminSheet, definition.projectCode);
      definition.timeline.forEach(function(event) {
        AdminInputImportTools.appendAdminInputLog({
          projectCode: definition.projectCode,
          documentNo: definition.documentNo,
          ownerUnit: definition.ownerUnit,
          receivedDate: definition.receivedDate,
          projectName: definition.projectName,
          budgetAmount: definition.budgetAmount,
          budgetSource: definition.budgetSource,
          fiscalYear: definition.fiscalYear,
          phase: event.phase,
          wiStep: event.wiStep,
          phaseEntryDate: event.phaseEntryDate,
          responsibleUnit: MasterTrackingTools.getResponsibleUnitByPhase(event.phase),
          note: event.note,
          recordedBy: 'TEST_DATA_TOOL_REFRESH',
          recordedAt: event.recordedAt,
          recordStatus: IMPORTED_STATUS,
        });
      });

      if (existingRows[definition.projectCode]) {
        writeMasterRow_(masterSheet, existingRows[definition.projectCode], definition);
        refreshed.push(definition.projectCode);
      } else {
        var newRow = masterSheet.getLastRow() + 1;
        writeMasterRow_(masterSheet, newRow, definition);
        existingRows[definition.projectCode] = newRow;
        created.push(definition.projectCode);
      }
    });

    return {
      success: true,
      version: VERSION,
      created: created,
      refreshed: refreshed,
      statusMix: {
        normal: 8,
        nearDue: 8,
        delayed: 8,
        pending: 3,
        rejected: 1,
        cancelled: 1,
        completed: 1,
      },
      message: 'รีเฟรชวันที่ชุดจำลองเดิมและเพิ่มรายการจำลองใหม่ 10 รายการแล้ว โดยไม่ล้างข้อมูลจริง',
    };
  }

  function addTenFreshSampleProjectsOnly() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    var adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    if (!masterSheet || !adminSheet) {
      throw new Error('ไม่พบชีต Master Tracking หรือ Admin_Input');
    }

    var now = new Date();
    var existingRows = readExistingRows_(masterSheet);
    var plans = [
      ['Phase 1', '1.สำรวจพื้นที่', 10],
      ['Phase 1', '2.เสนอคณะทำงานกลั่นกรองพิจารณา', 28],
      ['Phase 1', '3.แจ้งมติที่ประชุม', 12],
      ['Phase 2', '1.เขียนแบบ', 8],
      ['Phase 2', '2.คณะกรรมการพิจารณาแบบ', 28],
      ['Phase 3', '1.ส่งมอบเอกสารให้ผู้รับบริการ (User)', 8],
      ['Phase 4', '2.เสนอผู้บริหารพิจารณาแหล่งงบประมาณ', 3],
      ['Phase 5', 'รอรับเรื่อง', 0],
      ['Phase 6', '1.ตรวจรับงานจ้าง', 363],
      ['Phase 7', '1.บันทึกทะเบียนคุมสินทรัพย์', 36],
    ];
    var created = [];
    var skipped = [];

    plans.forEach(function(plan, offset) {
      var definition = buildNewSampleDefinition_(21 + offset, plan, now);
      if (existingRows[definition.projectCode]) {
        skipped.push(definition.projectCode);
        return;
      }

      definition.timeline.forEach(function(event) {
        AdminInputImportTools.appendAdminInputLog({
          projectCode: definition.projectCode,
          documentNo: definition.documentNo,
          ownerUnit: definition.ownerUnit,
          receivedDate: definition.receivedDate,
          projectName: definition.projectName,
          budgetAmount: definition.budgetAmount,
          budgetSource: definition.budgetSource,
          fiscalYear: definition.fiscalYear,
          phase: event.phase,
          wiStep: event.wiStep,
          phaseEntryDate: event.phaseEntryDate,
          responsibleUnit: MasterTrackingTools.getResponsibleUnitByPhase(event.phase),
          note: event.note,
          recordedBy: 'TEST_DATA_TOOL_ADD10',
          recordedAt: event.recordedAt,
          recordStatus: IMPORTED_STATUS,
        });
      });

      var newRow = masterSheet.getLastRow() + 1;
      writeMasterRow_(masterSheet, newRow, definition);
      existingRows[definition.projectCode] = newRow;
      created.push(definition.projectCode);
    });

    return {
      success: true,
      version: VERSION,
      created: created,
      skipped: skipped,
      message: 'เพิ่มรายการจำลองใหม่ 10 รายการ โดยไม่ล้างข้อมูลเดิม',
    };
  }

  /**
   * Rebuilds meaningful Admin_Input and Phase_History records for the ten
   * newest sample rows. This is deliberately limited to A-0021..A-0030 so
   * operational rows and earlier sample data are never changed.
   */
  function enrichNewSampleTimelines() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    var adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    if (!masterSheet || !adminSheet) {
      throw new Error('ไม่พบชีต Master Tracking หรือ Admin_Input');
    }

    PhaseHistoryTools.setupPhaseHistory();
    var historySheet = ss.getSheetByName('Phase_History');
    var now = new Date();
    var existingRows = readExistingRows_(masterSheet);
    var plans = [
      ['Phase 1', '1.สำรวจพื้นที่', 10],
      ['Phase 1', '2.เสนอคณะทำงานกลั่นกรองพิจารณา', 28],
      ['Phase 1', '3.แจ้งมติที่ประชุม', 12],
      ['Phase 2', '1.เขียนแบบ', 8],
      ['Phase 2', '2.คณะกรรมการพิจารณาแบบ', 28],
      ['Phase 3', '1.ส่งมอบเอกสารให้ผู้รับบริการ (User)', 8],
      ['Phase 4', '2.เสนอผู้บริหารพิจารณาแหล่งงบประมาณ', 3],
      ['Phase 5', 'รอรับเรื่อง', 0],
      ['Phase 6', '1.ตรวจรับงานจ้าง', 363],
      ['Phase 7', '1.บันทึกทะเบียนคุมสินทรัพย์', 36],
    ];
    var updated = [];
    var missing = [];

    plans.forEach(function(plan, offset) {
      var number = 21 + offset;
      var definition = buildNewSampleDefinition_(number, plan, now);
      if (!existingRows[definition.projectCode]) {
        missing.push(definition.projectCode);
        return;
      }

      removeAdminLogsForCode_(adminSheet, definition.projectCode);
      removeRowsByCode_(historySheet, definition.projectCode, 1);
      var timeline = buildMeaningfulSampleTimeline_(definition, now, number);

      timeline.events.forEach(function(event) {
        AdminInputImportTools.appendAdminInputLog({
          projectCode: definition.projectCode,
          documentNo: definition.documentNo,
          ownerUnit: definition.ownerUnit,
          receivedDate: definition.receivedDate,
          projectName: definition.projectName,
          budgetAmount: definition.budgetAmount,
          budgetSource: definition.budgetSource,
          fiscalYear: definition.fiscalYear,
          phase: event.phase,
          wiStep: event.wiStep,
          phaseEntryDate: event.recordedAt,
          responsibleUnit: MasterTrackingTools.getResponsibleUnitByPhase(event.phase),
          note: event.note,
          recordedBy: 'TEST_DATA_TIMELINE',
          recordedAt: event.recordedAt,
          recordStatus: IMPORTED_STATUS,
        });
      });

      timeline.closedPhases.forEach(function(item) {
        PhaseHistoryTools.recordClosedPhase({
          projectCode: definition.projectCode,
          phase: item.phase,
          phaseEntryDate: item.entryDate,
          phaseTotalSla: MasterTrackingTools.getPhaseTotalSla(item.phase),
          wiStep: item.finalWi,
        }, item.exitDate, item.finalWi, '[TEST DATA] ดำเนินการครบขั้นตอนและส่งต่อ Phase ถัดไป');
      });
      updated.push(definition.projectCode);
    });

    return {
      success: true,
      updated: updated,
      missing: missing,
      message: 'สร้าง Timeline จำลองตาม Phase, WI และ SLA สำหรับ A-0021 ถึง A-0030 แล้ว',
    };
  }

  /**
   * Adds coherent histories to every dashboard sample (A-0001, A-0002, ...).
   * Master Tracking fields are read only; only the matching test histories are
   * replaced in Admin_Input and Phase_History.
   */
  function enrichAllSampleTimelines() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    var adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    if (!masterSheet || !adminSheet) {
      throw new Error('ไม่พบชีต Master Tracking หรือ Admin_Input');
    }
    PhaseHistoryTools.setupPhaseHistory();
    var historySheet = ss.getSheetByName('Phase_History');
    var now = new Date();
    var rows = masterSheet.getLastRow() < 2
      ? []
      : masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 22).getValues();
    var updated = [];
    var ignored = [];

    rows.forEach(function(row, index) {
      var code = String(row[0] || '').trim();
      if (!/^A-\d{4}$/.test(code)) {
        if (code) ignored.push(code);
        return;
      }
      var definition = {
        projectCode: code,
        documentNo: row[1] || ('SAMPLE/' + code.slice(-4)),
        ownerUnit: row[2] || '[TEST] หน่วยงานจำลอง',
        receivedDate: asDate_(row[3]) || addDays_(now, -30),
        projectName: row[4] || '[TEST] รายการจำลอง',
        budgetAmount: row[5] || 0,
        budgetSource: row[6] || 'งบจำลอง',
        fiscalYear: row[7] || '2569',
        phase: String(row[8] || 'Phase 1'),
        wiStep: String(row[9] || ''),
        phaseEntryDate: asDate_(row[10]) || addDays_(now, -1),
      };
      replaceSampleTimeline_(adminSheet, historySheet, definition, now, index + 1);
      updated.push(code);
    });
    return {
      success: true,
      updated: updated,
      ignored: ignored,
      message: 'สร้าง Timeline จำลองตามสถานะและ SLA ให้ทุกรายการ A-xxxx แล้ว',
    };
  }

  function buildFreshSampleDefinitions_(masterSheet, existingRows, now) {
    var definitions = [];
    var existingValues = masterSheet.getLastRow() < 2
      ? {}
      : masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 22).getValues()
        .reduce(function(result, row) {
          if (row[0]) result[String(row[0])] = row;
          return result;
        }, {});

    for (var index = 1; index <= 20; index += 1) {
      var code = 'A-' + String(index).padStart(4, '0');
      if (!existingRows[code]) continue;
      if (index < 20 && String(existingValues[code][20] || '').indexOf('ปรับวันที่ใหม่เพื่อจำลองสถานะหลายรูปแบบ') !== -1) {
        continue;
      }
      definitions.push(buildExistingSampleDefinition_(code, existingValues[code], index, now));
    }

    var newPlans = [
      ['Phase 1', '1.สำรวจพื้นที่', 10, 'normal'],
      ['Phase 1', '2.เสนอคณะทำงานกลั่นกรองพิจารณา', 28, 'nearDue'],
      ['Phase 1', '3.แจ้งมติที่ประชุม', 12, 'delayed'],
      ['Phase 2', '1.เขียนแบบ', 8, 'normal'],
      ['Phase 2', '2.คณะกรรมการพิจารณาแบบ', 28, 'nearDue'],
      ['Phase 3', '1.ส่งมอบเอกสารให้ผู้รับบริการ (User)', 8, 'delayed'],
      ['Phase 4', '2.เสนอผู้บริหารพิจารณาแหล่งงบประมาณ', 3, 'normal'],
      ['Phase 5', 'รอรับเรื่อง', 0, 'pending'],
      ['Phase 6', '1.ตรวจรับงานจ้าง', 363, 'nearDue'],
      ['Phase 7', '1.บันทึกทะเบียนคุมสินทรัพย์', 36, 'delayed'],
    ];

    newPlans.forEach(function(plan, offset) {
      var number = 21 + offset;
      definitions.push(buildNewSampleDefinition_(number, plan, now));
    });
    return definitions;
  }

  function buildExistingSampleDefinition_(code, row, index, now) {
    var phase = String(row[8] || 'Phase 1');
    var wiStep = String(row[9] || '1.สำรวจพื้นที่');
    var statusPlan = getExistingStatusPlan_(index);
    var phaseSla = wiStep === 'รอรับเรื่อง'
      ? MasterTrackingTools.getPhaseTotalSla(phase)
      : MasterTrackingTools.getSlaDays(phase, wiStep);
    var phaseEntryDate = addDays_(now, -statusPlan.age);
    var forcedStatus = statusPlan.forcedStatus || '';

    if (statusPlan.type === 'pending') {
      wiStep = 'รอรับเรื่อง';
    } else if (statusPlan.type === 'rejected') {
      phase = 'Phase 1';
      wiStep = '3.แจ้งมติที่ประชุม';
    } else if (statusPlan.type === 'cancelled') {
      phase = 'Phase 5';
      wiStep = 'ยกเลิกรายการ';
    } else if (statusPlan.type === 'completed') {
      // Keep the terminal item in Phase 7. The phase column validates only
      // Phase 1-7, while the terminal state is represented by WI/status.
      phase = 'Phase 7';
      wiStep = 'เสร็จสิ้น';
    }

    if (statusPlan.type === 'pending') {
      phaseSla = MasterTrackingTools.getPhaseTotalSla(phase);
    } else if (statusPlan.type === 'rejected' || statusPlan.type === 'cancelled' || statusPlan.type === 'completed') {
      phaseSla = 30;
    } else {
      phaseSla = MasterTrackingTools.getSlaDays(phase, wiStep);
    }

    return {
      projectCode: code,
      documentNo: row[1] || ('SAMPLE/' + code.slice(-4)),
      ownerUnit: row[2] || '[TEST] หน่วยงานจำลอง',
      receivedDate: addDays_(now, -(statusPlan.age + 14)),
      projectName: row[4] || '[TEST] รายการจำลองปรับเวลาใหม่',
      budgetAmount: row[5] || 1000000,
      budgetSource: row[6] || 'งบจำลอง',
      fiscalYear: row[7] || '2569',
      phase: phase,
      wiStep: wiStep,
      phaseEntryDate: phaseEntryDate,
      slaDays: phaseSla,
      forcedStatus: forcedStatus,
      note: '[TEST DATA] ปรับวันที่ใหม่เพื่อจำลองสถานะหลายรูปแบบ',
      timeline: makeTimeline_(now, [[
        phase,
        wiStep,
        -statusPlan.age,
        9 + (index % 8),
        (index * 7) % 60,
        index,
        '[TEST DATA] เหตุการณ์ล่าสุดหลังรีเฟรชเวลา',
      ]]),
    };
  }

  function getExistingStatusPlan_(index) {
    if (index <= 5) return { type: 'normal', age: 4 + index };
    if (index <= 10) return { type: 'nearDue', age: 25 + (index % 3) };
    if (index <= 15) return { type: 'delayed', age: 38 + (index % 5) };
    if (index <= 17) return { type: 'pending', age: 0 };
    if (index === 18) return { type: 'rejected', age: 8, forcedStatus: '❌ ไม่อนุมัติ' };
    if (index === 19) return { type: 'cancelled', age: 8, forcedStatus: '⚫ ยกเลิกรายการ' };
    if (index === 20) return { type: 'completed', age: 1, forcedStatus: '🔵 เสร็จสิ้น' };
    return { type: 'normal', age: 7 };
  }

  function buildNewSampleDefinition_(number, plan, now) {
    var phase = plan[0];
    var wiStep = plan[1];
    var age = plan[2];
    var type = plan[3];
    var isPending = type === 'pending';
    var slaDays = isPending
      ? MasterTrackingTools.getPhaseTotalSla(phase)
      : MasterTrackingTools.getSlaDays(phase, wiStep);
    var code = 'A-' + String(number).padStart(4, '0');
    return {
      projectCode: code,
      documentNo: 'SAMPLE/' + String(number).padStart(3, '0'),
      ownerUnit: '[TEST] หน่วยงานจำลองชุดใหม่',
      receivedDate: addDays_(now, -(Math.max(age, 1) + 14)),
      projectName: '[TEST] รายการจำลองสถานะ ' + String(number).padStart(2, '0'),
      budgetAmount: 750000 + (number * 125000),
      budgetSource: 'งบจำลอง',
      fiscalYear: '2569',
      phase: phase,
      wiStep: wiStep,
      phaseEntryDate: addDays_(now, -age),
      slaDays: slaDays,
      forcedStatus: '',
      note: '[TEST DATA] รายการใหม่สำหรับทดสอบสถานะและเวลา',
      timeline: makeTimeline_(now, [[
        phase,
        wiStep,
        -age,
        10 + (number % 6),
        (number * 5) % 60,
        number,
        '[TEST DATA] สร้างรายการใหม่เพื่อทดสอบ Dashboard',
      ]]),
    };
  }

  function buildMeaningfulSampleTimeline_(definition, now, seed) {
    var phases = MasterTrackingTools.getPhaseList();
    var currentIndex = phases.indexOf(definition.phase);
    var currentEntry = definition.phaseEntryDate;
    var events = [];
    var closedPhases = [];
    var cursor = Math.max(1, Math.round((now.getTime() - currentEntry.getTime()) / 86400000) + 1);

    for (var index = currentIndex - 1; index >= 0; index -= 1) {
      var phase = phases[index];
      var wiList = MasterTrackingTools.getWiListByPhase(phase) || [];
      var firstWi = wiList[0] || '';
      var finalWi = wiList[wiList.length - 1] || firstWi;
      var phaseSla = MasterTrackingTools.getPhaseTotalSla(phase);
      var duration = Math.max(2, Math.min(10 + ((seed + index) % 6), phaseSla - 1));
      var exitDate = dateAt_(addDays_(now, -cursor), 15, 10 + ((seed + index) % 40), seed + index);
      var entryDate = dateAt_(addDays_(exitDate, -duration + 1), 9, 5 + ((seed + index) % 30), seed + index);
      events.push({ phase: phase, wiStep: firstWi, recordedAt: entryDate, note: '[TEST DATA] รับเรื่องและเริ่มดำเนินการ' });
      if (finalWi && finalWi !== firstWi) {
        events.push({ phase: phase, wiStep: finalWi, recordedAt: exitDate, note: '[TEST DATA] ดำเนินการครบขั้นตอนและส่งต่อ Phase ถัดไป' });
      }
      closedPhases.push({ phase: phase, entryDate: entryDate, exitDate: exitDate, finalWi: finalWi });
      cursor += duration + 1;
    }

    var currentWiList = MasterTrackingTools.getWiListByPhase(definition.phase) || [];
    var firstCurrentWi = currentWiList[0] || definition.wiStep;
    var currentStart = dateAt_(currentEntry, 9 + (seed % 4), 5 + ((seed * 3) % 40), seed);
    events.push({
      phase: definition.phase,
      wiStep: firstCurrentWi,
      recordedAt: currentStart,
      note: definition.wiStep === 'รอรับเรื่อง'
        ? '[TEST DATA] ส่งเรื่องระหว่างหน่วยงาน รอผู้รับเรื่องกดรับ'
        : '[TEST DATA] เริ่มดำเนินการใน Phase ปัจจุบัน',
    });
    if (definition.wiStep && definition.wiStep !== firstCurrentWi) {
      var currentAge = Math.max(1, Math.floor((now.getTime() - currentEntry.getTime()) / 86400000));
      var currentWiAt = dateAt_(addDays_(currentEntry, Math.max(1, Math.floor(currentAge / 2))), 13 + (seed % 3), 10, seed + 1);
      events.push({
        phase: definition.phase,
        wiStep: definition.wiStep,
        recordedAt: currentWiAt,
        note: '[TEST DATA] อยู่ระหว่างดำเนินการตาม WI ปัจจุบัน',
      });
    }
    return { events: events.sort(function(a, b) { return a.recordedAt.getTime() - b.recordedAt.getTime(); }), closedPhases: closedPhases };
  }

  function replaceSampleTimeline_(adminSheet, historySheet, definition, now, seed) {
    removeAdminLogsForCode_(adminSheet, definition.projectCode);
    removeRowsByCode_(historySheet, definition.projectCode, 1);
    var phaseList = MasterTrackingTools.getPhaseList();
    var isCompleted = phaseList.indexOf(definition.phase) === -1;
    var timelineDefinition = definition;
    if (isCompleted) {
      var finalPhase = phaseList[phaseList.length - 1];
      var finalWiList = MasterTrackingTools.getWiListByPhase(finalPhase) || [];
      timelineDefinition = {
        projectCode: definition.projectCode,
        phase: finalPhase,
        wiStep: finalWiList[finalWiList.length - 1] || '',
        phaseEntryDate: definition.phaseEntryDate,
      };
    }
    var timeline = buildMeaningfulSampleTimeline_(timelineDefinition, now, seed);
    if (isCompleted) {
      timeline.events.push({
        phase: phaseList[phaseList.length - 1],
        wiStep: definition.wiStep || 'เสร็จสิ้น',
        recordedAt: dateAt_(definition.phaseEntryDate, 16, 10, seed),
        note: '[TEST DATA] ดำเนินการเสร็จสิ้น',
      });
    }
    timeline.events.sort(function(a, b) { return a.recordedAt.getTime() - b.recordedAt.getTime(); })
      .forEach(function(event) {
        AdminInputImportTools.appendAdminInputLog({
          projectCode: definition.projectCode,
          documentNo: definition.documentNo,
          ownerUnit: definition.ownerUnit,
          receivedDate: definition.receivedDate,
          projectName: definition.projectName,
          budgetAmount: definition.budgetAmount,
          budgetSource: definition.budgetSource,
          fiscalYear: definition.fiscalYear,
          phase: event.phase,
          wiStep: event.wiStep,
          phaseEntryDate: event.recordedAt,
          responsibleUnit: MasterTrackingTools.getResponsibleUnitByPhase(event.phase),
          note: event.note,
          recordedBy: 'TEST_DATA_TIMELINE_ALL',
          recordedAt: event.recordedAt,
          recordStatus: IMPORTED_STATUS,
        });
      });
    timeline.closedPhases.forEach(function(item) {
      PhaseHistoryTools.recordClosedPhase({
        projectCode: definition.projectCode,
        phase: item.phase,
        phaseEntryDate: item.entryDate,
        phaseTotalSla: MasterTrackingTools.getPhaseTotalSla(item.phase),
        wiStep: item.finalWi,
      }, item.exitDate, item.finalWi, '[TEST DATA] ดำเนินการครบขั้นตอนและส่งต่อ Phase ถัดไป');
    });
  }

  function buildTestDefinitions_(now) {
    var received = addDays_(now, -45);
    return [
      {
        projectCode: 'A-TEST-001',
        documentNo: 'TEST/001',
        ownerUnit: '[TEST] หน่วยงานทดสอบระบบ',
        receivedDate: received,
        projectName: '[TEST] ตรวจลำดับ Timeline หลาย Phase',
        budgetAmount: 1250000,
        budgetSource: 'งบทดสอบ',
        fiscalYear: '2569',
        phase: 'Phase 3',
        wiStep: 'ส่งมอบ BOQ ให้หน่วยงานเพื่อทำโครงการ',
        phaseEntryDate: addDays_(now, -2),
        note: '[TEST DATA] ใช้ตรวจลำดับ Timeline และเครื่องหมาย Phase',
        timeline: makeTimeline_(now, [
          ['Phase 1', '1.รอสำรวจ', -12, 9, 0, 1],
          ['Phase 1', '2.รอกรรมการกลั่นกรองพิจารณา', -11, 10, 15, 2],
          ['Phase 1', '3.พิจารณารายการ', -10, 14, 1, 3],
          ['Phase 2', '1.กำลังออกแบบ', -9, 8, 30, 4],
          ['Phase 2', '4.รอกรรมการ Factor F อนุมัติ', -5, 13, 20, 5],
          ['Phase 2', '5.เสนอคณบดีเห็นชอบราคากลาง', -4, 15, 5, 6],
          ['Phase 3', 'ส่งมอบ BOQ ให้หน่วยงานเพื่อทำโครงการ', -2, 11, 45, 7],
        ]),
      },
      {
        projectCode: 'A-TEST-002',
        documentNo: 'TEST/002',
        ownerUnit: '[TEST] หน่วยงานทดสอบระบบ',
        receivedDate: addDays_(now, -20),
        projectName: '[TEST] รายการไม่อนุมัติ',
        budgetAmount: 500000,
        budgetSource: 'งบทดสอบ',
        fiscalYear: '2569',
        phase: 'Phase 1',
        wiStep: '3.พิจารณารายการ',
        phaseEntryDate: addDays_(now, -5),
        note: '[TEST DATA] ไม่อนุมัติรายการ',
        timeline: makeTimeline_(now, [
          ['Phase 1', '1.รอสำรวจ', -8, 9, 0, 1],
          ['Phase 1', '2.รอกรรมการกลั่นกรองพิจารณา', -7, 10, 0, 2],
          ['Phase 1', '3.พิจารณารายการ', -5, 14, 30, 3, '[TEST DATA] ไม่อนุมัติรายการ'],
        ]),
      },
      {
        projectCode: 'A-TEST-003',
        documentNo: 'TEST/003',
        ownerUnit: '[TEST] หน่วยงานทดสอบระบบ',
        receivedDate: addDays_(now, -75),
        projectName: '[TEST] รายการยกเลิก',
        budgetAmount: 780000,
        budgetSource: 'งบทดสอบ',
        fiscalYear: '2569',
        phase: 'Phase 5',
        wiStep: 'ยกเลิกรายการ',
        phaseEntryDate: addDays_(now, -3),
        note: '[TEST DATA] ยกเลิกรายการ',
        timeline: makeTimeline_(now, [
          ['Phase 4', '3.แจ้งผลการจัดสรร(ใบขวาง)', -10, 9, 5, 1],
          ['Phase 5', '1.แต่งตั้งคณะกรรมการ', -6, 10, 10, 2],
          ['Phase 5', 'ยกเลิกรายการ', -3, 16, 25, 3, '[TEST DATA] ยกเลิกรายการ'],
        ]),
      },
      {
        projectCode: 'A-TEST-004',
        documentNo: 'TEST/004',
        ownerUnit: '[TEST] หน่วยงานทดสอบระบบ',
        receivedDate: addDays_(now, -100),
        projectName: '[TEST] โครงการเสร็จสิ้น',
        budgetAmount: 2100000,
        budgetSource: 'งบทดสอบ',
        fiscalYear: '2569',
        phase: 'Phase 7',
        wiStep: 'เสร็จสิ้น',
        phaseEntryDate: addDays_(now, -1),
        note: '[TEST DATA] เสร็จสิ้น',
        timeline: makeTimeline_(now, [
          ['Phase 6', '2.เบิกจ่าย', -8, 10, 20, 1],
          ['Phase 7', 'บันทึกทะเบียนคุมสินทรัพย์', -4, 13, 15, 2],
          ['Phase 7', 'เสร็จสิ้น', -1, 16, 40, 3, '[TEST DATA] เสร็จสิ้น'],
        ]),
      },
      {
        projectCode: 'A-TEST-005',
        documentNo: 'TEST/005',
        ownerUnit: '[TEST] หน่วยงานทดสอบระบบ',
        receivedDate: addDays_(now, -90),
        projectName: '[TEST] Phase ล่าช้า',
        budgetAmount: 920000,
        budgetSource: 'งบทดสอบ',
        fiscalYear: '2569',
        phase: 'Phase 4',
        wiStep: '2.เสนอผู้บริหารพิจารณาแหล่งงบประมาณ',
        phaseEntryDate: addDays_(now, -45),
        note: '[TEST DATA] ใช้ทดสอบสถานะล่าช้า',
        timeline: makeTimeline_(now, [
          ['Phase 3', 'ส่งมอบ BOQ ให้หน่วยงานเพื่อทำโครงการ', -55, 9, 30, 1],
          ['Phase 4', '1.ตรวจสอบเอกสาร', -45, 10, 20, 2],
          ['Phase 4', '2.เสนอผู้บริหารพิจารณาแหล่งงบประมาณ', -40, 14, 50, 3],
        ]),
      },
    ];
  }

  function makeTimeline_(now, rows) {
    return rows.map(function(item) {
      var recordedAt = dateAt_(addDays_(now, item[2]), item[3], item[4], item[5]);
      return {
        phase: item[0],
        wiStep: item[1],
        phaseEntryDate: recordedAt,
        recordedAt: recordedAt,
        note: item[6] || '[TEST DATA] ขั้นตอนจำลอง',
      };
    });
  }

  function writeMasterRow_(sheet, row, definition) {
    var completedSystemDays = definition.forcedStatus === '🔵 เสร็จสิ้น'
      ? getInclusiveDays_(definition.receivedDate, definition.phaseEntryDate || new Date())
      : '';
    var values = [
      definition.projectCode,
      definition.documentNo,
      definition.ownerUnit,
      definition.receivedDate,
      definition.projectName,
      definition.budgetAmount,
      definition.budgetSource,
      definition.fiscalYear,
      definition.phase,
      definition.wiStep,
      definition.phaseEntryDate,
      definition.slaDays || 30,
      '',
      '',
      definition.phase === 'เสร็จสิ้น' ? '-' : MasterTrackingTools.getResponsibleUnitByPhase(definition.phase),
      '',
      180,
      completedSystemDays,
      '',
      '',
      definition.note,
      new Date(),
    ];

    sheet.getRange(row, 1, 1, values.length).setValues([values]);
    if (definition.phase !== 'เสร็จสิ้น') {
      MasterTrackingTools.setWiValidationForRow(row, definition.phase);
    }
    if (definition.wiStep === 'รอรับเรื่อง') {
      sheet.getRange(row, 13).clearContent();
      sheet.getRange(row, 14).setValue('⏳ รอรับเรื่อง');
      sheet.getRange(row, 20).clearContent();
    } else if (definition.forcedStatus) {
      sheet.getRange(row, 14).setValue(definition.forcedStatus);
      sheet.getRange(row, 20).setValue(definition.forcedStatus);
    } else {
      MasterTrackingTools.setMasterFormulasForRow(row);
    }
  }

  function readExistingRows_(sheet) {
    var result = {};
    if (sheet.getLastRow() < 2) return result;
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(function(value, index) {
      if (value[0]) result[String(value[0])] = index + 2;
    });
    return result;
  }

  function removeAdminLogsForCode_(sheet, projectCode) {
    if (sheet.getLastRow() < 2) return;
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var index = values.length - 1; index >= 0; index -= 1) {
      if (String(values[index][0] || '') === String(projectCode)) {
        sheet.deleteRow(index + 2);
      }
    }
  }

  function removeRowsByCode_(sheet, projectCode, codeColumn) {
    if (!sheet || sheet.getLastRow() < 2) return;
    var values = sheet.getRange(2, codeColumn, sheet.getLastRow() - 1, 1).getValues();
    for (var index = values.length - 1; index >= 0; index -= 1) {
      if (String(values[index][0] || '') === String(projectCode)) sheet.deleteRow(index + 2);
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

  function asDate_(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (!value) return null;
    var parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function getInclusiveDays_(startValue, endValue) {
    var start = asDate_(startValue);
    var end = asDate_(endValue);
    if (!start || !end) return 0;
    start = new Date(start.getFullYear(), start.getMonth(), start.getDate());
    end = new Date(end.getFullYear(), end.getMonth(), end.getDate());
    return Math.max(1, Math.floor((end.getTime() - start.getTime()) / 86400000) + 1);
  }

  return {
    VERSION: VERSION,
    seedSafeDashboardTestData: seedSafeDashboardTestData,
    refreshSampleTimesAndAddTen: refreshSampleTimesAndAddTen,
    addTenFreshSampleProjectsOnly: addTenFreshSampleProjectsOnly,
    enrichNewSampleTimelines: enrichNewSampleTimelines,
    enrichAllSampleTimelines: enrichAllSampleTimelines,
  };
})();

function seedSafeDashboardTestData() {
  return TestDataTools.seedSafeDashboardTestData();
}

function refreshSampleTimesAndAddTen() {
  return TestDataTools.refreshSampleTimesAndAddTen();
}

function addTenFreshSampleProjectsOnly() {
  return TestDataTools.addTenFreshSampleProjectsOnly();
}

function enrichNewSampleTimelines() {
  return TestDataTools.enrichNewSampleTimelines();
}

function enrichAllSampleTimelines() {
  return TestDataTools.enrichAllSampleTimelines();
}
