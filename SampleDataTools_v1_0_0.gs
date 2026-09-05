var SampleDataTools = (function() {
  var VERSION = '1.0.0';
  var MASTER_SHEET_NAME = 'Master Tracking';
  var ADMIN_SHEET_NAME = 'Admin_Input';
  var PHASE_HISTORY_SHEET_NAME = 'Phase_History';
  var IMPORTED_STATUS = 'นำเข้าแล้ว';
  var PENDING_HANDOFF_WI = 'รอรับเรื่อง';
  var PENDING_HANDOFF_STATUS = '⏳ รอรับเรื่อง';

  function seedTwentyDiverseSampleProjects() {
    setupSheets_();
    clearOperationalSheets_();

    var now = new Date();
    var definitions = buildDefinitions_(now);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

    definitions.forEach(function(definition, index) {
      writeMasterRow_(masterSheet, index + 2, definition, now);
      writeTimeline_(definition, now);
      writePhaseHistory_(definition, now);
    });

    return {
      success: true,
      version: VERSION,
      created: definitions.length,
      projectCodes: definitions.map(function(item) { return item.projectCode; }),
      groups: summarizeGroups_(definitions),
    };
  }

  function setupSheets_() {
    MasterTrackingTools.setupMasterTracking();
    AdminInputImportTools.setupAdminInput();
    AuthApi.setupAdminAuthUsers();
    if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PHASE_HISTORY_SHEET_NAME)) {
      PhaseHistoryTools.setupAndBackfillPhaseHistory();
    }
  }

  function clearOperationalSheets_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    [MASTER_SHEET_NAME, ADMIN_SHEET_NAME, PHASE_HISTORY_SHEET_NAME].forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      if (sheet && sheet.getMaxRows() > 1) {
        sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clearContent();
      }
    });
  }

  function buildDefinitions_(now) {
    return [
      def_(1, 'ปรับปรุงห้องตรวจ OPD ชั้น 2', 'Phase 1', '1.สำรวจพื้นที่', -2, 450000, '', '', 'เริ่มรับเรื่องใหม่', []),
      def_(2, 'ซ่อมแซมทางเดินเชื่อมอาคาร A-B', 'Phase 1', '2.เสนอคณะทำงานกลั่นกรองพิจารณา', -8, 880000, '', '', 'รอผลคณะทำงานกลั่นกรอง', [
        hist_('Phase 1', -20, -8, 65, 13, 'ปกติ', '1.สำรวจพื้นที่', 'สำรวจพื้นที่แล้ว')
      ]),
      def_(3, 'ปรับปรุงห้องพักแพทย์เวร', 'Phase 1', '2.เสนอคณะทำงานกลั่นกรองพิจารณา', -12, 620000, '', '', 'ไม่อนุมัติรายการ', [], '❌ ไม่อนุมัติ'),
      def_(4, 'ปรับปรุงห้องประชุมเล็ก', 'Phase 2', '1.เขียนแบบ', -3, 1250000, '', '', 'อนุมัติและเริ่มออกแบบ', [
        hist_('Phase 1', -40, -4, 65, 37, 'ปกติ', '2.เสนอคณะทำงานกลั่นกรองพิจารณา', 'อนุมัติรายการและส่งต่อ Phase 2')
      ]),
      def_(5, 'ก่อสร้างกันสาดหน้าอาคารบริการ', 'Phase 2', '4.ถอดแบบ (BOQ) คำนวณราคากลาง (Factor F)', -42, 2100000, '', '', 'กำลังถอดแบบและคำนวณราคา', [
        hist_('Phase 1', -70, -50, 65, 21, 'ปกติ', '2.เสนอคณะทำงานกลั่นกรองพิจารณา', 'ผ่านการกลั่นกรอง'),
        hist_('Phase 2', -50, -42, 130, 9, 'ปกติ', '3.ผู้รับบริการรับรองแบบ (User)', 'เจ้าของเรื่องรับรองแบบ')
      ]),
      def_(6, 'ปรับปรุงระบบระบายน้ำลานจอดรถ', 'Phase 2', '5.คณะกรรมการพิจารณาราคาประมาณการ', -145, 3650000, '', '', 'ล่าช้า รอคณะกรรมการพิจารณาราคา', [
        hist_('Phase 1', -190, -150, 65, 41, 'ปกติ', '2.เสนอคณะทำงานกลั่นกรองพิจารณา', 'อนุมัติส่งต่อออกแบบ')
      ]),
      def_(7, 'ซ่อมแซมฝ้าเพดานอาคารเรียนรวม', 'Phase 3', '1.ส่งมอบเอกสารให้ผู้รับบริการ (User)', -2, 790000, '', '', 'เตรียมส่งมอบเอกสาร BOQ', [
        hist_('Phase 1', -80, -55, 65, 26, 'ปกติ', '2.เสนอคณะทำงานกลั่นกรองพิจารณา', 'ผ่านกลั่นกรอง'),
        hist_('Phase 2', -55, -3, 130, 53, 'ปกติ', '6.เสนอคณบดีเห็นชอบราคาประมาณการ', 'เห็นชอบราคาประมาณการ')
      ]),
      def_(8, 'ปรับปรุงห้องน้ำผู้ป่วยนอก', 'Phase 4', PENDING_HANDOFF_WI, '', 1475000, '', '', 'ส่งต่อให้งานนโยบายและแผน รอรับเรื่อง', [
        hist_('Phase 1', -75, -60, 65, 16, 'ปกติ', '2.เสนอคณะทำงานกลั่นกรองพิจารณา', 'ผ่านกลั่นกรอง'),
        hist_('Phase 2', -60, -10, 130, 51, 'ปกติ', '6.เสนอคณบดีเห็นชอบราคาประมาณการ', 'เห็นชอบราคา'),
        hist_('Phase 3', -10, -1, 5, 10, 'ล่าช้า', '1.ส่งมอบเอกสารให้ผู้รับบริการ (User)', 'ส่งต่อให้นโยบายและแผน')
      ], PENDING_HANDOFF_STATUS),
      def_(9, 'ปรับปรุงระบบไฟฟ้าอาคารวิจัย', 'Phase 4', '1.ตรวจสอบเอกสาร', -1, 3200000, '', '', 'รับเรื่องแล้ว ตรวจเอกสารงบประมาณ', [
        hist_('Phase 3', -8, -2, 5, 7, 'ล่าช้า', '1.ส่งมอบเอกสารให้ผู้รับบริการ (User)', 'งานนโยบายรับเรื่อง')
      ]),
      def_(10, 'ก่อสร้างห้องเก็บเวชภัณฑ์', 'Phase 4', '2.เสนอผู้บริหารพิจารณาแหล่งงบประมาณ', -6, 1850000, '', '', 'เสนอผู้บริหารพิจารณาแหล่งเงิน', [
        hist_('Phase 3', -20, -7, 5, 14, 'ล่าช้า', '1.ส่งมอบเอกสารให้ผู้รับบริการ (User)', 'ส่งแผนงบประมาณ')
      ]),
      def_(11, 'ปรับปรุงห้อง Lab กลาง', 'Phase 4', '3.แจ้งผลการอนุมัติ(ใบขวาง) (User)', -3, 5420000, 'งบลงทุน', '2570', 'รอแจ้งผลใบขวางให้ผู้รับบริการ', [
        hist_('Phase 3', -25, -12, 5, 14, 'ล่าช้า', '1.ส่งมอบเอกสารให้ผู้รับบริการ (User)', 'รับเรื่องงบประมาณ'),
        hist_('Phase 4', -12, -3, 15, 10, 'ปกติ', '2.เสนอผู้บริหารพิจารณาแหล่งงบประมาณ', 'ผู้บริหารเห็นชอบแหล่งเงิน')
      ]),
      def_(12, 'ปรับปรุงหอพักบุคลากร', 'Phase 5', PENDING_HANDOFF_WI, '', 7950000, 'งบลงทุน', '2570', 'ส่งต่อให้งานพัสดุฯ รอรับเรื่อง', [
        hist_('Phase 4', -18, -1, 15, 18, 'ล่าช้า', '3.แจ้งผลการอนุมัติ(ใบขวาง) (User)', 'ส่งต่อพัสดุหลังได้รับใบขวาง')
      ], PENDING_HANDOFF_STATUS),
      def_(13, 'ซ่อมแซมลิฟต์อาคารบริการ', 'Phase 5', '1.แต่งตั้งคณะกรรมการ', -2, 2680000, 'เงินบำรุง', '2570', 'งานพัสดุรับเรื่องแล้ว', [
        hist_('Phase 4', -20, -3, 15, 18, 'ล่าช้า', '3.แจ้งผลการอนุมัติ(ใบขวาง) (User)', 'ส่งต่อพัสดุ')
      ]),
      def_(14, 'ปรับปรุงระบบปรับอากาศ ICU', 'Phase 5', '4.ประกวดราคา', -75, 9300000, 'งบลงทุน', '2570', 'ล่าช้าในขั้นประกวดราคา', [
        hist_('Phase 4', -120, -90, 15, 31, 'ล่าช้า', '3.แจ้งผลการอนุมัติ(ใบขวาง) (User)', 'ได้รับอนุมัติงบประมาณ'),
        hist_('Phase 5', -90, -75, 120, 16, 'ปกติ', '3.จัดทำ TOR', 'TOR แล้วเสร็จ')
      ]),
      def_(15, 'ปรับปรุงหลังคาอาคารคลังพัสดุ', 'Phase 5', 'ยกเลิกรายการ', -5, 1150000, 'เงินรายได้', '2569', 'ยกเลิกรายการ เนื่องจากรวมกับโครงการอื่น', [
        hist_('Phase 4', -40, -20, 15, 21, 'ล่าช้า', '3.แจ้งผลการอนุมัติ(ใบขวาง) (User)', 'อนุมัติงบประมาณ'),
        hist_('Phase 5', -20, -5, 120, 16, 'ปกติ', '2.ราคากลาง', 'ยกเลิกรายการ')
      ], '⚫ ยกเลิกรายการ'),
      def_(16, 'ก่อสร้างทางลาดผู้พิการ', 'Phase 6', '1.ตรวจรับงานจ้าง', -34, 2450000, 'งบลงทุน', '2569', 'อยู่ระหว่างตรวจรับงานจ้าง', [
        hist_('Phase 5', -120, -35, 120, 86, 'ปกติ', '5.ลงนามในสัญญา', 'ลงนามในสัญญาแล้ว')
      ]),
      def_(17, 'ปรับปรุงอาคารจอดรถ', 'Phase 6', '2.เบิกจ่ายเงินงวด', -380, 15400000, 'งบลงทุน', '2569', 'ล่าช้าในขั้นเบิกจ่าย', [
        hist_('Phase 5', -500, -390, 120, 111, 'ปกติ', '5.ลงนามในสัญญา', 'ลงนามสัญญา'),
        hist_('Phase 6', -390, -380, 365, 11, 'ปกติ', '1.ตรวจรับงานจ้าง', 'ตรวจรับแล้ว')
      ]),
      def_(18, 'ซื้อครุภัณฑ์ระบบกล้องวงจรปิด', 'Phase 7', '1.บันทึกทะเบียนคุมสินทรัพย์', -7, 1780000, 'งบครุภัณฑ์', '2569', 'รอบันทึกทะเบียนสินทรัพย์', [
        hist_('Phase 6', -80, -8, 365, 73, 'ปกติ', '2.เบิกจ่ายเงินงวด', 'เบิกจ่ายแล้ว')
      ]),
      def_(19, 'ปรับปรุงห้องฉุกเฉิน', 'Phase 7', 'เสร็จสิ้น', -1, 22600000, 'งบลงทุน', '2569', 'เสร็จสิ้นและปิดรายการ', [
        hist_('Phase 5', -250, -120, 120, 131, 'ล่าช้า', '5.ลงนามในสัญญา', 'ลงนามล่าช้า'),
        hist_('Phase 6', -120, -20, 365, 101, 'ปกติ', '2.เบิกจ่ายเงินงวด', 'เบิกจ่ายครบ'),
        hist_('Phase 7', -20, -1, 30, 20, 'ปกติ', '1.บันทึกทะเบียนคุมสินทรัพย์', 'ปิดรายการ')
      ], '🔵 เสร็จสิ้น'),
      def_(20, 'ปรับปรุงพื้นที่รอญาติผู้ป่วย', 'Phase 2', '2.คณะกรรมการพิจารณาแบบ', -12, 980000, '', '', 'คณะกรรมการพิจารณาแบบรอบแรก', [
        hist_('Phase 1', -35, -15, 65, 21, 'ปกติ', '2.เสนอคณะทำงานกลั่นกรองพิจารณา', 'อนุมัติส่งต่อออกแบบ')
      ])
    ];
  }

  function def_(number, name, phase, wiStep, phaseEntryOffset, budget, source, fiscalYear, note, history, forcedStatus) {
    var code = 'A-' + String(number).padStart(4, '0');
    var now = new Date();
    var receivedOffset = -Math.max(10, Math.abs(Number(phaseEntryOffset) || 1) + 30);
    return {
      projectCode: code,
      documentNo: 'SAMPLE/' + String(number).padStart(3, '0'),
      ownerUnit: sampleOwner_(number),
      receivedDate: addDays_(now, receivedOffset),
      projectName: name,
      budgetAmount: budget,
      budgetSource: source || '',
      fiscalYear: fiscalYear || '',
      phase: phase,
      wiStep: wiStep,
      phaseEntryDate: phaseEntryOffset === '' ? '' : addDays_(now, phaseEntryOffset),
      note: note,
      history: history || [],
      forcedStatus: forcedStatus || '',
    };
  }

  function hist_(phase, startOffset, exitOffset, sla, usedDays, status, finalWi, note) {
    var now = new Date();
    return {
      phase: phase,
      phaseEntryDate: addDays_(now, startOffset),
      phaseExitDate: addDays_(now, exitOffset),
      phaseSla: sla,
      usedDays: usedDays,
      overdueDays: status === 'ล่าช้า' ? -Math.max(1, usedDays - sla) : 0,
      status: status,
      finalWi: finalWi,
      note: note || '',
    };
  }

  function writeMasterRow_(sheet, row, definition, now) {
    var phaseSla = definition.wiStep === PENDING_HANDOFF_WI
      ? MasterTrackingTools.getPhaseTotalSla(definition.phase)
      : MasterTrackingTools.getSlaDays(definition.phase, definition.wiStep);
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
      phaseSla,
      '',
      definition.forcedStatus || '',
      MasterTrackingTools.getResponsibleUnitByPhase(definition.phase),
      '',
      MasterTrackingTools.DEFAULT_MAIN_SLA_DAYS,
      '',
      '',
      definition.forcedStatus || '',
      definition.note,
      now,
    ];
    sheet.getRange(row, 1, 1, values.length).setValues([values]);
    MasterTrackingTools.setWiValidationForRow(row, definition.phase);
    if (definition.wiStep === PENDING_HANDOFF_WI) {
      sheet.getRange(row, 13).clearContent();
      sheet.getRange(row, 14).setValue(PENDING_HANDOFF_STATUS);
      sheet.getRange(row, 20).setValue('');
    } else if (definition.forcedStatus) {
      sheet.getRange(row, 14).setValue(definition.forcedStatus);
      sheet.getRange(row, 20).setValue(definition.forcedStatus);
    } else {
      MasterTrackingTools.setMasterFormulasForRow(row);
    }
  }

  function writeTimeline_(definition, now) {
    var steps = buildTimelineSteps_(definition);
    steps.forEach(function(step) {
      AdminInputImportTools.appendAdminInputLog({
        projectCode: definition.projectCode,
        documentNo: definition.documentNo,
        ownerUnit: definition.ownerUnit,
        receivedDate: definition.receivedDate,
        projectName: definition.projectName,
        budgetAmount: definition.budgetAmount,
        budgetSource: definition.budgetSource,
        fiscalYear: definition.fiscalYear,
        phase: step.phase,
        wiStep: step.wiStep,
        phaseEntryDate: step.phaseEntryDate,
        responsibleUnit: MasterTrackingTools.getResponsibleUnitByPhase(step.phase),
        note: step.note,
        recordedBy: 'SAMPLE_DATA',
        recordedAt: step.recordedAt,
        recordStatus: IMPORTED_STATUS,
      });
    });
  }

  function buildTimelineSteps_(definition) {
    var result = [];
    definition.history.forEach(function(item) {
      result.push({
        phase: item.phase,
        wiStep: item.finalWi,
        phaseEntryDate: item.phaseEntryDate,
        recordedAt: item.phaseExitDate,
        note: item.note || 'ส่งต่อขั้นตอนถัดไป',
      });
    });
    result.push({
      phase: definition.phase,
      wiStep: definition.wiStep,
      phaseEntryDate: definition.phaseEntryDate || new Date(),
      recordedAt: new Date(),
      note: definition.note,
    });
    return result;
  }

  function writePhaseHistory_(definition) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PHASE_HISTORY_SHEET_NAME);
    if (!sheet) return;
    definition.history.forEach(function(item) {
      sheet.appendRow([
        definition.projectCode,
        item.phase,
        item.phaseEntryDate,
        item.phaseExitDate,
        item.phaseSla,
        item.usedDays,
        item.overdueDays,
        item.status,
        item.finalWi,
        item.note,
        new Date(),
      ]);
    });
  }

  function summarizeGroups_(definitions) {
    var result = {};
    definitions.forEach(function(item) {
      var key = item.wiStep === PENDING_HANDOFF_WI ? 'รอรับเรื่อง' : item.forcedStatus || item.phase;
      result[key] = (result[key] || 0) + 1;
    });
    return result;
  }

  function sampleOwner_(number) {
    var owners = [
      'ภ.จุลชีววิทยา',
      'ภ.อายุรศาสตร์',
      'ภ.ศัลยศาสตร์',
      'สำนักงานคณบดี',
      'งานบริการพยาบาล',
      'ศูนย์วิจัยคลินิก',
      'ภ.รังสีวิทยา',
      'ภ.เวชศาสตร์ฉุกเฉิน',
    ];
    return owners[(number - 1) % owners.length];
  }

  function addDays_(date, days) {
    var result = new Date(date.getTime());
    result.setDate(result.getDate() + Number(days || 0));
    return result;
  }

  return {
    VERSION: VERSION,
    seedTwentyDiverseSampleProjects: seedTwentyDiverseSampleProjects,
  };
})();

function seedTwentyDiverseSampleProjects() {
  return SampleDataTools.seedTwentyDiverseSampleProjects();
}
