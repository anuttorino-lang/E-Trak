var MasterTrackingTools = (function() {
  var VERSION = '1.4.0';
  var MASTER_SHEET_NAME = 'Master Tracking';
  var ADMIN_INPUT_SHEET_NAME = 'Admin_Input';
  var PROJECT_PREFIX = 'A';
  var PROJECT_DIGITS = 4;
  var DEFAULT_PHASE_SLA_DAYS = 30;
  var DEFAULT_MAIN_SLA_DAYS = 730;

  var MASTER_HEADERS = [
    'รหัสโครงการ',
    'เลขที่หนังสือ',
    'หน่วยงานเจ้าของเรื่อง / ต้นเรื่อง',
    'วันที่รับเรื่อง',
    'ชื่อโครงการ / รายการ',
    'วงเงินงบประมาณ',
    'แหล่งงบประมาณ',
    'ปีงบประมาณ',
    'SOP / Phase ปัจจุบัน',
    'WI / ขั้นตอนย่อยปัจจุบัน',
    'วันที่รับเข้า Phase ปัจจุบัน',
    'SLA Phase (วัน)',
    'วันคงเหลือของ Phase',
    'สถานะ SLA Phase',
    'ฝ่ายรับผิดชอบปัจจุบัน',
    'วันที่ครบกำหนดตามระเบียบ',
    'SLA รายการหลัก (วัน)',
    'จำนวนวันที่ใช้ไปของรายการหลัก',
    'วันคงเหลือของรายการหลัก',
    'สถานะ SLA รายการหลัก (ไม่ใช้)',
    'หมายเหตุ / Action Required',
    'อัปเดตล่าสุด',
  ];

  var PHASES = [
    'Phase 1',
    'Phase 2',
    'Phase 3',
    'Phase 4',
    'Phase 5',
    'Phase 6',
    'Phase 7',
    
  ];

  var WI_BY_PHASE = {
    'Phase 1': [
      '1.สำรวจพื้นที่',
      '2.เสนอคณะทำงานกลั่นกรองพิจารณา',
      '3.แจ้งมติที่ประชุม',
    ],
    'Phase 2': [
      '1.เขียนแบบ',
      '2.คณะกรรมการพิจารณาแบบ',
      '3.ผู้รับบริการรับรองแบบ (User)',
      '4.ถอดแบบ (BOQ) คำนวณราคากลาง (Factor F)',
      '5.คณะกรรมการพิจารณาราคาประมาณการ',
      '6.เสนอคณบดีเห็นชอบราคาประมาณการ',
    ],
    'Phase 3': [
      '1.ส่งมอบเอกสารให้ผู้รับบริการ (User)',
    ],
    'Phase 4': [
      '1.ตรวจสอบเอกสาร',
      '2.เสนอผู้บริหารพิจารณาแหล่งงบประมาณ',
      '3.แจ้งผลการอนุมัติ(ใบขวาง) (User)',
    ],
    'Phase 5': [
      '1.แต่งตั้งคณะกรรมการ',
      '2.ราคากลาง',
      '3.จัดทำ TOR',
      '4.ประกวดราคา',
      '5.ลงนามในสัญญา',
    ],
    'Phase 6': [
      '1.ตรวจรับงานจ้าง',
      '2.เบิกจ่ายเงินงวด',
    ],
    'Phase 7': [
      '1.บันทึกทะเบียนคุมสินทรัพย์',
    ],
  };

  var PHASE_DEFINITIONS = {
    'Phase 1': 'รับเรื่องและกลั่นกรอง (Survey & Preliminary Design)',
    'Phase 2': 'ออกแบบและประเมินราคา (Drawing & BOQ)',
    'Phase 3': 'ส่งมอบแบบรูปรายการ (Handover of Construction Drawings)',
    'Phase 4': 'การอนุมัติงบประมาณ (Budget Approval)',
    'Phase 5': 'กระบวนการจัดจ้าง (Procurement & Contract)',
    'Phase 6': 'บริหารสัญญา (Contract Management)',
    'Phase 7': 'บันทึกทะเบียนพัสดุ (Inventory Management)',
  };

  var SLA_BY_PHASE_WI = {
    'Phase 1': {
      '1.สำรวจพื้นที่': 30,
      '2.เสนอคณะทำงานกลั่นกรองพิจารณา': 30,
      '3.แจ้งมติที่ประชุม': 5,
    },
    'Phase 2': {
      '1.เขียนแบบ': 30,
      '2.คณะกรรมการพิจารณาแบบ': 30,
      '3.ผู้รับบริการรับรองแบบ (User)': 5,
      '4.ถอดแบบ (BOQ) คำนวณราคากลาง (Factor F)': 30,
      '5.คณะกรรมการพิจารณาราคาประมาณการ': 30,
      '6.เสนอคณบดีเห็นชอบราคาประมาณการ': 5,
    },
    'Phase 3': {
      '1.ส่งมอบเอกสารให้ผู้รับบริการ (User)': 5,
    },
    'Phase 4': {
      '1.ตรวจสอบเอกสาร': 5,
      '2.เสนอผู้บริหารพิจารณาแหล่งงบประมาณ': 5,
      '3.แจ้งผลการอนุมัติ(ใบขวาง) (User)': 5,
    },
    'Phase 5': {
      '1.แต่งตั้งคณะกรรมการ': 10,
      '2.ราคากลาง': 30,
      '3.จัดทำ TOR': 10,
      '4.ประกวดราคา': 60,
      '5.ลงนามในสัญญา': 10,
    },
    'Phase 6': {
      '1.ตรวจรับงานจ้าง': 365,
      '2.เบิกจ่ายเงินงวด': 0,
    },
    'Phase 7': {
      '1.บันทึกทะเบียนคุมสินทรัพย์': 30,
    },
  };

  var TOTAL_SLA_BY_PHASE = {
    'Phase 1': 65,
    'Phase 2': 130,
    'Phase 3': 5,
    'Phase 4': 15,
    'Phase 5': 120,
    'Phase 6': 365,
    'Phase 7': 30,
  };

  var LEGACY_WI_ALIASES = {
    'Phase 1': {
      '1.รอสำรวจ': '1.สำรวจพื้นที่',
      '2.รอกรรมการกลั่นกรองพิจารณา': '2.เสนอคณะทำงานกลั่นกรองพิจารณา',
      '3.พิจารณารายการ': '3.แจ้งมติที่ประชุม',
    },
    'Phase 2': {
      '1.กำลังออกแบบ': '1.เขียนแบบ',
      '2.รอเจ้าของเรื่องเซ็นรับรองแบบ': '3.ผู้รับบริการรับรองแบบ (User)',
      '3.กำลังจัดทำราคากลาง': '4.ถอดแบบ (BOQ) คำนวณราคากลาง (Factor F)',
      '4.รอกรรมการ Factor F อนุมัติ': '5.คณะกรรมการพิจารณาราคาประมาณการ',
      '5.เสนอคณบดีเห็นชอบราคากลาง': '6.เสนอคณบดีเห็นชอบราคาประมาณการ',
    },
    'Phase 3': {
      'ส่งมอบ BOQ ให้หน่วยงานเพื่อทำโครงการ': '1.ส่งมอบเอกสารให้ผู้รับบริการ (User)',
    },
    'Phase 4': {
      '3.แจ้งผลการจัดสรร(ใบขวาง)': '3.แจ้งผลการอนุมัติ(ใบขวาง) (User)',
    },
    'Phase 5': {
      '2.จัดทำ TOR': '3.จัดทำ TOR',
      '3.ประกวดราคา': '4.ประกวดราคา',
      '4.ลงนามในสัญญา': '5.ลงนามในสัญญา',
    },
    'Phase 6': {
      '1.ตรวจรับพัสดุ': '1.ตรวจรับงานจ้าง',
      '2.เบิกจ่าย': '2.เบิกจ่ายเงินงวด',
    },
    'Phase 7': {
      'บันทึกทะเบียนคุมสินทรัพย์': '1.บันทึกทะเบียนคุมสินทรัพย์',
    },
  };

  var RESPONSIBLE_UNIT_BY_PHASE = {
    'Phase 1': 'งานอาคารสถานที่',
    'Phase 2': 'งานอาคารสถานที่',
    'Phase 3': 'งานอาคารสถานที่',
    'Phase 4': 'งานนโยบายและแผน',
    'Phase 5': 'งานพัสดุและยานพาหนะ',
    'Phase 6': 'งานพัสดุและยานพาหนะ',
    'Phase 7': 'งานพัสดุและยานพาหนะ',
  };

  function setupMasterTracking() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet_(ss, MASTER_SHEET_NAME);

    ensureMasterShape_(sheet);
    setMasterHeaders_(sheet);
    formatMasterSheet_(sheet);
    setMasterValidation_(sheet);
    applyWiValidationByExistingPhase_(sheet);
    setMasterFormulas_(sheet);

    return {
      success: true,
      version: VERSION,
      sheetName: MASTER_SHEET_NAME,
    };
  }

  function installOnEditTrigger() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var triggers = ScriptApp.getProjectTriggers();
    var exists = triggers.some(function(trigger) {
      return trigger.getHandlerFunction() === 'handleMasterTrackingEdit';
    });

    if (!exists) {
      ScriptApp.newTrigger('handleMasterTrackingEdit')
        .forSpreadsheet(ss)
        .onEdit()
        .create();
    }

    return {
      success: true,
      installed: !exists,
    };
  }

  function handleEdit(e) {
    if (!e || !e.range) return;

    var sheet = e.range.getSheet();
    if (sheet.getName() !== MASTER_SHEET_NAME) return;

    var row = e.range.getRow();
    var column = e.range.getColumn();
    if (row < 2 || column !== 9) return;

    var phase = sheet.getRange(row, 9).getValue();
    var defaultWi = getWiListByPhase(phase)[0] || '';

    setWiValidationForRow_(sheet, row, phase);
    sheet.getRange(row, 10).setValue(defaultWi);
    sheet.getRange(row, 11).setValue(new Date());
    sheet.getRange(row, 12).setValue(getSlaDays(phase, defaultWi));
    sheet.getRange(row, 22).setValue(new Date());
    setMasterFormulasForRow_(sheet, row);
  }

  function generateNextProjectCode() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var codes = []
      .concat(readProjectCodes_(ss.getSheetByName(MASTER_SHEET_NAME)))
      .concat(readProjectCodes_(ss.getSheetByName(ADMIN_INPUT_SHEET_NAME)));
    var maxNumber = 0;
    var pattern = new RegExp('^' + PROJECT_PREFIX + '-(\\d+)$');

    codes.forEach(function(code) {
      var match = String(code).match(pattern);
      if (match) {
        maxNumber = Math.max(maxNumber, Number(match[1]));
      }
    });

    return PROJECT_PREFIX + '-' + String(maxNumber + 1).padStart(PROJECT_DIGITS, '0');
  }

  function getWiListByPhase(phase) {
    return WI_BY_PHASE[phase] || [];
  }

  function normalizeWiStep(phase, wiStep) {
    var text = String(wiStep || '').trim();
    return (LEGACY_WI_ALIASES[phase] && LEGACY_WI_ALIASES[phase][text]) || text;
  }

  function getSlaDays(phase, wiStep) {
    var normalizedWi = normalizeWiStep(phase, wiStep);
    var phaseSla = SLA_BY_PHASE_WI[phase] || {};
    return Object.prototype.hasOwnProperty.call(phaseSla, normalizedWi)
      ? Number(phaseSla[normalizedWi])
      : DEFAULT_PHASE_SLA_DAYS;
  }

  function getPhaseTotalSla(phase) {
    return Number(TOTAL_SLA_BY_PHASE[phase] || DEFAULT_PHASE_SLA_DAYS);
  }

  function getResponsibleUnitByPhase(phase) {
    return RESPONSIBLE_UNIT_BY_PHASE[phase] || '';
  }

  function getPhaseList() {
    return PHASES.slice();
  }

  function getAllWiList() {
    var seen = {};
    var result = [];
    Object.keys(WI_BY_PHASE).forEach(function(phase) {
      WI_BY_PHASE[phase].forEach(function(wi) {
        if (!seen[wi]) {
          seen[wi] = true;
          result.push(wi);
        }
      });
    });
    return result;
  }

  function setMasterFormulas() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
    if (!sheet) return;
    setMasterFormulas_(sheet);
  }

  function setMasterFormulasForRow(row) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
    if (!sheet || row < 2) return;
    setMasterFormulasForRow_(sheet, row);
  }

  function setWiValidationForRow(row, phase) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
    if (!sheet || row < 2) return;
    setWiValidationForRow_(sheet, row, phase);
  }

  function ensureMasterShape_(sheet) {
    if (sheet.getMaxColumns() < MASTER_HEADERS.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), MASTER_HEADERS.length - sheet.getMaxColumns());
    }
    if (sheet.getMaxColumns() > MASTER_HEADERS.length) {
      sheet.deleteColumns(MASTER_HEADERS.length + 1, sheet.getMaxColumns() - MASTER_HEADERS.length);
    }
    if (sheet.getMaxRows() < 500) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 500 - sheet.getMaxRows());
    }
  }

  function setMasterHeaders_(sheet) {
    sheet.getRange(1, 1, 1, MASTER_HEADERS.length).setValues([MASTER_HEADERS]);
  }

  function formatMasterSheet_(sheet) {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, MASTER_HEADERS.length)
      .setFontWeight('bold')
      .setFontColor('#ffffff')
      .setBackground('#075f6c')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, MASTER_HEADERS.length)
      .setVerticalAlignment('middle')
      .setWrap(true);

    sheet.getRange(2, 4, sheet.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(2, 6, sheet.getMaxRows() - 1, 1).setNumberFormat('#,##0.00');
    sheet.getRange(2, 11, sheet.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(2, 16, sheet.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(2, 22, sheet.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  }

  function setMasterValidation_(sheet) {
    var maxRows = sheet.getMaxRows() - 1;
    var phaseRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(PHASES, true)
      .setAllowInvalid(false)
      .build();
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['🟢 ปกติ', '🟡 ใกล้กำหนด', '🔴 ล่าช้า', '❌ ไม่อนุมัติ', '⚫ ยกเลิกรายการ', '🔵 เสร็จสิ้น'], true)
      .setAllowInvalid(true)
      .build();
    var dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .build();

    sheet.getRange(2, 4, maxRows, 1).setDataValidation(dateRule);
    sheet.getRange(2, 9, maxRows, 1).setDataValidation(phaseRule);
    sheet.getRange(2, 11, maxRows, 1).setDataValidation(dateRule);
    sheet.getRange(2, 14, maxRows, 1).setDataValidation(statusRule);
  }

  function setMasterFormulas_(sheet) {
    var lastRow = Math.max(sheet.getLastRow(), 2);
    for (var row = 2; row <= lastRow; row += 1) {
      if (sheet.getRange(row, 1).getValue()) {
        setMasterFormulasForRow_(sheet, row);
      }
    }
  }

  function setMasterFormulasForRow_(sheet, row) {
    sheet.getRange(row, 13).setFormula('=IF($K' + row + '="","",$L' + row + '-(INT(TODAY())-INT($K' + row + ')+1))');
    sheet.getRange(row, 14).setFormula('=IF(REGEXMATCH($J' + row + '&$U' + row + ',"ไม่อนุมัติ"),"❌ ไม่อนุมัติ",IF($J' + row + '="เสร็จสิ้น","🔵 เสร็จสิ้น",IF(REGEXMATCH($J' + row + ',"ยกเลิก"),"⚫ ยกเลิกรายการ",IF($K' + row + '="","",IF($M' + row + '<0,"🔴 ล่าช้า",IF($M' + row + '<=5,"🟡 ใกล้กำหนด","🟢 ปกติ"))))))');
    sheet.getRange(row, 15).setFormula('=SWITCH($I' + row + ',"Phase 1","งานอาคารสถานที่","Phase 2","งานอาคารสถานที่","Phase 3","งานอาคารสถานที่","Phase 4","งานนโยบายและแผน","Phase 5","งานพัสดุและยานพาหนะ","Phase 6","งานพัสดุและยานพาหนะ","Phase 7","งานพัสดุและยานพาหนะ","")');
    sheet.getRange(row, 16).setFormula('=IF(OR($D' + row + '="",$Q' + row + '=""),"",INT($D' + row + ')+$Q' + row + '-1)');
    sheet.getRange(row, 18).setFormula('=IF($D' + row + '="","",INT(TODAY())-INT($D' + row + ')+1)');
    sheet.getRange(row, 19).setFormula('=IF($P' + row + '="","",INT($P' + row + ')-INT(TODAY()))');
    sheet.getRange(row, 20).clearContent();
  }

  function applyWiValidationByExistingPhase_(sheet) {
    var lastRow = Math.max(sheet.getLastRow(), 2);
    for (var row = 2; row <= lastRow; row += 1) {
      var phase = sheet.getRange(row, 9).getValue();
      if (phase) {
        setWiValidationForRow_(sheet, row, phase);
      }
    }
  }

  function setWiValidationForRow_(sheet, row, phase) {
    var wiList = getWiListByPhase(phase);
    if (!wiList.length) {
      wiList = getAllWiList();
    }

    var wiRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(wiList, true)
      .setAllowInvalid(true)
      .build();

    sheet.getRange(row, 10).setDataValidation(wiRule);
  }

  function readProjectCodes_(sheet) {
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
      .map(function(row) { return row[0]; })
      .filter(Boolean);
  }

  function getOrCreateSheet_(ss, sheetName) {
    return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  }

  return {
    VERSION: VERSION,
    MASTER_SHEET_NAME: MASTER_SHEET_NAME,
    ADMIN_INPUT_SHEET_NAME: ADMIN_INPUT_SHEET_NAME,
    MASTER_HEADERS: MASTER_HEADERS,
    PHASES: PHASES,
    WI_BY_PHASE: WI_BY_PHASE,
    PHASE_DEFINITIONS: PHASE_DEFINITIONS,
    SLA_BY_PHASE_WI: SLA_BY_PHASE_WI,
    TOTAL_SLA_BY_PHASE: TOTAL_SLA_BY_PHASE,
    LEGACY_WI_ALIASES: LEGACY_WI_ALIASES,
    RESPONSIBLE_UNIT_BY_PHASE: RESPONSIBLE_UNIT_BY_PHASE,
    DEFAULT_PHASE_SLA_DAYS: DEFAULT_PHASE_SLA_DAYS,
    DEFAULT_MAIN_SLA_DAYS: DEFAULT_MAIN_SLA_DAYS,
    setupMasterTracking: setupMasterTracking,
    installOnEditTrigger: installOnEditTrigger,
    handleEdit: handleEdit,
    generateNextProjectCode: generateNextProjectCode,
    getWiListByPhase: getWiListByPhase,
    normalizeWiStep: normalizeWiStep,
    getSlaDays: getSlaDays,
    getPhaseTotalSla: getPhaseTotalSla,
    getResponsibleUnitByPhase: getResponsibleUnitByPhase,
    getPhaseList: getPhaseList,
    getAllWiList: getAllWiList,
    setMasterFormulas: setMasterFormulas,
    setMasterFormulasForRow: setMasterFormulasForRow,
    setWiValidationForRow: setWiValidationForRow,
  };
})();

function setupMasterTrackingTools() {
  return MasterTrackingTools.setupMasterTracking();
}

function installMasterTrackingEditTrigger() {
  return MasterTrackingTools.installOnEditTrigger();
}

function handleMasterTrackingEdit(e) {
  return MasterTrackingTools.handleEdit(e);
}
