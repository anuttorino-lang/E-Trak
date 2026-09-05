var FlowRegressionTestTools = (function() {
  var VERSION = '1.0.0';
  var MASTER_SHEET_NAME = 'Master Tracking';
  var ADMIN_SHEET_NAME = 'Admin_Input';
  var PHASE_HISTORY_SHEET_NAME = 'Phase_History';

  function runLatestFlowRegressionTest() {
    setupCoreSheets_();
    clearOperationalSheets_();

    var tokens = loginAllAdmins_();
    var created = createBuildingProject_(tokens.building);
    var code = created.projectCode;
    var checks = [];

    assertProject_(checks, code, 'Phase 1', '1.สำรวจพื้นที่', 'created phase 1 wi1');

    transferNext_(tokens.building, code, 'building p1 wi1 to wi2');
    assertProject_(checks, code, 'Phase 1', '2.เสนอคณะทำงานกลั่นกรองพิจารณา', 'phase 1 wi2');

    approvePhase1_(tokens.building, code);
    assertProject_(checks, code, 'Phase 2', '1.เขียนแบบ', 'approval sends to phase 2 wi1');
    assertTimelineHas_(checks, code, 'Phase 1', '2.เสนอคณะทำงานกลั่นกรองพิจารณา', 'อนุมัติรายการและส่งต่อ Phase 2', 'approval note stays on phase 1 wi2');
    assertTimelineHas_(checks, code, 'Phase 2', '1.เขียนแบบ', 'ส่งต่อขั้นตอนถัดไป', 'phase 2 wi1 log exists');

    repeatTransfer_(tokens.building, code, 6, 'building to phase 3');
    assertProject_(checks, code, 'Phase 3', '1.ส่งมอบเอกสารให้ผู้รับบริการ (User)', 'phase 3 wi1');

    transferNext_(tokens.building, code, 'building handoff to plan');
    assertPendingHandoff_(checks, code, 'Phase 4', 'งานนโยบายและแผน', 'building to plan pending handoff');
    assertNoPhaseEntryDate_(checks, code, 'pending phase 4 has no phase entry date');

    assertActiveIncludes_(checks, tokens.plan, code, 'plan inbox sees pending item');
    acceptHandoff_(tokens.plan, code, 'plan accepts');
    assertProject_(checks, code, 'Phase 4', '1.ตรวจสอบเอกสาร', 'plan accept starts phase 4 wi1');
    assertHasPhaseEntryDate_(checks, code, 'phase 4 starts after accept');

    transferNext_(tokens.plan, code, 'phase 4 wi1 to wi2');
    assertProject_(checks, code, 'Phase 4', '2.เสนอผู้บริหารพิจารณาแหล่งงบประมาณ', 'phase 4 wi2');
    transferNext_(tokens.plan, code, 'phase 4 wi2 to wi3 without budget');
    assertProject_(checks, code, 'Phase 4', '3.แจ้งผลการอนุมัติ(ใบขวาง) (User)', 'phase 4 wi3');

    assertBudgetRequired_(checks, tokens.plan, code);
    transferNext_(tokens.plan, code, 'phase 4 wi3 to procurement pending', {
      budgetSource: 'งบทดสอบถดถอย',
      budgetAmount: 9876543,
      fiscalYear: '2570',
      note: 'ทดสอบกรอกข้อมูลอนุมัติงบประมาณและส่งต่องานพัสดุ',
    });
    assertPendingHandoff_(checks, code, 'Phase 5', 'งานพัสดุและยานพาหนะ', 'plan to procurement pending handoff');
    assertBudget_(checks, code, 9876543, 'งบทดสอบถดถอย', '2570');

    assertActiveIncludes_(checks, tokens.procurement, code, 'procurement inbox sees pending item');
    acceptHandoff_(tokens.procurement, code, 'procurement accepts');
    assertProject_(checks, code, 'Phase 5', '1.แต่งตั้งคณะกรรมการ', 'procurement accept starts phase 5 wi1');

    repeatTransfer_(tokens.procurement, code, 7, 'procurement phases 5-7');
    assertProject_(checks, code, 'Phase 7', '1.บันทึกทะเบียนคุมสินทรัพย์', 'phase 7 wi1');

    PhaseTransferApi.completeProject({
      authToken: tokens.procurement,
      projectCode: code,
      note: 'ทดสอบเสร็จสิ้น',
    });
    assertProject_(checks, code, 'Phase 7', 'เสร็จสิ้น', 'complete project');

    var finalProject = getProject_(code);
    var timeline = PhaseTransferApi.getProjectTimelineForPhaseTransfer(code);
    var history = PhaseHistoryTools.getProjectPhaseHistory(code);

    return {
      success: checks.every(function(check) { return check.pass; }),
      version: VERSION,
      projectCode: code,
      checks: checks,
      finalProject: summarizeProject_(finalProject),
      timelineCount: timeline.length,
      phaseHistoryCount: history.length,
      timeline: timeline.slice(0, 12),
      phaseHistory: history,
    };
  }

  function setupCoreSheets_() {
    MasterTrackingTools.setupMasterTracking();
    AdminInputImportTools.setupAdminInput();
    AuthApi.setupAdminAuthUsers();
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

  function loginAllAdmins_() {
    return {
      building: AuthApi.loginAdmin({ username: 'building', password: '1234' }).token,
      plan: AuthApi.loginAdmin({ username: 'plan', password: '1234' }).token,
      procurement: AuthApi.loginAdmin({ username: 'procurement', password: '1234' }).token,
    };
  }

  function createBuildingProject_(token) {
    return ProjectCreateApi.createNewProjectFromBuilding({
      authToken: token,
      documentNo: 'FLOW-TEST/001',
      receivedDate: formatDateInput_(new Date()),
      ownerUnit: '[TEST] หน่วยงานเจ้าของเรื่อง',
      projectName: '[TEST] ทดสอบ flow ล่าสุดทุก phase',
      note: 'สร้างข้อมูลทดสอบ flow ล่าสุด',
    });
  }

  function transferNext_(token, code, label, extra) {
    var payload = {
      authToken: token,
      projectCode: code,
      note: extra && extra.note ? extra.note : label,
    };
    if (extra) {
      if (Object.prototype.hasOwnProperty.call(extra, 'budgetSource')) payload.budgetSource = extra.budgetSource;
      if (Object.prototype.hasOwnProperty.call(extra, 'budgetAmount')) payload.budgetAmount = extra.budgetAmount;
      if (Object.prototype.hasOwnProperty.call(extra, 'fiscalYear')) payload.fiscalYear = extra.fiscalYear;
    }
    return PhaseTransferApi.transferProjectToNextStep(payload);
  }

  function repeatTransfer_(token, code, count, label) {
    for (var i = 0; i < count; i += 1) {
      transferNext_(token, code, label + ' #' + (i + 1));
    }
  }

  function approvePhase1_(token, code) {
    return AdminInputImportTools.approveProjectFromConsiderationSafe({
      authToken: token,
      projectCode: code,
      note: 'อนุมัติรายการและส่งต่อ Phase 2',
    });
  }

  function acceptHandoff_(token, code, note) {
    return PhaseTransferApi.acceptIncomingHandoff({
      authToken: token,
      projectCode: code,
      note: note || 'รับเรื่อง',
    });
  }

  function assertBudgetRequired_(checks, token, code) {
    try {
      PhaseTransferApi.transferProjectToNextStep({
        authToken: token,
        projectCode: code,
        note: 'ควรถูกปฏิเสธเพราะไม่กรอกงบประมาณ',
      });
      checks.push({ name: 'phase 4 wi3 requires budget fields', pass: false, detail: 'no error thrown' });
    } catch (error) {
      checks.push({
        name: 'phase 4 wi3 requires budget fields',
        pass: String(error.message || '').indexOf('กรุณากรอก') !== -1,
        detail: String(error.message || ''),
      });
    }
  }

  function assertProject_(checks, code, phase, wi, name) {
    var project = getProject_(code);
    checks.push({
      name: name,
      pass: Boolean(project && String(project.phase) === phase && String(project.wiStep) === wi),
      expected: phase + ' / ' + wi,
      actual: project ? project.phase + ' / ' + project.wiStep : 'not found',
    });
  }

  function assertPendingHandoff_(checks, code, phase, unit, name) {
    var project = getProject_(code);
    checks.push({
      name: name,
      pass: Boolean(project && project.phase === phase && project.wiStep === 'รอรับเรื่อง' && project.responsibleUnit === unit && String(project.phaseSlaStatus || '') === '⏳ รอรับเรื่อง'),
      expected: phase + ' / รอรับเรื่อง / ' + unit,
      actual: project ? project.phase + ' / ' + project.wiStep + ' / ' + project.responsibleUnit + ' / ' + project.phaseSlaStatus : 'not found',
    });
  }

  function assertNoPhaseEntryDate_(checks, code, name) {
    var project = getProject_(code);
    checks.push({
      name: name,
      pass: Boolean(project && !project.phaseEntryDate),
      actual: project ? project.phaseEntryDate : 'not found',
    });
  }

  function assertHasPhaseEntryDate_(checks, code, name) {
    var project = getProject_(code);
    checks.push({
      name: name,
      pass: Boolean(project && project.phaseEntryDate),
      actual: project ? project.phaseEntryDate : 'not found',
    });
  }

  function assertActiveIncludes_(checks, token, code, name) {
    var list = PhaseTransferApi.getActiveProjectsForPhaseTransfer(token);
    checks.push({
      name: name,
      pass: list.some(function(item) { return item.projectCode === code && item.wiStep === 'รอรับเรื่อง'; }),
      count: list.length,
    });
  }

  function assertTimelineHas_(checks, code, phase, wi, note, name) {
    var timeline = PhaseTransferApi.getProjectTimelineForPhaseTransfer(code);
    checks.push({
      name: name,
      pass: timeline.some(function(item) {
        return item.phase === phase && item.wiStep === wi && String(item.note || '') === note;
      }),
      expected: phase + ' / ' + wi + ' / ' + note,
    });
  }

  function assertBudget_(checks, code, amount, source, year) {
    var project = getProject_(code);
    checks.push({
      name: 'budget approval fields saved before procurement handoff',
      pass: Boolean(project && Number(project.budgetAmount) === Number(amount) && String(project.budgetSource) === source && String(project.fiscalYear) === year),
      expected: [amount, source, year].join(' / '),
      actual: project ? [project.budgetAmount, project.budgetSource, project.fiscalYear].join(' / ') : 'not found',
    });
  }

  function getProject_(code) {
    var list = PhaseTransferApi.getActiveProjectsForPhaseTransfer(AuthApi.loginAdmin({ username: 'building', password: '1234' }).token)
      .concat(PhaseTransferApi.getActiveProjectsForPhaseTransfer(AuthApi.loginAdmin({ username: 'plan', password: '1234' }).token))
      .concat(PhaseTransferApi.getActiveProjectsForPhaseTransfer(AuthApi.loginAdmin({ username: 'procurement', password: '1234' }).token));
    for (var i = 0; i < list.length; i += 1) {
      if (list[i].projectCode === code) return list[i];
    }
    return getProjectFromMaster_(code);
  }

  function getProjectFromMaster_(code) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return null;
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, MasterTrackingTools.MASTER_HEADERS.length).getValues();
    for (var i = 0; i < rows.length; i += 1) {
      if (String(rows[i][0]) === String(code)) {
        return {
          projectCode: rows[i][0],
          projectName: rows[i][4],
          budgetAmount: rows[i][5],
          budgetSource: rows[i][6],
          fiscalYear: rows[i][7],
          phase: rows[i][8],
          wiStep: rows[i][9],
          phaseEntryDate: rows[i][10],
          phaseSlaStatus: rows[i][13],
          responsibleUnit: rows[i][14],
        };
      }
    }
    return null;
  }

  function summarizeProject_(project) {
    if (!project) return null;
    return {
      projectCode: project.projectCode,
      phase: project.phase,
      wiStep: project.wiStep,
      phaseEntryDate: project.phaseEntryDate ? String(project.phaseEntryDate) : '',
      phaseSlaStatus: project.phaseSlaStatus,
      responsibleUnit: project.responsibleUnit,
      budgetAmount: project.budgetAmount,
      budgetSource: project.budgetSource,
      fiscalYear: project.fiscalYear,
    };
  }

  function formatDateInput_(date) {
    return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0') + '-' + String(date.getDate()).padStart(2, '0');
  }

  return {
    VERSION: VERSION,
    runLatestFlowRegressionTest: runLatestFlowRegressionTest,
  };
})();

function runLatestFlowRegressionTest() {
  return FlowRegressionTestTools.runLatestFlowRegressionTest();
}
