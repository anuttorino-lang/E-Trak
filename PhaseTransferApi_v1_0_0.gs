var PhaseTransferApi = (function() {
  var VERSION = '1.0.5';
  var MASTER_SHEET_NAME = 'Master Tracking';
  var ADMIN_SHEET_NAME = 'Admin_Input';
  var UNIT_BUILDING = 'งานอาคารสถานที่';
  var UNIT_PLAN = 'งานนโยบายและแผน';
  var UNIT_PROCUREMENT = 'งานพัสดุและยานพาหนะ';
  var PHASE_COMPLETED = 'เสร็จสิ้น';
  var WI_CANCELLED = 'ยกเลิกรายการ';
  var PENDING_HANDOFF_WI = 'รอรับเรื่อง';
  var PENDING_HANDOFF_STATUS = '⏳ รอรับเรื่อง';

  function transferProjectPhase(payload) {
    var adminUser = AuthApi.requireAdminSession(payload && payload.authToken);
    validatePayload_(payload);
    validatePhasePermission_(adminUser, payload.phase);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) throw new Error('ไม่พบชีต Master Tracking');

    var project = getProjectByCode_(masterSheet, payload.projectCode);
    if (!project) throw new Error('ไม่พบรหัสโครงการ: ' + payload.projectCode);

    var now = new Date();
    var phase = payload.phase;
    var wiStep = payload.wiStep;
    var responsibleUnit = MasterTrackingTools.getResponsibleUnitByPhase(phase);
    validateBudgetApprovalPayloadIfNeeded_(payload, project.phase, project.wiStep);
    var logProject = applyBudgetApprovalPayloadToProject_(project, payload, project.phase, project.wiStep);

    recordPhaseClosureIfNeeded_(project, phase, now, payload.note || '');
    appendTransferLog_(logProject, phase, wiStep, responsibleUnit, payload.note || '', adminUser.username, now);
    var result = AdminInputImportTools.importPendingAdminInputToMasterTracking();
    MasterTrackingTools.setMasterFormulasForRow(project.rowNumber);

    return {
      success: true,
      projectCode: project.projectCode,
      phase: phase,
      wiStep: wiStep,
      responsibleUnit: responsibleUnit,
      imported: result.imported,
      failed: result.failed,
    };
  }

  function transferProjectToNextUnit(payload) {
    var adminUser = AuthApi.requireAdminSession(payload && payload.authToken);
    if (!payload || !payload.projectCode) throw new Error('กรุณาเลือกรหัสโครงการ');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) throw new Error('ไม่พบชีต Master Tracking');

    var project = getProjectByCode_(masterSheet, payload.projectCode);
    if (!project) throw new Error('ไม่พบรหัสโครงการ: ' + payload.projectCode);

    var targetPhase = getNextTransferTarget_(adminUser, project.phase);
    if (!targetPhase) {
      throw new Error('โครงการนี้ยังไม่อยู่ Phase สุดท้ายของหน่วยงานคุณ จึงยังส่งต่อไปงานถัดไปไม่ได้');
    }

    var now = new Date();
    var wiStep = MasterTrackingTools.getWiListByPhase(targetPhase)[0] || '';
    var responsibleUnit = MasterTrackingTools.getResponsibleUnitByPhase(targetPhase);
    var note = payload.note || 'ส่งต่อไปยัง ' + responsibleUnit;

    recordPhaseClosureIfNeeded_(project, targetPhase, now, note);
    createPendingHandoff_(masterSheet, project, targetPhase, responsibleUnit, note, adminUser.username, now);

    return {
      success: true,
      projectCode: project.projectCode,
      phase: targetPhase,
      wiStep: PENDING_HANDOFF_WI,
      responsibleUnit: responsibleUnit,
      pendingHandoff: true,
    };
  }

  function saveProjectNoteOnly(payload) {
    var adminUser = AuthApi.requireAdminSession(payload && payload.authToken);
    if (!payload || !payload.projectCode) throw new Error('กรุณาเลือกรหัสโครงการ');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) throw new Error('ไม่พบชีต Master Tracking');

    var project = getProjectByCode_(masterSheet, payload.projectCode);
    if (!project) throw new Error('ไม่พบรหัสโครงการ: ' + payload.projectCode);
    validatePhasePermission_(adminUser, project.phase);

    var now = new Date();
    var note = payload.note || '';
    var hasBudgetPayload = isPhase4BudgetApprovalStep_(project.phase, project.wiStep) && hasBudgetApprovalPayload_(payload);
    if (!note && !hasBudgetPayload) throw new Error('กรุณากรอกหมายเหตุก่อนบันทึก');
    validateBudgetApprovalPayloadIfNeeded_(payload, project.phase, project.wiStep);
    var logProject = applyBudgetApprovalPayloadToProject_(project, payload, project.phase, project.wiStep);

    appendTransferLog_(
      logProject,
      project.phase,
      project.wiStep,
      project.responsibleUnit || MasterTrackingTools.getResponsibleUnitByPhase(project.phase),
      note || 'อัปเดตข้อมูลอนุมัติงบประมาณ',
      adminUser.username,
      now,
      project.phaseEntryDate || now
    );
    var result = AdminInputImportTools.importPendingAdminInputToMasterTracking();
    MasterTrackingTools.setMasterFormulasForRow(project.rowNumber);

    return {
      success: true,
      projectCode: project.projectCode,
      phase: project.phase,
      wiStep: project.wiStep,
      responsibleUnit: project.responsibleUnit,
      imported: result.imported,
      failed: result.failed,
    };
  }

  function transferProjectToNextStep(payload) {
    var adminUser = AuthApi.requireAdminSession(payload && payload.authToken);
    if (!payload || !payload.projectCode) throw new Error('กรุณาเลือกรหัสโครงการ');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) throw new Error('ไม่พบชีต Master Tracking');

    var project = getProjectByCode_(masterSheet, payload.projectCode);
    if (!project) throw new Error('ไม่พบรหัสโครงการ: ' + payload.projectCode);
    validatePhasePermission_(adminUser, project.phase);

    var target = getNextStepTarget_(project);
    if (!target) throw new Error('รายการนี้ไม่มีขั้นตอนถัดไป');
    validateNextStepPermission_(adminUser, project.phase, target.phase);

    var now = new Date();
    var responsibleUnit = MasterTrackingTools.getResponsibleUnitByPhase(target.phase);
    var note = payload.note || 'ส่งต่อขั้นตอนถัดไป';
    validateBudgetApprovalPayloadIfNeeded_(payload, project.phase, project.wiStep);
    var logProject = applyBudgetApprovalPayloadToProject_(project, payload, project.phase, project.wiStep);

    if (isCrossUnitTransfer_(project.phase, target.phase)) {
      recordPhaseClosureIfNeeded_(project, target.phase, now, note);
      createPendingHandoff_(masterSheet, logProject, target.phase, responsibleUnit, note, adminUser.username, now);
      return {
        success: true,
        projectCode: project.projectCode,
        phase: target.phase,
        wiStep: PENDING_HANDOFF_WI,
        responsibleUnit: responsibleUnit,
        pendingHandoff: true,
      };
    }

    recordPhaseClosureIfNeeded_(project, target.phase, now, note);
    appendTransferLog_(logProject, target.phase, target.wiStep, responsibleUnit, note, adminUser.username, now);
    var result = AdminInputImportTools.importPendingAdminInputToMasterTracking();
    MasterTrackingTools.setMasterFormulasForRow(project.rowNumber);

    return {
      success: true,
      projectCode: project.projectCode,
      phase: target.phase,
      wiStep: target.wiStep,
      responsibleUnit: responsibleUnit,
      imported: result.imported,
      failed: result.failed,
    };
  }

  function transferProjectToUnit(payload) {
    var adminUser = AuthApi.requireAdminSession(payload && payload.authToken);
    if (!payload || !payload.projectCode) throw new Error('กรุณาเลือกรหัสโครงการ');
    if (!payload.targetUnit) throw new Error('กรุณาเลือกหน่วยงานปลายทาง');

    var targetPhase = getFirstPhaseByUnit_(payload.targetUnit);
    if (!targetPhase) throw new Error('ไม่พบ Phase แรกของหน่วยงานปลายทาง: ' + payload.targetUnit);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) throw new Error('ไม่พบชีต Master Tracking');

    var project = getProjectByCode_(masterSheet, payload.projectCode);
    if (!project) throw new Error('ไม่พบรหัสโครงการ: ' + payload.projectCode);
    if (String(project.responsibleUnit || '') === String(payload.targetUnit || '')) {
      throw new Error('หน่วยงานปลายทางต้องเป็นงานอื่น ไม่ใช่ฝ่ายรับผิดชอบปัจจุบัน');
    }

    var now = new Date();
    var wiStep = MasterTrackingTools.getWiListByPhase(targetPhase)[0] || '';
    var responsibleUnit = MasterTrackingTools.getResponsibleUnitByPhase(targetPhase);
    var note = payload.note || 'ส่งต่อไปยัง ' + responsibleUnit;

    recordPhaseClosureIfNeeded_(project, targetPhase, now, note);
    createPendingHandoff_(masterSheet, project, targetPhase, responsibleUnit, note, adminUser.username, now);

    return {
      success: true,
      projectCode: project.projectCode,
      phase: targetPhase,
      wiStep: PENDING_HANDOFF_WI,
      responsibleUnit: responsibleUnit,
      pendingHandoff: true,
    };
  }

  function acceptIncomingHandoff(payload) {
    var adminUser = AuthApi.requireAdminSession(payload && payload.authToken);
    if (!payload || !payload.projectCode) throw new Error('กรุณาเลือกรหัสโครงการ');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) throw new Error('ไม่พบชีต Master Tracking');

    var project = getProjectByCode_(masterSheet, payload.projectCode);
    if (!project) throw new Error('ไม่พบรหัสโครงการ: ' + payload.projectCode);
    validatePhasePermission_(adminUser, project.phase);

    if (String(project.wiStep || '').trim() !== PENDING_HANDOFF_WI) {
      throw new Error('รายการนี้ไม่ได้อยู่ในสถานะรอรับเรื่อง');
    }

    var now = new Date();
    var phase = String(project.phase || '').trim();
    var firstWi = MasterTrackingTools.getWiListByPhase(phase)[0] || '';
    var responsibleUnit = MasterTrackingTools.getResponsibleUnitByPhase(phase);
    var note = payload.note || 'รับเรื่องเข้า ' + phase;

    appendTransferLog_(project, phase, firstWi, responsibleUnit, note, adminUser.username, now, now, AdminInputImportTools.IMPORTED_STATUS);
    MasterTrackingTools.setWiValidationForRow(project.rowNumber, phase);
    masterSheet.getRange(project.rowNumber, 10).setValue(firstWi);
    masterSheet.getRange(project.rowNumber, 11).setValue(now);
    masterSheet.getRange(project.rowNumber, 12).setValue(MasterTrackingTools.getSlaDays(phase, firstWi));
    masterSheet.getRange(project.rowNumber, 15).setValue(responsibleUnit);
    masterSheet.getRange(project.rowNumber, 21).setValue(note);
    masterSheet.getRange(project.rowNumber, 22).setValue(now);
    MasterTrackingTools.setMasterFormulasForRow(project.rowNumber);

    return {
      success: true,
      projectCode: project.projectCode,
      phase: phase,
      wiStep: firstWi,
      responsibleUnit: responsibleUnit,
      accepted: true,
    };
  }

  function completeProject(payload) {
    var adminUser = AuthApi.requireAdminSession(payload && payload.authToken);
    if (!payload || !payload.projectCode) throw new Error('กรุณาเลือกรหัสโครงการ');
    if (!isProcurementAdmin_(adminUser)) {
      throw new Error('เฉพาะงานพัสดุและยานพาหนะเท่านั้นที่กดเสร็จสิ้นได้');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) throw new Error('ไม่พบชีต Master Tracking');

    var project = getProjectByCode_(masterSheet, payload.projectCode);
    if (!project) throw new Error('ไม่พบรหัสโครงการ: ' + payload.projectCode);
    if (String(project.phase || '').trim() !== 'Phase 7' ||
        MasterTrackingTools.normalizeWiStep('Phase 7', project.wiStep) !== '1.บันทึกทะเบียนคุมสินทรัพย์') {
      throw new Error('กดเสร็จสิ้นได้เฉพาะรายการที่อยู่ Phase 7 / 1.บันทึกทะเบียนคุมสินทรัพย์');
    }

    var now = new Date();
    var phase = String(project.phase || 'Phase 7');
    var responsibleUnit = project.responsibleUnit || MasterTrackingTools.getResponsibleUnitByPhase(phase);
    var note = payload.note || PHASE_COMPLETED;

    PhaseHistoryTools.recordClosedPhase(project, now, project.wiStep, note);
    appendTransferLog_(project, phase, PHASE_COMPLETED, responsibleUnit, note, adminUser.username, now);
    var result = AdminInputImportTools.importPendingAdminInputToMasterTracking();
    MasterTrackingTools.setMasterFormulasForRow(project.rowNumber);
    forceApplyTerminalStatus_(masterSheet, project.rowNumber, phase, PHASE_COMPLETED, '🔵 เสร็จสิ้น', responsibleUnit, note, now);

    return {
      success: true,
      projectCode: project.projectCode,
      phase: phase,
      wiStep: PHASE_COMPLETED,
      responsibleUnit: responsibleUnit,
      imported: result.imported,
      failed: result.failed,
    };
  }

  function cancelProject(payload) {
    var adminUser = AuthApi.requireAdminSession(payload && payload.authToken);
    if (!payload || !payload.projectCode) throw new Error('กรุณาเลือกรหัสโครงการ');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) throw new Error('ไม่พบชีต Master Tracking');

    var project = getProjectByCode_(masterSheet, payload.projectCode);
    if (!project) throw new Error('ไม่พบรหัสโครงการ: ' + payload.projectCode);
    validatePhasePermission_(adminUser, project.phase);

    var now = new Date();
    var phase = String(project.phase || '');
    var responsibleUnit = project.responsibleUnit || MasterTrackingTools.getResponsibleUnitByPhase(phase);
    var note = payload.note || WI_CANCELLED;

    PhaseHistoryTools.recordClosedPhase(project, now, project.wiStep, note);
    appendTransferLog_(project, phase, WI_CANCELLED, responsibleUnit, note, adminUser.username, now);
    var result = AdminInputImportTools.importPendingAdminInputToMasterTracking();
    MasterTrackingTools.setMasterFormulasForRow(project.rowNumber);
    forceApplyTerminalStatus_(masterSheet, project.rowNumber, phase, WI_CANCELLED, '⚫ ยกเลิกรายการ', responsibleUnit, note, now);

    return {
      success: true,
      projectCode: project.projectCode,
      phase: phase,
      wiStep: WI_CANCELLED,
      responsibleUnit: responsibleUnit,
      imported: result.imported,
      failed: result.failed,
    };
  }

  function forceApplyTerminalStatus_(masterSheet, rowNumber, phase, wiStep, phaseStatus, responsibleUnit, note, now) {
    masterSheet.getRange(rowNumber, 9).setValue(phase);
    MasterTrackingTools.setWiValidationForRow(rowNumber, phase);
    masterSheet.getRange(rowNumber, 10).setValue(wiStep);
    masterSheet.getRange(rowNumber, 11).setValue(now);
    masterSheet.getRange(rowNumber, 14).setValue(phaseStatus);
    masterSheet.getRange(rowNumber, 15).setValue(responsibleUnit);
    // Freeze the final system duration at completion so it remains visible
    // after the live TODAY()-based formula is no longer applicable.
    if (phaseStatus === '🔵 เสร็จสิ้น') {
      var receivedDate = masterSheet.getRange(rowNumber, 4).getValue();
      masterSheet.getRange(rowNumber, 18).setValue(calculateInclusiveDays_(receivedDate, now));
    }
    masterSheet.getRange(rowNumber, 21).setValue(note || wiStep);
    masterSheet.getRange(rowNumber, 22).setValue(now);
  }

  function appendTransferLog_(project, phase, wiStep, responsibleUnit, note, recordedBy, now, phaseEntryDate, recordStatus) {
    AdminInputImportTools.appendAdminInputLog({
      projectCode: project.projectCode,
      documentNo: project.documentNo,
      ownerUnit: project.ownerUnit,
      receivedDate: project.receivedDate,
      projectName: project.projectName,
      budgetAmount: project.budgetAmount,
      budgetSource: project.budgetSource,
      fiscalYear: project.fiscalYear,
      phase: phase,
      wiStep: wiStep,
      phaseEntryDate: phaseEntryDate || now,
      responsibleUnit: responsibleUnit,
      note: note,
      recordedBy: recordedBy,
      recordedAt: now,
      recordStatus: recordStatus || AdminInputImportTools.PENDING_STATUS,
    });
  }

  function recordPhaseClosureIfNeeded_(project, targetPhase, now, note) {
    if (!project || String(project.phase || '') === String(targetPhase || '')) return null;
    return PhaseHistoryTools.recordClosedPhase(project, now, project.wiStep, note);
  }

  function isCrossUnitTransfer_(currentPhase, targetPhase) {
    return MasterTrackingTools.getResponsibleUnitByPhase(currentPhase) !==
      MasterTrackingTools.getResponsibleUnitByPhase(targetPhase);
  }

  function createPendingHandoff_(masterSheet, project, targetPhase, responsibleUnit, note, recordedBy, now) {
    appendTransferLog_(project, targetPhase, PENDING_HANDOFF_WI, responsibleUnit, note || PENDING_HANDOFF_WI, recordedBy, now, '', AdminInputImportTools.IMPORTED_STATUS);
    masterSheet.getRange(project.rowNumber, 6).setValue(project.budgetAmount || '');
    masterSheet.getRange(project.rowNumber, 7).setValue(project.budgetSource || '');
    masterSheet.getRange(project.rowNumber, 8).setValue(project.fiscalYear || '');
    masterSheet.getRange(project.rowNumber, 9).setValue(targetPhase);
    MasterTrackingTools.setWiValidationForRow(project.rowNumber, targetPhase);
    masterSheet.getRange(project.rowNumber, 10).setValue(PENDING_HANDOFF_WI);
    masterSheet.getRange(project.rowNumber, 11).clearContent();
    masterSheet.getRange(project.rowNumber, 12).setValue(MasterTrackingTools.getPhaseTotalSla(targetPhase));
    masterSheet.getRange(project.rowNumber, 13).clearContent();
    masterSheet.getRange(project.rowNumber, 14).setValue(PENDING_HANDOFF_STATUS);
    masterSheet.getRange(project.rowNumber, 15).setValue(responsibleUnit);
    masterSheet.getRange(project.rowNumber, 21).setValue(note || PENDING_HANDOFF_WI);
    masterSheet.getRange(project.rowNumber, 22).setValue(now);
  }

  function isPhase4BudgetApprovalStep_(phase, wiStep) {
    return String(phase || '').trim() === 'Phase 4' &&
      MasterTrackingTools.normalizeWiStep('Phase 4', wiStep) === '3.แจ้งผลการอนุมัติ(ใบขวาง) (User)';
  }

  function hasBudgetApprovalPayload_(payload) {
    return Boolean(payload && (
      String(payload.budgetSource || '').trim() ||
      String(payload.budgetAmount || '').trim() ||
      String(payload.fiscalYear || '').trim()
    ));
  }

  function validateBudgetApprovalPayloadIfNeeded_(payload, phase, wiStep) {
    if (!isPhase4BudgetApprovalStep_(phase, wiStep)) return;
    if (!String(payload && payload.budgetSource || '').trim()) throw new Error('กรุณากรอกแหล่งเงิน');
    if (!String(payload && payload.budgetAmount || '').trim()) throw new Error('กรุณากรอกวงเงิน');
    if (!String(payload && payload.fiscalYear || '').trim()) throw new Error('กรุณากรอกปีงบประมาณ');
  }

  function applyBudgetApprovalPayloadToProject_(project, payload, phase, wiStep) {
    if (!isPhase4BudgetApprovalStep_(phase, wiStep)) return project;
    var copy = {};
    Object.keys(project || {}).forEach(function(key) {
      copy[key] = project[key];
    });
    if (payload && Object.prototype.hasOwnProperty.call(payload, 'budgetAmount')) copy.budgetAmount = payload.budgetAmount;
    if (payload && Object.prototype.hasOwnProperty.call(payload, 'budgetSource')) copy.budgetSource = payload.budgetSource;
    if (payload && Object.prototype.hasOwnProperty.call(payload, 'fiscalYear')) copy.fiscalYear = payload.fiscalYear;
    return copy;
  }

  function getActiveProjectsForPhaseTransfer(authToken) {
    var adminUser = AuthApi.requireAdminSession(authToken);
    var allowedPhases = getAllowedTransferPhases_(adminUser);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet || masterSheet.getLastRow() < 2) return [];

    var rows = masterSheet
      .getRange(2, 1, masterSheet.getLastRow() - 1, MasterTrackingTools.MASTER_HEADERS.length)
      .getValues();

    return rows
      .filter(function(row) {
        var projectCode = String(row[0] || '');
        var currentPhase = String(row[8] || '');
        return /^A-\d+$/.test(projectCode) && allowedPhases.indexOf(currentPhase) !== -1;
      })
      .map(function(row) {
        return {
          projectCode: String(row[0] || ''),
          documentNo: String(row[1] || ''),
          ownerUnit: String(row[2] || ''),
          receivedDate: formatDateForClient_(row[3]),
          projectName: String(row[4] || ''),
          budgetAmount: row[5] || '',
          budgetSource: String(row[6] || ''),
          fiscalYear: String(row[7] || ''),
          phase: String(row[8] || ''),
          wiStep: String(row[9] || ''),
          phaseEntryDate: formatDateForClient_(row[10]),
          phaseSla: row[11] || '',
          phaseRemainingDays: row[12] || '',
          phaseSlaStatus: String(row[13] || ''),
          responsibleUnit: String(row[14] || ''),
          regulationDueDate: formatDateForClient_(row[15]),
          mainSla: row[16] || '',
          mainUsedDays: row[17] || '',
          mainRemainingDays: row[18] || '',
          mainSlaStatus: String(row[19] || ''),
          note: String(row[20] || ''),
          updatedAt: formatDateForClient_(row[21]),
        };
      });
  }

  function getProjectTimelineForPhaseTransfer(projectCode) {
    if (!projectCode) return [];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    if (!adminSheet || adminSheet.getLastRow() < 2) return [];

    var rows = adminSheet.getRange(2, 1, adminSheet.getLastRow() - 1, 16).getValues();
    return rows
      .map(function(row, index) {
        return {
          rowNumber: index + 2,
          projectCode: String(row[0] || ''),
          phase: String(row[8] || ''),
          wiStep: String(row[9] || ''),
          phaseEntryDate: formatDateForClient_(row[10]),
          responsibleUnit: String(row[11] || ''),
          note: String(row[12] || ''),
          recordedBy: String(row[13] || ''),
          recordedAt: formatDateTimeForClient_(row[14]),
          recordedAtMs: toTime_(row[14]),
          recordStatus: String(row[15] || ''),
          sortTime: toTime_(row[14]),
        };
      })
      .filter(function(log) {
        return log.projectCode === String(projectCode);
      })
      .sort(function(a, b) {
        if (a.sortTime !== b.sortTime) return b.sortTime - a.sortTime;
        return b.rowNumber - a.rowNumber;
      })
      .slice(0, 20)
      .map(function(log) {
        delete log.sortTime;
        return log;
      });
  }

  function getProjectTimelineBundle(projectCode) {
    return {
      timeline: getProjectTimelineForPhaseTransfer(projectCode),
      phaseHistory: PhaseHistoryTools.getProjectPhaseHistory(projectCode),
    };
  }

  function getProjectByCode_(sheet, projectCode) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    var rows = sheet.getRange(2, 1, lastRow - 1, MasterTrackingTools.MASTER_HEADERS.length).getValues();
    for (var i = 0; i < rows.length; i += 1) {
      var row = rows[i];
      if (String(row[0]) === String(projectCode)) {
        return {
          rowNumber: i + 2,
          projectCode: row[0],
          documentNo: row[1],
          ownerUnit: row[2],
          receivedDate: row[3],
          projectName: row[4],
          budgetAmount: row[5],
          budgetSource: row[6],
          fiscalYear: row[7],
          phase: row[8],
          wiStep: row[9],
          phaseEntryDate: row[10],
          phaseSla: row[11],
          phaseRemainingDays: row[12],
          phaseSlaStatus: row[13],
          responsibleUnit: row[14],
          regulationDueDate: row[15],
          mainSla: row[16],
          mainUsedDays: row[17],
          mainRemainingDays: row[18],
          mainSlaStatus: row[19],
          note: row[20],
          updatedAt: row[21],
        };
      }
    }
    return null;
  }

  function validatePayload_(payload) {
    if (!payload) throw new Error('ไม่พบข้อมูลส่งต่อ Phase');
    if (!payload.projectCode) throw new Error('กรุณาเลือกรหัสโครงการ');
    if (!payload.phase) throw new Error('กรุณาเลือก SOP / Phase');
    if (!payload.wiStep) throw new Error('กรุณาเลือก WI / ขั้นตอนย่อย');
    var wiList = MasterTrackingTools.getWiListByPhase(payload.phase);
    if (!wiList.length) {
      throw new Error('ไม่พบ WI สำหรับ Phase ที่เลือก');
    }
    payload.wiStep = MasterTrackingTools.normalizeWiStep(payload.phase, payload.wiStep);
    if (wiList.indexOf(payload.wiStep) === -1) {
      throw new Error('WI ที่เลือกไม่ตรงกับ Phase ที่เลือก');
    }
  }

  function validatePhasePermission_(adminUser, phase) {
    var allowedPhases = getAllowedTransferPhases_(adminUser);
    if (allowedPhases.indexOf(phase) === -1) {
      throw new Error('หน่วยงานของคุณไม่มีสิทธิ์อัปเดต ' + phase);
    }
  }

  function isProcurementAdmin_(adminUser) {
    return adminUser && adminUser.unit === UNIT_PROCUREMENT;
  }

  function getAllowedTransferPhases_(adminUser) {
    if (!adminUser) return [];
    if (adminUser.unit === UNIT_BUILDING) return ['Phase 1', 'Phase 2', 'Phase 3'];
    if (adminUser.unit === UNIT_PLAN) return ['Phase 4'];
    if (adminUser.unit === UNIT_PROCUREMENT) return ['Phase 5', 'Phase 6', 'Phase 7'];
    return [];
  }

  function getNextTransferTarget_(adminUser, currentPhase) {
    if (!adminUser) return '';
    if (adminUser.unit === UNIT_BUILDING && currentPhase === 'Phase 3') return 'Phase 4';
    if (adminUser.unit === UNIT_PLAN && currentPhase === 'Phase 4') return 'Phase 5';
    
    return '';
  }

  function getNextStepTarget_(project) {
    var phaseOrder = MasterTrackingTools.PHASES || [];
    var phase = String(project.phase || '').trim();
    var wiStep = MasterTrackingTools.normalizeWiStep(phase, project.wiStep);
    var wiList = MasterTrackingTools.getWiListByPhase(phase) || [];
    var wiIndex = wiList.indexOf(wiStep);

    if (wiIndex !== -1 && wiIndex < wiList.length - 1) {
      return {
        phase: phase,
        wiStep: wiList[wiIndex + 1],
      };
    }

    if (wiStep === PENDING_HANDOFF_WI) return null;

    var phaseIndex = phaseOrder.indexOf(phase);
    if (phaseIndex === -1 || phaseIndex >= phaseOrder.length - 1) return null;

    var nextPhase = phaseOrder[phaseIndex + 1];
    var nextWiList = MasterTrackingTools.getWiListByPhase(nextPhase) || [];
    return {
      phase: nextPhase,
      wiStep: nextWiList[0] || nextPhase,
    };
  }

  function validateNextStepPermission_(adminUser, currentPhase, targetPhase) {
    var allowed = getAllowedTransferPhases_(adminUser);
    if (allowed.indexOf(currentPhase) === -1) {
      throw new Error('คุณไม่มีสิทธิ์ส่งต่อรายการใน Phase นี้');
    }
    if (allowed.indexOf(targetPhase) !== -1) return;
    if (adminUser.unit === UNIT_BUILDING && currentPhase === 'Phase 3' && targetPhase === 'Phase 4') return;
    if (adminUser.unit === UNIT_PLAN && currentPhase === 'Phase 4' && targetPhase === 'Phase 5') return;
    throw new Error('คุณไม่มีสิทธิ์ส่งต่อไปยัง Phase ปลายทางนี้');
  }

  function getFirstPhaseByUnit_(unit) {
    if (unit === UNIT_BUILDING) return 'Phase 1';
    if (unit === UNIT_PLAN) return 'Phase 4';
    if (unit === UNIT_PROCUREMENT) return 'Phase 5';
    return '';
  }

  function formatDateForClient_(value) {
    if (!(value instanceof Date)) return value ? String(value) : '';
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  function formatDateTimeForClient_(value) {
    if (!(value instanceof Date)) return value ? String(value) : '';
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss.SSS');
  }

  function toTime_(value) {
    if (value instanceof Date) return value.getTime();
    var parsed = value ? new Date(value) : null;
    return parsed && !Number.isNaN(parsed.getTime()) ? parsed.getTime() : 0;
  }

  function calculateInclusiveDays_(startValue, endValue) {
    var start = startValue instanceof Date ? startValue : new Date(startValue);
    var end = endValue instanceof Date ? endValue : new Date(endValue);
    if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return 0;
    start = new Date(start.getFullYear(), start.getMonth(), start.getDate());
    end = new Date(end.getFullYear(), end.getMonth(), end.getDate());
    return Math.max(1, Math.floor((end.getTime() - start.getTime()) / 86400000) + 1);
  }

  return {
    VERSION: VERSION,
    transferProjectPhase: transferProjectPhase,
    transferProjectToNextUnit: transferProjectToNextUnit,
    saveProjectNoteOnly: saveProjectNoteOnly,
    transferProjectToNextStep: transferProjectToNextStep,
    transferProjectToUnit: transferProjectToUnit,
    acceptIncomingHandoff: acceptIncomingHandoff,
    completeProject: completeProject,
    cancelProject: cancelProject,
    getActiveProjectsForPhaseTransfer: getActiveProjectsForPhaseTransfer,
    getProjectTimelineForPhaseTransfer: getProjectTimelineForPhaseTransfer,
    getProjectTimelineBundle: getProjectTimelineBundle,
  };
})();

function transferProjectPhase(payload) {
  return PhaseTransferApi.transferProjectPhase(payload);
}

function getActiveProjectsForPhaseTransfer(authToken) {
  return PhaseTransferApi.getActiveProjectsForPhaseTransfer(authToken);
}

function getProjectTimelineForPhaseTransfer(projectCode) {
  return PhaseTransferApi.getProjectTimelineForPhaseTransfer(projectCode);
}

function getProjectTimelineBundle(projectCode) {
  return PhaseTransferApi.getProjectTimelineBundle(projectCode);
}

function transferProjectToNextUnit(payload) {
  return PhaseTransferApi.transferProjectToNextUnit(payload);
}

function saveProjectNoteOnly(payload) {
  return PhaseTransferApi.saveProjectNoteOnly(payload);
}

function transferProjectToNextStep(payload) {
  return PhaseTransferApi.transferProjectToNextStep(payload);
}

function transferProjectToUnit(payload) {
  return PhaseTransferApi.transferProjectToUnit(payload);
}

function acceptIncomingHandoff(payload) {
  return PhaseTransferApi.acceptIncomingHandoff(payload);
}

function completeProject(payload) {
  return PhaseTransferApi.completeProject(payload);
}

function cancelProject(payload) {
  return PhaseTransferApi.cancelProject(payload);
}
