var ProjectCreateApi = (function() {
  var VERSION = '1.0.2';
  var UNIT_BUILDING = 'งานอาคารสถานที่';
  var UNIT_PLAN = 'งานนโยบายและแผน';
  var UNIT_PROCUREMENT = 'งานพัสดุและยานพาหนะ';
function createNewProjectFromBuilding(payload) {
    var adminUser = AuthApi.requireAdminSession(payload && payload.authToken);
    requireBuildingAdmin_(adminUser);
    validatePayload_(payload);

    var now = new Date();
    var projectCode = MasterTrackingTools.generateNextProjectCode();
    var receivedDate = toDateOrDefault_(payload.receivedDate, now);
    var phase = 'Phase 1';
    var wiStep = MasterTrackingTools.getWiListByPhase(phase)[0];
    var responsibleUnit = MasterTrackingTools.getResponsibleUnitByPhase(phase);
    validatePhaseAndWi_(phase, wiStep);

    AdminInputImportTools.appendAdminInputLog({
      projectCode: projectCode,
      documentNo: payload.documentNo,
      ownerUnit: payload.ownerUnit,
      receivedDate: receivedDate,
      projectName: payload.projectName,
      budgetAmount: payload.budgetAmount,
      budgetSource: payload.budgetSource,
      fiscalYear: payload.fiscalYear,
      phase: phase,
      wiStep: wiStep,
      phaseEntryDate: receivedDate,
      responsibleUnit: responsibleUnit,
      note: payload.note,
      recordedBy: adminUser.username,
      recordedAt: now,
      recordStatus: AdminInputImportTools.PENDING_STATUS,
    });

    var result = AdminInputImportTools.importPendingAdminInputToMasterTracking();
    fillMainSlaInMaster_(projectCode);

    return {
      success: true,
      projectCode: projectCode,
      imported: result.imported,
      failed: result.failed,
    };
  }

  function getProjectCreateOptions(authToken) {
    var adminUser = AuthApi.requireAdminSession(authToken);
    return {
      success: true,
      version: VERSION,
      canCreateProject: isBuildingAdmin_(adminUser),
      canCompleteProject: isProcurementAdmin_(adminUser),
      createPhases: ['Phase 1'],
      transferPhases: getAllowedTransferPhases_(adminUser),
      transferUnits: [UNIT_BUILDING, UNIT_PLAN, UNIT_PROCUREMENT],
      nextPhaseByCurrentPhase: getNextPhaseByCurrentPhase_(adminUser),
      wiByPhase: MasterTrackingTools.WI_BY_PHASE,
      responsibleUnitByPhase: MasterTrackingTools.RESPONSIBLE_UNIT_BY_PHASE,
      user: adminUser,
    };
  }

  function requireBuildingAdmin_(adminUser) {
    if (!isBuildingAdmin_(adminUser)) {
      throw new Error('เฉพาะงานอาคารสถานที่เท่านั้นที่เพิ่มโครงการใหม่ได้');
    }
  }

  function isBuildingAdmin_(adminUser) {
    return adminUser && adminUser.unit === UNIT_BUILDING;
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

  function getNextPhaseByCurrentPhase_(adminUser) {
    if (!adminUser) return {};
    if (adminUser.unit === UNIT_BUILDING) return { 'Phase 3': 'Phase 4' };
    if (adminUser.unit === UNIT_PLAN) return { 'Phase 4': 'Phase 5' };
    if (adminUser.unit === UNIT_PROCUREMENT) return {};
    return {};
  }

  function fillMainSlaInMaster_(projectCode) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(MasterTrackingTools.MASTER_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return;

    var codes = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < codes.length; i += 1) {
      if (String(codes[i][0]) === String(projectCode)) {
        var row = i + 2;
        if (!sheet.getRange(row, 17).getValue()) {
          sheet.getRange(row, 17).setValue(MasterTrackingTools.DEFAULT_MAIN_SLA_DAYS);
        }
        MasterTrackingTools.setMasterFormulasForRow(row);
        return;
      }
    }
  }

  function validatePayload_(payload) {
    if (!payload) throw new Error('ไม่พบข้อมูลโครงการ');
    if (!payload.documentNo) throw new Error('กรุณากรอกเลขที่หนังสือ');
    if (!payload.ownerUnit) throw new Error('กรุณากรอกหน่วยงานเจ้าของเรื่อง / ต้นเรื่อง');
    if (!payload.receivedDate) throw new Error('กรุณากรอกวันที่รับเรื่อง');
    if (!payload.projectName) throw new Error('กรุณากรอกชื่อโครงการ / รายการ');
  }

  function validatePhaseAndWi_(phase, wiStep) {
    if (MasterTrackingTools.getPhaseList().indexOf(phase) === -1) {
      throw new Error('SOP / Phase ไม่ถูกต้อง');
    }
    if (MasterTrackingTools.getWiListByPhase(phase).indexOf(wiStep) === -1) {
      throw new Error('WI ไม่ตรงกับ Phase ที่เลือก');
    }
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
    return fallback instanceof Date ? fallback : new Date(fallback);
  }

  return {
    VERSION: VERSION,
    createNewProjectFromBuilding: createNewProjectFromBuilding,
    getProjectCreateOptions: getProjectCreateOptions,
  };
})();

function createNewProjectFromBuilding(payload) {
  return ProjectCreateApi.createNewProjectFromBuilding(payload);
}

function getProjectCreateOptions(authToken) {
  return ProjectCreateApi.getProjectCreateOptions(authToken);
}
