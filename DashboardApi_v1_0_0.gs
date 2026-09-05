var DashboardApi = (function() {
  var VERSION = '1.0.4';
  var MASTER_SHEET_NAME = 'Master Tracking';

  var STATUS_NORMAL = '🟢 ปกติ';
  var STATUS_NEAR_DUE = '🟡 ใกล้กำหนด';
  var STATUS_DELAYED = '🔴 ล่าช้า';
  var STATUS_REJECTED = '❌ ไม่อนุมัติ';
  var STATUS_CANCELLED = '⚫ ยกเลิกรายการ';
  var STATUS_COMPLETED = '🔵 เสร็จสิ้น';
  var FILTER_ALL = 'ทั้งหมด';

  function getDashboardData(filters) {
    var projects = readMasterProjects_();
    var normalizedFilters = normalizeFilters_(filters || {});
    var filteredProjects = applyFilters_(projects, normalizedFilters);

    return {
      version: VERSION,
      generatedAt: formatDateTimeForClient_(new Date()),
      filters: buildFilterOptions_(projects),
      summary: buildSummary_(filteredProjects),
      phaseSummary: buildCountSummary_(filteredProjects, 'phase'),
      unitSummary: buildCountSummary_(filteredProjects, 'responsibleUnit'),
      unitWorkloadSummary: buildUnitWorkloadSummary_(filteredProjects),
      statusSummary: buildStatusSummary_(filteredProjects),
      actionRequired: buildActionRequired_(filteredProjects),
      oldestProjects: buildOldestProjects_(filteredProjects),
      projects: filteredProjects,
    };
  }

  function readMasterProjects_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }

    var rows = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, MasterTrackingTools.MASTER_HEADERS.length)
      .getValues();

    return rows
      .filter(function(row) {
        return /^(A-\d+|A-TEST-\d+)$/.test(String(row[0] || ''));
      })
      .map(function(row) {
        return {
          projectCode: String(row[0] || ''),
          documentNo: String(row[1] || ''),
          ownerUnit: String(row[2] || ''),
          receivedDate: formatDateForClient_(row[3]),
          projectName: String(row[4] || ''),
          budgetAmount: toNumber_(row[5]),
          budgetSource: String(row[6] || ''),
          fiscalYear: String(row[7] || ''),
          phase: String(row[8] || ''),
          wiStep: String(row[9] || ''),
          phaseEntryDate: formatDateForClient_(row[10]),
          phaseSla: toNumber_(row[11]),
          phaseRemainingDays: toNumberOrBlank_(row[12]),
          phaseSlaStatus: String(row[13] || ''),
          responsibleUnit: String(row[14] || ''),
          regulationDueDate: formatDateForClient_(row[15]),
          mainSla: toNumber_(row[16]),
          mainUsedDays: toNumberOrBlank_(row[17]),
          note: String(row[20] || ''),
          updatedAt: formatDateTimeForClient_(row[21]),
        };
      });
  }

  function buildSummary_(projects) {
    return {
      totalProjects: projects.length,
      normal: countByField_(projects, 'phaseSlaStatus', STATUS_NORMAL),
      nearDue: countByField_(projects, 'phaseSlaStatus', STATUS_NEAR_DUE),
      delayed: countByField_(projects, 'phaseSlaStatus', STATUS_DELAYED),
      rejected: countByField_(projects, 'phaseSlaStatus', STATUS_REJECTED),
      cancelled: countByField_(projects, 'phaseSlaStatus', STATUS_CANCELLED),
      completed: countByField_(projects, 'phaseSlaStatus', STATUS_COMPLETED),
      totalBudget: projects.reduce(function(total, project) {
        var status = String(project.phaseSlaStatus || '');
        if (status === STATUS_REJECTED || status === STATUS_CANCELLED) return total;
        return total + toNumber_(project.budgetAmount);
      }, 0),
    };
  }

  function buildStatusSummary_(projects) {
    return getStatusList_().map(function(status) {
      return {
        name: status,
        phaseCount: countByField_(projects, 'phaseSlaStatus', status),
        count: countByField_(projects, 'phaseSlaStatus', status),
      };
    });
  }

  function buildCountSummary_(projects, field) {
    var counts = {};
    projects.forEach(function(project) {
      var value = project[field] || '-';
      counts[value] = (counts[value] || 0) + 1;
    });

    return Object.keys(counts).sort().map(function(name) {
      return {
        name: name,
        count: counts[name],
      };
    });
  }

  function buildUnitWorkloadSummary_(projects) {
    var units = [
      'งานอาคารสถานที่',
      'งานนโยบายและแผน',
      'งานพัสดุและยานพาหนะ',
      '-',
    ];

    return units.map(function(unit) {
      var unitProjects = projects.filter(function(project) {
        return (project.responsibleUnit || '-') === unit;
      });

      return {
        name: unit,
        total: unitProjects.length,
        normal: countByField_(unitProjects, 'phaseSlaStatus', STATUS_NORMAL),
        nearDue: countByField_(unitProjects, 'phaseSlaStatus', STATUS_NEAR_DUE),
        delayed: countByField_(unitProjects, 'phaseSlaStatus', STATUS_DELAYED),
        rejected: countByField_(unitProjects, 'phaseSlaStatus', STATUS_REJECTED),
        cancelled: countByField_(unitProjects, 'phaseSlaStatus', STATUS_CANCELLED),
        completed: countByField_(unitProjects, 'phaseSlaStatus', STATUS_COMPLETED),
      };
    }).filter(function(row) {
      return row.total > 0;
    });
  }

  function buildActionRequired_(projects) {
    return projects
      .filter(function(project) {
        if (isClosedProject_(project)) {
          return false;
        }
        return project.phaseSlaStatus === STATUS_DELAYED ||
          project.phaseSlaStatus === STATUS_NEAR_DUE ||
          project.note;
      })
      .sort(function(a, b) {
        return priorityScore_(b) - priorityScore_(a);
      })
      .slice(0, 50);
  }

  function buildOldestProjects_(projects) {
    return projects
      .filter(function(project) {
        return !isClosedProject_(project) && Number.isFinite(Number(project.mainUsedDays));
      })
      .sort(function(a, b) {
        return Number(b.mainUsedDays) - Number(a.mainUsedDays);
      })
      .slice(0, 10);
  }

  function isClosedProject_(project) {
    var status = String(project.phaseSlaStatus || '');
    var wiStep = String(project.wiStep || '');
    return status === STATUS_CANCELLED ||
      status === STATUS_COMPLETED ||
      status === STATUS_REJECTED ||
      wiStep.indexOf('ไม่อนุมัติ') !== -1 ||
      wiStep.indexOf('ยกเลิก') !== -1;
  }

  function priorityScore_(project) {
    var score = 0;
    if (project.phaseSlaStatus === STATUS_DELAYED) score += 100;
    if (project.phaseSlaStatus === STATUS_NEAR_DUE) score += 50;
    if (project.note) score += 10;
    if (typeof project.phaseRemainingDays === 'number') score -= project.phaseRemainingDays;
    return score;
  }

  function buildFilterOptions_(projects) {
    return {
      fiscalYears: uniqueValues_(projects, 'fiscalYear'),
      phases: mergeOrderedValues_(getStandardPhaseList_(), uniqueValues_(projects, 'phase')),
      responsibleUnits: mergeOrderedValues_(getStandardResponsibleUnitList_(), uniqueValues_(projects, 'responsibleUnit')),
      statuses: [FILTER_ALL].concat(getStatusList_()),
    };
  }

  function getStandardPhaseList_() {
    return MasterTrackingTools.getPhaseList();
  }

  function getStandardResponsibleUnitList_() {
    var phaseMap = MasterTrackingTools.RESPONSIBLE_UNIT_BY_PHASE || {};
    return mergeOrderedValues_([], Object.keys(phaseMap).map(function(phase) {
      return phaseMap[phase];
    }));
  }

  function mergeOrderedValues_(baseValues, extraValues) {
    var seen = {};
    var merged = [];

    (baseValues || []).concat(extraValues || []).forEach(function(value) {
      value = String(value || '').trim();
      if (!value || seen[value]) return;
      seen[value] = true;
      merged.push(value);
    });

    return merged;
  }

  function applyFilters_(projects, filters) {
    return projects.filter(function(project) {
      if (filters.keyword && !matchesKeyword_(project, filters.keyword)) return false;
      if (filters.fiscalYear && project.fiscalYear !== filters.fiscalYear) return false;
      if (filters.phase && project.phase !== filters.phase) return false;
      if (filters.responsibleUnit && project.responsibleUnit !== filters.responsibleUnit) return false;
      if (filters.status && filters.status !== FILTER_ALL) {
        return project.phaseSlaStatus === filters.status;
      }
      return true;
    });
  }

  function normalizeFilters_(filters) {
    return {
      keyword: normalizeText_(filters.keyword || ''),
      fiscalYear: filters.fiscalYear || '',
      phase: filters.phase || '',
      responsibleUnit: filters.responsibleUnit || '',
      status: filters.status || '',
    };
  }

  function matchesKeyword_(project, keyword) {
    var haystack = [
      project.projectCode,
      project.documentNo,
      project.ownerUnit,
      project.projectName,
      project.budgetSource,
      project.fiscalYear,
      project.phase,
      project.wiStep,
      project.responsibleUnit,
      project.phaseSlaStatus,
      project.note,
    ].join(' ');

    return normalizeText_(haystack).indexOf(keyword) !== -1;
  }

  function normalizeText_(value) {
    return String(value || '').toLowerCase().trim();
  }

  function uniqueValues_(projects, field) {
    var seen = {};
    projects.forEach(function(project) {
      if (project[field]) {
        seen[project[field]] = true;
      }
    });
    return Object.keys(seen).sort();
  }

  function countByField_(projects, field, value) {
    return projects.filter(function(project) {
      return project[field] === value;
    }).length;
  }

  function getStatusList_() {
    return [
      STATUS_NORMAL,
      STATUS_NEAR_DUE,
      STATUS_DELAYED,
      STATUS_REJECTED,
      STATUS_CANCELLED,
      STATUS_COMPLETED,
    ];
  }

  function toNumber_(value) {
    var numberValue = Number(value);
    return Number.isFinite(numberValue) ? numberValue : 0;
  }

  function toNumberOrBlank_(value) {
    if (value === '' || value === null || value === undefined) {
      return '';
    }
    var numberValue = Number(value);
    return Number.isFinite(numberValue) ? numberValue : '';
  }

  function formatDateForClient_(value) {
    if (!(value instanceof Date)) {
      return value ? String(value) : '';
    }
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  function formatDateTimeForClient_(value) {
    if (!(value instanceof Date)) {
      return value ? String(value) : '';
    }
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  }

  return {
    VERSION: VERSION,
    getDashboardData: getDashboardData,
  };
})();

function getDashboardData(filters) {
  return DashboardApi.getDashboardData(filters);
}
