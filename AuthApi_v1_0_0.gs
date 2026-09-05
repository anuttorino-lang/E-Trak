var AuthApi = (function() {
  var VERSION = '1.0.1';
  var ADMIN_USERS_SHEET_NAME = 'Admin_Users';
  var SESSION_PREFIX = 'ADMIN_SESSION_';
  var SESSION_TTL_SECONDS = 21600;

  var ADMIN_USER_HEADERS = [
    'ชื่อผู้ใช้',
    'รหัสผ่าน Hash',
    'ชื่อแสดงผล',
    'หน่วยงาน',
    'บทบาท',
    'สถานะใช้งาน',
    'อัปเดตล่าสุด',
  ];

  var DEFAULT_USERS = [
    {
      username: 'building',
      password: '1234',
      displayName: 'แอดมินงานอาคารสถานที่',
      unit: 'งานอาคารสถานที่',
      role: 'admin',
    },
    {
      username: 'plan',
      password: '1234',
      displayName: 'แอดมินงานนโยบายและแผน',
      unit: 'งานนโยบายและแผน',
      role: 'admin',
    },
    {
      username: 'procurement',
      password: '1234',
      displayName: 'แอดมินงานพัสดุและยานพาหนะ',
      unit: 'งานพัสดุและยานพาหนะ',
      role: 'admin',
    },
  ];

  function setupAdminAuthUsers() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ADMIN_USERS_SHEET_NAME) || ss.insertSheet(ADMIN_USERS_SHEET_NAME);

    ensureShape_(sheet);
    clearOldValidations_(sheet);
    sheet.getRange(1, 1, 1, ADMIN_USER_HEADERS.length).setValues([ADMIN_USER_HEADERS]);
    formatSheet_(sheet);
    appendMissingDefaultUsers_(sheet);
    repairDefaultUsers_(sheet);

    return {
      success: true,
      version: VERSION,
      sheetName: ADMIN_USERS_SHEET_NAME,
      defaultPassword: '1234',
      users: DEFAULT_USERS.map(function(user) {
        return {
          username: user.username,
          displayName: user.displayName,
          unit: user.unit,
        };
      }),
    };
  }

  function loginAdmin(payload) {
    if (!payload || !payload.username || !payload.password) {
      throw new Error('กรุณากรอกชื่อผู้ใช้และรหัสผ่าน');
    }

    setupAdminAuthUsers();
    var user = findActiveUser_(normalizeLoginUsername_(payload.username));
    if (!user || user.passwordHash !== hashPassword_(payload.password)) {
      throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
    }

    var token = Utilities.getUuid();
    var session = {
      username: user.username,
      displayName: user.displayName,
      unit: user.unit,
      role: user.role,
      loginAt: new Date().toISOString(),
    };

    CacheService.getScriptCache().put(SESSION_PREFIX + token, JSON.stringify(session), SESSION_TTL_SECONDS);

    return {
      success: true,
      token: token,
      user: session,
      expiresInSeconds: SESSION_TTL_SECONDS,
    };
  }

  function validateAdminSession(token) {
    if (!token) {
      return {
        success: false,
        authenticated: false,
      };
    }

    var cached = CacheService.getScriptCache().get(SESSION_PREFIX + token);
    if (!cached) {
      return {
        success: false,
        authenticated: false,
      };
    }

    return {
      success: true,
      authenticated: true,
      user: JSON.parse(cached),
    };
  }

  function requireAdminSession(token) {
    var session = validateAdminSession(token);
    if (!session.authenticated) {
      throw new Error('กรุณาเข้าสู่ระบบแอดมินก่อนใช้งาน');
    }
    return session.user;
  }

  function logoutAdmin(token) {
    if (token) {
      CacheService.getScriptCache().remove(SESSION_PREFIX + token);
    }
    return {
      success: true,
    };
  }

  function setAdminPassword(username, newPassword) {
    if (!username || !newPassword) {
      throw new Error('กรุณาระบุ username และ newPassword');
    }

    setupAdminAuthUsers();
    username = normalizeLoginUsername_(username);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMIN_USERS_SHEET_NAME);
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, ADMIN_USER_HEADERS.length).getValues();

    for (var i = 0; i < rows.length; i += 1) {
      if (String(rows[i][0] || '') === String(username)) {
        var row = i + 2;
        sheet.getRange(row, 2).setValue(hashPassword_(newPassword));
        sheet.getRange(row, 7).setValue(new Date());

        return {
          success: true,
          username: username,
        };
      }
    }

    throw new Error('ไม่พบชื่อผู้ใช้: ' + username);
  }

  function findActiveUser_(username) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMIN_USERS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      return null;
    }

    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, ADMIN_USER_HEADERS.length).getValues();

    for (var i = 0; i < rows.length; i += 1) {
      if (String(rows[i][0] || '').trim() === String(username || '').trim()) {
        return {
          username: String(rows[i][0] || ''),
          passwordHash: String(rows[i][1] || ''),
          displayName: String(rows[i][2] || ''),
          unit: String(rows[i][3] || ''),
          role: String(rows[i][4] || ''),
          active: String(rows[i][5] || '') === 'ใช้งาน',
        };
      }
    }

    return null;
  }

  function normalizeLoginUsername_(username) {
    var text = String(username || '').trim().toLowerCase();
    var aliases = {
      'building_admin': 'building',
      'anuttorino@gmail.com': 'building',
      'plan_admin': 'plan',
      'gydeview@gmail.com': 'plan',
      'procurement_admin': 'procurement',
      'gyde002@gmail.com': 'procurement',
    };

    return aliases[text] || text;
  }

  function appendMissingDefaultUsers_(sheet) {
    var existing = {};
    if (sheet.getLastRow() >= 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(function(row) {
        if (row[0]) {
          existing[String(row[0])] = true;
        }
      });
    }

    DEFAULT_USERS.forEach(function(user) {
      if (existing[user.username]) {
        return;
      }

      sheet.appendRow([
        user.username,
        hashPassword_(user.password),
        user.displayName,
        user.unit,
        user.role,
        'ใช้งาน',
        new Date(),
      ]);
    });
  }

  function repairDefaultUsers_(sheet) {
    if (sheet.getLastRow() < 2) {
      return;
    }

    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, ADMIN_USER_HEADERS.length).getValues();
    var defaultsByUsername = {};

    DEFAULT_USERS.forEach(function(user) {
      defaultsByUsername[user.username] = user;
    });

    rows.forEach(function(row, index) {
      var user = defaultsByUsername[String(row[0] || '')];
      if (!user) {
        return;
      }

      var rowNumber = index + 2;
      sheet.getRange(rowNumber, 3, 1, 5).setValues([[
        user.displayName,
        user.unit,
        user.role,
        'ใช้งาน',
        new Date(),
      ]]);
    });
  }

  function ensureShape_(sheet) {
    if (sheet.getMaxColumns() < ADMIN_USER_HEADERS.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), ADMIN_USER_HEADERS.length - sheet.getMaxColumns());
    }
    if (sheet.getMaxRows() < 20) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 20 - sheet.getMaxRows());
    }
  }

  function clearOldValidations_(sheet) {
    sheet.getRange(1, 1, sheet.getMaxRows(), ADMIN_USER_HEADERS.length).clearDataValidations();
  }

  function formatSheet_(sheet) {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, ADMIN_USER_HEADERS.length)
      .setFontWeight('bold')
      .setFontColor('#ffffff')
      .setBackground('#075f6c')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    sheet.getRange(2, 7, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('dd/mm/yyyy hh:mm');
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 220);
    sheet.setColumnWidth(4, 220);
    sheet.setColumnWidth(5, 110);
    sheet.setColumnWidth(6, 110);
    sheet.setColumnWidth(7, 150);
  }

  function hashPassword_(password) {
    var bytes = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      String(password || ''),
      Utilities.Charset.UTF_8
    );

    return bytes.map(function(byte) {
      var value = byte < 0 ? byte + 256 : byte;
      return ('0' + value.toString(16)).slice(-2);
    }).join('');
  }

  return {
    VERSION: VERSION,
    setupAdminAuthUsers: setupAdminAuthUsers,
    loginAdmin: loginAdmin,
    validateAdminSession: validateAdminSession,
    requireAdminSession: requireAdminSession,
    logoutAdmin: logoutAdmin,
    setAdminPassword: setAdminPassword,
  };
})();

function setupAdminAuthUsers() {
  return AuthApi.setupAdminAuthUsers();
}

function loginAdmin(payload) {
  return AuthApi.loginAdmin(payload);
}

function validateAdminSession(token) {
  return AuthApi.validateAdminSession(token);
}

function requireAdminSession(token) {
  return AuthApi.requireAdminSession(token);
}

function logoutAdmin(token) {
  return AuthApi.logoutAdmin(token);
}

function setAdminPassword(username, newPassword) {
  return AuthApi.setAdminPassword(username, newPassword);
}
