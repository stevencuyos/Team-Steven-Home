// ============================================================
// Utilities.gs — Shared helpers for Play Ops Store
// Enhanced with error handling, validation, and monitoring
// ============================================================

// ============================================================
// ERROR HANDLING & RECOVERY
// ============================================================

/**
 * Global error handler wrapper for safe execution
 * @param {Function} fn - Function to execute
 * @param {Object} context - Context to execute in
 * @param {string} errorContext - Error context for logging
 */
function executeWithErrorHandling(fn, context, errorContext) {
  try {
    return fn.call(context || this);
  } catch(e) {
    Logger.log('[ERROR] ' + errorContext + ': ' + e.message);
    Logger.log('[STACK] ' + e.stack);
    logErrorToSheet({
      error: e.message,
      stack: e.stack,
      context: errorContext,
      timestamp: new Date(),
      user: getCurrentLdap()
    });
    throw e;
  }
}

/**
 * Retry logic with exponential backoff
 * @param {Function} fn - Function to retry
 * @param {number} maxRetries - Maximum number of attempts
 * @param {number} initialDelayMs - Initial delay in milliseconds
 */
function executeWithRetry(fn, maxRetries, initialDelayMs) {
  maxRetries = maxRetries || 3;
  initialDelayMs = initialDelayMs || 1000;
  
  var delayMs = initialDelayMs;
  
  for (var i = 0; i < maxRetries; i++) {
    try {
      return fn();
    } catch(error) {
      if (i === maxRetries - 1) {
        Logger.log('[RETRY] Final attempt failed: ' + error.message);
        throw error;
      }
      
      Logger.log('[RETRY] Attempt ' + (i + 1) + '/' + maxRetries + ' failed. Retrying in ' + delayMs + 'ms...');
      Utilities.sleep(delayMs);
      delayMs *= 2; // Exponential backoff
    }
  }
}

/**
 * Log errors to ErrorLog sheet for monitoring
 */
function logErrorToSheet(errorData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var errorSheet = ss.getSheetByName('ErrorLog');
    
    if (!errorSheet) {
      errorSheet = ss.insertSheet('ErrorLog');
      errorSheet.appendRow(['Timestamp', 'Error', 'Stack', 'Context', 'LDAP', 'View']);
      errorSheet.setFrozenRows(1);
    }
    
    if (errorSheet.getLastRow() < 1000) { // Prevent unbounded growth
      errorSheet.appendRow([
        errorData.timestamp || new Date(),
        errorData.error || 'Unknown',
        errorData.stack || '',
        errorData.context || '',
        errorData.user || 'unknown',
        errorData.view || ''
      ]);
    }
  } catch(e) {
    Logger.log('[ERROR LOG FAILED] ' + e.message);
  }
}

// ============================================================
// AUDIT LOGGING
// ============================================================

/**
 * Log important actions for audit trail
 */
function auditLog(action, details, ldap) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var auditSheet = ss.getSheetByName('AuditLog');
    
    if (!auditSheet) {
      auditSheet = ss.insertSheet('AuditLog');
      auditSheet.appendRow(['Timestamp', 'Action', 'LDAP', 'Details', 'IPAddress']);
      auditSheet.setFrozenRows(1);
    }
    
    if (auditSheet.getLastRow() < 5000) { // Prevent unbounded growth
      auditSheet.appendRow([
        new Date().toISOString(),
        action,
        ldap || getCurrentLdap(),
        JSON.stringify(details),
        Session.getActiveUser().getEmail()
      ]);
    }
    
    Logger.log('[AUDIT] ' + action + ' by ' + (ldap || getCurrentLdap()));
  } catch(e) {
    Logger.log('[AUDIT LOG FAILED] ' + e.message);
  }
}

// ============================================================
// VALIDATION HELPERS
// ============================================================

/**
 * Validate task input data
 */
function validateTaskInput(taskData) {
  var errors = [];
  
  if (!taskData.title || taskData.title.trim().length < 3) {
    errors.push('Title must be at least 3 characters');
  }
  
  if (!taskData.category) {
    errors.push('Category is required');
  }
  
  if (isNaN(taskData.basePoints) || parseInt(taskData.basePoints) < 1) {
    errors.push('Points must be a positive number');
  }
  
  if (taskData.deadline && isNaN(new Date(taskData.deadline).getTime())) {
    errors.push('Invalid deadline format');
  }
  
  if (errors.length > 0) {
    throw new Error(errors.join('; '));
  }
  
  return true;
}

/**
 * Validate kudos submission
 */
function validateKudosInput(kudosData) {
  var errors = [];
  
  if (!kudosData.ldap || kudosData.ldap.trim().length === 0) {
    errors.push('Agent LDAP is required');
  }
  
  if (!kudosData.whyKudos || kudosData.whyKudos.trim().length < 10) {
    errors.push('Kudos reason must be at least 10 characters');
  }
  
  if (errors.length > 0) {
    throw new Error(errors.join('; '));
  }
  
  return true;
}

/**
 * Sanitize HTML input to prevent XSS
 */
function sanitizeInput(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ============================================================
// PERFORMANCE MONITORING
// ============================================================

/**
 * Track function execution time
 */
function measurePerformance(functionName, fn) {
  var startTime = new Date().getTime();
  var result = fn();
  var duration = new Date().getTime() - startTime;
  
  Logger.log('[PERF] ' + functionName + ': ' + duration + 'ms');
  
  // Store in cache for monitoring
  var perfKey = 'perf_' + functionName;
  var perfData = getCached(perfKey) || { times: [], count: 0 };
  perfData.times.push(duration);
  perfData.count++;
  perfData.avgTime = perfData.times.reduce(function(a, b) { return a + b; }) / perfData.times.length;
  
  // Keep only last 100 measurements
  if (perfData.times.length > 100) {
    perfData.times.shift();
  }
  
  setCached(perfKey, perfData, 3600);
  
  return result;
}

/**
 * Get performance metrics for monitoring
 */
function getPerformanceMetrics(functionName) {
  return getCached('perf_' + functionName) || null;
}

// ============================================================
// RATE LIMITING
// ============================================================

/**
 * Simple rate limiter to prevent abuse
 */
var RateLimiter = {
  limits: {},
  
  isAllowed: function(key, maxRequests, windowMs) {
    maxRequests = maxRequests || 50;
    windowMs = windowMs || 60000;
    
    var now = Date.now();
    if (!this.limits[key]) {
      this.limits[key] = [];
    }
    
    var recentRequests = this.limits[key].filter(function(time) {
      return now - time < windowMs;
    });
    
    if (recentRequests.length >= maxRequests) {
      return false;
    }
    
    recentRequests.push(now);
    this.limits[key] = recentRequests;
    return true;
  },
  
  reset: function(key) {
    delete this.limits[key];
  }
};

// ============================================================
// SHEET ACCESS
// ============================================================

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet not found: ' + name);
  return sheet;
}

var _sheetDataCache = {};

function getSheetData(name) {
  if (_sheetDataCache[name]) return _sheetDataCache[name];

  var sheet = getSheet(name);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    _sheetDataCache[name] = [];
    return [];
  }
  var headers = data[0];
  var result = data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });

  _sheetDataCache[name] = result;
  return result;
}

function clearSheetDataCache(name) {
  if (name) {
    delete _sheetDataCache[name];
  } else {
    _sheetDataCache = {};
  }
}

var _configCache = null;

function getConfig(setting) {
  if (_configCache && _configCache[setting] !== undefined) {
    return _configCache[setting];
  }

  // Try script cache
  var cached = getCached('app_config');
  if (cached) {
    _configCache = cached;
    if (_configCache[setting] !== undefined) return _configCache[setting];
  }

  // Load from sheet
  var rows = getSheetData('Config');
  var config = {};
  for (var i = 0; i < rows.length; i++) {
    config[rows[i]['Setting']] = rows[i]['Value'];
  }

  _configCache = config;
  setCached('app_config', config, 1800); // 30 min cache

  return config[setting] !== undefined ? config[setting] : null;
}

// ============================================================
// ID GENERATION
// ============================================================

function generateTaskId() {
  return generateId('Tasks', 'TASK');
}

function generateCompletionId() {
  return generateId('Completions', 'COMP');
}

function generateKudosId() {
  return generateId('Kudos', 'KUD');
}

function generateDemeritId() {
  return generateId('Demerits', 'DEM');
}

function generateNotificationId() {
  return generateId('Notifications', 'NOTIF');
}

function generateId(sheetName, prefix) {
  return generateIds(sheetName, prefix, 1)[0];
}

function generateIds(sheetName, prefix, count) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = getSheet(sheetName);
    var lastRow = sheet.getLastRow();
    var startNum = lastRow < 2 ? 1 : lastRow;
    var timestamp = new Date().getTime().toString().slice(-4);

    var ids = [];
    for (var i = 0; i < count; i++) {
      var padded = String(startNum + i).padStart(4, '0');
      ids.push(prefix + '-' + padded + '-' + timestamp);
    }
    return ids;
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// DATE & TIME HELPERS
// ============================================================

function now() {
  return new Date();
}

function formatDate(date) {
  if (!date) return '';
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'MMM dd, yyyy');
}

function formatDateTime(date) {
  if (!date) return '';
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'MMM dd, yyyy hh:mm a');
}

function isExpired(deadline) {
  if (!deadline) return false;
  return new Date(deadline) < now();
}

function hoursUntil(deadline) {
  var diff = new Date(deadline) - now();
  return diff / (1000 * 60 * 60);
}

function minutesToHours(minutes) {
  return Math.round(minutes / 60 * 10) / 10;
}

// Calculate TAT in hours between two dates
function calcTAT(claimedAt, completedAt) {
  if (!claimedAt || !completedAt) return null;
  var diff = new Date(completedAt) - new Date(claimedAt);
  return Math.round(diff / (1000 * 60 * 60) * 10) / 10; // hours, 1 decimal
}

// Days before deadline the task was completed (excluding weekends)
function daysEarly(completedAt, deadline) {
  if (!completedAt || !deadline) return 0;
  var start = new Date(completedAt);
  var end = new Date(deadline);
  if (start >= end) return 0;

  var businessDays = 0;
  var current = new Date(start);
  while (current < end) {
    current.setDate(current.getDate() + 1);
    var dayOfWeek = current.getDay();
    if (dayOfWeek !== 0 && dayOfWeek !== 6) { // 0 = Sunday, 6 = Saturday
      businessDays++;
    }
  }
  return businessDays;
}

// ============================================================
// POINTS & BONUS HELPERS
// ============================================================

function calcEarlyBonus(basePoints, completedAt, deadline) {
  var days = daysEarly(completedAt, deadline);
  if (days >= 2) {
    var rate = parseFloat(getConfig('EarlyBonus2PlusDays')) || 0.2;
    return Math.round(basePoints * rate);
  } else if (days >= 1) {
    var rate = parseFloat(getConfig('EarlyBonus1Day')) || 0.1;
    return Math.round(basePoints * rate);
  }
  return 0;
}

function applyPointsFloor(points) {
  var floor = parseInt(getConfig('PointsFloor')) || 0;
  return Math.max(points, floor);
}

// ============================================================
// SESSION & AUTH HELPERS
// ============================================================

function getSessionEmail() {
  return Session.getActiveUser().getEmail();
}

function getLdapFromEmail(email) {
  if (!email) return null;
  return email.split('@')[0].toLowerCase();
}

function getCurrentLdap() {
  return getLdapFromEmail(getSessionEmail());
}

// ============================================================
// RESPONSE HELPERS (for doGet / doPost JSON responses)
// ============================================================

function jsonSuccess(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonError(message) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: message }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// APPEND ROW HELPER
// ============================================================

// Pass an object with column headers as keys; appends a row in correct column order
function appendRow(sheetName, rowObj) {
  var sheet = getSheet(sheetName);
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var row = headers.map(function(h) {
    return rowObj.hasOwnProperty(h) ? rowObj[h] : '';
  });
  sheet.appendRow(row);
  clearSheetDataCache(sheetName);
}

// Update a row where a column matches a value
function updateRow(sheetName, matchCol, matchVal, updates) {
  var sheet = getSheet(sheetName);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var matchIdx = headers.indexOf(matchCol);
  if (matchIdx === -1) throw new Error('Column not found: ' + matchCol);

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][matchIdx]) === String(matchVal)) {
      var rowValues = data[i];
      Object.keys(updates).forEach(function(key) {
        var colIdx = headers.indexOf(key);
        if (colIdx !== -1) {
          rowValues[colIdx] = updates[key];
        }
      });
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowValues]);
      clearSheetDataCache(sheetName);
      return true;
    }
  }
  return false;
}

// Find a single row as an object
function findRow(sheetName, matchCol, matchVal) {
  var rows = getSheetData(sheetName);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][matchCol]) === String(matchVal)) return rows[i];
  }
  return null;
}

// Find all rows matching a column value
function findRows(sheetName, matchCol, matchVal) {
  var rows = getSheetData(sheetName);
  return rows.filter(function(r) {
    return String(r[matchCol]) === String(matchVal);
  });
}

// ============================================================
// CACHE HELPERS
// ============================================================

function getCached(key) {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(key);
    return cached ? JSON.parse(cached, function(k, v) {
      if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d*)?(?:[-+]\d{2}:?\d{2}|Z)?$/.test(v)) {
        return new Date(v);
      }
      return v;
    }) : null;
  } catch(e) {
    Logger.log('[Cache] Read error: ' + e.message);
    return null;
  }
}

function setCached(key, data, ttlSeconds) {
  try {
    var cache = CacheService.getScriptCache();
    var serialized = JSON.stringify(data);
    // Apps Script cache has a 100KB per-item limit
    if (serialized.length < 100000) {
      cache.put(key, serialized, ttlSeconds || 300);
    }
  } catch(e) {
    Logger.log('[Cache] Write error: ' + e.message);
  }
}

function invalidateCache(keys) {
  try {
    var cache = CacheService.getScriptCache();
    cache.removeAll(keys);
    // Also clear local cache if app_config is invalidated
    if (keys.indexOf('app_config') !== -1) _configCache = null;
  } catch(e) {
    Logger.log('[Cache] Invalidate error: ' + e.message);
  }
}

// ============================================================
// BATCH HELPERS
// ============================================================

function batchAppendRows(sheetName, rowObjs) {
  if (!rowObjs || rowObjs.length === 0) return;
  var sheet = getSheet(sheetName);
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) throw new Error('Sheet ' + sheetName + ' has no headers.');

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  var rows = rowObjs.map(function(rowObj) {
    return headers.map(function(h) {
      return rowObj.hasOwnProperty(h) ? rowObj[h] : '';
    });
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  clearSheetDataCache(sheetName);
}

function batchUpdateRows(sheetName, matchCol, matchValToUpdates) {
  var sheet = getSheet(sheetName);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return false;

  var headers = data[0];
  var matchIdx = headers.indexOf(matchCol);
  if (matchIdx === -1) throw new Error('Column not found: ' + matchCol);

  var changed = false;
  for (var i = 1; i < data.length; i++) {
    var matchVal = String(data[i][matchIdx]);
    if (matchValToUpdates.hasOwnProperty(matchVal)) {
      var updates = matchValToUpdates[matchVal];
      Object.keys(updates).forEach(function(key) {
        var colIdx = headers.indexOf(key);
        if (colIdx !== -1) {
          data[i][colIdx] = updates[key];
          changed = true;
        }
      });
    }
  }

  if (changed) {
    sheet.getDataRange().setValues(data);
    clearSheetDataCache(sheetName);
  }
  return changed;
}
