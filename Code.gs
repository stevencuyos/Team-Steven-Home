// ============================================================
// Code.gs — Main router, doGet(), session handling, triggers
// Play Ops Store
// Enhanced with health checks and system monitoring
// ============================================================

// ============================================================
// SYSTEM HEALTH & MONITORING
// ============================================================

/**
 * Get system health status for monitoring dashboard
 */
function clientGetSystemHealth() {
  var checks = {
    timestamp: new Date().toISOString(),
    status: 'HEALTHY'
  };
  
  // Check Sheets connectivity
  try {
    SpreadsheetApp.getActiveSpreadsheet();
    checks.sheets = 'OK';
  } catch(e) {
    checks.sheets = 'ERROR: ' + e.message;
    checks.status = 'DEGRADED';
  }
  
  // Check Cache service
  try {
    var cache = CacheService.getScriptCache();
    cache.put('health_check', 'ok', 60);
    checks.cache = 'OK';
  } catch(e) {
    checks.cache = 'ERROR: ' + e.message;
    checks.status = 'DEGRADED';
  }
  
  // Check critical sheets exist
  var criticalSheets = ['Tasks', 'Agents', 'Config', 'Leaderboard'];
  checks.sheets_available = {};
  
  criticalSheets.forEach(function(sheetName) {
    try {
      var sheet = getSheet(sheetName);
      checks.sheets_available[sheetName] = sheet.getLastRow() + ' rows';
    } catch(e) {
      checks.sheets_available[sheetName] = 'MISSING';
      checks.status = 'DEGRADED';
    }
  });
  
  // Check API rate limiting
  checks.rateLimiter = {
    activeKeys: Object.keys(RateLimiter.limits).length,
    status: 'OK'
  };
  
  // Overall status
  if (checks.sheets && checks.cache && checks.status !== 'DEGRADED') {
    checks.status = 'HEALTHY';
  }
  
  auditLog('SYSTEM_HEALTH_CHECK', checks, getCurrentLdap());
  return checks;
}

/**
 * Automated backup of critical sheets (run weekly)
 */
function backupCriticalData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var criticalSheets = ['Tasks', 'Completions', 'Kudos', 'Demerits'];
    
    // Create or get Backups folder
    var folders = DriveApp.getFoldersByName('PlayOps_Backups');
    var backupFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder('PlayOps_Backups');
    
    var timestamp = new Date().toISOString().split('T')[0];
    var backupResults = {};
    
    criticalSheets.forEach(function(sheetName) {
      try {
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) return;
        
        var backupName = sheetName + '_backup_' + timestamp;
        var data = sheet.getDataRange().getValues();
        
        // Create backup spreadsheet
        var backupDoc = SpreadsheetApp.create(backupName);
        var backupSheet = backupDoc.getSheets()[0];
        
        // Copy data
        if (data.length > 0) {
          backupSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
        }
        
        // Move to backup folder
        DriveApp.getFileById(backupDoc.getId()).moveTo(backupFolder);
        
        backupResults[sheetName] = 'OK';
        Logger.log('[BACKUP] ' + backupName + ' created successfully');
      } catch(e) {
        backupResults[sheetName] = 'ERROR: ' + e.message;
        Logger.log('[BACKUP ERROR] ' + sheetName + ': ' + e.message);
      }
    });
    
    auditLog('BACKUP_COMPLETED', backupResults, 'system');
    return { success: true, results: backupResults };
  } catch(e) {
    Logger.log('[BACKUP FAILED] ' + e.message);
    logErrorToSheet({
      error: e.message,
      context: 'backupCriticalData',
      timestamp: new Date()
    });
    return { success: false, error: e.message };
  }
}

// ============================================================
// ENTRY POINT
// ============================================================

function doGet(e) {
  try {
    var ldap = getCurrentLdap();

    // Access check — must be in Agents or Managers sheet
    var role = getUserRole(ldap);
    if (!role) {
      return serveAccessDenied(ldap);
    }

    var template = HtmlService.createTemplateFromFile('Index');
    template.ldap = ldap;
    template.role = role;
    template.email = getSessionEmail();

    auditLog('APP_ACCESSED', { role: role }, ldap);

    return template.evaluate()
      .setTitle('Play Ops Store')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch(e) {
    Logger.log('[doGet ERROR] ' + e.message);
    logErrorToSheet({
      error: e.message,
      context: 'doGet',
      timestamp: new Date()
    });
    return HtmlService.createHtmlOutput('<h2>Error: ' + escHtml(e.message) + '</h2>');
  }
}

function getScriptScopes() {
  // This function exists to ensure UrlFetchApp scope is requested
  UrlFetchApp.fetch('https://www.google.com');
}

// Serve a locked-out page for unauthorized users
function serveAccessDenied(ldap) {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;text-align:center;padding:80px 40px;">' +
    '<h2 style="color:#c0392b;">Access Denied</h2>' +
    '<p>Your account (<strong>' + sanitizeInput(ldap || 'unknown') + '</strong>) is not authorized to access Play Ops Store.</p>' +
    '<p style="color:#888;font-size:13px;">Contact your manager to be added to the system.</p>' +
    '</div>'
  );
  html.setTitle('Play Ops Store — Access Denied');
  auditLog('ACCESS_DENIED', { ldap: ldap }, ldap);
  return html;
}

// Included in HTML templates via <?= include('Styles') ?>
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// HTML escape helper
function escHtml(str) {
  return sanitizeInput(str);
}

// ============================================================
// CLIENT-CALLABLE FUNCTIONS (called via google.script.run)
// These are the API surface exposed to the frontend.
// ============================================================

// --- Session ---
function getSessionInfo() {
  return executeWithErrorHandling(function() {
    var ldap = getCurrentLdap();
    var role = getUserRole(ldap);
    var agent = getAgentProfile(ldap);
    
    // Figure out the team name based on role
    var teamName = 'Play Ops';
    if (role === 'manager' || role === 'supervisor') {
      var mgrRow = findRow('Managers', 'LDAP', ldap);
      if (mgrRow) teamName = mgrRow['Team'];
    } else if (agent) {
      teamName = agent.teamLead || agent.workgroup || 'Play Ops';
    }

    return {
      ldap: ldap,
      email: getSessionEmail(),
      role: role,
      displayName: agent ? agent.displayName : ldap,
      photoUrl: getMomaPhotoUrl(ldap),
      team: teamName
    };
  }, this, 'getSessionInfo');
}

// --- Tasks ---
function clientGetHomepageTasks() {
  if (!RateLimiter.isAllowed(getCurrentLdap(), 20, 60000)) {
    throw new Error('Rate limit exceeded. Please try again later.');
  }
  return getHomepageTasks(getCurrentLdap());
}

function clientGetTaskDetail(taskId) {
  return executeWithErrorHandling(function() {
    return getTaskDetail(taskId, getCurrentLdap());
  }, this, 'clientGetTaskDetail');
}

function clientClaimTask(taskId) {
  return executeWithErrorHandling(function() {
    var result = claimTask(taskId, getCurrentLdap());
    if (result && result.success) {
      bustServerCaches();
      auditLog('TASK_CLAIMED', { taskId: taskId }, getCurrentLdap());
    }
    return result;
  }, this, 'clientClaimTask');
}

function clientCompleteTask(taskId) {
  return executeWithErrorHandling(function() {
    var result = completeTask(taskId, getCurrentLdap());
    if (result && result.success) {
      bustServerCaches();
      auditLog('TASK_COMPLETED', { taskId: taskId }, getCurrentLdap());
    }
    return result;
  }, this, 'clientCompleteTask');
}

function clientAcknowledgeTask(taskId) {
  return executeWithErrorHandling(function() {
    return acknowledgeTask(taskId, getCurrentLdap());
  }, this, 'clientAcknowledgeTask');
}

function clientSearchTasks(query) {
  return executeWithErrorHandling(function() {
    if (!RateLimiter.isAllowed(getCurrentLdap() + '_search', 10, 60000)) {
      throw new Error('Search rate limit exceeded. Try again later.');
    }
    var sanitized = sanitizeInput(query);
    return searchTasks(sanitized, getCurrentLdap());
  }, this, 'clientSearchTasks');
}

// --- Leaderboard ---
function clientGetLeaderboard() {
  return executeWithErrorHandling(function() {
    return getLeaderboard(getCurrentLdap(), getUserRole(getCurrentLdap()));
  }, this, 'clientGetLeaderboard');
}

// --- Profile ---
function clientGetMyProfile() {
  return executeWithErrorHandling(function() {
    return getAgentFullProfile(getCurrentLdap());
  }, this, 'clientGetMyProfile');
}

function clientGetAgentProfile(ldap) {
  return executeWithErrorHandling(function() {
    var requesterLdap = getCurrentLdap();
    var role = getUserRole(requesterLdap);

    if (role === 'manager') return getAgentFullProfile(ldap, true);

    if (role === 'supervisor') {
      var managed = getManagedLdaps(requesterLdap);
      if (managed.indexOf(ldap) !== -1 || ldap === requesterLdap) {
        return getAgentFullProfile(ldap, true);
      }
    }

    // Default to public profile
    return getAgentFullProfile(ldap, false);
  }, this, 'clientGetAgentProfile');
}

// --- Kudos ---
function clientSubmitKudos(formData) {
  return executeWithErrorHandling(function() {
    validateKudosInput(formData);
    var result = submitKudos(formData, getCurrentLdap());
    auditLog('KUDOS_SUBMITTED', { kudosId: result.id }, getCurrentLdap());
    return result;
  }, this, 'clientSubmitKudos');
}

// --- Notifications ---
function clientGetNotifications() {
  return executeWithErrorHandling(function() {
    return getNotifications(getCurrentLdap());
  }, this, 'clientGetNotifications');
}

function clientMarkNotificationsRead() {
  return executeWithErrorHandling(function() {
    return markAllNotificationsRead(getCurrentLdap());
  }, this, 'clientMarkNotificationsRead');
}

// --- Supervisor/Manager functions ---
function clientCreateTask(taskData) {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    validateTaskInput(taskData);
    var result = createTask(taskData, getCurrentLdap());
    auditLog('TASK_CREATED', { taskId: result.id, title: taskData.title }, getCurrentLdap());
    bustServerCaches();
    return result;
  }, this, 'clientCreateTask');
}

function clientUpdateTask(taskId, taskData) {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    validateTaskInput(taskData);
    var result = updateTask(taskId, taskData, getCurrentLdap());
    auditLog('TASK_UPDATED', { taskId: taskId }, getCurrentLdap());
    bustServerCaches();
    return result;
  }, this, 'clientUpdateTask');
}

function clientGetKudosQueue() {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    var managedLdaps = getManagedLdaps(getCurrentLdap());
    return getKudosQueue(managedLdaps);
  }, this, 'clientGetKudosQueue');
}

function clientReviewKudos(kudosId, decision, note) {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    var result = reviewKudos(kudosId, decision, note, getCurrentLdap());
    auditLog('KUDOS_REVIEWED', { kudosId: kudosId, decision: decision }, getCurrentLdap());
    return result;
  }, this, 'clientReviewKudos');
}

function clientAddDemerit(demeritData) {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    // Ensure supervisor can only add demerit to their team
    var managedLdaps = getManagedLdaps(getCurrentLdap());
    if (managedLdaps && managedLdaps.indexOf(demeritData.ldap) === -1) {
      throw new Error('Unauthorized: You can only add demerits to your own team.');
    }
    var result = addDemerit(demeritData, getCurrentLdap());
    auditLog('DEMERIT_ADDED', { ldap: demeritData.ldap, type: demeritData.type }, getCurrentLdap());
    return result;
  }, this, 'clientAddDemerit');
}

function clientGetTeamAnalytics() {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    var managedLdaps = getManagedLdaps(getCurrentLdap());
    return getTeamAnalytics(managedLdaps);
  }, this, 'clientGetTeamAnalytics');
}

function clientGetAgentLookup(ldap) {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    var managedLdaps = getManagedLdaps(getCurrentLdap());
    if (managedLdaps && managedLdaps.indexOf(ldap) === -1 && ldap !== getCurrentLdap()) {
      return getAgentFullProfile(ldap, false); // Public view only
    }
    return getAgentFullProfile(ldap, true);
  }, this, 'clientGetAgentLookup');
}

function clientGetAllAgents() {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    var managedLdaps = getManagedLdaps(getCurrentLdap());
    return getAllAgents(managedLdaps);
  }, this, 'clientGetAllAgents');
}

function clientUpdateAgentDisplayName(ldap, displayName) {
  return executeWithErrorHandling(function() {
    requireManager(); // Only Manager can edit display names
    var result = updateAgentDisplayName(ldap, displayName);
    auditLog('AGENT_DISPLAY_NAME_UPDATED', { ldap: ldap, displayName: displayName }, getCurrentLdap());
    return result;
  }, this, 'clientUpdateAgentDisplayName');
}

// --- Supervisor: Task Manager list ---
function clientGetAllTasks() {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    var managedLdaps = getManagedLdaps(getCurrentLdap());
    return getAllTasksForManager(managedLdaps);
  }, this, 'clientGetAllTasks');
}

// --- Concerns ---
function clientSubmitConcern(concernData) {
  return executeWithErrorHandling(function() {
    var ldap = getCurrentLdap();
    var result = submitConcern(concernData, ldap);
    auditLog('CONCERN_SUBMITTED', { concernId: result.id }, ldap);
    return result;
  }, this, 'clientSubmitConcern');
}

function clientGetConcerns() {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    var ldap = getCurrentLdap();
    Logger.log('[clientGetConcerns] Fetching concerns for ' + ldap);
    var managedLdaps = getManagedLdaps(ldap);
    var result = getConcerns(managedLdaps, ldap);
    Logger.log('[clientGetConcerns] Found ' + (result ? result.length : 0) + ' concerns.');
    return result;
  }, this, 'clientGetConcerns');
}

function clientGetLeadershipList() {
  return executeWithErrorHandling(function() {
    return getLeadershipList();
  }, this, 'clientGetLeadershipList');
}

function clientUpdateConcern(id, updates) {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    var result = updateConcern(id, updates);
    auditLog('CONCERN_UPDATED', { concernId: id, updates: updates }, getCurrentLdap());
    return result;
  }, this, 'clientUpdateConcern');
}

// ============================================================
// ROLE GUARD
// ============================================================

function requireManager() {
  var ldap = getCurrentLdap();
  if (!isManager(ldap)) {
    throw new Error('Unauthorized: manager access required.');
  }
}

function requireSupervisor() {
  var ldap = getCurrentLdap();
  if (!isSupervisor(ldap)) {
    throw new Error('Unauthorized: supervisor/manager access required.');
  }
}

// --- Quality Bridge ---
function clientGetMyQuality(ldap, month) {
  return executeWithErrorHandling(function() {
    var requesterLdap = getCurrentLdap();
    var requesterRole = getUserRole(requesterLdap);
    var isMgmt = requesterRole === 'manager' || requesterRole === 'supervisor';
    var targetLdap = (ldap && isMgmt) ? ldap : requesterLdap;

    if (requesterRole === 'supervisor' && targetLdap !== requesterLdap) {
      var managed = getManagedLdaps(requesterLdap);
      if (managed.indexOf(targetLdap) === -1) {
        targetLdap = requesterLdap;
      }
    }

    return getMyQualityData(targetLdap, month);
  }, this, 'clientGetMyQuality');
}

function clientGetTeamQuality(month) {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    return getTeamQualityData(getCurrentLdap(), month);
  }, this, 'clientGetTeamQuality');
}

function clientGetAllTeamsQuality(month) {
  return executeWithErrorHandling(function() {
    requireManager();
    return getAllTeamsQualityData(month);
  }, this, 'clientGetAllTeamsQuality');
}

function clientGetAvailableQualityMonths() {
  return executeWithErrorHandling(function() {
    return getAvailableQualityMonths();
  }, this, 'clientGetAvailableQualityMonths');
}

// --- CSAT Bridge ---
function clientGetMyCsat(ldap, month) {
  return executeWithErrorHandling(function() {
    var requesterLdap = getCurrentLdap();
    var requesterRole = getUserRole(requesterLdap);
    var isMgmt = requesterRole === 'manager' || requesterRole === 'supervisor';
    var targetLdap = (ldap && isMgmt) ? ldap : requesterLdap;

    if (requesterRole === 'supervisor' && targetLdap !== requesterLdap) {
      var managed = getManagedLdaps(requesterLdap);
      if (managed.indexOf(targetLdap) === -1) {
        targetLdap = requesterLdap; // restrict to self
      }
    }

    return getMyCsatData(targetLdap, month);
  }, this, 'clientGetMyCsat');
}

function clientGetTeamCsat(month) {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    return getTeamCsatData(getCurrentLdap(), month);
  }, this, 'clientGetTeamCsat');
}

function clientGetAllTeamsCsat(month) {
  return executeWithErrorHandling(function() {
    requireManager(); // Strictly Manager only
    return getAllTeamsCsatData(month);
  }, this, 'clientGetAllTeamsCsat');
}

function clientGetAvailableCsatMonths() {
  return executeWithErrorHandling(function() {
    return getAvailableCsatMonths();
  }, this, 'clientGetAvailableCsatMonths');
}


function clientRunQualityAggregation() {
  return executeWithErrorHandling(function() {
    requireManager();
    clearQualityCache();
    auditLog('QUALITY_AGGREGATION_RUN', {}, getCurrentLdap());
    return { success: true };
  }, this, 'clientRunQualityAggregation');
}

function clientRunCsatAggregation() {
  return executeWithErrorHandling(function() {
    requireManager();
    runCsatAggregation();
    auditLog('CSAT_AGGREGATION_RUN', {}, getCurrentLdap());
    return { success: true };
  }, this, 'clientRunCsatAggregation');
}

// ============================================================
// MOMA PHOTO URL
// ============================================================

function getMomaPhotoUrl(ldap) {
  // Standard Google internal Moma photo URL pattern
  return 'https://moma.corp.google.com/person/' + ldap + '/photo';
}

// ============================================================
// AUTOMATED TRIGGERS (set these up in Apps Script trigger menu)
// ============================================================

// Nightly at 00:00 — flags expired/abandoned tasks, deducts points
function triggerNightlyExpiry() {
  try {
    runExpiryCheck();
    sendExpiryWarnings();
  } catch(e) {
    logErrorToSheet({
      error: e.message,
      context: 'triggerNightlyExpiry',
      timestamp: new Date()
    });
  }
}

// 1st of month at 00:01 — resets monthly points, updates tiers
function triggerMonthlyReset() {
  try {
    var today = new Date();
    var resetDay = parseInt(getConfig('MonthlyResetDay')) || 1;
    if (today.getDate() === resetDay) {
      runMonthlyReset();
    }
  } catch(e) {
    logErrorToSheet({
      error: e.message,
      context: 'triggerMonthlyReset',
      timestamp: new Date()
    });
  }
}

// Weekly backup trigger (every Sunday at 2 AM)
function triggerWeeklyBackup() {
  try {
    var today = new Date();
    if (today.getDay() === 0) { // 0 = Sunday
      backupCriticalData();
    }
  } catch(e) {
    logErrorToSheet({
      error: e.message,
      context: 'triggerWeeklyBackup',
      timestamp: new Date()
    });
  }
}

// onEdit trigger — detects new demerit rows
function triggerOnEdit(e) {
  try {
    var sheet = e.range.getSheet();
    if (sheet.getName() === 'Demerits') {
      var row = e.range.getRow();
      if (row > 1) {
        // Small delay to let the full row populate
        Utilities.sleep(500);
        onDemeritRowAdded(row);
      }
    }
  } catch (err) {
    Logger.log('triggerOnEdit error: ' + err.message);
    logErrorToSheet({
      error: err.message,
      context: 'triggerOnEdit',
      timestamp: new Date()
    });
  }
}

// ============================================================
// SETUP — Run once to create all sheets and seed data
// ============================================================

function setupSpreadsheet() {
  try {
    createAllSheets();
    seedConfigSheet();
    seedBadgeDefsSheet();
    seedAgentsSheet();
    Logger.log('Play Ops Store setup complete!');
    auditLog('SYSTEM_SETUP', {}, 'system');
  } catch(e) {
    logErrorToSheet({
      error: e.message,
      context: 'setupSpreadsheet',
      timestamp: new Date()
    });
    throw e;
  }
}

function createAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var schemas = {
    'Tasks': ['ID','Title','Description','Category','TargetType','TargetValue','Deadline','BasePoints','Difficulty','EarlyBonusRate','Status','CreatedBy','CreatedAt','IsFeatured'],
    'Completions': ['ID','TaskID','LDAP','ClaimedAt','CompletedAt','TAT','BasePoints','BonusPoints','TotalPoints','IsFirst','IsEarly','Type'],
    'Kudos': ['ID','LDAP','CaseID','Channel','IssueType','Resolution','WhyKudos','Status','ReviewedBy','ReviewNote','SubmittedAt','ReviewedAt'],
    'Demerits': ['ID','Timestamp','LDAP','Type','Details','Points','EnteredBy','NotificationSent'],
    'Notifications': ['ID','LDAP','Type','Message','IsRead','CreatedAt'],
    'Agents': ['LDAP','Email','DisplayName','Channel','Site','Workgroup','TeamLead'],
    'Managers': ['LDAP','Email','Role','Team'],
    'Leaderboard': ['LDAP','Month','MonthlyPoints','AllTimePoints','CurrentStreak','BestStreak','Tier'],
    'Badges': ['LDAP','BadgeID','BadgeName','AwardedAt'],
    'BadgeDefs': ['BadgeID','BadgeName','Description','SVGIcon','Trigger'],
    'Config': ['Setting','Value'],
    'Concerns': ['ID','Timestamp','LDAP','AddressedTo','Type','Nature','Status','Resolution']
  };

  Object.keys(schemas).forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    // Only write headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(schemas[name]);
      sheet.getRange(1, 1, 1, schemas[name].length)
        .setFontWeight('bold')
        .setBackground('#4285F4')
        .setFontColor('#FFFFFF');
    }
  });
}

function seedConfigSheet() {
  var sheet = getSheet('Config');
  if (sheet.getLastRow() > 1) return; // already seeded

  var defaults = [
    {Setting: 'PointsRTA', Value: -20},
    {Setting: 'PointsQAMarkdown', Value: -15},
    {Setting: 'PointsMissedTask', Value: -10},
    {Setting: 'PointsAbandonedTask', Value: -5},
    {Setting: 'PointsPerfectAttendance', Value: 50},
    {Setting: 'PointsKudosValidated', Value: 30},
    {Setting: 'PointsFirstToComplete', Value: 10},
    {Setting: 'EarlyBonus1Day', Value: 0.10},
    {Setting: 'EarlyBonus2PlusDays', Value: 0.20},
    {Setting: 'Streak3Bonus', Value: 15},
    {Setting: 'Streak5Bonus', Value: 25},
    {Setting: 'PointsFloor', Value: 0},
    {Setting: 'MonthlyResetDay', Value: 1},
    {Setting: 'ExpiryCheckTime', Value: '00:00'},
    {Setting: 'AdminEmail', Value: 'stevenjosephc@google.com'}
  ];

  batchAppendRows('Config', defaults);
}

function seedBadgeDefsSheet() {
  var sheet = getSheet('BadgeDefs');
  if (sheet.getLastRow() > 1) return;

  var badges = [
    {BadgeID: 'FIRST_BLOOD', BadgeName: 'First Blood', Description: 'First task ever completed', SVGIcon: '', Trigger: 'first_completion'},
    {BadgeID: 'SPEED_DEMON', BadgeName: 'Speed Demon', Description: 'Complete a task 3+ days before deadline', SVGIcon: '', Trigger: 'early_3days'},
    {BadgeID: 'PERFECTIONIST', BadgeName: 'Perfectionist', Description: '5 consecutive on-time completions', SVGIcon: '', Trigger: 'ontime_streak_5'},
    {BadgeID: 'KUDOS_KING', BadgeName: 'Kudos King/Queen', Description: '3 validated Kudos in a single month', SVGIcon: '', Trigger: 'kudos_3_month'},
    {BadgeID: 'CLEAN_SLATE', BadgeName: 'Clean Slate', Description: 'Full month with zero demerits', SVGIcon: '', Trigger: 'zero_demerits_month'},
    {BadgeID: 'OVERACHIEVER', BadgeName: 'Overachiever', Description: 'Complete 10+ tasks in a single month', SVGIcon: '', Trigger: 'tasks_10_month'},
    {BadgeID: 'STREAK_MASTER', BadgeName: 'Streak Master', Description: '10-task on-time streak', SVGIcon: '', Trigger: 'ontime_streak_10'},
    {BadgeID: 'ACKNOWLEDGED', BadgeName: 'Acknowledged', Description: 'First announcement acknowledgement', SVGIcon: '', Trigger: 'first_acknowledge'},
    {BadgeID: 'VETERAN', BadgeName: 'Veteran', Description: 'Active for 3 consecutive months', SVGIcon: '', Trigger: 'active_3months'},
    {BadgeID: 'LEGEND', BadgeName: 'Legend', Description: 'Reach Legend tier for the first time', SVGIcon: '', Trigger: 'tier_legend'},
    {BadgeID: 'COMEBACK_KID', BadgeName: 'Comeback Kid', Description: 'Reach Gold+ tier after being Bronze the previous month', SVGIcon: '', Trigger: 'comeback_gold'}
  ];

  batchAppendRows('BadgeDefs', badges);
}

function seedAgentsSheet() {
  var sheet = getSheet('Agents');
  if (sheet.getLastRow() > 1) return;

  var agents = [
    'anggam','maekaila','caranoo','codeniera','genlee',
    'nicoleflores','jecylangela','chilado','kennethwhesley',
    'leonoral','krisangelo','kareenm','palomarj','johnnoelp',
    'refugio','torreon','vequezo','cedricanthony'
  ];

  var domain = '@google.com';
  var rowObjs = agents.map(function(ldap) {
    return {
      LDAP: ldap,
      Email: ldap + domain,
      DisplayName: '',
      Channel: 'Chat',
      Site: 'Cebu',
      Workgroup: 'Play Ops',
      TeamLead: 'stevenjosephc'
    };
  });

  batchAppendRows('Agents', rowObjs);

  // Seed manager
  var managerSheet = getSheet('Managers');
  if (managerSheet.getLastRow() < 2) {
    batchAppendRows('Managers', [{
      LDAP: 'stevenjosephc',
      Email: 'stevenjosephc@google.com',
      Role: 'manager',
      Team: 'Team Steven'
    }]);
  }
}

// ============================================================
// TEST & DEBUG FUNCTIONS
// ============================================================

function testTaskManager() {
  try {
    var result = getAllTasksForManager();
    Logger.log(JSON.stringify(result));
  } catch(e) {
    Logger.log('ERROR: ' + e.message + ' | Line: ' + e.lineNumber);
  }
}

function testSheets() {
  var names = ['Tasks','Completions','Kudos','Demerits','Notifications','Agents','Managers','Leaderboard','Badges','BadgeDefs','Config'];
  names.forEach(function(n) {
    try {
      var s = getSheet(n);
      Logger.log(n + ': OK (' + s.getLastRow() + ' rows)');
    } catch(e) {
      Logger.log(n + ': MISSING');
    }
  });
}

function testManagerAuth() {
  var ldap = getCurrentLdap();
  var role = getUserRole(ldap);
  Logger.log('LDAP: ' + ldap);
  Logger.log('Role: ' + role);
  Logger.log('Manager row: ' + JSON.stringify(findRow('Managers', 'LDAP', ldap)));
}

function testClientGetAllTasks() {
  try {
    requireManager();
    var result = getAllTasksForManager();
    Logger.log('Result type: ' + typeof result);
    Logger.log('Result: ' + JSON.stringify(result));
  } catch(e) {
    Logger.log('ERROR: ' + e.message);
    Logger.log('Stack: ' + e.stack);
  }
}

function clientGetPendingAgents(taskId) {
  requireManager();
  return getPendingAgentsForTask(taskId);
}

function deletePlayOpsTask(taskId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  var data = sheet.getDataRange().getValues();
  
  // Loop through to find the matching Task ID (assuming it's in column A/index 0)
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === taskId) {
      sheet.deleteRow(i + 1); // +1 because sheet rows are 1-indexed
      clearSheetDataCache('Tasks');
      auditLog('TASK_DELETED', { taskId: taskId }, getCurrentLdap());
      return { success: true };
    }
  }
  return { error: 'Task not found in database.' };
}

// ============================================================
// FEEDBACK SYSTEM
// ============================================================

function clientSubmitFeedback(type, text) {
  return executeWithErrorHandling(function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Feedback');
    
    // Auto-create the tab if it doesn't exist yet!
    if (!sheet) {
      sheet = ss.insertSheet('Feedback');
      sheet.appendRow(['Timestamp', 'Agent LDAP', 'Feedback Type', 'Details', 'Status']);
      sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f3f3");
      sheet.setFrozenRows(1);
    }
    
    // Grab the user who submitted it
    var ldap = getCurrentLdap();
    var timestamp = new Date();
    
    // Log the feedback into the sheet
    sheet.appendRow([timestamp, ldap, sanitizeInput(type), sanitizeInput(text), 'New']);
    clearSheetDataCache('Feedback');
    
    auditLog('FEEDBACK_SUBMITTED', { type: type }, ldap);
    
    return { success: true };
  }, this, 'clientSubmitFeedback');
}

// ============================================================
// CSAT & COACHING ENDPOINTS
// ============================================================

function clientGetAgentCoaching(ldap, month) {
  return executeWithErrorHandling(function() {
    var requesterLdap = getCurrentLdap();
    var requesterRole = getUserRole(requesterLdap);
    var isMgmt = requesterRole === 'manager' || requesterRole === 'supervisor';
    var targetLdap = (ldap && isMgmt) ? ldap : requesterLdap;

    if (requesterRole === 'supervisor' && targetLdap !== requesterLdap) {
      var managed = getManagedLdaps(requesterLdap);
      if (managed.indexOf(targetLdap) === -1) {
        targetLdap = requesterLdap;
      }
    }

    return getAgentCsatCoaching(targetLdap, month);
  }, this, 'clientGetAgentCoaching');
}

function clientGetTeamCoaching(month) {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    return getTeamCsatCoaching(getCurrentLdap(), month);
  }, this, 'clientGetTeamCoaching');
}

// ============================================================
// SCHEDULE ENDPOINTS
// ============================================================

function clientGetAgentSchedule(month, targetLdap) {
  return executeWithErrorHandling(function() {
    var ldap = getCurrentLdap();
    var role = getUserRole(ldap);
    var isMgmt = role === 'manager' || role === 'supervisor';
    var effectiveLdap = (targetLdap && isMgmt) ? targetLdap : ldap;

    if (role === 'supervisor' && effectiveLdap !== ldap) {
      var managed = getManagedLdaps(ldap);
      if (managed.indexOf(effectiveLdap) === -1) {
        effectiveLdap = ldap;
      }
    }

    return getAgentScheduleData(effectiveLdap, month);
  }, this, 'clientGetAgentSchedule');
}

function clientGetScheduleAgentList() {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    var managedLdaps = getManagedLdaps(getCurrentLdap());
    var agents = getAllAgents(managedLdaps);
    return agents.map(function(a) {
      return { ldap: a.ldap, displayName: a.displayName || a.ldap };
    }).filter(function(a) { return a.ldap; });
  }, this, 'clientGetScheduleAgentList');
}

function clientGetTeamSchedule(dateKey) {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    return getTeamScheduleData(getCurrentLdap(), dateKey);
  }, this, 'clientGetTeamSchedule');
}

function clientGetScheduleMonths() {
  return executeWithErrorHandling(function() {
    return getAvailableScheduleMonths();
  }, this, 'clientGetScheduleMonths');
}

// ============================================================
// CACHE MANAGEMENT
// ============================================================

function bustServerCaches() {
  var cache = CacheService.getScriptCache();
  cache.removeAll([
    'shifts_display_v1',
    'breaks_display_v1',
    'csat_raw_dump',
    'leaderboard_agent',
    'leaderboard_manager',
    'team_analytics',
    'app_config'
  ]);
  Logger.log('[Cache] All server caches busted at ' + new Date());
}

function warmCsatCache() {
  try {
    readCsatDump();
    Logger.log('[CSAT Cache] Warmed successfully at ' + new Date());
  } catch(e) {
    Logger.log('[CSAT Cache] Warm error: ' + e.message);
  }
}

function clientGetDashboardData() {
  return executeWithErrorHandling(function() {
    var ldap = getCurrentLdap();
    var currentMonth = currentCsatMonth(); // Uses YYYY-MM format

    var data = {
      tasks: getHomepageTasks(ldap),
      profile: getAgentFullProfile(ldap, true),
      notifications: getNotifications(ldap)
    };

    // Background load Quality stats for current month to improve perceived speed
    try {
      data.quality = getMyQualityData(ldap, currentMonth);
    } catch(e) {
      Logger.log('[Dashboard] Quality pre-load failed: ' + e.message);
    }

    return data;
  }, this, 'clientGetDashboardData');
}
