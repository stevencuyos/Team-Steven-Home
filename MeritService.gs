// ============================================================
// MeritService.gs — Streaks, badges, demerits, expiry check
// Play Ops Store
// ============================================================

// ------------------------------------------------------------
// BADGE AWARDING
// ------------------------------------------------------------

// Awards a badge only if the agent doesn't already have it
function awardBadgeIfNew(ldap, badgeId) {
  var existing = findRows('Badges', 'LDAP', ldap);
  var alreadyHas = existing.some(function(b) { return b['BadgeID'] === badgeId; });
  if (alreadyHas) return false;

  var def = findRow('BadgeDefs', 'BadgeID', badgeId);
  if (!def) return false;

  appendRow('Badges', {
    LDAP: ldap,
    BadgeID: badgeId,
    BadgeName: def['BadgeName'],
    AwardedAt: now()
  });

  // Notify agent
  createNotification(ldap, 'badge', 'You earned the "' + def['BadgeName'] + '" badge!');
  Logger.log('[Badge] ' + ldap + ' awarded: ' + badgeId);
  return true;
}

// ------------------------------------------------------------
// BADGE EVALUATION — called after every completion/point event
// ------------------------------------------------------------

function evaluateBadges(ldap, context) {
  // context = { completionCount, isEarly, daysEarly, isAcknowledge, onTimeStreak }

  // First Blood — first task ever completed
  if (context.completionCount === 1) {
    awardBadgeIfNew(ldap, 'FIRST_BLOOD');
  }

  // First Acknowledge
  if (context.isAcknowledge && context.completionCount === 1) {
    awardBadgeIfNew(ldap, 'ACKNOWLEDGED');
  }

  // Speed Demon — 3+ days early
  if (context.isEarly && context.daysEarly >= 3) {
    awardBadgeIfNew(ldap, 'SPEED_DEMON');
  }

  // Perfectionist — 5 consecutive on-time
  if (context.onTimeStreak >= 5) {
    awardBadgeIfNew(ldap, 'PERFECTIONIST');
  }

  // Streak Master — 10 consecutive on-time
  if (context.onTimeStreak >= 10) {
    awardBadgeIfNew(ldap, 'STREAK_MASTER');
  }

  // Overachiever — 10+ tasks this month
  var monthlyCount = getMonthlyCompletionCount(ldap);
  if (monthlyCount >= 10) {
    awardBadgeIfNew(ldap, 'OVERACHIEVER');
  }

  // Kudos King/Queen — 3 validated kudos this month
  var monthlyKudos = getMonthlyKudosCount(ldap);
  if (monthlyKudos >= 3) {
    awardBadgeIfNew(ldap, 'KUDOS_KING');
  }
}

// ------------------------------------------------------------
// STREAK MANAGEMENT
// ------------------------------------------------------------

function updateStreak(ldap, wasOnTime) {
  ensureLeaderboardRow(ldap);
  var row = findRow('Leaderboard', 'LDAP', ldap);
  if (!row) return 0;

  var current = parseInt(row['CurrentStreak']) || 0;
  var best = parseInt(row['BestStreak']) || 0;

  if (wasOnTime) {
    current += 1;
  } else {
    current = 0; // broken streak
  }

  if (current > best) best = current;

  updateRow('Leaderboard', 'LDAP', ldap, {
    CurrentStreak: current,
    BestStreak: best
  });

  // Streak bonuses
  if (wasOnTime) {
    if (current === 3) {
      var bonus3 = parseInt(getConfig('Streak3Bonus')) || 15;
      addPoints(ldap, bonus3, 'Streak bonus (3)');
      createNotification(ldap, 'streak', '3-task streak! +' + bonus3 + ' bonus points.');
    } else if (current === 5) {
      var bonus5 = parseInt(getConfig('Streak5Bonus')) || 25;
      addPoints(ldap, bonus5, 'Streak bonus (5)');
      createNotification(ldap, 'streak', '5-task streak! +' + bonus5 + ' bonus points.');
    } else if (current > 0 && current % 10 === 0) {
      // Every 10-streak milestone
      var bonusMile = parseInt(getConfig('Streak5Bonus')) || 25;
      addPoints(ldap, bonusMile, 'Streak milestone (' + current + ')');
      createNotification(ldap, 'streak', current + '-task streak! +' + bonusMile + ' bonus points.');
    }
  }

  return current;
}

// ------------------------------------------------------------
// DEMERIT HANDLING
// ------------------------------------------------------------

function addDemerit(demeritData, enteredBy) {
  var id = generateDemeritId();
  var timestamp = now();

  var points = resolveDemeritPoints(demeritData.type);

  appendRow('Demerits', {
    ID: id,
    Timestamp: timestamp,
    LDAP: demeritData.ldap,
    Type: demeritData.type,
    Details: demeritData.details || '',
    Points: points,
    EnteredBy: enteredBy,
    NotificationSent: false
  });

  // Deduct points
  deductPoints(demeritData.ldap, Math.abs(points), 'Demerit: ' + demeritData.type);

  // Notify agent (in-app + email)
  var message = 'A demerit was added to your record: ' + demeritData.type + ' (' + points + ' pts). Details: ' + (demeritData.details || 'N/A');
  createNotification(demeritData.ldap, 'demerit', message);
  sendDemeritEmail(demeritData.ldap, demeritData.type, demeritData.details, points);

  // Mark notification sent
  updateRow('Demerits', 'ID', id, { NotificationSent: true });

  // Check Clean Slate badge eligibility will be done at monthly reset
  invalidateCache(['leaderboard_agent', 'leaderboard_manager', 'team_analytics']);
  Logger.log('[Demerit] ' + demeritData.ldap + ' | ' + demeritData.type + ' | ' + points + ' pts');
  return { success: true, id: id, points: points };
}

function onDemeritRowAdded(row) {
  // Called from onEdit trigger when a new row appears in Demerits sheet
  var sheet = getSheet('Demerits');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rowData = data[row - 1];

  var obj = {};
  headers.forEach(function(h, i) { obj[h] = rowData[i]; });

  if (!obj['LDAP'] || obj['NotificationSent'] === true) return;

  var points = obj['Points'] || resolveDemeritPoints(obj['Type']);
  var message = 'A demerit was added to your record: ' + obj['Type'] + ' (' + points + ' pts).';
  createNotification(obj['LDAP'], 'demerit', message);
  sendDemeritEmail(obj['LDAP'], obj['Type'], obj['Details'], points);
  sheet.getRange(row, headers.indexOf('NotificationSent') + 1).setValue(true);
}

function resolveDemeritPoints(type) {
  var map = {
    'RTA': parseInt(getConfig('PointsRTA')) || -20,
    'QA Markdown': parseInt(getConfig('PointsQAMarkdown')) || -15,
    'Missed Task': parseInt(getConfig('PointsMissedTask')) || -10,
    'Abandoned Task': parseInt(getConfig('PointsAbandonedTask')) || -5
  };
  return map[type] || -10;
}

// ------------------------------------------------------------
// CLEAN SLATE BADGE — checked at monthly reset
// ------------------------------------------------------------

function checkCleanSlateBadge(ldap, month) {
  // Check if agent had zero demerits in the given month
  var demerits = findRows('Demerits', 'LDAP', ldap);
  var monthStart = getMonthStart(month);
  var monthEnd = getMonthEnd(month);

  var hasDemerits = demerits.some(function(d) {
    var ts = new Date(d['Timestamp']);
    return ts >= monthStart && ts <= monthEnd;
  });

  if (!hasDemerits) {
    awardBadgeIfNew(ldap, 'CLEAN_SLATE');
  }
}

// ------------------------------------------------------------
// NIGHTLY EXPIRY CHECK
// ------------------------------------------------------------

function runExpiryCheck() {
  var tasks = getSheetData('Tasks');
  var completions = getSheetData('Completions');
  var leaderboard = getSheetData('Leaderboard');

  var lbMap = {};
  leaderboard.forEach(function(r) { lbMap[r.LDAP] = r; });

  var rightNow = now();
  var missedPoints = Math.abs(parseInt(getConfig('PointsMissedTask')) || -10);

  var taskUpdates = {};
  var completionUpdates = {};
  var lbUpdates = {};
  var notifications = [];

  tasks.forEach(function(task) {
    if (task['Status'] === 'Expired' || task['Status'] === 'Completed') return;
    if (!task['Deadline']) return;

    var deadline = new Date(task['Deadline']);
    if (deadline >= rightNow) return; // not expired yet

    var taskId = task['ID'];

    // Find claimed-but-uncompleted entries for this task
    var claimedEntries = completions.filter(function(c) {
      return c['TaskID'] === taskId && c['ClaimedAt'] && !c['CompletedAt'];
    });

    if (claimedEntries.length > 0) {
      claimedEntries.forEach(function(c) {
        completionUpdates[c['ID']] = { CompletedAt: '', Type: 'Missed' };

        var ldap = c['LDAP'];
        var lbRow = lbMap[ldap];
        if (lbRow) {
          var currentMonthly = lbUpdates[ldap] ? lbUpdates[ldap].MonthlyPoints : (lbRow.MonthlyPoints || 0);
          lbUpdates[ldap] = {
            MonthlyPoints: applyPointsFloor(currentMonthly - missedPoints),
            CurrentStreak: 0
          };
        }

        notifications.push({
          ldap: ldap,
          type: 'missed',
          message: 'You missed the deadline for task: ' + task['Title'] + '. ' + (-missedPoints) + ' pts deducted.'
        });
        Logger.log('[Expiry] Missed: ' + ldap + ' | Task: ' + taskId);
      });
      taskUpdates[taskId] = { Status: 'Expired' };

    } else {
      // Published or Claimed but no completion row — just expire
      taskUpdates[taskId] = { Status: 'Expired' };
      if (task['Status'] === 'Published') {
        Logger.log('[Expiry] Unclaimed expired: ' + taskId);
      }
    }
  });

  if (Object.keys(taskUpdates).length > 0) batchUpdateRows('Tasks', 'ID', taskUpdates);
  if (Object.keys(completionUpdates).length > 0) batchUpdateRows('Completions', 'ID', completionUpdates);
  if (Object.keys(lbUpdates).length > 0) batchUpdateRows('Leaderboard', 'LDAP', lbUpdates);
  if (notifications.length > 0) createNotifications(notifications);

  SpreadsheetApp.flush();
  invalidateCache(['leaderboard_agent', 'leaderboard_manager', 'team_analytics']);
  Logger.log('[Expiry Check] Done at ' + formatDateTime(rightNow));
}

// ------------------------------------------------------------
// HELPER: Monthly counts
// ------------------------------------------------------------

function getMonthlyCompletionCount(ldap) {
  var completions = findRows('Completions', 'LDAP', ldap);
  var now = new Date();
  return completions.filter(function(c) {
    if (!c['CompletedAt']) return false;
    var d = new Date(c['CompletedAt']);
    return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
  }).length;
}

function getMonthlyKudosCount(ldap) {
  var kudos = findRows('Kudos', 'LDAP', ldap).filter(function(k) {
    return k['Status'] === 'Approved';
  });
  var now = new Date();
  return kudos.filter(function(k) {
    if (!k['ReviewedAt']) return false;
    var d = new Date(k['ReviewedAt']);
    return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
  }).length;
}

function getMonthStart(month) {
  // month = 'M/YYYY'
  var parts = month.split('/');
  return new Date(parseInt(parts[1]), parseInt(parts[0]) - 1, 1);
}

function getMonthEnd(month) {
  var parts = month.split('/');
  var d = new Date(parseInt(parts[1]), parseInt(parts[0]), 0); // day 0 = last day of previous month
  return d;
}

// ------------------------------------------------------------
// WEEKLY CSAT POINT EVALUATION (Manager UX)
// ------------------------------------------------------------
function evaluateWeeklyCsatPoints(offsetDays) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // NEW: Time Travel offset to look at previous weeks
    var targetDate = new Date();
    if (offsetDays) {
      targetDate.setDate(targetDate.getDate() - parseInt(offsetDays));
    }

    // Generate a unique ID for the target week (e.g., "W14-2026")
    var weekNum = Math.ceil(Math.floor((targetDate - new Date(targetDate.getFullYear(), 0, 1)) / (24 * 60 * 60 * 1000)) / 7);
    var evalKey = 'CSAT-W' + weekNum + '-' + targetDate.getFullYear();

    var configSheet = getSheet('Config');
    var configData = configSheet.getDataRange().getValues();
    
    // Check if already run this week to prevent duplicate points!
    var alreadyRun = configData.some(function(row) { return row[0] === 'LastCsatEval' && row[1] === evalKey; });
    if (alreadyRun) return { success: false, error: 'CSAT Evaluation for ' + evalKey + ' was already completed.' };

    // Targets (Mapped to your exact requirements)
    var targetChat = 84;  // 84%
    var targetPhone = 86; // 86%
    var targetEmail = 50; // 50%
    
    // Point Economy
    var passPts = 20;  // Earn 20 pts per passing metric
    var failPts = -10; // Lose 10 pts per failing metric

    // FIX: Forces the engine to use the month of the specific week we are evaluating
    var month = currentCsatMonth();
    var allCsatRows = readCsatDump();
    var agents = getSheetData('Agents');
    var notifications = [];
    var totalAwarded = 0;

    agents.forEach(function(a) {
      var ldap = a['LDAP'];
      var rows = allCsatRows.filter(function(r) { return String(r.ldap).toLowerCase() === String(ldap).toLowerCase() && r.month === month; });
      if (rows.length === 0) return;

      var agg = aggregateRows(rows); // Gets overall, chat, phone, email & volumes
      var ptsToAdd = 0;
      var reasons = [];

      // 1. Volumes
      var volChat = agg.chatVol || 0;
      var volPhone = agg.phoneVol || 0;
      var volEmail = agg.emailVol || 0;
      var totalVol = volChat + volPhone + volEmail;

      if (totalVol === 0) return; // Skip if no surveys

      // 2. Calculate Custom Target
      var customTarget = ((targetChat * volChat) + (targetPhone * volPhone) + (targetEmail * volEmail)) / totalVol;
      
      // 3. Calculate PTG (Scaled to 100)
      var ptg = customTarget > 0 ? (agg.overall / customTarget) * 100 : 0;

      // 4. Evaluate Metrics (Upgraded with Perfect Sweep tracking)
      var perfectSweep = true;

      if (ptg >= 100.0) { ptsToAdd += passPts; reasons.push('Overall PTG Pass'); }
      else { ptsToAdd += failPts; reasons.push('Overall PTG Fail'); perfectSweep = false; }

      if (volChat > 0) {
        if (agg.chat >= targetChat) { ptsToAdd += passPts; reasons.push('Chat Pass'); }
        else { ptsToAdd += failPts; reasons.push('Chat Fail'); perfectSweep = false; }
      }
      
      if (volPhone > 0) {
        if (agg.phone >= targetPhone) { ptsToAdd += passPts; reasons.push('Phone Pass'); }
        else { ptsToAdd += failPts; reasons.push('Phone Fail'); perfectSweep = false; }
      }

      if (volEmail > 0) {
        if (agg.email >= targetEmail) { ptsToAdd += passPts; reasons.push('Email Pass'); }
        else { ptsToAdd += failPts; reasons.push('Email Fail'); perfectSweep = false; }
      }

      // NEW: Award the CSAT Champion Badge for a Perfect Sweep!
      if (perfectSweep && totalVol > 0) {
        var earned = awardBadgeIfNew(ldap, 'CSAT_CHAMPION');
        if (earned) {
          reasons.push('🏆 CSAT Champion Badge Earned!');
          // Optional: Add a bonus +50 points for the perfect week!
          // ptsToAdd += 50; 
        }
      }

      // 5. Award/Deduct Points
      if (ptsToAdd !== 0) {
        addPoints(ldap, ptsToAdd, 'Weekly CSAT Eval (' + evalKey + ')');
        totalAwarded++;
        var msg = 'Weekly CSAT Review: You ' + (ptsToAdd > 0 ? 'gained ' : 'lost ') + Math.abs(ptsToAdd) + ' pts. [' + reasons.join(', ') + ']';
        notifications.push({ ldap: ldap, type: ptsToAdd > 0 ? 'streak' : 'demerit', message: msg });
      }
    });

    if (notifications.length > 0) createNotifications(notifications);

    // Record that we ran it this week
    updateRow('Config', 'Setting', 'LastCsatEval', { Value: evalKey });
    if (!findRow('Config', 'Setting', 'LastCsatEval')) { appendRow('Config', { Setting: 'LastCsatEval', Value: evalKey }); }

    invalidateCache(['leaderboard_agent', 'leaderboard_manager', 'team_analytics']);
    return { success: true, key: evalKey, count: totalAwarded };
  } finally {
    lock.releaseLock();
  }
}
