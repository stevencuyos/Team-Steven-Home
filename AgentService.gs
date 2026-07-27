// ============================================================
// AgentService.gs — Access control, role checking, profiles
// Play Ops Store
// ============================================================

// ------------------------------------------------------------
// ROLE & ACCESS CONTROL
// ------------------------------------------------------------

// Returns 'manager', 'supervisor', 'delegate', 'agent', or null
function getUserRole(ldap) {
  if (!ldap) return null;
  var managerRow = findRow('Managers', 'LDAP', ldap);
  if (managerRow) {
    var role = String(managerRow['Role']).toLowerCase();
    // NEW: Allow 'delegate' as a valid elevated role
    if (role === 'manager' || role === 'supervisor' || role === 'delegate') return role;
    return 'manager'; // fallback
  }

  var agent = findRow('Agents', 'LDAP', ldap);
  if (agent) return 'agent';

  return null;
}

function isManager(ldap) {
  return getUserRole(ldap) === 'manager';
}

function isSupervisor(ldap) {
  var role = getUserRole(ldap);
  // NEW: Delegates pass the Supervisor check, giving them access to the Manager UI
  return role === 'supervisor' || role === 'manager' || role === 'delegate';
}

function isAgent(ldap) {
  var role = getUserRole(ldap);
  return role === 'agent' || role === 'supervisor' || role === 'manager' || role === 'delegate';
}

function getManagedLdaps(ldap) {
  var role = getUserRole(ldap);
  if (role === 'manager') return null; // null means "all"

  if (role === 'supervisor') {
    var agents = getSheetData('Agents');
    return agents
      .filter(function(a) { return String(a['TeamLead']).toLowerCase() === ldap.toLowerCase(); })
      .map(function(a) { return a['LDAP']; });
  }

  // NEW: Delegate Scope Logic
  if (role === 'delegate') {
    var agents = getSheetData('Agents');
    // Find the delegate's own record to identify their team lead
    var myAgentRecord = agents.find(function(a) { return String(a['LDAP']).toLowerCase() === ldap.toLowerCase(); });
    
    if (myAgentRecord && myAgentRecord['TeamLead']) {
      var myLead = String(myAgentRecord['TeamLead']).toLowerCase();
      // Return everyone who shares the same TeamLead
      return agents
        .filter(function(a) { return String(a['TeamLead']).toLowerCase() === myLead; })
        .map(function(a) { return a['LDAP']; });
    }
  }

  return [];
}

// ------------------------------------------------------------
// AGENT PROFILE — BASIC
// ------------------------------------------------------------

function getAgentProfile(ldap) {
  if (!ldap) return null;

  var agentRow = findRow('Agents', 'LDAP', ldap);
  var managerRow = findRow('Managers', 'LDAP', ldap);

  if (!agentRow && !managerRow) return null;

  var leaderboardRow = findRow('Leaderboard', 'LDAP', ldap);
  var tier = leaderboardRow ? leaderboardRow['Tier'] : 'Bronze';
  var monthlyPoints = leaderboardRow ? (leaderboardRow['MonthlyPoints'] || 0) : 0;
  var allTimePoints = leaderboardRow ? (leaderboardRow['AllTimePoints'] || 0) : 0;
  var currentStreak = leaderboardRow ? (leaderboardRow['CurrentStreak'] || 0) : 0;

  return {
    ldap: ldap,
    email: agentRow ? agentRow['Email'] : (managerRow ? managerRow['Email'] : ''),
    displayName: formatDisplayName(ldap),
    channel: agentRow ? agentRow['Channel'] : '',
    site: agentRow ? agentRow['Site'] : '',
    workgroup: agentRow ? agentRow['Workgroup'] : '',
    teamLead: agentRow ? agentRow['TeamLead'] : '',
    role: managerRow ? String(managerRow['Role']).toLowerCase() : 'agent',
    tier: tier || 'Bronze',
    monthlyPoints: monthlyPoints,
    allTimePoints: allTimePoints,
    currentStreak: currentStreak,
    photoUrl: getMomaPhotoUrl(ldap)
  };
}

// ------------------------------------------------------------
// AGENT PROFILE — FULL (for profile page)
// ------------------------------------------------------------

function getAgentFullProfile(ldap, includeManagerData) {
  var base = getAgentProfile(ldap);
  if (!base) return null;

  // Pre-load necessary data
  var allBadgeDefs = getSheetData('BadgeDefs');
  var allTasks = getSheetData('Tasks');

  var badgeDefsMap = {};
  allBadgeDefs.forEach(function(d) { badgeDefsMap[d.BadgeID] = d; });

  var tasksMap = {};
  allTasks.forEach(function(t) { tasksMap[t.ID] = t; });

  // Badges
  var badgeRows = findRows('Badges', 'LDAP', ldap);
  var badges = badgeRows.map(function(b) {
    var def = badgeDefsMap[b['BadgeID']];
    return {
      badgeId: b['BadgeID'],
      name: b['BadgeName'],
      awardedAt: formatDate(b['AwardedAt']),
      description: def ? def['Description'] : '',
      icon: def ? def['SVGIcon'] : ''
    };
  });

  // Completion history (last 20)
  var completions = findRows('Completions', 'LDAP', ldap);
  completions.sort(function(a, b) {
    return new Date(b['CompletedAt']) - new Date(a['CompletedAt']);
  });
  var recentCompletions = completions.slice(0, 20).map(function(c) {
    var task = tasksMap[c['TaskID']];
    return {
      taskId: c['TaskID'],
      taskTitle: task ? task['Title'] : c['TaskID'],
      completedAt: formatDate(c['CompletedAt']),
      totalPoints: c['TotalPoints'],
      isEarly: c['IsEarly'],
      tat: c['TAT'],
      type: c['Type']
    };
  });

  // Kudos history
  var kudosList = findRows('Kudos', 'LDAP', ldap).filter(function(k) {
    return k['Status'] === 'Approved';
  });

  // Demerits (for agent's own view or manager view)
  var demerits = [];
  if (includeManagerData) {
    var demeritRows = findRows('Demerits', 'LDAP', ldap);
    demeritRows.sort(function(a, b) {
      return new Date(b['Timestamp']) - new Date(a['Timestamp']);
    });
    demerits = demeritRows.map(function(d) {
      return {
        id: d['ID'],
        timestamp: formatDateTime(d['Timestamp']),
        type: d['Type'],
        details: d['Details'],
        points: d['Points'],
        enteredBy: d['EnteredBy']
      };
    });
  }

  // Leaderboard rank
  var rank = getAgentRank(ldap);

  var profile = Object.assign({}, base, {
    badges: badges,
    recentCompletions: recentCompletions,
    kudosCount: kudosList.length,
    tasksCompleted: completions.length,
    rank: rank
  });

  if (includeManagerData) {
    profile.demerits = demerits;
    profile.demeritCount = demerits.length;
    profile.demeritPoints = demerits.reduce(function(sum, d) {
      return sum + (parseInt(d.points) || 0);
    }, 0);
    profile.lastActive = recentCompletions.length > 0 ? recentCompletions[0].completedAt : 'No activity';
  }

  return profile;
}

// ------------------------------------------------------------
// ALL AGENTS (for manager views)
// ------------------------------------------------------------

function getAllAgents(managedLdaps) {
  var agentsData = getSheetData('Agents');
  if (managedLdaps) {
    agentsData = agentsData.filter(function(a) { return managedLdaps.indexOf(a['LDAP']) !== -1; });
  }
  var managersData = getSheetData('Managers');
  var leaderboardData = getSheetData('Leaderboard');

  var lbMap = {};
  leaderboardData.forEach(function(r) { lbMap[r.LDAP] = r; });

  var mgrMap = {};
  managersData.forEach(function(m) { mgrMap[m.LDAP] = m; });

  return agentsData.map(function(a) {
    var ldap = a['LDAP'];
    if (!ldap) return null;

    var lbRow = lbMap[ldap];
    var mgrRow = mgrMap[ldap];

    return {
      ldap: ldap,
      email: a['Email'] || '',
      displayName: a['DisplayName'] ? a['DisplayName'].trim() : ldap.toLowerCase(),
      channel: a['Channel'] || '',
      site: a['Site'] || '',
      workgroup: a['Workgroup'] || '',
      teamLead: a['TeamLead'] || '',
      role: mgrRow ? String(mgrRow['Role']).toLowerCase() : 'agent',
      tier: lbRow ? (lbRow['Tier'] || 'Bronze') : 'Bronze',
      monthlyPoints: lbRow ? (lbRow['MonthlyPoints'] || 0) : 0,
      allTimePoints: lbRow ? (lbRow['AllTimePoints'] || 0) : 0,
      currentStreak: lbRow ? (lbRow['CurrentStreak'] || 0) : 0,
      photoUrl: getMomaPhotoUrl(ldap)
    };
  }).filter(Boolean);
}

// ------------------------------------------------------------
// LEADERBOARD RANK
// ------------------------------------------------------------

function getAgentRank(ldap) {
  var board = getSheetData('Leaderboard');
  board.sort(function(a, b) {
    return (b['MonthlyPoints'] || 0) - (a['MonthlyPoints'] || 0);
  });
  for (var i = 0; i < board.length; i++) {
    if (board[i]['LDAP'] === ldap) return i + 1;
  }
  return null;
}

// ------------------------------------------------------------
// LEADERBOARD
// ------------------------------------------------------------

function getLeaderboard(requestingLdap, role) {
  var cacheKey = 'leaderboard_' + role;
  var cached = getCached(cacheKey);
  if (cached) {
    // Re-apply isMe flag since it's user-specific
    cached.forEach(function(e) { e.isMe = e.ldap === requestingLdap; });
    return cached;
  }

  var board = getSheetData('Leaderboard');
  board.sort(function(a, b) {
    return (b['MonthlyPoints'] || 0) - (a['MonthlyPoints'] || 0);
  });

  // Load all sheets once if manager
  var allCompletions = [], allKudos = [], allDemerits = [];
  if (role === 'manager') {
    allCompletions = getSheetData('Completions');
    allKudos = getSheetData('Kudos');
    allDemerits = getSheetData('Demerits');
  }

      // --- NEW: FETCH LIVE CSAT DATA FOR LEADERBOARD ---
  var csatMonth = currentCsatMonth(); 
  var allCsatRows = readCsatDump();
  var curCsatRows = allCsatRows.filter(function(r) { return r.month === csatMonth; });


  // Pre-load agents to avoid individual findRow calls
  var allAgents = getSheetData('Agents');
  var agentsMap = {};
  allAgents.forEach(function(a) { agentsMap[a.LDAP] = a; });

  var result = board.map(function(row, index) {
    var ldap = row['LDAP'];
    var agent = agentsMap[ldap];
    var dName = (agent && agent.DisplayName) ? agent.DisplayName.trim() : ldap.toLowerCase();

    var entry = {
      rank: index + 1,
      ldap: ldap,
      displayName: dName,
      tier: row['Tier'] || 'Bronze',
      monthlyPoints: row['MonthlyPoints'] || 0,
      allTimePoints: row['AllTimePoints'] || 0,
      currentStreak: row['CurrentStreak'] || 0,
      photoUrl: getMomaPhotoUrl(ldap),
      isMe: ldap === requestingLdap
    };

    if (role === 'manager') {
      var ldap = row['LDAP'];
      var completions = allCompletions.filter(function(c) { return c['LDAP'] === ldap; });
      var kudos = allKudos.filter(function(k) { return k['LDAP'] === ldap && k['Status'] === 'Approved'; });
      var demerits = allDemerits.filter(function(d) { return d['LDAP'] === ldap; });
      var doneCompletions = completions.filter(function(c) { return c['CompletedAt']; });
      doneCompletions.sort(function(a, b) { return new Date(b['CompletedAt']) - new Date(a['CompletedAt']); });

      entry.tasksCompleted = doneCompletions.length;
      entry.kudosCount = kudos.length;
      entry.demeritCount = demerits.length;
      entry.lastActive = doneCompletions.length > 0 ? formatDate(doneCompletions[0]['CompletedAt']) : 'No activity';
    }

    // --- NEW: ATTACH CSAT AGGREGATION & TARGETS ---
    var agentCsatRows = curCsatRows.filter(function(r) { return String(r.ldap).toLowerCase() === String(ldap).toLowerCase(); });
    var agentCsat = aggregateRows(agentCsatRows);
    
    // Updated to match your exact aggregateRows variables!
    var vChat = agentCsat.chatDen || 0;
    var vPhone = agentCsat.phoneDen || 0;
    var vEmail = agentCsat.emailDen || 0;
    var vTotal = vChat + vPhone + vEmail;
    
    // Targets switched to whole numbers (84, 86, 50) to match the raw data scale
    var cTarget = vTotal > 0 ? ((84 * vChat) + (86 * vPhone) + (50 * vEmail)) / vTotal : 0;

    entry.csatOverall = agentCsat.overall;
    entry.csatChat = agentCsat.chat;
    entry.csatPhone = agentCsat.phone;
    entry.csatEmail = agentCsat.email;
    entry.csatTarget = cTarget;
    entry.csatPtg = cTarget > 0 ? (agentCsat.overall / cTarget) * 100 : 0;
    // ----------------------------------------------
    return entry;
  });

  setCached(cacheKey, result, 300); // 5 minute TTL
  return result;
}

// ------------------------------------------------------------
// LEADERBOARD TIER CALCULATION
// ------------------------------------------------------------

// Called after every point event to recalculate tier
function recalculateTier(ldap) {
  var row = findRow('Leaderboard', 'LDAP', ldap);
  if (!row) return;

  var points = row['AllTimePoints'] || 0;
  var tier = getTierForPoints(points);

  // Check for Legend badge trigger
  var previousTier = row['Tier'];
  updateRow('Leaderboard', 'LDAP', ldap, { Tier: tier });

  if (tier === 'Legend' && previousTier !== 'Legend') {
    awardBadgeIfNew(ldap, 'LEGEND');
  }

  // Comeback Kid: Gold+ this month after Bronze last month
  checkComebackKid(ldap, previousTier, tier);
}

function getTierForPoints(allTimePoints) {
  if (allTimePoints >= 2000) return 'Legend';
  if (allTimePoints >= 1000) return 'Platinum';
  if (allTimePoints >= 500)  return 'Gold';
  if (allTimePoints >= 200)  return 'Silver';
  return 'Bronze';
}

function checkComebackKid(ldap, previousTier, currentTier) {
  var goldTiers = ['Gold', 'Platinum', 'Legend'];
  if (previousTier === 'Bronze' && goldTiers.indexOf(currentTier) !== -1) {
    awardBadgeIfNew(ldap, 'COMEBACK_KID');
  }
}

// ------------------------------------------------------------
// LEADERBOARD ROW INIT (called on first point event for agent)
// ------------------------------------------------------------

function ensureLeaderboardRow(ldap) {
  var existing = findRow('Leaderboard', 'LDAP', ldap);
  if (!existing) {
    var month = getCurrentMonth();
    appendRow('Leaderboard', {
      LDAP: ldap,
      Month: month,
      MonthlyPoints: 0,
      AllTimePoints: 0,
      CurrentStreak: 0,
      BestStreak: 0,
      Tier: 'Bronze'
    });
  }
}

function getCurrentMonth() {
  var d = new Date();
  return (d.getMonth() + 1) + '/' + d.getFullYear();
}

// ------------------------------------------------------------
// POINTS MANAGEMENT
// ------------------------------------------------------------

function addPoints(ldap, points, reason) {
  ensureLeaderboardRow(ldap);

  var row = findRow('Leaderboard', 'LDAP', ldap);
  if (!row) return;

  var newMonthly = applyPointsFloor((row['MonthlyPoints'] || 0) + points);
  // If points are negative (demerit), don't touch AllTimePoints. Lifetime XP only goes up!
  var newAllTime = points < 0 ? (row['AllTimePoints'] || 0) : Math.max(0, (row['AllTimePoints'] || 0) + points);
  
  updateRow('Leaderboard', 'LDAP', ldap, {
    MonthlyPoints: newMonthly,
    AllTimePoints: newAllTime
  });

  recalculateTier(ldap);
  invalidateCache(['leaderboard_agent', 'leaderboard_manager', 'team_analytics']);
  Logger.log('[Points] ' + ldap + ' | ' + (points > 0 ? '+' : '') + points + ' | ' + reason + ' | Monthly: ' + newMonthly + ' | AllTime: ' + newAllTime);
}

function deductPoints(ldap, points, reason) {
  // points should be negative or will be made negative here
  var deduction = points > 0 ? -points : points;
  addPoints(ldap, deduction, reason);
}

// ------------------------------------------------------------
// MONTHLY RESET
// ------------------------------------------------------------

function runMonthlyReset() {
  var agents = getSheetData('Agents');
  var month = getCurrentMonth();
  var prevMonth = getPreviousMonth();

  var leaderboardData = getSheetData('Leaderboard');
  var completionsData = getSheetData('Completions');
  var demeritsData = getSheetData('Demerits');
  var badgesData = getSheetData('Badges');
  var badgeDefsData = getSheetData('BadgeDefs');

  var leaderboardLdaps = leaderboardData.map(function(r) { return r.LDAP; });
  var badgesMap = {};
  badgesData.forEach(function(b) {
    if (!badgesMap[b.LDAP]) badgesMap[b.LDAP] = [];
    badgesMap[b.LDAP].push(b.BadgeID);
  });

  var badgeDefsMap = {};
  badgeDefsData.forEach(function(d) { badgeDefsMap[d.BadgeID] = d; });

  var monthStart = getMonthStart(prevMonth);
  var monthEnd = getMonthEnd(prevMonth);

  // Group completions and demerits by LDAP to avoid nested filter (O(N^2))
  var completionsByLdap = {};
  completionsData.forEach(function(c) {
    if (!completionsByLdap[c.LDAP]) completionsByLdap[c.LDAP] = [];
    completionsByLdap[c.LDAP].push(c);
  });

  var demeritsByLdap = {};
  demeritsData.forEach(function(d) {
    if (!demeritsByLdap[d.LDAP]) demeritsByLdap[d.LDAP] = [];
    demeritsByLdap[d.LDAP].push(d);
  });

  var newLeaderboardRows = [];
  var leaderboardUpdates = {};
  var newBadgeRows = [];
  var newNotifications = [];

  agents.forEach(function(a) {
    var ldap = a['LDAP'];
    if (!ldap) return;

    // 1. Ensure Leaderboard Row
    if (leaderboardLdaps.indexOf(ldap) === -1) {
       newLeaderboardRows.push({
         LDAP: ldap, Month: month, MonthlyPoints: 0, AllTimePoints: 0,
         CurrentStreak: 0, BestStreak: 0, Tier: 'Bronze'
       });
       leaderboardLdaps.push(ldap);
    }

    // 2. Veteran Badge Check (simplified in-memory)
    var agentCompletions = completionsByLdap[ldap] || [];
    var distinctMonths = {};
    agentCompletions.forEach(function(c) {
      if (c.CompletedAt) {
        var d = new Date(c.CompletedAt);
        distinctMonths[d.getFullYear() + '-' + d.getMonth()] = true;
      }
    });
    if (Object.keys(distinctMonths).length >= 3) {
      if (!badgesMap[ldap] || badgesMap[ldap].indexOf('VETERAN') === -1) {
        var def = badgeDefsMap['VETERAN'];
        if (def) {
          newBadgeRows.push({ LDAP: ldap, BadgeID: 'VETERAN', BadgeName: def.BadgeName, AwardedAt: now() });
          newNotifications.push({ ldap: ldap, type: 'badge', message: 'You earned the "' + def.BadgeName + '" badge!' });
          if (!badgesMap[ldap]) badgesMap[ldap] = [];
          badgesMap[ldap].push('VETERAN');
        }
      }
    }

    // 3. Clean Slate Badge Check
    var agentDemerits = demeritsByLdap[ldap] || [];
    var hasDemeritsInPrevMonth = agentDemerits.some(function(d) {
      var ts = new Date(d.Timestamp);
      return ts >= monthStart && ts <= monthEnd;
    });
    if (!hasDemeritsInPrevMonth) {
      if (!badgesMap[ldap] || badgesMap[ldap].indexOf('CLEAN_SLATE') === -1) {
        var def = badgeDefsMap['CLEAN_SLATE'];
        if (def) {
          newBadgeRows.push({ LDAP: ldap, BadgeID: 'CLEAN_SLATE', BadgeName: def.BadgeName, AwardedAt: now() });
          newNotifications.push({ ldap: ldap, type: 'badge', message: 'You earned the "' + def.BadgeName + '" badge!' });
          if (!badgesMap[ldap]) badgesMap[ldap] = [];
          badgesMap[ldap].push('CLEAN_SLATE');
        }
      }
    }

    // 4. Collect Leaderboard Update
    leaderboardUpdates[ldap] = {
      MonthlyPoints: 0,
      Month: month
    };
  });

  if (newLeaderboardRows.length > 0) batchAppendRows('Leaderboard', newLeaderboardRows);
  if (newBadgeRows.length > 0) batchAppendRows('Badges', newBadgeRows);
  if (newNotifications.length > 0) createNotifications(newNotifications);
  batchUpdateRows('Leaderboard', 'LDAP', leaderboardUpdates);

  Logger.log('[Monthly Reset] Completed for month: ' + month);
}

function checkVeteranBadge(ldap) {
  // Award Veteran badge if agent has completions in 3+ distinct months
  var completions = findRows('Completions', 'LDAP', ldap);
  var months = {};
  completions.forEach(function(c) {
    if (c['CompletedAt']) {
      var d = new Date(c['CompletedAt']);
      var key = d.getFullYear() + '-' + d.getMonth();
      months[key] = true;
    }
  });
  if (Object.keys(months).length >= 3) {
    awardBadgeIfNew(ldap, 'VETERAN');
  }
}

// ------------------------------------------------------------
// HELPERS
// ------------------------------------------------------------
function updateAgentDisplayName(ldap, displayName) {
  if (!ldap || !displayName || displayName.trim() === '') {
    return { success: false, error: 'Display name cannot be empty.' };
  }
  var result = updateRow('Agents', 'LDAP', ldap, { DisplayName: displayName.trim() });
  if (!result) return { success: false, error: 'Agent not found.' };
  invalidateCache(['leaderboard_agent', 'leaderboard_manager', 'team_analytics']);
  Logger.log('[DisplayName] ' + ldap + ' updated to: ' + displayName.trim());
  return { success: true };
}

function formatDisplayName(ldap) {
  if (!ldap) return '';
  var agent = findRow('Agents', 'LDAP', ldap);
  if (agent && agent['DisplayName'] && agent['DisplayName'].trim() !== '') {
    return agent['DisplayName'].trim();
  }
  return ldap.toLowerCase();
}
function getPreviousMonth() {
  var d = new Date();
  d.setMonth(d.getMonth() - 1);
  return (d.getMonth() + 1) + '/' + d.getFullYear();
}
