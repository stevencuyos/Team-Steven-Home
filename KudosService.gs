// ============================================================
// KudosService.gs — Kudos submission, queue, review
// Play Ops Store
// ============================================================

// ------------------------------------------------------------
// KUDOS ISSUE TYPES
// ------------------------------------------------------------

var KUDOS_ISSUE_TYPES = [
  'Above and Beyond',
  'Empathy & Communication',
  'Technical Excellence',
  'First Contact Resolution',
  'Process Compliance',
  'Teamwork & Collaboration',
  'Customer Advocacy'
];

var KUDOS_CHANNELS = [
  'Chat',
  'Email',
  'Phone',
  'Social',
  'Other'
];

function getKudosFormConfig() {
  return {
    issueTypes: KUDOS_ISSUE_TYPES,
    channels: KUDOS_CHANNELS
  };
}

// ------------------------------------------------------------
// SUBMIT KUDOS (agent)
// ------------------------------------------------------------

function submitKudos(formData, ldap) {
  // Basic validation
  if (!formData.caseId || !formData.channel || !formData.issueType || !formData.whyKudos) {
    return { success: false, error: 'Please fill in all required fields.' };
  }

  if (formData.whyKudos.trim().length < 20) {
    return { success: false, error: 'Please provide more detail (at least 20 characters).' };
  }

  // Check for duplicate case ID submissions by same agent
  var existing = findRows('Kudos', 'LDAP', ldap).find(function(k) {
    return k['CaseID'] === formData.caseId && k['Status'] !== 'Rejected';
  });
  if (existing) {
    return { success: false, error: 'You have already submitted a Kudos for case ' + formData.caseId + '.' };
  }

  var id = generateKudosId();
  var submittedAt = now();

  appendRow('Kudos', {
    ID: id,
    LDAP: ldap,
    CaseID: formData.caseId,
    Channel: formData.channel,
    IssueType: formData.issueType,
    Resolution: formData.resolution || '',
    WhyKudos: formData.whyKudos,
    Status: 'Pending',
    ReviewedBy: '',
    ReviewNote: '',
    SubmittedAt: submittedAt,
    ReviewedAt: ''
  });

  // Notify agent of submission received
  createNotification(ldap, 'kudos_submitted', 'Your Kudos for case ' + formData.caseId + ' has been submitted and is pending review.');

  // Notify managers of new kudos in queue
  notifyManagersNewKudos(ldap, formData.caseId);

  Logger.log('[Kudos Submitted] ' + ldap + ' | Case: ' + formData.caseId + ' | ID: ' + id);
  return { success: true, kudosId: id };
}

// ------------------------------------------------------------
// GET KUDOS QUEUE (manager)
// ------------------------------------------------------------

function getKudosQueue(managedLdaps) {
  var rows = getSheetData('Kudos');
  rows.sort(function(a, b) {
    return new Date(a['SubmittedAt']) - new Date(b['SubmittedAt']); // oldest first
  });

  var pending = rows.filter(function(r) {
    if (managedLdaps && managedLdaps.indexOf(r['LDAP']) === -1) return false;
    return r['Status'] === 'Pending';
  });
  var reviewed = rows.filter(function(r) {
    if (managedLdaps && managedLdaps.indexOf(r['LDAP']) === -1) return false;
    return r['Status'] !== 'Pending';
  }).slice(0, 20); // last 20 reviewed

  return {
    pending: pending.map(formatKudosRow),
    reviewed: reviewed.map(formatKudosRow),
    pendingCount: pending.length
  };
}

function formatKudosRow(k) {
  return {
    id: k['ID'],
    ldap: k['LDAP'],
    displayName: formatDisplayName(k['LDAP']),
    caseId: k['CaseID'],
    channel: k['Channel'],
    issueType: k['IssueType'],
    resolution: k['Resolution'],
    whyKudos: k['WhyKudos'],
    status: k['Status'],
    reviewedBy: k['ReviewedBy'],
    reviewNote: k['ReviewNote'],
    submittedAt: formatDateTime(k['SubmittedAt']),
    reviewedAt: k['ReviewedAt'] ? formatDateTime(k['ReviewedAt']) : null,
    photoUrl: getMomaPhotoUrl(k['LDAP'])
  };
}

// ------------------------------------------------------------
// REVIEW KUDOS (manager)
// ------------------------------------------------------------

function reviewKudos(kudosId, decision, note, reviewedBy) {
  var kudos = findRow('Kudos', 'ID', kudosId);
  if (!kudos) return { success: false, error: 'Kudos not found.' };
  if (kudos['Status'] !== 'Pending') return { success: false, error: 'This Kudos has already been reviewed.' };

  if (decision !== 'Approved' && decision !== 'Rejected') {
    return { success: false, error: 'Decision must be Approved or Rejected.' };
  }

  if (decision === 'Rejected' && (!note || note.trim() === '')) {
    return { success: false, error: 'A rejection note is required.' };
  }

  var reviewedAt = now();
  updateRow('Kudos', 'ID', kudosId, {
    Status: decision,
    ReviewedBy: reviewedBy,
    ReviewNote: note || '',
    ReviewedAt: reviewedAt
  });

  var ldap = kudos['LDAP'];

  if (decision === 'Approved') {
    var points = parseInt(getConfig('PointsKudosValidated')) || 30;
    addPoints(ldap, points, 'Kudos approved: ' + kudosId);

    createNotification(ldap, 'kudos_approved', 'Your Kudos for case ' + kudos['CaseID'] + ' was approved! +' + points + ' pts.');
    sendKudosApprovedEmail(ldap, points, kudos['CaseID']);

    // Badge check
    evaluateBadges(ldap, {
      completionCount: 0,
      isEarly: false,
      daysEarly: 0,
      isAcknowledge: false,
      onTimeStreak: 0
    });

  } else {
    createNotification(ldap, 'kudos_rejected', 'Your Kudos for case ' + kudos['CaseID'] + ' was not approved.' + (note ? ' Note: ' + note : ''));
    sendKudosRejectedEmail(ldap, kudos['CaseID'], note);
  }

  invalidateCache(['leaderboard_agent', 'leaderboard_manager', 'team_analytics']);
  Logger.log('[Kudos Reviewed] ' + kudosId + ' | ' + decision + ' | by ' + reviewedBy);
  return { success: true, decision: decision };
}

// ------------------------------------------------------------
// NOTIFY MANAGERS OF NEW KUDOS
// ------------------------------------------------------------

function notifyManagersNewKudos(agentLdap, caseId) {
  var managers = getSheetData('Managers');
  var notifications = managers.map(function(m) {
    return {
      ldap: m['LDAP'],
      type: 'new_kudos',
      message: formatDisplayName(agentLdap) + ' submitted a Kudos for case ' + caseId + '. Review it in the Kudos Queue.'
    };
  });
  createNotifications(notifications);
}

// ------------------------------------------------------------
// KUDOS STATS (for profile / analytics)
// ------------------------------------------------------------

function getKudosStats(ldap) {
  var all = findRows('Kudos', 'LDAP', ldap);
  var approved = all.filter(function(k) { return k['Status'] === 'Approved'; });
  var pending  = all.filter(function(k) { return k['Status'] === 'Pending'; });
  var rejected = all.filter(function(k) { return k['Status'] === 'Rejected'; });

  // Breakdown by issue type
  var byType = {};
  approved.forEach(function(k) {
    var t = k['IssueType'] || 'Other';
    byType[t] = (byType[t] || 0) + 1;
  });

  return {
    total: all.length,
    approved: approved.length,
    pending: pending.length,
    rejected: rejected.length,
    byType: byType
  };
}

// ------------------------------------------------------------
// TEAM ANALYTICS (manager)
// ------------------------------------------------------------

function getTeamAnalytics(managedLdaps) {
  var cacheKey = 'team_analytics' + (managedLdaps ? '_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, managedLdaps.join(',')).map(function(b) { return (b & 0xFF).toString(16); }).join('') : '');
  var cached = getCached(cacheKey);
  if (cached) return cached;

  var agents = getSheetData('Agents');
  if (managedLdaps) {
    agents = agents.filter(function(a) { return managedLdaps.indexOf(a['LDAP']) !== -1; });
  }
  var now = new Date();
  var monthStart = new Date(now.getFullYear(), now.getMonth(), 1);

  // Load all sheets once
  var allCompletions = getSheetData('Completions');
  var allKudos = getSheetData('Kudos');
  var allDemerits = getSheetData('Demerits');
  var allLeaderboard = getSheetData('Leaderboard');

  // Pre-group data by LDAP to avoid nested filters
  var completionsByLdap = {};
  allCompletions.forEach(function(c) {
    if (!completionsByLdap[c.LDAP]) completionsByLdap[c.LDAP] = [];
    completionsByLdap[c.LDAP].push(c);
  });

  var kudosByLdap = {};
  allKudos.forEach(function(k) {
    if (k.Status === 'Approved') {
      if (!kudosByLdap[k.LDAP]) kudosByLdap[k.LDAP] = [];
      kudosByLdap[k.LDAP].push(k);
    }
  });

  var demeritsByLdap = {};
  allDemerits.forEach(function(d) {
    if (!demeritsByLdap[d.LDAP]) demeritsByLdap[d.LDAP] = [];
    demeritsByLdap[d.LDAP].push(d);
  });

  var lbMap = {};
  allLeaderboard.forEach(function(r) { lbMap[r.LDAP] = r; });

  var stats = agents.map(function(a) {
    var ldap = a['LDAP'];
    if (!ldap) return null;

    var completions = completionsByLdap[ldap] || [];
    var monthlyCompletions = completions.filter(function(c) {
      return c['CompletedAt'] && new Date(c['CompletedAt']) >= monthStart;
    });
    var kudos = kudosByLdap[ldap] || [];
    var demerits = demeritsByLdap[ldap] || [];
    var monthlyDemerits = demerits.filter(function(d) {
      return new Date(d['Timestamp']) >= monthStart;
    });
    var lbRow = lbMap[ldap] || null;

    var doneCompletions = completions.filter(function(c) { return c['CompletedAt']; });
    doneCompletions.sort(function(a, b) { return new Date(b['CompletedAt']) - new Date(a['CompletedAt']); });

    var dName = (a && a.DisplayName) ? a.DisplayName.trim() : ldap.toLowerCase();
    return {
      ldap: ldap,
      displayName: dName,
      photoUrl: getMomaPhotoUrl(ldap),
      tier: lbRow ? lbRow['Tier'] : 'Bronze',
      monthlyPoints: lbRow ? (lbRow['MonthlyPoints'] || 0) : 0,
      allTimePoints: lbRow ? (lbRow['AllTimePoints'] || 0) : 0,
      currentStreak: lbRow ? (lbRow['CurrentStreak'] || 0) : 0,
      totalCompletions: doneCompletions.length,
      monthlyCompletions: monthlyCompletions.length,
      kudosApproved: kudos.length,
      totalDemerits: demerits.length,
      monthlyDemerits: monthlyDemerits.length,
      lastActive: doneCompletions.length > 0 ? formatDate(doneCompletions[0]['CompletedAt']) : 'No activity'
    };
  });

  var totalPoints = stats.reduce(function(s, a) { return s + a.monthlyPoints; }, 0);
  var totalCompletions = stats.reduce(function(s, a) { return s + a.monthlyCompletions; }, 0);
  var totalKudos = stats.reduce(function(s, a) { return s + a.kudosApproved; }, 0);
  var totalDemerits = stats.reduce(function(s, a) { return s + a.monthlyDemerits; }, 0);

  var result = {
    agents: stats,
    summary: {
      totalAgents: agents.length,
      totalMonthlyPoints: totalPoints,
      totalMonthlyCompletions: totalCompletions,
      totalKudosApproved: totalKudos,
      totalMonthlyDemerits: totalDemerits,
      month: now.toLocaleString('default', { month: 'long', year: 'numeric' })
    }
  };

  setCached(cacheKey, result, 300); // 5 minute TTL
  return result;
}
