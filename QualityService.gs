// ============================================================
// QualityService.gs — Aggregations and logic for Quality Audits
// Source: 'QualityAudits' tab in the same spreadsheet
// ============================================================

var QUALITY_SHEET_NAME = 'QualityAudits';

// ── SCHEMA MAPPING ────────────────────────────────────────────────────────

var Q_COLS = {
  CASE_ID: 0,            // A
  ENTITY_GROUP: 1,       // B
  AGENT_LDAP: 2,         // C
  OPENING_CHANNEL: 3,    // D
  REVIEW_DATE: 4,        // E
  REVIEW_WEEK: 5,        // F
  REVIEW_MONTH: 6,       // G
  CASE_DATE: 7,          // H
  CASE_WEEK: 8,          // I
  CASE_MONTH: 9,         // J
  CUSTOMER_CRITICAL: 10, // K
  BUSINESS_CRITICAL: 11, // L
  COMPLIANCE_CRITICAL: 12,// M
  REVIEWER_COMMENTS: 13, // N

  // Customer Critical Parameters (O-W)
  LISTENING: 14,
  PROBING: 15,
  COMPLETE_RESOLUTION: 16,
  TROUBLESHOOTING: 17,
  USER_EXPECTATIONS: 18,
  EMPATHY: 19,
  OWNERSHIP: 20,
  REFUNDS: 21,
  RESPONSIVENESS: 22,

  // Business Critical Parameters (X-AD)
  CONSULTS_ESCALATIONS: 23,
  CASE_DETAILS: 24,
  CATEGORIZATION: 25,
  CSAT_REMINDER: 26,
  CASE_STATE: 27,
  OPENING_CLOSING: 28,
  LANGUAGE_PROFICIENCY: 29,

  // Compliance Critical Parameters (AE-AH)
  AUTHENTICATION: 30,
  GOOGLE_ONLY_INFO: 31,
  PROFESSIONAL_CONDUCT: 32,
  PAYMENT_COMPLAINTS: 33,

  TEAM: 36,              // AK (index 36)
  AGENT_WORKFLOW: 37     // AL (index 37)
};

var Q_TARGETS = {
  CUSTOMER: 95,
  BUSINESS: 90,
  COMPLIANCE: 99.50
};

// ── DATA LOADING ──────────────────────────────────────────────────────────

function getRawQualityData() {
  var cacheKey = 'quality_raw_v1';
  var cache = CacheService.getScriptCache();

  try {
    var chunkCount = cache.get(cacheKey + '_chunks');
    if (chunkCount) {
      var assembled = '';
      for (var c = 0; c < parseInt(chunkCount); c++) {
        var chunk = cache.get(cacheKey + '_chunk_' + c);
        if (!chunk) { assembled = null; break; }
        assembled += chunk;
      }
      if (assembled) return JSON.parse(assembled);
    }
  } catch(e) {
    Logger.log('[Quality] Cache read error: ' + e.message);
  }

  var sheet = getSheet(QUALITY_SHEET_NAME);
  var raw = sheet.getDataRange().getValues();
  if (raw.length < 2) return [];

  var data = raw.slice(1);

  // Cache result (chunked if needed)
  try {
    var serialized = JSON.stringify(data);
    if (serialized.length < 100000) {
      cache.put(cacheKey, serialized, 1800); // 30 min
    } else {
      var chunkSize = 90000;
      var chunks = [];
      for (var ci = 0; ci < serialized.length; ci += chunkSize) {
        chunks.push(serialized.slice(ci, ci + chunkSize));
      }
      chunks.forEach(function(chunk, idx) {
        cache.put(cacheKey + '_chunk_' + idx, chunk, 1800);
      });
      cache.put(cacheKey + '_chunks', String(chunks.length), 1800);
    }
  } catch(e) {
    Logger.log('[Quality] Cache write error: ' + e.message);
  }

  return data;
}

function getAvailableQualityMonths() {
  var rows = getRawQualityData();
  var seen = {};
  rows.forEach(function(r) {
    var month = r[Q_COLS.REVIEW_MONTH];
    if (month) {
      if (month instanceof Date) {
        month = Utilities.formatDate(month, Session.getScriptTimeZone(), 'yyyy-MM');
      }
      seen[month] = true;
    }
  });
  return Object.keys(seen).sort().reverse();
}

function normalizeQualityMonth(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  return String(val).trim();
}

function normalizeLdap(val) {
  if (!val) return '';
  return String(val).trim().toLowerCase().split('@')[0];
}

// ── AGGREGATION ───────────────────────────────────────────────────────────

function aggregateQualityRows(rows) {
  if (rows.length === 0) return null;

  var customerSum = 0, businessSum = 0, complianceSum = 0;

  var params = {};
  var paramGroups = {
    customer: ['LISTENING', 'PROBING', 'COMPLETE_RESOLUTION', 'TROUBLESHOOTING', 'USER_EXPECTATIONS', 'EMPATHY', 'OWNERSHIP', 'REFUNDS', 'RESPONSIVENESS'],
    business: ['CONSULTS_ESCALATIONS', 'CASE_DETAILS', 'CATEGORIZATION', 'CSAT_REMINDER', 'CASE_STATE', 'OPENING_CLOSING', 'LANGUAGE_PROFICIENCY'],
    compliance: ['AUTHENTICATION', 'GOOGLE_ONLY_INFO', 'PROFESSIONAL_CONDUCT', 'PAYMENT_COMPLAINTS']
  };

  var paramCols = [].concat(paramGroups.customer, paramGroups.business, paramGroups.compliance);
  paramCols.forEach(function(p) { params[p] = { yes: 0, total: 0 }; });

  rows.forEach(function(r) {
    customerSum += (parseFloat(r[Q_COLS.CUSTOMER_CRITICAL]) || 0);
    businessSum += (parseFloat(r[Q_COLS.BUSINESS_CRITICAL]) || 0);
    complianceSum += (parseFloat(r[Q_COLS.COMPLIANCE_CRITICAL]) || 0);

    paramCols.forEach(function(p) {
      var val = String(r[Q_COLS[p]]).trim().toLowerCase();
      if (val === 'yes' || val === 'no') {
        params[p].total++;
        if (val === 'yes') params[p].yes++;
      }
    });
  });

  var count = rows.length;
  var paramScores = {};
  paramCols.forEach(function(p) {
    paramScores[p] = params[p].total > 0 ? (params[p].yes / params[p].total) * 100 : null;
  });

  // Grouped params
  var groupedParams = {
    customer: paramGroups.customer.map(p => ({ name: p, score: paramScores[p] })),
    business: paramGroups.business.map(p => ({ name: p, score: paramScores[p] })),
    compliance: paramGroups.compliance.map(p => ({ name: p, score: paramScores[p] }))
  };

  return {
    customer: (customerSum / count) * 100,
    business: (businessSum / count) * 100,
    compliance: (complianceSum / count) * 100,
    count: count,
    params: paramScores,
    groupedParams: groupedParams,
    targets: Q_TARGETS
  };
}

function aggregateTrends(rows) {
  var daily = {};
  var weekly = {};

  rows.forEach(function(r) {
    var dateRaw = r[Q_COLS.REVIEW_DATE];
    var date = (dateRaw instanceof Date) ? Utilities.formatDate(dateRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(dateRaw);
    var week = r[Q_COLS.REVIEW_WEEK];

    [ {obj: daily, key: date}, {obj: weekly, key: week} ].forEach(function(t) {
      if (!t.key) return;
      if (!t.obj[t.key]) t.obj[t.key] = { customer: 0, business: 0, compliance: 0, count: 0 };
      t.obj[t.key].customer += (parseFloat(r[Q_COLS.CUSTOMER_CRITICAL]) || 0);
      t.obj[t.key].business += (parseFloat(r[Q_COLS.BUSINESS_CRITICAL]) || 0);
      t.obj[t.key].compliance += (parseFloat(r[Q_COLS.COMPLIANCE_CRITICAL]) || 0);
      t.obj[t.key].count++;
    });
  });

  var formatTrend = function(obj) {
    return Object.keys(obj).sort().map(function(k) {
      var d = obj[k];
      return {
        label: k,
        customer: (d.customer / d.count) * 100,
        business: (d.business / d.count) * 100,
        compliance: (d.compliance / d.count) * 100,
        avg: ((d.customer + d.business + d.compliance) / (d.count * 3)) * 100
      };
    });
  };

  return { daily: formatTrend(daily), weekly: formatTrend(weekly) };
}

// ── VIEW DATA FETCHERS ────────────────────────────────────────────────────

function getMyQualityData(ldap, month) {
  var cacheKey = 'quality_agent_' + normalizeLdap(ldap) + '_' + month;
  var cached = getCached(cacheKey);
  if (cached) return cached;

  var allRows = getRawQualityData();
  var filtered = allRows.filter(function(r) {
    return normalizeLdap(r[Q_COLS.AGENT_LDAP]) === normalizeLdap(ldap) &&
           normalizeQualityMonth(r[Q_COLS.REVIEW_MONTH]) === month;
  });

  var stats = aggregateQualityRows(filtered);
  var trends = aggregateTrends(filtered);

  var caseLog = filtered.map(function(r) {
    return {
      caseId: r[Q_COLS.CASE_ID],
      reviewDate: formatDate(r[Q_COLS.REVIEW_DATE]),
      customer: r[Q_COLS.CUSTOMER_CRITICAL],
      business: r[Q_COLS.BUSINESS_CRITICAL],
      compliance: r[Q_COLS.COMPLIANCE_CRITICAL],
      comments: r[Q_COLS.REVIEWER_COMMENTS],
      details: {
        customer: {
          LISTENING: r[Q_COLS.LISTENING],
          PROBING: r[Q_COLS.PROBING],
          COMPLETE_RESOLUTION: r[Q_COLS.COMPLETE_RESOLUTION],
          TROUBLESHOOTING: r[Q_COLS.TROUBLESHOOTING],
          USER_EXPECTATIONS: r[Q_COLS.USER_EXPECTATIONS],
          EMPATHY: r[Q_COLS.EMPATHY],
          OWNERSHIP: r[Q_COLS.OWNERSHIP],
          REFUNDS: r[Q_COLS.REFUNDS],
          RESPONSIVENESS: r[Q_COLS.RESPONSIVENESS]
        },
        business: {
          CONSULTS_ESCALATIONS: r[Q_COLS.CONSULTS_ESCALATIONS],
          CASE_DETAILS: r[Q_COLS.CASE_DETAILS],
          CATEGORIZATION: r[Q_COLS.CATEGORIZATION],
          CSAT_REMINDER: r[Q_COLS.CSAT_REMINDER],
          CASE_STATE: r[Q_COLS.CASE_STATE],
          OPENING_CLOSING: r[Q_COLS.OPENING_CLOSING],
          LANGUAGE_PROFICIENCY: r[Q_COLS.LANGUAGE_PROFICIENCY]
        },
        compliance: {
          AUTHENTICATION: r[Q_COLS.AUTHENTICATION],
          GOOGLE_ONLY_INFO: r[Q_COLS.GOOGLE_ONLY_INFO],
          PROFESSIONAL_CONDUCT: r[Q_COLS.PROFESSIONAL_CONDUCT],
          PAYMENT_COMPLAINTS: r[Q_COLS.PAYMENT_COMPLAINTS]
        }
      }
    };
  });

  var result = {
    ldap: ldap,
    month: month,
    stats: stats,
    trends: trends,
    caseLog: caseLog,
    hasData: filtered.length > 0
  };

  setCached(cacheKey, result, 600); // 10 min
  return result;
}

function getTeamQualityData(managerLdap, month) {
  var cacheKey = 'quality_team_' + normalizeLdap(managerLdap) + '_' + month;
  var cached = getCached(cacheKey);
  if (cached) return cached;

  var managedLdaps = getManagedLdaps(managerLdap);
  var allRows = getRawQualityData();

  var teamRows = allRows.filter(function(r) {
    var agentLdap = normalizeLdap(r[Q_COLS.AGENT_LDAP]);
    return (managedLdaps === null || managedLdaps.indexOf(agentLdap) !== -1) &&
           normalizeQualityMonth(r[Q_COLS.REVIEW_MONTH]) === month;
  });

  var teamStats = aggregateQualityRows(teamRows);
  var trends = aggregateTrends(teamRows);

  var agentStats = {};
  var uniqueLdaps = [];
  teamRows.forEach(function(r) {
    var ldap = normalizeLdap(r[Q_COLS.AGENT_LDAP]);
    if (!agentStats[ldap]) {
      agentStats[ldap] = [];
      uniqueLdaps.push(ldap);
    }
    agentStats[ldap].push(r);
  });

  var agents = uniqueLdaps.map(function(ldap) {
    var stats = aggregateQualityRows(agentStats[ldap]);
    return {
      ldap: ldap,
      displayName: formatDisplayName(ldap),
      stats: stats
    };
  }).sort(function(a, b) {
    // Sort by compliance then customer then business? Or just average.
    var avgA = (a.stats.customer + a.stats.business + a.stats.compliance) / 3;
    var avgB = (b.stats.customer + b.stats.business + b.stats.compliance) / 3;
    return avgB - avgA;
  });

  var result = {
    managerLdap: managerLdap,
    month: month,
    stats: teamStats,
    trends: trends,
    agents: agents,
    hasData: teamRows.length > 0
  };

  setCached(cacheKey, result, 600); // 10 min
  return result;
}

function getAllTeamsQualityData(month) {
  var cacheKey = 'quality_allteams_' + month;
  var cached = getCached(cacheKey);
  if (cached) return cached;

  var allRows = getRawQualityData();
  var monthRows = allRows.filter(function(r) {
    return normalizeQualityMonth(r[Q_COLS.REVIEW_MONTH]) === month;
  });

  var siteStats = aggregateQualityRows(monthRows);

  // Group by TeamLead (or Team column if reliable)
  // User said "Use the existing manager/supervisor mapping already established in the Managers and Agents sheets."
  // So I'll fetch all managers/supervisors and their teams.

  var managers = getSheetData('Managers');
  var teams = managers.filter(function(m) { return m.Role === 'manager' || m.Role === 'supervisor'; }).map(function(m) {
    var mLdap = normalizeLdap(m.LDAP);
    var managed = getManagedLdaps(mLdap);
    var teamRows = monthRows.filter(function(r) {
      var agentLdap = normalizeLdap(r[Q_COLS.AGENT_LDAP]);
      return (managed === null || managed.indexOf(agentLdap) !== -1);
    });

    return {
      managerLdap: mLdap,
      teamName: m.Team || ('Team ' + formatDisplayName(mLdap)),
      stats: aggregateQualityRows(teamRows),
      agentCount: managed ? managed.length : 0
    };
  }).filter(function(t) { return t.stats !== null; })
    .sort(function(a, b) {
      var avgA = (a.stats.customer + a.stats.business + a.stats.compliance) / 3;
      var avgB = (b.stats.customer + b.stats.business + b.stats.compliance) / 3;
      return avgB - avgA;
    });

  var result = {
    month: month,
    siteStats: siteStats,
    teams: teams
  };

  setCached(cacheKey, result, 600); // 10 min
  return result;
}
