// ============================================================
// CsatService.gs — Direct read from source CSATdump sheet
// Source: ONEPlayCebu Performance Dashboard
// One row per case — aggregated on-the-fly
// ============================================================

var CSAT_SHEET_ID  = '1wH2AVGJ9jyZJUX1EIvhBkMn-i7gOAE9z9gxB7STjC0g';
var CSAT_DUMP_NAME = 'CSATdump';

// ── HELPERS ───────────────────────────────────────────────────────────────

function buildColMap(headers) {
  var map = {};
  headers.forEach(function(h, i) {
    if (h) map[String(h).trim()] = i;
  });
  return map;
}

function currentCsatMonth() {
  var d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function getPrevCsatMonth(monthStr) {
  var d = new Date(monthStr + '-01');
  d.setMonth(d.getMonth() - 1);
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function parseRowMonth(rawMonth) {
  if (!rawMonth) return '';
  if (rawMonth instanceof Date) {
    return rawMonth.getFullYear() + '-' + String(rawMonth.getMonth() + 1).padStart(2, '00').slice(-2);
  }
  // Handle "2026-04" or "2026-04-26" formats
  var s = String(rawMonth).trim();
  var iso = s.match(/^(\d{4})-(\d{2})/);
  if (iso) return iso[1] + '-' + iso[2];
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  }
  return s;
}

function csatStatus(pct) {
  if (pct === null || pct === undefined || isNaN(pct)) return 'no_data';
  if (pct >= 85) return 'on_target';
  if (pct >= 75) return 'monitor';
  return 'at_risk';
}

function csatDelta(current, prev) {
  if (current === null || prev === null ||
      current === undefined || prev === undefined ||
      isNaN(current) || isNaN(prev)) return null;
  return Math.round((current - prev) * 10) / 10;
}

function safeFloat(val) {
  if (val === null || val === '' || val === 'null' || val === undefined) return null;
  var n = parseFloat(val);
  return isNaN(n) ? null : n;
}

// ── READ AHT DUMP RAW DATA ───────────────────────────────────────────────

function readAhtDumpCached() {
  var cacheKey = 'aht_raw_v1';
  var currentYear = new Date().getFullYear();

  try {
    var cache = CacheService.getScriptCache();

    // Check for chunked cache
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

    // Check single-key cache
    var cachedStr = cache.get(cacheKey);
    if (cachedStr) return JSON.parse(cachedStr);
  } catch(e) {
    Logger.log('[AHT] Cache read error: ' + e.message);
  }

  // Fetch from source
  var ext = SpreadsheetApp.openById(CSAT_SHEET_ID);
  var dump = ext.getSheetByName('AHTDump');
  if (!dump) throw new Error('AHTDump sheet not found.');

  var raw = dump.getDataRange().getValues();
  if (raw.length < 2) return [];

  var headers = raw[0];
  var C = buildColMap(headers);
  var rows = [];

  var C_lower = {};
  for (var k in C) {
    C_lower[k.toLowerCase()] = C[k];
  }

  function getIndex(names) {
    for (var i = 0; i < names.length; i++) {
      var n = names[i].toLowerCase();
      if (C_lower[n] !== undefined) return C_lower[n];
    }
    return undefined;
  }

  var idxMonth = getIndex(['Month', 'MONTH']);
  var idxLdap = getIndex(['LDAP']);
  var idxTc = getIndex(['Team Captain / Manager LDAP', 'Team Capitan']);
  var idxChannel = getIndex(['Channel']);
  var idxSymptom = getIndex(['Symptom / TUI', 'TUI CHECKER']);
  var idxChatAht = getIndex(['Chat AHT']);
  var idxPhoneAht = getIndex(['Phone AHT']);
  var idxSurveyRate = getIndex(['Survey Offer Rate']);
  var idxTotalChats = getIndex(['Number of Total Chats', 'Number of total chats']);
  var idxIncomingPhone = getIndex(['Incoming Phone Calls']);

  for (var i = 1; i < raw.length; i++) {
    var row = raw[i];

    var rawMonth = idxMonth !== undefined ? row[idxMonth] : null;
    var month = parseRowMonth(rawMonth);
    if (!month) continue;

    var yearPart = parseInt(month.split('-')[0]);
    if (yearPart !== currentYear) continue;

    var ldap = idxLdap !== undefined ? String(row[idxLdap] || '').trim().toLowerCase().split('@')[0] : '';
    if (!ldap) continue;

    var tcLdap = (idxTc !== undefined ? row[idxTc] : row[22]);
    tcLdap = String(tcLdap || '').trim().toLowerCase().split('@')[0];

    rows.push({
      month: month,
      ldap: ldap,
      tcLdap: tcLdap,
      channel: idxChannel !== undefined ? String(row[idxChannel] || '').trim().toLowerCase() : '',
      symptom: idxSymptom !== undefined ? String(row[idxSymptom] || '').trim() : String(row[23] || '').trim(),
      chatAht: safeFloat(idxChatAht !== undefined ? row[idxChatAht] : null),
      phoneAht: safeFloat(idxPhoneAht !== undefined ? row[idxPhoneAht] : null),
      surveyOfferRate: safeFloat(idxSurveyRate !== undefined ? row[idxSurveyRate] : null),
      totalChats: safeFloat(idxTotalChats !== undefined ? row[idxTotalChats] : null),
      incomingPhone: safeFloat(idxIncomingPhone !== undefined ? row[idxIncomingPhone] : null)
    });
  }

  // Cache result
  try {
    var cache = CacheService.getScriptCache();
    var serialized = JSON.stringify(rows);
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
    Logger.log('[AHT] Cache write error: ' + e.message);
  }

  return rows;
}

function getAhtMetrics(ahtRows) {
  var chatAhtSeconds = 0, chatVolume = 0;
  var phoneAhtSeconds = 0, phoneVolume = 0;
  var surveyOfferSum = 0, surveyOfferCount = 0;
  var symptoms = {};

  ahtRows.forEach(function(r) {
    if (r.channel === 'chat') {
      var vol = r.totalChats || 0;
      var aht = r.chatAht || 0;
      if (vol > 0 && aht > 0) {
        chatAhtSeconds += (aht * vol);
        chatVolume += vol;
      }
    } else if (r.channel === 'phone') {
      var vol = r.incomingPhone || 0;
      var aht = r.phoneAht || 0;
      if (vol > 0 && aht > 0) {
        phoneAhtSeconds += (aht * vol);
        phoneVolume += vol;
      }
    }

    if (r.surveyOfferRate !== null && r.surveyOfferRate !== undefined) {
      var sor = parseFloat(r.surveyOfferRate);
      if (!isNaN(sor)) {
        surveyOfferSum += (sor > 1 ? sor : sor * 100);
        surveyOfferCount++;
      }
    }

    if (r.symptom) {
      if (!symptoms[r.symptom]) {
        symptoms[r.symptom] = { chatAhtSeconds: 0, chatVol: 0, phoneAhtSeconds: 0, phoneVol: 0, totalCases: 0 };
      }

      var sym = symptoms[r.symptom];
      if (r.channel === 'chat') {
        var vol = r.totalChats || 0;
        var aht = r.chatAht || 0;
        if (vol > 0 && aht > 0) {
          sym.chatAhtSeconds += (aht * vol);
          sym.chatVol += vol;
          sym.totalCases += vol;
        }
      } else if (r.channel === 'phone') {
        var vol = r.incomingPhone || 0;
        var aht = r.phoneAht || 0;
        if (vol > 0 && aht > 0) {
          sym.phoneAhtSeconds += (aht * vol);
          sym.phoneVol += vol;
          sym.totalCases += vol;
        }
      }
    }
  });

  var symptomBreakdown = Object.keys(symptoms).map(function(k) {
    var s = symptoms[k];
    var cAht = s.chatVol > 0 ? Math.round(s.chatAhtSeconds / s.chatVol) : null;
    var pAht = s.phoneVol > 0 ? Math.round(s.phoneAhtSeconds / s.phoneVol) : null;
    var oAht = (s.chatVol + s.phoneVol) > 0 ? Math.round((s.chatAhtSeconds + s.phoneAhtSeconds) / (s.chatVol + s.phoneVol)) : null;
    return {
      symptom: k,
      chatAht: cAht,
      phoneAht: pAht,
      overallAht: oAht,
      caseVolume: s.totalCases
    };
  }).sort(function(a, b) {
    return b.caseVolume - a.caseVolume;
  });

  return {
    surveyOfferRate: surveyOfferCount > 0 ? Math.round((surveyOfferSum / surveyOfferCount) * 10) / 10 : null,
    chatAht: chatVolume > 0 ? Math.round(chatAhtSeconds / chatVolume) : null,
    phoneAht: phoneVolume > 0 ? Math.round(phoneAhtSeconds / phoneVolume) : null,
    symptomAhtBreakdown: symptomBreakdown
  };
}

// ── READ RAW DATA FROM SOURCE ─────────────────────────────────────────────

function readCsatDump() {
  var cacheKey = 'csat_raw_v3';

  try {
    var cache = CacheService.getScriptCache();

    // Check for chunked cache first
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

    // Check single-key cache
    var cachedStr = cache.get(cacheKey);
    if (cachedStr) return JSON.parse(cachedStr);
  } catch(e) {
    Logger.log('[CSAT] Cache read error: ' + e.message);
  }

  // ── Open source spreadsheet directly ──
  var ext  = SpreadsheetApp.openById(CSAT_SHEET_ID);
  var dump = ext.getSheetByName(CSAT_DUMP_NAME);
  if (!dump) throw new Error('CSATdump sheet not found in source spreadsheet.');

  var raw = dump.getDataRange().getValues();
  if (raw.length < 2) return [];

  var headers = raw[0];
  var C       = buildColMap(headers);

  Logger.log('[CSAT] Columns found: ' + Object.keys(C).join(', '));

  var currentYear = new Date().getFullYear();
  var rows = [];

  for (var i = 1; i < raw.length; i++) {
    var row = raw[i];

    // Month — from the "Month" or "Week" column
    // Screenshots show Month = "2026-04", Week = "2026-04-26"
    var rawMonth = C['Month'] !== undefined ? row[C['Month']] : null;
    var month    = parseRowMonth(rawMonth);
    if (!month) continue;

    // Only load current year data to keep payload small
    var yearPart = parseInt(month.split('-')[0]);
    if (yearPart !== currentYear) continue;

    // LDAP — use "Final LDAP" as the agent identifier
    // Falls back to "First LDAP" if Final is blank
    var finalLdap = C['Final LDAP'] !== undefined
      ? String(row[C['Final LDAP']] || '').trim().toLowerCase().split('@')[0]
      : '';
    var firstLdap = C['First LDAP'] !== undefined
      ? String(row[C['First LDAP']] || '').trim().toLowerCase().split('@')[0]
      : '';
    var ldap = finalLdap || firstLdap;
    if (!ldap) continue;

    // TC LDAP
    var tcLdap = C['TC LDAP'] !== undefined
      ? String(row[C['TC LDAP']] || '').trim().toLowerCase().split('@')[0]
      : '';

    // SATISFACTION: 1 = CSAT, 0 = DSAT, null = no survey
    var satRaw = C['SATISFACTION'] !== undefined ? row[C['SATISFACTION']] : null;
    var satisfaction = safeFloat(satRaw);

    // Channel CSAT columns (per-row percentage values)
    var chatPct   = safeFloat(C['Chat CSAT (%)']         !== undefined ? row[C['Chat CSAT (%)']]         : null);
    var chatResp  = safeFloat(C['Chat CSAT Responses']    !== undefined ? row[C['Chat CSAT Responses']]    : null);
    var phonePct  = safeFloat(C['Phone CSAT (%)']         !== undefined ? row[C['Phone CSAT (%)']]         : null);
    var phoneResp = safeFloat(C['Phone CSAT Responses']   !== undefined ? row[C['Phone CSAT Responses']]   : null);
    var emailPct  = safeFloat(C['Email CSAT (%)']         !== undefined ? row[C['Email CSAT (%)']]         : null);
    var emailResp = safeFloat(C['Email CSAT Responses']   !== undefined ? row[C['Email CSAT Responses']]   : null);

    // Quality metrics
    var repeatContact    = safeFloat(C['Repeat Contact Rate (7 Days)'] !== undefined ? row[C['Repeat Contact Rate (7 Days)']] : null);
    var resolutionRate   = safeFloat(C['Cases Resolution Rate']         !== undefined ? row[C['Cases Resolution Rate']]         : null);
    // var surveyOfferRate  = safeFloat(C['Survey Offer Rate']             !== undefined ? row[C['Survey Offer Rate']]             : null);
    var satisfiedResp    = safeFloat(C['Satisfied CSAT Responses']      !== undefined ? row[C['Satisfied CSAT Responses']]      : null);
    var surveyRespCount  = safeFloat(C['Survey Response Count']         !== undefined ? row[C['Survey Response Count']]         : null);
    var caseId = C['Case ID'] !== undefined ? String(row[C['Case ID']] || '').trim() : '';

    rows.push({
      ldap:           ldap,
      tcLdap:         tcLdap,
      month:          month,
      symptom:        C['Symptom']      !== undefined ? String(row[C['Symptom']]      || '').trim() : '',
      channel:        C['Channel']      !== undefined ? String(row[C['Channel']]      || '').trim() : '',
      caseId:       caseId,
      csatScore:      C['CSat Score']   !== undefined ? row[C['CSat Score']]   : null,
      csatComment:    C['CSat Comments']!== undefined ? String(row[C['CSat Comments']] || '').trim() : '',
      satisfaction:   satisfaction,
      chatPct:        chatPct,
      chatResp:       chatResp,
      phonePct:       phonePct,
      phoneResp:      phoneResp,
      emailPct:       emailPct,
      emailResp:      emailResp,
      repeatContact:  repeatContact,
      resolutionRate: resolutionRate,
      // surveyOfferRate:surveyOfferRate,
      satisfiedResp:  satisfiedResp,
      surveyRespCount:surveyRespCount
    });
  }

  Logger.log('[CSAT] Loaded ' + rows.length + ' rows for year ' + currentYear);

  // ── Cache result (chunked if needed) ──
  try {
    var cache      = CacheService.getScriptCache();
    var serialized = JSON.stringify(rows);
    if (serialized.length < 100000) {
      cache.put(cacheKey, serialized, 1800); // 30 min
    } else {
      var chunkSize = 90000;
      var chunks    = [];
      for (var ci = 0; ci < serialized.length; ci += chunkSize) {
        chunks.push(serialized.slice(ci, ci + chunkSize));
      }
      chunks.forEach(function(chunk, idx) {
        cache.put(cacheKey + '_chunk_' + idx, chunk, 1800);
      });
      cache.put(cacheKey + '_chunks', String(chunks.length), 1800);
    }
  } catch(e) {
    Logger.log('[CSAT] Cache write error: ' + e.message);
  }

  return rows;
}

// ── AGGREGATE ROWS ────────────────────────────────────────────────────────

function aggregateRows(rows) {
  var chatNum = 0, chatDen = 0;
  var phoneNum = 0, phoneDen = 0;
  var emailNum = 0, emailDen = 0;
  var overallSat = 0, overallResp = 0;
  var csatCount = 0, dsatCount = 0;
  var dsatThemesObj = {};
  var repeatVals = [], resolutionVals = [], offerVals = [];

  rows.forEach(function(r) {
    // Chat channel
    if (r.chatResp !== null && r.chatResp > 0) {
      chatDen += r.chatResp;
      if (r.chatPct !== null) {
        chatNum += r.chatPct > 1 ? (r.chatPct / 100) * r.chatResp : r.chatPct * r.chatResp;
      }
    }
    // Phone channel
    if (r.phoneResp !== null && r.phoneResp > 0) {
      phoneDen += r.phoneResp;
      if (r.phonePct !== null) {
        phoneNum += r.phonePct > 1 ? (r.phonePct / 100) * r.phoneResp : r.phonePct * r.phoneResp;
      }
    }
    // Email channel
    if (r.emailResp !== null && r.emailResp > 0) {
      emailDen += r.emailResp;
      if (r.emailPct !== null) {
        emailNum += r.emailPct > 1 ? (r.emailPct / 100) * r.emailResp : r.emailPct * r.emailResp;
      }
    }
    // Overall SATISFACTION
    if (r.satisfaction !== null) {
      overallResp++;
      if (r.satisfaction === 1) {
        csatCount++;
        overallSat++;
      } else if (r.satisfaction === 0) {
        dsatCount++;
        if (r.symptom) {
          dsatThemesObj[r.symptom] = (dsatThemesObj[r.symptom] || 0) + 1;
        }
      }
    }
    // Quality metrics
    if (r.repeatContact !== null && r.repeatContact >= 0) {
      repeatVals.push(r.repeatContact > 1 ? r.repeatContact : r.repeatContact * 100);
    }
    if (r.resolutionRate !== null && r.resolutionRate >= 0) {
      resolutionVals.push(r.resolutionRate > 1 ? r.resolutionRate : r.resolutionRate * 100);
    }

  });

  // 1. Convert DSAT Themes object to a sorted array for the UI!
  var themeArr = [];
  for (var k in dsatThemesObj) {
    themeArr.push({ symptom: k, count: dsatThemesObj[k] });
  }
  themeArr.sort(function(a, b) { return b.count - a.count; });

  // 2. Average the quality metric arrays
  function avg(arr) {
    if (!arr || arr.length === 0) return null;
    var sum = arr.reduce(function(a,b){return a+b}, 0);
    return Math.round(sum / arr.length);
  }

  return {
    // Core CSAT Percentages
    overall: pct(overallSat, overallResp),
    chat: pct(chatNum, chatDen),
    phone: pct(phoneNum, phoneDen),
    email: pct(emailNum, emailDen),
    
    // UI Dashboard Variables
    overallResponses: overallResp,
    chatResponses: chatDen,
    phoneResponses: phoneDen,
    emailResponses: emailDen,
    csatCount: csatCount,
    dsatCount: dsatCount,
    totalSurveyed: overallResp,
    totalCases: overallResp,
    dsatThemes: themeArr,
    repeatContactRate: avg(repeatVals),
    casesResolutionRate: avg(resolutionVals),

    
    // Gamification Engine Variables
    chatVol: chatDen,
    phoneVol: phoneDen,
    emailVol: emailDen,
    totalVol: overallResp
  };
}

// Helper function SAFELY PLACED OUTSIDE so other endpoints can access it!
function pct(num, den) {
  if (!den || den === 0) return null;
  return Math.round((num / den) * 1000) / 10;
}

// ── AVAILABLE MONTHS ──────────────────────────────────────────────────────

function getAvailableCsatMonths() {
  var rows = readCsatDump();
  var seen = {};
  rows.forEach(function(r) { if (r.month) seen[r.month] = true; });
  return Object.keys(seen).sort().reverse();
}

// ── MY CSAT ───────────────────────────────────────────────────────────────

function getMyCsatData(ldap, selectedMonth) {
  var allRows  = readCsatDump();
  var allAhtRows = readAhtDumpCached();
  var curMonth = selectedMonth || currentCsatMonth();
  var prvMonth = getPrevCsatMonth(curMonth);

  var curRows = allRows.filter(function(r) { return r.ldap === ldap && r.month === curMonth; });
  var prvRows = allRows.filter(function(r) { return r.ldap === ldap && r.month === prvMonth; });

  if (curRows.length === 0) return { hasData: false, month: curMonth };

  var tcLdap = curRows[0].tcLdap;
  var cur    = aggregateRows(curRows);
  var prv    = aggregateRows(prvRows);

  var curAhtRows = allAhtRows.filter(function(r) { return r.ldap === ldap && r.month === curMonth; });
  var ahtMetrics = getAhtMetrics(curAhtRows);

  // Team comparison
  var teamRows = allRows.filter(function(r) { return r.tcLdap === tcLdap && r.month === curMonth; });
  var teamAgg  = aggregateRows(teamRows);

  // Rank within team
  var agentLdaps = [];
  teamRows.forEach(function(r) {
    if (agentLdaps.indexOf(r.ldap) === -1) agentLdaps.push(r.ldap);
  });
  var ranked = agentLdaps.map(function(l) {
    var lRows = teamRows.filter(function(r) { return r.ldap === l; });
    return { ldap: l, pct: aggregateRows(lRows).overall };
  }).filter(function(x) { return x.pct !== null; })
    .sort(function(a, b) { return b.pct - a.pct; });
  var myRank = ranked.findIndex(function(x) { return x.ldap === ldap; }) + 1 || null;

  // Month trend (all months this year for this agent)
  var monthMap = {};
  allRows.filter(function(r) { return r.ldap === ldap; }).forEach(function(r) {
    if (!monthMap[r.month]) monthMap[r.month] = [];
    monthMap[r.month].push(r);
  });
  var trendData = Object.keys(monthMap).sort().map(function(m) {
    return { month: m, overall: aggregateRows(monthMap[m]).overall };
  });

  // Case log — only surveyed cases
  var caseLog = curRows
    .filter(function(r) {
      // Only rows that received a survey response (satisfaction 1=CSAT, 0=DSAT)
      return r.satisfaction !== null;
    })
    .map(function(r) {
      return {
        caseId:       r.caseId,
        channel:      r.channel,
        symptom:      r.symptom,
        csatScore:    r.csatScore,
        csatComment:  r.csatComment,
        satisfaction: r.satisfaction
      };
    });

  return {
    hasData:          true,
    month:            curMonth,
    ldap:             ldap,
    displayName:      formatDisplayName(ldap),
    tcLdap:           tcLdap,
    overall:          cur.overall,
    chat:             cur.chat,    chatResponses:  cur.chatResponses,
    phone:            cur.phone,   phoneResponses: cur.phoneResponses,
    email:            cur.email,   emailResponses: cur.emailResponses,
    overallResponses: cur.overallResponses,
    csatCount:        cur.csatCount,
    dsatCount:        cur.dsatCount,
    totalSurveyed:    cur.totalSurveyed,
    totalCases:       cur.totalCases,
    dsatThemes:       cur.dsatThemes,
    status:           csatStatus(cur.overall),
    casesResolutionRate: cur.casesResolutionRate,
    repeatContactRate:   cur.repeatContactRate,
    surveyOfferRate:     ahtMetrics.surveyOfferRate,
    chatAht:             ahtMetrics.chatAht,
    phoneAht:            ahtMetrics.phoneAht,
    symptomAhtBreakdown: ahtMetrics.symptomAhtBreakdown,
    teamRank:         myRank,
    teamSize:         ranked.length,
    trendData:        trendData,
    vsTeam: {
      overall: teamAgg.overall,
      chat:    teamAgg.chat,
      phone:   teamAgg.phone,
      email:   teamAgg.email
    },
    delta: {
      overall: csatDelta(cur.overall, prv.overall),
      chat:    csatDelta(cur.chat,    prv.chat),
      phone:   csatDelta(cur.phone,   prv.phone),
      email:   csatDelta(cur.email,   prv.email)
    },
    caseLog: caseLog
  };
}

// ── TEAM CSAT ─────────────────────────────────────────────────────────────

function getTeamCsatData(managerLdap, selectedMonth) {
  var allRows  = readCsatDump();
  var allAhtRows = readAhtDumpCached();
  var curMonth = selectedMonth || currentCsatMonth();
  var prvMonth = getPrevCsatMonth(curMonth);

  var curRows = allRows.filter(function(r) { return r.tcLdap === managerLdap && r.month === curMonth; });
  var prvRows = allRows.filter(function(r) { return r.tcLdap === managerLdap && r.month === prvMonth; });

  if (curRows.length === 0) {
    return { month: curMonth, overall: null, agents: [], atRiskCount: 0,
             csatCount: 0, dsatCount: 0, totalSurveyed: 0, totalCases: 0,
             chatResponses: 0, phoneResponses: 0, emailResponses: 0,
             chat: null, phone: null, email: null, surveyOfferRate: null, chatAht: null, phoneAht: null,
             topDsatThemes: [], delta: { overall: null, chat: null, phone: null, email: null } };
  }

  var cur = aggregateRows(curRows);
  var prv = aggregateRows(prvRows);
  var curAhtRows = allAhtRows.filter(function(r) { return r.tcLdap === managerLdap && r.month === curMonth; });
  var teamAhtMetrics = getAhtMetrics(curAhtRows);

  // Per-agent breakdown
  var agentLdaps = [];
  curRows.forEach(function(r) {
    if (agentLdaps.indexOf(r.ldap) === -1) agentLdaps.push(r.ldap);
  });

  var agents = agentLdaps.map(function(l) {
    var lRows    = curRows.filter(function(r) { return r.ldap === l; });
    var lAgg     = aggregateRows(lRows);
    var lPrvRows = prvRows.filter(function(r) { return r.ldap === l; });
    var lPrv     = aggregateRows(lPrvRows);
    var lAhtRows = curAhtRows.filter(function(r) { return r.ldap === l; });
    var lAhtMetrics = getAhtMetrics(lAhtRows);

    // Trend for this agent
    var monthMap = {};
    allRows.filter(function(r) { return r.ldap === l; }).forEach(function(r) {
      if (!monthMap[r.month]) monthMap[r.month] = [];
      monthMap[r.month].push(r);
    });
    var trendData = Object.keys(monthMap).sort().map(function(m) {
      return { month: m, overall: aggregateRows(monthMap[m]).overall };
    });

    return {
      ldap:            l,
      displayName:     formatDisplayName(l),
      overall:         lAgg.overall,
      chat:            lAgg.chat,    chatResponses:  lAgg.chatResponses,
      phone:           lAgg.phone,   phoneResponses: lAgg.phoneResponses,
      email:           lAgg.email,   emailResponses: lAgg.emailResponses,
      csatCount:       lAgg.csatCount,
      dsatCount:       lAgg.dsatCount,
      totalSurveyed:   lAgg.totalSurveyed,
      totalCases:      lAgg.totalCases,
      surveyOfferRate: lAhtMetrics.surveyOfferRate,
      chatAht:         lAhtMetrics.chatAht,
      phoneAht:        lAhtMetrics.phoneAht,
      resolutionRate:  lAgg.casesResolutionRate,
      repeatContact:   lAgg.repeatContactRate,
      dsatThemes:      lAgg.dsatThemes,
      delta:           csatDelta(lAgg.overall, lPrv.overall),
      trendData:       trendData,
      status:          csatStatus(lAgg.overall),
      caseLog:         curRows.filter(function(r) { return r.ldap === l && r.satisfaction !== null; }).map(function(r) {
        return {
          caseId:       r.caseId,
          channel:      r.channel,
          symptom:      r.symptom,
          csatScore:    r.csatScore,
          csatComment:  r.csatComment,
          satisfaction: r.satisfaction
        };
      })
    };
  }).sort(function(a, b) {
    if (a.overall === null) return 1;
    if (b.overall === null) return -1;
    return b.overall - a.overall;
  });

  return {
    month:            curMonth,
    overall:          cur.overall,
    chat:             cur.chat,    chatResponses:  cur.chatResponses,
    phone:            cur.phone,   phoneResponses: cur.phoneResponses,
    email:            cur.email,   emailResponses: cur.emailResponses,
    overallResponses: cur.overallResponses,
    csatCount:        cur.csatCount,
    dsatCount:        cur.dsatCount,
    totalSurveyed:    cur.totalSurveyed,
    totalCases:       cur.totalCases,
    surveyOfferRate:  teamAhtMetrics.surveyOfferRate,
    chatAht:          teamAhtMetrics.chatAht,
    phoneAht:         teamAhtMetrics.phoneAht,
    atRiskCount:      agents.filter(function(a) { return a.status === 'at_risk'; }).length,
    topDsatThemes:    cur.dsatThemes,
    delta: {
      overall: csatDelta(cur.overall, prv.overall),
      chat:    csatDelta(cur.chat,    prv.chat),
      phone:   csatDelta(cur.phone,   prv.phone),
      email:   csatDelta(cur.email,   prv.email)
    },
    agents: agents
  };
}

// ── ALL TEAMS ─────────────────────────────────────────────────────────────

function getAllTeamsCsatData(selectedMonth) {
  var allRows  = readCsatDump();
  var curMonth = selectedMonth || currentCsatMonth();
  var prvMonth = getPrevCsatMonth(curMonth);

  var curRows = allRows.filter(function(r) { return r.month === curMonth; });
  var prvRows = allRows.filter(function(r) { return r.month === prvMonth; });

  var tcLdaps = [];
  curRows.forEach(function(r) {
    if (tcLdaps.indexOf(r.tcLdap) === -1) tcLdaps.push(r.tcLdap);
  });

  var teams = tcLdaps.filter(function(tc) { return tc; }).map(function(tc) {
    var tCur = curRows.filter(function(r) { return r.tcLdap === tc; });
    var tPrv = prvRows.filter(function(r) { return r.tcLdap === tc; });
    var tAgg = aggregateRows(tCur);
    var pAgg = aggregateRows(tPrv);

    var uniqueAgents = [];
    tCur.forEach(function(r) {
      if (uniqueAgents.indexOf(r.ldap) === -1) uniqueAgents.push(r.ldap);
    });

    return {
      tcLdap:        tc,
      teamName:      'Team ' + (formatDisplayName(tc) || tc),
      agentCount:    uniqueAgents.length,
      overall:       tAgg.overall,
      chat:          tAgg.chat,
      phone:         tAgg.phone,
      email:         tAgg.email,
      csatCount:     tAgg.csatCount,
      dsatCount:     tAgg.dsatCount,
      totalSurveyed: tAgg.totalSurveyed,
      status:        csatStatus(tAgg.overall),
      delta:         csatDelta(tAgg.overall, pAgg.overall)
    };
  }).sort(function(a, b) {
    if (a.overall === null) return 1;
    if (b.overall === null) return -1;
    return b.overall - a.overall;
  });

  var siteAgg = aggregateRows(curRows);
  var sitePrv = aggregateRows(prvRows);

  return {
    month:         curMonth,
    siteOverall:   siteAgg.overall,
    siteChat:      siteAgg.chat,
    sitePhone:     siteAgg.phone,
    siteEmail:     siteAgg.email,
    siteDelta:     csatDelta(siteAgg.overall, sitePrv.overall),
    totalSurveyed: siteAgg.totalSurveyed,
    teamCount:     teams.length,
    atRiskTeams:   teams.filter(function(t) { return t.status === 'at_risk'; }).length,
    teams:         teams
  };
}

// ── CACHE MANAGEMENT ──────────────────────────────────────────────────────

function clearCsatCache() {
  var cache    = CacheService.getScriptCache();
  var cacheKey = 'csat_raw_v3';
  var ahtCacheKey = 'aht_raw_v1';
  cache.remove(cacheKey);
  var chunkCount = cache.get(cacheKey + '_chunks');
  if (chunkCount) {
    for (var i = 0; i < parseInt(chunkCount); i++) {
      cache.remove(cacheKey + '_chunk_' + i);
    }
    cache.remove(cacheKey + '_chunks');
  }

  cache.remove(ahtCacheKey);
  var ahtChunkCount = cache.get(ahtCacheKey + '_chunks');
  if (ahtChunkCount) {
    for (var j = 0; j < parseInt(ahtChunkCount); j++) {
      cache.remove(ahtCacheKey + '_chunk_' + j);
    }
    cache.remove(ahtCacheKey + '_chunks');
  }
  // Also clear old cache keys
  cache.remove('csat_raw_dump');
  cache.remove('csat_summary_v2');
  Logger.log('[CSAT] All caches cleared at ' + new Date());
}

// Called via clientRunCsatAggregation in Code.gs
function runCsatAggregation() {
  clearCsatCache();
  readCsatDump(); // pre-warm fresh from source
  readAhtDumpCached(); // pre-warm AHT
  Logger.log('[CSAT] Re-read from source CSATdump and AHTDump and cached at ' + new Date());
}

// Test function — run from Apps Script editor
function testCsatService() {
  clearCsatCache();
  var rows = readCsatDump();
  Logger.log('[CSAT] Total rows: ' + rows.length);
  if (rows.length > 0) {
    Logger.log('[CSAT] Sample row: ' + JSON.stringify(rows[0]));
    Logger.log('[CSAT] Months found: ' + getAvailableCsatMonths().join(', '));
  }
}
