// ============================================================
// ScheduleService.gs — FIXED VERSION
// Changes:
//   1. Added missing readShiftsSheet() function with caching
//   2. Fixed SHIFTS_HEADER_ROW to correctly target the date/LDAP header row
//   3. Fixed DOW row detection — now looks for the row ABOVE the header
//      that contains day abbreviations (Thu, Fri, Sat…)
//   4. Defensive LDAP matching with trim + toLowerCase on both sides
//   5. Month string parsing handles "May 2026" format correctly
// ============================================================

var SHIFTS_SS_ID   = '1YnQun0_ANNmteT2anXtOD_VIYk3X_gihHMGPPdh3OSc';
var SHIFTS_GID     = '1772105674';
var BREAKS_SS_ID   = '1bfKu5rYixmQy3v-wEH86G6h_qtxOCvm8-gmwmazFQk4';
var BREAKS_GID     = '1464763579';

// ── ROW INDICES (0-based) ──────────────────────────────────
// From the screenshot:
//   Row 10 (idx  9): blank / title area
//   Row 11 (idx 10): blank
//   Row 12 (idx 11): Thu Apr-30, Fri May-1, Sat May-2 … ← DATE LABEL row
//   Row 13 (idx 12): 4:00, 4:00, 4:00 …                ← sample shift times (not header)
//   Row 14 (idx 13): Thu Apr-30, Fri May-1 …            ← SECOND date row (duplicate?)
//   Row 15 (idx 14): EMP ID | LDAP | Workflow | Channel | Apr-30 | May-1 … ← HEADER ROW
//   Row 16+         : agent data rows
//
// NOTE: If your sheet has a DIFFERENT structure, change SHIFTS_HEADER_ROW to
//       the 0-based index of the row that contains "LDAP" in the LDAP column.
var SHIFTS_HEADER_ROW = 14; // row 15 in sheet = index 14

// Cache TTL (seconds) — 15 minutes
var SCHEDULE_CACHE_TTL = 900;


// ============================================================
// ============================================================
function readShiftsSheet() {
  var cache  = CacheService.getScriptCache();
  var cached = cache.get('shifts_display_v2');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var ss    = SpreadsheetApp.openById(SHIFTS_SS_ID);
  var sheet = getSheetByGid(ss, SHIFTS_GID);
  if (!sheet) throw new Error('Shifts sheet not found (GID: ' + SHIFTS_GID + ')');

  var data = sheet.getDataRange().getDisplayValues();

  try {
    var s = JSON.stringify(data);
    // Only cache if under Apps Script's 100 KB per-key limit
    if (s.length < 100000) cache.put('shifts_display_v2', s, SCHEDULE_CACHE_TTL);
  } catch(e) {
    Logger.log('[ScheduleService] readShiftsSheet cache write failed: ' + e.message);
  }
  return data;
}


// ============================================================
// PUBLIC — called from Code.gs client functions
// ============================================================

function getAgentScheduleData(ldap, month, targetLdap) {
  var effectiveLdap = (targetLdap || ldap || '').trim().toLowerCase();
  try {
    var shifts  = readShiftsForAgent(effectiveLdap, month);
    var breaks  = readBreaksForAgent(effectiveLdap, month);
    var today   = getTodayString();
    return {
      success:    true,
      ldap:       effectiveLdap,
      month:      shifts.month,
      shiftDays:  shifts.days,
      breakDays:  breaks.byDate,
      today:      today,
      summary:    buildAgentSummary(shifts.days, breaks.byDate)
    };
  } catch(e) {
    Logger.log('[ScheduleService] getAgentScheduleData error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function getTeamScheduleData(managerLdap, dateKey) {
  try {
    var managedLdaps = getManagedLdaps(managerLdap);
    var agents       = getAllAgents(managedLdaps);
    var date         = dateKey || getTodayString();

    var shiftsByLdap = getTeamShiftsMap(managedLdaps);
    var breaksByLdap = getTeamBreaksMap(managedLdaps);

    var result = agents.map(function(agent) {
      var agentLdap   = agent.ldap.toLowerCase();
      var agentShifts = shiftsByLdap[agentLdap] || [];
      var shiftInfo   = agentShifts.find(function(d) { return d.dateKey === date; }) || { type: 'OFF' };
      var breakMap    = breaksByLdap[agentLdap] || {};
      var breakSlots  = breakMap[date] || [];

      return {
        ldap:        agent.ldap,
        displayName: agent.displayName,
        channel:     agent.channel,
        workflow:    agent.workgroup || agent.workflow || '',
        shiftStart:  shiftInfo.shiftStart || null,
        shiftEnd:    shiftInfo.shiftEnd   || null,
        shiftType:   shiftInfo.type       || 'OFF',
        breaks:      breakSlots,
        hasRemarks:  breakSlots.some(function(b) {
          return b.status === 'Check RTA Remarks' || b.status === 'Check RTA';
        })
      };
    });

    result.sort(function(a, b) {
      var aOff = (a.shiftType === 'OFF' || a.shiftType === 'VL' || a.shiftType === 'MVL');
      var bOff = (b.shiftType === 'OFF' || b.shiftType === 'VL' || b.shiftType === 'MVL');
      if (aOff && !bOff) return 1;
      if (!aOff && bOff) return -1;
      if (!a.shiftStart) return 1;
      if (!b.shiftStart) return -1;
      return a.shiftStart.localeCompare(b.shiftStart);
    });

    return { success: true, date: date, agents: result };
  } catch(e) {
    Logger.log('[ScheduleService] getTeamScheduleData error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function getAvailableScheduleMonths() {
  try {
    var raw     = readShiftsSheet();
    var headers = raw[SHIFTS_HEADER_ROW];
    var months  = {};
    headers.forEach(function(h) {
      var m = String(h).trim().match(/^([A-Za-z]+)-(\d+)$/);
      if (m) months[m[1]] = true;
    });
    return { success: true, months: Object.keys(months) };
  } catch(e) {
    Logger.log('[ScheduleService] getAvailableScheduleMonths error: ' + e.message);
    return { success: false, error: e.message };
  }
}


// ============================================================
// SHIFT READING
// ============================================================

function readShiftsForAgent(ldap, monthLabel) {
  var raw     = readShiftsSheet();
  var headers = raw[SHIFTS_HEADER_ROW];

  var ldapCol = -1;
  headers.forEach(function(h, i) {
    if (String(h).trim().toUpperCase() === 'LDAP') ldapCol = i;
  });
  if (ldapCol === -1) {
    // Fallback: scan all rows near the header for LDAP
    for (var scan = Math.max(0, SHIFTS_HEADER_ROW - 3); scan <= SHIFTS_HEADER_ROW + 1; scan++) {
      if (!raw[scan]) continue;
      var found = raw[scan].findIndex(function(h) { return String(h).trim().toUpperCase() === 'LDAP'; });
      if (found !== -1) { ldapCol = found; break; }
    }
  }
  if (ldapCol === -1) throw new Error('LDAP column not found near row ' + (SHIFTS_HEADER_ROW + 1));

  // ── Parse month filter ─────────────────────────────────────
  // monthLabel arrives as "May 2026" from the dropdown
  var monthPrefix = monthLabel ? monthLabel.split(' ')[0].toLowerCase() : null;

  var dateCols = [];
  headers.forEach(function(h, idx) {
    var str = String(h).trim();
    var m   = str.match(/^([A-Za-z]+)-(\d+)$/);
    if (m) {
      var monthMatches = !monthPrefix || m[1].toLowerCase() === monthPrefix;
      if (monthMatches) {
        dateCols.push({ idx: idx, month: m[1], day: parseInt(m[2]), label: str });
      }
    }
  });

  var dowRow    = null;
  var dowNames  = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
  for (var i = SHIFTS_HEADER_ROW - 1; i >= Math.max(0, SHIFTS_HEADER_ROW - 5); i--) {
    var checkRow = raw[i];
    if (!checkRow) continue;
    var hitCount = checkRow.filter(function(c) {
      return dowNames.indexOf(String(c).trim()) !== -1;
    }).length;
    // Require at least 3 DOW matches to avoid false positives
    if (hitCount >= 3) { dowRow = checkRow; break; }
  }

  // ── Find agent row ─────────────────────────────────────────
  var targetLdap = String(ldap).trim().toLowerCase();
  var agentRow   = null;
  for (var r = SHIFTS_HEADER_ROW + 1; r < raw.length; r++) {
    var rowLdap = String(raw[r][ldapCol] || '').trim().toLowerCase();
    if (rowLdap === targetLdap) { agentRow = raw[r]; break; }
  }

  if (!agentRow) {
    Logger.log('[ScheduleService] Agent not found: ' + targetLdap);
    return { month: monthLabel || '', days: [] };
  }

  var days = dateCols.map(function(dc) {
    var val   = String(agentRow[dc.idx] || '').trim();
    var dow   = dowRow ? String(dowRow[dc.idx] || '').trim() : '';
    var valUp = val.toUpperCase();

    var type = 'WORK';
    if (!val || valUp === 'OFF' || valUp === 'RD') type = 'OFF';
    else if (valUp === 'VL')                       type = 'VL';
    else if (valUp === 'MVL' || valUp === 'MWL')   type = 'MVL';

    var shiftStart = null, shiftEnd = null;
    if (type === 'WORK' && val) {
      shiftStart = normalizeTime(val);
      if (shiftStart) shiftEnd = addHours(shiftStart, 9);
    }

    return {
      dateKey:    dc.label,
      dateLabel:  dc.month + ' ' + dc.day,
      day:        dc.day,
      dayOfWeek:  dow,
      type:       type,
      shiftStart: shiftStart,
      shiftEnd:   shiftEnd
    };
  });

  var month = dateCols.length > 0 ? dateCols[0].month + ' 2026' : (monthLabel || '');
  return { month: month, days: days };
}


// ============================================================
// BREAK READING
// ============================================================

function readBreaksForAgent(ldap, monthLabel) {
  var raw = readBreaksSheet();
  if (!raw || raw.length < 2) return { byDate: {} };

  var headerRowIdx = 0;
  for (var i = 0; i < Math.min(raw.length, 10); i++) {
    if (raw[i].some(function(c) { return String(c).trim().toUpperCase() === 'LDAP'; })) {
      headerRowIdx = i;
      break;
    }
  }

  var headers    = raw[headerRowIdx].map(function(h) { return String(h).trim().toUpperCase(); });
  var ldapIdx    = headers.indexOf('LDAP');
  var timeIdx    = headers.indexOf('TIME');
  var rtaIdx     = headers.findIndex(function(h) { return h === 'RTA APPROVAL STATUS' || (h.includes('RTA') && h.includes('APPROVAL')); });
  var remarksIdx = headers.findIndex(function(h) { return h === 'TIME REMARKS' || (h.includes('TIME') && h.includes('REMARKS')); });
  var sosIdx     = headers.indexOf('SOS');
  var eosIdx     = headers.indexOf('EOS');

  if (ldapIdx === -1 || timeIdx === -1) return { byDate: {} };

  var byDate          = {};
  var monthNames      = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var targetLdap      = String(ldap).trim().toLowerCase();
  var targetMonthPfx  = monthLabel ? monthLabel.split(' ')[0].toLowerCase() : null;

  for (var r = headerRowIdx + 1; r < raw.length; r++) {
    var row     = raw[r];
    var rowLdap = String(row[ldapIdx] || '').trim().toLowerCase();
    if (rowLdap !== targetLdap) continue;

    var timeVal = String(row[timeIdx] || '').trim();
    if (!timeVal) continue;

    var dateObj = new Date(timeVal);
    if (isNaN(dateObj.getTime())) continue;

    var monthAbbr = monthNames[dateObj.getMonth()];
    if (targetMonthPfx && monthAbbr.toLowerCase() !== targetMonthPfx) continue;

    var dateKey    = monthAbbr + '-' + dateObj.getDate();
    var schedTime  = formatTime12h(dateObj);
    var actualTime = schedTime;
    var rtaVal     = rtaIdx     !== -1 ? String(row[rtaIdx]     || '').trim() : '';
    var remarksVal = remarksIdx !== -1 ? String(row[remarksIdx] || '').trim() : '';
    var sosVal     = sosIdx     !== -1 ? String(row[sosIdx]     || '').trim() : null;
    var eosVal     = eosIdx     !== -1 ? String(row[eosIdx]     || '').trim() : null;

    var status = 'No result yet';
    if (rtaVal) {
      var rtaLower = rtaVal.toLowerCase();
      if (rtaLower.includes('adhere')) {
        status = 'Adhere';
      } else if (rtaLower.includes('check') || rtaLower.includes('rta')) {
        status = 'Check RTA';
        if (remarksVal) {
          var rd = new Date(remarksVal);
          actualTime = !isNaN(rd.getTime()) ? formatTime12h(rd) : remarksVal;
        }
      }
    }

    if (!byDate[dateKey]) byDate[dateKey] = [];
    byDate[dateKey].push({
      schedTime:   schedTime,
      actualTime:  actualTime,
      status:      status,
      timeRemarks: remarksVal,
      sos:         sosVal,
      eos:         eosVal
    });
  }

  return { byDate: byDate };
}

function readBreaksSheet() {
  var cache  = CacheService.getScriptCache();
  var cached = cache.get('breaks_display_v1');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var ss    = SpreadsheetApp.openById(BREAKS_SS_ID);
  var sheet = getSheetByGid(ss, BREAKS_GID);
  if (!sheet) throw new Error('Breaks sheet not found (GID: ' + BREAKS_GID + ')');

  var data = sheet.getDataRange().getDisplayValues();

  try {
    var s = JSON.stringify(data);
    if (s.length < 100000) cache.put('breaks_display_v1', s, SCHEDULE_CACHE_TTL);
  } catch(e) {}
  return data;
}


// ============================================================
// SUMMARY BUILDER
// ============================================================

function buildAgentSummary(days, breaksByDate) {
  var workDays = 0, offDays = 0, vlDays = 0, remainingWork = 0;
  var today    = new Date();

  days.forEach(function(d) {
    if (d.type === 'WORK')                        workDays++;
    else if (d.type === 'OFF')                    offDays++;
    else if (d.type === 'VL' || d.type === 'MVL') vlDays++;
    if (d.type === 'WORK' && d.day >= today.getDate()) remainingWork++;
  });

  var todayKey    = getTodayString();
  var todayBreaks = breaksByDate[todayKey] || [];
  var todayAdhere = todayBreaks.filter(function(b) { return b.status === 'Adhere'; }).length;

  return {
    workDays:        workDays,
    offDays:         offDays,
    vlDays:          vlDays,
    remainingWork:   remainingWork,
    todayBreakCount: todayBreaks.length,
    todayAdhere:     todayAdhere
  };
}


// ============================================================
// HELPERS
// ============================================================

function getSheetByGid(ss, gid) {
  var gidInt = parseInt(gid);
  var sheets  = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gidInt) return sheets[i];
  }
  return null;
}

function getTodayString() {
  var d      = new Date();
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return months[d.getMonth()] + '-' + d.getDate();
}

function formatTime12h(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return String(d);
  var h = d.getHours(), m = d.getMinutes();
  var ampm = h >= 12 ? 'PM' : 'AM';
  h = h % 12 || 12;
  return h + ':' + (m < 10 ? '0' : '') + m + ' ' + ampm;
}

function normalizeTime(str) {
  if (!str) return null;
  str = String(str).trim();
  if (str.match(/\d{1,2}:\d{2}\s*(AM|PM)/i)) return str;

  var m = str.match(/^(\d{1,2}):(\d{2})$/);
  if (m) {
    var h   = parseInt(m[1]);
    var min = m[2];
    var sfx = h >= 12 ? 'PM' : 'AM';
    h = h % 12 || 12;
    return h + ':' + min + ' ' + sfx;
  }

  // Plain hour number, e.g. "5" → "5:00 AM/PM"
  var h2 = parseInt(str);
  if (!isNaN(h2)) {
    var sfx2 = h2 >= 12 ? 'PM' : 'AM';
    h2 = h2 % 12 || 12;
    return h2 + ':00 ' + sfx2;
  }

  return str;
}

function addHours(timeStr, hours) {
  if (!timeStr) return null;
  var d = new Date('2000/01/01 ' + timeStr);
  if (isNaN(d.getTime())) return null;
  d.setHours(d.getHours() + hours);
  return formatTime12h(d);
}

function invalidateScheduleCache() {
  var cache = CacheService.getScriptCache();
  cache.remove('shifts_display_v2');
  cache.remove('breaks_display_v1');
}


// ============================================================
// TEAM OPTIMIZATION HELPERS
// ============================================================

function getTeamShiftsMap(ldaps) {
  var raw     = readShiftsSheet();
  var headers = raw[SHIFTS_HEADER_ROW];
  var ldapCol = headers.findIndex(function(h) { return String(h).trim().toUpperCase() === 'LDAP'; });

  var dateCols = [];
  headers.forEach(function(h, idx) {
    var str = String(h).trim();
    var m   = str.match(/^([A-Za-z]+)-(\d+)$/);
    if (m) dateCols.push({ idx: idx, label: str });
  });

  var map = {};
  if (ldaps) ldaps.forEach(function(l) { map[l.toLowerCase()] = []; });

  for (var r = SHIFTS_HEADER_ROW + 1; r < raw.length; r++) {
    var rowLdap = String(raw[r][ldapCol] || '').trim().toLowerCase();
    if (!ldaps || map.hasOwnProperty(rowLdap)) {
      var agentRow = raw[r];
      map[rowLdap] = dateCols.map(function(dc) {
        var val   = String(agentRow[dc.idx] || '').trim();
        var valUp = val.toUpperCase();
        var type  = 'WORK';
        if (!val || valUp === 'OFF' || valUp === 'RD') type = 'OFF';
        else if (valUp === 'VL')                       type = 'VL';
        else if (valUp === 'MVL' || valUp === 'MWL')   type = 'MVL';

        var shiftStart = null, shiftEnd = null;
        if (type === 'WORK' && val) {
          shiftStart = normalizeTime(val);
          if (shiftStart) shiftEnd = addHours(shiftStart, 9);
        }
        return { dateKey: dc.label, type: type, shiftStart: shiftStart, shiftEnd: shiftEnd };
      });
    }
  }
  return map;
}

function getTeamBreaksMap(ldaps) {
  var raw = readBreaksSheet();
  if (!raw || raw.length < 2) return {};

  var headerRowIdx = 0;
  for (var i = 0; i < Math.min(raw.length, 10); i++) {
    if (raw[i].some(function(c) { return String(c).trim().toUpperCase() === 'LDAP'; })) {
      headerRowIdx = i;
      break;
    }
  }

  var headers    = raw[headerRowIdx].map(function(h) { return String(h).trim().toUpperCase(); });
  var ldapIdx    = headers.indexOf('LDAP');
  var timeIdx    = headers.indexOf('TIME');
  var rtaIdx     = headers.findIndex(function(h) { return h === 'RTA APPROVAL STATUS' || (h.includes('RTA') && h.includes('APPROVAL')); });
  var remarksIdx = headers.findIndex(function(h) { return h === 'TIME REMARKS' || (h.includes('TIME') && h.includes('REMARKS')); });

  var map        = {};
  var monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  if (ldaps) ldaps.forEach(function(l) { map[l.toLowerCase()] = {}; });

  for (var r = headerRowIdx + 1; r < raw.length; r++) {
    var row     = raw[r];
    var rowLdap = String(row[ldapIdx] || '').trim().toLowerCase();
    if (ldaps && !map.hasOwnProperty(rowLdap)) continue;
    if (!map[rowLdap]) map[rowLdap] = {};

    var timeVal = String(row[timeIdx] || '').trim();
    if (!timeVal) continue;
    var dateObj = new Date(timeVal);
    if (isNaN(dateObj.getTime())) continue;

    var dateKey   = monthNames[dateObj.getMonth()] + '-' + dateObj.getDate();
    var schedTime = formatTime12h(dateObj);
    var actualTime = schedTime;
    var rtaVal    = rtaIdx     !== -1 ? String(row[rtaIdx]     || '').trim() : '';
    var remarksVal = remarksIdx !== -1 ? String(row[remarksIdx] || '').trim() : '';

    var status = 'No result yet';
    if (rtaVal) {
      var rtaLower = rtaVal.toLowerCase();
      if (rtaLower.includes('adhere')) {
        status = 'Adhere';
      } else if (rtaLower.includes('check') || rtaLower.includes('rta')) {
        status = 'Check RTA';
        if (remarksVal) {
          var rd = new Date(remarksVal);
          actualTime = !isNaN(rd.getTime()) ? formatTime12h(rd) : remarksVal;
        }
      }
    }

    if (!map[rowLdap][dateKey]) map[rowLdap][dateKey] = [];
    map[rowLdap][dateKey].push({ schedTime: schedTime, actualTime: actualTime, status: status });
  }
  return map;
}
