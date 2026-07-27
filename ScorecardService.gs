// =========================================================================
// ScorecardService.gs — Enhanced Scorecard with correct column mapping
// Column B: Month, C: LDAP, D: Channel, E: Symptom
// F: Chat AHT, G: Phone AHT, H: Phone Talk Time, I: Phone Wait Time
// J: Phone Hold Time, K: Phone ACW, L: Net Cases Closed
// N: Incoming Phone Calls, O: Incoming Emails Total, P: Number of Total Chats
// Q: Survey Offer Rate, R: Chat SLA%, S: Phone SLA%, T: Email SLA%
// U: Repeat Contact Rate (7 Days), W: Team Captain/Supervisor
// =========================================================================

var PERF_SS_ID = '1wH2AVGJ9jyZJUX1EIvhBkMn-i7gOAE9z9gxB7STjC0g';
var AHT_SHEET_NAME = 'AHTDump';

// ── TARGETS ──────────────────────────────────────────────────────────────
var SCORECARD_TARGETS = {
  chatAht:       1200,  // 20:00 in seconds
  phoneAht:       820,  // 13:40 in seconds
  chatSla:         90,  // 90%
  phoneSla:        90,  // 90%
  emailSla:        90,  // 90%
  surveyOfferRate: 85,  // 85%
  repeatContact:   10   // ≤10% is good
};

// ── COLUMN INDEX HELPERS (0-based from column A) ─────────────────────────
// B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9, K=10, L=11, N=13, O=14,
// P=15, Q=16, R=17, S=18, T=19, U=20, W=22
var COL = {
  month:           1,  // B
  ldap:            2,  // C
  channel:         3,  // D
  symptom:         4,  // E
  chatAht:         5,  // F
  phoneAht:        6,  // G
  phoneTalkTime:   7,  // H
  phoneWaitTime:   8,  // I
  phoneHoldTime:   9,  // J
  phoneAcw:       10,  // K
  netCasesClosed: 11,  // L
  incomingPhone:  13,  // N
  incomingEmail:  14,  // O
  totalChats:     15,  // P
  surveyOfferRate:16,  // Q
  chatSla:        17,  // R
  phoneSla:       18,  // S
  emailSla:       19,  // T
  repeatContact:  20,  // U
  supervisor:     22   // W
};

// ── MAIN AGENT SCORECARD ──────────────────────────────────────────────────
function clientGetAgentScorecard(ldap, monthLabel) {
  return executeWithErrorHandling(function() {
    var sheet = SpreadsheetApp.openById(PERF_SS_ID).getSheetByName(AHT_SHEET_NAME);
    if (!sheet) throw new Error('AHTDump sheet not found.');

    var data = sheet.getDataRange().getValues();
    var targetLdap = String(ldap || '').trim().toLowerCase();
    var targetMonth = String(monthLabel || '').trim();

    // Accumulators
    var chatAhtSeconds = 0, chatVolume = 0;
    var phoneAhtSeconds = 0, phoneVolume = 0;
    var phoneTalkTime = 0, phoneWaitTime = 0, phoneHoldTime = 0, phoneAcw = 0;
    var netCasesClosed = 0;
    var incomingPhone = 0, incomingEmail = 0, totalChats = 0;
    var chatSlaSum = 0, chatSlaCount = 0;
    var phoneSlaSum = 0, phoneSlaCount = 0;
    var emailSlaSum = 0, emailSlaCount = 0;
    var surveyOfferSum = 0, surveyOfferCount = 0;
    var repeatContactSum = 0, repeatContactCount = 0;
    var supervisor = '';
    var symptoms = {};
    var rowCount = 0;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      // Parse month
      var rowMonthRaw = row[COL.month];
      var rowMonthStr = '';
      if (rowMonthRaw instanceof Date) {
        rowMonthStr = rowMonthRaw.getFullYear() + '-' + String(rowMonthRaw.getMonth() + 1).padStart(2, '0');
      } else {
        rowMonthStr = String(rowMonthRaw || '').trim();
      }

      var rowLdap = String(row[COL.ldap] || '').trim().toLowerCase();
      if (rowLdap !== targetLdap || rowMonthStr !== targetMonth) continue;

      rowCount++;

      // Channel
      var channel = String(row[COL.channel] || '').trim().toLowerCase();

      // Symptom tracking
      var symptom = String(row[COL.symptom] || '').trim();
      if (symptom) {
        symptoms[symptom] = (symptoms[symptom] || 0) + 1;
      }

      // Supervisor
      if (!supervisor && row[COL.supervisor]) {
        supervisor = String(row[COL.supervisor]).trim();
      }

      // ── CHAT AHT (Column F) ──────────────────────────────────────────
      var chatAhtVal = parseFloat(row[COL.chatAht]) || 0;
      var chatVol    = parseInt(row[COL.totalChats]) || 0;
      if (chatVol > 0 && chatAhtVal > 0) {
        chatAhtSeconds += chatAhtVal * chatVol;
        chatVolume += chatVol;
      }

      // ── PHONE AHT (Column G) — only when channel is phone ────────────
      var phoneAhtVal = parseFloat(row[COL.phoneAht]) || 0;
      var phoneVol    = parseInt(row[COL.incomingPhone]) || 0;
      if (phoneVol > 0 && phoneAhtVal > 0) {
        phoneAhtSeconds += phoneAhtVal * phoneVol;
        phoneVolume += phoneVol;
      }

      // Phone decomposition
      var pTalk = parseFloat(row[COL.phoneTalkTime]) || 0;
      var pWait = parseFloat(row[COL.phoneWaitTime]) || 0;
      var pHold = parseFloat(row[COL.phoneHoldTime]) || 0;
      var pAcw  = parseFloat(row[COL.phoneAcw]) || 0;
      if (phoneVol > 0) {
        phoneTalkTime += pTalk * phoneVol;
        phoneWaitTime += pWait * phoneVol;
        phoneHoldTime += pHold * phoneVol;
        phoneAcw      += pAcw  * phoneVol;
      }

      // Volume
      netCasesClosed += parseInt(row[COL.netCasesClosed]) || 0;
      incomingPhone  += parseInt(row[COL.incomingPhone])  || 0;
      incomingEmail  += parseInt(row[COL.incomingEmail])  || 0;
      totalChats     += parseInt(row[COL.totalChats])     || 0;

      // SLA
      var cSla = parseFloat(row[COL.chatSla]);
      if (!isNaN(cSla) && cSla > 0) { chatSlaSum += cSla; chatSlaCount++; }

      var pSla = parseFloat(row[COL.phoneSla]);
      if (!isNaN(pSla) && pSla > 0) { phoneSlaSum += pSla; phoneSlaCount++; }

      var eSla = parseFloat(row[COL.emailSla]);
      if (!isNaN(eSla) && eSla > 0) { emailSlaSum += eSla; emailSlaCount++; }

      var sor = parseFloat(row[COL.surveyOfferRate]);
      if (!isNaN(sor) && sor > 0) { surveyOfferSum += sor; surveyOfferCount++; }

      var rc = parseFloat(row[COL.repeatContact]);
      if (!isNaN(rc)) { repeatContactSum += rc; repeatContactCount++; }
    }

    if (rowCount === 0) {
      return { success: true, hasData: false, ldap: ldap, period: monthLabel };
    }

    // Final calculations
    var chatAhtFinal    = chatVolume  > 0 ? Math.round(chatAhtSeconds  / chatVolume)  : null;
    var phoneAhtFinal   = phoneVolume > 0 ? Math.round(phoneAhtSeconds / phoneVolume) : null;

    var totalVol = chatVolume + phoneVolume;
    var blendedAhtFinal = totalVol > 0
      ? Math.round((chatAhtSeconds + phoneAhtSeconds) / totalVol)
      : null;

    // Weighted blended target
    var blendedTarget = totalVol > 0
      ? Math.round((SCORECARD_TARGETS.chatAht * chatVolume + SCORECARD_TARGETS.phoneAht * phoneVolume) / totalVol)
      : null;

    // Phone decomp (per call)
    var phoneTalkFinal = phoneVolume > 0 ? Math.round(phoneTalkTime / phoneVolume) : null;
    var phoneWaitFinal = phoneVolume > 0 ? Math.round(phoneWaitTime / phoneVolume) : null;
    var phoneHoldFinal = phoneVolume > 0 ? Math.round(phoneHoldTime / phoneVolume) : null;
    var phoneAcwFinal  = phoneVolume > 0 ? Math.round(phoneAcw      / phoneVolume) : null;

    // SLA averages (Fixed decimal to percentage scaling)
    var chatSlaFinal    = chatSlaCount    > 0 ? Math.round((chatSlaSum    / chatSlaCount) * 100)    : null;
    var phoneSlaFinal   = phoneSlaCount   > 0 ? Math.round((phoneSlaSum   / phoneSlaCount) * 100)   : null;
    var emailSlaFinal   = emailSlaCount   > 0 ? Math.round((emailSlaSum   / emailSlaCount) * 100)   : null;
    var surveyFinal     = surveyOfferCount > 0 ? Math.round((surveyOfferSum / surveyOfferCount) * 100) : null;
    var repeatFinal     = repeatContactCount > 0 ? parseFloat(((repeatContactSum / repeatContactCount) * 100).toFixed(1)) : null;

    // Top symptoms
    var topSymptoms = Object.keys(symptoms)
      .map(function(s) { return { symptom: s, count: symptoms[s] }; })
      .sort(function(a, b) { return b.count - a.count; })
      .slice(0, 8);

    // PTG for blended
    var blendedPtg = blendedAhtFinal !== null && blendedTarget > 0
      ? parseFloat((blendedTarget / blendedAhtFinal * 100).toFixed(1))
      : null;

    return {
      success:        true,
      hasData:        true,
      ldap:           ldap,
      period:         monthLabel,
      supervisor:     supervisor,

      // AHT
      chatAht:        chatAhtFinal,
      phoneAht:       phoneAhtFinal,
      blendedAht:     blendedAhtFinal,
      blendedTarget:  blendedTarget,
      blendedPtg:     blendedPtg,

      // Targets
      chatAhtTarget:  SCORECARD_TARGETS.chatAht,
      phoneAhtTarget: SCORECARD_TARGETS.phoneAht,

      // Phone decomposition
      phoneTalkTime:  phoneTalkFinal,
      phoneWaitTime:  phoneWaitFinal,
      phoneHoldTime:  phoneHoldFinal,
      phoneAcw:       phoneAcwFinal,

      // Volume
      totalChats:     chatVolume,
      incomingPhone:  phoneVolume,
      incomingEmail:  incomingEmail,
      netCasesClosed: netCasesClosed,
      totalVolume:    totalVol,

      // SLA
      chatSla:        chatSlaFinal,
      phoneSla:       phoneSlaFinal,
      emailSla:       emailSlaFinal,
      chatSlaTarget:  SCORECARD_TARGETS.chatSla,
      phoneSlaTarget: SCORECARD_TARGETS.phoneSla,
      emailSlaTarget: SCORECARD_TARGETS.emailSla,

      // Other metrics
      surveyOfferRate: surveyFinal,
      repeatContact:   repeatFinal,

      // Symptoms
      topSymptoms:    topSymptoms
    };
  }, this, 'clientGetAgentScorecard');
}

// ── TEAM SCORECARD (for supervisors/managers) ─────────────────────────────
function clientGetTeamScorecard(monthLabel) {
  return executeWithErrorHandling(function() {
    requireSupervisor();

    var myLdap = getCurrentLdap().toLowerCase();
    var managedLdaps = getManagedLdaps(myLdap);
    
    var sheet = SpreadsheetApp.openById(PERF_SS_ID).getSheetByName(AHT_SHEET_NAME);
    if (!sheet) throw new Error('AHTDump sheet not found.');

    var data = sheet.getDataRange().getValues();
    var targetMonth = String(monthLabel || '').trim();

    // Aggregate per agent, then roll up
    var agentMap = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      var rowMonthRaw = row[COL.month];
      var rowMonthStr = '';
      if (rowMonthRaw instanceof Date) {
        rowMonthStr = rowMonthRaw.getFullYear() + '-' + String(rowMonthRaw.getMonth() + 1).padStart(2, '0');
      } else {
        rowMonthStr = String(rowMonthRaw || '').trim();
      }
      if (rowMonthStr !== targetMonth) continue;

      var rowLdap = String(row[COL.ldap] || '').trim().toLowerCase();
      var rowSupervisor = String(row[COL.supervisor] || '').trim().toLowerCase();
      
      // STRICT FILTER: Only include if they are explicitly mapped to the logged-in manager
      var isMyAgent = false;
      if (rowSupervisor === myLdap) isMyAgent = true;
      if (managedLdaps && managedLdaps.indexOf(rowLdap) !== -1) isMyAgent = true;
      
      if (!isMyAgent) continue;

      if (!agentMap[rowLdap]) {
        agentMap[rowLdap] = {
          ldap: rowLdap,
          chatAhtS: 0, chatVol: 0,
          phoneAhtS: 0, phoneVol: 0,
          netCasesClosed: 0,
          chatSlaSum: 0, chatSlaCount: 0,
          phoneSlaSum: 0, phoneSlaCount: 0,
          emailSlaSum: 0, emailSlaCount: 0,
          surveyOfferSum: 0, surveyOfferCount: 0,
          repeatContactSum: 0, repeatContactCount: 0,
          supervisor: ''
        };
      }
      var a = agentMap[rowLdap];

      if (!a.supervisor && row[COL.supervisor]) a.supervisor = String(row[COL.supervisor]).trim();

      var chatAhtVal = parseFloat(row[COL.chatAht]) || 0;
      var chatVol    = parseInt(row[COL.totalChats]) || 0;
      if (chatVol > 0 && chatAhtVal > 0) { a.chatAhtS += chatAhtVal * chatVol; a.chatVol += chatVol; }

      var phoneAhtVal = parseFloat(row[COL.phoneAht]) || 0;
      var phoneVol    = parseInt(row[COL.incomingPhone]) || 0;
      if (phoneVol > 0 && phoneAhtVal > 0) { a.phoneAhtS += phoneAhtVal * phoneVol; a.phoneVol += phoneVol; }

      a.netCasesClosed += parseInt(row[COL.netCasesClosed]) || 0;

      var cSla = parseFloat(row[COL.chatSla]);
      if (!isNaN(cSla) && cSla > 0) { a.chatSlaSum += cSla; a.chatSlaCount++; }
      var pSla = parseFloat(row[COL.phoneSla]);
      if (!isNaN(pSla) && pSla > 0) { a.phoneSlaSum += pSla; a.phoneSlaCount++; }
      var eSla = parseFloat(row[COL.emailSla]);
      if (!isNaN(eSla) && eSla > 0) { a.emailSlaSum += eSla; a.emailSlaCount++; }
      var sor = parseFloat(row[COL.surveyOfferRate]);
      if (!isNaN(sor) && sor > 0) { a.surveyOfferSum += sor; a.surveyOfferCount++; }
      var rc = parseFloat(row[COL.repeatContact]);
      if (!isNaN(rc)) { a.repeatContactSum += rc; a.repeatContactCount++; }
    }

    // Build per-agent summaries
    var agents = Object.keys(agentMap).map(function(ldap) {
      var a = agentMap[ldap];
      var agentData = getAgentProfile(ldap);
      var displayName = agentData ? (agentData.displayName || ldap) : ldap;

      var chatAht  = a.chatVol  > 0 ? Math.round(a.chatAhtS  / a.chatVol)  : null;
      var phoneAht = a.phoneVol > 0 ? Math.round(a.phoneAhtS / a.phoneVol) : null;
      var tv = a.chatVol + a.phoneVol;
      var blended = tv > 0 ? Math.round((a.chatAhtS + a.phoneAhtS) / tv) : null;
      var blendedTarget = tv > 0
        ? Math.round((SCORECARD_TARGETS.chatAht * a.chatVol + SCORECARD_TARGETS.phoneAht * a.phoneVol) / tv)
        : null;
      var blendedPtg = blended !== null && blendedTarget > 0
        ? parseFloat((blendedTarget / blended * 100).toFixed(1))
        : null;

      var chatSla   = a.chatSlaCount   > 0 ? Math.round((a.chatSlaSum   / a.chatSlaCount) * 100)   : null;
      var phoneSla  = a.phoneSlaCount  > 0 ? Math.round((a.phoneSlaSum  / a.phoneSlaCount) * 100)  : null;
      var emailSla  = a.emailSlaCount  > 0 ? Math.round((a.emailSlaSum  / a.emailSlaCount) * 100)  : null;
      var surveyOff = a.surveyOfferCount > 0 ? Math.round((a.surveyOfferSum / a.surveyOfferCount) * 100) : null;
      var repeatC   = a.repeatContactCount > 0 ? parseFloat(((a.repeatContactSum / a.repeatContactCount) * 100).toFixed(1)) : null;

      // Status: good if blended PTG >= 100
      var status = 'no_data';
      if (blendedPtg !== null) {
        status = blendedPtg >= 100 ? 'on_target' : blendedPtg >= 85 ? 'monitor' : 'at_risk';
      }

      return {
        ldap:           ldap,
        displayName:    displayName,
        chatAht:        chatAht,
        phoneAht:       phoneAht,
        blendedAht:     blended,
        blendedTarget:  blendedTarget,
        blendedPtg:     blendedPtg,
        chatVolume:     a.chatVol,
        phoneVolume:    a.phoneVol,
        totalVolume:    tv,
        netCasesClosed: a.netCasesClosed,
        chatSla:        chatSla,
        phoneSla:       phoneSla,
        emailSla:       emailSla,
        surveyOfferRate:surveyOff,
        repeatContact:  repeatC,
        status:         status
      };
    });

    // Sort by blended PTG descending
    agents.sort(function(a, b) {
      if (a.blendedPtg === null && b.blendedPtg === null) return 0;
      if (a.blendedPtg === null) return 1;
      if (b.blendedPtg === null) return -1;
      return b.blendedPtg - a.blendedPtg;
    });

    // Team totals
    var teamChatVol = 0, teamChatAhtS = 0, teamPhoneVol = 0, teamPhoneAhtS = 0;
    var teamCases = 0, teamAtRisk = 0;
    agents.forEach(function(a) {
      teamChatVol   += a.chatVolume   || 0;
      teamChatAhtS  += (a.chatAht    || 0) * (a.chatVolume  || 0);
      teamPhoneVol  += a.phoneVolume  || 0;
      teamPhoneAhtS += (a.phoneAht   || 0) * (a.phoneVolume || 0);
      teamCases     += a.netCasesClosed || 0;
      if (a.status === 'at_risk') teamAtRisk++;
    });

    var teamTV = teamChatVol + teamPhoneVol;
    var teamBlended = teamTV > 0 ? Math.round((teamChatAhtS + teamPhoneAhtS) / teamTV) : null;
    var teamBlendedTarget = teamTV > 0
      ? Math.round((SCORECARD_TARGETS.chatAht * teamChatVol + SCORECARD_TARGETS.phoneAht * teamPhoneVol) / teamTV)
      : null;
    var teamBlendedPtg = teamBlended && teamBlendedTarget
      ? parseFloat((teamBlendedTarget / teamBlended * 100).toFixed(1))
      : null;

    return {
      success:        true,
      period:         monthLabel,
      agents:         agents,
      agentCount:     agents.length,
      atRiskCount:    teamAtRisk,
      teamBlendedAht: teamBlended,
      teamBlendedTarget: teamBlendedTarget,
      teamBlendedPtg: teamBlendedPtg,
      teamChatVol:    teamChatVol,
      teamPhoneVol:   teamPhoneVol,
      teamCases:      teamCases
    };
  }, this, 'clientGetTeamScorecard');
}

// ── AVAILABLE MONTHS ──────────────────────────────────────────────────────
function clientGetAvailableAhtMonths() {
  try {
    var sheet = SpreadsheetApp.openById(PERF_SS_ID).getSheetByName(AHT_SHEET_NAME);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var seen = {}, months = [];
    for (var i = 1; i < data.length; i++) {
      var raw = data[i][COL.month];
      var monthStr = '';
      if (raw instanceof Date) {
        monthStr = raw.getFullYear() + '-' + String(raw.getMonth() + 1).padStart(2, '0');
      } else {
        monthStr = String(raw || '').trim();
      }
      if (monthStr && !seen[monthStr]) {
        seen[monthStr] = true;
        months.push(monthStr);
      }
    }
    return months.sort().reverse();
  } catch(e) {
    return [];
  }
}

// ── AI SCORECARD COACHING ─────────────────────────────────────────────────
function clientGetScorecardCoaching(ldap, monthLabel) {
  return executeWithErrorHandling(function() {
    var requesterLdap = getCurrentLdap();
    var requesterRole = getUserRole(requesterLdap);
    var isMgmt = requesterRole === 'manager' || requesterRole === 'supervisor';
    var targetLdap = (ldap && isMgmt) ? ldap : requesterLdap;

    var data = clientGetAgentScorecard(targetLdap, monthLabel);
    if (!data || !data.success || !data.hasData) {
      return { success: false, error: 'No scorecard data available for this agent.' };
    }

    var agentProfile = getAgentProfile(targetLdap);
    var displayName = agentProfile ? (agentProfile.displayName || targetLdap) : targetLdap;

    var secsToStr = function(s) {
      if (!s) return 'N/A';
      var m = Math.floor(s / 60);
      var sec = s % 60;
      return m + ':' + String(sec).padStart(2, '0');
    };

    var prompt;
    if (isMgmt) {
      // SMART format for managers
      prompt =
        'You are a Google Play Operations senior manager reviewing an agent performance scorecard. ' +
        'Provide a SMART (Specific, Measurable, Achievable, Relevant, Time-bound) coaching report. ' +
        'Format your response using these EXACT headers:\n' +
        '## Performance Summary\n' +
        '## SMART Goals for This Month\n' +
        '## Coaching Actions (This Week)\n' +
        '## Risk Flags\n' +
        '## Recognition Opportunities\n\n' +
        'Agent: ' + displayName + ' (' + targetLdap + ')\n' +
        'Period: ' + monthLabel + '\n\n' +
        '--- AHT DATA ---\n' +
        'Blended AHT: ' + secsToStr(data.blendedAht) + ' (target: ' + secsToStr(data.blendedTarget) + ', PTG: ' + (data.blendedPtg !== null ? data.blendedPtg + '%' : 'N/A') + ')\n' +
        'Chat AHT: ' + secsToStr(data.chatAht) + ' (' + (data.totalChats || 0) + ' chats, target: ' + secsToStr(data.chatAhtTarget) + ')\n' +
        'Phone AHT: ' + secsToStr(data.phoneAht) + ' (' + (data.incomingPhone || 0) + ' calls, target: ' + secsToStr(data.phoneAhtTarget) + ')\n' +
        'Phone Talk: ' + secsToStr(data.phoneTalkTime) + ' | Hold: ' + secsToStr(data.phoneHoldTime) + ' | ACW: ' + secsToStr(data.phoneAcw) + '\n\n' +
        '--- SLA DATA ---\n' +
        'Chat SLA: ' + (data.chatSla !== null ? data.chatSla + '%' : 'N/A') + ' (target: ' + SCORECARD_TARGETS.chatSla + '%)\n' +
        'Phone SLA: ' + (data.phoneSla !== null ? data.phoneSla + '%' : 'N/A') + ' (target: ' + SCORECARD_TARGETS.phoneSla + '%)\n' +
        'Email SLA: ' + (data.emailSla !== null ? data.emailSla + '%' : 'N/A') + ' (target: ' + SCORECARD_TARGETS.emailSla + '%)\n\n' +
        '--- VOLUME & QUALITY ---\n' +
        'Net Cases Closed: ' + (data.netCasesClosed || 0) + '\n' +
        'Survey Offer Rate: ' + (data.surveyOfferRate !== null ? data.surveyOfferRate + '%' : 'N/A') + ' (target: ' + SCORECARD_TARGETS.surveyOfferRate + '%)\n' +
        'Repeat Contact Rate: ' + (data.repeatContact !== null ? data.repeatContact + '%' : 'N/A') + ' (target: ≤' + SCORECARD_TARGETS.repeatContact + '%)\n\n' +
        'Top symptoms handled: ' + (data.topSymptoms || []).slice(0, 5).map(function(t) { return t.symptom + ' (' + t.count + ')'; }).join(', ') + '\n\n' +
        'Write SMART goals for each metric that is off-target. ' +
        'Each goal must state the current value, target value, and a concrete action plan with timeline. ' +
        'Keep each section to 3-4 bullet points max.';
    } else {
      // Insights format for agents
      prompt =
        'You are a Google Play Operations team leader providing performance coaching to a customer support agent. ' +
        'Be encouraging, specific, and actionable. Use a warm but professional tone. ' +
        'Format your response using these EXACT headers:\n' +
        '## Your Performance This Month\n' +
        '## What You Are Doing Well\n' +
        '## Areas to Improve\n' +
        '## Your Action Plan\n' +
        '## Quick Tips for Next Week\n\n' +
        'Agent: ' + displayName + '\n' +
        'Period: ' + monthLabel + '\n\n' +
        'Blended AHT: ' + secsToStr(data.blendedAht) + ' (target: ' + secsToStr(data.blendedTarget) + ')\n' +
        'Chat AHT: ' + secsToStr(data.chatAht) + ' from ' + (data.totalChats || 0) + ' chats (target: ' + secsToStr(data.chatAhtTarget) + ')\n' +
        'Phone AHT: ' + secsToStr(data.phoneAht) + ' from ' + (data.incomingPhone || 0) + ' calls (target: ' + secsToStr(data.phoneAhtTarget) + ')\n' +
        'Chat SLA: ' + (data.chatSla !== null ? data.chatSla + '%' : 'N/A') + '\n' +
        'Phone SLA: ' + (data.phoneSla !== null ? data.phoneSla + '%' : 'N/A') + '\n' +
        'Survey Offer Rate: ' + (data.surveyOfferRate !== null ? data.surveyOfferRate + '%' : 'N/A') + '\n' +
        'Repeat Contact Rate: ' + (data.repeatContact !== null ? data.repeatContact + '%' : 'N/A') + '\n' +
        'Cases Closed: ' + (data.netCasesClosed || 0) + '\n\n' +
        'Provide specific, encouraging coaching. Focus on what the agent can control. ' +
        'Keep each section to 2-3 bullet points.';
    }

    try {
      var response = callGroq(prompt);
      return { success: true, coaching: response, displayName: displayName, period: monthLabel, isSmart: isMgmt };
    } catch(e) {
      return { success: false, error: e.message };
    }
  }, this, 'clientGetScorecardCoaching');
}

// ── TEAM AI COACHING ──────────────────────────────────────────────────────
function clientGetTeamScorecardCoaching(monthLabel) {
  return executeWithErrorHandling(function() {
    requireSupervisor();
    var data = clientGetTeamScorecard(monthLabel);
    if (!data || !data.success) {
      return { success: false, error: 'No team scorecard data available.' };
    }

    var secsToStr = function(s) {
      if (!s) return 'N/A';
      var m = Math.floor(s / 60);
      var sec = s % 60;
      return m + ':' + String(sec).padStart(2, '0');
    };

    var agentSummary = (data.agents || []).map(function(a) {
      return a.displayName + ' (' + a.ldap + '): ' +
        'Blended AHT ' + secsToStr(a.blendedAht) +
        ' PTG ' + (a.blendedPtg !== null ? a.blendedPtg + '%' : 'N/A') +
        ' [' + a.status + ']' +
        ' | Chat ' + secsToStr(a.chatAht) + ' (' + (a.chatVolume || 0) + ')' +
        ' | Phone ' + secsToStr(a.phoneAht) + ' (' + (a.phoneVolume || 0) + ')';
    }).join('\n');

    var prompt =
      'You are a senior Google Play Operations manager reviewing your team\'s efficiency scorecard. ' +
      'Provide a SMART team coaching report. Format using these EXACT headers:\n' +
      '## Team Performance Summary\n' +
      '## SMART Team Goals\n' +
      '## Agents Needing Immediate Coaching\n' +
      '## Team Strengths\n' +
      '## Manager Action Plan (This Week)\n\n' +
      'Period: ' + monthLabel + '\n' +
      'Team Blended AHT: ' + secsToStr(data.teamBlendedAht) + ' (target: ' + secsToStr(data.teamBlendedTarget) + ', PTG: ' + (data.teamBlendedPtg !== null ? data.teamBlendedPtg + '%' : 'N/A') + ')\n' +
      'Total Chat Volume: ' + data.teamChatVol + '\n' +
      'Total Phone Volume: ' + data.teamPhoneVol + '\n' +
      'Total Cases Closed: ' + data.teamCases + '\n' +
      'Agents at risk: ' + data.atRiskCount + ' of ' + data.agentCount + '\n\n' +
      'Individual breakdown:\n' + agentSummary + '\n\n' +
      'Write SMART goals for the team. Name specific agents who need priority coaching. ' +
      'Keep each section to 3-5 bullet points.';

    try {
      var response = callGroq(prompt);
      return { success: true, coaching: response, period: monthLabel };
    } catch(e) {
      return { success: false, error: e.message };
    }
  }, this, 'clientGetTeamScorecardCoaching');
}
