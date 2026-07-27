// ============================================================
// GeminiService.gs — Gemini AI coaching for CSAT Dashboard
// ============================================================

var GEMINI_MODEL   = 'gemini-3.5-flash';
var GEMINI_API_URL = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent';

function getGeminiApiKey() {
  // Request the property by its name
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

function callGemini(prompt) {
  var apiKey = getGeminiApiKey();
  if (!apiKey) throw new Error('Gemini API key not configured.');

  var payload = {
    systemInstruction: {
      parts: [
        { text: 'You are a Google Play Operations coaching assistant. Be encouraging, specific, and actionable. ' +
          'Formatting rules: use "## " for section headers, "- " for bullet points, and "**text**" only for bold emphasis. ' +
          'Do not use backticks, code formatting, or single-asterisk italics anywhere. ' +
          'When referencing an agent name or number for emphasis, wrap it in double asterisks like **cabilete** or **70.6%**, never in backticks.' }
      ]
    },
    contents: [
      {
        role: 'user',
        parts: [ { text: prompt } ]
      }
    ],
    generationConfig: {
      maxOutputTokens: 4096,
      temperature: 0.7
    }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var url = GEMINI_API_URL + '?key=' + apiKey;
  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());
  if (json.error) throw new Error(json.error.message);

  var candidate = json.candidates && json.candidates[0];
  if (!candidate || !candidate.content || !candidate.content.parts || !candidate.content.parts[0]) {
    throw new Error('Gemini returned no usable content. Finish reason: ' + (candidate ? candidate.finishReason : 'unknown'));
  }
  return candidate.content.parts[0].text;
}


// ── COACHING CACHE ───────────────────────────────────────────
function getCoachingCacheSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CoachingCache');
  if (!sheet) {
    sheet = ss.insertSheet('CoachingCache');
    sheet.appendRow(['CacheKey', 'Timestamp', 'Payload']);
    sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#f3f3f3");
    sheet.setFrozenRows(1);
    sheet.hideSheet(); // Keep it hidden from users
  }
  return sheet;
}

function getCoachingFromCache(cacheKey) {
  var sheet = getCoachingCacheSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === cacheKey) {
      return data[i][2]; // Return payload
    }
  }
  return null;
}

function saveCoachingToCache(cacheKey, payload) {
  var sheet = getCoachingCacheSheet();
  var data = sheet.getDataRange().getValues();
  var timestamp = new Date();

  // Check if it exists to update
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === cacheKey) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[timestamp, payload]]);
      return;
    }
  }

  // Or append new
  sheet.appendRow([cacheKey, timestamp, payload]);
}

// ── AGENT COACHING ───────────────────────────────────────────
function getAgentCsatCoaching(ldap, month) {
  var data = getMyCsatData(ldap, month);
  if (!data || !data.hasData) {
    return { success: false, error: 'No CSAT data available for this agent.' };
  }

  var prompt =
    'You are speaking directly to ' + (data.displayName || ldap) + ', a Google Play Operations support agent, as their team leader delivering 1:1 coaching. ' +
    'Address them by name in the opening line. Be warm, encouraging, and specific — reference their actual numbers, not generic advice. ' +
    'Write in full, realistic sentences (not just fragments) inside each section, as if this were a real coaching note you would send them. ' +
    'Format your response in clear sections using these exact headers:\n' +
    '## Performance Summary\n' +
    '## What You Are Doing Well\n' +
    '## Areas to Focus On\n' +
    '## Your Coaching Plan This Month\n' +
    '## Quick Wins for Next Week\n\n' +
    'Here is the agent data for ' + data.month + ':\n' +
    'Agent: ' + (data.displayName || ldap) + '\n' +
    'Overall CSAT: ' + (data.overall !== null ? data.overall + '%' : 'No data') + '\n' +
    'Chat CSAT: ' + (data.chat !== null ? data.chat + '% (' + data.chatResponses + ' responses)' : 'No data') + '\n' +
    'Phone CSAT: ' + (data.phone !== null ? data.phone + '% (' + data.phoneResponses + ' responses)' : 'No data') + '\n' +
    'Email CSAT: ' + (data.email !== null ? data.email + '% (' + data.emailResponses + ' responses)' : 'No data') + '\n' +
    'CSAT count: ' + data.csatCount + '\n' +
    'DSAT count: ' + data.dsatCount + '\n' +
    'Total surveyed: ' + data.totalSurveyed + '\n' +
    'Team rank: ' + (data.teamRank ? '#' + data.teamRank + ' of ' + data.teamSize : 'N/A') + '\n' +
    'Status: ' + (data.status || 'unknown') + '\n' +
    'Month-over-month overall delta: ' + (data.delta && data.delta.overall !== null ? data.delta.overall + 'pp' : 'N/A') + '\n' +
    'Top DSAT themes: ' + (data.dsatThemes && data.dsatThemes.length > 0
      ? data.dsatThemes.map(function(t) { return t.symptom + ' (' + t.count + ')'; }).join(', ')
      : 'None') + '\n' +
    'Cases resolution rate: ' + (data.casesResolutionRate !== null ? data.casesResolutionRate + '%' : 'N/A') + '\n' +
    'Repeat contact rate: ' + (data.repeatContactRate !== null ? data.repeatContactRate + '%' : 'N/A') + '\n\n' +
    'Write a realistic, detailed coaching note grounded in this specific data — call out their exact CSAT numbers, rank, and DSAT themes by name. ' +
    'If DSAT themes exist, explain what likely went wrong in those interactions and give concrete handling tips for each one. ' +
    'If they are trending up or down month-over-month, acknowledge it explicitly and explain what to keep doing or change. ' +
    'Each section should be 3-5 sentences or bullet points — enough to feel like genuine, individualized coaching rather than a template filled with their name.';

  try {
    var cacheKey = 'agent_' + ldap + '_' + month;
    var cachedResponse = getCoachingFromCache(cacheKey);
    var response;

    if (cachedResponse) {
      response = cachedResponse;
    } else {
      response = callGemini(prompt);
      saveCoachingToCache(cacheKey, response);
    }

    return { success: true, coaching: response, agentName: data.displayName || ldap, month: data.month };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── TEAM COACHING ────────────────────────────────────────────
function getTeamCsatCoaching(managerLdap, month) {
  var data = getTeamCsatData(managerLdap, month);
  if (!data || data.overall === null) {
    return { success: false, error: 'No team CSAT data available.' };
  }

  var agentSummary = (data.agents || []).map(function(a) {
    return a.displayName + ': ' + (a.overall !== null ? a.overall + '%' : 'no data') +
      ' (' + a.status + ')' +
      (a.dsatThemes && a.dsatThemes.length > 0
        ? ' — top DSAT: ' + a.dsatThemes[0].symptom
        : '');
  }).join('\n');

  var prompt =
    'You are speaking directly to the supervisor of this team as a senior Google Play Operations coaching partner, reviewing this month\'s CSAT performance with them. ' +
    'Address the reader as the team\'s manager throughout — this is written for them to read and act on, not for the agents themselves. ' +
    'Be data-driven, direct, and realistic. Write in full sentences inside each section, referencing real numbers and real agent names from the data below — not generic management advice. ' +
    'Format your response using these exact headers:\n' +
    '## Team Performance Summary\n' +
    '## Team Strengths\n' +
    '## Key Risk Areas\n' +
    '## Recommended Coaching Actions\n' +
    '## Focus for Next Month\n\n' +
    'Here is the team data for ' + data.month + ':\n' +
    'Team overall CSAT: ' + (data.overall !== null ? data.overall + '%' : 'No data') + '\n' +
    'Chat CSAT: ' + (data.chat !== null ? data.chat + '% (' + data.chatResponses + ' responses)' : 'No data') + '\n' +
    'Phone CSAT: ' + (data.phone !== null ? data.phone + '% (' + data.phoneResponses + ' responses)' : 'No data') + '\n' +
    'Email CSAT: ' + (data.email !== null ? data.email + '% (' + data.emailResponses + ' responses)' : 'No data') + '\n' +
    'Total surveys: ' + data.totalSurveyed + '\n' +
    'CSAT count: ' + data.csatCount + '\n' +
    'DSAT count: ' + data.dsatCount + '\n' +
    'Agents at risk: ' + data.atRiskCount + '\n' +
    'Month-over-month delta: ' + (data.delta && data.delta.overall !== null ? data.delta.overall + 'pp' : 'N/A') + '\n' +
    'Top team DSAT themes: ' + (data.topDsatThemes && data.topDsatThemes.length > 0
      ? data.topDsatThemes.map(function(t) { return t.symptom + ' (' + t.count + ')'; }).join(', ')
      : 'None') + '\n\n' +
    'Individual agent breakdown:\n' + agentSummary + '\n\n' +
    'Write a detailed, realistic team review grounded in this specific data. Name specific agents by their displayName who are at risk or trending down, and specific agents who are strong performers worth recognizing. ' +
    'For each at-risk agent, briefly note what their DSAT theme suggests and one concrete coaching action for them. ' +
    'Suggest concrete, doable actions the supervisor can take this week — not vague suggestions like "monitor performance," but specific next steps (e.g., who to schedule a 1:1 with, what theme to address team-wide). ' +
    'Each section should be 3-6 bullet points with enough detail to act on immediately.';

  try {
    var cacheKey = 'team_' + managerLdap + '_' + month;
    var cachedResponse = getCoachingFromCache(cacheKey);
    var response;

    if (cachedResponse) {
      response = cachedResponse;
    } else {
      response = callGemini(prompt);
      saveCoachingToCache(cacheKey, response);
    }

    return { success: true, coaching: response, month: data.month };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── TEST FUNCTION ─────────────────────────────────────────────
function testGemini() {
  try {
    var result = callGemini('Say hello in one sentence.');
    Logger.log('SUCCESS: ' + result);
  } catch(e) {
    Logger.log('ERROR: ' + e.message);
  }
}

// ── OPUS AI CHAT ASSISTANT ──────────────────────────────────
function processOpusChat(ldap, message) {
  var agentInfo = findRow('Agents', 'LDAP', ldap);
  var name = agentInfo ? agentInfo['DisplayName'] : ldap;
  var role = agentInfo ? 'an agent' : 'a team member';
  
  var prompt = 
    "You are 'Opus', the friendly, highly concise internal AI assistant for Google Play Ops. " +
    "The user talking to you is " + name + " (" + ldap + "), who is " + role + ". " +
    "Keep your answers extremely brief, helpful, and formatted with basic markdown (bolding, lists). " +
    "Do not hallucinate internal links. If you don't know something, tell them to check the Help Center. " +
    "User says: " + message;
    
  try {
    var reply = callGemini(prompt);
    return { success: true, reply: reply };
  } catch(e) {
    return { success: false, error: e.message };
  }
}
