// ============================================================
// NotificationService.gs — In-app notifications + email sends
// Play Ops Store
// ============================================================

// ------------------------------------------------------------
// IN-APP NOTIFICATIONS
// ------------------------------------------------------------

function createNotification(ldap, type, message) {
  return createNotifications([{ ldap: ldap, type: type, message: message }])[0];
}

function createNotifications(notifications) {
  if (!notifications || notifications.length === 0) return [];

  var ids = generateIds('Notifications', 'NOTIF', notifications.length);

  var rowObjs = notifications.map(function(n, index) {
    return {
      ID: ids[index],
      LDAP: n.ldap,
      Type: n.type,
      Message: n.message,
      IsRead: false,
      CreatedAt: now()
    };
  });

  batchAppendRows('Notifications', rowObjs);
  return ids;
}

function getNotifications(ldap) {
  var rows = findRows('Notifications', 'LDAP', ldap);
  rows.sort(function(a, b) {
    return new Date(b['CreatedAt']) - new Date(a['CreatedAt']);
  });

  var recent = rows.slice(0, 30);
  var unreadCount = rows.filter(function(r) {
    return r['IsRead'] === false || r['IsRead'] === 'FALSE';
  }).length;

  return {
    notifications: recent.map(function(r) {
      return {
        id: r['ID'],
        type: r['Type'],
        message: r['Message'],
        isRead: r['IsRead'] === true || r['IsRead'] === 'TRUE',
        createdAt: formatDateTime(r['CreatedAt']),
        icon: getNotificationIcon(r['Type'])
      };
    }),
    unreadCount: unreadCount
  };
}

function markAllNotificationsRead(ldap) {
  var sheet = getSheet('Notifications');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var ldapIdx = headers.indexOf('LDAP');
  var isReadIdx = headers.indexOf('IsRead');

  var changed = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][ldapIdx] === ldap && data[i][isReadIdx] !== true) {
      data[i][isReadIdx] = true;
      changed = true;
    }
  }

  if (changed) {
    sheet.getDataRange().setValues(data);
  }

  return { success: true };
}

function getNotificationIcon(type) {
  var icons = {
    'claimed':        '📋',
    'completed':      '✅',
    'expiry_warning': '⏰',
    'new_task':       '🆕',
    'kudos_approved': '🌟',
    'kudos_rejected': '❌',
    'kudos_submitted':'📤',
    'new_kudos':      '📬',
    'demerit':        '⚠️',
    'badge':          '🏅',
    'streak':         '🔥',
    'missed':         '❗',
    'acknowledged':   '👁️'
  };
  return icons[type] || '🔔';
}

// ------------------------------------------------------------
// EMAIL — Shared builder
// ------------------------------------------------------------

function buildEmailHtml(title, bodyHtml, agentLdap) {
  var primaryColor = '#01875f'; // Google Play green
  var accentColor  = '#4285F4';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>' +
    'body{margin:0;padding:0;background:#f1f3f4;font-family:Google Sans,Roboto,sans-serif;}' +
    '.wrap{max-width:560px;margin:32px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.12);}' +
    '.header{background:' + primaryColor + ';padding:28px 32px;text-align:center;}' +
    '.header h1{margin:0;color:#fff;font-size:20px;font-weight:500;letter-spacing:.5px;}' +
    '.header p{margin:4px 0 0;color:rgba(255,255,255,.75);font-size:13px;}' +
    '.body{padding:32px;}' +
    '.body p{margin:0 0 16px;color:#3c4043;font-size:15px;line-height:1.6;}' +
    '.highlight{background:#e6f4ea;border-left:4px solid ' + primaryColor + ';border-radius:4px;padding:16px 20px;margin:20px 0;}' +
    '.highlight .pts{font-size:28px;font-weight:700;color:' + primaryColor + ';}' +
    '.highlight .label{font-size:13px;color:#5f6368;margin-top:2px;}' +
    '.badge-row{display:flex;align-items:center;gap:12px;background:#fafafa;border:1px solid #e8eaed;border-radius:8px;padding:14px 16px;margin:16px 0;}' +
    '.badge-icon{font-size:28px;}' +
    '.badge-name{font-weight:600;color:#202124;font-size:15px;}' +
    '.badge-desc{font-size:13px;color:#5f6368;}' +
    '.btn{display:inline-block;margin-top:8px;padding:12px 28px;background:' + primaryColor + ';color:#fff;text-decoration:none;border-radius:24px;font-size:14px;font-weight:500;}' +
    '.footer{padding:20px 32px;border-top:1px solid #e8eaed;text-align:center;}' +
    '.footer p{margin:0;font-size:12px;color:#9aa0a6;}' +
    '</style></head><body>' +
    '<div class="wrap">' +
    '<div class="header"><h1>Play Ops Store</h1><p>' + (agentLdap || '') + '</p></div>' +
    '<div class="body">' + bodyHtml + '</div>' +
    '<div class="footer"><p>Play Ops Store &mdash; Team Steven &mdash; Do not reply to this email.</p></div>' +
    '</div></body></html>';
}

function sendEmail(toEmail, subject, bodyHtml, agentLdap) {
  try {
    var html = buildEmailHtml(subject, bodyHtml, agentLdap);
    MailApp.sendEmail({
      to: toEmail,
      subject: subject,
      htmlBody: html
    });
  } catch (e) {
    Logger.log('[Email Error] ' + e.message + ' | To: ' + toEmail);
  }
}

function getLdapEmail(ldap) {
  var agent = findRow('Agents', 'LDAP', ldap);
  if (agent && agent['Email']) return agent['Email'];
  return ldap + '@google.com';
}

// ------------------------------------------------------------
// EMAIL — Task Claimed
// ------------------------------------------------------------

function sendTaskClaimedEmail(ldap, task, claimedAt) {
  var email = getLdapEmail(ldap);
  var subject = 'You claimed: ' + task['Title'];
  var deadline = task['Deadline'] ? formatDate(task['Deadline']) : 'No deadline';
  var body =
    '<p>Hi <strong>' + formatDisplayName(ldap) + '</strong>,</p>' +
    '<p>You\'ve successfully claimed a task. The TAT clock has started.</p>' +
    '<div class="highlight">' +
    '<div class="pts">' + task['BasePoints'] + ' pts</div>' +
    '<div class="label">' + task['Title'] + '</div>' +
    '</div>' +
    '<p><strong>Deadline:</strong> ' + deadline + '</p>' +
    '<p><strong>Claimed at:</strong> ' + formatDateTime(claimedAt) + '</p>' +
    '<p>Complete it before the deadline to earn your points. Finishing early earns a bonus!</p>';
  sendEmail(email, subject, body, ldap);
}

// ------------------------------------------------------------
// EMAIL — Task Completed
// ------------------------------------------------------------

function sendTaskCompletedEmail(ldap, task, totalPoints, bonusPoints, isFirst) {
  var email = getLdapEmail(ldap);
  var subject = 'Task completed: ' + task['Title'] + ' — ' + totalPoints + ' pts';
  var extras = '';
  if (isFirst) extras += '<p>🥇 <strong>First to complete!</strong> +' + getConfig('PointsFirstToComplete') + ' bonus points awarded.</p>';
  if (bonusPoints > 0) extras += '<p>⚡ <strong>Early completion bonus:</strong> +' + bonusPoints + ' pts.</p>';
  var body =
    '<p>Hi <strong>' + formatDisplayName(ldap) + '</strong>,</p>' +
    '<p>Great work! You\'ve completed a task.</p>' +
    '<div class="highlight">' +
    '<div class="pts">+' + totalPoints + ' pts</div>' +
    '<div class="label">' + task['Title'] + '</div>' +
    '</div>' +
    extras +
    '<p>Your points have been updated on the leaderboard.</p>';
  sendEmail(email, subject, body, ldap);
}

// ------------------------------------------------------------
// EMAIL — Demerit
// ------------------------------------------------------------

function sendDemeritEmail(ldap, type, details, points) {
  var email = getLdapEmail(ldap);
  var subject = 'Demerit notice: ' + type;
  var body =
    '<p>Hi <strong>' + formatDisplayName(ldap) + '</strong>,</p>' +
    '<p>A demerit has been added to your record.</p>' +
    '<div class="highlight">' +
    '<div class="pts" style="color:#c0392b;">' + points + ' pts</div>' +
    '<div class="label">' + type + '</div>' +
    '</div>' +
    '<p><strong>Details:</strong> ' + (details || 'N/A') + '</p>' +
    '<p>If you believe this is an error, please speak with your team lead.</p>';
  sendEmail(email, subject, body, ldap);
}

// ------------------------------------------------------------
// EMAIL — Kudos Approved
// ------------------------------------------------------------

function sendKudosApprovedEmail(ldap, points, caseId) {
  var email = getLdapEmail(ldap);
  var subject = 'Your Kudos was approved! +' + points + ' pts';
  var body =
    '<p>Hi <strong>' + formatDisplayName(ldap) + '</strong>,</p>' +
    '<p>Your Kudos submission has been reviewed and <strong>approved</strong>.</p>' +
    '<div class="highlight">' +
    '<div class="pts">+' + points + ' pts</div>' +
    '<div class="label">Case ID: ' + (caseId || 'N/A') + '</div>' +
    '</div>' +
    '<p>Points have been added to your leaderboard score. Keep up the great work!</p>';
  sendEmail(email, subject, body, ldap);
}

// ------------------------------------------------------------
// EMAIL — Kudos Rejected
// ------------------------------------------------------------

function sendKudosRejectedEmail(ldap, caseId, note) {
  var email = getLdapEmail(ldap);
  var subject = 'Kudos update: Case ' + (caseId || 'N/A');
  var body =
    '<p>Hi <strong>' + formatDisplayName(ldap) + '</strong>,</p>' +
    '<p>Your Kudos submission for case <strong>' + (caseId || 'N/A') + '</strong> was reviewed but could not be approved at this time.</p>' +
    (note ? '<p><strong>Reviewer note:</strong> ' + note + '</p>' : '') +
    '<p>You can submit a new Kudos for a different case. If you have questions, speak with your team lead.</p>';
  sendEmail(email, subject, body, ldap);
}
