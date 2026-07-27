// ============================================================
// TaskService.gs — Task CRUD, claim, complete, acknowledge,
//                  search, homepage feed
// Play Ops Store
// ============================================================

// ------------------------------------------------------------
// CREATE TASK (manager only)
// ------------------------------------------------------------

function createTask(taskData, createdBy) {
  var id = generateTaskId();
  var createdAt = now();

  appendRow('Tasks', {
    ID: id,
    Title: taskData.title,
    Description: taskData.description || '',
    Category: taskData.category || 'General',
    TargetType: taskData.targetType || 'team',   // 'team', 'specific', 'group'
    TargetValue: taskData.targetValue || '',      // LDAP or group name
    Deadline: taskData.deadline ? new Date(taskData.deadline) : '',
    BasePoints: taskData.basePoints || 10,
    Difficulty: taskData.difficulty || 'Easy',
    EarlyBonusRate: taskData.earlyBonusRate || '',
    Status: 'Published',
    CreatedBy: createdBy,
    CreatedAt: createdAt,
    IsFeatured: taskData.isFeatured || false
  });

  // Notify targeted agents
  notifyNewTask(id, taskData);

  Logger.log('[Task Created] ' + id + ' by ' + createdBy);
  return { success: true, taskId: id };
}

// ------------------------------------------------------------
// UPDATE TASK (manager only)
// ------------------------------------------------------------

function updateTask(taskId, taskData, updatedBy) {
  var task = findRow('Tasks', 'ID', taskId);
  if (!task) return { success: false, error: 'Task not found.' };

  var updates = {};
  if (taskData.title !== undefined)       updates['Title'] = taskData.title;
  if (taskData.description !== undefined) updates['Description'] = taskData.description;
  if (taskData.category !== undefined)    updates['Category'] = taskData.category;
  if (taskData.deadline !== undefined)    updates['Deadline'] = new Date(taskData.deadline);
  if (taskData.basePoints !== undefined)  updates['BasePoints'] = taskData.basePoints;
  if (taskData.difficulty !== undefined)  updates['Difficulty'] = taskData.difficulty;
  if (taskData.isFeatured !== undefined)  updates['IsFeatured'] = taskData.isFeatured;
  if (taskData.status !== undefined)      updates['Status'] = taskData.status;

  updateRow('Tasks', 'ID', taskId, updates);
  Logger.log('[Task Updated] ' + taskId + ' by ' + updatedBy);
  return { success: true };
}

// ------------------------------------------------------------
// GET ALL TASKS (manager task manager view)
// ------------------------------------------------------------

function getAllTasksForManager(managedLdaps) {
  var tasks = getSheetData('Tasks');
  var allCompletions = getSheetData('Completions');

  // Group completions by TaskID once, instead of scanning full sheet per task
  var completionsByTask = {};
  allCompletions.forEach(function(c) {
    var tid = c['TaskID'];
    if (!completionsByTask[tid]) completionsByTask[tid] = [];
    completionsByTask[tid].push(c);
  });

  // If supervisor, filter tasks they created or targeted at their team
  if (managedLdaps) {
    tasks = tasks.filter(function(t) {
      if (t['TargetType'] === 'team' || t['TargetType'] === 'announcement') return true;
      if (t['TargetType'] === 'specific' && managedLdaps.indexOf(t['TargetValue']) !== -1) return true;
      if (t['TargetType'] === 'group') {
        var targets = String(t['TargetValue']).split(',').map(function(s) { return s.trim(); });
        return targets.some(function(target) { return managedLdaps.indexOf(target) !== -1; });
      }
      return false;
    });
  }

  return tasks.map(function(t) {
    var completions = completionsByTask[t['ID']] || [];
    return {
      id: t['ID'],
      title: t['Title'],
      category: t['Category'],
      targetType: t['TargetType'],
      targetValue: t['TargetValue'],
      deadline: formatDate(t['Deadline']),
      basePoints: t['BasePoints'],
      difficulty: t['Difficulty'],
      status: t['Status'],
      isFeatured: t['IsFeatured'],
      createdBy: t['CreatedBy'],
      createdAt: formatDate(t['CreatedAt']),
      completionCount: completions.filter(function(c) { return c['CompletedAt']; }).length,
      claimCount: completions.length
    };
  });
}

// ------------------------------------------------------------
// HOMEPAGE FEED
// ------------------------------------------------------------

function getHomepageTasks(ldap) {
  var allTasks = getSheetData('Tasks');
  var myCompletions = findRows('Completions', 'LDAP', ldap);

  var completedTaskIds = [];
  var claimedTaskIds = [];
  myCompletions.forEach(function(c) {
    if (c['CompletedAt']) completedTaskIds.push(c['TaskID']);
    else if (c['ClaimedAt']) claimedTaskIds.push(c['TaskID']);
  });

  // Load agent data once for visibility check optimization
  var agent = findRow('Agents', 'LDAP', ldap);

  // Filter to tasks visible to this agent
  var visible = allTasks.filter(function(t) {
    return isTaskVisibleToOptimized(t, ldap, agent, completedTaskIds) && t['Status'] !== 'Expired';
  });

  // Enrich each task with agent-specific status
  var enriched = visible.map(function(t) {
    return enrichTask(t, ldap, completedTaskIds, claimedTaskIds);
  });

  // --- ADD THIS BLOCK HERE ---
  // Hide unclaimed tasks that have passed their deadline
  enriched = enriched.filter(function(t) {
    return !(t.urgency === 'expired' && t.agentStatus === 'Available');
  });
  // ---------------------------

  var rightNow = now();

  // Featured / pinned
  var featured = enriched.filter(function(t) {
    return t.isFeatured && t.status !== 'Expired';
  }).slice(0, 3);

  // For You — targeted specifically at this agent
  var forYou = enriched.filter(function(t) {
    return (t.targetType === 'specific' && t.targetValue === ldap) && t.agentStatus === 'Available';
  });

  // New This Week
  var oneWeekAgo = new Date(rightNow - 7 * 24 * 60 * 60 * 1000);
  var newThisWeek = enriched.filter(function(t) {
    return new Date(t.createdAt) >= oneWeekAgo && t.agentStatus === 'Available';
  });

  // Closing Soon — expires within 48 hours
  var in48hrs = new Date(rightNow.getTime() + 48 * 60 * 60 * 1000);
  var closingSoon = enriched.filter(function(t) {
    if (!t.deadline) return false;
    var dl = new Date(t.deadline);
    return dl <= in48hrs && dl > rightNow && t.agentStatus !== 'Completed';
  });

  // Quick Wins — Easy difficulty
  var quickWins = enriched.filter(function(t) {
    return t.difficulty === 'Easy' && t.agentStatus === 'Available';
  });

  // ADDED: Grab all completed tasks before they get thrown out
  var completedTasks = enriched.filter(function(t) {
    return t.agentStatus === 'Completed';
  });

  return {
    featured: featured,
    forYou: forYou,
    newThisWeek: newThisWeek,
    closingSoon: closingSoon,
    quickWins: quickWins,
    completedTasks: completedTasks, // ADDED: Send to frontend
    categories: getTaskCategories(enriched)
  };
}

// ------------------------------------------------------------
// TASK DETAIL PAGE
// ------------------------------------------------------------

function getTaskDetail(taskId, ldap) {
  var task = findRow('Tasks', 'ID', taskId);
  if (!task) return null;

  var myCompletions = findRows('Completions', 'LDAP', ldap);
  var completedTaskIds = [];
  var claimedTaskIds = [];
  myCompletions.forEach(function(c) {
    if (c['CompletedAt']) completedTaskIds.push(c['TaskID']);
    else if (c['ClaimedAt']) claimedTaskIds.push(c['TaskID']);
  });

  var agent = findRow('Agents', 'LDAP', ldap);
  if (!isTaskVisibleToOptimized(task, ldap, agent, completedTaskIds)) return null;

  var enriched = enrichTask(task, ldap, completedTaskIds, claimedTaskIds);

  // Team completion history (who else completed it)
  var allCompletionsForTask = findRows('Completions', 'TaskID', taskId).filter(function(c) { return c['CompletedAt']; });

  // Pre-load agents to avoid findRow in loop
  var agentsData = getSheetData('Agents');
  var agentsMap = {};
  agentsData.forEach(function(a) { agentsMap[a.LDAP] = a; });

  enriched.completionHistory = allCompletionsForTask.map(function(c) {
    var a = agentsMap[c['LDAP']];
    var dName = (a && a['DisplayName']) ? a['DisplayName'].trim() : c['LDAP'].toLowerCase();
    return {
      ldap: c['LDAP'],
      displayName: dName,
      completedAt: formatDate(c['CompletedAt']),
      totalPoints: c['TotalPoints'],
      photoUrl: getMomaPhotoUrl(c['LDAP'])
    };
  });

  // If claimed by this agent, add TAT info
  var myClaim = myCompletions.find(function(c) {
    return c['TaskID'] === taskId && c['ClaimedAt'] && !c['CompletedAt'];
  });
  if (myClaim) {
    enriched.claimedAt = formatDateTime(myClaim['ClaimedAt']);
    enriched.tatSoFar = calcTAT(myClaim['ClaimedAt'], now());
  }

  return enriched;
}

// ------------------------------------------------------------
// CLAIM TASK
// ------------------------------------------------------------

function claimTask(taskId, ldap) {
  var task = findRow('Tasks', 'ID', taskId);
  if (!task) return { success: false, error: 'Task not found.' };
  if (!isTaskVisibleTo(task, ldap)) return { success: false, error: 'Task not available to you.' };
  if (task['Status'] === 'Expired') return { success: false, error: 'This task has expired.' };
  if (isExpired(task['Deadline'])) return { success: false, error: 'This task deadline has passed.' };

  // Check if already claimed by this agent
  var existing = findRows('Completions', 'TaskID', taskId).find(function(c) {
    return c['LDAP'] === ldap && c['ClaimedAt'];
  });
  if (existing) return { success: false, error: 'You have already claimed this task.' };
  // --- NEW: TASK HOARDING LIMIT ---
  var allTasks = getSheetData('Tasks');
  var myActiveClaims = findRows('Completions', 'LDAP', ldap).filter(function(c) {
    // 1. Must be an active, uncompleted claim
    if (c['CompletedAt'] || c['Type'] !== 'Claimed') return false;
    
    // 2. Cross-reference: Does the task actually still exist in the database?
    var relatedTask = allTasks.find(function(t) { return t['ID'] === c['TaskID']; });
    
    // If the manager deleted the task, don't count it against the agent!
    if (!relatedTask) return false; 
    
    // If the task is explicitly marked as Expired, don't count it against the agent!
    if (relatedTask['Status'] === 'Expired') return false;

    return true; // This is a real, visible, valid active claim
  });

  if (myActiveClaims.length >= 3) {
    return { success: false, error: 'Limit reached: You already have 3 active claims. Check your "In Progress" filter in Browse Tasks to complete them!' };
  }
  // --------------------------------

  var compId = generateCompletionId();
  var claimedAt = now();

  appendRow('Completions', {
    ID: compId,
    TaskID: taskId,
    LDAP: ldap,
    ClaimedAt: claimedAt,
    CompletedAt: '',
    TAT: '',
    BasePoints: task['BasePoints'],
    BonusPoints: '',
    TotalPoints: '',
    IsFirst: '',
    IsEarly: '',
    Type: 'Claimed'
  });

  // Update task status for specific tasks (Case-Insensitive)
  if (String(task['TargetType']).trim().toLowerCase() === 'specific') {
    updateRow('Tasks', 'ID', taskId, { Status: 'Claimed' });
  }

  // Notify agent
  var deadlineStr = task['Deadline'] ? formatDate(task['Deadline']) : 'No deadline';
  createNotification(ldap, 'claimed', 'You claimed "' + task['Title'] + '" — deadline: ' + deadlineStr);
  sendTaskClaimedEmail(ldap, task, claimedAt);

  // Expiry warning (24hr) — scheduled via notification, actual check in nightly trigger
  Logger.log('[Claimed] ' + ldap + ' | Task: ' + taskId);
  return { success: true, compId: compId, claimedAt: formatDateTime(claimedAt) };
}

// ------------------------------------------------------------
// COMPLETE TASK
// ------------------------------------------------------------

function completeTask(taskId, ldap) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000); // Wait up to 10 seconds for other clicks to process
  
  try {
    var task = findRow('Tasks', 'ID', taskId);
    if (!task) return { success: false, error: 'Task not found.' };

  // Find the active claim row (Bulletproof Case-Insensitive & Active Type Check)
          var claimRow = findRows('Completions', 'TaskID', taskId).find(function(c) {
            return String(c['LDAP']).toLowerCase() === String(ldap).toLowerCase() && 
                   c['ClaimedAt'] && 
                   (!c['CompletedAt'] || String(c['CompletedAt']).trim() === '') && 
                   c['Type'] === 'Claimed';
          });

          if (!claimRow) return { success: false, error: 'You do not have an active claim for this task.' };

  var completedAt = now();
  var tat = calcTAT(claimRow['ClaimedAt'], completedAt);
  var basePoints = parseInt(task['BasePoints']) || 10;
  var bonusPoints = calcEarlyBonus(basePoints, completedAt, task['Deadline']);
  var days = daysEarly(completedAt, task['Deadline']);
  var isEarly = days > 0;
  var isOnTime = !isExpired(task['Deadline']);

  // First to complete bonus
  var allDone = findRows('Completions', 'TaskID', taskId).filter(function(c) { return c['CompletedAt']; });
  var isFirst = allDone.length === 0;
  if (isFirst) {
    bonusPoints += parseInt(getConfig('PointsFirstToComplete')) || 10;
  }

  var totalPoints = basePoints + bonusPoints;

  // Update completion row
  updateRow('Completions', 'ID', claimRow['ID'], {
    CompletedAt: completedAt,
    TAT: tat,
    BonusPoints: bonusPoints,
    TotalPoints: totalPoints,
    IsFirst: isFirst,
    IsEarly: isEarly,
    Type: isOnTime ? 'Completed' : 'Late'
  });

  // --- SMART STATUS UPDATE (Case-Insensitive & Auto-Close) ---
  try {
    var tType = String(task['TargetType']).trim().toLowerCase();
    
    if (tType === 'specific') {
      // Instantly complete specific tasks
      updateRow('Tasks', 'ID', taskId, { Status: 'Completed' });
    } else {
      // For Team/Group tasks, auto-complete the task globally ONLY IF everyone has finished it
      var pendingCheck = getPendingAgentsForTask(taskId, null);
      if (pendingCheck && pendingCheck.success && pendingCheck.pending.length === 0) {
        updateRow('Tasks', 'ID', taskId, { Status: 'Completed' });
      }
    }
  } catch(e) {
    Logger.log('Status Update Error: ' + e.message);
  }
  // -----------------------------------------------------------

  // Award points
  addPoints(ldap, totalPoints, 'Task completed: ' + taskId);

  // Update streak
  var streak = updateStreak(ldap, isOnTime);

  // Evaluate badges
  var totalCompletions = findRows('Completions', 'LDAP', ldap).filter(function(c) { return c['CompletedAt']; }).length;
  evaluateBadges(ldap, {
    completionCount: totalCompletions,
    isEarly: isEarly,
    daysEarly: days,
    isAcknowledge: false,
    onTimeStreak: streak
  });

  // Notify agent
  var msg = 'You completed "' + task['Title'] + '" — ' + totalPoints + ' pts awarded';
  if (isFirst) msg += ' (First to complete! +' + getConfig('PointsFirstToComplete') + ' bonus)';
  if (isEarly) msg += ' (Early bonus: +' + bonusPoints + ')';
  createNotification(ldap, 'completed', msg);
  sendTaskCompletedEmail(ldap, task, totalPoints, bonusPoints, isFirst);

  Logger.log('[Completed] ' + ldap + ' | Task: ' + taskId + ' | Points: ' + totalPoints);
  return {
      success: true,
      totalPoints: totalPoints,
      basePoints: basePoints,
      bonusPoints: bonusPoints,
      isFirst: isFirst,
      isEarly: isEarly,
      tat: tat,
      streak: streak
    };
  } finally {
    lock.releaseLock(); // Release the queue for the next agent
  }
}

// ------------------------------------------------------------
// ACKNOWLEDGE ANNOUNCEMENT
// ------------------------------------------------------------

function acknowledgeTask(taskId, ldap) {
  var task = findRow('Tasks', 'ID', taskId);
  if (!task) return { success: false, error: 'Task not found.' };
  if (task['TargetType'] !== 'announcement') return { success: false, error: 'This task is not an announcement.' };

  // Check if already acknowledged
  var existing = findRows('Completions', 'TaskID', taskId).find(function(c) {
    return c['LDAP'] === ldap && c['CompletedAt'];
  });
  if (existing) return { success: false, error: 'Already acknowledged.' };

  var compId = generateCompletionId();
  var completedAt = now();
  var basePoints = parseInt(task['BasePoints']) || 5;

  appendRow('Completions', {
    ID: compId,
    TaskID: taskId,
    LDAP: ldap,
    ClaimedAt: completedAt,
    CompletedAt: completedAt,
    TAT: 0,
    BasePoints: basePoints,
    BonusPoints: 0,
    TotalPoints: basePoints,
    IsFirst: false,
    IsEarly: false,
    Type: 'Acknowledged'
  });

  addPoints(ldap, basePoints, 'Announcement acknowledged: ' + taskId);

  // Badge check
  var totalAck = findRows('Completions', 'LDAP', ldap).filter(function(c) {
    return c['Type'] === 'Acknowledged';
  }).length;
  if (totalAck === 1) {
    awardBadgeIfNew(ldap, 'ACKNOWLEDGED');
  }

  createNotification(ldap, 'acknowledged', 'You acknowledged "' + task['Title'] + '" — ' + basePoints + ' pts awarded.');
  Logger.log('[Acknowledged] ' + ldap + ' | Task: ' + taskId);
  return { success: true, points: basePoints };
}

// ------------------------------------------------------------
// SEARCH TASKS
// ------------------------------------------------------------

function searchTasks(query, ldap) {
  if (!query || query.trim() === '') return [];
  var q = query.toLowerCase();

  var allTasks = getSheetData('Tasks');
  var myCompletions = findRows('Completions', 'LDAP', ldap);
  var completedTaskIds = [];
  var claimedTaskIds = [];
  myCompletions.forEach(function(c) {
    if (c['CompletedAt']) completedTaskIds.push(c['TaskID']);
    else if (c['ClaimedAt']) claimedTaskIds.push(c['TaskID']);
  });

  var agent = findRow('Agents', 'LDAP', ldap);

  return allTasks
    .filter(function(t) {
      return isTaskVisibleToOptimized(t, ldap, agent, completedTaskIds) &&
        t['Status'] !== 'Expired' &&
        (t['Title'].toLowerCase().indexOf(q) !== -1 ||
         t['Description'].toLowerCase().indexOf(q) !== -1 ||
         t['Category'].toLowerCase().indexOf(q) !== -1);
    })
    .map(function(t) {
      return enrichTask(t, ldap, completedTaskIds, claimedTaskIds);
    })
    // --- ADD THIS FILTER TO THE CHAIN ---
    .filter(function(t) {
      return !(t.urgency === 'expired' && t.agentStatus === 'Available');
    });
}

// ------------------------------------------------------------
// TASK VISIBILITY CHECK
// ------------------------------------------------------------

function isTaskVisibleTo(task, ldap) {
  var agent = findRow('Agents', 'LDAP', ldap);
  return isTaskVisibleToOptimized(task, ldap, agent);
}

function isTaskVisibleToOptimized(task, ldap, agent, completedTaskIds) {
  // FIX: Prevent creators from seeing their own tasks in their agent feed
  if (task['CreatedBy'] && String(task['CreatedBy']).toLowerCase() === ldap.toLowerCase()) {
    return false;
  }

  var type = task['TargetType'];
  if (type === 'team') return true;
  if (type === 'announcement') {
    // Only visible if not already acknowledged
    if (completedTaskIds) {
      return completedTaskIds.indexOf(task['ID']) === -1;
    }
    var existing = findRows('Completions', 'TaskID', task['ID']).find(function(c) {
      return c['LDAP'] === ldap && c['CompletedAt'];
    });
    return !existing;
  }
  if (type === 'specific') return String(task['TargetValue']).toLowerCase() === ldap.toLowerCase();
  if (type === 'group') {
    // Force both targets and agent data to lowercase for flawless matching
    var targets = String(task['TargetValue']).split(',').map(function(s) { return s.trim().toLowerCase(); });
    var ldapLower = ldap.toLowerCase();
    if (targets.indexOf(ldapLower) !== -1) return true;
    
    if (agent && agent['Workgroup'] && targets.indexOf(agent['Workgroup'].toLowerCase()) !== -1) return true;
    if (agent && agent['TeamLead'] && targets.indexOf(agent['TeamLead'].toLowerCase()) !== -1) return true;
  }
  return false;
}

// ------------------------------------------------------------
// TASK ENRICHMENT (adds agent-specific status + display fields)
// ------------------------------------------------------------

function enrichTask(task, ldap, completedTaskIds, claimedTaskIds) {
  var taskId = task['ID'];
  var agentStatus = 'Available';
  if (completedTaskIds.indexOf(taskId) !== -1) agentStatus = 'Completed';
  else if (claimedTaskIds.indexOf(taskId) !== -1) agentStatus = 'In Progress';

  var deadline = task['Deadline'] ? new Date(task['Deadline']) : null;
  var hoursLeft = deadline ? hoursUntil(deadline) : null;
  var urgency = 'normal';
  if (hoursLeft !== null) {
    if (hoursLeft < 0) urgency = 'expired';
    else if (hoursLeft < 24) urgency = 'critical';
    else if (hoursLeft < 48) urgency = 'warning';
  }

  return {
    id: taskId,
    title: task['Title'],
    description: task['Description'],
    category: task['Category'],
    targetType: task['TargetType'],
    targetValue: task['TargetValue'],
    deadline: deadline ? formatDate(deadline) : null,
    deadlineRaw: deadline ? deadline.toISOString() : null,
    basePoints: parseInt(task['BasePoints']) || 0,
    difficulty: task['Difficulty'],
    status: task['Status'],
    isFeatured: task['IsFeatured'] === true || task['IsFeatured'] === 'TRUE',
    createdAt: task['CreatedAt'] ? new Date(task['CreatedAt']).toISOString() : null,
    agentStatus: agentStatus,
    urgency: urgency,
    hoursLeft: hoursLeft !== null ? Math.round(hoursLeft) : null
  };
}

// ------------------------------------------------------------
// CATEGORIES
// ------------------------------------------------------------

function getTaskCategories(enrichedTasks) {
  var cats = {};
  enrichedTasks.forEach(function(t) {
    if (t.category) cats[t.category] = (cats[t.category] || 0) + 1;
  });
  return Object.keys(cats).map(function(c) {
    return { name: c, count: cats[c] };
  }).sort(function(a, b) { return b.count - a.count; });
}

// ------------------------------------------------------------
// TASK NOTIFICATION HELPERS
// ------------------------------------------------------------

function notifyNewTask(taskId, taskData) {
  var type = taskData.targetType || 'team';
  var agents = [];

  if (type === 'team' || type === 'announcement') {
    agents = getSheetData('Agents').map(function(a) { return a['LDAP']; });
  } else if (type === 'specific') {
    agents = [taskData.targetValue];
  } else if (type === 'group') {
    var targets = taskData.targetValue.split(',').map(function(s) { return s.trim(); });
    var allAgents = getSheetData('Agents');
    agents = allAgents
      .filter(function(a) {
        return targets.indexOf(a['LDAP']) !== -1 || targets.indexOf(a['Workgroup']) !== -1;
      })
      .map(function(a) { return a['LDAP']; });
  }

  var notifications = agents.map(function(ldap) {
    return {
      ldap: ldap,
      type: 'new_task',
      message: 'New task available: "' + taskData.title + '" — ' + (taskData.basePoints || 10) + ' pts'
    };
  });
  createNotifications(notifications);
}

// ------------------------------------------------------------
// EXPIRY WARNING NOTIFICATIONS (called from nightly trigger)
// ------------------------------------------------------------

function sendExpiryWarnings() {
  var completions = getSheetData('Completions');
  var tasks = getSheetData('Tasks');
  var allNotifications = getSheetData('Notifications');

  var newNotifications = [];

  completions.forEach(function(c) {
    if (!c['ClaimedAt'] || c['CompletedAt']) return;

    var task = tasks.find(function(t) { return t['ID'] === c['TaskID']; });
    if (!task || !task['Deadline']) return;

    var hrs = hoursUntil(task['Deadline']);
    if (hrs > 0 && hrs <= 24) {
      // Check we haven't already warned them
      var alreadyWarned = allNotifications.some(function(n) {
        return n['LDAP'] === c['LDAP'] && n['Type'] === 'expiry_warning' && n['Message'].indexOf(c['TaskID']) !== -1;
      });

      if (!alreadyWarned) {
        newNotifications.push({
          ldap: c['LDAP'],
          type: 'expiry_warning',
          message: '"' + task['Title'] + '" expires in ' + Math.round(hrs) + ' hours! Complete it now. (Task ID: ' + c['TaskID'] + ')'
        });
      }
    }
  });

  if (newNotifications.length > 0) {
    createNotifications(newNotifications);
  }
}
// ------------------------------------------------------------
// TRACK MISSING AGENTS (Manager Only)
// ------------------------------------------------------------
function getPendingAgentsForTask(taskId, managedLdaps) {
  var task = findRow('Tasks', 'ID', taskId);
  if (!task) return { success: false, error: 'Task not found' };

  var allAgents = getSheetData('Agents');

  // FIX: Remove the task creator from the required agents list
  var creatorLdap = String(task['CreatedBy']).trim().toLowerCase();
  allAgents = allAgents.filter(function(a) {
    return String(a['LDAP']).trim().toLowerCase() !== creatorLdap;
  });

  if (managedLdaps) {
    allAgents = allAgents.filter(function(a) { return managedLdaps.indexOf(a['LDAP']) !== -1; });
  }
  var targetAgents = [];

  // 1. Determine who SHOULD complete this task
  var type = String(task['TargetType']).trim().toLowerCase(); // FIX: Bulletproof casing
  if (type === 'team' || type === 'announcement') {
    targetAgents = allAgents;
  } else if (type === 'specific') {
    var targetLdap = String(task['TargetValue']).trim().toLowerCase();
    targetAgents = allAgents.filter(function(a) { return String(a['LDAP']).toLowerCase() === targetLdap; });
  } else if (type === 'group') {
    var targets = String(task['TargetValue']).split(',').map(function(s) { return s.trim().toLowerCase(); });
    targetAgents = allAgents.filter(function(a) {
      // Force everything to lowercase to prevent matching errors
      var aLdap = String(a['LDAP']).toLowerCase();
      var aWorkgroup = a['Workgroup'] ? String(a['Workgroup']).toLowerCase() : '';
      var aTeamLead = a['TeamLead'] ? String(a['TeamLead']).toLowerCase() : '';
      
      return targets.indexOf(aLdap) !== -1 || 
             targets.indexOf(aWorkgroup) !== -1 || 
             targets.indexOf(aTeamLead) !== -1;
    });
  }

  // 2. Find who HAS completed it
  var completions = findRows('Completions', 'TaskID', taskId).filter(function(c) {
    // Strictly ensure they actually have a completion timestamp, not just a claim
    return c['CompletedAt'] && String(c['CompletedAt']).trim() !== '';
  });
  // Convert all completed LDAPs to lowercase for safe matching
  var completedLdaps = completions.map(function(c) { return String(c['LDAP']).trim().toLowerCase(); });

  // 3. Find the missing agents
  var pending = targetAgents.filter(function(a) {
    return completedLdaps.indexOf(String(a['LDAP']).trim().toLowerCase()) === -1;
  }).map(function(a) {
    // Fallback to LDAP if DisplayName is somehow blank in the DB
    var dName = a['DisplayName'] && a['DisplayName'] !== '' ? a['DisplayName'] : a['LDAP'];
    return { ldap: a['LDAP'], displayName: dName };
  });

  return { success: true, pending: pending, taskTitle: task['Title'] };
}
// ------------------------------------------------------------
// COMPLIANCE MATRIX (Manager View)
// ------------------------------------------------------------

// FIX: Added categoryFilter to the function arguments right here!
function getComplianceMatrixData(managedLdaps, timeframe, statusFilter, categoryFilter) {
  var allTasks = getSheetData('Tasks');
  var allAgents = getSheetData('Agents');
  var allCompletions = getSheetData('Completions');
  
  // 1. Filter Agents (Only show those the manager is allowed to see)
  if (managedLdaps) {
    allAgents = allAgents.filter(function(a) { return managedLdaps.indexOf(a['LDAP']) !== -1; });
  }

  // 2. Filter Tasks based on filters
  var now = new Date();
  var filteredTasks = allTasks.filter(function(t) {
    
    // FIX: Category Filter (Now safely outside the status block!)
    if (categoryFilter && categoryFilter !== 'all' && t['Category'] !== categoryFilter) return false;

    // Status Filter
    if (statusFilter !== 'all') {
      if (statusFilter === 'active' && (t['Status'] === 'Expired' || t['Status'] === 'Completed')) return false;
      if (statusFilter === 'expired' && t['Status'] !== 'Expired') return false;
      if (statusFilter === 'completed' && t['Status'] !== 'Completed') return false;
    }

    // Timeframe Filter (Based on Deadline)
    if (timeframe !== 'all' && t['Deadline']) {
      var dl = new Date(t['Deadline']);
      var diffDays = (dl - now) / (1000 * 60 * 60 * 24);
      
      if (timeframe === 'today' && (dl.getDate() !== now.getDate() || dl.getMonth() !== now.getMonth())) return false;
      if (timeframe === 'week' && (diffDays < -7 || diffDays > 7)) return false;
      if (timeframe === 'month' && dl.getMonth() !== now.getMonth()) return false;
    }
    
    return true;
  });

  // Sort tasks by deadline
  filteredTasks.sort(function(a, b) {
    if (!a['Deadline']) return 1;
    if (!b['Deadline']) return -1;
    return new Date(a['Deadline']) - new Date(b['Deadline']);
  });

  // 3. Build the Matrix
  var matrix = allAgents.map(function(agent) {
    var ldap = String(agent['LDAP']).toLowerCase();
    var dName = agent['DisplayName'] ? agent['DisplayName'] : agent['LDAP'];
    
    var taskStatuses = filteredTasks.map(function(task) {
      var taskId = task['ID'];
      var isTargeted = false;
      
      // Check if this task is actually assigned to this agent
      var tType = String(task['TargetType']).toLowerCase();
      if (tType === 'team' || tType === 'announcement') {
        isTargeted = true;
      } else if (tType === 'specific' && String(task['TargetValue']).toLowerCase() === ldap) {
        isTargeted = true;
      } else if (tType === 'group') {
        var targets = String(task['TargetValue']).toLowerCase().split(',').map(function(s) { return s.trim(); });
        var aWorkgroup = agent['Workgroup'] ? String(agent['Workgroup']).toLowerCase() : '';
        var aTeamLead = agent['TeamLead'] ? String(agent['TeamLead']).toLowerCase() : '';
        if (targets.indexOf(ldap) !== -1 || targets.indexOf(aWorkgroup) !== -1 || targets.indexOf(aTeamLead) !== -1) {
          isTargeted = true;
        }
      }

      // If not assigned to them, return N/A
      if (!isTargeted) return { taskId: taskId, status: 'N/A' };

      // Check their completion status
      var myClaims = allCompletions.filter(function(c) { 
        return String(c['LDAP']).toLowerCase() === ldap && c['TaskID'] === taskId; 
      });

      var cellStatus = 'Available'; // Default
      
      if (myClaims.length > 0) {
        var claim = myClaims[myClaims.length - 1]; // Get latest
        if (claim['CompletedAt']) cellStatus = 'Completed';
        else if (claim['Type'] === 'Missed') cellStatus = 'Missed';
        else cellStatus = 'In Progress';
      } else if (task['Status'] === 'Expired') {
        cellStatus = 'Missed';
      }

      return { taskId: taskId, status: cellStatus };
    });

    return {
      ldap: agent['LDAP'],
      displayName: dName,
      photoUrl: 'https://moma-teams-photos.corp.google.com/photos/' + agent['LDAP'] + '?sz=600&type=PLUS',
      tasks: taskStatuses
    };
  });

  return {
    success: true,
    // FIX: Included 'category' in the return map so the frontend can populate the dropdown
    tasks: filteredTasks.map(function(t) { return { id: t['ID'], title: t['Title'], category: t['Category'], deadline: t['Deadline'] ? Utilities.formatDate(new Date(t['Deadline']), Session.getScriptTimeZone(), 'MMM dd') : 'No Date' }; }),
    matrix: matrix
  };
}

// ------------------------------------------------------------
// RETURN TASK (Agent UX)
// ------------------------------------------------------------
function returnTask(taskId, ldap) {
  var claimRow = findRows('Completions', 'TaskID', taskId).find(function(c) {
    return String(c['LDAP']).toLowerCase() === String(ldap).toLowerCase() && 
           c['Type'] === 'Claimed' && 
           (!c['CompletedAt'] || String(c['CompletedAt']).trim() === '');
  });
  
  if (!claimRow) return { success: false, error: 'You do not have an active claim.' };

  // Mark the claim as Returned so it frees up the agent's slot
  updateRow('Completions', 'ID', claimRow['ID'], { Type: 'Returned' });

  // If it was a Specific task, revert it to Published so they (or a manager) can interact with it again
  var task = findRow('Tasks', 'ID', taskId);
  if (task && String(task['TargetType']).trim().toLowerCase() === 'specific') {
    updateRow('Tasks', 'ID', taskId, { Status: 'Published' });
  }
  
  Logger.log('[Task Returned] ' + ldap + ' | Task: ' + taskId);
  return { success: true };
}

// ------------------------------------------------------------
// BULK OPERATIONS (Manager UX)
// ------------------------------------------------------------
function bulkDeleteTasks(taskIds) {
  var sheet = getSheet('Tasks');
  var data = sheet.getDataRange().getValues();
  var deletedCount = 0;
  
  // Loop backwards to safely delete multiple rows without messing up the indexes
  for (var i = data.length - 1; i > 0; i--) {
    if (taskIds.indexOf(data[i][0]) !== -1) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }
  clearSheetDataCache('Tasks');
  return { success: true, count: deletedCount };
}

function bulkExpireTasks(taskIds) {
  var updates = {};
  taskIds.forEach(function(id) { updates[id] = { Status: 'Expired' }; });
  batchUpdateRows('Tasks', 'ID', updates);
  return { success: true, count: taskIds.length };
}
