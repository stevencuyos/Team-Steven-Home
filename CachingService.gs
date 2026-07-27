// ============================================================
// CachingService.gs — Advanced caching layer for performance
// Implements server-side and client-side caching strategies
// ============================================================

/**
 * Wrapper for cached function execution with TTL
 * @param {string} cacheKey - Unique cache key
 * @param {Function} fn - Function to execute if cache miss
 * @param {number} ttlSeconds - Time to live in seconds
 * @param {Array} dependencies - Array of cache keys to invalidate when this updates
 */
function cachedExecution(cacheKey, fn, ttlSeconds, dependencies) {
  var cached = getCached(cacheKey);
  if (cached) {
    Logger.log('[CACHE HIT] ' + cacheKey);
    return cached;
  }
  
  Logger.log('[CACHE MISS] ' + cacheKey + ' - Executing...');
  var result = fn();
  
  setCached(cacheKey, result, ttlSeconds || 300);
  
  // If dependencies provided, invalidate them
  if (dependencies && dependencies.length > 0) {
    invalidateCache(dependencies);
  }
  
  return result;
}

/**
 * Batch API execution - combine multiple requests into one
 * Reduces number of google.script.run calls
 */
var BatchQueue = {
  queue: {},
  timers: {},
  
  /**
   * Add function to batch queue
   * @param {string} batchKey - Batch identifier
   * @param {string} fnName - Function name to call
   * @param {Array} args - Arguments
   * @param {number} delayMs - Delay before executing batch (default 100ms)
   */
  add: function(batchKey, fnName, args, delayMs) {
    delayMs = delayMs || 100;
    
    if (!this.queue[batchKey]) {
      this.queue[batchKey] = [];
    }
    
    this.queue[batchKey].push({
      fnName: fnName,
      args: args
    });
    
    // Clear existing timer for this batch
    if (this.timers[batchKey]) {
      clearTimeout(this.timers[batchKey]);
    }
    
    // Set new timer to execute batch
    var self = this;
    this.timers[batchKey] = setTimeout(function() {
      self.execute(batchKey);
    }, delayMs);
  },
  
  /**
   * Execute all queued functions in batch
   */
  execute: function(batchKey) {
    var items = this.queue[batchKey];
    if (!items || items.length === 0) return;
    
    Logger.log('[BATCH] Executing ' + items.length + ' operations for ' + batchKey);
    
    var results = {};
    items.forEach(function(item, index) {
      var globalFn = window[item.fnName];
      if (globalFn) {
        results[index] = globalFn.apply(null, item.args);
      }
    });
    
    delete this.queue[batchKey];
    delete this.timers[batchKey];
    
    return results;
  },
  
  clear: function(batchKey) {
    if (this.timers[batchKey]) {
      clearTimeout(this.timers[batchKey]);
    }
    delete this.queue[batchKey];
    delete this.timers[batchKey];
  }
};

/**
 * Client-side cache manager (for frontend use)
 * Stores data in localStorage with TTL
 */
var ClientCacheManager = {
  prefix: 'playops_cache_',
  
  /**
   * Set a value in cache with TTL
   */
  set: function(key, value, ttlMinutes) {
    ttlMinutes = ttlMinutes || 30;
    
    var expiresAt = Date.now() + (ttlMinutes * 60 * 1000);
    try {
      localStorage.setItem(this.prefix + key, JSON.stringify({
        data: value,
        expiresAt: expiresAt,
        createdAt: Date.now()
      }));
      Logger.log('[CLIENT CACHE] Set: ' + key + ' (TTL: ' + ttlMinutes + 'min)');
    } catch(e) {
      Logger.log('[CLIENT CACHE ERROR] ' + e.message);
    }
  },
  
  /**
   * Get a value from cache if not expired
   */
  get: function(key) {
    try {
      var stored = localStorage.getItem(this.prefix + key);
      if (!stored) return null;
      
      var parsed = JSON.parse(stored);
      if (Date.now() > parsed.expiresAt) {
        localStorage.removeItem(this.prefix + key);
        Logger.log('[CLIENT CACHE] Expired: ' + key);
        return null;
      }
      
      Logger.log('[CLIENT CACHE] Hit: ' + key);
      return parsed.data;
    } catch(e) {
      Logger.log('[CLIENT CACHE ERROR] ' + e.message);
      return null;
    }
  },
  
  /**
   * Remove a key from cache
   */
  remove: function(key) {
    localStorage.removeItem(this.prefix + key);
    Logger.log('[CLIENT CACHE] Removed: ' + key);
  },
  
  /**
   * Clear all cache entries matching pattern
   */
  clearPattern: function(pattern) {
    var keysToRemove = [];
    for (var i = 0; i < localStorage.length; i++) {
      var key = localStorage.key(i);
      if (key.startsWith(this.prefix) && key.includes(pattern)) {
        keysToRemove.push(key);
      }
    }
    keysToRemove.forEach(function(key) {
      localStorage.removeItem(key);
    });
    Logger.log('[CLIENT CACHE] Cleared ' + keysToRemove.length + ' entries matching: ' + pattern);
  },
  
  /**
   * Get cache stats
   */
  getStats: function() {
    var stats = {
      total: 0,
      expired: 0,
      valid: 0,
      size: 0
    };
    
    for (var i = 0; i < localStorage.length; i++) {
      var key = localStorage.key(i);
      if (key.startsWith(this.prefix)) {
        stats.total++;
        var value = localStorage.getItem(key);
        stats.size += value.length;
        
        try {
          var parsed = JSON.parse(value);
          if (Date.now() > parsed.expiresAt) {
            stats.expired++;
          } else {
            stats.valid++;
          }
        } catch(e) {
          // Skip parsing errors
        }
      }
    }
    
    return stats;
  }
};

/**
 * Intelligent batch task retrieval with smart caching
 */
function clientGetHomepageTasksOptimized() {
  return cachedExecution(
    'homepage_tasks_' + new Date().toISOString().split('T')[0],
    function() {
      return clientGetHomepageTasks();
    },
    600, // 10 minute cache
    [] // No dependencies
  );
}

/**
 * Batch request handler for multiple dashboard components
 * Instead of 5 separate API calls, make 1 batch call
 */
function clientGetDashboardData() {
  return cachedExecution(
    'dashboard_data_' + new Date().toISOString().split('T')[0],
    function() {
      return {
        tasks: clientGetHomepageTasks(),
        profile: clientGetMyProfile(),
        notifications: clientGetNotifications(),
        leaderboard: clientGetLeaderboard()
      };
    },
    600, // 10 minute cache
    []
  );
}

/**
 * Smart team analytics cache with invalidation
 */
function clientGetTeamAnalyticsOptimized() {
  return cachedExecution(
    'team_analytics_' + getCurrentLdap(),
    function() {
      return clientGetTeamAnalytics();
    },
    1800, // 30 minute cache
    ['leaderboard_' + getCurrentLdap()] // Invalidate leaderboard on update
  );
}

/**
 * Leaderboard with daily caching
 */
function clientGetLeaderboardOptimized() {
  return cachedExecution(
    'leaderboard_' + new Date().toISOString().split('T')[0],
    function() {
      return clientGetLeaderboard();
    },
    3600, // 1 hour cache
    []
  );
}

/**
 * Agent profile with smart caching
 * Cache per-user data longer, invalidate on updates
 */
function clientGetMyProfileOptimized() {
  return cachedExecution(
    'profile_' + getCurrentLdap(),
    function() {
      return clientGetMyProfile();
    },
    7200, // 2 hour cache - profiles change less frequently
    []
  );
}

/**
 * CSAT data with month-based caching
 */
function clientGetMyCsatOptimized(ldap, month) {
  month = month || new Date().toISOString().substring(0, 7);
  
  return cachedExecution(
    'csat_' + (ldap || getCurrentLdap()) + '_' + month,
    function() {
      return clientGetMyCsat(ldap, month);
    },
    1800, // 30 minute cache
    []
  );
}

/**
 * Team CSAT data with month-based caching
 */
function clientGetTeamCsatOptimized(month) {
  month = month || new Date().toISOString().substring(0, 7);
  
  return cachedExecution(
    'team_csat_' + getCurrentLdap() + '_' + month,
    function() {
      return clientGetTeamCsat(month);
    },
    1800, // 30 minute cache
    []
  );
}

/**
 * Batch search with caching
 * Cache search results for 5 minutes
 */
function clientSearchTasksOptimized(query) {
  if (!query || query.trim().length < 3) {
    return [];
  }
  
  var cacheKey = 'search_' + sanitizeInput(query).substring(0, 50);
  
  return cachedExecution(
    cacheKey,
    function() {
      return clientSearchTasks(query);
    },
    300, // 5 minute cache
    []
  );
}

/**
 * Invalidate all user-specific caches (call on updates)
 */
function invalidateUserCaches() {
  var ldap = getCurrentLdap();
  var today = new Date().toISOString().split('T')[0];
  
  var keysToInvalidate = [
    'profile_' + ldap,
    'dashboard_data_' + today,
    'homepage_tasks_' + today,
    'team_analytics_' + ldap,
    'leaderboard_' + today
  ];
  
  invalidateCache(keysToInvalidate);
  Logger.log('[CACHE] Invalidated ' + keysToInvalidate.length + ' user caches');
}

/**
 * Invalidate all caches (use sparingly - call on major data updates)
 */
function invalidateAllCaches() {
  var cache = CacheService.getScriptCache();
  cache.removeAll(cache.getAll());
  Logger.log('[CACHE] All caches invalidated');
}

/**
 * Get cache statistics for monitoring
 */
function getCacheStats() {
  return {
    timestamp: new Date(),
    app: 'Play Ops Store',
    cacheService: CacheService.getScriptCache() ? 'Available' : 'Unavailable',
    notes: 'Check browser console for CLIENT CACHE stats'
  };
}
