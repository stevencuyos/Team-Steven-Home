function normalizeQualityMonth(val) {
  if (!val) return '';
  if (val instanceof Date) {
    // mock Utilities
    return "date";
  }
  var strVal = String(val).trim();

  // Handle ISO date strings from JSON parsing
  if (strVal.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/)) {
    return strVal.substring(0, 7);
  }

  // Handle M/D/YYYY or MM/DD/YYYY strings from raw sheet data
  var mdYyyyMatch = strVal.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mdYyyyMatch) {
    var month = mdYyyyMatch[1].length === 1 ? '0' + mdYyyyMatch[1] : mdYyyyMatch[1];
    var year = mdYyyyMatch[3];
    return year + '-' + month;
  }

  return strVal;
}

console.log(normalizeQualityMonth("2026-06-27T16:00:00.000Z"));
console.log(normalizeQualityMonth("2026-07"));
console.log(normalizeQualityMonth("7/11/2026"));
console.log(normalizeQualityMonth("07/11/2026"));
