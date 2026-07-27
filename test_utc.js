const strVal = "2026-06-27T16:00:00.000Z";
console.log(strVal.substring(0, 7)); // This incorrectly truncates UTC string to GMT.

// Let's test a reviver that correctly parses to Apps script format date string maybe?
