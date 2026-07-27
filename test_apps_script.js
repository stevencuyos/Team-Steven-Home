// If the sheet contains actual Google Apps script Date objects, they are serialized via JSON.stringify to `2026-06-27T16:00:00.000Z`.
// What is the local month for `2026-06-27T16:00:00.000Z` in Manila (+8)?
const d = new Date("2026-06-27T16:00:00.000Z");
// local hours is 16 + 8 = 24 (next day, June 28th)
// What if it is `2026-06-30T16:00:00.000Z` ?
const d2 = new Date("2026-06-30T16:00:00.000Z");
// local hours is 16+8 = 24 -> July 1st!
console.log(d2.getFullYear() + "-" + String(d2.getMonth() + 1).padStart(2, '0'));
