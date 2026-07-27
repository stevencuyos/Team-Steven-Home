const d = new Date("2026-06-30T16:00:00.000Z");
// In node, this prints the local timezone date.
console.log(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
console.log(d.toString());
