const str = JSON.stringify({ d: new Date(), a: "2026-06-27T16:00:00.000Z", b: "test" });
const out = JSON.parse(str, function(k, v) {
  if (typeof v === 'string' && v.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/)) {
    return new Date(v);
  }
  return v;
});
console.log(out.d instanceof Date);
console.log(out.a instanceof Date);
console.log(typeof out.b);
