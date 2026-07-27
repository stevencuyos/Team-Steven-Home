const serialized = JSON.stringify([new Date("2026-07-01T00:00:00.000+08:00")]);
console.log(serialized);

const parsed = JSON.parse(serialized, function(key, value) {
  if (typeof value === 'string' && value.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/)) {
    return new Date(value);
  }
  return value;
});

console.log(parsed[0] instanceof Date);
console.log(parsed[0].toISOString());
