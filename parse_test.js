function reviver(key, value) {
  var isoDateRegex = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d*)?(?:[-+]\d{2}:?\d{2}|Z)?$/;
  if (typeof value === 'string' && isoDateRegex.test(value)) {
    return new Date(value);
  }
  return value;
}

const s = JSON.stringify({d: new Date()});
const out = JSON.parse(s, reviver);
console.log(out.d instanceof Date);
