// Display-only Excel number formats. Values and formula caches remain untouched.
function sections(format) {
  const result = [''];
  let quoted = false;
  for (let i = 0; i < format.length; i++) {
    const ch = format[i];
    if (ch === '\\' && i + 1 < format.length) { result[result.length - 1] += ch + format[++i]; continue; }
    if (ch === '"') quoted = !quoted;
    if (ch === ';' && !quoted) result.push(''); else result[result.length - 1] += ch;
  }
  return quoted || result.length > 4 ? null : result;
}
function readable(format) {
  let visible = '', active = '', quoted = false;
  const literal = value => { visible += value; active += '\0'.repeat(value.length); };
  for (let i = 0; i < format.length; i++) {
    const ch = format[i];
    if (ch === '"') { quoted = !quoted; continue; }
    if (quoted) { literal(ch); continue; }
    if (ch === '\\') { literal(format[++i] || ''); continue; }
    if (ch === '_' || ch === '*') { i++; if (ch === '_') literal(' '); continue; }
    if (ch === '[') {
      const end = format.indexOf(']', i);
      if (end < 0) return null;
      const value = format.slice(i + 1, end);
      if (/^[hms]$/i.test(value)) { visible += `[${value}]`; active += `[${value}]`; }
      else if (/^\$/.test(value)) literal(value.slice(1).split('-')[0]);
      else if (!/^(?:Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\d+)$/i.test(value)) return null;
      i = end; continue;
    }
    visible += ch; active += ch;
  }
  return { visible, active };
}
function dateDisplay(value, format, date1904) {
  if (value < 0 || value > 2958465) return null;
  const { active, visible } = format;
  const tokens = [...active.matchAll(/\[[hms]\]|AM\/PM|A\/P|y{1,4}|m{1,5}|d{1,4}|h{1,2}|s{1,2}/ig)];
  if (!tokens.length) return null;
  const days = Math.floor(value);
  const date = new Date(Date.UTC(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 31) + (days - (!date1904 && days >= 60 ? 1 : 0)) * 86400000);
  const fakeLeapDay = !date1904 && days === 60;
  const year = fakeLeapDay ? 1900 : date.getUTCFullYear();
  const month = fakeLeapDay ? 2 : !date1904 && days === 0 ? 1 : date.getUTCMonth() + 1;
  const day = fakeLeapDay ? 29 : !date1904 && days === 0 ? 0 : date.getUTCDate();
  const seconds = Math.floor((value - days) * 86400 + 0.00001);
  const hour = Math.floor(seconds / 3600) % 24, minute = Math.floor(seconds / 60) % 60, second = seconds % 60;
  const ampm = tokens.some(token => /a/i.test(token[0]));
  const pad = (number, count) => String(number).padStart(count, '0');
  let output = '', cursor = 0;
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i], code = token[0].toLowerCase(), n = code.length;
    output += visible.slice(cursor, token.index);
    let text;
    if (code.startsWith('[')) text = String(Math.floor(value * ({ h: 24, m: 1440, s: 86400 }[code[1]]) + 0.00001));
    else if (code === 'am/pm' || code === 'a/p') text = (hour < 12 ? 'AM' : 'PM').slice(0, code === 'a/p' ? 1 : 2);
    else if (code[0] === 'y') text = n <= 2 ? pad(year % 100, n) : pad(year, 4);
    else if (code[0] === 'd') text = n <= 2 ? pad(day, n) : ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][date.getUTCDay()] + (n >= 4 ? ['day', 'day', 'sday', 'nesday', 'rsday', 'day', 'urday'][date.getUTCDay()] : '');
    else if (code[0] === 'h') text = pad(ampm ? hour % 12 || 12 : hour, n);
    else if (code[0] === 's') text = pad(second, n);
    else {
      const isMinute = /h/i.test(tokens[i - 1]?.[0] || '') || /^s/i.test(tokens[i + 1]?.[0] || '');
      text = isMinute ? pad(minute, n) : n <= 2 ? pad(month, n) : ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][month - 1];
      if (!isMinute && n >= 4) text = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][month - 1];
      if (!isMinute && n === 5) text = text[0];
    }
    output += text; cursor = token.index + token[0].length;
  }
  return output + visible.slice(cursor);
}

export function formatSpreadsheetValue(value, code = 'General', { date1904 = false } = {}) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value !== 'number' || !Number.isFinite(value) || !code || /^general$|^@$/i.test(code)) return String(value);
  const parts = sections(String(code));
  if (!parts) return String(value);
  const chosen = value < 0 && parts.length > 1 ? 1 : value === 0 && parts.length > 2 ? 2 : 0;
  if (parts[chosen] === '') return '';
  const format = readable(parts[chosen]);
  if (!format) return String(value);
  if (/[ymdhs]/i.test(format.active)) return dateDisplay(value, format, date1904) ?? String(value);
  const match = /[0#?][0#?,]*(?:\.[0#?]+)?(?:[Ee][+-]0+)?/.exec(format.active);
  if (!match) return format.visible.includes('@') ? format.visible.replaceAll('@', String(value)) : format.visible;
  if (match[0].includes('?')) return String(value); // Fraction/alignment patterns are not interpreted yet.
  let number = Math.abs(value);
  const percentCount = (format.active.match(/%/g) || []).length;
  number *= 100 ** percentCount;
  const scaling = (match[0].match(/,+$/) || [''])[0].length;
  number /= 1000 ** scaling;
  const decimal = (match[0].split(/[Ee]/)[0].split('.')[1] || '');
  const maximum = Math.min(15, decimal.length), minimum = Math.min(maximum, (decimal.match(/0/g) || []).length);
  let text;
  if (/[Ee]/.test(match[0])) {
    const [mantissa, exponent] = number.toExponential(maximum).split('e');
    const exponentWidth = match[0].split(/[Ee][+-]/)[1]?.length || 2;
    text = mantissa + 'E' + (Number(exponent) < 0 ? '-' : '+') + String(Math.abs(Number(exponent))).padStart(exponentWidth, '0');
  } else {
    const integerMask = match[0].split('.')[0].replace(/,+$/, '');
    text = new Intl.NumberFormat('en-US', { useGrouping: integerMask.includes(','), minimumFractionDigits: minimum, maximumFractionDigits: maximum }).format(number);
    if (/^0+$/.test(integerMask)) { const [integer, fraction] = text.split('.'); text = integer.padStart(integerMask.length, '0') + (fraction === undefined ? '' : '.' + fraction); }
  }
  const sign = value < 0 && chosen === 0 ? '-' : '';
  return sign + format.visible.slice(0, match.index) + text + format.visible.slice(match.index + match[0].length);
}
