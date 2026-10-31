(function (root) {
  const MAX_GROUP = 1000000000;
  const YEAR_MIN_START = 1900;
  const YEAR_MAX_START = 2198;
  const YEAR_SOFT_SPAN = 10;
  const SPACE_RE = /[\s\u00A0\u2009\u202F]/g;

  function toLatin(str) {
    return String(str)
      .replace(/[\u0660-\u0669]/g, d => String(d.charCodeAt(0) - 0x0660))
      .replace(/[\u06F0-\u06F9]/g, d => String(d.charCodeAt(0) - 0x06F0))
      .replace(/\u066B/g, '.')
      .replace(/\u066C/g, '');
  }

  const normalizeDigits = str => toLatin(str).replace(SPACE_RE, '');

  function sanitizeDecimal(str) {
    let v = normalizeDigits(str).replace(/[^\d.]/g, '');
    const i = v.indexOf('.');
    if (i !== -1) v = v.slice(0, i + 1) + v.slice(i + 1).replace(/\./g, '');
    return v;
  }

  function smartNumber(str) {
    let s = normalizeDigits(str);
    const commas = (s.match(/,/g) || []).length;
    if (commas) {
      if (s.includes('.') || commas > 1 || (/,\d{3}$/.test(s) && !/^0,/.test(s))) s = s.replace(/,/g, '');
      else s = s.replace(',', '.');
    }
    return sanitizeDecimal(s);
  }

  function groupLive(clean) {
    if (!clean) return '';
    let [int, frac] = clean.split('.');
    int = int.replace(/^0+(?=\d)/, '');
    if (!int && frac !== undefined) int = '0';
    const g = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return frac !== undefined ? `${g}.${frac.slice(0, 3)}` : g;
  }

  function parseDecimal(str) {
    const n = parseFloat(sanitizeDecimal(str));
    if (isNaN(n) || n <= 0) return 0;
    return Math.min(Math.round(n * 1000) / 1000, MAX_GROUP);
  }

  const foldArabic = s => String(s)
    .replace(/[\u064B-\u065F\u0670ـ]/g, '')
    .replace(/[أإآٱ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه').replace(/ؤ/g, 'و').replace(/ئ/g, 'ي')
    .replace(/\s+/g, ' ').trim().toLowerCase();

  function wordScore(w, s) {
    if (s.startsWith(w)) return 0;
    if ((' ' + s).includes(' ' + w)) return 1;
    if (s.includes(w)) return 2;
    if (w.length > 2) {
      let k = 0;
      for (const ch of s) if (ch === w[k]) k++;
      if (k === w.length) return 4;
    }
    return Infinity;
  }

  function textScore(q, title, kw = '') {
    const t = foldArabic(title), k = foldArabic(kw), f = foldArabic(q);
    if (!f) return 0;
    if (t === f) return -2;
    if (t.startsWith(f)) return -1;
    let total = 0;
    for (const w of f.split(' ')) {
      const s = Math.min(wordScore(w, t), wordScore(w, k) + 0.6);
      if (s === Infinity) return null;
      total += s;
    }
    return total;
  }

  function findGroups(q, names, folded) {
    const f = foldArabic(q);
    if (!f) return [];
    const pool = folded || names.map(foldArabic);
    const score = n => n.startsWith(f) ? 0 : n.split(' ').some(w => w.startsWith(f)) ? 1 : n.includes(f) ? 2 : 9;
    return pool.map((n, i) => [i, score(n)]).filter(x => x[1] < 9).sort((a, b) => a[1] - b[1] || a[0] - b[0]).map(x => x[0]);
  }

  function evalExpr(expr, current) {
    const e = String(expr).replace(/[\s\u00A0]+/g, '');
    if (!e || /[^\d.,+\-]/.test(e)) return null;
    const relative = /^[+\-]/.test(e);
    const terms = e.split(/(?=[+\-])/).filter(Boolean);
    let sum = 0;
    for (const t of terms) {
      const sign = t[0] === '-' ? -1 : 1;
      const c = smartNumber(t.replace(/^[+\-]/, ''));
      if (!c || isNaN(parseFloat(c))) return null;
      sum += sign * parseFloat(c);
    }
    const value = Math.round(((relative ? current : 0) + sum) * 1000) / 1000;
    if (value < 0) return { error: 'النتيجة أقل من صفر' };
    return { value: Math.min(value, MAX_GROUP), relative, compound: terms.length > 1 || relative };
  }

  function parseEntry(raw, ctx) {
    const names = ctx.names, values = ctx.values, target = ctx.target || 0;
    const s = toLatin(raw).trim();
    if (!s) return { kind: 'skip' };
    const wrap = (index, v) => !v ? { kind: 'bad' } : v.error ? { kind: 'bad', msg: v.error } : { kind: 'set', index, ...v };
    let m = s.match(/^#?(\d{1,2})\s*[=:]\s*(.+)$/);
    if (m && +m[1] >= 1 && +m[1] <= names.length) {
      const i = +m[1] - 1;
      return wrap(i, evalExpr(m[2], values[i]));
    }
    m = s.match(/^(.*?\p{L}[^\d+\-]*?)\s*([+\-]?\d[\d.,+\-\s]*)?$/u);
    if (m && /\p{L}/u.test(m[1])) {
      const hits = findGroups(m[1], names, ctx.folded);
      if (!hits.length) return { kind: 'nomatch', query: m[1].trim() };
      if (!m[2]) return { kind: 'goto', index: hits[0], more: hits.length - 1 };
      return wrap(hits[0], evalExpr(m[2], values[hits[0]]));
    }
    return wrap(target, evalExpr(s, values[target]));
  }

  function splitBatch(raw) {
    return String(raw).split(/[\u061B;\n\r]+/).map(p => p.trim()).filter(Boolean);
  }

  function currentFiscalStart(now) {
    const d = now || new Date();
    const end = d.getMonth() + 1 >= 7 ? d.getFullYear() + 1 : d.getFullYear();
    return end - 1;
  }

  function expandTwoDigit(yy, currentStart) {
    const modern = 2000 + yy;
    return modern <= currentStart + 30 ? modern : 1900 + yy;
  }

  const fiscalLabel = start => `${start} \u2013 ${start + 1}`;

  function parseFiscalInput(raw, opts) {
    const o = opts || {};
    const phase = o.phase === 'commit' ? 'commit' : 'typing';
    const cur = currentFiscalStart(o.now);
    const suggest = String(cur);
    const s = toLatin(raw == null ? '' : raw).replace(SPACE_RE, '');
    const sepAt = s.search(/[-\u2010-\u2015/\\]/);
    let a, b = '', hasSep = sepAt !== -1;
    if (hasSep) {
      a = s.slice(0, sepAt).replace(/\D/g, '');
      const mb = s.slice(sepAt + 1).match(/\d+/);
      b = mb ? mb[0] : '';
    } else {
      a = s.replace(/\D/g, '');
      if (a.length === 8) { b = a.slice(4); a = a.slice(0, 4); hasSep = true; }
    }
    if (!a) return { state: 'empty', ghost: phase === 'typing' && !hasSep ? suggest : '' };
    if (a.length > 4) return { state: 'invalid', reason: 'format' };

    let start;
    if (a.length === 4) start = +a;
    else if (a.length === 2 && (phase === 'commit' || hasSep)) start = expandTwoDigit(+a, cur);
    else if (phase === 'typing') {
      const res = { state: 'partial', ghost: !hasSep && suggest.startsWith(a) ? suggest.slice(a.length) : '' };
      if (a.length === 2) res.preview = expandTwoDigit(+a, cur);
      return res;
    } else return { state: 'invalid', reason: 'incomplete' };

    if (start < YEAR_MIN_START || start > YEAR_MAX_START) return { state: 'invalid', reason: 'range', start };
    const end = start + 1;
    if (hasSep && b) {
      const es = String(end);
      if (b.length === 4) {
        if (b !== es) return { state: 'invalid', reason: 'end-mismatch', start, end };
      } else if (b.length === 2) {
        if (b !== es.slice(2) && !(phase === 'typing' && b === es.slice(0, 2))) return { state: 'invalid', reason: 'end-mismatch', start, end };
      } else if (phase === 'commit') return { state: 'invalid', reason: 'format', start, end };
    }
    const state = Math.abs(start - cur) > YEAR_SOFT_SPAN ? 'far' : 'valid';
    let settled = !hasSep;
    if (hasSep && b) {
      const es = String(end);
      settled = (b.length === 4 && b === es) || (b.length === 2 && b === es.slice(2));
    }
    return { state, start, end, settled };
  }

  let compact = null;
  function compactFormatter() {
    if (compact !== null) return compact;
    try {
      compact = new Intl.NumberFormat('ar-EG-u-nu-latn', { notation: 'compact', maximumFractionDigits: 2 });
    } catch (e) {
      compact = false;
    }
    return compact;
  }

  function magnitudeLabel(n) {
    if (!(n >= 1000000)) return '';
    const f = compactFormatter();
    if (f) {
      const out = f.format(n);
      if (/[\u0621-\u064A]/.test(out) && !/[\u0660-\u0669]/.test(out)) return out.replace(/\s/g, '\u00A0');
    }
    const trim = x => String(Math.round(x * 100) / 100);
    return n >= 1000000000 ? `${trim(n / 1000000000)}\u00A0مليار` : `${trim(n / 1000000)}\u00A0مليون`;
  }

  root.PharmaCore = Object.freeze({
    MAX_GROUP,
    YEAR_MIN_START,
    YEAR_MAX_START,
    YEAR_SOFT_SPAN,
    toLatin,
    normalizeDigits,
    sanitizeDecimal,
    smartNumber,
    groupLive,
    parseDecimal,
    foldArabic,
    wordScore,
    textScore,
    findGroups,
    evalExpr,
    parseEntry,
    splitBatch,
    currentFiscalStart,
    expandTwoDigit,
    fiscalLabel,
    parseFiscalInput,
    magnitudeLabel
  });
})(typeof window !== 'undefined' ? window : globalThis);
