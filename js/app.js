document.addEventListener('DOMContentLoaded', () => {
  const $ = id => document.getElementById(id);
  const { MAX_GROUP, YEAR_MIN_START, YEAR_MAX_START, currentFiscalStart, fiscalLabel, parseFiscalInput, toLatin, normalizeDigits, sanitizeDecimal, smartNumber, groupLive, parseDecimal, foldArabic, evalExpr, wordScore, textScore, findGroups } = window.PharmaCore;
  const escapeHTML = str => String(str).replace(/[&<>'"]/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[m]));
  const reducedMotion = () => window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const wait = ms => new Promise(r => setTimeout(r, ms));
  const store = {
    get(key, fallback) { try { const v = localStorage.getItem(key); return v === null ? fallback : JSON.parse(v); } catch (e) { return fallback; } },
    set(key, value) { try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) {} },
    remove(key) { try { localStorage.removeItem(key); } catch (e) {} }
  };

  const docEl = document.documentElement;
  const appShell = $('app-shell');
  const cardsContainer = $('cards-container');
  const totalEl = $('total');
  const totalCard = $('total-card');
  const totalMeta = $('total-meta');
  const totalProgressFill = $('total-progress-fill');
  const dockTotal = $('dock-total');
  const dockGroups = $('dock-groups');
  const dockGroupsCount = $('dock-groups-count');
  const ringFill = $('ring-fill');
  const saveState = $('save-state');
  const groupsCount = $('groups-count');
  const mastheadSub = $('masthead-sub');
  const yearSel = $('year');
  const monthSel = $('month');
  const regionSel = $('region');
  const pharmacySel = $('pharmacy');
  const selects = [yearSel, monthSel, regionSel, pharmacySel];
  const yearPrev = $('year-prev');
  const yearNext = $('year-next');
  const yearInput = $('year-input');
  const yearGhost = $('yt-ghost');
  const yearTail = $('year-tail');
  const yearEnd = $('yt-end');
  const yearType = $('year-type');
  const yearHint = $('year-hint');
  const monthChoices = $('month-choices');
  const regionChoices = $('region-choices');
  const pharmacyChoices = $('pharmacy-choices');
  const pharmacyHint = $('pharmacy-hint');
  const groupsEmpty = $('groups-empty');
  const filterBtns = Array.from(document.querySelectorAll('#group-filter .seg-btn'));
  const viewBtns = Array.from(document.querySelectorAll('#view-switch .seg-btn'));
  const submitBtn = $('submitButton');
  const submitInner = $('submitInner');
  const resetBtn = $('resetButton');
  const statusBar = $('status-bar');
  const statusDismiss = $('status-dismiss');
  const statusReopen = $('status-reopen');
  const chipClass = $('chip-class');
  const chipFields = $('chip-fields');
  const chipCards = $('chip-cards');
  const chipReady = $('chip-ready');
  const totalTicketsInput = $('totalTickets');
  const totalExternalInput = $('totalExternal');
  const optionalInputs = [totalTicketsInput, totalExternalInput];
  const toastArea = $('notification-area');
  const themeColorMeta = $('theme-color-meta');
  const themeBtns = Array.from(document.querySelectorAll('.theme-switch-btn'));
  const classBtns = Array.from(document.querySelectorAll('.class-btn'));

  const SCRIPT_URL = 'https://firecloud.rubberylock7.workers.dev/submit/pharmacy-sheet';
  const XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  const DATA_KEY = 'pharmacyAppData';
  const THEME_KEY = 'pharmacyAppTheme';
  const STATUS_KEY = 'pharmacyAppStatusHidden';
  const SENT_KEY = 'pharmacyAppSent';
  const VIEW_KEY = 'pharmacyAppView';
  const MAX_OPTIONAL = 1000000;
  const RING_LEN = 2 * Math.PI * 16;

  const cardNames = ['المسكنات', 'مضادات حيوية', 'كلي', 'نفسية و عصبية', 'قلب', 'سكر حقن', 'سكر فم', 'نسا', 'كبد', 'جهاز هضمي', 'جهاز تنفسي', 'امراض جلدية', 'عيون و رمد', 'انف و اذن', 'مضادات تقلصات', 'اورام', 'فيتامينات', 'مختلفة'];
  const HEADERS = ['date', 'year', 'class', 'month', 'region', 'pharmacy', 'totalPrescription', 'insuranceCoveredPrescription', ...cardNames];
  const CLASS_THEMES = { students: 'theme-students', workforce: 'theme-workforce', infants: 'theme-infants' };

  const pharmacyOptions = {
    students: {
      first: ['كفر الدوار طلاب', 'البيضا طلاب', 'النوبارية طلاب', 'ابو المطامير طلاب', 'رشيد طلاب', 'ادكو طلاب'],
      second: ['دمنهور طلاب', 'حوش عيسي طلاب', 'المحمودية طلاب', 'الرحمانية طلاب', 'ابو حمص طلاب', 'اورام طلاب'],
      third: ['ايتاي طلاب', 'كوم حمادة طلاب', 'بدر طلاب', 'شبراخيت طلاب', 'الدلنجات طلاب']
    },
    workforce: {
      first: ['كفر الدوار سكر', 'كفر الدوار الشاملة 1', 'كفر الدوار الشاملة 2', 'كفر الدوار مسائي', 'النوبارية', 'البيضا', 'ابو المطامير', 'رشيد', 'ادكو'],
      second: ['دمنهور الشاملة 1', 'دمنهور الشاملة 2', 'دمنهور مسائي', 'دمنهور سكر', 'دمنهور اورام', 'الشرطة', 'حوش عيسي', 'المحمودية مسائي', 'المحمودية', 'الرحمانية', 'ابو حمص'],
      third: ['ايتاي الشاملة 1', 'ايتاي الشاملة 2', 'ايتاي مسائي', 'كوم حمادة', 'كوم حمادة مسائي', 'بدر', 'شبراخيت مسائي', 'شبراخيت', 'الدلنجات مسائي', 'الدلنجات', 'وادي النطرون', 'قليشان']
    },
    infants: {
      first: ['كفر الدوار مواليد', 'البيضا مواليد', 'النوبارية مواليد', 'ابو المطامير مواليد', 'رشيد مواليد', 'ادكو مواليد'],
      second: ['دمنهور مواليد', 'حوش عيسي مواليد', 'المحمودية مواليد', 'الرحمانية مواليد', 'ابو حمص مواليد', 'اورام مواليد'],
      third: ['ايتاي مواليد', 'كوم حمادة مواليد', 'بدر مواليد', 'شبراخيت مواليد', 'الدلنجات مواليد']
    }
  };

  const ICONS = {
    success: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9.5"/><polyline points="8 12.5 11 15.5 16 9.5"/></svg>',
    error: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9.5"/><line x1="12" y1="7.5" x2="12" y2="13"/><line x1="12" y1="16.5" x2="12.01" y2="16.5"/></svg>',
    info: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9.5"/><line x1="12" y1="11" x2="12" y2="16.5"/><line x1="12" y1="7.5" x2="12.01" y2="7.5"/></svg>',
    close: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>',
    send: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>',
    check: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>',
    fail: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>'
  };

  const fmt3 = n => n.toLocaleString('en-US', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
  const fmtShort = n => n.toLocaleString('en-US', { maximumFractionDigits: 3 });

  function setKeepingCaret(inp, raw, shown) {
    const caret = inp.selectionStart ?? raw.length;
    const sig = raw.slice(0, caret).replace(/[^\d.\u0660-\u0669\u06F0-\u06F9\u066B]/g, '').length;
    if (inp.value !== shown) inp.value = shown;
    if (document.activeElement !== inp) return;
    let pos = 0, count = 0;
    while (pos < shown.length && count < sig) { if (/[\d.]/.test(shown[pos])) count++; pos++; }
    try { inp.setSelectionRange(pos, pos); } catch (e) {}
  }
  const displayDecimal = n => n > 0 ? String(Math.round(n * 1000) / 1000) : '';
  const groupedDecimal = n => n > 0 ? n.toLocaleString('en-US', { maximumFractionDigits: 3 }) : '';
  const fitChars = (el, text) => el.style.setProperty('--chars', Math.max(6, String(text).length));

  let cardValues = Array(cardNames.length).fill(0);
  let activeClass = null;
  let isSubmitting = false;
  let lastTotal = 0;
  let groupFilter = 'all';

  const darkMql = window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)') : null;
  const resolveTheme = mode => mode === 'system' ? (darkMql && darkMql.matches ? 'dark' : 'light') : mode;
  function applyTheme(mode, persist) {
    const resolved = resolveTheme(mode);
    docEl.setAttribute('data-theme', resolved);
    docEl.setAttribute('data-theme-pref', mode);
    themeColorMeta.setAttribute('content', resolved === 'dark' ? '#14151c' : '#f6efe1');
    themeBtns.forEach((btn, i) => {
      const on = btn.dataset.mode === mode;
      btn.classList.toggle('active', on);
      btn.setAttribute('aria-checked', String(on));
      btn.tabIndex = on ? 0 : -1;
      if (on) $('theme-switch').style.setProperty('--pos', i);
    });
    if (persist) { try { localStorage.setItem(THEME_KEY, mode); } catch (e) {} }
  }
  function rovingKeys(list, onPick) {
    list.forEach((btn, i) => btn.addEventListener('keydown', e => {
      const map = { ArrowLeft: 1, ArrowDown: 1, ArrowRight: -1, ArrowUp: -1 };
      let next = null;
      if (e.key in map) next = list[(i + map[e.key] + list.length) % list.length];
      else if (e.key === 'Home') next = list[0];
      else if (e.key === 'End') next = list[list.length - 1];
      if (next) { e.preventDefault(); next.focus(); onPick(next); }
    }));
  }
  themeBtns.forEach(btn => btn.addEventListener('click', () => applyTheme(btn.dataset.mode, true)));
  rovingKeys(themeBtns, btn => applyTheme(btn.dataset.mode, true));
  if (darkMql) {
    const onSys = () => { if (docEl.getAttribute('data-theme-pref') === 'system') applyTheme('system', false); };
    darkMql.addEventListener ? darkMql.addEventListener('change', onSys) : darkMql.addListener(onSys);
  }
  applyTheme(docEl.getAttribute('data-theme-pref') || 'system', false);

  function showToast(message, type = 'info', opts = {}) {
    const duration = opts.duration || 4500;
    const t = document.createElement('div');
    t.className = `toast ${type}`;
    t.setAttribute('role', type === 'error' ? 'alert' : 'status');
    t.innerHTML = `<span class="toast-icon">${ICONS[type] || ICONS.info}</span><span class="toast-msg">${escapeHTML(message)}</span>${opts.action ? `<button type="button" class="toast-action">${escapeHTML(opts.action.label)}</button>` : ''}<button type="button" class="toast-close" aria-label="إغلاق">${ICONS.close}</button><span class="toast-progress"></span>`;
    toastArea.appendChild(t);
    while (toastArea.children.length > 4) toastArea.firstElementChild.remove();
    let closed = false;
    const close = () => {
      if (closed) return;
      closed = true;
      const out = t.animate([{ opacity: 1, transform: 'none' }, { opacity: 0, transform: 'translateY(-10px) scale(0.96)' }], { duration: reducedMotion() ? 1 : 220, easing: 'ease-in', fill: 'forwards' });
      out.onfinish = () => t.remove();
    };
    t.animate([{ opacity: 0, transform: 'translateY(-16px) scale(0.94)' }, { opacity: 1, transform: 'none' }], { duration: reducedMotion() ? 1 : 360, easing: 'cubic-bezier(0.34,1.56,0.64,1)' });
    const bar = t.querySelector('.toast-progress').animate([{ transform: 'scaleX(1)' }, { transform: 'scaleX(0)' }], { duration, easing: 'linear', fill: 'forwards' });
    bar.onfinish = close;
    t.addEventListener('mouseenter', () => bar.pause());
    t.addEventListener('mouseleave', () => { if (!closed) bar.play(); });
    t.querySelector('.toast-close').addEventListener('click', close);
    if (opts.action) t.querySelector('.toast-action').addEventListener('click', () => { opts.action.onClick(); close(); });
    return close;
  }

  function confirmDialog({ title, desc, ok = 'تأكيد', danger = false }) {
    const dlg = $('confirm-dialog');
    const okBtn = $('confirm-ok-btn');
    const cancelBtn = $('confirm-cancel-btn');
    $('confirm-dialog-title').textContent = title;
    $('confirm-dialog-desc').textContent = desc;
    okBtn.textContent = ok;
    okBtn.classList.toggle('btn-danger', danger);
    okBtn.classList.toggle('btn-primary', !danger);
    return new Promise(resolve => {
      const done = val => { okBtn.onclick = cancelBtn.onclick = null; dlg.onclose = null; if (dlg.open) dlg.close(); resolve(val); };
      okBtn.onclick = () => done(true);
      cancelBtn.onclick = () => done(false);
      dlg.onclose = () => done(false);
      dlg.showModal();
      cancelBtn.focus();
    });
  }
  document.querySelectorAll('.comic-dialog').forEach(dlg => dlg.addEventListener('click', e => {
    if (e.target !== dlg) return;
    const r = dlg.getBoundingClientRect();
    if (e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom) dlg.close();
  }));

  cardNames.forEach((name, i) => {
    const card = document.createElement('div');
    card.className = 'card group-card reveal';
    card.style.setProperty('--i', i);
    card.dataset.index = i;
    card.dataset.search = `${foldArabic(name)} ${i + 1}`;
    const id = `drug-${i}`;
    card.innerHTML = `<div class="group-head"><span class="group-index">${String(i + 1).padStart(2, '0')}</span><label class="group-name" for="${id}">${escapeHTML(name)}</label><span class="group-share">—</span><button type="button" class="group-clear" data-index="${i}" aria-label="مسح ${escapeHTML(name)}" tabindex="-1">${ICONS.close}</button></div><div class="group-field"><input type="text" class="num-input drug-input" id="${id}" data-index="${i}" inputmode="decimal" autocomplete="off" dir="ltr" placeholder="0.000" enterkeyhint="next"></div><div class="group-meter" aria-hidden="true"><span class="group-meter-track"><span class="group-meter-fill"></span></span></div>`;
    cardsContainer.appendChild(card);
  });
  const groupCards = Array.from(cardsContainer.querySelectorAll('.group-card'));
  const drugInputs = Array.from(cardsContainer.querySelectorAll('.drug-input'));
  const visibleDrugInputs = () => drugInputs.filter(inp => !inp.closest('.group-card').hidden);
  const focusOrder = () => [...visibleDrugInputs(), totalTicketsInput, totalExternalInput];

  appShell.addEventListener('pointermove', e => {
    const el = e.target.closest('.reveal');
    if (!el) return;
    const r = el.getBoundingClientRect();
    el.style.setProperty('--mx', `${e.clientX - r.left}px`);
    el.style.setProperty('--my', `${e.clientY - r.top}px`);
  }, { passive: true });

  let saveTimer = null, flashTimer = null;
  function save() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      store.set(DATA_KEY, {
        cardValues,
        activeButton: activeClass ? classBtnFor(activeClass).id : null,
        year: yearSel.value, month: monthSel.value, region: regionSel.value, pharmacy: pharmacySel.value,
        totalTickets: totalTicketsInput.value, totalExternal: totalExternalInput.value
      });
      saveState.textContent = 'تم الحفظ';
      saveState.classList.add('is-flash');
      clearTimeout(flashTimer);
      flashTimer = setTimeout(() => { saveState.textContent = 'محفوظ تلقائيًا'; saveState.classList.remove('is-flash'); }, 1400);
    }, 250);
  }
  function snapshot() {
    return {
      cardValues: [...cardValues], activeClass,
      year: yearSel.value, month: monthSel.value, region: regionSel.value, pharmacy: pharmacySel.value,
      totalTickets: totalTicketsInput.value, totalExternal: totalExternalInput.value
    };
  }
  function restore(s) {
    applyClass(s.activeClass, { keepPharmacy: true });
    ensureYearOption(String(s.year || ''));
    yearSel.value = s.year || '';
    monthSel.value = s.month || '';
    regionSel.value = s.region || '';
    updatePharmacyDropdown(s.pharmacy || '');
    const cleanInt = v => { const d = normalizeDigits(v || '').replace(/\D/g, '').replace(/^0+(?=\d)/, ''); return d && +d > MAX_OPTIONAL ? String(MAX_OPTIONAL) : d; };
    totalTicketsInput.value = cleanInt(s.totalTickets);
    totalExternalInput.value = cleanInt(s.totalExternal);
    cardValues = cardNames.map((_, i) => parseDecimal(String((s.cardValues || [])[i] ?? '')));
    drugInputs.forEach((inp, i) => { inp.value = document.activeElement === inp ? displayDecimal(cardValues[i]) : groupedDecimal(cardValues[i]); });
    refresh();
    flushNow();
  }

  const classBtnFor = key => classBtns.find(b => b.dataset.class === key);
  function applyClass(key, { keepPharmacy = false } = {}) {
    const changed = key !== activeClass;
    activeClass = CLASS_THEMES[key] ? key : null;
    classBtns.forEach((btn, i) => {
      const on = btn.dataset.class === activeClass;
      btn.classList.toggle('btn-active', on);
      btn.setAttribute('aria-checked', String(on));
      btn.tabIndex = on || (!activeClass && i === 0) ? 0 : -1;
    });
    docEl.classList.remove(...Object.values(CLASS_THEMES));
    if (activeClass) docEl.classList.add(CLASS_THEMES[activeClass]);
    if (changed && !keepPharmacy) updatePharmacyDropdown('');
  }
  classBtns.forEach(btn => btn.addEventListener('click', () => A.setClass(btn.dataset.class)));
  rovingKeys(classBtns, btn => A.setClass(btn.dataset.class));

  const periodKey = () => activeClass && yearSel.value && monthSel.value ? `${activeClass}|${yearSel.value}|${monthSel.value}` : null;
  const loadSent = () => { const v = store.get(SENT_KEY, {}); return v && typeof v === 'object' && !Array.isArray(v) ? v : {}; };
  let sentCache = loadSent();
  const sentLog = () => sentCache;
  window.addEventListener('storage', e => {
    if (e.key !== SENT_KEY) return;
    sentCache = loadSent();
    updatePharmacyDropdown();
    schedule(D.PICKERS | D.STATUS, false);
  });
  function isSent(pharmacyValue) {
    const k = periodKey();
    return !!(k && regionSel.value && (sentLog()[k] || []).includes(`${regionSel.value}:${pharmacyValue}`));
  }
  function markSent() {
    const k = periodKey();
    if (!k) return;
    const log = sentLog();
    const entry = `${regionSel.value}:${pharmacySel.value}`;
    log[k] = Array.from(new Set([...(log[k] || []), entry]));
    store.set(SENT_KEY, log);
  }
  const currentList = () => (activeClass && regionSel.value && pharmacyOptions[activeClass][regionSel.value]) || [];
  const pharmacyName = () => currentList()[(parseInt(pharmacySel.value, 10) || 0) - 1] || '';

  function updatePharmacyDropdown(preferred = pharmacySel.value) {
    const list = currentList();
    const placeholder = !activeClass ? 'اختر الفئة أولًا' : !regionSel.value ? 'اختر المنطقة أولًا' : 'اختر الصيدلية';
    pharmacySel.innerHTML = '';
    pharmacySel.appendChild(new Option(placeholder, ''));
    list.forEach((name, i) => {
      const val = String(i + 1);
      const sent = isSent(val);
      const opt = new Option(sent ? `${name}  ✓ أُرسلت` : name, val);
      opt.dataset.name = name;
      if (sent) opt.dataset.sent = '1';
      pharmacySel.appendChild(opt);
    });
    pharmacySel.disabled = list.length === 0;
    pharmacySel.value = list.length && preferred && +preferred <= list.length ? String(preferred) : '';
  }


  const now = new Date();
  const curStart = currentFiscalStart(now);
  const currentFiscal = String(curStart + 1);
  const suggestStart = String(curStart);
  const suggestedMonth = String(((now.getMonth() + 11) % 12) + 1);
  const YEAR_MIN_END = YEAR_MIN_START + 1;
  const YEAR_MAX_END = YEAR_MAX_START + 1;
  const EASE = 'cubic-bezier(.22,.8,.24,1)';
  const validEnd = v => /^\d{4}$/.test(v) && +v >= YEAR_MIN_END && +v <= YEAR_MAX_END;
  const hasYear = validEnd;
  const validStart = s => /^\d{4}$/.test(s) && validEnd(String(+s + 1));
  const committedStart = () => validEnd(yearSel.value) ? +yearSel.value - 1 : null;
  const committedText = () => { const c = committedStart(); return c === null ? '' : String(c); };
  const parseYear = (raw, phase) => parseFiscalInput(raw, { now, phase });

  function ensureYearOption(v) {
    if (!validEnd(v) || Array.from(yearSel.options).some(o => o.value === v)) return;
    const after = Array.from(yearSel.options).find(o => o.value && +o.value > +v);
    yearSel.insertBefore(new Option(fiscalLabel(+v - 1), v), after || null);
  }
  const commitYearStart = start => validStart(String(start)) && A.setYearStart(+start);

  const fieldYear = $('field-year');
  const yearBox = $('yt-box');
  const yearQuick = $('year-quick');
  const yearMsg = $('year-msg');
  let yearArmed = null;
  let yearNotice = null;
  let yearNoticeTimer = 0;
  let suppressMouseUp = false;
  let quickSig = '';

  const yearFocused = () => document.activeElement === yearInput;
  const cleanYearText = raw => toLatin(raw).replace(/[\u2010-\u2015\\/]/g, '-').replace(/[^\d-]/g, '').replace(/-+/g, '-').replace(/^-/, '').slice(0, 9);
  function sizeYear() { yearInput.style.width = `${Math.max(yearInput.value.length, 1) + 0.12}ch`; }
  function caretToEnd() { try { yearInput.setSelectionRange(yearInput.value.length, yearInput.value.length); } catch (e) {} }

  function setTail(end, state) {
    if (end && yearEnd.textContent !== String(end)) {
      const wasOn = yearTail.dataset.state !== 'none';
      yearEnd.textContent = String(end);
      if (wasOn && !reducedMotion()) yearEnd.animate([{ opacity: 0, transform: 'translateY(6px)' }, { opacity: 1, transform: 'none' }], { duration: 220, easing: EASE });
    }
    yearTail.dataset.state = end ? state : 'none';
  }

  const YEAR_REASON = {
    range: () => `خارج النطاق (${YEAR_MIN_START} – ${YEAR_MAX_START})`,
    'end-mismatch': r => `سنة النهاية يجب أن تكون ${r.end}`,
    incomplete: () => 'أكمل السنة بأربعة أرقام',
    format: () => 'صيغة غير مفهومة، اكتب مثل 2024 أو 2024/25'
  };

  function setYearMsg(html, cls) {
    const next = `${cls || ''}|${html}`;
    if (yearMsg.dataset.sig === next) return;
    yearMsg.dataset.sig = next;
    yearMsg.className = `year-msg${cls ? ` ${cls}` : ''}`;
    yearMsg.innerHTML = html;
  }

  function noticeYear(html, cls, ms) {
    clearTimeout(yearNoticeTimer);
    yearNotice = { html, cls };
    yearNoticeTimer = setTimeout(() => { yearNotice = null; renderYearTyping(); }, ms);
    renderYearTyping();
  }
  function clearNotice() {
    clearTimeout(yearNoticeTimer);
    yearNotice = null;
  }

  function renderYearTyping() {
    const focused = yearFocused();
    const committed = committedStart();
    let r;
    if (!focused) r = committed === null ? { state: 'empty' } : { state: 'valid', start: committed, end: committed + 1 };
    else {
      r = parseYear(yearInput.value, 'typing');
      if (r.state === 'far' && r.start === committed) r = { ...r, state: 'valid' };
    }
    yearGhost.textContent = focused && r.ghost ? r.ghost : '';
    let tailEnd = 0, tailState = 'none';
    if (r.state === 'valid') { tailEnd = r.end; tailState = 'on'; }
    else if (r.state === 'far') { tailEnd = r.end; tailState = 'far'; }
    else if (focused && r.ghost) { tailEnd = curStart + 1; tailState = 'ghost'; }
    setTail(tailEnd, tailState);
    yearType.classList.toggle('is-invalid', r.state === 'invalid');
    yearType.classList.toggle('is-far', r.state === 'far');

    if (yearNotice) { setYearMsg(yearNotice.html, yearNotice.cls); return; }
    if (r.state === 'invalid') setYearMsg(escapeHTML(YEAR_REASON[r.reason](r)), 'is-bad');
    else if (r.state === 'far') {
      const label = fiscalLabel(r.start);
      setYearMsg(`سنة بعيدة؟ اضغط <kbd>Enter</kbd> مرة أخرى للتأكيد · <button type="button" class="link-btn" data-year-confirm="${r.start}">نعم، <span dir="ltr">${label}</span></button>`, 'is-warn');
    } else if (r.state === 'partial' && r.preview) setYearMsg(`سيصبح <span dir="ltr">${fiscalLabel(r.preview)}</span>`, 'is-dim');
    else setYearMsg('', '');
  }

  function renderYearHint() {
    const sel = yearSel.value;
    let html;
    if (yearFocused()) html = 'اكتب سنة البداية · <kbd>↑</kbd> <kbd>↓</kbd> للتغيير';
    else if (!validEnd(sel)) html = 'اكتب سنة البداية مثل <b dir="ltr">2024</b>';
    else if (sel === currentFiscal) html = 'السنة الحالية';
    else html = '';
    if (yearHint.dataset.html === html) return;
    yearHint.dataset.html = html;
    yearHint.innerHTML = html;
  }

  function renderYearQuick() {
    const committed = committedStart();
    const starts = [curStart, curStart - 1];
    if (committed !== null && !starts.includes(committed)) starts.push(committed);
    const sig = starts.join(',');
    if (sig !== quickSig) {
      quickSig = sig;
      yearQuick.innerHTML = starts.map(st => `<button type="button" class="choice year-chip" role="radio" data-start="${st}" aria-checked="false" tabindex="-1"><span class="choice-text" dir="ltr">${fiscalLabel(st)}</span>${st === curStart ? '<span class="chip-note">الحالية</span>' : ''}</button>`).join('');
    }
    const btns = Array.from(yearQuick.querySelectorAll('.choice'));
    btns.forEach(b => {
      const on = committed !== null && +b.dataset.start === committed;
      b.classList.toggle('is-on', on);
      b.setAttribute('aria-checked', String(on));
    });
    const tabbable = btns.find(b => b.classList.contains('is-on')) || btns[0];
    btns.forEach(b => { b.tabIndex = b === tabbable ? 0 : -1; });
  }

  function renderYear() {
    const committed = committedStart();
    if (!yearFocused()) {
      const t = committedText();
      if (yearInput.value !== t) yearInput.value = t;
      sizeYear();
    }
    renderYearTyping();
    renderYearQuick();
    yearPrev.disabled = committed !== null && committed <= YEAR_MIN_START;
    yearNext.disabled = committed !== null && committed >= YEAR_MAX_START;
    fieldYear.classList.toggle('is-filled', committed !== null);
    renderYearHint();
  }

  function rollYear(oldText, dir) {
    if (reducedMotion() || !dir || !oldText) return;
    const box = yearBox.getBoundingClientRect();
    const inp = yearInput.getBoundingClientRect();
    const ghost = document.createElement('span');
    ghost.className = 'yt-roll';
    ghost.setAttribute('aria-hidden', 'true');
    ghost.textContent = oldText;
    ghost.style.insetInlineStart = '0';
    ghost.style.left = `${inp.left - box.left}px`;
    ghost.style.top = `${inp.top - box.top}px`;
    yearBox.appendChild(ghost);
    const dist = dir > 0 ? '-55%' : '55%';
    const out = ghost.animate([{ opacity: 1, transform: 'none' }, { opacity: 0, transform: `translateY(${dist})` }], { duration: 180, easing: EASE, fill: 'forwards' });
    out.onfinish = () => ghost.remove();
    yearInput.animate([{ opacity: 0, transform: `translateY(${dir > 0 ? '55%' : '-55%'})` }, { opacity: 1, transform: 'none' }], { duration: 180, easing: EASE });
  }

  function shakeYear() {
    if (reducedMotion()) return;
    yearType.animate([{ transform: 'translateX(0)' }, { transform: 'translateX(-4px)' }, { transform: 'translateX(4px)' }, { transform: 'translateX(0)' }], { duration: 240 });
  }

  function baseStart() {
    const r = parseYear(yearInput.value, 'typing');
    if (r.start !== undefined && r.state !== 'invalid') return r.start;
    const c = committedStart();
    return c === null ? curStart : c;
  }

  function stepYearTo(next) {
    const clamped = Math.max(YEAR_MIN_START, Math.min(YEAR_MAX_START, next));
    const before = yearInput.value;
    const dir = clamped - baseStart();
    yearArmed = null;
    clearNotice();
    A.setYearStart(clamped);
    yearInput.value = String(clamped);
    sizeYear();
    renderYear();
    if (yearFocused()) { yearInput.select(); } else caretToEnd();
    rollYear(before, dir);
  }
  const stepYear = delta => stepYearTo(baseStart() + delta);

  function revertYear() {
    yearArmed = null;
    clearNotice();
    yearInput.value = committedText();
    sizeYear();
    renderYear();
  }

  function finalizeYear({ enter = false } = {}) {
    const raw = yearInput.value;
    const r = parseYear(raw, 'commit');
    const committed = committedStart();
    if (r.state === 'empty') { revertYear(); return committed !== null; }
    if (r.state === 'valid' || (r.state === 'far' && (r.start === committed || yearArmed === r.start))) {
      A.setYearStart(r.start);
      yearArmed = null;
      clearNotice();
      yearInput.value = String(r.start);
      sizeYear();
      renderYear();
      return true;
    }
    if (r.state === 'far') {
      if (enter) { yearArmed = r.start; renderYearTyping(); shakeYear(); return false; }
      revertYear();
      return false;
    }
    const msg = YEAR_REASON[r.reason](r);
    revertYear();
    noticeYear(escapeHTML(msg), 'is-bad', 2500);
    shakeYear();
    return false;
  }

  function onYearInput() {
    yearArmed = null;
    clearNotice();
    const raw = yearInput.value;
    const clean = cleanYearText(raw);
    if (clean !== raw) {
      const caret = yearInput.selectionStart ?? raw.length;
      const kept = cleanYearText(raw.slice(0, caret)).length;
      yearInput.value = clean;
      try { yearInput.setSelectionRange(kept, kept); } catch (e) {}
    }
    sizeYear();
    const r = parseYear(clean, 'typing');
    if (r.state === 'valid' && r.settled) A.setYearStart(r.start);
    renderYearTyping();
    renderYearHint();
  }

  yearInput.addEventListener('beforeinput', e => {
    if (e.inputType === 'insertText' && e.data && /[^\d\u0660-\u0669\u06F0-\u06F9\-\u2010-\u2015/\\]/.test(e.data)) e.preventDefault();
  });
  yearInput.addEventListener('input', onYearInput);
  yearInput.addEventListener('paste', e => {
    e.preventDefault();
    const clean = cleanYearText((e.clipboardData || window.clipboardData).getData('text'));
    if (!clean) return;
    yearInput.setRangeText(clean, yearInput.selectionStart, yearInput.selectionEnd, 'end');
    yearInput.dispatchEvent(new Event('input', { bubbles: true }));
  });
  yearInput.addEventListener('keydown', e => {
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    const k = e.key;
    if (k === 'Enter') { e.preventDefault(); if (finalizeYear({ enter: true })) advanceFrom(yearInput); }
    else if (k === 'ArrowUp') { e.preventDefault(); stepYear(e.shiftKey ? 10 : 1); }
    else if (k === 'ArrowDown') { e.preventDefault(); stepYear(e.shiftKey ? -10 : -1); }
    else if (k === 'PageUp') { e.preventDefault(); stepYear(10); }
    else if (k === 'PageDown') { e.preventDefault(); stepYear(-10); }
    else if (k === 'Home') { e.preventDefault(); stepYearTo(curStart); }
    else if (k === 'ArrowRight' && yearGhost.textContent && yearInput.selectionStart === yearInput.value.length) {
      e.preventDefault();
      yearInput.value = suggestStart;
      onYearInput();
      caretToEnd();
    } else if (k === 'Escape' && (cleanYearText(yearInput.value) !== committedText() || yearArmed !== null || yearNotice)) {
      e.preventDefault();
      e.stopPropagation();
      revertYear();
      yearInput.select();
    }
  });
  yearInput.addEventListener('focus', () => {
    fieldYear.classList.add('is-typing');
    yearInput.select();
    suppressMouseUp = true;
    setTimeout(() => { suppressMouseUp = false; }, 250);
    renderYearTyping();
    renderYearHint();
  });
  yearInput.addEventListener('mouseup', e => { if (suppressMouseUp) { e.preventDefault(); suppressMouseUp = false; } });
  yearInput.addEventListener('blur', () => {
    fieldYear.classList.remove('is-typing');
    finalizeYear();
    renderYearHint();
  });
  yearBox.addEventListener('pointerdown', e => { if (e.target !== yearInput) { e.preventDefault(); yearInput.focus(); } });
  yearMsg.addEventListener('pointerdown', e => { if (e.target.closest('button')) e.preventDefault(); });
  yearMsg.addEventListener('click', e => {
    const b = e.target.closest('[data-year-confirm]');
    if (!b) return;
    const start = +b.dataset.yearConfirm;
    yearArmed = start;
    if (yearFocused()) { finalizeYear({ enter: true }); } else { A.setYearStart(start); renderYear(); }
  });
  yearPrev.addEventListener('click', () => stepYear(-1));
  yearNext.addEventListener('click', () => stepYear(1));

  yearQuick.addEventListener('click', e => {
    const b = e.target.closest('.choice');
    if (!b) return;
    clearNotice();
    yearArmed = null;
    A.setYearStart(+b.dataset.start);
    const again = yearQuick.querySelector(`.choice[data-start="${b.dataset.start}"]`);
    if (again && document.activeElement !== again) again.focus({ preventScroll: true });
  });
  yearQuick.addEventListener('keydown', e => {
    const b = e.target.closest('.choice');
    if (!b) return;
    const btns = Array.from(yearQuick.querySelectorAll('.choice'));
    const i = btns.indexOf(b);
    const map = { ArrowLeft: 1, ArrowDown: 1, ArrowRight: -1, ArrowUp: -1 };
    if (e.key === 'Enter' && !e.ctrlKey && !e.metaKey) {
      e.preventDefault();
      A.setYearStart(+b.dataset.start);
      advanceFrom(yearInput);
      return;
    }
    let next = null;
    if (e.key in map) next = btns[(i + map[e.key] + btns.length) % btns.length];
    else if (e.key === 'Home') next = btns[0];
    else if (e.key === 'End') next = btns[btns.length - 1];
    if (!next) return;
    e.preventDefault();
    A.setYearStart(+next.dataset.start);
    const again = yearQuick.querySelector(`.choice[data-start="${next.dataset.start}"]`);
    if (again) again.focus();
  });

  const choiceSig = new WeakMap();
  function renderChoices(container, sel, { empty = '', decorate } = {}) {
    const opts = Array.from(sel.options).filter(o => o.value);
    const sig = opts.map(o => `${o.value}:${o.dataset.name || o.text}:${o.dataset.sent || ''}`).join('|') + `#${sel.disabled}`;
    if (choiceSig.get(container) !== sig) {
      choiceSig.set(container, sig);
      if (!opts.length || sel.disabled) {
        container.innerHTML = `<p class="choice-empty">${escapeHTML(empty || sel.options[0]?.text || '')}</p>`;
      } else {
        container.innerHTML = opts.map(o => {
          const label = o.dataset.name || o.text;
          const extra = decorate ? decorate(o) : '';
          return `<button type="button" class="choice${o.dataset.sent ? ' is-sent' : ''}${extra}" role="radio" data-value="${escapeHTML(o.value)}" aria-checked="false" tabindex="-1"><span class="choice-text">${escapeHTML(label)}</span>${o.dataset.sent ? `<span class="choice-sent" title="أُرسلت">${ICONS.check}</span>` : ''}</button>`;
        }).join('');
      }
    }
    const btns = Array.from(container.querySelectorAll('.choice'));
    btns.forEach(b => {
      const on = b.dataset.value === sel.value;
      b.classList.toggle('is-on', on);
      b.setAttribute('aria-checked', String(on));
    });
    const tabbable = btns.find(b => b.classList.contains('is-on')) || btns[0];
    btns.forEach(b => { b.tabIndex = b === tabbable ? 0 : -1; });
    container.closest('.field').classList.toggle('is-filled', !!sel.value && !sel.disabled);
  }
  function bindChoices(container, sel, setter) {
    container.addEventListener('click', e => {
      const b = e.target.closest('.choice');
      if (!b) return;
      setter(b.dataset.value);
      const again = container.querySelector(`.choice[data-value="${CSS.escape(b.dataset.value)}"]`);
      if (again && document.activeElement !== again) again.focus({ preventScroll: true });
    });
    let typed = '', typedTimer = null;
    container.addEventListener('keydown', e => {
      const b = e.target.closest('.choice');
      if (!b) return;
      const btns = Array.from(container.querySelectorAll('.choice'));
      const i = btns.indexOf(b);
      const map = { ArrowLeft: 1, ArrowDown: 1, ArrowRight: -1, ArrowUp: -1 };
      let next = null;
      if (e.key === 'Enter' && !e.ctrlKey && !e.metaKey) {
        e.preventDefault();
        if (sel.value !== b.dataset.value) setter(b.dataset.value);
        advanceFrom(container);
        return;
      }
      if (e.key in map) next = btns[(i + map[e.key] + btns.length) % btns.length];
      else if (e.key === 'Home') next = btns[0];
      else if (e.key === 'End') next = btns[btns.length - 1];
      else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey && e.key !== ' ') {
        clearTimeout(typedTimer);
        typed += foldArabic(normalizeDigits(e.key));
        typedTimer = setTimeout(() => { typed = ''; }, 900);
        if (/^\d+$/.test(typed)) next = btns[+typed - 1] || btns[+typed.slice(-1) - 1] || null;
        else {
          const label = x => foldArabic(x.textContent);
          next = btns.find(x => label(x).startsWith(typed)) || btns.find(x => label(x).split(' ').some(w => w.startsWith(typed))) || btns.find(x => label(x).includes(typed)) || null;
        }
        if (!next) return;
      }
      if (!next) return;
      e.preventDefault();
      setter(next.dataset.value);
      const again = container.querySelector(`.choice[data-value="${CSS.escape(next.dataset.value)}"]`);
      if (again) again.focus();
    });
  }
  bindChoices(monthChoices, monthSel, v => A.setMonth(v));
  bindChoices(regionChoices, regionSel, v => A.setRegion(v));
  bindChoices(pharmacyChoices, pharmacySel, v => A.setPharmacy(activeClass, regionSel.value, +v - 1));

  function advanceFrom(container) {
    const flow = [$('class-group'), yearInput, monthChoices, regionChoices, pharmacyChoices];
    const filled = [() => !!activeClass, () => !!yearSel.value, () => !!monthSel.value, () => !!regionSel.value, () => !!pharmacySel.value];
    const at = flow.indexOf(container);
    for (let k = at + 1; k < flow.length; k++) {
      if (filled[k]() && k !== flow.length - 1) continue;
      const c = flow[k];
      const btn = c === yearInput ? c : (c.querySelector('[tabindex="0"]') || c.querySelector('button'));
      if (btn) {
        btn.focus({ preventScroll: true });
        btn.closest('.field, .class-grid')?.scrollIntoView({ behavior: reducedMotion() ? 'auto' : 'smooth', block: 'nearest' });
        return;
      }
    }
    goToEntry();
  }
  function goToEntry() {
    const empty = cardValues.findIndex(v => !(v > 0));
    if (empty === -1) {
      $('step-totals').scrollIntoView({ behavior: reducedMotion() ? 'auto' : 'smooth', block: 'center' });
      setTimeout(() => totalTicketsInput.focus({ preventScroll: true }), reducedMotion() ? 0 : 320);
      return;
    }
    setTarget(empty);
    jumpToGroup(empty);
  }
  classBtns.forEach(btn => btn.addEventListener('keydown', e => {
    if (e.key === 'Enter' && !e.ctrlKey && !e.metaKey) {
      e.preventDefault();
      A.setClass(btn.dataset.class);
      advanceFrom($('class-group'));
    } else if (/^[123\u0661-\u0663]$/.test(e.key)) {
      e.preventDefault();
      const b = classBtns[+normalizeDigits(e.key) - 1];
      b.focus(); A.setClass(b.dataset.class);
    }
  }));

  function renderPickers() {
    renderYear();
    renderChoices(monthChoices, monthSel, { decorate: o => o.value === suggestedMonth ? ' is-suggested' : '' });
    renderChoices(regionChoices, regionSel);
    renderChoices(pharmacyChoices, pharmacySel, { empty: !activeClass ? 'اختر الفئة أولًا لعرض الصيدليات' : 'اختر المنطقة لعرض الصيدليات' });
    const list = currentList();
    const sentCount = list.filter((_, i) => isSent(String(i + 1))).length;
    pharmacyHint.textContent = list.length && periodKey() ? `${sentCount} من ${list.length} أُرسلت هذا الشهر` : list.length ? `${list.length} صيدلية` : '';
    pharmacyHint.classList.toggle('is-complete', list.length > 0 && sentCount === list.length);
  }

  function setCardValue(i, raw, { format = false } = {}) {
    const inp = drugInputs[i];
    const clean = sanitizeDecimal(raw);
    let n = parseFloat(clean);
    if (isNaN(n) || n <= 0) n = 0;
    if (n > MAX_GROUP) n = MAX_GROUP;
    cardValues[i] = Math.round(n * 1000) / 1000;
    const shown = format ? (document.activeElement === inp ? groupLive(displayDecimal(cardValues[i])) : groupedDecimal(cardValues[i])) : groupLive(n === MAX_GROUP ? String(MAX_GROUP) : clean);
    return shown;
  }
  function writeCard(i, raw, opts) {
    const shown = setCardValue(i, raw, opts);
    if (drugInputs[i].value !== shown) drugInputs[i].value = shown;
  }

  const focusStart = {};
  cardsContainer.addEventListener('input', e => {
    if (!e.target.classList.contains('drug-input')) return;
    const inp = e.target;
    let raw = inp.value;
    if (e.inputType === 'insertText' && e.data === ',' && !raw.includes('.')) {
      const c = inp.selectionStart;
      raw = raw.slice(0, c - 1) + '.' + raw.slice(c);
      inp.value = raw;
      try { inp.setSelectionRange(c, c); } catch (err) {}
    }
    setKeepingCaret(inp, raw, setCardValue(+inp.dataset.index, raw));
    schedule(D.TOTALS | D.GROUPS | D.STATUS);
  });
  cardsContainer.addEventListener('focusout', e => {
    if (!e.target.classList.contains('drug-input')) return;
    const idx = +e.target.dataset.index;
    writeCard(idx, e.target.value, { format: true });
    if (focusStart[idx] !== undefined) pushHistory(idx, focusStart[idx], cardValues[idx]);
    delete focusStart[idx];
    schedule(D.TOTALS | D.GROUPS | D.STATUS);
    setTimeout(applyGroupFilter, 0);
  });
  cardsContainer.addEventListener('focusin', e => {
    if (!e.target.classList.contains('drug-input')) return;
    const inp = e.target;
    const idx = +inp.dataset.index;
    focusStart[idx] = cardValues[idx];
    const raw = groupLive(displayDecimal(cardValues[idx]));
    if (inp.value !== raw) inp.value = raw;
    fitChars(inp, inp.value);
    inp.select();
    setTarget(idx);
  });
  cardsContainer.addEventListener('click', e => {
    const btn = e.target.closest('.group-clear');
    if (btn) {
      const i = +btn.dataset.index;
      const prev = A.setGroup(i, 0);
      drugInputs[i].focus();
      showToast(`تم مسح «${cardNames[i]}»`, 'info', { duration: 4000, action: { label: 'تراجع', onClick: () => A.restoreGroup(i, prev) } });
      return;
    }
    const card = e.target.closest('.group-card');
    if (card && !e.target.closest('input, label, button')) drugInputs[+card.dataset.index].focus();
  });
  cardsContainer.addEventListener('paste', e => {
    if (!e.target.classList.contains('drug-input')) return;
    const text = (e.clipboardData || window.clipboardData).getData('text');
    if (!/[\t\r\n]/.test(text.trim())) return;
    e.preventDefault();
    let cells = text.replace(/\r/g, '').replace(/\n+$/, '').split('\n');
    if (cells.length === 1) cells = cells[0].split('\t');
    else cells = cells.map(row => row.split('\t')[0]);
    const start = +e.target.dataset.index;
    const changes = [];
    cells.forEach((cell, k) => {
      const i = start + k;
      if (i >= cardNames.length) return;
      changes.push({ i, value: parseDecimal(smartNumber(cell)) });
    });
    const filled = changes.length;
    A.setGroups(changes);
    drugInputs[Math.min(start + filled, cardNames.length) - 1].focus();
    showToast(`تم لصق ${filled} قيمة بدءًا من «${cardNames[start]}»`, 'success', { duration: 3000 });
  });

  function matchesGroup(card) {
    const v = cardValues[+card.dataset.index];
    if (groupFilter === 'filled' && !(v > 0)) return false;
    if (groupFilter === 'empty' && v > 0) return false;
    return true;
  }
  function applyGroupFilter() {
    const active = document.activeElement;
    let shown = 0;
    groupCards.forEach(card => {
      const keep = card.contains(active) && active.classList.contains('drug-input');
      const show = matchesGroup(card) || keep;
      if (card.hidden === show) card.hidden = !show;
      if (show) shown++;
    });
    groupsEmpty.hidden = shown > 0;
    cardsContainer.classList.toggle('is-filtered', groupFilter !== 'all');
  }
  function setGroupFilter(f) {
    groupFilter = f;
    filterBtns.forEach(b => {
      const on = b.dataset.filter === f;
      b.classList.toggle('is-on', on);
      b.setAttribute('aria-checked', String(on));
      b.tabIndex = on ? 0 : -1;
    });
    applyGroupFilter();
  }
  filterBtns.forEach(b => b.addEventListener('click', () => setGroupFilter(b.dataset.filter)));
  rovingKeys(filterBtns, b => setGroupFilter(b.dataset.filter));

  function setView(v, persist) {
    const view = v === 'grid' ? 'grid' : 'list';
    cardsContainer.classList.toggle('is-list', view === 'list');
    cardsContainer.classList.toggle('is-grid', view === 'grid');
    viewBtns.forEach(b => {
      const on = b.dataset.view === view;
      b.classList.toggle('is-on', on);
      b.setAttribute('aria-checked', String(on));
      b.tabIndex = on ? 0 : -1;
    });
    if (persist) store.set(VIEW_KEY, view);
  }
  viewBtns.forEach(b => b.addEventListener('click', () => setView(b.dataset.view, true)));
  rovingKeys(viewBtns, b => setView(b.dataset.view, true));
  setView(store.get(VIEW_KEY, 'list'), false);

  let target = 0;
  const history = [];
  function pushHistoryBatch(changes) {
    const list = changes.filter(c => c.prev !== c.next);
    if (!list.length) return;
    history.push(list);
    if (history.length > 100) history.shift();
  }
  const pushHistory = (i, prev, next) => pushHistoryBatch([{ i, prev, next }]);
  const foldedNames = cardNames.map(foldArabic);
  const parseEntry = raw => window.PharmaCore.parseEntry(raw, { names: cardNames, values: cardValues, target, folded: foldedNames });
  const nextAfter = i => (i + 1 < cardNames.length ? i + 1 : (cardValues.findIndex(v => !(v > 0)) + 1 || cardNames.length) - 1);
  const groupNum = n => n.toLocaleString('en-US', { maximumFractionDigits: 3 });

  function setTarget(i, { scroll = false } = {}) {
    target = Math.max(0, Math.min(cardNames.length - 1, i));
    groupCards.forEach((c, k) => c.classList.toggle('is-target', k === target));
    if (scroll) {
      const card = groupCards[target];
      if (card.hidden) { setGroupFilter('all'); }
      card.scrollIntoView({ behavior: reducedMotion() ? 'auto' : 'smooth', block: 'nearest' });
    }
  }




  function flashCard(i) {
    const card = groupCards[i];
    card.classList.remove('is-landed'); void card.offsetWidth; card.classList.add('is-landed');
  }
  function undo() { return A.undo(); }
  function applyGroup(i, value) {
    const prev = A.setGroup(i, value);
    flashCard(i);
    return prev;
  }

  $('groups-empty-reset').addEventListener('click', () => setGroupFilter('all'));

  function jumpToGroup(i) {
    const card = groupCards[i];
    if (card.hidden) setGroupFilter('all');
    card.scrollIntoView({ behavior: reducedMotion() ? 'auto' : 'smooth', block: 'center' });
    card.classList.remove('needs-attention'); void card.offsetWidth; card.classList.add('needs-attention');
    setTimeout(() => drugInputs[i].focus({ preventScroll: true }), reducedMotion() ? 0 : 320);
  }
  dockGroups.addEventListener('click', goToEntry);

  optionalInputs.forEach(inp => {
    inp.addEventListener('input', () => {
      let v = normalizeDigits(inp.value).replace(/\D/g, '').replace(/^0+(?=\d)/, '');
      if (v !== '' && +v > MAX_OPTIONAL) v = String(MAX_OPTIONAL);
      if (inp.value !== v) inp.value = v;
      schedule(D.GROUPS | D.STATUS);
    });
    inp.addEventListener('focus', () => inp.select());
  });

  const isTyping = el => el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.tagName === 'SELECT' || el.isContentEditable);
  document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); submitBtn.click(); return; }
    const modal = document.querySelector('dialog[open]');
    if (e.key === '/' && !isTyping(e.target) && !modal && !e.ctrlKey && !e.metaKey && !e.altKey) {
      e.preventDefault();
      openPal();
      return;
    }
    if ((e.target === document.body || e.target === document.documentElement) && /^[\d\u0660-\u0669\u06F0-\u06F9]$/.test(e.key) && !e.ctrlKey && !e.metaKey && !e.altKey && !modal) {
      e.preventDefault();
      openPal(e.key);
      return;
    }
    if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && e.key.toLowerCase() === 'z' && !isTyping(e.target) && !modal) {
      e.preventDefault();
      A.undo();
      return;
    }
    const order = focusOrder();
    const idx = order.indexOf(e.target);
    if (idx === -1) return;
    let dir = 0;
    if (e.key === 'Enter') dir = e.shiftKey ? -1 : 1;
    else if (e.key === 'ArrowDown') dir = 1;
    else if (e.key === 'ArrowUp') dir = -1;
    else if (e.key === 'Escape' && e.target.classList.contains('drug-input')) { e.preventDefault(); e.target.blur(); return; }
    if (!dir) return;
    e.preventDefault();
    const next = order[idx + dir];
    if (next) next.focus();
    else if (dir > 0) submitBtn.focus();
    else if (e.key === 'ArrowUp' && e.target === order[0]) firstChoice(pharmacyChoices).focus();
  });

  function validOptional(inp) { const v = inp.value.trim(); return v === '' || (/^\d+$/.test(v) && +v <= MAX_OPTIONAL); }
  const firstChoice = c => c.querySelector('.choice[tabindex="0"]') || c.querySelector('.choice') || c;
  function missingList() {
    const m = [];
    if (!activeClass) m.push({ label: 'الفئة', el: classBtns[0], step: 'step-class' });
    if (!yearSel.value) m.push({ label: 'السنة', el: yearInput, step: 'step-fields' });
    if (!monthSel.value) m.push({ label: 'الشهر', el: firstChoice(monthChoices), step: 'step-fields' });
    if (!regionSel.value) m.push({ label: 'المنطقة', el: firstChoice(regionChoices), step: 'step-fields' });
    if (!pharmacySel.value) m.push({ label: 'الصيدلية', el: firstChoice(pharmacyChoices), step: 'step-fields' });
    if (!cardValues.some(v => v > 0)) m.push({ label: 'مجموعة دوائية واحدة على الأقل', el: drugInputs[0], step: 'step-groups' });
    optionalInputs.forEach(inp => { if (!validOptional(inp)) m.push({ label: 'قيمة إجمالي صحيحة', el: inp, step: 'step-totals' }); });
    return m;
  }
  const isFormEmpty = () => !activeClass && selects.every(s => !s.value || (s === yearSel && s.value === currentFiscal)) && cardValues.every(v => !v) && optionalInputs.every(i => !i.value);

  function animateTotal(to) {
    const from = lastTotal;
    lastTotal = to;
    dockTotal.textContent = fmt3(to);
    fitChars(dockTotal, fmt3(to));
    fitChars(totalEl, fmt3(to));
    if (Math.abs(to - from) < 0.0005) { totalEl.textContent = fmt3(to); return; }
    if (reducedMotion()) { totalEl.textContent = fmt3(to); return; }
    totalCard.classList.remove('is-bumped'); void totalCard.offsetWidth; totalCard.classList.add('is-bumped');
    const start = performance.now(), dur = 480;
    const step = t => {
      const p = Math.min(1, (t - start) / dur);
      const eased = 1 - Math.pow(1 - p, 3);
      if (lastTotal !== to) return;
      totalEl.textContent = fmt3(from + (to - from) * eased);
      if (p < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  }

  function setChip(chip, text, cls) { chip.textContent = text; chip.className = `status-chip ${cls}`; }

  const D = { TOTALS: 1, GROUPS: 2, PICKERS: 4, STATUS: 8, ALL: 15 };
  let dirty = 0, raf = 0;
  function schedule(flags = D.ALL, persist = true) {
    dirty |= flags;
    if (!raf) raf = requestAnimationFrame(flush);
    if (persist) save();
  }
  function flushNow() {
    if (raf) cancelAnimationFrame(raf);
    flush();
  }
  const refresh = () => schedule(D.ALL, false);

  function groupStats() {
    let total = 0, filled = 0, max = 0;
    for (const v of cardValues) {
      total += v;
      if (v > 0) { filled++; if (v > max) max = v; }
    }
    return { total: Math.round(total * 1000) / 1000, filled, max };
  }

  function flush() {
    raf = 0;
    const f = dirty;
    dirty = 0;
    if (!f) return;
    const st = groupStats();
    if (f & D.TOTALS) renderTotals(st);
    if (f & D.GROUPS) renderGroupStates(st);
    if (f & D.PICKERS) renderPickers();
    if (f & D.STATUS) renderStatus(st);
  }

  function renderTotals({ total, filled, max }) {
    animateTotal(total);
    drugInputs.forEach((inp, i) => {
      const card = groupCards[i];
      const v = cardValues[i];
      card.querySelector('.group-meter-fill').style.width = max > 0 ? `${(v / max) * 100}%` : '0';
      card.querySelector('.group-share').textContent = total > 0 && v > 0 ? `${((v / total) * 100).toFixed(1)}%` : '—';
    });
    totalProgressFill.style.width = `${(filled / cardNames.length) * 100}%`;
    totalMeta.textContent = filled ? `${filled} مجموعة · متوسط ${fmtShort(total / filled)}` : '';
    dockGroupsCount.textContent = filled;
    ringFill.style.strokeDasharray = `${RING_LEN}`;
    ringFill.style.strokeDashoffset = `${RING_LEN * (1 - filled / cardNames.length)}`;
    dockGroups.classList.toggle('is-complete', filled === cardNames.length);
    dockGroups.title = filled === cardNames.length ? 'كل المجموعات مُدخلة' : 'انتقل إلى أول مجموعة فارغة';
  }

  function renderGroupStates({ filled }) {
    drugInputs.forEach((inp, i) => {
      groupCards[i].classList.toggle('card-active', cardValues[i] > 0);
      fitChars(inp, inp.value || inp.placeholder);
    });
    optionalInputs.forEach(inp => {
      const card = inp.closest('.card');
      card.classList.toggle('card-active', inp.value !== '' && validOptional(inp));
      card.classList.toggle('card-invalid', !validOptional(inp));
    });
    groupsCount.textContent = `${filled}/${cardNames.length}`;
    $('count-all').textContent = cardNames.length;
    $('count-filled').textContent = filled;
    $('count-empty').textContent = cardNames.length - filled;
    applyGroupFilter();
  }

  function renderStatus({ filled }) {
    const missing = missingList();
    const fieldsMissing = missing.filter(m => m.step === 'step-fields').map(m => m.label);
    $('step-class').classList.toggle('is-done', !!activeClass);
    $('step-fields').classList.toggle('is-done', fieldsMissing.length === 0);
    $('step-groups').classList.toggle('is-done', filled > 0);
    $('step-totals').classList.toggle('is-done', optionalInputs.every(validOptional) && filled > 0);

    const activeBtn = activeClass && classBtnFor(activeClass);
    setChip(chipClass, activeBtn ? activeBtn.dataset.label : 'لم تُحدد الفئة', activeBtn ? 'chip-good' : 'chip-warning');
    setChip(chipFields, fieldsMissing.length ? `ناقص: ${fieldsMissing.join('، ')}` : 'بيانات الصيدلية مكتملة', fieldsMissing.length ? 'chip-warning' : 'chip-good');
    setChip(chipCards, `${filled} من ${cardNames.length} مجموعة`, filled === cardNames.length ? 'chip-good' : filled ? 'chip-progress' : '');

    const ready = missing.length === 0;
    if (isSubmitting) setChip(chipReady, 'جاري الإرسال…', 'chip-progress');
    else if (!optionalInputs.every(validOptional)) setChip(chipReady, 'تحقق من الأرقام المدخلة', 'chip-warning');
    else if (ready) setChip(chipReady, 'جاهز للإرسال ✓', 'chip-ready');
    else setChip(chipReady, `متبقٍ ${missing.length} ${missing.length === 1 ? 'عنصر' : 'عناصر'}`, '');

    const name = pharmacyName();
    const monthName = monthSel.value ? monthSel.options[monthSel.selectedIndex].text : '';
    const yearName = yearSel.value ? yearSel.options[yearSel.selectedIndex].text : '';
    const parts = [activeBtn && activeBtn.dataset.label, name, [monthName, yearName].filter(Boolean).join(' ')].filter(Boolean);
    mastheadSub.textContent = parts.join(' · ');

    if (!isSubmitting) submitBtn.setAttribute('aria-disabled', String(!ready));
    resetBtn.disabled = isSubmitting || isFormEmpty();
  }

  const A = {
    setClass(key, opts) {
      if (key === activeClass) return;
      applyClass(key, opts);
      schedule(D.ALL);
    },
    setYearStart(start) {
      const v = String(start + 1);
      if (!validEnd(v)) return false;
      ensureYearOption(v);
      if (yearSel.value !== v) {
        yearSel.value = v;
        updatePharmacyDropdown();
        schedule(D.PICKERS | D.STATUS);
      }
      return true;
    },
    setMonth(v) {
      if (monthSel.value === v) return;
      monthSel.value = v;
      updatePharmacyDropdown();
      schedule(D.PICKERS | D.STATUS);
    },
    setRegion(v) {
      if (regionSel.value === v) return;
      regionSel.value = v;
      updatePharmacyDropdown('');
      schedule(D.PICKERS | D.STATUS);
    },
    setPharmacy(cls, reg, idx) {
      if (cls !== activeClass) applyClass(cls, { keepPharmacy: true });
      regionSel.value = reg;
      updatePharmacyDropdown(String(idx + 1));
      schedule(D.ALL);
    },
    setGroup(i, value, { record = true } = {}) {
      const prev = cardValues[i];
      if (record) pushHistory(i, prev, value);
      writeCard(i, String(value), { format: true });
      schedule(D.TOTALS | D.GROUPS | D.STATUS);
      return prev;
    },
    setGroups(list, { record = true } = {}) {
      const changes = list.map(({ i, value }) => ({ i, prev: cardValues[i], next: value }));
      if (record) pushHistoryBatch(changes);
      changes.forEach(c => writeCard(c.i, String(c.next), { format: true }));
      schedule(D.TOTALS | D.GROUPS | D.STATUS);
      return changes;
    },
    restoreGroup(i, prev) {
      const top = history[history.length - 1];
      if (top && top.length === 1 && top[0].i === i && top[0].prev === prev) history.pop();
      A.setGroup(i, prev, { record: false });
    },
    setTotal(which, n) {
      const inp = which === 'tickets' ? totalTicketsInput : totalExternalInput;
      const v = Math.max(0, Math.min(MAX_OPTIONAL, Math.floor(+n) || 0));
      inp.value = v ? String(v) : '';
      schedule(D.GROUPS | D.STATUS);
    },
    undo() {
      const h = history.pop();
      if (!h) { showToast('لا يوجد ما يمكن التراجع عنه', 'info', { duration: 1800 }); return false; }
      for (let k = h.length - 1; k >= 0; k--) A.setGroup(h[k].i, h[k].prev, { record: false });
      setTarget(h[0].i, { scroll: true });
      h.forEach(c => flashCard(c.i));
      const msg = h.length === 1 ? `أُعيدت «${cardNames[h[0].i]}» إلى ${h[0].prev > 0 ? groupNum(h[0].prev) : 'فارغة'}` : `أُعيدت ${h.length} مجموعات`;
      showToast(msg, 'info', { duration: 2200 });
      return true;
    }
  };

  function guideTo(item) {
    const step = $(item.step);
    if (item.step === 'step-groups') setGroupFilter('all');
    step.scrollIntoView({ behavior: reducedMotion() ? 'auto' : 'smooth', block: 'center' });
    const target = item.el.closest('.field, .card, .class-grid') || step;
    target.classList.remove('needs-attention'); void target.offsetWidth; target.classList.add('needs-attention');
    setTimeout(() => { if (!item.el.disabled && item.el.focus) item.el.focus({ preventScroll: true }); }, reducedMotion() ? 0 : 350);
  }
  [chipClass, chipFields, chipCards].forEach(chip => chip.addEventListener('click', () => {
    const m = missingList().find(x => x.step === chip.dataset.target);
    if (m) guideTo(m); else $(chip.dataset.target).scrollIntoView({ behavior: reducedMotion() ? 'auto' : 'smooth', block: 'center' });
  }));
  document.addEventListener('animationend', e => e.target.classList.remove('needs-attention'));

  function setSubmitVisual(state) {
    submitBtn.classList.remove('is-sending', 'is-sent', 'is-failed');
    const labels = { idle: 'إرسال البيانات', sending: 'جاري الإرسال…', sent: 'تم الإرسال', failed: 'فشل الإرسال' };
    const icon = state === 'sending' ? '<span class="spinner" aria-hidden="true"></span>' : state === 'sent' ? ICONS.check : state === 'failed' ? ICONS.fail : ICONS.send;
    if (state !== 'idle') submitBtn.classList.add(`is-${state}`);
    submitInner.innerHTML = `${icon}<span id="submitButtonLabel">${labels[state]}</span>${state === 'idle' ? '<span class="submit-kbd">Ctrl ↵</span>' : ''}`;
  }

  function download(content, name, type) {
    const a = document.createElement('a');
    const url = URL.createObjectURL(new Blob([content], { type }));
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }
  const csvCell = v => { const s = String(v ?? ''); return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
  const safeName = s => (s || 'data').replace(/[^\p{L}\p{N}]+/gu, '_').replace(/^_+|_+$/g, '') || 'data';

  function buildPayload() {
    const btn = classBtnFor(activeClass);
    return {
      date: new Date().toISOString(),
      year: yearSel.value,
      class: btn ? btn.dataset.label : '',
      month: monthSel.options[monthSel.selectedIndex]?.text || '',
      region: regionSel.options[regionSel.selectedIndex]?.text || '',
      pharmacy: pharmacyName(),
      totalPrescription: +totalTicketsInput.value || 0,
      insuranceCoveredPrescription: +totalExternalInput.value || 0,
      ...Object.fromEntries(cardNames.map((n, i) => [n, cardValues[i] || 0])),
      calculatedTotal: Math.round(cardValues.reduce((s, v) => s + v, 0) * 1000) / 1000
    };
  }

  async function sendPayload(data) {
    await fetch(`${SCRIPT_URL}?data=${encodeURIComponent(JSON.stringify(data))}`, { method: 'GET', mode: 'no-cors' });
    const base = `${safeName(data.pharmacy)}_${data.date.replace(/[:.]/g, '-')}`;
    const keys = Object.keys(data);
    download(JSON.stringify(data, null, 2), `${base}.json`, 'application/json');
    await wait(350);
    download('\uFEFF' + keys.map(csvCell).join(',') + '\n' + keys.map(k => csvCell(data[k])).join(','), `${base}.csv`, 'text/csv;charset=utf-8;');
  }

  submitBtn.addEventListener('click', async () => {
    if (isSubmitting) return;
    const missing = missingList();
    if (missing.length) {
      showToast(`أكمل أولًا: ${missing.map(m => m.label).join('، ')}`, 'error', { duration: 5000 });
      submitBtn.animate([{ transform: 'translateX(0)' }, { transform: 'translateX(-6px)' }, { transform: 'translateX(6px)' }, { transform: 'translateX(-3px)' }, { transform: 'translateX(0)' }], { duration: 380 });
      guideTo(missing[0]);
      return;
    }
    if (isSent(pharmacySel.value)) {
      const again = await confirmDialog({ title: 'أُرسلت مسبقًا', desc: `«${pharmacyName()}» أُرسلت لهذا الشهر من هذا الجهاز.`, ok: 'إرسال مجددًا' });
      if (!again) return;
    }
    isSubmitting = true;
    submitBtn.setAttribute('aria-busy', 'true');
    setSubmitVisual('sending');
    refresh();
    const sentName = pharmacyName();
    try {
      await sendPayload(buildPayload());
      markSent();
      setSubmitVisual('sent');
      const burst = submitBtn.querySelector('.success-burst');
      if (!reducedMotion()) { burst.classList.remove('is-active'); void burst.offsetWidth; burst.classList.add('is-active'); }
      showToast(`أُرسلت «${sentName}» مع نسخة احتياطية.`, 'success');
      await wait(1300);
      clearForNext();
    } catch (err) {
      setSubmitVisual('failed');
      showToast('تعذّر الإرسال. بياناتك محفوظة.', 'error', { duration: 6500 });
      await wait(1800);
    } finally {
      isSubmitting = false;
      submitBtn.removeAttribute('aria-busy');
      submitBtn.querySelector('.success-burst')?.classList.remove('is-active');
      setSubmitVisual('idle');
      refresh();
    }
  });

  function clearForNext() {
    cardValues.fill(0);
    drugInputs.forEach(i => { i.value = ''; });
    optionalInputs.forEach(i => { i.value = ''; });
    updatePharmacyDropdown('');
    lastTotal = 0;
    totalEl.textContent = fmt3(0);
    history.length = 0;
    setGroupFilter('all');
    setTarget(0);
    schedule(D.ALL);
    flushNow();
    showToast('اختر الصيدلية التالية.', 'info', { duration: 4000 });
    $('field-pharmacy').scrollIntoView({ behavior: reducedMotion() ? 'auto' : 'smooth', block: 'center' });
    const nextUnsent = pharmacyChoices.querySelector('.choice:not(.is-sent)') || firstChoice(pharmacyChoices);
    setTimeout(() => { if (nextUnsent.focus) nextUnsent.focus({ preventScroll: true }); }, 400);
  }

  function resetAll() {
    applyClass(null);
    selects.forEach(s => { s.value = ''; });
    ensureYearOption(currentFiscal);
    yearSel.value = currentFiscal;
    updatePharmacyDropdown('');
    cardValues.fill(0);
    drugInputs.forEach(i => { i.value = ''; });
    optionalInputs.forEach(i => { i.value = ''; });
    history.length = 0;
    setGroupFilter('all');
    setTarget(0);
    refresh();
    flushNow();
    store.remove(DATA_KEY);
  }
  resetBtn.addEventListener('click', () => {
    if (isFormEmpty() || isSubmitting) return;
    const snap = snapshot();
    resetAll();
    resetBtn.querySelector('svg').animate([{ transform: 'rotate(0)' }, { transform: 'rotate(-360deg)' }], { duration: reducedMotion() ? 1 : 600, easing: 'cubic-bezier(0.34,1.56,0.64,1)' });
    window.scrollTo({ top: 0, behavior: reducedMotion() ? 'auto' : 'smooth' });
    showToast('تمت إعادة تعيين النموذج.', 'info', { duration: 6000, action: { label: 'تراجع', onClick: () => { restore(snap); save(); showToast('تمت استعادة البيانات.', 'success', { duration: 2500 }); } } });
  });

  function setStatusHidden(hidden) {
    statusBar.hidden = hidden;
    statusReopen.hidden = !hidden;
    try { localStorage.setItem(STATUS_KEY, hidden ? '1' : '0'); } catch (e) {}
  }
  statusDismiss.addEventListener('click', () => { setStatusHidden(true); statusReopen.focus(); });
  statusReopen.addEventListener('click', () => { setStatusHidden(false); statusDismiss.focus(); });
  try { setStatusHidden(localStorage.getItem(STATUS_KEY) === '1'); } catch (e) { setStatusHidden(false); }

  (function load() {
    const d = store.get(DATA_KEY, null);
    if (!d || typeof d !== 'object') { updatePharmacyDropdown(''); refresh(); return; }
    const btn = d.activeButton && document.getElementById(d.activeButton);
    restore({ ...d, activeClass: btn ? btn.dataset.class : null });
  })();
  (function initEntry() {
    if (!yearSel.value && hasYear(currentFiscal)) { ensureYearOption(currentFiscal); yearSel.value = currentFiscal; save(); }
    const firstEmpty = cardValues.findIndex(v => !(v > 0));
    setTarget(firstEmpty === -1 ? 0 : firstEmpty);
    refresh();
  })();

  const uploadToggle = $('upload-toggle');
  const uploadPanel = $('upload-panel');
  const dropArea = $('drop-area');
  const browseBtn = $('browse-btn');
  const fileInput = $('file-input');
  const fileListDOM = $('file-list');
  const uploadBtn = $('upload-btn');
  const uploadStopBtn = $('upload-stop-btn');
  const resultsContainer = $('results-container');
  const progressContainer = $('progress-container');
  const progressTrack = $('progress-track');
  const progressBar = $('progress-bar');
  const progressText = $('progress-text');
  const sheetDialog = $('sheet-dialog');
  const sheetList = $('sheet-list');
  const sheetConfirm = $('sheet-confirm-btn');
  const sheetCancel = $('sheet-cancel-btn');
  const sheetToggleAll = $('sheet-toggle-all');
  const sheetCount = $('sheet-dialog-count');

  $('header-chips').innerHTML = HEADERS.map(h => `<span>${escapeHTML(h)}</span>`).join('');

  let filesToUpload = [];
  let isUploading = false;
  let stopRequested = false;
  let xlsxPromise = null;
  function loadXLSX() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (!xlsxPromise) {
      xlsxPromise = new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = XLSX_URL;
        s.onload = () => resolve(window.XLSX);
        s.onerror = () => { xlsxPromise = null; reject(new Error('تعذّر تحميل مكتبة قراءة ملفات Excel.')); };
        document.head.appendChild(s);
      });
    }
    return xlsxPromise;
  }

  function setUploadOpen(open) {
    uploadPanel.classList.toggle('open', open);
    uploadToggle.classList.toggle('open', open);
    uploadToggle.setAttribute('aria-expanded', String(open));
    uploadPanel.inert = !open;
    if (open) loadXLSX().catch(() => {});
  }
  uploadToggle.addEventListener('click', () => setUploadOpen(!uploadPanel.classList.contains('open')));

  $('download-csv-btn').addEventListener('click', () => download('\uFEFF' + HEADERS.map(csvCell).join(','), 'template.csv', 'text/csv;charset=utf-8;'));
  $('download-json-btn').addEventListener('click', () => download(JSON.stringify([Object.fromEntries(HEADERS.map(h => [h, '']))], null, 2), 'template.json', 'application/json'));

  const stop = e => { e.preventDefault(); e.stopPropagation(); };
  let dragDepth = 0;
  dropArea.addEventListener('dragenter', e => { stop(e); dragDepth++; dropArea.classList.add('drag-active'); });
  dropArea.addEventListener('dragover', stop);
  dropArea.addEventListener('dragleave', e => { stop(e); dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) dropArea.classList.remove('drag-active'); });
  dropArea.addEventListener('drop', e => { stop(e); dragDepth = 0; dropArea.classList.remove('drag-active'); handleFiles(e.dataTransfer.files); });
  window.addEventListener('dragover', e => e.preventDefault());
  window.addEventListener('drop', e => { e.preventDefault(); if (!dropArea.contains(e.target) && e.dataTransfer.files.length) { setUploadOpen(true); handleFiles(e.dataTransfer.files); } });
  const openPicker = () => { if (!isUploading) fileInput.click(); };
  dropArea.addEventListener('click', openPicker);
  dropArea.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openPicker(); } });
  browseBtn.addEventListener('click', e => { e.stopPropagation(); openPicker(); });
  fileInput.addEventListener('change', e => handleFiles(e.target.files));

  const extOf = name => name.split('.').pop().toLowerCase();
  async function handleFiles(list) {
    if (isUploading) return;
    const files = Array.from(list || []);
    fileInput.value = '';
    for (const file of files) {
      const ext = extOf(file.name);
      if (!['json', 'csv', 'xlsx', 'xls'].includes(ext)) { showToast(`«${file.name}» نوع غير مدعوم.`, 'error'); continue; }
      if (filesToUpload.some(f => f.file.name === file.name && f.file.size === file.size)) { showToast(`«${file.name}» مضاف بالفعل.`, 'info', { duration: 2500 }); continue; }
      if (ext === 'xlsx' || ext === 'xls') {
        try {
          const sheets = await pickSheets(file);
          if (sheets && sheets.length) filesToUpload.push({ file, selectedSheets: sheets });
        } catch (err) { showToast(`تعذّرت قراءة «${file.name}»: ${err.message}`, 'error'); }
      } else filesToUpload.push({ file, selectedSheets: null });
      renderFileList();
    }
    renderFileList();
  }

  async function pickSheets(file) {
    const XLSX = await loadXLSX();
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', bookSheets: true });
    const sheets = wb.SheetNames;
    if (sheets.length <= 1) return sheets;
    $('sheet-dialog-file').textContent = file.name;
    sheetList.innerHTML = sheets.map((n, i) => `<li class="sheet-item"><label><input type="checkbox" data-i="${i}" checked><span>${escapeHTML(n)}</span></label></li>`).join('');
    const boxes = () => Array.from(sheetList.querySelectorAll('input'));
    const sync = () => {
      const c = boxes().filter(b => b.checked).length;
      sheetCount.textContent = `${c} من ${sheets.length} محددة`;
      sheetConfirm.disabled = c === 0;
      sheetToggleAll.textContent = c === sheets.length ? 'إلغاء تحديد الكل' : 'تحديد الكل';
    };
    sync();
    sheetList.onchange = sync;
    sheetToggleAll.onclick = () => { const all = boxes().every(b => b.checked); boxes().forEach(b => { b.checked = !all; }); sync(); };
    return new Promise(resolve => {
      const done = val => { sheetConfirm.onclick = sheetCancel.onclick = sheetDialog.onclose = null; if (sheetDialog.open) sheetDialog.close(); resolve(val); };
      sheetConfirm.onclick = () => done(boxes().filter(b => b.checked).map(b => sheets[+b.dataset.i]));
      sheetCancel.onclick = () => done(null);
      sheetDialog.onclose = () => done(null);
      sheetDialog.showModal();
      sheetConfirm.focus();
    });
  }

  function formatSize(bytes) { return bytes < 1024 ? `${bytes} B` : bytes < 1048576 ? `${(bytes / 1024).toFixed(1)} KB` : `${(bytes / 1048576).toFixed(1)} MB`; }
  function renderFileList() {
    fileListDOM.innerHTML = '';
    filesToUpload.forEach((fw, idx) => {
      const item = document.createElement('div');
      item.className = 'file-item';
      const details = [formatSize(fw.file.size), fw.selectedSheets ? `${fw.selectedSheets.length} ورقة عمل` : ''].filter(Boolean).join(' · ');
      item.innerHTML = `<span class="file-badge">${escapeHTML(extOf(fw.file.name))}</span><div class="file-info"><div class="file-name" title="${escapeHTML(fw.file.name)}">${escapeHTML(fw.file.name)}</div><div class="file-details">${escapeHTML(details)}</div></div><button class="file-remove" type="button" aria-label="إزالة ${escapeHTML(fw.file.name)}" ${isUploading ? 'disabled' : ''}>${ICONS.close}</button>`;
      item.querySelector('.file-remove').addEventListener('click', () => { filesToUpload.splice(idx, 1); renderFileList(); });
      fileListDOM.appendChild(item);
    });
    uploadBtn.disabled = isUploading || !filesToUpload.length;
    if (!isUploading) uploadBtn.textContent = filesToUpload.length ? `رفع ${filesToUpload.length === 1 ? 'الملف' : `${filesToUpload.length} ملفات`}` : 'رفع الملفات';
  }

  function setProgress(pct, text) {
    progressBar.style.width = `${pct}%`;
    progressTrack.setAttribute('aria-valuenow', String(pct));
    progressText.textContent = text;
  }
  function addResult(title, ok, msg) {
    const item = document.createElement('div');
    item.className = `result-item ${ok ? 'result-success' : 'result-error'}`;
    item.innerHTML = `${ok ? ICONS.success : ICONS.error}<div><div class="result-title">${escapeHTML(title)}</div><div class="result-msg">${escapeHTML(msg)}</div></div>`;
    resultsContainer.appendChild(item);
    resultsContainer.hidden = false;
  }

  async function processFile(fw) {
    const ext = extOf(fw.file.name);
    let rows = [];
    if (ext === 'json') {
      let d;
      try { d = JSON.parse(await fw.file.text()); } catch (e) { throw new Error('صيغة JSON غير صالحة.'); }
      rows = Array.isArray(d) ? d : [d];
    } else {
      const XLSX = await loadXLSX();
      const wb = XLSX.read(await fw.file.arrayBuffer(), { type: 'array' });
      (fw.selectedSheets || wb.SheetNames).forEach(name => { if (wb.Sheets[name]) rows.push(...XLSX.utils.sheet_to_json(wb.Sheets[name], { raw: false, dateNF: 'YYYY-MM-DD' })); });
    }
    rows = rows.filter(r => r && typeof r === 'object' && Object.values(r).some(v => String(v ?? '').trim() !== ''));
    rows.forEach((row, i) => {
      const missing = HEADERS.filter(h => !(h in row));
      if (missing.length) throw new Error(`الصف ${i + 2} ينقصه: ${missing.join('، ')}`);
    });
    return rows;
  }

  uploadStopBtn.addEventListener('click', () => { stopRequested = true; uploadStopBtn.disabled = true; uploadStopBtn.textContent = 'جاري الإيقاف…'; });

  uploadBtn.addEventListener('click', async () => {
    if (!filesToUpload.length || isUploading) return;
    isUploading = true;
    stopRequested = false;
    resultsContainer.innerHTML = '';
    resultsContainer.hidden = true;
    uploadBtn.disabled = true;
    uploadBtn.textContent = 'جاري التحضير…';
    renderFileList();

    const rows = [];
    for (const fw of filesToUpload) {
      try {
        const data = await processFile(fw);
        if (data.length) rows.push(...data.map(r => ({ ...r, _sourceFile: fw.file.name })));
        else addResult(fw.file.name, false, 'لا توجد بيانات صالحة في الملف.');
      } catch (err) { addResult(fw.file.name, false, `فشل معالجة الملف: ${err.message}`); }
    }

    if (!rows.length) {
      isUploading = false;
      if (resultsContainer.hidden) addResult('الرفع', false, 'لم يتم العثور على بيانات صالحة.');
      renderFileList();
      return;
    }

    progressContainer.hidden = false;
    uploadStopBtn.hidden = false;
    uploadStopBtn.disabled = false;
    uploadStopBtn.textContent = 'إيقاف';
    uploadBtn.textContent = 'جاري الإرسال…';
    const sent = [];
    let errors = 0;
    for (let i = 0; i < rows.length; i++) {
      if (stopRequested) break;
      setProgress(Math.round((i / rows.length) * 100), `يتم إرسال الصف ${i + 1} من ${rows.length}`);
      try {
        await fetch(`${SCRIPT_URL}?data=${encodeURIComponent(JSON.stringify(rows[i]))}`, { method: 'GET', mode: 'no-cors' });
        sent.push(rows[i]);
      } catch (e) { errors++; }
      if (i < rows.length - 1 && !stopRequested) await wait(800);
    }
    const processed = sent.length + errors;
    setProgress(Math.round((processed / rows.length) * 100), stopRequested ? `تم الإيقاف بعد ${processed} من ${rows.length}` : 'اكتمل الإرسال');
    uploadStopBtn.hidden = true;

    if (sent.length) {
      addResult('ملخص الإرسال', true, `أُرسل ${sent.length} ${sent.length === 1 ? 'صف' : 'صفوف'} دون تأكيد استلام. حُفظت نسخة احتياطية.`);
      download(JSON.stringify(sent, null, 2), `bulk_upload_backup_${new Date().toISOString().replace(/[:.]/g, '-')}.json`, 'application/json');
    }
    if (errors) addResult('أخطاء الإرسال', false, `تعذّر إرسال ${errors} ${errors === 1 ? 'صف' : 'صفوف'}.`);
    if (stopRequested && rows.length > processed) addResult('تم الإيقاف', false, `لم يُرسل ${rows.length - processed} ${rows.length - processed === 1 ? 'صف' : 'صفوف'}.`);

    isUploading = false;
    filesToUpload = [];
    renderFileList();
    uploadBtn.textContent = 'تم إرسال الملفات';
    setTimeout(() => { if (!isUploading) { progressContainer.hidden = true; setProgress(0, ''); renderFileList(); } }, 4000);
  });

  renderFileList();

  const pal = $('palette');
  const palInput = $('pal-input');
  const palList = $('pal-list');
  const palBody = $('pal-body');
  const palHl = $('pal-hl');
  const palEmpty = $('pal-empty');
  const palFlash = $('pal-flash');
  const isMac = /Mac|iPhone|iPad|iPod/i.test(navigator.platform || navigator.userAgent || '');
  const modK = isMac ? '⌘ K' : 'Ctrl K';
  document.querySelectorAll('.js-cmdk, #cmd-trigger-kbd').forEach(k => { k.textContent = modK; });
  $('pal-keys').innerHTML = '<span><kbd>↑↓</kbd>تنقّل</span><span><kbd>↵</kbd>تطبيق</span><span><kbd>Tab</kbd>إكمال</span><span><kbd>Esc</kbd>إغلاق</span>';

  const svg = d => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">${d}</svg>`;
  const PICON = {
    class: svg('<circle cx="12" cy="8.5" r="3.5"/><path d="M5 20c1-4 3.8-6 7-6s6 2 7 6"/>'),
    year: svg('<rect x="4" y="5.5" width="16" height="14" rx="2.5"/><path d="M4 10h16M8.5 3.5v3M15.5 3.5v3"/>'),
    month: svg('<rect x="4" y="5.5" width="16" height="14" rx="2.5"/><path d="M4 10h16M9 14h.01M12 14h.01M15 14h.01"/>'),
    region: svg('<path d="M12 20s6-5.4 6-10.2A6 6 0 0 0 6 9.8C6 14.6 12 20 12 20z"/><circle cx="12" cy="10" r="2"/>'),
    pharmacy: svg('<rect x="3.5" y="9" width="17" height="7" rx="3.5" transform="rotate(-35 12 12.5)"/><path d="m10 8.8 4.2 6"/>'),
    total: svg('<path d="M17 5H7l6 7-6 7h10"/>'),
    action: svg('<path d="M13 3.5 6 13h5.5l-1 7.5L18 11h-5.5z"/>'),
    hint: svg('<path d="M5 12h12M13 7l5 5-5 5"/>'),
    error: svg('<circle cx="12" cy="12" r="8.5"/><path d="M12 8v4.5M12 16h.01"/>')
  };
  const KIND = { group: 'مجموعة', class: 'فئة', year: 'سنة', month: 'شهر', region: 'منطقة', pharmacy: 'صيدلية', total: 'إجمالي', action: 'أمر', hint: 'التالي' };
  const P = { items: [], active: -1, returnFocus: null, closing: false };

  const regionLabel = v => regionSel.querySelector(`option[value="${v}"]`)?.text || '';
  const classLabel = k => classBtnFor(k)?.dataset.label || '';
  const pad2 = n => String(n).padStart(2, '0');


  function prefill(text) { palInput.value = text; palRender(); palInput.focus(); try { palInput.setSelectionRange(text.length, text.length); } catch (e) {} }
  function done(html) {
    palFlash.innerHTML = html;
    if (!reducedMotion()) palFlash.animate([{ opacity: 0, transform: 'translateY(4px)' }, { opacity: 1, transform: 'none' }], { duration: 260, easing: 'cubic-bezier(.22,.8,.24,1)' });
    palInput.value = '';
    palRender();
  }
  const ok = s => `<span class="pf-ok">${ICONS.check}</span>${s}`;

  function groupItem(i, browse) {
    const v = cardValues[i];
    return { kind: 'group', badge: pad2(i + 1), title: cardNames[i], kw: `مجموعه ${i + 1}`, on: i === target && browse,
      sub: browse && i === target ? 'الهدف الحالي' : '', val: v > 0 ? groupNum(v) : '—', dim: !(v > 0),
      complete: cardNames[i] + ' ', run: () => prefill(cardNames[i] + ' ') };
  }
  function setGroupItem(i, expr, score, isTarget) {
    const v = evalExpr(expr, cardValues[i]);
    if (!v) return null;
    if (v.error) return { kind: 'error', title: cardNames[i], sub: v.error, score, disabled: true };
    const prev = cardValues[i];
    const sub = (v.relative ? `${groupNum(prev)} ${expr.trim()[0] === '-' ? '−' : '+'} …` : prev > 0 ? `بدلًا من ${groupNum(prev)}` : 'فارغة الآن') + (isTarget ? ' · المجموعة الحالية' : '');
    return { kind: 'group', badge: pad2(i + 1), title: cardNames[i], sub, val: groupNum(v.value), strong: true, score,
      run: () => { applyGroup(i, v.value); setTarget(nextAfter(i)); done(ok(`${escapeHTML(cardNames[i])} ← <b dir="ltr">${groupNum(v.value)}</b>`)); } };
  }
  function yearItem(start, score, sub) {
    const s = +start;
    const on = yearSel.value === String(s + 1);
    return { kind: 'year', title: `<span dir="ltr">${s} – ${s + 1}</span>`, html: true, plain: `${s}`, kw: 'سنه ماليه عام', on, score,
      sub: sub || (String(s + 1) === currentFiscal ? 'السنة الحالية' : ''),
      run: () => { A.setYearStart(s); renderYear(); done(ok(`السنة المالية <b dir="ltr">${s} – ${s + 1}</b>`)); } };
  }
  function totalItems(expr, q, score0) {
    return optionalInputs.map((inp, which) => {
      const title = inp.closest('.total-field').querySelector('.total-field-head').childNodes[2].textContent.trim();
      const s = q ? textScore(q, title, 'اجمالي تذاكر خارجي') : 0;
      if (s === null) return null;
      if (expr != null) {
        const d = normalizeDigits(expr).replace(/[^\d]/g, '');
        if (!d) return null;
        const n = Math.min(+d, MAX_OPTIONAL);
        return { kind: 'total', title, sub: inp.value ? `بدلًا من ${inp.value}` : 'اختياري', val: n.toLocaleString('en-US'), strong: true, score: s + score0,
          run: () => { A.setTotal(which === 0 ? 'tickets' : 'external', n); done(ok(`${escapeHTML(title)} ← <b dir="ltr">${n.toLocaleString('en-US')}</b>`)); } };
      }
      return { kind: 'total', title, sub: 'اكتب الرقم بعد الاسم', val: inp.value ? (+inp.value).toLocaleString('en-US') : '—', dim: !inp.value, score: s + score0, complete: title + ' ', run: () => prefill(title + ' ') };
    }).filter(Boolean);
  }
  function choiceItems() {
    const items = [];
    classBtns.forEach(b => items.push({ kind: 'class', title: b.dataset.label, sub: b.querySelector('.class-desc').textContent, kw: 'فئه', on: b.dataset.class === activeClass,
      run: () => { A.setClass(b.dataset.class); done(ok(`الفئة: ${escapeHTML(b.dataset.label)}`)); } }));
    Array.from(monthSel.options).filter(o => o.value).forEach(o => items.push({ kind: 'month', badge: o.value, title: o.text, kw: `شهر ${o.value}`, on: monthSel.value === o.value,
      sub: o.value === suggestedMonth ? 'الشهر السابق' : '',
      run: () => { A.setMonth(o.value); done(ok(`الشهر: ${escapeHTML(o.text)}`)); } }));
    Array.from(regionSel.options).filter(o => o.value).forEach(o => items.push({ kind: 'region', title: `المنطقة ${o.text}`, kw: 'منطقه', on: regionSel.value === o.value,
      run: () => { A.setRegion(o.value); done(ok(`المنطقة ${escapeHTML(o.text)}`)); } }));
    Object.entries(pharmacyOptions).forEach(([cls, regions]) => Object.entries(regions).forEach(([reg, list]) => list.forEach((name, idx) => {
      const here = cls === activeClass && reg === regionSel.value;
      const on = here && pharmacySel.value === String(idx + 1);
      const sent = here && isSent(String(idx + 1));
      items.push({ kind: 'pharmacy', title: name, kw: `صيدليه ${classLabel(cls)} ${regionLabel(reg)}`, on, boost: here ? -0.5 : cls === activeClass ? -0.25 : 0,
        sub: `${classLabel(cls)} · المنطقة ${regionLabel(reg)}${sent ? ' · أُرسلت' : ''}`,
        run: () => { A.setPharmacy(cls, reg, idx); done(ok(`الصيدلية: ${escapeHTML(name)}`)); } });
    })));
    return items;
  }
  function actionItems() {
    const close = fn => () => { closePal(false); fn(); };
    const ready = missingList().length === 0;
    return [
      { kind: 'action', title: 'إرسال البيانات', kw: 'ارسال submit send', sub: ready ? 'جاهز' : `ينقص ${missingList().length}`, run: close(() => submitBtn.click()) },
      { kind: 'action', title: 'الانتقال إلى أول مجموعة فارغة', kw: 'فارغه التالي ادخال سريع', run: close(goToEntry) },
      { kind: 'action', title: 'تراجع عن آخر إدخال', kw: 'تراجع undo', sub: history.length ? '' : 'لا يوجد', run: () => { const had = history.length; undo(); if (had) done(ok('تم التراجع')); } },
      { kind: 'action', title: 'عرض المجموعات الفارغة فقط', kw: 'تصفيه فارغه', run: close(() => { setGroupFilter('empty'); $('step-groups').scrollIntoView({ behavior: reducedMotion() ? 'auto' : 'smooth', block: 'start' }); }) },
      { kind: 'action', title: 'عرض الدفتر', kw: 'قائمه list view', on: cardsContainer.classList.contains('is-list'), run: () => { setView('list', true); done(ok('عرض الدفتر')); } },
      { kind: 'action', title: 'عرض البطاقات', kw: 'شبكه grid view', on: cardsContainer.classList.contains('is-grid'), run: () => { setView('grid', true); done(ok('عرض البطاقات')); } },
      { kind: 'action', title: 'المظهر الفاتح', kw: 'فاتح light theme', on: docEl.dataset.themePref === 'light', run: () => { applyTheme('light', true); done(ok('المظهر الفاتح')); } },
      { kind: 'action', title: 'المظهر الداكن', kw: 'داكن dark theme ليلي', on: docEl.dataset.themePref === 'dark', run: () => { applyTheme('dark', true); done(ok('المظهر الداكن')); } },
      { kind: 'action', title: 'المظهر التلقائي', kw: 'تلقائي system theme', on: docEl.dataset.themePref === 'system', run: () => { applyTheme('system', true); done(ok('المظهر التلقائي')); } },
      { kind: 'action', title: 'رفع دفعة من الملفات', kw: 'رفع upload excel csv', run: close(() => { setUploadOpen(true); uploadToggle.scrollIntoView({ behavior: reducedMotion() ? 'auto' : 'smooth', block: 'start' }); }) },
      { kind: 'action', title: 'إعادة تعيين النموذج', kw: 'مسح reset', danger: true, disabled: resetBtn.disabled, run: close(() => resetBtn.click()) }
    ];
  }
  function hintItems() {
    const map = { 'الفئة': 'فئة ', 'السنة': 'سنة ', 'الشهر': 'شهر ', 'المنطقة': 'منطقة ', 'الصيدلية': 'صيدلية ' };
    return missingList().filter(m => map[m.label]).map(m => ({ kind: 'hint', title: `اختر ${m.label}`, sub: 'اكتب للبحث', complete: map[m.label], run: () => prefill(map[m.label]) }));
  }

  function palItems(raw) {
    const q = toLatin(raw).replace(/\s+/g, ' ').trim();
    if (!q) {
      const out = [];
      const hints = hintItems();
      if (hints.length) out.push({ head: 'التالي' }, ...hints);
      out.push({ head: 'المجموعات الدوائية' }, ...cardNames.map((_, i) => groupItem(i, true)));
      out.push({ head: 'أوامر' }, ...actionItems());
      return out;
    }
    const res = [];
    const push = it => { if (it) res.push(it); };
    let m = q.match(/^#?(\d{1,2})\s*[=:]\s*(.+)$/);
    if (m && +m[1] >= 1 && +m[1] <= cardNames.length) return [setGroupItem(+m[1] - 1, m[2], 0)].filter(Boolean);
    m = q.match(/^(.*?)\s*([+\-]?\d[\d.,+\-\s]*)$/);
    const hasLetters = /\p{L}/u.test(q);
    const text = m && /\p{L}/u.test(m[1]) ? m[1].trim() : (hasLetters && !m ? q : '');
    const num = m ? m[2].trim() : '';
    const yearM = num.match(/^(\d{4})(?:\s*[-–\/]\s*\d{2,4})?$/);

    if (num && text) {
      findGroups(text, cardNames, foldedNames).slice(0, 6).forEach((i, k) => push(setGroupItem(i, num, k * 0.1)));
      totalItems(num, text, 0.2).forEach(push);
      const ys = textScore(text, 'سنه ماليه', 'عام');
      if (ys !== null && yearM && validStart(yearM[1])) push(yearItem(yearM[1], ys));
      const ms = textScore(text, 'شهر');
      if (ms !== null && /^\d{1,2}$/.test(num) && +num >= 1 && +num <= 12) {
        const o = monthSel.options[+num];
        push({ kind: 'month', badge: num, title: o.text, score: ms, on: monthSel.value === num, run: () => { A.setMonth(num); done(ok(`الشهر: ${escapeHTML(o.text)}`)); } });
      }
    } else if (num && !hasLetters) {
      if (yearM && /[-–\/]/.test(num) && validStart(yearM[1])) push(yearItem(yearM[1], -1));
      else {
        push(setGroupItem(target, num, 0, true));
        if (yearM && validStart(yearM[1])) push(yearItem(yearM[1], 0.5, 'السنة المالية'));
        if (/^\d{1,2}$/.test(num) && +num >= 1 && +num <= cardNames.length) push({ ...groupItem(+num - 1, false), score: 0.8, sub: 'انتقال إلى المجموعة', run: () => { setTarget(+num - 1); prefill(cardNames[+num - 1] + ' '); } });
      }
    } else {
      const pool = [
        ...cardNames.map((_, i) => groupItem(i, false)),
        ...choiceItems(),
        ...totalItems(null, null, 0),
        ...actionItems().map(a => ({ ...a, boost: 0.3 }))
      ];
      const yq = textScore(text, 'سنه ماليه', 'عام');
      if (yq !== null) [-1, 0, 1].forEach(d => pool.push({ ...yearItem(String(+suggestStart + d), 0), boost: Math.abs(d) * 0.05 }));
      pool.forEach(it => {
        const s = it.kind === 'year' ? yq : textScore(text, it.plain || it.title, it.kw);
        if (s === null || s === undefined) return;
        it.score = s + (it.boost || 0);
        res.push(it);
      });
    }
    return res.sort((a, b) => (a.score ?? 0) - (b.score ?? 0)).slice(0, 40);
  }

  function itemHTML(it, i) {
    if (it.head) return `<li class="pal-sec" role="presentation">${escapeHTML(it.head)}</li>`;
    const icon = it.badge ? `<span class="pi-badge">${escapeHTML(it.badge)}</span>` : (PICON[it.kind] || PICON.action);
    const title = it.html ? it.title : escapeHTML(it.title);
    return `<li class="pal-item pi-${it.kind}${it.on ? ' is-on' : ''}${it.disabled ? ' is-disabled' : ''}${it.danger ? ' is-danger' : ''}" role="option" id="pal-opt-${i}" data-i="${i}" aria-selected="false"${it.disabled ? ' aria-disabled="true"' : ''}>`
      + `<span class="pi-icon">${icon}</span>`
      + `<span class="pi-main"><span class="pi-title">${title}</span>${it.sub ? `<span class="pi-sub">${escapeHTML(it.sub)}</span>` : ''}</span>`
      + (it.val ? `<span class="pi-val${it.strong ? ' is-strong' : ''}${it.dim ? ' is-dim' : ''}" dir="ltr">${escapeHTML(it.val)}</span>` : '')
      + (it.on ? `<span class="pi-check">${ICONS.check}</span>` : '')
      + (P.flat ? `<span class="pi-kind">${KIND[it.kind] || ''}</span>` : '')
      + `</li>`;
  }

  function palRender(instant) {
    const raw = palInput.value;
    P.flat = !!raw.trim();
    P.items = palItems(raw);
    palList.innerHTML = P.items.map(itemHTML).join('');
    const has = P.items.some(it => !it.head);
    palEmpty.hidden = has;
    if (!has) palEmpty.innerHTML = /\d/.test(raw) ? 'صيغة غير مفهومة — جرّب <b>قلب 4500</b> أو <b dir="ltr">5=4500</b>' : `لا نتائج لـ «${escapeHTML(raw.trim())}»`;
    const first = P.items.findIndex(it => !it.head && !it.disabled);
    setActive(first === -1 ? P.items.findIndex(it => !it.head) : first, { instant: true, scroll: 'top' });
  }

  function setActive(i, { instant = false, scroll = 'nearest' } = {}) {
    P.active = i;
    palList.querySelectorAll('[aria-selected="true"]').forEach(el => el.setAttribute('aria-selected', 'false'));
    const el = i >= 0 ? palList.querySelector(`[data-i="${i}"]`) : null;
    if (!el) { palHl.style.opacity = '0'; palInput.removeAttribute('aria-activedescendant'); return; }
    el.setAttribute('aria-selected', 'true');
    palInput.setAttribute('aria-activedescendant', el.id);
    palHl.classList.toggle('no-anim', instant);
    palHl.style.opacity = '1';
    palHl.style.transform = `translateY(${el.offsetTop}px)`;
    palHl.style.height = `${el.offsetHeight}px`;
    if (instant) { void palHl.offsetWidth; palHl.classList.remove('no-anim'); }
    if (scroll === 'top') { palBody.scrollTop = 0; return; }
    const top = el.offsetTop, bottom = top + el.offsetHeight, pad = 28;
    if (top - pad < palBody.scrollTop) palBody.scrollTop = Math.max(0, top - pad - (el.previousElementSibling?.classList.contains('pal-sec') ? 30 : 0));
    else if (bottom + 8 > palBody.scrollTop + palBody.clientHeight) palBody.scrollTop = bottom + 8 - palBody.clientHeight;
  }
  function moveActive(dir) {
    const sel = P.items.map((it, i) => (!it.head ? i : -1)).filter(i => i !== -1);
    if (!sel.length) return;
    const at = sel.indexOf(P.active);
    setActive(sel[(at + dir + sel.length) % sel.length]);
  }
  function runActive() {
    const it = P.items[P.active];
    if (!it || it.head) { rejectPal(); return; }
    if (it.disabled) { rejectPal(); return; }
    const el = palList.querySelector(`[data-i="${P.active}"]`);
    if (el && !reducedMotion()) el.animate([{ transform: 'scale(1)' }, { transform: 'scale(.985)' }, { transform: 'scale(1)' }], { duration: 180, easing: 'ease-out' });
    it.run();
  }
  function rejectPal() {
    if (reducedMotion()) return;
    pal.animate([{ transform: 'translateX(0)' }, { transform: 'translateX(-5px)' }, { transform: 'translateX(5px)' }, { transform: 'translateX(0)' }], { duration: 240, composite: 'add' });
  }

  function openPal(text = '') {
    if (pal.open) { palInput.focus(); palInput.select(); return; }
    P.returnFocus = document.activeElement;
    palFlash.textContent = '';
    palInput.value = text;
    pal.classList.remove('is-closing');
    pal.showModal();
    palRender(true);
    palInput.focus();
  }
  function closePal(restore = true) {
    if (!pal.open || P.closing) return;
    P.closing = true;
    const finish = () => {
      pal.close();
      pal.classList.remove('is-closing');
      P.closing = false;
      const back = P.returnFocus;
      if (restore && back && document.contains(back) && back !== document.body) back.focus({ preventScroll: true });
    };
    if (reducedMotion()) finish();
    else { pal.classList.add('is-closing'); setTimeout(finish, 150); }
  }

  palInput.addEventListener('input', () => { palFlash.textContent = ''; palRender(); });
  palInput.addEventListener('keydown', e => {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); moveActive(e.key === 'ArrowDown' ? 1 : -1); }
    else if (e.key === 'Enter' && !e.ctrlKey && !e.metaKey) { e.preventDefault(); runActive(); }
    else if (e.key === 'Tab') {
      const it = P.items[P.active];
      e.preventDefault();
      if (it && it.complete && palInput.value !== it.complete) prefill(it.complete);
      else moveActive(e.shiftKey ? -1 : 1);
    } else if (e.key === 'Escape') {
      e.preventDefault(); e.stopPropagation();
      if (palInput.value) { palInput.value = ''; palRender(); } else closePal();
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z' && !palInput.value) {
      e.preventDefault(); const had = history.length; undo(); if (had) done(ok('تم التراجع'));
    }
  });
  pal.addEventListener('cancel', e => { e.preventDefault(); if (palInput.value) { palInput.value = ''; palRender(); } else closePal(); });
  let palPointer = null;
  palList.addEventListener('pointermove', e => {
    if (palPointer && palPointer.x === e.clientX && palPointer.y === e.clientY) return;
    palPointer = { x: e.clientX, y: e.clientY };
    const li = e.target.closest('.pal-item');
    if (li && +li.dataset.i !== P.active) setActive(+li.dataset.i, { scroll: 'none' });
  });
  palList.addEventListener('click', e => {
    const li = e.target.closest('.pal-item');
    if (!li) return;
    setActive(+li.dataset.i, { scroll: 'none' });
    runActive();
    if (pal.open && !(window.matchMedia && window.matchMedia('(pointer: coarse)').matches)) palInput.focus({ preventScroll: true });
  });
  pal.addEventListener('pointerdown', e => {
    if (e.target !== pal) return;
    const r = pal.getBoundingClientRect();
    if (e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom) closePal();
  });
  $('pal-close').addEventListener('click', () => closePal());
  $('cmd-trigger').addEventListener('click', () => openPal());
  document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && !e.altKey && (e.code === 'KeyK' || e.key.toLowerCase() === 'k')) {
      if (document.querySelector('dialog[open]:not(#palette)')) return;
      e.preventDefault();
      pal.open ? closePal() : openPal();
    }
  });
});
