document.addEventListener('DOMContentLoaded', () => {
  const $ = id => document.getElementById(id);
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
  const yearCurrent = $('year-current');
  const monthChoices = $('month-choices');
  const regionChoices = $('region-choices');
  const pharmacyChoices = $('pharmacy-choices');
  const pharmacyHint = $('pharmacy-hint');
  const groupSearch = $('group-search');
  const searchClear = $('search-clear');
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
  const MAX_GROUP = 1000000000;
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

  function normalizeDigits(str) {
    return String(str)
      .replace(/[\u0660-\u0669]/g, d => String(d.charCodeAt(0) - 0x0660))
      .replace(/[\u06F0-\u06F9]/g, d => String(d.charCodeAt(0) - 0x06F0))
      .replace(/[\u066B,]/g, '.')
      .replace(/[\u066C\s]/g, '');
  }
  function sanitizeDecimal(str) {
    let v = normalizeDigits(str).replace(/[^\d.]/g, '');
    const i = v.indexOf('.');
    if (i !== -1) v = v.slice(0, i + 1) + v.slice(i + 1).replace(/\./g, '');
    return v;
  }
  function parseDecimal(str) {
    const n = parseFloat(sanitizeDecimal(str));
    if (isNaN(n) || n <= 0) return 0;
    return Math.min(Math.round(n * 1000) / 1000, MAX_GROUP);
  }
  const displayDecimal = n => n > 0 ? String(Math.round(n * 1000) / 1000) : '';
  const groupedDecimal = n => n > 0 ? n.toLocaleString('en-US', { maximumFractionDigits: 3 }) : '';
  const fitChars = (el, text) => el.style.setProperty('--chars', Math.max(6, String(text).length));
  const foldArabic = s => String(s)
    .replace(/[\u064B-\u065F\u0670\u0640]/g, '')
    .replace(/[أإآٱ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه').replace(/ؤ/g, 'و').replace(/ئ/g, 'ي')
    .replace(/\s+/g, ' ').trim().toLowerCase();

  let cardValues = Array(cardNames.length).fill(0);
  let activeClass = null;
  let isSubmitting = false;
  let lastTotal = 0;
  let groupFilter = 'all';
  let groupQuery = '';

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
    setClass(s.activeClass, { keepPharmacy: true });
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
  }

  const classBtnFor = key => classBtns.find(b => b.dataset.class === key);
  function setClass(key, { keepPharmacy = false } = {}) {
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
  classBtns.forEach(btn => btn.addEventListener('click', () => {
    if (btn.dataset.class === activeClass) return;
    setClass(btn.dataset.class);
    refresh(); save();
  }));
  rovingKeys(classBtns, btn => { if (btn.dataset.class !== activeClass) { setClass(btn.dataset.class); refresh(); save(); } });

  const periodKey = () => activeClass && yearSel.value && monthSel.value ? `${activeClass}|${yearSel.value}|${monthSel.value}` : null;
  const sentLog = () => store.get(SENT_KEY, {});
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

  selects.forEach(s => s.addEventListener('change', () => {
    if (s === regionSel) updatePharmacyDropdown('');
    if (s === yearSel || s === monthSel) updatePharmacyDropdown();
    refresh(); save();
  }));

  function setSelect(sel, value) {
    if (sel.value === value) return;
    sel.value = value;
    sel.dispatchEvent(new Event('change'));
  }

  const now = new Date();
  const currentFiscal = String(now.getMonth() + 1 >= 7 ? now.getFullYear() + 1 : now.getFullYear());
  const suggestedMonth = String(((now.getMonth() + 11) % 12) + 1);
  const yearValues = Array.from(yearSel.options).map(o => o.value).filter(Boolean);
  const hasYear = v => yearValues.includes(v);

  function stepYear(dir) {
    const idx = yearValues.indexOf(yearSel.value);
    if (idx === -1) { setSelect(yearSel, hasYear(currentFiscal) ? currentFiscal : yearValues[0]); return; }
    const next = yearValues[Math.max(0, Math.min(yearValues.length - 1, idx + dir))];
    setSelect(yearSel, next);
  }
  yearPrev.addEventListener('click', () => stepYear(-1));
  yearNext.addEventListener('click', () => stepYear(1));
  yearCurrent.addEventListener('click', () => {
    if (!yearSel.value) stepYear(0);
    else if (hasYear(currentFiscal) && yearSel.value !== currentFiscal) { setSelect(yearSel, currentFiscal); showToast('تم الرجوع إلى السنة المالية الحالية', 'info', { duration: 2200 }); }
  });
  [yearPrev, yearCurrent, yearNext].forEach(b => b.addEventListener('keydown', e => {
    if (e.key === 'ArrowLeft' || e.key === 'ArrowDown') { e.preventDefault(); stepYear(1); }
    if (e.key === 'ArrowRight' || e.key === 'ArrowUp') { e.preventDefault(); stepYear(-1); }
  }));

  function renderYear() {
    const opt = yearSel.options[yearSel.selectedIndex];
    const has = !!yearSel.value;
    yearCurrent.innerHTML = has ? `<span class="stepper-num">${escapeHTML(opt.text)}</span>${yearSel.value === currentFiscal ? '<span class="stepper-tag">الحالية</span>' : ''}` : `<span class="stepper-empty">اختر السنة</span>`;
    yearCurrent.title = has && yearSel.value !== currentFiscal ? 'اضغط للرجوع إلى السنة الحالية' : '';
    const idx = yearValues.indexOf(yearSel.value);
    yearPrev.disabled = has && idx <= 0;
    yearNext.disabled = has && idx >= yearValues.length - 1;
    yearCurrent.closest('.year-stepper').classList.toggle('is-set', has);
  }

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
  function bindChoices(container, sel) {
    container.addEventListener('click', e => {
      const b = e.target.closest('.choice');
      if (!b) return;
      setSelect(sel, b.dataset.value);
      const again = container.querySelector(`.choice[data-value="${CSS.escape(b.dataset.value)}"]`);
      if (again && document.activeElement !== again) again.focus({ preventScroll: true });
    });
    container.addEventListener('keydown', e => {
      const b = e.target.closest('.choice');
      if (!b) return;
      const btns = Array.from(container.querySelectorAll('.choice'));
      const i = btns.indexOf(b);
      const map = { ArrowLeft: 1, ArrowDown: 1, ArrowRight: -1, ArrowUp: -1 };
      let next = null;
      if (e.key in map) next = btns[(i + map[e.key] + btns.length) % btns.length];
      else if (e.key === 'Home') next = btns[0];
      else if (e.key === 'End') next = btns[btns.length - 1];
      if (!next) return;
      e.preventDefault();
      setSelect(sel, next.dataset.value);
      const again = container.querySelector(`.choice[data-value="${CSS.escape(next.dataset.value)}"]`);
      if (again) again.focus();
    });
  }
  bindChoices(monthChoices, monthSel);
  bindChoices(regionChoices, regionSel);
  bindChoices(pharmacyChoices, pharmacySel);

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
    const shown = format ? (document.activeElement === inp ? displayDecimal(cardValues[i]) : groupedDecimal(cardValues[i])) : (n === MAX_GROUP ? String(MAX_GROUP) : clean);
    if (inp.value !== shown) inp.value = shown;
  }

  cardsContainer.addEventListener('input', e => {
    if (!e.target.classList.contains('drug-input')) return;
    const inp = e.target;
    const before = inp.value;
    const caret = inp.selectionStart;
    setCardValue(+inp.dataset.index, before);
    if (inp.value !== before) { const pos = Math.max(0, caret - (before.length - inp.value.length)); try { inp.setSelectionRange(pos, pos); } catch (err) {} }
    refresh(); save();
  });
  cardsContainer.addEventListener('focusout', e => {
    if (!e.target.classList.contains('drug-input')) return;
    setCardValue(+e.target.dataset.index, e.target.value, { format: true });
    refresh(); save();
    setTimeout(applyGroupFilter, 0);
  });
  cardsContainer.addEventListener('focusin', e => {
    if (!e.target.classList.contains('drug-input')) return;
    const inp = e.target;
    const raw = displayDecimal(cardValues[+inp.dataset.index]);
    if (inp.value !== raw) inp.value = raw;
    fitChars(inp, inp.value);
    inp.select();
  });
  cardsContainer.addEventListener('click', e => {
    const btn = e.target.closest('.group-clear');
    if (btn) {
      const i = +btn.dataset.index;
      const prev = cardValues[i];
      setCardValue(i, '', { format: true });
      refresh(); save();
      drugInputs[i].focus();
      showToast(`تم مسح «${cardNames[i]}»`, 'info', { duration: 4000, action: { label: 'تراجع', onClick: () => { setCardValue(i, String(prev), { format: true }); refresh(); save(); } } });
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
    let filled = 0;
    cells.forEach((cell, k) => {
      const i = start + k;
      if (i >= cardNames.length) return;
      setCardValue(i, cell, { format: true });
      filled++;
    });
    refresh(); save();
    drugInputs[Math.min(start + filled, cardNames.length) - 1].focus();
    showToast(`تم لصق ${filled} قيمة بدءًا من «${cardNames[start]}»`, 'success', { duration: 3000 });
  });

  function matchesGroup(card) {
    const v = cardValues[+card.dataset.index];
    if (groupFilter === 'filled' && !(v > 0)) return false;
    if (groupFilter === 'empty' && v > 0) return false;
    if (!groupQuery) return true;
    const q = foldArabic(normalizeDigits(groupQuery).replace(/^0+(?=\d)/, ''));
    if (/^\d+$/.test(q)) return String(+card.dataset.index + 1) === q || String(+card.dataset.index + 1).startsWith(q);
    return card.dataset.search.includes(q);
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
    cardsContainer.classList.toggle('is-filtered', !!groupQuery || groupFilter !== 'all');
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

  groupSearch.addEventListener('input', () => {
    groupQuery = groupSearch.value.trim();
    searchClear.hidden = !groupQuery;
    applyGroupFilter();
  });
  groupSearch.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      e.preventDefault();
      const first = visibleDrugInputs()[0];
      if (first) { first.focus(); first.closest('.group-card').scrollIntoView({ behavior: reducedMotion() ? 'auto' : 'smooth', block: 'nearest' }); }
    } else if (e.key === 'Escape') {
      if (groupSearch.value) { e.preventDefault(); clearSearch(); }
    } else if (e.key === 'ArrowDown') {
      const first = visibleDrugInputs()[0];
      if (first) { e.preventDefault(); first.focus(); }
    }
  });
  function clearSearch() {
    groupSearch.value = '';
    groupQuery = '';
    searchClear.hidden = true;
    applyGroupFilter();
  }
  searchClear.addEventListener('click', () => { clearSearch(); groupSearch.focus(); });
  $('groups-empty-reset').addEventListener('click', () => { clearSearch(); setGroupFilter('all'); });

  function jumpToGroup(i) {
    const card = groupCards[i];
    if (card.hidden) { clearSearch(); setGroupFilter('all'); }
    card.scrollIntoView({ behavior: reducedMotion() ? 'auto' : 'smooth', block: 'center' });
    card.classList.remove('needs-attention'); void card.offsetWidth; card.classList.add('needs-attention');
    setTimeout(() => drugInputs[i].focus({ preventScroll: true }), reducedMotion() ? 0 : 320);
  }
  dockGroups.addEventListener('click', () => {
    const empty = cardValues.findIndex(v => !(v > 0));
    if (empty === -1) { $('step-totals').scrollIntoView({ behavior: reducedMotion() ? 'auto' : 'smooth', block: 'center' }); setTimeout(() => totalTicketsInput.focus({ preventScroll: true }), 300); return; }
    jumpToGroup(empty);
  });

  optionalInputs.forEach(inp => {
    inp.addEventListener('input', () => {
      let v = normalizeDigits(inp.value).replace(/\D/g, '').replace(/^0+(?=\d)/, '');
      if (v !== '' && +v > MAX_OPTIONAL) v = String(MAX_OPTIONAL);
      if (inp.value !== v) inp.value = v;
      refresh(); save();
    });
    inp.addEventListener('focus', () => inp.select());
  });

  const isTyping = el => el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.tagName === 'SELECT' || el.isContentEditable);
  document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); submitBtn.click(); return; }
    if (e.key === '/' && !isTyping(e.target) && !document.querySelector('dialog[open]')) {
      e.preventDefault();
      groupSearch.focus();
      groupSearch.select();
      return;
    }
    const order = focusOrder();
    const idx = order.indexOf(e.target);
    if (idx === -1) return;
    let dir = 0;
    if (e.key === 'Enter') dir = e.shiftKey ? -1 : 1;
    else if (e.key === 'ArrowDown') dir = 1;
    else if (e.key === 'ArrowUp') dir = -1;
    else if (e.key === 'Escape' && e.target.classList.contains('drug-input')) { e.preventDefault(); groupSearch.focus(); groupSearch.select(); return; }
    if (!dir) return;
    e.preventDefault();
    const next = order[idx + dir];
    if (next) next.focus();
    else if (dir > 0) submitBtn.focus();
    else if (e.key === 'ArrowUp' && e.target === order[0]) groupSearch.focus();
  });

  function validOptional(inp) { const v = inp.value.trim(); return v === '' || (/^\d+$/.test(v) && +v <= MAX_OPTIONAL); }
  const firstChoice = c => c.querySelector('.choice[tabindex="0"]') || c.querySelector('.choice') || c;
  function missingList() {
    const m = [];
    if (!activeClass) m.push({ label: 'الفئة', el: classBtns[0], step: 'step-class' });
    if (!yearSel.value) m.push({ label: 'السنة', el: yearCurrent, step: 'step-fields' });
    if (!monthSel.value) m.push({ label: 'الشهر', el: firstChoice(monthChoices), step: 'step-fields' });
    if (!regionSel.value) m.push({ label: 'المنطقة', el: firstChoice(regionChoices), step: 'step-fields' });
    if (!pharmacySel.value) m.push({ label: 'الصيدلية', el: firstChoice(pharmacyChoices), step: 'step-fields' });
    if (!cardValues.some(v => v > 0)) m.push({ label: 'مجموعة دوائية واحدة على الأقل', el: drugInputs[0], step: 'step-groups' });
    optionalInputs.forEach(inp => { if (!validOptional(inp)) m.push({ label: 'قيمة إجمالي صحيحة', el: inp, step: 'step-totals' }); });
    return m;
  }
  const isFormEmpty = () => !activeClass && selects.every(s => !s.value) && cardValues.every(v => !v) && optionalInputs.every(i => !i.value);

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

  function refresh() {
    const total = Math.round(cardValues.reduce((a, v) => a + v, 0) * 1000) / 1000;
    const filled = cardValues.filter(v => v > 0).length;
    const max = Math.max(...cardValues);
    animateTotal(total);

    drugInputs.forEach((inp, i) => {
      const card = groupCards[i];
      const v = cardValues[i];
      card.classList.toggle('card-active', v > 0);
      fitChars(inp, inp.value || inp.placeholder);
      card.querySelector('.group-meter-fill').style.width = max > 0 ? `${(v / max) * 100}%` : '0';
      card.querySelector('.group-share').textContent = total > 0 && v > 0 ? `${((v / total) * 100).toFixed(1)}%` : '—';
    });
    optionalInputs.forEach(inp => {
      const card = inp.closest('.card');
      card.classList.toggle('card-active', inp.value !== '' && validOptional(inp));
      card.classList.toggle('card-invalid', !validOptional(inp));
    });

    renderPickers();

    groupsCount.textContent = `${filled}/${cardNames.length}`;
    $('count-all').textContent = cardNames.length;
    $('count-filled').textContent = filled;
    $('count-empty').textContent = cardNames.length - filled;
    totalProgressFill.style.width = `${(filled / cardNames.length) * 100}%`;
    totalMeta.textContent = filled ? `${filled} مجموعة · متوسط ${fmtShort(total / filled)}` : '';
    dockGroupsCount.textContent = filled;
    ringFill.style.strokeDasharray = `${RING_LEN}`;
    ringFill.style.strokeDashoffset = `${RING_LEN * (1 - filled / cardNames.length)}`;
    dockGroups.classList.toggle('is-complete', filled === cardNames.length);
    dockGroups.title = filled === cardNames.length ? 'كل المجموعات مُدخلة' : 'انتقل إلى أول مجموعة فارغة';

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
    applyGroupFilter();
  }

  function guideTo(item) {
    const step = $(item.step);
    if (item.step === 'step-groups') { clearSearch(); setGroupFilter('all'); }
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
    clearSearch();
    setGroupFilter('all');
    refresh(); save();
    showToast('اختر الصيدلية التالية.', 'info', { duration: 4000 });
    $('field-pharmacy').scrollIntoView({ behavior: reducedMotion() ? 'auto' : 'smooth', block: 'center' });
    const nextUnsent = pharmacyChoices.querySelector('.choice:not(.is-sent)') || firstChoice(pharmacyChoices);
    setTimeout(() => { if (nextUnsent.focus) nextUnsent.focus({ preventScroll: true }); }, 400);
  }

  function resetAll() {
    setClass(null);
    selects.forEach(s => { s.value = ''; });
    updatePharmacyDropdown('');
    cardValues.fill(0);
    drugInputs.forEach(i => { i.value = ''; });
    optionalInputs.forEach(i => { i.value = ''; });
    clearSearch();
    setGroupFilter('all');
    refresh();
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
});
