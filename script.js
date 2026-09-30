(() => {
  'use strict';

  /*
   * Primary source: xlsx file.xlsx
   * Sheets: الفصل الأول 2023, الفصل الثاني 2023, الفصل الثالث 2023, الفصل الرابع 2023.
   * A faithful normalized representation of the 64 detail rows is embedded below
   * so this static page also works when index.html is opened directly with file://.
   * Fields: [term, programName, courseCount|null, durationArabic, trainees, location|null]
   */
  const rows = [
    ['Term 1','التطبيقات المكتبية',6,'أســبوعين',119,'معهد تقنية المعلومات'], ['Term 1','إعداد المراسلات',6,'خمسة أيام',117,null], ['Term 1','MS - Excel الجداول الالكترونية',5,'خمسة أيام',102,null], ['Term 1','أساسيات تشغيل الحاسب الآلي',6,'خمسة أيام',116,null], ['Term 1','معـالجة النصـوص MS - Word',6,'خمسة أيام',113,null], ['Term 1','MS-PowerPoint العروض التقديمية',4,'خمسة أيام',74,null], ['Term 1','الإنترنت والبريد الالكتروني',4,'خمسة أيام',78,null], ['Term 1','مبادئ الحوسبة السحابية',1,'خمسة أيام',21,null], ['Term 1','تطوير ذكاء الاعمال باستخدام SQL Server',1,'خمسة أيام',16,null], ['Term 1','Introduction to Python Programming',1,'خمسة أيام',18,null], ['Term 1','PMP إدارة المشاريع',1,'خمسة أيام',20,null], ['Term 1','صيانة الحاسب الآلي +A',1,'خمسة أيام',16,null], ['Term 1','(IT Governance) مقدمة في حوكمة تقنية المعلومات',1,'ثلاثة أيام',18,null], ['Term 1','مقدمة في تحليل الأعمال Foundation in Business Analysis',1,'ثلاثة أيام',21,null], ['Term 1','إدارة التغيير',1,'ثلاثة أيام',26,null], ['Term 1','مقدمة في أمن المعلومات',4,'يومين',358,null], ['Term 1','التوعية الأمنية للأجهزة التقنية',4,'يوم واحد',301,null], ['Term 1','ورشة العمل عن الجدار الناري NGFW/SDWAN',1,'يومين',13,null], ['Term 1','الأنظمة الآلية',97,'ثلاثة أشهر',2854,null],
    ['Term 2','التطبيقات المكتبية',5,'أســبوعين',97,'معهد تقنية المعلومات'], ['Term 2','MS - Excel الجداول الالكترونية',5,'خمسة أيام',97,null], ['Term 2','الإنترنت والبريد الالكتروني',4,'خمسة أيام',68,null], ['Term 2','إعداد المراسلات',5,'خمسة أيام',95,null], ['Term 2','أساسيات تشغيل الحاسب الآلي',5,'خمسة أيام',100,null], ['Term 2','صيانة الحاسب الآلي +A',3,'خمسة أيام',43,null], ['Term 2','MS-PowerPoint العروض التقديمية',4,'خمسة أيام',79,null], ['Term 2','معـالجة النصـوص MS - Word',5,'خمسة أيام',92,null], ['Term 2','إدارة التغيير',1,'ثلاثة أيام',30,null], ['Term 2','السلوك الوظيفي في بيئة العمل (للصم وضعاف السمع)',1,'يومين',20,null], ['Term 2','مقدمة في أمن المعلومات',3,'يومين',268,null], ['Term 2','مفاهيم ومبادئ إدارة الجودة ISO 9000',1,'يومين',60,null], ['Term 2','مهارات تطبيق نظام الجودة ISO 9001',1,'يومين',55,null], ['Term 2','التوعية الأمنية للأجهزة التقنية',5,'يوم واحد',381,null], ['Term 2','الأنظمة الآلية',97,'ثلاثة أشهر',2859,null],
    ['Term 3','التطبيقات المكتبية',4,'أســبوعين',81,'معهد تقنية المعلومات'], ['Term 3','MS - Excel الجداول الالكترونية',4,'خمسة أيام',79,null], ['Term 3','الإنترنت والبريد الالكتروني',2,'خمسة أيام',36,null], ['Term 3','إعداد المراسلات',4,'خمسة أيام',77,null], ['Term 3','أساسيات تشغيل الحاسب الآلي',4,'خمسة أيام',77,null], ['Term 3','صيانة الحاسب الآلي +A',4,'خمسة أيام',61,null], ['Term 3','MS-PowerPoint العروض التقديمية',3,'خمسة أيام',57,null], ['Term 3','معـالجة النصـوص MS - Word',3,'خمسة أيام',59,null], ['Term 3','البنية المؤسسية',2,'ثلاثة أيام',58,null], ['Term 3','إدارة التغيير',1,'ثلاثة أيام',26,null], ['Term 3','السلوك الوظيفي في بيئة العمل (للصم وضعاف السمع)',1,'يومين',19,null], ['Term 3','مهارات إدارة الوقت وضغوط العمل (للصم وضعاف السمع)',1,'يومين',24,null], ['Term 3','مقدمة في أمن المعلومات',1,'يومين',93,null], ['Term 3','التوعية الأمنية للأجهزة التقنية',1,'يوم واحد',90,null], ['Term 3','الأنظمة الآلية',null,'ثلاثة أشهر',2145,null],
    ['Term 4','التطبيقات المكتبية',5,'أســبوعين',95,'معهد تقنية المعلومات'], ['Term 4','MS - Excel الجداول الالكترونية',4,'خمسة أيام',70,null], ['Term 4','الإنترنت والبريد الالكتروني',4,'خمسة أيام',80,null], ['Term 4','إعداد المراسلات',4,'خمسة أيام',79,null], ['Term 4','أساسيات تشغيل الحاسب الآلي',4,'خمسة أيام',79,null], ['Term 4','صيانة الحاسب الآلي +A',2,'خمسة أيام',26,null], ['Term 4','MS-PowerPoint العروض التقديمية',4,'خمسة أيام',74,null], ['Term 4','معـالجة النصـوص MS - Word',4,'خمسة أيام',73,null], ['Term 4','PMP إدارة المشاريع',1,'خمسة أيام',24,null], ['Term 4','إدارة التغيير',1,'ثلاثة أيام',27,null], ['Term 4','السلوك الوظيفي في بيئة العمل (للصم وضعاف السمع)',1,'يومين',16,null], ['Term 4','مهارات إدارة الوقت وضغوط العمل (للصم وضعاف السمع)',1,'يومين',28,null], ['Term 4','مقدمة في أمن المعلومات',1,'يومين',91,null], ['Term 4','التوعية الأمنية للأجهزة التقنية',1,'يوم واحد',87,null], ['Term 4','الأنظمة الآلية',100,'ثلاثة أشهر',1999,null]
  ].map(([term, name, courses, duration, trainees, location], index) => ({ id: index + 1, term, name, courses, duration, trainees, location }));

  const SHEETS = [
    { term: 'Term 1', source: 'الفصل الأول 2023', reported: 4401 },
    { term: 'Term 2', source: 'الفصل الثاني 2023', reported: 4344 },
    { term: 'Term 3', source: 'الفصل الثالث 2023', reported: 2982 },
    { term: 'Term 4', source: 'الفصل الرابع 2023', reported: 2848 }
  ];
  const durationNames = { 'يوم واحد': 'One day', 'يومين': 'Two days', 'ثلاثة أيام': 'Three days', 'خمسة أيام': 'Five days', 'أســبوعين': 'Two weeks', 'ثلاثة أشهر': 'Three months' };
  const PROGRAM_NAMES = Object.freeze({
    'التطبيقات المكتبية': 'Office Applications',
    'إعداد المراسلات': 'Business Correspondence Preparation',
    'MS - Excel الجداول الالكترونية': 'Microsoft Excel Spreadsheets',
    'أساسيات تشغيل الحاسب الآلي': 'Computer Operation Fundamentals',
    'معـالجة النصـوص MS - Word': 'Microsoft Word Word Processing',
    'MS-PowerPoint العروض التقديمية': 'Microsoft PowerPoint Presentations',
    'الإنترنت والبريد الالكتروني': 'Internet and Email',
    'مبادئ الحوسبة السحابية': 'Cloud Computing Fundamentals',
    'تطوير ذكاء الاعمال باستخدام SQL Server': 'Business Intelligence Development Using SQL Server',
    'Introduction to Python Programming': 'Introduction to Python Programming',
    'PMP إدارة المشاريع': 'Project Management Professional (PMP)',
    'صيانة الحاسب الآلي +A': 'A+ Computer Maintenance',
    '(IT Governance) مقدمة في حوكمة تقنية المعلومات': 'Introduction to IT Governance',
    'مقدمة في تحليل الأعمال Foundation in Business Analysis': 'Foundation in Business Analysis',
    'إدارة التغيير': 'Change Management',
    'مقدمة في أمن المعلومات': 'Introduction to Information Security',
    'التوعية الأمنية للأجهزة التقنية': 'Information Security Awareness for Technical Devices',
    'ورشة العمل عن الجدار الناري NGFW/SDWAN': 'NGFW/SD-WAN Firewall Workshop',
    'الأنظمة الآلية': 'Automated Systems',
    'السلوك الوظيفي في بيئة العمل (للصم وضعاف السمع)': 'Workplace Conduct for Deaf and Hard-of-Hearing Participants',
    'مفاهيم ومبادئ إدارة الجودة ISO 9000': 'ISO 9000 Quality Management Concepts and Principles',
    'مهارات تطبيق نظام الجودة ISO 9001': 'ISO 9001 Quality Management System Implementation Skills',
    'البنية المؤسسية': 'Enterprise Architecture',
    'مهارات إدارة الوقت وضغوط العمل (للصم وضعاف السمع)': 'Time and Work-Stress Management for Deaf and Hard-of-Hearing Participants'
  });
  const durationOrder = ['يوم واحد', 'يومين', 'ثلاثة أيام', 'خمسة أيام', 'أســبوعين', 'ثلاثة أشهر'];
  const palette = ['#155eef', '#047857', '#7e22ce', '#c2410c', '#087ea4', '#b54708'];
  const viewOrder = ['overview', 'terms', 'programs', 'duration', 'participation'];
  const hashViews = { '#overview': 'overview', '#terms': 'terms', '#programs': 'programs', '#duration': 'duration', '#participation': 'participation' };
  const state = { activeView: 'overview', rendered: new Set(), search: '', term: 'all', duration: 'all' };
  const $ = (selector, root = document) => root.querySelector(selector);
  const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];
  const numberFormatter = new Intl.NumberFormat('en-US');
  const reduceMotion = matchMedia('(prefers-reduced-motion: reduce)').matches;
  const tooltip = $('#tooltip');

  const format = (value) => numberFormatter.format(value);
  const formatValue = (value, decimals = 0) => new Intl.NumberFormat('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(value);
  const percentage = (value, total) => total ? `${formatValue(value / total * 100, 1)}%` : '0.0%';
  const escapeHTML = (value) => String(value).replace(/[&<>'"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[char]));
  const programDisplayName = (sourceTitle) => PROGRAM_NAMES[sourceTitle] || sourceTitle;
  const originalArabicTitle = (sourceTitle) => sourceTitle === 'Introduction to Python Programming' ? null : sourceTitle;
  const normalizeSearch = (value) => String(value).toLocaleLowerCase().normalize('NFKD').replace(/[ـً-ٰٟ]/g, '').replace(/[أإآ]/g, 'ا').replace(/\s+/g, ' ').trim();
  const programTitleMarkup = (sourceTitle, { compact = false } = {}) => {
    const original = originalArabicTitle(sourceTitle);
    return `<span class="program-title${compact ? ' is-compact' : ''}"><strong>${escapeHTML(programDisplayName(sourceTitle))}</strong>${original ? `<span class="original-title"><b>Original Arabic Title:</b> <span lang="ar" dir="rtl">${escapeHTML(original)}</span></span>` : '<span class="original-title"><b>Original Arabic Title:</b> Not applicable — source title is already English.</span>'}</span>`;
  };
  const originalTitleTooltip = (sourceTitle) => {
    const original = originalArabicTitle(sourceTitle);
    return original ? `<span class="original-title"><b>Original Arabic Title:</b> <span lang="ar" dir="rtl">${escapeHTML(original)}</span></span>` : '<span class="original-title"><b>Original Arabic Title:</b> Not applicable — source title is already English.</span>';
  };

  const totalTrainees = rows.reduce((sum, row) => sum + row.trainees, 0);
  const totalCourses = rows.reduce((sum, row) => sum + (row.courses ?? 0), 0);
  const missingCourses = rows.filter((row) => row.courses === null).length;
  const missingLocations = rows.filter((row) => !row.location).length;

  function totalsByTerm() {
    return SHEETS.map((sheet, index) => {
      const termRows = rows.filter((row) => row.term === sheet.term);
      const trainees = termRows.reduce((sum, row) => sum + row.trainees, 0);
      const courses = termRows.reduce((sum, row) => sum + (row.courses ?? 0), 0);
      const missing = termRows.filter((row) => row.courses === null).length;
      const previous = index ? null : undefined;
      return { ...sheet, trainees, records: termRows.length, courses, missing, share: trainees / totalTrainees, previous };
    }).map((item, index, data) => index ? { ...item, change: item.trainees - data[index - 1].trainees, changeRate: (item.trainees - data[index - 1].trainees) / data[index - 1].trainees } : { ...item, change: null, changeRate: null });
  }

  function groupPrograms() {
    const map = new Map();
    rows.forEach((row) => {
      if (!map.has(row.name)) map.set(row.name, { name: row.name, displayName: programDisplayName(row.name), originalArabicTitle: originalArabicTitle(row.name), trainees: 0, courses: 0, terms: 0, courseMissing: 0, byTerm: {} });
      const item = map.get(row.name);
      item.trainees += row.trainees;
      item.courses += row.courses ?? 0;
      item.terms += 1;
      item.courseMissing += row.courses === null ? 1 : 0;
      item.byTerm[row.term] = row.trainees;
    });
    return [...map.values()].map((item) => ({ ...item, share: item.trainees / totalTrainees })).sort((a, b) => b.trainees - a.trainees);
  }

  function durationStats() {
    return durationOrder.map((duration) => {
      const subset = rows.filter((row) => row.duration === duration);
      const trainees = subset.reduce((sum, row) => sum + row.trainees, 0);
      return { duration, label: durationNames[duration], records: subset.length, programs: new Set(subset.map((row) => row.name)).size, courses: subset.reduce((sum, row) => sum + (row.courses ?? 0), 0), missing: subset.filter((row) => row.courses === null).length, trainees, share: trainees / totalTrainees };
    });
  }

  function getDataQuality() {
    const sourceTitles = new Set(rows.map((row) => row.name));
    const mappingComplete = [...sourceTitles].every((title) => Object.hasOwn(PROGRAM_NAMES, title)) && Object.keys(PROGRAM_NAMES).length === sourceTitles.size;
    return { missingCourses, missingLocations, mappingComplete, valid: rows.length === 64 && totalTrainees === 14575 && totalCourses === 468 && totalsByTerm().every((item) => item.trainees === item.reported) && mappingComplete };
  }

  function icon(type) {
    const icons = {
      people: '<svg viewBox="0 0 24 24"><circle cx="9" cy="8" r="3"/><path d="M3.5 20c.4-3.8 2.4-6 5.5-6s5.1 2.2 5.5 6M16 5.5a3 3 0 0 1 0 5.7M18 14c1.7.6 2.7 2.5 3 5"/></svg>',
      records: '<svg viewBox="0 0 24 24"><path d="M7 3h8l3 3v15H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2Z"/><path d="M15 3v4h4M9 12h6M9 16h6"/></svg>',
      courses: '<svg viewBox="0 0 24 24"><path d="M4 5.5A2.5 2.5 0 0 1 6.5 3H20v16H6.5A2.5 2.5 0 0 0 4 21.5v-16Z"/><path d="M4 5.5A2.5 2.5 0 0 1 6.5 8H20M9 11h7M9 15h5"/></svg>',
      average: '<svg viewBox="0 0 24 24"><path d="M4 19V9M10 19V5M16 19v-7M22 19V3"/><path d="M2 20h20"/></svg>'
    };
    return icons[type];
  }

  function chartButton({ chart, index, label, style = '', className = 'chart-hit' }) {
    return `<button class="${className}" type="button" data-chart="${chart}" data-index="${index}" style="${style}" aria-label="${escapeHTML(label)}"></button>`;
  }

  function bindTooltip(element, siblings, content) {
    const show = () => showTooltip(element, content());
    element.addEventListener('pointerenter', show);
    element.addEventListener('focus', show);
    element.addEventListener('pointerleave', hideTooltip);
    element.addEventListener('blur', hideTooltip);
    element.addEventListener('keydown', (event) => keyboardNavigate(event, siblings));
  }

  function bindChartTooltips(root, chart, content) {
    const items = $$(`[data-chart="${chart}"]`, root);
    items.forEach((element, index) => bindTooltip(element, items, () => content(index)));
  }

  function keyboardNavigate(event, siblings) {
    const index = siblings.indexOf(event.currentTarget);
    let next = null;
    if (event.key === 'ArrowRight' || event.key === 'ArrowDown') next = Math.min(siblings.length - 1, index + 1);
    if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') next = Math.max(0, index - 1);
    if (event.key === 'Home') next = 0;
    if (event.key === 'End') next = siblings.length - 1;
    if (event.key === 'Escape') { hideTooltip(); event.currentTarget.blur(); return; }
    if (next !== null) { event.preventDefault(); siblings[next].focus(); }
  }

  function showTooltip(element, content) {
    const rect = element.getBoundingClientRect();
    tooltip.innerHTML = content;
    tooltip.style.left = `${Math.min(Math.max(rect.left + rect.width / 2, 135), innerWidth - 135)}px`;
    tooltip.style.top = `${Math.max(rect.top, 70)}px`;
    tooltip.classList.add('is-visible');
    tooltip.setAttribute('aria-hidden', 'false');
    element.setAttribute('aria-describedby', 'tooltip');
  }

  function hideTooltip() {
    tooltip.classList.remove('is-visible');
    tooltip.setAttribute('aria-hidden', 'true');
    $$('[aria-describedby="tooltip"]').forEach((element) => element.removeAttribute('aria-describedby'));
  }

  function animateNumbers(root) {
    $$('[data-count]', root).forEach((element) => {
      if (element.dataset.animated) return;
      const target = Number(element.dataset.count);
      const decimals = Number(element.dataset.decimals || 0);
      element.dataset.animated = 'true';
      if (reduceMotion) { element.textContent = formatValue(target, decimals); return; }
      const start = performance.now();
      const tick = (now) => {
        const ratio = Math.min((now - start) / 520, 1);
        element.textContent = formatValue(target * (1 - Math.pow(1 - ratio, 3)), decimals);
        if (ratio < 1) requestAnimationFrame(tick);
      };
      requestAnimationFrame(tick);
    });
  }

  function panelHeading(number, title, summary) {
    return `<div class="view-title-row"><div><p class="section-kicker">${number}</p><h2>${title}</h2></div><p>${summary}</p></div>`;
  }

  function evidencePanels(evidence, questions) {
    return `<div class="evidence-grid"><article class="evidence-card"><span class="evidence-label">Evidence from the dataset</span><h3>${evidence.title}</h3><p>${evidence.text}</p></article><article class="question-card"><span class="evidence-label">Questions for further investigation</span><h3>${questions.title}</h3><p>${questions.text}</p></article></div>`;
  }

  function emptyState(message, action = '') {
    return `<div class="empty-state"><strong>No displayable data</strong><p>${message}</p>${action}</div>`;
  }

  function renderOverview(root) {
    const terms = totalsByTerm();
    const firstToFourth = (terms[3].trainees - terms[0].trainees) / terms[0].trainees;
    root.innerHTML = `${panelHeading('01 · Overview', 'The annual training picture, at a glance', 'Four workbook sheets contribute 64 program records. Empty source fields remain missing rather than estimated.')}
      <div class="kpi-grid" aria-label="Key measures from the local workbook">
        ${[
          { label: 'Reported trainee counts', value: totalTrainees, context: 'Sum across 64 program records', color: '#155eef', type: 'people' },
          { label: 'Program records', value: rows.length, context: 'Term-sheet detail rows, not unique programs', color: '#047857', type: 'records' },
          { label: 'Available course count', value: totalCourses, context: `${missingCourses} blank course-count cell excluded`, color: '#7e22ce', type: 'courses' },
          { label: 'Average trainees / record', value: totalTrainees / rows.length, context: 'Reported participation volume, not completion', color: '#c2410c', type: 'average', decimals: 1 }
        ].map((card) => `<article class="kpi-card"><div><p class="kpi-label">${card.label}</p><p class="kpi-value" data-count="${card.value}" data-decimals="${card.decimals || 0}">${reduceMotion ? formatValue(card.value, card.decimals) : '0'}</p></div><div class="kpi-bottom"><p class="kpi-context">${card.context}</p><span class="kpi-mark" style="--kpi-color:${card.color}" aria-hidden="true">${icon(card.type)}</span></div></article>`).join('')}
      </div>
      <div class="chart-grid overview-grid">
        <figure class="panel line-panel" aria-labelledby="overview-trend-title"><div class="panel-header"><div><p class="panel-kicker">Reported trainee trend</p><h3 id="overview-trend-title">Four training terms</h3></div><span class="source-badge">4 points</span></div><p class="chart-summary">Each point is the sum of trainee counts across the program records in one workbook sheet. It is term order, not a continuous time series.</p>${renderLineChart(terms, 'overview-trend')}<figcaption>Use Tab then arrow keys to inspect each term. Direct labels and the table below provide the same values.</figcaption></figure>
        <article class="panel scope-card"><p class="panel-kicker">Scope & data quality</p><h3>Useful signals, bounded by the workbook</h3><ul class="scope-list"><li><strong>24</strong> exact recorded program names</li><li><strong>${format(missingLocations)} / ${rows.length}</strong> locations not recorded</li><li><strong>${missingCourses}</strong> course-count value missing</li><li><strong>0</strong> outcome or satisfaction fields</li></ul><p>Reported trainee counts can repeat people across terms; the workbook has no learner identifier for deduplication.</p></article>
      </div>
      <div class="data-table-card"><h3>Term totals behind the trend</h3><div class="compact-table-wrap"><table class="compact-table"><thead><tr><th>Term</th><th>Trainees</th><th>Annual share</th><th>Records</th></tr></thead><tbody>${terms.map((term) => `<tr><td>${term.term}</td><td>${format(term.trainees)}</td><td>${percentage(term.trainees, totalTrainees)}</td><td>${term.records}</td></tr>`).join('')}</tbody></table></div></div>
      ${evidencePanels({ title: `Trainee totals are ${percentage(Math.abs(firstToFourth), 1)} lower in Term 4 than Term 1`, text: `The records sum to ${format(terms[0].trainees)} in Term 1 and ${format(terms[3].trainees)} in Term 4, a ${formatValue(Math.abs(firstToFourth) * 100, 1)}% decrease. The workbook shows the pattern, not its cause.` }, { title: 'What changed between terms?', text: 'Review scheduling, availability, registration conditions, and reporting practice with operational records not present in this workbook. This dataset alone cannot identify a cause.' })}`;
    animateNumbers(root);
  }

  function renderLineChart(data, chart) {
    const width = 620; const height = 270; const pad = { top: 26, right: 26, bottom: 55, left: 48 };
    const max = Math.ceil(Math.max(...data.map((item) => item.trainees)) / 500) * 500;
    const min = 0;
    const x = (index) => pad.left + index * ((width - pad.left - pad.right) / (data.length - 1));
    const y = (value) => height - pad.bottom - ((value - min) / (max - min)) * (height - pad.top - pad.bottom);
    const points = data.map((item, index) => `${x(index)},${y(item.trainees)}`).join(' ');
    const grid = [0, .5, 1].map((fraction) => { const value = max * fraction; return `<g><line x1="${pad.left}" x2="${width - pad.right}" y1="${y(value)}" y2="${y(value)}"/><text x="${pad.left - 10}" y="${y(value) + 4}" text-anchor="end">${format(value)}</text></g>`; }).join('');
    return `<div class="svg-chart line-chart" role="img" aria-label="Line chart of reported trainee counts by term"><svg viewBox="0 0 ${width} ${height}" aria-hidden="true"><g class="chart-gridlines">${grid}</g><polyline class="trend-line" points="${points}"/>${data.map((item, index) => `<g><circle class="trend-dot" cx="${x(index)}" cy="${y(item.trainees)}" r="7"/><text class="point-value" x="${x(index)}" y="${y(item.trainees) - 15}" text-anchor="middle">${format(item.trainees)}</text><text class="axis-label" x="${x(index)}" y="${height - 17}" text-anchor="middle">${item.term}</text></g>`).join('')}</svg><div class="svg-hit-layer">${data.map((item, index) => chartButton({ chart, index, label: `${item.term}, ${format(item.trainees)} reported trainees`, style: `left:${(x(index) / width) * 100}%;top:${(y(item.trainees) / height) * 100}%` })).join('')}</div></div>`;
  }

  function renderTerms(root) {
    const data = totalsByTerm();
    root.innerHTML = `${panelHeading('02 · Term comparison', 'Compare the four reported training terms', 'The two charts show different measures in separate panels. Course totals exclude one blank Term 3 source cell; they are not shown on a dual axis.')}
      <div class="chart-grid">
        <figure class="panel" aria-labelledby="term-bar-title"><div class="panel-header"><div><p class="panel-kicker">Participation volume</p><h3 id="term-bar-title">Reported trainees by term</h3></div><span class="source-badge">Sum of trainees</span></div><p class="chart-summary">The bar height represents the total trainee count in each term’s detail rows.</p>${renderVerticalBars(data, 'term-trainees', 'trainees', '#155eef')}<figcaption>Each total is validated against the matching source term-total row.</figcaption></figure>
        <figure class="panel" aria-labelledby="term-line-title"><div class="panel-header"><div><p class="panel-kicker">Term-to-term pattern</p><h3 id="term-line-title">Reported trainee line</h3></div><span class="source-badge">Same totals</span></div><p class="chart-summary">The line repeats the same trainee totals to emphasize change across the ordered training terms.</p>${renderLineChart(data, 'term-line')}<figcaption>A line does not imply values between terms; it only connects four recorded totals.</figcaption></figure>
      </div>
      <div class="data-table-card"><div class="panel-header"><div><p class="panel-kicker">Detailed term summary</p><h3>Calculated from program records</h3></div><span class="source-badge">4 rows</span></div><div class="table-wrap"><table><caption class="sr-only">Term comparison summary</caption><thead><tr><th>Term</th><th>Reported trainees</th><th>Annual share</th><th>Program records</th><th>Available courses</th><th>Change from prior term</th></tr></thead><tbody>${data.map((item) => `<tr><td><strong>${item.term}</strong><br><span dir="auto">${escapeHTML(item.source)}</span></td><td><strong>${format(item.trainees)}</strong></td><td>${percentage(item.trainees, totalTrainees)}</td><td>${item.records}</td><td>${format(item.courses)}${item.missing ? ' <span class="missing">(1 missing)</span>' : ''}</td><td>${item.change === null ? 'Not applicable' : `<span class="delta ${item.change < 0 ? 'is-negative' : 'is-positive'}">${item.change > 0 ? '+' : ''}${format(item.change)} · ${item.changeRate > 0 ? '+' : ''}${formatValue(item.changeRate * 100, 1)}%</span>`}</td></tr>`).join('')}</tbody></table></div></div>
      ${evidencePanels({ title: 'The largest reported change occurs from Term 2 to Term 3', text: `Trainee totals decrease by ${format(Math.abs(data[2].change))} (${formatValue(Math.abs(data[2].changeRate) * 100, 1)}%) between those adjacent terms. That calculation reflects record totals only.` }, { title: 'Why does Term 3 differ?', text: 'The workbook cannot establish whether schedule, program mix, registrations, capacity, or reporting conditions contributed to the change. Compare those operational sources before drawing a conclusion.' })}`;
  }

  function renderVerticalBars(data, chart, metric, color) {
    const maximum = Math.max(...data.map((item) => item[metric]));
    return `<div class="vertical-bar-chart">${data.map((item, index) => `<div class="vertical-bar-group"><span class="bar-value">${format(item[metric])}</span><div class="bar-slot"><span class="vertical-fill" style="--bar-height:${Math.max(6, item[metric] / maximum * 100)}%;--bar-color:${color}"></span>${chartButton({ chart, index, label: `${item.term}, ${format(item[metric])} ${metric === 'trainees' ? 'reported trainees' : 'available courses'}`, className: 'bar-hit' })}</div><span class="bar-label">${item.term}</span></div>`).join('')}</div>`;
  }

  function renderPrograms(root) {
    const programs = groupPrograms();
    const top = programs.slice(0, 8);
    root.innerHTML = `${panelHeading('03 · Program participation', 'Which program titles have the greatest reported trainee volume?', 'Program titles are grouped only when the exact source title matches. English labels are provided for readability, while the original Arabic titles remain visible for source traceability. These are reported participation counts, not unique people or measures of program quality.')}
      <div class="chart-grid program-grid">
        <figure class="panel" aria-labelledby="program-rank-title"><div class="panel-header"><div><p class="panel-kicker">Ranked trainee volume</p><h3 id="program-rank-title">Top 8 recorded programs</h3></div><span class="source-badge">Exact name groups</span></div><p class="chart-summary">Each bar sums trainee counts for identical recorded program names across all four terms.</p>${renderHorizontalBars(top, 'program-rank', 'trainees', 'reported trainees')}<figcaption>Use focus or hover to see the per-term breakdown and available course count.</figcaption></figure>
        <figure class="panel scatter-panel" aria-labelledby="course-trainee-title"><div class="panel-header"><div><p class="panel-kicker">Two measures, separate axes</p><h3 id="course-trainee-title">Available courses vs. trainees</h3></div><span class="source-badge">24 programs</span></div><p class="chart-summary">Each point is an exact-name program grouping: x is available course count and y is reported trainees. One blank source course count is excluded.</p>${renderScatter(programs, 'program-scatter')}<figcaption>Position supports comparison only; it does not indicate causation or program quality.</figcaption></figure>
      </div>
      <div class="data-table-card full-ranking"><div class="panel-header"><div><p class="panel-kicker">All program groups</p><h3>Annual ranked detail</h3></div><span class="source-badge">${programs.length} exact source titles</span></div><div class="table-wrap"><table><caption class="sr-only">Annual program participation ranking</caption><thead><tr><th>Rank</th><th>Program title</th><th>Reported trainees</th><th>Annual share</th><th>Available courses</th><th>Terms recorded</th></tr></thead><tbody>${programs.map((item, index) => `<tr><td>${index + 1}</td><td>${programTitleMarkup(item.name)}</td><td><strong>${format(item.trainees)}</strong></td><td>${percentage(item.trainees, totalTrainees)}</td><td>${format(item.courses)}${item.courseMissing ? ' <span class="missing">(1 missing)</span>' : ''}</td><td>${item.terms}</td></tr>`).join('')}</tbody></table></div></div>
      <section class="data-table-card record-explorer" aria-labelledby="record-explorer-title"><div class="panel-header"><div><p class="panel-kicker">Explore source records</p><h3 id="record-explorer-title">Search and filter program records</h3></div><span class="source-badge">64 source rows</span></div><p class="chart-summary">Filters apply only to the detailed source-derived records below. Search in English or Arabic. Missing values remain visible rather than being converted to zero.</p><div class="filters" aria-label="Filter program records"><label class="filter-field"><span>Search English or Arabic program title</span><input id="program-search" type="search" value="${escapeHTML(state.search)}" placeholder="e.g., Office Applications or التطبيقات المكتبية"></label><label class="filter-field"><span>Training term</span><select id="term-filter"><option value="all">All terms</option>${SHEETS.map((sheet) => `<option value="${sheet.term}"${state.term === sheet.term ? ' selected' : ''}>${sheet.term}</option>`).join('')}</select></label><label class="filter-field"><span>Duration</span><select id="duration-filter"><option value="all">All durations</option>${durationOrder.map((duration) => `<option value="${duration}"${state.duration === duration ? ' selected' : ''}>${durationNames[duration]}</option>`).join('')}</select></label><button class="clear-filters" id="clear-filters" type="button">Clear filters</button></div><p class="table-status" id="program-table-status" role="status" aria-live="polite"></p><div class="table-wrap"><table><caption class="sr-only">Filtered Institute of Information Technology 2023 training-program records</caption><thead><tr><th>Term</th><th>Program title</th><th>Duration</th><th>Courses</th><th>Reported trainees</th><th>Training location</th></tr></thead><tbody id="program-table-body"></tbody></table></div></section>
      ${evidencePanels({ title: 'Reported trainee volume is concentrated in one program title', text: `${escapeHTML(programs[0].displayName)} sums to ${format(programs[0].trainees)} trainee counts, ${percentage(programs[0].trainees, totalTrainees)} of the annual total. This is a volume concentration, not evidence of outcomes or demand quality.` }, { title: 'What capacity or reporting factors should be reviewed?', text: 'Operational teams can compare staffing, capacity, availability, registrations, and reporting practices for high-volume offerings. Those factors are not recorded in this workbook.' })}`;
  }

  function renderHorizontalBars(data, chart, metric, metricLabel) {
    const maximum = Math.max(...data.map((item) => item[metric]));
    return `<div class="horizontal-bar-chart">${data.map((item, index) => {
      const sourceTitle = item.name;
      const label = item.displayName || item.label || programDisplayName(sourceTitle);
      const titleMarkup = sourceTitle && item.displayName ? programTitleMarkup(sourceTitle, { compact: true }) : `<strong>${escapeHTML(label)}</strong>`;
      return `<div class="horizontal-row"><span class="rank-number">${String(index + 1).padStart(2, '0')}</span><span class="horizontal-name" title="${escapeHTML(label)}">${titleMarkup}</span><span class="horizontal-track"><span class="horizontal-fill" style="--fill:${item[metric] / maximum * 100}%;--bar-color:#155eef"></span>${chartButton({ chart, index, label: `${label}, ${format(item[metric])} ${metricLabel}`, className: 'horizontal-hit' })}</span><strong class="horizontal-value">${format(item[metric])}</strong></div>`;
    }).join('')}</div>`;
  }

  function renderScatter(data, chart) {
    const width = 560; const height = 315; const pad = { top: 25, right: 25, bottom: 48, left: 55 };
    const maxX = Math.ceil(Math.max(...data.map((item) => item.courses)) / 50) * 50;
    const maxY = Math.ceil(Math.max(...data.map((item) => item.trainees)) / 2000) * 2000;
    const xTicks = [0, .5, 1];
    const x = (value) => pad.left + value / maxX * (width - pad.left - pad.right);
    const y = (value) => height - pad.bottom - value / maxY * (height - pad.top - pad.bottom);
    const ticks = [0, .5, 1];
    return `<div class="svg-chart scatter-chart" role="img" aria-label="Scatter chart comparing available course count and reported trainee counts by program"><svg viewBox="0 0 ${width} ${height}" aria-hidden="true"><g class="chart-gridlines">${ticks.map((fraction) => `<line x1="${pad.left}" x2="${width - pad.right}" y1="${y(maxY * fraction)}" y2="${y(maxY * fraction)}"/><text x="${pad.left - 9}" y="${y(maxY * fraction) + 4}" text-anchor="end">${format(maxY * fraction)}</text>`).join('')}</g><line class="chart-axis" x1="${pad.left}" x2="${width - pad.right}" y1="${height - pad.bottom}" y2="${height - pad.bottom}"/>${xTicks.map((fraction) => `<text class="axis-label" x="${x(maxX * fraction)}" y="${height - pad.bottom + 18}" text-anchor="middle">${format(maxX * fraction)}</text>`).join('')}<text class="axis-title" x="${width / 2}" y="${height - 10}" text-anchor="middle">Available course count</text><text class="axis-title" transform="translate(15 ${height / 2}) rotate(-90)" text-anchor="middle">Reported trainees</text>${data.map((item, index) => `<circle class="scatter-dot" cx="${x(item.courses)}" cy="${y(item.trainees)}" r="${index === 0 ? 8 : 5}" style="--dot-color:${index === 0 ? '#c2410c' : '#155eef'}"/>`).join('')}</svg><div class="svg-hit-layer">${data.map((item, index) => chartButton({ chart, index, label: `${item.displayName || programDisplayName(item.name)}, ${format(item.courses)} available courses and ${format(item.trainees)} reported trainees`, style: `left:${(x(item.courses) / width) * 100}%;top:${(y(item.trainees) / height) * 100}%` })).join('')}</div></div>`;
  }

  function renderDuration(root) {
    const data = durationStats();
    root.innerHTML = `${panelHeading('04 · Duration', 'Compare duration categories without mixing their measures', 'The six labels are the source workbook’s categories. They are displayed in duration order and are not converted to a common instructional-time measure.')}
      <div class="chart-grid">
        <figure class="panel" aria-labelledby="duration-record-title"><div class="panel-header"><div><p class="panel-kicker">Offering-row volume</p><h3 id="duration-record-title">Program records by duration</h3></div><span class="source-badge">64 rows</span></div><p class="chart-summary">Each row is a recorded program offering in a term; repeated program names can appear in more than one term.</p>${renderHorizontalBars(data, 'duration-records', 'records', 'program records')}<figcaption>Five-day offerings occur most often by record count.</figcaption></figure>
        <figure class="panel" aria-labelledby="duration-trainee-title"><div class="panel-header"><div><p class="panel-kicker">Participation volume</p><h3 id="duration-trainee-title">Reported trainees by duration</h3></div><span class="source-badge">14,575 total</span></div><p class="chart-summary">This separate chart sums reported trainee counts, so it does not confuse participation volume with record frequency.</p>${renderHorizontalBars(data, 'duration-trainees', 'trainees', 'reported trainees')}<figcaption>Three-month records hold the greatest reported trainee volume.</figcaption></figure>
      </div>
      <div class="data-table-card"><div class="panel-header"><div><p class="panel-kicker">Duration summary</p><h3>All source categories</h3></div><span class="source-badge">6 categories</span></div><div class="table-wrap"><table><caption class="sr-only">Duration analysis summary</caption><thead><tr><th>Duration</th><th>Program records</th><th>Distinct recorded names</th><th>Reported trainees</th><th>Annual share</th><th>Available courses</th></tr></thead><tbody>${data.map((item) => `<tr><td><strong>${item.label}</strong><br><span dir="auto">${escapeHTML(item.duration)}</span></td><td>${item.records}</td><td>${item.programs}</td><td><strong>${format(item.trainees)}</strong></td><td>${percentage(item.trainees, totalTrainees)}</td><td>${format(item.courses)}${item.missing ? ' <span class="missing">(1 missing)</span>' : ''}</td></tr>`).join('')}</tbody></table></div></div>
      ${evidencePanels({ title: 'Record frequency and reported trainee volume answer different questions', text: `Five-day categories account for ${data.find((item) => item.label === 'Five days').records} program records, while three-month categories sum to ${format(data.find((item) => item.label === 'Three months').trainees)} trainee counts. The separate charts retain that distinction.` }, { title: 'What explains the duration mix?', text: 'The workbook does not document curriculum design, capacity, registration rules, or strategic intent. Review those sources before interpreting duration patterns.' })}`;
  }

  function renderParticipation(root) {
    const programs = groupPrograms();
    const top = programs.slice(0, 5);
    const other = programs.slice(5).reduce((sum, item) => sum + item.trainees, 0);
    const buckets = [...top, { name: 'Other recorded programs', displayName: 'Other recorded programs', trainees: other, share: other / totalTrainees, terms: null, courses: null, isAggregate: true }];
    root.innerHTML = `${panelHeading('05 · Participation', 'How is reported trainee volume distributed across programs?', 'Share means a program group’s reported trainee count divided by the total of 14,575. It is not attendance rate, completion rate, or engagement quality.')}
      <div class="participation-layout"><figure class="panel donut-panel" aria-labelledby="donut-title"><div class="panel-header"><div><p class="panel-kicker">Share of annual volume</p><h3 id="donut-title">Top 5 programs + other</h3></div><span class="source-badge">6 buckets</span></div><p class="chart-summary">The final bucket is the exact combined total of all remaining source program titles, keeping the donut readable.</p>${renderDonut(buckets, 'participation-donut')}<figcaption>Legend values and the table are available without relying on slice color.</figcaption></figure>
        <article class="panel participation-note"><p class="panel-kicker">Interpret carefully</p><h3>Participation is a volume signal</h3><p>The workbook has no learner identifier, so counts cannot be deduplicated across programs or terms. It also does not measure completion, satisfaction, outcomes, or effectiveness.</p><div class="participation-total"><span>Total reported trainees</span><strong>${format(totalTrainees)}</strong></div></article></div>
      <div class="data-table-card"><div class="panel-header"><div><p class="panel-kicker">Participation shares</p><h3>Donut calculation detail</h3></div><span class="source-badge">Sums to 100%</span></div><div class="table-wrap"><table><caption class="sr-only">Program participation share table</caption><thead><tr><th>Program bucket</th><th>Reported trainees</th><th>Annual share</th><th>Calculation</th></tr></thead><tbody>${buckets.map((item) => `<tr><td>${item.isAggregate ? '<span class="program-title"><strong>Other recorded programs</strong><span class="original-title"><b>Original Arabic Title:</b> Not applicable — derived aggregate bucket.</span></span>' : programTitleMarkup(item.name)}</td><td><strong>${format(item.trainees)}</strong></td><td>${percentage(item.trainees, totalTrainees)}</td><td>${format(item.trainees)} ÷ ${format(totalTrainees)}</td></tr>`).join('')}</tbody></table></div></div>
      ${evidencePanels({ title: 'One program accounts for more than two-thirds of reported trainee volume', text: `${escapeHTML(top[0].displayName)} contributes ${format(top[0].trainees)} of ${format(totalTrainees)} trainee counts (${percentage(top[0].trainees, totalTrainees)}). The concentration is a calculated fact, not an assessment of that program’s quality or demand.` }, { title: 'What should be investigated before resource decisions?', text: 'Compare capacity, enrolment, program availability, and participant-level records outside this workbook. They are needed to understand who participates and why the reported counts concentrate.' })}`;
  }

  function renderDonut(data, chart) {
    const radius = 82; const circumference = 2 * Math.PI * radius; let offset = 0;
    const slices = data.map((item, index) => { const dash = item.trainees / totalTrainees * circumference; const slice = { ...item, dash, offset }; offset += dash; return slice; });
    return `<div class="donut-layout"><div class="donut-wrap"><svg class="donut" viewBox="0 0 220 220" role="img" aria-label="Donut chart of reported trainee share by program"><circle class="donut-track" cx="110" cy="110" r="${radius}"/><g transform="rotate(-90 110 110)">${slices.map((item, index) => `<circle class="donut-slice" cx="110" cy="110" r="${radius}" style="--slice-color:${palette[index]};stroke-dasharray:${item.dash} ${circumference - item.dash};stroke-dashoffset:${-item.offset}"/>`).join('')}</g></svg><div class="donut-center"><strong>${format(totalTrainees)}</strong><span>reported trainees</span></div><div class="donut-hit-layer">${slices.map((item, index) => chartButton({ chart, index, label: `${item.displayName || item.name}, ${format(item.trainees)} reported trainees, ${percentage(item.trainees, totalTrainees)}`, style: `--hit-index:${index}` })).join('')}</div></div><ul class="chart-legend">${slices.map((item, index) => `<li><i style="--legend-color:${palette[index]}" aria-hidden="true"></i><span title="${escapeHTML(item.displayName || item.name)}">${escapeHTML(item.displayName || item.name)}</span><strong>${percentage(item.trainees, totalTrainees)}</strong></li>`).join('')}</ul></div>`;
  }

  function renderView(view) {
    const root = $(`#${view}`);
    if (!root || state.rendered.has(view)) return;
    const quality = getDataQuality();
    if (!quality.valid) {
      root.innerHTML = emptyState('The embedded workbook rows did not pass the data-integrity checks. Review the source and method note before using the dashboard.', '<a class="button" href="#source">View source &amp; method</a>');
      root.setAttribute('aria-busy', 'false');
      state.rendered.add(view);
      return;
    }
    const renderers = { overview: renderOverview, terms: renderTerms, programs: renderPrograms, duration: renderDuration, participation: renderParticipation };
    requestAnimationFrame(() => {
      renderers[view](root);
      root.setAttribute('aria-busy', 'false');
      state.rendered.add(view);
      bindViewInteractions(view, root);
    });
  }

  function filterRows() {
    const needle = normalizeSearch(state.search);
    return rows.filter((row) => {
      const searchableTitle = normalizeSearch(`${programDisplayName(row.name)} ${row.name}`);
      return (state.term === 'all' || row.term === state.term) && (state.duration === 'all' || row.duration === state.duration) && (!needle || searchableTitle.includes(needle));
    });
  }

  function renderProgramTable(root) {
    const body = $('#program-table-body', root);
    const status = $('#program-table-status', root);
    if (!body || !status) return;
    const results = filterRows();
    const filteredTrainees = results.reduce((sum, row) => sum + row.trainees, 0);
    body.innerHTML = results.length ? results.map((row) => `<tr><td><strong>${row.term}</strong></td><td>${programTitleMarkup(row.name)}</td><td>${durationNames[row.duration]}</td><td>${row.courses === null ? '<span class="missing">Missing</span>' : format(row.courses)}</td><td><strong>${format(row.trainees)}</strong></td><td>${row.location ? `<span lang="ar" dir="rtl">${escapeHTML(row.location)}</span>` : '<span class="missing">Not recorded</span>'}</td></tr>`).join('') : '<tr><td colspan="6"><div class="table-empty"><strong>No program records match these filters.</strong><span>Clear or adjust a filter to restore the source-derived records.</span></div></td></tr>';
    status.textContent = `${results.length} of ${rows.length} program records shown · ${format(filteredTrainees)} reported trainee counts`;
  }

  function setupProgramFilters(root) {
    const search = $('#program-search', root);
    const term = $('#term-filter', root);
    const duration = $('#duration-filter', root);
    const clear = $('#clear-filters', root);
    const redraw = () => renderProgramTable(root);
    search.addEventListener('input', (event) => { state.search = event.target.value; redraw(); });
    term.addEventListener('change', (event) => { state.term = event.target.value; redraw(); });
    duration.addEventListener('change', (event) => { state.duration = event.target.value; redraw(); });
    clear.addEventListener('click', () => { state.search = ''; state.term = 'all'; state.duration = 'all'; search.value = ''; term.value = 'all'; duration.value = 'all'; redraw(); search.focus(); });
    renderProgramTable(root);
  }

  function bindViewInteractions(view, root) {
    if (view === 'overview') {
      const data = totalsByTerm();
      bindChartTooltips(root, 'overview-trend', (index) => `<strong>${data[index].term} · ${format(data[index].trainees)} trainees</strong>${data[index].records} program records · ${percentage(data[index].trainees, totalTrainees)} of annual total`);
    }
    if (view === 'terms') {
      const data = totalsByTerm();
      bindChartTooltips(root, 'term-trainees', (index) => `<strong>${data[index].term} · ${format(data[index].trainees)} reported trainees</strong>${data[index].records} program records · ${format(data[index].courses)} available courses${data[index].missing ? ' (1 source cell missing)' : ''}`);
      bindChartTooltips(root, 'term-line', (index) => `<strong>${data[index].term} · ${format(data[index].trainees)} trainees</strong>${data[index].change === null ? 'First reported term' : `${data[index].change > 0 ? '+' : ''}${format(data[index].change)} (${data[index].changeRate > 0 ? '+' : ''}${formatValue(data[index].changeRate * 100, 1)}%) from prior term`}`);
    }
    if (view === 'programs') {
      const top = groupPrograms().slice(0, 8); const all = groupPrograms();
      bindChartTooltips(root, 'program-rank', (index) => { const item = top[index]; return `<strong>${escapeHTML(item.displayName)}</strong>${originalTitleTooltip(item.name)}${format(item.trainees)} reported trainees · ${percentage(item.trainees, totalTrainees)} of annual total<br>${format(item.courses)} available courses across ${item.terms} terms`; });
      bindChartTooltips(root, 'program-scatter', (index) => { const item = all[index]; const terms = SHEETS.map((sheet) => `${sheet.term}: ${format(item.byTerm[sheet.term] || 0)}`).join(' · '); return `<strong>${escapeHTML(item.displayName)}</strong>${originalTitleTooltip(item.name)}${format(item.courses)} available courses · ${format(item.trainees)} reported trainees<br>${terms}${item.courseMissing ? '<br>1 course-count source cell missing' : ''}`; });
      setupProgramFilters(root);
    }
    if (view === 'duration') {
      const data = durationStats();
      bindChartTooltips(root, 'duration-records', (index) => `<strong>${data[index].label}</strong>${data[index].records} program records · ${data[index].programs} distinct recorded names<br>${format(data[index].trainees)} reported trainees`);
      bindChartTooltips(root, 'duration-trainees', (index) => `<strong>${data[index].label}</strong>${format(data[index].trainees)} reported trainees · ${percentage(data[index].trainees, totalTrainees)} of annual total<br>${format(data[index].courses)} available courses${data[index].missing ? ' (1 source cell missing)' : ''}`);
    }
    if (view === 'participation') {
      const programs = groupPrograms(); const top = programs.slice(0, 5); const other = programs.slice(5).reduce((sum, item) => sum + item.trainees, 0); const data = [...top, { name: 'Other recorded programs', displayName: 'Other recorded programs', trainees: other, isAggregate: true }];
      bindChartTooltips(root, 'participation-donut', (index) => { const item = data[index]; return `<strong>${escapeHTML(item.displayName)}</strong>${item.isAggregate ? '<span class="original-title"><b>Original Arabic Title:</b> Not applicable — derived aggregate bucket.</span>' : originalTitleTooltip(item.name)}${format(item.trainees)} reported trainees · ${percentage(item.trainees, totalTrainees)} of annual total`; });
    }
  }

  function activateView(view, { updateHash = false, focus = false } = {}) {
    const next = viewOrder.includes(view) ? view : 'overview';
    hideTooltip();
    viewOrder.forEach((name) => {
      const panel = $(`#${name}`); const tab = $(`#tab-${name}`);
      const active = name === next;
      panel.hidden = !active;
      panel.classList.toggle('is-active', active);
      tab.setAttribute('aria-selected', String(active));
      tab.tabIndex = active ? 0 : -1;
    });
    state.activeView = next;
    renderView(next);
    if (updateHash && location.hash !== `#${next}`) history.pushState(null, '', `#${next}`);
    if (focus) requestAnimationFrame(() => $(`#${next}`).focus({ preventScroll: true }));
  }

  function setupTabs() {
    const tabs = $$('.dashboard-tab');
    tabs.forEach((tab, index) => {
      tab.addEventListener('click', () => activateView(tab.dataset.view, { updateHash: true, focus: true }));
      tab.addEventListener('keydown', (event) => {
        let next = null;
        if (event.key === 'ArrowRight') next = (index + 1) % tabs.length;
        if (event.key === 'ArrowLeft') next = (index - 1 + tabs.length) % tabs.length;
        if (event.key === 'Home') next = 0;
        if (event.key === 'End') next = tabs.length - 1;
        if (next !== null) { event.preventDefault(); tabs[next].focus(); return; }
        if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); activateView(tab.dataset.view, { updateHash: true, focus: true }); }
      });
    });
    window.addEventListener('hashchange', () => activateView(hashViews[location.hash] || 'overview'));
  }

  function setupTheme() {
    const button = $('#theme-toggle'); const label = $('#theme-label');
    const render = () => { const dark = document.documentElement.dataset.theme !== 'light'; button.setAttribute('aria-pressed', String(dark)); label.textContent = dark ? 'Switch to light theme' : 'Switch to dark theme'; button.setAttribute('aria-label', dark ? 'Switch to light theme' : 'Switch to dark theme'); };
    button.addEventListener('click', () => { document.documentElement.dataset.theme = document.documentElement.dataset.theme === 'light' ? 'dark' : 'light'; localStorage.setItem('sdaia-dashboard-theme', document.documentElement.dataset.theme); render(); });
    render();
  }

  function updateHero() {
    const terms = totalsByTerm(); const change = (terms[3].trainees - terms[0].trainees) / terms[0].trainees;
    $('#hero-start').textContent = format(terms[0].trainees);
    $('#hero-end').textContent = format(terms[3].trainees);
    $('#hero-summary').textContent = `That is a ${formatValue(Math.abs(change) * 100, 1)}% decrease between the first and fourth reported terms. The workbook shows the pattern, not its cause.`;
  }

  function validateData() {
    const terms = totalsByTerm(); const quality = getDataQuality();
    if (!quality.valid || terms.some((item) => item.trainees !== item.reported) || groupPrograms().length !== 24) console.warn('Dataset validation failed', { totalTrainees, totalCourses, records: rows.length, terms, quality });
  }

  document.addEventListener('keydown', (event) => { if (event.key === 'Escape') hideTooltip(); });
  updateHero();
  setupTheme();
  setupTabs();
  validateData();
  activateView(hashViews[location.hash] || 'overview');
})();
