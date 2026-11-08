
const STORE = 'jinan-manager-changes-v1';
const CONFIG_STORE = 'jinan-manager-config-v1';

const STATUSES = [
  'لم يُستلم',
  'تم الاستلام',
  'قيد المراجعة',
  'مُقفل',
  'متعثر'
];

const MONTHS = [
  'تموز','آب','أيلول','تشرين الأول',
  'تشرين الثاني','كانون الأول','كانون الثاني',
  'شباط','آذار','نيسان','أيار','حزيران'
];

const LABELS = {
  dashboard: 'لوحة القيادة',
  departments: 'الأقسام الإدارية',
  department: 'مهام القسم',
  calendar: 'الروزنامة السنوية',
  routine: 'الأعمال الثابتة',
  government: 'المعاملات الرسمية',
  settings: 'الإعدادات والربط'
};

const content = document.querySelector('#content');
const dialog = document.querySelector('#taskDialog');

const loadLocal = (key, fallback) => {
  try {
    return JSON.parse(localStorage.getItem(key)) || fallback;
  } catch {
    return fallback;
  }
};

const saveLocal = (key, value) =>
  localStorage.setItem(key, JSON.stringify(value));

const escapeHtml = value =>
  String(value ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[c]));

const e = escapeHtml;

let tasks = [];
let departments = [];

let changes = loadLocal(STORE, {});
let config = {
  clientId: '',
  calendarId: 'primary',
  sheetId: '',
  ...loadLocal(CONFIG_STORE, {})
};

let view = 'dashboard';
let department = '';
let month = '';
let tab = 'all';
let search = '';
let statusFilter = '';

let selectedTask = null;
let googleToken = '';
let tokenExpires = 0;
let googleScope = '';

const task = original => ({
  ...original,
  ...(changes[original.id] || {})
});

const byId = id => tasks.find(t => t.id === id);

const today = () => {
  const items = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Beirut',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(new Date());

  const get = name => items.find(p => p.type === name).value;

  return `${get('year')}-${get('month')}-${get('day')}`;
};

const completed = t => task(t).status === 'مُقفل';

const overdue = t =>
  Boolean(task(t).date && task(t).date < today() && !completed(t));

const upcoming = t => {
  const date = task(t).date;

  const days = date
    ? (Date.parse(date + 'T00:00:00Z') -
       Date.parse(today() + 'T00:00:00Z')) / 86400000
    : Infinity;

  return !completed(t) && days >= 0 && days <= 14;
};

const formatDate = date =>
  date
    ? new Date(date + 'T12:00:00Z').toLocaleDateString('ar-LB', {
        day: 'numeric',
        month: 'short',
        year: 'numeric',
        timeZone: 'UTC'
      })
    : 'غير محدد';

function alertBox(message) {
  const box = document.querySelector('#toast');
  box.textContent = message;
  box.style.display = 'block';
  setTimeout(() => box.style.display = 'none', 4500);
}

function stats(list) {
  return {
    total: list.length,
    done: list.filter(completed).length,
    late: list.filter(overdue).length,
    soon: list.filter(upcoming).length
  };
}

function kpis(list) {
  const n = stats(list);

  return `
    <div class="kpis">
      <div class="kpi">
        <span>إجمالي الأعمال</span>
        <strong>${n.total}</strong>
      </div>
      <div class="kpi">
        <span>الأعمال المقفلة</span>
        <strong>${n.done}</strong>
      </div>
      <div class="kpi">
        <span>الأعمال المتأخرة</span>
        <strong>${n.late}</strong>
      </div>
      <div class="kpi">
        <span>المستحقة خلال 14 يوماً</span>
        <strong>${n.soon}</strong>
      </div>
    </div>`;
}

function cards() {
  return `<div class="departments">${
    departments.map(d => {
      const s = stats(tasks.filter(t => t.department === d));
      const percent = s.total
        ? Math.round(s.done * 100 / s.total)
        : 0;

      return `
        <button class="dept" data-dept="${e(d)}">
          <strong>${e(d)}</strong>
          <small>
            ${s.total} عمل · ${s.late} متأخر · ${percent}% منجز
          </small>
          <span class="progress">
            <i style="width:${percent}%"></i>
          </span>
        </button>`;
    }).join('')
  }</div>`;
}

function table(list) {
  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>العمل المطلوب</th>
            <th>التصنيف</th>
            <th>الموعد</th>
            <th>الأولوية</th>
            <th>الحالة</th>
            <th>المتابعة</th>
          </tr>
        </thead>
        <tbody id="rows">${rows(list)}</tbody>
      </table>
    </div>
    <p class="muted" id="count">${list.length} مهمة</p>`;
}

function rows(list) {
  if (!list.length) {
    return '<tr><td colspan="6">لا توجد مهام مطابقة.</td></tr>';
  }

  return list.map(raw => {
    const t = task(raw);

    return `
      <tr>
        <td>
          <b>${e(t.title)}</b>
          <small>${e(t.id)} | ${e(t.department)}</small>
        </td>
        <td>${e(t.frequency)}</td>
        <td class="${overdue(raw) ? 'late' : ''}">
          ${e(t.date ? formatDate(t.date) :
              t.deadlineText || 'حسب الحاجة')}
        </td>
        <td>${e(t.priority || '—')}</td>
        <td>
          <span class="tag" data-status="${e(t.status)}">
            ${e(t.status)}
          </span>
        </td>
        <td>
          <button class="row-action" data-task="${e(t.id)}">
            فتح المهمة
          </button>
        </td>
      </tr>`;
  }).join('');
}

function toolbar() {
  return `
    <div class="filters">
      <input id="search"
        placeholder="ابحث عن تقرير أو مسؤول أو مهمة"
        value="${e(search)}">
      <select id="statusFilter">
        <option value="">جميع الحالات</option>
        ${STATUSES.map(x =>
          `<option ${statusFilter === x ? 'selected' : ''}>
            ${e(x)}
          </option>`
        ).join('')}
      </select>
    </div>`;
}

function filtered(list) {
  const q = search.trim().toLocaleLowerCase('ar');

  return list.filter(raw => {
    const t = task(raw);

    const match = [
      t.title,
      t.department,
      t.sourceDepartment,
      t.description,
      t.assignee
    ].join(' ').toLocaleLowerCase('ar').includes(q);

    const category =
      tab === 'all' ||
      (tab === 'seasonal' && t.frequency === 'موسمي') ||
      (tab === 'monthly' && t.frequency === 'شهري') ||
      (tab === 'routine' && t.kind === 'routine') ||
      (tab === 'government' && t.kind === 'government');

    return match &&
      (!statusFilter || t.status === statusFilter) &&
      category;
  }).sort((a, b) =>
    (task(a).date || '9999').localeCompare(task(b).date || '9999')
  );
}

function tabButtons() {
  const options = [
    ['all', 'جميع الأعمال'],
    ['seasonal', 'السنوية والموسمية'],
    ['monthly', 'الدورية والشهرية'],
    ['routine', 'الأعمال الثابتة'],
    ['government', 'المعاملات الرسمية']
  ];

  return `<div class="tabs">${
    options.map(([key, label]) =>
      `<button class="tab ${tab === key ? 'active' : ''}"
        data-tab="${key}">${label}</button>`
    ).join('')
  }</div>`;
}

function taskSection(list, withTabs = false) {
  return `
    ${withTabs ? tabButtons() : ''}
    ${toolbar()}
    <div class="panel" id="results">
      ${table(filtered(list))}
    </div>`;
}

function render() {
  document.querySelector('#pageTitle').textContent =
    view === 'department' ? department : LABELS[view];

  document.querySelectorAll('#menu button').forEach(b =>
    b.classList.toggle('active', b.dataset.view === view)
  );

  if (view === 'dashboard') {
    const urgent = tasks
      .filter(t => overdue(t) || upcoming(t))
      .sort((a, b) =>
        (task(a).date || '9999').localeCompare(task(b).date || '9999')
      )
      .slice(0, 15);

    content.innerHTML = `
      <section class="hero">
        <h2>لوحة متابعة المدير العام</h2>
        <p>الأعمال السنوية، التقارير الثابتة،
        المعاملات الرسمية وقرارات المتابعة.</p>
      </section>
      ${kpis(tasks)}
      <div class="section-title"><h2>الأقسام الإدارية</h2></div>
      ${cards()}
      <div class="section-title"><h2>المتابعات العاجلة</h2></div>
      <div class="panel">${table(urgent)}</div>`;

  } else if (view === 'departments') {
    content.innerHTML = `
      <section class="hero">
        <h2>الأقسام الإدارية</h2>
        <p>اختر القسم للاطلاع على أعماله بالتفصيل.</p>
      </section>
      ${cards()}`;

  } else if (view === 'department') {
    const list = tasks.filter(t => t.department === department);

    content.innerHTML = `
      <section class="hero">
        <h2>${e(department)}</h2>
        <p>تقارير القسم ومسؤولياته والمواعيد وقرارات المدير.</p>
      </section>
      ${kpis(list)}
      ${taskSection(list, true)}`;

  } else if (view === 'calendar') {
    const list = tasks.filter(t =>
      t.date && (!month || t.month === month)
    );

    content.innerHTML = `
      <div class="panel">
        <div class="actions">
          <button class="outline" data-action="ics">
            تصدير المواعيد إلى Google Calendar (ICS)
          </button>
        </div>
        <div class="months">
          <button data-month="" class="${!month ? 'active' : ''}">
            جميع الأشهر
          </button>
          ${MONTHS.map(m =>
            `<button data-month="${e(m)}"
              class="${month === m ? 'active' : ''}">
              ${e(m)}
            </button>`
          ).join('')}
        </div>
      </div>
      ${taskSection(list)}`;

  } else if (view === 'routine' || view === 'government') {
    const list = tasks.filter(t => t.kind === view);

    content.innerHTML = `
      <section class="hero">
        <h2>${LABELS[view]}</h2>
        <p>${view === 'government'
          ? 'تُراجع المواعيد الرسمية وفق أحدث تعميم للجهة المختصة.'
          : 'متابعات وتقارير ثابتة على مدار السنة.'}</p>
      </section>
      ${kpis(list)}
      ${taskSection(list)}`;

  } else if (view === 'settings') {
    renderSettings();
  }
}

function updateRows() {
  const result = document.querySelector('#results');
  if (!result) return;

  const base =
    view === 'department'
      ? tasks.filter(t => t.department === department)
      : view === 'calendar'
        ? tasks.filter(t => t.date && (!month || t.month === month))
        : tasks.filter(t => t.kind === view);

  result.innerHTML = table(filtered(base));
}

function openTask(id) {
  selectedTask = byId(id);
  if (!selectedTask) return;

  const t = task(selectedTask);

  document.querySelector('#taskTitle').textContent = t.title;

  const information = [
    ['القسم', t.department],
    ['التصنيف', t.frequency],
    ['الموعد المرجعي', t.deadlineText || t.date || 'حسب الحاجة'],
    ['المطلوب في التقرير', t.description],
    ['الأدلة أو المؤشرات', t.evidence],
    ['دور المدير', t.managerAction],
    ['المصدر', t.reference]
  ];

  document.querySelector('#taskInfo').innerHTML =
    information.map(([a, b]) =>
      `<div><b>${a}:</b> ${e(b || '—')}</div>`
    ).join('');

  const fields = [
    ['taskStatus', 'status'],
    ['taskAssignee', 'assignee'],
    ['taskDate', 'date'],
    ['taskReceived', 'receivedAt'],
    ['taskDecision', 'decision'],
    ['taskNotes', 'notes'],
    ['taskProof', 'proofUrl']
  ];

  for (const [id, key] of fields) {
    document.querySelector('#' + id).value = t[key] || '';
  }

  dialog.showModal();
}

function saveTask() {
  if (!selectedTask) return false;

  const id = selectedTask.id;
  const old = changes[id] || {};

  const updated = {
    ...old,
    status: document.querySelector('#taskStatus').value,
    assignee: document.querySelector('#taskAssignee').value.trim(),
    date: document.querySelector('#taskDate').value,
    receivedAt: document.querySelector('#taskReceived').value,
    decision: document.querySelector('#taskDecision').value.trim(),
    notes: document.querySelector('#taskNotes').value.trim(),
    proofUrl: document.querySelector('#taskProof').value.trim(),
    updatedAt: new Date().toISOString()
  };

  if (updated.proofUrl &&
      !/^https:\/\//i.test(updated.proofUrl)) {
    alertBox('رابط المستند يجب أن يبدأ بـ https://');
    return false;
  }

  changes[id] = updated;
  saveLocal(STORE, changes);
  render();

  return true;
}

function calendarDraft() {
  if (!selectedTask) return;
  if (!saveTask()) return;

  const t = task(selectedTask);
  if (!t.date) return alertBox('حدّد تاريخ المهمة أولاً');

  const start = t.date.replaceAll('-', '');
  const next = new Date(
    Date.parse(t.date + 'T00:00:00Z') + 86400000
  ).toISOString().slice(0, 10).replaceAll('-', '');

  const params = new URLSearchParams({
    action: 'TEMPLATE',
    text: t.title,
    dates: `${start}/${next}`,
    details: `القسم: ${t.department}
المسؤول: ${t.assignee || t.sourceDepartment}
دور المدير: ${t.managerAction}`
  });

  window.open(
    'https://calendar.google.com/calendar/render?' + params,
    '_blank',
    'noopener,noreferrer'
  );
}

function download(name, text, mime) {
  const link = document.createElement('a');
  const url = URL.createObjectURL(new Blob([text], {type: mime}));

  link.href = url;
  link.download = name;
  link.click();

  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function backup() {
  download(
    'jinan-followups-backup.json',
    JSON.stringify({
      changes,
      config,
      exportedAt: new Date().toISOString()
    }, null, 2),
    'application/json'
  );
}

function exportIcs() {
  let out = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Al Jinan High School//Management//AR',
    'CALSCALE:GREGORIAN'
  ];

  for (const original of tasks) {
    const t = task(original);
    if (!t.date) continue;

    const start = t.date.replaceAll('-', '');
    const end = new Date(
      Date.parse(t.date + 'T00:00:00Z') + 86400000
    ).toISOString().slice(0, 10).replaceAll('-', '');

    const clean = v => String(v || '')
      .replaceAll('\\', '\\\\')
      .replaceAll(',', '\\,')
      .replaceAll(';', '\\;')
      .replaceAll('\n', '\\n');

    out.push(
      'BEGIN:VEVENT',
      `UID:${t.id}@jinan-management`,
      `DTSTART;VALUE=DATE:${start}`,
      `DTEND;VALUE=DATE:${end}`,
      `SUMMARY:${clean(t.title)}`,
      `DESCRIPTION:${clean(t.department + ' - ' + t.managerAction)}`,
      'END:VEVENT'
    );
  }

  out.push('END:VCALENDAR');

  download(
    'jinan-management.ics',
    out.join('\r\n'),
    'text/calendar;charset=utf-8'
  );
}

function renderSettings() {
  content.innerHTML = `
    <div class="settings">
      <div class="panel">
        <h2>ربط Google Calendar</h2>

        <label>OAuth Client ID
          <input id="clientId" value="${e(config.clientId)}"
            placeholder="...apps.googleusercontent.com">
        </label>

        <label>معرّف التقويم
          <input id="calendarId" value="${e(config.calendarId)}">
        </label>

        <div class="actions">
          <button class="primary" data-action="saveConfig">
            حفظ الإعدادات
          </button>
          <button class="outline" data-action="connect">
            ربط حساب Google
          </button>
        </div>

        <div class="note">
          يلزم تفعيل Google Calendar API وإنشاء OAuth Client ID
          من نوع Web application وإضافة رابط GitHub Pages
          إلى Authorized JavaScript origins.
          لا تضع Client Secret في الموقع.
        </div>
      </div>

      <div class="panel">
        <h2>ربط Google Sheets</h2>

        <label>Spreadsheet ID
          <input id="sheetId" value="${e(config.sheetId)}">
        </label>

        <div class="actions">
          <button class="outline" data-action="push">
            رفع المتابعات إلى Sheets
          </button>
          <button class="outline" data-action="pull">
            استيراد المتابعات من Sheets
          </button>
        </div>

        <p class="muted">
          هذه مزامنة يدوية عند الضغط، وليست فورية.
        </p>
      </div>
    </div>

    <div class="panel">
      <h2>النسخة الاحتياطية والتقويم</h2>

      <div class="actions">
        <button class="outline" data-action="backup">
          تنزيل نسخة احتياطية
        </button>
        <button class="outline" data-action="restore">
          استيراد نسخة احتياطية
        </button>
        <button class="outline" data-action="ics">
          ملف التقويم ICS
        </button>
        <input type="file" id="backupFile"
          accept=".json,application/json" hidden>
      </div>

      <div class="note">
        GitHub Pages موقع عام عادةً. لا ترفع بيانات شخصية
        أو ملفات داخلية حساسة من دون نظام صلاحيات حقيقي.
      </div>
    </div>`;
}

function saveConfig() {
  config = {
    clientId: document.querySelector('#clientId').value.trim(),
    calendarId:
      document.querySelector('#calendarId').value.trim() || 'primary',
    sheetId: document.querySelector('#sheetId').value.trim()
  };

  saveLocal(CONFIG_STORE, config);
  googleToken = '';
  alertBox('تم حفظ إعدادات Google');
}

function getScopes() {
  return [
    'https://www.googleapis.com/auth/calendar.events',
    ...(config.sheetId
      ? ['https://www.googleapis.com/auth/spreadsheets']
      : [])
  ].join(' ');
}

async function getGoogleToken() {
  if (!config.clientId) {
    throw new Error('أدخل OAuth Client ID من صفحة الإعدادات');
  }

  if (!window.google?.accounts?.oauth2) {
    throw new Error('مكتبة Google غير متاحة');
  }

  const scope = getScopes();

  if (googleToken &&
      Date.now() < tokenExpires - 60000 &&
      scope === googleScope) {
    return googleToken;
  }

  return new Promise((resolve, reject) => {
    const client = google.accounts.oauth2.initTokenClient({
      client_id: config.clientId,
      scope,

      callback: result => {
        if (result.error) return reject(new Error(result.error));

        googleToken = result.access_token;
        tokenExpires =
          Date.now() + Number(result.expires_in || 3600) * 1000;
        googleScope = scope;

        resolve(googleToken);
      },

      error_callback: x =>
        reject(new Error(x.message || 'تعذّر الاتصال بحساب Google'))
    });

    client.requestAccessToken();
  });
}

async function googleApi(url, method = 'GET', body) {
  const token = await getGoogleToken();

  const response = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(body ? {'Content-Type': 'application/json'} : {})
    },
    ...(body ? {body: JSON.stringify(body)} : {})
  });

  const data = await response.json().catch(() => ({}));

  if (!response.ok) {
    throw new Error(
      data.error?.message || 'فشل طلب Google API'
    );
  }

  return data;
}

async function syncCalendar() {
  if (!selectedTask) return;
  if (!saveTask()) return;

  const t = task(selectedTask);

  if (!t.date) return alertBox('حدّد موعد المهمة أولاً');

  const base =
    `https://www.googleapis.com/calendar/v3/calendars/` +
    `${encodeURIComponent(config.calendarId || 'primary')}/events`;

  const end = new Date(
    Date.parse(t.date + 'T00:00:00Z') + 86400000
  ).toISOString().slice(0, 10);

  const body = {
    summary: t.title,
    description:
      `القسم: ${t.department}\n` +
      `المسؤول: ${t.assignee || t.sourceDepartment}\n` +
      `دور المدير: ${t.managerAction}\n` +
      `رقم المهمة: ${t.id}`,
    start: {date: t.date},
    end: {date: end},
    reminders: {
      useDefault: false,
      overrides: [{method: 'popup', minutes: 1440}]
    }
  };

  try {
    let event;
    const old = changes[t.id]?.calendarEventId;

    if (old) {
      try {
        event = await googleApi(
          `${base}/${encodeURIComponent(old)}`, 'PUT', body
        );
      } catch (err) {
        if (!/not found|404/i.test(err.message)) throw err;
        event = await googleApi(base, 'POST', body);
      }
    } else {
      event = await googleApi(base, 'POST', body);
    }

    changes[t.id] = {
      ...changes[t.id],
      calendarEventId: event.id,
      updatedAt: new Date().toISOString()
    };

    saveLocal(STORE, changes);
    alertBox('تم إنشاء/تحديث الموعد في Google Calendar');

    if (event.htmlLink) {
      window.open(event.htmlLink, '_blank', 'noopener,noreferrer');
    }
  } catch (err) {
    alertBox(err.message);
  }
}

function sheetBase() {
  if (!config.sheetId) {
    throw new Error('أدخل Spreadsheet ID أولاً');
  }

  return `https://sheets.googleapis.com/v4/spreadsheets/${
    encodeURIComponent(config.sheetId)
  }`;
}

async function ensureSheet() {
  const base = sheetBase();
  const meta = await googleApi(
    base + '?fields=sheets.properties.title'
  );

  if (!meta.sheets?.some(s =>
    s.properties?.title === 'Jinan_Status'
  )) {
    await googleApi(base + ':batchUpdate', 'POST', {
      requests: [
        {addSheet: {properties: {title: 'Jinan_Status'}}}
      ]
    });
  }
}

async function remoteChanges() {
  const base = sheetBase();
  let result;

  try {
    result = await googleApi(
      `${base}/values/${
        encodeURIComponent("'Jinan_Status'!A2:J10000")
      }`
    );
  } catch (err) {
    if (/Unable to parse range|not found/i.test(err.message)) {
      return {};
    }
    throw err;
  }

  const valid = new Set(tasks.map(t => t.id));
  const updates = {};

  for (const row of result.values || []) {
    const [
      id, status, assignee, date, receivedAt,
      decision, notes, proofUrl, calendarEventId, updatedAt
    ] = row;

    if (valid.has(id)) {
      updates[id] = {
        status, assignee, date, receivedAt,
        decision, notes, proofUrl, calendarEventId, updatedAt
      };
    }
  }

  return updates;
}

function mergeChanges(local, remote) {
  const merged = {...local};

  for (const [id, value] of Object.entries(remote)) {
    if (!merged[id] ||
        String(value.updatedAt || '') >
        String(merged[id].updatedAt || '')) {
      merged[id] = value;
    }
  }

  return merged;
}

async function pullSheet() {
  try {
    changes = mergeChanges(changes, await remoteChanges());
    saveLocal(STORE, changes);
    render();
    alertBox('تم استيراد المتابعات');
  } catch (err) {
    alertBox(err.message);
  }
}

async function pushSheet() {
  try {
    await ensureSheet();
    changes = mergeChanges(changes, await remoteChanges());

    const columns = [
      'id','status','assignee','date','receivedAt',
      'decision','notes','proofUrl','calendarEventId','updatedAt'
    ];

    const ids = new Set(tasks.map(t => t.id));

    const values = [
      columns,
      ...Object.entries(changes)
        .filter(([id]) => ids.has(id))
        .map(([id, x]) => [
          id,
          ...columns.slice(1).map(key => x[key] || '')
        ])
    ];

    const range = `'Jinan_Status'!A1:J${values.length}`;

    await googleApi(
      `${sheetBase()}/values/${
        encodeURIComponent(range)
      }?valueInputOption=RAW`,
      'PUT',
      {range, majorDimension: 'ROWS', values}
    );

    saveLocal(STORE, changes);
    alertBox('تم رفع المتابعات إلى Google Sheets');
  } catch (err) {
    alertBox(err.message);
  }
}

async function restore(file) {
  if (!file) return;

  try {
    const data = JSON.parse(await file.text());
    const valid = new Set(tasks.map(t => t.id));

    if (!data.changes || typeof data.changes !== 'object') {
      throw new Error('ملف النسخة الاحتياطية غير صالح');
    }

    changes = Object.fromEntries(
      Object.entries(data.changes)
        .filter(([id]) => valid.has(id))
    );

    saveLocal(STORE, changes);
    render();
    alertBox('تم استيراد النسخة الاحتياطية');
  } catch (err) {
    alertBox(err.message);
  }
}

function navigate(next) {
  view = next;
  search = '';
  statusFilter = '';
  tab = 'all';
  render();
}

document.addEventListener('click', async event => {
  const button = event.target.closest('button');
  if (!button) return;

  if (button.dataset.view) {
    navigate(button.dataset.view);
    return;
  }

  if (button.dataset.dept) {
    department = button.dataset.dept;
    navigate('department');
    return;
  }

  if (button.dataset.task) {
    openTask(button.dataset.task);
    return;
  }

  if (button.dataset.tab) {
    tab = button.dataset.tab;
    render();
    return;
  }

  if (button.dataset.month !== undefined) {
    month = button.dataset.month;
    render();
    return;
  }

  switch (button.dataset.action) {
    case 'ics':
      exportIcs();
      break;

    case 'backup':
      backup();
      break;

    case 'restore':
      document.querySelector('#backupFile').click();
      break;

    case 'saveConfig':
      saveConfig();
      break;

    case 'connect':
      try {
        await getGoogleToken();
        alertBox('تم الاتصال بحساب Google');
      } catch (err) {
        alertBox(err.message);
      }
      break;

    case 'pull':
      await pullSheet();
      break;

    case 'push':
      await pushSheet();
      break;
  }
});

document.addEventListener('input', event => {
  if (event.target.id === 'search') {
    search = event.target.value;
    updateRows();
  }
});

document.addEventListener('change', event => {
  if (event.target.id === 'statusFilter') {
    statusFilter = event.target.value;
    updateRows();
  }

  if (event.target.id === 'backupFile') {
    restore(event.target.files[0]);
  }
});

document.querySelector('#exportBtn')
  .addEventListener('click', backup);

document.querySelector('#closeDialog')
  .addEventListener('click', () => dialog.close());

document.querySelector('#taskForm')
  .addEventListener('submit', event => {
    event.preventDefault();

    if (saveTask()) {
      dialog.close();
      alertBox('تم حفظ المتابعة');
    }
  });

document.querySelector('#calendarDraft')
  .addEventListener('click', calendarDraft);

document.querySelector('#calendarSync')
  .addEventListener('click', syncCalendar);

// قراءة الأعمال من ملف البيانات
(async () => {
  try {
    const response = await fetch('./data/tasks.json', {
      cache: 'no-store'
    });

    if (!response.ok) {
      throw new Error('تعذّر الوصول إلى ملف البيانات');
    }

    const result = await response.json();

    tasks = result.tasks;
    departments = result.departments;

    if (!Array.isArray(tasks) ||
        !Array.isArray(departments)) {
      throw new Error('صيغة البيانات غير صحيحة');
    }

    render();

  } catch (err) {
    content.innerHTML = `
      <div class="note">
        ${e(err.message)}.
        افتح الموقع عبر GitHub Pages وليس بالنقر المزدوج
        على index.html.
      </div>`;
  }
})();
