
'use strict';

/* ===== الإعدادات العامة ===== */

const $ = id => document.getElementById(id);

const e = value => String(value ?? '').replace(
  /[&<>"']/g,
  x => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[x])
);

const STORE = 'jinan-manager-changes-word-v2';
const CONFIG_STORE = 'jinan-manager-config-v1';
const EXTRA_STORE = 'jinan-extra-tasks-v3';
const JOB_STORE = 'jinan-job-descriptions-v3';
const PENDING_STORE = 'jinan-pending-cloud-v3';

const SHEET_NAME = 'Jinan_Updates';

const SCOPE = {
  calendar: 'https://www.googleapis.com/auth/calendar.events',
  sheets: 'https://www.googleapis.com/auth/spreadsheets',
  gmail: 'https://www.googleapis.com/auth/gmail.send'
};

const STATUSES = [
  'لم يُستلم',
  'تم الاستلام',
  'قيد المراجعة',
  'مُقفل',
  'متعثر'
];

const MONTHS = [
  'تموز', 'آب', 'أيلول', 'تشرين الأول',
  'تشرين الثاني', 'كانون الأول',
  'كانون الثاني', 'شباط', 'آذار',
  'نيسان', 'أيار', 'حزيران'
];

const TITLES = {
  dashboard: 'لوحة القيادة',
  organization: 'الهيكلية الإدارية',
  departments: 'الأقسام الإدارية',
  department: 'أعمال القسم',
  jobs: 'الأوصاف الوظيفية',
  calendar: 'الروزنامة السنوية',
  routine: 'الأعمال الثابتة',
  government: 'المعاملات الرسمية',
  settings: 'الإعدادات والربط مع Google'
};

const KIND_LABEL = {
  plan: 'عمل سنوي',
  routine: 'تقرير ثابت',
  oversight: 'متابعة ورقابة',
  calendar: 'روزنامة',
  government: 'معاملة رسمية'
};

/* ===== الحفظ المحلي ===== */

function loadLocal(key, fallback) {
  try {
    const value = localStorage.getItem(key);
    return value ? JSON.parse(value) : fallback;
  } catch (error) {
    console.warn('التخزين المحلي غير متاح', error);
    return fallback;
  }
}

function saveLocal(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
    return true;
  } catch (error) {
    console.error('تعذّر التخزين المحلي', error);
    return false;
  }
}

function stamp() {
  return new Date().toISOString();
}

function stampView(value) {
  if (!value) return 'لم يُسجل بعد';

  if (/^\d{4}-\d{2}-\d{2}$/.test(String(value))) {
    return `${value} (من النسخة السابقة؛ الوقت غير مسجل)`;
  }

  const d = new Date(value);

  if (Number.isNaN(d.getTime())) return String(value);

  return d.toLocaleString('ar-LB', {
    timeZone: 'Asia/Beirut',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  });
}

function today() {
  const formatter = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Beirut',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  });

  const parts = formatter.formatToParts(new Date());

  const get = key => parts.find(p => p.type === key).value;

  return `${get('year')}-${get('month')}-${get('day')}`;
}

function dateView(date) {
  return date
    ? new Date(date + 'T12:00:00Z').toLocaleDateString(
        'ar-LB',
        {
          day: 'numeric',
          month: 'short',
          year: 'numeric',
          timeZone: 'UTC'
        }
      )
    : 'غير محدد';
}

function safeUrl(value) {
  return /^https:\/\/[^\s]+$/i.test(String(value || ''));
}

function validEmail(value) {
  return /^[^\s@\r\n]+@[^\s@\r\n]+\.[^\s@\r\n]+$/.test(
    String(value || '')
  );
}

function notify(message, warning = false) {
  const box = $('toast');

  box.textContent = message;
  box.style.display = 'block';
  box.style.background = warning ? '#80501b' : '#143c51';

  clearTimeout(notify.timer);

  notify.timer = setTimeout(() => {
    box.style.display = 'none';
  }, 5700);
}

function indicator(message, mode = '') {
  const box = $('syncIndicator');

  box.textContent = message;
  box.className = 'sync-indicator ' + mode;
}

/* ===== ألوان الأقسام ===== */

function deptTheme(name = '') {
  if (
    name.includes('الإشراف الإداري') ||
    name.includes('الاشراف الاداري')
  ) {
    return 'administrative';
  }

  if (
    name.includes('شؤون المتعلمين') ||
    name.includes('شؤون الطلاب')
  ) {
    return 'students';
  }

  if (name.includes('الإشراف التربوي')) {
    return 'pedagogical';
  }

  if (name.includes('الموارد البشرية')) {
    return 'hr';
  }

  if (
    name.includes('المالية') ||
    name.includes('المحاسبة')
  ) {
    return 'finance';
  }

  if (name.includes('الجودة')) {
    return 'quality';
  }

  return 'neutral';
}

/* ===== بيانات النظام ===== */

let baseTasks = [];
let tasks = [];
let departments = [];

let changes = loadLocal(STORE, {});
let customTasks = loadLocal(EXTRA_STORE, {});
let jobs = loadLocal(JOB_STORE, {});
let pending = loadLocal(PENDING_STORE, []);

let config = {
  clientId: '',
  calendarId: 'primary',
  sheetId: '',
  editorName: '',
  ...loadLocal(CONFIG_STORE, {})
};

let view = 'dashboard';
let department = '';
let month = '';
let tab = 'all';
let search = '';
let statusFilter = '';
let jobSearch = '';
let jobDepartment = '';

let selectedTask = null;
let editingTaskId = null;
let editingJobId = null;

let oauthToken = '';
let tokenExpires = 0;
let oauthGranted = new Set();
let syncBusy = false;

/* ===== معالجة الأعمال ===== */

function buildTasks() {
  const extra = Object.values(customTasks).filter(
    item => !item.deleted
  );

  tasks = [...baseTasks, ...extra];
}

function task(raw) {
  return {
    ...raw,
    ...(changes[raw.id] || {})
  };
}

function byId(id) {
  return tasks.find(t => t.id === id);
}

function isDone(raw) {
  return task(raw).status === 'مُقفل';
}

function isLate(raw) {
  const t = task(raw);

  return Boolean(
    t.date &&
    t.date < today() &&
    !isDone(raw)
  );
}

function isSoon(raw) {
  const t = task(raw);

  return (
    !isDone(raw) &&
    t.date >= today() &&
    (
      Date.parse(t.date + 'T00:00:00Z') -
      Date.parse(today() + 'T00:00:00Z')
    ) / 86400000 <= 14
  );
}

function stats(list) {
  return {
    total: list.length,
    done: list.filter(isDone).length,
    late: list.filter(isLate).length,
    soon: list.filter(isSoon).length
  };
}

function kpis(list) {
  const s = stats(list);

  const values = [
    ['الأعمال', s.total],
    ['المقفلة', s.done],
    ['المتأخرة (المؤرخة)', s.late],
    ['خلال 14 يوماً', s.soon]
  ];

  return `
    <div class="kpis">
      ${values.map(([label, value]) => `
        <div class="kpi">
          <span>${label}</span>
          <strong>${value}</strong>
        </div>
      `).join('')}
    </div>
  `;
}

function cards() {
  return `
    <div class="departments">
      ${departments.map(d => {
        const s = stats(
          tasks.filter(t => t.department === d)
        );

        const percent = s.total
          ? Math.round(100 * s.done / s.total)
          : 0;

        return `
          <button class="dept"
                  data-theme="${deptTheme(d)}"
                  data-dept="${e(d)}">

            <strong>${e(d)}</strong>

            <small>
              ${s.total} بند ·
              ${s.late} متأخر ·
              ${percent}% منجز
            </small>

            <span class="progress">
              <i style="width:${percent}%"></i>
            </span>
          </button>
        `;
      }).join('')}
    </div>
  `;
}

function linkButton(url, label = 'فتح التقرير') {
  return safeUrl(url)
    ? `<a class="open-doc"
          href="${e(url)}"
          target="_blank"
          rel="noopener noreferrer">${e(label)}</a>`
    : '';
}

function rows(list) {
  if (!list.length) {
    return '<tr><td colspan="7">لا توجد بنود مطابقة.</td></tr>';
  }

  return list.map(raw => {
    const t = task(raw);
    const delivered = t.proofUrl && t.receivedAt;

    const report = delivered
      ? `
        <span class="receipt">
          ${e(stampView(t.receivedAt))}
        </span>
        <br>
        ${linkButton(t.proofUrl)}
      `
      : t.proofUrl
        ? linkButton(t.proofUrl)
        : '<span class="muted">لم يصل رابط</span>';

    return `
      <tr>
        <td>
          <b>${e(t.title)}</b>
          <small>
            ${e(t.section || KIND_LABEL[t.kind] || 'عمل')}
            | ${e(t.id)}
          </small>
        </td>

        <td>${e(t.frequency || KIND_LABEL[t.kind] || '—')}</td>

        <td class="${isLate(raw) ? 'late' : ''}">
          ${e(
            t.deadlineText ||
            (t.date ? dateView(t.date) : 'تُحدّد')
          )}
        </td>

        <td>
          <span class="tag"
                data-status="${e(t.status || 'لم يُستلم')}">
            ${e(t.status || 'لم يُستلم')}
          </span>
        </td>

        <td>${report}</td>

        <td>
          ${e(t.assignee || t.sourceAssignee || '—')}
        </td>

        <td>
          <div class="quick-actions">
            <button class="row-action"
                    data-task="${e(t.id)}"
                    type="button">
              فتح
            </button>

            <button
              class="row-action row-reminder ${t.calendarEventId ? 'ready' : ''}"
              data-remind="${e(t.id)}"
              type="button">
              🔔 ${t.calendarEventId ? 'تحديث التذكير' : 'ذكّرني'}
            </button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

function table(list) {
  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>العمل</th>
            <th>النوع</th>
            <th>الموعد</th>
            <th>الحالة</th>
            <th>التقرير ووقت الوصول</th>
            <th>المسؤول</th>
            <th>إجراءات</th>
          </tr>
        </thead>
        <tbody>
          ${rows(list)}
        </tbody>
      </table>
    </div>
    <p class="muted">${list.length} بند</p>
  `;
}

function toolbar() {
  return `
    <div class="filters">
      <input id="search"
             placeholder="بحث عن تقرير أو عمل أو مسؤول"
             value="${e(search)}">

      <select id="statusFilter">
        <option value="">كل الحالات</option>
        ${STATUSES.map(x => `
          <option value="${e(x)}"
                  ${statusFilter === x ? 'selected' : ''}>
            ${e(x)}
          </option>
        `).join('')}
      </select>
    </div>
  `;
}

function filtered(list) {
  const q = search.trim().toLocaleLowerCase('ar');

  return list.filter(raw => {
    const t = task(raw);

    const text = [
      t.title,
      t.department,
      t.section,
      t.deadlineText,
      t.evidence,
      t.assignee,
      t.sourceAssignee,
      t.managerAction,
      t.followUp,
      t.description
    ].join(' ').toLocaleLowerCase('ar');

    return (
      text.includes(q) &&
      (!statusFilter || t.status === statusFilter) &&
      (tab === 'all' || t.kind === tab)
    );
  });
}

function taskArea(list, withTabs = false) {
  const types = [
    ['all', 'كل الأعمال'],
    ['plan', 'الأعمال السنوية'],
    ['routine', 'التقارير الثابتة'],
    ['oversight', 'الرقابة والمتابعة'],
    ['calendar', 'الروزنامة'],
    ['government', 'المعاملات الرسمية']
  ];

  return `
    ${withTabs ? `
      <div class="tabs">
        ${types.map(([key, label]) => `
          <button
            class="tab ${tab === key ? 'active' : ''}"
            type="button"
            data-tab="${key}">
            ${label}
          </button>
        `).join('')}
      </div>
    ` : ''}

    ${toolbar()}

    <div class="panel" id="results">
      ${table(filtered(list))}
    </div>
  `;
}

function mainList() {
  if (view === 'department') {
    return tasks.filter(t => t.department === department);
  }

  if (view === 'calendar') {
    return tasks.filter(
      t => t.kind === 'calendar' &&
      (!month || t.month === month)
    );
  }

  return tasks.filter(t => t.kind === view);
}

function updateRows() {
  const box = $('results');

  if (box) {
    box.innerHTML = table(filtered(mainList()));
  }
}

/* =====================================
   الهيكلية الإدارية الرسمية
   ===================================== */

const OFFICIAL_ORG = [
  {
    title: 'الإشراف الإداري والجودة',
    person: 'أ. سامي حلاق',
    group: 'الإشراف الإداري والعمليات والخدمات المدرسية',
    children: [
      ['قسم الإعلام والتسويق', 'أ. سمية النجار'],
      ['مسؤول مكننة ومعلوماتية', ''],
      ['ممرض', ''],
      ['المشرف الصحي', ''],
      ['مسؤول مخزون', ''],
      ['مسؤول مشتريات', ''],
      ['مسؤول الصيانة', ''],
      ['حارس', ''],
      ['سائق', '']
    ]
  },
  {
    title: 'نائب المدير',
    person: 'د. يكانة يكن',
    group: 'الإدارة العامة والقيادة المدرسية',
    children: []
  },
  {
    title: 'شؤون الطلاب',
    person: 'أ. مصطفى كمون',
    group: 'شؤون المتعلمين',
    children: []
  },
  {
    title: 'المدقق المالي',
    person: '',
    group: 'الشؤون المالية والمحاسبة',
    children: []
  },
  {
    title: 'المدير المالي',
    person: 'د. يكانة يكن',
    group: 'الشؤون المالية والمحاسبة',
    children: [
      ['محاسب', 'أ. فاطمة أبو زيد'],
      ['أمين صندوق', 'أ. هنادي لاذقاني']
    ]
  },
  {
    title: 'أمين السر',
    person: '',
    group: 'أمانة السر والمراسلات والأرشفة',
    children: []
  },
  {
    title: 'موارد بشرية',
    person: 'أ. صبا الخضر',
    group: 'إدارة الموارد البشرية',
    children: []
  },
  {
    title: 'الإشراف التربوي',
    person: 'أ. محمد صهيوني',
    group: 'الإشراف التربوي والأكاديمي',
    children: [
      ['مسؤولة الأنشطة', 'أ. رحمة كبارة'],
      ['عامل صحي اجتماعي', 'أ. هبة شعاع'],
      ['مسؤول مختبر العلوم', 'أ. ريما حمود'],
      ['مسؤول مكتبة', ''],
      ['مسؤول مختبر الكمبيوتر', 'أ. غنى درويش'],
      ['إشراف المرحلة الثانوية', 'أ. نهى الزهر'],
      ['إشراف الحلقتين الثالثة والرابعة', 'أ. زينة ولي الدين'],
      ['إشراف الحلقتين الثانية والأولى', 'أ. بشرى الزيلع'],
      ['إشراف مرحلة الروضات', 'أ. ماجدة حداد'],
      ['إشراف قسم الدعم التربوي', 'أ. رفاه حيدر']
    ]
  }
];

function organization() {
  function node(title, person, group) {
    const saved = Object.values(jobs).find(
      j => !j.deleted && j.title === title
    );

    const action = saved
      ? `data-job="${e(saved.id)}"`
      : `data-add-org-role="${e(title)}"
         data-add-org-dept="${e(group)}"`;

    return `
      <div class="org-position-row">
        <div>
          <strong>${e(title)}</strong>
          ${person ? `<small>(${e(person)})</small>` : ''}
        </div>

        <button type="button"
                class="row-action"
                ${action}>
          ${saved ? 'عرض الوصف' : 'إضافة وصف'}
        </button>
      </div>
    `;
  }

  return `
    <section class="hero">
      <h2>الهيكلية الإدارية المعتمدة</h2>
      <p>
        العام الدراسي 2026–2027،
        وفق مخطط الهيكلية الإدارية لثانوية الجنان.
      </p>
    </section>

    <div class="organization-root">
      <b>المدير العام</b>
      <small>د. سالم يكن</small>
    </div>

    <div class="org-connector"></div>

    <div class="org-grid">
      ${OFFICIAL_ORG.map(group => `
        <article class="org-card"
                 data-theme="${deptTheme(group.group)}">

          <header>
            <h3>${e(group.title)}</h3>
            ${group.person ? `<small>${e(group.person)}</small>` : ''}
          </header>

          <div class="org-roles">
            ${node(group.title, group.person, group.group)}

            ${group.children.map(([name, person]) =>
              node(name, person, group.group)
            ).join('')}
          </div>
        </article>
      `).join('')}
    </div>

    <p class="note">
      المسميات والأسماء مأخوذة من مخطط الهيكلية.
      إضافة وصف وظيفي لا تعني اعتماد مهام
      أو صلاحيات لم تُقرّها المؤسسة.
    </p>
  `;
}

/* ===== الأوصاف الوظيفية ===== */

function jobsPage() {
  const available = Object.values(jobs).filter(j => {
    if (j.deleted) return false;

    if (
      jobDepartment &&
      j.department !== jobDepartment
    ) {
      return false;
    }

    const text = [
      j.title,
      j.department,
      j.purpose,
      j.responsibilities,
      j.reportsTo
    ].join(' ').toLowerCase();

    return text.includes(jobSearch.toLowerCase());
  });

  return `
    <section class="hero">
      <h2>بطاقات الوصف الوظيفي</h2>
      <p>
        الأوصاف الوظيفية المعتمدة لكل وظيفة
        في ثانوية الجنان.
      </p>
    </section>

    <div class="note">
      لا يتضمن ملف أعمال المدير العام الأوصاف
      التفصيلية المعتمدة لجميع الوظائف.
      تُستكمل هذه البطاقات من الوثائق الرسمية للمؤسسة.
    </div>

    <div class="button-row">
      <h2>الوظائف المسجّلة (${available.length})</h2>
      <button class="primary" data-action="addJob">
        + إضافة وظيفة ووصفها
      </button>
    </div>

    <div class="filters">
      <input id="jobSearch"
             placeholder="بحث باسم الوظيفة أو مسؤولياتها"
             value="${e(jobSearch)}">

      <select id="jobDepartmentFilter">
        <option value="">جميع الأقسام</option>

        ${departments.map(d => `
          <option value="${e(d)}"
                  ${jobDepartment === d ? 'selected' : ''}>
            ${e(d)}
          </option>
        `).join('')}
      </select>
    </div>

    ${!available.length
      ? `<div class="panel">
           لا توجد بطاقات وصف وظيفي مسجّلة بعد.
         </div>`
      : `
        <div class="job-grid">
          ${available.map(j => `
            <article class="job-card"
                     data-theme="${deptTheme(j.department)}">

              <h3>${e(j.title)}</h3>
              <div class="muted">${e(j.department)}</div>

              <div class="job-fields">
                ${[
                  ['يرفع تقاريره إلى', j.reportsTo],
                  ['الغاية من الوظيفة', j.purpose],
                  ['المسؤوليات', j.responsibilities],
                  ['الصلاحيات', j.authority],
                  ['المؤهلات', j.qualifications],
                  ['مؤشرات الأداء', j.kpis]
                ].filter(([, value]) => value)
                 .map(([key, value]) => `
                   <p>
                     <b>${key}: </b>${e(value)}
                   </p>
                 `).join('') || '<p>الوصف لم يُستكمل بعد.</p>'}
              </div>

              <div class="quick-actions">
                <button class="row-action"
                        data-job="${e(j.id)}">
                  عرض وتعديل
                </button>

                ${linkButton(j.docUrl, 'الوصف المعتمد')}
              </div>
            </article>
          `).join('')}
        </div>
      `}
  `;
}

/* ===== إعدادات Google ===== */

function renderSettings() {
  $('content').innerHTML = `
    <section class="hero">
      <h2>الربط مع Google</h2>
      <p>
        التقويم والبريد ومشاركة التحديثات
        عبر Google Sheets.
      </p>
    </section>

    <div class="settings">
      <div class="panel">
        <h2>إعدادات حساب Google</h2>

        <label>
          OAuth Web Client ID
          <input id="clientId"
                 dir="ltr"
                 value="${e(config.clientId)}"
                 placeholder="...apps.googleusercontent.com">
        </label>

        <label>
          Calendar ID
          <input id="calendarId"
                 dir="ltr"
                 value="${e(config.calendarId)}">
        </label>

        <label>
          اسم المستخدم / المرسل
          <input id="editorName"
                 value="${e(config.editorName)}">
        </label>

        <div class="actions">
          <button class="primary"
                  data-action="saveConfig">
            حفظ الإعدادات
          </button>

          <button class="outline"
                  data-action="connectCalendar">
            ربط Google Calendar
          </button>

          <button class="outline"
                  data-action="connectMail">
            ربط Gmail للإرسال
          </button>
        </div>

        <p class="field-hint">
          فعّل Calendar API وGmail API،
          واضبط OAuth Web Client في Google Cloud.
        </p>
      </div>

      <div class="panel">
        <h2>مزامنة الأعمال بين الأجهزة</h2>

        <label>
          Google Spreadsheet ID
          <input id="sheetId"
                 dir="ltr"
                 value="${e(config.sheetId)}">
        </label>

        <div class="actions">
          <button class="outline"
                  data-action="connectSheets">
            ربط Google Sheets
          </button>

          <button class="outline"
                  data-action="syncCloud">
            مزامنة الآن
          </button>
        </div>

        <p class="field-hint">
          تستخدم المزامنة ورقة Jinan_Updates
          وتعمل كل 90 ثانية ما دام الحساب
          متصلاً والتبويب مفتوحاً.
        </p>

        <div class="note">
          وقت تسليم التقرير مأخوذ من ساعة جهاز
          المستخدم، وليس وقت خادم رسمي.
        </div>
      </div>
    </div>

    <div class="panel">
      <h2>النسخ الاحتياطية</h2>

      <div class="actions">
        <button class="outline"
                data-action="backup">
          تنزيل نسخة احتياطية
        </button>

        <button class="outline"
                data-action="restore">
          استيراد نسخة احتياطية
        </button>

        <button class="outline"
                data-action="ics">
          تصدير ICS
        </button>

        <input id="backupFile"
               type="file"
               accept=".json,application/json"
               hidden>
      </div>

      <p class="note">
        لا تستخدم GitHub Pages لحفظ بيانات
        الطلاب أو الملفات الحساسة بصورة عامة.
        المزامنة الحالية لا تفرض صلاحيات مختلفة
        بين المدير والموظف.
      </p>
    </div>
  `;
}

/* ===== صفحات المنظومة ===== */

function render() {
  $('content').dataset.theme =
    view === 'department'
      ? deptTheme(department)
      : 'neutral';

  $('pageTitle').textContent =
    view === 'department'
      ? department
      : TITLES[view];

  document.querySelectorAll('#menu button').forEach(b => {
    b.classList.toggle(
      'active',
      b.dataset.view === view
    );
  });

  if (view === 'dashboard') {
    const central = tasks.filter(
      t => t.kind === 'calendar'
    ).slice(0, 15);

    $('content').innerHTML = `
      <section class="hero">
        <h2>لوحة متابعة المدير العام</h2>
        <p>
          الأعمال والتقارير والقرارات
          والتذكيرات والمراسلات.
        </p>
      </section>

      <div class="legend">
        يُسجّل وقت وصول التقرير عند حفظ رابطه
        لأول مرة. ولا تظهر التغييرات على جهاز آخر
        إلا بعد مزامنتها مع Google Sheets.
      </div>

      ${kpis(tasks)}

      <div class="button-row">
        <h2>الأقسام الإدارية</h2>
        <button class="primary" data-action="newTask">
          + عمل جديد
        </button>
      </div>

      ${cards()}

      <div class="section-title">
        <h2>عينة من الروزنامة المركزية</h2>
      </div>

      <div class="panel">
        ${table(central)}
      </div>
    `;

  } else if (view === 'organization') {
    $('content').innerHTML = organization();

  } else if (view === 'departments') {
    $('content').innerHTML = `
      <section class="hero">
        <h2>الأقسام الإدارية</h2>
        <p>الأعمال والتقارير بحسب كل قسم.</p>
      </section>

      <div class="button-row">
        <h2>أقسام المؤسسة</h2>
        <button class="primary" data-action="newTask">
          + عمل جديد
        </button>
      </div>

      ${cards()}
    `;

  } else if (view === 'jobs') {
    $('content').innerHTML = jobsPage();

  } else if (view === 'department') {
    const list = mainList();

    $('content').innerHTML = `
      <section class="hero">
        <h2>${e(department)}</h2>
        <p>الاستحقاقات والتقارير وقرارات المدير.</p>
      </section>

      ${kpis(list)}

      <div class="button-row">
        <h2>سجل أعمال القسم</h2>

        <button class="primary" data-action="newTask">
          + إضافة عمل لهذا القسم
        </button>
      </div>

      ${taskArea(list, true)}
    `;

  } else if (view === 'calendar') {
    $('content').innerHTML = `
      <div class="panel">
        <div class="actions">
          <button class="outline" data-action="ics">
            تصدير المواعيد المؤرّخة ICS
          </button>
        </div>

        <div class="months">
          <button data-month=""
                  class="${!month ? 'active' : ''}">
            كل الأشهر
          </button>

          ${MONTHS.map(m => `
            <button data-month="${e(m)}"
                    class="${month === m ? 'active' : ''}">
              ${e(m)}
            </button>
          `).join('')}
        </div>
      </div>

      ${taskArea(mainList())}
    `;

  } else if (view === 'routine' ||
             view === 'government') {
    $('content').innerHTML = `
      <section class="hero">
        <h2>${TITLES[view]}</h2>
      </section>

      ${kpis(mainList())}
      ${taskArea(mainList())}
    `;

  } else if (view === 'settings') {
    renderSettings();
  }
}

function navigate(target) {
  view = target;
  search = '';
  tab = 'all';
  statusFilter = '';
  render();
}

/* ===== نافذة تفاصيل المهمة ===== */

function openTask(id) {
  selectedTask = byId(id);
  if (!selectedTask) return;

  const t = task(selectedTask);
  const dlg = $('taskDialog');

  dlg.dataset.theme = deptTheme(t.department);
  $('taskTitle').textContent = t.title;

  const details = [
    ['القسم', t.department],
    ['نوع العمل', t.section || KIND_LABEL[t.kind]],
    ['الاستحقاق كما ورد في Word', t.deadlineText],
    ['تفاصيل المطلوب', t.description],
    ['مصادر المعلومات والأدلة', t.evidence],
    ['المسؤولون وفق Word', t.sourceAssignee],
    ['دور المدير الأصلي', t.managerAction],
    ['المتابعة اللاحقة', t.followUp],
    ['المرجع', t.reference]
  ];

  $('taskInfo').innerHTML = details
    .filter(([, value]) => value)
    .map(([label, value]) => `
      <div>
        <b>${e(label)}:</b>
        ${e(value)}
      </div>
    `).join('');

  const fields = [
    ['taskStatus', 'status'],
    ['taskAssignee', 'assignee'],
    ['taskDate', 'date'],
    ['taskProof', 'proofUrl'],
    ['submittedBy', 'submittedBy'],
    ['taskDecision', 'decision'],
    ['taskNotes', 'notes'],
    ['recipientEmail', 'recipientEmail']
  ];

  fields.forEach(([element, key]) => {
    $(element).value = t[key] || '';
  });

  $('reminderMinutes').value =
    String(t.reminderMinutes || 1440);

  $('editCustomTask').hidden =
    !customTasks[t.id] || customTasks[t.id].deleted;

  $('taskReceivedAt').value =
    t.receivedAt ? stampView(t.receivedAt) : '';

  $('reportLinkBox').innerHTML =
    t.proofUrl && safeUrl(t.proofUrl)
      ? `<div class="report-box">
           التقرير المسلّم:
           ${linkButton(t.proofUrl, 'فتح رابط التقرير')}
           · ${e(stampView(t.receivedAt))}
         </div>`
      : '';

  $('emailAudit').textContent =
    t.commentSentAt
      ? `آخر تعليق أُرسل إلى ${t.commentSentTo || '—'}
         بتاريخ ${stampView(t.commentSentAt)}`
      : 'لم يُرسل تعليق بالبريد بعد.';

  dlg.showModal();
}

/* ===== حفظ التعديلات ===== */

function persist(type, id, value) {
  const key = type === 'task_change'
    ? STORE
    : type === 'custom_task'
      ? EXTRA_STORE
      : JOB_STORE;

  const target = type === 'task_change'
    ? changes
    : type === 'custom_task'
      ? customTasks
      : jobs;

  const next = {...target, [id]: value};

  if (!saveLocal(key, next)) {
    notify('تعذّر الحفظ في المتصفح', true);
    return false;
  }

  if (type === 'task_change') {
    changes = next;
  } else if (type === 'custom_task') {
    customTasks = next;
    buildTasks();
  } else {
    jobs = next;
  }

  const event = {
    timestamp: stamp(),
    type,
    id,
    payload: value,
    editor: config.editorName || ''
  };

  pending = [...pending, event];
  saveLocal(PENDING_STORE, pending);

  indicator(
    `بانتظار مزامنة (${pending.length})`,
    'warn'
  );

  if (canUseScope(SCOPE.sheets) && config.sheetId) {
    void syncCloud({
      silent: true,
      interactive: false
    });
  }

  return true;
}

function saveTask() {
  if (!selectedTask) return false;

  const original = task(selectedTask);
  const id = selectedTask.id;

  const link = $('taskProof').value.trim();
  const email = $('recipientEmail').value.trim();

  if (link && !safeUrl(link)) {
    notify('يجب أن يبدأ رابط التقرير بـ https://', true);
    return false;
  }

  if (email && !validEmail(email)) {
    notify('صيغة البريد الإلكتروني غير صحيحة', true);
    return false;
  }

  const updated = {
    ...(changes[id] || {}),
    status: $('taskStatus').value,
    assignee: $('taskAssignee').value.trim(),
    date: $('taskDate').value,
    reminderMinutes: Number($('reminderMinutes').value),
    proofUrl: link,
    submittedBy: $('submittedBy').value.trim(),
    recipientEmail: email,
    decision: $('taskDecision').value.trim(),
    notes: $('taskNotes').value.trim(),
    updatedAt: stamp()
  };

  if (!original.receivedAt && link) {
    updated.receivedAt = stamp();

    if (updated.status === 'لم يُستلم') {
      updated.status = 'تم الاستلام';
    }

  } else if (original.receivedAt) {
    updated.receivedAt = original.receivedAt;
  }

  if (!persist('task_change', id, updated)) {
    return false;
  }

  $('taskReceivedAt').value =
    updated.receivedAt
      ? stampView(updated.receivedAt)
      : '';

  render();
  return true;
}

/* ===== إضافة أعمال جديدة ===== */

function openNewTask(id = null, preset = '') {
  editingTaskId = id;
  const old = id ? customTasks[id] : null;

  $('customTaskHeading').textContent =
    old ? 'تعديل العمل الإضافي' : 'إضافة عمل جديد';

  $('newDepartment').innerHTML = departments.map(d => `
    <option value="${e(d)}">${e(d)}</option>
  `).join('');

  $('newTitle').value = old?.title || '';

  $('newDepartment').value =
    old?.department || preset || department || departments[0];

  $('newKind').value = old?.kind || 'plan';
  $('newDeadline').value = old?.deadlineText || '';
  $('newDate').value = old?.date || '';
  $('newEvidence').value = old?.evidence || '';
  $('newAssignee').value = old?.sourceAssignee || '';
  $('newManagerAction').value = old?.managerAction || '';
  $('newFollowUp').value = old?.followUp || '';

  $('deleteCustomTask').hidden = !old;
  $('customTaskDialog').showModal();
}

function saveCustomTask() {
  const title = $('newTitle').value.trim();

  if (!title) {
    notify('عنوان العمل مطلوب', true);
    return false;
  }

  const previous = editingTaskId
    ? customTasks[editingTaskId]
    : null;

  const id = editingTaskId ||
    `CUSTOM-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

  const kind = $('newKind').value;

  const record = {
    ...previous,
    id,
    kind,
    department: $('newDepartment').value,
    section: 'أعمال أضيفت إلى المنظومة',
    title,
    frequency: KIND_LABEL[kind],
    deadlineText: $('newDeadline').value.trim(),
    date: $('newDate').value,
    evidence: $('newEvidence').value.trim(),
    sourceAssignee: $('newAssignee').value.trim(),
    managerAction: $('newManagerAction').value.trim(),
    followUp: $('newFollowUp').value.trim(),
    status: previous?.status || 'لم يُستلم',
    reference: 'إضافة إلكترونية — خارج ملف Word الأصلي',
    createdAt: previous?.createdAt || stamp(),
    updatedAt: stamp(),
    deleted: false
  };

  return persist('custom_task', id, record);
}

function deleteCustomTask() {
  if (!editingTaskId ||
      !customTasks[editingTaskId]) return;

  if (!confirm('حذف هذا العمل المضاف؟')) return;

  const id = editingTaskId;
  const old = customTasks[id];

  if (persist('custom_task', id, {
    ...old,
    deleted: true,
    updatedAt: stamp()
  })) {
    $('customTaskDialog').close();
    render();
    notify('تم إلغاء العمل الإضافي');
  }
}

/* ===== إضافة الأوصاف الوظيفية ===== */

function openJob(id = null, preset = '', proposedTitle = '') {
  editingJobId = id;

  const old = id ? jobs[id] : null;

  $('jobHeading').textContent =
    old ? 'تعديل الوصف الوظيفي' : 'إضافة وصف وظيفي';

  $('jobDepartment').innerHTML = departments.map(d => `
    <option value="${e(d)}">${e(d)}</option>
  `).join('');

  const fields = {
    jobTitle: 'title',
    jobReportsTo: 'reportsTo',
    jobPurpose: 'purpose',
    jobResponsibilities: 'responsibilities',
    jobAuthority: 'authority',
    jobQualifications: 'qualifications',
    jobKpis: 'kpis',
    jobDocUrl: 'docUrl'
  };

  Object.entries(fields).forEach(([element, key]) => {
    $(element).value = old?.[key] || '';
  });

  $('jobDepartment').value =
    old?.department || preset || departments[0];

  if (!old && proposedTitle) {
    $('jobTitle').value = proposedTitle;
  }

  $('deleteJob').hidden = !old;
  $('jobDialog').showModal();
}

function saveJob() {
  const title = $('jobTitle').value.trim();
  const docUrl = $('jobDocUrl').value.trim();

  if (!title) {
    notify('المسمى الوظيفي مطلوب', true);
    return false;
  }

  if (docUrl && !safeUrl(docUrl)) {
    notify('رابط الوصف الوظيفي يجب أن يبدأ بـ https://', true);
    return false;
  }

  const old = editingJobId ? jobs[editingJobId] : null;

  const id = editingJobId ||
    `ROLE-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

  const record = {
    ...old,
    id,
    title,
    department: $('jobDepartment').value,
    reportsTo: $('jobReportsTo').value.trim(),
    purpose: $('jobPurpose').value.trim(),
    responsibilities: $('jobResponsibilities').value.trim(),
    authority: $('jobAuthority').value.trim(),
    qualifications: $('jobQualifications').value.trim(),
    kpis: $('jobKpis').value.trim(),
    docUrl,
    updatedAt: stamp(),
    createdAt: old?.createdAt || stamp(),
    deleted: false
  };

  return persist('job', id, record);
}

function deleteJob() {
  if (!editingJobId || !jobs[editingJobId]) return;
  if (!confirm('هل تريد حذف بطاقة الوصف الوظيفي؟')) return;

  const id = editingJobId;

  if (persist('job', id, {
    ...jobs[id],
    deleted: true,
    updatedAt: stamp()
  })) {
    $('jobDialog').close();
    render();
    notify('حُذفت البطاقة');
  }
}

/* ===== النسخ الاحتياطية ===== */

function download(name, text, mime) {
  const a = document.createElement('a');

  const url = URL.createObjectURL(
    new Blob([text], {type: mime})
  );

  a.href = url;
  a.download = name;
  a.click();

  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

function backup() {
  download(
    'jinan-director-backup-v3.json',
    JSON.stringify({
      changes,
      customTasks,
      jobs,
      pending,
      config,
      exportedAt: stamp()
    }, null, 2),
    'application/json'
  );
}

async function restore(file) {
  if (!file) return;

  try {
    const data = JSON.parse(await file.text());

    if (
      !data.changes ||
      typeof data.changes !== 'object'
    ) {
      throw Error('نسخة احتياطية غير صحيحة');
    }

    if (!confirm(
      'استيراد النسخة سيستبدل المعلومات المحلية الحالية. متابعة؟'
    )) return;

    changes = data.changes;
    customTasks = data.customTasks || {};
    jobs = data.jobs || {};
    pending = data.pending || [];

    const saved = [
      saveLocal(STORE, changes),
      saveLocal(EXTRA_STORE, customTasks),
      saveLocal(JOB_STORE, jobs),
      saveLocal(PENDING_STORE, pending)
    ];

    if (!saved.every(Boolean)) {
      throw Error('تعذّر حفظ جزء من النسخة المستوردة');
    }

    buildTasks();
    render();

    notify('استُعيدت النسخة الاحتياطية محلياً');

  } catch (err) {
    notify(err.message, true);
  }
}

function exportIcs() {
  const esc = s => String(s || '')
    .replaceAll('\\', '\\\\')
    .replaceAll(',', '\\,')
    .replaceAll(';', '\\;')
    .replaceAll('\n', '\\n');

  const out = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Al Jinan High School//Management//AR',
    'CALSCALE:GREGORIAN'
  ];

  for (const original of tasks) {
    const t = task(original);
    if (!t.date) continue;

    const a = t.date.replaceAll('-', '');

    const b = new Date(
      Date.parse(t.date + 'T00:00:00Z') + 86400000
    ).toISOString().slice(0, 10).replaceAll('-', '');

    out.push(
      'BEGIN:VEVENT',
      `UID:${t.id}@jinan-management`,
      `DTSTAMP:${stamp()
        .replace(/[-:]/g, '')
        .replace(/\.\d{3}Z$/, 'Z')}`,
      `DTSTART;VALUE=DATE:${a}`,
      `DTEND;VALUE=DATE:${b}`,
      `SUMMARY:${esc(t.title)}`,
      `DESCRIPTION:${esc(t.department)}`,
      'END:VEVENT'
    );
  }

  out.push('END:VCALENDAR');

  download(
    'jinan-calendar.ics',
    out.join('\r\n'),
    'text/calendar;charset=utf-8'
  );
}

/* ===== Google OAuth ===== */

function canUseScope(scope) {
  return Boolean(
    oauthToken &&
    Date.now() < tokenExpires - 45000 &&
    oauthGranted.has(scope)
  );
}

async function tokenFor(scope, interactive = true) {
  if (canUseScope(scope)) return oauthToken;

  if (!interactive) {
    throw Error('اتصال Google يحتاج إعادة تفويض');
  }

  if (!config.clientId) {
    throw Error('أضف OAuth Client ID في الإعدادات أولاً');
  }

  if (!window.google?.accounts?.oauth2) {
    throw Error('مكتبة Google غير متاحة');
  }

  return new Promise((resolve, reject) => {
    const client = google.accounts.oauth2.initTokenClient({
      client_id: config.clientId,
      scope,
      include_granted_scopes: true,

      callback: result => {
        if (result.error) {
          reject(Error(result.error));
          return;
        }

        oauthToken = result.access_token;

        tokenExpires =
          Date.now() +
          Number(result.expires_in || 3600) * 1000;

        oauthGranted = new Set(
          String(result.scope || scope)
            .split(/\s+/)
            .filter(Boolean)
        );

        resolve(oauthToken);
      },

      error_callback: result => {
        reject(
          Error(
            result.message || 'لم تكتمل موافقة Google'
          )
        );
      }
    });

    client.requestAccessToken();
  });
}

async function googleApi(
  url,
  method = 'GET',
  body,
  scope = SCOPE.calendar,
  interactive = true
) {
  const accessToken = await tokenFor(scope, interactive);

  const response = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(body !== undefined
        ? {'Content-Type': 'application/json'}
        : {})
    },
    ...(body !== undefined
      ? {body: JSON.stringify(body)}
      : {})
  });

  const data = await response.json().catch(() => ({}));

  if (!response.ok) {
    const error = Error(
      data.error?.message ||
      `فشل Google API: HTTP ${response.status}`
    );

    error.status = response.status;
    throw error;
  }

  return data;
}

function saveConfig() {
  config = {
    clientId: $('clientId').value.trim(),
    calendarId: $('calendarId').value.trim() || 'primary',
    sheetId: $('sheetId').value.trim(),
    editorName: $('editorName').value.trim()
  };

  if (!saveLocal(CONFIG_STORE, config)) {
    notify('تعذّر حفظ الإعدادات', true);
    return;
  }

  oauthToken = '';
  oauthGranted = new Set();

  indicator('إعدادات محفوظة محلياً');
  notify('تم حفظ الإعدادات. اربط خدمات Google المطلوبة.');
}

async function connectGoogle(scope, label) {
  try {
    await tokenFor(scope);
    notify(`تم السماح باستخدام ${label}`);

    if (scope === SCOPE.sheets) {
      await syncCloud({
        silent: false,
        interactive: false
      });
    }

  } catch (err) {
    notify(err.message, true);
  }
}

/* =====================================
   Google Calendar - زر ذكّرني
   ===================================== */

async function remindTask(id) {
  const source = byId(id);
  if (!source) return;

  const t = task(source);

  if (!t.date) {
    openTask(id);
    $('taskDate').focus();

    notify(
      'أدخل تاريخ الاستحقاق أولاً ثم اضغط ذكّرني',
      true
    );
    return;
  }

  if (!/^\d{4}-\d{2}-\d{2}$/.test(t.date)) {
    notify('صيغة الموعد غير صحيحة', true);
    return;
  }

  try {
    const end = new Date(
      Date.parse(t.date + 'T00:00:00Z') + 86400000
    ).toISOString().slice(0, 10);

    const base =
      'https://www.googleapis.com/calendar/v3/calendars/' +
      encodeURIComponent(config.calendarId || 'primary') +
      '/events';

    const minutes = Number(t.reminderMinutes || 1440);

    const body = {
      summary: 'متابعة: ' + t.title,

      description: [
        `القسم: ${t.department}`,
        `رقم العمل: ${t.id}`,
        `المسؤول: ${t.assignee || t.sourceAssignee || ''}`,
        `مطلوب من المدير: ${t.managerAction || ''}`
      ].join('\n'),

      start: {date: t.date},
      end: {date: end},

      reminders: {
        useDefault: false,
        overrides: [
          {method: 'popup', minutes},
          {method: 'email', minutes}
        ]
      }
    };

    let event;

    if (t.calendarEventId) {
      try {
        event = await googleApi(
          `${base}/${encodeURIComponent(t.calendarEventId)}`,
          'PUT',
          body,
          SCOPE.calendar
        );

      } catch (err) {
        if (err.status !== 404) throw err;

        event = await googleApi(
          base,
          'POST',
          body,
          SCOPE.calendar
        );
      }

    } else {
      event = await googleApi(
        base,
        'POST',
        body,
        SCOPE.calendar
      );
    }

    const next = {
      ...(changes[t.id] || {}),
      calendarEventId: event.id,
      calendarLink: event.htmlLink || '',
      updatedAt: stamp()
    };

    if (!persist('task_change', t.id, next)) {
      notify(
        'أُنشئ التذكير، لكن تعذّر حفظ مرجعه محلياً',
        true
      );
      return;
    }

    notify('تم حفظ التذكير في Google Calendar');

    if ($('taskDialog').open) {
      openTaskAgain(t.id);
    }

    render();

  } catch (err) {
    notify(
      'تعذّر إنشاء التذكير: ' + err.message,
      true
    );
  }
}

function openTaskAgain(id) {
  if ($('taskDialog').open) {
    $('taskDialog').close();
  }

  openTask(id);
}

async function remindFromDialog() {
  if (!selectedTask || !saveTask()) return;

  await remindTask(selectedTask.id);
}

function calendarDraft() {
  if (!selectedTask || !saveTask()) return;

  const t = task(selectedTask);

  if (!t.date) {
    notify('حدّد موعد المهمة أولاً', true);
    return;
  }

  const from = t.date.replaceAll('-', '');

  const to = new Date(
    Date.parse(t.date + 'T00:00:00Z') + 86400000
  ).toISOString().slice(0, 10).replaceAll('-', '');

  const params = new URLSearchParams({
    action: 'TEMPLATE',
    text: 'متابعة: ' + t.title,
    dates: `${from}/${to}`,

    details:
      `القسم: ${t.department}\n` +
      `المسؤول: ${t.assignee || t.sourceAssignee || ''}\n` +
      `${t.managerAction || ''}`
  });

  window.open(
    'https://calendar.google.com/calendar/render?' + params,
    '_blank',
    'noopener,noreferrer'
  );
}

/* =====================================
   Gmail - إرسال تعليق المدير
   ===================================== */

function base64UrlUtf8(text) {
  const bytes = new TextEncoder().encode(String(text));

  let result = '';

  for (let i = 0; i < bytes.length; i += 8192) {
    result += String.fromCharCode(
      ...bytes.subarray(i, i + 8192)
    );
  }

  return btoa(result)
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

async function sendCommentEmail() {
  if (!selectedTask) return;

  const recipient = $('recipientEmail').value.trim();
  const decision = $('taskDecision').value.trim();

  if (!validEmail(recipient)) {
    notify('أدخل بريداً إلكترونياً صحيحاً', true);
    return;
  }

  if (!decision) {
    notify('اكتب تعليق المدير أولاً', true);
    return;
  }

  const approved = confirm(
    `إرسال تعليق المدير إلى ${recipient} عبر Gmail؟`
  );

  if (!approved) return;
  if (!saveTask()) return;

  const t = task(selectedTask);

  const subject = `تعليق المدير على تقرير: ${t.title}`;

  const body = [
    'ثانوية الجنان — متابعة أعمال الإدارة',
    '',
    `العمل: ${t.title}`,
    `القسم: ${t.department}`,
    `رقم المتابعة: ${t.id}`,
    '',
    'تعليق / قرار المدير:',
    decision,
    '',
    t.proofUrl ? `رابط التقرير: ${t.proofUrl}` : '',
    '',
    'هذه الرسالة صادرة من منظومة متابعة المدير العام.'
  ].join('\r\n');

  const toMimeBase64 = text => {
    const base = base64UrlUtf8(text)
      .replace(/-/g, '+')
      .replace(/_/g, '/');

    return base.padEnd(
      Math.ceil(base.length / 4) * 4,
      '='
    );
  };

  const encodedSubject = toMimeBase64(subject);

  const encodedBody = toMimeBase64(body)
    .match(/.{1,76}/g)
    .join('\r\n');

  const mime = [
    `To: ${recipient}`,
    `Subject: =?UTF-8?B?${encodedSubject}?=`,
    'MIME-Version: 1.0',
    'Content-Type: text/plain; charset=UTF-8',
    'Content-Transfer-Encoding: base64',
    '',
    encodedBody
  ].join('\r\n');

  try {
    const result = await googleApi(
      'https://gmail.googleapis.com/gmail/v1/users/me/messages/send',
      'POST',
      {
        raw: base64UrlUtf8(mime)
      },
      SCOPE.gmail
    );

    const next = {
      ...(changes[t.id] || {}),
      commentSentAt: stamp(),
      commentSentTo: recipient,
      gmailMessageId: result.id || '',
      updatedAt: stamp()
    };

    persist('task_change', t.id, next);

    notify(`تم إرسال التعليق إلى ${recipient}`);

    if ($('taskDialog').open) {
      $('emailAudit').textContent =
        `تم إرسال التعليق إلى ${recipient} ` +
        `في ${stampView(next.commentSentAt)}`;
    }

  } catch (err) {
    notify(
      'لم يُرسل البريد: ' + err.message,
      true
    );
  }
}

/* =====================================
   Google Sheets - المزامنة بين الأجهزة
   ===================================== */

function sheetBase() {
  if (!config.sheetId) {
    throw Error('أدخل Spreadsheet ID في الإعدادات');
  }

  return (
    'https://sheets.googleapis.com/v4/spreadsheets/' +
    encodeURIComponent(config.sheetId)
  );
}

async function ensureLogSheet(interactive = true) {
  const base = sheetBase();

  const meta = await googleApi(
    base + '?fields=sheets.properties.title',
    'GET',
    undefined,
    SCOPE.sheets,
    interactive
  );

  const exists = meta.sheets?.some(
    s => s.properties?.title === SHEET_NAME
  );

  if (!exists) {
    await googleApi(
      base + ':batchUpdate',
      'POST',
      {
        requests: [
          {
            addSheet: {
              properties: {
                title: SHEET_NAME
              }
            }
          }
        ]
      },
      SCOPE.sheets,
      interactive
    );
  }
}

async function appendEvent(event, interactive = true) {
  const base = sheetBase();

  const range = encodeURIComponent(
    `'${SHEET_NAME}'!A:E`
  );

  await googleApi(
    `${base}/values/${range}:append` +
      '?valueInputOption=RAW&insertDataOption=INSERT_ROWS',

    'POST',

    {
      values: [
        [
          event.timestamp,
          event.type,
          event.id,
          JSON.stringify(event.payload),
          event.editor || ''
        ]
      ]
    },

    SCOPE.sheets,
    interactive
  );
}

function applyRemoteEvent(record) {
  const {type, id, payload} = record;

  if (
    !id ||
    !payload ||
    typeof payload !== 'object'
  ) return;

  const old =
    type === 'task_change'
      ? changes[id]
      : type === 'custom_task'
        ? customTasks[id]
        : type === 'job'
          ? jobs[id]
          : null;

  const latest = String(
    payload.updatedAt || record.timestamp || ''
  );

  if (
    old &&
    String(old.updatedAt || '') >= latest
  ) {
    return;
  }

  if (
    type === 'task_change' &&
    tasks.some(t => t.id === id)
  ) {
    changes[id] = payload;

  } else if (type === 'custom_task') {
    customTasks[id] = payload;

  } else if (type === 'job') {
    jobs[id] = payload;
  }
}

async function readCloud(interactive = true) {
  const base = sheetBase();

  const range = encodeURIComponent(
    `'${SHEET_NAME}'!A2:E10000`
  );

  const data = await googleApi(
    `${base}/values/${range}`,
    'GET',
    undefined,
    SCOPE.sheets,
    interactive
  );

  let count = 0;

  for (const row of data.values || []) {
    const [
      timestamp,
      type,
      id,
      payloadJSON,
      editor
    ] = row;

    try {
      applyRemoteEvent({
        timestamp,
        type,
        id,
        payload: JSON.parse(payloadJSON),
        editor
      });

      count++;

    } catch (err) {
      console.warn('سجل سحابي غير صالح', err);
    }
  }

  buildTasks();

  saveLocal(STORE, changes);
  saveLocal(EXTRA_STORE, customTasks);
  saveLocal(JOB_STORE, jobs);

  return count;
}

async function syncCloud({
  silent = false,
  interactive = true
} = {}) {
  if (syncBusy) return;

  if (!config.sheetId) {
    if (!silent) {
      notify('أدخل Spreadsheet ID في الإعدادات', true);
    }
    return;
  }

  syncBusy = true;

  try {
    await ensureLogSheet(interactive);

    let uploaded = 0;

    while (pending.length) {
      const entry = pending[0];

      await appendEvent(entry, interactive);

      pending = pending.slice(1);

      if (!saveLocal(PENDING_STORE, pending)) {
        notify('فشل حفظ قائمة المزامنة محلياً', true);
      }

      uploaded++;
    }

    const received = await readCloud(interactive);

    indicator('متزامن مع Google Sheets', 'ok');

    if (!silent) {
      notify(
        `تم رفع ${uploaded} تحديثات ` +
        `واستقبال ${received} سجلاً`
      );
    }

    if (
      !$('taskDialog').open &&
      !$('customTaskDialog').open &&
      !$('jobDialog').open
    ) {
      render();
    }

  } catch (err) {
    indicator(
      pending.length
        ? `غير متزامن (${pending.length})`
        : 'تعذّر التحديث',
      'warn'
    );

    if (!silent) {
      notify('تعذّرت المزامنة: ' + err.message, true);
    } else {
      console.warn('مزامنة Google', err);
    }

  } finally {
    syncBusy = false;
  }
}

/* =====================================
   أحداث الواجهة
   ===================================== */

document.addEventListener('click', async event => {
  const btn = event.target.closest('button');

  if (!btn) return;

  if (btn.dataset.view) {
    navigate(btn.dataset.view);
    return;
  }

  if (btn.dataset.dept) {
    department = btn.dataset.dept;
    navigate('department');
    return;
  }

  if (btn.dataset.task) {
    openTask(btn.dataset.task);
    return;
  }

  if (btn.dataset.remind) {
    await remindTask(btn.dataset.remind);
    return;
  }

  if (btn.dataset.tab) {
    tab = btn.dataset.tab;
    render();
    return;
  }

  if (btn.dataset.month !== undefined) {
    month = btn.dataset.month;
    render();
    return;
  }

  if (btn.dataset.job) {
    openJob(btn.dataset.job);
    return;
  }

  if (btn.dataset.addOrgRole) {
    openJob(
      null,
      btn.dataset.addOrgDept || '',
      btn.dataset.addOrgRole
    );
    return;
  }

  if (btn.dataset.newJobDept) {
    openJob(null, btn.dataset.newJobDept);
    return;
  }

  if (btn.dataset.close) {
    $(btn.dataset.close)?.close();
    return;
  }

  switch (btn.dataset.action) {
    case 'newTask':
      openNewTask(
        null,
        view === 'department' ? department : ''
      );
      break;

    case 'addJob':
      openJob();
      break;

    case 'backup':
      backup();
      break;

    case 'restore':
      $('backupFile')?.click();
      break;

    case 'ics':
      exportIcs();
      break;

    case 'saveConfig':
      saveConfig();
      break;

    case 'connectCalendar':
      await connectGoogle(
        SCOPE.calendar,
        'Google Calendar'
      );
      break;

    case 'connectMail':
      await connectGoogle(SCOPE.gmail, 'Gmail');
      break;

    case 'connectSheets':
      await connectGoogle(
        SCOPE.sheets,
        'Google Sheets'
      );
      break;

    case 'syncCloud':
      await syncCloud();
      break;
  }
});

document.addEventListener('input', event => {
  if (event.target.id === 'search') {
    search = event.target.value;
    updateRows();
  }

  if (event.target.id === 'jobSearch') {
    const cursor = event.target.selectionStart;

    jobSearch = event.target.value;
    render();

    const input = $('jobSearch');

    if (input) {
      input.focus();
      input.setSelectionRange(cursor, cursor);
    }
  }
});

document.addEventListener('change', event => {
  if (event.target.id === 'statusFilter') {
    statusFilter = event.target.value;
    updateRows();
  }

  if (event.target.id === 'jobDepartmentFilter') {
    jobDepartment = event.target.value;
    render();
  }

  if (event.target.id === 'backupFile') {
    void restore(event.target.files?.[0]);
  }
});

/* =====================================
   ربط أزرار الصفحات والنوافذ
   ===================================== */

$('exportBtn').addEventListener('click', backup);

$('refreshCloud').addEventListener(
  'click',
  () => syncCloud()
);

$('closeDialog').addEventListener(
  'click',
  () => $('taskDialog').close()
);

$('taskForm').addEventListener('submit', event => {
  event.preventDefault();

  if (saveTask()) {
    $('taskDialog').close();

    notify(
      'تم حفظ التغييرات محلياً' +
      (config.sheetId
        ? ' وسيُحاول إرسالها إلى Google Sheets'
        : '؛ لعرضها على جهاز آخر اربط Google Sheets')
    );
  }
});

$('remindDirect').addEventListener(
  'click',
  remindFromDialog
);

$('editCustomTask').addEventListener('click', () => {
  if (
    !selectedTask ||
    !customTasks[selectedTask.id]
  ) return;

  $('taskDialog').close();
  openNewTask(selectedTask.id);
});

$('sendComment').addEventListener(
  'click',
  sendCommentEmail
);

$('calendarDraft').addEventListener(
  'click',
  calendarDraft
);

$('customTaskForm').addEventListener('submit', event => {
  event.preventDefault();

  if (saveCustomTask()) {
    $('customTaskDialog').close();
    render();
    notify('أُضيف العمل بنجاح');
  }
});

$('deleteCustomTask').addEventListener(
  'click',
  deleteCustomTask
);

$('jobForm').addEventListener('submit', event => {
  event.preventDefault();

  if (saveJob()) {
    $('jobDialog').close();
    render();
    notify('تم حفظ بطاقة الوصف الوظيفي');
  }
});

$('deleteJob').addEventListener(
  'click',
  deleteJob
);

/* =====================================
   تشغيل المنظومة وتحميل JSON
   ===================================== */

(async function initialize() {
  try {
    const embedded = $('embeddedData');

    let data;

    if (embedded?.textContent.trim()) {
      data = JSON.parse(embedded.textContent);

    } else {
      const response = await fetch(
        './data/tasks.json',
        {cache: 'no-store'}
      );

      if (!response.ok) {
        throw Error(
          `لم يُعثر على data/tasks.json (${response.status})`
        );
      }

      data = await response.json();
    }

    if (
      !Array.isArray(data.tasks) ||
      !Array.isArray(data.departments) ||
      data.tasks.some(
        t => !t.id || !t.title || !t.department
      )
    ) {
      throw Error('صيغة ملف الأعمال غير صحيحة');
    }

    baseTasks = data.tasks;
    departments = data.departments;

    buildTasks();

    indicator(
      pending.length
        ? `غير متزامن (${pending.length})`
        : 'حفظ محلي',
      pending.length ? 'warn' : ''
    );

    render();

  } catch (error) {
    console.error(error);

    $('content').innerHTML = `
      <div class="note">
        ${e(error.message)}
        — تأكد من رفع data/tasks.json
        مع الملفات الأخرى.
      </div>
    `;
  }
})();

/* =====================================
   التحديث الدوري من Google Sheets
   ===================================== */

setInterval(() => {
  if (
    document.visibilityState === 'visible' &&
    config.sheetId &&
    canUseScope(SCOPE.sheets)
  ) {
    void syncCloud({
      silent: true,
      interactive: false
    });
  }
}, 90000);

