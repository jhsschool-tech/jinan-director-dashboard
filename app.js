
const STORE = 'jinan-manager-changes-word-v2';
const GOOGLE_STORE = 'jinan-manager-config-v1';

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
  departments: 'الأقسام الإدارية',
  department: 'أعمال القسم',
  calendar: 'الروزنامة السنوية',
  routine: 'الأعمال الثابتة',
  government: 'المعاملات الرسمية',
  settings: 'الإعدادات والربط'
};

const content = document.getElementById('content');
const dialog = document.getElementById('taskDialog');
const el = id => document.getElementById(id);

const escapeHTML = value =>
  String(value ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[c]));

const e = escapeHTML;

function load(key, fallback) {
  try {
    return JSON.parse(localStorage.getItem(key)) || fallback;
  } catch {
    return fallback;
  }
}

function save(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
    return true;
  } catch {
    return false;
  }
}

let tasks = [];
let departments = [];
let changes = load(STORE, {});

let googleConfig = {
  clientId: '',
  calendarId: 'primary',
  sheetId: '',
  ...load(GOOGLE_STORE, {})
};

let view = 'dashboard';
let department = '';
let month = '';
let category = 'all';
let query = '';
let filterStatus = '';

let selected = null;
let accessToken = '';
let tokenExpiry = 0;
let tokenScope = '';

// =======================================
// ألوان الأقسام
// =======================================

function theme(name = '') {
  if (name.includes('الإشراف الإداري')) {
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

// =======================================
// بيانات المهام والحالات
// =======================================

const current = task => ({
  ...task,
  ...(changes[task.id] || {})
});

const byId = id => tasks.find(t => t.id === id);

const isDone = t => current(t).status === 'مُقفل';

function today() {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Beirut',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(new Date());

  const get = name =>
    parts.find(p => p.type === name).value;

  return `${get('year')}-${get('month')}-${get('day')}`;
}

function isLate(t) {
  const item = current(t);

  return !!(
    item.date &&
    item.date < today() &&
    !isDone(t)
  );
}

function isSoon(t) {
  const date = current(t).date;

  if (!date || isDone(t)) return false;

  const days = (
    Date.parse(date + 'T00:00:00Z') -
    Date.parse(today() + 'T00:00:00Z')
  ) / 86400000;

  return days >= 0 && days <= 14;
}

function dateView(date) {
  if (!date) return 'غير محدد';

  return new Date(date + 'T12:00:00Z')
    .toLocaleDateString('ar-LB', {
      day: 'numeric',
      month: 'short',
      year: 'numeric',
      timeZone: 'UTC'
    });
}

function notice(message) {
  const toast = el('toast');

  toast.textContent = message;
  toast.style.display = 'block';

  setTimeout(() => {
    toast.style.display = 'none';
  }, 4000);
}

// =======================================
// قراءة ملف Word الأصلي
// =======================================

// استخراج ملف document.xml من داخل DOCX
// بواسطة إمكانات المتصفح الحديثة.

async function wordXml(filename) {
  const response = await fetch(filename, {
    cache: 'no-store'
  });

  if (!response.ok) {
    throw Error(
      `لم يُعثر على ملف Word (${response.status}): ${filename}`
    );
  }

  const bytes = new Uint8Array(
    await response.arrayBuffer()
  );

  const dv = new DataView(bytes.buffer);

  const u16 = n => dv.getUint16(n, true);
  const u32 = n => dv.getUint32(n, true);

  let end = -1;

  for (
    let i = bytes.length - 22;
    i >= Math.max(0, bytes.length - 65557);
    i--
  ) {
    if (u32(i) === 0x06054b50) {
      end = i;
      break;
    }
  }

  if (end < 0) {
    throw Error('ملف Word غير صالح');
  }

  const decode = new TextDecoder('utf-8');
  let pos = u32(end + 16);

  for (let i = 0; i < u16(end + 10); i++) {
    if (u32(pos) !== 0x02014b50) {
      throw Error('خطأ في بيانات Word');
    }

    const method = u16(pos + 10);
    const size = u32(pos + 20);
    const nameLen = u16(pos + 28);
    const extra = u16(pos + 30);
    const comment = u16(pos + 32);
    const offset = u32(pos + 42);

    const name = decode.decode(
      bytes.slice(pos + 46, pos + 46 + nameLen)
    );

    pos += 46 + nameLen + extra + comment;

    if (name !== 'word/document.xml') {
      continue;
    }

    if (u32(offset) !== 0x04034b50) {
      throw Error('محتوى Word غير صالح');
    }

    const start =
      offset +
      30 +
      u16(offset + 26) +
      u16(offset + 28);

    const compressed = bytes.slice(
      start,
      start + size
    );

    if (method === 0) {
      return decode.decode(compressed);
    }

    if (
      method !== 8 ||
      !('DecompressionStream' in window)
    ) {
      throw Error(
        'استخدم Chrome أو Edge محدّثاً لقراءة ملف Word'
      );
    }

    const stream = new Blob([compressed])
      .stream()
      .pipeThrough(
        new DecompressionStream('deflate-raw')
      );

    return decode.decode(
      await new Response(stream).arrayBuffer()
    );
  }

  throw Error(
    'ملف Word لا يحتوي على جداول قابلة للقراءة'
  );
}

// =======================================
// استخراج الجداول من XML
// =======================================

function extractTables(xmlText) {
  const xml = new DOMParser().parseFromString(
    xmlText,
    'application/xml'
  );

  if (
    xml.getElementsByTagName('parsererror').length
  ) {
    throw Error('تعذّر تحليل ملف Word');
  }

  const body = Array.from(
    xml.documentElement.getElementsByTagName('*')
  ).find(x => x.localName === 'body');

  if (!body) {
    throw Error('الملف لا يحتوي على جداول Word');
  }

  const children = node =>
    Array.from(node.children);

  return children(body)
    .filter(x => x.localName === 'tbl')
    .map(tbl =>
      children(tbl)
        .filter(x => x.localName === 'tr')
        .map(row =>
          children(row)
            .filter(x => x.localName === 'tc')
            .map(cell =>
              children(cell)
                .filter(x => x.localName === 'p')
                .map(p =>
                  Array.from(
                    p.getElementsByTagName('*')
                  )
                    .filter(x => x.localName === 't')
                    .map(x => x.textContent)
                    .join('')
                )
                .join('\n')
                .trim()
            )
        )
    );
}

// =======================================
// تحويل الجداول إلى أعمال المدير
// =======================================

function rowsFromWord(tables, departmentNames) {
  const annual = [
    1, 3, 5, 7, 10,
    13, 16, 20, 22, 25
  ];

  const fixed = [
    2, 4, 6, 8, 11,
    14, 17, 21, 23, 26
  ];

  const months = [
    'آذار',
    'نيسان',
    'أيار',
    'حزيران',
    'تموز',
    'آب',
    'أيلول',
    'تشرين الأول',
    'تشرين الثاني',
    'كانون الأول',
    'كانون الثاني',
    'شباط'
  ];

  const extras = {
    9: [
      3,
      'محور الرقابة المالية وإدارة المخاطر'
    ],
    12: [
      4,
      'محور قياس فعالية التسويق واستقطاب المتعلمين'
    ],
    15: [
      5,
      'محور متابعة الأندية والبرامج اللامنهجية'
    ],
    18: [
      6,
      'محور متابعة خدمات الإرشاد والدعم'
    ],
    19: [
      7,
      'نطاق الإشراف الإداري'
    ],
    24: [
      9,
      'المجالات التي تخضع لمتابعة الجودة'
    ]
  };

  const official = [
    'معاملات وزارة التربية والتعليم العالي',
    'صندوق تعويضات أفراد الهيئة التعليمية',
    'الصندوق الوطني للضمان الاجتماعي',
    'وزارة المالية والالتزامات الضريبية',
    'معاملات صحية وموافقات رسمية أخرى'
  ];

  const kinds = {
    plan: 'أعمال سنوية',
    routine: 'تقارير دورية ثابتة',
    oversight: 'محاور المتابعة والرقابة',
    government: 'معاملة رسمية',
    calendar: 'روزنامة مركزية',
    central: 'تقارير دورية ثابتة للمدير العام'
  };

  const output = [];

  function add(tableIndex, kind, department, section) {
    const table = tables[tableIndex];

    if (!table) {
      throw Error(
        `جدول رقم ${tableIndex + 1} غير موجود`
      );
    }

    for (let i = 1; i < table.length; i++) {
      let cells = table[i];

      if (!cells.some(Boolean)) continue;

      const result = {
        id:
          `WORD-T${String(tableIndex).padStart(2, '0')}` +
          `-R${String(i).padStart(2, '0')}`,

        kind:
          kind === 'central'
            ? 'routine'
            : kind,

        department,
        section,

        originalTable: tableIndex + 1,
        originalRow: i + 1,

        frequency: kinds[kind],
        status: 'لم يُستلم'
      };

      if (
        kind === 'plan' ||
        kind === 'routine'
      ) {
        if (tableIndex === 3) {
          cells = [0, 2, 3, 4, 5, 6]
            .map(j => cells[j]);
        }

        if (cells.length !== 6) {
          throw Error(
            'تغيّر عدد أعمدة ملف Word'
          );
        }

        const [
          deadlineText,
          title,
          evidence,
          sourceAssignee,
          managerAction,
          followUp
        ] = cells;

        Object.assign(result, {
          deadlineText,
          title,
          evidence,
          sourceAssignee,
          managerAction,
          followUp
        });

      } else if (kind === 'oversight') {
        if (![2, 3, 4].includes(cells.length)) {
          throw Error('خطأ في أعمدة المتابعة');
        }

        result.title = cells[0];

        if (cells[1]) {
          result.description = cells[1];
        }

        if (cells[2]) {
          result.evidence = cells[2];
        }

        if (cells[3]) {
          result.followUp = cells[3];
        }

      } else if (kind === 'government') {
        const [
          title,
          deadlineText,
          evidence,
          sourceAssignee
        ] = cells;

        Object.assign(result, {
          title,
          deadlineText,
          evidence,
          sourceAssignee
        });

      } else if (kind === 'calendar') {
        const [
          deadlineText,
          title,
          sourceAssignee,
          managerAction
        ] = cells;

        Object.assign(result, {
          title,
          deadlineText,
          sourceAssignee,
          managerAction,
          month: months[tableIndex - 32]
        });

      } else {
        const [
          deadlineText,
          title,
          evidence,
          sourceAssignee,
          managerAction
        ] = cells;

        Object.assign(result, {
          deadlineText,
          title,
          evidence,
          sourceAssignee,
          managerAction
        });
      }

      result.reference =
        `أعمال المدير العام.docx — ${section} — ` +
        `الجدول ${tableIndex + 1}، الصف ${i + 1}`;

      if (!result.title) {
        throw Error(
          'أحد أعمال Word لا يحمل عنواناً'
        );
      }

      output.push(result);
    }
  }

  // أعمال الأقسام السنوية والثابتة
  for (let i = 0; i < 10; i++) {
    add(
      annual[i],
      'plan',
      departmentNames[i],
      'جدول المتابعة السنوية'
    );

    add(
      fixed[i],
      'routine',
      departmentNames[i],
      'التقارير الدورية الثابتة'
    );
  }

  // محاور الرقابة
  for (
    const [number, [index, title]]
    of Object.entries(extras)
  ) {
    add(
      Number(number),
      'oversight',
      departmentNames[index],
      title
    );
  }

  // المعاملات الرسمية
  for (let i = 27; i <= 31; i++) {
    add(
      i,
      'government',
      'المعاملات والملفات الرسمية مع الدولة',
      official[i - 27]
    );
  }

  // الروزنامة المركزية
  for (let i = 32; i <= 43; i++) {
    add(
      i,
      'calendar',
      'الإدارة العامة والقيادة المدرسية',
      'الروزنامة السنوية المركزية للمدير العام'
    );
  }

  add(
    44,
    'central',
    'الإدارة العامة والقيادة المدرسية',
    'التقارير الدورية الثابتة للمدير العام'
  );

  return output;
}

// =======================================
// مؤشرات لوحة القيادة
// =======================================

function numbers(list) {
  return {
    total: list.length,
    done: list.filter(isDone).length,
    late: list.filter(isLate).length,
    soon: list.filter(isSoon).length
  };
}

function counters(list) {
  const s = numbers(list);

  const items = [
    ['إجمالي البنود', s.total],
    ['المقفلة', s.done],
    ['المتأخرة بعد تحديد تاريخ', s.late],
    ['المستحقة خلال 14 يوماً', s.soon]
  ];

  return `
    <div class="kpis">
      ${items.map(([name, value]) => `
        <div class="kpi">
          <span>${name}</span>
          <strong>${value}</strong>
        </div>
      `).join('')}
    </div>
  `;
}

// =======================================
// بطاقات الأقسام الملونة
// =======================================

function cards() {
  return `
    <div class="departments">

      ${departments.map(d => {

        const s = numbers(
          tasks.filter(t => t.department === d)
        );

        const percent = s.total
          ? Math.round(100 * s.done / s.total)
          : 0;

        return `
          <button
            class="dept"
            data-theme="${theme(d)}"
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

// =======================================
// جدول عرض الأعمال
// =======================================

function rows(list) {
  if (!list.length) {
    return `
      <tr>
        <td colspan="6">
          لا توجد أعمال مطابقة.
        </td>
      </tr>
    `;
  }

  return list.map(raw => {
    const t = current(raw);

    return `
      <tr>
        <td>
          <b>${e(t.title)}</b>

          <small>
            ${e(t.section)} |
            ${e(t.id)}
          </small>
        </td>

        <td>${e(t.frequency)}</td>

        <td class="${isLate(raw) ? 'late' : ''}">
          ${e(
            t.deadlineText ||
            (t.date ? dateView(t.date) : 'حسب الخطة')
          )}
        </td>

        <td>
          ${e(
            t.assignee ||
            t.sourceAssignee ||
            '—'
          )}
        </td>

        <td>
          <span
            class="tag"
            data-status="${e(t.status)}">
            ${e(t.status)}
          </span>
        </td>

        <td>
          <button
            class="row-action"
            data-task="${e(t.id)}">
            التفاصيل
          </button>
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
            <th>العمل كما في Word</th>
            <th>النوع</th>
            <th>الموعد الأصلي</th>
            <th>المسؤول</th>
            <th>الحالة</th>
            <th>عرض</th>
          </tr>
        </thead>

        <tbody>
          ${rows(list)}
        </tbody>
      </table>
    </div>

    <p class="muted">
      ${list.length} بند
    </p>
  `;
}

// =======================================
// البحث والتصفية
// =======================================

function matches(list) {
  const q = query.trim().toLowerCase();

  return list.filter(raw => {
    const t = current(raw);

    const text = [
      t.title,
      t.department,
      t.section,
      t.evidence,
      t.sourceAssignee,
      t.assignee,
      t.managerAction,
      t.followUp,
      t.deadlineText
    ].join(' ').toLowerCase();

    return (
      text.includes(q) &&
      (!filterStatus || t.status === filterStatus) &&
      (category === 'all' || t.kind === category)
    );
  });
}

function taskArea(list, tabs = false) {
  const types = [
    ['all', 'جميع الأعمال'],
    ['plan', 'الأعمال السنوية'],
    ['routine', 'التقارير الثابتة'],
    ['oversight', 'محاور المتابعة'],
    ['calendar', 'الروزنامة المركزية'],
    ['government', 'المعاملات الرسمية']
  ];

  return `
    ${tabs ? `
      <div class="tabs">

        ${types.map(([code, label]) => `
          <button
            class="tab ${category === code ? 'active' : ''}"
            data-tab="${code}">
            ${label}
          </button>
        `).join('')}

      </div>
    ` : ''}

    <div class="filters">

      <input
        id="search"
        placeholder="بحث في الأعمال والتقارير"
        value="${e(query)}">

      <select id="statusFilter">
        <option value="">جميع الحالات</option>

        ${STATUSES.map(s => `
          <option ${filterStatus === s ? 'selected' : ''}>
            ${e(s)}
          </option>
        `).join('')}

      </select>
    </div>

    <div class="panel" id="results">
      ${table(matches(list))}
    </div>
  `;
}

function baseList() {
  if (view === 'department') {
    return tasks.filter(t =>
      t.department === department
    );
  }

  if (view === 'calendar') {
    return tasks.filter(t =>
      t.kind === 'calendar' &&
      (!month || t.month === month)
    );
  }

  return tasks.filter(t => t.kind === view);
}

function updateTable() {
  const box = el('results');

  if (box) {
    box.innerHTML = table(
      matches(baseList())
    );
  }
}

// =======================================
// رسم صفحات البرنامج
// =======================================

function render() {
  content.dataset.theme =
    view === 'department'
      ? theme(department)
      : 'neutral';

  el('pageTitle').textContent =
    view === 'department'
      ? department
      : TITLES[view];

  document.querySelectorAll('#menu button')
    .forEach(button => {
      button.classList.toggle(
        'active',
        button.dataset.view === view
      );
    });

  if (view === 'dashboard') {
    content.innerHTML = `
      <section class="hero">
        <h2>خطة أعمال المدير العام</h2>

        <p>
          وفق جداول ثانوية الجنان الأصلية،
          مع مهام المدير والمتابعة والأدلة.
        </p>
      </section>

      <div class="note">
        مواعيد Word محفوظة كما وردت.
        حدّد التاريخ الكامل في تفاصيل المهمة
        قبل إضافتها إلى التقويم.
      </div>

      ${counters(tasks)}

      <div class="section-title">
        <h2>الأقسام الإدارية</h2>
      </div>

      ${cards()}

      <div class="section-title">
        <h2>من الروزنامة المركزية</h2>
      </div>

      <div class="panel">
        ${table(
          tasks.filter(t => t.kind === 'calendar')
            .slice(0, 15)
        )}
      </div>
    `;

  } else if (view === 'departments') {
    content.innerHTML = `
      <section class="hero">
        <h2>الأقسام الإدارية</h2>
        <p>
          اختر القسم للانتقال إلى أعماله
          السنوية والثابتة.
        </p>
      </section>

      ${cards()}
    `;

  } else if (view === 'department') {
    const list = baseList();

    content.innerHTML = `
      <section class="hero">
        <h2>${e(department)}</h2>
        <p>
          الأعمال، المسؤولون، المواعيد،
          والتقارير المطلوبة.
        </p>
      </section>

      ${counters(list)}
      ${taskArea(list, true)}
    `;

  } else if (view === 'calendar') {
    content.innerHTML = `
      <div class="panel">
        <p>
          الروزنامة المركزية
          (69 بنداً كما في Word).
        </p>

        <button
          class="outline"
          data-action="ics">
          تصدير المواعيد المؤرّخة ICS
        </button>

        <div class="months">

          <button
            data-month=""
            class="${!month ? 'active' : ''}">
            كل الأشهر
          </button>

          ${MONTHS.map(m => `
            <button
              data-month="${e(m)}"
              class="${m === month ? 'active' : ''}">
              ${e(m)}
            </button>
          `).join('')}

        </div>
      </div>

      ${taskArea(baseList())}
    `;

  } else if (
    view === 'routine' ||
    view === 'government'
  ) {
    content.innerHTML = `
      <section class="hero">
        <h2>${e(TITLES[view])}</h2>
      </section>

      ${counters(baseList())}
      ${taskArea(baseList())}
    `;

  } else if (view === 'settings') {
    renderSettings();
  }
}

// =======================================
// نافذة تفاصيل المهمة
// =======================================

function openTask(id) {
  selected = byId(id);

  if (!selected) return;

  const t = current(selected);

  dialog.dataset.theme = theme(t.department);

  el('taskTitle').textContent = t.title;

  const fields = [
    ['القسم', t.department],
    ['جدول Word', t.section],
    ['التصنيف', t.frequency],
    ['الفترة/الموعد الأصلي', t.deadlineText],
    ['تفاصيل العمل', t.description],
    ['مصادر المعلومات والأدلة', t.evidence],
    ['المسؤولون', t.sourceAssignee],
    ['مهام المدير', t.managerAction],
    ['المتابعة اللاحقة', t.followUp],
    ['مرجع Word', t.reference]
  ];

  el('taskInfo').innerHTML = fields
    .filter(([, value]) => value)
    .map(([key, value]) => `
      <div>
        <b>${e(key)}:</b>
        ${e(value)}
      </div>
    `).join('');

  const inputs = [
    ['taskStatus', 'status'],
    ['taskAssignee', 'assignee'],
    ['taskDate', 'date'],
    ['taskReceived', 'receivedAt'],
    ['taskDecision', 'decision'],
    ['taskNotes', 'notes'],
    ['taskProof', 'proofUrl']
  ];

  for (const [id, key] of inputs) {
    el(id).value = t[key] || '';
  }

  dialog.showModal();
}

function saveTask() {
  if (!selected) return false;

  const change = {
    ...(changes[selected.id] || {}),

    status: el('taskStatus').value,
    assignee: el('taskAssignee').value.trim(),
    date: el('taskDate').value,
    receivedAt: el('taskReceived').value,
    decision: el('taskDecision').value.trim(),
    notes: el('taskNotes').value.trim(),
    proofUrl: el('taskProof').value.trim(),
    updatedAt: new Date().toISOString()
  };

  if (
    change.proofUrl &&
    !change.proofUrl.startsWith('https://')
  ) {
    notice('يجب أن يبدأ الرابط بـ https://');
    return false;
  }

  const next = {
    ...changes,
    [selected.id]: change
  };

  if (!save(STORE, next)) {
    notice('تعذّر حفظ التعديلات في المتصفح');
    return false;
  }

  changes = next;
  render();

  return true;
}

// =======================================
// النسخ الاحتياطي
// =======================================

function download(name, data, type) {
  const blob = URL.createObjectURL(
    new Blob([data], {type})
  );

  const a = document.createElement('a');

  a.href = blob;
  a.download = name;
  a.click();

  setTimeout(() => {
    URL.revokeObjectURL(blob);
  }, 1000);
}

function backup() {
  download(
    'jinan-backup.json',
    JSON.stringify({
      changes,
      googleConfig,
      createdAt: new Date().toISOString()
    }, null, 2),
    'application/json'
  );
}

async function importBackup(file) {
  if (!file) return;

  try {
    const data = JSON.parse(
      await file.text()
    );

    if (
      !data.changes ||
      typeof data.changes !== 'object'
    ) {
      throw Error('الملف غير صالح');
    }

    const ids = new Set(
      tasks.map(t => t.id)
    );

    changes = Object.fromEntries(
      Object.entries(data.changes)
        .filter(([id]) => ids.has(id))
    );

    if (!save(STORE, changes)) {
      throw Error('تعذّر حفظ التعديلات');
    }

    render();
    notice('تم استيراد النسخة الاحتياطية');

  } catch (err) {
    notice(err.message);
  }
}

// =======================================
// فتح المهمة في Google Calendar
// =======================================

function draftCalendar() {
  if (!saveTask()) return;

  const t = current(selected);

  if (!t.date) {
    notice('حدّد تاريخاً أولاً');
    return;
  }

  const from = t.date.replaceAll('-', '');

  const to = new Date(
    Date.parse(t.date + 'T00:00:00Z') + 86400000
  ).toISOString().slice(0, 10).replaceAll('-', '');

  const params = new URLSearchParams({
    action: 'TEMPLATE',
    text: t.title,
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

// =======================================
// تصدير مواعيد ICS
// =======================================

function exportICS() {
  const clean = s => String(s || '')
    .replaceAll('\\', '\\\\')
    .replaceAll(',', '\\,')
    .replaceAll(';', '\\;')
    .replaceAll('\n', '\\n');

  const output = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Al Jinan High School//AR',
    'CALSCALE:GREGORIAN'
  ];

  for (const source of tasks) {
    const t = current(source);

    if (!t.date) continue;

    const a = t.date.replaceAll('-', '');

    const b = new Date(
      Date.parse(t.date + 'T00:00:00Z') + 86400000
    ).toISOString().slice(0, 10).replaceAll('-', '');

    output.push(
      'BEGIN:VEVENT',

      `UID:${t.id}@jinan-management`,

      `DTSTAMP:${new Date().toISOString()
        .replace(/[-:]/g, '')
        .replace(/\.\d{3}Z$/, 'Z')}`,

      `DTSTART;VALUE=DATE:${a}`,
      `DTEND;VALUE=DATE:${b}`,

      `SUMMARY:${clean(t.title)}`,
      `DESCRIPTION:${clean(t.department)}`,

      'END:VEVENT'
    );
  }

  output.push('END:VCALENDAR');

  download(
    'jinan-calendar.ics',
    output.join('\r\n'),
    'text/calendar;charset=utf-8'
  );
}

// =======================================
// إعدادات Google
// =======================================

function renderSettings() {
  content.innerHTML = `
    <div class="settings">

      <div class="panel">

        <h2>Google Calendar</h2>

        <label>
          OAuth Client ID
          <input
            id="clientId"
            value="${e(googleConfig.clientId)}">
        </label>

        <label>
          Calendar ID
          <input
            id="calendarId"
            value="${e(googleConfig.calendarId)}">
        </label>

        <div class="actions">
          <button
            class="primary"
            data-action="saveGoogle">
            حفظ الإعدادات
          </button>

          <button
            class="outline"
            data-action="connectGoogle">
            توصيل Google
          </button>
        </div>

        <p class="note">
          فعّل Calendar API وOAuth Web Client،
          وأضف عنوان GitHub Pages إلى
          Authorized JavaScript origins.
          لا تضع Client Secret هنا.
        </p>

      </div>

      <div class="panel">

        <h2>Google Sheets</h2>

        <label>
          Spreadsheet ID
          <input
            id="sheetId"
            value="${e(googleConfig.sheetId)}">
        </label>

        <div class="actions">

          <button
            class="outline"
            data-action="pushSheet">
            رفع التعديلات
          </button>

          <button
            class="outline"
            data-action="pullSheet">
            استيراد التعديلات
          </button>

        </div>

        <p>
          يجب تفعيل Google Sheets API.
          المزامنة يدوية.
        </p>

      </div>
    </div>

    <div class="panel">

      <h2>النسخ الاحتياطي</h2>

      <button
        class="outline"
        data-action="backup">
        تنزيل نسخة احتياطية
      </button>

      <button
        class="outline"
        data-action="restore">
        استيراد نسخة احتياطية
      </button>

      <button
        class="outline"
        data-action="ics">
        ملف ICS
      </button>

      <input
        id="backupFile"
        type="file"
        accept=".json,application/json"
        hidden>

      <p class="note">
        GitHub Pages عام.
        لا تنشر بيانات الطلاب والموظفين الخاصة.
        الحفظ المحلي ليس نظام صلاحيات مشتركة.
      </p>

    </div>
  `;
}

function saveGoogle() {
  googleConfig = {
    clientId: el('clientId').value.trim(),
    calendarId:
      el('calendarId').value.trim() || 'primary',
    sheetId: el('sheetId').value.trim()
  };

  if (!save(GOOGLE_STORE, googleConfig)) {
    notice('تعذّر حفظ الإعدادات');
    return;
  }

  accessToken = '';

  notice('تم حفظ إعدادات Google');
}

// =======================================
// Google OAuth
// =======================================

function scopes() {
  return [
    'https://www.googleapis.com/auth/calendar.events',

    ...(googleConfig.sheetId
      ? ['https://www.googleapis.com/auth/spreadsheets']
      : [])
  ].join(' ');
}

async function googleToken() {
  if (!googleConfig.clientId) {
    throw Error('أدخل Client ID');
  }

  if (!window.google?.accounts?.oauth2) {
    throw Error('مكتبة Google غير متاحة');
  }

  const scope = scopes();

  if (
    accessToken &&
    Date.now() < tokenExpiry - 60000 &&
    scope === tokenScope
  ) {
    return accessToken;
  }

  return new Promise((resolve, reject) => {
    const client =
      google.accounts.oauth2.initTokenClient({
        client_id: googleConfig.clientId,
        scope,

        callback: result => {
          if (result.error) {
            reject(Error(result.error));
            return;
          }

          accessToken = result.access_token;

          tokenExpiry =
            Date.now() +
            Number(result.expires_in || 3600) * 1000;

          tokenScope = scope;

          resolve(accessToken);
        },

        error_callback: err => {
          reject(
            Error(err.message || 'فشل الاتصال')
          );
        }
      });

    client.requestAccessToken();
  });
}

async function googleRequest(
  url,
  method = 'GET',
  body
) {
  const token = await googleToken();

  const response = await fetch(url, {
    method,

    headers: {
      Authorization: `Bearer ${token}`,

      ...(body
        ? {'Content-Type': 'application/json'}
        : {})
    },

    ...(body
      ? {body: JSON.stringify(body)}
      : {})
  });

  const result = await response.json()
    .catch(() => ({}));

  if (!response.ok) {
    throw Error(
      result.error?.message ||
      'فشل اتصال Google'
    );
  }

  return result;
}

// =======================================
// المزامنة المباشرة مع Google Calendar
// =======================================

async function syncCalendar() {
  if (!selected || !saveTask()) return;

  const t = current(selected);

  if (!t.date) {
    notice('حدّد تاريخ المهمة');
    return;
  }

  const base =
    'https://www.googleapis.com/calendar/v3/calendars/' +
    encodeURIComponent(
      googleConfig.calendarId || 'primary'
    ) +
    '/events';

  const next = new Date(
    Date.parse(t.date + 'T00:00:00Z') + 86400000
  ).toISOString().slice(0, 10);

  const body = {
    summary: t.title,

    description:
      `القسم: ${t.department}\n` +
      `المسؤول: ${t.assignee || t.sourceAssignee || ''}\n` +
      `دور المدير: ${t.managerAction || ''}`,

    start: {date: t.date},
    end: {date: next},

    reminders: {
      useDefault: false,
      overrides: [
        {method: 'popup', minutes: 1440}
      ]
    }
  };

  try {
    const eventId =
      changes[t.id]?.calendarEventId;

    let event;

    if (eventId) {
      try {
        event = await googleRequest(
          base + '/' + encodeURIComponent(eventId),
          'PUT',
          body
        );

      } catch (err) {
        if (!/404|not found/i.test(err.message)) {
          throw err;
        }

        event = await googleRequest(
          base,
          'POST',
          body
        );
      }

    } else {
      event = await googleRequest(
        base,
        'POST',
        body
      );
    }

    changes[t.id] = {
      ...changes[t.id],
      calendarEventId: event.id,
      updatedAt: new Date().toISOString()
    };

    if (!save(STORE, changes)) {
      notice(
        'تمت الإضافة للتقويم ولكن تعذّر حفظ معرّف الحدث محلياً'
      );
      return;
    }

    notice('تمت مزامنة Google Calendar');

  } catch (err) {
    notice(err.message);
  }
}

// =======================================
// Google Sheets
// =======================================

function sheetUrl() {
  if (!googleConfig.sheetId) {
    throw Error('أدخل Spreadsheet ID');
  }

  return (
    'https://sheets.googleapis.com/v4/spreadsheets/' +
    encodeURIComponent(googleConfig.sheetId)
  );
}

async function pullRemote() {
  const base = sheetUrl();
  let result;

  try {
    result = await googleRequest(
      base + '/values/' +
      encodeURIComponent(
        "'Jinan_Status'!A2:J10000"
      )
    );

  } catch (err) {
    if (
      /Unable to parse range|not found/i
        .test(err.message)
    ) {
      return {};
    }

    throw err;
  }

  const existing = new Set(
    tasks.map(t => t.id)
  );

  const updates = {};

  for (const row of result.values || []) {
    const [
      id,
      status,
      assignee,
      date,
      receivedAt,
      decision,
      notes,
      proofUrl,
      calendarEventId,
      updatedAt
    ] = row;

    if (existing.has(id)) {
      updates[id] = {
        status,
        assignee,
        date,
        receivedAt,
        decision,
        notes,
        proofUrl,
        calendarEventId,
        updatedAt
      };
    }
  }

  return updates;
}

function reconcile(local, remote) {
  const all = {...local};

  for (
    const [id, value]
    of Object.entries(remote)
  ) {
    if (
      !all[id] ||
      String(value.updatedAt || '') >
      String(all[id].updatedAt || '')
    ) {
      all[id] = value;
    }
  }

  return all;
}

async function pullSheet() {
  try {
    changes = reconcile(
      changes,
      await pullRemote()
    );

    if (!save(STORE, changes)) {
      throw Error('فشل الحفظ المحلي');
    }

    render();
    notice('تم الاستيراد');

  } catch (err) {
    notice(err.message);
  }
}

async function pushSheet() {
  try {
    const base = sheetUrl();

    const meta = await googleRequest(
      base + '?fields=sheets.properties.title'
    );

    const exists = meta.sheets?.some(
      x => x.properties.title === 'Jinan_Status'
    );

    if (!exists) {
      await googleRequest(
        base + ':batchUpdate',
        'POST',
        {
          requests: [{
            addSheet: {
              properties: {
                title: 'Jinan_Status'
              }
            }
          }]
        }
      );
    }

    changes = reconcile(
      changes,
      await pullRemote()
    );

    const keys = [
      'id',
      'status',
      'assignee',
      'date',
      'receivedAt',
      'decision',
      'notes',
      'proofUrl',
      'calendarEventId',
      'updatedAt'
    ];

    const ids = new Set(
      tasks.map(t => t.id)
    );

    const values = [
      keys,

      ...Object.entries(changes)
        .filter(([id]) => ids.has(id))
        .map(([id, t]) => [
          id,
          ...keys.slice(1)
            .map(key => t[key] || '')
        ])
    ];

    const range =
      `'Jinan_Status'!A1:J${values.length}`;

    await googleRequest(
      base + '/values/' +
      encodeURIComponent(range) +
      '?valueInputOption=RAW',

      'PUT',

      {
        range,
        majorDimension: 'ROWS',
        values
      }
    );

    const clearRange =
      `'Jinan_Status'!A${values.length + 1}:J10000`;

    await googleRequest(
      base + '/values/' +
      encodeURIComponent(clearRange) +
      ':clear',

      'POST',
      {}
    );

    save(STORE, changes);
    notice('تم رفع التعديلات');

  } catch (err) {
    notice(err.message);
  }
}

// =======================================
// التنقل والأحداث
// =======================================

function navigate(next) {
  view = next;
  category = 'all';
  query = '';
  filterStatus = '';

  render();
}

document.addEventListener(
  'click',
  async event => {
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
      category = button.dataset.tab;
      render();
      return;
    }

    if (button.dataset.month !== undefined) {
      month = button.dataset.month;
      render();
      return;
    }

    switch (button.dataset.action) {
      case 'backup':
        backup();
        break;

      case 'restore':
        el('backupFile').click();
        break;

      case 'ics':
        exportICS();
        break;

      case 'saveGoogle':
        saveGoogle();
        break;

      case 'connectGoogle':
        try {
          await googleToken();
          notice('تم الاتصال بـ Google');
        } catch (err) {
          notice(err.message);
        }
        break;

      case 'pullSheet':
        await pullSheet();
        break;

      case 'pushSheet':
        await pushSheet();
        break;
    }
  }
);

document.addEventListener('input', event => {
  if (event.target.id === 'search') {
    query = event.target.value;
    updateTable();
  }
});

document.addEventListener('change', event => {
  if (event.target.id === 'statusFilter') {
    filterStatus = event.target.value;
    updateTable();
  }

  if (event.target.id === 'backupFile') {
    importBackup(event.target.files[0]);
  }
});

el('exportBtn').addEventListener(
  'click',
  backup
);

el('closeDialog').addEventListener(
  'click',
  () => dialog.close()
);

el('taskForm').addEventListener(
  'submit',
  event => {
    event.preventDefault();

    if (saveTask()) {
      dialog.close();
      notice('تم حفظ التغييرات');
    }
  }
);

el('calendarDraft').addEventListener(
  'click',
  draftCalendar
);

el('calendarSync').addEventListener(
  'click',
  syncCalendar
);

// =======================================
// تشغيل البرنامج
// =======================================

(async function start() {
  try {
    const response = await fetch(
      './data/tasks.json',
      {cache: 'no-store'}
    );

    if (!response.ok) {
      throw Error(
        `لم يتم العثور على data/tasks.json (${response.status})`
      );
    }

    const config = await response.json();

    departments = config.departments;

    if (!Array.isArray(departments)) {
      throw Error('الأقسام غير موجودة في JSON');
    }

    if (Array.isArray(config.tasks)) {
      tasks = config.tasks;

    } else {
      const xml = await wordXml(
        config.sourceDocx
      );

      tasks = rowsFromWord(
        extractTables(xml),
        departments
      );
    }

    if (
      config.expectedTasks &&
      tasks.length !== config.expectedTasks
    ) {
      throw Error(
        `المهام ${tasks.length} بدلاً من ` +
        `${config.expectedTasks}. راجع نسخة Word`
      );
    }

    render();

  } catch (err) {
    console.error(err);

    content.innerHTML = `
      <div class="note">
        ${e(err.message)}.
        تأكد من رفع Word وJSON
        داخل مجلد data.
      </div>
    `;
  }
})();
