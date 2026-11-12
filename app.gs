/*
 * ثانوية الجنان — منظومة متابعة المدير العام (الإصدار الثالث)
 * البيانات الأصلية من: data/tasks.json (476 بنداً من جداول Word)
 * تحديثات الأعمال والأوصاف تحفظ محلياً؛ Google Sheets اختياري للمشاركة.
 * GitHub Pages ليس بوابة مصادقة أو نظام صلاحيات للموظفين.
 */
'use strict';

const $ = id => document.getElementById(id);
const e = value => String(value ?? '').replace(/[&<>"']/g, x => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[x]));
const STORE = 'jinan-manager-changes-word-v2';
const CONFIG_STORE = 'jinan-manager-config-v1';
const EXTRA_STORE = 'jinan-extra-tasks-v3';
const JOB_STORE = 'jinan-job-descriptions-v3';
const PENDING_STORE = 'jinan-pending-cloud-v3';
const MODE_STORE = 'jinan-workspace-mode-v4';
const MODE_DEPT_STORE = 'jinan-workspace-dept-v4';
const SHEET_NAME = 'Jinan_Updates';
const SCOPE = {
  calendar: 'https://www.googleapis.com/auth/calendar.events',
  sheets: 'https://www.googleapis.com/auth/spreadsheets',
  gmail: 'https://www.googleapis.com/auth/gmail.send'
};
const STATUSES = ['لم يُستلم','تم الاستلام','قيد المراجعة','مُقفل','متعثر'];
const MONTHS = ['تموز','آب','أيلول','تشرين الأول','تشرين الثاني','كانون الأول','كانون الثاني','شباط','آذار','نيسان','أيار','حزيران'];
const TITLES = {dashboard:'لوحة القيادة',organization:'الهيكلية الإدارية',building:'خريطة المبنى والصفوف',departments:'الأقسام الإدارية',department:'أعمال القسم',jobs:'الأوصاف الوظيفية',calendar:'الروزنامة السنوية',routine:'الأعمال الثابتة',government:'المعاملات الرسمية',settings:'الإعدادات والربط مع Google'};
const KIND_LABEL = {plan:'عمل سنوي',routine:'تقرير ثابت',oversight:'متابعة ورقابة',calendar:'روزنامة',government:'معاملة رسمية'};

function loadLocal(key, fallback) {
  try {const v = localStorage.getItem(key); return v ? JSON.parse(v) : fallback;}
  catch (error) {console.warn('التخزين المحلي غير متاح', error);return fallback;}
}
function saveLocal(key, value) {
  try {localStorage.setItem(key, JSON.stringify(value));return true;}
  catch (error) {console.error('تعذّر التخزين المحلي',error);return false;}
}
function stamp() {return new Date().toISOString();}
function stampView(value) {
  if (!value) return 'لم يُسجل بعد';
  if (/^\d{4}-\d{2}-\d{2}$/.test(String(value))) return `${value} (من النسخة السابقة؛ الوقت غير مسجل)`;
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value);
  return d.toLocaleString('ar-LB',{timeZone:'Asia/Beirut',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit'});
}
function today() {
  const fmt = new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Beirut',year:'numeric',month:'2-digit',day:'2-digit'});
  const p = fmt.formatToParts(new Date());
  const get = key => p.find(x=>x.type===key).value;
  return `${get('year')}-${get('month')}-${get('day')}`;
}
function dateView(date) {return date ? new Date(date+'T12:00:00Z').toLocaleDateString('ar-LB',{day:'numeric',month:'short',year:'numeric',timeZone:'UTC'}) : 'غير محدد';}
function safeUrl(value) {return /^https:\/\/[^\s]+$/i.test(String(value||''));}
function validEmail(value) {return /^[^\s@\r\n]+@[^\s@\r\n]+\.[^\s@\r\n]+$/.test(String(value||''));}
function notify(message, warning=false) {const box=$('toast');box.textContent=message;box.style.display='block';box.style.background=warning?'#80501b':'#143c51';clearTimeout(notify.timer);notify.timer=setTimeout(()=>box.style.display='none',5700);}
function indicator(message, mode='') {const box=$('syncIndicator');box.textContent=message;box.className='sync-indicator '+mode;}
function deptTheme(name='') {
  if(name.includes('الإشراف الإداري')||name.includes('الاشراف الاداري'))return 'administrative';
  if(name.includes('شؤون المتعلمين')||name.includes('شؤون الطلاب'))return 'students';
  if(name.includes('الإشراف التربوي'))return 'pedagogical';
  if(name.includes('الموارد البشرية'))return 'hr';
  if(name.includes('المالية')||name.includes('المحاسبة'))return 'finance';
  if(name.includes('الجودة'))return 'quality';
  return 'neutral';
}

let baseTasks=[];
let tasks=[];
let departments=[];
let changes=loadLocal(STORE,{});
let customTasks=loadLocal(EXTRA_STORE,{});
let jobs=loadLocal(JOB_STORE,{});
let pending=loadLocal(PENDING_STORE,[]);
let config={clientId:'',calendarId:'primary',sheetId:'',editorName:'',...loadLocal(CONFIG_STORE,{})};
let view='dashboard',department='',month='',tab='all',search='',statusFilter='',jobSearch='',jobDepartment='';
let selectedTask=null,editingTaskId=null,editingJobId=null;
let oauthToken='',tokenExpires=0,oauthGranted=new Set();
let syncBusy=false;
let workspaceMode=loadLocal(MODE_STORE,'manager');
let workspaceDepartment=loadLocal(MODE_DEPT_STORE,'');

function buildTasks(){
  const extra=Object.values(customTasks).filter(x=>!x.deleted);
  tasks=[...baseTasks,...extra];
}
function task(raw){return {...raw,...(changes[raw.id]||{})};}
function byId(id){return tasks.find(t=>t.id===id);}
function isDone(raw){return task(raw).status==='مُقفل';}
function isLate(raw){const t=task(raw);return Boolean(t.date && t.date<today() && !isDone(raw));}
function isSoon(raw){const t=task(raw);return !isDone(raw)&&t.date>=today()&&(Date.parse(t.date+'T00:00:00Z')-Date.parse(today()+'T00:00:00Z'))/86400000<=14;}
function stats(list){return {total:list.length,done:list.filter(isDone).length,late:list.filter(isLate).length,soon:list.filter(isSoon).length};}
function kpis(list){const s=stats(list);return `<div class="kpis">${[['الأعمال',s.total],['المقفلة',s.done],['المتأخرة (المؤرخة)',s.late],['خلال 14 يوماً',s.soon]].map(([a,b])=>`<div class="kpi"><span>${a}</span><strong>${b}</strong></div>`).join('')}</div>`;}
function cards(){return `<div class="departments">${departments.filter(d=>workspaceMode!=='employee'||d===workspaceDepartment).map(d=>{const s=stats(tasks.filter(t=>t.department===d)),percent=s.total?Math.round(100*s.done/s.total):0;return `<button class="dept" data-theme="${deptTheme(d)}" data-dept="${e(d)}"><strong>${e(d)}</strong><small>${s.total} بند · ${s.late} متأخر · ${percent}% منجز</small><span class="progress"><i style="width:${percent}%"></i></span></button>`;}).join('')}</div>`;}
function linkButton(url,label='فتح التقرير'){return safeUrl(url)?`<a class="open-doc" href="${e(url)}" target="_blank" rel="noopener noreferrer">${e(label)}</a>`:'';}
function rows(list){
  if(!list.length)return '<tr><td colspan="7">لا توجد بنود مطابقة.</td></tr>';
  return list.map(raw=>{
    const t=task(raw),delivered=t.proofUrl&&t.receivedAt;
    return `<tr>
      <td><b>${e(t.title)}</b><small>${e(t.section||KIND_LABEL[t.kind]||'عمل')} | ${e(t.id)}</small></td>
      <td>${e(t.frequency||KIND_LABEL[t.kind]||'—')}</td>
      <td class="${isLate(raw)?'late':''}">${e(t.deadlineText||(t.date?dateView(t.date):'تُحدّد'))}</td>
      <td><span class="tag" data-status="${e(t.status||'لم يُستلم')}">${e(t.status||'لم يُستلم')}</span></td>
      <td>${delivered?`<span class="receipt">${e(stampView(t.receivedAt))}</span><br>${linkButton(t.proofUrl)}`:t.proofUrl?linkButton(t.proofUrl):'<span class="muted">لم يصل رابط</span>'}</td>
      <td>${e(t.assignee||t.sourceAssignee||'—')}</td>
      <td><div class="quick-actions"><button class="row-action" data-task="${e(t.id)}" type="button">فتح</button><button class="row-action row-reminder " data-remind="${e(t.id)}" type="button">🔔 أضف إلى تقويمي</button></div></td>
      </tr>`;
  }).join('');
}
function table(list){return `<div class="table-wrap"><table><thead><tr><th>العمل</th><th>النوع</th><th>الموعد</th><th>الحالة</th><th>التقرير ووقت الوصول</th><th>المسؤول</th><th>إجراءات</th></tr></thead><tbody>${rows(list)}</tbody></table></div><p class="muted">${list.length} بند</p>`;}
function toolbar(){return `<div class="filters"><input id="search" placeholder="بحث عن تقرير أو عمل أو مسؤول" value="${e(search)}"><select id="statusFilter"><option value="">كل الحالات</option>${STATUSES.map(x=>`<option value="${e(x)}" ${statusFilter===x?'selected':''}>${e(x)}</option>`).join('')}</select></div>`;}
function filtered(list){const q=search.trim().toLocaleLowerCase('ar');return list.filter(raw=>{const t=task(raw),text=[t.title,t.department,t.section,t.deadlineText,t.evidence,t.assignee,t.sourceAssignee,t.managerAction,t.followUp,t.description].join(' ').toLocaleLowerCase('ar');return text.includes(q)&&(!statusFilter||t.status===statusFilter)&&(tab==='all'||t.kind===tab);});}
function taskArea(list,withTabs=false){const types=[['all','كل الأعمال'],['plan','الأعمال السنوية'],['routine','التقارير الثابتة'],['oversight','الرقابة والمتابعة'],['calendar','الروزنامة'],['government','المعاملات الرسمية']];return `${withTabs?`<div class="tabs">${types.map(([k,v])=>`<button class="tab ${tab===k?'active':''}" type="button" data-tab="${k}">${v}</button>`).join('')}</div>`:''}${toolbar()}<div class="panel" id="results">${table(filtered(list))}</div>`;}
function mainList(){
  const selected=view==='department'?tasks.filter(t=>t.department===department):view==='calendar'?tasks.filter(t=>t.kind==='calendar'&&(!month||t.month===month)):tasks.filter(t=>t.kind===view);
  return workspaceMode==='employee'&&workspaceDepartment ? selected.filter(t=>t.department===workspaceDepartment):selected;
}
function filteredByWorkspace(items){return workspaceMode==='employee'&&workspaceDepartment?items.filter(t=>t.department===workspaceDepartment):items;}
function setWorkspaceDisplay(){
  $('workspaceMode').value=workspaceMode;
  $('workspaceDepartment').innerHTML=departments.map(d=>`<option value="${e(d)}">${e(d)}</option>`).join('');
  if(!departments.includes(workspaceDepartment))workspaceDepartment=departments[0]||'';
  $('workspaceDepartment').value=workspaceDepartment;
  $('workspaceDepartment').hidden=workspaceMode!=='employee';
  $('refreshCloud').textContent='تحديث تقارير الأقسام من Google';
}

function updateRows(){const el=$('results');if(el)el.innerHTML=table(filtered(mainList()));}

/* الهيكلية المعتمدة في المخطط المرفق لعام 2026-2027.
   حقول الوصف الوظيفي تظل فارغة إلى أن تُعتمد رسمياً. */
const OFFICIAL_ORG = [
  {title:'الإشراف الإداري والجودة',person:'أ. سامي حلاق',group:'الإشراف الإداري والعمليات والخدمات المدرسية',children:[
    ['قسم الإعلام والتسويق','أ. سمية النجار'],['مسؤول مكننة ومعلوماتية',''],['ممرض',''],
    ['المشرف الصحي',''],['مسؤول مخزون',''],['مسؤول مشتريات',''],['مسؤول الصيانة',''],['حارس',''],['سائق','']
  ]},
  {title:'نائب المدير',person:'د. يكانة يكن',group:'الإدارة العامة والقيادة المدرسية',children:[]},
  {title:'شؤون الطلاب',person:'أ. مصطفى كمون',group:'شؤون المتعلمين',children:[]},
  {title:'المدقق المالي',person:'',group:'الشؤون المالية والمحاسبة',children:[]},
  {title:'المدير المالي',person:'د. يكانة يكن',group:'الشؤون المالية والمحاسبة',children:[
    ['محاسب','أ. فاطمة أبو زيد'],['أمين صندوق','أ. هنادي لاذقاني']
  ]},
  {title:'أمين السر',person:'',group:'أمانة السر والمراسلات والأرشفة',children:[]},
  {title:'موارد بشرية',person:'أ. صبا الخضر',group:'إدارة الموارد البشرية',children:[]},
  {title:'الإشراف التربوي',person:'أ. محمد صهيوني',group:'الإشراف التربوي والأكاديمي',children:[
    ['مسؤولة الأنشطة','أ. رحمة كبارة'],['عامل صحي اجتماعي','أ. هبة شعاع'],
    ['مسؤول مختبر العلوم','أ. ريما حمود'],['مسؤول مكتبة',''],
    ['مسؤول مختبر الكمبيوتر','أ. غنى درويش'],
    ['إشراف المرحلة الثانوية','أ. نهى الزهر'],
    ['إشراف الحلقتين الثالثة والرابعة','أ. زينة ولي الدين'],
    ['إشراف الحلقتين الثانية والأولى','أ. بشرى الزيلع'],
    ['إشراف مرحلة الروضات','أ. ماجدة حداد'],
    ['إشراف قسم الدعم التربوي','أ. رفاه حيدر']
  ]}
];
/* الوثيقتان الأصليتان (بدون إعادة رسم أو تغيير في النص أو الألوان). */
const OFFICIAL_DOCUMENTS = {
  organization: {
    title: 'الهيكلية الإدارية المعتمدة',
    subtitle: 'المخطط الرسمي لثانوية الجنان للعام الدراسي 2026–2027',
    image: './assets/organizational-structure-2026-2027.png',
    pdf: './assets/organizational-structure-2026-2027.pdf',
    alt: 'الهيكلية الإدارية الأصلية لثانوية الجنان: المدير العام والتبعية الإدارية والأقسام والمشرفون',
    hint: 'هذا المخطط مطابق للصفحة الرسمية المرفقة، بما في ذلك جميع الأسماء والأسهم والألوان.'
  },
  building: {
    title: 'خريطة توزيع الهيكلية الإدارية والصفوف',
    subtitle: 'المخطط الأصلي للطوابق من الأرضي إلى الخامس للعام الدراسي 2026–2027',
    image: './assets/building-floor-map-2026-2027.png',
    pdf: './assets/building-floor-map-2026-2027.pdf',
    alt: 'خريطة ثانوية الجنان الأصلية: مواقع الأقسام الإدارية والصفوف موزعة على الطوابق الستة',
    hint: 'اضغط على تكبير لقراءة أسماء الصفوف والمكاتب، أو افتح ملف PDF الأصلي للتكبير بالدقة الأعلى.'
  }
};
function officialDocumentPage(kind) {
  const d = OFFICIAL_DOCUMENTS[kind];
  const isOrg = kind === 'organization';
  return `<section class="hero">
      <h2>${e(d.title)}</h2><p>${e(d.subtitle)}</p>
    </section>
    <div class="document-switch" role="group" aria-label="اختيار المخطط">
      <button type="button" data-view="organization" class="${isOrg ? 'selected' : ''}">الهيكلية الإدارية الأصلية</button>
      <button type="button" data-view="building" class="${!isOrg ? 'selected' : ''}">خريطة المبنى والصفوف</button>
      <button type="button" data-view="jobs">الأوصاف الوظيفية</button>
    </div>
    <section class="panel document-panel" aria-label="${e(d.title)}">
      <div class="document-toolbar">
        <div><h2>${e(d.title)}</h2><p class="muted">${e(d.hint)}</p></div>
        <div class="document-buttons">
          <button type="button" class="outline" data-visual-zoom="25" aria-label="تكبير المخطط">+ تكبير</button>
          <button type="button" class="outline" data-visual-zoom="-25" aria-label="تصغير المخطط">− تصغير</button>
          <button type="button" class="outline" data-visual-zoom="reset">إعادة الحجم</button>
          <span id="visualZoomLabel" class="document-zoom-label" aria-live="polite">100%</span>
        </div>
      </div>
      <div class="official-document-scroll" tabindex="0" aria-label="اسحب أفقياً عند التكبير لرؤية تفاصيل المخطط">
        <img id="visualDocumentImage" src="${d.image}" data-zoom="100"
          alt="${e(d.alt)}" width="1819" height="2572" draggable="false" />
      </div>
      <div class="document-foot">
        <a href="${d.pdf}" target="_blank" rel="noopener noreferrer" class="primary document-link">فتح ملف PDF الأصلي بالدقة الكاملة</a>
        <a href="${d.image}" target="_blank" rel="noopener noreferrer" class="outline document-link">عرض الصورة بحجمها الكامل</a>
      </div>
    </section>
    <p class="note">تنبيه: مخططات المدرسة المنشورة على GitHub Pages ستكون متاحة للجمهور. يُفضّل نشرها هناك فقط إذا كانت المدرسة تسمح بعرضها علناً.</p>`;
}
function organization(){return officialDocumentPage('organization');}
function buildingMap(){return officialDocumentPage('building');}
function jobsPage(){
  const available=Object.values(jobs).filter(j=>!j.deleted && (!jobDepartment||j.department===jobDepartment) && [j.title,j.department,j.purpose,j.responsibilities,j.reportsTo].join(' ').toLowerCase().includes(jobSearch.toLowerCase()));
  return `<section class="hero"><h2>بطاقات الوصف الوظيفي</h2><p>أرشيف عملي يُستكمل من الوصف الوظيفي المعتمد لكل وظيفة في المؤسسة.</p></section>
    <div class="note">ملف Word الخاص بأعمال المدير لا يحتوي بطاقات توصيف كاملة لكل وظيفة. الحقول التالية فارغة إلى أن يملأها المسؤول ويعتمدها.</div>
    <div class="button-row"><h2>الوظائف المسجّلة (${available.length})</h2><button class="primary" data-action="addJob">+ إضافة وظيفة ووصفها</button></div>
    <div class="filters"><input id="jobSearch" placeholder="بحث باسم الوظيفة أو مسؤولياتها" value="${e(jobSearch)}"><select id="jobDepartmentFilter"><option value="">جميع الأقسام</option>${departments.map(d=>`<option value="${e(d)}" ${jobDepartment===d?'selected':''}>${e(d)}</option>`).join('')}</select></div>
    ${!available.length?'<div class="panel">لا توجد بطاقات وصف وظيفي في هذا القسم بعد. اضغط «إضافة وظيفة ووصفها» لإنشاء أول بطاقة.</div>':`<div class="job-grid">${available.map(j=>`<article class="job-card" data-theme="${deptTheme(j.department)}"><h3>${e(j.title)}</h3><div class="muted">${e(j.department)}</div><div class="job-fields">${[['يرفع تقاريره إلى',j.reportsTo],['الغاية من الوظيفة',j.purpose],['المسؤوليات',j.responsibilities],['الصلاحيات',j.authority],['المؤهلات',j.qualifications],['مؤشرات الأداء',j.kpis]].filter(([,v])=>v).map(([k,v])=>`<p><b>${k}: </b>${e(v)}</p>`).join('')||'<p class="muted">الوصف لم يُستكمل بعد.</p>'}</div><div class="quick-actions"><button class="row-action" data-job="${e(j.id)}">عرض وتعديل</button>${linkButton(j.docUrl,'الوصف المعتمد')}</div></article>`).join('')}</div>`}`;
}
function renderSettings(){
  $('content').innerHTML=`<section class="hero"><h2>الربط مع Google</h2><p>التقويم، إرسال تعليقات المدير بالبريد، ومشاركة تحديثات الأعمال على Google Sheets.</p></section>
    <div class="settings">
      <div class="panel"><h2>إعدادات حساب Google</h2>
        <label>OAuth Web Client ID<input id="clientId" value="${e(config.clientId)}" dir="ltr" placeholder="...apps.googleusercontent.com"></label>
        <label>Calendar ID<input id="calendarId" dir="ltr" value="${e(config.calendarId)}"></label>
        <label>اسم المستخدم / المرسل في سجل المتابعة<input id="editorName" value="${e(config.editorName)}" placeholder="اسم الموظف أو المدير"></label>
        <div class="actions"><button class="primary" data-action="saveConfig">حفظ الإعدادات</button><button class="outline" data-action="connectCalendar">ربط Google Calendar</button><button class="outline" data-action="connectMail">ربط Gmail للإرسال</button></div>
        <p class="field-hint">فعّل Calendar API وGmail API في Google Cloud، وأضف عنوان موقع GitHub Pages إلى Authorized JavaScript origins. إذن gmail.send حساس وقد يتطلب التحقق من التطبيق أو إضافتك إلى Test users.</p>
      </div>
      <div class="panel"><h2>مزامنة الأعمال بين الأجهزة</h2>
        <label>Google Spreadsheet ID<input id="sheetId" dir="ltr" value="${e(config.sheetId)}"></label>
        <div class="actions"><button class="outline" data-action="connectSheets">ربط Google Sheets</button><button class="outline" data-action="syncCloud">مزامنة الآن</button></div>
        <p class="field-hint">يُنشأ جدول Jinan_Updates في ملف Sheets، وتُرسل إليه التحديثات بعد ربط الحساب. يُحدّث الموقع البيانات عند المزامنة أو آلياً أثناء فتحه كل 90 ثانية ما دام الإذن سارياً. يجب مشاركة ملف Sheets فقط مع الحسابات المصرّح لها؛ هذا النموذج لا يفرض أدوار المدير/الموظف.</p>
        <div class="note">تاريخ التسليم يُسجَّل من ساعة جهاز من أضاف رابط التقرير، وليس ختم وصول معتمداً من خادم.</div>
      </div>
    </div>
    <div class="panel"><h2>نسخ احتياطي واستعادة</h2><div class="actions"><button class="outline" data-action="backup">تنزيل نسخة احتياطية</button><button class="outline" data-action="restore">استيراد نسخة احتياطية</button><button class="outline" data-action="ics">تصدير ICS</button><input id="backupFile" type="file" accept=".json,application/json" hidden></div>
    <div class="note">لا تضع روابط حساسة أو أسماء متعلمين أو مستندات شخصية في ملفات GitHub العامة. وحماية ملفات Google Drive تعتمد على صلاحياتها الفعلية، وليس على إخفاء الرابط في الموقع.</div></div>`;
}
function render(){
  setWorkspaceDisplay();
  $('content').dataset.theme=view==='department'?deptTheme(department):'neutral';
  $('pageTitle').textContent=view==='department'?department:TITLES[view];
  document.querySelectorAll('#menu button').forEach(b=>b.classList.toggle('active',b.dataset.view===view));
  if(view==='dashboard') {
    const central=tasks.filter(t=>t.kind==='calendar').slice(0,15);
    $('content').innerHTML=`<section class="hero"><h2>لوحة متابعة المدير العام</h2><p>الأعمال الأصلية من Word، مع روابط التقارير ووقت التسليم وملاحظات المدير وتذكيرات Google.</p></section>
      <div class="legend">بمجرد حفظ رابط التقرير داخل العمل يُسجّل وقت وصوله على الجهاز. لن يظهر لمستخدم آخر إلا بعد مزامنة البيانات باستخدام Google Sheets أو مشاركة النسخة الاحتياطية.</div>
      ${kpis(filteredByWorkspace(tasks))}<div class="button-row"><h2>الأقسام الإدارية</h2><button class="primary" data-action="newTask">+ عمل جديد</button></div>${cards()}
      <div class="section-title"><h2>عينة من الروزنامة المركزية</h2></div><div class="panel">${table(central)}</div>`;
  } else if(view==='organization') $('content').innerHTML=organization();
  else if(view==='building') $('content').innerHTML=buildingMap();
  else if(view==='departments') $('content').innerHTML=`<section class="hero"><h2>الأقسام الإدارية</h2><p>كل قسم بلونه، مع إمكانية إضافة أعمال وتوثيق تقارير.</p></section><div class="button-row"><h2>أقسام المؤسسة</h2><button class="primary" data-action="newTask">+ عمل جديد</button></div>${cards()}`;
  else if(view==='jobs') $('content').innerHTML=jobsPage();
  else if(view==='department') {const list=mainList();$('content').innerHTML=`<section class="hero"><h2>${e(department)}</h2><p>الاستحقاقات والتقارير وقرارات المدير ومواعيدها.</p></section>${kpis(list)}<div class="button-row"><h2>سجل أعمال القسم</h2><button class="primary" data-action="newTask">+ إضافة عمل لهذا القسم</button></div>${taskArea(list,true)}`;}
  else if(view==='calendar') {$('content').innerHTML=`<div class="panel"><div class="actions"><button class="outline" data-action="ics">تصدير المواعيد المؤرّخة ICS</button></div><div class="months"><button data-month="" class="${!month?'active':''}">كل الأشهر</button>${MONTHS.map(m=>`<button data-month="${e(m)}" class="${month===m?'active':''}">${e(m)}</button>`).join('')}</div></div>${taskArea(mainList())}`;}
  else if(view==='routine'||view==='government') $('content').innerHTML=`<section class="hero"><h2>${TITLES[view]}</h2></section>${kpis(mainList())}${taskArea(mainList())}`;
  else if(view==='settings')renderSettings();
}
function navigate(target){view=target;search='';tab='all';statusFilter='';render();}

function renderReportHistory(t){
  const items=[];
  if(t.proofUrl)items.push(`<div class="report-step"><strong>التقرير الأول</strong>${linkButton(t.proofUrl,'فتح التقرير')}<small> · ${e(stampView(t.receivedAt))}</small></div>`);
  if(t.decision)items.push(`<div class="report-step"><strong>تعليق المدير</strong><p>${e(t.decision)}</p><small>${e(stampView(t.decisionAt))}</small></div>`);
  if(t.revisedProofUrl)items.push(`<div class="report-step"><strong>النسخة المعدّلة</strong>${linkButton(t.revisedProofUrl,'فتح النسخة المعدّلة')}<small> · ${e(stampView(t.revisedAt))}</small></div>`);
  return items.length?`<div class="report-history">${items.join('')}</div>`:'<span class="muted">لم تصل تقارير بعد.</span>';
}
function openTask(id){
  selectedTask=byId(id);if(!selectedTask)return;
  const t=task(selectedTask),dlg=$('taskDialog');
  dlg.dataset.theme=deptTheme(t.department);
  dlg.dataset.mode=workspaceMode;
  $('taskTitle').textContent=t.title;
  const details=[['القسم',t.department],['نوع العمل',t.section||KIND_LABEL[t.kind]],['الموعد بحسب Word',t.deadlineText],['تفاصيل المطلوب',t.description],['مصادر المعلومات والأدلة',t.evidence],['المسؤولون وفق Word',t.sourceAssignee],['دور المدير',t.managerAction],['المتابعة اللاحقة',t.followUp],['المرجع',t.reference]];
  $('taskInfo').innerHTML=details.filter(([,v])=>v).map(([label,val])=>`<div><b>${e(label)}:</b> ${e(val)}</div>`).join('');
  [['taskStatus','status'],['taskAssignee','assignee'],['taskProof','proofUrl'],['submittedBy','submittedBy'],['taskDecision','decision'],['taskNotes','notes'],['recipientEmail','recipientEmail'],['taskRevisedProof','revisedProofUrl']].forEach(([element,key])=>$(element).value=t[key]||'');
  $('taskDate').value=t.date||dateOfTask(t)||'';
  $('taskReceivedAt').value=t.receivedAt?stampView(t.receivedAt):'';
  $('taskDecisionAt').value=t.decisionAt?stampView(t.decisionAt):'';
  $('taskRevisedAt').value=t.revisedAt?stampView(t.revisedAt):'';
  $('reportLinkBox').innerHTML=t.proofUrl?`<div class="report-box">${linkButton(t.proofUrl,'فتح التقرير الأصلي')}</div>`:'';
  $('reportHistory').innerHTML=renderReportHistory(t);
  $('emailAudit').textContent=t.commentSentAt?`آخر تعليق أرسل إلى ${t.commentSentTo||'—'} بتاريخ ${stampView(t.commentSentAt)}`:'لم يُرسل تعليق بالبريد بعد.';
  $('editCustomTask').hidden=!customTasks[t.id]||customTasks[t.id].deleted||workspaceMode!=='manager';
  const isManager=workspaceMode==='manager';
  ['taskStatus','taskAssignee','taskDate','taskDecision','taskNotes','recipientEmail'].forEach(id=>$(id).disabled=!isManager);
  ['taskProof','submittedBy','taskRevisedProof'].forEach(id=>$(id).disabled=isManager);
  $('sendComment').hidden=!isManager;
  $('saveWorkflow').textContent=isManager?'حفظ تعليق المدير وقراره':'إرسال رابط التقرير / النسخة المعدّلة';
  dlg.showModal();
}
function persist(type,id,value){
  const key=type==='task_change'?STORE:type==='custom_task'?EXTRA_STORE:JOB_STORE;
  const target=type==='task_change'?changes:type==='custom_task'?customTasks:jobs;
  const next={...target,[id]:value};
  if(!saveLocal(key,next)){notify('تعذّر الحفظ في المتصفح',true);return false;}
  if(type==='task_change')changes=next;
  else if(type==='custom_task'){customTasks=next;buildTasks();}
  else jobs=next;
  queueCloudEvent({timestamp:stamp(),type,id,payload:value,editor:config.editorName||''});
  return true;
}
function queueCloudEvent(event){
  pending=[...pending,event];
  saveLocal(PENDING_STORE,pending);
  indicator(`بانتظار مزامنة (${pending.length})`,'warn');
  if(canUseScope(SCOPE.sheets)&&config.sheetId)void syncCloud({silent:true,interactive:false});
}
// نحفظ الحقول المتغيّرة فقط، لئلا يطغى تعليق المدير على رابط الموظف عند المزامنة.
function persistTaskPatch(id,patch,kind='task_patch'){
  const changedAt=stamp();
  const old=changes[id]||{};
  const next={...old,...patch,updatedAt:changedAt};
  if(!saveLocal(STORE,{...changes,[id]:next})){notify('تعذّر الحفظ المحلي',true);return false;}
  changes={...changes,[id]:next};
  queueCloudEvent({timestamp:changedAt,type:kind,id,payload:{...patch,updatedAt:changedAt},editor:config.editorName||''});
  return true;
}
function saveTask(){
  if(!selectedTask)return false;
  const t=task(selectedTask),id=t.id;
  let patch={};
  if(workspaceMode==='manager'){
    const comment=$('taskDecision').value.trim();
    const recipient=$('recipientEmail').value.trim();
    if(recipient&&!validEmail(recipient)){notify('البريد غير صالح',true);return false;}
    patch={status:$('taskStatus').value,assignee:$('taskAssignee').value.trim(),date:$('taskDate').value,decision:comment,notes:$('taskNotes').value.trim(),recipientEmail:recipient};
    if(comment!==String(t.decision||''))patch.decisionAt=comment?stamp():'';
  }else{
    const link=$('taskProof').value.trim();
    const revised=$('taskRevisedProof').value.trim();
    if((link&&!safeUrl(link))||(revised&&!safeUrl(revised))){notify('رابط التقرير يجب أن يبدأ بـ https://',true);return false;}
    if(revised&&!link){notify('أرسل التقرير الأصلي أولاً',true);return false;}
    patch={proofUrl:link,submittedBy:$('submittedBy').value.trim(),revisedProofUrl:revised};
    if(link!==String(t.proofUrl||''))patch.receivedAt=link?stamp():'';
    if(revised!==String(t.revisedProofUrl||''))patch.revisedAt=revised?stamp():'';
    if((link&&!t.proofUrl)||(revised&&revised!==t.revisedProofUrl))patch.status='تم الاستلام';
  }
  if(!persistTaskPatch(id,patch,workspaceMode==='manager'?'manager_note':'submission'))return false;
  render();
  return true;
}

function openNewTask(id=null,preset=''){
  editingTaskId=id;
  const old=id?customTasks[id]:null;
  $('customTaskHeading').textContent=old?'تعديل العمل الإضافي':'إضافة عمل جديد';
  $('newDepartment').innerHTML=departments.map(d=>`<option value="${e(d)}">${e(d)}</option>`).join('');
  $('newTitle').value=old?.title||'';
  $('newDepartment').value=old?.department||preset||department||departments[0];
  $('newKind').value=old?.kind||'plan';
  $('newDeadline').value=old?.deadlineText||'';
  $('newDate').value=old?.date||'';
  $('newEvidence').value=old?.evidence||'';
  $('newAssignee').value=old?.sourceAssignee||'';
  $('newManagerAction').value=old?.managerAction||'';
  $('newFollowUp').value=old?.followUp||'';
  $('deleteCustomTask').hidden=!old;
  $('customTaskDialog').showModal();
}
function saveCustomTask(){
  const title=$('newTitle').value.trim();if(!title){notify('عنوان العمل مطلوب',true);return false;}
  const previous=editingTaskId?customTasks[editingTaskId]:null;
  const id=editingTaskId||`CUSTOM-${Date.now()}-${Math.random().toString(36).slice(2,8)}`;
  const kind=$('newKind').value;
  const record={...previous,id,kind,department:$('newDepartment').value,section:'أعمال أضيفت إلى المنظومة',title,frequency:KIND_LABEL[kind],deadlineText:$('newDeadline').value.trim(),date:$('newDate').value,evidence:$('newEvidence').value.trim(),sourceAssignee:$('newAssignee').value.trim(),managerAction:$('newManagerAction').value.trim(),followUp:$('newFollowUp').value.trim(),status:previous?.status||'لم يُستلم',reference:'إضافة إلكترونية — خارج ملف Word الأصلي',createdAt:previous?.createdAt||stamp(),updatedAt:stamp(),deleted:false};
  return persist('custom_task',id,record);
}
function deleteCustomTask(){
  if(!editingTaskId||!customTasks[editingTaskId])return;
  if(!confirm('حذف هذا العمل المضاف؟ الأعمال الأصلية من Word لا يمكن حذفها هنا.'))return;
  const id=editingTaskId,old=customTasks[id];
  if(persist('custom_task',id,{...old,deleted:true,updatedAt:stamp()})){$('customTaskDialog').close();render();notify('تم إلغاء العمل الإضافي');}
}

function openJob(id=null,preset='',proposedTitle=''){
  editingJobId=id;
  const old=id?jobs[id]:null;
  $('jobHeading').textContent=old?'تعديل الوصف الوظيفي':'إضافة وصف وظيفي';
  $('jobDepartment').innerHTML=departments.map(d=>`<option value="${e(d)}">${e(d)}</option>`).join('');
  const fields={jobTitle:'title',jobReportsTo:'reportsTo',jobPurpose:'purpose',jobResponsibilities:'responsibilities',jobAuthority:'authority',jobQualifications:'qualifications',jobKpis:'kpis',jobDocUrl:'docUrl'};
  Object.entries(fields).forEach(([el,key])=>$(el).value=old?.[key]||'');
  $('jobDepartment').value=old?.department||preset||departments[0];
  if(!old&&proposedTitle)$('jobTitle').value=proposedTitle;
  $('deleteJob').hidden=!old;
  $('jobDialog').showModal();
}
function saveJob(){
  const title=$('jobTitle').value.trim(),docUrl=$('jobDocUrl').value.trim();
  if(!title){notify('المسمى الوظيفي مطلوب',true);return false;}
  if(docUrl&&!safeUrl(docUrl)){notify('رابط الوصف الوظيفي يجب أن يبدأ بـ https://',true);return false;}
  const old=editingJobId?jobs[editingJobId]:null;
  const id=editingJobId||`ROLE-${Date.now()}-${Math.random().toString(36).slice(2,8)}`;
  const record={...old,id,title,department:$('jobDepartment').value,reportsTo:$('jobReportsTo').value.trim(),purpose:$('jobPurpose').value.trim(),responsibilities:$('jobResponsibilities').value.trim(),authority:$('jobAuthority').value.trim(),qualifications:$('jobQualifications').value.trim(),kpis:$('jobKpis').value.trim(),docUrl,updatedAt:stamp(),createdAt:old?.createdAt||stamp(),deleted:false};
  return persist('job',id,record);
}
function deleteJob(){
  if(!editingJobId||!jobs[editingJobId])return;
  if(!confirm('هل تريد حذف بطاقة الوصف الوظيفي؟'))return;
  const id=editingJobId;
  if(persist('job',id,{...jobs[id],deleted:true,updatedAt:stamp()})){$('jobDialog').close();render();notify('حُذفت البطاقة');}
}

function download(name,text,mime){const a=document.createElement('a'),url=URL.createObjectURL(new Blob([text],{type:mime}));a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1500);}
function backup(){download('jinan-director-backup-v3.json',JSON.stringify({changes,customTasks,jobs,pending,config,exportedAt:stamp()},null,2),'application/json');}
async function restore(file){
  if(!file)return;
  try{
    const data=JSON.parse(await file.text());
    if(!data.changes||typeof data.changes!=='object')throw Error('نسخة احتياطية غير صحيحة');
    if(!confirm('استيراد النسخة الاحتياطية سيستبدل المعلومات المحلية الحالية. متابعة؟'))return;
    changes=data.changes;customTasks=data.customTasks||{};jobs=data.jobs||{};pending=data.pending||[];
    if(![saveLocal(STORE,changes),saveLocal(EXTRA_STORE,customTasks),saveLocal(JOB_STORE,jobs),saveLocal(PENDING_STORE,pending)].every(Boolean))throw Error('تعذّر حفظ جزء من النسخة المستوردة');
    buildTasks();render();notify('استُعيدت النسخة الاحتياطية محلياً');
  }catch(err){notify(err.message,true);}
}
function exportIcs(){
  const esc=s=>String(s||'').replaceAll('\\','\\\\').replaceAll(',','\\,').replaceAll(';','\\;').replaceAll('\n','\\n');
  const out=['BEGIN:VCALENDAR','VERSION:2.0','PRODID:-//Al Jinan High School//Management//AR','CALSCALE:GREGORIAN'];
  for(const original of tasks){const t=task(original);if(!t.date)continue;const a=t.date.replaceAll('-',''),b=new Date(Date.parse(t.date+'T00:00:00Z')+86400000).toISOString().slice(0,10).replaceAll('-','');out.push('BEGIN:VEVENT',`UID:${t.id}@jinan-management`,`DTSTAMP:${stamp().replace(/[-:]/g,'').replace(/\.\d{3}Z$/,'Z')}`,`DTSTART;VALUE=DATE:${a}`,`DTEND;VALUE=DATE:${b}`,`SUMMARY:${esc(t.title)}`,`DESCRIPTION:${esc(t.department)}`,'END:VEVENT');}
  out.push('END:VCALENDAR');download('jinan-calendar.ics',out.join('\r\n'),'text/calendar;charset=utf-8');
}

/* ============= Google OAuth: الصلاحيات عند الحاجة فقط ============= */
function canUseScope(scope) {return Boolean(oauthToken && Date.now()<tokenExpires-45000 && oauthGranted.has(scope));}
async function tokenFor(scope, interactive=true) {
  if(canUseScope(scope))return oauthToken;
  if(!interactive)throw Error('اتصال Google يحتاج إعادة تفويض');
  if(!config.clientId)throw Error('أضف OAuth Client ID في الإعدادات أولاً');
  if(!window.google?.accounts?.oauth2)throw Error('مكتبة تسجيل الدخول إلى Google غير متاحة');
  return new Promise((resolve,reject)=>{
    const client=google.accounts.oauth2.initTokenClient({
      client_id:config.clientId,
      scope,
      include_granted_scopes:true,
      prompt:'select_account',
      callback:result=>{
        if(result.error){reject(Error(result.error));return;}
        oauthToken=result.access_token;
        tokenExpires=Date.now()+Number(result.expires_in||3600)*1000;
        oauthGranted=new Set(String(result.scope||scope).split(/\s+/).filter(Boolean));
        resolve(oauthToken);
      },
      error_callback:result=>reject(Error(result.message||'لم تكتمل موافقة Google'))
    });
    client.requestAccessToken();
  });
}
async function googleApi(url,method='GET',body,scope=SCOPE.calendar,interactive=true) {
  const accessToken=await tokenFor(scope,interactive);
  const response=await fetch(url,{method,headers:{Authorization:`Bearer ${accessToken}`,...(body!==undefined?{'Content-Type':'application/json'}:{})},...(body!==undefined?{body:JSON.stringify(body)}:{})});
  const data=await response.json().catch(()=>({}));
  if(!response.ok){const err=Error(data.error?.message||`فشل Google API: HTTP ${response.status}`);err.status=response.status;throw err;}
  return data;
}
function saveConfig(){
  config={clientId:$('clientId').value.trim(),calendarId:$('calendarId').value.trim()||'primary',sheetId:$('sheetId').value.trim(),editorName:$('editorName').value.trim()};
  if(!saveLocal(CONFIG_STORE,config)){notify('تعذّر حفظ الإعدادات',true);return;}
  oauthToken='';oauthGranted=new Set();
  indicator('إعدادات محفوظة محلياً');notify('تم حفظ الإعدادات. اربط خدمات Google المطلوبة.');
}
async function connectGoogle(scope,label){try{await tokenFor(scope);notify(`تم السماح باستخدام ${label}`);if(scope===SCOPE.sheets)await syncCloud({silent:false,interactive:false});}catch(err){notify(err.message,true);}}

/* ============= إضافة العمل نفسه إلى تقويم الحساب المصرّح ============= */
function dateOfTask(t){
  if(/^\d{4}-\d{2}-\d{2}$/.test(t.date||''))return t.date;
  const raw=String(t.deadlineText||'');
  const m=raw.match(/(?:^|[^\d])(\d{1,2})\s*[\/\-.]\s*(\d{1,2})(?!\d)/);
  if(!m)return null;
  const day=Number(m[1]),mon=Number(m[2]);
  if(mon<1||mon>12||day<1||day>31)return null;
  // تاريخ قادم لأن Word يحتوي على اليوم والشهر فقط دون تحديد السنة.
  const currentYear=Number(today().slice(0,4));
  const format=year=>`${year}-${String(mon).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
  for(const year of [currentYear,currentYear+1]){
    const iso=format(year);
    const valid=new Date(iso+'T12:00:00Z').toISOString().slice(0,10)===iso;
    if(valid&&iso>=today())return iso;
  }
  return null;
}
async function remindTask(id){
  const source=byId(id);if(!source)return;
  const t=task(source);
  const due=dateOfTask(t);
  if(!due){
    openTask(id);
    $('taskDate').focus();
    notify('موعد هذا العمل غير محدد بيوم وشهر واضحين. أدخل التاريخ أولاً.',true);
    return;
  }
  try{
    // تفويض Google يتم فقط عند الضغط، عبر الحساب الذي يختاره المستخدم.
    await tokenFor(SCOPE.calendar);
    const base='https://www.googleapis.com/calendar/v3/calendars/primary/events';
    const lookup=await googleApi(base+'?'+new URLSearchParams({privateExtendedProperty:'jinanTaskId='+t.id,maxResults:'3'}),'GET',undefined,SCOPE.calendar,false);
    const existing=(lookup.items||[]).find(ev=>ev.status!=='cancelled');
    const end=new Date(Date.parse(due+'T00:00:00Z')+86400000).toISOString().slice(0,10);
    const body={
      summary:t.title,
      description:`القسم: ${t.department}\nالمسؤول: ${t.assignee||t.sourceAssignee||''}\nمرجع العمل: ${t.id}`,
      start:{date:due},end:{date:end},
      reminders:{useDefault:true},
      extendedProperties:{private:{jinanTaskId:t.id}}
    };
    const event=existing
      ? await googleApi(`${base}/${encodeURIComponent(existing.id)}`,'PUT',body,SCOPE.calendar,false)
      : await googleApi(base,'POST',body,SCOPE.calendar,false);
    notify(`تم ${existing?'تحديث':'إضافة'} العمل في تقويم الحساب الذي اخترته: ${dateView(due)}`);
    // حدث تقويم كل مستخدم مستقل عن الموظفين الآخرين، فلا ننشر معرّفه في سجل المتابعات المشترك.
    // إتمام الإضافة مباشرة من دون فتح نافذة تحرير الحدث أو طلب خطوة إضافية.
  }catch(error){notify('لم تتم إضافة الموعد: '+error.message,true);}
}
async function remindFromDialog(){
  if(!selectedTask)return;
  // لا نحتاج إلى حفظ/تعديل التقرير لإنشاء موعد في التقويم.
  const chosen=$('taskDate').value;
  if(chosen && chosen!==task(selectedTask).date){
    if(!persistTaskPatch(selectedTask.id,{date:chosen},'task_patch'))return;
  }
  await remindTask(selectedTask.id);
}
/* ============= Gmail API: تعليق المدير عبر البريد من الحساب المتصل ============= */
function base64UrlUtf8(text){
  const bytes=new TextEncoder().encode(String(text));
  let result='';
  for(let i=0;i<bytes.length;i+=8192){result+=String.fromCharCode(...bytes.subarray(i,i+8192));}
  return btoa(result).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
}
async function sendCommentEmail(){
  if(!selectedTask)return;
  const recipient=$('recipientEmail').value.trim(),decision=$('taskDecision').value.trim();
  if(!validEmail(recipient)){notify('أدخل بريد الجهة المسؤولة بطريقة صحيحة',true);return;}
  if(!decision){notify('اكتب تعليق المدير قبل الإرسال',true);return;}
  if(!confirm(`إرسال تعليق المدير إلى ${recipient} باستخدام حساب Gmail المتصل؟`))return;
  if(!saveTask())return;
  const t=task(selectedTask);
  const subject=`تعليق المدير على تقرير: ${t.title}`;
  const body=[`ثانوية الجنان — متابعة أعمال الإدارة`,``, `العمل: ${t.title}`,`القسم: ${t.department}`,`رقم المتابعة: ${t.id}`,``, `تعليق / قرار المدير:`,decision,'',t.proofUrl?`رابط التقرير: ${t.proofUrl}`:'',``, `هذه الرسالة صادرة من منظومة متابعة المدير العام.`].filter(x=>x!==null&&x!==undefined).join('\r\n');
  const toMimeBase64=text=>{const base=base64UrlUtf8(text).replace(/-/g,'+').replace(/_/g,'/');return base.padEnd(Math.ceil(base.length/4)*4,'=');};
  const encodedSubject=toMimeBase64(subject);
  const encodedBody=toMimeBase64(body).match(/.{1,76}/g).join('\r\n');
  const mime=[`To: ${recipient}`,`Subject: =?UTF-8?B?${encodedSubject}?=`,`MIME-Version: 1.0`,`Content-Type: text/plain; charset=UTF-8`,`Content-Transfer-Encoding: base64`,'',encodedBody].join('\r\n');
  try{
    const result=await googleApi('https://gmail.googleapis.com/gmail/v1/users/me/messages/send','POST',{raw:base64UrlUtf8(mime)},SCOPE.gmail);
    const next={commentSentAt:stamp(),commentSentTo:recipient,gmailMessageId:result.id||''};
    persistTaskPatch(t.id,next,'manager_note');
    notify(`تم إرسال التعليق عبر Gmail إلى ${recipient}`);
    if($('taskDialog').open){$('emailAudit').textContent=`تم إرسال التعليق إلى ${recipient} في ${stampView(next.commentSentAt)}`;}
  }catch(err){notify('لم يُرسل البريد: '+err.message,true);}
}

/* ============= Google Sheets: سجل إضافات فقط لتقليل تضارب التحديثات ============= */
function sheetBase(){if(!config.sheetId)throw Error('أدخل Spreadsheet ID في الإعدادات');return 'https://sheets.googleapis.com/v4/spreadsheets/'+encodeURIComponent(config.sheetId);}
async function ensureLogSheet(interactive=true){
  const base=sheetBase(),meta=await googleApi(base+'?fields=sheets.properties.title','GET',undefined,SCOPE.sheets,interactive);
  if(!meta.sheets?.some(s=>s.properties?.title===SHEET_NAME)){
    await googleApi(base+':batchUpdate','POST',{requests:[{addSheet:{properties:{title:SHEET_NAME}}}]},SCOPE.sheets,interactive);
  }
}
async function appendEvent(event,interactive=true){
  const base=sheetBase(),range=encodeURIComponent(`'${SHEET_NAME}'!A:E`);
  await googleApi(`${base}/values/${range}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,'POST',{values:[[event.timestamp,event.type,event.id,JSON.stringify(event.payload),event.editor||'']]},SCOPE.sheets,interactive);
}
function applyRemoteEvent(record){
  const {type,id,payload}=record;
  if(!id||!payload||typeof payload!=='object')return;
  if(['task_patch','submission','manager_note'].includes(type)){
    if(!tasks.some(t=>t.id===id))return;
    const old=changes[id]||{};
    const versions={...(old._fieldVersions||{})};
    const next={...old};
    const timestamp=String(record.timestamp||payload.updatedAt||'');
    for(const [key,value] of Object.entries(payload)){
      if(key==='updatedAt'||key==='_fieldVersions')continue;
      const localTime=String(versions[key]||(old._fieldVersions ? '' : old.updatedAt||''));
      if(timestamp>=localTime){next[key]=value;versions[key]=timestamp;}
    }
    next._fieldVersions=versions;
    next.updatedAt=String(old.updatedAt||'')>timestamp?old.updatedAt:timestamp;
    changes[id]=next;
    return;
  }
  const old=type==='task_change'?changes[id]:type==='custom_task'?customTasks[id]:type==='job'?jobs[id]:null;
  const latest=String(payload.updatedAt||record.timestamp||'');
  if(old&&String(old.updatedAt||'')>=latest)return;
  if(type==='task_change'&&tasks.some(t=>t.id===id))changes[id]={...(old||{}),...payload};
  else if(type==='custom_task')customTasks[id]=payload;
  else if(type==='job')jobs[id]=payload;
}
async function readCloud(interactive=true){
  const base=sheetBase(),range=encodeURIComponent(`'${SHEET_NAME}'!A2:E10000`);
  const data=await googleApi(`${base}/values/${range}`,'GET',undefined,SCOPE.sheets,interactive);
  let count=0;
  for(const row of data.values||[]){
    const [timestamp,type,id,payloadJSON,editor]=row;
    try{applyRemoteEvent({timestamp,type,id,payload:JSON.parse(payloadJSON),editor});count++;}
    catch(err){console.warn('سجل سحابي غير صالح',err);}
  }
  buildTasks();
  saveLocal(STORE,changes);saveLocal(EXTRA_STORE,customTasks);saveLocal(JOB_STORE,jobs);
  return count;
}
async function syncCloud({silent=false,interactive=true}={}){
  if(syncBusy)return;
  if(!config.sheetId){if(!silent)notify('أدخل Spreadsheet ID في الإعدادات',true);return;}
  syncBusy=true;
  try{
    await ensureLogSheet(interactive);
    let n=0;
    // سجل العمليات غير المرفوعة يُكتب أولاً؛ لا يُحذف إلا بعد نجاح كل طلب.
    while(pending.length){
      const entry=pending[0];
      await appendEvent(entry,interactive);
      pending=pending.slice(1);
      if(!saveLocal(PENDING_STORE,pending))notify('فشل حفظ قائمة المزامنة محلياً',true);
      n++;
    }
    const count=await readCloud(interactive);
    indicator('متزامن مع Google Sheets','ok');
    if(!silent)notify(`تم رفع ${n} تحديثات واستقبال ${count} سجلاً من Google Sheets`);
    if(!$('taskDialog').open&&!$('customTaskDialog').open&&!$('jobDialog').open)render();
  }catch(err){
    indicator(pending.length?`غير متزامن (${pending.length})`:'تعذّر التحديث','warn');
    if(!silent)notify('تعذّرت المزامنة: '+err.message,true);
    else console.warn('مزامنة Google',err);
  }finally{syncBusy=false;}
}

/* ============= جميع أحداث الواجهة ============= */
document.addEventListener('click',async ev=>{
  const btn=ev.target.closest('button');if(!btn)return;
  if(btn.dataset.visualZoom) {
    const img=$('visualDocumentImage');
    if(!img)return;
    const action=btn.dataset.visualZoom;
    const current=Number(img.dataset.zoom||100);
    const next=action==='reset'?100:Math.max(75,Math.min(300,current+Number(action)));
    img.dataset.zoom=String(next);
    img.style.width=next+'%';
    $('visualZoomLabel').textContent=next+'%';
    return;
  }
  if(btn.dataset.view){navigate(btn.dataset.view);return;}
  if(btn.dataset.dept){department=btn.dataset.dept;navigate('department');return;}
  if(btn.dataset.task){openTask(btn.dataset.task);return;}
  if(btn.dataset.remind){await remindTask(btn.dataset.remind);return;}
  if(btn.dataset.tab){tab=btn.dataset.tab;render();return;}
  if(btn.dataset.month!==undefined){month=btn.dataset.month;render();return;}
  if(btn.dataset.job){openJob(btn.dataset.job);return;}
  if(btn.dataset.addOrgRole){openJob(null,btn.dataset.addOrgDept||'',btn.dataset.addOrgRole);return;}
  if(btn.dataset.newJobDept){openJob(null,btn.dataset.newJobDept);return;}
  if(btn.dataset.close){$(btn.dataset.close)?.close();return;}
  switch(btn.dataset.action){
    case 'newTask':openNewTask(null,view==='department'?department:'');break;
    case 'addJob':openJob();break;
    case 'backup':backup();break;
    case 'restore':$('backupFile')?.click();break;
    case 'ics':exportIcs();break;
    case 'saveConfig':saveConfig();break;
    case 'connectCalendar':await connectGoogle(SCOPE.calendar,'Google Calendar');break;
    case 'connectMail':await connectGoogle(SCOPE.gmail,'Gmail');break;
    case 'connectSheets':await connectGoogle(SCOPE.sheets,'Google Sheets');break;
    case 'syncCloud':await syncCloud();break;
  }
});
document.addEventListener('input',ev=>{
  if(ev.target.id==='search'){search=ev.target.value;updateRows();}
  if(ev.target.id==='jobSearch'){const cursor=ev.target.selectionStart;jobSearch=ev.target.value;render();const input=$('jobSearch');if(input){input.focus();input.setSelectionRange(cursor,cursor);}}
});
document.addEventListener('change',ev=>{
  if(ev.target.id==='workspaceMode'){
    workspaceMode=ev.target.value==='employee'?'employee':'manager';
    saveLocal(MODE_STORE,workspaceMode);
    render();
  }
  if(ev.target.id==='workspaceDepartment'){
    workspaceDepartment=ev.target.value;
    saveLocal(MODE_DEPT_STORE,workspaceDepartment);
    render();
  }

  if(ev.target.id==='statusFilter'){statusFilter=ev.target.value;updateRows();}
  if(ev.target.id==='jobDepartmentFilter'){jobDepartment=ev.target.value;render();}
  if(ev.target.id==='backupFile'){void restore(ev.target.files?.[0]);}
});
$('exportBtn').addEventListener('click',backup);
$('refreshCloud').addEventListener('click',()=>syncCloud());
$('closeDialog').addEventListener('click',()=>$('taskDialog').close());
$('taskForm').addEventListener('submit',ev=>{ev.preventDefault();if(saveTask()){$('taskDialog').close();notify('تم حفظ التغييرات محلياً'+(config.sheetId?' وسيُحاول إرسالها إلى Google Sheets':'؛ لعرضها للمدير من جهاز آخر، اربط Google Sheets'));}});
$('remindDirect').addEventListener('click',remindFromDialog);
$('editCustomTask').addEventListener('click',()=>{if(!selectedTask||!customTasks[selectedTask.id])return;$('taskDialog').close();openNewTask(selectedTask.id);});
$('sendComment').addEventListener('click',sendCommentEmail);

$('customTaskForm').addEventListener('submit',ev=>{ev.preventDefault();if(saveCustomTask()){$('customTaskDialog').close();render();notify('أُضيف العمل بنجاح');}});
$('deleteCustomTask').addEventListener('click',deleteCustomTask);
$('jobForm').addEventListener('submit',ev=>{ev.preventDefault();if(saveJob()){$('jobDialog').close();render();notify('تم حفظ بطاقة الوصف الوظيفي');}});
$('deleteJob').addEventListener('click',deleteJob);

(async function initialize(){
  try{
    const embedded=$('embeddedData');
    let data;
    if(embedded?.textContent.trim())data=JSON.parse(embedded.textContent);
    else {const response=await fetch('./data/tasks.json',{cache:'no-store'});if(!response.ok)throw Error(`لم يُعثر على data/tasks.json (${response.status})`);data=await response.json();}
    if(!Array.isArray(data.tasks)||!Array.isArray(data.departments)||data.tasks.some(t=>!t.id||!t.title||!t.department))throw Error('صيغة ملف الأعمال غير صحيحة');
    baseTasks=data.tasks;departments=data.departments;
    buildTasks();
    indicator(pending.length?`غير متزامن (${pending.length})`:'حفظ محلي',pending.length?'warn':'');
    render();
  }catch(error){console.error(error);$('content').innerHTML=`<div class="note">${e(error.message)} — تأكد من رفع data/tasks.json مع الملفات الأخرى.</div>`;}
})();

// تحديث دوري صامت للبيانات المشتركة؛ لا يفتح نافذة OAuth من تلقاء نفسه.
setInterval(()=>{
  if(document.visibilityState==='visible'&&config.sheetId&&canUseScope(SCOPE.sheets))void syncCloud({silent:true,interactive:false});
},90000);
