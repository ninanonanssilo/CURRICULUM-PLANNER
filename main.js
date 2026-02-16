const $ = (id) => document.getElementById(id);

// Sidebar nav (category)
(function initNav(){
  const sidebar = $("sidebar");
  const scrim = $("scrim");
  const btnOpen = $("navToggle");
  const btnClose = $("navClose");
  const open = ()=>{
    sidebar?.classList.add('isOpen');
    scrim?.classList.add('isOn');
    btnOpen?.setAttribute('aria-expanded','true');
    document.body.style.overflow = 'hidden';
  };
  const close = ()=>{
    sidebar?.classList.remove('isOpen');
    scrim?.classList.remove('isOn');
    btnOpen?.setAttribute('aria-expanded','false');
    document.body.style.overflow = '';
  };
  btnOpen?.addEventListener('click', open);
  btnClose?.addEventListener('click', close);
  scrim?.addEventListener('click', close);
  document.addEventListener('keydown', (e)=>{ if(e.key==='Escape') close(); });
})();

const els = {
  file: $("file"),
  sheet: $("sheet"),
  rangePolicy: $("rangePolicy"),
  gradeBand: $("gradeBand"),
  gradeFilter: $("gradeFilter"),
  subjectPlan: $("subjectPlan"),
  subjectCurriculum: $("subjectCurriculum"),
  termStartDate: $("termStartDate"),
  scheduleDate: $("scheduleDate"),
  scheduleType: $("scheduleType"),
  scheduleMemo: $("scheduleMemo"),
  btnAddSchedule: $("btnAddSchedule"),
  btnLoadNational: $("btnLoadNational"),
  btnClearSchedule: $("btnClearSchedule"),
  btnTotals: $("btnTotals"),
  btnCSV: $("btnCSV"),
  btnDocx: $("btnDocx"),
  schoolYear: $("schoolYear"),
  status: $("status"),
  totalsBody: $("totalsBody"),
  curriculumBody: $("curriculumBody"),
  timetableWrap: $("timetableWrap"),
  scheduleBody: $("scheduleBody"),
  previewWrap: $("previewWrap"),
  toast: $("toast"),
  btnGuide: $("btnGuide"),
  guideOverlay: $("guideOverlay"),
  guideStep: $("guideStep"),
  guideTitle: $("guideTitle"),
  guideBody: $("guideBody"),
  btnGuideNext: $("btnGuideNext"),
  btnGuideSkip: $("btnGuideSkip"),
};

const state = {
  wb: null,
  rows: null,
  headerMap: null,
  activeSheet: null,
  lastTotals: null,
  scheduleRows: [],
};

function toast(msg, ms = 1800) {
  const el = els.toast;
  if (!el) return;
  const t = String(msg || "").trim();
  if (!t) return;
  el.textContent = t;
  el.classList.add("show");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.classList.remove("show"), Math.max(700, ms | 0));
}

function setStatus(msg, { error = false } = {}) {
  const el = els.status;
  if (!el) return;
  const t = String(msg || "").trim();
  if (!t) {
    el.classList.remove("show", "is-error");
    el.textContent = "";
    return;
  }
  el.textContent = t;
  el.classList.add("show");
  el.classList.toggle("is-error", !!error);
}

const guideSteps = [
  { title: "1단계 · 기본 설정", body: "학년도, 학년군, 학년을 먼저 고르면 이후 계산 결과가 정확해집니다.", target: "schoolYear" },
  { title: "2단계 · 학사 일정", body: "학기 시작일과 공휴일/대체공휴일/휴업일을 선택형으로 추가하면 주차별 수업가능일이 자동 반영됩니다.", target: "scheduleDate" },
  { title: "3단계 · 편제 시수", body: "총론 기준 과목별 편제 시수를 입력하면 합계표에서 편제 대비 차이를 바로 볼 수 있습니다.", target: "subjectPlan" },
  { title: "4단계 · 과목별 교육과정", body: "과목 | 성취기준 | 성취수준 형식으로 넣으면 요약표가 자동 생성됩니다.", target: "subjectCurriculum" },
  { title: "5단계 · 파일 업로드", body: "xlsx를 올리고 시트를 선택하면 원본 데이터를 자동 인식합니다.", target: "file" },
  { title: "6단계 · 결과 확인", body: "교과별 시수 합계와 연간 시간표(수업가능일 포함)를 확인하세요.", target: "timetableWrap" },
  { title: "7단계 · 다운로드", body: "확인이 끝나면 CSV/DOCX로 내려받아 학년별 업무자료로 활용하면 됩니다.", target: "btnCSV" },
];
let guideIndex = 0;
let guideFocusedEl = null;

function clearGuideFocus(){
  if (guideFocusedEl) guideFocusedEl.classList.remove('guideFocus');
  guideFocusedEl = null;
}
function showGuideStep(){
  if (!els.guideOverlay) return;
  const step = guideSteps[guideIndex];
  if (!step) return;

  els.guideStep.textContent = `${guideIndex + 1} / ${guideSteps.length}`;
  els.guideTitle.textContent = step.title;
  els.guideBody.textContent = step.body;

  clearGuideFocus();
  const target = document.getElementById(step.target);
  if (target){
    guideFocusedEl = target;
    target.classList.add('guideFocus');
    target.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  els.btnGuideNext.textContent = (guideIndex === guideSteps.length - 1) ? '완료' : 'OK';
}
function closeGuide(){
  els.guideOverlay?.classList.remove('show');
  els.guideOverlay?.setAttribute('aria-hidden', 'true');
  clearGuideFocus();
}
function openGuide(){
  guideIndex = 0;
  els.guideOverlay?.classList.add('show');
  els.guideOverlay?.setAttribute('aria-hidden', 'false');
  showGuideStep();
}

function normalizeKey(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");
}

function parseHourCell(v, policy = "avg") {
  // Accept: number, "3", "7-9", "7~9", "7–9", "7~ 9"
  if (v === null || v === undefined) return null;
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const s0 = String(v).trim();
  if (!s0) return null;

  const s = s0
    .replace(/[–—〜~]/g, "-")
    .replace(/[^0-9\-.]/g, "")
    .trim();

  if (!s) return null;
  if (/^-?\d+(?:\.\d+)?$/.test(s)) {
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  }

  const m = s.match(/^(-?\d+(?:\.\d+)?)\s*-\s*(-?\d+(?:\.\d+)?)$/);
  if (m) {
    const a = Number(m[1]);
    const b = Number(m[2]);
    if (!Number.isFinite(a) || !Number.isFinite(b)) return null;
    if (policy === "min") return Math.min(a, b);
    if (policy === "max") return Math.max(a, b);
    return (a + b) / 2;
  }

  return null;
}

function detectHeaders(rows) {
  // rows: array of arrays (first row is header candidate)
  if (!rows || !rows.length) return null;

  // Heuristic: choose first row with >=4 non-empty cells
  let headerRow = null;
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    const r = rows[i] || [];
    const nonEmpty = r.filter((x) => String(x || "").trim()).length;
    if (nonEmpty >= 4) {
      headerRow = r;
      break;
    }
  }
  if (!headerRow) headerRow = rows[0];

  const map = {};
  headerRow.forEach((h, idx) => {
    const k = normalizeKey(h);
    if (!k) return;

    // Korean header synonyms
    if (["주차", "주", "week"].includes(k)) map.week = idx;
    if (["학년", "grade"].includes(k)) map.grade = idx;
    if (["학기", "semester"].includes(k)) map.semester = idx;
    if (["교과", "과목", "subject"].includes(k)) map.subject = idx;
    if (["단원", "대단원", "unit"].includes(k)) map.unit = idx;
    if (["학습주제", "학습내용", "주제", "내용", "topic", "lesson"].includes(k)) map.topic = idx;
    if (["시수", "차시", "시간", "hours"].includes(k)) map.hours = idx;
  });

  // Minimum requirements
  if (map.subject == null || map.hours == null) return map; // allow partial but warn later
  return map;
}

function rowsToObjects(rows, headerMap) {
  // Determine actual data start row (after header row we picked in detectHeaders)
  // For simplicity: assume header is first row.
  const header = rows[0] || [];
  const data = rows.slice(1);

  const get = (arr, idx) => (idx == null ? "" : arr[idx]);

  return data
    .map((r) => {
      if (!r) return null;
      // Skip empty rows
      const nonEmpty = r.filter((x) => String(x || "").trim()).length;
      if (!nonEmpty) return null;

      return {
        week: String(get(r, headerMap.week) ?? "").trim(),
        grade: String(get(r, headerMap.grade) ?? "").trim(),
        semester: String(get(r, headerMap.semester) ?? "").trim(),
        subject: String(get(r, headerMap.subject) ?? "").trim(),
        unit: String(get(r, headerMap.unit) ?? "").trim(),
        topic: String(get(r, headerMap.topic) ?? "").trim(),
        hoursRaw: get(r, headerMap.hours),
        _row: r,
      };
    })
    .filter(Boolean);
}

function renderPreview(header, dataRows, limit = 30) {
  const wrap = els.previewWrap;
  if (!wrap) return;
  wrap.innerHTML = "";

  const table = document.createElement("table");
  table.className = "table";
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  header.forEach((h) => {
    const th = document.createElement("th");
    th.textContent = String(h ?? "");
    trh.appendChild(th);
  });
  thead.appendChild(trh);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  (dataRows || []).slice(0, limit).forEach((r) => {
    const tr = document.createElement("tr");
    (r || []).forEach((v) => {
      const td = document.createElement("td");
      td.textContent = String(v ?? "");
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  wrap.appendChild(table);

  const note = document.createElement("div");
  note.className = "hint";
  note.style.marginTop = "10px";
  note.textContent = `미리보기는 상위 ${limit}행만 표시됩니다.`;
  wrap.appendChild(note);
}

function parseSubjectPlan(text){
  const map = new Map();
  const lines = String(text || "").split(/\r?\n/);
  for (const lineRaw of lines){
    const line = String(lineRaw || "").trim();
    if (!line) continue;
    const parts = line.split(/[:：\t]/).map(s=>String(s||"").trim()).filter(Boolean);
    if (parts.length < 2) continue;
    const subject = parts[0];
    const n = Number(String(parts.slice(1).join(" ")).replace(/[^0-9.\-]/g, ""));
    if (!subject || !Number.isFinite(n)) continue;
    map.set(subject, n);
  }
  return map;
}

function bandGrades(band){
  if (band === "1-2") return ["1","2"];
  if (band === "3-4") return ["3","4"];
  if (band === "5-6") return ["5","6"];
  return [];
}

function syncGradeOptions(){
  const band = String(els.gradeBand?.value || "");
  const all = ["1","2","3","4","5","6"];
  const allow = bandGrades(band);
  const options = (allow.length ? allow : all).map(g=>`<option value="${g}">${g}학년</option>`).join("");
  const current = String(els.gradeFilter?.value || "");
  if (els.gradeFilter){
    els.gradeFilter.innerHTML = `<option value="">${band ? "학년군 전체" : "전체 학년"}</option>` + options;
    if ((allow.length && !allow.includes(current)) || (!allow.length && current && !all.includes(current))) {
      els.gradeFilter.value = "";
    } else {
      els.gradeFilter.value = current;
    }
  }
}

function filterByGrade(objs, band, grade){
  const g = String(grade || "").trim();
  const b = String(band || "").trim();
  if (g) return objs.filter(o => String(o.grade || "").trim() === g);
  const allow = bandGrades(b);
  if (!allow.length) return objs;
  return objs.filter(o => allow.includes(String(o.grade || "").trim()));
}

function parseSubjectCurriculum(text){
  return String(text || "").split(/\r?\n/).map(line=>{
    const t = line.trim();
    if (!t) return null;
    const parts = t.split(/[|｜]/).map(x=>x.trim());
    if (!parts[0]) return null;
    return {
      subject: parts[0] || "",
      achievement: parts[1] || "",
      level: parts[2] || ""
    };
  }).filter(Boolean);
}

function renderCurriculum(rows){
  const body = els.curriculumBody;
  if (!body) return;
  body.innerHTML = "";
  if (!rows.length){
    body.innerHTML = `<tr><td colspan="3" class="muted">과목별 교육과정 입력란을 채우면 표시됩니다.</td></tr>`;
    return;
  }
  for (const r of rows){
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escapeHtml(r.subject)}</td><td>${escapeHtml(r.achievement)}</td><td>${escapeHtml(r.level)}</td>`;
    body.appendChild(tr);
  }
}

function isOffType(type){
  const offKeys = ["공휴일","대체공휴일","휴업일","재량휴업","방학","시험"];
  return offKeys.some(k => String(type || "").includes(k));
}

function normalizeScheduleRows(rows){
  const map = new Map();
  for (const r of (rows || [])){
    if (!r?.date || !/^\d{4}-\d{2}-\d{2}$/.test(r.date)) continue;
    const key = `${r.date}|${r.type || ""}|${r.memo || ""}`;
    if (map.has(key)) continue;
    map.set(key, {
      date: r.date,
      type: r.type || "기타",
      memo: r.memo || "",
      isOff: isOffType(r.type || "")
    });
  }
  return Array.from(map.values()).sort((a,b)=>a.date.localeCompare(b.date) || a.type.localeCompare(b.type,'ko-KR'));
}

function renderSchedule(rows){
  const body = els.scheduleBody;
  if (!body) return;
  body.innerHTML = "";
  if (!rows.length){
    body.innerHTML = `<tr><td colspan="4" class="muted">학사 일정을 추가하면 표시됩니다.</td></tr>`;
    return;
  }
  rows.forEach((r, idx)=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escapeHtml(r.date)}</td><td>${escapeHtml(r.type)}</td><td>${escapeHtml(r.memo || "")}</td><td><button type="button" class="btn ghost" data-del-schedule="${idx}" style="padding:6px 10px;">삭제</button></td>`;
    body.appendChild(tr);
  });
}

function addDays(date, n){
  const d = new Date(date.getTime());
  d.setDate(d.getDate() + n);
  return d;
}

function weekAvailableDays(weekNo, termStartDate, scheduleRows){
  const n = Number(String(weekNo).replace(/[^0-9.\-]/g, ""));
  if (!Number.isFinite(n) || n < 1 || !termStartDate) return "";
  const start = addDays(new Date(termStartDate), (Math.floor(n)-1) * 7);
  const offSet = new Set((scheduleRows || []).filter(x=>x.isOff).map(x=>x.date));
  let days = 0;
  for (let i=0;i<5;i++){
    const d = addDays(start, i);
    const key = d.toISOString().slice(0,10);
    if (!offSet.has(key)) days++;
  }
  return days;
}

function renderAnnualTimetable(objs, policy, termStartDate, scheduleRows){
  const wrap = els.timetableWrap;
  if (!wrap) return;
  if (!objs.length){
    wrap.innerHTML = `<div class="muted">선택한 조건에 해당하는 데이터가 없습니다.</div>`;
    return;
  }

  const subjects = Array.from(new Set(objs.map(o=>(o.subject||"").trim()).filter(Boolean))).sort((a,b)=>a.localeCompare(b,'ko-KR'));
  const weekMap = new Map();

  for (const o of objs){
    const w = String(o.week || "").trim() || "미지정";
    const s = String(o.subject || "").trim();
    if (!s) continue;
    const h = parseHourCell(o.hoursRaw, policy) || 0;
    if (!weekMap.has(w)) weekMap.set(w, new Map());
    const m = weekMap.get(w);
    m.set(s, (m.get(s) || 0) + h);
  }

  const weekKeys = Array.from(weekMap.keys()).sort((a,b)=>{
    const na = Number(String(a).replace(/[^0-9.\-]/g, ""));
    const nb = Number(String(b).replace(/[^0-9.\-]/g, ""));
    const fa = Number.isFinite(na), fb = Number.isFinite(nb);
    if (fa && fb) return na - nb;
    if (fa) return -1;
    if (fb) return 1;
    return a.localeCompare(b,'ko-KR');
  });

  const table = document.createElement("table");
  table.className = "table";
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  headRow.innerHTML = `<th>주차</th><th style="text-align:right;">수업가능일</th>${subjects.map(s=>`<th style="text-align:right;">${escapeHtml(s)}</th>`).join("")}<th style="text-align:right;">합계</th>`;
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (const wk of weekKeys){
    const rowMap = weekMap.get(wk);
    let sum = 0;
    const tr = document.createElement("tr");
    const cells = subjects.map(s=>{
      const v = Math.round(((rowMap.get(s) || 0) + Number.EPSILON) * 10) / 10;
      sum += v;
      return `<td style="text-align:right;">${v ? v : ""}</td>`;
    }).join("");
    const avail = weekAvailableDays(wk, termStartDate, scheduleRows);
    tr.innerHTML = `<td>${escapeHtml(wk)}</td><td style="text-align:right;">${avail === "" ? "" : avail}</td>${cells}<td style="text-align:right; font-weight:700;">${Math.round(sum*10)/10}</td>`;
    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  wrap.innerHTML = "";
  wrap.appendChild(table);
}

function computeTotals(objs, policy, planMap) {
  const by = new Map();
  let badHours = 0;

  for (const o of objs) {
    const subj = (o.subject || "").trim();
    if (!subj) continue;
    const h = parseHourCell(o.hoursRaw, policy);
    if (h == null) badHours++;

    if (!by.has(subj)) by.set(subj, { subject: subj, sum: 0, rows: 0, bad: 0 });
    const rec = by.get(subj);
    rec.rows++;
    if (h == null) rec.bad++;
    else rec.sum += h;
  }

  const out = Array.from(by.values()).map((it)=>{
    const planned = planMap?.has(it.subject) ? planMap.get(it.subject) : null;
    const diff = planned == null ? null : (Math.round((it.sum - planned) * 10) / 10);
    return {...it, planned, diff};
  }).sort((a, b) => b.sum - a.sum);

  return { out, badHours, totalRows: objs.length };
}

function renderTotals(items) {
  const body = els.totalsBody;
  if (!body) return;
  body.innerHTML = "";
  if (!items || !items.length) {
    body.innerHTML = `<tr><td colspan="5" class="muted">표시할 데이터가 없습니다.</td></tr>`;
    return;
  }

  for (const it of items) {
    const tr = document.createElement("tr");

    const td1 = document.createElement("td");
    td1.textContent = it.subject;

    const tdPlan = document.createElement("td");
    tdPlan.style.textAlign = "right";
    tdPlan.textContent = it.planned == null ? "-" : (Math.round(it.planned * 10) / 10).toString();

    const td2 = document.createElement("td");
    td2.style.textAlign = "right";
    td2.textContent = (Math.round(it.sum * 10) / 10).toString();

    const tdDiff = document.createElement("td");
    tdDiff.style.textAlign = "right";
    tdDiff.textContent = it.diff == null ? "-" : (it.diff > 0 ? `+${it.diff}` : `${it.diff}`);

    const td3 = document.createElement("td");
    td3.style.textAlign = "right";
    td3.textContent = String(it.rows);

    tr.appendChild(td1);
    tr.appendChild(tdPlan);
    tr.appendChild(td2);
    tr.appendChild(tdDiff);
    tr.appendChild(td3);

    if (it.bad) {
      tr.title = `시수 파싱 실패 ${it.bad}건`;
    }

    body.appendChild(tr);
  }
}

function downloadCSV(objs, totalsBySubject = null) {
  const cols = [
    ["주차", "week"],
    ["학년", "grade"],
    ["학기", "semester"],
    ["교과", "subject"],
    ["단원", "unit"],
    ["학습주제", "topic"],
    ["시수", "hoursRaw"],
    ["편제시수", "planned"],
    ["편제대비차이", "diff"],
  ];

  const rows = [cols.map((c) => c[0])];
  for (const o of objs) {
    const t = totalsBySubject?.get((o.subject || "").trim()) || null;
    rows.push([
      String(o.week ?? ""),
      String(o.grade ?? ""),
      String(o.semester ?? ""),
      String(o.subject ?? ""),
      String(o.unit ?? ""),
      String(o.topic ?? ""),
      String(o.hoursRaw ?? ""),
      t?.planned == null ? "" : String(t.planned),
      t?.diff == null ? "" : String(t.diff),
    ]);
  }

  const csv = rows
    .map((r) =>
      r
        .map((x) => {
          const s = String(x ?? "");
          if (/[",\n]/.test(s)) return `"${s.replaceAll('"', '""')}"`;
          return s;
        })
        .join(",")
    )
    .join("\n");

  const blob = new Blob(["\ufeff" + csv], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "진도표_정리.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
}

async function handleFile(file) {
  if (!file) return;
  if (!window.XLSX) {
    setStatus("엑셀 파서(XLSX) 로딩 중입니다. 잠시 후 다시 시도해 주세요.", { error: true });
    return;
  }

  // default school year from file or current year
  if (els.schoolYear && !String(els.schoolYear.value || "").trim()) {
    els.schoolYear.value = String(new Date().getFullYear());
  }

  setStatus("파일을 읽는 중…");

  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  state.wb = wb;

  const names = wb.SheetNames || [];
  els.sheet.disabled = !names.length;
  els.sheet.innerHTML = names.map((n, i) => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join("");

  const active = names[0] || null;
  state.activeSheet = active;
  if (active) {
    els.sheet.value = active;
    await loadSheet(active);
  }

  els.btnTotals.disabled = false;
  els.btnCSV.disabled = false;
  if (els.btnDocx) els.btnDocx.disabled = false;

  setStatus("파일을 불러왔습니다. 시수 합계를 계산해 보세요.");
  toast("파일을 불러왔습니다.");
}

async function loadSheet(name) {
  if (!state.wb || !name) return;

  const sheet = state.wb.Sheets[name];
  if (!sheet) return;

  // Use sheet_to_json to preserve rows
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  state.rows = rows;

  const header = rows[0] || [];
  const headerMap = detectHeaders(rows);
  state.headerMap = headerMap;

  renderPreview(header, rows.slice(1));

  // warn if required keys missing
  const missing = [];
  if (headerMap?.subject == null) missing.push("교과");
  if (headerMap?.hours == null) missing.push("시수");
  if (missing.length) {
    setStatus(`헤더 인식이 완전하지 않습니다. 누락: ${missing.join(", ")} (그래도 미리보기는 가능합니다)`, { error: true });
  } else {
    setStatus("시트를 불러왔습니다.");
  }

  if (!missing.length) applyTotals();
}

function escapeHtml(s) {
  return String(s || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

els.file?.addEventListener("change", async (e) => {
  const f = e.target.files && e.target.files[0];
  await handleFile(f);
});

els.sheet?.addEventListener("change", async () => {
  const name = els.sheet.value;
  state.activeSheet = name;
  await loadSheet(name);
});

function applyTotals(){
  if (!state.rows || !state.headerMap) return null;

  const policy = els.rangePolicy?.value || "avg";
  const gradeBand = String(els.gradeBand?.value || "").trim();
  const grade = String(els.gradeFilter?.value || "").trim();
  const planMap = parseSubjectPlan(els.subjectPlan?.value || "");
  const scheduleRows = state.scheduleRows || [];
  const termStartDate = els.termStartDate?.value || "";

  const objsAll = rowsToObjects(state.rows, state.headerMap);
  const objs = filterByGrade(objsAll, gradeBand, grade);
  const { out, badHours, totalRows } = computeTotals(objs, policy, planMap);
  state.lastTotals = out;

  renderTotals(out);
  renderCurriculum(parseSubjectCurriculum(els.subjectCurriculum?.value || ""));
  renderSchedule(scheduleRows);
  renderAnnualTimetable(objs, policy, termStartDate, scheduleRows);

  const bandLabel = gradeBand ? `${gradeBand}학년군 / ` : "";
  const gradeLabel = grade ? `${grade}학년 / ` : bandLabel;
  if (badHours) {
    setStatus(`완료: ${gradeLabel}${out.length}개 교과 합산. 시수 파싱 실패 ${badHours}건(총 ${totalRows}행).`, { error: true });
  } else {
    setStatus(`완료: ${gradeLabel}${out.length}개 교과 합산(총 ${totalRows}행).`);
  }

  return { objs, out };
}

els.btnTotals?.addEventListener("click", () => {
  applyTotals();
});

els.btnCSV?.addEventListener("click", () => {
  if (!state.rows || !state.headerMap) return;
  const result = applyTotals();
  if (!result) return;
  const map = new Map((result.out || []).map(x=>[x.subject, x]));
  downloadCSV(result.objs, map);
});

async function downloadDOCX(){
  if (!state.rows || !state.headerMap) return;
  if (!window.docx) {
    toast("DOCX 라이브러리 로딩 중입니다. 잠시 후 다시 시도해 주세요.", 2200);
    return;
  }

  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, HeadingLevel } = window.docx;

  const policy = els.rangePolicy?.value || "avg";
  const gradeBand = String(els.gradeBand?.value || "").trim();
  const gradeSelected = String(els.gradeFilter?.value || "").trim();
  const planMap = parseSubjectPlan(els.subjectPlan?.value || "");
  const curriculumRows = parseSubjectCurriculum(els.subjectCurriculum?.value || "");
  const scheduleRows = state.scheduleRows || [];

  const objsAll = rowsToObjects(state.rows, state.headerMap);
  const objs = filterByGrade(objsAll, gradeBand, gradeSelected);
  const { out: totals } = computeTotals(objs, policy, planMap);

  const numOrInf = (s)=>{
    const n = parseFloat(String(s||"").replace(/[^0-9.\-]/g, ""));
    return Number.isFinite(n) ? n : Infinity;
  };
  objs.sort((a,b)=>{
    const wa = numOrInf(a.week), wb = numOrInf(b.week);
    if (wa !== wb) return wa - wb;
    const sa = (a.subject||"");
    const sb = (b.subject||"");
    const c = sa.localeCompare(sb, 'ko-KR');
    if (c) return c;
    return (a.unit||"").localeCompare(b.unit||"", 'ko-KR');
  });

  const grade = (gradeSelected || objs.find(o=>o.grade)?.grade || "").trim();
  const semester = (objs.find(o=>o.semester)?.semester || "").trim();
  const year = String(els.schoolYear?.value || new Date().getFullYear()).trim();

  if (!objs.length){
    toast("선택한 학년에 해당하는 데이터가 없어 DOCX를 만들 수 없습니다.", 2400);
    return;
  }

  const title = `${year}학년도 ${grade?grade+"학년 ":""}${semester?semester+"학기 ":""}학교교육과정 편성·운영 및 평가 계획`;
  const meta = `생성일: ${new Date().toISOString().slice(0,10)} / 시수범위: ${policy}`;

  const makeTable = (headers, rows)=> new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: headers.map(h=>new TableCell({ children:[new Paragraph({children:[new TextRun({text:h, bold:true})]})] })) }),
      ...rows.map(r=>new TableRow({ children: r.map(v=>new TableCell({ children:[new Paragraph(String(v ?? ""))] })) }))
    ]
  });

  const totalsTable = makeTable(
    ["교과", "편제", "합계", "차이", "행 개수"],
    (totals.length ? totals : [{subject:"(데이터 없음)", planned:"", sum:"", diff:"", rows:""}]).map(t=>[
      t.subject || "",
      t.planned == null ? "-" : t.planned,
      Math.round((t.sum || 0) * 10) / 10,
      t.diff == null ? "-" : (t.diff > 0 ? `+${t.diff}` : `${t.diff}`),
      t.rows || ""
    ])
  );

  const curriculumTable = makeTable(
    ["과목", "성취기준(핵심내용)", "성취수준"],
    (curriculumRows.length ? curriculumRows : [{subject:"(입력 없음)", achievement:"", level:""}]).map(r=>[r.subject, r.achievement, r.level])
  );

  const scheduleTable = makeTable(
    ["날짜", "유형", "메모"],
    (scheduleRows.length ? scheduleRows : [{date:"(입력 없음)", type:"", memo:""}]).map(r=>[r.date, r.type, r.memo || ""])
  );

  const progressTable = makeTable(
    ["주차","교과","단원","학습주제","시수"],
    objs.map(o=>{
      const h = parseHourCell(o.hoursRaw, policy);
      return [
        o.week,
        o.subject,
        o.unit,
        o.topic,
        (h==null? String(o.hoursRaw??"") : (Math.round(h*10)/10).toString()),
      ];
    })
  );

  const children = [
    new Paragraph({ text: title, heading: HeadingLevel.HEADING_1 }),
    new Paragraph({ children: [ new TextRun({ text: meta, color: "666666" }) ] }),
    new Paragraph({ text: "" }),

    new Paragraph({ text: "1. 학교교육과정 편성·운영의 기본 방향", heading: HeadingLevel.HEADING_2 }),
    new Paragraph(`- 대상: ${grade ? grade + "학년" : (gradeBand ? gradeBand + "학년군" : "전체")}`),
    new Paragraph("- 2022 개정 교육과정 및 학교알리미 2-가 공시항목 기준에 따라 작성함."),
    new Paragraph(""),

    new Paragraph({ text: "2. 편제 및 시간 배당", heading: HeadingLevel.HEADING_2 }),
    totalsTable,
    new Paragraph(""),

    new Paragraph({ text: "3. 교육과정 편성·운영 계획", heading: HeadingLevel.HEADING_2 }),
    curriculumTable,
    new Paragraph(""),

    new Paragraph({ text: "4. 학교교육과정 편성·운영 평가 계획", heading: HeadingLevel.HEADING_2 }),
    new Paragraph("- 학기 중(4월/9월) 운영 점검 및 성취기준 도달도 점검을 실시한다."),
    new Paragraph("- 교과별 계획 대비 실제 운영 시수, 학사일정 변동, 평가 결과를 반영해 개선한다."),
    new Paragraph(""),

    new Paragraph({ text: "5. 연간학사일정", heading: HeadingLevel.HEADING_2 }),
    scheduleTable,
    new Paragraph(""),

    new Paragraph({ text: "부록. 주차별 진도표", heading: HeadingLevel.HEADING_2 }),
    progressTable,
  ];

  const doc = new Document({
    sections: [{ properties: {}, children }]
  });

  const blob = await Packer.toBlob(doc);
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = title.replace(/\s+/g,'_') + '.docx';
  document.body.appendChild(a);
  a.click();
  a.remove();
}

function refreshSchedule(rows){
  state.scheduleRows = normalizeScheduleRows(rows);
  renderSchedule(state.scheduleRows);
  if (state.rows && state.headerMap) applyTotals();
}

function addScheduleFromInputs(){
  const date = String(els.scheduleDate?.value || "").trim();
  const type = String(els.scheduleType?.value || "기타").trim() || "기타";
  const memo = String(els.scheduleMemo?.value || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)){
    toast("날짜를 먼저 선택해 주세요.", 1800);
    return;
  }
  refreshSchedule([...(state.scheduleRows || []), {date, type, memo, isOff: isOffType(type)}]);
  if (els.scheduleMemo) els.scheduleMemo.value = "";
}

async function loadNationalHolidays(){
  const y = Number(els.schoolYear?.value || new Date().getFullYear());
  try {
    setStatus("국가공휴일을 불러오는 중...");
    const res = await fetch(`https://date.nager.at/api/v3/PublicHolidays/${y}/KR`);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const arr = await res.json();
    const rows = (Array.isArray(arr) ? arr : []).map(x=>{
      const name = String(x.localName || x.name || "공휴일");
      const t = /(대체|대체휴일|Alternative|Substitute)/i.test(name) ? "대체공휴일" : "공휴일";
      return { date: x.date, type: t, memo: name, isOff: true };
    });
    refreshSchedule([...(state.scheduleRows || []), ...rows]);
    setStatus(`국가공휴일 ${rows.length}건을 반영했습니다.`);
    toast("국가공휴일 자동 반영 완료", 1800);
  } catch (e) {
    setStatus("국가공휴일 자동 불러오기에 실패했습니다. 직접 입력해 주세요.", { error: true });
    toast("공휴일 자동 불러오기 실패", 2000);
  }
}

els.btnAddSchedule?.addEventListener('click', addScheduleFromInputs);
els.btnClearSchedule?.addEventListener('click', ()=>refreshSchedule([]));
els.btnLoadNational?.addEventListener('click', ()=>{ loadNationalHolidays(); });
els.scheduleBody?.addEventListener('click', (e)=>{
  const btn = e.target.closest('[data-del-schedule]');
  if (!btn) return;
  const idx = Number(btn.getAttribute('data-del-schedule'));
  if (!Number.isFinite(idx)) return;
  const next = state.scheduleRows.slice();
  next.splice(idx, 1);
  refreshSchedule(next);
});

[els.rangePolicy, els.gradeBand, els.gradeFilter].forEach((el)=>{
  el?.addEventListener('change', ()=>{
    if (el === els.gradeBand) syncGradeOptions();
    if (state.rows && state.headerMap) applyTotals();
  });
});
[els.subjectPlan, els.subjectCurriculum].forEach((el)=>{
  el?.addEventListener('input', ()=>{
    if (state.rows && state.headerMap) applyTotals();
  });
});
els.termStartDate?.addEventListener('change', ()=>{
  if (state.rows && state.headerMap) applyTotals();
});

syncGradeOptions();
renderCurriculum(parseSubjectCurriculum(els.subjectCurriculum?.value || ""));
renderSchedule(state.scheduleRows);
if (els.termStartDate && !els.termStartDate.value){
  const y = Number(els.schoolYear?.value || new Date().getFullYear());
  els.termStartDate.value = `${y}-03-02`;
}

els.btnGuide?.addEventListener('click', openGuide);
els.btnGuideSkip?.addEventListener('click', closeGuide);
els.btnGuideNext?.addEventListener('click', ()=>{
  if (guideIndex >= guideSteps.length - 1) { closeGuide(); return; }
  guideIndex++;
  showGuideStep();
});
els.guideOverlay?.addEventListener('click', (e)=>{
  if (e.target === els.guideOverlay) closeGuide();
});
document.addEventListener('keydown', (e)=>{
  if (e.key === 'Escape' && els.guideOverlay?.classList.contains('show')) closeGuide();
});

els.btnDocx?.addEventListener('click', ()=>{
  downloadDOCX().catch(()=>toast('DOCX 생성에 실패했습니다.', 2400));
});
