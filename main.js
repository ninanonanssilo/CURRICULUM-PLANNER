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
  btnTotals: $("btnTotals"),
  btnCSV: $("btnCSV"),
  status: $("status"),
  totalsBody: $("totalsBody"),
  previewWrap: $("previewWrap"),
  toast: $("toast"),
};

const state = {
  wb: null,
  rows: null,
  headerMap: null,
  activeSheet: null,
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

function computeTotals(objs, policy) {
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

  const out = Array.from(by.values()).sort((a, b) => b.sum - a.sum);
  return { out, badHours, totalRows: objs.length };
}

function renderTotals(items) {
  const body = els.totalsBody;
  if (!body) return;
  body.innerHTML = "";
  if (!items || !items.length) {
    body.innerHTML = `<tr><td colspan="3" class="muted">표시할 데이터가 없습니다.</td></tr>`;
    return;
  }

  for (const it of items) {
    const tr = document.createElement("tr");

    const td1 = document.createElement("td");
    td1.textContent = it.subject;

    const td2 = document.createElement("td");
    td2.style.textAlign = "right";
    td2.textContent = (Math.round(it.sum * 10) / 10).toString();

    const td3 = document.createElement("td");
    td3.style.textAlign = "right";
    td3.textContent = String(it.rows);

    tr.appendChild(td1);
    tr.appendChild(td2);
    tr.appendChild(td3);

    if (it.bad) {
      tr.title = `시수 파싱 실패 ${it.bad}건`;
    }

    body.appendChild(tr);
  }
}

function downloadCSV(objs, headerMap) {
  const cols = [
    ["주차", "week"],
    ["학년", "grade"],
    ["학기", "semester"],
    ["교과", "subject"],
    ["단원", "unit"],
    ["학습주제", "topic"],
    ["시수", "hoursRaw"],
  ];

  const rows = [cols.map((c) => c[0])];
  for (const o of objs) {
    rows.push(cols.map((c) => String(o[c[1]] ?? "")));
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

els.btnTotals?.addEventListener("click", () => {
  if (!state.rows || !state.headerMap) return;
  const policy = els.rangePolicy?.value || "avg";
  const objs = rowsToObjects(state.rows, state.headerMap);
  const { out, badHours, totalRows } = computeTotals(objs, policy);

  renderTotals(out);
  if (badHours) {
    setStatus(`완료: ${out.length}개 교과 합산. 시수 파싱 실패 ${badHours}건(총 ${totalRows}행). 시수 칸 형식을 확인하세요.`, { error: true });
  } else {
    setStatus(`완료: ${out.length}개 교과 합산(총 ${totalRows}행).`);
  }
});

els.btnCSV?.addEventListener("click", () => {
  if (!state.rows || !state.headerMap) return;
  const objs = rowsToObjects(state.rows, state.headerMap);
  downloadCSV(objs, state.headerMap);
});
