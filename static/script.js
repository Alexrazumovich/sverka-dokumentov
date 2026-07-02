"use strict";

let fileA = null, fileB = null, downloadUrl = null;
let hasHeaderA = true, hasHeaderB = true;

const $ = id => document.getElementById(id);

const steps = {
  upload:  $("step-upload"),
  columns: $("step-columns"),
  loading: $("step-loading"),
  result:  $("step-result"),
};

function showStep(name) {
  Object.values(steps).forEach(s => s.classList.add("hidden"));
  steps[name].classList.remove("hidden");
  $("error-box").classList.add("hidden");
}

function showError(msg) {
  const box = $("error-box");
  box.textContent = "Ошибка: " + msg;
  box.classList.remove("hidden");
}

// ── File drop zones ──────────────────────────────────────────────

function setupZone(zoneId, inputId, which) {
  const zone  = $(zoneId);
  const input = $(inputId);

  // <label> handles click → file dialog natively; no JS needed for that

  zone.addEventListener("dragover", e => {
    e.preventDefault();
    zone.classList.add("dragover");
  });
  zone.addEventListener("dragleave", e => {
    if (!zone.contains(e.relatedTarget)) zone.classList.remove("dragover");
  });
  zone.addEventListener("drop", e => {
    e.preventDefault();
    zone.classList.remove("dragover");
    const file = e.dataTransfer.files[0];
    if (file) applyFile(file, which);
  });

  input.addEventListener("change", () => {
    if (input.files[0]) applyFile(input.files[0], which);
  });
}

function applyFile(file, which) {
  if (!/\.xlsx?$/i.test(file.name)) {
    showError("Поддерживаются только .xlsx и .xls");
    return;
  }
  $("error-box").classList.add("hidden");
  if (which === "a") {
    fileA = file;
    $("name-a").textContent = file.name;
    $("drop-a").classList.add("has-file");
  } else {
    fileB = file;
    $("name-b").textContent = file.name;
    $("drop-b").classList.add("has-file");
  }
  $("btn-preview").disabled = !(fileA && fileB);
}

setupZone("drop-a", "input-a", "a");
setupZone("drop-b", "input-b", "b");

// "No header" checkboxes
$("no-header-a").addEventListener("change", e => { hasHeaderA = !e.target.checked; });
$("no-header-b").addEventListener("change", e => { hasHeaderB = !e.target.checked; });

// ── Step 1 → 2: preview columns ─────────────────────────────────

$("btn-preview").addEventListener("click", async () => {
  showStep("loading");
  $("loading-text").textContent = "Читаем заголовки файлов...";

  const fd = new FormData();
  fd.append("file_a",       fileA);
  fd.append("file_b",       fileB);
  fd.append("has_header_a", hasHeaderA ? "true" : "false");
  fd.append("has_header_b", hasHeaderB ? "true" : "false");

  try {
    const res  = await fetch("/api/preview", { method: "POST", body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || "Неизвестная ошибка");

    $("name-a-lbl").textContent = data.name_a;
    $("name-b-lbl").textContent = data.name_b;
    populateSelect("date-col-a",   data.columns_a, guessDate(data.columns_a));
    populateSelect("amount-col-a", data.columns_a, guessAmount(data.columns_a));
    populateSelect("date-col-b",   data.columns_b, guessDate(data.columns_b));
    populateSelect("amount-col-b", data.columns_b, guessAmount(data.columns_b));

    showStep("columns");
  } catch (e) {
    showStep("upload");
    showError(e.message);
  }
});

function populateSelect(id, cols, defaultVal) {
  const sel = $(id);
  sel.innerHTML = "";
  cols.forEach(c => {
    const opt = document.createElement("option");
    opt.value = c;
    opt.textContent = c;
    if (c === defaultVal) opt.selected = true;
    sel.appendChild(opt);
  });
}

function guessDate(cols) {
  const kw = ["дата", "date", "период", "день", "time"];
  return cols.find(c => kw.some(k => String(c).toLowerCase().includes(k))) ?? cols[0];
}

function guessAmount(cols) {
  const kw = ["сумма", "amount", "итог", "деньги", "руб", "rub", "value", "цена"];
  return (
    cols.find(c => kw.some(k => String(c).toLowerCase().includes(k))) ??
    (cols.length > 1 ? cols[1] : cols[0])
  );
}

// ── Back ─────────────────────────────────────────────────────────

$("btn-back").addEventListener("click", () => showStep("upload"));

// ── Step 2 → result: reconcile ───────────────────────────────────

$("btn-reconcile").addEventListener("click", async () => {
  showStep("loading");
  $("loading-text").textContent = "Выполняется сверка...";

  const tol = document.querySelector("input[name='tol']:checked").value;
  const fd  = new FormData();
  fd.append("file_a",         fileA);
  fd.append("file_b",         fileB);
  fd.append("date_col_a",     $("date-col-a").value);
  fd.append("amount_col_a",   $("amount-col-a").value);
  fd.append("date_col_b",     $("date-col-b").value);
  fd.append("amount_col_b",   $("amount-col-b").value);
  fd.append("date_tolerance", tol);
  fd.append("has_header_a",   hasHeaderA ? "true" : "false");
  fd.append("has_header_b",   hasHeaderB ? "true" : "false");

  try {
    const res  = await fetch("/api/reconcile", { method: "POST", body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || "Неизвестная ошибка");

    downloadUrl = data.download_url;
    renderSummary(data.summary);
    showStep("result");
  } catch (e) {
    showStep("columns");
    showError(e.message);
  }
});

// ── Summary ───────────────────────────────────────────────────────

function fmt(n) {
  return Number(n).toLocaleString("ru-RU", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function card(val, label, cls = "") {
  return `<div class="summary-card">
    <div class="val ${cls}">${val}</div>
    <div class="lbl">${label}</div>
  </div>`;
}

function renderSummary(s) {
  const diffCls = Math.abs(s.diff) < 0.005 ? "ok" : "bad";
  const unmACls = s.unmatched_a === 0 ? "ok" : "bad";
  const unmBCls = s.unmatched_b === 0 ? "ok" : "bad";

  $("summary-grid").innerHTML = [
    card(fmt(s.total_a),    `Итого сумма А (${s.rows_a} стр.)`),
    card(fmt(s.total_b),    `Итого сумма Б (${s.rows_b} стр.)`),
    card(fmt(s.diff),       "Разница итогов", diffCls),
    card(s.exact,           "Точных совпадений", "ok"),
    card(s.tolerance,       s.workday ? "По раб. дню (пт→пн)" : "По допуску даты"),
    card(s.group,           "Групповых совпадений"),
    card(s.total_pairs,     "Совпавших пар всего"),
    card(s.unmatched_a,     "Не найдено в Б", unmACls),
    card(s.unmatched_b,     "Не найдено в А", unmBCls),
  ].join("");
}

// ── Download & restart ────────────────────────────────────────────

$("btn-download").addEventListener("click", () => {
  if (downloadUrl) window.location.href = downloadUrl;
});

$("btn-restart").addEventListener("click", () => {
  fileA = fileB = downloadUrl = null;
  hasHeaderA = true; hasHeaderB = true;
  ["name-a", "name-b"].forEach(id => $(id).textContent = "");
  ["drop-a", "drop-b"].forEach(id => $(id).classList.remove("has-file"));
  ["input-a", "input-b"].forEach(id => $(id).value = "");
  ["no-header-a", "no-header-b"].forEach(id => $(id).checked = false);
  $("btn-preview").disabled = true;
  showStep("upload");
});

showStep("upload");
