/* ─── 정산 검증 페이지 스크립트 v3 ─── */

let currentSessionId = null;
let currentFilter = "ALL";
let allRows = [];          // 현재 세션+필터의 전체 결과
let sortState = { col: "row_number", dir: "asc" };
const EMPTY_COLS = 29;     // colspan for empty message

document.addEventListener("DOMContentLoaded", () => {
  setupDropZone();
  loadSessionList();

  const params = new URLSearchParams(location.search);
  const sid = params.get("session");
  if (sid) loadSession(sid);
});

// ─── 드래그앤드롭 ─────────────────────────────────────────
function setupDropZone() {
  const zone  = document.getElementById("drop-zone");
  const input = document.getElementById("file-input");

  zone.addEventListener("click", () => input.click());
  zone.addEventListener("dragover",  (e) => { e.preventDefault(); zone.classList.add("drag-over"); });
  zone.addEventListener("dragleave", ()  => zone.classList.remove("drag-over"));
  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    zone.classList.remove("drag-over");
    if (e.dataTransfer.files.length) selectFile(e.dataTransfer.files[0]);
  });
  input.addEventListener("change", () => { if (input.files.length) selectFile(input.files[0]); });
}

function selectFile(file) {
  document.getElementById("drop-zone").style.display = "none";
  document.getElementById("file-selected").style.display = "flex";
  document.getElementById("file-name").textContent = file.name;
  document.getElementById("upload-btn").onclick = () => uploadFile(file);
}

function resetUpload() {
  document.getElementById("drop-zone").style.display = "";
  document.getElementById("file-selected").style.display = "none";
  document.getElementById("file-input").value = "";
  document.getElementById("upload-error").style.display = "none";
}

async function uploadFile(file) {
  if (!file) {
    file = document.getElementById("file-input").files[0];
  }
  if (!file) return;

  document.getElementById("file-selected").style.display = "none";
  document.getElementById("upload-progress").style.display = "block";
  document.getElementById("upload-error").style.display = "none";

  const form = new FormData();
  form.append("file", file);

  try {
    const res = await fetch("/api/verification/upload", { method: "POST", body: form });
    document.getElementById("upload-progress").style.display = "none";
    if (!res.ok) {
      const err = await res.json();
      showError(err.detail || "업로드 실패");
      document.getElementById("drop-zone").style.display = "";
      return;
    }
    const session = await res.json();
    resetUpload();
    await loadSessionList();
    renderSummary(session);
    currentSessionId = session.id;
    await loadResults();
  } catch (e) {
    document.getElementById("upload-progress").style.display = "none";
    showError("네트워크 오류: " + e.message);
    document.getElementById("drop-zone").style.display = "";
  }
}

// ─── 세션 목록 ───────────────────────────────────────────
async function loadSessionList() {
  const res = await fetch("/api/verification/sessions");
  const sessions = await res.json();
  const sel = document.getElementById("session-select");
  sel.innerHTML = '<option value="">세션 선택...</option>' +
    sessions.map(s => `<option value="${s.id}">[${s.id}] ${s.filename} (${fmtDate(s.uploaded_at)})</option>`).join("");
}

async function loadSession(id) {
  if (!id) return;
  currentSessionId = parseInt(id);
  document.getElementById("session-select").value = id;
  const res = await fetch(`/api/verification/sessions/${id}`);
  if (!res.ok) return;
  const session = await res.json();
  renderSummary(session);
  await loadResults();
}

// ─── 요약 렌더 ───────────────────────────────────────────
function renderSummary(s) {
  document.getElementById("summary-section").style.display = "";
  document.getElementById("summary-filename").textContent = s.filename;
  document.getElementById("s-total").textContent = s.total_rows.toLocaleString();
  document.getElementById("s-trkv-pass").textContent = s.trkv_pass;
  document.getElementById("s-trkv-fail").textContent = s.trkv_fail;
  document.getElementById("s-trkv-no-rate").textContent = s.trkv_no_rate;
  document.getElementById("s-storage-pass").textContent = s.storage_pass;
  document.getElementById("s-storage-fail").textContent = s.storage_fail;
  document.getElementById("s-storage-no-rate").textContent = s.storage_no_rate;
  document.getElementById("s-handling-pass").textContent = s.handling_pass;
  document.getElementById("s-handling-fail").textContent = s.handling_fail;
  document.getElementById("s-handling-no-rate").textContent = s.handling_no_rate;
  document.getElementById("s-shuttle-pass").textContent = s.shuttle_pass;
  document.getElementById("s-shuttle-fail").textContent = s.shuttle_fail;
  document.getElementById("s-shuttle-no-rate").textContent = s.shuttle_no_rate;
  const diff = s.total_diff || 0;
  const diffEl = document.getElementById("s-total-diff");
  diffEl.textContent = fmtMoney(diff);
  diffEl.style.color = diff > 0 ? "#c62828" : "#1e7e34";
}

// ─── 필터 ────────────────────────────────────────────────
function setFilter(filter, btn) {
  currentFilter = filter;
  document.querySelectorAll(".filter-btn").forEach(b => b.classList.remove("active"));
  btn.classList.add("active");
  loadResults();
}

// ─── 결과 로드 (전체 한번에) ─────────────────────────────
async function loadResults() {
  if (!currentSessionId) return;
  const url = `/api/verification/sessions/${currentSessionId}/results?status_filter=${currentFilter}&skip=0&limit=99999`;
  const res = await fetch(url);
  allRows = await res.json();
  applySortAndRender();
}

// ─── 정렬 ────────────────────────────────────────────────
function sortBy(col) {
  if (sortState.col === col) {
    sortState.dir = sortState.dir === "asc" ? "desc" : "asc";
  } else {
    sortState.col = col;
    sortState.dir = "asc";
  }
  applySortAndRender();
}

function applySortAndRender() {
  const sorted = [...allRows].sort((a, b) => {
    let av = a[sortState.col];
    let bv = b[sortState.col];
    if (av == null && bv == null) return 0;
    if (av == null) return 1;
    if (bv == null) return -1;
    const na = parseFloat(av), nb = parseFloat(bv);
    const cmp = (!isNaN(na) && !isNaN(nb))
      ? na - nb
      : String(av).localeCompare(String(bv), "ko");
    return sortState.dir === "asc" ? cmp : -cmp;
  });
  renderResults(sorted);
  updateSortArrows();
}

function updateSortArrows() {
  document.querySelectorAll(".sortable-th").forEach(th => {
    const match = (th.getAttribute("onclick") || "").match(/sortBy\('([^']+)'\)/);
    if (!match) return;
    const key = match[1];
    const arrow = th.querySelector(".sort-arrow");
    if (!arrow) return;
    arrow.textContent = sortState.col === key
      ? (sortState.dir === "asc" ? " ↑" : " ↓")
      : " ↕";
  });
}

// ─── 결과 테이블 렌더 (전체 행, 페이지 없음) ─────────────
function renderResults(sorted) {
  const tbody = document.getElementById("results-tbody");
  const countEl = document.getElementById("result-count");

  if (!sorted.length) {
    tbody.innerHTML = `<tr><td colspan="${EMPTY_COLS}" class="empty-msg">결과가 없습니다.</td></tr>`;
    if (countEl) countEl.textContent = "";
    return;
  }

  tbody.innerHTML = sorted.map(r => renderRow(r)).join("");
  if (countEl) countEl.textContent = `총 ${sorted.length.toLocaleString()}건`;
}

function renderRow(r) {
  const rowClass = r.overall_status === "DIFF" ? "row-diff"
                 : r.overall_status === "NO_RATE" ? "row-no-rate" : "";

  const pickupPort = r.pickup_port_resolved
    ? `<span class="port-resolved">${r.pickup_port_resolved}</span>`
    : "-";
  const destPort = r.dest_port_resolved
    ? `<span class="port-resolved">${r.dest_port_resolved}</span>`
    : "-";

  const qty = r.quantity != null ? Number(r.quantity).toLocaleString("ko-KR") : "1";

  return `<tr class="${rowClass}">
    <td>${r.row_number}</td>
    <td>${r.container_no || "-"}</td>
    <td>${r.transport_date || "-"}</td>
    <!-- 운송 구간 정보 (9열) -->
    <td>${r.pickup_name || "-"}</td>
    <td>${pickupPort}</td>
    <td>${r.departure_name || "-"}</td>
    <td>${r.departure_code_resolved ? `<span class="port-resolved">${r.departure_code_resolved}</span>` : "-"}</td>
    <td title="${r.odcy_name || ""}">${r.odcy_code || r.odcy_name || "-"}</td>
    <td>${r.dest_name || "-"}</td>
    <td>${destPort}</td>
    <td>${r.container_type || "-"}</td>
    <td class="money">${qty}</td>
    <!-- TRKV -->
    ${chargeCell(r.trkv_actual, r.trkv_expected, r.trkv_diff, r.trkv_status)}
    <!-- 보관료 -->
    ${chargeCell(r.storage_actual, r.storage_expected, r.storage_diff, r.storage_status)}
    <!-- 상하차료 -->
    ${chargeCell(r.handling_actual, r.handling_expected, r.handling_diff, r.handling_status)}
    <!-- 셔틀비용 -->
    ${chargeCell(r.shuttle_actual, r.shuttle_expected, r.shuttle_diff, r.shuttle_status)}
    <!-- 종합 -->
    <td class="${statusClass(r.overall_status)}">${r.overall_status || "-"}</td>
  </tr>`;
}

function chargeCell(actual, expected, diff, status) {
  const sc = statusClass(status);
  const diffStr = diff != null ? (diff >= 0 ? "+" : "") + fmtMoney(diff) : "-";
  const diffColor = diff > 0.5 ? 'style="color:#c62828"' : diff < -0.5 ? 'style="color:#1a73e8"' : "";
  return `
    <td class="money">${fmtMoney(actual)}</td>
    <td class="money">${expected != null ? fmtMoney(expected) : "-"}</td>
    <td class="money" ${diffColor}>${diffStr}</td>
    <td class="${sc}">${status || "-"}</td>
  `;
}

function statusClass(s) {
  return { OK: "status-ok", DIFF: "status-diff", NO_RATE: "status-no-rate", SKIP: "status-skip" }[s] || "";
}

function exportExcel() {
  if (!currentSessionId) return alert("세션을 선택하세요.");
  window.location.href = `/api/verification/sessions/${currentSessionId}/export`;
}

// ─── helpers ─────────────────────────────────────────────
function fmtMoney(n) {
  if (n == null) return "-";
  return Number(n).toLocaleString("ko-KR") + "원";
}
function fmtDate(s) {
  if (!s) return "-";
  return s.slice(0, 16).replace("T", " ");
}
function showError(msg) {
  const el = document.getElementById("upload-error");
  el.textContent = msg;
  el.style.display = "block";
}
