/* ─── TRKV 요율 설정 페이지 스크립트 ─── */

const CONT_TYPES = ["22G1", "22R1", "45G1", "45R1"];
const DG_OPTIONS = [false, true];  // false=없음, true=X

// ──────────────────────────────────────────────────────────────────
// 공통 유틸
// ──────────────────────────────────────────────────────────────────
function fmt(val) {
  if (val === null || val === undefined || val === "") return "-";
  return Number(val).toLocaleString("ko-KR");
}

function showMsg(elId, msg, isOk) {
  const el = document.getElementById(elId);
  el.textContent = msg;
  el.className = "form-msg " + (isOk ? "msg-ok" : "msg-error");
  el.style.display = "inline-block";
  setTimeout(() => { el.style.display = "none"; }, 3500);
}

function portBadge(port) {
  const cls = port === "부산신항" ? "badge-blue" : "badge-purple";
  return `<span class="badge ${cls}">${port}</span>`;
}

// ──────────────────────────────────────────────────────────────────
// ① 포트명 매핑
// ──────────────────────────────────────────────────────────────────
let pmData = [];

async function loadPortMappings() {
  const res = await fetch("/api/trkv/port-mappings");
  pmData = await res.json();
  renderPm();
}

function renderPm() {
  const tbody = document.getElementById("pm-tbody");
  if (!pmData.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="empty-msg">등록된 포트 매핑이 없습니다.</td></tr>';
    return;
  }
  tbody.innerHTML = pmData.map((d, i) => `
    <tr>
      <td><input type="checkbox" class="pm-chk" data-id="${d.id}" onchange="updatePmSelCount()" /></td>
      <td>${i + 1}</td>
      <td>${d.excel_name}</td>
      <td>${portBadge(d.port_type)}</td>
      <td>
        <button class="btn btn-sm btn-outline" onclick="startPmEdit(${d.id})">수정</button>
        <button class="btn btn-sm btn-danger" onclick="deletePm(${d.id})">삭제</button>
      </td>
    </tr>
  `).join("");
  updatePmSelCount();
}

function updatePmSelCount() {
  const checked = document.querySelectorAll(".pm-chk:checked").length;
  document.getElementById("pm-sel-count").textContent = checked;
  document.getElementById("pm-action-bar").style.display = checked > 0 ? "flex" : "none";
  // 전체 선택 체크박스 동기화
  const all = document.querySelectorAll(".pm-chk");
  document.getElementById("pm-check-all").checked = all.length > 0 && checked === all.length;
}

function toggleAllPmCheck(checked) {
  document.querySelectorAll(".pm-chk").forEach(cb => { cb.checked = checked; });
  updatePmSelCount();
}

async function deleteSelectedPm() {
  const ids = [...document.querySelectorAll(".pm-chk:checked")].map(cb => parseInt(cb.dataset.id));
  if (!ids.length) return;
  if (!confirm(`선택한 ${ids.length}건을 삭제하시겠습니까?`)) return;
  for (const id of ids) {
    await fetch(`/api/trkv/port-mappings/${id}`, { method: "DELETE" });
  }
  showMsg("pm-msg", `${ids.length}건 삭제되었습니다.`, true);
  loadPortMappings();
}

function startPmEdit(id) {
  const d = pmData.find(x => x.id === id);
  if (!d) return;
  document.getElementById("pm-edit-id").value = id;
  document.getElementById("pm-excel-name").value = d.excel_name;
  document.getElementById("pm-port-type").value = d.port_type;
  document.getElementById("pm-submit-btn").textContent = "수정";
  document.getElementById("pm-cancel-btn").style.display = "inline-flex";
}

function cancelPmEdit() {
  document.getElementById("pm-edit-id").value = "";
  document.getElementById("pm-excel-name").value = "";
  document.getElementById("pm-port-type").value = "";
  document.getElementById("pm-submit-btn").textContent = "추가";
  document.getElementById("pm-cancel-btn").style.display = "none";
}

document.getElementById("pm-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const editId = document.getElementById("pm-edit-id").value;
  const body = {
    excel_name: document.getElementById("pm-excel-name").value.trim(),
    port_type:  document.getElementById("pm-port-type").value,
  };
  if (!body.excel_name || !body.port_type) return;

  const url    = editId ? `/api/trkv/port-mappings/${editId}` : "/api/trkv/port-mappings";
  const method = editId ? "PUT" : "POST";
  const res    = await fetch(url, { method, headers: {"Content-Type": "application/json"}, body: JSON.stringify(body) });

  if (res.ok) {
    showMsg("pm-msg", editId ? "수정되었습니다." : "추가되었습니다.", true);
    cancelPmEdit();
    loadPortMappings();
  } else {
    const err = await res.json().catch(() => ({}));
    showMsg("pm-msg", err.detail || "오류가 발생했습니다.", false);
  }
});

async function deletePm(id) {
  if (!confirm("이 포트 매핑을 삭제하시겠습니까?")) return;
  const res = await fetch(`/api/trkv/port-mappings/${id}`, { method: "DELETE" });
  if (res.ok) { showMsg("pm-msg", "삭제되었습니다.", true); loadPortMappings(); }
  else showMsg("pm-msg", "삭제 실패", false);
}

// ─── 포트 매핑 엑셀 다운로드 / 업로드 ─────────────────────────────
function downloadPmTemplate() {
  window.location.href = "/api/trkv/port-mappings/template";
}

async function uploadPm() {
  const fileInput = document.getElementById("pm-file");
  const file = fileInput.files[0];
  if (!file) return;
  const mode = document.querySelector('input[name="pm-mode"]:checked').value;

  const fd = new FormData();
  fd.append("file", file);
  fd.append("mode", mode);

  const msgEl = document.getElementById("pm-upload-msg");
  msgEl.textContent = "업로드 중...";
  msgEl.className = "upload-result";
  msgEl.style.display = "inline";

  try {
    const res = await fetch("/api/trkv/port-mappings/upload", { method: "POST", body: fd });
    const data = await res.json();
    if (res.ok) {
      const failMsg = data.failed.length ? ` (실패 ${data.failed.length}건)` : "";
      msgEl.textContent = `✅ ${data.success}건 등록${failMsg}`;
      msgEl.className = "upload-result ok";
      loadPortMappings();
    } else {
      msgEl.textContent = `❌ ${data.detail || "업로드 실패"}`;
      msgEl.className = "upload-result err";
    }
  } catch {
    msgEl.textContent = "❌ 네트워크 오류";
    msgEl.className = "upload-result err";
  }
  fileInput.value = "";
  setTimeout(() => { msgEl.style.display = "none"; }, 5000);
}

// ──────────────────────────────────────────────────────────────────
// ② 컨테이너 티어 설정
// ──────────────────────────────────────────────────────────────────
let ctData = [];

async function loadContainerTiers() {
  const res = await fetch("/api/trkv/container-tiers");
  ctData = await res.json();
  renderContainerTiers();
}

function renderContainerTiers() {
  const tbody = document.getElementById("ct-tbody");
  const rows = [];
  for (const ct of CONT_TYPES) {
    for (const isDg of DG_OPTIONS) {
      const saved = ctData.find(d => d.cont_type === ct && d.is_dg === isDg);
      const tierVal = saved ? (saved.tier_number ?? "") : "";
      const dgLabel = isDg ? '<span class="badge badge-red">X</span>' : '<span class="badge badge-gray">없음</span>';
      const options = `<option value="">-</option>` +
        [1,2,3,4,5,6].map(n => `<option value="${n}" ${tierVal == n ? "selected" : ""}>${n}</option>`).join("");
      rows.push(`
        <tr>
          <td><strong>${ct}</strong></td>
          <td>${dgLabel}</td>
          <td>
            <select class="ct-select" data-cont="${ct}" data-dg="${isDg}">
              ${options}
            </select>
          </td>
        </tr>
      `);
    }
  }
  tbody.innerHTML = rows.join("");
}

async function saveContainerTiers() {
  const selects = document.querySelectorAll(".ct-select");
  const items = [];
  selects.forEach(sel => {
    items.push({
      cont_type: sel.dataset.cont,
      is_dg: sel.dataset.dg === "true",
      tier_number: sel.value ? parseInt(sel.value) : null,
    });
  });

  const res = await fetch("/api/trkv/container-tiers/bulk", {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify({ items }),
  });

  if (res.ok) {
    ctData = await res.json();
    showMsg("ct-msg", "저장되었습니다.", true);
  } else {
    showMsg("ct-msg", "저장 실패", false);
  }
}

// ──────────────────────────────────────────────────────────────────
// ③ 구간별 요율
// ──────────────────────────────────────────────────────────────────
let rtData = [];
let rtEditMode = false;  // 인라인 편집 모드 여부

async function loadRoutes() {
  const res = await fetch("/api/trkv/routes");
  rtData = await res.json();
  renderRoutes();
}

// ─── 일반 뷰 렌더링 ───────────────────────────────────────────────
function renderRoutes() {
  const tbody = document.getElementById("rt-tbody");
  if (!rtData.length) {
    tbody.innerHTML = '<tr><td colspan="13" class="empty-msg">등록된 구간 요율이 없습니다.</td></tr>';
    updateRtSelCount();
    return;
  }

  if (rtEditMode) {
    renderRoutesEditMode(tbody);
  } else {
    renderRoutesViewMode(tbody);
  }
  updateRtSelCount();
}

function renderRoutesViewMode(tbody) {
  tbody.innerHTML = rtData.map((r, i) => `
    <tr>
      <td><input type="checkbox" class="rt-chk" data-id="${r.id}" onchange="updateRtSelCount()" /></td>
      <td>${i + 1}</td>
      <td>${portBadge(r.pickup_port)}</td>
      <td>${r.departure_name}</td>
      <td>${portBadge(r.dest_port)}</td>
      <td>${fmt(r.tier1)}</td>
      <td>${fmt(r.tier2)}</td>
      <td>${fmt(r.tier3)}</td>
      <td>${fmt(r.tier4)}</td>
      <td>${fmt(r.tier5)}</td>
      <td>${fmt(r.tier6)}</td>
      <td>${r.memo || "-"}</td>
      <td>
        <button class="btn btn-sm btn-outline" onclick="startRtEdit(${r.id})">수정</button>
        <button class="btn btn-sm btn-danger" onclick="deleteRt(${r.id})">삭제</button>
      </td>
    </tr>
  `).join("");
}

// ─── 편집 모드 렌더링 ─────────────────────────────────────────────
function renderRoutesEditMode(tbody) {
  const portOpts = `<option value="부산신항">부산신항</option><option value="부산북항">부산북항</option>`;
  tbody.innerHTML = rtData.map((r, i) => `
    <tr data-id="${r.id}">
      <td><input type="checkbox" class="rt-chk" data-id="${r.id}" onchange="updateRtSelCount()" /></td>
      <td>${i + 1}</td>
      <td>
        <select class="rt-inline-pickup">
          ${`<option value="부산신항" ${r.pickup_port==="부산신항"?"selected":""}>부산신항</option>
             <option value="부산북항" ${r.pickup_port==="부산북항"?"selected":""}>부산북항</option>`}
        </select>
      </td>
      <td><input type="text" class="rt-inline-dep" value="${r.departure_name}" /></td>
      <td>
        <select class="rt-inline-dest">
          ${`<option value="부산신항" ${r.dest_port==="부산신항"?"selected":""}>부산신항</option>
             <option value="부산북항" ${r.dest_port==="부산북항"?"selected":""}>부산북항</option>`}
        </select>
      </td>
      ${[1,2,3,4,5,6].map(n => `<td><input type="number" class="rt-inline-tier" data-tier="${n}" value="${r["tier"+n] ?? ""}" min="0" /></td>`).join("")}
      <td><input type="text" class="rt-inline-memo" value="${r.memo || ""}" /></td>
      <td><button class="btn btn-sm btn-danger" onclick="deleteRt(${r.id})">삭제</button></td>
    </tr>
  `).join("");
}

// ─── 편집 모드 토글 ───────────────────────────────────────────────
function toggleRtEditMode() {
  rtEditMode = true;
  document.getElementById("rt-edit-mode-btn").style.display = "none";
  document.getElementById("rt-save-bulk-btn").style.display = "inline-flex";
  document.getElementById("rt-cancel-edit-btn").style.display = "inline-flex";
  renderRoutes();
}

function cancelRtEditMode() {
  rtEditMode = false;
  document.getElementById("rt-edit-mode-btn").style.display = "inline-flex";
  document.getElementById("rt-save-bulk-btn").style.display = "none";
  document.getElementById("rt-cancel-edit-btn").style.display = "none";
  renderRoutes();
}

async function saveRtBulk() {
  const rows = document.querySelectorAll("#rt-tbody tr[data-id]");
  if (!rows.length) { cancelRtEditMode(); return; }

  const updates = [];
  rows.forEach(tr => {
    const id = parseInt(tr.dataset.id);
    const body = {
      pickup_port:    tr.querySelector(".rt-inline-pickup").value,
      departure_name: tr.querySelector(".rt-inline-dep").value.trim(),
      dest_port:      tr.querySelector(".rt-inline-dest").value,
      memo: tr.querySelector(".rt-inline-memo").value.trim() || null,
    };
    tr.querySelectorAll(".rt-inline-tier").forEach(inp => {
      const n = inp.dataset.tier;
      const v = inp.value;
      body[`tier${n}`] = (v === "" || v === null) ? null : parseFloat(v);
    });
    updates.push({ id, body });
  });

  const results = await Promise.allSettled(
    updates.map(({ id, body }) =>
      fetch(`/api/trkv/routes/${id}`, {
        method: "PUT",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify(body),
      })
    )
  );

  const failed = results.filter(r => r.status === "rejected" || (r.value && !r.value.ok)).length;
  if (failed) {
    showMsg("rt-msg", `일부 저장 실패 (${failed}건)`, false);
  } else {
    showMsg("rt-msg", `${updates.length}건 저장되었습니다.`, true);
  }

  rtEditMode = false;
  document.getElementById("rt-edit-mode-btn").style.display = "inline-flex";
  document.getElementById("rt-save-bulk-btn").style.display = "none";
  document.getElementById("rt-cancel-edit-btn").style.display = "none";
  loadRoutes();
}

// ─── 체크박스 / 선택 삭제 ─────────────────────────────────────────
function updateRtSelCount() {
  const checked = document.querySelectorAll(".rt-chk:checked").length;
  document.getElementById("rt-sel-count").textContent = checked;
  document.getElementById("rt-action-bar").style.display = checked > 0 ? "flex" : "none";
  const all = document.querySelectorAll(".rt-chk");
  document.getElementById("rt-check-all").checked = all.length > 0 && checked === all.length;
}

function toggleAllRtCheck(checked) {
  document.querySelectorAll(".rt-chk").forEach(cb => { cb.checked = checked; });
  updateRtSelCount();
}

async function deleteSelectedRt() {
  const ids = [...document.querySelectorAll(".rt-chk:checked")].map(cb => parseInt(cb.dataset.id));
  if (!ids.length) return;
  if (!confirm(`선택한 ${ids.length}건을 삭제하시겠습니까?`)) return;
  for (const id of ids) {
    await fetch(`/api/trkv/routes/${id}`, { method: "DELETE" });
  }
  if (rtEditMode) { rtEditMode = false; }
  showMsg("rt-msg", `${ids.length}건 삭제되었습니다.`, true);
  loadRoutes();
}

// ─── 개별 등록/수정 폼 ────────────────────────────────────────────
function startRtEdit(id) {
  const r = rtData.find(x => x.id === id);
  if (!r) return;
  // 편집 모드 종료 후 폼 채우기
  if (rtEditMode) cancelRtEditMode();
  document.getElementById("rt-edit-id").value = id;
  document.getElementById("rt-pickup").value   = r.pickup_port;
  document.getElementById("rt-departure").value = r.departure_name;
  document.getElementById("rt-dest").value     = r.dest_port;
  [1,2,3,4,5,6].forEach(n => {
    document.getElementById(`rt-tier${n}`).value = r[`tier${n}`] ?? "";
  });
  document.getElementById("rt-memo").value = r.memo || "";
  document.getElementById("rt-submit-btn").textContent = "수정";
  document.getElementById("rt-cancel-btn").style.display = "inline-flex";
  document.getElementById("rt-form").scrollIntoView({ behavior: "smooth" });
}

function cancelRtEdit() {
  document.getElementById("rt-edit-id").value = "";
  document.getElementById("rt-pickup").value  = "";
  document.getElementById("rt-departure").value = "";
  document.getElementById("rt-dest").value    = "";
  [1,2,3,4,5,6].forEach(n => document.getElementById(`rt-tier${n}`).value = "");
  document.getElementById("rt-memo").value = "";
  document.getElementById("rt-submit-btn").textContent = "등록";
  document.getElementById("rt-cancel-btn").style.display = "none";
}

function nullOrFloat(val) {
  if (val === "" || val === null || val === undefined) return null;
  const n = parseFloat(val);
  return isNaN(n) ? null : n;
}

document.getElementById("rt-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const editId = document.getElementById("rt-edit-id").value;
  const body = {
    pickup_port:    document.getElementById("rt-pickup").value,
    departure_name: document.getElementById("rt-departure").value.trim(),
    dest_port:      document.getElementById("rt-dest").value,
    tier1: nullOrFloat(document.getElementById("rt-tier1").value),
    tier2: nullOrFloat(document.getElementById("rt-tier2").value),
    tier3: nullOrFloat(document.getElementById("rt-tier3").value),
    tier4: nullOrFloat(document.getElementById("rt-tier4").value),
    tier5: nullOrFloat(document.getElementById("rt-tier5").value),
    tier6: nullOrFloat(document.getElementById("rt-tier6").value),
    memo: document.getElementById("rt-memo").value.trim() || null,
  };
  if (!body.pickup_port || !body.departure_name || !body.dest_port) {
    showMsg("rt-msg", "픽업항, 출하지명, 도착항은 필수 입력입니다.", false);
    return;
  }

  const url    = editId ? `/api/trkv/routes/${editId}` : "/api/trkv/routes";
  const method = editId ? "PUT" : "POST";
  const res    = await fetch(url, { method, headers: {"Content-Type": "application/json"}, body: JSON.stringify(body) });

  if (res.ok) {
    showMsg("rt-msg", editId ? "수정되었습니다." : "등록되었습니다.", true);
    cancelRtEdit();
    loadRoutes();
  } else {
    showMsg("rt-msg", "오류가 발생했습니다.", false);
  }
});

async function deleteRt(id) {
  if (!confirm("이 구간 요율을 삭제하시겠습니까?")) return;
  const res = await fetch(`/api/trkv/routes/${id}`, { method: "DELETE" });
  if (res.ok) { showMsg("rt-msg", "삭제되었습니다.", true); loadRoutes(); }
  else showMsg("rt-msg", "삭제 실패", false);
}

// ─── 구간요율 엑셀 다운로드 / 업로드 ─────────────────────────────
function downloadRtTemplate() {
  window.location.href = "/api/trkv/routes/template";
}

async function uploadRt() {
  const fileInput = document.getElementById("rt-file");
  const file = fileInput.files[0];
  if (!file) return;
  const mode = document.querySelector('input[name="rt-mode"]:checked').value;

  const fd = new FormData();
  fd.append("file", file);
  fd.append("mode", mode);

  const msgEl = document.getElementById("rt-upload-msg");
  msgEl.textContent = "업로드 중...";
  msgEl.className = "upload-result";
  msgEl.style.display = "inline";

  try {
    const res = await fetch("/api/trkv/routes/upload", { method: "POST", body: fd });
    const data = await res.json();
    if (res.ok) {
      const failMsg = data.failed.length ? ` (실패 ${data.failed.length}건)` : "";
      msgEl.textContent = `✅ ${data.success}건 등록${failMsg}`;
      msgEl.className = "upload-result ok";
      if (rtEditMode) cancelRtEditMode();
      loadRoutes();
    } else {
      msgEl.textContent = `❌ ${data.detail || "업로드 실패"}`;
      msgEl.className = "upload-result err";
    }
  } catch {
    msgEl.textContent = "❌ 네트워크 오류";
    msgEl.className = "upload-result err";
  }
  fileInput.value = "";
  setTimeout(() => { msgEl.style.display = "none"; }, 5000);
}

// ──────────────────────────────────────────────────────────────────
// 페이지 초기화
// ──────────────────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
  loadPortMappings();
  loadContainerTiers();
  loadRoutes();
});
