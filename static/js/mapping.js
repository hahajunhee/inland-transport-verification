/* ─── 매핑설정 페이지 스크립트 v5 ─── */

function showMsg(elId, msg, isOk) {
  const el = document.getElementById(elId);
  if (!el) return;
  el.textContent = msg;
  el.className = "form-msg " + (isOk ? "msg-ok" : "msg-error");
  el.style.display = "inline-block";
  setTimeout(() => { el.style.display = "none"; }, 3500);
}

// ══════════════════════════════════════════════════════════════════
// ① 포트명 매핑
// ══════════════════════════════════════════════════════════════════
let pmData = [];
let changedPmIds = new Set();

async function loadPortMappings(prevSnapshot) {
  try {
    const res = await fetch("/api/trkv/port-mappings");
    if (!res.ok) throw new Error("API 오류");
    pmData = await res.json();
    if (prevSnapshot) {
      changedPmIds = new Set();
      const prevMap = new Map(prevSnapshot.map(r => [r.id, r]));
      for (const r of pmData) {
        const prev = prevMap.get(r.id);
        if (!prev || String(r.excel_name     ?? "") !== String(prev.excel_name     ?? "")
                  || String(r.port_type     ?? "") !== String(prev.port_type     ?? "")
                  || String(r.terminal_type ?? "") !== String(prev.terminal_type ?? "")) {
          changedPmIds.add(r.id);
        }
      }
    }
    renderPm();
  } catch {
    const tbody = document.getElementById("pm-tbody");
    if (tbody) tbody.innerHTML = '<tr><td colspan="5" class="empty-msg" style="color:#ef4444">불러오기 실패</td></tr>';
  }
}

function renderPm() {
  const tbody = document.getElementById("pm-tbody");
  if (!tbody) return;
  if (!pmData.length) {
    tbody.innerHTML = '<tr><td colspan="6" class="empty-msg">등록된 포트 매핑이 없습니다.</td></tr>';
    updatePmSelCount(); return;
  }
  tbody.innerHTML = pmData.map((d, i) => `
    <tr class="${changedPmIds.has(d.id) ? "row-changed" : ""}">
      <td><input type="checkbox" class="pm-chk" data-id="${d.id}" onchange="updatePmSelCount()" /></td>
      <td>${i + 1}</td>
      <td>${d.excel_name}</td>
      <td>${d.port_type}</td>
      <td>${d.terminal_type ? `<span class="badge badge-green">${d.terminal_type}</span>` : '<span style="color:#9ca3af">-</span>'}</td>
      <td>
        <button class="btn btn-sm btn-outline" onclick="startPmEdit(${d.id})">수정</button>
        <button class="btn btn-sm btn-danger"  onclick="deletePm(${d.id})">삭제</button>
      </td>
    </tr>`).join("");
  updatePmSelCount();
}

function updatePmSelCount() {
  const checked = document.querySelectorAll(".pm-chk:checked").length;
  const countEl = document.getElementById("pm-sel-count");
  const barEl   = document.getElementById("pm-action-bar");
  if (countEl) countEl.textContent = checked;
  if (barEl)   barEl.style.display = checked > 0 ? "flex" : "none";
  const allChk = document.getElementById("pm-check-all");
  const all    = document.querySelectorAll(".pm-chk");
  if (allChk) allChk.checked = all.length > 0 && checked === all.length;
}

function toggleAllPmCheck(checked) {
  document.querySelectorAll(".pm-chk").forEach(cb => { cb.checked = checked; });
  updatePmSelCount();
}

async function deleteSelectedPm() {
  const ids = [...document.querySelectorAll(".pm-chk:checked")].map(cb => parseInt(cb.dataset.id));
  if (!ids.length) return;
  if (!confirm(`선택한 ${ids.length}건을 삭제하시겠습니까?`)) return;
  for (const id of ids) await fetch(`/api/trkv/port-mappings/${id}`, { method: "DELETE" });
  showMsg("pm-msg", `${ids.length}건 삭제되었습니다.`, true);
  loadPortMappings();
}

function startPmEdit(id) {
  const d = pmData.find(x => x.id === id);
  if (!d) return;
  document.getElementById("pm-edit-id").value       = id;
  document.getElementById("pm-excel-name").value    = d.excel_name;
  document.getElementById("pm-port-type").value     = d.port_type;
  document.getElementById("pm-terminal-type").value = d.terminal_type || "";
  document.getElementById("pm-submit-btn").textContent = "수정";
  document.getElementById("pm-cancel-btn").style.display = "inline-flex";
}

function cancelPmEdit() {
  document.getElementById("pm-edit-id").value       = "";
  document.getElementById("pm-excel-name").value    = "";
  document.getElementById("pm-port-type").value     = "";
  document.getElementById("pm-terminal-type").value = "";
  document.getElementById("pm-submit-btn").textContent = "추가";
  document.getElementById("pm-cancel-btn").style.display = "none";
}

document.getElementById("pm-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const editId = document.getElementById("pm-edit-id").value;
  const body   = {
    excel_name:    document.getElementById("pm-excel-name").value.trim(),
    port_type:     document.getElementById("pm-port-type").value.trim(),
    terminal_type: document.getElementById("pm-terminal-type").value.trim(),
  };
  if (!body.excel_name || !body.port_type) return;
  const url    = editId ? `/api/trkv/port-mappings/${editId}` : "/api/trkv/port-mappings";
  const method = editId ? "PUT" : "POST";
  const res    = await fetch(url, { method, headers: {"Content-Type": "application/json"}, body: JSON.stringify(body) });
  if (res.ok) {
    showMsg("pm-msg", editId ? "수정되었습니다." : "추가되었습니다.", true);
    cancelPmEdit(); loadPortMappings();
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

// ══════════════════════════════════════════════════════════════════
// ② 출하지 매핑
// ══════════════════════════════════════════════════════════════════
let dmData = [];
let changedDmIds = new Set();

async function loadDepartureMappings(prevSnapshot) {
  try {
    const res = await fetch("/api/trkv/departure-mappings");
    if (!res.ok) throw new Error("API 오류");
    dmData = await res.json();
    if (prevSnapshot) {
      changedDmIds = new Set();
      const prevMap = new Map(prevSnapshot.map(r => [r.id, r]));
      for (const r of dmData) {
        const prev = prevMap.get(r.id);
        if (!prev || String(r.departure_name ?? "") !== String(prev.departure_name ?? "")
                  || String(r.departure_code ?? "") !== String(prev.departure_code ?? "")) {
          changedDmIds.add(r.id);
        }
      }
    }
    renderDm();
  } catch {
    const tbody = document.getElementById("dm-tbody");
    if (tbody) tbody.innerHTML = '<tr><td colspan="5" class="empty-msg" style="color:#ef4444">불러오기 실패</td></tr>';
  }
}

function renderDm() {
  const tbody = document.getElementById("dm-tbody");
  if (!tbody) return;
  if (!dmData.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="empty-msg">등록된 출하지 매핑이 없습니다.</td></tr>';
    updateDmSelCount(); return;
  }
  tbody.innerHTML = dmData.map((d, i) => `
    <tr class="${changedDmIds.has(d.id) ? "row-changed" : ""}">
      <td><input type="checkbox" class="dm-chk" data-id="${d.id}" onchange="updateDmSelCount()" /></td>
      <td>${i + 1}</td>
      <td>${d.departure_name}</td>
      <td><span class="badge badge-blue">${d.departure_code}</span></td>
      <td>
        <button class="btn btn-sm btn-outline" onclick="startDmEdit(${d.id})">수정</button>
        <button class="btn btn-sm btn-danger"  onclick="deleteDm(${d.id})">삭제</button>
      </td>
    </tr>`).join("");
  updateDmSelCount();
}

function updateDmSelCount() {
  const checked = document.querySelectorAll(".dm-chk:checked").length;
  const countEl = document.getElementById("dm-sel-count");
  const barEl   = document.getElementById("dm-action-bar");
  if (countEl) countEl.textContent = checked;
  if (barEl)   barEl.style.display = checked > 0 ? "flex" : "none";
  const allChk = document.getElementById("dm-check-all");
  const all    = document.querySelectorAll(".dm-chk");
  if (allChk) allChk.checked = all.length > 0 && checked === all.length;
}

function toggleAllDmCheck(checked) {
  document.querySelectorAll(".dm-chk").forEach(cb => { cb.checked = checked; });
  updateDmSelCount();
}

async function deleteSelectedDm() {
  const ids = [...document.querySelectorAll(".dm-chk:checked")].map(cb => parseInt(cb.dataset.id));
  if (!ids.length) return;
  if (!confirm(`선택한 ${ids.length}건을 삭제하시겠습니까?`)) return;
  for (const id of ids) await fetch(`/api/trkv/departure-mappings/${id}`, { method: "DELETE" });
  showMsg("dm-msg", `${ids.length}건 삭제되었습니다.`, true);
  loadDepartureMappings();
}

function startDmEdit(id) {
  const d = dmData.find(x => x.id === id);
  if (!d) return;
  document.getElementById("dm-edit-id").value       = id;
  document.getElementById("dm-departure-name").value = d.departure_name;
  document.getElementById("dm-departure-code").value = d.departure_code;
  document.getElementById("dm-submit-btn").textContent = "수정";
  document.getElementById("dm-cancel-btn").style.display = "inline-flex";
}

function cancelDmEdit() {
  document.getElementById("dm-edit-id").value        = "";
  document.getElementById("dm-departure-name").value = "";
  document.getElementById("dm-departure-code").value = "";
  document.getElementById("dm-submit-btn").textContent = "추가";
  document.getElementById("dm-cancel-btn").style.display = "none";
}

document.getElementById("dm-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const editId = document.getElementById("dm-edit-id").value;
  const body   = {
    departure_name: document.getElementById("dm-departure-name").value.trim(),
    departure_code: document.getElementById("dm-departure-code").value.trim(),
  };
  if (!body.departure_name || !body.departure_code) return;
  const url    = editId ? `/api/trkv/departure-mappings/${editId}` : "/api/trkv/departure-mappings";
  const method = editId ? "PUT" : "POST";
  const res    = await fetch(url, { method, headers: {"Content-Type": "application/json"}, body: JSON.stringify(body) });
  if (res.ok) {
    showMsg("dm-msg", editId ? "수정되었습니다." : "추가되었습니다.", true);
    cancelDmEdit(); loadDepartureMappings();
  } else {
    const err = await res.json().catch(() => ({}));
    showMsg("dm-msg", err.detail || "오류가 발생했습니다.", false);
  }
});

async function deleteDm(id) {
  if (!confirm("이 출하지 매핑을 삭제하시겠습니까?")) return;
  const res = await fetch(`/api/trkv/departure-mappings/${id}`, { method: "DELETE" });
  if (res.ok) { showMsg("dm-msg", "삭제되었습니다.", true); loadDepartureMappings(); }
  else showMsg("dm-msg", "삭제 실패", false);
}

// ══════════════════════════════════════════════════════════════════
// 통합 업로드 (전체 교체)
// ══════════════════════════════════════════════════════════════════
function downloadUnified() {
  window.location.href = "/api/trkv/template";
}

async function uploadUnified() {
  const fileInput = document.getElementById("unified-file");
  const file = fileInput.files[0];
  if (!file) return;

  const prevPmData = pmData.map(r => ({ ...r }));
  const prevDmData = dmData.map(r => ({ ...r }));
  const prevOmData = omData.map(r => ({ ...r }));

  const fd = new FormData();
  fd.append("file", file);

  const msgEl = document.getElementById("unified-msg");
  msgEl.textContent = "업로드 중...";
  msgEl.className   = "upload-result";
  msgEl.style.display = "inline";

  try {
    const res  = await fetch("/api/trkv/upload", { method: "POST", body: fd });
    const data = await res.json();
    if (res.ok) {
      const parts = [];
      if (data.sheets["포트명 매핑"])    parts.push(`포트매핑 ${data.sheets["포트명 매핑"].success}건`);
      if (data.sheets["출하지 매핑"])    parts.push(`출하지매핑 ${data.sheets["출하지 매핑"].success}건`);
      if (data.sheets["ODCY 매핑"])      parts.push(`ODCY매핑 ${data.sheets["ODCY 매핑"].success}건`);
      if (data.sheets["TRKV 구간 요율"]) parts.push(`구간요율 ${data.sheets["TRKV 구간 요율"].success}건`);
      msgEl.textContent = `✅ ${parts.join(" · ")} 교체 완료`;
      msgEl.className   = "upload-result ok";
      await loadPortMappings(prevPmData);
      await loadDepartureMappings(prevDmData);
      await loadOdcyMappings(prevOmData);
    } else {
      msgEl.textContent = `❌ ${data.detail || "업로드 실패"}`;
      msgEl.className   = "upload-result err";
    }
  } catch {
    msgEl.textContent = "❌ 네트워크 오류";
    msgEl.className   = "upload-result err";
  }
  fileInput.value = "";
  setTimeout(() => { msgEl.style.display = "none"; }, 6000);
}

// ══════════════════════════════════════════════════════════════════
// 페이지 초기화
// ══════════════════════════════════════════════════════════════════
document.addEventListener("DOMContentLoaded", () => {
  loadPortMappings();
  loadDepartureMappings();
  loadOdcyMappings();
});


// ══════════════════════════════════════════════════════════════════
// ③ ODCY 매핑
// ══════════════════════════════════════════════════════════════════
let omData = [];
let changedOmIds = new Set();

async function loadOdcyMappings(prevSnapshot) {
  try {
    const res = await fetch("/api/trkv/odcy-mappings");
    if (!res.ok) throw new Error("API 오류");
    omData = await res.json();
    if (prevSnapshot) {
      changedOmIds = new Set();
      const prevMap = new Map(prevSnapshot.map(r => [r.id, r]));
      for (const r of omData) {
        const prev = prevMap.get(r.id);
        if (!prev || String(r.odcy_destination_name ?? "") !== String(prev.odcy_destination_name ?? "")
                  || String(r.odcy_name        ?? "") !== String(prev.odcy_name        ?? "")
                  || String(r.terminal_type    ?? "") !== String(prev.terminal_type    ?? "")) {
          changedOmIds.add(r.id);
        }
      }
    }
    renderOm();
  } catch {
    const tbody = document.getElementById("om-tbody");
    if (tbody) tbody.innerHTML = '<tr><td colspan="5" class="empty-msg" style="color:#ef4444">불러오기 실패</td></tr>';
  }
}

function renderOm() {
  const tbody = document.getElementById("om-tbody");
  if (!tbody) return;
  if (!omData.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="empty-msg">등록된 ODCY 매핑이 없습니다.</td></tr>';
    updateOmSelCount(); return;
  }
  tbody.innerHTML = omData.map((d, i) => `
    <tr class="${changedOmIds.has(d.id) ? "row-changed" : ""}">
      <td><input type="checkbox" class="om-chk" data-id="${d.id}" onchange="updateOmSelCount()" /></td>
      <td>${i + 1}</td>
      <td>${d.odcy_destination_name}</td>
      <td><span class="badge badge-blue">${d.odcy_name}</span></td>
      <td>${d.terminal_type ? `<span class="badge badge-green">${d.terminal_type}</span>` : '<span style="color:#9ca3af">-</span>'}</td>
      <td>
        <button class="btn btn-sm btn-outline" onclick="startOmEdit(${d.id})">수정</button>
        <button class="btn btn-sm btn-danger"  onclick="deleteOm(${d.id})">삭제</button>
      </td>
    </tr>`).join("");
  updateOmSelCount();
}

function updateOmSelCount() {
  const checked = document.querySelectorAll(".om-chk:checked").length;
  const countEl = document.getElementById("om-sel-count");
  const barEl   = document.getElementById("om-action-bar");
  if (countEl) countEl.textContent = checked;
  if (barEl)   barEl.style.display = checked > 0 ? "flex" : "none";
  const allChk = document.getElementById("om-check-all");
  const all    = document.querySelectorAll(".om-chk");
  if (allChk) allChk.checked = all.length > 0 && checked === all.length;
}

function toggleAllOmCheck(checked) {
  document.querySelectorAll(".om-chk").forEach(cb => { cb.checked = checked; });
  updateOmSelCount();
}

async function deleteSelectedOm() {
  const ids = [...document.querySelectorAll(".om-chk:checked")].map(cb => parseInt(cb.dataset.id));
  if (!ids.length) return;
  if (!confirm(`선택한 ${ids.length}건을 삭제하시겠습니까?`)) return;
  for (const id of ids) await fetch(`/api/trkv/odcy-mappings/${id}`, { method: "DELETE" });
  showMsg("om-msg", `${ids.length}건 삭제되었습니다.`, true);
  loadOdcyMappings();
}

function startOmEdit(id) {
  const d = omData.find(x => x.id === id);
  if (!d) return;
  document.getElementById("om-edit-id").value        = id;
  document.getElementById("om-dest-name").value      = d.odcy_destination_name;
  document.getElementById("om-odcy-name").value      = d.odcy_name;
  document.getElementById("om-terminal-type").value  = d.terminal_type || "";
  document.getElementById("om-submit-btn").textContent = "수정";
  document.getElementById("om-cancel-btn").style.display = "inline-flex";
}

function cancelOmEdit() {
  document.getElementById("om-edit-id").value       = "";
  document.getElementById("om-dest-name").value     = "";
  document.getElementById("om-odcy-name").value     = "";
  document.getElementById("om-terminal-type").value = "";
  document.getElementById("om-submit-btn").textContent = "추가";
  document.getElementById("om-cancel-btn").style.display = "none";
}

document.getElementById("om-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const editId = document.getElementById("om-edit-id").value;
  const body   = {
    odcy_destination_name: document.getElementById("om-dest-name").value.trim(),
    odcy_name:             document.getElementById("om-odcy-name").value.trim(),
    terminal_type:         document.getElementById("om-terminal-type").value.trim(),
  };
  if (!body.odcy_destination_name || !body.odcy_name) return;
  const url    = editId ? `/api/trkv/odcy-mappings/${editId}` : "/api/trkv/odcy-mappings";
  const method = editId ? "PUT" : "POST";
  const res    = await fetch(url, { method, headers: {"Content-Type": "application/json"}, body: JSON.stringify(body) });
  if (res.ok) {
    showMsg("om-msg", editId ? "수정되었습니다." : "추가되었습니다.", true);
    cancelOmEdit(); loadOdcyMappings();
  } else {
    const err = await res.json().catch(() => ({}));
    showMsg("om-msg", err.detail || "오류가 발생했습니다.", false);
  }
});

async function deleteOm(id) {
  if (!confirm("이 ODCY 매핑을 삭제하시겠습니까?")) return;
  const res = await fetch(`/api/trkv/odcy-mappings/${id}`, { method: "DELETE" });
  if (res.ok) { showMsg("om-msg", "삭제되었습니다.", true); loadOdcyMappings(); }
  else showMsg("om-msg", "삭제 실패", false);
}
