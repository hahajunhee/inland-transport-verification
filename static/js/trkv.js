/* ─── TRKV 요율 설정 페이지 스크립트 v5 ─── */

const CONT_TYPES = ["22G1", "22R1", "45G1", "45R1"];
const DG_OPTIONS = [false, true];

// ──────────────────────────────────────────────────────────────────
// 공통 유틸
// ──────────────────────────────────────────────────────────────────
function fmt(val) {
  if (val === null || val === undefined || val === "") return "-";
  return Number(val).toLocaleString("ko-KR");
}

function showMsg(elId, msg, isOk) {
  const el = document.getElementById(elId);
  if (!el) return;
  el.textContent = msg;
  el.className = "form-msg " + (isOk ? "msg-ok" : "msg-error");
  el.style.display = "inline-block";
  setTimeout(() => { el.style.display = "none"; }, 3500);
}

function portBadge(port) {
  const cls = port === "부산신항" ? "badge-blue" : port === "부산북항" ? "badge-purple" : "badge-gray";
  return `<span class="badge ${cls}">${port || "-"}</span>`;
}

// ──────────────────────────────────────────────────────────────────
// 통합 업로드 (전체 교체, 변경 행 하이라이트)
// ──────────────────────────────────────────────────────────────────
let changedRtKeys = new Set(); // "pickup|dep|dest" 키로 변경된 행 추적

function downloadUnified() {
  window.location.href = "/api/trkv/template";
}

async function uploadUnified() {
  const fileInput = document.getElementById("unified-file");
  const file = fileInput.files[0];
  if (!file) return;

  // 업로드 전 현재 상태 스냅샷
  const prevRtData = rtData.map(r => ({ ...r }));

  const fd = new FormData();
  fd.append("file", file);

  const msgEl = document.getElementById("unified-msg");
  msgEl.textContent = "업로드 중...";
  msgEl.className = "upload-result";
  msgEl.style.display = "inline";

  try {
    const res = await fetch("/api/trkv/upload", { method: "POST", body: fd });
    const data = await res.json();
    if (res.ok) {
      const parts = [];
      if (data.sheets["포트명 매핑"]) parts.push(`포트매핑 ${data.sheets["포트명 매핑"].success}건`);
      if (data.sheets["TRKV 구간 요율"]) parts.push(`구간요율 ${data.sheets["TRKV 구간 요율"].success}건`);
      msgEl.textContent = `✅ ${parts.join(" · ")} 교체 완료`;
      msgEl.className = "upload-result ok";

      // 변경 감지 후 reload
      await loadRoutes(prevRtData);
    } else {
      msgEl.textContent = `❌ ${data.detail || "업로드 실패"}`;
      msgEl.className = "upload-result err";
    }
  } catch {
    msgEl.textContent = "❌ 네트워크 오류";
    msgEl.className = "upload-result err";
  }
  fileInput.value = "";
  setTimeout(() => { msgEl.style.display = "none"; }, 6000);
}

// ──────────────────────────────────────────────────────────────────
// ② 컨테이너 티어 설정(TRKV티어)
// ──────────────────────────────────────────────────────────────────
let ctData = [];

async function loadContainerTiers() {
  try {
    const res = await fetch("/api/trkv/container-tiers");
    if (!res.ok) throw new Error("API 오류");
    ctData = await res.json();
    renderContainerTiers();
  } catch (e) {
    const tbody = document.getElementById("ct-tbody");
    if (tbody) tbody.innerHTML = '<tr><td colspan="3" class="empty-msg" style="color:#ef4444">불러오기 실패</td></tr>';
  }
}

function renderContainerTiers() {
  const tbody = document.getElementById("ct-tbody");
  if (!tbody) return;
  const rows = [];
  for (const ct of CONT_TYPES) {
    for (const isDg of DG_OPTIONS) {
      const saved = ctData.find(d => d.cont_type === ct && d.is_dg === isDg);
      const tierVal = saved ? (saved.tier_number ?? "") : "";
      const dgLabel = isDg
        ? '<span class="badge badge-red">X</span>'
        : '<span class="badge badge-gray">없음</span>';
      const options = `<option value="">-</option>` +
        [1,2,3,4,5,6].map(n =>
          `<option value="${n}" ${String(tierVal) === String(n) ? "selected" : ""}>${n}</option>`
        ).join("");
      rows.push(`
        <tr>
          <td><strong>${ct}</strong></td>
          <td>${dgLabel}</td>
          <td>
            <select class="ct-select" data-cont="${ct}" data-dg="${isDg}">
              ${options}
            </select>
          </td>
        </tr>`);
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
let rtEditMode = false;

// prevSnapshot: 업로드 전 스냅샷 (변경 감지용)
// changedRtIds: 변경/신규 행 ID Set (다음 업로드 전까지 유지)
let changedRtIds = new Set();

async function loadRoutes(prevSnapshot) {
  try {
    const res = await fetch("/api/trkv/routes");
    if (!res.ok) throw new Error("API 오류");
    rtData = await res.json();

    if (prevSnapshot) {
      // 변경된 행 감지
      changedRtIds = new Set();
      const prevMap = new Map(prevSnapshot.map(r => [r.id, r]));
      for (const r of rtData) {
        const prev = prevMap.get(r.id);
        if (!prev) {
          // 신규 행 (업로드 후 새 ID로 삽입된 경우 → 전체 교체이므로 모두 새 ID)
          changedRtIds.add(r.id);
        } else {
          // 값이 달라진 행
          const fields = ["pickup_port","departure_code","dest_port","tier1","tier2","tier3","tier4","tier5","tier6","memo"];
          if (fields.some(f => String(r[f] ?? "") !== String(prev[f] ?? ""))) {
            changedRtIds.add(r.id);
          }
        }
      }
    }

    renderRoutes();
  } catch {
    const tbody = document.getElementById("rt-tbody");
    if (tbody) tbody.innerHTML = '<tr><td colspan="13" class="empty-msg" style="color:#ef4444">불러오기 실패</td></tr>';
  }
}

function renderRoutes() {
  const tbody = document.getElementById("rt-tbody");
  if (!tbody) return;
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
  tbody.innerHTML = rtData.map((r, i) => {
    const changed = changedRtIds.has(r.id) ? " row-changed" : "";
    return `
    <tr class="${changed}">
      <td><input type="checkbox" class="rt-chk" data-id="${r.id}" onchange="updateRtSelCount()" /></td>
      <td>${i + 1}</td>
      <td>${portBadge(r.pickup_port)}</td>
      <td>${r.departure_code ?? r.departure_name ?? "-"}</td>
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
    </tr>`;
  }).join("");
}

function renderRoutesEditMode(tbody) {
  tbody.innerHTML = rtData.map((r, i) => {
    const changed = changedRtIds.has(r.id) ? " row-changed" : "";
    return `
    <tr data-id="${r.id}" class="${changed}">
      <td><input type="checkbox" class="rt-chk" data-id="${r.id}" onchange="updateRtSelCount()" /></td>
      <td>${i + 1}</td>
      <td><input type="text" class="rt-inline-pickup" list="port-list" value="${r.pickup_port || ""}" /></td>
      <td><input type="text" class="rt-inline-dep" list="dep-code-list" value="${r.departure_code ?? r.departure_name ?? ""}" /></td>
      <td><input type="text" class="rt-inline-dest" list="port-list" value="${r.dest_port || ""}" /></td>
      ${[1,2,3,4,5,6].map(n =>
        `<td><input type="number" class="rt-inline-tier" data-tier="${n}" value="${r["tier"+n] ?? ""}" min="0" /></td>`
      ).join("")}
      <td><input type="text" class="rt-inline-memo" value="${r.memo || ""}" /></td>
      <td><button class="btn btn-sm btn-danger" onclick="deleteRt(${r.id})">삭제</button></td>
    </tr>`;
  }).join("");
}

// ─── 편집 모드 ───────────────────────────────────────────────────
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
      pickup_port:    tr.querySelector(".rt-inline-pickup").value.trim(),
      departure_code: tr.querySelector(".rt-inline-dep").value.trim(),
      dest_port:      tr.querySelector(".rt-inline-dest").value.trim(),
      memo:           tr.querySelector(".rt-inline-memo").value.trim() || null,
    };
    tr.querySelectorAll(".rt-inline-tier").forEach(inp => {
      const v = inp.value;
      body[`tier${inp.dataset.tier}`] = v === "" ? null : parseFloat(v);
    });
    updates.push({ id, body });
  });
  const results = await Promise.allSettled(
    updates.map(({ id, body }) =>
      fetch(`/api/trkv/routes/${id}`, {
        method: "PUT", headers: {"Content-Type": "application/json"}, body: JSON.stringify(body),
      })
    )
  );
  const failed = results.filter(r => r.status === "rejected" || (r.value && !r.value.ok)).length;
  showMsg("rt-msg", failed ? `일부 저장 실패(${failed}건)` : `${updates.length}건 저장`, !failed);
  rtEditMode = false;
  document.getElementById("rt-edit-mode-btn").style.display = "inline-flex";
  document.getElementById("rt-save-bulk-btn").style.display = "none";
  document.getElementById("rt-cancel-edit-btn").style.display = "none";
  loadRoutes();
}

// ─── 체크박스 / 선택 삭제 ────────────────────────────────────────
function updateRtSelCount() {
  const checked = document.querySelectorAll(".rt-chk:checked").length;
  document.getElementById("rt-sel-count").textContent = checked;
  document.getElementById("rt-action-bar").style.display = checked > 0 ? "flex" : "none";
  const all = document.querySelectorAll(".rt-chk");
  const allChk = document.getElementById("rt-check-all");
  if (allChk) allChk.checked = all.length > 0 && checked === all.length;
}

function toggleAllRtCheck(checked) {
  document.querySelectorAll(".rt-chk").forEach(cb => { cb.checked = checked; });
  updateRtSelCount();
}

async function deleteSelectedRt() {
  const ids = [...document.querySelectorAll(".rt-chk:checked")].map(cb => parseInt(cb.dataset.id));
  if (!ids.length) return;
  if (!confirm(`선택한 ${ids.length}건을 삭제하시겠습니까?`)) return;
  for (const id of ids) await fetch(`/api/trkv/routes/${id}`, { method: "DELETE" });
  if (rtEditMode) rtEditMode = false;
  showMsg("rt-msg", `${ids.length}건 삭제되었습니다.`, true);
  loadRoutes();
}

// ─── 개별 등록/수정 ──────────────────────────────────────────────
function startRtEdit(id) {
  const r = rtData.find(x => x.id === id);
  if (!r) return;
  if (rtEditMode) cancelRtEditMode();
  document.getElementById("rt-edit-id").value = id;
  document.getElementById("rt-pickup").value = r.pickup_port;
  document.getElementById("rt-departure").value = r.departure_code ?? r.departure_name ?? "";
  document.getElementById("rt-dest").value = r.dest_port;
  [1,2,3,4,5,6].forEach(n => { document.getElementById(`rt-tier${n}`).value = r[`tier${n}`] ?? ""; });
  document.getElementById("rt-memo").value = r.memo || "";
  document.getElementById("rt-submit-btn").textContent = "수정";
  document.getElementById("rt-cancel-btn").style.display = "inline-flex";
  document.getElementById("rt-form").scrollIntoView({ behavior: "smooth" });
}

function cancelRtEdit() {
  document.getElementById("rt-edit-id").value = "";
  document.getElementById("rt-pickup").value = "";
  document.getElementById("rt-departure").value = "";
  document.getElementById("rt-dest").value = "";
  [1,2,3,4,5,6].forEach(n => { document.getElementById(`rt-tier${n}`).value = ""; });
  document.getElementById("rt-memo").value = "";
  document.getElementById("rt-submit-btn").textContent = "등록";
  document.getElementById("rt-cancel-btn").style.display = "none";
}

function nullOrFloat(v) {
  if (v === "" || v === null || v === undefined) return null;
  const n = parseFloat(v);
  return isNaN(n) ? null : n;
}

document.getElementById("rt-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const editId = document.getElementById("rt-edit-id").value;
  const body = {
    pickup_port:    document.getElementById("rt-pickup").value,
    departure_code: document.getElementById("rt-departure").value.trim(),
    dest_port:      document.getElementById("rt-dest").value,
    tier1: nullOrFloat(document.getElementById("rt-tier1").value),
    tier2: nullOrFloat(document.getElementById("rt-tier2").value),
    tier3: nullOrFloat(document.getElementById("rt-tier3").value),
    tier4: nullOrFloat(document.getElementById("rt-tier4").value),
    tier5: nullOrFloat(document.getElementById("rt-tier5").value),
    tier6: nullOrFloat(document.getElementById("rt-tier6").value),
    memo:  document.getElementById("rt-memo").value.trim() || null,
  };
  if (!body.pickup_port || !body.departure_code || !body.dest_port) {
    showMsg("rt-msg", "픽업항, 출하지코드, 도착항은 필수입니다.", false);
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

// ──────────────────────────────────────────────────────────────────
// 페이지 초기화
// ──────────────────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
  loadContainerTiers();
  loadRoutes();
  loadDepCodeList();
});

async function loadDepCodeList() {
  try {
    const res = await fetch("/api/trkv/departure-mappings");
    if (!res.ok) return;
    const items = await res.json();
    const dl = document.getElementById("dep-code-list");
    if (!dl) return;
    dl.innerHTML = items.map(d => `<option value="${d.departure_code}">${d.departure_name} → ${d.departure_code}</option>`).join("");
  } catch { /* 무시 */ }
}
