/* ─── 매핑설정 페이지 스크립트 ─── */

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
  const cls = port === "부산신항" ? "badge-blue" : "badge-purple";
  return `<span class="badge ${cls}">${port}</span>`;
}

// ──────────────────────────────────────────────────────────────────
// ① 포트명 매핑
// ──────────────────────────────────────────────────────────────────
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
        if (!prev) {
          changedPmIds.add(r.id);
        } else {
          if (String(r.excel_name ?? "") !== String(prev.excel_name ?? "") ||
              String(r.port_type   ?? "") !== String(prev.port_type   ?? "")) {
            changedPmIds.add(r.id);
          }
        }
      }
    }

    renderPm();
  } catch (e) {
    const tbody = document.getElementById("pm-tbody");
    if (tbody) tbody.innerHTML = '<tr><td colspan="5" class="empty-msg" style="color:#ef4444">불러오기 실패</td></tr>';
  }
}

function renderPm() {
  const tbody = document.getElementById("pm-tbody");
  if (!tbody) return;
  if (!pmData.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="empty-msg">등록된 포트 매핑이 없습니다.</td></tr>';
    updatePmSelCount();
    return;
  }
  tbody.innerHTML = pmData.map((d, i) => {
    const changed = changedPmIds.has(d.id) ? " row-changed" : "";
    return `
    <tr class="${changed}">
      <td><input type="checkbox" class="pm-chk" data-id="${d.id}" onchange="updatePmSelCount()" /></td>
      <td>${i + 1}</td>
      <td>${d.excel_name}</td>
      <td>${portBadge(d.port_type)}</td>
      <td>
        <button class="btn btn-sm btn-outline" onclick="startPmEdit(${d.id})">수정</button>
        <button class="btn btn-sm btn-danger" onclick="deletePm(${d.id})">삭제</button>
      </td>
    </tr>
  `;
  }).join("");
  updatePmSelCount();
}

function updatePmSelCount() {
  const checked = document.querySelectorAll(".pm-chk:checked").length;
  const countEl = document.getElementById("pm-sel-count");
  const barEl = document.getElementById("pm-action-bar");
  if (countEl) countEl.textContent = checked;
  if (barEl) barEl.style.display = checked > 0 ? "flex" : "none";
  const all = document.querySelectorAll(".pm-chk");
  const allChk = document.getElementById("pm-check-all");
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

// ──────────────────────────────────────────────────────────────────
// 통합 업로드 (전체 교체)
// ──────────────────────────────────────────────────────────────────
function downloadUnified() {
  window.location.href = "/api/trkv/template";
}

async function uploadUnified() {
  const fileInput = document.getElementById("unified-file");
  const file = fileInput.files[0];
  if (!file) return;

  // 변경 감지를 위해 현재 데이터 스냅샷 저장
  const prevPmData = pmData.map(r => ({ ...r }));

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
      if (data.sheets["포트명 매핑"]) {
        const s = data.sheets["포트명 매핑"];
        parts.push(`포트매핑 ${s.success}건`);
      }
      if (data.sheets["TRKV 구간 요율"]) {
        const s = data.sheets["TRKV 구간 요율"];
        parts.push(`구간요율 ${s.success}건`);
      }
      msgEl.textContent = `✅ ${parts.join(" · ")} 등록`;
      msgEl.className = "upload-result ok";
      await loadPortMappings(prevPmData);
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
});
