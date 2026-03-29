/* ─── 보관료/상하차료 요율 페이지 스크립트 v1 ─── */

function showMsg(elId, msg, isOk) {
  const el = document.getElementById(elId);
  if (!el) return;
  el.textContent = msg;
  el.className = "form-msg " + (isOk ? "msg-ok" : "msg-error");
  el.style.display = "inline-block";
  setTimeout(() => { el.style.display = "none"; }, 3500);
}

function fmtMoney(v) {
  if (v == null) return "-";
  return Number(v).toLocaleString("ko-KR");
}

let srData = [];

async function loadStorageRates() {
  try {
    const res = await fetch("/api/storage-rates/");
    if (!res.ok) throw new Error("API 오류");
    srData = await res.json();
    renderSr();
  } catch {
    const tbody = document.getElementById("sr-tbody");
    if (tbody) tbody.innerHTML = '<tr><td colspan="8" class="empty-msg" style="color:#ef4444">불러오기 실패</td></tr>';
  }
}

function renderSr() {
  const tbody = document.getElementById("sr-tbody");
  if (!tbody) return;
  if (!srData.length) {
    tbody.innerHTML = '<tr><td colspan="8" class="empty-msg">등록된 요율이 없습니다.</td></tr>';
    updateSrSelCount(); return;
  }
  tbody.innerHTML = srData.map((d, i) => `
    <tr>
      <td><input type="checkbox" class="sr-chk" data-id="${d.id}" onchange="updateSrSelCount()" /></td>
      <td>${i + 1}</td>
      <td>${d.odcy_name || '<span style="color:#9ca3af">-</span>'}</td>
      <td>${d.zone_type ? `<span class="badge badge-green">${d.zone_type}</span>` : '<span style="color:#9ca3af">-</span>'}</td>
      <td class="money">${fmtMoney(d.storage_unit)}</td>
      <td class="money">${fmtMoney(d.handling_unit)}</td>
      <td style="font-size:12px;color:#6b7280">${d.memo || ""}</td>
      <td>
        <button class="btn btn-sm btn-outline" onclick="startSrEdit(${d.id})">수정</button>
        <button class="btn btn-sm btn-danger"  onclick="deleteSr(${d.id})">삭제</button>
      </td>
    </tr>`).join("");
  updateSrSelCount();
}

function updateSrSelCount() {
  const checked = document.querySelectorAll(".sr-chk:checked").length;
  const countEl = document.getElementById("sr-sel-count");
  const barEl   = document.getElementById("sr-action-bar");
  if (countEl) countEl.textContent = checked;
  if (barEl)   barEl.style.display = checked > 0 ? "flex" : "none";
  const allChk = document.getElementById("sr-check-all");
  const all    = document.querySelectorAll(".sr-chk");
  if (allChk) allChk.checked = all.length > 0 && checked === all.length;
}

function toggleAllSrCheck(checked) {
  document.querySelectorAll(".sr-chk").forEach(cb => { cb.checked = checked; });
  updateSrSelCount();
}

async function deleteSelectedSr() {
  const ids = [...document.querySelectorAll(".sr-chk:checked")].map(cb => parseInt(cb.dataset.id));
  if (!ids.length) return;
  if (!confirm(`선택한 ${ids.length}건을 삭제하시겠습니까?`)) return;
  for (const id of ids) await fetch(`/api/storage-rates/${id}`, { method: "DELETE" });
  showMsg("sr-msg", `${ids.length}건 삭제되었습니다.`, true);
  loadStorageRates();
}

function startSrEdit(id) {
  const d = srData.find(x => x.id === id);
  if (!d) return;
  document.getElementById("sr-edit-id").value       = id;
  document.getElementById("sr-odcy-name").value     = d.odcy_name || "";
  document.getElementById("sr-zone-type").value     = d.zone_type || "";
  document.getElementById("sr-storage-unit").value  = d.storage_unit ?? "";
  document.getElementById("sr-handling-unit").value = d.handling_unit ?? "";
  document.getElementById("sr-memo").value          = d.memo || "";
  document.getElementById("sr-submit-btn").textContent = "수정";
  document.getElementById("sr-cancel-btn").style.display = "inline-flex";
}

function cancelSrEdit() {
  document.getElementById("sr-edit-id").value       = "";
  document.getElementById("sr-odcy-name").value     = "";
  document.getElementById("sr-zone-type").value     = "";
  document.getElementById("sr-storage-unit").value  = "";
  document.getElementById("sr-handling-unit").value = "";
  document.getElementById("sr-memo").value          = "";
  document.getElementById("sr-submit-btn").textContent = "추가";
  document.getElementById("sr-cancel-btn").style.display = "none";
}

document.getElementById("sr-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const editId = document.getElementById("sr-edit-id").value;
  const su = document.getElementById("sr-storage-unit").value;
  const hu = document.getElementById("sr-handling-unit").value;
  const body = {
    odcy_name:     document.getElementById("sr-odcy-name").value.trim(),
    zone_type:     document.getElementById("sr-zone-type").value.trim(),
    storage_unit:  su !== "" ? parseFloat(su) : null,
    handling_unit: hu !== "" ? parseFloat(hu) : null,
    memo:          document.getElementById("sr-memo").value.trim(),
  };
  if (!body.odcy_name || !body.zone_type) return;
  const url    = editId ? `/api/storage-rates/${editId}` : "/api/storage-rates/";
  const method = editId ? "PUT" : "POST";
  const res    = await fetch(url, { method, headers: {"Content-Type": "application/json"}, body: JSON.stringify(body) });
  if (res.ok) {
    showMsg("sr-msg", editId ? "수정되었습니다." : "추가되었습니다.", true);
    cancelSrEdit(); loadStorageRates();
  } else {
    const err = await res.json().catch(() => ({}));
    showMsg("sr-msg", err.detail || "오류가 발생했습니다.", false);
  }
});

async function deleteSr(id) {
  if (!confirm("이 요율을 삭제하시겠습니까?")) return;
  const res = await fetch(`/api/storage-rates/${id}`, { method: "DELETE" });
  if (res.ok) { showMsg("sr-msg", "삭제되었습니다.", true); loadStorageRates(); }
  else showMsg("sr-msg", "삭제 실패", false);
}

function downloadTemplate() {
  window.location.href = "/api/storage-rates/template";
}

async function uploadRates() {
  const fileInput = document.getElementById("sr-upload-file");
  const file = fileInput.files[0];
  if (!file) return;

  const fd = new FormData();
  fd.append("file", file);

  const msgEl = document.getElementById("sr-upload-msg");
  msgEl.textContent = "업로드 중...";
  msgEl.className   = "upload-result";
  msgEl.style.display = "inline";

  try {
    const res  = await fetch("/api/storage-rates/upload", { method: "POST", body: fd });
    const data = await res.json();
    if (res.ok) {
      msgEl.textContent = `✅ ${data.success}건 교체 완료`;
      msgEl.className   = "upload-result ok";
      loadStorageRates();
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

document.addEventListener("DOMContentLoaded", () => {
  loadStorageRates();
});
