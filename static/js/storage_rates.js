/* ─── 보관료/상하차료/셔틀비 요율 페이지 스크립트 v4 ─── */

const TIERS = [1, 2, 3, 4, 5, 6];
const CONT_TYPES = ["22G1", "22R1", "45G1", "45R1"];
const DG_OPTIONS = [false, true];

function showMsg(elId, msg, isOk) {
  const el = document.getElementById(elId);
  if (!el) return;
  el.textContent = msg;
  el.className = "form-msg " + (isOk ? "msg-ok" : "msg-error");
  el.style.display = "inline-block";
  setTimeout(() => { el.style.display = "none"; }, 3500);
}

function fmtMoney(v) {
  if (v == null || v === "") return '<span style="color:#d1d5db">-</span>';
  return Number(v).toLocaleString("ko-KR");
}

// ──────────────────────────────────────────────────────────────────
// 보관료/상하차료/셔틀비 컨테이너 티어
// ──────────────────────────────────────────────────────────────────
let sctData = [];

async function loadStorageContainerTiers() {
  try {
    const res = await fetch("/api/trkv/storage-container-tiers");
    if (!res.ok) throw new Error("API 오류");
    sctData = await res.json();
    renderStorageContainerTiers();
  } catch (e) {
    const tbody = document.getElementById("sct-tbody");
    if (tbody) tbody.innerHTML = '<tr><td colspan="3" class="empty-msg" style="color:#ef4444">불러오기 실패</td></tr>';
  }
}

function renderStorageContainerTiers() {
  const tbody = document.getElementById("sct-tbody");
  if (!tbody) return;
  const rows = [];
  for (const ct of CONT_TYPES) {
    for (const isDg of DG_OPTIONS) {
      const saved = sctData.find(d => d.cont_type === ct && d.is_dg === isDg);
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
            <select class="sct-select" data-cont="${ct}" data-dg="${isDg}">
              ${options}
            </select>
          </td>
        </tr>`);
    }
  }
  tbody.innerHTML = rows.join("");
}

async function saveStorageContainerTiers() {
  const selects = document.querySelectorAll(".sct-select");
  const items = [];
  selects.forEach(sel => {
    items.push({
      cont_type: sel.dataset.cont,
      is_dg: sel.dataset.dg === "true",
      tier_number: sel.value ? parseInt(sel.value) : null,
    });
  });
  const res = await fetch("/api/trkv/storage-container-tiers/bulk", {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify({ items }),
  });
  if (res.ok) {
    sctData = await res.json();
    showMsg("sct-msg", "저장되었습니다.", true);
  } else {
    showMsg("sct-msg", "저장 실패", false);
  }
}

// ──────────────────────────────────────────────────────────────────
// 보관료/상하차료/셔틀비 요율 CRUD
// ──────────────────────────────────────────────────────────────────
let srData = [];

async function loadStorageRates() {
  try {
    const res = await fetch("/api/storage-rates/");
    if (!res.ok) throw new Error("API 오류");
    srData = await res.json();
    renderSr();
  } catch {
    const tbody = document.getElementById("sr-tbody");
    if (tbody) tbody.innerHTML = '<tr><td colspan="27" class="empty-msg" style="color:#ef4444">불러오기 실패</td></tr>';
  }
}

function renderSr() {
  const tbody = document.getElementById("sr-tbody");
  if (!tbody) return;
  if (!srData.length) {
    tbody.innerHTML = '<tr><td colspan="27" class="empty-msg">등록된 요율이 없습니다.</td></tr>';
    updateSrSelCount(); return;
  }
  tbody.innerHTML = srData.map((d, i) => {
    const storageCells  = TIERS.map(t => `<td class="money">${fmtMoney(d["storage_tier" + t])}</td>`).join("");
    const handlingCells = TIERS.map(t => `<td class="money">${fmtMoney(d["handling_tier" + t])}</td>`).join("");
    const shuttleCells  = TIERS.map(t => `<td class="money">${fmtMoney(d["shuttle_tier" + t])}</td>`).join("");
    const badge = (v) => v ? `<span class="badge badge-green">${v}</span>` : '<span style="color:#9ca3af">-</span>';
    return `<tr>
      <td><input type="checkbox" class="sr-chk" data-id="${d.id}" onchange="updateSrSelCount()" /></td>
      <td>${i + 1}</td>
      <td>${d.odcy_name || '<span style="color:#9ca3af">-</span>'}</td>
      <td>${badge(d.odcy_terminal_type || d.terminal_type)}</td>
      <td>${badge(d.odcy_location)}</td>
      <td>${badge(d.dest_port_type)}</td>
      <td>${badge(d.dest_terminal_type)}</td>
      ${storageCells}
      ${handlingCells}
      ${shuttleCells}
      <td style="font-size:12px;color:#6b7280">${d.memo || ""}</td>
      <td>
        <button class="btn btn-sm btn-outline" onclick="startSrEdit(${d.id})">수정</button>
        <button class="btn btn-sm btn-danger"  onclick="deleteSr(${d.id})">삭제</button>
      </td>
    </tr>`;
  }).join("");
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
  document.getElementById("sr-edit-id").value              = id;
  document.getElementById("sr-odcy-name").value            = d.odcy_name || "";
  document.getElementById("sr-odcy-terminal-type").value   = d.odcy_terminal_type || d.terminal_type || "";
  document.getElementById("sr-odcy-location").value        = d.odcy_location || "";
  document.getElementById("sr-dest-port-type").value       = d.dest_port_type || "";
  document.getElementById("sr-dest-terminal-type").value   = d.dest_terminal_type || "";
  document.getElementById("sr-memo").value                 = d.memo || "";
  for (const t of TIERS) {
    const sv = d["storage_tier" + t];
    const hv = d["handling_tier" + t];
    const shv = d["shuttle_tier" + t];
    document.getElementById("sr-storage-tier" + t).value  = sv != null ? sv : "";
    document.getElementById("sr-handling-tier" + t).value = hv != null ? hv : "";
    document.getElementById("sr-shuttle-tier" + t).value  = shv != null ? shv : "";
  }
  document.getElementById("sr-submit-btn").textContent = "수정";
  document.getElementById("sr-cancel-btn").style.display = "inline-flex";
  document.getElementById("sr-form").scrollIntoView({ behavior: "smooth" });
}

function cancelSrEdit() {
  document.getElementById("sr-edit-id").value            = "";
  document.getElementById("sr-odcy-name").value          = "";
  document.getElementById("sr-odcy-terminal-type").value = "";
  document.getElementById("sr-odcy-location").value      = "";
  document.getElementById("sr-dest-port-type").value     = "";
  document.getElementById("sr-dest-terminal-type").value = "";
  document.getElementById("sr-memo").value               = "";
  for (const t of TIERS) {
    document.getElementById("sr-storage-tier" + t).value  = "";
    document.getElementById("sr-handling-tier" + t).value = "";
    document.getElementById("sr-shuttle-tier" + t).value  = "";
  }
  document.getElementById("sr-submit-btn").textContent = "추가";
  document.getElementById("sr-cancel-btn").style.display = "none";
}

function buildBody() {
  const body = {
    odcy_name:          document.getElementById("sr-odcy-name").value.trim(),
    odcy_terminal_type: document.getElementById("sr-odcy-terminal-type").value.trim(),
    odcy_location:      document.getElementById("sr-odcy-location").value.trim(),
    dest_port_type:     document.getElementById("sr-dest-port-type").value.trim(),
    dest_terminal_type: document.getElementById("sr-dest-terminal-type").value.trim(),
    memo:               document.getElementById("sr-memo").value.trim(),
  };
  for (const t of TIERS) {
    const sv  = document.getElementById("sr-storage-tier" + t).value;
    const hv  = document.getElementById("sr-handling-tier" + t).value;
    const shv = document.getElementById("sr-shuttle-tier" + t).value;
    body["storage_tier" + t]  = sv  !== "" ? parseFloat(sv)  : null;
    body["handling_tier" + t] = hv  !== "" ? parseFloat(hv)  : null;
    body["shuttle_tier" + t]  = shv !== "" ? parseFloat(shv) : null;
  }
  return body;
}

document.getElementById("sr-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const editId = document.getElementById("sr-edit-id").value;
  const body   = buildBody();
  if (!body.odcy_name || !body.odcy_terminal_type) {
    showMsg("sr-msg", "ODCY명과 odcy터미널구분은 필수입니다.", false); return;
  }
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
  window.location.href = "/api/trkv/template";
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
    const res  = await fetch("/api/trkv/upload", { method: "POST", body: fd });
    const data = await res.json();
    if (res.ok) {
      const sheets = data.sheets || {};
      const srKey  = Object.keys(sheets).find(k => k.includes("보관료") || k.includes("셔틀"));
      if (srKey) {
        const s = sheets[srKey];
        const fails = s.failed?.length || 0;
        msgEl.textContent = `완료 — ${srKey} ${s.success}건 교체` + (fails ? ` (오류 ${fails}건)` : "");
      } else {
        const total = Object.values(sheets).reduce((a, s) => a + (s.success || 0), 0);
        msgEl.textContent = `완료 — 전체 ${total}건 교체 (보관료 시트 없음)`;
      }
      msgEl.className = "upload-result ok";
      loadStorageRates();
    } else {
      msgEl.textContent = `오류: ${data.detail || "업로드 실패"}`;
      msgEl.className   = "upload-result err";
    }
  } catch {
    msgEl.textContent = "네트워크 오류";
    msgEl.className   = "upload-result err";
  }
  fileInput.value = "";
  setTimeout(() => { msgEl.style.display = "none"; }, 6000);
}

document.addEventListener("DOMContentLoaded", () => {
  loadStorageContainerTiers();
  loadStorageRates();
});
