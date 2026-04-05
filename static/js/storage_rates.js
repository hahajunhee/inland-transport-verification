/* ─── 보관료/상하차료/셔틀비 요율 v5 ─── */

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

// ── 컨테이너 티어 (기존과 동일) ──
let sctData = [];

async function loadStorageContainerTiers() {
  try {
    const res = await fetch("/api/trkv/storage-container-tiers");
    if (!res.ok) throw new Error();
    sctData = await res.json();
    renderStorageContainerTiers();
  } catch { document.getElementById("sct-tbody").innerHTML = '<tr><td colspan="3" class="empty-msg" style="color:#ef4444">실패</td></tr>'; }
}

function renderStorageContainerTiers() {
  const tbody = document.getElementById("sct-tbody");
  const rows = [];
  for (const ct of CONT_TYPES) {
    for (const isDg of DG_OPTIONS) {
      const saved = sctData.find(d => d.cont_type === ct && d.is_dg === isDg);
      const tierVal = saved ? (saved.tier_number ?? "") : "";
      const dgLabel = isDg ? '<span class="badge badge-red">X</span>' : '<span class="badge badge-gray">없음</span>';
      const options = '<option value="">-</option>' + [1,2,3,4,5,6].map(n => `<option value="${n}" ${String(tierVal)===String(n)?"selected":""}>${n}</option>`).join("");
      rows.push(`<tr><td><strong>${ct}</strong></td><td>${dgLabel}</td><td><select class="sct-select" data-cont="${ct}" data-dg="${isDg}">${options}</select></td></tr>`);
    }
  }
  tbody.innerHTML = rows.join("");
}

async function saveStorageContainerTiers() {
  const items = [...document.querySelectorAll(".sct-select")].map(sel => ({
    cont_type: sel.dataset.cont, is_dg: sel.dataset.dg === "true",
    tier_number: sel.value ? parseInt(sel.value) : null,
  }));
  const res = await fetch("/api/trkv/storage-container-tiers/bulk", {
    method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({items})
  });
  if (res.ok) { sctData = await res.json(); showMsg("sct-msg","저장",true); }
  else showMsg("sct-msg","실패",false);
}

// ── 요율 CRUD ──
let srData = [];
let srBulkEditing = false;

async function loadStorageRates() {
  try {
    const res = await fetch("/api/storage-rates/");
    if (!res.ok) throw new Error();
    srData = await res.json();
    srBulkEditing = false;
    document.getElementById("sr-bulk-btn").textContent = "일괄 편집";
    renderSr();
  } catch {
    document.getElementById("sr-tbody").innerHTML = '<tr><td colspan="27" class="empty-msg" style="color:#ef4444">실패</td></tr>';
  }
}

function tierInputs(prefix, d, id) {
  return TIERS.map(t => {
    const v = d ? d[`${prefix}_tier${t}`] : null;
    return `<td><input class="inline-num" id="sr-${prefix}-t${t}-${id}" type="number" min="0" step="1" value="${v!=null?v:''}" /></td>`;
  }).join("");
}

function tierCells(prefix, d) {
  return TIERS.map(t => `<td class="money">${fmtMoney(d[`${prefix}_tier${t}`])}</td>`).join("");
}

function renderSrRowHtml(d, i, editing) {
  const esc = (v) => (v||'').replace(/"/g,'&quot;');
  const badge = (v) => v ? `<span class="badge badge-green">${v}</span>` : '<span style="color:#9ca3af">-</span>';

  if (editing) {
    return `<tr id="sr-row-${d.id}" class="editing-row">
      <td><input type="checkbox" class="sr-chk" data-id="${d.id}" onchange="updateSrSelCount()" /></td>
      <td>${i+1}</td>
      <td><input class="inline-input" id="sr-on-${d.id}" value="${esc(d.odcy_name)}" /></td>
      <td><input class="inline-input" id="sr-ot-${d.id}" value="${esc(d.odcy_terminal_type||d.terminal_type)}" /></td>
      <td><input class="inline-input" id="sr-ol-${d.id}" value="${esc(d.odcy_location)}" /></td>
      <td><input class="inline-input" id="sr-dp-${d.id}" value="${esc(d.dest_port_type)}" /></td>
      <td><input class="inline-input" id="sr-dt-${d.id}" value="${esc(d.dest_terminal_type)}" /></td>
      ${tierInputs("storage", d, d.id)}
      ${tierInputs("handling", d, d.id)}
      ${tierInputs("shuttle", d, d.id)}
      <td><input class="inline-input" id="sr-memo-${d.id}" value="${esc(d.memo)}" style="width:80px" /></td>
      <td style="white-space:nowrap">
        <button class="btn btn-sm btn-primary" onclick="saveSrRow(${d.id})">저장</button>
        <button class="btn btn-sm btn-outline" onclick="cancelSrRowEdit(${d.id})">취소</button>
      </td>
    </tr>`;
  }
  const autoStyle = d.auto_generated ? ' style="background:#fffde7"' : '';
  return `<tr id="sr-row-${d.id}"${autoStyle}>
    <td><input type="checkbox" class="sr-chk" data-id="${d.id}" onchange="updateSrSelCount()" /></td>
    <td>${i+1}</td>
    <td>${d.odcy_name||'<span style="color:#9ca3af">-</span>'}</td>
    <td>${badge(d.odcy_terminal_type||d.terminal_type)}</td>
    <td>${badge(d.odcy_location)}</td>
    <td>${badge(d.dest_port_type)}</td>
    <td>${badge(d.dest_terminal_type)}</td>
    ${tierCells("storage",d)} ${tierCells("handling",d)} ${tierCells("shuttle",d)}
    <td style="font-size:12px;color:#6b7280">${d.memo||""}</td>
    <td style="white-space:nowrap">
      <button class="btn btn-sm btn-outline" onclick="editSrRow(${d.id})">수정</button>
      <button class="btn btn-sm btn-danger" onclick="deleteSr(${d.id})">삭제</button>
    </td>
  </tr>`;
}

function renderSr() {
  const tbody = document.getElementById("sr-tbody");
  if (!srData.length) { tbody.innerHTML = '<tr><td colspan="27" class="empty-msg">등록된 요율이 없습니다.</td></tr>'; updateSrSelCount(); return; }
  tbody.innerHTML = srData.map((d,i) => renderSrRowHtml(d,i,srBulkEditing)).join("");
  updateSrSelCount();
}

function editSrRow(id) {
  const d = srData.find(x=>x.id===id);
  if (!d) return;
  const tr = document.getElementById(`sr-row-${id}`);
  if (tr) tr.outerHTML = renderSrRowHtml(d, srData.indexOf(d), true);
}

function cancelSrRowEdit(id) {
  const d = srData.find(x=>x.id===id);
  if (!d) return;
  const tr = document.getElementById(`sr-row-${id}`);
  if (tr) tr.outerHTML = renderSrRowHtml(d, srData.indexOf(d), false);
}

function buildSrBody(id) {
  const gv = (elId) => (document.getElementById(elId)?.value||"").trim();
  const body = {
    odcy_name: gv(`sr-on-${id}`),
    odcy_terminal_type: gv(`sr-ot-${id}`),
    odcy_location: gv(`sr-ol-${id}`),
    dest_port_type: gv(`sr-dp-${id}`),
    dest_terminal_type: gv(`sr-dt-${id}`),
    memo: gv(`sr-memo-${id}`),
  };
  for (const t of TIERS) {
    for (const p of ["storage","handling","shuttle"]) {
      const v = document.getElementById(`sr-${p}-t${t}-${id}`)?.value;
      body[`${p}_tier${t}`] = v!=="" ? parseFloat(v) : null;
    }
  }
  return body;
}

async function saveSrRow(id) {
  const body = buildSrBody(id);
  const res = await fetch(`/api/storage-rates/${id}`, {
    method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(body)
  });
  if (res.ok) { showMsg("sr-msg","저장 완료",true); loadStorageRates(); }
  else { const e=await res.json().catch(()=>({})); showMsg("sr-msg",e.detail||"실패",false); }
}

function toggleSrBulkEdit() {
  if (!srBulkEditing) {
    srBulkEditing = true;
    document.getElementById("sr-bulk-btn").textContent = "일괄 저장";
    renderSr();
  } else {
    saveAllSrRows();
  }
}

async function saveAllSrRows() {
  let ok=0, fail=0;
  for (const d of srData) {
    const name = document.getElementById(`sr-on-${d.id}`)?.value?.trim();
    if (name === undefined) continue;
    const body = buildSrBody(d.id);
    const res = await fetch(`/api/storage-rates/${d.id}`, {
      method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(body)
    });
    if (res.ok) ok++; else fail++;
  }
  showMsg("sr-msg",`일괄 저장: ${ok}건 성공${fail?`, ${fail}건 실패`:""}`,fail===0);
  loadStorageRates();
}

function showSrAddRow() {
  const tbody = document.getElementById("sr-tbody");
  if (document.getElementById("sr-new-row")) return;
  const tr = document.createElement("tr");
  tr.id = "sr-new-row";
  tr.className = "new-row";
  tr.innerHTML = `
    <td></td><td>신규</td>
    <td><input class="inline-input" id="sr-on-new" placeholder="ODCY명" /></td>
    <td><input class="inline-input" id="sr-ot-new" placeholder="터미널구분" /></td>
    <td><input class="inline-input" id="sr-ol-new" placeholder="ODCY_위치" /></td>
    <td><input class="inline-input" id="sr-dp-new" placeholder="포트구분" /></td>
    <td><input class="inline-input" id="sr-dt-new" placeholder="터미널구분" /></td>
    ${tierInputs("storage", null, "new")}
    ${tierInputs("handling", null, "new")}
    ${tierInputs("shuttle", null, "new")}
    <td><input class="inline-input" id="sr-memo-new" placeholder="비고" style="width:80px" /></td>
    <td style="white-space:nowrap">
      <button class="btn btn-sm btn-primary" onclick="addSrRow()">추가</button>
      <button class="btn btn-sm btn-outline" onclick="document.getElementById('sr-new-row').remove()">취소</button>
    </td>`;
  tbody.prepend(tr);
  document.getElementById("sr-on-new").focus();
}

async function addSrRow() {
  const body = buildSrBody("new");
  if (!body.odcy_name) { alert("ODCY명은 필수입니다."); return; }
  const res = await fetch("/api/storage-rates/", {
    method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(body)
  });
  if (res.ok) { showMsg("sr-msg","추가 완료",true); loadStorageRates(); }
  else { const e=await res.json().catch(()=>({})); showMsg("sr-msg",e.detail||"실패",false); }
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
  const ids = [...document.querySelectorAll(".sr-chk:checked")].map(cb=>parseInt(cb.dataset.id));
  if (!ids.length) return;
  if (!confirm(`선택한 ${ids.length}건을 삭제하시겠습니까?`)) return;
  for (const id of ids) await fetch(`/api/storage-rates/${id}`,{method:"DELETE"});
  showMsg("sr-msg",`${ids.length}건 삭제`,true);
  loadStorageRates();
}

async function deleteSr(id) {
  if (!confirm("삭제하시겠습니까?")) return;
  const res = await fetch(`/api/storage-rates/${id}`,{method:"DELETE"});
  if (res.ok) { showMsg("sr-msg","삭제 완료",true); loadStorageRates(); }
  else showMsg("sr-msg","삭제 실패",false);
}

function downloadTemplate() { window.location.href = "/api/trkv/template"; }

async function uploadRates() {
  const fileInput = document.getElementById("sr-upload-file");
  const file = fileInput.files[0];
  if (!file) return;
  const fd = new FormData();
  fd.append("file", file);
  const msgEl = document.getElementById("sr-upload-msg");
  msgEl.textContent = "업로드 중...";
  msgEl.className = "upload-result";
  msgEl.style.display = "inline";
  try {
    const res = await fetch("/api/trkv/upload", {method:"POST",body:fd});
    const data = await res.json();
    if (res.ok) {
      const sheets = data.sheets||{};
      const srKey = Object.keys(sheets).find(k=>k.includes("보관료")||k.includes("셔틀"));
      if (srKey) {
        const s = sheets[srKey];
        msgEl.textContent = `완료 — ${s.success}건 교체`;
      } else {
        msgEl.textContent = `완료 — 보관료 시트 없음`;
      }
      msgEl.className = "upload-result ok";
      loadStorageRates();
    } else {
      msgEl.textContent = `오류: ${data.detail||"실패"}`;
      msgEl.className = "upload-result err";
    }
  } catch {
    msgEl.textContent = "네트워크 오류";
    msgEl.className = "upload-result err";
  }
  fileInput.value = "";
  setTimeout(()=>{msgEl.style.display="none";},6000);
}

document.addEventListener("DOMContentLoaded", () => {
  loadStorageContainerTiers();
  loadStorageRates();
});
