/* ─── 매핑설정 페이지 스크립트 v6 ─── */

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
let pmBulkEditing = false;
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
        if (!prev
          || String(r.excel_name     ?? "") !== String(prev.excel_name     ?? "")
          || String(r.port_type      ?? "") !== String(prev.port_type      ?? "")
          || String(r.terminal_type  ?? "") !== String(prev.terminal_type  ?? "")) {
          changedPmIds.add(r.id);
        }
      }
    }
    pmBulkEditing = false;
    document.getElementById("pm-bulk-btn").textContent = "일괄 편집";
    renderPm();
  } catch {
    document.getElementById("pm-tbody").innerHTML =
      '<tr><td colspan="6" class="empty-msg" style="color:#ef4444">불러오기 실패</td></tr>';
  }
}

function renderPmRowHtml(d, i, editing) {
  if (editing) {
    return `<tr id="pm-row-${d.id}" class="editing-row">
      <td><input type="checkbox" class="pm-chk" data-id="${d.id}" onchange="updatePmSelCount()" /></td>
      <td>${i + 1}</td>
      <td><input class="inline-input" id="pm-excel-${d.id}" value="${(d.excel_name||'').replace(/"/g,'&quot;')}" /></td>
      <td><input class="inline-input" id="pm-port-${d.id}"  value="${(d.port_type||'').replace(/"/g,'&quot;')}" /></td>
      <td><input class="inline-input" id="pm-term-${d.id}"  value="${(d.terminal_type||'').replace(/"/g,'&quot;')}" /></td>
      <td>
        <button class="btn btn-sm btn-primary" onclick="savePmRow(${d.id})">저장</button>
        <button class="btn btn-sm btn-outline" onclick="cancelPmRowEdit(${d.id})">취소</button>
      </td>
    </tr>`;
  }
  return `<tr id="pm-row-${d.id}" class="${changedPmIds.has(d.id) ? 'row-changed' : ''}">
    <td><input type="checkbox" class="pm-chk" data-id="${d.id}" onchange="updatePmSelCount()" /></td>
    <td>${i + 1}</td>
    <td>${d.excel_name || ''}</td>
    <td>${d.port_type || ''}</td>
    <td>${d.terminal_type ? `<span class="badge badge-green">${d.terminal_type}</span>` : '<span style="color:#9ca3af">-</span>'}</td>
    <td>
      <button class="btn btn-sm btn-outline" onclick="editPmRow(${d.id})">수정</button>
      <button class="btn btn-sm btn-danger"  onclick="deletePm(${d.id})">삭제</button>
    </td>
  </tr>`;
}

function renderPm() {
  const tbody = document.getElementById("pm-tbody");
  if (!pmData.length) {
    tbody.innerHTML = '<tr><td colspan="6" class="empty-msg">등록된 포트 매핑이 없습니다.</td></tr>';
    updatePmSelCount(); return;
  }
  tbody.innerHTML = pmData.map((d, i) => renderPmRowHtml(d, i, pmBulkEditing)).join("");
  updatePmSelCount();
}

function editPmRow(id) {
  const d = pmData.find(x => x.id === id);
  if (!d) return;
  const idx = pmData.indexOf(d);
  const tr = document.getElementById(`pm-row-${id}`);
  if (tr) tr.outerHTML = renderPmRowHtml(d, idx, true);
}

function cancelPmRowEdit(id) {
  const d = pmData.find(x => x.id === id);
  if (!d) return;
  const idx = pmData.indexOf(d);
  const tr = document.getElementById(`pm-row-${id}`);
  if (tr) tr.outerHTML = renderPmRowHtml(d, idx, false);
}

async function savePmRow(id) {
  const body = {
    excel_name:    (document.getElementById(`pm-excel-${id}`)?.value || "").trim(),
    port_type:     (document.getElementById(`pm-port-${id}`)?.value  || "").trim(),
    terminal_type: (document.getElementById(`pm-term-${id}`)?.value  || "").trim(),
  };
  if (!body.excel_name || !body.port_type) { alert("포트명(엑셀원본명)과 포트 구분은 필수입니다."); return; }
  const res = await fetch(`/api/trkv/port-mappings/${id}`, {
    method: "PUT", headers: {"Content-Type":"application/json"}, body: JSON.stringify(body)
  });
  if (res.ok) { changedPmIds.add(id); loadPortMappings(pmData.map(r => ({...r}))); }
  else { const e = await res.json().catch(()=>({})); showMsg("pm-msg", e.detail || "저장 실패", false); }
}

function togglePmBulkEdit() {
  if (!pmBulkEditing) {
    pmBulkEditing = true;
    document.getElementById("pm-bulk-btn").textContent = "일괄 저장";
    renderPm();
  } else {
    saveAllPmRows();
  }
}

async function saveAllPmRows() {
  const prev = pmData.map(r => ({...r}));
  let ok = 0, fail = 0;
  for (const d of pmData) {
    const excel = document.getElementById(`pm-excel-${d.id}`)?.value?.trim();
    if (excel === undefined) continue;
    const body = {
      excel_name:    excel,
      port_type:     (document.getElementById(`pm-port-${d.id}`)?.value || "").trim(),
      terminal_type: (document.getElementById(`pm-term-${d.id}`)?.value || "").trim(),
    };
    if (!body.excel_name || !body.port_type) { fail++; continue; }
    const res = await fetch(`/api/trkv/port-mappings/${d.id}`, {
      method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(body)
    });
    if (res.ok) ok++; else fail++;
  }
  showMsg("pm-msg", `일괄 저장: ${ok}건 성공${fail ? `, ${fail}건 실패` : ""}`, fail === 0);
  loadPortMappings(prev);
}

function showPmAddRow() {
  const tbody = document.getElementById("pm-tbody");
  if (document.getElementById("pm-new-row")) return;
  const tr = document.createElement("tr");
  tr.id = "pm-new-row";
  tr.className = "new-row";
  tr.innerHTML = `
    <td></td>
    <td>신규</td>
    <td><input class="inline-input" id="pm-new-excel" placeholder="예: 부산신항BPTS" /></td>
    <td><input class="inline-input" id="pm-new-port"  placeholder="예: 부산신항" /></td>
    <td><input class="inline-input" id="pm-new-term"  placeholder="예: BPTS (선택)" /></td>
    <td>
      <button class="btn btn-sm btn-primary" onclick="addPmRow()">추가</button>
      <button class="btn btn-sm btn-outline" onclick="document.getElementById('pm-new-row').remove()">취소</button>
    </td>`;
  tbody.prepend(tr);
  document.getElementById("pm-new-excel").focus();
}

async function addPmRow() {
  const body = {
    excel_name:    (document.getElementById("pm-new-excel")?.value || "").trim(),
    port_type:     (document.getElementById("pm-new-port")?.value  || "").trim(),
    terminal_type: (document.getElementById("pm-new-term")?.value  || "").trim(),
  };
  if (!body.excel_name || !body.port_type) { alert("포트명(엑셀원본명)과 포트 구분은 필수입니다."); return; }
  const res = await fetch("/api/trkv/port-mappings", {
    method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(body)
  });
  if (res.ok) { showMsg("pm-msg", "추가되었습니다.", true); loadPortMappings(); }
  else { const e = await res.json().catch(()=>({})); showMsg("pm-msg", e.detail || "추가 실패", false); }
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
  for (const id of ids) await fetch(`/api/trkv/port-mappings/${id}`, { method:"DELETE" });
  showMsg("pm-msg", `${ids.length}건 삭제되었습니다.`, true);
  loadPortMappings();
}

async function deletePm(id) {
  if (!confirm("이 포트 매핑을 삭제하시겠습니까?")) return;
  const res = await fetch(`/api/trkv/port-mappings/${id}`, { method:"DELETE" });
  if (res.ok) { showMsg("pm-msg", "삭제되었습니다.", true); loadPortMappings(); }
  else showMsg("pm-msg", "삭제 실패", false);
}

// ══════════════════════════════════════════════════════════════════
// ② 출하지 매핑
// ══════════════════════════════════════════════════════════════════
let dmData = [];
let dmBulkEditing = false;
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
        if (!prev
          || String(r.departure_name ?? "") !== String(prev.departure_name ?? "")
          || String(r.departure_code ?? "") !== String(prev.departure_code ?? "")) {
          changedDmIds.add(r.id);
        }
      }
    }
    dmBulkEditing = false;
    document.getElementById("dm-bulk-btn").textContent = "일괄 편집";
    renderDm();
  } catch {
    document.getElementById("dm-tbody").innerHTML =
      '<tr><td colspan="5" class="empty-msg" style="color:#ef4444">불러오기 실패</td></tr>';
  }
}

function renderDmRowHtml(d, i, editing) {
  if (editing) {
    return `<tr id="dm-row-${d.id}" class="editing-row">
      <td><input type="checkbox" class="dm-chk" data-id="${d.id}" onchange="updateDmSelCount()" /></td>
      <td>${i + 1}</td>
      <td><input class="inline-input" id="dm-name-${d.id}" value="${(d.departure_name||'').replace(/"/g,'&quot;')}" /></td>
      <td><input class="inline-input" id="dm-code-${d.id}" value="${(d.departure_code||'').replace(/"/g,'&quot;')}" /></td>
      <td>
        <button class="btn btn-sm btn-primary" onclick="saveDmRow(${d.id})">저장</button>
        <button class="btn btn-sm btn-outline" onclick="cancelDmRowEdit(${d.id})">취소</button>
      </td>
    </tr>`;
  }
  return `<tr id="dm-row-${d.id}" class="${changedDmIds.has(d.id) ? 'row-changed' : ''}">
    <td><input type="checkbox" class="dm-chk" data-id="${d.id}" onchange="updateDmSelCount()" /></td>
    <td>${i + 1}</td>
    <td>${d.departure_name || ''}</td>
    <td><span class="badge badge-blue">${d.departure_code || ''}</span></td>
    <td>
      <button class="btn btn-sm btn-outline" onclick="editDmRow(${d.id})">수정</button>
      <button class="btn btn-sm btn-danger"  onclick="deleteDm(${d.id})">삭제</button>
    </td>
  </tr>`;
}

function renderDm() {
  const tbody = document.getElementById("dm-tbody");
  if (!dmData.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="empty-msg">등록된 출하지 매핑이 없습니다.</td></tr>';
    updateDmSelCount(); return;
  }
  tbody.innerHTML = dmData.map((d, i) => renderDmRowHtml(d, i, dmBulkEditing)).join("");
  updateDmSelCount();
}

function editDmRow(id) {
  const d = dmData.find(x => x.id === id);
  if (!d) return;
  const tr = document.getElementById(`dm-row-${id}`);
  if (tr) tr.outerHTML = renderDmRowHtml(d, dmData.indexOf(d), true);
}

function cancelDmRowEdit(id) {
  const d = dmData.find(x => x.id === id);
  if (!d) return;
  const tr = document.getElementById(`dm-row-${id}`);
  if (tr) tr.outerHTML = renderDmRowHtml(d, dmData.indexOf(d), false);
}

async function saveDmRow(id) {
  const body = {
    departure_name: (document.getElementById(`dm-name-${id}`)?.value || "").trim(),
    departure_code: (document.getElementById(`dm-code-${id}`)?.value || "").trim(),
  };
  if (!body.departure_name || !body.departure_code) { alert("출하지명과 출하지코드는 필수입니다."); return; }
  const res = await fetch(`/api/trkv/departure-mappings/${id}`, {
    method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(body)
  });
  if (res.ok) { changedDmIds.add(id); loadDepartureMappings(dmData.map(r=>({...r}))); }
  else { const e = await res.json().catch(()=>({})); showMsg("dm-msg", e.detail||"저장 실패", false); }
}

function toggleDmBulkEdit() {
  if (!dmBulkEditing) {
    dmBulkEditing = true;
    document.getElementById("dm-bulk-btn").textContent = "일괄 저장";
    renderDm();
  } else {
    saveAllDmRows();
  }
}

async function saveAllDmRows() {
  const prev = dmData.map(r=>({...r}));
  let ok=0, fail=0;
  for (const d of dmData) {
    const name = document.getElementById(`dm-name-${d.id}`)?.value?.trim();
    if (name === undefined) continue;
    const body = {
      departure_name: name,
      departure_code: (document.getElementById(`dm-code-${d.id}`)?.value||"").trim(),
    };
    if (!body.departure_name || !body.departure_code) { fail++; continue; }
    const res = await fetch(`/api/trkv/departure-mappings/${d.id}`, {
      method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(body)
    });
    if (res.ok) ok++; else fail++;
  }
  showMsg("dm-msg", `일괄 저장: ${ok}건 성공${fail?`, ${fail}건 실패`:""}`, fail===0);
  loadDepartureMappings(prev);
}

function showDmAddRow() {
  const tbody = document.getElementById("dm-tbody");
  if (document.getElementById("dm-new-row")) return;
  const tr = document.createElement("tr");
  tr.id = "dm-new-row";
  tr.className = "new-row";
  tr.innerHTML = `
    <td></td><td>신규</td>
    <td><input class="inline-input" id="dm-new-name" placeholder="예: 아산공장" /></td>
    <td><input class="inline-input" id="dm-new-code" placeholder="예: AS" /></td>
    <td>
      <button class="btn btn-sm btn-primary" onclick="addDmRow()">추가</button>
      <button class="btn btn-sm btn-outline" onclick="document.getElementById('dm-new-row').remove()">취소</button>
    </td>`;
  tbody.prepend(tr);
  document.getElementById("dm-new-name").focus();
}

async function addDmRow() {
  const body = {
    departure_name: (document.getElementById("dm-new-name")?.value||"").trim(),
    departure_code: (document.getElementById("dm-new-code")?.value||"").trim(),
  };
  if (!body.departure_name || !body.departure_code) { alert("출하지명과 출하지코드는 필수입니다."); return; }
  const res = await fetch("/api/trkv/departure-mappings", {
    method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(body)
  });
  if (res.ok) { showMsg("dm-msg","추가되었습니다.",true); loadDepartureMappings(); }
  else { const e=await res.json().catch(()=>({})); showMsg("dm-msg",e.detail||"추가 실패",false); }
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
  const ids = [...document.querySelectorAll(".dm-chk:checked")].map(cb=>parseInt(cb.dataset.id));
  if (!ids.length) return;
  if (!confirm(`선택한 ${ids.length}건을 삭제하시겠습니까?`)) return;
  for (const id of ids) await fetch(`/api/trkv/departure-mappings/${id}`,{method:"DELETE"});
  showMsg("dm-msg",`${ids.length}건 삭제되었습니다.`,true);
  loadDepartureMappings();
}

async function deleteDm(id) {
  if (!confirm("이 출하지 매핑을 삭제하시겠습니까?")) return;
  const res = await fetch(`/api/trkv/departure-mappings/${id}`,{method:"DELETE"});
  if (res.ok) { showMsg("dm-msg","삭제되었습니다.",true); loadDepartureMappings(); }
  else showMsg("dm-msg","삭제 실패",false);
}

// ══════════════════════════════════════════════════════════════════
// ③ ODCY 매핑
// ══════════════════════════════════════════════════════════════════
let omData = [];
let omBulkEditing = false;
let changedOmIds = new Set();

async function loadOdcyMappings(prevSnapshot) {
  try {
    const res = await fetch("/api/trkv/odcy-mappings");
    if (!res.ok) throw new Error("API 오류");
    omData = await res.json();
    if (prevSnapshot) {
      changedOmIds = new Set();
      const prevMap = new Map(prevSnapshot.map(r=>[r.id,r]));
      for (const r of omData) {
        const prev = prevMap.get(r.id);
        if (!prev
          || String(r.odcy_destination_name ?? "") !== String(prev.odcy_destination_name ?? "")
          || String(r.odcy_name             ?? "") !== String(prev.odcy_name             ?? "")
          || String(r.odcy_terminal_type    ?? "") !== String(prev.odcy_terminal_type    ?? "")
          || String(r.odcy_location         ?? "") !== String(prev.odcy_location         ?? "")) {
          changedOmIds.add(r.id);
        }
      }
    }
    omBulkEditing = false;
    document.getElementById("om-bulk-btn").textContent = "일괄 편집";
    renderOm();
  } catch {
    document.getElementById("om-tbody").innerHTML =
      '<tr><td colspan="7" class="empty-msg" style="color:#ef4444">불러오기 실패</td></tr>';
  }
}

function renderOmRowHtml(d, i, editing) {
  if (editing) {
    return `<tr id="om-row-${d.id}" class="editing-row">
      <td><input type="checkbox" class="om-chk" data-id="${d.id}" onchange="updateOmSelCount()" /></td>
      <td>${i + 1}</td>
      <td><input class="inline-input" id="om-dest-${d.id}"  value="${(d.odcy_destination_name||'').replace(/"/g,'&quot;')}" /></td>
      <td><input class="inline-input" id="om-name-${d.id}"  value="${(d.odcy_name||'').replace(/"/g,'&quot;')}" /></td>
      <td><input class="inline-input" id="om-term-${d.id}"  value="${(d.odcy_terminal_type||'').replace(/"/g,'&quot;')}" /></td>
      <td><input class="inline-input" id="om-loc-${d.id}"   value="${(d.odcy_location||'').replace(/"/g,'&quot;')}" /></td>
      <td>
        <button class="btn btn-sm btn-primary" onclick="saveOmRow(${d.id})">저장</button>
        <button class="btn btn-sm btn-outline" onclick="cancelOmRowEdit(${d.id})">취소</button>
      </td>
    </tr>`;
  }
  return `<tr id="om-row-${d.id}" class="${changedOmIds.has(d.id) ? 'row-changed' : ''}">
    <td><input type="checkbox" class="om-chk" data-id="${d.id}" onchange="updateOmSelCount()" /></td>
    <td>${i + 1}</td>
    <td>${d.odcy_destination_name || ''}</td>
    <td><span class="badge badge-blue">${d.odcy_name || ''}</span></td>
    <td>${d.odcy_terminal_type ? `<span class="badge badge-green">${d.odcy_terminal_type}</span>` : '<span style="color:#9ca3af">-</span>'}</td>
    <td>${d.odcy_location ? `<span class="badge badge-purple">${d.odcy_location}</span>` : '<span style="color:#9ca3af">-</span>'}</td>
    <td>
      <button class="btn btn-sm btn-outline" onclick="editOmRow(${d.id})">수정</button>
      <button class="btn btn-sm btn-danger"  onclick="deleteOm(${d.id})">삭제</button>
    </td>
  </tr>`;
}

function renderOm() {
  const tbody = document.getElementById("om-tbody");
  if (!omData.length) {
    tbody.innerHTML = '<tr><td colspan="7" class="empty-msg">등록된 ODCY 매핑이 없습니다.</td></tr>';
    updateOmSelCount(); return;
  }
  tbody.innerHTML = omData.map((d, i) => renderOmRowHtml(d, i, omBulkEditing)).join("");
  updateOmSelCount();
}

function editOmRow(id) {
  const d = omData.find(x=>x.id===id);
  if (!d) return;
  const tr = document.getElementById(`om-row-${id}`);
  if (tr) tr.outerHTML = renderOmRowHtml(d, omData.indexOf(d), true);
}

function cancelOmRowEdit(id) {
  const d = omData.find(x=>x.id===id);
  if (!d) return;
  const tr = document.getElementById(`om-row-${id}`);
  if (tr) tr.outerHTML = renderOmRowHtml(d, omData.indexOf(d), false);
}

async function saveOmRow(id) {
  const body = {
    odcy_destination_name: (document.getElementById(`om-dest-${id}`)?.value||"").trim(),
    odcy_name:             (document.getElementById(`om-name-${id}`)?.value||"").trim(),
    odcy_terminal_type:    (document.getElementById(`om-term-${id}`)?.value||"").trim(),
    odcy_location:         (document.getElementById(`om-loc-${id}`)?.value||"").trim(),
  };
  if (!body.odcy_destination_name || !body.odcy_name) { alert("ODCY 도착지명과 ODCY명은 필수입니다."); return; }
  const res = await fetch(`/api/trkv/odcy-mappings/${id}`, {
    method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(body)
  });
  if (res.ok) { changedOmIds.add(id); loadOdcyMappings(omData.map(r=>({...r}))); }
  else { const e=await res.json().catch(()=>({})); showMsg("om-msg",e.detail||"저장 실패",false); }
}

function toggleOmBulkEdit() {
  if (!omBulkEditing) {
    omBulkEditing = true;
    document.getElementById("om-bulk-btn").textContent = "일괄 저장";
    renderOm();
  } else {
    saveAllOmRows();
  }
}

async function saveAllOmRows() {
  const prev = omData.map(r=>({...r}));
  let ok=0, fail=0;
  for (const d of omData) {
    const dest = document.getElementById(`om-dest-${d.id}`)?.value?.trim();
    if (dest === undefined) continue;
    const body = {
      odcy_destination_name: dest,
      odcy_name:          (document.getElementById(`om-name-${d.id}`)?.value||"").trim(),
      odcy_terminal_type: (document.getElementById(`om-term-${d.id}`)?.value||"").trim(),
      odcy_location:      (document.getElementById(`om-loc-${d.id}`)?.value||"").trim(),
    };
    if (!body.odcy_destination_name || !body.odcy_name) { fail++; continue; }
    const res = await fetch(`/api/trkv/odcy-mappings/${d.id}`, {
      method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(body)
    });
    if (res.ok) ok++; else fail++;
  }
  showMsg("om-msg",`일괄 저장: ${ok}건 성공${fail?`, ${fail}건 실패`:""}`,fail===0);
  loadOdcyMappings(prev);
}

function showOmAddRow() {
  const tbody = document.getElementById("om-tbody");
  if (document.getElementById("om-new-row")) return;
  const tr = document.createElement("tr");
  tr.id = "om-new-row";
  tr.className = "new-row";
  tr.innerHTML = `
    <td></td><td>신규</td>
    <td><input class="inline-input" id="om-new-dest" placeholder="예: SB청암" /></td>
    <td><input class="inline-input" id="om-new-name" placeholder="예: 세방(주)" /></td>
    <td><input class="inline-input" id="om-new-term" placeholder="예: 배후단지" /></td>
    <td><input class="inline-input" id="om-new-loc"  placeholder="예: 부산신항" /></td>
    <td>
      <button class="btn btn-sm btn-primary" onclick="addOmRow()">추가</button>
      <button class="btn btn-sm btn-outline" onclick="document.getElementById('om-new-row').remove()">취소</button>
    </td>`;
  tbody.prepend(tr);
  document.getElementById("om-new-dest").focus();
}

async function addOmRow() {
  const body = {
    odcy_destination_name: (document.getElementById("om-new-dest")?.value||"").trim(),
    odcy_name:             (document.getElementById("om-new-name")?.value||"").trim(),
    odcy_terminal_type:    (document.getElementById("om-new-term")?.value||"").trim(),
    odcy_location:         (document.getElementById("om-new-loc")?.value||"").trim(),
  };
  if (!body.odcy_destination_name || !body.odcy_name) { alert("ODCY 도착지명과 ODCY명은 필수입니다."); return; }
  const res = await fetch("/api/trkv/odcy-mappings", {
    method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(body)
  });
  if (res.ok) { showMsg("om-msg","추가되었습니다.",true); loadOdcyMappings(); }
  else { const e=await res.json().catch(()=>({})); showMsg("om-msg",e.detail||"추가 실패",false); }
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
  const ids = [...document.querySelectorAll(".om-chk:checked")].map(cb=>parseInt(cb.dataset.id));
  if (!ids.length) return;
  if (!confirm(`선택한 ${ids.length}건을 삭제하시겠습니까?`)) return;
  for (const id of ids) await fetch(`/api/trkv/odcy-mappings/${id}`,{method:"DELETE"});
  showMsg("om-msg",`${ids.length}건 삭제되었습니다.`,true);
  loadOdcyMappings();
}

async function deleteOm(id) {
  if (!confirm("이 ODCY 매핑을 삭제하시겠습니까?")) return;
  const res = await fetch(`/api/trkv/odcy-mappings/${id}`,{method:"DELETE"});
  if (res.ok) { showMsg("om-msg","삭제되었습니다.",true); loadOdcyMappings(); }
  else showMsg("om-msg","삭제 실패",false);
}

// ══════════════════════════════════════════════════════════════════
// 통합 업로드
// ══════════════════════════════════════════════════════════════════
function downloadUnified() { window.location.href = "/api/trkv/template"; }

async function uploadUnified() {
  const fileInput = document.getElementById("unified-file");
  const file = fileInput.files[0];
  if (!file) return;
  const prevPmData = pmData.map(r=>({...r}));
  const prevDmData = dmData.map(r=>({...r}));
  const prevOmData = omData.map(r=>({...r}));
  const fd = new FormData();
  fd.append("file", file);
  const msgEl = document.getElementById("unified-msg");
  msgEl.textContent = "업로드 중...";
  msgEl.className   = "upload-result";
  msgEl.style.display = "inline";
  try {
    const res  = await fetch("/api/trkv/upload", { method:"POST", body:fd });
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
// 초기화
// ══════════════════════════════════════════════════════════════════
document.addEventListener("DOMContentLoaded", () => {
  loadPortMappings();
  loadDepartureMappings();
  loadOdcyMappings();
});
