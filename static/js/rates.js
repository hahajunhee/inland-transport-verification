let editingId = null;

document.addEventListener("DOMContentLoaded", () => {
  loadRates();

  document.getElementById("rate-form").addEventListener("submit", async (e) => {
    e.preventDefault();
    await saveRate();
  });

  document.getElementById("cancel-btn").addEventListener("click", cancelEdit);
});

async function loadRates() {
  const chargeType = document.getElementById("filter-charge").value;
  const params = chargeType ? `?charge_type=${encodeURIComponent(chargeType)}` : "";
  const res = await fetch(`/api/rates${params}`);
  const data = await res.json();
  renderRates(data);
}

function renderRates(rates) {
  const tbody = document.getElementById("rates-tbody");
  if (!rates.length) {
    tbody.innerHTML = '<tr><td colspan="10" class="empty-msg">등록된 요율이 없습니다.</td></tr>';
    return;
  }
  tbody.innerHTML = rates.map(r => `
    <tr>
      <td>${r.id}</td>
      <td><span class="badge ${chargeBadgeClass(r.charge_type)}">${r.charge_type}</span></td>
      <td>${r.pickup_code || '<span style="color:#bbb">전체</span>'}</td>
      <td>${r.odcy_code || '<span style="color:#bbb">전체</span>'}</td>
      <td>${r.dest_code || '<span style="color:#bbb">전체</span>'}</td>
      <td>${r.container_type || '<span style="color:#bbb">전체</span>'}</td>
      <td class="money">${fmtMoney(r.unit_price)}</td>
      <td>${r.memo || ''}</td>
      <td>${fmtDate(r.created_at)}</td>
      <td>
        <button class="btn btn-sm btn-outline" onclick="startEdit(${r.id})">수정</button>
        <button class="btn btn-sm btn-danger" onclick="deleteRate(${r.id})">삭제</button>
      </td>
    </tr>
  `).join("");
}

function chargeBadgeClass(ct) {
  const map = { TRKV: "badge-ok", "보관료": "badge-diff", "상하차료": "badge-no-rate", "셔틀비용": "" };
  return map[ct] || "";
}

async function saveRate() {
  const payload = {
    charge_type: document.getElementById("charge_type").value,
    pickup_code: val("pickup_code"),
    odcy_code: val("odcy_code"),
    dest_code: val("dest_code"),
    container_type: val("container_type"),
    unit_price: parseFloat(document.getElementById("unit_price").value),
    memo: val("memo"),
  };

  const method = editingId ? "PUT" : "POST";
  const url = editingId ? `/api/rates/${editingId}` : "/api/rates";
  const res = await fetch(url, {
    method,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  const msg = document.getElementById("form-msg");
  if (res.ok) {
    showMsg(msg, editingId ? "요율이 수정되었습니다." : "요율이 등록되었습니다.", "success");
    cancelEdit();
    loadRates();
  } else {
    const err = await res.json();
    showMsg(msg, `오류: ${err.detail || "저장 실패"}`, "error");
  }
}

async function startEdit(id) {
  const res = await fetch(`/api/rates?charge_type=`);
  const all = await res.json();
  const rate = all.find(r => r.id === id);
  if (!rate) return;

  editingId = id;
  document.getElementById("form-title").textContent = "요율 수정";
  document.getElementById("submit-btn").textContent = "수정 저장";
  document.getElementById("cancel-btn").style.display = "";
  document.getElementById("edit-id").value = id;

  setVal("charge_type", rate.charge_type);
  setVal("pickup_code", rate.pickup_code || "");
  setVal("odcy_code", rate.odcy_code || "");
  setVal("dest_code", rate.dest_code || "");
  setVal("container_type", rate.container_type || "");
  setVal("unit_price", rate.unit_price);
  setVal("memo", rate.memo || "");

  document.getElementById("rate-form").scrollIntoView({ behavior: "smooth" });
}

function cancelEdit() {
  editingId = null;
  document.getElementById("form-title").textContent = "새 요율 등록";
  document.getElementById("submit-btn").textContent = "등록";
  document.getElementById("cancel-btn").style.display = "none";
  document.getElementById("rate-form").reset();
  document.getElementById("form-msg").style.display = "none";
}

async function deleteRate(id) {
  if (!confirm("이 요율을 삭제하시겠습니까?")) return;
  const res = await fetch(`/api/rates/${id}`, { method: "DELETE" });
  if (res.ok) loadRates();
  else alert("삭제 실패");
}

// ─── helpers ───────────────────────────────────────────────
function val(id) {
  const v = document.getElementById(id).value.trim();
  return v || null;
}
function setVal(id, v) {
  document.getElementById(id).value = v ?? "";
}
function fmtMoney(n) {
  if (n == null) return "-";
  return Number(n).toLocaleString("ko-KR") + "원";
}
function fmtDate(s) {
  if (!s) return "-";
  return s.slice(0, 16).replace("T", " ");
}
function showMsg(el, text, type) {
  el.textContent = text;
  el.className = `form-msg ${type}`;
  el.style.display = "block";
  setTimeout(() => { el.style.display = "none"; }, 3000);
}
