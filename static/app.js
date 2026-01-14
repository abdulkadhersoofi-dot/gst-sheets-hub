// ---------- GLOBAL STATE ----------
let currentCompanyId = null;
let currentCompanyName = null;
let currentSheetName = null;

let sheetValues = [];   // 2D array of displayed values
let editableMask = [];  // 2D array of booleans

// ---------- HELPERS ----------
function setStatus(id, msg) {
  const el = document.getElementById(id);
  if (el) el.textContent = msg;
}

function showCompanyScreen() {
  document.getElementById("companyScreen").style.display = "block";
  document.getElementById("companyDetailScreen").style.display = "none";
}

function showCompanyDetailScreen() {
  document.getElementById("companyScreen").style.display = "none";
  document.getElementById("companyDetailScreen").style.display = "block";
}

// ---------- LOAD COMPANIES ----------
async function loadCompanies() {
  setStatus("statusHome", "Loading companies...");
  const resp = await fetch("/companies");
  const data = await resp.json();

  const listEl = document.getElementById("companyList");
  listEl.innerHTML = "";

  data.forEach((c) => {
    const row = document.createElement("div");
    row.className = "company-row";
    row.textContent = `${c.CompanyId} – ${c.CompanyName}`;
    row.onclick = () => selectCompany(c);
    listEl.appendChild(row);
  });

  setStatus("statusHome", "Type to search and select a company.");
}

// ---------- SELECT COMPANY ----------
async function selectCompany(company) {
  currentCompanyId = company.CompanyId;
  currentCompanyName = company.CompanyName;

  document.getElementById("currentCompanyHeading").textContent =
    `${company.CompanyId} – ${company.CompanyName}`;

  showCompanyDetailScreen();
  await loadSheetsForCompany();
}

// ---------- LOAD SHEET TABS ----------
async function loadSheetsForCompany() {
  if (!currentCompanyId) return;

  setStatus("statusLoading", "Loading sheets...");
  const resp = await fetch(`/company/${currentCompanyId}/sheets`);
  const sheets = await resp.json();

  const tabsEl = document.getElementById("sheetTabs");
  tabsEl.innerHTML = "";

  sheets.forEach((s, idx) => {
    const tab = document.createElement("div");
    tab.className = "sheet-tab";
    tab.textContent = s.sheetName;

    tab.onclick = () => {
      document
        .querySelectorAll(".sheet-tab")
        .forEach((t) => t.classList.remove("active"));
      tab.classList.add("active");
      loadSheet(s.sheetName);
    };

    if (idx === 0) tab.classList.add("active");
    tabsEl.appendChild(tab);
  });

  if (sheets.length > 0) {
    currentSheetName = sheets[0].sheetName;
    await loadSheet(currentSheetName);
  } else {
    setStatus("statusLoading", "No sheets found.");
  }
}

// ---------- LOAD A SHEET ----------
async function loadSheet(sheetName) {
  if (!currentCompanyId) return;
  currentSheetName = sheetName;

  setStatus("statusLoading", `Loading sheet: ${sheetName}...`);

  const resp = await fetch(
    `/sheet/${currentCompanyId}?sheet=${encodeURIComponent(sheetName)}`
  );
  const data = await resp.json();

  sheetValues = data.values || [];
  editableMask = data.editable || [];

  renderSheet();
  setStatus("statusLoading", `Loaded sheet: ${sheetName}`);
}

// ---------- RENDER SHEET TABLE ----------
function renderSheet() {
  const tbody = document.getElementById("sheetBody");
  tbody.innerHTML = "";

  sheetValues.forEach((row, rIdx) => {
    const tr = document.createElement("tr");

    row.forEach((cell, cIdx) => {
      const td = document.createElement("td");
      td.textContent = cell;

      const editable =
        editableMask[rIdx] && editableMask[rIdx][cIdx] === true;

      if (editable) {
        td.contentEditable = "true";
      } else {
        td.contentEditable = "false";
        td.classList.add("locked-cell"); // yellow in CSS
      }

      td.dataset.row = rIdx;
      td.dataset.col = cIdx;

      td.addEventListener("input", () => {
        sheetValues[rIdx][cIdx] = td.textContent;
      });

      tr.appendChild(td);
    });

    // Action cell: add row below this row
    const actionTd = document.createElement("td");
    const addBtn = document.createElement("button");
    addBtn.textContent = "+ row";
    addBtn.onclick = () => insertRowBelow(rIdx);
    actionTd.appendChild(addBtn);
    tr.appendChild(actionTd);

    tbody.appendChild(tr);
  });
}

// ---------- INSERT ROW BELOW ----------
async function insertRowBelow(rowIndex) {
  if (!currentCompanyId || !currentSheetName) return;

  setStatus("statusLoading", "Inserting row...");
  const resp = await fetch(
    `/sheet/${currentCompanyId}/insert-row`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        sheet: currentSheetName,
        row_index: rowIndex,
      }),
    }
  );

  const data = await resp.json();
  if (!resp.ok) {
    alert(data.error || "Failed to insert row");
    setStatus("statusLoading", "Insert row failed.");
    return;
  }

  // Refresh local state
  sheetValues = data.values || [];
  editableMask = data.editable || [];
  renderSheet();

  setStatus(
    "statusLoading",
    `Row inserted below row ${rowIndex + 1} (sheet reloaded).`
  );
}

// ---------- SAVE CHANGES ----------
async function saveChanges() {
  if (!currentCompanyId || !currentSheetName) return;

  setStatus("statusLoading", "Saving changes...");

  const resp = await fetch(
    `/sheet/${currentCompanyId}/update?sheet=${encodeURIComponent(
      currentSheetName
    )}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        values: sheetValues,
        editable: editableMask,
      }),
    }
  );

  const data = await resp.json();
  if (!resp.ok) {
    alert(data.error || "Save failed");
    setStatus("statusLoading", "Save failed.");
    return;
  }

  setStatus("statusLoading", "Saved successfully.");
}

// ---------- CLONE SHEET (APR -> NEW MONTH) ----------
async function cloneSheetFromApr() {
  if (!currentCompanyId) return;

  const newNameInput = document.getElementById("newSheetName");
  const newSheetName = (newNameInput.value || "").trim();
  if (!newSheetName) {
    alert("Enter the new sheet name, e.g. DEC 25");
    return;
  }

  setStatus("statusLoading", "Cloning sheet from APR...");
  const resp = await fetch(`/sheet/${currentCompanyId}/clone`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      source_sheet: "APR 25", // adjust if needed
      new_sheet: newSheetName,
    }),
  });

  const data = await resp.json();
  if (!resp.ok) {
    alert(data.error || "Clone failed");
    setStatus("statusLoading", "Clone failed.");
    return;
  }

  newNameInput.value = "";
  await loadSheetsForCompany();
  setStatus("statusLoading", `Created new sheet: ${newSheetName}`);
}

// ---------- EVENT WIRING ----------
document.addEventListener("DOMContentLoaded", () => {
  // Initial screen
  showCompanyScreen();
  loadCompanies();

  // Back button
  const backBtn = document.getElementById("backBtn");
  if (backBtn) backBtn.onclick = showCompanyScreen;

  // Reload button
  const reloadBtn = document.getElementById("reloadBtn");
  if (reloadBtn)
    reloadBtn.onclick = () => {
      if (currentSheetName) loadSheet(currentSheetName);
    };

  // Save button
  const saveBtn = document.getElementById("saveBtn");
  if (saveBtn) saveBtn.onclick = saveChanges;

  // Clone sheet button
  const cloneSheetBtn = document.getElementById("cloneSheetBtn");
  if (cloneSheetBtn) cloneSheetBtn.onclick = cloneSheetFromApr;

  // Company search filter
  const searchInput = document.getElementById("companySearch");
  if (searchInput) {
    searchInput.addEventListener("input", () => {
      const q = searchInput.value.toLowerCase();
      document
        .querySelectorAll("#companyList .company-row")
        .forEach((row) => {
          const txt = row.textContent.toLowerCase();
          row.style.display = txt.includes(q) ? "block" : "none";
        });
    });
  }
});
