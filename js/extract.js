console.log("extract.js loaded");

let rawData = [];
let selectedCols = [];

// ===============================
// READ EXCEL
// ===============================
document.getElementById("excelFile").addEventListener("change", e => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = ev => {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        rawData = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        renderColumnSelector();
        renderTable();
    };
    reader.readAsArrayBuffer(file);
});

// ===============================
// COLUMN CHECKLIST
// ===============================
function renderColumnSelector() {
    const box = document.getElementById("columnSelector");
    box.innerHTML = "<b>Pilih kolom untuk export:</b><br>";

    if (!rawData.length) return;

    const cols = Object.keys(rawData[0]);

    cols.forEach(c => {
        box.insertAdjacentHTML("beforeend", `
            <label style="margin-right:12px">
                <input type="checkbox" value="${c}" checked> ${c}
            </label>
        `);
    });
}

// ===============================
// PRESET FUNCTIONS
// ===============================
function selectPreset(cols) {
    document.querySelectorAll("#columnSelector input").forEach(cb => {
        cb.checked = cols.includes(cb.value);
    });
}

function selectAllCols() {
    document.querySelectorAll("#columnSelector input").forEach(cb => {
        cb.checked = true;
    });
}

function clearCols() {
    document.querySelectorAll("#columnSelector input").forEach(cb => {
        cb.checked = false;
    });
}

// ===============================
// RENDER TABLE PREVIEW
// ===============================
function renderTable() {
    const thead = document.querySelector("#dataTable thead");
    const tbody = document.querySelector("#dataTable tbody");

    thead.innerHTML = "";
    tbody.innerHTML = "";

    if (!rawData.length) return;

    const cols = Object.keys(rawData[0]);

    thead.innerHTML = `<tr>${cols.map(c => `<th>${c}</th>`).join("")}</tr>`;

    rawData.forEach(row => {
        tbody.insertAdjacentHTML("beforeend", `
            <tr>${cols.map(c => `<td>${row[c]}</td>`).join("")}</tr>
        `);
    });
}

// ===============================
// EXPORT EXCEL
// ===============================
function exportExcel() {
    const checked = document.querySelectorAll("#columnSelector input:checked");
    if (!checked.length) {
        alert("Pilih minimal 1 kolom");
        return;
    }

    selectedCols = Array.from(checked).map(cb => cb.value);

    const exportData = rawData.map(row => {
        let obj = {};
        selectedCols.forEach(col => obj[col] = row[col]);
        return obj;
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Extract Result");
    XLSX.writeFile(wb, "hasil_extract.xlsx");
}
