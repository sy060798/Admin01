console.log("extract.js loaded");

let rawData = [];
let exportData = [];

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
        renderTable();
    };
    reader.readAsArrayBuffer(file);
});

// ===============================
// RENDER TABLE
// ===============================
function renderTable() {
    const thead = document.querySelector("#dataTable thead");
    const tbody = document.querySelector("#dataTable tbody");

    thead.innerHTML = "";
    tbody.innerHTML = "";

    if (!rawData.length) return;

    const cols = Object.keys(rawData[0]);

    thead.innerHTML = `
        <tr>
            ${cols.map(c => `<th>${c}</th>`).join("")}
        </tr>
    `;

    rawData.forEach(row => {
        tbody.insertAdjacentHTML("beforeend", `
            <tr>
                ${cols.map(c => `<td>${row[c]}</td>`).join("")}
            </tr>
        `);
    });

    // ===============================
    // PILIH DATA YANG MAU DIEXPORT
    // ===============================
    exportData = rawData.map(r => ({
        "Cust ID Klien": r["Cust ID Klien"],
        "ONT": r["ONT"],
        "MAC ONT": r["MAC ONT"],
        "MAC STB": r["MAC STB"]
    }));
}

// ===============================
// EXPORT EXCEL
// ===============================
function exportExcel() {
    if (!exportData.length) {
        alert("Tidak ada data");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Extract Result");
    XLSX.writeFile(wb, "hasil_extract.xlsx");
}
