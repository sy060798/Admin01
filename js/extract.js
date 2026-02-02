console.log("extract.js loaded");

// ===============================
// ----- LIST -----
// FIELD YANG DIAMBIL
// ===============================
const EXPORT_FIELDS = [
    "Site",
    "Ticket/WO"
];

// ===============================
let rawData = [];
let resultData = [];

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

        document.getElementById("runBtn").disabled = false;
        alert("File berhasil di-upload. Klik JALAN untuk proses.");
    };
    reader.readAsArrayBuffer(file);
});

// ===============================
// PROSES EXTRACT
// ===============================
function runExtract() {
    resultData = [];

    rawData.forEach(row => {
        let obj = {};

        EXPORT_FIELDS.forEach(col => {
            obj[col] = row[col] || "";
        });

        resultData.push(obj);
    });

    renderTable();
}

// ===============================
// TAMPILKAN TABLE
// ===============================
function renderTable() {
    const thead = document.querySelector("#dataTable thead");
    const tbody = document.querySelector("#dataTable tbody");

    thead.innerHTML = "";
    tbody.innerHTML = "";

    if (!resultData.length) return;

    // header
    thead.innerHTML = `
        <tr>
            ${EXPORT_FIELDS.map(c => `<th>${c}</th>`).join("")}
        </tr>
    `;

    // body
    resultData.forEach(r => {
        tbody.insertAdjacentHTML("beforeend", `
            <tr>
                ${EXPORT_FIELDS.map(c => `<td>${r[c]}</td>`).join("")}
            </tr>
        `);
    });
}

// ===============================
// EXPORT EXCEL
// ===============================
function exportExcel() {
    if (!resultData.length) {
        alert("Tidak ada data untuk diexport");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(resultData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Extract Result");
    XLSX.writeFile(wb, "hasil_extract.xlsx");
}
