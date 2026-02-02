console.log("cirbackbone.js loaded");

// ===============================
// ----- LIST -----
// ===============================
const EXPORT_FIELDS = [
    "TT FLP",
    "Span Problem",
    "Root Cause",
    "Action"
];

let rawData = [];
let resultData = [];

// MODE SWITCH
document.querySelectorAll("input[name='mode']").forEach(r => {
    r.addEventListener("change", () => {
        document.getElementById("excelBox").style.display =
            r.value === "excel" && r.checked ? "block" : "none";
        document.getElementById("manualBox").style.display =
            r.value === "manual" && r.checked ? "block" : "none";
    });
});

// READ EXCEL
document.getElementById("excelFile").addEventListener("change", e => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = ev => {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        rawData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        alert("Excel siap diproses");
    };
    reader.readAsArrayBuffer(file);
});

// ADD MANUAL
function addManual() {
    const text = document.getElementById("cirText").value;
    if (!text.trim()) return alert("CIR kosong");

    let obj = {};
    EXPORT_FIELDS.forEach(f => obj[f] = "");

    text.split("\n").forEach(line => {
        EXPORT_FIELDS.forEach(key => {
            const m = line.match(new RegExp(`^${key}\\s*:\\s*(.*)$`, "i"));
            if (m) obj[key] = m[1].trim();
        });
    });

    resultData.push(obj);
    document.getElementById("cirText").value = "";
    renderTable();
}

// PROCESS
function runProcess() {
    if (rawData.length) {
        resultData = rawData.map(r => {
            let o = {};
            EXPORT_FIELDS.forEach(f => o[f] = r[f] || "");
            return o;
        });
    }

    if (!resultData.length) {
        alert("Tidak ada data");
        return;
    }

    renderTable();
}

// TABLE
function renderTable() {
    const thead = document.querySelector("#dataTable thead");
    const tbody = document.querySelector("#dataTable tbody");

    thead.innerHTML =
        `<tr>${EXPORT_FIELDS.map(f => `<th>${f}</th>`).join("")}</tr>`;

    tbody.innerHTML = "";
    resultData.forEach(r => {
        tbody.insertAdjacentHTML("beforeend",
            `<tr>${EXPORT_FIELDS.map(f => `<td>${r[f]}</td>`).join("")}</tr>`
        );
    });
}

// EXPORT
function exportExcel() {
    if (!resultData.length) return alert("Tidak ada data");

    const ws = XLSX.utils.json_to_sheet(resultData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "CIR Backbone");
    XLSX.writeFile(wb, "hasil_CIR_backbone.xlsx");
}
