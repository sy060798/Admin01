const fields = [
    "Cust ID Klien",
    "Customer Name",
    "ONT",
    "MAC ONT",
    "Kabel Precon 100 Old",
    "Kabel Precon 150 Old",
    "Kabel Precon 200 Old",
    "Detail Pekerjaan",
    "Status"
];

let cirCounter = 0;

// ================== UTIL ==================
function normalize(text) {
    return String(text || "")
        .trim()
        .toLowerCase()
        .replace(/\s+/g, " ");
}

// ================== FILE UPLOAD ==================
function openFileDialog() {
    document.getElementById("excelFile").click();
}

function loadExcel() {
    const input = document.getElementById("excelFile");
    if (!input.files.length) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        rows.forEach(row => addCIR(row));
    };

    reader.readAsArrayBuffer(input.files[0]);
    input.value = "";
}

// ================== MANUAL PARSE ==================
function parseCIR() {
    const text = document.getElementById("cirInput").value;
    if (!text.trim()) return;

    const row = {};
    fields.forEach(field => {
        const match = text.match(new RegExp(field + "\\s*:\\s*(.*)", "i"));
        row[field] = match ? match[1].trim() : "";
    });

    addCIR(row);
    document.getElementById("cirInput").value = "";
}

// ================== RENDER ==================
function addCIR(row) {
    const tbody = document.querySelector("#resultTable tbody");
    cirCounter++;

    const normalizedRow = {};
    Object.keys(row).forEach(k => {
        normalizedRow[normalize(k)] = row[k];
    });

    let html = `
        <tr>
            <td colspan="2" style="background:#d9edf7;font-weight:bold;">
                CIR #${cirCounter}
            </td>
        </tr>
    `;

    fields.forEach(field => {
        const value =
            normalizedRow[normalize(field)]
                ? normalizedRow[normalize(field)]
                : `<span style="color:#999">-</span>`;

        html += `
            <tr>
                <td>${field}</td>
                <td>${value}</td>
            </tr>
        `;
    });

    tbody.insertAdjacentHTML("afterbegin", html);
}

// ================== EXPORT ==================
function exportExcel() {
    const table = document.getElementById("resultTable");
    const blob = new Blob(
        ['\ufeff' + table.outerHTML],
        { type: 'application/vnd.ms-excel' }
    );

    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "CIR_Result.xls";
    a.click();
    URL.revokeObjectURL(a.href);
}

// ================== CLEAR ==================
function clearAll() {
    document.querySelector("#resultTable tbody").innerHTML = "";
    cirCounter = 0;
}
