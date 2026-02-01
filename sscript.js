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

// 🔹 normalisasi text (biar header Excel fleksibel)
function normalize(text) {
    return String(text || "")
        .trim()
        .toLowerCase()
        .replace(/\s+/g, " ");
}

// 🔥 BUKA FILE EXPLORER
function openFileDialog() {
    document.getElementById("excelFile").click();
}

// 📥 LOAD EXCEL SETELAH DIPILIH
function loadExcel() {
    const input = document.getElementById("excelFile");
    if (!input.files.length) return;

    const file = input.files[0];
    const reader = new FileReader();

    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        rows.forEach(row => addCIR(row));
    };

    reader.readAsArrayBuffer(file);

    // reset supaya bisa upload file yg sama lagi
    input.value = "";
}

function addCIR(row) {
    const tbody = document.querySelector("#resultTable tbody");
    cirCounter++;

    // buat map header excel
    const normalizedRow = {};
    Object.keys(row).forEach(key => {
        normalizedRow[normalize(key)] = row[key];
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
            normalizedRow[normalize(field)] !== undefined &&
            normalizedRow[normalize(field)] !== ""
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

// 🧹 CLEAR SEMUA DATA
function clearAll() {
    document.querySelector("#resultTable tbody").innerHTML = "";
    cirCounter = 0;
}
