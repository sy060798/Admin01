const fields = ["Cust ID Klien", "Customer Name", "ONT", "MAC ONT", "Kabel Precon 100 Old", "Kabel Precon 150 Old", "Kabel Precon 200 Old", "Detail Pekerjaan", "Status"];
let cirCounter = 0;

// === LOAD EXCEL ===
function loadExcel() {
    const fileInput = document.getElementById("excelFile");
    if (!fileInput.files.length) return alert("Upload file Excel dulu");

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet);

        json.forEach(row => addCIRFromExcel(row));
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
}

// === TAMBAH 1 CIR ===
function addCIRFromExcel(row) {
    const tbody = document.querySelector("#resultTable tbody");
    cirCounter++;

    let html = `
        <tr>
            <td colspan="2" style="background:#d9edf7;font-weight:bold;">
                CIR #${cirCounter}
            </td>
        </tr>
    `;

    fields.forEach(field => {
        html += `
            <tr>
                <td>${field}</td>
                <td>${row[field] || "-"}</td>
            </tr>
        `;
    });

    tbody.insertAdjacentHTML("afterbegin", html);
}

// === MANUAL PARSE (TEXTAREA) ===
function parseCIR() {
    const text = document.getElementById("cirInput").value;
    if (!text.trim()) return;

    let row = {};
    fields.forEach(f => {
        const match = text.match(new RegExp(f + "\\s*:\\s*(.*)", "i"));
        row[f] = match ? match[1].trim() : "-";
    });

    addCIRFromExcel(row);
    document.getElementById("cirInput").value = "";
}

// === CLEAR ===
function clearAll() {
    document.querySelector("#resultTable tbody").innerHTML = "";
    cirCounter = 0;
}
