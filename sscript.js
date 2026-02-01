const fields = [
    "Root Cause",
    "Action",
    "Impact Service",
    "Down Time",
    "Clear Time",
    "Duration",
    "Status",
    "Vendor",
    "Problem Location",
    "Area",
    "Span Problem"
];

let cirCounter = 0;

function parseCIR() {
    const text = document.getElementById("cirInput").value;
    const tbody = document.querySelector("#resultTable tbody");

    if (!text.trim()) return;

    cirCounter++;

    // === HEADER PEMBATAS CIR ===
    const headerRow = document.createElement("tr");
    headerRow.innerHTML = `
        <td colspan="2" style="background:#d9edf7;font-weight:bold;">
            CIR #${cirCounter}
        </td>
    `;
    tbody.prepend(headerRow);

    // === DATA FIELD ===
    fields.slice().reverse().forEach(field => {
        const regex = new RegExp(field + "\\s*:\\s*(.*)", "i");
        const match = text.match(regex);
        const value = match ? match[1].trim() : "-";

        const row = document.createElement("tr");
        row.innerHTML = `
            <td>${field}</td>
            <td>${value}</td>
        `;
        tbody.prepend(row);
    });

    // optional: clear textarea setelah parse
    document.getElementById("cirInput").value = "";
}

function exportExcel() {
    const table = document.getElementById("resultTable");
    const html = table.outerHTML;

    const blob = new Blob(
        ['\ufeff' + html],
        { type: 'application/vnd.ms-excel' }
    );

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "CIR_Result.xls";
    a.clic
