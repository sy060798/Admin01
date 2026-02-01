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

function parseCIR() {
    const text = document.getElementById("cirInput").value;
    const tbody = document.querySelector("#resultTable tbody");
    tbody.innerHTML = "";

    fields.forEach(field => {
        const regex = new RegExp(field + "\\s*:\\s*(.*)", "i");
        const match = text.match(regex);
        const value = match ? match[1].trim() : "-";

        const row = document.createElement("tr");

        const fieldCell = document.createElement("td");
        fieldCell.textContent = field;

        const valueCell = document.createElement("td");
        valueCell.textContent = value;

        row.appendChild(fieldCell);
        row.appendChild(valueCell);
        tbody.appendChild(row);
    });
}

function exportExcel() {
    const table = document.getElementById("resultTable");
    let html = table.outerHTML;

    const blob = new Blob(
        ['\ufeff' + html],
        { type: 'application/vnd.ms-excel' }
    );

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "CIR_Result.xls";
    a.click();
    URL.revokeObjectURL(url);
}
