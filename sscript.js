const fields = [
    "CRT",
    "SN ONT",
    "MAC",
    "PANJANG KABEL",
    "Status"
];

let cirCounter = 0;

function parseCIR() {
    const text = document.getElementById("cirInput").value;
    const tbody = document.querySelector("#resultTable tbody");

    if (!text.trim()) return;

    cirCounter++;

    let cirHTML = `
        <tr>
            <td colspan="2" style="background:#d9edf7;font-weight:bold;">
                CIR #${cirCounter}
            </td>
        </tr>
    `;

    fields.forEach(field => {
        const regex = new RegExp(field + "\\s*:\\s*(.*)", "i");
        const match = text.match(regex);
        const value = match ? match[1].trim() : "-";

        cirHTML += `
            <tr>
                <td>${field}</td>
                <td>${value}</td>
            </tr>
        `;
    });

    // 🔥 INSERT DI ATAS, TANPA HAPUS DATA LAMA
    tbody.insertAdjacentHTML("afterbegin", cirHTML);

    // optional: bersihin textarea
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
    a.click();
    URL.revokeObjectURL(url);
}

function clearAll() {
    document.querySelector("#resultTable tbody").innerHTML = "";
    cirCounter = 0;
}
