let dataA = [];
let dataB = [];
let resultData = [];

function readExcel(file, callback) {
    const reader = new FileReader();
    reader.onload = e => {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        callback(XLSX.utils.sheet_to_json(sheet, { defval: "" }));
    };
    reader.readAsArrayBuffer(file);
}

function compareExcel() {
    const fileA = document.getElementById("fileA").files[0];
    const fileB = document.getElementById("fileB").files[0];
    if (!fileA || !fileB) {
        alert("Upload Excel A & B dulu");
        return;
    }

    readExcel(fileA, rowsA => {
        dataA = rowsA;
        readExcel(fileB, rowsB => {
            dataB = rowsB;
            processCompare();
        });
    });
}

function processCompare() {
    const tbody = document.querySelector("#resultTable tbody");
    tbody.innerHTML = "";
    resultData = [];

    dataB.forEach(b => {
        const match = dataA.find(a => a["Cust ID Klien"] === b["Cust ID Klien"]);

        let status = "";
        let detail = "";
        let cls = "";

        if (!match) {
            status = "NOT FOUND";
            detail = "Tidak ada di Excel A";
            cls = "notfound";
        } else if (JSON.stringify(match) === JSON.stringify(b)) {
            status = "MATCH";
            detail = "Data cocok";
            cls = "match";
        } else {
            status = "FAILED";
            detail = "Data tidak sesuai";
            cls = "failed";
        }

        resultData.push({
            "Cust ID Klien": b["Cust ID Klien"],
            "Status Compare": status,
            "Detail": detail
        });

        tbody.innerHTML += `
            <tr class="${cls}">
                <td>${b["Cust ID Klien"]}</td>
                <td>${status}</td>
                <td>${detail}</td>
            </tr>
        `;
    });
}

function exportExcel() {
    const ws = XLSX.utils.json_to_sheet(resultData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Compare Result");
    XLSX.writeFile(wb, "hasil_komper.xlsx");
}

function clearAll() {
    document.querySelector("#resultTable tbody").innerHTML = "";
    resultData = [];
}
