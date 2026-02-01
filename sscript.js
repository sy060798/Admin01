let dataA = [];
let dataB = [];
let resultData = [];

// ====== SETTING DATA KOMPER (TINGGAL EDIT DI SINI) ======
const KEY_FIELD = "Cust ID Klien";   // kolom kunci utama (ID unik)

const COMPARE_FIELDS = [
    "Customer Name",
    "Status",
    "ONT",
    "MAC ONT"
];
// =======================================================

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
        const match = dataA.find(a => a[KEY_FIELD] === b[KEY_FIELD]);

        let status = "";
        let detailArr = [];
        let cls = "";

        if (!match) {
            status = "NOT FOUND";
            detailArr.push("Tidak ada di Excel A");
            cls = "notfound";
        } else {
            COMPARE_FIELDS.forEach(field => {
                if ((match[field] || "") !== (b[field] || "")) {
                    detailArr.push(`${field} tidak sama`);
                }
            });

            if (detailArr.length === 0) {
                status = "MATCH";
                detailArr.push("Semua data cocok");
                cls = "match";
            } else {
                status = "FAILED";
                cls = "failed";
            }
        }

        const detail = detailArr.join(" | ");

        resultData.push({
    [KEY_FIELD]: b[KEY_FIELD],
    "ONT": b["ONT"] || "",
    "MAC ONT": b["MAC ONT"] || "",
    "MAC STB": b["MAC STB"] || "",
    "Status Compare": status,
    "Detail": detail
});

        tbody.innerHTML += `
            <tr class="${cls}">
                <td>${b[KEY_FIELD]}</td>
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
