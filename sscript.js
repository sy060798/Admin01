const KEY_FIELD = "Cust ID Klien";

let dataA = [];
let dataB = [];
let resultData = [];

// ===============================
// BACA EXCEL
// ===============================
function readExcel(file, callback) {
    const reader = new FileReader();
    reader.onload = e => {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        callback(XLSX.utils.sheet_to_json(sheet, { defval: "" }));
    };
    reader.readAsArrayBuffer(file);
}

// ===============================
// LOAD & COMPARE
// ===============================
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

// ===============================
// PROSES COMPARE
// ===============================
function processCompare() {
    const tbody = document.querySelector("#resultTable tbody");
    tbody.innerHTML = "";
    resultData = [];

    dataB.forEach(b => {
        const match = dataA.find(a => a[KEY_FIELD] === b[KEY_FIELD]);

        let status = "";
        let detail = "";
        let cls = "";

        let ont = "";
        let macOnt = "";
        let macStb = "";

        if (!match) {
            status = "NOT FOUND";
            detail = "Tidak ada di Excel A";
            cls = "notfound";
        } else {
            // ONT
            if ((match["ONT"] || "").trim() === (b["ONT"] || "").trim()) {
                ont = b["ONT"];
            }

            // MAC ONT
            if ((match["MAC ONT"] || "").trim() === (b["MAC ONT"] || "").trim()) {
                macOnt = b["MAC ONT"];
            }

            // MAC STB
            if ((match["MAC STB"] || "").trim() === (b["MAC STB"] || "").trim()) {
                macStb = b["MAC STB"];
            }

            if (
                ont &&
                macOnt &&
                macStb
            ) {
                status = "MATCH";
                detail = "Data cocok";
                cls = "match";
            } else {
                status = "FAILED";
                detail = "Ada data tidak sesuai";
                cls = "failed";
            }
        }

        resultData.push({
            [KEY_FIELD]: b[KEY_FIELD],
            "ONT": ont,
            "MAC ONT": macOnt,
            "MAC STB": macStb,
            "Status": status,
            "Detail": detail
        });

        tbody.innerHTML += `
            <tr class="${cls}">
                <td>${b[KEY_FIELD]}</td>
                <td>${ont}</td>
                <td>${macOnt}</td>
                <td>${macStb}</td>
                <td>${status}</td>
                <td>${detail}</td>
            </tr>
        `;
    });
}

// ===============================
// EXPORT EXCEL
// ===============================
function exportExcel() {
    const ws = XLSX.utils.json_to_sheet(resultData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Compare Result");
    XLSX.writeFile(wb, "hasil_compare.xlsx");
}

// ===============================
// CLEAR
// ===============================
function clearAll() {
    document.querySelector("#resultTable tbody").innerHTML = "";
    resultData = [];
}
