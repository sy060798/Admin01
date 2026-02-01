// ===============================
// CONFIG
// ===============================
console.log("sscript.js loaded OK"); // cek di Console

const KEY_FIELD = "Cust ID Klien";

let dataA = [];
let dataB = [];
let resultData = [];

// ===============================
// NORMALIZE DATA (PENTING)
// ===============================
function normalize(val) {
    return String(val || "")
        .replace(/\u00A0/g, "") // hapus spasi aneh (non-breaking space)
        .replace(/[\s\-:]/g, "") // hapus spasi, - dan :
        .toUpperCase()
        .trim();
}

// ===============================
// BACA EXCEL
// ===============================
function readExcel(file, callback) {
    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        callback(rows);
    };
    reader.readAsArrayBuffer(file);
}

// ===============================
// LOAD & COMPARE
// ===============================
function compareExcel() {
    const fileAInput = document.getElementById("fileA");
    const fileBInput = document.getElementById("fileB");

    if (!fileAInput || !fileBInput) {
        alert("Input file tidak ditemukan");
        return;
    }

    const fileA = fileAInput.files[0];
    const fileB = fileBInput.files[0];

    if (!fileA || !fileB) {
        alert("Upload Excel A & B dulu");
        return;
    }

    readExcel(fileA, function (rowsA) {
        dataA = rowsA;
        readExcel(fileB, function (rowsB) {
            dataB = rowsB;
            processCompare();
        });
    });
}

// ===============================
// PROSES COMPARE (FINAL FIX)
// ===============================
function processCompare() {
    const tbody = document.querySelector("#resultTable tbody");
    if (!tbody) return;

    tbody.innerHTML = "";
    resultData = [];

    dataB.forEach(function (b) {
        const match = dataA.find(a =>
            normalize(a[KEY_FIELD]) === normalize(b[KEY_FIELD])
        );

        let status = "";
        let detail = "";
        let cls = "";

        let ont = b["ONT"] || "";
        let macOnt = b["MAC ONT"] || "";
        let macStb = b["MAC STB"] || "";

        if (!match) {
            status = "NOT FOUND";
            detail = "Tidak ada di Excel A";
            cls = "notfound";
        } else {
            const ontMatch =
                normalize(match["ONT"]) === normalize(b["ONT"]);

            const macOntMatch =
                normalize(match["MAC ONT"]) === normalize(b["MAC ONT"]);

            const macStbMatch =
                normalize(match["MAC STB"]) === normalize(b["MAC STB"]);

            if (ontMatch && macOntMatch && macStbMatch) {
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
            [KEY_FIELD]: b[KEY_FIELD] || "",
            "ONT": ont,
            "MAC ONT": macOnt,
            "MAC STB": macStb,
            "Status": status,
            "Detail": detail
        });

        tbody.insertAdjacentHTML("beforeend", `
            <tr class="${cls}">
                <td>${b[KEY_FIELD] || ""}</td>
                <td>${ont}</td>
                <td>${macOnt}</td>
                <td>${macStb}</td>
                <td>${status}</td>
                <td>${detail}</td>
            </tr>
        `);
    });
}

// ===============================
// EXPORT EXCEL
// ===============================
function exportExcel() {
    if (!resultData.length) {
        alert("Tidak ada data untuk di-export");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(resultData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Compare Result");
    XLSX.writeFile(wb, "hasil_compare.xlsx");
}

// ===============================
// CLEAR
// ===============================
function clearAll() {
    const tbody = document.querySelector("#resultTable tbody");
    if (tbody) tbody.innerHTML = "";
    resultData = [];
}
