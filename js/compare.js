// ===============================
// CONFIG
// ===============================
console.log("compare.js loaded OK");

const KEY_FIELD = "Cust ID Klien";

let dataA = [];
let dataB = [];
let resultData = [];

// ===============================
// NORMALIZE DATA
// ===============================
function normalize(val) {
    return String(val || "")
        .replace(/\u00A0/g, "")
        .replace(/[\s\-:]/g, "")
        .toUpperCase()
        .trim();
}

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
        const match = dataA.find(a =>
            normalize(a[KEY_FIELD]) === normalize(b[KEY_FIELD])
        );

        let status, detail, cls;

        const ont = b["ONT"] || "";
        const macOnt = b["MAC ONT"] || "";
        const macStb = b["MAC STB"] || "";

        if (!match) {
            status = "NOT FOUND";
            detail = "Tidak ada di Excel A";
            cls = "notfound";
        } else {
            const ok =
                normalize(match["ONT"]) === normalize(ont) &&
                normalize(match["MAC ONT"]) === normalize(macOnt) &&
                normalize(match["MAC STB"]) === normalize(macStb);

            status = ok ? "MATCH" : "FAILED";
            detail = ok ? "Data cocok" : "Ada data tidak sesuai";
            cls = ok ? "match" : "failed";
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
// EXPORT
// ===============================
function exportExcel() {
    if (!resultData.length) {
        alert("Tidak ada data");
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
    document.querySelector("#resultTable tbody").innerHTML = "";
    resultData = [];
}

// ===============================
// 🔥 FIX PENTING UNTUK GITHUB
// ===============================
window.compareExcel = compareExcel;
window.exportExcel = exportExcel;
window.clearAll = clearAll;
