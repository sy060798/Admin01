console.log("cirbackbone.js loaded");

let resultData = [];

// ===============================
// ----- list -----
// field yang ditampilkan & export
// ===============================
const exportFields = [
    "site",
    "ticket/wo",
    "span problem",
    "root cause action"
];

// ===============================
// proses utama
// ===============================
function processCir() {
    resultData = [];
    const tbody = document.querySelector("#resultTable tbody");
    tbody.innerHTML = "";

    // ===== baca excel =====
    const file = document.getElementById("excelFile").files[0];
    if (file) {
        readExcel(file, rows => {
            rows.forEach(r => {
                pushRow(
                    r["Site"] || r["site"],
                    r["Ticket/WO"] || r["ticket/wo"],
                    r["Span Problem"] || r["span problem"],
                    r["Root Cause Action"] || r["root cause action"]
                );
            });
        });
    }

    // ===== baca text cir =====
    const text = document.getElementById("cirText").value;
    if (text.trim()) {
        const data = parseCirText(text);
        if (data) {
            pushRow(
                data.site,
                data.ticket,
                data.span,
                data.root
            );
        }
    }
}

// ===============================
// baca excel
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
// parse text cir manual
// ===============================
function parseCirText(text) {
    const get = label => {
        const m = text.match(new RegExp(label + "\\s*:\\s*(.+)", "i"));
        return m ? m[1].trim() : "";
    };

    return {
        site: get("TR INDOSAT"),
        ticket: get("TT FLP"),
        span: get("Span Problem"),
        root: get("Root Cause|Action")
    };
}

// ===============================
// push ke table & array
// ===============================
function pushRow(site, ticket, span, root) {
    const row = {
        "site": site || "",
        "ticket/wo": ticket || "",
        "span problem": span || "",
        "root cause action": root || ""
    };

    resultData.push(row);

    document.querySelector("#resultTable tbody")
        .insertAdjacentHTML("beforeend", `
            <tr>
                <td>${row["site"]}</td>
                <td>${row["ticket/wo"]}</td>
                <td>${row["span problem"]}</td>
                <td>${row["root cause action"]}</td>
            </tr>
        `);
}

// ===============================
// export excel
// ===============================
function exportExcel() {
    if (!resultData.length) {
        alert("tidak ada data");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(resultData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "cir backbone");
    XLSX.writeFile(wb, "cir_backbone.xlsx");
}

// ===============================
// clear
// ===============================
function clearAll() {
    resultData = [];
    document.querySelector("#resultTable tbody").innerHTML = "";
}

// expose
window.processCir = processCir;
window.exportExcel = exportExcel;
window.clearAll = clearAll;
