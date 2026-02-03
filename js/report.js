let reconMap = {};
let reconData = [];

document.getElementById("upload").addEventListener("change", handleFile);

// ================== EXCEL DATE FIX ==================
function excelDateToJSDate(value) {
    if (!value) return "";

    if (typeof value === "number") {
        const utc_days = Math.floor(value - 25569);
        const utc_value = utc_days * 86400;
        const date_info = new Date(utc_value * 1000);

        const fractional_day = value - Math.floor(value);
        const total_seconds = Math.floor(86400 * fractional_day);

        const seconds = total_seconds % 60;
        const total_minutes = Math.floor(total_seconds / 60);
        const minutes = total_minutes % 60;
        const hours = Math.floor(total_minutes / 60);

        date_info.setHours(hours, minutes, seconds);
        return date_info;
    }

    if (typeof value === "string") {
        const m = value.match(
            /(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2})/
        );
        if (!m) return "";
        return new Date(m[1], m[2] - 1, m[3], m[4], m[5], m[6]);
    }

    return "";
}

const getDate = d => {
    const dt = excelDateToJSDate(d);
    return dt ? dt.toISOString().slice(0, 10) : "";
};

const getTime = d => {
    const dt = excelDateToJSDate(d);
    return dt ? dt.toTimeString().slice(0, 8) : "";
};

// ================= FILE HANDLER =================
function handleFile(e) {
    reconMap = {};
    reconData = [];

    const reader = new FileReader();
    reader.onload = function (evt) {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        rows.forEach(row => processRow(row));

        reconData = Object.values(reconMap).map(r => {
            delete r.__META;
            return r;
        });

        renderTable();
    };
    reader.readAsArrayBuffer(e.target.files[0]);
}

// ================= CORE PROCESS =================
function processRow(row) {
    let statusRaw = (row["Status"] || "").toUpperCase();
    let status = "";

    if (statusRaw.includes("DONE")) status = "DONE";
    else if (statusRaw.includes("CANCEL")) status = "CANCEL BTN";
    else return; // ⛔ status lain DIABAIKAN

    const wo = row["No Wo Klien"];
    if (!wo) return;

    let report = (row["Report Installation"] || "").replace(/\*/g, "");

    const extractText = regex => {
        const m = report.match(regex);
        return m ? m[1].trim() : "";
    };

    let rfoText = extractText(/RFO\s*[:\-]?\s*([\s\S]*?)(?=\n\s*\n|ACTION|$)/i);
    if (status === "CANCEL BTN") {
        const c = report.match(/CANCEL[^\n]*/i);
        if (c) rfoText = (rfoText ? rfoText + " | " : "") + c[0];
    }

    const recvDate = excelDateToJSDate(row["Datetime Receive"]);

    const newRow = {
        "ALARM DATE START": getDate(row["Datetime Receive"]),
        "ALARM TIME START": getTime(row["Datetime Receive"]),
        "CITY": row["Cabang"] || "",
        "INSIDEN TICKET": wo,
        "CIRCUIT ID": row["Cust ID Klien"] || "",
        "DESCRIPSI": extractText(
            /(TSHOOT|REQUEST)[\s\S]*?(?=\n\s*\n|RFO|ACTION|CANCEL|$)/i
        ),
        "ADDRESS": row["Alamat"] || "",
        "ALARM DATE CLEAR": getDate(row["Updated At"]),
        "ALARM TIME CLEAR": getTime(row["Updated At"]),
        "RFO": rfoText,
        "ACTION": extractText(/ACTION\s*[:\-]?\s*([\s\S]*?)(?=\n\s*\n|$)/i),
        "REPORTING": report,
        "STATUS": status,
        "__META": {
            status,
            recvDate
        }
    };

    if (!reconMap[wo]) {
        reconMap[wo] = newRow;
        return;
    }

    const old = reconMap[wo].__META;

    // ===== PRIORITAS STATUS =====
    if (old.status === "CANCEL BTN" && status === "DONE") {
        reconMap[wo] = newRow;
        return;
    }

    if (old.status === status) {
        if (recvDate > old.recvDate) {
            reconMap[wo] = newRow;
        }
    }
}

// ================= RENDER =================
function renderTable() {
    const tbody = document.querySelector("#resultTable tbody");
    tbody.innerHTML = "";

    reconData.forEach(row => {
        const tr = document.createElement("tr");
        Object.values(row).forEach(v => {
            const td = document.createElement("td");
            td.textContent = v;
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
}

// ================= EXPORT =================
function exportExcel() {
    if (!reconData.length) {
        alert("Data masih kosong");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(reconData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "RECONCILE");
    XLSX.writeFile(wb, "RECON_MEGA_AKSES.xlsx");
}
