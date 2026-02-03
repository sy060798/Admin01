let reconData = [];

document.getElementById("upload").addEventListener("change", handleFile);

// ================== EXCEL DATE FIX (AKTUAL, TANPA FALLBACK) ==================
function excelDateToJSDate(value) {
    if (!value) return "";

    // === KASUS 1: Excel date serial (number) ===
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

    // === KASUS 2: STRING (dibersihkan & diparse manual) ===
    if (typeof value === "string") {
        const clean = value
            .replace(/\u00A0/g, " ") // NBSP
            .replace(/\s+/g, " ")
            .trim();

        // format: yyyy-mm-dd hh:mm:ss
        const m = clean.match(
            /(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2})/
        );

        if (!m) return "";

        const [, y, mo, d, h, mi, s] = m.map(Number);
        return new Date(y, mo - 1, d, h, mi, s);
    }

    return "";
}

const getDate = d => {
    const dt = excelDateToJSDate(d);
    if (!dt || isNaN(dt)) return "";
    return dt.toISOString().slice(0, 10);
};

const getTime = d => {
    const dt = excelDateToJSDate(d);
    if (!dt || isNaN(dt)) return "";
    return dt.toTimeString().slice(0, 8);
};

// ================= FILE HANDLER =================
function handleFile(e) {
    reconData = [];

    const reader = new FileReader();
    reader.onload = function (evt) {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        rows.forEach(row => processRow(row));
        renderTable();
    };
    reader.readAsArrayBuffer(e.target.files[0]);
}

// ================= CORE PROCESS =================
function processRow(row) {
    let report = row["Report Installation"] || "";
    report = report.replace(/\*/g, ""); // abaikan tanda *

    const getNumber = val => {
        if (!val) return 0;
        const m = String(val).match(/\d+/);
        return m ? parseInt(m[0]) : 0;
    };

    const extractText = regex => {
        const m = report.match(regex);
        return m ? m[1].trim() : "";
    };

    const extractNumber = regex => {
        const m = report.match(regex);
        return m ? parseInt(m[1]) : 0;
    };

    // ================= DESCRIPSI =================
    const getDescription = () => {
        const m = report.match(
            /(TSHOOT|REQUEST)[\s\S]*?(?=\n\s*\n|RFO|ACTION|CANCEL|PIC|TEAM|$)/i
        );
        return m ? m[0].replace(/\s+/g, " ").trim() : "";
    };

    // ================= RFO + CANCEL =================
    let rfoText = extractText(/RFO\s*[:\-]?\s*([\s\S]*?)(?=\n\s*\n|ACTION|$)/i);

    if (/CANCEL/i.test(report)) {
        const cancelLine = report.match(/CANCEL[^\n]*/i);
        if (cancelLine) {
            rfoText = (rfoText ? rfoText + " | " : "") + cancelLine[0].trim();
        }
    }

    // ================= STATUS =================
    let status = (row["Status"] || "").toUpperCase();
    if (status.includes("RESCHEDULE")) status = "RESCHEDULE";
    else if (status.includes("CANCEL")) status = "CANCEL";
    else if (status.includes("DONE")) status = "DONE";

    reconData.push({
        "ALARM DATE START": getDate(row["Datetime Receive"]),
        "ALARM TIME START": getTime(row["Datetime Receive"]),
        "CITY": row["Cabang"] || "",
        "INSIDEN TICKET": row["No Wo Klien"] || "",
        "CIRCUIT ID": row["Cust ID Klien"] || "",
        "DESCRIPSI": getDescription(),
        "ADDRESS": row["Alamat"] || "",
        "ALARM DATE CLEAR": getDate(row["Updated At"]),
        "ALARM TIME CLEAR": getTime(row["Updated At"]),
        "RFO": rfoText,
        "ACTION": extractText(/ACTION\s*[:\-]?\s*([\s\S]*?)(?=\n\s*\n|$)/i),
        "REPORTING": report,

        "PRECON 50": getNumber(row["Kabel Precon 50 Old"]),
        "PRECON 75": getNumber(row["Kabel Precon 75 Old"]),
        "PRECON 80": getNumber(row["Kabel Precon 80 Old"]),
        "PRECON 100": getNumber(row["Kabel Precon 100 Old"]),
        "PRECON 125": getNumber(row["Kabel Precon 125 Old"]),
        "PRECON 150": getNumber(row["Kabel Precon 150 Old"]),
        "PRECON 200": getNumber(row["Kabel Precon 200 Old"]),
        "PRECON 225": getNumber(row["Kabel Precon 225 Old"]),
        "PRECON 250": getNumber(row["Kabel Precon 250 Old"]),

        "BAREL": extractNumber(/Barrel\s*[:\-]?\s*(\d+)/i),
        "PIGTAIL": extractNumber(/Pigtail\s*[:\-]?\s*(\d+)/i),
        "PATCHCORD": extractNumber(/Patchcord\s*[:\-]?\s*(\d+)/i),

        "STATUS": status
    });
}

// ================= RENDER TABLE =================
function renderTable() {
    const tbody = document.querySelector("#resultTable tbody");
    tbody.innerHTML = "";

    reconData.forEach(row => {
        const tr = document.createElement("tr");
        Object.values(row).forEach(val => {
            const td = document.createElement("td");
            td.textContent = val;
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
}

// ================= EXPORT EXCEL =================
function exportExcel() {
    if (reconData.length === 0) {
        alert("Data masih kosong");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(reconData, {
        header: Object.keys(reconData[0])
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "RECONCILE");
    XLSX.writeFile(wb, "RECON_MEGA_AKSES.xlsx");
}
