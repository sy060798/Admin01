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

        date_info.setHours(
            Math.floor(total_seconds / 3600),
            Math.floor((total_seconds % 3600) / 60),
            total_seconds % 60
        );
        return date_info;
    }

    if (typeof value === "string") {
        const clean = value.replace(/\u00A0/g, " ").trim();
        const m = clean.match(/(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2})/);
        if (!m) return "";
        const [, y, mo, d, h, mi, s] = m.map(Number);
        return new Date(y, mo - 1, d, h, mi, s);
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

function isNewer(a, b) {
    if (!a) return false;
    if (!b) return true;
    return a.getTime() > b.getTime();
}

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
            delete r.__DATE_OBJ;
            return r;
        });

        renderTable();
    };
    reader.readAsArrayBuffer(e.target.files[0]);
}

// ================= CORE PROCESS =================
function processRow(row) {
    let report = (row["Report Installation"] || "").replace(/\*/g, "");

    const getNumber = v => {
        const m = String(v || "").match(/\d+/);
        return m ? parseInt(m[0]) : 0;
    };

    const extractText = r => {
        const m = report.match(r);
        return m ? m[1].trim() : "";
    };

    const extractNumber = r => {
        const m = report.match(r);
        return m ? parseInt(m[1]) : 0;
    };

    const getDescription = () => {
        const m = report.match(
            /(TSHOOT|REQUEST)[\s\S]*?(?=\n\s*\n|RFO|ACTION|CANCEL|PIC|TEAM|$)/i
        );
        return m ? m[0].replace(/\s+/g, " ").trim() : "";
    };

    let rfoText = extractText(/RFO\s*[:\-]?\s*([\s\S]*?)(?=\n\s*\n|ACTION|$)/i);
    if (/CANCEL/i.test(report)) {
        const c = report.match(/CANCEL[^\n]*/i);
        if (c) rfoText = (rfoText ? rfoText + " | " : "") + c[0];
    }

    const rawStatus = (row["Status"] || "").toUpperCase();
    let status = "";

    if (rawStatus.includes("DONE")) status = "DONE";
    else if (rawStatus.includes("BTN")) status = "BTN";
    else return;

    const wo = row["No Wo Klien"];
    const dtReceive = excelDateToJSDate(row["Datetime Receive"]);

    // 🔥 CLEAR DATETIME DIGABUNG
    const clearDate = getDate(row["Updated At"]);
    const clearTime = getTime(row["Updated At"]);
    const clearDateTime = clearDate && clearTime ? `${clearDate} ${clearTime}` : "";

    const newData = {
        "ALARM DATE START": getDate(row["Datetime Receive"]),
        "ALARM TIME START": getTime(row["Datetime Receive"]),
        "CITY": row["Cabang"] || "",
        "INSIDEN TICKET": wo,
        "CIRCUIT ID": row["Cust ID Klien"] || "",
        "DESCRIPSI": getDescription(),
        "ADDRESS": row["Alamat"] || "",

        // ✅ DIGABUNG
        "ALARM DATE CLEAR": clearDateTime,
        "ALARM TIME CLEAR": "",

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

        "STATUS": status,
        "__DATE_OBJ": dtReceive
    };

    if (!reconMap[wo]) {
        reconMap[wo] = newData;
        return;
    }

    const old = reconMap[wo];

    if (old.STATUS !== "DONE" && status === "DONE") {
        reconMap[wo] = newData;
        return;
    }

    if (old.STATUS === "DONE" && status === "DONE") {
        if (isNewer(dtReceive, old.__DATE_OBJ)) reconMap[wo] = newData;
        return;
    }

    if (old.STATUS === "BTN" && status === "BTN") {
        if (isNewer(dtReceive, old.__DATE_OBJ)) reconMap[wo] = newData;
    }
}

// ================= RENDER =================
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

// ================= EXPORT =================
function exportExcel() {
    if (!reconData.length) return alert("Data masih kosong");

    const ws = XLSX.utils.json_to_sheet(reconData, {
        header: Object.keys(reconData[0])
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "RECONCILE");
    XLSX.writeFile(wb, "RECON_MEGA_AKSES.xlsx");
}
