let reconData = [];

document.getElementById("upload").addEventListener("change", handleFile);

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

// ================= CORE =================
function processRow(row) {
    const report = row["Report Installation"] || "";

    const getDate = d => d ? String(d).split(" ")[0] : "";
    const getTime = d => d ? String(d).split(" ")[1] : "";

    const getNumber = v => {
        const m = String(v).match(/\d+/);
        return m ? parseInt(m[0]) : 0;
    };

    // ================= DESCRIPSI =================
    const descMatch = report.match(
        /(\[?\s*(REQUEST|TSHOOT)\s*\]?[\s\S]*?)(?=\n\s*(CANCEL|RFO|TEAM|PIC)|_{3,}|\n\s*\n)/i
    );
    const descripsi = descMatch
        ? descMatch[1].replace(/\s+/g, " ").trim()
        : "";

    // ================= CANCEL DETECT =================
    const cancelMatch = report.match(/^(CANCEL.*)$/im);
    const cancelText = cancelMatch ? cancelMatch[1].trim() : "";

    // ================= RFO =================
    let rfoText = "";
    if (cancelText) {
        rfoText = cancelText;
    } else {
        const rfoMatch = report.match(/RFO\s*[:\-]?\s*([\s\S]*?)(?:\n|$)/i);
        rfoText = rfoMatch ? rfoMatch[1].trim() : "";
    }

    // ================= STATUS =================
    let statusFinal = (row["Status"] || "").toString().toUpperCase();
    if (cancelText) {
        statusFinal = "CANCEL";
    } else if (!statusFinal) {
        statusFinal = "RESCHEDULE";
    }

    reconData.push({
        "ALARM DATE START": getDate(row["Datetime Receive"]),
        "ALARM TIME START": getTime(row["Datetime Receive"]),
        "CITY": row["Cabang"] || "",
        "INSIDEN TICKET": row["No Wo Klien"] || "",
        "CIRCUIT ID": row["Cust ID Klien"] || "",

        "DESCRIPSI": descripsi,

        "ADDRESS": row["Alamat"] || "",
        "ALARM DATE CLEAR": getDate(row["Updated At"]),
        "ALARM TIME CLEAR": getTime(row["Updated At"]),

        "RFO": rfoText,
        "ACTION": (() => {
            const m = report.match(/ACTION\s*[:\-]?\s*([^\n]+)/i);
            return m ? m[1].trim() : "";
        })(),

        "REPORTING": report,

        // PRECON
        "PRECON 50": getNumber(row["Kabel Precon 50 Old"]),
        "PRECON 75": getNumber(row["Kabel Precon 75 Old"]),
        "PRECON 80": getNumber(row["Kabel Precon 80 Old"]),
        "PRECON 100": getNumber(row["Kabel Precon 100 Old"]),
        "PRECON 125": getNumber(row["Kabel Precon 125 Old"]),
        "PRECON 150": getNumber(row["Kabel Precon 150 Old"]),
        "PRECON 200": getNumber(row["Kabel Precon 200 Old"]),
        "PRECON 225": getNumber(row["Kabel Precon 225 Old"]),
        "PRECON 250": getNumber(row["Kabel Precon 250 Old"]),

        // MATERIAL
        "BAREL": (() => {
            const m = report.match(/Barrel\s*[:\-]?\s*(\d+)/i);
            return m ? parseInt(m[1]) : 0;
        })(),
        "PIGTAIL": (() => {
            const m = report.match(/Pigtail\s*[:\-]?\s*(\d+)/i);
            return m ? parseInt(m[1]) : 0;
        })(),
        "PATCHCORD": (() => {
            const m = report.match(/Patchcord\s*[:\-]?\s*(\d+)/i);
            return m ? parseInt(m[1]) : 0;
        })(),

        "STATUS": statusFinal
    });
}

// ================= TABLE =================
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
