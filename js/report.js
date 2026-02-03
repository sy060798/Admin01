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

// ================= CORE PROCESS =================
function processRow(row) {
    const report = row["Report Installation"] || "";

    const getDate = d => d ? String(d).split(" ")[0] : "";
    const getTime = d => d ? String(d).split(" ")[1] : "";

    const getNumber = val => {
        if (!val) return 0;
        const m = String(val).match(/\d+/);
        return m ? parseInt(m[0]) : 0;
    };

    const extractReport = (regex) => {
        const m = report.match(regex);
        return m ? m[1].trim() : "";
    };

    const extractReportNumber = (regex) => {
        const m = report.match(regex);
        return m ? parseInt(m[1]) : 0;
    };

    reconData.push({
        "ALARM DATE START": getDate(row["Datetime Receive"]),
        "ALARM TIME START": getTime(row["Datetime Receive"]),
        "CITY": row["Cabang"] || "",
        "INSIDEN TICKET": row["No Wo Klien"] || "",
        "CIRCUIT ID": row["Cust ID Klien"] || "",
        "DESCRIPSI": report.match(/\[TSHOOT\][\s\S]*/)?.[0] || "",
        "ADDRESS": row["Alamat"] || "",
        "ALARM DATE CLEAR": getDate(row["Updated At"]),
        "ALARM TIME CLEAR": getTime(row["Updated At"]),

        // 🔴 RFO & ACTION DARI REPORT INSTALLATION
        "RFO": extractReport(/RFO\s*[:\-]?\s*([\s\S]*?)(?:\n|$)/i),
        "ACTION": extractReport(/ACTION\s*[:\-]?\s*([^\n]+)/i),

        "REPORTING": report,

        // 🟢 PRECON DARI EXCEL (PT MEGA AKSES)
        "PRECON 50": getNumber(row["Kabel Precon 50 Old"]),
        "PRECON 75": getNumber(row["Kabel Precon 75 Old"]),
        "PRECON 80": getNumber(row["Kabel Precon 80 Old"]),
        "PRECON 100": getNumber(row["Kabel Precon 100 Old"]),
        "PRECON 125": getNumber(row["Kabel Precon 125 Old"]),
        "PRECON 150": getNumber(row["Kabel Precon 150 Old"]),
        "PRECON 200": getNumber(row["Kabel Precon 200 Old"]),
        "PRECON 225": getNumber(row["Kabel Precon 225 Old"]),
        "PRECON 250": getNumber(row["Kabel Precon 250 Old"]),

        // 🟡 MATERIAL DARI REPORT
        "BAREL": extractReportNumber(/Barrel\s*[:\-]?\s*(\d+)/i),
        "PIGTAIL": extractReportNumber(/Pigtail\s*[:\-]?\s*(\d+)/i),
        "PATCHCORD": extractReportNumber(/Patchcord\s*[:\-]?\s*(\d+)/i),
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
