let reconData = [];

document.getElementById("upload").addEventListener("change", handleFile);

function handleFile(e) {
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);

    rows.forEach(row => processRow(row));
    renderTable();
  };
  reader.readAsArrayBuffer(e.target.files[0]);
}

function processRow(row) {
  const report = row["Report Installation"] || "";

  const getDate = d => d ? d.split(" ")[0] : "";
  const getTime = d => d ? d.split(" ")[1] : "";

  const extract = (regex) => {
    const m = report.match(regex);
    return m ? m[1] : 0;
  };

  reconData.push({
    alarmDateStart: getDate(row["Datetime Receive"]),
    alarmTimeStart: getTime(row["Datetime Receive"]),
    city: row["Cabang"],
    ticket: row["No Wo Klien"],
    circuit: row["Cust ID Klien"],
    desc: report.match(/\[TSHOOT\][\s\S]*?\s{2,}/)?.[0] || "",
    address: row["Alamat"],
    alarmDateClear: getDate(row["Updated At"]),
    alarmTimeClear: getTime(row["Updated At"]),
    rfo: row["RFO"],
    action: (row["ACTION"] || "").split("\n")[0],
    reporting: report,

    p50: extract(/Precon 50.*?(\d+)/i),
    p75: extract(/Precon 75.*?(\d+)/i),
    p80: extract(/Precon 80.*?(\d+)/i),
    p100: extract(/Precon 100.*?(\d+)/i),
    p150: extract(/Precon 150.*?(\d+)/i),
    p200: extract(/Precon 200.*?(\d+)/i),
    p225: extract(/Precon 225.*?(\d+)/i),
    p250: extract(/Precon 250.*?(\d+)/i),
    barel: extract(/Barrel.*?(\d+)/i),
    pigtail: extract(/Pigtail.*?(\d+)/i),
    patchcord: extract(/Patchcord.*?(\d+)/i),
  });
}

function renderTable() {
  const tbody = document.querySelector("#resultTable tbody");
  tbody.innerHTML = "";

  reconData.forEach(d => {
    const tr = document.createElement("tr");
    Object.values(d).forEach(val => {
      const td = document.createElement("td");
      td.textContent = val;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

function exportExcel() {
  const ws = XLSX.utils.json_to_sheet(reconData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "RECON");
  XLSX.writeFile(wb, "RECON_MEGA_AKSES.xlsx");
}
