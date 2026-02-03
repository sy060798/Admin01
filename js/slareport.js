let excelData = [];
let batchIndex = 0;
let currentBatchWO = [];

// Upload Excel
document.getElementById('upload').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return alert('File tidak ditemukan!');
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    excelData = XLSX.utils.sheet_to_json(worksheet);
    batchIndex = 0;
    alert(`Excel berhasil diupload! Total WO Klien: ${excelData.length}`);
});

// Generate batch max 5 WO
function generateBatch() {
    if (excelData.length === 0) return alert('Upload Excel dulu!');
    const input = document.getElementById('ticketInput').value;
    if (!input) return alert('Masukkan No WO Klien!');
    const woList = input.split(',').map(t => t.trim()).slice(0,5);
    currentBatchWO = woList;
    batchIndex = 0;
    displayTickets(woList);
}

// Next batch
function nextBatch() {
    if (excelData.length === 0) return alert('Upload Excel dulu!');
    batchIndex += 5;
    const nextWO = excelData.slice(batchIndex, batchIndex + 5).map(t => t['No Wo Klien']);
    if (nextWO.length === 0) return alert('Tidak ada WO Klien berikutnya!');
    currentBatchWO = nextWO;
    displayTickets(nextWO);
}

// Ambil RFO dari Pending/Reschedule terakhir, case-insensitive, huruf asli
function getRFO(reportText) {
    if (!reportText) return '';
    const keywords = ['Rsch','PENDING','Cancel','TEAM VISIT','NOTE:','Status:','REQ'];
    const cleanText = reportText.replace(/\*/g, '');
    const fragments = cleanText.split(/[,;\n]/).map(f => f.trim());
    for (let f of fragments) {
        if (keywords.some(kw => f.toUpperCase().includes(kw.toUpperCase()))) return f;
    }
    return '';
}

// Tampilkan tiket horizontal per batch
function displayTickets(woList) {
    const container = document.getElementById('resultContainer');
    container.innerHTML = ''; // reset container

    const table = document.createElement('table');
    table.style.borderCollapse = "collapse";
    table.style.width = "100%";

    // Header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    ['No WO Klien','HOLD DATE','UNHOLD DATE','RFO','Report','Status'].forEach(h => {
        const th = document.createElement('th');
        th.innerText = h;
        th.style.border = "1px solid #333";
        th.style.padding = "5px";
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    woList.forEach(wo => {
        const rowData = excelData
            .filter(d => (d['No Wo Klien'] || '').trim().toLowerCase() === wo.toLowerCase())
            .sort((a,b) => new Date(a['Validate Date']) - new Date(b['Validate Date']));

        if (rowData.length === 0) return;

        // HOLD / UNHOLD logic
        const holdList = [];
        const unholdList = [];

        rowData.forEach((r, idx) => {
            if (r['Status'] === 'Pending' || r['Status'] === 'Reschedule') {
                holdList.push(r['Validate Date'] || '');
                // UNHOLD = Validate Date dari baris pertama setelah Pending/Resch sampai Done/Cancel/BTN
                let unholdDate = '';
                for (let j = idx+1; j<rowData.length; j++){
                    if(['Done','Cancel','BTN'].includes(rowData[j]['Status'])){
                        unholdDate = rowData[j]['Validate Date'] || '';
                        break;
                    }
                }
                if(!unholdDate) unholdDate = rowData[rowData.length-1]['Validate Date'] || '';
                unholdList.push(unholdDate);
            }
        });

        // Pending/Resch terakhir untuk RFO & Report
        const lastPending = [...rowData]
            .filter(r => r['Status']==='Pending' || r['Status']==='Reschedule')
            .sort((a,b) => new Date(b['Validate Date']) - new Date(a['Validate Date']))[0];

        // Status terakhir
        const latestRow = [...rowData].sort((a,b) => new Date(b['Validate Date']) - new Date(a['Validate Date']))[0];

        const tr = document.createElement('tr');

        const holdStr = holdList.join(', ');
        const unholdStr = unholdList.join(', ');

        [wo, holdStr, unholdStr,
         lastPending ? getRFO(lastPending['report']) : '',
         lastPending ? lastPending['Report Installation'] || '' : '',
         latestRow['Status'] || ''].forEach(val => {
            const td = document.createElement('td');
            td.innerText = val;
            td.style.border = "1px solid #333";
            td.style.padding = "5px";
            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    });

    table.appendChild(tbody);

    // Judul List
    const title = document.createElement('h4');
    title.innerText = `List ${batchIndex/5 + 1}`;
    container.appendChild(title);
    container.appendChild(table);
}

// Export Excel
function exportExcel() {
    const container = document.getElementById('resultContainer');
    if (!container) return alert('Tidak ada data untuk export!');
    const tables = container.querySelectorAll('table');
    if (tables.length === 0) return alert('Tidak ada data untuk export!');
    const wb = XLSX.utils.book_new();
    tables.forEach((tbl, idx)=>{
        const ws = XLSX.utils.table_to_sheet(tbl);
        XLSX.utils.book_append_sheet(wb, ws, `List ${idx+1}`);
    });
    XLSX.writeFile(wb, "SLA_Report.xlsx");
}
