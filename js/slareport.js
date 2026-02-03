let excelData = [];
let batchIndex = 0;

// Upload Excel
document.getElementById('upload').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    excelData = XLSX.utils.sheet_to_json(worksheet);
    batchIndex = 0;
    alert('Excel berhasil diupload! Total WO Klien: ' + excelData.length);
});

// Generate batch 5 WO Klien
function generateBatch() {
    const input = document.getElementById('ticketInput').value;
    if (!input) return alert('Masukkan No WO Klien terlebih dahulu!');
    const woList = input.split(',').map(t => t.trim()).slice(0,5); // max 5
    displayTickets(woList);
}

// Next batch
function nextBatch() {
    batchIndex += 5;
    const nextWO = excelData.slice(batchIndex, batchIndex + 5).map(t => t['No Wo Klien']);
    if (nextWO.length === 0) return alert('Tidak ada WO Klien berikutnya!');
    displayTickets(nextWO);
}

// Ambil RFO dari report: hapus *, ambil fragment pertama yang cocok keyword
function getRFO(reportText) {
    if (!reportText) return '';
    const keywords = ['Rsch','PENDING','Cancel','TEAM VISIT','NOTE:','Status:','REQ'];

    // Hapus tanda *
    const cleanText = reportText.replace(/\*/g, '');

    // Pecah report menjadi fragment berdasarkan koma atau titik koma
    const fragments = cleanText.split(/[,;]/).map(f => f.trim());

    // Ambil **fragment pertama** yang mengandung keyword (case-insensitive)
    for (let f of fragments) {
        if (keywords.some(kw => f.toUpperCase().includes(kw.toUpperCase()))) {
            return f; // ambil hanya satu fragment
        }
    }

    return '';
}

// Tampilkan tiket di tabel
function displayTickets(woList) {
    const tbody = document.querySelector('#resultTable tbody');
    tbody.innerHTML = ''; // reset tabel

    woList.forEach(wo => {
        const rowData = excelData
            .filter(d => (d['No Wo Klien'] || '').trim().toLowerCase() === wo.trim().toLowerCase())
            .sort((a,b) => new Date(a['Validate Date']) - new Date(b['Validate Date'])); // lama → baru

        if (rowData.length === 0) return;

        // Ambil HOLD / UNHOLD maksimal 5 dari Pending / Reschedule
        const hold = [], unhold = [];
        let count = 0;
        for (let r of rowData) {
            if (count >= 5) break;
            if (r['Status'] !== 'Pending' && r['Status'] !== 'Reschedule') continue; 
            hold.push(r['Validate Date'] || '');
            unhold.push(r['Validate Date'] || '');
            count++;
        }

        // Ambil baris Pending/Reschedule terakhir untuk RFO & Report
        const lastPending = [...rowData]
            .filter(r => r['Status'] === 'Pending' || r['Status'] === 'Reschedule')
            .sort((a,b) => new Date(b['Validate Date']) - new Date(a['Validate Date']))[0];

        // Ambil status terakhir dari baris terbaru (Validate Date terbaru)
        const latestRow = [...rowData].sort((a,b) => new Date(b['Validate Date']) - new Date(a['Validate Date']))[0];

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${wo}</td>
            <td>${hold[0]||''}</td>
            <td>${unhold[0]||''}</td>
            <td>${hold[1]||''}</td>
            <td>${unhold[1]||''}</td>
            <td>${hold[2]||''}</td>
            <td>${unhold[2]||''}</td>
            <td>${hold[3]||''}</td>
            <td>${unhold[3]||''}</td>
            <td>${hold[4]||''}</td>
            <td>${unhold[4]||''}</td>
            <td>${lastPending ? getRFO(lastPending['report']) : ''}</td>
            <td>${lastPending ? lastPending['Report Installation'] || '' : ''}</td>
            <td>${latestRow['Status'] || ''}</td>
        `;
        tbody.appendChild(tr);
    });
}

// Export tabel ke Excel
function exportExcel() {
    const table = document.getElementById('resultTable');
    const wb = XLSX.utils.table_to_book(table, {sheet: "SLA Report"});
    XLSX.writeFile(wb, "SLA_Report.xlsx");
}
