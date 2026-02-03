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
    const nextWO = excelData.slice(batchIndex, batchIndex + 5).map(t => t['No WO Klien']);
    if (nextWO.length === 0) return alert('Tidak ada WO Klien berikutnya!');
    displayTickets(nextWO);
}

// Tampilkan tiket di tabel
function displayTickets(woList) {
    const tbody = document.querySelector('#resultTable tbody');
    tbody.innerHTML = ''; // reset tabel

    woList.forEach(wo => {
        const rowData = excelData.filter(d => d['No WO Klien'] === wo);
        if (rowData.length === 0) return;

        // Ambil status terakhir berdasarkan Validate Date
        const lastStatusData = [...rowData].sort((a,b) => new Date(a['Validate Date']) - new Date(b['Validate Date'])).pop();

        // Ambil HOLD / UNHOLD maksimal 5
        const hold = [], unhold = [];
        let count = 0;
        for (let r of rowData) {
            if (count >= 5) break;
            // skip jika status Done/Cancel
            if (r['Status'] === 'Done' || r['Status'] === 'Cancel') continue;
            hold.push(r['HOLD DATE'] || '');
            unhold.push(r['UNHOLD DATE'] || '');
            count++;
        }

        // Tambahkan baris ke tabel
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
            <td>${lastStatusData['RFO'] || ''}</td>
            <td>${lastStatusData['Report'] || ''}</td>
            <td>${lastStatusData['Status'] || ''}</td>
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
