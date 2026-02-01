// ===============================
// validate.js
// ===============================
let uploadedFiles = [];

// ======= HANDLE UPLOAD =======
document.getElementById('imageFiles').addEventListener('change', function() {
    const files = this.files;
    const tbody = document.querySelector("#imageList tbody");

    Array.from(files).forEach((file) => {
        uploadedFiles.push(file);

        const reader = new FileReader();
        reader.onload = e => {
            tbody.insertAdjacentHTML("beforeend", `
                <tr>
                    <td>${tbody.children.length + 1}</td>
                    <td>${file.name}</td>
                    <td><img src="${e.target.result}" alt="${file.name}"></td>
                </tr>
            `);
        };
        reader.readAsDataURL(file);
    });

    this.value = ""; // reset supaya bisa upload file sama lagi
});

// ======= CLEAR =======
function clearAll() {
    document.querySelector("#imageList tbody").innerHTML = "";
    uploadedFiles = [];
}

// ======= DOWNLOAD PDF =======
async function downloadPDF() {
    if (uploadedFiles.length === 0) {
        alert("Tidak ada gambar untuk di-download");
        return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    for (let idx = 0; idx < uploadedFiles.length; idx++) {
        const file = uploadedFiles[idx];
        const imgData = await readFileAsDataURL(file);

        const imgProps = doc.getImageProperties(imgData);
        const pageWidth = doc.internal.pageSize.getWidth() - 20; // margin
        const pageHeight = doc.internal.pageSize.getHeight() - 30; // margin bawah untuk teks

        // hitung proporsi supaya gambar tidak pecah
        let pdfWidth = pageWidth;
        let pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
        if (pdfHeight > pageHeight) {
            pdfHeight = pageHeight;
            pdfWidth = (imgProps.width * pdfHeight) / imgProps.height;
        }

        if (idx > 0) doc.addPage();
        doc.setFontSize(12);
        doc.text(file.name, 10, 10); // nama asli file
        doc.addImage(imgData, 'JPEG', 10, 20, pdfWidth, pdfHeight);
    }

    doc.save("uploaded_images.pdf");
}

// helper untuk baca file jadi DataURL
function readFileAsDataURL(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = e => reject(e);
        reader.readAsDataURL(file);
    });
}
