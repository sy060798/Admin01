// ===============================
// validate.js
// ===============================
let uploadedFiles = [];

// handle upload
document.getElementById('imageFiles').addEventListener('change', function() {
    const files = this.files;
    const tbody = document.querySelector("#imageList tbody");

    Array.from(files).forEach(file => {
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

    this.value = ""; // reset
});

// clear list
function clearAll() {
    document.querySelector("#imageList tbody").innerHTML = "";
    uploadedFiles = [];
}

// helper async untuk read file
function readFileAsDataURL(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = e => reject(e);
        reader.readAsDataURL(file);
    });
}

// download PDF
async function downloadPDF() {
    if (uploadedFiles.length === 0) {
        alert("Tidak ada gambar untuk di-download");
        return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    for (let i = 0; i < uploadedFiles.length; i++) {
        const file = uploadedFiles[i];
        const imgData = await readFileAsDataURL(file); // tunggu gambar selesai

        const imgProps = doc.getImageProperties(imgData);
        const pageWidth = doc.internal.pageSize.getWidth() - 20;
        const pageHeight = doc.internal.pageSize.getHeight() - 30;

        let pdfWidth = pageWidth;
        let pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
        if (pdfHeight > pageHeight) {
            pdfHeight = pageHeight;
            pdfWidth = (imgProps.width * pdfHeight) / imgProps.height;
        }

        if (i > 0) doc.addPage();
        doc.setFontSize(12);
        doc.text(file.name, 10, 10);
        doc.addImage(imgData, 'JPEG', 10, 20, pdfWidth, pdfHeight);
    }

    doc.save("uploaded_images.pdf");
}
