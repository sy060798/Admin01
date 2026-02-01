// ===============================
// VALIDATE.JS
// ===============================
console.log("validate.js loaded");

// array untuk simpan file
let uploadedFiles = [];

// ======= HANDLE UPLOAD =======
function handleFiles(input) {
    const files = input.files;
    const tbody = document.querySelector("#imageList tbody");

    Array.from(files).forEach((file, index) => {
        // simpan file di array
        uploadedFiles.push(file);

        const reader = new FileReader();
        reader.onload = e => {
            tbody.insertAdjacentHTML("beforeend", `
                <tr>
                    <td>${tbody.children.length + 1}</td>
                    <td>${file.name}</td>
                    <td><img src="${e.target.result}" alt="${file.name}" style="max-width:200px; max-height:150px; object-fit:contain;"/></td>
                </tr>
            `);
        };
        reader.readAsDataURL(file);
    });

    // reset supaya bisa upload file sama lagi
    input.value = "";
}

// ======= CLEAR LIST =======
function clearList() {
    document.querySelector("#imageList tbody").innerHTML = "";
    uploadedFiles = [];
}

// ======= DOWNLOAD PDF =======
function downloadPDF() {
    if (uploadedFiles.length === 0) {
        alert("Tidak ada gambar untuk di-download");
        return;
    }

    const doc = new jsPDF();

    uploadedFiles.forEach((file, idx) => {
        const reader = new FileReader();
        reader.onload = e => {
            const imgData = e.target.result;
            // fit image ke halaman
            const imgProps = doc.getImageProperties(imgData);
            const pdfWidth = doc.internal.pageSize.getWidth() - 20; // margin 10
            const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;

            if (idx > 0) doc.addPage();
            doc.text(file.name, 10, 10);
            doc.addImage(imgData, 'JPEG', 10, 20, pdfWidth, pdfHeight);

            // kalau sudah file terakhir, simpan PDF
            if (idx === uploadedFiles.length - 1) {
                doc.save("uploaded_images.pdf");
            }
        };
        reader.readAsDataURL(file);
    });
}
