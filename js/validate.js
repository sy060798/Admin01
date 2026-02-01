// ===============================
// VALIDATE.JS REVISI
// ===============================
console.log("validate.js loaded");

// array untuk simpan file asli
let uploadedFiles = [];

// ======= HANDLE UPLOAD =======
function handleFiles(input) {
    const files = input.files;
    const tbody = document.querySelector("#imageList tbody");

    Array.from(files).forEach((file) => {
        // simpan file asli di array
        uploadedFiles.push(file);

        const reader = new FileReader();
        reader.onload = e => {
            // tampilkan di tabel
            tbody.insertAdjacentHTML("beforeend", `
                <tr>
                    <td>${tbody.children.length + 1}</td>
                    <td>${file.name}</td> <!-- pakai nama asli file -->
                    <td>
                        <img src="${e.target.result}" 
                             alt="${file.name}" 
                             style="max-width:200px; max-height:150px; object-fit:contain;"/>
                    </td>
                </tr>
            `);
        };
        reader.readAsDataURL(file);
    });

    // reset input supaya bisa upload file sama lagi
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

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    // loop untuk setiap file
    uploadedFiles.forEach((file, idx) => {
        const reader = new FileReader();
        reader.onload = e => {
            const imgData = e.target.result;

            const imgProps = doc.getImageProperties(imgData);
            const pdfWidth = doc.internal.pageSize.getWidth() - 20; // margin 10
            const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;

            if (idx > 0) doc.addPage();
            // tulis nama asli file di atas gambar
            doc.setFontSize(12);
            doc.text(file.name, 10, 10);
            doc.addImage(imgData, 'JPEG', 10, 20, pdfWidth, pdfHeight);

            // simpan PDF kalau file terakhir
            if (idx === uploadedFiles.length - 1) {
                doc.save("uploaded_images.pdf");
            }
        };
        reader.readAsDataURL(file);
    });
}
