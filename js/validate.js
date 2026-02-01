// ===============================
// VALIDATE.JS
// ===============================
console.log("validate.js loaded OK");

let imageFiles = [];

// ===============================
// HANDLE FILE UPLOAD
// ===============================
document.getElementById("imageFiles").addEventListener("change", function(e){
    const files = Array.from(e.target.files);
    files.forEach(file => {
        imageFiles.push(file);
    });
    renderList();
});

// ===============================
// RENDER LIST MENURUN
// ===============================
function renderList() {
    const tbody = document.querySelector("#imageList tbody");
    tbody.innerHTML = "";
    imageFiles.forEach((file, index) => {
        const reader = new FileReader();
        reader.onload = function(e){
            tbody.insertAdjacentHTML("beforeend", `
                <tr>
                    <td>${index + 1}</td>
                    <td>${file.name}</td>
                    <td><img src="${e.target.result}" alt="${file.name}"></td>
                </tr>
            `);
        };
        reader.readAsDataURL(file);
    });
}

// ===============================
// GENERATE PDF
// ===============================
function generatePDF() {
    if(!imageFiles.length){
        alert("Upload gambar dulu!");
        return;
    }

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF();

    let count = 0;

    function addImageToPDF(index){
        if(index >= imageFiles.length){
            pdf.save("images.pdf");
            return;
        }

        const reader = new FileReader();
        reader.onload = function(e){
            const imgData = e.target.result;
            const img = new Image();
            img.src = imgData;
            img.onload = function(){
                const pdfWidth = pdf.internal.pageSize.getWidth();
                const pdfHeight = pdf.internal.pageSize.getHeight();
                let ratio = Math.min(pdfWidth / img.width, pdfHeight / img.height);
                let width = img.width * ratio;
                let height = img.height * ratio;
                let x = (pdfWidth - width)/2;
                let y = (pdfHeight - height)/2;

                pdf.addImage(img, 'JPEG', x, y, width, height);

                if(index < imageFiles.length -1) pdf.addPage();
                addImageToPDF(index + 1);
            }
        };
        reader.readAsDataURL(imageFiles[index]);
    }

    addImageToPDF(0);
}

// ===============================
// CLEAR ALL
// ===============================
function clearAll(){
    imageFiles = [];
    document.querySelector("#imageList tbody").innerHTML = "";
    document.getElementById("imageFiles").value = "";
}
