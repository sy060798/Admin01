// ===============================
// Validate.js
// ===============================
console.log("validate.js loaded OK");

let imageFiles = [];

// ===============================
// UPLOAD & RENAME
// ===============================
document.getElementById("imageFiles").addEventListener("change", function(e){
    const files = Array.from(e.target.files);
    files.forEach((file, index) => {
        const newName = `Zahra_${imageFiles.length + 1}${file.name.slice(file.name.lastIndexOf('.'))}`;
        imageFiles.push({
            file: file,
            name: newName
        });
    });
    renderList();
});

// ===============================
// RENDER LIST MENURUN
// ===============================
function renderList(){
    const tbody = document.querySelector("#imageList tbody");
    tbody.innerHTML = "";
    imageFiles.forEach((item, index) => {
        const reader = new FileReader();
        reader.onload = function(e){
            tbody.insertAdjacentHTML("beforeend", `
                <tr>
                    <td>${index + 1}</td>
                    <td>${item.name}</td>
                    <td><img src="${e.target.result}" alt="${item.name}"></td>
                </tr>
            `);
        };
        reader.readAsDataURL(item.file);
    });
}

// ===============================
// GENERATE PDF
// ===============================
function generatePDF(){
    if(!imageFiles.length){ alert("Upload gambar dulu!"); return; }
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({orientation:'portrait', unit:'mm', format:'a4'});

    function addImageToPDF(i){
        if(i >= imageFiles.length){
            pdf.save("images.pdf");
            return;
        }
        const reader = new FileReader();
        reader.onload = function(e){
            const imgData = e.target.result;
            const img = new Image();
            img.src = imgData;
            img.onload = function(){
                const pdfW = pdf.internal.pageSize.getWidth();
                const pdfH = pdf.internal.pageSize.getHeight();
                let ratio = Math.min(pdfW/img.width, pdfH/img.height) * 0.9;
                let width = img.width * ratio;
                let height = img.height * ratio;
                let x = (pdfW - width)/2;
                let y = (pdfH - height)/2 + 10; // beri jarak nama

                pdf.setFontSize(12);
                pdf.text(imageFiles[i].name, pdfW/2, 10, {align:'center'});

                pdf.addImage(img,'JPEG',x,y,width,height);

                if(i < imageFiles.length -1) pdf.addPage();
                addImageToPDF(i+1);
            }
        };
        reader.readAsDataURL(imageFiles[i].file);
    }

    addImageToPDF(0);
}

// ===============================
// CLEAR ALL
// ===============================
function clearAll(){
    const tbody = document.querySelector("#imageList tbody");
    if(tbody) tbody.innerHTML = "";
    imageFiles = [];
}
