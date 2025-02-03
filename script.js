document.getElementById("fileInput").addEventListener("change", handleFiles);
document.getElementById("convertBtn").addEventListener("click", convertFiles);

let selectedFiles = [];

function handleFiles(event) {
    selectedFiles = event.target.files;
    displayFileList();
}

function displayFileList() {
    const fileList = document.getElementById("file-list");
    fileList.innerHTML = "<h3>File yang Diupload:</h3>";

    Array.from(selectedFiles).forEach(file => {
        fileList.innerHTML += `<p class="file-item">${file.name}</p>`;
    });
}

async function convertFiles() {
    if (selectedFiles.length === 0) {
        alert("Silakan upload file terlebih dahulu!");
        return;
    }

    const output = document.getElementById("output");
    output.innerHTML = "<h3>Hasil Konversi:</h3>";

    for (let file of selectedFiles) {
        const ext = file.name.split('.').pop().toLowerCase();
        
        if (ext === "vcf") {
            await convertVCF(file);
        } else if (ext === "xlsx" || ext === "xls") {
            await convertExcel(file);
        } else if (ext === "docx") {
            await convertDocx(file);
        } else {
            output.innerHTML += `<p style="color:red;">Format ${file.name} tidak didukung</p>`;
        }
    }
}

async function convertVCF(file) {
    const text = await file.text();
    const matches = text.match(/TEL[^:]*:([\d+\-() ]+)/g);
    
    let result = matches ? matches.map(m => m.split(":")[1].trim()).join("\n") : "Tidak ada nomor ditemukan";
    downloadTxt(file.name.replace(".vcf", ".txt"), result);
}

async function convertExcel(file) {
    const reader = new FileReader();
    reader.readAsArrayBuffer(file);
    reader.onload = async function () {
        const data = new Uint8Array(reader.result);
        const workbook = XLSX.read(data, { type: "array" });

        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const text = XLSX.utils.sheet_to_csv(sheet);
            downloadTxt(`${file.name.replace(/\..+$/, '')}_${sheetName}.txt`, text);
        });
    };
}

async function convertDocx(file) {
    const reader = new FileReader();
    reader.readAsArrayBuffer(file);
    reader.onload = async function () {
        const doc = await window.mammoth.extractRawText({ arrayBuffer: reader.result });
        downloadTxt(file.name.replace(".docx", ".txt"), doc.value);
    };
}

function downloadTxt(filename, content) {
    const blob = new Blob([content], { type: "text/plain" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.textContent = `Download ${filename}`;
    link.className = "download-btn";
    document.getElementById("output").appendChild(link);
    document.getElementById("output").appendChild(document.createElement("br"));
}

// Menambahkan pustaka eksternal untuk Excel & Word
const script1 = document.createElement("script");
script1.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.3/xlsx.full.min.js";
document.body.appendChild(script1);

const script2 = document.createElement("script");
script2.src = "https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.4.2/mammoth.browser.min.js";
document.body.appendChild(script2);
