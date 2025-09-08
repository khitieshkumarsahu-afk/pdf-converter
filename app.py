from flask import Flask, request, send_file, render_template_string
import os, zipfile
import pdfplumber
from pdf2docx import Converter
import openpyxl
import camelot
import tabula
import pandas as pd

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# HTML Page with Progress Bar
HTML_FORM = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PDF Converter</title>
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body {
            background: linear-gradient(135deg, #74ebd5 0%, #ACB6E5 100%);
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        .card {
            border-radius: 20px;
            box-shadow: 0px 8px 25px rgba(0, 0, 0, 0.2);
            max-width: 600px;
            width: 100%;
        }
        #dropZone {
            border: 3px dashed #0d6efd;
            border-radius: 15px;
            padding: 40px;
            text-align: center;
            background: #f8f9fa;
            cursor: pointer;
            transition: background 0.3s ease;
        }
        #dropZone.dragover {
            background: #e0f7fa;
        }
        #progressBar {
            width: 100%;
            background-color: #e9ecef;
            border-radius: 15px;
            overflow: hidden;
        }
        #progress {
            width: 0%;
            height: 25px;
            background: linear-gradient(to right, #00b09b, #96c93d);
            color: white;
            text-align: center;
            line-height: 25px;
            transition: width 0.4s ease;
        }
        /* File preview style */
        #fileList .file-item {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin: 5px 0;
            padding: 6px 10px;
            border: 1px solid #ccc;
            border-radius: 8px;
            background: #fff;
        }
        #fileList .file-info {
            display: flex;
            align-items: center;
        }
        #fileList .file-info span {
            margin-left: 8px;
        }
        .remove-btn {
            cursor: pointer;
            color: red;
            font-weight: bold;
            margin-left: 10px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="card p-5">
            <h2 class="text-center mb-4">📄 PDF to Word / Excel Converter</h2>

            <!-- Drag & Drop Upload Zone -->
            <div id="dropZone">
                <p class="mb-0">📂 Drag & Drop PDF here or <strong>Click to Upload</strong></p>
                <input type="file" id="fileInput" class="d-none" accept="application/pdf" multiple required>
            </div>
            
            <!-- File preview -->
            <div id="fileList" class="mt-3"></div>
            <br>

            <!-- Format Selector -->
            <form id="uploadForm" action="/convert" method="post" enctype="multipart/form-data" class="text-center">
                <div class="mb-3">
                    <select class="form-select" name="format" required>
                        <option value="word">Convert to Word (with images & layout)</option>
                        <option value="excel">Convert to Excel (tables only)</option>
                        <option value="excel_full">Convert to Excel (preserve full structure)</option>
                    </select>
                </div>
                <button type="submit" class="btn btn-success w-100">🚀 Convert</button>
            </form>
            <br>

            <!-- Progress -->
            <div id="progressBar"><div id="progress">0%</div></div>
            <p id="status" class="mt-3 text-center text-dark fw-bold"></p>
        </div>
    </div>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>

    <script>
    const dropZone = document.getElementById("dropZone");
    const fileInput = document.getElementById("fileInput");
    const form = document.getElementById("uploadForm");
    const progress = document.getElementById("progress");
    const status = document.getElementById("status");
    const fileList = document.getElementById("fileList");

    // Keep track of selected files
    let selectedFiles = [];

    // ✅ Show PDF name + size + remove option
    function showFileNames() {
        fileList.innerHTML = "";
        selectedFiles.forEach((file, index) => {
            let sizeKB = (file.size / 1024).toFixed(2);
            let sizeText = sizeKB > 1024 
                ? (sizeKB/1024).toFixed(2) + " MB" 
                : sizeKB + " KB";

            let div = document.createElement("div");
            div.className = "file-item";

            let info = document.createElement("div");
            info.className = "file-info";
            info.innerHTML = "📄 <span>" + file.name + " (" + sizeText + ")</span>";

            let removeBtn = document.createElement("span");
            removeBtn.className = "remove-btn";
            removeBtn.innerHTML = "❌";
            removeBtn.onclick = () => {
                selectedFiles.splice(index, 1);
                showFileNames();
            };

            div.appendChild(info);
            div.appendChild(removeBtn);
            fileList.appendChild(div);
        });
    }

    dropZone.onclick = () => fileInput.click();
    fileInput.onchange = () => {
        selectedFiles = [...selectedFiles, ...fileInput.files];
        showFileNames();
    };

    dropZone.ondragover = (e) => {
        e.preventDefault();
        dropZone.classList.add("dragover");
    };
    dropZone.ondragleave = () => dropZone.classList.remove("dragover");
    dropZone.ondrop = (e) => {
        e.preventDefault();
        dropZone.classList.remove("dragover");
        selectedFiles = [...selectedFiles, ...e.dataTransfer.files];
        showFileNames();
    };

    form.onsubmit = function(e) {
        e.preventDefault();
        let xhr = new XMLHttpRequest();
        xhr.open("POST", "/convert");

        xhr.upload.onprogress = function(e) {
            if (e.lengthComputable) {
                let percent = (e.loaded / e.total) * 40;
                progress.style.width = percent + "%";
                progress.innerHTML = Math.round(percent) + "%";
                status.innerHTML = "⏳ Uploading file...";
            }
        };

        xhr.onloadstart = function() {
            progress.style.width = "0%";
            progress.innerHTML = "0%";
            status.innerHTML = "⏳ Starting process...";
        };

        xhr.onload = function() {
            if (xhr.status === 200) {
                simulateConversionSteps(() => {
                    status.innerHTML = "✅ Done! Downloading...";
                    const blob = new Blob([xhr.response], {type: xhr.getResponseHeader("Content-Type")});
                    const link = document.createElement("a");
                    link.href = window.URL.createObjectURL(blob);
                    let filename = xhr.getResponseHeader("Content-Disposition").split("filename=")[1];
                    link.download = filename.replace(/"/g, '');
                    link.click();
                    status.innerHTML = "🎉 Download complete!";
                });
            } else {
                status.innerHTML = "❌ Error during conversion.";
            }
        };

        xhr.responseType = "blob";
        let formData = new FormData();
        for (let file of selectedFiles) {
            formData.append("files", file);
        }
        formData.append("format", form.querySelector("select").value);
        xhr.send(formData);
    };

    function simulateConversionSteps(callback) {
        let steps = [
            { msg: "⚙️ Extracting content...", progress: 60 },
            { msg: "📊 Converting to selected format...", progress: 80 },
            { msg: "📦 Preparing file for download...", progress: 100 }
        ];

        let i = 0;
        function nextStep() {
            if (i < steps.length) {
                setTimeout(() => {
                    progress.style.width = steps[i].progress + "%";
                    progress.innerHTML = steps[i].progress + "%";
                    status.innerHTML = steps[i].msg;
                    i++;
                    nextStep();
                }, 800);
            } else {
                callback();
            }
        }
        nextStep();
    }
    </script>
</body>
</html>
"""
@app.route('/')
def home():
    return render_template_string(HTML_FORM)

@app.route('/convert', methods=['POST'])
def convert_pdf():
    files = request.files.getlist('files')
    target_format = request.form['format']

    converted_files = []

    for file in files:
        filename = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filename)

        if target_format == "word":
            output_file = os.path.join(OUTPUT_FOLDER, file.filename.replace(".pdf", ".docx"))
            pdf_to_word(filename, output_file)
        elif target_format == "excel":
            output_file = os.path.join(OUTPUT_FOLDER, file.filename.replace(".pdf", ".xlsx"))
            pdf_to_excel(filename, output_file)
        else:
            output_file = os.path.join(OUTPUT_FOLDER, file.filename.replace(".pdf", "_structured.xlsx"))
            pdf_to_excel_full(filename, output_file)

        converted_files.append(output_file)

    if len(converted_files) == 1:
        return send_file(converted_files[0], as_attachment=True)

    zip_path = os.path.join(OUTPUT_FOLDER, "converted_files.zip")
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for f in converted_files:
            zipf.write(f, os.path.basename(f))

    return send_file(zip_path, as_attachment=True)


def pdf_to_word(pdf_path, docx_path):
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()


def pdf_to_excel(pdf_path, excel_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PDF Data"

    tables = camelot.read_pdf(pdf_path, pages='all')
    if tables:
        for t in tables:
            data = t.df.values.tolist()
            for line in data:
                ws.append(line)
    else:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    for line in text.split("\n"):
                        ws.append([line])

    wb.save(excel_path)


def pdf_to_excel_full(pdf_path, excel_path):
    try:
        dfs = tabula.read_pdf(pdf_path, pages="all", multiple_tables=True, stream=True)

        if dfs:
            writer = pd.ExcelWriter(excel_path, engine="openpyxl")
            for i, df in enumerate(dfs):
                sheet_name = f"Page_{i+1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            writer.close()
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        for line in text.split("\n"):
                            ws.append([line])
            wb.save(excel_path)

    except Exception as e:
        print("Tabula error:", e)
        pdf_to_excel(pdf_path, excel_path)


if __name__ == '__main__':
    import os
    port = int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0", port=port)
