# improved_pdf_to_word_ocr.py
from flask import Flask, request, send_file, render_template_string
import os, zipfile, logging, subprocess, uuid, tempfile, shutil
import pdfplumber
from pdf2docx import Converter
import pytesseract
from pdf2image import convert_from_path
from docx import Document
from PIL import Image
import numpy as np
import cv2

logging.basicConfig(level=logging.INFO)
app = Flask(__name__)

# ---------------- CONFIG - UPDATE THESE PATHS ----------------
# Windows example - change if installed elsewhere
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# Poppler bin folder you extracted (contains pdftoppm). Example:
POPPLER_PATH = r"C:\poppler-23.11.0\Library\bin"
# DPI to render PDF pages for OCR (higher -> better OCR but slower/more RAM)
OCR_DPI = 400
# ----------------------------------------------------------------

pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

# HTML - simple UI with language + mode
HTML_FORM = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>PDF → Word Converter</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css" rel="stylesheet">
  <style>
    body {
      background: linear-gradient(to right, #667eea, #764ba2);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      color: #fff;
    }
    .card {
      border-radius: 20px;
      box-shadow: 0 8px 20px rgba(0,0,0,0.4);
    }
    h4 {
      font-weight: 600;
      color: #333;
    }
    .form-label {
      font-weight: 500;
      color: #555;
    }
    .btn-primary {
      background: #764ba2;
      border: none;
    }
    .btn-primary:hover {
      background: #667eea;
    }
    .form-text {
      color: #888;
    }
    .note {
      font-size: 0.85rem;
      color: #eee;
      margin-top: 1rem;
    }
    .footer {
      font-size: 0.75rem;
      color: #ccc;
      margin-top: 1rem;
      text-align: center;
    }
  </style>
</head>
<body>
  <div class="container py-5">
    <div class="card mx-auto" style="max-width:800px;">
      <div class="card-body p-4">
        <h4 class="card-title mb-4 text-center">📄 PDF → Word Converter</h4>
        <p class="text-center text-muted">Supports OCR for Hindi, English & multi-language PDFs</p>
        <form method="post" action="/convert" enctype="multipart/form-data">
          <div class="mb-3">
            <label class="form-label"><i class="fa-solid fa-file-pdf"></i> Select PDF files</label>
            <input class="form-control" type="file" name="files" accept="application/pdf" multiple required>
          </div>

          <div class="mb-3">
            <label class="form-label"><i class="fa-solid fa-cogs"></i> Mode</label>
            <select name="mode" class="form-select">
              <option value="auto" selected>Auto (Detects text / OCR)</option>
              <option value="direct">Direct conversion (pdf2docx)</option>
              <option value="ocr">Force OCR (images + text extraction)</option>
            </select>
          </div>

          <div class="mb-3">
            <label class="form-label"><i class="fa-solid fa-language"></i> OCR Language</label>
            <select name="lang" class="form-select">
              <option value="eng+hin" selected>English + Hindi</option>
              <option value="eng">English</option>
              <option value="hin">Hindi</option>
              <option value="eng+hin+tam">English + Hindi + Tamil</option>
            </select>
            <div class="form-text">Ensure selected languages are installed in Tesseract OCR.</div>
          </div>

          <button class="btn btn-primary w-100 py-2" type="submit"><i class="fa-solid fa-arrow-down-to-line"></i> Convert PDF</button>
        </form>

        <p class="note text-center">
          
        </p>

        <p class="footer"></p>
      </div>
    </div>
  </div>
</body>
</html>
"""

@app.route("/", methods=["GET"])
def home():
    return render_template_string(HTML_FORM)


@app.route("/convert", methods=["POST"])
def convert_files():
    files = request.files.getlist("files")
    mode = request.form.get("mode", "auto")
    lang = request.form.get("lang", "eng+hin")

    tmpdir = tempfile.mkdtemp(prefix="pdf2word_")
    converted = []

    try:
        # check available tesseract langs
        available_langs = get_tesseract_langs()
        logging.info("Tesseract available langs: %s", available_langs)
        # if user asked for hin but it's missing, warn & fallback to english
        wanted = set(lang.split("+"))
        missing = [l for l in wanted if l not in available_langs]
        if missing:
            logging.warning("Requested languages %s not installed in Tesseract. Missing: %s. Falling back to 'eng'.",
                            wanted, missing)
            # fallback to english
            if "eng" in available_langs:
                lang = "eng"
            else:
                # use first available
                lang = available_langs[0] if available_langs else "eng"

        for up in files:
            safe_name = str(uuid.uuid4()) + "__" + secure_filename(up.filename)
            pdf_path = os.path.join(tmpdir, safe_name)
            up.save(pdf_path)
            out_docx = os.path.splitext(pdf_path)[0] + ".docx"

            try:
                if mode == "direct":
                    logging.info("Direct convert: %s", up.filename)
                    pdf_to_word(pdf_path, out_docx)
                elif mode == "ocr":
                    logging.info("Force OCR: %s", up.filename)
                    pdf_to_word_ocr(pdf_path, out_docx, lang=lang)
                else:  # auto
                    if pdf_has_text(pdf_path):
                        logging.info("PDF has selectable text -> direct convert: %s", up.filename)
                        pdf_to_word(pdf_path, out_docx)
                    else:
                        logging.info("PDF looks scanned -> OCR convert: %s", up.filename)
                        pdf_to_word_ocr(pdf_path, out_docx, lang=lang)

                converted.append(out_docx)
            except Exception as e:
                logging.exception("Conversion failed for %s: %s", up.filename, e)
                # final fallback: try OCR
                try:
                    pdf_to_word_ocr(pdf_path, out_docx, lang=lang)
                    converted.append(out_docx)
                except Exception:
                    logging.exception("OCR fallback also failed for %s", up.filename)

        # single file -> send it, multiple -> zip
        if len(converted) == 1:
            return send_file(converted[0], as_attachment=True, download_name=os.path.basename(converted[0]))
        else:
            zip_path = os.path.join(tmpdir, "converted.zip")
            with zipfile.ZipFile(zip_path, "w") as zf:
                for f in converted:
                    zf.write(f, os.path.basename(f))
            return send_file(zip_path, as_attachment=True, download_name="converted.zip")

    finally:
        # cleanup tmpdir after response (Flask will finish sending file first)
        shutil.rmtree(tmpdir, ignore_errors=True)


# ---------------- Utility & Conversion funcs ----------------

def get_tesseract_langs():
    """Return a list of installed tesseract language codes, or [] if cannot fetch."""
    try:
        cmd = [pytesseract.pytesseract.tesseract_cmd, "--list-langs"]
        out = subprocess.check_output(cmd, stderr=subprocess.STDOUT, universal_newlines=True)
        lines = [l.strip() for l in out.splitlines() if l.strip()]
        # first line is "List of available languages (N):"
        if len(lines) >= 2:
            return lines[1:]
        return []
    except Exception as e:
        logging.exception("Could not list tesseract languages: %s", e)
        return []


def pdf_has_text(pdf_path):
    """Return True if pdf contains selectable (extractable) text on any page."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t and t.strip():
                    return True
        return False
    except Exception as e:
        logging.exception("pdf_has_text check failed: %s", e)
        return False


def pdf_to_word(pdf_path, docx_path):
    """Direct conversion (pdf2docx) for PDFs that already include text."""
    cv = Converter(pdf_path)
    try:
        cv.convert(docx_path, start=0, end=None)
    finally:
        cv.close()


def pdf_to_word_ocr(pdf_path, docx_path, lang="eng"):
    """
    Convert scanned PDF to Word using PDF->images + preprocessing + pytesseract.
    Preprocessing improves OCR results for many scanned PDFs.
    """
    # render pages as images
    images = convert_from_path(pdf_path, dpi=OCR_DPI, poppler_path=POPPLER_PATH)
    doc = Document()

    for i, pil_img in enumerate(images, start=1):
        logging.info("OCR page %d (dpi=%d)", i, OCR_DPI)

        # Convert to OpenCV image (BGR)
        img_cv = cv2.cvtColor(np.array(pil_img.convert("RGB")), cv2.COLOR_RGB2BGR)

        # 1) Grayscale
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)

        # 2) Resize to improve small-text recognition
        h, w = gray.shape
        scale = 1.5
        gray = cv2.resize(gray, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_CUBIC)

        # 3) Denoise
        gray = cv2.medianBlur(gray, 3)

        # 4) Adaptive threshold (works well for uneven lighting)
        try:
            thresh = cv2.adaptiveThreshold(gray, 255,
                                           cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                           cv2.THRESH_BINARY, 21, 10)
        except Exception:
            # fallback to Otsu if adaptive fails
            _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

        # 5) Optional morphological open to remove small noise (fine-tune kernel)
        kernel = np.ones((1, 1), np.uint8)
        processed = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)

        # Convert processed back to PIL for pytesseract
        proc_pil = Image.fromarray(processed)

        # Try OCR with a few settings (order chosen because they often help)
        text = ""
        configs = [
            r'--oem 1 --psm 3',  # default block
            r'--oem 1 --psm 6',  # assume a uniform block of text
            r'--oem 1 --psm 11'  # sparse text
        ]
        for cfg in configs:
            try:
                text = pytesseract.image_to_string(proc_pil, lang=lang, config=cfg)
                if text and text.strip():
                    logging.info("OCR success with config: %s (len=%d)", cfg, len(text.strip()))
                    break
            except Exception as e:
                logging.exception("pytesseract failed with config %s: %s", cfg, e)
                text = ""

        # Final fallback: try raw image (no threshold)
        if not text.strip():
            try:
                text = pytesseract.image_to_string(pil_img, lang=lang, config=r'--oem 1 --psm 3')
            except Exception:
                text = ""

        # If still empty, embed original page image (but only after trying multiple OCR attempts)
        if text and text.strip() and len(text.strip()) > 10:
            # add extracted text
            doc.add_heading(f"Page {i}", level=2)
            # Preserve line breaks from OCR
            for para in text.splitlines():
                if para.strip():
                    doc.add_paragraph(para)
        else:
            # Save page image to temp and embed
            tmp_img_path = os.path.join(os.path.dirname(docx_path), f"page_{i}.png")
            pil_img.save(tmp_img_path, format="PNG")
            doc.add_heading(f"Page {i} (image embedded - OCR produced little/no text)", level=3)
            doc.add_paragraph("⚠️ OCR produced little or no readable text for this page. The page image is embedded below.")
            doc.add_picture(tmp_img_path, width=None)
            try:
                os.remove(tmp_img_path)
            except Exception:
                pass

        doc.add_page_break()

    doc.save(docx_path)


# small helper to create filesystem-safe filename
def secure_filename(name: str) -> str:
    return "".join(c if c.isalnum() or c in " ._-()" else "_" for c in name)


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)