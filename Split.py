from flask import Flask, request, send_file, render_template_string, redirect, url_for
import os, uuid, zipfile, fitz

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
PREVIEW_FOLDER = "static/previews"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(PREVIEW_FOLDER, exist_ok=True)

# HTML Template
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PDF Splitter — Preview & manual split</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    body {
      background: linear-gradient(135deg, #6a11cb, #2575fc);
      min-height: 100vh;
      padding: 30px;
      color: #fff;
    }
    .card {
      border-radius: 20px;
      box-shadow: 0px 10px 30px rgba(0,0,0,0.2);
    }
    .preview-img {
      max-width: 150px;
      border: 2px solid #ccc;
      border-radius: 8px;
      margin: 5px;
    }
    .progress { margin-top: 20px; height: 25px; }
  </style>
  <script>
    function showProgress() {
      document.getElementById("progressBar").style.display = "block";
    }
  </script>
</head>
<body>
  <div class="container">
    <div class="card p-4">
      <h2 class="text-center">📄 PDF Splitter — Preview & Manual Split</h2>
      <p class="text-center">Upload a PDF, preview pages, and select ranges to split</p>

      {% if error %}
      <div class="alert alert-danger">{{ error }}</div>
      {% endif %}

      {% if not previews %}
      <form method="post" enctype="multipart/form-data" onsubmit="showProgress()">
        <input type="file" name="pdf_file" accept="application/pdf" required class="form-control mb-3">
        <button type="submit" class="btn btn-primary w-100">Upload & Preview</button>
      </form>
      <div class="progress" id="progressBar" style="display:none;">
        <div class="progress-bar progress-bar-striped progress-bar-animated bg-success" 
             style="width: 100%">Processing...</div>
      </div>
      {% else %}
      <div class="text-center">
        {% for img, num in previews %}
          <div style="display:inline-block; text-align:center;">
            <img src="{{ img }}" class="preview-img"><br>
            <small>Page {{ num }}</small>
          </div>
        {% endfor %}
      </div>

      <form method="post" action="{{ url_for('split_pdf') }}">
        <input type="hidden" name="pdf_path" value="{{ pdf_path }}">
        <label class="mt-3">Enter page ranges (example: 1-3,5,7-9):</label>
        <input type="text" name="ranges" class="form-control mb-3" required>
        <button type="submit" class="btn btn-success w-100">Split & Download</button>
      </form>
      {% endif %}
    </div>
  </div>
</body>
</html>
"""

def create_previews(pdf_path):
    doc = fitz.open(pdf_path)
    previews = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        pix = page.get_pixmap(matrix=fitz.Matrix(0.3, 0.3))  # low-res preview
        img_filename = f"{uuid.uuid4()}.png"
        img_path = os.path.join(PREVIEW_FOLDER, img_filename)
        pix.save(img_path)
        previews.append((f"/static/previews/{img_filename}", page_num + 1))
    return previews

@app.route("/", methods=["GET", "POST"])
def upload_pdf():
    if request.method == "POST":
        try:
            file = request.files["pdf_file"]
            if not file:
                return render_template_string(HTML_TEMPLATE, error="No file uploaded", previews=None)

            pdf_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}.pdf")
            file.save(pdf_path)
            previews = create_previews(pdf_path)
            return render_template_string(HTML_TEMPLATE, previews=previews, pdf_path=pdf_path, error=None)
        except Exception as e:
            return render_template_string(HTML_TEMPLATE, error=f"Failed to create previews: {str(e)}", previews=None)
    return render_template_string(HTML_TEMPLATE, previews=None, error=None)

@app.route("/split", methods=["POST"])
def split_pdf():
    try:
        pdf_path = request.form["pdf_path"]
        ranges = request.form["ranges"]

        doc = fitz.open(pdf_path)
        output_folder = os.path.join(OUTPUT_FOLDER, str(uuid.uuid4()))
        os.makedirs(output_folder, exist_ok=True)

        # Parse ranges like 1-3,5,7-9
        for part_num, part in enumerate(ranges.split(","), start=1):
            part = part.strip()
            writer = fitz.open()

            if "-" in part:
                start, end = map(int, part.split("-"))
                for i in range(start - 1, end):
                    writer.insert_pdf(doc, from_page=i, to_page=i)
            else:
                i = int(part) - 1
                writer.insert_pdf(doc, from_page=i, to_page=i)

            out_path = os.path.join(output_folder, f"part_{part_num}.pdf")
            writer.save(out_path)
            writer.close()

        # Zip results
        zip_filename = f"{uuid.uuid4()}.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, zip_filename)
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for f in os.listdir(output_folder):
                zipf.write(os.path.join(output_folder, f), arcname=f)

        return send_file(zip_path, as_attachment=True, download_name="split_pdfs.zip")

    except Exception as e:
        return f"Error during split: {str(e)}"

if __name__ == "__main__":
    app.run(debug=True)