from flask import Flask, request, send_file, render_template_string
from docx import Document
from docx.shared import Inches, Pt
from PIL import Image
import io, zipfile, os

app = Flask(__name__)

HTML_PAGE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Folder to Word</title>
    <style>
        body {
            font-family: 'Segoe UI', sans-serif;
            background: #f2f2f2;
            margin: 0;
            padding: 0;
        }
        .container {
            max-width: 600px;
            margin: 5em auto;
            padding: 2em;
            background: white;
            box-shadow: 0 0 20px rgba(0,0,0,0.1);
            border-radius: 12px;
            text-align: center;
        }
        h1 {
            font-size: 2em;
            margin-bottom: 1em;
        }
        input[type="file"] {
            margin: 1.5em 0;
            font-size: 1.1em;
        }
        button {
            padding: 15px 30px;
            font-size: 1.2em;
            background-color: #007BFF;
            color: white;
            border: none;
            border-radius: 8px;
            cursor: pointer;
        }
        button:hover {
            background-color: #0056b3;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🖼️ Folder to Word Document</h1>
        <form method="POST" enctype="multipart/form-data">
            <input type="file" name="zipfile" accept=".zip" required>
            <br>
            <button type="submit">Upload & Generate</button>
        </form>
    </div>
</body>
</html>
'''

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        uploaded_file = request.files.get("zipfile")
        if not uploaded_file or not uploaded_file.filename.endswith(".zip"):
            return "Invalid ZIP file", 400

        zip_bytes = io.BytesIO(uploaded_file.read())

        try:
            with zipfile.ZipFile(zip_bytes) as zip_ref:
                file_list = zip_ref.namelist()
                extracted = {
                    name: zip_ref.read(name) for name in file_list
                    if not name.endswith("/") and name.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif'))
                }

                top_folders = [f.split('/')[0] for f in file_list if '/' in f]
                main_folder_name = top_folders[0] if top_folders else "images"

            folders = {}
            for path, data in extracted.items():
                folder = os.path.dirname(path)
                folders.setdefault(folder, []).append((path, data))

            doc = Document()
            doc.add_heading("📷 Image Gallery", 0)

            for folder, files in folders.items():
                heading = doc.add_paragraph()
                run = heading.add_run(folder or "Root")
                run.font.name = 'Times New Roman'
                run.font.size = Pt(14)

                for filename, img_bytes in sorted(files):
                    try:
                        caption = doc.add_paragraph()
                        run = caption.add_run(os.path.basename(filename))
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(12)

                        image_stream = io.BytesIO(img_bytes)
                        Image.open(image_stream).verify()
                        image_stream.seek(0)
                        doc.add_picture(image_stream, width=Inches(5.5))
                        doc.add_paragraph("\n\n\n")
                    except Exception as e:
                        doc.add_paragraph(f"[Error inserting {filename}: {e}]")

            output_stream = io.BytesIO()
            doc.save(output_stream)
            output_stream.seek(0)

            return send_file(output_stream,
                             as_attachment=True,
                             download_name=f"{main_folder_name}.docx",
                             mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

        except Exception as e:
            return f"Error processing ZIP: {e}", 500

    return render_template_string(HTML_PAGE)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Starting app on port {port}")
    app.run(host="0.0.0.0", port=port)
