from flask import Flask, render_template, request, send_file
import extract_msg
import os
import re
from PyPDF2 import PdfMerger
from weasyprint import HTML

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
ATTACHMENTS_DIR = "attachments"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("msgfile")

        if not file:
            return "No file uploaded"

        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        msg = extract_msg.Message(filepath)
        body = msg.body or ""

        # ---------- FIND SECOND EMAIL ----------
        from_pattern = re.compile(r'^From:\s.*', re.IGNORECASE | re.MULTILINE)
        matches = list(from_pattern.finditer(body))

        if len(matches) < 2:
            return "Second email not found"

        second_email = body[matches[0].start():matches[1].start()]

        # ---------- CREATE HTML ----------
        html_content = f"""
        <html>
        <head>
        <meta charset="UTF-8">
        <style>
        body {{
            font-family: Arial, sans-serif;
            font-size: 12px;
            line-height: 1.6;
        }}
        pre {{
            white-space: pre-wrap;
            word-wrap: break-word;
        }}
        </style>
        </head>
        <body>
        <pre>{second_email}</pre>
        </body>
        </html>
        """

        temp_pdf = "email.pdf"
        HTML(string=html_content).write_pdf(temp_pdf)

        # ---------- MERGE WITH PDF ATTACHMENTS ----------
        merger = PdfMerger()
        merger.append(temp_pdf)

        for attachment in msg.attachments:
            name = attachment.longFilename or attachment.shortFilename
            if name and name.lower().endswith(".pdf"):
                path = os.path.join(ATTACHMENTS_DIR, name)
                with open(path, "wb") as f:
                    f.write(attachment.data)
                merger.append(path)

        output_pdf = "final_output.pdf"
        merger.write(output_pdf)
        merger.close()

        return send_file(output_pdf, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
