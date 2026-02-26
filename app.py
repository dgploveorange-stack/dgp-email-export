from flask import Flask, render_template, request, send_file, jsonify
import extract_msg
import os
import re
import uuid
import zipfile
from PyPDF2 import PdfMerger
from weasyprint import HTML

app = Flask(__name__)

# Base folder for temporary workspace
BASE_FOLDER = "workspace"
os.makedirs(BASE_FOLDER, exist_ok=True)


def process_msg_file(msg_path, work_dir):
    """
    Process a single MSG file:
    - Extract second email
    - Convert to PDF
    - Merge PDF attachments
    Returns the final PDF path and list of temp files for cleanup.
    """
    temp_files = []

    msg = extract_msg.Message(msg_path)
    body = msg.body or ""
    if isinstance(body, bytes):
        body = body.decode("utf-8", errors="replace")

    # Robust From: detection, ignore leading whitespace
    from_pattern = re.compile(r'^\s*From:\s.*', re.IGNORECASE | re.MULTILINE)
    matches = list(from_pattern.finditer(body))
    print(f"DEBUG: Found {len(matches)} From: headers in {msg_path}")

    if len(matches) < 2:
        raise Exception("Thread does not contain a second email.")

    # Extract second email using matched From: positions
    second_start = matches[0].start()
    second_end = matches[1].start()
    second_email = body[second_start:second_end].strip()

    # Optional: remove subject line from search
    lines = second_email.splitlines()
    second_email_body = "\n".join([l for l in lines if not l.lower().startswith("subject:")])

    # Convert to HTML for PDF
    html_content = f"""
    <html>
    <body style="font-family:Arial; font-size:12px;">
    <pre style="white-space:pre-wrap;">{second_email}</pre>
    </body>
    </html>
    """

    # Save email PDF
    email_pdf = os.path.join(work_dir, f"{uuid.uuid4()}_email.pdf")
    HTML(string=html_content).write_pdf(email_pdf)
    temp_files.append(email_pdf)

    # Merge PDF attachments
    merger = PdfMerger()
    merger.append(email_pdf)

    for attachment in msg.attachments:
        name = attachment.longFilename or attachment.shortFilename
        if name and name.lower().endswith(".pdf"):
            attach_path = os.path.join(work_dir, f"{uuid.uuid4()}_{name}")
            with open(attach_path, "wb") as f:
                f.write(attachment.data)
            merger.append(attach_path)
            temp_files.append(attach_path)

    # Extract document number only from the second email
    match = re.search(r'DO\d{2}-\d{5}', second_email_body)
    if match:
        output_name = f"{match.group(0)}.pdf"
    else:
        output_name = f"output_{uuid.uuid4().hex[:6]}.pdf"

    final_path = os.path.join(work_dir, output_name)
    merger.write(final_path)
    merger.close()
    msg.close()

    temp_files.append(final_path)
    return final_path, temp_files


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "files" not in request.files:
        return jsonify({"error": "No files uploaded"}), 400

    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files uploaded"}), 400

    session_id = str(uuid.uuid4())
    work_dir = os.path.join(BASE_FOLDER, session_id)
    os.makedirs(work_dir, exist_ok=True)

    created_files = []
    all_temp_files = []

    try:
        for file in files:
            if not file.filename.lower().endswith(".msg"):
                continue

            msg_path = os.path.join(work_dir, f"{uuid.uuid4()}.msg")
            file.save(msg_path)
            all_temp_files.append(msg_path)

            final_pdf, temp_files = process_msg_file(msg_path, work_dir)
            created_files.append(final_pdf)
            all_temp_files.extend(temp_files)

        if not created_files:
            return jsonify({"error": "No valid MSG files processed"}), 400

        # Create ZIP of all PDFs
        zip_path = os.path.join(work_dir, "converted_files.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for pdf in created_files:
                zipf.write(pdf, os.path.basename(pdf))

        # Send ZIP to client
        response = send_file(zip_path, as_attachment=True)

        # Cleanup temp files after sending
        @response.call_on_close
        def cleanup():
            for f in all_temp_files:
                if os.path.exists(f):
                    os.remove(f)
            if os.path.exists(zip_path):
                os.remove(zip_path)
            if os.path.exists(work_dir):
                os.rmdir(work_dir)

        return response

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    # Run locally for testing
    app.run(host="0.0.0.0", port=10000, debug=True)
