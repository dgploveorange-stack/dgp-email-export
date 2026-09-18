from flask import Flask, render_template, request, send_file, jsonify
import extract_msg
import os
import re
import uuid
import zipfile
import html
import shutil
from PyPDF2 import PdfMerger
from weasyprint import HTML

app = Flask(__name__)

# ============================================================
# CONFIGURATION
# ============================================================

BASE_FOLDER = "workspace"
os.makedirs(BASE_FOLDER, exist_ok=True)


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def decode_content(content):
    """
    Convert MSG content into a normal Python string.
    """
    if content is None:
        return ""

    if isinstance(content, bytes):
        # Try UTF-8 first
        try:
            return content.decode("utf-8")
        except UnicodeDecodeError:
            pass

        # Fallback encodings commonly encountered in email
        for encoding in ["cp1252", "latin-1"]:
            try:
                return content.decode(encoding)
            except UnicodeDecodeError:
                continue

        return content.decode("utf-8", errors="replace")

    return str(content)


def get_msg_html(msg):
    """
    Get the HTML version of the MSG body if available.
    Returns empty string if no usable HTML body exists.
    """

    html_body = getattr(msg, "htmlBody", None)

    if not html_body:
        return ""

    html_body = decode_content(html_body).strip()

    if not html_body:
        return ""

    # Some MSG files can contain a complete HTML document,
    # while others contain only the body fragment.
    return html_body


def get_msg_plain_text(msg):
    """
    Get the plain-text version of the MSG body.
    """

    body = getattr(msg, "body", None)

    if not body:
        return ""

    return decode_content(body).strip()


def remove_quoted_email_plain_text(body):
    """
    Remove previous/quoted emails from a plain-text Outlook email.

    Outlook commonly uses:
        -----Original Message-----

    We also support common From:/Sent:/To:/Subject: blocks.
    """

    if not body:
        return ""

    # --------------------------------------------------------
    # Method 1: Outlook separator
    # --------------------------------------------------------

    separators = [
        r"^-{3,}\s*Original Message\s*-{3,}\s*$",
        r"^-{3,}\s*Original Appointment\s*-{3,}\s*$",
    ]

    for pattern in separators:
        match = re.search(
            pattern,
            body,
            flags=re.IGNORECASE | re.MULTILINE
        )

        if match:
            current_email = body[:match.start()].strip()

            if current_email:
                return current_email

    # --------------------------------------------------------
    # Method 2: Detect quoted From: block
    # --------------------------------------------------------

    # Look for a conventional quoted header block.
    quoted_header_pattern = re.compile(
        r"""
        ^\s*From:\s*.+$
        \s*
        ^\s*(?:Sent|Date):\s*.+$
        \s*
        ^\s*To:\s*.+$
        (?:\s*
        ^\s*Cc:\s*.+$)?
        \s*
        ^\s*Subject:\s*.+$
        """,
        re.IGNORECASE |
        re.MULTILINE |
        re.VERBOSE
    )

    match = quoted_header_pattern.search(body)

    if match:
        # Only remove it if there is meaningful content before it.
        before = body[:match.start()].strip()

        if before:
            return before

    # --------------------------------------------------------
    # Method 3: Fallback to From: headers
    # --------------------------------------------------------

    from_pattern = re.compile(
        r"^\s*From:\s*.+$",
        re.IGNORECASE | re.MULTILINE
    )

    matches = list(from_pattern.finditer(body))

    if len(matches) >= 2:
        # Assume the second From: starts the quoted email.
        before_second_from = body[:matches[1].start()].strip()

        if before_second_from:
            return before_second_from

    # No quoted content detected.
    return body.strip()


def remove_quoted_email_html(html_body):
    """
    Remove previous/quoted content from HTML email.

    This is intentionally conservative. If a recognizable
    Outlook quoted-email section is found, remove everything
    from that point onward.
    """

    if not html_body:
        return ""

    body = html_body

    # --------------------------------------------------------
    # Common Outlook HTML separator
    # --------------------------------------------------------

    separator_patterns = [
        r"-{3,}\s*Original Message\s*-{3,}",
        r"-{3,}\s*Original Appointment\s*-{3,}",
    ]

    for pattern in separator_patterns:
        match = re.search(
            pattern,
            body,
            flags=re.IGNORECASE
        )

        if match:
            before = body[:match.start()].strip()

            if before:
                return before

    # --------------------------------------------------------
    # Outlook quoted-message containers
    # --------------------------------------------------------

    quoted_patterns = [
        r'<div[^>]*class=["\'][^"\']*gmail_quote[^"\']*["\'][^>]*>',
        r'<blockquote[^>]*>',
        r'<div[^>]*class=["\'][^"\']*OutlookMessageHeader[^"\']*["\'][^>]*>',
    ]

    for pattern in quoted_patterns:
        match = re.search(
            pattern,
            body,
            flags=re.IGNORECASE
        )

        if match:
            before = body[:match.start()].strip()

            if before:
                return before

    return body.strip()


def build_pdf_html(content, is_html=False):
    """
    Convert email content into HTML suitable for WeasyPrint.
    """

    if is_html:
        email_content = content
    else:
        # Escape plain text so email characters such as
        # <, > and & don't break the generated HTML.
        email_content = html.escape(content)

        email_content = f"""
        <pre style="
            white-space: pre-wrap;
            word-wrap: break-word;
            font-family: Arial, sans-serif;
            font-size: 12px;
            line-height: 1.4;
        ">{email_content}</pre>
        """

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">

        <style>
            @page {{
                size: A4;
                margin: 18mm;
            }}

            body {{
                font-family: Arial, sans-serif;
                font-size: 12px;
                line-height: 1.4;
                color: #000;
                word-wrap: break-word;
            }}

            img {{
                max-width: 100%;
                height: auto;
            }}

            table {{
                max-width: 100%;
                border-collapse: collapse;
            }}

            td, th {{
                vertical-align: top;
            }}

            pre {{
                white-space: pre-wrap;
                word-wrap: break-word;
                font-family: Arial, sans-serif;
            }}
        </style>
    </head>

    <body>
        {email_content}
    </body>
    </html>
    """


def get_email_content(msg):
    """
    Get the best available email body.

    Priority:
        1. HTML
        2. Plain text

    Returns:
        (content, is_html)
    """

    # --------------------------------------------------------
    # Try HTML first
    # --------------------------------------------------------

    html_body = get_msg_html(msg)

    if html_body:
        cleaned_html = remove_quoted_email_html(html_body)

        if cleaned_html:
            return cleaned_html, True

    # --------------------------------------------------------
    # Fallback to plain text
    # --------------------------------------------------------

    plain_body = get_msg_plain_text(msg)

    if plain_body:
        cleaned_text = remove_quoted_email_plain_text(plain_body)

        if cleaned_text:
            return cleaned_text, False

    return "", False


def find_output_name(subject, body):
    """
    Find DOxx-xxxxx anywhere in subject/body.

    Example:
        DO12-12345
        (DO12-12345)
        Re: DO12-12345
    """

    pattern = r"\bDO\d{2}-\d{5}\b"

    # Body first
    match = re.search(pattern, body or "", re.IGNORECASE)

    if match:
        return f"{match.group(0).upper()}.pdf"

    # Subject second
    match = re.search(pattern, subject or "", re.IGNORECASE)

    if match:
        return f"{match.group(0).upper()}.pdf"

    # Fallback
    return f"output_{uuid.uuid4().hex[:6]}.pdf"


# ============================================================
# PROCESS MSG
# ============================================================

def process_msg_file(msg_path, work_dir):
    """
    Process a single MSG file:

    1. Read MSG
    2. Extract latest/current email body
    3. Prefer HTML body
    4. Convert email to PDF
    5. Add PDF attachments
    6. Name final PDF using DOxx-xxxxx
    """

    temp_files = []

    msg = None

    try:
        # ----------------------------------------------------
        # Open MSG
        # ----------------------------------------------------

        msg = extract_msg.Message(msg_path)

        subject = decode_content(
            getattr(msg, "subject", "") or ""
        )

        print("=" * 80)
        print("Processing:", msg_path)
        print("Subject:", subject)

        # ----------------------------------------------------
        # Extract email content
        # ----------------------------------------------------

        content, is_html = get_email_content(msg)

        if not content:
            raise Exception(
                "Unable to extract readable email body from MSG file"
            )

        print("Body format:", "HTML" if is_html else "Plain Text")
        print("Extracted body length:", len(content))

        # ----------------------------------------------------
        # Generate email PDF
        # ----------------------------------------------------

        html_content = build_pdf_html(
            content,
            is_html=is_html
        )

        email_pdf = os.path.join(
            work_dir,
            f"{uuid.uuid4()}_email.pdf"
        )

        HTML(
            string=html_content,
            base_url=os.path.dirname(os.path.abspath(msg_path))
        ).write_pdf(email_pdf)

        temp_files.append(email_pdf)

        # ----------------------------------------------------
        # Create PDF merger
        # ----------------------------------------------------

        merger = PdfMerger()

        try:
            merger.append(email_pdf)

            # ------------------------------------------------
            # Process PDF attachments
            # ------------------------------------------------

            attachments = getattr(msg, "attachments", []) or []

            for attachment in attachments:

                name = (
                    getattr(attachment, "longFilename", None)
                    or getattr(attachment, "shortFilename", None)
                )

                if not name:
                    continue

                name = os.path.basename(str(name))

                if not name.lower().endswith(".pdf"):
                    continue

                attachment_data = getattr(
                    attachment,
                    "data",
                    None
                )

                if not attachment_data:
                    print(
                        "Skipping empty attachment:",
                        name
                    )
                    continue

                attach_path = os.path.join(
                    work_dir,
                    f"{uuid.uuid4()}_{name}"
                )

                with open(
                    attach_path,
                    "wb"
                ) as f:
                    f.write(attachment_data)

                print(
                    "Adding PDF attachment:",
                    name
                )

                merger.append(attach_path)

                temp_files.append(attach_path)

            # ------------------------------------------------
            # Determine final filename
            # ------------------------------------------------

            # For filename searching, use both plain text
            # and HTML stripped of tags.
            plain_body = get_msg_plain_text(msg)

            html_for_search = get_msg_html(msg)

            html_text = re.sub(
                r"<[^>]+>",
                " ",
                html_for_search or ""
            )

            search_body = (
                plain_body
                + "\n"
                + html_text
            )

            output_name = find_output_name(
                subject,
                search_body
            )

            final_path = os.path.join(
                work_dir,
                output_name
            )

            # Avoid collision when two MSG files contain
            # the same DO number.
            if os.path.exists(final_path):

                base_name = os.path.splitext(
                    output_name
                )[0]

                final_path = os.path.join(
                    work_dir,
                    f"{base_name}_{uuid.uuid4().hex[:6]}.pdf"
                )

            # ------------------------------------------------
            # Write final PDF
            # ------------------------------------------------

            merger.write(final_path)

            temp_files.append(final_path)

            print("Created:", final_path)

            return final_path, temp_files

        finally:
            merger.close()

    finally:
        if msg is not None:
            try:
                msg.close()
            except Exception:
                pass


# ============================================================
# FLASK ROUTES
# ============================================================

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():

    if "files" not in request.files:
        return jsonify({
            "error": "No files uploaded"
        }), 400

    files = request.files.getlist("files")

    if not files:
        return jsonify({
            "error": "No files uploaded"
        }), 400

    # --------------------------------------------------------
    # Create unique workspace
    # --------------------------------------------------------

    session_id = str(uuid.uuid4())

    work_dir = os.path.join(
        BASE_FOLDER,
        session_id
    )

    os.makedirs(
        work_dir,
        exist_ok=True
    )

    created_files = []
    all_temp_files = []

    try:

        # ----------------------------------------------------
        # Process each uploaded file
        # ----------------------------------------------------

        for file in files:

            filename = file.filename or ""

            if not filename.lower().endswith(".msg"):
                print(
                    "Skipping non-MSG file:",
                    filename
                )
                continue

            msg_path = os.path.join(
                work_dir,
                f"{uuid.uuid4()}.msg"
            )

            file.save(msg_path)

            all_temp_files.append(msg_path)

            try:

                final_pdf, temp_files = process_msg_file(
                    msg_path,
                    work_dir
                )

                created_files.append(
                    final_pdf
                )

                all_temp_files.extend(
                    temp_files
                )

            except Exception as e:

                print(
                    f"Error processing {filename}: {e}"
                )

                # Continue processing other MSG files
                continue

        # ----------------------------------------------------
        # No successful files
        # ----------------------------------------------------

        if not created_files:

            return jsonify({
                "error": (
                    "No valid MSG files could be processed. "
                    "Check the server console for details."
                )
            }), 400

        # ----------------------------------------------------
        # Create ZIP
        # ----------------------------------------------------

        zip_path = os.path.join(
            work_dir,
            "converted_files.zip"
        )

        with zipfile.ZipFile(
            zip_path,
            "w",
            compression=zipfile.ZIP_DEFLATED
        ) as zipf:

            for pdf in created_files:

                if os.path.exists(pdf):

                    zipf.write(
                        pdf,
                        os.path.basename(pdf)
                    )

        # ----------------------------------------------------
        # Send ZIP
        # ----------------------------------------------------

        response = send_file(
            zip_path,
            as_attachment=True,
            download_name="converted_files.zip",
            mimetype="application/zip"
        )

        # ----------------------------------------------------
        # Cleanup after response
        # ----------------------------------------------------

        @response.call_on_close
        def cleanup():

            try:

                if os.path.exists(work_dir):
                    shutil.rmtree(
                        work_dir,
                        ignore_errors=True
                    )

            except Exception as e:

                print(
                    "Cleanup error:",
                    e
                )

        return response

    except Exception as e:

        print(
            "Upload error:",
            str(e)
        )

        # Cleanup immediately if something failed
        try:
            if os.path.exists(work_dir):
                shutil.rmtree(
                    work_dir,
                    ignore_errors=True
                )
        except Exception:
            pass

        return jsonify({
            "error": str(e)
        }), 500


# ============================================================
# RUN APPLICATION
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=10000,
        debug=True
    )
