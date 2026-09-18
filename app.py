from flask import Flask, render_template, request, send_file, jsonify
import extract_msg
import os
import re
import uuid
import zipfile
import html
import shutil
import traceback

from PyPDF2 import PdfMerger
from weasyprint import HTML
from striprtf.striprtf import rtf_to_text


app = Flask(__name__)

# ============================================================
# CONFIG
# ============================================================

BASE_FOLDER = "workspace"

os.makedirs(BASE_FOLDER, exist_ok=True)


# ============================================================
# GENERAL HELPERS
# ============================================================

def decode_content(value):
    """
    Safely convert MSG content to string.
    """

    if value is None:
        return ""

    if isinstance(value, str):
        return value

    if isinstance(value, bytes):

        # Try UTF-8
        try:
            return value.decode("utf-8")
        except Exception:
            pass

        # Try Windows encoding
        try:
            return value.decode("cp1252")
        except Exception:
            pass

        # Last resort
        return value.decode(
            "latin-1",
            errors="replace"
        )

    return str(value)


def clean_text(text):
    """
    Normalize strange whitespace and null characters.
    """

    if not text:
        return ""

    text = text.replace("\x00", "")

    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")

    # Remove excessive blank lines
    text = re.sub(
        r"\n{4,}",
        "\n\n\n",
        text
    )

    return text.strip()


# ============================================================
# BODY EXTRACTION
# ============================================================

def get_plain_body(msg):
    """
    Get normal plain-text body.
    """

    try:
        body = getattr(msg, "body", None)

        body = decode_content(body)

        return clean_text(body)

    except Exception as e:

        print(
            "Plain body extraction failed:",
            e
        )

        return ""


def get_html_body(msg):
    """
    Get HTML body.
    """

    try:

        body = getattr(
            msg,
            "htmlBody",
            None
        )

        body = decode_content(body)

        return body.strip()

    except Exception as e:

        print(
            "HTML body extraction failed:",
            e
        )

        return ""


def get_rtf_body(msg):
    """
    Get RTF body and convert it to plain text.

    Some Outlook MSG files contain useful content
    only in RTF.
    """

    try:

        rtf = getattr(
            msg,
            "rtfBody",
            None
        )

        if not rtf:
            return ""

        rtf = decode_content(rtf)

        if not rtf:
            return ""

        print(
            "RTF body found, length:",
            len(rtf)
        )

        try:

            text = rtf_to_text(rtf)

            return clean_text(text)

        except Exception as e:

            print(
                "RTF conversion failed:",
                e
            )

            return ""

    except Exception as e:

        print(
            "RTF body extraction failed:",
            e
        )

        return ""


# ============================================================
# QUOTED EMAIL REMOVAL
# ============================================================

def remove_quoted_plain_text(body):
    """
    Remove previous/quoted email from plain text.

    This is deliberately conservative.
    """

    if not body:
        return ""

    # --------------------------------------------------------
    # Outlook Original Message
    # --------------------------------------------------------

    patterns = [

        r"^-{3,}\s*Original Message\s*-{3,}",

        r"^-{3,}\s*Original Appointment\s*-{3,}",

        r"^-{3,}\s*Forwarded message\s*-{3,}",

    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            body,
            flags=re.IGNORECASE |
                  re.MULTILINE
        )

        if match:

            before = body[
                :match.start()
            ].strip()

            if len(before) > 10:
                return before

    # --------------------------------------------------------
    # Detect quoted From/Sent/To/Subject
    # --------------------------------------------------------

    quoted_header = re.compile(
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

    match = quoted_header.search(body)

    if match:

        before = body[
            :match.start()
        ].strip()

        if len(before) > 10:
            return before

    return body.strip()


def html_to_text(html_body):
    """
    Basic HTML -> text conversion.
    Used only for searching and fallback.
    """

    if not html_body:
        return ""

    text = html_body

    # Remove scripts/styles
    text = re.sub(
        r"<script.*?</script>",
        "",
        text,
        flags=re.IGNORECASE |
              re.DOTALL
    )

    text = re.sub(
        r"<style.*?</style>",
        "",
        text,
        flags=re.IGNORECASE |
              re.DOTALL
    )

    # Convert common breaks
    text = re.sub(
        r"<br\s*/?>",
        "\n",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"</p\s*>",
        "\n",
        text,
        flags=re.IGNORECASE
    )

    # Remove remaining tags
    text = re.sub(
        r"<[^>]+>",
        " ",
        text
    )

    # Decode HTML entities
    text = html.unescape(text)

    return clean_text(text)


def remove_quoted_html(body):
    """
    Remove common Outlook quoted HTML sections.
    """

    if not body:
        return ""

    # --------------------------------------------------------
    # Original Message
    # --------------------------------------------------------

    patterns = [

        r"-{3,}\s*Original Message\s*-{3,}",

        r"-{3,}\s*Original Appointment\s*-{3,}",

        r"-{3,}\s*Forwarded message\s*-{3,}",

    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            body,
            flags=re.IGNORECASE
        )

        if match:

            before = body[
                :match.start()
            ].strip()

            if len(before) > 10:
                return before

    # --------------------------------------------------------
    # Gmail-style quote
    # --------------------------------------------------------

    quote_patterns = [

        r'<blockquote[^>]*>',

        r'<div[^>]*class=["\'][^"\']*gmail_quote',

        r'<div[^>]*class=["\'][^"\']*OutlookMessageHeader',

    ]

    for pattern in quote_patterns:

        match = re.search(
            pattern,
            body,
            flags=re.IGNORECASE
        )

        if match:

            before = body[
                :match.start()
            ].strip()

            if len(before) > 10:
                return before

    return body.strip()


# ============================================================
# SELECT BEST BODY
# ============================================================

def extract_email_content(msg):
    """
    Extract the best available body.

    Priority:

        1. HTML
        2. Plain text
        3. RTF

    Returns:

        content
        is_html
        source
    """

    # --------------------------------------------------------
    # HTML
    # --------------------------------------------------------

    html_body = get_html_body(msg)

    if html_body:

        cleaned = remove_quoted_html(
            html_body
        )

        if cleaned:

            print(
                "Using HTML body"
            )

            return (
                cleaned,
                True,
                "HTML"
            )

    # --------------------------------------------------------
    # Plain text
    # --------------------------------------------------------

    plain_body = get_plain_body(msg)

    if plain_body:

        cleaned = remove_quoted_plain_text(
            plain_body
        )

        if cleaned:

            print(
                "Using plain-text body"
            )

            return (
                cleaned,
                False,
                "PLAIN"
            )

    # --------------------------------------------------------
    # RTF
    # --------------------------------------------------------

    rtf_body = get_rtf_body(msg)

    if rtf_body:

        cleaned = remove_quoted_plain_text(
            rtf_body
        )

        if cleaned:

            print(
                "Using RTF body"
            )

            return (
                cleaned,
                False,
                "RTF"
            )

    # --------------------------------------------------------
    # Nothing found
    # --------------------------------------------------------

    return (
        "",
        False,
        "NONE"
    )


# ============================================================
# PDF HTML
# ============================================================

def create_pdf_html(
    content,
    is_html=False
):
    """
    Create HTML document for WeasyPrint.
    """

    if is_html:

        email_content = content

    else:

        escaped = html.escape(
            content
        )

        email_content = f"""
        <pre class="plain-email">{escaped}</pre>
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
    font-family: Arial, Helvetica, sans-serif;
    font-size: 11pt;
    line-height: 1.45;
    color: #111;
    word-wrap: break-word;
    overflow-wrap: break-word;
}}

.plain-email {{
    white-space: pre-wrap;
    word-wrap: break-word;
    overflow-wrap: break-word;
    font-family: Arial, Helvetica, sans-serif;
    font-size: 11pt;
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

</style>

</head>

<body>

{email_content}

</body>

</html>
"""


# ============================================================
# OUTPUT FILE NAME
# ============================================================

def determine_output_name(
    subject,
    plain_body,
    html_body,
    rtf_body
):
    """
    Find DOxx-xxxxx from any available MSG content.
    """

    pattern = re.compile(
        r"\bDO\d{2}-\d{5}\b",
        re.IGNORECASE
    )

    # Search body first
    bodies = [
        plain_body,
        html_to_text(html_body),
        rtf_body,
        subject
    ]

    for body in bodies:

        if not body:
            continue

        match = pattern.search(
            body
        )

        if match:

            number = match.group(
                0
            ).upper()

            return f"{number}.pdf"

    return (
        f"output_"
        f"{uuid.uuid4().hex[:6]}"
        f".pdf"
    )


# ============================================================
# PROCESS ONE MSG
# ============================================================

def process_msg_file(
    msg_path,
    work_dir
):

    temp_files = []

    msg = None

    try:

        print("\n")
        print("=" * 80)
        print(
            "PROCESSING MSG:",
            msg_path
        )
        print("=" * 80)

        # ----------------------------------------------------
        # Open MSG
        # ----------------------------------------------------

        msg = extract_msg.Message(
            msg_path
        )

        # ----------------------------------------------------
        # Subject
        # ----------------------------------------------------

        subject = decode_content(
            getattr(
                msg,
                "subject",
                ""
            )
        )

        print(
            "Subject:",
            repr(subject)
        )

        # ----------------------------------------------------
        # Get ALL body versions
        # ----------------------------------------------------

        plain_body = get_plain_body(
            msg
        )

        html_body = get_html_body(
            msg
        )

        rtf_body = get_rtf_body(
            msg
        )

        print(
            "Plain body length:",
            len(plain_body)
        )

        print(
            "HTML body length:",
            len(html_body)
        )

        print(
            "RTF body length:",
            len(rtf_body)
        )

        # ----------------------------------------------------
        # Select best body
        # ----------------------------------------------------

        content, is_html, source = (
            extract_email_content(msg)
        )

        print(
            "Selected source:",
            source
        )

        print(
            "Selected content length:",
            len(content)
        )

        if not content:

            raise Exception(
                "No readable body found. "
                "MSG may contain unsupported/corrupt content."
            )

        # ----------------------------------------------------
        # Determine output filename
        # ----------------------------------------------------

        output_name = determine_output_name(
            subject,
            plain_body,
            html_body,
            rtf_body
        )

        print(
            "Output filename:",
            output_name
        )

        # ----------------------------------------------------
        # Generate email PDF
        # ----------------------------------------------------

        pdf_html = create_pdf_html(
            content,
            is_html=is_html
        )

        email_pdf = os.path.join(
            work_dir,
            f"{uuid.uuid4()}_email.pdf"
        )

        HTML(
            string=pdf_html,
            base_url=os.path.abspath(
                work_dir
            )
        ).write_pdf(
            email_pdf
        )

        temp_files.append(
            email_pdf
        )

        print(
            "Email PDF created"
        )

        # ----------------------------------------------------
        # Merge attachments
        # ----------------------------------------------------

        merger = PdfMerger()

        try:

            merger.append(
                email_pdf
            )

            attachments = (
                getattr(
                    msg,
                    "attachments",
                    []
                )
                or []
            )

            print(
                "Attachments:",
                len(attachments)
            )

            for attachment in attachments:

                try:

                    name = (
                        getattr(
                            attachment,
                            "longFilename",
                            None
                        )
                        or
                        getattr(
                            attachment,
                            "shortFilename",
                            None
                        )
                    )

                    if not name:
                        continue

                    name = os.path.basename(
                        str(name)
                    )

                    # Only PDF attachments
                    if not name.lower().endswith(
                        ".pdf"
                    ):
                        continue

                    data = getattr(
                        attachment,
                        "data",
                        None
                    )

                    if not data:
                        print(
                            "Empty attachment:",
                            name
                        )
                        continue

                    attachment_path = os.path.join(
                        work_dir,
                        f"{uuid.uuid4()}_{name}"
                    )

                    with open(
                        attachment_path,
                        "wb"
                    ) as f:

                        f.write(data)

                    print(
                        "Merging attachment:",
                        name
                    )

                    # Try to append
                    merger.append(
                        attachment_path
                    )

                    temp_files.append(
                        attachment_path
                    )

                except Exception as attachment_error:

                    # One bad attachment should NOT
                    # prevent the email from being converted.
                    print(
                        "Attachment error:",
                        attachment_error
                    )

                    continue

            # ------------------------------------------------
            # Final output
            # ------------------------------------------------

            final_path = os.path.join(
                work_dir,
                output_name
            )

            # Avoid duplicate filenames
            if os.path.exists(
                final_path
            ):

                base = os.path.splitext(
                    output_name
                )[0]

                final_path = os.path.join(
                    work_dir,
                    f"{base}_"
                    f"{uuid.uuid4().hex[:6]}"
                    f".pdf"
                )

            merger.write(
                final_path
            )

            print(
                "Final PDF created:",
                final_path
            )

            return (
                final_path,
                temp_files
            )

        finally:

            try:
                merger.close()
            except Exception:
                pass

    finally:

        if msg is not None:

            try:
                msg.close()
            except Exception:
                pass


# ============================================================
# HOME
# ============================================================

@app.route("/")
def index():

    return render_template(
        "index.html"
    )


# ============================================================
# UPLOAD
# ============================================================

@app.route(
    "/upload",
    methods=["POST"]
)
def upload():

    if "files" not in request.files:

        return jsonify({
            "error": "No files uploaded"
        }), 400

    files = request.files.getlist(
        "files"
    )

    if not files:

        return jsonify({
            "error": "No files uploaded"
        }), 400

    # --------------------------------------------------------
    # Workspace
    # --------------------------------------------------------

    session_id = str(
        uuid.uuid4()
    )

    work_dir = os.path.join(
        BASE_FOLDER,
        session_id
    )

    os.makedirs(
        work_dir,
        exist_ok=True
    )

    created_files = []

    errors = []

    try:

        # ----------------------------------------------------
        # Process files
        # ----------------------------------------------------

        for file in files:

            filename = (
                file.filename
                or ""
            )

            if not filename.lower().endswith(
                ".msg"
            ):

                errors.append({
                    "file": filename,
                    "error": "Not an MSG file"
                })

                continue

            msg_path = os.path.join(
                work_dir,
                f"{uuid.uuid4()}.msg"
            )

            try:

                file.save(
                    msg_path
                )

                final_pdf, temp_files = (
                    process_msg_file(
                        msg_path,
                        work_dir
                    )
                )

                created_files.append(
                    final_pdf
                )

            except Exception as e:

                print(
                    "\nERROR PROCESSING:",
                    filename
                )

                print(
                    traceback.format_exc()
                )

                errors.append({
                    "file": filename,
                    "error": str(e)
                })

        # ----------------------------------------------------
        # Nothing succeeded
        # ----------------------------------------------------

        if not created_files:

            return jsonify({
                "error": (
                    "No MSG files could be "
                    "successfully converted."
                ),
                "details": errors
            }), 400

        # ----------------------------------------------------
        # ZIP
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

                if os.path.exists(
                    pdf
                ):

                    zipf.write(
                        pdf,
                        os.path.basename(
                            pdf
                        )
                    )

        # ----------------------------------------------------
        # Send response
        # ----------------------------------------------------

        response = send_file(
            zip_path,
            as_attachment=True,
            download_name="converted_files.zip",
            mimetype="application/zip"
        )

        @response.call_on_close
        def cleanup():

            try:

                if os.path.exists(
                    work_dir
                ):

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
            traceback.format_exc()
        )

        try:

            if os.path.exists(
                work_dir
            ):

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
# START
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=10000,
        debug=True
    )
