import os
import re
import io
import uuid
import shutil
import zipfile
import tempfile
import subprocess
import platform
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st

from datetime import datetime

from docx import Document
from docx.oxml.ns import qn

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image
)

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="eCRF Document Automation System",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>

    /* ------------------------------------------------------
       MAIN APP
    ------------------------------------------------------ */

    .stApp {
        background-color: #0e1117;
        color: #ffffff;
    }

    .main .block-container {
        padding-top: 2.5rem;
        padding-bottom: 3rem;
        max-width: 100%;
    }

    /* ------------------------------------------------------
       HEADINGS
    ------------------------------------------------------ */

    h1 {
        color: #f5f5f5 !important;
        font-size: 30px !important;
        font-weight: 700 !important;
    }

    h2 {
        color: #f5f5f5 !important;
        font-size: 22px !important;
    }

    h3 {
        color: #f5f5f5 !important;
        font-size: 18px !important;
    }

    /* ------------------------------------------------------
       FILE UPLOADERS
    ------------------------------------------------------ */

    [data-testid="stFileUploader"] {
        background-color: #262934;
        border-radius: 6px;
        padding: 4px;
    }

    [data-testid="stFileUploaderDropzone"] {
        background-color: #262934;
        border: 1px solid #343741;
    }

    /* ------------------------------------------------------
       BUTTONS
    ------------------------------------------------------ */

    .stButton > button {
        width: 100%;
        border-radius: 5px;
        border: 1px solid #2e9f5b;
        background-color: #238636;
        color: #ffffff;
        min-height: 38px;
        font-weight: 500;
    }

    .stButton > button:hover {
        border-color: #3fbf6f;
        background-color: #2ea043;
        color: #ffffff;
    }

    /* ------------------------------------------------------
       GENERATE BUTTON
    ------------------------------------------------------ */

    div[data-testid="stButton"] button[kind="primary"] {
        background-color: #238636;
        border-color: #238636;
        color: white;
        font-weight: 600;
    }

    div[data-testid="stButton"] button[kind="primary"]:hover {
        background-color: #2ea043;
        border-color: #2ea043;
    }

    /* ------------------------------------------------------
       DOWNLOAD BUTTON
    ------------------------------------------------------ */

    div[data-testid="stDownloadButton"] button {
        background-color: #238636;
        border-color: #238636;
        color: white;
        font-weight: 600;
        width: 100%;
    }

    div[data-testid="stDownloadButton"] button:hover {
        background-color: #2ea043;
        border-color: #2ea043;
    }

    /* ------------------------------------------------------
       DATAFRAME
    ------------------------------------------------------ */

    [data-testid="stDataFrame"] {
        border: 1px solid #30333d;
        border-radius: 5px;
    }

    /* ------------------------------------------------------
       SUCCESS MESSAGE
    ------------------------------------------------------ */

    [data-testid="stAlert"] {
        border-radius: 5px;
    }

    /* ------------------------------------------------------
       METRIC BOX
    ------------------------------------------------------ */

    .summary-box {
        background-color: #171a21;
        border: 1px solid #30333d;
        border-radius: 6px;
        padding: 14px;
        text-align: center;
        margin-bottom: 10px;
    }

    .summary-title {
        color: #9da3af;
        font-size: 13px;
    }

    .summary-value {
        color: #ffffff;
        font-size: 24px;
        font-weight: 700;
        margin-top: 5px;
    }

    /* ------------------------------------------------------
       PACKAGE STRUCTURE
    ------------------------------------------------------ */

    .package-box {
        background-color: #15181f;
        border: 1px solid #30333d;
        border-radius: 6px;
        padding: 15px 20px;
        font-family: monospace;
        color: #d9dce3;
        line-height: 1.8;
    }

    /* ------------------------------------------------------
       SMALL TEXT
    ------------------------------------------------------ */

    .small-text {
        color: #9da3af;
        font-size: 12px;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# SESSION STATE INITIALIZATION
# ============================================================

DEFAULT_STATE = {
    "excel_bytes": None,
    "excel_filename": None,

    "template_bytes": None,
    "template_filename": None,

    "dataframe": None,

    "selected_rows": [],

    "generated": False,
    "zip_bytes": None,
    "zip_filename": None,

    "summary": None,
    "audit_records": [],

    "generation_id": None,

    "authenticated": False,
    "logged_in_user": None,
}


for key, value in DEFAULT_STATE.items():

    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# AUTHENTICATION CONFIGURATION
# Keep this section easy to edit for future migration to
# Streamlit secrets, password hashes, database auth, or SSO.
# ============================================================

AUTHORIZED_USERS = {
    "da@drklinisch.com": "DKRPL@SM2021",
    "sitemanagement7@drklinisch.com": "DKRPL@SM2021",
    "sitemanagement@drklinisch.com": "DKRPL@SM2021",
    "sitemanagement12@drklinisch.com": "DKRPL@SM2021",
    "ClinicalCoordinator@drklinisch.com": "DKRPL@SM2021",
}


def normalize_email(email):

    return str(email or "").strip().lower()


def authenticate_user(email, password):

    normalized_email = normalize_email(email)

    if not normalized_email:
        return False

    expected_password = AUTHORIZED_USERS.get(normalized_email)

    if expected_password is None:
        return False

    return expected_password == password


def find_project_logo():

    project_root = os.path.dirname(__file__)
    candidate_extensions = {".png", ".jpg", ".jpeg", ".svg", ".webp"}

    for root, _, files in os.walk(project_root):

        for filename in files:

            lowered_name = filename.lower()

            if (
                lowered_name.startswith("logo")
                or "company" in lowered_name
                or "brand" in lowered_name
            ) and os.path.splitext(lowered_name)[1] in candidate_extensions:

                return os.path.join(root, filename)

    return None


def logout():

    st.session_state["authenticated"] = False
    st.session_state["logged_in_user"] = None
    st.session_state.pop("login_email", None)
    st.session_state.pop("login_password", None)
    st.session_state.pop("login_error", None)


def show_login_page():

    st.markdown(
        """
        <style>
        .login-title {
            text-align: center;
            color: #f5f5f5;
            font-size: 26px;
            font-weight: 700;
            margin: 0 0 1rem 0;
            line-height: 1.2;
        }

        div[data-testid="stForm"] {
            max-width: 460px;
            width: 100%;
            margin: 0 auto;
            padding: 1.25rem 1.15rem;
            background: #171a21;
            border: 1px solid #30333d;
            border-radius: 12px;
            box-shadow: 0 10px 24px rgba(0, 0, 0, 0.18);
        }

        div[data-testid="stForm"] .stTextInput > div,
        div[data-testid="stForm"] input {
            width: 100%;
        }

        div[data-testid="stForm"] .stButton > button {
            display: block;
            margin: 0 auto;
            min-width: 120px;
            width: auto;
            border-radius: 6px;
            border: 1px solid #2e9f5b;
            background-color: #238636;
            color: #ffffff;
            font-weight: 600;
            padding: 0.55rem 1.2rem;
        }

        div[data-testid="stForm"] .stButton > button:hover {
            border-color: #3fbf6f;
            background-color: #2ea043;
            color: #ffffff;
        }

        .login-footer {
            margin-top: 1.25rem;
            text-align: center;
            color: #9da3af;
            font-size: 11px;
            letter-spacing: 0.02em;
            line-height: 1.5;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="login-title">eCRF Document Automation System</div>',
        unsafe_allow_html=True,
    )

    with st.form("login_form", clear_on_submit=False):

        email = st.text_input(
            "Email ID",
            key="login_email",
            placeholder="user@company.com",
        )

        password = st.text_input(
            "Password",
            type="password",
            key="login_password",
            placeholder="Enter your password",
        )

        submitted = st.form_submit_button(
            "Login",
            use_container_width=True,
        )

        if submitted:

            if authenticate_user(email, password):

                st.session_state["authenticated"] = True
                st.session_state["logged_in_user"] = normalize_email(email)
                st.session_state.pop("login_error", None)
                st.success("Login successful.")
                st.rerun()

            else:

                st.session_state["login_error"] = "Invalid user credentials."
                st.error("Invalid user credentials.")

    st.markdown(
        '<div class="login-footer">Copyright ©2026 Dr. Klinisch  All Rights Reserved Version 1.0.</div>',
        unsafe_allow_html=True,
    )


# ============================================================
# UTILITY FUNCTIONS
# ============================================================

PERSISTENT_STORAGE_DIRECTORIES = [
    os.path.join(
        os.path.dirname(__file__),
        directory_name
    )
    for directory_name in [
        "uploads",
        "downloads",
        "output"
    ]
]

USER_MANUAL_PATH = os.path.join(
    os.path.dirname(__file__),
    "manuals",
    "eCRF User Manual.pdf"
)


def clear_persistent_storage():

    for directory_path in PERSISTENT_STORAGE_DIRECTORIES:

        if not os.path.isdir(directory_path):
            continue

        for entry in os.scandir(directory_path):

            try:

                if entry.is_dir(follow_symlinks=False):
                    shutil.rmtree(
                        entry.path,
                        ignore_errors=True
                    )

                else:
                    os.remove(
                        entry.path
                    )

            except OSError:
                pass

def safe_filename(value):

    value = str(value).strip()

    value = re.sub(
        r'[<>:"/\\|?*]',
        "_",
        value
    )

    value = value.strip(". ")

    if not value:
        value = "file"

    return value


# ============================================================
# VALUE FORMATTER
# ============================================================

def format_value(value):

    if value is None:
        return ""

    try:

        if pd.isna(value):
            return ""

    except Exception:
        pass

    if isinstance(
        value,
        (datetime, pd.Timestamp)
    ):

        return value.strftime(
            "%d/%m/%Y"
        )

    val = str(value).strip().lower()

    if val in [
        "✔",
        "✓",
        "☑"
    ]:

        return "☑"

    if val in [
        "yes",
        "y",
        "true"
    ]:

        return "☑"

    if val in [
        "no",
        "n",
        "false"
    ]:

        return "☐"

    return str(value)


# ============================================================
# MCQ FORMATTER
# ============================================================

def format_mcq(
    selected_value,
    option
):

    if selected_value is None:
        return "☐"

    try:

        if pd.isna(selected_value):
            return "☐"

    except Exception:
        pass

    selected_value = str(
        selected_value
    ).strip()

    if not selected_value:
        return "☐"

    option = str(
        option
    ).strip()

    selected_options = re.split(
        r"[,;|]",
        selected_value
    )

    selected_options = [
        x.strip().lower()
        for x in selected_options
        if x.strip()
    ]

    if option.lower() in selected_options:

        return "☑"

    return "☐"


# ============================================================
# PROCESS TEXT
# ============================================================

def process_text(
    text,
    data
):

    if not text:
        return text

    # --------------------------------------------------------
    # NORMAL PLACEHOLDERS
    # --------------------------------------------------------

    for key, value in data.items():

        placeholder = (
            "{{"
            + str(key)
            + "}}"
        )

        formatted_value = format_value(
            value
        )

        text = text.replace(
            placeholder,
            formatted_value
        )

    # --------------------------------------------------------
    # MCQ PLACEHOLDERS
    # --------------------------------------------------------

    mcq_pattern = (
        r"\{\{MCQ:(.*?):(.*?)\}\}"
    )

    def mcq_replace(match):

        column_name = (
            match.group(1)
            .strip()
        )

        option = (
            match.group(2)
            .strip()
        )

        selected_value = data.get(
            column_name,
            ""
        )

        return format_mcq(
            selected_value,
            option
        )

    text = re.sub(
        mcq_pattern,
        mcq_replace,
        text
    )

    return text


# ============================================================
# PROCESS WORD PARAGRAPH
# ============================================================

def process_paragraph(
    paragraph,
    data
):

    if not paragraph.text:
        return

    original_text = paragraph.text

    new_text = process_text(
        original_text,
        data
    )

    if new_text == original_text:
        return

    # --------------------------------------------------------
    # Try to preserve run formatting when placeholder
    # is completely contained in a single run.
    # --------------------------------------------------------

    changed = False

    for run in paragraph.runs:

        if not run.text:
            continue

        processed_run_text = process_text(
            run.text,
            data
        )

        if processed_run_text != run.text:

            run.text = processed_run_text
            changed = True

    if changed:
        return

    # --------------------------------------------------------
    # Fallback for placeholders split across runs
    # --------------------------------------------------------

    paragraph.text = new_text


# ============================================================
# PROCESS TABLE
# ============================================================

def process_table(
    table,
    data
):

    for row in table.rows:

        for cell in row.cells:

            for paragraph in cell.paragraphs:

                process_paragraph(
                    paragraph,
                    data
                )

            for nested_table in cell.tables:

                process_table(
                    nested_table,
                    data
                )


# ============================================================
# REPLACE WORD PLACEHOLDERS
# ============================================================

def replace_placeholders(
    doc,
    data
):

    # --------------------------------------------------------
    # MAIN PARAGRAPHS
    # --------------------------------------------------------

    for paragraph in doc.paragraphs:

        process_paragraph(
            paragraph,
            data
        )

    # --------------------------------------------------------
    # MAIN TABLES
    # --------------------------------------------------------

    for table in doc.tables:

        process_table(
            table,
            data
        )

    # --------------------------------------------------------
    # HEADERS AND FOOTERS
    # --------------------------------------------------------

    for section in doc.sections:

        # Header

        for paragraph in section.header.paragraphs:

            process_paragraph(
                paragraph,
                data
            )

        for table in section.header.tables:

            process_table(
                table,
                data
            )

        # Footer

        for paragraph in section.footer.paragraphs:

            process_paragraph(
                paragraph,
                data
            )

        for table in section.footer.tables:

            process_table(
                table,
                data
            )


# ============================================================
# DOCX → PDF
# ============================================================

def render_docx_as_pdf(
    docx_path,
    pdf_path
):

    docx_document = Document(
        docx_path
    )

    pdf_document = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm
    )

    font_name = "Helvetica"

    for font_path in (
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "C:/Windows/Fonts/DejaVuSans.ttf",
        "C:/Windows/Fonts/arial.ttf"
    ):

        if os.path.exists(font_path):

            try:

                pdfmetrics.registerFont(
                    TTFont(
                        "TemplateUnicodeFont",
                        font_path
                    )
                )

                font_name = "TemplateUnicodeFont"
                break

            except Exception:
                continue

    styles = getSampleStyleSheet()

    body_style = ParagraphStyle(
        "GeneratedDocumentBody",
        parent=styles["BodyText"],
        fontName=font_name,
        fontSize=10,
        leading=14,
        spaceAfter=7
    )

    cell_style = ParagraphStyle(
        "GeneratedDocumentCell",
        parent=body_style,
        fontName=font_name,
        fontSize=8,
        leading=10,
        spaceAfter=0
    )

    elements = []

    def paragraph_text(text):

        if font_name == "Helvetica":

            text = (text or "").replace("☑", "[X]").replace("☐", "[ ]")

        return Paragraph(
            escape(text or "").replace("\n", "<br/>") or "<br/>",
            body_style
        )

    for paragraph in docx_document.paragraphs:

        if paragraph.text.strip():
            elements.append(
                paragraph_text(
                    paragraph.text
                )
            )

    available_width = A4[0] - (36 * mm)

    embedded_images = []

    for inline_shape in docx_document.inline_shapes:

        blip_elements = inline_shape._inline.xpath(
            ".//a:blip"
        )

        if not blip_elements:
            continue

        relationship_id = blip_elements[0].get(
            qn("r:embed")
        )

        image_part = docx_document.part.related_parts.get(
            relationship_id
        )

        if image_part is None:
            continue

        image_width = inline_shape.width / 12700
        image_height = inline_shape.height / 12700

        if image_width <= 0 or image_height <= 0:
            continue

        embedded_images.append(
            Image(
                io.BytesIO(image_part.blob),
                width=image_width,
                height=image_height
            )
        )

    for docx_table in docx_document.tables:

        row_count = len(docx_table.rows)

        column_count = max(
            (len(row.cells) for row in docx_table.rows),
            default=1
        )

        table_data = []

        for row in docx_table.rows:

            cells = []

            for cell in row.cells:

                cell_text = "\n".join(
                    paragraph.text
                    for paragraph in cell.paragraphs
                )

                if font_name == "Helvetica":

                    cell_text = (
                        cell_text
                        .replace("☑", "[X]")
                        .replace("☐", "[ ]")
                    )

                cells.append(
                    Paragraph(
                        escape(cell_text).replace("\n", "<br/>") or "<br/>",
                        cell_style
                    )
                )

            table_data.append(
                cells + [""] * (column_count - len(cells))
            )

        if table_data:

            table = Table(
                table_data,
                colWidths=[
                    available_width / column_count
                ] * column_count,
                repeatRows=1 if row_count > 1 else 0
            )

            table.setStyle(
                TableStyle(
                    [
                        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 4),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                        ("TOPPADDING", (0, 0), (-1, -1), 4),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 4)
                    ]
                )
            )

            elements.append(table)
            elements.append(Spacer(1, 8))

    for embedded_image in embedded_images:

        embedded_image._restrictSize(
            available_width,
            A4[1] - (36 * mm)
        )

        elements.append(
            embedded_image
        )

        elements.append(
            Spacer(1, 8)
        )

    if not elements:
        elements.append(
            paragraph_text(
                "The generated document is empty."
            )
        )

    pdf_document.build(
        elements
    )

def convert_docx_to_pdf(
    docx_path,
    pdf_path
):

    output_dir = os.path.dirname(
        pdf_path
    )

    os.makedirs(
        output_dir,
        exist_ok=True
    )

    # --------------------------------------------------------
    # METHOD 1: docx2pdf on Windows
    # --------------------------------------------------------

    if platform.system().lower() == "windows":

        try:

            from docx2pdf import convert

            convert(
                docx_path,
                pdf_path
            )

            if os.path.exists(pdf_path):

                return

        except Exception as e:

            windows_error = str(e)

        else:

            windows_error = (
                "DOCX to PDF conversion did not "
                "create the PDF file."
            )

    # --------------------------------------------------------
    # METHOD 2: LibreOffice / soffice
    # --------------------------------------------------------

    soffice_candidates = []

    for candidate in (
        "soffice",
        "libreoffice",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/usr/lib/libreoffice/program/soffice"
    ):

        resolved_candidate = shutil.which(candidate) or candidate

        if resolved_candidate not in soffice_candidates:
            soffice_candidates.append(resolved_candidate)

    conversion_errors = []

    for soffice in soffice_candidates:

        try:

            result = subprocess.run(
                [
                    soffice,
                    "--headless",
                    "-env:UserInstallation=file:///tmp/lo-profile",
                    "--convert-to",
                    "pdf:writer_pdf_Export",
                    "--outdir",
                    output_dir,
                    docx_path
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=120
            )

            generated_name = (
                os.path.splitext(
                    os.path.basename(docx_path)
                )[0]
                + ".pdf"
            )

            generated_pdf = os.path.join(
                output_dir,
                generated_name
            )

            if os.path.exists(
                generated_pdf
            ):

                if os.path.abspath(
                    generated_pdf
                ) != os.path.abspath(
                    pdf_path
                ):

                    shutil.move(
                        generated_pdf,
                        pdf_path
                    )

                return

        except (
            FileNotFoundError,
            subprocess.SubprocessError
        ) as e:
            conversion_errors.append(f"{soffice}: {e}")
            continue

        if result.returncode != 0:
            conversion_errors.append(
                f"{soffice}: {result.stderr.strip() or result.stdout.strip()}"
            )

    # --------------------------------------------------------
    # METHOD 3: Pure Python fallback for Python-only hosts
    # --------------------------------------------------------

    try:

        render_docx_as_pdf(
            docx_path,
            pdf_path
        )

        if os.path.exists(pdf_path):
            return

    except Exception as e:

        conversion_errors.append(
            f"ReportLab fallback: {e}"
        )

    # --------------------------------------------------------
    # If everything failed
    # --------------------------------------------------------

    if platform.system().lower() == "windows":

        raise Exception(
            "PDF conversion failed. "
            "Microsoft Word/docx2pdf or LibreOffice "
            "could not be found. "
            f"Details: {windows_error}"
        )

    raise Exception(
        "PDF conversion failed. "
        "LibreOffice (soffice) is not available or could not convert the DOCX "
        "on the server. "
        f"Details: {'; '.join(conversion_errors) or 'no executable was found.'}"
    )


# ============================================================
# PDF AUDIT REPORT
# ============================================================

def create_pdf_report(
    summary,
    details,
    full_path
):

    # --------------------------------------------------------
    # LANDSCAPE A4
    # This prevents the audit table from overlapping.
    # --------------------------------------------------------

    doc = SimpleDocTemplate(
        full_path,
        pagesize=landscape(A4),
        rightMargin=10 * mm,
        leftMargin=10 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "AuditTitle",
        parent=styles["Title"],
        fontSize=18,
        leading=22,
        alignment=TA_CENTER,
        spaceAfter=12
    )

    normal_style = ParagraphStyle(
        "AuditNormal",
        parent=styles["Normal"],
        fontSize=9,
        leading=12
    )

    small_style = ParagraphStyle(
        "AuditSmall",
        parent=styles["Normal"],
        fontSize=7,
        leading=9
    )

    header_style = ParagraphStyle(
        "AuditHeader",
        parent=styles["Normal"],
        fontSize=7,
        leading=8,
        textColor=colors.white,
        alignment=TA_CENTER
    )

    cell_style = ParagraphStyle(
        "AuditCell",
        parent=styles["Normal"],
        fontSize=7,
        leading=8,
        alignment=TA_LEFT
    )

    elements = []

    # --------------------------------------------------------
    # TITLE
    # --------------------------------------------------------

    elements.append(
        Paragraph(
            "eCRF Document Automation Audit Report",
            title_style
        )
    )

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    for key, value in summary.items():

        if isinstance(
            value,
            (datetime, pd.Timestamp)
        ):

            value = value.strftime(
                "%d/%m/%Y %H:%M:%S"
            )

        elements.append(
            Paragraph(
                f"<b>{key}:</b> "
                f"{str(value)}",
                normal_style
            )
        )

    elements.append(
        Spacer(1, 12)
    )

    # --------------------------------------------------------
    # DETAIL TABLE
    # --------------------------------------------------------

    table_data = [
        [
            Paragraph(
                "Row",
                header_style
            ),

            Paragraph(
                "Subject ID",
                header_style
            ),

            Paragraph(
                "File",
                header_style
            ),

            Paragraph(
                "Status",
                header_style
            ),

            Paragraph(
                "Error",
                header_style
            )
        ]
    ]

    for record in details:

        table_data.append(
            [
                Paragraph(
                    str(
                        record.get(
                            "row",
                            ""
                        )
                    ),
                    cell_style
                ),

                Paragraph(
                    str(
                        record.get(
                            "subject_id",
                            ""
                        )
                    ),
                    cell_style
                ),

                Paragraph(
                    str(
                        record.get(
                            "file",
                            ""
                        )
                    ),
                    cell_style
                ),

                Paragraph(
                    str(
                        record.get(
                            "status",
                            ""
                        )
                    ),
                    cell_style
                ),

                Paragraph(
                    str(
                        record.get(
                            "error",
                            ""
                        )
                    ),
                    cell_style
                )
            ]
        )

    # --------------------------------------------------------
    # COLUMN WIDTHS
    # --------------------------------------------------------

    col_widths = [
        18 * mm,
        35 * mm,
        95 * mm,
        28 * mm,
        90 * mm
    ]

    table = Table(
        table_data,
        colWidths=col_widths,
        repeatRows=1
    )

    table.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.grey
                ),

                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.white
                ),

                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.black
                ),

                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "TOP"
                ),

                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                )
            ]
        )
    )

    elements.append(
        table
    )

    elements.append(
        Spacer(1, 20)
    )

    # --------------------------------------------------------
    # AUTHORIZATION
    # --------------------------------------------------------

    elements.append(
        Paragraph(
            "<b>Authorized By:</b>",
            normal_style
        )
    )

    elements.append(
        Spacer(1, 6)
    )

    elements.append(
        Paragraph(
            "Name: Linu Thommi",
            normal_style
        )
    )

    elements.append(
        Paragraph(
            "Role: CTO",
            normal_style
        )
    )

    elements.append(
        Paragraph(
            "Date: "
            + datetime.now().strftime(
                "%d/%m/%Y %H:%M:%S"
            ),
            normal_style
        )
    )

    elements.append(
        Spacer(1, 12)
    )

    elements.append(
        Paragraph(
            "System-generated audit report.",
            small_style
        )
    )

    # --------------------------------------------------------
    # BUILD
    # --------------------------------------------------------

    doc.build(
        elements
    )


# ============================================================
# GENERATE PACKAGE
# ============================================================

def generate_package(
    excel_bytes,
    excel_filename,
    template_bytes,
    template_filename,
    selected_indices
):

    process_start = datetime.now()

    # --------------------------------------------------------
    # CREATE TEMP WORK DIRECTORY
    # --------------------------------------------------------

    work_dir = tempfile.mkdtemp(
        prefix="ecrf_"
    )

    try:

        # ----------------------------------------------------
        # INPUT DIRECTORY
        # ----------------------------------------------------

        input_dir = os.path.join(
            work_dir,
            "Input_Files"
        )

        word_dir = os.path.join(
            work_dir,
            "Word_Documents"
        )

        pdf_dir = os.path.join(
            work_dir,
            "PDF_Documents"
        )

        audit_dir = os.path.join(
            work_dir,
            "Audit_Report"
        )

        os.makedirs(
            input_dir,
            exist_ok=True
        )

        os.makedirs(
            word_dir,
            exist_ok=True
        )

        os.makedirs(
            pdf_dir,
            exist_ok=True
        )

        os.makedirs(
            audit_dir,
            exist_ok=True
        )

        # ----------------------------------------------------
        # SAVE INPUT FILES
        # ----------------------------------------------------

        excel_input_path = os.path.join(
            input_dir,
            safe_filename(
                excel_filename
            )
        )

        template_input_path = os.path.join(
            input_dir,
            safe_filename(
                template_filename
            )
        )

        with open(
            excel_input_path,
            "wb"
        ) as f:

            f.write(
                excel_bytes
            )

        with open(
            template_input_path,
            "wb"
        ) as f:

            f.write(
                template_bytes
            )

        # ----------------------------------------------------
        # READ EXCEL
        # ----------------------------------------------------

        df = pd.read_excel(
            io.BytesIO(
                excel_bytes
            )
        )

        # ----------------------------------------------------
        # SUBJECT ID CHECK
        # ----------------------------------------------------

        if "subject_id" not in df.columns:

            raise Exception(
                "Excel must contain "
                "'subject_id' column."
            )

        records = []

        generated_word_files = []

        generated_pdf_files = []

        success = 0
        failure = 0

        # ----------------------------------------------------
        # PROCESS SELECTED ROWS ONLY
        # ----------------------------------------------------

        for index in selected_indices:

            row = df.iloc[index]

            data = row.to_dict()

            status = "Success"

            error = ""

            subject_id = ""

            word_filename = ""

            pdf_filename = ""

            try:

                # --------------------------------------------
                # SUBJECT ID
                # --------------------------------------------

                raw_subject_id = data.get(
                    "subject_id"
                )

                if (
                    raw_subject_id is None
                    or pd.isna(raw_subject_id)
                ):

                    raise Exception(
                        f"Missing subject_id "
                        f"at Excel row {index + 2}"
                    )

                subject_id = str(
                    raw_subject_id
                ).strip()

                if not subject_id:

                    raise Exception(
                        f"Missing subject_id "
                        f"at Excel row {index + 2}"
                    )

                subject_id = safe_filename(
                    subject_id
                )

                # --------------------------------------------
                # DATE AND TIME
                # --------------------------------------------

                timestamp = datetime.now().strftime(
                    "%Y%m%d_%H%M%S"
                )

                base_filename = (
                    f"{subject_id}_"
                    f"{timestamp}"
                )

                # --------------------------------------------
                # CHECK DUPLICATE FILE NAME
                # --------------------------------------------

                word_filename = (
                    f"{base_filename}.docx"
                )

                pdf_filename = (
                    f"{base_filename}.pdf"
                )

                counter = 1

                while (
                    os.path.exists(
                        os.path.join(
                            word_dir,
                            word_filename
                        )
                    )
                    or
                    os.path.exists(
                        os.path.join(
                            pdf_dir,
                            pdf_filename
                        )
                    )
                ):

                    word_filename = (
                        f"{base_filename}_"
                        f"{counter:02d}.docx"
                    )

                    pdf_filename = (
                        f"{base_filename}_"
                        f"{counter:02d}.pdf"
                    )

                    counter += 1

                # --------------------------------------------
                # WORD DOCUMENT
                # --------------------------------------------

                doc = Document(
                    template_input_path
                )

                replace_placeholders(
                    doc,
                    data
                )

                word_path = os.path.join(
                    word_dir,
                    word_filename
                )

                doc.save(
                    word_path
                )

                # --------------------------------------------
                # PDF DOCUMENT
                # --------------------------------------------

                pdf_path = os.path.join(
                    pdf_dir,
                    pdf_filename
                )

                convert_docx_to_pdf(
                    word_path,
                    pdf_path
                )

                if not os.path.exists(
                    pdf_path
                ):

                    raise Exception(
                        "PDF conversion failed."
                    )

                generated_word_files.append(
                    word_path
                )

                generated_pdf_files.append(
                    pdf_path
                )

                success += 1

            except Exception as e:

                status = "Failed"

                error = str(e)

                failure += 1

            # -----------------------------------------------
            # AUDIT RECORD
            # -----------------------------------------------

            records.append(
                {
                    "row": index + 2,

                    "subject_id": subject_id,

                    "file": (
                        pdf_filename
                        if status == "Success"
                        else ""
                    ),

                    "status": status,

                    "error": error
                }
            )

        # ----------------------------------------------------
        # END TIME
        # ----------------------------------------------------

        process_end = datetime.now()

        # ----------------------------------------------------
        # SUMMARY
        # ----------------------------------------------------

        summary = {
            "Start Time": process_start,

            "End Time": process_end,

            "Total Selected Records": len(
                selected_indices
            ),

            "Success": success,

            "Failure": failure
        }

        # ----------------------------------------------------
        # AUDIT REPORT
        # ----------------------------------------------------

        audit_path = os.path.join(
            audit_dir,
            "audit_report.pdf"
        )

        create_pdf_report(
            summary,
            records,
            audit_path
        )

        # ----------------------------------------------------
        # ZIP FILE
        # ----------------------------------------------------

        package_timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        zip_filename = (
            f"eCRF_Document_Package_"
            f"{package_timestamp}.zip"
        )

        zip_path = os.path.join(
            work_dir,
            zip_filename
        )

        # ----------------------------------------------------
        # CREATE ZIP
        # ----------------------------------------------------

        with zipfile.ZipFile(
            zip_path,
            "w",
            zipfile.ZIP_DEFLATED
        ) as z:

            for root, dirs, files in os.walk(
                work_dir
            ):

                # Do not include the ZIP itself

                if root == work_dir:

                    files = [
                        f
                        for f in files
                        if f != zip_filename
                    ]

                for file in files:

                    full_path = os.path.join(
                        root,
                        file
                    )

                    relative_path = os.path.relpath(
                        full_path,
                        work_dir
                    )

                    z.write(
                        full_path,
                        relative_path
                    )

        # ----------------------------------------------------
        # READ ZIP INTO MEMORY
        # ----------------------------------------------------

        with open(
            zip_path,
            "rb"
        ) as f:

            zip_bytes = f.read()

        return (
            zip_bytes,
            zip_filename,
            summary,
            records
        )

    finally:

        # ----------------------------------------------------
        # CLEAN TEMP FILES
        # ----------------------------------------------------

        shutil.rmtree(
            work_dir,
            ignore_errors=True
        )


# ============================================================
# RESET GENERATED PACKAGE
# ============================================================

def clear_generated_package():

    st.session_state.generated = False

    st.session_state.zip_bytes = None

    st.session_state.zip_filename = None

    st.session_state.summary = None

    st.session_state.audit_records = []

    st.session_state.generation_id = None


if not st.session_state.get("authenticated", False):

    show_login_page()
    st.stop()


# ============================================================
# HEADER
# ============================================================

header_col, logout_col = st.columns([5, 1])

with header_col:

    st.markdown(
        """
        <div style="
            font-size:30px;
            font-weight:700;
            color:#f5f5f5;
        ">
            eCRF Document Automation System
        </div>
        """,
        unsafe_allow_html=True
    )

    st.caption(
        f"Logged in as: {st.session_state.get('logged_in_user', 'Unknown user')}"
    )

with logout_col:

    if st.button(
        "Logout",
        use_container_width=True,
        key="logout_button"
    ):

        logout()
        st.rerun()

if os.path.isfile(USER_MANUAL_PATH):

    manual_col, _ = st.columns([1, 5])

    with manual_col:

        with open(USER_MANUAL_PATH, "rb") as manual_file:

            st.download_button(
                label="Download User Manual",
                data=manual_file.read(),
                file_name="eCRF_User_Manual.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="download_user_manual_button",
                on_click="ignore"
            )


# ============================================================
# 1. UPLOAD FILES
# ============================================================

st.markdown(
    "### 1. Upload Files"
)

col1, col2 = st.columns(
    2,
    gap="medium"
)

with col1:

    excel_file = st.file_uploader(
        "Upload Excel File",
        type=[
            "xlsx",
            "xls"
        ],
        key="excel_uploader"
    )

with col2:

    template_file = st.file_uploader(
        "Upload Word Template",
        type=[
            "docx"
        ],
        key="template_uploader"
    )


# ============================================================
# STORE UPLOADS
# ============================================================

if excel_file is not None:

    new_excel_bytes = excel_file.getvalue()

    if (
        st.session_state.excel_filename
        != excel_file.name
        or
        st.session_state.excel_bytes
        != new_excel_bytes
    ):

        st.session_state.excel_filename = (
            excel_file.name
        )

        st.session_state.excel_bytes = (
            new_excel_bytes
        )

        # Read Excel

        try:

            loaded_df = pd.read_excel(
                io.BytesIO(
                    new_excel_bytes
                )
            )

            st.session_state.dataframe = (
                loaded_df
            )

            st.session_state.selected_rows = []

            clear_generated_package()

        except Exception as e:

            st.error(
                "Unable to read Excel file: "
                + str(e)
            )

            st.session_state.dataframe = None


if template_file is not None:

    new_template_bytes = (
        template_file.getvalue()
    )

    if (
        st.session_state.template_filename
        != template_file.name
        or
        st.session_state.template_bytes
        != new_template_bytes
    ):

        st.session_state.template_filename = (
            template_file.name
        )

        st.session_state.template_bytes = (
            new_template_bytes
        )

        clear_generated_package()


# ============================================================
# FILE VALIDATION
# ============================================================

if (
    st.session_state.dataframe is not None
    and
    st.session_state.excel_bytes is not None
):

    df = st.session_state.dataframe

    st.success(
        f"Excel loaded successfully: "
        f"{len(df)} records, "
        f"{len(df.columns)} columns."
    )


# ============================================================
# 3. EXCEL DATA PREVIEW
# ============================================================

if st.session_state.dataframe is not None:

    df = st.session_state.dataframe

    st.markdown(
        "### 3. Excel Data Preview"
    )

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=False,
        height=250
    )


# ============================================================
# 4. SELECT RECORDS
# ============================================================

if st.session_state.dataframe is not None:

    df = st.session_state.dataframe.copy()

    st.markdown(
        "### 4. Select Records"
    )

    # --------------------------------------------------------
    # Create selection dataframe
    # --------------------------------------------------------

    selection_df = df.copy()

    selection_df.insert(
        0,
        "Select",
        selection_df.index.isin(
            st.session_state.selected_rows
        )
    )

    action_col1, action_col2 = st.columns(
        2,
        gap="small"
    )

    with action_col1:

        if st.button(
            "Select All",
            use_container_width=True,
            key="select_all_records_button"
        ):

            st.session_state.selected_rows = (
                df.index.tolist()
            )

            st.rerun()

    with action_col2:

        if st.button(
            "Clear",
            use_container_width=True,
            key="clear_selected_records_button"
        ):

            st.session_state.selected_rows = []

            st.rerun()

    # --------------------------------------------------------
    # Unique editor key based on uploaded Excel
    # --------------------------------------------------------

    excel_key = (
        str(
            hash(
                st.session_state.excel_bytes
            )
        )
    )

    editor_key = (
        "record_selector_"
        + excel_key
    )

    edited_df = st.data_editor(
        selection_df,
        use_container_width=True,
        hide_index=False,
        num_rows="fixed",
        key=editor_key,
        column_config={
            "Select": st.column_config.CheckboxColumn(
                "Select",
                help="Select records to generate",
                default=False
            )
        }
    )

    # --------------------------------------------------------
    # Selected rows
    # --------------------------------------------------------

    selected_mask = (
        edited_df["Select"]
        == True
    )

    selected_indices = (
        edited_df.index[
            selected_mask
        ].tolist()
    )

    st.session_state.selected_rows = (
        selected_indices
    )

    selected_count = len(
        selected_indices
    )

    total_count = len(
        df
    )

    st.markdown(
        f"""
        <div style="
            color:#f5f5f5;
            font-weight:600;
            margin-top:8px;
            margin-bottom:8px;
        ">
            Selected records: {selected_count}/{total_count}
        </div>
        """,
        unsafe_allow_html=True
    )

    # --------------------------------------------------------
    # Show selected subject IDs
    # --------------------------------------------------------

    with st.expander(
        "View selected Subject IDs"
    ):

        if selected_count == 0:

            st.write(
                "No records selected."
            )

        else:

            selected_subjects = []

            for idx in selected_indices:

                selected_subjects.append(
                    str(
                        df.iloc[idx][
                            "subject_id"
                        ]
                    )
                )

            st.write(
                ", ".join(
                    selected_subjects
                )
            )


# ============================================================
# 5. GENERATE DOCUMENTS
# ============================================================

st.markdown(
    "### 5. Generate Documents"
)

can_generate = (
    st.session_state.excel_bytes is not None
    and
    st.session_state.template_bytes is not None
    and
    st.session_state.dataframe is not None
    and
    len(
        st.session_state.selected_rows
    ) > 0
)


if not can_generate:

    st.info(
        "Upload an Excel file, upload a Word template, "
        "and select at least one record."
    )


if st.button(
    "⚡ Generate Documents",
    type="primary",
    use_container_width=True,
    disabled=not can_generate,
    key="generate_documents_button"
):

    # --------------------------------------------------------
    # Clear old package before new generation
    # --------------------------------------------------------

    clear_generated_package()

    try:

        with st.spinner(
            "Generating Word documents, PDFs and audit report..."
        ):

            (
                zip_bytes,
                zip_filename,
                summary,
                audit_records
            ) = generate_package(

                st.session_state.excel_bytes,

                st.session_state.excel_filename,

                st.session_state.template_bytes,

                st.session_state.template_filename,

                st.session_state.selected_rows
            )

        clear_persistent_storage()

        # ----------------------------------------------------
        # IMPORTANT:
        # Store everything in session state.
        # Therefore download does not depend on the
        # st.button() state.
        # ----------------------------------------------------

        st.session_state.zip_bytes = (
            zip_bytes
        )

        st.session_state.zip_filename = (
            zip_filename
        )

        st.session_state.summary = (
            summary
        )

        st.session_state.audit_records = (
            audit_records
        )

        st.session_state.generated = True

        st.session_state.generation_id = (
            uuid.uuid4().hex
        )

        st.success(
            "Documents generated successfully."
        )

    except Exception as e:

        st.error(
            "An error occurred while processing "
            "the Excel file: "
            + str(e)
        )


# ============================================================
# RESULTS
# ============================================================

if (
    st.session_state.generated
    and
    st.session_state.zip_bytes is not None
):

    st.markdown(
        "---"
    )

    st.markdown(
        "### 6. Generation Summary"
    )

    summary = (
        st.session_state.summary
    )

    # --------------------------------------------------------
    # SUMMARY METRICS
    # --------------------------------------------------------

    c1, c2, c3, c4 = st.columns(
        4
    )

    with c1:

        st.markdown(
            f"""
            <div class="summary-box">
                <div class="summary-title">
                    Selected Records
                </div>
                <div class="summary-value">
                    {summary["Total Selected Records"]}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c2:

        st.markdown(
            f"""
            <div class="summary-box">
                <div class="summary-title">
                    Successful
                </div>
                <div class="summary-value">
                    {summary["Success"]}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c3:

        st.markdown(
            f"""
            <div class="summary-box">
                <div class="summary-title">
                    Failed
                </div>
                <div class="summary-value">
                    {summary["Failure"]}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c4:

        st.markdown(
            f"""
            <div class="summary-box">
                <div class="summary-title">
                    Package
                </div>
                <div class="summary-value">
                    ZIP
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )


    # ========================================================
    # DOWNLOAD ZIP
    # ========================================================

    st.markdown(
        "### 7. Download Package"
    )

    st.download_button(

        label="Download ZIP Package",

        data=st.session_state.zip_bytes,

        file_name=st.session_state.zip_filename,

        mime="application/zip",

        key=(
            "download_zip_"
            + str(
                st.session_state.generation_id
            )
        ),

        type="primary",

        use_container_width=True,

        # IMPORTANT:
        # This prevents Streamlit from rerunning
        # the application when the user downloads.
        on_click="ignore"
    )


    st.caption(
        "The ZIP contains the Word documents, "
        "PDF documents, audit report and input files. "
        "The audit report does not need to be downloaded separately."
    )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div style="
        text-align:center;
        color:#9da3af;
        font-size:12px;
        margin-top:40px;
        padding:16px 0;
    ">
        Copyright ©2026 Dr. Klinisch. All Rights Reserved Version 1.0.
    </div>
    """,
    unsafe_allow_html=True
)