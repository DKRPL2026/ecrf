
import io
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image,
)

# ============================================================
# STREAMLIT CONFIG
# ============================================================

st.set_page_config(
    page_title="eCRF Document Automation",
    page_icon="📄",
    layout="wide",
)

# ============================================================
# HELPERS
# ============================================================

def format_value(value):
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime("%d/%m/%Y")

    val = str(value).strip().lower()

    if val in ["✔", "✓", "☑", "yes", "y", "true"]:
        return "☑"

    if val in ["no", "n", "false"]:
        return "☐"

    return str(value)


def format_mcq(selected_value, option):
    """Return checked/unchecked symbol for an MCQ option."""
    if selected_value is None:
        return "☐"

    try:
        if pd.isna(selected_value):
            return "☐"
    except Exception:
        pass

    selected_value = str(selected_value).strip()

    if not selected_value:
        return "☐"

    selected_options = re.split(r"[,;|]", selected_value)
    selected_options = [
        x.strip().lower()
        for x in selected_options
        if x.strip()
    ]

    return "☑" if option.strip().lower() in selected_options else "☐"


def process_text(text, data):
    if not text:
        return text

    # Normal placeholders
    for key, value in data.items():
        placeholder = "{{" + str(key) + "}}"
        text = text.replace(
            placeholder,
            format_value(value)
        )

    # MCQ placeholders:
    # {{MCQ:Gender:Male}}
    # {{MCQ:Symptoms:Fever}}
    mcq_pattern = r"\{\{MCQ:(.*?):(.*?)\}\}"

    def mcq_replace(match):
        column_name = match.group(1).strip()
        option = match.group(2).strip()
        selected_value = data.get(column_name, "")

        return format_mcq(
            selected_value,
            option
        )

    return re.sub(
        mcq_pattern,
        mcq_replace,
        text
    )


def process_paragraph(paragraph, data):
    if not paragraph.text:
        return

    # Do not overwrite paragraphs containing images/drawings.
    if any(
        run.element.xpath(".//w:drawing")
        for run in paragraph.runs
    ):
        return

    original_text = paragraph.text
    new_text = process_text(original_text, data)

    if new_text != original_text:
        paragraph.text = new_text


def process_table(table, data):
    for row in table.rows:
        for cell in row.cells:

            for paragraph in cell.paragraphs:
                process_paragraph(paragraph, data)

            for nested_table in cell.tables:
                process_table(nested_table, data)


def replace_placeholders(doc, data):
    # Main document
    for paragraph in doc.paragraphs:
        process_paragraph(paragraph, data)

    for table in doc.tables:
        process_table(table, data)

    # Headers and footers
    for section in doc.sections:

        for paragraph in section.header.paragraphs:
            process_paragraph(paragraph, data)

        for table in section.header.tables:
            process_table(table, data)

        for paragraph in section.footer.paragraphs:
            process_paragraph(paragraph, data)

        for table in section.footer.tables:
            process_table(table, data)


def safe_filename(value):
    value = str(value).strip()

    value = re.sub(
        r'[<>:"/\\|?*]',
        "_",
        value
    )

    return value or "document"


def convert_docx_to_pdf(docx_path, pdf_path):
    """
    Convert DOCX to PDF using LibreOffice.
    This is suitable for Linux/Docker/Render.
    """

    output_dir = os.path.dirname(pdf_path)

    result = subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            output_dir,
            docx_path,
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )

    expected_pdf = os.path.join(
        output_dir,
        Path(docx_path).stem + ".pdf"
    )

    if result.returncode != 0:
        raise RuntimeError(
            "LibreOffice conversion failed: "
            + (result.stderr or result.stdout)
        )

    if not os.path.exists(expected_pdf):
        raise RuntimeError(
            "PDF was not created by LibreOffice. "
            + (result.stderr or result.stdout)
        )

    if os.path.abspath(expected_pdf) != os.path.abspath(pdf_path):
        os.replace(expected_pdf, pdf_path)


def create_audit_report(summary, details, output_path):
    doc = SimpleDocTemplate(
        output_path,
        pagesize=landscape(A4),
        rightMargin=10 * mm,
        leftMargin=10 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
    )

    styles = getSampleStyleSheet()
    elements = []

    elements.append(
        Paragraph(
            "eCRF Document Automation Audit Report",
            styles["Title"]
        )
    )

    elements.append(Spacer(1, 8))

    for key, value in summary.items():

        if isinstance(value, (datetime, pd.Timestamp)):
            value = value.strftime("%d/%m/%Y %H:%M:%S")

        elements.append(
            Paragraph(
                f"<b>{key}:</b> {value}",
                styles["Normal"]
            )
        )

    elements.append(Spacer(1, 12))

    table_data = [
        [
            "Row",
            "Subject ID",
            "File",
            "Status",
            "Error",
        ]
    ]

    for record in details:
        table_data.append(
            [
                str(record.get("row", "")),
                str(record.get("subject_id", "")),
                str(record.get("file", "")),
                str(record.get("status", "")),
                str(record.get("error", "")),
            ]
        )

    table = Table(
        table_data,
        repeatRows=1,
        colWidths=[
            15 * mm,
            30 * mm,
            75 * mm,
            25 * mm,
            110 * mm,
        ],
    )

    table.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.grey,
                ),
                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.white,
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.black,
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "TOP",
                ),
                (
                    "FONTSIZE",
                    (0, 0),
                    (-1, -1),
                    7,
                ),
            ]
        )
    )

    elements.append(table)
    elements.append(Spacer(1, 15))

    elements.append(
        Paragraph(
            "<b>Authorized By:</b>",
            styles["Normal"]
        )
    )

    signature_path = os.path.join(
        os.path.dirname(output_path),
        "signature.png"
    )

    if os.path.exists(signature_path):
        elements.append(
            Image(
                signature_path,
                width=45 * mm,
                height=18 * mm,
            )
        )
    else:
        elements.append(
            Paragraph(
                "Signature not available",
                styles["Normal"]
            )
        )

    elements.append(
        Paragraph(
            "Name: Linu Thommi",
            styles["Normal"]
        )
    )

    elements.append(
        Paragraph(
            "Role: CTO",
            styles["Normal"]
        )
    )

    elements.append(
        Paragraph(
            f"Date: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
            styles["Normal"]
        )
    )

    elements.append(Spacer(1, 10))

    elements.append(
        Paragraph(
            "System-generated audit report.",
            styles["Italic"]
        )
    )

    doc.build(elements)


def create_zip_package(
    excel_bytes,
    excel_filename,
    template_bytes,
    template_filename,
    generated_files,
    audit_report_path,
    signature_bytes=None,
):
    package_buffer = io.BytesIO()

    with zipfile.ZipFile(
        package_buffer,
        "w",
        zipfile.ZIP_DEFLATED
    ) as z:

        # Generated PDFs
        for pdf_path in generated_files:
            z.write(
                pdf_path,
                arcname=os.path.join(
                    "outputs",
                    os.path.basename(pdf_path)
                )
            )

        # Audit report
        z.write(
            audit_report_path,
            arcname="logs/audit_report.pdf"
        )

        # Input Excel
        z.writestr(
            os.path.join(
                "uploads",
                safe_filename(excel_filename)
            ),
            excel_bytes
        )

        # Input Word template
        z.writestr(
            os.path.join(
                "uploads",
                safe_filename(template_filename)
            ),
            template_bytes
        )

        # Signature if uploaded
        if signature_bytes:
            z.writestr(
                "assets/signature.png",
                signature_bytes
            )

    package_buffer.seek(0)
    return package_buffer.getvalue()


def process_records(
    df,
    selected_indices,
    template_bytes,
    work_dir,
):
    output_dir = os.path.join(work_dir, "outputs")
    os.makedirs(output_dir, exist_ok=True)

    generated_files = []
    records = []

    start_time = datetime.now()
    success = 0
    failure = 0

    for index in selected_indices:

        data = df.iloc[index].to_dict()

        status = "Success"
        error = ""
        pdf_filename = ""
        subject_id = ""

        try:
            raw_subject_id = data.get("subject_id")

            if raw_subject_id is None or pd.isna(raw_subject_id):
                raise Exception(
                    f"Missing subject_id at Excel row {index + 2}"
                )

            subject_id = str(raw_subject_id).strip()

            if not subject_id:
                raise Exception(
                    f"Missing subject_id at Excel row {index + 2}"
                )

            subject_id = safe_filename(subject_id)

            timestamp = datetime.now().strftime(
                "%Y%m%d_%H%M%S_%f"
            )

            pdf_filename = (
                f"{subject_id}_{timestamp}.pdf"
            )

            pdf_path = os.path.join(
                output_dir,
                pdf_filename
            )

            # Create a fresh DOCX for each record
            template_path = os.path.join(
                work_dir,
                "template.docx"
            )

            with open(template_path, "wb") as f:
                f.write(template_bytes)

            doc = Document(template_path)

            replace_placeholders(
                doc,
                data
            )

            temp_docx = os.path.join(
                work_dir,
                f"{subject_id}_{timestamp}.docx"
            )

            doc.save(temp_docx)

            convert_docx_to_pdf(
                temp_docx,
                pdf_path
            )

            if not os.path.exists(pdf_path):
                raise Exception(
                    "PDF conversion failed."
                )

            generated_files.append(pdf_path)
            success += 1

        except Exception as e:
            status = "Failed"
            error = str(e)
            failure += 1

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
                "error": error,
            }
        )

    end_time = datetime.now()

    summary = {
        "User": st.session_state.get(
            "username",
            "Unknown User"
        ),
        "Start Time": start_time,
        "End Time": end_time,
        "Total Selected Records": len(selected_indices),
        "Success": success,
        "Failure": failure,
    }

    return summary, records, generated_files


# ============================================================
# STREAMLIT UI
# ============================================================

st.title("📄 eCRF Document Automation System")

st.caption(
    "Excel → Word Template → Individual PDFs → Audit Report → ZIP"
)

with st.sidebar:
    st.header("User Information")

    username = st.text_input(
        "Username",
        value="",
        placeholder="Enter your name"
    )

    st.session_state["username"] = (
        username.strip()
        if username.strip()
        else "Unknown User"
    )

    st.divider()

    st.info(
        "Word placeholders:\n"
        "{{subject_id}}\n\n"
        "MCQ placeholders:\n"
        "{{MCQ:Gender:Male}}\n"
        "{{MCQ:Gender:Female}}"
    )

st.subheader("1. Upload Files")

col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader(
        "Upload Excel file",
        type=["xlsx", "xls"],
        key="excel_upload"
    )

with col2:
    template_file = st.file_uploader(
        "Upload Word template",
        type=["docx"],
        key="template_upload"
    )

st.subheader("2. Optional Signature")

signature_file = st.file_uploader(
    "Upload signature PNG",
    type=["png"],
    key="signature_upload"
)

if excel_file:

    try:
        excel_bytes = excel_file.getvalue()

        df = pd.read_excel(
            io.BytesIO(excel_bytes)
        )

        st.success(
            f"Excel loaded successfully: "
            f"{len(df)} records, {len(df.columns)} columns."
        )

        if "subject_id" not in df.columns:
            st.error(
                "Excel must contain a 'subject_id' column."
            )
            st.stop()

        st.subheader("3. Excel Data Preview")

        st.dataframe(
            df,
            use_container_width=True,
            height=300,
        )

        st.subheader("4. Select Records")

        # Add a selection column without modifying original data
        display_df = df.copy()

        display_df.insert(
            0,
            "Select",
            True
        )

        selected_table = st.data_editor(
            display_df,
            use_container_width=True,
            height=400,
            hide_index=True,
            column_config={
                "Select": st.column_config.CheckboxColumn(
                    "Select",
                    default=True,
                )
            },
            disabled=[
                col
                for col in display_df.columns
                if col != "Select"
            ],
            key="record_selector",
        )

        selected_indices = selected_table.index[
            selected_table["Select"]
        ].tolist()

        st.write(
            f"**Selected records: {len(selected_indices)} "
            f"/ {len(df)}**"
        )

        if selected_indices:
            selected_subjects = [
                str(df.iloc[i]["subject_id"])
                for i in selected_indices
            ]

            with st.expander(
                "View selected Subject IDs"
            ):
                st.write(
                    ", ".join(selected_subjects)
                )

        st.subheader("5. Generate Documents")

        can_process = (
            template_file is not None
            and len(selected_indices) > 0
        )

        if not template_file:
            st.warning(
                "Please upload a Word template."
            )

        if len(selected_indices) == 0:
            st.warning(
                "Please select at least one record."
            )

        if st.button(
            "🚀 Generate Documents",
            type="primary",
            disabled=not can_process,
            use_container_width=True,
        ):

            with st.spinner(
                "Generating documents. Please wait..."
            ):

                with tempfile.TemporaryDirectory() as work_dir:

                    signature_bytes = (
                        signature_file.getvalue()
                        if signature_file
                        else None
                    )

                    summary, records, generated_files = (
                        process_records(
                            df=df,
                            selected_indices=selected_indices,
                            template_bytes=template_file.getvalue(),
                            work_dir=work_dir,
                        )
                    )

                    # Prepare audit report
                    audit_dir = os.path.join(
                        work_dir,
                        "logs"
                    )
                    os.makedirs(
                        audit_dir,
                        exist_ok=True
                    )

                    # Copy signature for audit report
                    if signature_bytes:
                        signature_path = os.path.join(
                            audit_dir,
                            "signature.png"
                        )

                        with open(
                            signature_path,
                            "wb"
                        ) as f:
                            f.write(signature_bytes)

                    audit_report_path = os.path.join(
                        audit_dir,
                        "audit_report.pdf"
                    )

                    create_audit_report(
                        summary,
                        records,
                        audit_report_path,
                    )

                    zip_bytes = create_zip_package(
                        excel_bytes=excel_bytes,
                        excel_filename=excel_file.name,
                        template_bytes=template_file.getvalue(),
                        template_filename=template_file.name,
                        generated_files=generated_files,
                        audit_report_path=audit_report_path,
                        signature_bytes=signature_bytes,
                    )

                    st.session_state[
                        "generated_zip"
                    ] = zip_bytes

                    st.session_state[
                        "summary"
                    ] = summary

                    st.session_state[
                        "records"
                    ] = records

                    st.session_state[
                        "generated_count"
                    ] = len(generated_files)

            st.success(
                "Document generation completed."
            )

    except Exception as error:
        st.error(f"An error occurred while processing the Excel file: {error}")

# ============================================================
# RESULTS
# ============================================================

if "summary" in st.session_state:

    st.subheader("6. Processing Summary")

    summary = st.session_state["summary"]

    c1, c2, c3 = st.columns(3)

    with c1:
        st.metric(
            "Selected",
            summary["Total Selected Records"]
        )

    with c2:
        st.metric(
            "Success",
            summary["Success"]
        )

    with c3:
        st.metric(
            "Failure",
            summary["Failure"]
        )

    st.subheader("7. Audit Results")

    result_df = pd.DataFrame(
        st.session_state["records"]
    )

    st.dataframe(
        result_df,
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("8. Download Package")

    st.download_button(
        label="⬇️ Download ZIP Package",
        data=st.session_state["generated_zip"],
        file_name=(
            f"eCRF_document_package_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        ),
        mime="application/zip",
        type="primary",
        use_container_width=True,
    )

    st.caption(
        "The ZIP contains generated PDFs, audit_report.pdf, "
        "and the original uploaded Excel/Word files."
    )
