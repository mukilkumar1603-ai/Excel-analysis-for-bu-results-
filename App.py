import io
import json
import os
import time
import pandas as pd
from PIL import Image
from pydantic import BaseModel, Field
import streamlit as st
from docx import Document
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

from google import genai
from google.genai import types

st.set_page_config(page_title="Universal Marksheet Converter", layout="wide")

st.title("🎓 Multi-Format Marksheet Analyzer & Converter")
st.write(
    "Upload up to 50 student marksheets in **Image (JPG/PNG)**, **PDF**, or **Word (.docx)** format. "
    "Consolidate all student records and export directly to **Excel (.xlsx)**, **Word (.docx)**, or **PDF**."
)

st.sidebar.header("🔑 AI Settings")
api_key = st.sidebar.text_input(
    "Gemini API Key",
    type="password",
    value=os.environ.get("GEMINI_API_KEY", ""),
    help="Get an API key from Google AI Studio (aistudio.google.com)",
)

available_models = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"]
if api_key.strip():
    try:
        temp_client = genai.Client(api_key=api_key.strip())
        fetched_models = [
            m.name.replace("models/", "")
            for m in temp_client.models.list()
            if "generateContent" in (m.supported_actions or [])
        ]
        if fetched_models:
            available_models = fetched_models
    except Exception:
        pass

selected_model = st.sidebar.selectbox("Gemini Model", available_models, index=0)

class SubjectEntry(BaseModel):
    sub_code: str = Field(description="Subject code, e.g., 43A, 43B")
    subject_name: str = Field(description="Full name of subject")
    marks: str = Field(description="Marks string e.g. 018+048 or total mark")
    result: str = Field(description="'P' or 'F'")

class StudentResult(BaseModel):
    register_number: str = Field(description="Student Register Number")
    student_name: str = Field(description="Student Full Name")
    subjects: list[SubjectEntry] = Field(description="All subjects listed on the marksheet")

def parse_marks(mark_str: str) -> int:
    try:
        parts = mark_str.replace(" ", "").split("+")
        if len(parts) == 2:
            return int(parts[0]) + int(parts[1])
        elif len(parts) == 1 and parts[0].isdigit():
            return int(parts[0])
    except Exception:
        pass
    return 0

def extract_from_gemini(client: genai.Client, file_bytes: bytes, mime_type: str, model_name: str) -> dict:
    prompt = (
        "Extract register number, student name, and all subject rows (sub code, subject name, marks, and result). "
        "Return strictly valid JSON matching schema."
    )
    res = client.models.generate_content(
        model=model_name,
        contents=[
            prompt,
            types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
        ],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=StudentResult,
            temperature=0.0,
        ),
    )
    return json.loads(res.text)

def extract_from_docx_text(client: genai.Client, file_bytes: bytes, model_name: str) -> dict:
    doc = Document(io.BytesIO(file_bytes))
    full_text = []
    for p in doc.paragraphs:
        if p.text.strip():
            full_text.append(p.text)
    for table in doc.tables:
        for row in table.rows:
            full_text.append(" | ".join([cell.text.strip() for cell in row.cells]))
    raw_content = "\n".join(full_text)

    prompt = (
        "Extract register number, student name, and all subject rows (sub code, subject name, marks, and result) "
        f"from this document text:\n\n{raw_content}\n\nReturn strictly valid JSON matching schema."
    )
    res = client.models.generate_content(
        model=model_name,
        contents=[prompt],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=StudentResult,
            temperature=0.0,
        ),
    )
    return json.loads(res.text)

# File uploader supporting Images, PDF, and Word docs
uploaded_files = st.file_uploader(
    "Upload files (Photos, PDFs, or Word .docx up to 50)",
    type=["png", "jpg", "jpeg", "webp", "pdf", "docx"],
    accept_multiple_files=True,
)

if uploaded_files:
    if len(uploaded_files) > 50:
        st.warning("⚠️ Processing capped at first 50 files.")
        uploaded_files = uploaded_files[:50]

    if not api_key.strip():
        st.warning("⚠️ Enter your Gemini API key in the sidebar to proceed.")
    else:
        if st.button("🚀 Process Marksheets & Generate Reports", type="primary"):
            client = genai.Client(api_key=api_key.strip())
            progress_bar = st.progress(0)
            status_text = st.empty()

            all_records = []
            all_discovered_subjects = []

            for idx, f in enumerate(uploaded_files):
                status_text.text(f"Extracting ({idx+1}/{len(uploaded_files)}): {f.name}")
                try:
                    f_bytes = f.read()
                    name_lower = f.name.lower()

                    if name_lower.endswith((".jpg", ".jpeg", ".png", ".webp")):
                        pil_img = Image.open(io.BytesIO(f_bytes)).convert("RGB")
                        pil_img.thumbnail((1600, 1600))
                        buf = io.BytesIO()
                        pil_img.save(buf, format="JPEG", quality=85)
                        data = extract_from_gemini(client, buf.getvalue(), "image/jpeg", selected_model)
                    elif name_lower.endswith(".pdf"):
                        data = extract_from_gemini(client, f_bytes, "application/pdf", selected_model)
                    elif name_lower.endswith(".docx"):
                        data = extract_from_docx_text(client, f_bytes, selected_model)
                    else:
                        raise ValueError("Unsupported format")

                    reg_no = data.get("register_number", "Unknown")
                    s_name = data.get("student_name", "Unknown")
                    subjects = data.get("subjects", [])

                    total_marks = 0
                    failed_subjects = []
                    row_data = {"Register Number": reg_no, "Student Name": s_name}

                    for sub in subjects:
                        code = sub.get("sub_code", "").strip()
                        sub_title = sub.get("subject_name", "").strip()
                        mark_str = sub.get("marks", "").strip()
                        res_val = "P" if "P" in sub.get("result", "").upper() else "F"

                        subject_col_name = f"{sub_title} ({code})" if code else sub_title
                        if subject_col_name not in all_discovered_subjects:
                            all_discovered_subjects.append(subject_col_name)

                        sub_total = parse_marks(mark_str)
                        total_marks += sub_total

                        if res_val == "F":
                            failed_subjects.append(f"{sub_title} ({code})")
                            row_data[subject_col_name] = f"{sub_total} [FAIL]"
                        else:
                            row_data[subject_col_name] = sub_total

                    row_data["Total Marks"] = total_marks
                    row_data["Result"] = "PASS" if len(failed_subjects) == 0 else "FAIL"
                    row_data["Arrear Count"] = len(failed_subjects)
                    row_data["Failed Subjects"] = ", ".join(failed_subjects) if failed_subjects else "None"

                    all_records.append(row_data)

                except Exception as e:
                    all_records.append({
                        "Register Number": "Error",
                        "Student Name": f.name,
                        "Result": f"Error: {str(e)}",
                        "Total Marks": 0,
                        "Failed Subjects": "Error",
                        "Arrear Count": 0,
                    })

                progress_bar.progress((idx + 1) / len(uploaded_files))
                time.sleep(1)  # Keeps queries within rate limits

            status_text.text("Extraction complete!")

            final_columns = ["Register Number", "Student Name"] + all_discovered_subjects + [
                "Total Marks",
                "Result",
                "Failed Subjects",
                "Arrear Count",
            ]

            df = pd.DataFrame(all_records)
            for c in final_columns:
                if c not in df.columns:
                    df[c] = "-"
            df = df[final_columns]

            st.subheader("📋 Class Consolidated Report")
            st.dataframe(df, use_container_width=True)

            # ----------------- 1. EXCEL EXPORT -----------------
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Consolidated_Report", index=False)
                ws = writer.sheets["Consolidated_Report"]
                for col in ws.columns:
                    header = str(col[0].value or "")
                    max_len = max(len(str(cell.value or "")) for cell in col)
                    col_letter = col[0].column_letter
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 15)

            # ----------------- 2. WORD (.DOCX) EXPORT -----------------
            doc_buffer = io.BytesIO()
            doc = Document()
            doc.add_heading("Class Consolidated Marksheet Report", level=1)
            t = doc.add_table(rows=1, cols=len(final_columns))
            t.style = "Table Grid"
            hdr_cells = t.rows[0].cells
            for i, col_name in enumerate(final_columns):
                hdr_cells[i].text = col_name

            for _, row in df.iterrows():
                row_cells = t.add_row().cells
                for i, col_name in enumerate(final_columns):
                    row_cells[i].text = str(row[col_name])
            doc.save(doc_buffer)

            # ----------------- 3. PDF EXPORT -----------------
            pdf_buffer = io.BytesIO()
            doc_pdf = SimpleDocTemplate(pdf_buffer, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
            elements = []
            styles = getSampleStyleSheet()
            elements.append(Paragraph("Class Consolidated Marksheet Report", styles['Title']))
            elements.append(Spacer(1, 10))

            cell_style = ParagraphStyle(name='CellText', fontSize=7, leading=9)
            pdf_data = []
            pdf_data.append([Paragraph(f"<b>{c}</b>", cell_style) for c in final_columns])
            for _, row in df.iterrows():
                pdf_data.append([Paragraph(str(row[c]), cell_style) for c in final_columns])

            pdf_table = Table(pdf_data, repeatRows=1)
            pdf_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#4B6584")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ]))
            elements.append(pdf_table)
            doc_pdf.build(elements)

            st.write("### 📥 Download Report in Any Format:")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.download_button(
                    label="📊 Download Excel (.xlsx)",
                    data=excel_buffer.getvalue(),
                    file_name="Class_Consolidated_Marksheet.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with col2:
                st.download_button(
                    label="📄 Download Word (.docx)",
                    data=doc_buffer.getvalue(),
                    file_name="Class_Consolidated_Marksheet.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            with col3:
                st.download_button(
                    label="📑 Download PDF (.pdf)",
                    data=pdf_buffer.getvalue(),
                    file_name="Class_Consolidated_Marksheet.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
