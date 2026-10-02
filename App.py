import io
import json
import os
import re
import pandas as pd
from PIL import Image
from pydantic import BaseModel, Field
import streamlit as st
from google import genai
from google.genai import types

st.set_page_config(page_title="College Marksheet Consolidation", layout="wide")

st.title("🎓 Student Marksheet Consolidated Report")
st.write(
    "Upload up to 50 result screenshots. Each student will get a single row "
    "showing their marks under actual subject names, Grand Total, and list of failed subjects."
)

st.sidebar.header("🔑 AI Settings")
api_key = st.sidebar.text_input(
    "Gemini API Key",
    type="password",
    value=os.environ.get("GEMINI_API_KEY", ""),
    help="Get an API key from Google AI Studio",
)

class SubjectEntry(BaseModel):
    sub_code: str = Field(description="Subject code, e.g., 43A, 43B")
    subject_name: str = Field(description="Full name of subject")
    marks: str = Field(description="Marks string e.g. 018+048")
    result: str = Field(description="'P' or 'F'")

class StudentResult(BaseModel):
    register_number: str = Field(description="Student Register Number")
    student_name: str = Field(description="Student Full Name")
    subjects: list[SubjectEntry] = Field(description="All subjects listed on the marksheet")

def parse_marks(mark_str: str) -> tuple[int, int]:
    """Parse '018+048' into (total_marks, status_flag)."""
    try:
        parts = mark_str.replace(" ", "").split("+")
        if len(parts) == 2:
            return int(parts[0]) + int(parts[1])
        elif len(parts) == 1 and parts[0].isdigit():
            return int(parts[0])
    except Exception:
        pass
    return 0

def extract_result(client: genai.Client, pil_img: Image.Image) -> dict:
    prompt = (
        "Extract register number, student name, and all subject rows (sub code, subject name, marks, and result). "
        "Return strictly valid JSON matching schema."
    )
    
    # Convert image to bytes to ensure safe transfer
    img_byte_arr = io.BytesIO()
    pil_img.save(img_byte_arr, format="JPEG")
    img_bytes = img_byte_arr.getvalue()

    res = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=[
            prompt,
            types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg"),
        ],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=StudentResult,
            temperature=0.0,
        ),
    )
    return json.loads(res.text)

uploaded_files = st.file_uploader(
    "Upload marksheet screenshots (up to 50)",
    type=["png", "jpg", "jpeg", "webp"],
    accept_multiple_files=True,
)

if uploaded_files:
    if len(uploaded_files) > 50:
        st.warning("⚠️ Processing capped at the first 50 files.")
        uploaded_files = uploaded_files[:50]

    if not api_key.strip():
        st.warning("⚠️ Enter your Gemini API key in the sidebar to proceed.")
    else:
        if st.button("🚀 Process Marksheets & Generate Excel", type="primary"):
            client = genai.Client(api_key=api_key.strip())
            progress_bar = st.progress(0)
            status_text = st.empty()

            all_records = []
            all_discovered_subjects = []

            for idx, f in enumerate(uploaded_files):
                status_text.text(f"Extracting ({idx+1}/{len(uploaded_files)}): {f.name}")
                try:
                    img = Image.open(f).convert("RGB")
                    data = extract_result(client, img)

                    reg_no = data.get("register_number", "Unknown")
                    name = data.get("student_name", "Unknown")
                    subjects = data.get("subjects", [])

                    total_marks = 0
                    failed_subjects = []
                    row_data = {
                        "Register Number": reg_no,
                        "Student Name": name,
                    }

                    for sub in subjects:
                        code = sub.get("sub_code", "").strip()
                        s_name = sub.get("subject_name", "").strip()
                        mark_str = sub.get("marks", "").strip()
                        res_val = "P" if "P" in sub.get("result", "").upper() else "F"

                        subject_col_name = f"{s_name} ({code})" if code else s_name
                        if subject_col_name not in all_discovered_subjects:
                            all_discovered_subjects.append(subject_col_name)

                        sub_total = parse_marks(mark_str)
                        total_marks += sub_total

                        # Show mark string along with fail marker if failed
                        if res_val == "F":
                            failed_subjects.append(f"{s_name} ({code})")
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

            status_text.text("Extraction complete!")

            # Reorder columns neatly: Reg No, Name, [All Subjects], Total, Result, Failed Subjects, Arrears
            final_columns = ["Register Number", "Student Name"] + all_discovered_subjects + [
                "Total Marks",
                "Result",
                "Failed Subjects",
                "Arrear Count",
            ]

            df = pd.DataFrame(all_records)
            # Ensure any missing columns across different images are safely filled with '-'
            for c in final_columns:
                if c not in df.columns:
                    df[c] = "-"
            df = df[final_columns]

            st.subheader("📋 Class Consolidated Marks")
            st.dataframe(df, use_container_width=True)

            # Export Excel
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Consolidated_Marks", index=False)
                ws = writer.sheets["Consolidated_Marks"]

                for col in ws.columns:
                    header = str(col[0].value or "")
                    max_len = max(len(str(cell.value or "")) for cell in col)
                    col_letter = col[0].column_letter
                    if "Subject" in header or len(header) > 20:
                        ws.column_dimensions[col_letter].width = max(max_len + 3, 25)
                    else:
                        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

            st.download_button(
                label="📥 Download Consolidated Marksheet (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name="Class_Consolidated_Marksheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
