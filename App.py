import io
import json
import os
import numpy as np
import pandas as pd
from PIL import Image, ImageStat
from PIL.ExifTags import TAGS
from pydantic import BaseModel, Field
import streamlit as st
from google import genai
from google.genai import types

st.set_page_config(page_title="AI Batch Image Analyzer (50 Pics)", layout="wide")

st.title("🤖 50 Photo AI Vision & Metadata Analyzer")
st.write(
    "Upload up to 50 photos. The app extracts technical attributes, EXIF data, "
    "and queries Gemini for semantic descriptions, keywords, and safety checks before exporting to Excel."
)

# Sidebar for Configuration
st.sidebar.header("🔑 AI Settings")
api_key_input = st.sidebar.text_input(
    "Gemini API Key",
    type="password",
    value=os.environ.get("GEMINI_API_KEY", ""),
    help="Get an API key from Google AI Studio. You can also export GEMINI_API_KEY in your shell.",
)

model_name = st.sidebar.selectbox(
    "Vision Model",
    options=["gemini-2.5-flash", "gemini-2.0-flash"],
    index=0,
)

# File uploader (up to 50 files)
uploaded_files = st.file_uploader(
    "Choose up to 50 images (JPG, PNG, WEBP)",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
)

# Pydantic schema for structured AI vision output
class ImageAIAnalysis(BaseModel):
    caption: str = Field(description="A 1-2 sentence detailed description of the scene.")
    subject_category: str = Field(description="Main category, e.g. Landscape, Portrait, Architecture, Document, Product, Animal.")
    tags: list[str] = Field(description="5 to 8 relevant semantic tags or keywords.")
    safety_flag: str = Field(description="'Safe' or specific concern like 'Explicit', 'Violent', 'PII/Sensitive'.")

def get_exif_data(image: Image.Image) -> dict:
    """Extract standard camera EXIF parameters."""
    details = {"Camera Model": "N/A", "Date Taken": "N/A", "ISO": "N/A"}
    try:
        exif = image.getexif()
        if exif:
            for tag_id, val in exif.items():
                tag_name = TAGS.get(tag_id, tag_id)
                if tag_name == "Model":
                    details["Camera Model"] = str(val)
                elif tag_name == "DateTime":
                    details["Date Taken"] = str(val)
                elif tag_name == "ISOSpeedRatings":
                    details["ISO"] = str(val)
    except Exception:
        pass
    return details

def run_gemini_analysis(client: genai.Client, image: Image.Image, model_id: str) -> dict:
    """Query Gemini multimodal API with structured JSON output schema."""
    prompt = (
        "Analyze this image carefully. Provide a concise, clear caption, identify the "
        "primary subject category, extract descriptive tags, and flag any safety concerns."
    )
    try:
        response = client.models.generate_content(
            model=model_id,
            contents=[prompt, image],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=ImageAIAnalysis,
                temperature=0.2,
            ),
        )
        parsed = json.loads(response.text)
        return {
            "AI Caption": parsed.get("caption", "N/A"),
            "AI Category": parsed.get("subject_category", "N/A"),
            "AI Tags": ", ".join(parsed.get("tags", [])),
            "AI Safety Flag": parsed.get("safety_flag", "Safe"),
        }
    except Exception as e:
        return {
            "AI Caption": f"API Error: {str(e)}",
            "AI Category": "Error",
            "AI Tags": "",
            "AI Safety Flag": "Unchecked",
        }

def analyze_photo(file, index, client, model_id):
    """Run full pipeline: technical specs, EXIF, and AI analysis."""
    try:
        file.seek(0)
        img_bytes = file.read()
        file_size_kb = round(len(img_bytes) / 1024, 2)
        
        with Image.open(io.BytesIO(img_bytes)) as pil_img:
            # Normalize orientation / convert to RGB
            rgb_img = pil_img.convert("RGB")
            width, height = pil_img.size
            img_format = pil_img.format or "UNKNOWN"
            aspect_ratio = f"{round(width / max(height, 1), 2)}:1"
            orientation = "Landscape" if width > height else ("Portrait" if height > width else "Square")

            # Technical Color & Detail Metrics
            stat = ImageStat.Stat(rgb_img)
            r, g, b = stat.mean[:3]
            brightness = round(0.299 * r + 0.587 * g + 0.114 * b, 1)

            gray_img = pil_img.convert("L")
            contrast_score = round(ImageStat.Stat(gray_img).stddev[0], 2)

            exif = get_exif_data(pil_img)

            # AI Vision Call
            if client:
                ai_data = run_gemini_analysis(client, rgb_img, model_id)
            else:
                ai_data = {
                    "AI Caption": "No API Key Provided",
                    "AI Category": "Skipped",
                    "AI Tags": "",
                    "AI Safety Flag": "Skipped",
                }

            return {
                "Item #": index + 1,
                "File Name": file.name,
                "Format": img_format,
                "Size (KB)": file_size_kb,
                "Dimensions": f"{width}x{height}",
                "Orientation": orientation,
                "Brightness (0-255)": brightness,
                "Sharpness/Contrast": contrast_score,
                "Camera Model": exif["Camera Model"],
                "Date Taken": exif["Date Taken"],
                "AI Caption": ai_data["AI Caption"],
                "AI Category": ai_data["AI Category"],
                "AI Tags": ai_data["AI Tags"],
                "Safety Check": ai_data["AI Safety Flag"],
                "Status": "Processed",
            }
    except Exception as e:
        return {
            "Item #": index + 1,
            "File Name": file.name,
            "Status": f"Failed: {str(e)}",
        }

# Process Trigger
if uploaded_files:
    if len(uploaded_files) > 50:
        st.warning(f"You selected {len(uploaded_files)} files. Processing capped at the first 50.")
        uploaded_files = uploaded_files[:50]

    # Initialize Gemini client if key present
    client = None
    if api_key_input.strip():
        client = genai.Client(api_key=api_key_input.strip())
    else:
        st.info("💡 Running in metadata-only mode. Enter a Gemini API Key in the left sidebar to enable AI captions and tags.")

    if st.button("🚀 Analyze Photos & Build Excel", type="primary"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        results = []
        for idx, f in enumerate(uploaded_files):
            status_text.text(f"Analyzing {idx + 1}/{len(uploaded_files)}: {f.name}")
            data = analyze_photo(f, idx, client, model_name)
            results.append(data)
            progress_bar.progress((idx + 1) / len(uploaded_files))

        status_text.text("Analysis complete!")
        df = pd.DataFrame(results)

        # Overview Metrics
        st.subheader("📋 Results Preview")
        st.dataframe(df, use_container_width=True)

        # Build Formatted Excel Output
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Photo_Analysis", index=False)
            ws = writer.sheets["Photo_Analysis"]

            # Format Column Widths automatically
            for col in ws.columns:
                header_val = str(col[0].value or "")
                max_content_len = max(len(str(cell.value or "")) for cell in col)
                col_letter = col[0].column_letter
                
                # Give captions and tags adequate space
                if "Caption" in header_val:
                    ws.column_dimensions[col_letter].width = 45
                elif "Tags" in header_val:
                    ws.column_dimensions[col_letter].width = 30
                else:
                    ws.column_dimensions[col_letter].width = max(max_content_len + 3, 12)

        st.download_button(
            label="📥 Download Excel Spreadsheet (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name="ai_picture_analysis_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
