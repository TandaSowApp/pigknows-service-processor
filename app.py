from datetime import datetime, timedelta
import io
import json
from google import genai
from google.genai import types
import openpyxl
import pandas as pd
import streamlit as st

try:
  import fitz  # PyMuPDF
except ImportError:
  fitz = None

st.set_page_config(
    page_title="PigKnows Service Sheet Processor (Free)",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🐖 PigKnows Service Sheet Processor (Gemini Free Tier)")
st.write(
    "Upload your scanned service sheet (PDF or image). Powered by Google's"
    " free Gemini vision model."
)

api_key = st.secrets.get("GEMINI_API_KEY", "")

with st.sidebar:
  st.header("Rules Enforced")
  st.markdown("- Year: **2026**")
  st.markdown("- Event Name: `SERVICE`")
  st.markdown("- Service Type: `3= MS (Mixed Semen)`")
  st.markdown("- Service Location: *Blank*")
  st.markdown("- Remark Precedence: `Early` > `Tra`")

uploaded_file = st.file_uploader(
    "Upload Service Sheet Scan (PDF, PNG, JPG)",
    type=["pdf", "png", "jpg", "jpeg"],
)


def convert_pdf_to_images(pdf_bytes):
  images = []
  if not fitz:
    st.error("PyMuPDF (fitz) is missing from requirements.txt")
    return images
  doc = fitz.open(stream=pdf_bytes, filetype="pdf")
  for page_num in range(len(doc)):
    page = doc[page_num]
    pix = page.get_pixmap(dpi=200)
    images.append(pix.tobytes("jpeg"))
  return images


if uploaded_file:
  image_bytes_list = []

  if uploaded_file.type == "application/pdf":
    st.info("PDF scan detected. Converting pages for vision processing...")
    pdf_bytes = uploaded_file.getvalue()
    image_bytes_list = convert_pdf_to_images(pdf_bytes)
  else:
    image_bytes_list = [uploaded_file.getvalue()]
    st.image(uploaded_file, caption="Uploaded Scan", use_container_width=True)

  if st.button("🔍 Extract Data with Gemini Free AI", type="primary"):
    if not api_key:
      st.error(
          "API key not found. Please add GEMINI_API_KEY to your Streamlit App"
          " Secrets."
      )
    else:
      with st.spinner(
          "Reading handwriting and applying agricultural business rules..."
      ):
        try:
          client = genai.Client(api_key=api_key)
          all_parsed_rows = []

          for img_bytes in image_bytes_list:
            image_part = types.Part.from_bytes(
                data=img_bytes, mime_type="image/jpeg"
            )

            prompt = """
                        You are an expert agricultural data assistant parsing handwritten Pig Breeding Service sheets.
                        Extract every row from the image into a JSON list of objects with these exact keys:
                        - "Sow ID" (string, e.g., "6142")
                        - "Event Date" (string in YYYY-MM-DD format, assuming year 2026)
                        - "Tech 1" (string, e.g., "VA" or "V2")
                        - "Tech 2" (string, e.g., "V2", "VA", or "X" if skipped/blank)
                        - "Remark" (string, e.g., "Early", "Tra", or blank if none)
                        - "Confidence" (float between 0.0 and 1.0 indicating OCR certainty on the Sow ID digits)
                        
                        Follow technician propagation rules: if Tech 1 or Tech 2 is blank on a row, it carries over from above.
                        Return ONLY valid JSON list. No markdown wrappers or conversational text.
                        """

            # Updated to current standard model identifier
            response = client.models.generate_content(
                model="gemini-3.8-flash", contents=[image_part, prompt]
            )

            content = response.text.strip()
            if "```json" in content:
              content = content.split("```json")[1].split("```")[0].strip()
            elif "```" in content:
              content = content.split("```")[1].split("```")[0].strip()

            page_rows = json.loads(content)
            all_parsed_rows.extend(page_rows)

          st.session_state["df_extracted"] = pd.DataFrame(all_parsed_rows)
          st.success("Extraction complete!")
          st.rerun()

        except Exception as e:
          st.error(f"Error during AI vision parsing: {e}")

if "df_extracted" in st.session_state:
  st.markdown("---")
  st.subheader("📋 Step 2: Verify & Edit Extracted Data")

  edited_df = st.data_editor(
      st.session_state["df_extracted"],
      column_config={
          "Sow ID": st.column_config.TextColumn(
              "Sow ID (Verify digits)", required=True
          ),
          "Confidence": st.column_config.ProgressColumn(
              "AI Confidence", min_value=0.0, max_value=1.0, format="%.2f"
          ),
          "Event Date": st.column_config.TextColumn("Event Date (YYYY-MM-DD)"),
          "Tech 1": st.column_config.TextColumn("Tech 1"),
          "Tech 2": st.column_config.TextColumn("Tech 2"),
          "Remark": st.column_config.TextColumn("Remark"),
      },
      disabled=["Confidence"],
      num_rows="dynamic",
      use_container_width=True,
  )

  if st.button(
      "🚀 Generate PigKnows-Ready Excel File",
      type="primary",
      use_container_width=True,
  ):
    final_rows = []
    last_tech1 = "VA"
    last_tech2 = "V2"

    for _, row in edited_df.iterrows():
      sow_id = str(row["Sow ID"]).strip()
      event_date = str(row["Event Date"]).strip()
      t1 = str(row["Tech 1"]).strip() if pd.notna(row["Tech 1"]) else ""
      t2 = str(row["Tech 2"]).strip() if pd.notna(row["Tech 2"]) else ""
      comment = str(row["Remark"]).strip() if pd.notna(row["Remark"]) else ""

      if t1 != "":
        last_tech1 = t1
      else:
        t1 = last_tech1

      if t2 != "":
        last_tech2 = t2
      else:
        t2 = last_tech2

      actual_t2 = "" if t2.upper() == "X" else t2

      cleaned_remark = ""
      if "early" in comment.lower():
        cleaned_remark = "Early"
      elif "tra" in comment.lower():
        cleaned_remark = "Tra"

      date2 = ""
      if actual_t2 and event_date:
        try:
          dt = datetime.strptime(event_date, "%Y-%m-%d")
          date2 = (dt + timedelta(days=1)).strftime("%Y-%m-%d")
        except ValueError:
          date2 = ""

      pigknows_row = {
          ";Sow ID": sow_id,
          "Event Date": event_date,
          "Event Name": "SERVICE",
          "Service Group": 1,
          "Service Type": "3= MS (Mixed Semen)",
          "Service Location": "",
          "Boar 1": "",
          "Tech 1": t1,
          "Date 2": date2,
          "Boar 2": "",
          "Tech 2": actual_t2,
          "Date 3": "",
          "Boar 3": "",
          "Tech 3": "",
          "Remark": cleaned_remark,
          "Semen Age 1": "",
          "Semen Age 2": "",
          "Semen Age 3": "",
          "Serv Score 1": "",
          "Serv Time 1": "",
          "Serv Time 2": "",
          "Serv Time 3": "",
          "Serv Score 2": "",
          "Serv Score 3": "",
          "Weight": "",
          "Sow Weight": "",
          "Date 4": "",
          "Boar 4": "",
          "Tech 4": "",
          "Date 5": "",
          "Boar 5": "",
          "Tech 5": "",
      }
      final_rows.append(pigknows_row)

    df_output = pd.DataFrame(final_rows)

    template_header_rows = [
        {
            ";Sow ID": ";",
            "Event Date": None,
            "Event Name": "SERVICE",
            "Service Group": None,
            "Service Type": "1= AI (Artificial Insemination)",
            "Service Location": None,
            "Boar 1": None,
            "Tech 1": None,
            "Date 2": None,
            "Boar 2": None,
            "Tech 2": None,
            "Date 3": None,
            "Boar 3": None,
            "Tech 3": None,
            "Remark": None,
            "Semen Age 1": None,
            "Semen Age 2": None,
            "Semen Age 3": None,
            "Serv Score 1": None,
            "Serv Time 1": None,
            "Serv Time 2": None,
            "Serv Time 3": None,
            "Serv Score 2": None,
            "Serv Score 3": None,
            "Weight": None,
            "Sow Weight": None,
            "Date 4": None,
            "Boar 4": None,
            "Tech 4": None,
            "Date 5": None,
            "Boar 5": None,
            "Tech 5": None,
        },
        {
            ";Sow ID": ";",
            "Event Date": None,
            "Event Name": "SERVICE",
            "Service Group": None,
            "Service Type": "2= NS (Natural Service)",
            "Service Location": None,
            "Boar 1": None,
            "Tech 1": None,
            "Date 2": None,
            "Boar 2": None,
            "Tech 2": None,
            "Date 3": None,
            "Boar 3": None,
            "Tech 3": None,
            "Remark": None,
            "Semen Age 1": None,
            "Semen Age 2": None,
            "Semen Age 3": None,
            "Serv Score 1": None,
            "Serv Time 1": None,
            "Serv Time 2": None,
            "Serv Time 3": None,
            "Serv Score 2": None,
            "Serv Score 3": None,
            "Weight": None,
            "Sow Weight": None,
            "Date 4": None,
            "Boar 4": None,
            "Tech 4": None,
            "Date 5": None,
            "Boar 5": None,
            "Tech 5": None,
        },
        {
            ";Sow ID": ";",
            "Event Date": None,
            "Event Name": "SERVICE",
            "Service Group": None,
            "Service Type": "3= MS (Mixed Semen)",
            "Service Location": None,
            "Boar 1": None,
            "Tech 1": None,
            "Date 2": None,
            "Boar 2": None,
            "Tech 2": None,
            "Date 3": None,
            "Boar 3": None,
            "Tech 3": None,
            "Remark": None,
            "Semen Age 1": None,
            "Semen Age 2": None,
            "Semen Age 3": None,
            "Serv Score 1": None,
            "Serv Time 1": None,
            "Serv Time 2": None,
            "Serv Time 3": None,
            "Serv Score 2": None,
            "Serv Score 3": None,
            "Weight": None,
            "Sow Weight": None,
            "Date 4": None,
            "Boar 4": None,
            "Tech 4": None,
            "Date 5": None,
            "Boar 5": None,
            "Tech 5": None,
        },
        {
            ";Sow ID": ";",
            "Event Date": None,
            "Event Name": "SERVICE",
            "Service Group": None,
            "Service Type": "4= AN (Artificial, Natural)",
            "Service Location": None,
            "Boar 1": None,
            "Tech 1": None,
            "Date 2": None,
            "Boar 2": None,
            "Tech 2": None,
            "Date 3": None,
            "Boar 3": None,
            "Tech 3": None,
            "Remark": None,
            "Semen Age 1": None,
            "Semen Age 2": None,
            "Semen Age 3": None,
            "Serv Score 1": None,
            "Serv Time 1": None,
            "Serv Time 2": None,
            "Serv Time 3": None,
            "Serv Score 2": None,
            "Serv Score 3": None,
            "Weight": None,
            "Sow Weight": None,
            "Date 4": None,
            "Boar 4": None,
            "Tech 4": None,
            "Date 5": None,
            "Boar 5": None,
            "Tech 5": None,
        },
        {
            ";Sow ID": ";",
            "Event Date": None,
            "Event Name": "SERVICE",
            "Service Group": None,
            "Service Type": "5= NA (Natural, Artificial)",
            "Service Location": None,
            "Boar 1": None,
            "Tech 1": None,
            "Date 2": None,
            "Boar 2": None,
            "Tech 2": None,
            "Date 3": None,
            "Boar 3": None,
            "Tech 3": None,
            "Remark": None,
            "Semen Age 1": None,
            "Semen Age 2": None,
            "Semen Age 3": None,
            "Serv Score 1": None,
            "Serv Time 1": None,
            "Serv Time 2": None,
            "Serv Time 3": None,
            "Serv Score 2": None,
            "Serv Score 3": None,
            "Weight": None,
            "Sow Weight": None,
            "Date 4": None,
            "Boar 4": None,
            "Tech 4": None,
            "Date 5": None,
            "Boar 5": None,
            "Tech 5": None,
        },
    ]

    df_final = pd.concat(
        [pd.DataFrame(template_header_rows), df_output], ignore_index=True
    )

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
      df_final.to_excel(writer, sheet_name="SERVICE", index=False)

    st.success("Excel file successfully built!")
    st.download_button(
        label="📥 Download PigKnows_Services.xlsx",
        data=buffer.getvalue(),
        file_name="PigKnows_Services.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
