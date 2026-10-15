from datetime import datetime, timedelta
import io
import json
import re
from google import genai
from google.genai import types
import openpyxl
import pandas as pd
from PIL import Image
import streamlit as st

try:
  import fitz  # PyMuPDF
except ImportError:
  fitz = None

st.set_page_config(
    page_title="PigKnows Service & Arrival Processor",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🐖 PigKnows Native PDF & Batch Processor")
st.write(
    "Upload your scanned service sheets (PDF or images). Fast, robust, and"
    " crash-proof."
)

api_key = st.secrets.get("GEMINI_API_KEY", "")

with st.sidebar:
  st.header("Rules Enforced")
  st.markdown("- Year: **2026**")
  st.markdown("- Service Group: **Dynamic (Detected from page header)**")
  st.markdown("- Event Name: `SERVICE`")
  st.markdown("- Service Type: `3= MS (Mixed Semen)`")
  st.markdown("- Service Location: *Blank*")
  st.markdown("- Remark Precedence: `Early` > `Tra`")
  st.markdown("- Arrival Date: **1 day before 1st breeding**")

uploaded_files = st.file_uploader(
    "Upload Service Sheet Scans (PDF, PNG, JPG)",
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
)


def convert_pdf_to_images(pdf_bytes):
  images_bytes = []
  if not fitz:
    st.error("PyMuPDF (fitz) is missing from requirements.txt")
    return images_bytes
  try:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page_num in range(len(doc)):
      page = doc[page_num]
      # 150 DPI is crisp for text and keeps file size tiny
      pix = page.get_pixmap(dpi=150)
      img = Image.open(io.BytesIO(pix.tobytes("jpeg")))
      img.thumbnail((1600, 1600))
      buffer = io.BytesIO()
      img.save(buffer, format="JPEG", quality=85)
      images_bytes.append(buffer.getvalue())
  except Exception as e:
    st.error(f"Error converting PDF: {e}")
  return images_bytes


if uploaded_files:
  image_bytes_list = []

  for uploaded_file in uploaded_files:
    if uploaded_file.type == "application/pdf":
      pdf_bytes = uploaded_file.getvalue()
      pages_imgs = convert_pdf_to_images(pdf_bytes)
      image_bytes_list.extend(pages_imgs)
    else:
      img = Image.open(uploaded_file)
      img.thumbnail((1600, 1600))
      buffer = io.BytesIO()
      img.save(buffer, format="JPEG", quality=85)
      image_bytes_list.append(buffer.getvalue())

  st.info(f"Successfully loaded {len(image_bytes_list)} page(s) ready for AI.")

  if st.button("🔍 Extract Data Instantly", type="primary"):
    if not api_key:
      st.error(
          "API key not found. Please add GEMINI_API_KEY to your Streamlit App"
          " Secrets."
      )
    else:
      with st.spinner("Processing pages with Gemini AI..."):
        try:
          client = genai.Client(api_key=api_key)
          all_parsed_rows = []

          for idx, img_bytes in enumerate(image_bytes_list):
            image_part = types.Part.from_bytes(
                data=img_bytes, mime_type="image/jpeg"
            )

            prompt = f"""
                        You are an expert agricultural data assistant parsing handwritten Pig Breeding Service sheets (Page {idx+1} of {len(image_bytes_list)}).
                        1. Look at the top of the page header (e.g., 'Batch 2') and extract the batch/group number as an integer (1 to 10).
                        2. Extract every row from the image into a JSON list of objects with these exact keys:
                           - "Batch" (integer, detected from page header, e.g. 2)
                           - "Sow ID" (string, e.g., "6142")
                           - "Event Date" (string in YYYY-MM-DD format, assuming year 2026. Dates carry down if blank).
                           - "Tech 1" (string, e.g., "VA" or "V2")
                           - "Tech 2" (string, e.g., "V2", "VA", or "X" if skipped/blank)
                           - "Comments" (string containing full comments text, e.g., "Alt 38310 B Early Tra" or blank)
                           - "Confidence" (float between 0.0 and 1.0)
                        
                        Follow technician propagation rules: if Tech 1 or Tech 2 is blank on a row, it carries over from above.
                        Return ONLY valid JSON list. No markdown wrappers or conversational text.
                        """

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
          "Batch": st.column_config.NumberColumn(
              "Batch / Group", min_value=1, max_value=10, step=1, required=True
          ),
          "Sow ID": st.column_config.TextColumn("Sow ID", required=True),
          "Confidence": st.column_config.ProgressColumn(
              "AI Confidence", min_value=0.0, max_value=1.0, format="%.2f"
          ),
          "Event Date": st.column_config.TextColumn("Event Date (YYYY-MM-DD)"),
          "Tech 1": st.column_config.TextColumn("Tech 1"),
          "Tech 2": st.column_config.TextColumn("Tech 2"),
          "Comments": st.column_config.TextColumn("Comments"),
      },
      disabled=["Confidence"],
      num_rows="dynamic",
      use_container_width=True,
  )

  if st.button(
      "🚀 Generate PigKnows Excel Files (Services & Arrivals)",
      type="primary",
      use_container_width=True,
  ):
    service_rows = []
    arrival_rows = []

    last_tech1 = "VA"
    last_tech2 = "V2"

    first_breeding_date = None
    for _, row in edited_df.iterrows():
      dt_str = str(row["Event Date"]).strip()
      if dt_str and dt_str != "nan":
        try:
          first_breeding_date = datetime.strptime(dt_str, "%Y-%m-%d")
          break
        except ValueError:
          pass

    if not first_breeding_date:
      first_breeding_date = datetime(2026, 10, 9)

    arrival_date_str = (first_breeding_date - timedelta(days=1)).strftime(
        "%Y-%m-%d"
    )

    for _, row in edited_df.iterrows():
      sow_id = str(row["Sow ID"]).strip()
      event_date = str(row["Event Date"]).strip()
      t1 = str(row["Tech 1"]).strip() if pd.notna(row["Tech 1"]) else ""
      t2 = str(row["Tech 2"]).strip() if pd.notna(row["Tech 2"]) else ""
      comments = str(row["Comments"]).strip() if pd.notna(row["Comments"]) else ""

      try:
        batch_num = int(row["Batch"])
      except (ValueError, TypeError):
        batch_num = 1

      if t1 != "":
        last_tech1 = t1
      else:
        t1 = last_tech1

      if t2 != "":
        last_tech2 = t2
      else:
        t2 = last_tech2

      actual_t2 = "" if t2.upper() == "X" else t2

      alt_id = ""
      alt_match = re.search(r"alt\s*([0-9a-zA-Z-]+)", comments, re.IGNORECASE)
      if alt_match:
        alt_id = alt_match.group(1)

      cleaned_remark = ""
      if "early" in comments.lower():
        cleaned_remark = "Early"
      elif "tra" in comments.lower():
        cleaned_remark = "Tra"

      date2 = ""
      if actual_t2 and event_date:
        try:
          dt = datetime.strptime(event_date, "%Y-%m-%d")
          date2 = (dt + timedelta(days=1)).strftime("%Y-%m-%d")
        except ValueError:
          date2 = ""

      if sow_id and sow_id != "nan":
        service_rows.append({
            ";Sow ID": sow_id,
            "Event Date": event_date,
            "Event Name": "SERVICE",
            "Service Group": batch_num,
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
        })

        if alt_id:
          arrival_rows.append({
              ";Sow ID": sow_id,
              "Event Date": arrival_date_str,
              "Event Name": "ARRIVAL",
              "Alt ID": alt_id,
              "Batch": batch_num,
          })

    # --- Generate Service Excel ---
    df_services = pd.DataFrame(service_rows)
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

    df_services_final = pd.concat(
        [pd.DataFrame(template_header_rows), df_services], ignore_index=True
    )

    serv_buffer = io.BytesIO()
    with pd.ExcelWriter(serv_buffer, engine="openpyxl") as writer:
      df_services_final.to_excel(writer, sheet_name="SERVICE", index=False)

    # --- Generate Arrivals Excel ---
    df_arrivals = pd.DataFrame(arrival_rows)
    arr_buffer = io.BytesIO()
    with pd.ExcelWriter(arr_buffer, engine="openpyxl") as writer:
      df_arrivals.to_excel(writer, sheet_name="ARRIVAL", index=False)

    st.success("Excel files successfully built!")

    col1, col2 = st.columns(2)
    with col1:
      st.download_button(
          label="📥 Download PigKnows_Services.xlsx",
          data=serv_buffer.getvalue(),
          file_name="PigKnows_Services.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          use_container_width=True,
      )
    with col2:
      st.download_button(
          label="📥 Download PigKnows_Arrivals.xlsx",
          data=arr_buffer.getvalue(),
          file_name="PigKnows_Arrivals.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          use_container_width=True,
      )
