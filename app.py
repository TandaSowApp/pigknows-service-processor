from datetime import datetime, timedelta
import io
import json
import os
import openpyxl
import pandas as pd
import streamlit as st

try:
  from openai import OpenAI
except ImportError:
  OpenAI = None

st.set_page_config(
    page_title="PigKnows Service Sheet Processor",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🐖 PigKnows Service Sheet Processor")
st.write(
    "Upload your scanned service sheet. The app extracts the handwriting,"
    " allows you to verify any low-confidence entries, and exports a clean"
    " PigKnows-ready Excel file."
)

# Sidebar for API Key & Settings
with st.sidebar:
  st.header("Settings & API")
  api_key = st.text_input(
      "OpenAI API Key",
      type="password",
      value=os.environ.get("OPENAI_API_KEY", ""),
      help="Required for automated vision parsing of your scanned sheets.",
  )

  st.markdown("---")
  st.markdown("**Enforced Rules:**")
  st.markdown("- Year: **2026**")
  st.markdown("- Event Name: `SERVICE`")
  st.markdown("- Service Type: `3= MS (Mixed Semen)`")
  st.markdown("- Service Location: *Blank*")
  st.markdown("- Remark Precedence: `Early` > `Tra`")

uploaded_file = st.file_uploader(
    "Upload Scanned Service Sheet (PNG, JPG, JPEG)",
    type=["png", "jpg", "jpeg"],
)

if uploaded_file:
  st.image(uploaded_file, caption="Uploaded Scan", use_container_width=True)

  if st.button("🔍 Extract Data from Scan", type="primary"):
    if not api_key and OpenAI:
      st.error("Please enter your OpenAI API key in the sidebar first.")
    else:
      with st.spinner("Analyzing handwriting and parsing rows..."):
        try:
          # Convert uploaded image to bytes for Vision API
          image_bytes = uploaded_file.getvalue()

          client = OpenAI(api_key=api_key)

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
                    Return ONLY valid JSON. No markdown wrappers around the JSON if possible, or standard json format.
                    """

          import base64

          base64_image = base64.b64encode(image_bytes).decode("utf-8")

          response = client.chat.completions.create(
              model="gpt-4o",
              messages=[{
                  "role": "user",
                  "content": [
                      {"type": "text", "text": prompt},
                      {
                          "type": "image_url",
                          "image_url": {
                              "url": f"data:image/jpeg;base64,{base64_image}"
                          },
                      },
                  ],
              }],
              max_tokens=4000,
          )

          content = response.choices[0].message.content
          # Clean up response if it contains markdown code blocks
          if "```json" in content:
            content = content.split("```json")[1].split("```")[0].strip()
          elif "```" in content:
            content = content.split("```")[1].split("```")[0].strip()

          parsed_data = json.loads(content)
          st.session_data = pd.DataFrame(parsed_data)
          st.success("Extraction complete!")
          st.rerun()

        except Exception as e:
          st.error(
              f"Error during AI vision parsing: {e}. Falling back to sample"
              " structure."
          )
          # Fallback sample data if API call fails
          fallback_data = [
              {
                  "Sow ID": "6142",
                  "Confidence": 0.98,
                  "Event Date": "2026-10-09",
                  "Tech 1": "VA",
                  "Tech 2": "V2",
                  "Remark": "Early Tra",
              },
              {
                  "Sow ID": "6143",
                  "Confidence": 0.95,
                  "Event Date": "2026-10-09",
                  "Tech 1": "V2",
                  "Tech 2": "V2",
                  "Remark": "Early Tra",
              },
              {
                  "Sow ID": "6152",
                  "Confidence": 0.65,
                  "Event Date": "2026-10-10",
                  "Tech 1": "V2",
                  "Tech 2": "X",
                  "Remark": "Tra",
              },
          ]
          st.session_state["df_extracted"] = pd.DataFrame(fallback_data)

# If extracted data exists in session state, show interactive editor
if "df_extracted" in st.session_state:
  st.markdown("---")
  st.subheader("📋 Step 2: Verify & Edit Extracted Data")
  st.markdown(
      "Review the parsed rows. Any low-confidence Sow ID digits can be edited"
      " directly in the table before generating your file."
  )

  edited_df = st.data_editor(
      st.session_state["df_extracted"],
      column_config={
          "Sow ID": st.column_config.TextColumn(
              "Sow ID (Verify digits)", required=True
          ),
          "Confidence": st.column_config.ProgressColumn(
              "OCR Confidence", min_value=0.0, max_value=1.0, format="%.2f"
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
      "🚀 Generate PigKnows Excel Import File",
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

      # Technician propagation
      if t1 != "":
        last_tech1 = t1
      else:
        t1 = last_tech1

      if t2 != "":
        last_tech2 = t2
      else:
        t2 = last_tech2

      # Handle 'X' as blank secondary service
      actual_t2 = "" if t2.upper() == "X" else t2

      # Remark precedence logic ("Early" beats "Tra")
      cleaned_remark = ""
      if "early" in comment.lower():
        cleaned_remark = "Early"
      elif "tra" in comment.lower():
        cleaned_remark = "Tra"

      # Date 2 calculation (+1 day if Tech 2 present)
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
          "Service Location": "",  # Blank per template
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

    # Prepend 5 standard PigKnows template header rows
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
