from docxtpl import DocxTemplate
import streamlit as st
from datetime import datetime, date, timedelta
from io import BytesIO
import pandas as pd
import os
import warnings
from PIL import Image
import base64


st.set_page_config(page_title="SOW Generator", layout="wide")
# --- Display UI Header ---


st.set_page_config(layout="wide")

# --- Convert local logo to base64 so HTML <img> can display it ---
def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

logo_base64 = get_base64_image("logo-clbs- (1).png")

# --- Full-width header style ---
st.markdown(f"""
<style>
/* Make only the header section full width */
.header-full {{
    width: 100vw; /* full viewport width */
    position: relative;
    left: 50%;
    right: 50%;
    margin-left: -50vw;
    margin-right: -50vw;
    background: linear-gradient(90deg, #0a0f1e, #13203d, #1f3d6d);
    padding: 10px 60px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    box-shadow: 0 4px 15px rgba(0,0,0,0.4);
    border-bottom: 2px solid #2c4e8a;
    z-index: 10;
}}

.header-logo img {{
    height: 40px;
}}

.header-text h1 {{
    font-size: 34px;
    font-weight: 800;
    color: #ffffff;
    margin: 0;
    letter-spacing: 1px;
}}

.header-text p {{
    font-size: 16px;
    color: #b0c4de;
    margin-top: 5px;
}}
</style>

<div class="header-full">
    <div class="header-logo">
        <img src="data:image/png;base64,{logo_base64}" alt="CloudLabs Logo">
    </div>
    <div class="header-text">
        <h1>SOW Generator</h1>
        <p>Single Click Word SOW Generator</p>
    </div>
</div>
""", unsafe_allow_html=True)




warnings.filterwarnings("ignore", category=UserWarning, module='pkg_resources')


# st.title("SOW Generator — Single Click Word SOW")
# st.markdown("Fill fields below and click **Generate SOW**. Uses a Word template with Jinja placeholders.")

# --- Upload or choose template ---
template_file = st.file_uploader("Upload client Word template (.docx)", type=["docx"])

# --- Basic fields ---
Client_Name = st.selectbox(
    "Select Client",
    ("BSC", "Abiomed", "Cognex", "Itaros")      
)

sow_num = st.text_input("SOW Number", "1234")

colA, colB = st.columns([1, 1])
with colA:
    option = st.selectbox(
    "Select Project Type",
    ("Fixed Fee", "T&M")      
)
with colB:
    sow_name = st.text_input("SOW Name", "SOW - Implementation")

colA, colB = st.columns([1, 1])
with colA:
    start_date = st.date_input("Start Date", date.today())
with colB:
    end_date = st.date_input("End Date", date.today())

colA, colB = st.columns([1, 1])
with colA:
    pm_client = st.text_input("PM Client", "John Client")
with colB:
    pm_sp = st.text_input("PM Service Provider", "Project PM")
mg_client = st.text_input("Mgmt Client", "Mgmt Client")
mg_sp = st.text_input("Mgmt Service Provider", "Umang Naik")
scope_text = st.text_area("Scope / Responsibilities", "Describe scope here...")
ser_del = st.text_area("Services / Deliverables", "Describe Services/Del. here...")

# --- Format dates ---
generated_date = datetime.today().strftime("%B %d, %Y")
start_str = start_date.strftime("%B %d, %Y")
end_str = end_date.strftime("%B %d, %Y")

# --- Helper to calculate working days (like Excel NETWORKDAYS) ---
def networkdays(start_date, end_date):
    day_count = 0
    current = start_date
    while current <= end_date:
        if current.weekday() < 5:  # Mon-Fri only
            day_count += 1
        current += timedelta(days=1)
    return day_count

workdays = networkdays(start_date, end_date)
st.write(f"📅 Total working days (Mon–Fri) between selected dates: **{workdays}**")

# --- Resources Table ---
if option == "T&M":
    st.subheader("Resource Details")

    resources_df = st.data_editor(
        pd.DataFrame(
            columns=[
                "Role", "Location", "Start Date", "End Date",
                "Allocation %", "Hrs/Day", "Rate/hr ($)"
            ],
        data=[[ "", "", start_date, end_date, 100, 8, 100 ]]
        ),
        num_rows="dynamic",
        key="resources_table"
    )

    # --- Calculate Estimated $ per row ---
    if not resources_df.empty:
        def calc_value(row):
            try:
                start = pd.to_datetime(row["Start Date"])
                end = pd.to_datetime(row["End Date"])
                days = len(pd.bdate_range(start, end))
                return round(days * (row["Allocation %"]/100) * row["Hrs/Day"] * row["Rate/hr ($)"], 2)
            except Exception:
                return 0.0

        resources_df["Estimated $"] = resources_df.apply(calc_value, axis=1)
        st.dataframe(resources_df)

# --- Total Contract Value ---
    currency_value = resources_df["Estimated $"].sum()
    currency_value_str = f"${currency_value:,.2f}"
    st.write(f"💰 Total Contract Value: **{currency_value_str}**")

# --- Generate Word SOW ---
if st.button("Generate SOW Document"):

    if template_file is None:
        st.warning("Please upload a Word template (.docx) before generating.")
    else:
        # Save uploaded template temporarily
        template_path = os.path.join("generated_sows", "template.docx")
        os.makedirs("generated_sows", exist_ok=True)
        with open(template_path, "wb") as f:
            f.write(template_file.getbuffer())

        # --- Context for template ---
        context = {
            "sow_num": sow_num,
            "sow_name": sow_name,
            "pm_client": pm_client,
            "pm_SP": pm_sp,
            "mg_client": mg_client,
            "mg_sp": mg_sp,
            "ser_del": ser_del,
            "scope_text": scope_text,
            "start_date": start_str,
            "end_date": end_str,
            "resources": resources_df.to_dict(orient="records"),
            "generated_date": generated_date,
            "currency_value_str": currency_value_str,
            "currency_value": currency_value
        }

        # --- Render Word template ---
        doc = DocxTemplate(template_path)
        doc.render(context)

        # --- Save generated file ---
        output_file = os.path.join(
            "generated_sows",
            f"{sow_num} - {sow_name} - {start_str} to {end_str}.docx"
        )
        doc.save(output_file)

        st.success(f"SOW Document generated: {output_file}")
        with open(output_file, "rb") as f:
            st.download_button(
                "📄 Download SOW",
                data=f,
                file_name=os.path.basename(output_file),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )