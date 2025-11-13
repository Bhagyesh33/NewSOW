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
st.markdown("""
<style>
/* Hide Streamlit default header, footer, and menu */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}

/* Remove top padding Streamlit leaves behind */
div.block-container {
    padding-top: 0rem;
    padding-bottom: 1rem;
}

/* Optional: make app slightly wider */
section.main > div {
    padding-top: 0rem;
}
</style>
""", unsafe_allow_html=True)

# --- Display UI Header ---

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
st.markdown("<br>", unsafe_allow_html=True)

template_file = st.file_uploader("Upload client Word template (.docx)", type=["docx"])

# --- Basic fields ---
Client_Name = st.selectbox(
    "Select Client",
    ("BSC", "Abiomed", "Cognex", "Itaros")      
)
option = st.selectbox(
    "Select Project Type",
    ("Fixed Fee", "T&M", "Change Order")      
)

if option == "Change Order":
    Change = st.text_input("Change Order", "10")
colA, colB = st.columns([1, 1])
with colA:
    sow_num = st.text_input("SOW Number", "1234")
with colB:
    sow_name = st.text_input("SOW Name", "SOW - Implementation")

if option == "Change Order":
    colA, colB = st.columns([1, 1])
    with colA:
        sow_start_date = st.date_input("SOW Start Date", date.today())
    with colB:
        sow_end_date = st.date_input("SOW End Date", date.today())

colA, colB = st.columns([1, 1])
with colA:
    start_date = st.date_input("Start Date", date.today())
with colB:
    end_date = st.date_input("End Date", date.today())

colA, colB = st.columns([1, 1])
with colA:
    pm_client = st.text_input("Client (Project Management)", "John Client")
with colB:
    pm_sp = st.text_input("Service Provider (Project Management)", "Project PM")

colA, colB = st.columns([1, 1])
with colA:
    mg_client = st.text_input("Client (Management)", "Mgmt Client")
with colB:
    mg_sp = st.text_input("Service Provider (Management)", "Umang Naik")

scope_text = st.text_area("Scope / Responsibilities", "Describe scope here...")
ser_del = st.text_area("Services / Deliverables", "Describe Services/Del. here...")
if option == "Fixed Fee":
    Fees_al = st.text_input("Fees", "100")

if option == "Change Order":
    colA, colB = st.columns([1, 1])
    with colA:
        Fees_co = st.text_input("Change Order Fees", "100")
    with colB:
        Fees_sow = st.text_input("SOW Fees", "100")  
    
    difference = float(Fees_co) - float(Fees_sow)

# --- Format dates ---
generated_date = datetime.today().strftime("%B %d, %Y")
start_str = start_date.strftime("%B %d, %Y")
end_str = end_date.strftime("%B %d, %Y")
if option == "Change Order":
    sow_str = sow_start_date.strftime("%B %d, %Y")
    sow_end = sow_end_date.strftime("%B %d, %Y")

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

if option == "Fixed Fee":
    st.subheader("Milestone Schedule / Payment Breakdown")

    # Convert Fees to numeric
    try:
        total_fees = float(Fees_al)
    except:
        total_fees = 0

    default_data = [
        ["1", "", date.today(), ""]
    ]

    # ✅ Editable table (NO Net Payment here)
    milestone_input_df = st.data_editor(
        pd.DataFrame(
            default_data,
            columns=[
                "Milestone #",
                "Services / Deliverables",
                "Milestone Due Date",
                "Payment Allocation (%)"
            ]
        ),
        num_rows="dynamic",
        key="milestone_table"
    )

    # ✅ Calculate Net Payment column separately
    milestone_df = milestone_input_df.copy()

    def calc_net(row):
        try:
            alloc = float(row["Payment Allocation (%)"])
            return round(total_fees * (alloc / 100), 2)
        except:
            return 0

    milestone_df["Net Milestone Payment ($)"] = milestone_df.apply(calc_net, axis=1)

    # ✅ Show final results
    st.write("🔹 Calculated Milestones")
    st.dataframe(milestone_df)

    total_payment = milestone_df["Net Milestone Payment ($)"].sum()
    st.write(f"✅ Total Net Milestone Payment: **${total_payment:,.2f}**")
    # Fix column keys for Jinja compatibility
    milestone_df = milestone_df.rename(columns={
        "Milestone #": "milestone_no",
        "Services / Deliverables": "services",
        "Milestone Due Date": "due_date",
        "Payment Allocation (%)": "allocation",
        "Net Milestone Payment ($)": "net_pay"
    })




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

        if option == "T&M":
        # --- Context for t&m template ---
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

        if option == "Fixed Fee":
        # --- Context for fixedfee template ---
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
                # "resources": resources_df.to_dict(orient="records"),
                "generated_date": generated_date,
                # "currency_value_str": currency_value_str,
                # "currency_value": currency_value
                "milestones": milestone_df.to_dict(orient="records"),
                "milestone_total": total_payment,
                "Fees" : Fees_al
            }

        if option == "Change Order":
        # --- Context for fixedfee template ---
            context = {
                "Change": Change,
                "sow_num": sow_num,
                "sow_name": sow_name,
                "scope_text": scope_text,
                "start_date": start_str,
                "end_date": end_str,
                "sow_end" : sow_end,
                "sow_str" : sow_str,
                "Fees_co" : Fees_co,
                "Fees_sow" : Fees_sow,
                "difference" : difference
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