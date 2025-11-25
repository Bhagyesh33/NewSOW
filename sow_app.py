from docxtpl import DocxTemplate
import streamlit as st
from datetime import datetime, date, timedelta
from io import BytesIO
import pandas as pd
import os
import warnings
from PIL import Image
import base64




st.set_page_config(page_title="SOW Generator", layout="wide", page_icon="📋")
st.markdown("""
<style>
/* REMOVE top-right sidebar toggle ( » ) */
.css-1rs6os.edgvbvh3 { 
    display: none !important;
}

/* Remove the top padding Streamlit adds */
.block-container {
    padding-top: 0 !important;
}

/* Hide the default Streamlit header completely */
header {visibility: hidden !important;}
</style>
""", unsafe_allow_html=True)

# st.markdown("""
# <style>
# /* Only hide the default header and footer */
# footer {visibility: hidden;}
# header {visibility: hidden;}

# /* Remove top padding */
# .block-container {
#     padding-top: 0rem;
# }

# /* Keep your custom header styles */
# .header-full {
#     width: 100vw;
#     position: relative;
#     left: 50%;
#     right: 50%;
#     margin-left: -50vw;
#     margin-right: -50vw;
#     background: linear-gradient(90deg, #0a0f1e, #13203d, #1f3d6d);
#     padding: 10px 60px;
#     display: flex;
#     align-items: center;
#     justify-content: space-between;
#     box-shadow: 0 4px 15px rgba(0,0,0,0.4);
#     border-bottom: 2px solid #2c4e8a;
#     z-index: 10;
# }

# .header-logo img {
#     height: 40px;
# }

# .header-text h1 {
#     font-size: 34px;
#     font-weight: 800;
#     color: #ffffff;
#     margin: 0;
#     letter-spacing: 1px;
# }

# .header-text p {
#     font-size: 16px;
#     color: #b0c4de;
#     margin-top: 5px;
# }
# </style>
# """, unsafe_allow_html=True)
# --- Initialize session state ---
if 'should_increment_on_download' not in st.session_state:
    st.session_state.should_increment_on_download = False
if 'generated_file_path' not in st.session_state:
    st.session_state.generated_file_path = None
if 'file_data' not in st.session_state:
    st.session_state.file_data = None
if 'reset_trigger' not in st.session_state:
    st.session_state.reset_trigger = 0
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


# ----------------------------
# SIDEBAR MULTI-PAGE NAVIGATION
# ----------------------------
# ------- Custom Sidebar Area Under Header -------
with st.container():
    # st.markdown("""
    # <div style='background:#f7f9fc;padding:15px;border-radius:12px;
    #             margin-top:15px;margin-bottom:20px;'>
    #     <h3 style='margin:0;padding:0;'>📌 Navigation</h3>
    # </div>
    # """, unsafe_allow_html=True)

    page = st.radio(
        "",
        ["SOW Generator", "SOW Records"],
        horizontal=True
    )



warnings.filterwarnings("ignore", category=UserWarning, module='pkg_resources')

def save_sow_record(record):
    """Append SOW details to CSV file."""
    folder = "sow_records"
    os.makedirs(folder, exist_ok=True)
    file_path = os.path.join(folder, "sow_records.csv")

    df_new = pd.DataFrame([record])

    # If file exists → append
    if os.path.exists(file_path):
        df_existing = pd.read_csv(file_path)
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        df_combined.to_csv(file_path, index=False)
    else:
        # First time → create file
        df_new.to_csv(file_path, index=False)


def get_next_sow_number(peek_only=False):
    counter_file = "sow_counter.txt"
    start_num = 1000  # starting number

    # Ensure the file exists
    if not os.path.exists(counter_file):
        with open(counter_file, "w") as f:
            f.write(str(start_num))

    # Read the current number safely
    try:
        with open(counter_file, "r") as f:
            content = f.read().strip()
            current = int(content) if content else start_num
    except ValueError:
        current = start_num

    if peek_only:
        # Just preview, don’t increment file
        return current

    # If not peeking, increment and save
    next_num = current + 1
    with open(counter_file, "w") as f:
        f.write(str(next_num))

    return current

def reset_all_fields():
    """Clear all form-related session state to reset inputs"""
    keys_to_keep = ['should_increment_on_download', 'generated_file_path', 'file_data', 'reset_trigger']
    keys_to_remove = [key for key in st.session_state.keys() if key not in keys_to_keep]
    for key in keys_to_remove:
        del st.session_state[key]
    st.session_state.reset_trigger += 1

# st.title("SOW Generator — Single Click Word SOW")
# st.markdown("Fill fields below and click **Generate SOW**. Uses a Word template with Jinja placeholders.")

# --- Upload or choose template ---
st.markdown("<br>", unsafe_allow_html=True)



if page == "SOW Generator":

    template_file = st.file_uploader("Upload client Word template (.docx)", type=["docx"], key=f"template_{st.session_state.reset_trigger}")

    # --- Basic fields ---
    Client_Name = st.selectbox(
        "Select Client",
        ("BSC", "Abiomed", "Cognex", "Itaros"),
        key=f"client_{st.session_state.reset_trigger}"      
    )
    option = st.selectbox(
        "Select Project Type",
        ("Fixed Fee", "T&M", "Change Order"),
        key=f"project_type_{st.session_state.reset_trigger}"      
    )

    if option == "Change Order":
        Change = st.text_input("Change Order", "10", key=f"change_{st.session_state.reset_trigger}")
    colA, colB = st.columns([1, 1])
    with colA:
        auto_sow_num = get_next_sow_number(peek_only=True)
        sow_num = st.text_input("SOW Number", str(auto_sow_num), key=f"sow_num_{st.session_state.reset_trigger}")



    with colB:
        sow_name = st.text_input("SOW Name", key=f"sow_name_{st.session_state.reset_trigger}")

    if option == "Change Order":
        colA, colB = st.columns([1, 1])
        with colA:
            sow_start_date = st.date_input("SOW Start Date", date.today(), key=f"sow_start_{st.session_state.reset_trigger}")
        with colB:
            sow_end_date = st.date_input("SOW End Date", date.today(), key=f"sow_end_{st.session_state.reset_trigger}")

    colA, colB = st.columns([1, 1])
    with colA:
        start_date = st.date_input("Start Date", date.today(), key=f"start_date_{st.session_state.reset_trigger}")
    with colB:
        end_date = st.date_input("End Date", date.today(), key=f"end_date_{st.session_state.reset_trigger}")

    colA, colB = st.columns([1, 1])
    with colA:
        pm_client = st.text_input("Client (Project Management)", key=f"pm_client_{st.session_state.reset_trigger}",  help="Name of the client used specifically for project management documentation and reporting.")
    with colB:
        pm_sp = st.text_input("Service Provider (Project Management)", key=f"pm_sp_{st.session_state.reset_trigger}", help="Name of the service provider responsible for project management activities.")

    colA, colB = st.columns([1, 1])
    with colA:
        mg_client = st.text_input("Client (Management)", key=f"mg_client_{st.session_state.reset_trigger}", help="Official client name used for management-level communication and approvals.")
    with colB:
        mg_sp = st.text_input("Service Provider (Management)", key=f"mg_sp_{st.session_state.reset_trigger}", help="Service provider name used for management, governance, and contract-level processes.")

    scope_text = st.text_area("Scope / Responsibilities", key=f"scope_{st.session_state.reset_trigger}")
    ser_del = st.text_area("Services / Deliverables", key=f"ser_del_{st.session_state.reset_trigger}")
    if option == "Fixed Fee":
        Fees_al = st.text_input("Fees", key=f"fees_al_{st.session_state.reset_trigger}")

    if option == "Change Order":
        colA, colB = st.columns([1, 1])
        with colA:
            Fees_co = st.text_input("Change Order Fees", "10", key=f"fees_co_{st.session_state.reset_trigger}")
        with colB:
            Fees_sow = st.text_input("SOW Fees", "10", key=f"fees_sow_{st.session_state.reset_trigger}")  
        
        difference = float(Fees_co) - float(Fees_sow)

    additional_personnel = st.text_input(
    "Additional Personnel",
    key=f"additional_personnel_{st.session_state.reset_trigger}"
)
    

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
                    "pm_sp": pm_sp,
                    "mg_client": mg_client,
                    "mg_sp": mg_sp,
                    "ser_del": ser_del,
                    "scope_text": scope_text,
                    "start_date": start_str,
                    "end_date": end_str,
                    "resources": resources_df.to_dict(orient="records"),
                    "generated_date": generated_date,
                    "currency_value_str": currency_value_str,
                    "currency_value": currency_value,
                    "additional_personnel": additional_personnel
                }

            if option == "Fixed Fee":
            # --- Context for fixedfee template ---
                context = {
                    "sow_num": sow_num,
                    "sow_name": sow_name,
                    "pm_client": pm_client,
                    "pm_sp": pm_sp,
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
                    "Fees" : Fees_al,
                    "additional_personnel": additional_personnel
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
                    "difference" : difference,
                    "additional_personnel": additional_personnel
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
            # Store file data in session state for download
        with open(output_file, "rb") as f:
            st.session_state.file_data = f.read()
        
        st.session_state.generated_file_path = output_file
        st.session_state.should_increment_on_download = True
        # --- END OF REPLACEMENT ---

    # --- ADD THIS NEW SECTION AFTER THE GENERATION BUTTON BLOCK ---
    # --- Show download button if file is generated ---
    if st.session_state.should_increment_on_download and st.session_state.file_data:
        # Create download button that increments counter when clicked
        if st.download_button(
            "📄 Download SOW",
            data=st.session_state.file_data,
            file_name=os.path.basename(st.session_state.generated_file_path),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="download_sow"
        ):
            
            # Prepare record for storage
            sow_record = {
                "sow_num": sow_num,
                "sow_name": sow_name,
                "client": Client_Name,
                "project_type": option,
                "generated_date": generated_date,
                "start_date": start_str,
                "end_date": end_str,
            }

            # Project-type specific fields
            if option == "Fixed Fee":
                sow_record["fees"] = Fees_al
                # sow_record["milestones"] = milestone_df.to_json()

            if option == "T&M":
                sow_record["total_value"] = currency_value
                # sow_record["resources"] = resources_df.to_json()

            if option == "Change Order":
                sow_record["change_order"] = Change
                sow_record["fees_co"] = Fees_co
                sow_record["fees_sow"] = Fees_sow
                sow_record["difference"] = difference
                sow_record["sow_start_date"] = sow_str
                sow_record["sow_end_date"] = sow_end

            # SAVE HERE 👇
            save_sow_record(sow_record)

            # This block runs when download is clicked
            if st.session_state.should_increment_on_download:
                get_next_sow_number(peek_only=False)  # Increment counter
                st.session_state.should_increment_on_download = False  # Reset flag
                st.session_state.file_data = None  # Clear file data
                st.session_state.generated_file_path = None  # Clear file path
                reset_all_fields()  # Clear all input fields
                st.rerun()  # Refresh to show updated SOW number

# ----------------------------
# PAGE 2 — SOW RECORDS VIEWER
# ----------------------------
elif page == "SOW Records":
    st.title("📄 SOW Records Dashboard")

    file_path = "sow_records/sow_records.csv"

    if not os.path.exists(file_path):
        st.info("No SOW records found yet.")
    else:
        df = pd.read_csv(file_path)

        # Column filters for each project type
        column_filters = {
            "T&M": ["sow_num", "sow_name", "client", "project_type",
                    "generated_date", "start_date", "end_date", "total_value"],

            "Fixed Fee": ["sow_num", "sow_name", "client", "project_type",
                          "generated_date", "start_date", "end_date", "fees"],

            "Change Order": ["sow_num", "sow_name", "client", "project_type",
                             "generated_date", "start_date", "end_date",
                             "change_order", "fees_co", "fees_sow",
                             "difference", "sow_start_date", "sow_end_date"]
        }

        st.subheader("Filter by Project Type")

        col1, col2, col3 = st.columns([1, 1,3])

        selected_df = None
        selected_title = ""

        with col1:
            if st.button("T&M Records"):
                selected_title = "T&M Records"
                temp_df = df[df["project_type"] == "T&M"]
                selected_df = temp_df[column_filters["T&M"]]

        with col2:
            if st.button("Fixed Fee Records"):
                selected_title = "Fixed Fee Records"
                temp_df = df[df["project_type"] == "Fixed Fee"]
                selected_df = temp_df[column_filters["Fixed Fee"]]

        with col3:
            if st.button("Change Order Records"):
                selected_title = "Change Order Records"
                temp_df = df[df["project_type"] == "Change Order"]
                selected_df = temp_df[column_filters["Change Order"]]

        # Show filtered results below buttons
        if selected_df is not None:
            st.markdown(f"### {selected_title}")
            st.dataframe(selected_df if not selected_df.empty else "No data found.")
