import streamlit as st
from docxtpl import DocxTemplate
from datetime import datetime, date, timedelta
import pandas as pd
import os
import json
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module='pkg_resources')

# --- Shared Functions ---
def get_next_sow_number(peek_only=False):
    counter_file = "sow_counter.txt"
    start_num = 1000

    if not os.path.exists(counter_file):
        with open(counter_file, "w") as f:
            f.write(str(start_num))

    try:
        with open(counter_file, "r") as f:
            content = f.read().strip()
            current = int(content) if content else start_num
    except ValueError:
        current = start_num

    if peek_only:
        return current

    next_num = current + 1
    with open(counter_file, "w") as f:
        f.write(str(next_num))

    return current

def reset_all_fields():
    keys_to_keep = ['should_increment_on_download', 'generated_file_path', 'file_data', 'reset_trigger', 'user_role']
    keys_to_remove = [key for key in st.session_state.keys() if key not in keys_to_keep]
    for key in keys_to_remove:
        del st.session_state[key]
    st.session_state.reset_trigger += 1

def save_sow_to_local(sow_data, file_data):
    folder = "sow_approvals"
    os.makedirs(folder, exist_ok=True)
    
    file_name = f"{sow_data['sow_num']}_{sow_data['sow_name'].replace(' ', '_')}.docx"
    file_path = os.path.join(folder, file_name)
    with open(file_path, "wb") as f:
        f.write(file_data)
    
    metadata = {
        **sow_data,
        "file_path": file_path,
        "file_name": file_name,
        "status": "pending",
        "created_date": datetime.now().isoformat(),
        "approved_date": None,
        "approved_by": None
    }
    
    metadata_file = os.path.join(folder, f"{sow_data['sow_num']}_metadata.json")
    with open(metadata_file, "w") as f:
        json.dump(metadata, f, indent=2)
    
    return file_path

def networkdays(start_date, end_date):
    day_count = 0
    current = start_date
    while current <= end_date:
        if current.weekday() < 5:
            day_count += 1
        current += timedelta(days=1)
    return day_count

# --- Main SOW Generator Function ---
def show_sow_generator():
    st.markdown("### 📝 Generate New SOW")
    
    template_file = st.file_uploader("Upload client Word template (.docx)", type=["docx"], key=f"template_{st.session_state.reset_trigger}")

    # Form Fields
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
        if option in ["T&M", "Fixed Fee"]:
            auto_sow_num = get_next_sow_number(peek_only=True)
            sow_num = st.text_input(
                "SOW Number",
                str(auto_sow_num),
                key=f"sow_num_{st.session_state.reset_trigger}"
            )
        else:
            sow_num = st.text_input(
                "SOW Number",
                "",
                key=f"sow_num_{st.session_state.reset_trigger}",
                placeholder="Enter SOW number manually for Change Order"
            )

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
        pm_client = st.text_input("Client (Project Management)", key=f"pm_client_{st.session_state.reset_trigger}")
    with colB:
        pm_sp = st.text_input("Service Provider (Project Management)", key=f"pm_sp_{st.session_state.reset_trigger}")

    colA, colB = st.columns([1, 1])
    with colA:
        mg_client = st.text_input("Client (Management)", key=f"mg_client_{st.session_state.reset_trigger}")
    with colB:
        mg_sp = st.text_input("Service Provider (Management)", key=f"mg_sp_{st.session_state.reset_trigger}")

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
        
        try:
            difference = float(Fees_co) - float(Fees_sow)
        except:
            difference = 0

    additional_personnel = st.text_input(
        "Additional Personnel",
        key=f"additional_personnel_{st.session_state.reset_trigger}"
    )
    
    # Format dates
    generated_date = datetime.today().strftime("%B %d, %Y")
    start_str = start_date.strftime("%B %d, %Y")
    end_str = end_date.strftime("%B %d, %Y")
    if option == "Change Order":
        sow_str = sow_start_date.strftime("%B %d, %Y")
        sow_end = sow_end_date.strftime("%B %d, %Y")

    # Calculate working days
    workdays = networkdays(start_date, end_date)
    st.write(f"📅 Total working days (Mon–Fri) between selected dates: **{workdays}**")

    # Resources Table for T&M
    resources_df = pd.DataFrame()
    if option == "T&M":
        st.subheader("Resource Details")
        resources_df = st.data_editor(
            pd.DataFrame(
                columns=["Role", "Location", "Start Date", "End Date", "Allocation %", "Hrs/Day", "Rate/hr ($)"],
                data=[["Developer", "Remote", start_date, end_date, 100, 8, 100]]
            ),
            num_rows="dynamic",
            key="resources_table"
        )

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

        currency_value = resources_df["Estimated $"].sum()
        currency_value_str = f"${currency_value:,.2f}"
        st.write(f"💰 Total Contract Value: **{currency_value_str}**")

    # Milestone Schedule for Fixed Fee
    milestone_df = pd.DataFrame()
    if option == "Fixed Fee":
        st.subheader("Milestone Schedule / Payment Breakdown")
        try:
            total_fees = float(Fees_al) if Fees_al else 0
        except:
            total_fees = 0

        default_data = [
            ["1", "Initial Delivery", date.today(), "50"],
            ["2", "Final Delivery", date.today() + timedelta(days=30), "50"]
        ]

        milestone_input_df = st.data_editor(
            pd.DataFrame(
                default_data,
                columns=["Milestone #", "Services / Deliverables", "Milestone Due Date", "Payment Allocation (%)"]
            ),
            num_rows="dynamic",
            key="milestone_table"
        )

        milestone_df = milestone_input_df.copy()

        def calc_net(row):
            try:
                alloc = float(row["Payment Allocation (%)"]) if row["Payment Allocation (%)"] else 0
                return round(total_fees * (alloc / 100), 2)
            except:
                return 0

        milestone_df["Net Milestone Payment ($)"] = milestone_df.apply(calc_net, axis=1)
        st.dataframe(milestone_df)

        total_payment = milestone_df["Net Milestone Payment ($)"].sum()
        st.write(f"✅ Total Net Milestone Payment: **${total_payment:,.2f}**")
        
        milestone_df = milestone_df.rename(columns={
            "Milestone #": "milestone_no",
            "Services / Deliverables": "services",
            "Milestone Due Date": "due_date",
            "Payment Allocation (%)": "allocation",
            "Net Milestone Payment ($)": "net_pay"
        })

    # Generate SOW Button
    if st.button("🚀 Generate SOW Document", type="primary"):
        if template_file is None:
            st.warning("Please upload a Word template (.docx) before generating.")
        else:
            with st.spinner("Generating SOW document..."):
                template_path = os.path.join("generated_sows", "template.docx")
                os.makedirs("generated_sows", exist_ok=True)
                with open(template_path, "wb") as f:
                    f.write(template_file.getbuffer())

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
                    "generated_date": generated_date,
                    "additional_personnel": additional_personnel
                }

                if option == "T&M":
                    context.update({
                        "resources": resources_df.to_dict(orient="records"),
                        "currency_value_str": currency_value_str,
                        "currency_value": currency_value
                    })

                if option == "Fixed Fee":
                    context.update({
                        "milestones": milestone_df.to_dict(orient="records"),
                        "milestone_total": total_payment,
                        "Fees": Fees_al
                    })

                if option == "Change Order":
                    context.update({
                        "Change": Change,
                        "sow_end": sow_end,
                        "sow_str": sow_str,
                        "Fees_co": Fees_co,
                        "Fees_sow": Fees_sow,
                        "difference": difference
                    })

                doc = DocxTemplate(template_path)
                doc.render(context)

                output_file = os.path.join("generated_sows", f"{sow_num} - {sow_name}.docx")
                doc.save(output_file)

                with open(output_file, "rb") as f:
                    file_data = f.read()

                sow_data = {
                    'sow_num': sow_num,
                    'sow_name': sow_name,
                    'client': Client_Name,
                    'project_type': option,
                    'created_by': 'User'
                }

                file_path = save_sow_to_local(sow_data, file_data)
                
                st.success("✅ SOW generated and submitted for approval!")
                st.info("📋 The Legal Team will review and approve your SOW in the Approval Dashboard.")
                
                st.session_state.file_data = file_data
                st.session_state.generated_file_path = output_file
                st.session_state.should_increment_on_download = True

    # Quick Download Option
    if 'should_increment_on_download' in st.session_state and st.session_state.should_increment_on_download and st.session_state.file_data:
        st.markdown("---")
        st.warning("⚠️ Quick Download (Bypasses Approval)")
        if st.download_button(
            "📄 Download Local Copy",
            data=st.session_state.file_data,
            file_name=os.path.basename(st.session_state.generated_file_path),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="download_sow_local"
        ):
            if option in ["T&M", "Fixed Fee"]:
                get_next_sow_number(peek_only=False)
            
            st.session_state.should_increment_on_download = False
            st.session_state.file_data = None
            st.session_state.generated_file_path = None
            reset_all_fields()
            st.rerun()