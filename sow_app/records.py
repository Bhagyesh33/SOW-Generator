import streamlit as st
import pandas as pd
from approval import get_sow_approvals

def show_sow_records():
    st.title("📄 SOW Records")
    
    approvals = get_sow_approvals()
    
    if not approvals:
        st.info("No SOW records found yet.")
        return
    
    # Convert to DataFrame
    records_data = []
    for approval in approvals:
        records_data.append({
            'SOW Number': approval.get('sow_num', ''),
            'SOW Name': approval.get('sow_name', ''),
            'Client': approval.get('client', ''),
            'Project Type': approval.get('project_type', ''),
            'Status': approval.get('status', '').title(),
            'Created Date': approval.get('created_date', '')[:10],
            'Approved Date': approval.get('approved_date', '')[:10] if approval.get('approved_date') else 'N/A',
            'Approved By': approval.get('approved_by', 'N/A')
        })
    
    df = pd.DataFrame(records_data)
    
    # Filters
    col1, col2, col3 = st.columns(3)
    with col1:
        status_filter = st.selectbox("Filter by Status", ["All", "Pending", "Approved", "Rejected"])
    with col2:
        type_filter = st.selectbox("Filter by Type", ["All", "Fixed Fee", "T&M", "Change Order"])
    with col3:
        client_filter = st.selectbox("Filter by Client", ["All", "BSC", "Abiomed", "Cognex", "Itaros"])
    
    # Apply filters
    filtered_df = df.copy()
    if status_filter != "All":
        filtered_df = filtered_df[filtered_df['Status'] == status_filter]
    if type_filter != "All":
        filtered_df = filtered_df[filtered_df['Project Type'] == type_filter]
    if client_filter != "All":
        filtered_df = filtered_df[filtered_df['Client'] == client_filter]
    
    st.subheader(f"Records ({len(filtered_df)} found)")
    st.dataframe(filtered_df, use_container_width=True)
    
    # Export option
    if st.button("📊 Export to CSV"):
        csv = filtered_df.to_csv(index=False)
        st.download_button(
            "Download CSV",
            data=csv,
            file_name="sow_records.csv",
            mime="text/csv"
        )