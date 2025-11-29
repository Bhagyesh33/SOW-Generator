import streamlit as st
import os
import json
from datetime import datetime

def get_sow_approvals():
    folder = "sow_approvals"
    if not os.path.exists(folder):
        return []
    
    approvals = []
    for file in os.listdir(folder):
        if file.endswith("_metadata.json"):
            try:
                with open(os.path.join(folder, file), "r") as f:
                    metadata = json.load(f)
                    approvals.append(metadata)
            except:
                continue
    
    return sorted(approvals, key=lambda x: x.get('created_date', ''), reverse=True)

def update_approval_status(sow_number, status, approved_by="Legal Team"):
    metadata_file = f"sow_approvals/{sow_number}_metadata.json"
    
    if os.path.exists(metadata_file):
        with open(metadata_file, "r") as f:
            metadata = json.load(f)
        
        metadata["status"] = status
        metadata["approved_date"] = datetime.now().isoformat() if status == "approved" else None
        metadata["approved_by"] = approved_by if status == "approved" else None
        
        with open(metadata_file, "w") as f:
            json.dump(metadata, f, indent=2)
        
        return True
    return False

def show_approval_dashboard():
    st.title("📊 SOW Approval Dashboard")
    
    # Role selection
    st.session_state.user_role = st.selectbox(
        "Select Your Role",
        ["User", "Legal Team"],
        key="role_selector"
    )
    
    if st.button("🔄 Refresh Approvals"):
        st.rerun()
    
    approvals = get_sow_approvals()
    
    if not approvals:
        st.info("📝 No SOW submissions found. Generate some SOWs first!")
        return
    
    st.subheader(f"📋 SOW Requests ({len(approvals)} total)")
    
    # Filter options
    col1, col2, col3 = st.columns(3)
    with col1:
        show_all = st.checkbox("Show All", value=True)
    with col2:
        show_pending = st.checkbox("Show Pending", value=True)
    with col3:
        show_approved = st.checkbox("Show Approved", value=True)
    
    filtered_approvals = []
    for approval in approvals:
        status = approval.get('status', 'pending')
        if show_all or (show_pending and status == 'pending') or (show_approved and status == 'approved'):
            filtered_approvals.append(approval)
    
    if not filtered_approvals:
        st.warning("No SOWs match your filter criteria.")
        return
    
    # Display each approval
    for approval in filtered_approvals:
        with st.container():
            st.markdown("---")
            col1, col2, col3 = st.columns([3, 2, 2])
            
            with col1:
                st.subheader(f"{approval.get('sow_name', 'N/A')}")
                st.write(f"**SOW #:** {approval.get('sow_num', 'N/A')}")
                st.write(f"**Client:** {approval.get('client', 'N/A')}")
                st.write(f"**Type:** {approval.get('project_type', 'N/A')}")
                
                created_date = approval.get('created_date', '')
                if created_date:
                    try:
                        pretty_date = datetime.fromisoformat(created_date).strftime("%B %d, %Y %H:%M")
                        st.write(f"**Submitted:** {pretty_date}")
                    except:
                        st.write(f"**Submitted:** {created_date[:10]}")
            
            with col2:
                status = approval.get('status', 'pending')
                if status == 'pending':
                    st.warning("⏳ Pending Approval")
                elif status == 'approved':
                    st.success("✅ Approved")
                    approved_by = approval.get('approved_by')
                    approved_date = approval.get('approved_date')
                    if approved_by:
                        st.write(f"**By:** {approved_by}")
                    if approved_date:
                        try:
                            pretty_date = datetime.fromisoformat(approved_date).strftime("%b %d, %Y")
                            st.write(f"**On:** {pretty_date}")
                        except:
                            st.write(f"**On:** {approved_date[:10]}")
                else:
                    st.error("❌ Rejected")
            
            with col3:
                file_path = approval.get('file_path')
                if file_path and os.path.exists(file_path):
                    if status == 'approved' or st.session_state.user_role == "Legal Team":
                        with open(file_path, "rb") as f:
                            file_data = f.read()
                        
                        st.download_button(
                            "📥 Download SOW",
                            data=file_data,
                            file_name=approval.get('file_name', f"SOW_{approval.get('sow_num', 'N/A')}.docx"),
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"download_{approval.get('sow_num', '')}_{id(approval)}"
                        )
                    
                    if st.session_state.user_role == "Legal Team" and status == 'pending':
                        col_approve, col_reject = st.columns(2)
                        with col_approve:
                            if st.button("✅ Approve", key=f"approve_{approval.get('sow_num', '')}"):
                                if update_approval_status(approval['sow_num'], 'approved'):
                                    st.success("SOW Approved!")
                                    st.rerun()
                        with col_reject:
                            if st.button("❌ Reject", key=f"reject_{approval.get('sow_num', '')}"):
                                if update_approval_status(approval['sow_num'], 'rejected'):
                                    st.error("SOW Rejected!")
                                    st.rerun()
                else:
                    st.error("File not found")