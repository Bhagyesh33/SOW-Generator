import streamlit as st
import base64

# Import pages
from sow_generator import show_sow_generator
from approval import show_approval_dashboard
from records import show_sow_records

st.set_page_config(page_title="SOW Generator", layout="wide", page_icon="📋")
st.markdown("""
<style>
.css-1rs6os.edgvbvh3 { 
    display: none !important;
}
.block-container {
    padding-top: 0 !important;
}
header {visibility: hidden !important;}
</style>
""", unsafe_allow_html=True)

# --- Initialize session state ---
if 'reset_trigger' not in st.session_state:
    st.session_state.reset_trigger = 0
if 'user_role' not in st.session_state:
    st.session_state.user_role = 'user'

# --- Convert local logo to base64 ---
def get_base64_image(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except:
        return ""

logo_base64 = get_base64_image("logo-clbs- (1).png")

# --- Header ---
st.markdown(f"""
<style>
.header-full {{
    width: 100vw;
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

# --- Navigation ---
st.markdown("<br>", unsafe_allow_html=True)

with st.container():
    page = st.radio(
        "",
        ["SOW Generator", "Approval Dashboard", "SOW Records"],
        horizontal=True
    )

# --- Page Routing ---
if page == "SOW Generator":
    show_sow_generator()
elif page == "Approval Dashboard":
    show_approval_dashboard()
elif page == "SOW Records":
    show_sow_records()

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #666;'>SOW Generator v2.0</div>",
    unsafe_allow_html=True
)