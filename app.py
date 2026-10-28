import io
import os
import glob
from typing import Dict, Optional, Tuple

import pandas as pd
import streamlit as st

import tracker_engine

# =============================================================================
# STREAMLIT PAGE CONFIGURATION
# =============================================================================
st.set_page_config(
    page_title="Reachability Tracker Generator",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =============================================================================
# CUSTOM CSS FOR MODERN STYLING
# =============================================================================
st.markdown("""
<style>
    /* Global Sidebar Styling */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #090D16 0%, #0F172A 40%, #161F36 100%) !important;
        border-right: 1px solid rgba(148, 163, 184, 0.15) !important;
    }
    [data-testid="stSidebar"] .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 2rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }

    /* Sidebar Brand Card with Inlined SVG Icon */
    .brand-container {
        display: flex;
        align-items: center;
        gap: 12px;
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.8) 0%, rgba(15, 23, 42, 0.95) 100%);
        border: 1px solid rgba(56, 189, 248, 0.28);
        border-radius: 12px;
        padding: 10px 14px;
        margin-bottom: 12px;
        box-shadow: 0 4px 16px -2px rgba(0, 0, 0, 0.35);
    }
    .brand-icon-box {
        width: 40px;
        height: 40px;
        border-radius: 10px;
        background: linear-gradient(135deg, rgba(56, 189, 248, 0.2) 0%, rgba(99, 102, 241, 0.25) 100%);
        border: 1px solid rgba(56, 189, 248, 0.4);
        display: flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
        box-shadow: 0 0 10px rgba(56, 189, 248, 0.25);
    }
    .brand-text-box {
        display: flex;
        flex-direction: column;
    }
    .brand-title-text {
        font-size: 1.15rem;
        font-weight: 800;
        color: #F8FAFC;
        line-height: 1.15;
        letter-spacing: -0.3px;
    }
    .brand-subtitle-text {
        font-size: 0.73rem;
        color: #94A3B8;
        font-weight: 500;
        margin-top: 2px;
    }

    /* Sidebar Divider */
    .sidebar-glow-divider {
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(56, 189, 248, 0.3), transparent);
        margin: 14px 0;
        border: none;
    }

    /* Radio Button Styling in Sidebar */
    [data-testid="stSidebar"] [data-testid="stRadio"] > div {
        gap: 6px;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] label {
        background: rgba(30, 41, 59, 0.6) !important;
        border: 1px solid rgba(148, 163, 184, 0.18) !important;
        border-radius: 9px !important;
        padding: 8px 12px !important;
        margin: 0 !important;
        transition: all 0.2s ease !important;
        cursor: pointer !important;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] label:hover {
        background: rgba(56, 189, 248, 0.12) !important;
        border-color: rgba(56, 189, 248, 0.5) !important;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) {
        background: linear-gradient(90deg, rgba(56, 189, 248, 0.2) 0%, rgba(99, 102, 241, 0.2) 100%) !important;
        border-color: #38BDF8 !important;
        box-shadow: 0 0 10px rgba(56, 189, 248, 0.25) !important;
    }

    /* Required files card */
    .req-files-container {
        background: rgba(15, 23, 42, 0.6);
        border: 1px solid rgba(148, 163, 184, 0.15);
        border-radius: 10px;
        padding: 10px 12px;
        font-size: 0.78rem;
        color: #94A3B8;
        margin-top: 10px;
    }
    .req-file-item {
        display: flex;
        align-items: center;
        gap: 7px;
        margin-bottom: 6px;
        color: #CBD5E1;
    }
    .req-file-item:last-child {
        margin-bottom: 0;
    }
    .req-badge {
        background: rgba(56, 189, 248, 0.15);
        border: 1px solid rgba(56, 189, 248, 0.3);
        color: #38BDF8;
        padding: 2px 6px;
        border-radius: 5px;
        font-family: monospace;
        font-size: 0.72rem;
        font-weight: 600;
    }

    /* Main Page Headings */
    .main-title {
        font-size: 2.3rem;
        font-weight: 800;
        background: linear-gradient(90deg, #38BDF8 0%, #60A5FA 35%, #818CF8 70%, #C084FC 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.25rem;
        letter-spacing: -0.5px;
        line-height: 1.2;
    }
    .sub-title {
        font-size: 1.05rem;
        color: #94A3B8;
        font-weight: 400;
        margin-bottom: 1.5rem;
    }

    /* Metrics Cards */
    .metric-card {
        background: rgba(30, 41, 59, 0.65);
        border: 1px solid rgba(148, 163, 184, 0.25);
        border-radius: 12px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.15);
    }
    .metric-val {
        font-size: 1.85rem;
        font-weight: 800;
        background: linear-gradient(90deg, #38BDF8, #818CF8);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .metric-lbl {
        font-size: 0.85rem;
        color: #94A3B8;
        font-weight: 500;
        margin-top: 4px;
    }

    /* File Status Cards */
    .file-ready-box {
        background: rgba(16, 185, 129, 0.12);
        border: 1px solid rgba(16, 185, 129, 0.45);
        border-radius: 10px;
        padding: 14px 16px;
        margin-bottom: 10px;
    }
    .file-ready-title {
        color: #34D399;
        font-weight: 600;
        font-size: 0.95rem;
    }
    .file-ready-name {
        color: #A7F3D0;
        font-size: 0.82rem;
        word-break: break-all;
        margin-top: 4px;
        display: block;
    }
    .file-missing-box {
        background: rgba(239, 68, 68, 0.12);
        border: 1px solid rgba(239, 68, 68, 0.4);
        border-radius: 10px;
        padding: 14px 16px;
        margin-bottom: 10px;
    }
    .file-missing-title {
        color: #F87171;
        font-weight: 600;
        font-size: 0.95rem;
    }
    .file-missing-name {
        color: #FECACA;
        font-size: 0.82rem;
        margin-top: 4px;
        display: block;
    }

    /* File Error Card */
    .file-error-box {
        background: rgba(239, 68, 68, 0.15);
        border: 1px solid rgba(239, 68, 68, 0.55);
        border-radius: 10px;
        padding: 14px 16px;
        margin-bottom: 10px;
        box-shadow: 0 4px 12px rgba(239, 68, 68, 0.15);
    }
    .file-error-title {
        color: #F87171;
        font-weight: 700;
        font-size: 0.95rem;
    }
    .file-error-name {
        color: #FECACA;
        font-size: 0.82rem;
        word-break: break-all;
        margin-top: 4px;
        display: block;
    }
    .file-error-detail {
        margin-top: 8px;
        padding: 6px 10px;
        border-radius: 6px;
        background: rgba(15, 23, 42, 0.6);
        border-left: 3px solid #EF4444;
        font-size: 0.74rem;
        color: #FCA5A5;
        line-height: 1.4;
    }

    /* Verification Banners */
    .validation-success-banner {
        background: linear-gradient(135deg, rgba(16, 185, 129, 0.15) 0%, rgba(5, 150, 105, 0.22) 100%);
        border: 1px solid rgba(16, 185, 129, 0.45);
        border-radius: 10px;
        padding: 14px 18px;
        margin: 14px 0 16px 0;
        color: #D1FAE5;
        font-size: 0.92rem;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .validation-blocked-banner {
        background: linear-gradient(135deg, rgba(239, 68, 68, 0.18) 0%, rgba(185, 28, 28, 0.25) 100%);
        border: 1px solid rgba(239, 68, 68, 0.55);
        border-radius: 10px;
        padding: 16px 20px;
        margin: 14px 0 16px 0;
        color: #FEE2E2;
        font-size: 0.92rem;
    }

    /* Download Button */
    .stDownloadButton > button {
        width: 100%;
        background-color: #16A34A !important;
        color: white !important;
        font-size: 1.1rem !important;
        font-weight: 600 !important;
        padding: 0.6rem 1.2rem !important;
        border-radius: 8px !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(22, 163, 74, 0.3) !important;
        transition: all 0.2s ease !important;
    }
    .stDownloadButton > button:hover {
        background-color: #15803D !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 16px rgba(22, 163, 74, 0.4) !important;
    }
</style>
""", unsafe_allow_html=True)


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================
def get_sample_files(sample_dir: str) -> Dict[str, Optional[str]]:
    """Scan sample folder for the 3 target files."""
    files = {"ciq": None, "data": None, "ims": None}
    if not os.path.exists(sample_dir):
        return files

    ciq_hits = glob.glob(os.path.join(sample_dir, "*CIQ*.xlsx"))
    data_hits = glob.glob(os.path.join(sample_dir, "*UP-DATA-IP-Plan*.xlsx"))
    ims_hits = glob.glob(os.path.join(sample_dir, "*UP-IMS-IP-Plan*.xlsx"))

    if ciq_hits:
        files["ciq"] = ciq_hits[0]
    if data_hits:
        files["data"] = data_hits[0]
    if ims_hits:
        files["ims"] = ims_hits[0]
    return files


def classify_uploaded_files(uploaded_list) -> Tuple[Dict[str, any], List[any]]:
    """Auto-detect which uploaded file is CIQ, UP-DATA, or UP-IMS."""
    classified = {"ciq": None, "data": None, "ims": None}
    unclassified = []

    for f in uploaded_list:
        name_lower = f.name.lower()
        if "ciq" in name_lower:
            if classified["ciq"] is None:
                classified["ciq"] = f
            else:
                unclassified.append(f)
        elif "ims" in name_lower and "ip-plan" in name_lower:
            if classified["ims"] is None:
                classified["ims"] = f
            else:
                unclassified.append(f)
        elif "data" in name_lower and "ip-plan" in name_lower:
            if classified["data"] is None:
                classified["data"] = f
            else:
                unclassified.append(f)
        else:
            unclassified.append(f)

    # For any remaining unclassified files, read workbook sheets
    remaining_unclassified = []
    for f in unclassified:
        try:
            import openpyxl
            if hasattr(f, "seek"):
                f.seek(0)
            wb = openpyxl.load_workbook(f, read_only=True)
            ftype = tracker_engine.detect_file_type(wb, f.name)
            wb.close()
            if hasattr(f, "seek"):
                f.seek(0)
            if ftype in classified and classified[ftype] is None:
                classified[ftype] = f
            else:
                remaining_unclassified.append(f)
        except Exception:
            remaining_unclassified.append(f)

    return classified, remaining_unclassified


# =============================================================================
# SIDEBAR
# =============================================================================
with st.sidebar:
    # Inlined SVG Telecom Switch / Core Emblem (Zero broken images, completely offline safe)
    st.markdown("""
    <div class="brand-container">
        <div class="brand-icon-box">
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#38BDF8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <rect x="2" y="2" width="20" height="8" rx="2" ry="2"></rect>
                <rect x="2" y="14" width="20" height="8" rx="2" ry="2"></rect>
                <line x1="6" y1="6" x2="6.01" y2="6"></line>
                <line x1="10" y1="6" x2="10.01" y2="6"></line>
                <line x1="6" y1="18" x2="6.01" y2="18"></line>
                <line x1="10" y1="18" x2="10.01" y2="18"></line>
                <circle cx="17" cy="6" r="1.5" fill="#34D399" stroke="none"></circle>
                <circle cx="17" cy="18" r="1.5" fill="#38BDF8" stroke="none"></circle>
            </svg>
        </div>
        <div class="brand-text-box">
            <div class="brand-title-text">Settings &amp; Config</div>
            <div class="brand-subtitle-text">Nokia CMG Reachability Engine</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("Configure source options, site identity, and DCGW VPRN parameters.")

    available_samples = []
    if os.path.isdir("Data 1"):
        available_samples.append("Data 1 (Karnal-2)")
    if os.path.isdir("Data 2"):
        available_samples.append("Data 2 (Mohali)")
    if os.path.isdir("Data 3"):
        available_samples.append("Data 3 (Ashok Marg / Lucknow)")

    source_options = ["📁 Upload Excel Files"]
    if available_samples:
        source_options.append("🧪 Load Sample Dataset")

    input_mode = st.radio(
        "Source Mode",
        source_options,
        help="Choose whether to upload custom Excel files or load bundled test datasets."
    )

    sample_selection = None
    if input_mode == "🧪 Load Sample Dataset" and available_samples:
        sample_selection = st.selectbox(
            "Select Sample Dataset",
            available_samples
        )

    st.markdown('<div class="sidebar-glow-divider"></div>', unsafe_allow_html=True)
    st.subheader("Circle / Regional Profile")
    circle_selection = st.selectbox(
        "Telecom Circle",
        [
            "Auto-Detect from Site / Files",
            "Karnal / Haryana",
            "Mohali / Punjab",
            "Lucknow / UP East",
            "Delhi",
            "Gujarat",
            "Mumbai",
            "Maharashtra",
            "Generic / Other Circle"
        ],
        help="Select your circle region to auto-populate default eNodeB, DNS, and router destinations."
    )

    st.subheader("Site Information")
    site_name_override = st.text_input(
        "Site / Sheet Name",
        value="",
        placeholder="Auto-detected if empty",
        help="Will be used as the active worksheet tab name and in the report header."
    )

    with st.expander("⚙️ Advanced Parameters (Manual Overrides)", expanded=False):
        use_manual_overrides = st.checkbox("Enable manual parameter overrides", value=False,
                                           help="If unchecked, parameters auto-adapt to the detected circle and CIQ data.")
        st.markdown("**UE Pool Settings**")
        sgi_pool_val = st.number_input("SGi Public Pool Number", min_value=1, max_value=50, value=2)
        sgi_v4_line_val = st.number_input("SGi IPv4 Pool Line (0-based)", min_value=0, max_value=20, value=1)
        gm_v6_line_val = st.number_input("Gm IMS Pool Line (0-based)", min_value=0, max_value=20, value=1)

        st.markdown("**Test Nodes**")
        s1u_test_enb = st.text_input("S1-U Test eNB IP", value="100.91.17.85")
        ntp_ip1 = st.text_input("NTP Server 1", value="192.168.130.1")
        ntp_ip2 = st.text_input("NTP Server 2", value="192.168.130.2")

        st.markdown("**DCGW VPRNs**")
        col_v1, col_v2 = st.columns(2)
        with col_v1:
            dcgw_ntp = st.number_input("NTP VPRN", value=1000)
            dcgw_sx = st.number_input("SX VPRN", value=1004)
            dcgw_s11 = st.number_input("S11 VPRN", value=1004)
            dcgw_s1u = st.number_input("S1-U VPRN", value=1004)
            dcgw_sgi = st.number_input("SGi VPRN", value=1001)
        with col_v2:
            dcgw_gx = st.number_input("Gx VPRN", value=1000)
            dcgw_gy = st.number_input("Gy VPRN", value=1000)
            dcgw_bp = st.number_input("Bp VPRN", value=1000)
            dcgw_gm = st.number_input("Gm VPRN", value=1003)
            dcgw_s5s8 = st.number_input("S5-S8 VPRN", value=1004)

    st.markdown('<div class="sidebar-glow-divider"></div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="req-files-container">
        <div style="font-weight:700; color:#E2E8F0; margin-bottom:6px;">Required 3 Source Files:</div>
        <div class="req-file-item">
            <span class="req-badge">1. CIQ</span>
            <span>Nokia CMG (*CIQ*.xlsx)</span>
        </div>
        <div class="req-file-item">
            <span class="req-badge">2. DATA</span>
            <span>UP-DATA (*UP-DATA-IP-Plan*.xlsx)</span>
        </div>
        <div class="req-file-item">
            <span class="req-badge">3. IMS</span>
            <span>UP-IMS (*UP-IMS-IP-Plan*.xlsx)</span>
        </div>
        <div style="margin-top:8px; font-size:0.73rem; color:#7DD3FC; border-top:1px dashed rgba(148,163,184,0.2); padding-top:6px;">
            💡 Check "Expected File Formats" on the main page for sheet details.
        </div>
    </div>
    """, unsafe_allow_html=True)


# =============================================================================
# MAIN CONTENT
# =============================================================================
st.markdown('<div class="main-title">📡 Reachability Tracker Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Automate the creation of Nokia CMG Reachability Workbooks directly from CIQ and IP-Plan Excel sheets.</div>', unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# GENERAL EXPECTED FORMAT FOR 3 INPUT FILES
# -----------------------------------------------------------------------------
with st.expander("📋 General Expected Format for the 3 Input Files (CIQ, DATA IP-Plan, IMS IP-Plan)", expanded=False):
    st.markdown("""
    To ensure the generator accurately extracts all CMG interfaces, peer IPs, service VPRNs, and command scripts, the **3 input Excel files** should follow standard Nokia CMG engineering templates:
    """)

    fmt_tab1, fmt_tab2, fmt_tab3, fmt_tab4 = st.tabs([
        "📄 1. Nokia CMG CIQ",
        "📄 2. UP-DATA IP-Plan",
        "📄 3. UP-IMS IP-Plan",
        "⚙️ Auto-Classification & Rules"
    ])

    with fmt_tab1:
        st.markdown("""
        #### 1. Nokia CMG CIQ (`*CIQ*.xlsx`)
        - **Filename Pattern**: Must contain `CIQ` (e.g. `VIL_Modernization-<Site>-CMG-*-CIQ*.xlsx`).
        - **Primary Function**: Customer Information Questionnaire containing external peer IP addresses, Diameter endpoints, eNodeB transport configurations, and UE IP pool allocations.
        - **Mandatory Sheets & Extracted Parameters**:
        """)
        ciq_fmt_df = pd.DataFrame([
            {"Sheet Name": "S11", "Starting Row": "Row 5+", "Key Columns": "Col A: Interface, Col C: CMG SRC IP, Col F: Peer Interface, Col H: MME Peer IP", "Extracted Parameters": "CMG S11 Source IP, MME Destination IPs"},
            {"Sheet Name": "Gx", "Starting Row": "Row 4+", "Key Columns": "Row 'Gx-Peer-IP-1', 'CMG Diameter-Lo-1', Col headers matching 'DRA' & 'DATA'", "Extracted Parameters": "PCRF / DRA Peer IPs, CMG Diameter Loopback"},
            {"Sheet Name": "Gy", "Starting Row": "Row 5+", "Key Columns": "Row keys 'Gy-Peer-IP-*', Col headers matching 'RECC' & 'DRA'", "Extracted Parameters": "OCS / DRA Peer IPs, RECC server IPs"},
            {"Sheet Name": "S6b", "Starting Row": "Row 4+", "Key Columns": "Row 's6b-Peer-IP-1', Col headers matching 'DRA'", "Extracted Parameters": "AAA / HSS Diameter Peer IPs"},
            {"Sheet Name": "Bp", "Starting Row": "Row 5+", "Key Columns": "Rows with Billing Mediation server IPs", "Extracted Parameters": "Offline Billing Mediation Collector IPs"},
            {"Sheet Name": "X1-X2-X3", "Starting Row": "Row 4+", "Key Columns": "Col A: NE Type, Col B: Interface (X1/X2/X3), Col C: CMG SRC, Col D: DST IP", "Extracted Parameters": "Lawful Interception (ADMF, MDF2, MDF3) IPs"},
            {"Sheet Name": "S1U-eNb", "Starting Row": "Row 6+", "Key Columns": "Col B: NE Type ('DATA' or 'IMS'), Col C: CMG Source IP", "Extracted Parameters": "S1-U UPF01 (DATA) and UPF02 (IMS) Source IPs"},
            {"Sheet Name": "Gn-Gp-S5-S8", "Starting Row": "Row 6+", "Key Columns": "Col A: Node Name, Col B: NE Type ('DATA'/'IMS'), Col C: CMG SRC / Peer IP", "Extracted Parameters": "SGW-C, PGW-C, SGW-U, PGW-U peer interfaces"},
            {"Sheet Name": "S2b", "Starting Row": "Row 6+", "Key Columns": "Col A: Interface ('S2b-C'/'S2b-U'), Col C: Source IPv4, Col F: Peer Iface, Col H: Peer IPv4", "Extracted Parameters": "ePDG / VoWiFi Signaling & Media endpoints"},
            {"Sheet Name": "P-CSCF_SBC", "Starting Row": "Row 6+", "Key Columns": "Col F: Signaling IP (IPv6), Col J: Media IP (IPv6 host prefix)", "Extracted Parameters": "P-CSCF SBC Gm Signaling & Media Peer IPs"},
            {"Sheet Name": "UE IP pool", "Starting Row": "Any Row", "Key Columns": "Rows containing 'Public Pool' & 'UE Pool' / 'UE Pool IMS'", "Extracted Parameters": "SGi Public IPv4/IPv6 Subnets, Gm IMS Voice IPv6 Subnet"}
        ])
        st.dataframe(ciq_fmt_df, width="stretch", hide_index=True)

    with fmt_tab2:
        st.markdown("""
        #### 2. UP-DATA IP-Plan (`*UP-DATA-IP-Plan*.xlsx`)
        - **Filename Pattern**: Must contain `DATA` and `IP-Plan` (e.g. `VIL_Modernization-<Site>-*-UP-DATA-IP-Plan*.xlsx`).
        - **Primary Function**: Detailed IP address plan for CPF1 (Control Plane) and UPF1 (Data User Plane).
        - **Mandatory Sheets & Extracted Parameters**:
        """)
        data_fmt_df = pd.DataFrame([
            {"Sheet Name": "External Interfaces", "Key Columns": "Col D: CP Interface Name, Col F: CP IP Address, Col I: UPF Interface Name, Col K: UPF IP Address", "Extracted Parameters": "CPF1 & UPF1 SYSTEM IP, Gx, Gy, and Sx interface loopbacks"},
            {"Sheet Name": "CP-SIGNAL", "Key Columns": "Col G: Service ID / VPRN number", "Extracted Parameters": "Signaling VPRN Service ID (e.g. 10)"},
            {"Sheet Name": "CP-BILLING", "Key Columns": "Col G: Service ID / VPRN number", "Extracted Parameters": "Billing VPRN Service ID (e.g. 11)"},
            {"Sheet Name": "CP-LI", "Key Columns": "Col G: Service ID / VPRN number", "Extracted Parameters": "Lawful Interception VPRN Service ID (e.g. 12)"},
            {"Sheet Name": "CP-CORE", "Key Columns": "Col G: Service ID / VPRN number", "Extracted Parameters": "Control Plane Core Transport VPRN (e.g. 100)"},
            {"Sheet Name": "UP-CORE", "Key Columns": "Col G: Service ID / VPRN numbers (rows 2+)", "Extracted Parameters": "Data UPF User Plane routing groups and transport VPRNs"},
            {"Sheet Name": "NAT_13XX_INSIDE VPRN", "Key Columns": "Col G: Service ID, Col J: Interface Name ('NAT-Inside-Grp*-PORT-1')", "Extracted Parameters": "SGi NAT inside VPRN (e.g. 1301)"}
        ])
        st.dataframe(data_fmt_df, width="stretch", hide_index=True)

    with fmt_tab3:
        st.markdown("""
        #### 3. UP-IMS IP-Plan (`*UP-IMS-IP-Plan*.xlsx`)
        - **Filename Pattern**: Must contain `IMS` and `IP-Plan` (e.g. `VIL_Modernization-<Site>-*-UP-IMS-IP-Plan*.xlsx`).
        - **Primary Function**: Engineering IP address plan for UPF2 (IMS Voice Plane Function).
        - **Mandatory Sheets & Extracted Parameters**:
        """)
        ims_fmt_df = pd.DataFrame([
            {"Sheet Name": "External Interfaces", "Key Columns": "Col I: UPF Interface Name, Col K: UPF IP Address", "Extracted Parameters": "UPF2 SYSTEM IP, Sx loopback, and Gm voice interface IPs"},
            {"Sheet Name": "UP-CORE", "Key Columns": "Col G: Service ID / VPRN numbers (rows 2+)", "Extracted Parameters": "IMS UPF User Plane routing groups and transport VPRNs"},
            {"Sheet Name": "UP-IMS", "Key Columns": "Col G: Service ID / VPRN numbers", "Extracted Parameters": "Gm Voice VPRN Service ID (e.g. 200)"},
            {"Sheet Name": "UE Pool (Optional)", "Key Columns": "Supplementary pool definitions", "Extracted Parameters": "Voice pool references"}
        ])
        st.dataframe(ims_fmt_df, width="stretch", hide_index=True)

    with fmt_tab4:
        st.markdown("""
        #### ⚙️ Auto-Classification & General Guidelines
        - **File Format**: Standard Excel openxml files (`.xlsx` or `.xlsm`). Files must **not** be password-encrypted.
        - **Automatic Classification Logic**:
          - Files containing `ciq` in their filename or containing sheets `['s1u-enb', 'p-cscf_sbc', 'ue ip pool', 'gx']` are classified as **CIQ**.
          - Files containing `data` and `ip-plan` or containing sheets `['cp-signal', 'nat_13xx_inside vprn']` are classified as **UP-DATA IP-Plan**.
          - Files containing `ims` and `ip-plan` or containing sheet `['up-ims']` are classified as **UP-IMS IP-Plan**.
        - **Site Name Auto-Detection**:
          - Detected from filename pattern `Modernization-<Site>-...` (e.g. `Mohali-5`, `Karnal-2`).
          - You can override the detected site name anytime via the **Site / Sheet Name** input in the sidebar.
        - **Circle Adaptation & Self-Healing Multi-Circle Engine**:
          - **Resilient Sheet Discovery**: Automatically matches sheet names even with variations, spaces, underscores, or circle suffixes (e.g. `P_CSCF_SBC_JK`, `UE IP pool `, `S1U-eNb_DL`, `X1_X2_X3`).
          - **Dynamic Column & Row Header Detection**: Interface loopbacks, service IDs, and peer IPs are detected dynamically by analyzing header rows and cell formats rather than rigid column numbers.
          - **Multi-Circle Profiles**: Supports **Karnal/Haryana**, **Mohali/Punjab**, **Delhi**, **Gujarat**, **Mumbai**, **Maharashtra**, and **Generic Circle** profiles with full override control.
          - **CIQ Priority Auto-Extraction**: If your CIQ file contains circle-specific destination IPs in `S1U-eNb` (test eNodeB IP) or `Gi DNS` (DNS servers) or `CGNAT`, the engine automatically prioritizes those actual circle addresses!
        """)

ciq_file = None
data_file = None
ims_file = None
detected_site = "Site"
fn_map = {}

# -----------------------------------------------------------------------------
# MODE 1: UPLOAD FILES
# -----------------------------------------------------------------------------
if input_mode == "📁 Upload Excel Files":
    upload_tab1, upload_tab2 = st.tabs(["⚡ Drag & Drop All 3 Files (Auto-Detect)", "📋 Individual File Uploaders"])

    with upload_tab1:
        st.markdown("Drag and drop all **3 Excel files** at once. The system will automatically classify each file:")
        all_uploaded = st.file_uploader(
            "Upload 3 Excel files (CIQ, DATA IP-Plan, IMS IP-Plan)",
            type=["xlsx", "xlsm"],
            accept_multiple_files=True,
            key="multi_upload"
        )
        if all_uploaded:
            classified, unclassified = classify_uploaded_files(all_uploaded)
            ciq_file = classified.get("ciq")
            data_file = classified.get("data")
            ims_file = classified.get("ims")
            if unclassified:
                unf_names = ", ".join(f"`{f.name}`" for f in unclassified)
                st.warning(f"⚠️ Unclassified or duplicate file(s) ignored: {unf_names}. Expected 1 CIQ, 1 UP-DATA, and 1 UP-IMS file.")

    with upload_tab2:
        st.markdown("Or upload each file individually into its designated field:")
        col1, col2, col3 = st.columns(3)
        with col1:
            u_ciq = st.file_uploader("1. Nokia CMG CIQ (*CIQ*.xlsx)", type=["xlsx", "xlsm"], key="ind_ciq")
            if u_ciq:
                ciq_file = u_ciq
        with col2:
            u_data = st.file_uploader("2. UP-DATA IP Plan (*DATA*.xlsx)", type=["xlsx", "xlsm"], key="ind_data")
            if u_data:
                data_file = u_data
        with col3:
            u_ims = st.file_uploader("3. UP-IMS IP Plan (*IMS*.xlsx)", type=["xlsx", "xlsm"], key="ind_ims")
            if u_ims:
                ims_file = u_ims

# -----------------------------------------------------------------------------
# MODE 2: LOAD SAMPLE DATASET
# -----------------------------------------------------------------------------
else:
    sample_dir = None
    if sample_selection:
        if "Data 1" in sample_selection:
            sample_dir = "Data 1"
        elif "Data 2" in sample_selection:
            sample_dir = "Data 2"
        elif "Data 3" in sample_selection:
            sample_dir = "Data 3"

    if sample_dir and os.path.isdir(sample_dir):
        s_files = get_sample_files(sample_dir)
        ciq_file = s_files["ciq"]
        data_file = s_files["data"]
        ims_file = s_files["ims"]
        st.info(f"Loaded bundled dataset from `{sample_dir}` directory.")
    else:
        st.warning("⚠️ No bundled sample dataset found. Please switch to '📁 Upload Excel Files' above.")


# -----------------------------------------------------------------------------
# FILE STATUS & REAL-TIME FORMAT VALIDATION CARDS
# -----------------------------------------------------------------------------
st.markdown("#### Source Files Status & Format Validation")
c1, c2, c3 = st.columns(3)

val_ciq = None
val_data = None
val_ims = None

with c1:
    if ciq_file:
        name = os.path.basename(ciq_file) if isinstance(ciq_file, str) else ciq_file.name
        fn_map["ciq"] = name
        detected_site = tracker_engine.extract_site_name(name)
        val_ciq = tracker_engine.validate_source_file(ciq_file, "ciq", filename=name)
        if val_ciq["is_valid"]:
            st.markdown(f"""
            <div class="file-ready-box">
                <div class="file-ready-title">✅ 1. Nokia CMG CIQ <span style="font-size:0.72rem; background:rgba(16,185,129,0.25); color:#6EE7B7; padding:2px 6px; border-radius:4px; margin-left:6px; font-weight:600;">Format Validated</span></div>
                <span class="file-ready-name">{name}</span>
                <div style="font-size:0.75rem; color:#A7F3D0; margin-top:5px;">
                    ✔ All {len(val_ciq['mandatory_sheets'])} mandatory sheets verified
                </div>
            </div>
            """, unsafe_allow_html=True)
            with st.expander("🔍 View Verified Sheets", expanded=False):
                st.caption(f"**Sheets found:** {', '.join(val_ciq['verified_sheets'])}")
        else:
            err_html = "<br>• ".join(val_ciq["errors"])
            st.markdown(f"""
            <div class="file-error-box">
                <div class="file-error-title">❌ 1. Nokia CMG CIQ <span style="font-size:0.72rem; background:rgba(239,68,68,0.3); color:#FCA5A5; padding:2px 6px; border-radius:4px; margin-left:6px; font-weight:600;">Format Error</span></div>
                <span class="file-error-name">{name}</span>
                <div class="file-error-detail">
                    <strong>Format Check Failed:</strong><br>• {err_html}
                </div>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="file-missing-box">
            <div class="file-missing-title">⏳ 1. Nokia CMG CIQ</div>
            <span class="file-missing-name">Awaiting file (*CIQ*.xlsx)</span>
        </div>
        """, unsafe_allow_html=True)

with c2:
    if data_file:
        name = os.path.basename(data_file) if isinstance(data_file, str) else data_file.name
        fn_map["data"] = name
        val_data = tracker_engine.validate_source_file(data_file, "data", filename=name)
        if val_data["is_valid"]:
            st.markdown(f"""
            <div class="file-ready-box">
                <div class="file-ready-title">✅ 2. UP-DATA IP-Plan <span style="font-size:0.72rem; background:rgba(16,185,129,0.25); color:#6EE7B7; padding:2px 6px; border-radius:4px; margin-left:6px; font-weight:600;">Format Validated</span></div>
                <span class="file-ready-name">{name}</span>
                <div style="font-size:0.75rem; color:#A7F3D0; margin-top:5px;">
                    ✔ All {len(val_data['mandatory_sheets'])} mandatory sheets verified
                </div>
            </div>
            """, unsafe_allow_html=True)
            with st.expander("🔍 View Verified Sheets", expanded=False):
                st.caption(f"**Sheets found:** {', '.join(val_data['verified_sheets'])}")
        else:
            err_html = "<br>• ".join(val_data["errors"])
            st.markdown(f"""
            <div class="file-error-box">
                <div class="file-error-title">❌ 2. UP-DATA IP-Plan <span style="font-size:0.72rem; background:rgba(239,68,68,0.3); color:#FCA5A5; padding:2px 6px; border-radius:4px; margin-left:6px; font-weight:600;">Format Error</span></div>
                <span class="file-error-name">{name}</span>
                <div class="file-error-detail">
                    <strong>Format Check Failed:</strong><br>• {err_html}
                </div>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="file-missing-box">
            <div class="file-missing-title">⏳ 2. UP-DATA IP-Plan</div>
            <span class="file-missing-name">Awaiting file (*UP-DATA-IP-Plan*.xlsx)</span>
        </div>
        """, unsafe_allow_html=True)

with c3:
    if ims_file:
        name = os.path.basename(ims_file) if isinstance(ims_file, str) else ims_file.name
        fn_map["ims"] = name
        val_ims = tracker_engine.validate_source_file(ims_file, "ims", filename=name)
        if val_ims["is_valid"]:
            st.markdown(f"""
            <div class="file-ready-box">
                <div class="file-ready-title">✅ 3. UP-IMS IP-Plan <span style="font-size:0.72rem; background:rgba(16,185,129,0.25); color:#6EE7B7; padding:2px 6px; border-radius:4px; margin-left:6px; font-weight:600;">Format Validated</span></div>
                <span class="file-ready-name">{name}</span>
                <div style="font-size:0.75rem; color:#A7F3D0; margin-top:5px;">
                    ✔ All {len(val_ims['mandatory_sheets'])} mandatory sheets verified
                </div>
            </div>
            """, unsafe_allow_html=True)
            with st.expander("🔍 View Verified Sheets", expanded=False):
                st.caption(f"**Sheets found:** {', '.join(val_ims['verified_sheets'])}")
        else:
            err_html = "<br>• ".join(val_ims["errors"])
            st.markdown(f"""
            <div class="file-error-box">
                <div class="file-error-title">❌ 3. UP-IMS IP-Plan <span style="font-size:0.72rem; background:rgba(239,68,68,0.3); color:#FCA5A5; padding:2px 6px; border-radius:4px; margin-left:6px; font-weight:600;">Format Error</span></div>
                <span class="file-error-name">{name}</span>
                <div class="file-error-detail">
                    <strong>Format Check Failed:</strong><br>• {err_html}
                </div>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="file-missing-box">
            <div class="file-missing-title">⏳ 3. UP-IMS IP-Plan</div>
            <span class="file-missing-name">Awaiting file (*UP-IMS-IP-Plan*.xlsx)</span>
        </div>
        """, unsafe_allow_html=True)

# Final site name decision
final_site = site_name_override.strip() if site_name_override.strip() else detected_site


def resolve_circle_profile(circle_sel: str, site_str: str) -> Tuple[str, Dict[str, Any]]:
    c_low = circle_sel.lower()
    if "haryana" in c_low or "karnal" in c_low:
        return "Karnal / Haryana", tracker_engine.CIRCLE_PROFILES["karnal"].copy()
    if "punjab" in c_low or "mohali" in c_low:
        return "Mohali / Punjab", tracker_engine.CIRCLE_PROFILES["mohali"].copy()
    if "lucknow" in c_low or "up east" in c_low or "upeast" in c_low or "ashok" in c_low:
        return "Lucknow / UP East", tracker_engine.CIRCLE_PROFILES["lucknow"].copy()
    if "delhi" in c_low:
        return "Delhi", tracker_engine.CIRCLE_PROFILES["delhi"].copy()
    if "gujarat" in c_low:
        return "Gujarat", tracker_engine.CIRCLE_PROFILES["gujarat"].copy()
    if "mumbai" in c_low:
        return "Mumbai", tracker_engine.CIRCLE_PROFILES["mumbai"].copy()
    if "maharashtra" in c_low:
        return "Maharashtra", tracker_engine.CIRCLE_PROFILES["maharashtra"].copy()
    if "generic" in c_low or "other" in c_low:
        return "Generic / Other", tracker_engine.CIRCLE_PROFILES["generic"].copy()
    # Auto-detect from site name
    prof = tracker_engine.get_circle_profile(site_str)
    s_low = site_str.lower()
    if "mohali" in s_low or "punjab" in s_low:
        lbl = "Mohali / Punjab (Auto-detected)"
    elif "karnal" in s_low or "haryana" in s_low:
        lbl = "Karnal / Haryana (Auto-detected)"
    elif "ashok" in s_low or "lucknow" in s_low or "up east" in s_low or "upeast" in s_low or "up-east" in s_low:
        lbl = "Lucknow / UP East (Auto-detected)"
    elif "delhi" in s_low or "dl" in s_low.split():
        lbl = "Delhi (Auto-detected)"
    elif "gujarat" in s_low or "ahmedabad" in s_low or "gj" in s_low.split():
        lbl = "Gujarat (Auto-detected)"
    elif "mumbai" in s_low:
        lbl = "Mumbai (Auto-detected)"
    else:
        lbl = "Generic Profile (Auto-detected)"
    return lbl, prof


active_circle_label, active_circle_profile = resolve_circle_profile(circle_selection, final_site)

# -----------------------------------------------------------------------------
# GENERATION ACTION & FORMAT ENFORCEMENT
# -----------------------------------------------------------------------------
all_ready = (ciq_file is not None and data_file is not None and ims_file is not None)
all_valid = all_ready and bool(val_ciq and val_ciq["is_valid"]) and bool(val_data and val_data["is_valid"]) and bool(val_ims and val_ims["is_valid"])

st.markdown("<br>", unsafe_allow_html=True)

if not all_ready:
    st.warning("⚠️ Please provide all 3 Excel files above to enable reachability tracker generation.")
elif all_ready and not all_valid:
    # Clear any previous generation results so invalid states do not show stale workbooks
    if "out_buf" in st.session_state:
        for k in ["out_buf", "stats", "warnings", "summary_data", "detail_data", "final_site"]:
            st.session_state.pop(k, None)

    total_errors = []
    if val_ciq and not val_ciq["is_valid"]:
        total_errors.extend(val_ciq["errors"])
    if val_data and not val_data["is_valid"]:
        total_errors.extend(val_data["errors"])
    if val_ims and not val_ims["is_valid"]:
        total_errors.extend(val_ims["errors"])

    err_list_html = "".join(f"<li style='margin-bottom:6px;'>{e}</li>" for e in total_errors)
    st.markdown(f"""
    <div class="validation-blocked-banner">
        <div style="font-weight:700; font-size:1.05rem; display:flex; align-items:center; gap:8px;">
            <span style="font-size:1.3rem;">🛑</span> Generation Stopped: File Format Check Failed
        </div>
        <div style="margin-top:6px; font-size:0.85rem; color:#FEE2E2;">
            The uploaded files do not match the expected Nokia CMG templates. To ensure data accuracy and prevent producing corrupted workbooks, <strong>generation is stopped</strong>. Please fix or replace the invalid file(s) listed below:
        </div>
        <ul style="margin-top:10px; margin-bottom:4px; font-size:0.83rem; color:#FECACA; padding-left:20px;">
            {err_list_html}
        </ul>
    </div>
    """, unsafe_allow_html=True)
    st.button("🚫 Generation Blocked: Fix Format Errors Above", disabled=True, width="stretch")
else:
    st.markdown("""
    <div class="validation-success-banner">
        <span style="font-size:1.35rem;">✅</span>
        <div>
            <strong style="color:#A7F3D0;">File Formats Validated:</strong> All 3 source files match the expected Nokia CMG templates (CIQ, DATA IP-Plan, IMS IP-Plan). Ready to generate.
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.expander(f"🌐 Active Circle Profile & Target Routing Parameters ({active_circle_label})", expanded=False):
        p_c1, p_c2 = st.columns(2)
        with p_c1:
            st.markdown(f"**Applied Circle Profile:** `{active_circle_label}`")
            st.markdown(f"**S1-U Test eNodeB IP:** `{active_circle_profile.get('s1u_test_enb')}`")
            pool_str = str(active_circle_profile.get('sgi_pool')) if active_circle_profile.get('sgi_pool') is not None else "Auto-detected from NAT Inside VPRN"
            st.markdown(f"**SGi NAT Public Pool:** `{pool_str}`")
        with p_c2:
            st.markdown(f"**SGi Destination Targets (IPv4):** `{', '.join(active_circle_profile.get('sgi_dst_v4', []))}`")
            st.markdown(f"**NTP Servers:** `{', '.join(active_circle_profile.get('ntp', []))}`")
            st.caption("💡 The engine automatically discovers and prioritizes circle destinations in CIQ sheets (e.g. S1U-eNb, Gi DNS) if populated.")

    generate_btn = st.button("🚀 Generate Reachability Tracker Workbook", type="primary", width="stretch")

    if generate_btn:
        if not all_valid:
            st.error("❌ Generation blocked: Source files failed format validation.")
            st.stop()
        with st.spinner("Extracting parameters and constructing Reachability Tracker workbook..."):
            try:
                if use_manual_overrides:
                    custom_dcgw = {
                        **tracker_engine.DEFAULT_DCGW_VPRN,
                        "ntp": dcgw_ntp, "sx": dcgw_sx, "s11": dcgw_s11, "s1u": dcgw_s1u, "sgi": dcgw_sgi,
                        "gx": dcgw_gx, "gy": dcgw_gy, "bp": dcgw_bp, "gm": dcgw_gm, "s5s8_u": dcgw_s5s8,
                    }
                    custom_static = {
                        **active_circle_profile,
                        "ntp": [ntp_ip1, ntp_ip2],
                        "s1u_test_enb": s1u_test_enb,
                    }
                    p_pool = sgi_pool_val
                    p_v4_line = sgi_v4_line_val
                    p_gm_line = gm_v6_line_val
                elif circle_selection != "Auto-Detect from Site / Files":
                    custom_dcgw = None
                    custom_static = active_circle_profile
                    p_pool = active_circle_profile.get("sgi_pool")
                    p_v4_line = active_circle_profile.get("sgi_v4_line", 0)
                    p_gm_line = active_circle_profile.get("gm_v6_line", 0)
                else:
                    custom_dcgw = None
                    custom_static = None
                    p_pool = None
                    p_v4_line = None
                    p_gm_line = None

                # Reset file pointers if BytesIO
                for f_obj in [ciq_file, data_file, ims_file]:
                    if hasattr(f_obj, "seek"):
                        f_obj.seek(0)

                out_buf, stats, warnings_list, summary_data, detail_data = tracker_engine.generate_reachability_tracker(
                    ciq_source=ciq_file,
                    data_source=data_file,
                    ims_source=ims_file,
                    site_name=final_site,
                    sgi_pool=p_pool,
                    sgi_v4_line=p_v4_line,
                    gm_v6_line=p_gm_line,
                    dcgw_vprn=custom_dcgw,
                    static_cfg=custom_static,
                    filenames=fn_map
                )

                st.session_state["out_buf"] = out_buf
                st.session_state["stats"] = stats
                st.session_state["warnings"] = warnings_list
                st.session_state["summary_data"] = summary_data
                st.session_state["detail_data"] = detail_data
                st.session_state["final_site"] = final_site

                st.success(f"🎉 Successfully built Reachability Tracker for **{final_site}**!")

            except Exception as e:
                st.error(f"❌ Error during generation: {str(e)}")
                st.exception(e)


# -----------------------------------------------------------------------------
# DISPLAY RESULTS & DOWNLOAD
# -----------------------------------------------------------------------------
if "out_buf" in st.session_state:
    stats = st.session_state["stats"]
    warnings_list = st.session_state["warnings"]
    summary_data = st.session_state["summary_data"]
    detail_data = st.session_state["detail_data"]
    final_site = st.session_state["final_site"]
    out_buf = st.session_state["out_buf"]

    st.markdown("---")
    st.subheader(f"📊 Reachability Tracker Output: {final_site}")

    # Metric Row
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-val">{stats['total_detail_rows']}</div>
            <div class="metric-lbl">Total Detail Rows</div>
        </div>
        """, unsafe_allow_html=True)
    with m2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-val">{stats['total_blocks']}</div>
            <div class="metric-lbl">Configured Blocks</div>
        </div>
        """, unsafe_allow_html=True)
    with m3:
        nodes_str = ", ".join(stats['nodes'])
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-val" style="font-size:1.3rem; padding-top:6px;">{nodes_str}</div>
            <div class="metric-lbl">Covered Nodes</div>
        </div>
        """, unsafe_allow_html=True)
    with m4:
        kb_size = round(stats['workbook_size_bytes'] / 1024, 1)
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-val">{kb_size} KB</div>
            <div class="metric-lbl">Excel Workbook Size</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Download Button
    file_download_name = f"Reachability_tracker_of_{final_site.replace(' ', '_')}.xlsx"
    st.download_button(
        label=f"📥 Download {file_download_name}",
        data=out_buf.getvalue(),
        file_name=file_download_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch"
    )

    # Warnings if any
    if warnings_list:
        with st.expander(f"⚠️ Validation Notes ({len(warnings_list)})", expanded=False):
            for w in warnings_list:
                st.warning(w)

    # Interactive Previews
    st.markdown("### 🔍 Workbook Content Preview & Address Verification")
    tab_sum, tab_det, tab_addr, tab_sources = st.tabs([
        "📋 Summary Rollup Tables",
        "📄 All Interface Rows Preview",
        "🎯 CMG SRC & DST Address Audit (Cols E & G)",
        "ℹ️ Sources & Assumptions"
    ])

    with tab_sum:
        st.markdown("**Status from CMG / DCGW Summary Matrix**")
        df_sum = pd.DataFrame(summary_data)
        st.dataframe(df_sum, width="stretch", hide_index=True)
        st.caption("Note: Dynamic Excel formulas `=COUNTIF(...)` and `=SUM(...)` will compute automatically when opened in Microsoft Excel or LibreOffice.")

    with tab_det:
        st.markdown(f"**Generated Interface Detail Breakdown (Rows {stats['start_row']} - {stats['end_row']})**")
        df_det = pd.DataFrame(detail_data)
        df_det = df_det.astype(str).replace({"None": "", "nan": ""})

        # Filters
        node_filter = st.multiselect("Filter by Node:", options=stats['nodes'], default=stats['nodes'])
        if node_filter:
            filtered_df = df_det[df_det["Node"].isin(node_filter)]
        else:
            filtered_df = df_det

        st.dataframe(filtered_df, width="stretch", hide_index=True)

    with tab_addr:
        st.markdown("#### 🎯 Strict Verification of CMG SRC Address (Col E) & DST Address (Col G)")
        st.markdown(
            "Every automated `ping` and `traceroute` script strictly relies on 100% accurate source and destination IP addresses. "
            "Use this view to verify the extracted, cleaned, and formatted IP pairs across all nodes and interfaces."
        )

        df_det_full = pd.DataFrame(detail_data)
        df_det_full = df_det_full.astype(str).replace({"None": "", "nan": ""})

        # Address summary metrics
        valid_src = df_det_full[df_det_full["CMG SRC Address (Col E)"] != ""]["CMG SRC Address (Col E)"].nunique()
        valid_dst = df_det_full[df_det_full["DST Address (Col G)"] != ""]["DST Address (Col G)"].nunique()
        active_pairs = len(df_det_full[df_det_full["Generated Ping Command"] != ""])

        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            st.metric("Unique CMG Source IPs (Col E)", valid_src)
        with col_m2:
            st.metric("Unique Target IPs (Col G)", valid_dst)
        with col_m3:
            st.metric("Active Ping/Trace Pairs", active_pairs)

        # Filters for address review
        c_search, c_node = st.columns([2, 1])
        with c_search:
            search_query = st.text_input("Search IP address or destination:", placeholder="e.g. 10.216, 203.88, 2402:, PCRF, NTP...", key="ip_search")
        with c_node:
            node_sel = st.multiselect("Filter by Node:", options=stats['nodes'], default=stats['nodes'], key="addr_node_filter")

        addr_cols = ["Node", "Interface", "Destination", "CMG SRC Address (Col E)", "DST Address (Col G)", "Generated Ping Command"]
        df_addr = df_det_full[addr_cols]

        if node_sel:
            df_addr = df_addr[df_addr["Node"].isin(node_sel)]
        if search_query.strip():
            sq = search_query.strip().lower()
            df_addr = df_addr[
                df_addr["CMG SRC Address (Col E)"].str.lower().str.contains(sq) |
                df_addr["DST Address (Col G)"].str.lower().str.contains(sq) |
                df_addr["Destination"].str.lower().str.contains(sq) |
                df_addr["Interface"].str.lower().str.contains(sq)
            ]

        st.dataframe(df_addr, width="stretch", hide_index=True)

    with tab_sources:
        st.markdown("""
        **Automated Mapping Architecture:**
        - **CMG system / loopbacks**: Extracted from `'External Interfaces'` sheet in DATA and IMS IP-Plans.
        - **VPRN Service IDs**: Extracted from `'CP-SIGNAL'`, `'CP-BILLING'`, `'CP-LI'`, `'CP-CORE'`, and `'UP-CORE'`.
        - **S11 Peer & MME**: Extracted from CIQ `'S11'` sheet.
        - **Gx / Gy / S6b Peers**: Extracted from CIQ `'Gx'`, `'Gy'`, and `'S6b'` sheets.
        - **LI Interfaces (X1, X2, X3)**: Extracted from CIQ `'X1-X2-X3'` sheet.
        - **S2b-C & S2b-U**: Extracted from CIQ `'S2b'` sheet.
        - **S1-U Sources**: Extracted from CIQ `'S1U-eNb'` sheet per DATA / IMS group.
        - **S5-S8 SGW/PGW Sources**: Extracted from CIQ `'Gn-Gp-S5-S8'` sheet.
        - **Gm Peers & Media**: Extracted from CIQ `'P-CSCF_SBC'` sheet.
        - **SGi & Gm Sources**: Extracted from CIQ `'UE IP pool'` sheet.
        - **Automated Scripts**: Formulaic generation of `ping` and `traceroute` commands for each peer in Columns P, Q, S.
        """)
