# ============================================================
# CS WORKLOAD DASHBOARD
# ============================================================

from __future__ import annotations

import re
import html
import hashlib
import logging
import tempfile
import textwrap
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ============================================================
# CONFIG
# ============================================================

st.set_page_config(
    page_title="CS OPERATIONS PERFORMANCE DASHBOARD",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)
# Khởi tạo trạng thái đăng nhập
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

DASHBOARD_PASSWORD = "1234"  # Mật khẩu của bạn


# ============================================================
# 1. ĐỊNH NGHĨA POPUP NHẬP MẬT KHẨU (ST.DIALOG)
# ============================================================
@st.dialog("🔒 DASHBOARD ACCESS AUTHENTICATION")
def password_popup():
    # Ép chiều rộng modal rộng ra để chữ không bị rớt dòng
    st.markdown(
        """
        <style>
        div[data-testid="stModal"] > div {
            width: 1000px !important;
            max-width: 1000px !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )
    st.write("Please enter the password to access the dashboard.")
    
    with st.form("popup_login_form"):
        entered_password = st.text_input("Password", type="password", placeholder="Input password...")
        submit_btn = st.form_submit_button("Confirm", use_container_width=True)
        
        if submit_btn:
            if entered_password == DASHBOARD_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()  # Đóng popup và load trực tiếp vào dashboard
            else:
                st.error("Password incorrect! Please try again.")


# ============================================================
# 2. PHÂN LUỒNG HIỂN THỊ
# ============================================================
if st.session_state.authenticated:
    # 🌟 KHI ĐÃ ĐĂNG NHẬP: Bỏ qua (pass) để code dashboard gốc ở dưới chạy bình thường
    pass

else:
    # 🌟 KHI CHƯA ĐĂNG NHẬP: CHỈ HIỂN THỊ TRANG COVER (TRANG BÌA)
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 4, 1])
    with col2:
        st.markdown(
            """
            <div style="padding: 30px; border: 1px solid #D8E1EA; border-radius: 14px; background: #FFFFFF; box-shadow: 0 4px 15px rgba(0,0,0,0.05);">
                <h1 style="color: #0A192F; font-size: 26px;">CS OPERATIONS PERFORMANCE DASHBOARD</h1>
                <p style="color: #666; font-size: 14px;">Capacity • Workload • Utilization • Performance</p>
                <hr style="border: 0; height: 3px; background: #FF7A00; width: 100%;">
                <br>
                <ul>
                    <li><b>CAPACITY:</b> HC Capacity, Requirement & Gap</li>
                    <li><b>WORKLOAD:</b> Customer, Volume, Segment & Activity</li>
                    <li><b>UTILIZATION:</b> Office Workload, CS Allocation</li>
                    <li><b>PERFORMANCE:</b> CS Resolution, YVF Booking Adoption</li>
                </ul>
            </div>
            """,
            unsafe_allow_html=True
        )
        
        st.write("")
        # Nút bấm kích hoạt Popup nhập mật khẩu
        if st.button("VIEW DASHBOARD ➔", type="primary", use_container_width=True):
            password_popup()
            
    # Dừng app ở đây khi chưa đăng nhập để không bị lộ nội dung dashboard bên dưới
    st.stop()

APP_TITLE = "CS OPERATIONS PERFORMANCE DASHBOARD"
APP_SUBTITLE = "Capacity • Workload • Utilization • Performance"
DEFAULT_FILE = "(Not for Office Input) MASTER DATA SOURCE.xlsm"
APP_DIR = Path(__file__).resolve().parent
DEFAULT_FILE_PATH = APP_DIR / DEFAULT_FILE
CAPACITY_HOURS_PER_FTE = 8 * 0.95 * 22  # 167.2 hours/FTE/month
STANDARD_OFFICES = ["HAN", "HAD", "HLC", "HCM"]
SERVICE_ORDER = ["AE", "AI", "OE", "OI", "CC", "TR", "WH"]
SERVICE_LABELS = {
    "AE": "Air Export",
    "AI": "Air Import",
    "OE": "Ocean Export",
    "OI": "Ocean Import",
    "CC": "Customs Clearance",
    "TR": "Trucking",
    "WH": "Warehouse",
}

# ============================================================
# YUSEN 3C-INSPIRED CORPORATE COLOR SYSTEM
# Visual reference: Yusen Logistics corporate web presence.
# These HEX values are used as a consistent dashboard design system and
# are NOT asserted here as official corporate-brand specifications.
# ============================================================
YUSEN_THEME = {
    "primary": "#06183F",          # Yusen Navy - primary
    "primary_dark": "#041532",     # deeper navy for sidebar gradient
    "secondary": "#0DBAEE",        # Yusen Cyan - supporting/data series
    "secondary_mid": "#3F5B81",    # Mid Blue
    "secondary_light": "#D5EAF8",  # Light Blue
    "secondary_pale": "#EEF7FC",
    "accent": "#E6761B",           # Yusen Orange - accent/attention
    "accent_pale": "#FFF2E8",
    "background": "#F6F8FB",
    "surface": "#FFFFFF",
    "text_primary": "#06183F",
    "text_secondary": "#5B6575",
    "border": "#D5E1EA",
    "grid": "#E8EEF3",
    "hover": "#F3F7FA",
}

COLORS = {
    "navy": YUSEN_THEME["primary"],
    "blue": YUSEN_THEME["secondary"],
    "light_blue": YUSEN_THEME["secondary_pale"],
    "red": "#D92D20",
    "green": "#95C947",
    # Legacy alias retained for any older status branch still referencing
    # COLORS["purple"]. It intentionally matches the approved High Load color.
    "purple": "#F2B84B",
    "high_load": "#F2B84B",
    "amber": YUSEN_THEME["accent"],
    "gray": "#98A2B3",
    "gray_dark": YUSEN_THEME["text_secondary"],
    "grid": YUSEN_THEME["grid"],
    "bg": YUSEN_THEME["background"],
    "white": YUSEN_THEME["surface"],
    "text": YUSEN_THEME["text_primary"],
    "muted": YUSEN_THEME["text_secondary"],
    "border": YUSEN_THEME["border"],
}

# Business-meaning color map.
# UI only: Same Business Meaning = Same Color Everywhere.
BUSINESS_COLORS = {
    "actual": COLORS["blue"],
    "approved": COLORS["navy"],
    "required": COLORS["amber"],
    "positive": COLORS["green"],
    "negative": COLORS["red"],
    "critical": COLORS["red"],
    "supporting": YUSEN_THEME["secondary_light"],
}


# ============================================================
# SHARED EXECUTIVE UI CONSTANTS
# UI only — no business logic / calculation changes
# ============================================================

UI = {
    "font_family": "Inter, 'Segoe UI', Arial, sans-serif",
    "title_size": 30,
    "section_title_size": 19,
    "chart_title_size": 17,
    "kpi_value_size": 32,
    "kpi_label_size": 13,
    "body_size": 12,
    "axis_size": 11,
    "note_size": 11,
    "radius": 12,
    "card_padding": 16,
    "section_gap": 18,
    "chart_height": 380,
    "chart_height_tall": 520,
}

# Shared height for Shipment Volume chart/detail pairs
SHIPMENT_PAIR_HEIGHT = 500

logger = logging.getLogger(__name__)


class WorkbookLoadError(RuntimeError):
    """Raised when the selected workbook or one of its sheets cannot be read."""

CORPORATE_PALETTE = [
    YUSEN_THEME["secondary"],
    YUSEN_THEME["primary"],
    "#2E73AA",
    "#5D91BC",
    "#8EB7D8",
    "#B7D0E3",
    YUSEN_THEME["accent"],
]

SHEET_NAMES = {
    "hc": " 1.  Office Cap. & Workload",
    "resolution": "9. CS Resolutions Rate",
    "workload": "4. Workload by Activity",
    "yvf": "10. YVF",
    "shipment": "3. Active Cus - Vol.",
    "customer_ns": "11. Vol. by Customer",

    # The new master workbook no longer keeps separate office-specific customer sheets.
    # Keep these aliases intentionally unmatched so prepare_customer() falls back to
    # the combined Customer Volume-N&S sheet without changing downstream logic.
    "customer_had": "__NOT_USED_CUSTOMER_HAD__",
    "customer_han": "__NOT_USED_CUSTOMER_HAN__",
    "customer_hlc": "__NOT_USED_CUSTOMER_HLC__",
    "customer_hcm": "__NOT_USED_CUSTOMER_HCM__",

    "fte": " 2. FTE Workload",
    "core": "5. C Vol.",
    "ancillary": "6. A Vol.",
    "supporting": "7. S Vol.",
    "exception": "8. E Vol.",

    # The new master workbook does not contain the former "Ghi chú" sheet.
    # Existing downstream fallback behavior is retained.
    "notes": "__NOT_USED_NOTES__",
}

# ============================================================
# STYLE
# ============================================================

st.markdown(
    f"""
    <style>
    .stApp {{
        background: {COLORS['bg']};
        color: {COLORS['text']};
    }}

/* Reduce top whitespace and move dashboard content upward */
[data-testid="stMainBlockContainer"],
.block-container {{
    padding-top: 0rem !important;
    margin-top: -3.5rem !important;
}}

.main-header {{
    margin-top: 0 !important;
}}
    section[data-testid="stSidebar"] {{
        background: {COLORS['navy']};
    }}
    /* Sidebar: high-contrast labels, captions and controls */
    section[data-testid="stSidebar"] {{
        background: {COLORS['navy']};
    }}
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] small,
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {{
        color: #FFFFFF !important;
        opacity: 1 !important;
    }}
    section[data-testid="stSidebar"] div[data-baseweb="select"] > div {{
        background: #FFFFFF !important;
        border-color: #D9E2EC !important;
    }}
    section[data-testid="stSidebar"] div[data-baseweb="select"] span,
    section[data-testid="stSidebar"] div[data-baseweb="select"] input,
    section[data-testid="stSidebar"] div[data-baseweb="select"] svg {{
        color: {COLORS['navy']} !important;
        fill: {COLORS['navy']} !important;
        opacity: 1 !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section {{
        background: #FFFFFF !important;
        border-color: #D9E2EC !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section * {{
        color: {COLORS['navy']} !important;
        opacity: 1 !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] button {{
        background: #FFFFFF !important;
        color: {COLORS['navy']} !important;
        border: 1px solid #B8C7D6 !important;
        opacity: 1 !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] button * {{
        color: {COLORS['navy']} !important;
        opacity: 1 !important;
    }}
    .main-header {{
        background: {COLORS['white']};
        border: 1px solid {COLORS['border']};
        border-left: 7px solid {COLORS['blue']};
        border-radius: 14px;
        padding: 10px 16px;
        margin-bottom: 8px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    }}
    .main-title {{
        color: {COLORS['navy']};
        font-size: 30px;
        font-weight: 800;
        margin: 0;
        letter-spacing: -0.02em;
    }}
    .subtitle {{
        color: {COLORS['muted']};
        font-size: 14px;
        margin-top: 4px;
    }}
    .section-title {{
        color: {COLORS['navy']};
        font-size: 18px;
        font-weight: 800;
        margin: 22px 0 8px 0;
        border-left: 5px solid {COLORS['amber']};
        padding-left: 10px;
    }}
    .kpi-card {{
        background: {COLORS['white']};
        border: 1px solid {COLORS['border']};
        border-radius: 14px;
        padding: 16px 14px;
        min-height: 110px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.035);
    }}

    .hc-kpi-card {{
        background: #FFFFFF;
        border: 1px solid #D9E2EC;
        border-radius: 14px;
        padding: 16px 16px 14px 16px;
        min-height: 190px;
        height: 190px;
        box-sizing: border-box;
        box-shadow: 0 2px 10px rgba(0,0,0,0.035);
        display: flex;
        flex-direction: column;
        justify-content: flex-start;
        text-align: center;
    }}

    .hc-kpi-card .kpi-label {{
        text-align: center;
        width: 100%;
    }}

    .hc-kpi-total {{
        color: #003B70;
        font-size: 32px;
        line-height: 1.05;
        font-weight: 850;
        margin-top: 12px;
        margin-bottom: 8px;
        text-align: center;
        width: 100%;
    }}

    .hc-detail-row {{
        display: grid;
        grid-template-columns: 1fr 1px 1fr;
        align-items: stretch;
        gap: 12px;
        margin-top: auto;
        padding-top: 12px;
        border-top: 1px solid #E5E7EB;
    }}

    .hc-detail-divider {{
        background: #E5E7EB;
        width: 1px;
    }}

    .hc-detail-item {{
        text-align: center;
    }}

    .hc-detail-label {{
        color: #64748B;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.04em;
    }}

    .hc-detail-value {{
        color: #003B70;
        font-size: 20px;
        line-height: 1.2;
        font-weight: 800;
        margin-top: 3px;
    }}

    .hc-variance-card {{
        justify-content: flex-start;
    }}

    .hc-variance-formula {{
        color: #64748B;
        font-size: 12px;
        font-weight: 600;
        margin-top: 12px;
        margin-bottom: 10px;
        text-align: center;
        width: 100%;
    }}

    .hc-variance-status {{
        display: block;
        width: 100%;
        box-sizing: border-box;
        text-align: center;
        margin-top: 0 !important;
    }}


    /* Section 2 - Shipment KPI cards */
    .shipment-kpi-card {{
        background: #FFFFFF;
        border: 1px solid #D9E2EC;
        border-radius: 14px;
        padding: 18px 18px;
        min-height: 148px;
        height: 148px;
        box-sizing: border-box;
        box-shadow: 0 2px 10px rgba(0,0,0,0.035);
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        text-align: center;
    }}

    .shipment-kpi-label {{
        color: #64748B;
        font-size: 12px;
        font-weight: 700;
        text-transform: none;
        letter-spacing: 0.04em;
        margin-bottom: 14px;
    }}

    .shipment-kpi-value {{
        color: #003B70;
        font-size: 36px;
        line-height: 1.05;
        font-weight: 850;
    }}

    .shipment-kpi-note {{
        color: #64748B;
        font-size: 11px;
        margin-top: 10px;
    }}


    /* Section 3 - Workload by PIC */
    .pic-kpi-card {{
        background: #FFFFFF;
        border: 1px solid #D9E2EC;
        border-radius: 14px;
        padding: 14px 14px;
        min-height: 142px;
        height: 142px;
        box-sizing: border-box;
        box-shadow: 0 2px 10px rgba(0,0,0,0.035);
        display: flex;
        flex-direction: column;
        justify-content: center;
        text-align: center;
    }}

    .pic-kpi-label {{
        color: #64748B;
        font-size: 11px;
        font-weight: 700;
        text-transform: none;
        letter-spacing: 0.035em;
        line-height: 1.25;
        min-height: 30px;
    }}

    .pic-kpi-unit {{
        color: #7A8699;
        font-size: 10px;
        line-height: 1.2;
        font-weight: 600;
        margin-top: 4px;
        margin-bottom: 8px;
        text-align: center;
        letter-spacing: 0.02em;
    }}

    .pic-kpi-value {{
        color: #003B70;
        font-size: 27px;
        line-height: 1.05;
        font-weight: 850;
        margin-top: 8px;
    }}

    .pic-kpi-note {{
        color: #64748B;
        font-size: 10.5px;
        margin-top: 7px;
        line-height: 1.25;
    }}

    .pic-status-card {{
        background: #FFFFFF;
        border: 1px solid #D9E2EC;
        border-radius: 14px;
        padding: 14px 16px;
        min-height: 110px;
        height: 110px;
        box-sizing: border-box;
        box-shadow: 0 2px 10px rgba(0,0,0,0.035);
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
    }}

    .pic-status-left {{
        min-width: 180px;
    }}

    .pic-status-title {{
        color: #64748B;
        font-size: 11px;
        font-weight: 700;
        text-transform: none;
        letter-spacing: 0.035em;
    }}

    .pic-status-value {{
        color: #003B70;
        font-size: 30px;
        font-weight: 850;
        margin-top: 4px;
    }}

    .pic-progress-track {{
        flex: 1;
        height: 10px;
        background: #EAF3F8;
        border-radius: 999px;
        overflow: hidden;
    }}

    .pic-progress-fill {{
        height: 100%;
        border-radius: 999px;
    }}

    .workload-status-panel {{
        background: #FFFFFF;
        border: 1px solid #D9E2EC;
        border-radius: 14px;
        padding: 14px 16px;
        min-height: 110px;
        height: 110px;
        box-sizing: border-box;
        box-shadow: 0 2px 10px rgba(0,0,0,0.035);
        display: flex;
        flex-direction: column;
        justify-content: center;
        text-align: center;
    }}

    .workload-status-text {{
        font-size:30px ;
        font-weight: 850;
        margin-top: 8px;
    }}
    .kpi-label {{
        color: {COLORS['muted']};
        font-size: 12px;
        font-weight: 700;
        text-transform: none;
        letter-spacing: 0.04em;
        margin-bottom: 6px;
    }}
    .kpi-value {{
        color: {COLORS['navy']};
        font-size: 26px;
        line-height: 1.05;
        font-weight: 850;
        margin-bottom: 4px;
    }}
    .kpi-note {{
        color: {COLORS['muted']};
        font-size: 11px;
    }}
    .status-badge {{
        display:inline-block;
        padding: 4px 10px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 800;
        margin-top: 8px;
    }}
    .chart-box {{
        background: {COLORS['white']};
        border: 1px solid {COLORS['border']};
        border-radius: 14px;
        padding: 12px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.03);
    }}
    .warning-box {{
        background: #FFF7ED;
        color: #92400E;
        border: 1px solid #FED7AA;
        border-radius: 12px;
        padding: 10px 12px;
        margin-bottom: 10px;
        font-size: 13px;
    }}
    .stTabs [data-baseweb="tab-list"] {{ gap: 8px; }}
    .stTabs [data-baseweb="tab"] {{
        background: {COLORS['white']};
        border-radius: 10px 10px 0 0;
        border: 1px solid {COLORS['border']};
        padding: 8px 16px;
    }}
    
    /* Section 3 compact executive layout */
    .compact-workload-kpi {{
        min-height: 138px !important;
        height: 138px !important;
        padding: 16px 18px 14px 18px !important;
        display: flex !important;
        flex-direction: column !important;
        justify-content: center !important;
    }}
    .compact-workload-kpi .kpi-label {{
        margin-bottom: 8px !important;
        line-height: 1.15 !important;
    }}
    .compact-workload-kpi .kpi-value {{
        margin: 2px 0 6px 0 !important;
        line-height: 1.05 !important;
    }}
    .compact-workload-kpi .kpi-note {{
        margin-top: 4px !important;
        line-height: 1.15 !important;
    }}
    .workload-util-card, .workload-status-card {{
        min-height: 94px !important;
        height: 94px !important;
        padding: 13px 16px !important;
    }}
    .workload-util-card {{
        display: grid !important;
        grid-template-columns: 175px minmax(0, 1fr) !important;
        align-items: center !important;
        column-gap: 18px !important;
    }}
    .workload-status-card {{
        display: flex !important;
        flex-direction: column !important;
        justify-content: center !important;
    }}
    @media (max-width: 1100px) {{
        .workload-util-card {{
            grid-template-columns: 155px minmax(0, 1fr) !important;
            column-gap: 12px !important;
        }}
    }}

</style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# EXECUTIVE / CORPORATE UI OVERRIDES
# Shared styling layer only — business logic remains unchanged
# ============================================================

st.markdown(
    f"""
    <style>
    :root {{
        --font-main: {UI['font_family']};
        --navy: {COLORS['navy']};
        --blue: {COLORS['blue']};
        --orange: {COLORS['amber']};
        --green: {COLORS['green']};
        --red: {COLORS['red']};
        --text: {COLORS['text']};
        --muted: #667085;
        --border: #D8E1EA;
        --surface: #FFFFFF;
        --background: #F6F8FA;
        --radius: {UI['radius']}px;
    }}

    html, body, [class*="css"], .stApp,
    button, input, textarea, select {{
        font-family: var(--font-main) !important;
    }}

    .stApp {{
        background: var(--background);
        color: var(--text);
    }}

    .block-container {{
        max-width: 1680px;
        padding-top: 1.35rem !important;
        padding-left: 1.45rem !important;
        padding-right: 1.45rem !important;
        padding-bottom: 2rem !important;
    }}

    /* Header */
    .main-header {{
        border-radius: var(--radius);
        padding: 16px 20px;
        margin: 0 0 14px 0 !important;
        border: 1px solid var(--border);
        border-left: 5px solid var(--blue);
        box-shadow: 0 1px 4px rgba(16, 24, 40, 0.05);
    }}

    .main-title {{
        font-size: {UI['title_size']}px !important;
        line-height: 1.15 !important;
        font-weight: 700 !important;
        letter-spacing: -0.015em !important;
        color: var(--navy) !important;
    }}

    .subtitle {{
        font-size: {UI['body_size']}px !important;
        line-height: 1.45 !important;
        color: var(--muted) !important;
        margin-top: 3px !important;
    }}

    /* Section titles */
    .section-title {{
        font-size: {UI['section_title_size']}px !important;
        line-height: 1.25 !important;
        font-weight: 700 !important;
        color: var(--navy) !important;
        margin: 20px 0 10px 0 !important;
        padding: 0 0 0 10px !important;
        border-left: 4px solid var(--orange) !important;
    }}

    /* Shared card language */
    .kpi-card,
    .hc-kpi-card,
    .shipment-kpi-card,
    .pic-kpi-card,
    .pic-status-card,
    .workload-status-panel,
    .chart-box {{
        background: var(--surface) !important;
        border: 1px solid var(--border) !important;
        border-radius: var(--radius) !important;
        box-shadow: 0 1px 4px rgba(16, 24, 40, 0.045) !important;
        box-sizing: border-box !important;
    }}

    .kpi-card,
    .hc-kpi-card,
    .shipment-kpi-card,
    .pic-kpi-card {{
        padding: {UI['card_padding']}px !important;
    }}

    /* KPI labels */
    .kpi-label,
    .shipment-kpi-label,
    .pic-kpi-label {{
        color: #5F6B7A !important;
        font-size: {UI['kpi_label_size']}px !important;
        line-height: 1.25 !important;
        font-weight: 600 !important;
        letter-spacing: 0.025em !important;
        text-transform: none !important;
        margin-bottom: 7px !important;
    }}

    /* KPI values */
    .kpi-value,
    .hc-kpi-total,
    .shipment-kpi-value,
    .pic-kpi-value,
    .pic-status-value {{
        color: var(--navy) !important;
        font-size: {UI['kpi_value_size']}px !important;
        line-height: 1.05 !important;
        font-weight: 700 !important;
        letter-spacing: -0.02em !important;
    }}

    /* Notes / formulas / sources */
    .kpi-note,
    .shipment-kpi-note,
    .pic-kpi-note,
    .hc-variance-formula,
    [data-testid="stCaptionContainer"],
    [data-testid="stCaptionContainer"] p {{
        color: var(--muted) !important;
        font-size: {UI['note_size']}px !important;
        line-height: 1.4 !important;
        font-weight: 400 !important;
    }}

    /* HC cards — equal structure */
    .hc-kpi-card {{
        height: 184px !important;
        min-height: 184px !important;
    }}

    .hc-detail-row {{
        margin-top: auto !important;
        padding-top: 11px !important;
        gap: 10px !important;
        border-top: 1px solid #E7ECF1 !important;
    }}

    .hc-detail-label {{
        color: var(--muted) !important;
        font-size: 11px !important;
        font-weight: 600 !important;
    }}

    .hc-detail-value {{
        color: var(--navy) !important;
        font-size: 19px !important;
        font-weight: 700 !important;
    }}

    /* Office Capacity Snapshot — semantic color hierarchy */
    .hc-total-approved {{
        color: var(--navy) !important;
    }}
    .hc-total-actual {{
        color: var(--blue) !important;
    }}
    .hc-total-required {{
        color: var(--orange) !important;
    }}

    /* Shipment KPI cards */
    .shipment-kpi-card {{
        height: 132px !important;
        min-height: 132px !important;
    }}

    /* PIC KPI cards */
    .pic-kpi-card {{
        height: 140px !important;
        min-height: 140px !important;
    }}

    .pic-status-card,
    .workload-status-panel {{
        height: 104px !important;
        min-height: 104px !important;
        padding: 14px 16px !important;
    }}

    .pic-status-title {{
        font-size: 12px !important;
        font-weight: 600 !important;
        color: var(--muted) !important;
        letter-spacing: 0.025em !important;
    }}

    /* Chart wrappers */
    .chart-box {{
        padding: 12px 14px 10px 14px !important;
        overflow: visible !important;
        min-width: 0 !important;
    }}

    .chart-box [data-testid="stPlotlyChart"] {{
        margin: 0 !important;
    }}

    .chart-box .js-plotly-plot,
    .chart-box .plot-container,
    .chart-box .svg-container {{
        width: 100% !important;
    }}

    /* Status */
    .status-badge {{
        font-size: 11px !important;
        font-weight: 700 !important;
        padding: 4px 9px !important;
    }}

    /* Streamlit dataframe */
    [data-testid="stDataFrame"] {{
        border: 1px solid var(--border);
        border-radius: var(--radius);
        overflow: hidden;
    }}

    /* Vertical spacing between Streamlit blocks */
    div[data-testid="stVerticalBlock"] > div {{
        gap: 0.35rem;
    }}

    /* Sidebar remains high contrast */
    section[data-testid="stSidebar"] {{
        background: var(--navy) !important;
    }}

    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] small,
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {{
        color: #FFFFFF !important;
        opacity: 1 !important;
    }}

    section[data-testid="stSidebar"] div[data-baseweb="select"] > div,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section {{
        background: #FFFFFF !important;
        border: 1px solid #C9D5E1 !important;
        border-radius: 8px !important;
    }}

    /* Laptop responsive behavior */
    @media (max-width: 1200px) {{
        .block-container {{
            padding-left: 1rem !important;
            padding-right: 1rem !important;
        }}

        .main-title {{
            font-size: 28px !important;
        }}

        .kpi-value,
        .hc-kpi-total,
        .shipment-kpi-value,
        .pic-kpi-value,
        .pic-status-value {{
            font-size: 28px !important;
        }}

        .kpi-label,
        .shipment-kpi-label,
        .pic-kpi-label {{
            font-size: 12px !important;
        }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# FINAL UI/UX QUALITY OVERRIDES
# UI only — no business logic / calculation changes
# ============================================================

st.markdown(
    f"""
    <style>
    /* Management-first density and consistent vertical rhythm */
    .block-container {{
        max-width: 1680px !important;
        padding-top: 1.10rem !important;
        padding-bottom: 1.75rem !important;
    }}

    .section-title {{
        margin-top: 22px !important;
        margin-bottom: 12px !important;
    }}

    /* Standard Streamlit charts as real cards.
       Avoid raw HTML wrappers around Streamlit widgets. */
    [data-testid="stPlotlyChart"] {{
        background: #FFFFFF;
        border: 1px solid {COLORS['border']};
        border-radius: {UI['radius']}px;
        box-shadow: 0 1px 4px rgba(16, 24, 40, 0.045);
        padding: 8px 10px 4px 10px;
        box-sizing: border-box;
    }}

    [data-testid="stDataFrame"] {{
        background: #FFFFFF;
        border: 1px solid {COLORS['border']} !important;
        border-radius: {UI['radius']}px !important;
        box-shadow: 0 1px 4px rgba(16, 24, 40, 0.035);
    }}

    /* Equal KPI-card language */
    .kpi-card {{
        min-height: 124px !important;
        height: 124px !important;
        display: flex !important;
        flex-direction: column !important;
        justify-content: center !important;
        align-items: center !important;
        text-align: center !important;
    }}

    .kpi-card .kpi-label,
    .kpi-card .kpi-value,
    .kpi-card .kpi-note {{
        width: 100% !important;
        text-align: center !important;
    }}

    .kpi-label {{
        min-height: 18px;
    }}

    /* Make notes subordinate to management metrics */
    .kpi-note {{
        margin-top: 5px !important;
        line-height: 1.35 !important;
    }}

    /* Paired chart + detail-table layout */
    .paired-detail-card {{
        background: #FFFFFF;
        border: 1px solid var(--border);
        border-radius: var(--radius);
        box-shadow: 0 1px 4px rgba(16, 24, 40, 0.045);
        padding: 12px 12px 10px 12px;
        box-sizing: border-box;
        width: 100%;
        margin-top: 12px;
        overflow: hidden;
    }}

    .paired-detail-title {{
        color: var(--navy);
        font-size: 15px;
        line-height: 1.25;
        font-weight: 700;
        margin: 1px 0 10px 2px;
    }}

    .paired-detail-table {{
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        table-layout: fixed;
        font-family: var(--font-main);
        font-size: 11px;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        overflow: hidden;
    }}

    .paired-detail-table thead th {{
        background: #F8FAFC;
        color: var(--navy);
        font-weight: 700;
        padding: 7px 7px;
        border-bottom: 1px solid #D8E1EA;
        text-align: left;
        line-height: 1.15;
    }}

    .paired-detail-table tbody td {{
        padding: 6px 7px;
        border-bottom: 1px solid #EDF1F5;
        color: var(--text);
        line-height: 1.15;
        vertical-align: middle;
        background: #FFFFFF;
    }}

    .paired-detail-table tbody tr:nth-child(even) td {{
        background: #FBFCFD;
    }}

    .paired-detail-table tbody tr:last-child td {{
        border-bottom: 0;
    }}

    .paired-detail-table .pair-rank {{
        text-align: center;
        color: #667085;
        font-variant-numeric: tabular-nums;
    }}

    .paired-detail-table .pair-number,
    .paired-detail-table .pair-share {{
        text-align: right;
        white-space: nowrap;
        font-variant-numeric: tabular-nums;
    }}

    .paired-detail-table td.pair-number {{
        color: var(--navy);
        font-weight: 600;
    }}

    .paired-detail-table .pair-name {{
        text-align: left;
        overflow-wrap: anywhere;
    }}

    .customer-name-cell {{
        font-size: 10.5px;
    }}

    .paired-detail-foot {{
        color: var(--muted);
        font-size: 10.5px;
        line-height: 1.3;
        margin: 7px 2px 0 2px;
    }}

    /* Compact tabs and avoid visual competition */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 4px !important;
        border-bottom: 1px solid {COLORS['grid']} !important;
    }}

    .stTabs [data-baseweb="tab"] {{
        border: 0 !important;
        border-radius: 0 !important;
        background: transparent !important;
        padding: 7px 10px !important;
        font-size: 12px !important;
    }}

    /* Responsive laptop refinements */
    @media (max-width: 1366px) {{
        .block-container {{
            padding-left: 0.85rem !important;
            padding-right: 0.85rem !important;
        }}
        .main-title {{
            font-size: 28px !important;
        }}
        .section-title {{
            font-size: 18px !important;
        }}
        .kpi-value,
        .hc-kpi-total,
        .shipment-kpi-value,
        .pic-kpi-value,
        .pic-status-value {{
            font-size: 28px !important;
        }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# YUSEN 3C-INSPIRED FINAL UI LAYER
# UI-only overrides — business logic/data/calculations unchanged.
# ============================================================

st.markdown(
    f"""
    <style>
    :root {{
        --y-primary: {YUSEN_THEME['primary']};
        --y-primary-dark: {YUSEN_THEME['primary_dark']};
        --y-secondary: {YUSEN_THEME['secondary']};
        --y-secondary-mid: {YUSEN_THEME['secondary_mid']};
        --y-secondary-light: {YUSEN_THEME['secondary_light']};
        --y-accent: {YUSEN_THEME['accent']};
        --y-bg: {YUSEN_THEME['background']};
        --y-surface: {YUSEN_THEME['surface']};
        --y-text: {YUSEN_THEME['text_primary']};
        --y-muted: {YUSEN_THEME['text_secondary']};
        --y-border: {YUSEN_THEME['border']};
        --y-grid: {YUSEN_THEME['grid']};
    }}

    html, body, .stApp, [class*="css"],
    button, input, textarea, select {{
        font-family: {UI['font_family']} !important;
    }}

    .stApp {{
        background: var(--y-bg) !important;
        color: var(--y-text) !important;
    }}

    .block-container {{
        max-width: 1680px !important;
        padding-top: 1rem !important;
        padding-left: 1.35rem !important;
        padding-right: 1.35rem !important;
        padding-bottom: 1.8rem !important;
    }}

    /* Main header */
    .main-header {{
        background: var(--y-surface) !important;
        border: 1px solid var(--y-border) !important;
        border-left: 5px solid var(--y-secondary) !important;
        border-radius: 12px !important;
        box-shadow: 0 2px 8px rgba(0,59,112,0.055) !important;
        padding: 15px 19px !important;
        margin-bottom: 12px !important;
    }}
    .main-title {{
        color: var(--y-primary) !important;
        font-size: 30px !important;
        font-weight: 750 !important;
        line-height: 1.12 !important;
        letter-spacing: -0.02em !important;
    }}

    /* Section hierarchy */
    .section-title {{
        color: var(--y-primary) !important;
        font-size: 19px !important;
        font-weight: 700 !important;
        line-height: 1.25 !important;
        border-left: 4px solid var(--y-accent) !important;
        padding-left: 10px !important;
        margin: 22px 0 11px 0 !important;
    }}

    /* Cards */
    .kpi-card,
    .hc-kpi-card,
    .shipment-kpi-card,
    .pic-kpi-card,
    .pic-status-card,
    .workload-status-panel,
    [data-testid="stPlotlyChart"],
    [data-testid="stDataFrame"] {{
        background: var(--y-surface) !important;
        border: 1px solid var(--y-border) !important;
        border-radius: 12px !important;
        box-shadow: 0 2px 7px rgba(0,59,112,0.045) !important;
    }}

    .kpi-label,
    .shipment-kpi-label,
    .pic-kpi-label,
    .pic-status-title {{
        color: var(--y-muted) !important;
        font-weight: 650 !important;
        letter-spacing: 0.025em !important;
    }}

    .kpi-value,
    .hc-kpi-total,
    .shipment-kpi-value,
    .pic-kpi-value,
    .pic-status-value {{
        color: var(--y-primary) !important;
        font-weight: 750 !important;
    }}

    .hc-total-approved {{ color: var(--y-primary) !important; }}
    .hc-total-actual {{ color: var(--y-secondary) !important; }}
    .hc-total-required {{ color: var(--y-accent) !important; }}

    /* Plotly cards */
    [data-testid="stPlotlyChart"] {{
        padding: 7px 9px 3px 9px !important;
        overflow: visible !important;
    }}

    /* Dataframes */
    [data-testid="stDataFrame"] {{
        overflow: hidden !important;
    }}

    /* Sidebar */
    section[data-testid="stSidebar"] {{
        background:
            linear-gradient(180deg, var(--y-primary-dark) 0%, var(--y-primary) 100%) !important;
    }}
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] small,
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {{
        color: #FFFFFF !important;
        opacity: 1 !important;
    }}

    section[data-testid="stSidebar"] div[data-baseweb="select"] > div,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section {{
        background: #FFFFFF !important;
        color: var(--y-primary) !important;
        border: 1px solid #C9D6E1 !important;
        border-radius: 8px !important;
    }}
    section[data-testid="stSidebar"] div[data-baseweb="select"] span,
    section[data-testid="stSidebar"] div[data-baseweb="select"] input,
    section[data-testid="stSidebar"] div[data-baseweb="select"] svg {{
        color: var(--y-primary) !important;
        fill: var(--y-primary) !important;
    }}

    /* HOME button — guarantee contrast */
    section[data-testid="stSidebar"] div[data-testid="stButton"] > button,
    section[data-testid="stSidebar"] .stButton > button {{
        background: #FFFFFF !important;
        color: var(--y-primary) !important;
        border: 1px solid #D2DEE8 !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
    }}
    section[data-testid="stSidebar"] div[data-testid="stButton"] > button *,
    section[data-testid="stSidebar"] .stButton > button * {{
        color: var(--y-primary) !important;
    }}
    section[data-testid="stSidebar"] div[data-testid="stButton"] > button:hover {{
        background: {YUSEN_THEME['hover']} !important;
        border-color: var(--y-secondary-mid) !important;
    }}

    /* Primary action */
    div[data-testid="stButton"] > button[kind="primary"] {{
        background: var(--y-primary) !important;
        color: #FFFFFF !important;
        border: 1px solid var(--y-primary) !important;
        border-radius: 9px !important;
        font-weight: 700 !important;
        box-shadow: 0 5px 14px rgba(0,59,112,0.14) !important;
    }}
    div[data-testid="stButton"] > button[kind="primary"]:hover {{
        background: var(--y-secondary) !important;
        border-color: var(--y-secondary) !important;
    }}

    /* Tabs */
    .stTabs [data-baseweb="tab"] {{
        color: var(--y-muted) !important;
        font-weight: 600 !important;
    }}
    .stTabs [aria-selected="true"] {{
        color: var(--y-primary) !important;
    }}

    /* Notes/captions */
    .kpi-note,
    .shipment-kpi-note,
    .pic-kpi-note,
    .hc-variance-formula,
    [data-testid="stCaptionContainer"],
    [data-testid="stCaptionContainer"] p {{
        color: var(--y-muted) !important;
    }}

    @media (max-width: 1366px) {{
        .block-container {{
            padding-left: 0.9rem !important;
            padding-right: 0.9rem !important;
        }}
        .main-title {{ font-size: 28px !important; }}
        .section-title {{ font-size: 18px !important; }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# FINAL VISUAL POLISH OVERRIDES
# ============================================================
st.markdown(
    f"""
    <style>
    .shipment-kpi-card {{
        height: 132px !important;
        min-height: 132px !important;
    }}

    [data-testid="stDataFrame"] {{
        font-size: 12px !important;
    }}

    [data-testid="stDataFrame"] [role="columnheader"] {{
        font-weight: 650 !important;
        color: {YUSEN_THEME['primary']} !important;
    }}

    .paired-detail-title {{
        font-size: 16px !important;
        color: {YUSEN_THEME['primary']} !important;
        font-weight: 700 !important;
    }}

    .pic-progress-track {{
        height: 10px !important;
    }}

    /* Section 3 KPI alignment: equal title area and value baseline */
    .pic-kpi-card {{
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        justify-content: flex-start !important;
        text-align: center !important;
        padding-top: 18px !important;
    }}

    .pic-kpi-label {{
        width: 100% !important;
        min-height: 40px !important;
        height: 40px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        text-align: center !important;
        font-size: 14px !important;
        font-weight: 600 !important;
        line-height: 1.25 !important;
        margin: 0 0 8px 0 !important;
    }}

    .pic-kpi-value {{
        margin-top: 0 !important;
    }}

    .pic-kpi-note {{
        margin-top: 8px !important;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# HOME / SIDEBAR / HEADER POLISH — YUSEN EXECUTIVE FORMAT
# Consolidated UI layer for Sidebar + Main Header + Filter Summary + KPI hierarchy
# UI only — business logic, calculations, filters and data mappings are unchanged
# ============================================================
st.markdown(
    f"""
    <style>
    header[data-testid="stHeader"] {{
        background:transparent !important;
        height:0.35rem !important;
        min-height:0.35rem !important;
    }}

    [data-testid="stToolbar"] {{
        top:0.15rem !important;
    }}

    /* ------------------------------------------------------------
       LAPTOP-FIRST MAIN CANVAS
       ------------------------------------------------------------ */
    .block-container {{
        max-width: 1680px !important;
        padding-top: 0.10rem !important;
        padding-left: 1.05rem !important;
        padding-right: 1.05rem !important;
        padding-bottom: 1.5rem !important;
    }}

    /* ------------------------------------------------------------
       SIDEBAR — 248px clean executive navigation / filter rail
       ------------------------------------------------------------ */
    section[data-testid="stSidebar"] {{
        width: 248px !important;
        min-width: 248px !important;
        max-width: 248px !important;
        background: linear-gradient(180deg, #041532 0%, #06183F 100%) !important;
        border-right: 1px solid rgba(255,255,255,0.07) !important;
    }}

    section[data-testid="stSidebar"] > div:first-child {{
        width: 248px !important;
        padding-top: 0.25rem !important;
        padding-left: 0.85rem !important;
        padding-right: 0.85rem !important;
        padding-bottom: 0.65rem !important;
    }}

    .sidebar-brand {{
        display:flex;
        align-items:center;
        gap:0;
        min-height:10px;
        height:10px;
        color:#FFFFFF;
        margin:0 0 2px 1px;
    }}

    .sidebar-brand-compact {{
        width:100%;
        justify-content:flex-start;
    }}

    .sidebar-brand-mark {{
        width:0;
        height:0;
        display:none;
        position:relative;
        display:inline-block;
        flex:0 0 32px;
    }}

    .sidebar-brand-mark::before,
    .sidebar-brand-mark::after {{
        content:"";
        position:absolute;
        left:0;
        height:3px;
        border-radius:999px;
        background:#FFFFFF;
        transform:skewX(-28deg);
    }}
    .sidebar-brand-mark::before {{ width:33px; top:4px; }}
    .sidebar-brand-mark::after {{ width:26px; top:13px; left:6px; }}

    /* HOME */
    section[data-testid="stSidebar"] div[data-testid="stButton"] > button {{
        min-height:41px !important;
        height:41px !important;
        background:#FFFFFF !important;
        color:#06183F !important;
        border:1px solid #D5E1EA !important;
        border-radius:9px !important;
        font-size:13px !important;
        font-weight:700 !important;
        box-shadow:0 3px 10px rgba(0,0,0,0.09) !important;
        margin:0 0 10px 0 !important;
    }}

    section[data-testid="stSidebar"] div[data-testid="stButton"] > button *,
    section[data-testid="stSidebar"] .stButton > button * {{
        color:#06183F !important;
    }}

    section[data-testid="stSidebar"] div[data-testid="stButton"] > button:hover {{
        color:#E6761B !important;
        border-color:#E6761B !important;
        background:#FFFFFF !important;
    }}

    section[data-testid="stSidebar"] div[data-testid="stButton"] > button:hover * {{
        color:#E6761B !important;
    }}

    /* FILTERS heading */
    .sidebar-filter-title {{
        display:flex;
        align-items:center;
        gap:7px;
        color:#FFFFFF;
        font-size:15px;
        line-height:1.2;
        font-weight:800;
        margin:3px 0 8px 0;
        letter-spacing:0.01em;
    }}

    .sidebar-filter-title::after {{
        content:"";
        width:20px;
        height:2px;
        border-radius:999px;
        background:#E6761B;
        display:block;
    }}

    .sidebar-filter-caption {{
        display:none !important;
        color:#D5EAF8;
        font-size:11.5px;
        line-height:1.3;
        font-weight:500;
        margin:0 0 11px 0;
    }}

    /* Field labels */
    section[data-testid="stSidebar"] label {{
        color:#FFFFFF !important;
        font-size:12.5px !important;
        line-height:1.25 !important;
        font-weight:700 !important;
        letter-spacing:0.015em !important;
        margin-bottom:5px !important;
    }}

    /* Select boxes */
    section[data-testid="stSidebar"] div[data-baseweb="select"] > div {{
        min-height:41px !important;
        height:41px !important;
        background:#FFFFFF !important;
        border:1px solid #D5E1EA !important;
        border-radius:9px !important;
        box-shadow:none !important;
    }}

    section[data-testid="stSidebar"] div[data-baseweb="select"] span,
    section[data-testid="stSidebar"] div[data-baseweb="select"] input {{
        color:#06183F !important;
        font-size:13.5px !important;
        font-weight:500 !important;
        opacity:1 !important;
    }}

    section[data-testid="stSidebar"] div[data-baseweb="select"] svg {{
        color:#06183F !important;
        fill:#06183F !important;
        opacity:1 !important;
    }}

    section[data-testid="stSidebar"] [data-testid="stSelectbox"] {{
        margin-bottom:0.28rem !important;
    }}

    section[data-testid="stSidebar"] hr {{
        border:0 !important;
        border-top:1px solid rgba(213,234,248,0.20) !important;
        margin:10px 0 9px 0 !important;
    }}

    /* Upload — compact white card */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] {{
        margin-top:1px !important;
    }}

    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section {{
        min-height:78px !important;
        background:#FFFFFF !important;
        border:1px solid #D5E1EA !important;
        border-radius:9px !important;
        padding:6px 8px !important;
        box-shadow:0 2px 9px rgba(0,0,0,0.07) !important;
    }}

    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section * {{
        color:#06183F !important;
        opacity:1 !important;
        font-size:11.5px !important;
    }}

    section[data-testid="stSidebar"] [data-testid="stFileUploader"] button {{
        background:#FFFFFF !important;
        color:#06183F !important;
        border:1px solid #B9CAD8 !important;
        border-radius:8px !important;
        font-size:12px !important;
        font-weight:650 !important;
        min-height:34px !important;
    }}

    /* Sidebar footer — compact application metadata */
    .sidebar-footer {{
        margin-top:14px;
        padding-top:10px;
        border-top:1px solid rgba(255,255,255,0.16);
        display:flex;
        align-items:center;
        justify-content:center;
        gap:4px;
        white-space:nowrap;
        color:rgba(255,255,255,0.72);
        font-size:8.7px;
        line-height:1.25;
        font-weight:500;
        letter-spacing:-0.01em;
    }}
    .sidebar-footer .footer-sep {{
        color:#E6761B;
        opacity:0.95;
        font-weight:700;
    }}

    /* ------------------------------------------------------------
       MAIN HEADER — compact executive card
       ------------------------------------------------------------ */
    .main-header {{
        position:relative !important;
        overflow:hidden !important;
        background:rgba(255,255,255,0.985) !important;
        backdrop-filter:blur(8px);
        -webkit-backdrop-filter:blur(8px);
        border:1px solid #D5E1EA !important;
        border-left:5px solid #0DBAEE !important;
        border-radius:11px !important;
        padding:9px 14px 8px 14px !important;
        margin:0 !important;
        box-shadow:0 4px 14px rgba(6,24,63,0.08) !important;
    }}

    /* Make the Streamlit element wrapper sticky, not the inner HTML.
       This avoids clipping/stacking issues on Streamlit Cloud. */
    div[data-testid="stVerticalBlock"] > div:has(.main-header) {{
        position:sticky !important;
        top:0.10rem !important;
        z-index:1000 !important;
        background:{COLORS['bg']} !important;
        padding-top:0.15rem !important;
        padding-bottom:0.30rem !important;
    }}

    /* Remove the old decorative arrow to preserve laptop width. */
    .main-header::after {{
        content:none !important;
        display:none !important;
    }}

    .main-title {{
        color:#06183F !important;
        font-size:28px !important;
        line-height:1.06 !important;
        font-weight:800 !important;
        letter-spacing:-0.025em !important;
        padding-right:0 !important;
        margin:0 !important;
    }}

    .subtitle {{
        color:#5B6575 !important;
        font-size:12.5px !important;
        line-height:1.25 !important;
        font-weight:450 !important;
        margin-top:2px !important;
    }}

    /* ------------------------------------------------------------
       SELECTED FILTER SUMMARY
       ------------------------------------------------------------ */
    .filter-summary-card {{
        display:flex;
        align-items:center;
        min-height:40px;
        gap:0;
        background:#F8FBFE;
        border:1px solid #E0E7EE;
        border-radius:8px;
        box-shadow:none;
        margin:6px 0 0 0;
        padding:3px 9px;
    }}

    .filter-summary-item {{
        display:grid;
        grid-template-columns:30px minmax(0,1fr);
        grid-template-rows:auto auto;
        column-gap:8px;
        align-items:center;
        min-width:185px;
        padding:0 16px 0 0;
        margin-right:16px;
        border-right:1px solid #E4EBF1;
    }}

    .filter-summary-item:last-child {{
        border-right:0;
        margin-right:0;
    }}

    .filter-summary-icon {{
        grid-row:1 / span 2;
        width:25px;
        height:25px;
        border-radius:7px;
        display:flex;
        align-items:center;
        justify-content:center;
        color:#06183F;
        background:#EEF7FC;
        border:1px solid #D5EAF8;
    }}

    .filter-summary-icon svg {{
        width:16px;
        height:16px;
        fill:none;
        stroke:currentColor;
        stroke-width:1.8;
        stroke-linecap:round;
        stroke-linejoin:round;
    }}

    .filter-summary-label {{
        color:#5B6575;
        font-size:11.5px;
        line-height:1.2;
        font-weight:600;
        margin:0;
    }}

    .filter-summary-value {{
        color:#06183F;
        font-size:13.5px;
        line-height:1.2;
        font-weight:750;
        margin-top:1px;
    }}

    /* ------------------------------------------------------------
       SECTION HIERARCHY
       ------------------------------------------------------------ */
    .section-title {{
        color:#06183F !important;
        font-size:19px !important;
        line-height:1.22 !important;
        font-weight:800 !important;
        border-left:4px solid #E6761B !important;
        padding-left:9px !important;
        margin:18px 0 9px 0 !important;
    }}


    /* ------------------------------------------------------------
       KPI ICON SYSTEM — clean executive layout
       Icons are concentrated in HC KPI cards and filter summary
       ------------------------------------------------------------ */
    .kpi-icon-circle {{
        width:44px;
        height:44px;
        min-width:44px;
        border-radius:50%;
        display:inline-flex;
        align-items:center;
        justify-content:center;
        flex:0 0 44px;
    }}

    .kpi-icon-circle svg {{
        width:23px;
        height:23px;
        fill:none;
        stroke:currentColor;
        stroke-width:1.8;
        stroke-linecap:round;
        stroke-linejoin:round;
    }}

    .hc-main-row {{
        display:flex;
        align-items:center;
        justify-content:center;
        gap:10px;
        width:100%;
        margin:7px 0 5px 0;
    }}

    .hc-main-row .hc-kpi-total {{
        width:auto !important;
        margin:0 !important;
        text-align:left !important;
    }}



    /* ------------------------------------------------------------
       KPI / CARD HIERARCHY
       ------------------------------------------------------------ */
    .kpi-card,
    .hc-kpi-card,
    .shipment-kpi-card,
    .pic-kpi-card,
    .pic-status-card,
    .workload-status-panel,
    [data-testid="stPlotlyChart"],
    [data-testid="stDataFrame"] {{
        background:#FFFFFF !important;
        border:1px solid #D5E1EA !important;
        border-radius:11px !important;
        box-shadow:0 2px 7px rgba(6,24,63,0.035) !important;
    }}

    .kpi-label,
    .shipment-kpi-label,
    .pic-kpi-label {{
        color:#5B6575 !important;
        font-size:13px !important;
        line-height:1.22 !important;
        font-weight:650 !important;
        letter-spacing:0.015em !important;
        text-transform:none !important;
    }}

    .kpi-value,
    .hc-kpi-total,
    .shipment-kpi-value,
    .pic-kpi-value,
    .pic-status-value {{
        font-size:32px !important;
        line-height:1.04 !important;
        font-weight:780 !important;
        letter-spacing:-0.02em !important;
    }}

    /* Section 1 HC cards — compact but fully readable */
    .hc-kpi-card {{
        height:158px !important;
        min-height:158px !important;
        padding:14px 14px 12px 14px !important;
    }}

    .hc-kpi-total {{
        margin-top:8px !important;
        margin-bottom:5px !important;
    }}

    .hc-detail-row {{
        padding-top:9px !important;
        gap:8px !important;
    }}

    .hc-detail-label {{
        font-size:11.5px !important;
        font-weight:650 !important;
    }}

    .hc-detail-value {{
        font-size:19px !important;
        font-weight:750 !important;
        margin-top:2px !important;
    }}

    .hc-variance-formula {{
        font-size:11.5px !important;
        line-height:1.25 !important;
        margin-top:7px !important;
        margin-bottom:6px !important;
    }}

    .status-badge {{
        font-size:11.5px !important;
        line-height:1.2 !important;
        font-weight:750 !important;
    }}

    /* Other KPI cards */
    .shipment-kpi-card {{
        height:124px !important;
        min-height:124px !important;
        padding:14px 14px !important;
    }}

    .pic-kpi-card {{
        height:136px !important;
        min-height:136px !important;
        padding:15px 14px 13px 14px !important;
    }}

    .pic-kpi-label {{
        min-height:34px !important;
        height:34px !important;
        font-size:13px !important;
        margin:0 0 6px 0 !important;
    }}

    .pic-kpi-note,
    .kpi-note,
    .shipment-kpi-note {{
        font-size:11px !important;
        line-height:1.3 !important;
    }}


    /* FTE Workload Status — align status text with adjacent KPI value */
    .workload-status-text {{
        font-size:32px !important;
        line-height:1.05 !important;
        font-weight:800 !important;
        padding:7px 18px !important;
        min-width:180px;
        text-align:center;
        margin-top:8px !important;
    }}

    @media (max-width:1366px) {{
        .workload-status-text {{
            font-size:28px !important;
            padding:6px 16px !important;
            min-width:165px;
        }}
    }}

    /* ------------------------------------------------------------
       TABLE READABILITY
       ------------------------------------------------------------ */
    .paired-detail-table {{
        font-size:12px !important;
    }}

    .paired-detail-table thead th {{
        font-size:12px !important;
        font-weight:700 !important;
        padding:7px 7px !important;
    }}

    .paired-detail-table tbody td {{
        font-size:12px !important;
        padding:6px 7px !important;
    }}

    .customer-name-cell {{
        font-size:11.5px !important;
    }}

    .paired-detail-foot {{
        font-size:11px !important;
    }}

    [data-testid="stDataFrame"] {{
        font-size:12px !important;
    }}

    /* ------------------------------------------------------------
       PLOTLY / CHART CARD DENSITY
       ------------------------------------------------------------ */
    [data-testid="stPlotlyChart"] {{
        padding:6px 8px 3px 8px !important;
    }}

    /* ------------------------------------------------------------
       LAPTOP 1366 × 768
       Reduce whitespace, not core readability.
       ------------------------------------------------------------ */
    @media (max-width:1366px) {{
        .kpi-icon-circle {{
            width:42px;
            height:42px;
            min-width:42px;
            flex-basis:42px;
        }}
        .kpi-icon-circle svg {{
            width:22px;
            height:22px;
        }}
        .hc-main-row {{
            gap:8px;
            margin:5px 0 4px 0;
        }}

        section[data-testid="stSidebar"],
        section[data-testid="stSidebar"] > div:first-child {{
            width:260px !important;
            min-width:260px !important;
            max-width:260px !important;
        }}

        .block-container {{
            padding-top:0.25rem !important;
            padding-left:0.80rem !important;
            padding-right:0.80rem !important;
            padding-bottom:1.25rem !important;
        }}

        .main-header {{
            padding:8px 12px 7px 12px !important;
        }}
        div[data-testid="stVerticalBlock"] > div:has(.main-header) {{
            top:0.05rem !important;
        }}

        .main-title {{
            font-size:27px !important;
        }}

        .subtitle {{
            font-size:12px !important;
        }}

        .filter-summary-card {{
            min-height:50px !important;
            padding:6px 12px !important;
            margin-bottom:10px !important;
        }}

        .filter-summary-item {{
            min-width:185px !important;
            padding-right:18px !important;
            margin-right:18px !important;
        }}

        .section-title {{
            font-size:18px !important;
            margin-top:16px !important;
            margin-bottom:8px !important;
        }}

        .hc-kpi-card {{
            height:150px !important;
            min-height:150px !important;
            padding:12px 12px 10px 12px !important;
        }}

        .kpi-value,
        .hc-kpi-total,
        .shipment-kpi-value,
        .pic-kpi-value,
        .pic-status-value {{
            font-size:30px !important;
        }}

        .kpi-label,
        .shipment-kpi-label,
        .pic-kpi-label {{
            font-size:12.5px !important;
        }}

        .shipment-kpi-card {{
            height:118px !important;
            min-height:118px !important;
        }}

        .pic-kpi-card {{
            height:130px !important;
            min-height:130px !important;
        }}
    }}

    @media (max-width:1100px) {{
        .filter-summary-card {{
            flex-wrap:wrap;
            gap:7px 0;
        }}
        .filter-summary-item {{
            min-width:165px;
        }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# OFFICE COMPARISON UI — ALL OFFICES ONLY
# UI layer only; KPI formulas / filters / business logic unchanged
# ============================================================
st.markdown(
    f"""
    <style>
    .office-comparison-heading {{
        display:flex; align-items:center; gap:8px; color:#06183F;
        font-size:15px; line-height:1.2; font-weight:750;
        margin:10px 0 8px 1px;
    }}
    .office-comparison-heading::before {{
        content:""; width:22px; height:3px; border-radius:999px;
        background:#E6761B; display:inline-block;
    }}
    .office-compare-card {{
        --office-status:#3F5B81; --office-status-bg:#EEF3F8;
        position:relative; background:#FFFFFF; border:1px solid #D5E1EA;
        border-top:4px solid var(--office-status); border-radius:11px;
        min-height:154px; padding:11px 12px 10px 12px; box-sizing:border-box;
        box-shadow:0 2px 7px rgba(6,24,63,0.035); overflow:hidden;
    }}
    .office-compare-top {{
        display:flex; align-items:center; justify-content:space-between;
        gap:8px; margin-bottom:9px;
    }}
    .office-compare-name {{
        color:#06183F; font-size:18px; line-height:1; font-weight:800;
    }}
    .office-compare-status {{
        display:inline-flex; align-items:center; justify-content:center;
        max-width:125px; min-height:24px; padding:3px 9px; border-radius:999px;
        background:var(--office-status-bg); color:var(--office-status);
        font-size:10.5px; line-height:1.1; font-weight:800; white-space:nowrap;
    }}
    .office-compare-primary {{
        display:flex; align-items:baseline; justify-content:space-between;
        gap:10px; padding-bottom:8px; margin-bottom:7px;
        border-bottom:1px solid #E8EEF3;
    }}
    .office-compare-primary-label {{
        color:#667085; font-size:10.5px; line-height:1.15; font-weight:600;
    }}
    .office-compare-primary-value {{
        color:#06183F; font-size:24px; line-height:1; font-weight:800;
        letter-spacing:-0.02em; white-space:nowrap;
    }}
    .office-compare-grid {{
        display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:7px 10px;
    }}
    .office-compare-metric {{ min-width:0; }}
    .office-compare-metric-label {{
        color:#7A8699; font-size:9.8px; line-height:1.15; font-weight:600;
        margin-bottom:2px; white-space:nowrap;
    }}
    .office-compare-metric-value {{
        color:#06183F; font-size:14px; line-height:1.15; font-weight:750;
        font-variant-numeric:tabular-nums; white-space:nowrap;
    }}
    .office-compare-metric-value.negative {{ color:#D92D20; }}
    .office-compare-metric-value.positive {{ color:#6EA52B; }}
    @media (max-width:1366px) {{
        .office-compare-card {{ min-height:146px; padding:10px 10px 9px 10px; }}
        .office-compare-name {{ font-size:17px; }}
        .office-compare-primary-value {{ font-size:22px; }}
        .office-compare-status {{ font-size:9.8px; padding:3px 7px; }}
        .office-compare-metric-label {{ font-size:9.3px; }}
        .office-compare-metric-value {{ font-size:13px; }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# HELPER FUNCTIONS
# ============================================================


def safe_float(value) -> float:
    if pd.isna(value) or value == "":
        return 0.0
    if isinstance(value, str):
        value = value.replace(",", "").strip()
        if value.startswith("="):
            return 0.0
    try:
        return float(value)
    except (TypeError, ValueError, OverflowError):
        return 0.0


def clean_col(col) -> str:
    col = str(col).replace("\n", " ").replace("\r", " ")
    col = re.sub(r"\s+", " ", col).strip()
    return col


def normalize_office(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def parse_month(value) -> pd.Timestamp | pd.NaT:
    if pd.isna(value) or value == "":
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return pd.Timestamp(year=value.year, month=value.month, day=1)
    text = str(value).strip()

    # Adapter for the new MASTER DATA SOURCE workbook:
    # Month is stored as Apr..Mar while the dashboard logic still works with
    # real MonthDate values. FY2026 = Apr-Dec 2026 + Jan-Mar 2027.
    month_only = {
        "jan": (2027, 1), "feb": (2027, 2), "mar": (2027, 3),
        "apr": (2026, 4), "may": (2026, 5), "jun": (2026, 6),
        "jul": (2026, 7), "aug": (2026, 8), "sep": (2026, 9),
        "oct": (2026, 10), "nov": (2026, 11), "dec": (2026, 12),
    }
    key = text[:3].lower()
    if len(text) <= 4 and key in month_only:
        year, month = month_only[key]
        return pd.Timestamp(year=year, month=month, day=1)

    for fmt in ["%b-%y", "%b-%Y", "%Y-%m", "%m/%Y", "%Y/%m"]:
        try:
            dt = pd.to_datetime(text, format=fmt)
            return pd.Timestamp(year=dt.year, month=dt.month, day=1)
        except (TypeError, ValueError, OverflowError):
            pass
    try:
        dt = pd.to_datetime(text, errors="coerce")
        if pd.isna(dt):
            return pd.NaT
        return pd.Timestamp(year=dt.year, month=dt.month, day=1)
    except (TypeError, ValueError, OverflowError):
        return pd.NaT


def format_month(ts) -> str:
    if pd.isna(ts):
        return ""
    return pd.Timestamp(ts).strftime("%b-%y")


def numeric_series(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0)


def read_sheet(path: str | Path, sheet: str, header: int = 1) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=sheet, header=header, engine="openpyxl")
        df.columns = [clean_col(c) for c in df.columns]
        df = df.dropna(how="all")
        return df
    except Exception:
        logger.exception("Unable to read sheet %r from workbook %s", sheet, path)
        return pd.DataFrame()


def ensure_cols(df: pd.DataFrame, cols: Iterable[str]) -> pd.DataFrame:
    for c in cols:
        if c not in df.columns:
            df[c] = np.nan
    return df


def first_existing(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    clean_map = {clean_col(c).lower(): c for c in df.columns}
    for c in candidates:
        key = clean_col(c).lower()
        if key in clean_map:
            return clean_map[key]
    return None


def weighted_period_avg(df: pd.DataFrame, value_col: str, group_col: str = "MonthDate") -> float:
    """Average of valid monthly totals only; blank future months are excluded."""
    if df.empty or value_col not in df.columns or group_col not in df.columns:
        return 0.0

    valid = df[[group_col, value_col]].copy()
    valid[value_col] = pd.to_numeric(valid[value_col], errors="coerce")
    valid = valid.dropna(subset=[group_col, value_col])

    if valid.empty:
        return 0.0

    monthly = (
        valid.groupby(group_col, dropna=True)[value_col]
        .sum(min_count=1)
        .reset_index()
        .dropna(subset=[value_col])
    )
    if monthly.empty:
        return 0.0
    return float(monthly[value_col].mean())


def safe_div(num: float, den: float) -> float:
    return float(num) / float(den) if den not in [0, 0.0] and not pd.isna(den) else 0.0


def fmt_num(value: float, decimals: int = 1, suffix: str = "") -> str:
    if value is None or pd.isna(value):
        return "N/A"
    return f"{value:,.{decimals}f}{suffix}"


def fmt_int(value: float) -> str:
    if value is None or pd.isna(value):
        return "N/A"
    return f"{value:,.0f}"


def fmt_pct(value: float) -> str:
    if value is None or pd.isna(value):
        return "N/A"
    return f"{value * 100:,.1f}%"


def status_from_util(util: float) -> Tuple[str, str, str]:
    """Standard workload status rule used across KPI/status displays."""
    if util <= 0:
        return "NO DATA", COLORS["muted"], COLORS["light_blue"]
    if util < 0.90:
        return "LESS LOAD", COLORS["green"], "#DCFCE7"
    if util <= 0.95:
        return "BALANCED", COLORS["blue"], "#DBEAFE"
    if util <= 1.00:
        return "HIGH LOAD", COLORS["high_load"], "#FFF4D6"
    return "OVERLOAD", COLORS["red"], "#FEE2E2"



def _office_compare_card(
    office_name: str,
    primary_label: str,
    primary_value: str,
    metrics: List[Tuple[str, str, str]],
    status_text: str,
    status_color: str,
    status_bg: str,
) -> None:
    # Render one compact office benchmark card. UI only.
    metric_html = "".join(
        f'<div class="office-compare-metric"><div class="office-compare-metric-label">{html.escape(str(label))}</div><div class="office-compare-metric-value {css_class}">{html.escape(str(value))}</div></div>'
        for label, value, css_class in metrics
    )
    st.markdown(
        f'<div class="office-compare-card" style="--office-status:{status_color};--office-status-bg:{status_bg};"><div class="office-compare-top"><div class="office-compare-name">{html.escape(str(office_name))}</div><div class="office-compare-status">{html.escape(str(status_text))}</div></div><div class="office-compare-primary"><div class="office-compare-primary-label">{html.escape(str(primary_label))}</div><div class="office-compare-primary-value">{html.escape(str(primary_value))}</div></div><div class="office-compare-grid">{metric_html}</div></div>',
        unsafe_allow_html=True,
    )


def _office_comparison_heading(title: str) -> None:
    st.markdown(
        f'<div class="office-comparison-heading">{html.escape(title)}</div>',
        unsafe_allow_html=True,
    )


def render_hc_office_comparison(hc_filtered_all_offices: pd.DataFrame) -> None:
    # Reuse existing HC source-of-truth functions and status thresholds.
    _office_comparison_heading("Utilization by Office")
    cols = st.columns(4, gap="small")
    for col, office_name in zip(cols, STANDARD_OFFICES):
        if hc_filtered_all_offices is not None and not hc_filtered_all_offices.empty and "Office" in hc_filtered_all_offices.columns:
            office_df = hc_filtered_all_offices[hc_filtered_all_offices["Office"] == office_name].copy()
        else:
            office_df = pd.DataFrame()

        if office_df.empty:
            actual = required = gap = util = float("nan")
            status_text, status_color, status_bg = "NO DATA", COLORS["muted"], COLORS["light_blue"]
        else:
            actual = weighted_period_avg(office_df, "Total Actual HC")
            required = weighted_period_avg(office_df, "Total Required HC")
            gap = required - actual
            util = hc_capacity_utilization(office_df)
            # Office status is determined by Office Workload (utilization),
            # using the standard workload thresholds:
            # < 90%       -> LESS LOAD / Green
            # 90% - 95%   -> BALANCED / Blue
            # >95% - 100% -> HIGH LOAD / Amber
            # >100%       -> OVERLOAD / Red
            if pd.isna(util):
                status_text, status_color, status_bg = "NO DATA", COLORS["muted"], COLORS["light_blue"]
            else:
                status_text, status_color, status_bg = status_from_util(util)

        gap_class = "negative" if (not pd.isna(gap) and gap > 0) else ("positive" if (not pd.isna(gap) and gap < 0) else "")
        gap_text = "N/A" if pd.isna(gap) else f"{gap:+,.2f}"
        with col:
            _office_compare_card(
                office_name,
                "Office Workload",
                "N/A" if pd.isna(util) else fmt_pct(util),
                [
                    ("Actual HC", "N/A" if pd.isna(actual) else fmt_num(actual, 2), ""),
                    ("Required HC", "N/A" if pd.isna(required) else fmt_num(required, 2), ""),
                    ("HC Gap", gap_text, gap_class),
                    ("Status", status_text.title(), ""),
                ],
                status_text, status_color, status_bg,
            )


def _fte_office_summary(office_fte: pd.DataFrame, selected_month: str) -> Tuple[float, float, float, Tuple[str, str, str]]:
    # Apply the exact Section 3 month logic to one office.
    if office_fte is None or office_fte.empty:
        return float("nan"), float("nan"), float("nan"), ("NO DATA", COLORS["muted"], COLORS["light_blue"])
    d = office_fte.copy()
    d["Available Time"] = pd.to_numeric(d.get("Available Time"), errors="coerce")
    d["Actual Working Time"] = pd.to_numeric(d.get("Actual Working Time"), errors="coerce")
    monthly = (
        d.dropna(subset=["MonthDate", "Available Time", "Actual Working Time"])
        .groupby("MonthDate", as_index=False)
        .agg(
            Total_Available_Time=("Available Time", "sum"),
            Total_Actual_Working_Time=("Actual Working Time", "sum"),
        )
    )
    if monthly.empty:
        return float("nan"), float("nan"), float("nan"), ("NO DATA", COLORS["muted"], COLORS["light_blue"])
    if str(selected_month).strip().lower() == "all":
        total_available = float(monthly["Total_Available_Time"].sum())
        total_actual = float(monthly["Total_Actual_Working_Time"].sum())
    else:
        row = monthly.sort_values("MonthDate").iloc[-1]
        total_available = float(row["Total_Available_Time"])
        total_actual = float(row["Total_Actual_Working_Time"])
    workload = safe_div(total_actual, total_available)
    return total_available, total_actual, workload, status_from_util(workload)


def render_fte_office_comparison(fte_filtered_all_offices: pd.DataFrame, selected_month: str) -> None:
    # Reuse exact Section 3 FTE formula / month handling.
    _office_comparison_heading("Office Workload per PIC")
    cols = st.columns(4, gap="small")
    for col, office_name in zip(cols, STANDARD_OFFICES):
        if fte_filtered_all_offices is not None and not fte_filtered_all_offices.empty and "Office" in fte_filtered_all_offices.columns:
            office_df = fte_filtered_all_offices[fte_filtered_all_offices["Office"] == office_name].copy()
        else:
            office_df = pd.DataFrame()
        available, actual, workload, status = _fte_office_summary(office_df, selected_month)
        status_text, status_color, status_bg = status
        variance = actual - available if not pd.isna(actual) and not pd.isna(available) else float("nan")
        variance_class = "negative" if (not pd.isna(variance) and variance > 0) else ("positive" if not pd.isna(variance) else "")
        with col:
            _office_compare_card(
                office_name,
                "Average PIC Workload",
                "N/A" if pd.isna(workload) else f"{int(np.ceil(workload * 100))}%",
                [
                    ("Available Time (hour)", "N/A" if pd.isna(available) else fmt_num(available, 0), ""),
                    ("Actual Time (hour)", "N/A" if pd.isna(actual) else fmt_num(actual, 0), ""),
                    ("Gap", "N/A" if pd.isna(variance) else fmt_num(variance, 0), variance_class),
                    ("Status", status_text.title(), ""),
                ],
                status_text, status_color, status_bg,
            )

    st.markdown(
        f"""
        <div style="
            margin:8px 0 0;
            padding:13px 16px;
            display:flex;
            justify-content:center;
            align-items:center;
            flex-wrap:wrap;
            gap:10px 28px;
            color:#667085;
            background:#F4F6F8;
            border-top:1px solid #E3E8EF;
            border-bottom:1px solid #D8E0E8;
            font-size:12px;
            font-weight:500;
            line-height:1.4;">
            <span style="display:inline-flex;align-items:center;gap:7px;white-space:nowrap;">
                <span style="width:10px;height:10px;border-radius:2px;background:{COLORS['red']};"></span>
                Overload &gt;100%
            </span>
            <span style="display:inline-flex;align-items:center;gap:7px;white-space:nowrap;">
                <span style="width:10px;height:10px;border-radius:2px;background:{COLORS['high_load']};"></span>
                High Load &gt;95–100%
            </span>
            <span style="display:inline-flex;align-items:center;gap:7px;white-space:nowrap;">
                <span style="width:10px;height:10px;border-radius:2px;background:{COLORS['blue']};"></span>
                Balanced 90–95%
            </span>
            <span style="display:inline-flex;align-items:center;gap:7px;white-space:nowrap;">
                <span style="width:10px;height:10px;border-radius:2px;background:{COLORS['green']};"></span>
                Less Load &lt;90%
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )



def ui_icon_svg(kind: str, color: str, bg: str, circle_class: str = "kpi-icon-circle") -> str:
    """Inline corporate SVG icon - UI only, no external dependency."""
    icons = {
        "people": """
            <circle cx="12" cy="8" r="3"></circle>
            <circle cx="5.5" cy="10" r="2.2"></circle>
            <circle cx="18.5" cy="10" r="2.2"></circle>
            <path d="M7.5 20v-2.2c0-3 2-5.3 4.5-5.3s4.5 2.3 4.5 5.3V20"></path>
            <path d="M2 19v-1.4c0-2.3 1.5-4 3.6-4"></path>
            <path d="M22 19v-1.4c0-2.3-1.5-4-3.6-4"></path>
        """,
        "people_active": """
            <circle cx="10" cy="8" r="3"></circle>
            <path d="M4.5 20v-2.3c0-3.1 2.4-5.4 5.5-5.4s5.5 2.3 5.5 5.4V20"></path>
            <path d="M16 9l1.8 1.8L21 7"></path>
        """,
        "people_required": """
            <circle cx="9" cy="8" r="3"></circle>
            <path d="M3.5 20v-2.3c0-3.1 2.4-5.4 5.5-5.4s5.5 2.3 5.5 5.4V20"></path>
            <path d="M18 8v7M14.5 11.5h7"></path>
        """,
        "balance": """
            <path d="M12 4v16"></path>
            <path d="M6 7h12"></path>
            <path d="M6 7l-3 6h6L6 7z"></path>
            <path d="M18 7l-3 6h6l-3-6z"></path>
            <path d="M8 20h8"></path>
        """,
        "package": """
            <path d="M4 8l8-4 8 4-8 4-8-4z"></path>
            <path d="M4 8v8l8 4 8-4V8"></path>
            <path d="M12 12v8"></path>
        """,
        "customers": """
            <circle cx="9" cy="8" r="3"></circle>
            <circle cx="17" cy="9" r="2.5"></circle>
            <path d="M3.5 20v-2c0-3.1 2.4-5.3 5.5-5.3s5.5 2.2 5.5 5.3v2"></path>
            <path d="M15 14c2.8 0 5 1.8 5 4.5V20"></path>
        """,
        "clipboard": """
            <rect x="5" y="5" width="14" height="16" rx="2"></rect>
            <path d="M9 5V3h6v2"></path>
            <path d="M8.5 10h7M8.5 14h7M8.5 18h5"></path>
        """,
        "gauge": """
            <path d="M4 17a8 8 0 0 1 16 0"></path>
            <path d="M12 17l4-5"></path>
            <circle cx="12" cy="17" r="1.4"></circle>
        """,
        "target": """
            <circle cx="12" cy="12" r="8"></circle>
            <circle cx="12" cy="12" r="4"></circle>
            <path d="M12 12l6-6"></path>
            <path d="M16 6h3v3"></path>
        """,
    }
    path = icons.get(kind, icons["clipboard"])
    return f'<span class="{circle_class}" style="color:{color};background:{bg};"><svg viewBox="0 0 24 24" aria-hidden="true">{path}</svg></span>'


def hc_icon_for_label(label: str) -> str:
    key = str(label).upper()
    if "APPROVED" in key:
        return ui_icon_svg("people", "#06183F", "#EEF3F8")
    if "ACTUAL" in key:
        return ui_icon_svg("people_active", "#0DBAEE", "#E8F8FD")
    if "REQUIRED" in key:
        return ui_icon_svg("people_required", "#E6761B", "#FFF2E8")
    return ui_icon_svg("balance", "#6EA52B", "#F1F8E8")


def general_kpi_icon(label: str) -> str:
    key = str(label).lower()
    if "shipment" in key:
        return ui_icon_svg("package", "#0DBAEE", "#E8F8FD")
    if "customer" in key:
        return ui_icon_svg("customers", "#06183F", "#EEF3F8")
    if "workload" in key or "working time" in key or "available time" in key:
        return ui_icon_svg("clipboard", "#0DBAEE", "#E8F8FD")
    return ui_icon_svg("clipboard", "#06183F", "#EEF3F8")


def kpi_card(label: str, value: str, note: str = "", status: Optional[Tuple[str, str, str]] = None):
    badge = ""
    if status:
        txt, color, bg = status
        badge = f'<span class="status-badge" style="color:{color};background:{bg};">{txt}</span>'
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
            {badge}
        </div>
        """,
        unsafe_allow_html=True,
    )



def hc_detail_card(
    label: str,
    total_value: float,
    mng_value: Optional[float] = None,
    pic_value: Optional[float] = None,
    note_left: str = "MNG",
    note_right: str = "PIC",
    status_text: Optional[str] = None,
    status_color: Optional[str] = None,
    status_bg: Optional[str] = None,
):
    """Executive HC card with equal height and two aligned detail blocks at the bottom."""
    details_html = ""
    if mng_value is not None or pic_value is not None:
        left_val = fmt_num(mng_value or 0, 2)
        right_val = fmt_num(pic_value or 0, 2)
        details_html = f"""
        <div class="hc-detail-row">
            <div class="hc-detail-item">
                <div class="hc-detail-label">{note_left}</div>
                <div class="hc-detail-value">{left_val}</div>
            </div>
            <div class="hc-detail-divider"></div>
            <div class="hc-detail-item">
                <div class="hc-detail-label">{note_right}</div>
                <div class="hc-detail-value">{right_val}</div>
            </div>
        </div>
        """

    status_html = ""
    if status_text:
        status_html = (
            f'<span class="status-badge" '
            f'style="color:{status_color};background:{status_bg};margin-top:10px;">'
            f'{status_text}</span>'
        )

    # UI-only semantic color hierarchy for Office Capacity Snapshot:
    # Approved = Navy, Actual = Corporate Blue, Required = Orange.
    label_upper = label.upper()
    if "REQUIRED" in label_upper:
        total_color_class = "hc-total-required"
    elif "ACTUAL" in label_upper:
        total_color_class = "hc-total-actual"
    else:
        total_color_class = "hc-total-approved"

    st.markdown(
        f"""
        <div class="hc-kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="hc-main-row">
                {hc_icon_for_label(label)}
                <div class="hc-kpi-total {total_color_class}">{fmt_num(total_value, 2)}</div>
            </div>
            {status_html}
            {details_html}
        </div>
        """,
        unsafe_allow_html=True,
    )



def hc_variance_card(
    label: str,
    value: float,
    formula_text: str,
    status_text: str,
    status_color: str,
    status_bg: str,
):
    """Centered variance card to visually balance the HC cards."""
    st.markdown(
        f"""
        <div class="hc-kpi-card hc-variance-card">
            <div class="kpi-label">{label}</div>
            <div class="hc-main-row">
                {ui_icon_svg("balance", "#6EA52B", "#F1F8E8")}
                <div class="hc-kpi-total" style="color:{status_color} !important;">{fmt_num(value, 2)}</div>
            </div>
            <div class="hc-variance-formula">{formula_text}</div>
            <span class="status-badge hc-variance-status"
                  style="color:{status_color};background:{status_bg};">
                {status_text}
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )



def shipment_kpi_card(label: str, value: str, note: str = ""):
    """Equal-size centered KPI card for Shipment Volume section."""
    st.markdown(
        f"""
        <div class="shipment-kpi-card">
            <div class="shipment-kpi-label">{label}</div>
            <div class="shipment-kpi-value">{value}</div>
            <div class="shipment-kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )



def pic_kpi_card(label: str, value: str, note: str = "", unit: str = ""):
    """Compact numeric KPI card: Title -> Unit -> Value -> Note."""
    unit_html = f'<div class="pic-kpi-unit">Unit: {unit}</div>' if unit else ""
    st.markdown(
        f"""
        <div class="pic-kpi-card">
            <div class="pic-kpi-label">{label}</div>
            {unit_html}
            <div class="pic-kpi-value">{value}</div>
            <div class="pic-kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def filtered_monthly_metric(
    df: pd.DataFrame,
    value_col: str,
    agg: str = "sum",
) -> float:
    """
    KPI helper:
    - Single month => exact selected-month aggregate.
    - All months => average of monthly aggregates.
    """
    if df is None or df.empty or value_col not in df.columns:
        return float("nan")

    d = df[["MonthDate", value_col]].copy()
    d[value_col] = pd.to_numeric(d[value_col], errors="coerce")
    d = d.dropna(subset=["MonthDate", value_col])
    if d.empty:
        return float("nan")

    if agg == "sum":
        monthly = d.groupby("MonthDate")[value_col].sum(min_count=1)
    elif agg == "mean":
        monthly = d.groupby("MonthDate")[value_col].mean()
    else:
        raise ValueError("agg must be 'sum' or 'mean'")

    monthly = monthly.dropna()
    if monthly.empty:
        return float("nan")
    return float(monthly.mean())


def hc_capacity_utilization(df: pd.DataFrame) -> float:
    """
    Capacity Utilization KPI source of truth:
    sheet HC -> column "Capacity Utilization (%)".

    Single Office / Month:
        use the source value directly.

    All Offices:
        calculate each month's overall utilization as the Actual-HC-weighted
        average of the office source percentages, then average across selected months.

    Blank future months are excluded.
    """
    if df is None or df.empty or "HC Utilization" not in df.columns:
        return float("nan")

    cols = ["MonthDate", "HC Utilization"]
    if "Total Actual HC" in df.columns:
        cols.append("Total Actual HC")

    d = df[cols].copy()
    d["HC Utilization"] = pd.to_numeric(d["HC Utilization"], errors="coerce")
    if "Total Actual HC" in d.columns:
        d["Total Actual HC"] = pd.to_numeric(d["Total Actual HC"], errors="coerce")

    d = d.dropna(subset=["MonthDate", "HC Utilization"])
    if d.empty:
        return float("nan")

    monthly_values = []

    for _, g in d.groupby("MonthDate"):
        # If only one row/office for the month, this is the exact source value.
        if len(g) == 1:
            monthly_values.append(float(g["HC Utilization"].iloc[0]))
            continue

        # All Offices: weight by Actual HC so larger offices contribute appropriately.
        if "Total Actual HC" in g.columns:
            valid = g["Total Actual HC"].notna() & (g["Total Actual HC"] > 0)
            if valid.any():
                weighted = (
                    g.loc[valid, "HC Utilization"]
                    * g.loc[valid, "Total Actual HC"]
                ).sum() / g.loc[valid, "Total Actual HC"].sum()
                monthly_values.append(float(weighted))
                continue

        # Fallback only if HC weights are unavailable.
        monthly_values.append(float(g["HC Utilization"].mean()))

    monthly_values = [v for v in monthly_values if not pd.isna(v)]
    return float(np.mean(monthly_values)) if monthly_values else float("nan")



def pic_utilization_card(util: float):
    if pd.isna(util):
        value = "N/A"
        pct = 0
        color = COLORS["muted"]
    else:
        value = fmt_pct(util)
        pct = max(0, min(util * 100, 125))
        _, color, _ = status_from_util(util)

    width_pct = min(pct / 125 * 100, 100)
    st.markdown(
        f"""
        <div class="pic-status-card">
            <div class="pic-status-left">
                <div class="pic-status-title">CAPACITY UTILIZATION</div>
                <div class="pic-status-value">{value}</div>
            </div>
            <div class="pic-progress-track">
                <div class="pic-progress-fill"
                     style="width:{width_pct:.1f}%;background:{color};"></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def overall_workload_status_card(util: float):
    if pd.isna(util):
        status_text, color, bg = "NO DATA", COLORS["muted"], COLORS["light_blue"]
    else:
        status_text, color, bg = status_from_util(util)

    st.markdown(
        f"""
        <div class="workload-status-panel">
            <div class="pic-status-title">OVERALL WORKLOAD STATUS</div>
            <div class="workload-status-text"
                 style="color:{color};background:{bg};
                        border-radius:999px;padding:8px 14px;">
                {status_text}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_title(text: str):
    st.markdown(f'<div class="section-title">{text}</div>', unsafe_allow_html=True)


def pair_panel_title(text: str):
    """Shared title style for chart/table pairs so both columns align visually."""
    st.markdown(
        f"""
        <div style="
            color:{COLORS['navy']};
            font-family:{UI['font_family']};
            font-size:{UI['chart_title_size']}px;
            line-height:1.25;
            font-weight:700;
            min-height:28px;
            display:flex;
            align-items:center;
            margin:0 0 8px 2px;">
            {text}
        </div>
        """,
        unsafe_allow_html=True,
    )


def plotly_layout(
    fig: go.Figure,
    height: int = UI["chart_height"],
    *,
    show_legend: bool = True,
    legend_position: str = "top",
    margin_left: int = 52,
    margin_right: int = 36,
    margin_top: int = 62,
    margin_bottom: int = 44,
) -> go.Figure:
    """Shared Executive/Corporate Plotly layout — UI only."""
    legend_cfg = dict(
        font=dict(size=UI["axis_size"]),
        bgcolor="rgba(0,0,0,0)",
        borderwidth=0,
    )

    if legend_position == "top":
        legend_cfg.update(
            orientation="h",
            yanchor="bottom",
            y=1.015,
            xanchor="right",
            x=1,
        )
    elif legend_position == "bottom":
        legend_cfg.update(
            orientation="h",
            yanchor="top",
            y=-0.16,
            xanchor="left",
            x=0,
        )
    else:
        legend_cfg.update(
            orientation="v",
            yanchor="top",
            y=1,
            xanchor="left",
            x=1.02,
        )

    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(
            color=COLORS["text"],
            family=UI["font_family"],
            size=UI["axis_size"],
        ),
        # Preserve an existing Plotly title when one is explicitly set.
        # If the chart uses an external pair_panel_title(), keep Plotly title text blank
        # to prevent some Plotly/Streamlit versions from rendering "undefined".
        title=dict(
            text=(fig.layout.title.text or "") if getattr(fig.layout, "title", None) else "",
            font=dict(
                size=UI["chart_title_size"],
                color=COLORS["navy"],
                family=UI["font_family"],
            ),
            x=0.0,
            xanchor="left",
            y=0.985,
            yanchor="top",
            pad=dict(t=0, b=8),
        ),
        margin=dict(
            l=margin_left,
            r=margin_right,
            t=margin_top,
            b=margin_bottom,
        ),
        legend=legend_cfg,
        showlegend=show_legend,
        hoverlabel=dict(
            font=dict(family=UI["font_family"], size=UI["axis_size"]),
            bgcolor="#FFFFFF",
            bordercolor=COLORS["border"],
        ),
        hovermode="closest",
    )

    fig.update_xaxes(
        gridcolor=COLORS["grid"],
        gridwidth=0.7,
        zeroline=False,
        showline=False,
        tickfont=dict(size=UI["axis_size"]),
        title_font=dict(size=UI["axis_size"], color=COLORS["gray_dark"]),
        automargin=True,
        ticks="",
    )
    fig.update_yaxes(
        gridcolor=COLORS["grid"],
        gridwidth=0.7,
        zeroline=False,
        showline=False,
        tickfont=dict(size=UI["axis_size"]),
        title_font=dict(size=UI["axis_size"], color=COLORS["gray_dark"]),
        automargin=True,
        ticks="",
    )
    return fig

# ============================================================
# DATA LOADING
# ============================================================

@st.cache_data(show_spinner=False)
def load_data(path: str, cache_token: str = "") -> Dict[str, pd.DataFrame]:
    """
    FAST workbook loader:
    - Open Excel workbook only ONCE with pd.ExcelFile.
    - Parse required sheets from the same workbook handle.
    - cache_token invalidates Streamlit cache when file content changes.
    """
    data: Dict[str, pd.DataFrame] = {}

    try:
        xls = pd.ExcelFile(path, engine="openpyxl")
        available_sheets = set(xls.sheet_names)
    except Exception as exc:
        logger.exception("Unable to open workbook %s", path)
        raise WorkbookLoadError(
            "Không thể mở file Excel đã chọn. Vui lòng kiểm tra định dạng "
            "file, mật khẩu bảo vệ và cấu trúc workbook."
        ) from exc

    for key, sheet in SHEET_NAMES.items():
        if sheet not in available_sheets:
            data[key] = pd.DataFrame()
            continue

        try:
            df = pd.read_excel(
                xls,
                sheet_name=sheet,
                header=1,
            )
            df.columns = [clean_col(c) for c in df.columns]
            df = df.dropna(how="all")
            data[key] = df
        except Exception as exc:
            logger.exception("Unable to read sheet %r from workbook %s", sheet, path)
            raise WorkbookLoadError(
                f"Không thể đọc sheet '{sheet}'. Vui lòng kiểm tra dữ liệu "
                "hoặc định dạng của sheet này."
            ) from exc

    return data



@st.cache_data(show_spinner=False)
def prepare_hc(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Office", "MonthDate"])
    df = df.copy()
    office_col = first_existing(df, ["Office", "OFFICE"])
    month_col = first_existing(df, ["Month"])
    if not office_col or not month_col:
        return pd.DataFrame(columns=["Office", "MonthDate"])
    df["Office"] = df[office_col].map(normalize_office)
    df["MonthDate"] = df[month_col].map(parse_month)
    mapping = {
        "Approved HC MNG": ["Approved HC – MNG", "Approved HC (MNG)"],
        "Approved HC PIC": ["Approved HC – PIC", "Approved HC (PIC)"],
        "Total Approved HC": ["Total Approved HC"],
        "Actual HC MNG": ["Actual HC – MNG", "Actual HC (MNG)"],
        "Actual HC PIC": ["Actual HC – PIC", "Actual HC (PIC)"],
        "Total Actual HC": ["Total Actual HC", "Total Actual  HC"],
        "Required HC MNG": ["Required HC – MNG", "Required HC (MNG)"],
        "Required HC PIC": ["Required HC – PIC", "Required HC (PIC)"],
        "Total Required HC": ["Total Required HC"],
        "HC Available Hours": [
            "Total Available Standard Time (95%x8x22xPIC)",
            "Total Available Time (95%x8x22x total PIC) (i)",
        ],
        "HC Actual Working Hours": [
            "Total actual Working Time (=C+A+S+E)",
            "Total actual Working Time (=C+A+S+E) (ii)",
        ],
        "HC Actual Workload per PIC": ["Actual workload/PIC (hour)"],
        "HC Utilization": [
            "Capacity Utilization (%)",
            "HC Utilization (%)",
            "Office Workload (%) (ii /i)",
        ],
        "HC Status": [
            "Overal Workload Status",
            "Overal  Workload Status",
            "Overall Workload Status",
            "Office Workload Status",
            "HC Status",
        ],
    }
    for new, candidates in mapping.items():
        col = first_existing(df, candidates)
        if col:
            df[new] = df[col]
    # Fallback calculations
    if "Total Approved HC" not in df.columns:
        df["Total Approved HC"] = numeric_series(df.get("Approved HC MNG", 0)) + numeric_series(df.get("Approved HC PIC", 0))
    if "Total Actual HC" not in df.columns:
        df["Total Actual HC"] = numeric_series(df.get("Actual HC MNG", 0)) + numeric_series(df.get("Actual HC PIC", 0))
    hc_numeric_cols = [
        "Approved HC MNG", "Approved HC PIC", "Total Approved HC",
        "Actual HC MNG", "Actual HC PIC", "Total Actual HC",
        "Required HC MNG", "Required HC PIC", "Total Required HC",
        "HC Available Hours", "HC Actual Working Hours",
        "HC Actual Workload per PIC", "HC Utilization"
    ]
    for col in hc_numeric_cols:
        if col in df.columns:
            # IMPORTANT: keep blank Excel cells as NaN.
            # Do not convert future blank months to 0 because that distorts averages/trends.
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=["MonthDate"])

    # Keep only HC rows that actually contain HC data.
    hc_key_cols = [
        c for c in [
            "Total Approved HC", "Total Actual HC", "Total Required HC",
            "Approved HC MNG", "Approved HC PIC",
            "Actual HC MNG", "Actual HC PIC",
            "Required HC MNG", "Required HC PIC"
        ] if c in df.columns
    ]
    if hc_key_cols:
        df = df.dropna(subset=hc_key_cols, how="all")

    return df


@st.cache_data(show_spinner=False)
def prepare_workload(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Office", "MonthDate", "Segment"])

    df = df.copy()
    office_col = first_existing(df, ["Office", "OFFICE"])
    month_col = first_existing(df, ["Month"])
    segment_col = first_existing(df, ["Segment"])

    if not office_col or not month_col or not segment_col:
        return pd.DataFrame(columns=["Office", "MonthDate", "Segment"])

    df["Office"] = df[office_col].map(normalize_office)
    df["MonthDate"] = df[month_col].map(parse_month)
    df["Segment"] = df[segment_col].astype(str).str.strip().str.upper()

    # Canonical dashboard fields remain unchanged.
    # Only source-header aliases are expanded for the renamed MASTER DATA SOURCE workbook.
    component_map = {
        "Core Workload (min)": [
            "Core Workload (min)",
            "C Total Time (min)",
        ],
        "Ancillary Workload (min)": [
            "Ancillary Workload (min)",
            "A Total Time (min)",
        ],
        "Supporting Workload (min)": [
            "Supporting Workload (min)",
            "S Total time (min)",
            "S Total Time (min)",
        ],
        "Exception Workload (min)": [
            "Exception Workload (min)",
            "E Total Time (min)",
        ],
        "Total Workload (min)": [
            "Total Workload (min)",
            "Total time (min)",
        ],
        "Workload Share": [
            "% of Network",
            "CS Allocation (%)",
        ],
        "Office HC Allocation Ratio": [
            "OFFICE HC ALLOCATION RATIO TO Bus",
            "CS Allocation (FTE)",
        ],
        "Core Volume": [
            "Core Volume",
            "C Volume",
        ],
        "Ancillary Volume": [
            "Ancillary Volume",
            "A Volume",
        ],
        "Supporting Volume": [
            "Supporting Volume",
            "S Volume",
        ],
        "Exception Volume": [
            "Exception Volume",
            "E Volume",
        ],
    }

    for canonical, candidates in component_map.items():
        col = first_existing(df, candidates)
        if col:
            df[canonical] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[canonical] = 0.0

    if df["Total Workload (min)"].sum() == 0:
        df["Total Workload (min)"] = (
            df["Core Workload (min)"]
            + df["Ancillary Workload (min)"]
            + df["Supporting Workload (min)"]
            + df["Exception Workload (min)"]
        )

    df["Workload Hours"] = df["Total Workload (min)"] / 60
    df["Core Hours"] = df["Core Workload (min)"] / 60
    df["Ancillary Hours"] = df["Ancillary Workload (min)"] / 60
    df["Supporting Hours"] = df["Supporting Workload (min)"] / 60
    df["Exception Hours"] = df["Exception Workload (min)"] / 60
    df["Service Label"] = df["Segment"].map(SERVICE_LABELS).fillna(df["Segment"])

    return df.dropna(subset=["MonthDate"])


@st.cache_data(show_spinner=False)
def prepare_fte(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize PIC workload while preserving the new source fields from
    sheet "2. FTE Workload".

    Canonical output:
        Office
        CS PIC
        MonthDate
        Available Time
        Actual Working Time
        Actual FTE
        FTE Workload Status

    New MASTER DATA SOURCE:
        Office
        Month
        CS PIC Name
        Total Available Time (95%x8x22x1) (i)
        Total Actual Working Time (=C+A+S+E) (ii)
        FTE Workload (ii /i)
        FTE Workload Status

    Legacy wide-format input remains supported as a fallback.
    """
    base = [
        "Office",
        "CS PIC",
        "MonthDate",
        "Available Time",
        "Actual Working Time",
        "Actual FTE",
        "FTE Workload Status",
    ]

    if df.empty:
        return pd.DataFrame(columns=base)

    df = df.copy()

    office_col = first_existing(df, ["OFFICE", "Office"])
    pic_col = first_existing(df, ["CS PIC", "PIC", "CS PIC Name"])

    if not office_col or not pic_col:
        return pd.DataFrame(columns=base)

    # --------------------------------------------------------
    # New long-format source
    # --------------------------------------------------------
    month_col = first_existing(df, ["Month"])

    # Section 3: Total Available Time is sourced strictly from Excel column D
    # of sheet "2. FTE Workload". In pandas, column D has zero-based index 3.
    available_col = df.columns[3] if len(df.columns) > 3 else None

    actual_time_col = first_existing(
        df,
        [
            "Total Actual Working Time (=C+A+S+E) (ii)",
            "Total Actual Working Time",
            "Actual Working Time",
        ],
    )

    factor_col = first_existing(
        df,
        [
            "FTE Workload (ii /i)",
            "FTE Workload (ii / i)",
            "FTE Workload",
            "FTE Workload",
        ],
    )

    status_col = first_existing(
        df,
        [
            "FTE Workload Status",
            "Workload Status",
        ],
    )

    if month_col and (factor_col or actual_time_col):
        keep_cols = [office_col, month_col, pic_col]
        for c in [available_col, actual_time_col, factor_col, status_col]:
            if c and c not in keep_cols:
                keep_cols.append(c)

        long = df[keep_cols].copy()

        long["Office"] = long[office_col].map(normalize_office)
        long["CS PIC"] = long[pic_col].astype(str).str.strip()
        long["MonthDate"] = long[month_col].map(parse_month)

        # Use Total Available Time directly from source column D.
        if available_col:
            long["Available Time"] = pd.to_numeric(
                long[available_col], errors="coerce"
            )
        else:
            long["Available Time"] = np.nan

        if actual_time_col:
            long["Actual Working Time"] = pd.to_numeric(
                long[actual_time_col], errors="coerce"
            )
        else:
            long["Actual Working Time"] = np.nan

        if factor_col:
            long["Actual FTE"] = pd.to_numeric(
                long[factor_col], errors="coerce"
            )
        else:
            long["Actual FTE"] = np.nan

        # Fallback only when a source field is missing.
        missing_fte = long["Actual FTE"].isna()
        long.loc[missing_fte, "Actual FTE"] = (
            long.loc[missing_fte, "Actual Working Time"]
            / long.loc[missing_fte, "Available Time"].replace(0, np.nan)
        )

        missing_actual = long["Actual Working Time"].isna()
        long.loc[missing_actual, "Actual Working Time"] = (
            long.loc[missing_actual, "Actual FTE"]
            * long.loc[missing_actual, "Available Time"]
        )

        if status_col:
            long["FTE Workload Status"] = (
                long[status_col].astype(str).str.strip()
            )
        else:
            long["FTE Workload Status"] = long["Actual FTE"].apply(
                lambda x: status_from_util(float(x))[0]
                if pd.notna(x) else "NO DATA"
            )

        long = long[
            (long["Office"] != "")
            & (long["CS PIC"] != "")
            & (~long["MonthDate"].isna())
            & (long["Actual FTE"].notna())
        ]

        return long[base].reset_index(drop=True)

    # --------------------------------------------------------
    # Legacy wide-format source
    # --------------------------------------------------------
    month_cols = [
        c for c in df.columns
        if not pd.isna(parse_month(c))
    ]

    if not month_cols:
        return pd.DataFrame(columns=base)

    long = df.melt(
        id_vars=[office_col, pic_col],
        value_vars=month_cols,
        var_name="Month",
        value_name="Actual FTE",
    )

    long["Office"] = long[office_col].map(normalize_office)
    long["CS PIC"] = long[pic_col].astype(str).str.strip()
    long["MonthDate"] = long["Month"].map(parse_month)
    long["Actual FTE"] = pd.to_numeric(long["Actual FTE"], errors="coerce")

    long["Available Time"] = CAPACITY_HOURS_PER_FTE
    long["Actual Working Time"] = (
        long["Actual FTE"] * long["Available Time"]
    )
    long["FTE Workload Status"] = long["Actual FTE"].apply(
        lambda x: status_from_util(float(x))[0]
        if pd.notna(x) else "NO DATA"
    )

    long = long[
        (long["Office"] != "")
        & (~long["MonthDate"].isna())
        & (long["Actual FTE"].notna())
    ]

    return long[base].reset_index(drop=True)



@st.cache_data(show_spinner=False)
def prepare_shipment(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if df.empty:
        return (
            pd.DataFrame(columns=["Office", "MonthDate", "Total Shipment", "Active Customers"]),
            pd.DataFrame(columns=["Office", "MonthDate", "Mode", "Volume"]),
        )

    df = df.copy()
    office_col = first_existing(df, ["Office", "OFFICE"])
    month_col = first_existing(df, ["Month"])
    active_col = first_existing(
        df,
        [
            "Active Customers",
            "Total No. of Active Customers",
        ],
    )
    total_col = first_existing(
        df,
        [
            "TOTAL",
            "Total",
            "Total No. of shipment",
            "Total No. of Shipment",
        ],
    )

    if not office_col or not month_col:
        return pd.DataFrame(), pd.DataFrame()

    df["Office"] = df[office_col].map(normalize_office)
    df["MonthDate"] = df[month_col].map(parse_month)

    df["Total Shipment"] = (
        pd.to_numeric(df[total_col], errors="coerce")
        if total_col else np.nan
    )
    df["Active Customers"] = (
        pd.to_numeric(df[active_col], errors="coerce")
        if active_col else np.nan
    )

    df = df.dropna(subset=["MonthDate"])
    df = df.dropna(subset=["Total Shipment", "Active Customers"], how="all")
    df["Total Shipment"] = df["Total Shipment"].fillna(0)
    df["Active Customers"] = df["Active Customers"].fillna(0)

    # Legacy workbook contained mode columns directly.
    # New master summary sheet contains only Active Customers + Total Shipment.
    excluded = {
        office_col, month_col, active_col, total_col,
        "Office", "MonthDate", "Active Customers", "Total Shipment"
    }
    mode_cols = [
        c for c in df.columns
        if c not in excluded and not str(c).startswith("Unnamed")
    ]

    if mode_cols:
        mode_long = df.melt(
            id_vars=["Office", "MonthDate"],
            value_vars=mode_cols,
            var_name="Mode",
            value_name="Volume",
        )
        mode_long["Volume"] = pd.to_numeric(mode_long["Volume"], errors="coerce").fillna(0)
        mode_long = mode_long[mode_long["Volume"] > 0]
    else:
        mode_long = pd.DataFrame(columns=["Office", "MonthDate", "Mode", "Volume"])

    return df, mode_long


@st.cache_data(show_spinner=False)
def prepare_customer(data: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    # Prefer office-specific customer sheets when available; otherwise use 11. Vol. by Customer.
    office_sheets = ["customer_had", "customer_han", "customer_hlc", "customer_hcm"]
    frames = []
    for key in office_sheets:
        df = data.get(key, pd.DataFrame())
        if df.empty:
            continue
        frames.append(customer_wide_to_long(df))
    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    if combined.empty:
        combined = customer_wide_to_long(data.get("customer_ns", pd.DataFrame()))
    return combined


@st.cache_data(show_spinner=False)
def customer_wide_to_long(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Office", "Customer", "MonthDate", "Volume"])
    df = df.copy()
    office_col = first_existing(df, ["Office", "OFFICE"])
    cust_col = first_existing(df, ["Customer", "CUSTOMER", "Customer Name"])
    if not office_col or not cust_col:
        return pd.DataFrame(columns=["Office", "Customer", "MonthDate", "Volume"])
    month_cols = [c for c in df.columns if parse_month(c) is not pd.NaT and not pd.isna(parse_month(c))]
    if not month_cols:
        return pd.DataFrame(columns=["Office", "Customer", "MonthDate", "Volume"])
    long = df.melt(id_vars=[office_col, cust_col], value_vars=month_cols, var_name="Month", value_name="Volume")
    long["Office"] = long[office_col].map(normalize_office)
    long["Customer"] = long[cust_col].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    long["MonthDate"] = long["Month"].map(parse_month)
    long["Volume"] = numeric_series(long["Volume"])
    long = long[(long["Volume"] > 0) & (long["Customer"] != "") & (~long["MonthDate"].isna())]
    return long[["Office", "Customer", "MonthDate", "Volume"]]


@st.cache_data(show_spinner=False)
def prepare_case_detail(
    df: pd.DataFrame,
    activity_type: str,
) -> pd.DataFrame:
    """
    Normalize detail sheets C / A / S / E into long format.

    C/A/S expected pattern:
        Office | Scope/Code | Apr-26 ... Mar-27 | Total

    E expected pattern:
        Office | Code | BU | Criteria | Exception Detail | Apr-26 ... Mar-27 | Total

    The parser is intentionally tolerant to header wording so the dashboard
    remains usable when the source template adds descriptive columns.
    """
    base_cols = [
        "Activity Type", "Office", "Code", "BU", "Criteria",
        "Detail", "MonthDate", "Volume"
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=base_cols)

    d = df.copy()
    d.columns = [clean_col(c) for c in d.columns]

    office_col = first_existing(d, ["Office", "OFFICE"])
    if not office_col:
        return pd.DataFrame(columns=base_cols)

    # Identify month columns strictly by parseable month headers.
    month_cols = []
    for c in d.columns:
        parsed = parse_month(c)
        if not pd.isna(parsed):
            month_cols.append(c)

    if not month_cols:
        return pd.DataFrame(columns=base_cols)

    # Descriptive columns before month columns.
    code_col = first_existing(d, ["Scope", "CODE", "Code", "Service Code"])
    bu_col = first_existing(d, ["BU", "Segment", "Service"])
    criteria_col = first_existing(
        d,
        ["Criteria", "CRITERIA", "Criterion", "Criteria (N/M/S)"],
    )
    detail_col = first_existing(
        d,
        [
            "Scope details",       # Sheet A
            "Job details",         # Sheet S
            "EXCEPTION DETAIL",    # Sheet E
            "Exception Detail",
            "Detail",
            "Description",
            "Activity",
        ],
    )

    id_cols = [office_col]
    for c in [code_col, bu_col, criteria_col, detail_col]:
        if c and c not in id_cols:
            id_cols.append(c)

    long = d.melt(
        id_vars=id_cols,
        value_vars=month_cols,
        var_name="Month",
        value_name="Volume",
    )

    long["Office"] = long[office_col].map(normalize_office)
    long["MonthDate"] = long["Month"].map(parse_month)
    long["Volume"] = pd.to_numeric(long["Volume"], errors="coerce")

    long["Code"] = (
        long[code_col].astype(str).str.strip()
        if code_col else ""
    )
    long["BU"] = (
        long[bu_col].astype(str).str.strip().str.upper()
        if bu_col else ""
    )
    long["Criteria"] = (
        long[criteria_col].astype(str).str.strip()
        if criteria_col else ""
    )
    long["Detail"] = (
        long[detail_col].astype(str).str.strip()
        if detail_col else ""
    )
    long["Activity Type"] = activity_type

    # Keep only rows that contain real activity data.
    # Blank cells and 0-volume months are excluded so detail tabs show only months
    # that actually have C/A/S/E data in the corresponding source sheet.
    long = long[
        (long["Office"] != "")
        & (~long["MonthDate"].isna())
        & (long["Volume"].notna())
        & (long["Volume"] > 0)
    ].copy()

    return long[base_cols].reset_index(drop=True)



@st.cache_data(show_spinner=False)
def prepare_code_note_map(df: pd.DataFrame) -> Dict[str, str]:
    """
    Sheet 'Ghi chú':
        Col A = Scope of Job code
        Col B = description.

    Used mainly for Core Service codes such as AE-CTAB:
        suffix CTAB -> "Customs + Trucking + Air"

    A/S/E sheets already contain their own descriptive columns
    (Scope details / Job details / EXCEPTION DETAIL), which take priority.
    """
    if df is None or df.empty or df.shape[1] < 2:
        return {}

    d = df.copy()
    code_col = d.columns[0]
    desc_col = d.columns[1]

    d[code_col] = d[code_col].astype(str).str.strip().str.upper()
    d[desc_col] = d[desc_col].astype(str).str.strip()

    # Remove title/header/blank rows.
    d = d[
        d[code_col].ne("")
        & d[desc_col].ne("")
        & d[code_col].ne("SCOPE OF JOB")
        & d[code_col].ne("NAN")
        & d[desc_col].ne("nan")
    ].copy()

    return dict(zip(d[code_col], d[desc_col]))


def add_code_description(
    df: pd.DataFrame,
    activity_type: str,
    note_map: Dict[str, str],
) -> pd.DataFrame:
    """
    Create the dashboard field 'Code Description'.

    Priority:
    1) A/S/E: source description already available in the corresponding sheet.
    2) C: lookup suffix of Scope code against sheet 'Ghi chú'
       e.g. AE-ABBB -> ABBB -> Air Freight Only.
    3) Fallback: lookup whole code in 'Ghi chú'.
    """
    if df is None or df.empty:
        return df

    d = df.copy()
    d["Code"] = d["Code"].fillna("").astype(str).str.strip()
    d["Detail"] = d["Detail"].fillna("").astype(str).str.strip()

    def _decode(row):
        code = str(row.get("Code", "")).strip().upper()
        source_detail = str(row.get("Detail", "")).strip()

        # Ancillary / Supporting / Exception: use source wording first.
        if activity_type != "Core Service" and source_detail:
            return source_detail

        # Core: AE-CTAB -> CTAB
        suffix = code.split("-", 1)[1] if "-" in code else code
        if suffix in note_map:
            return note_map[suffix]

        if code in note_map:
            return note_map[code]

        # Defensive fallback if a source detail exists.
        return source_detail

    d["Code Description"] = d.apply(_decode, axis=1)
    return d


def workload_breakdown_table(df: pd.DataFrame) -> pd.DataFrame:
    """
    Section 5 summary:
    Segment | Core Service (min) | Ancillary Service (min) |
    Supporting Activity (min) | Exception Handling (min) |
    Total Workload (min) | Ratio
    """
    cols = [
        "Segment",
        "Core Service (min)",
        "Ancillary Service (min)",
        "Supporting Activity (min)",
        "Exception Handling (min)",
        "Total Workload (min)",
        "Ratio",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)

    d = df.copy()
    for c in [
        "Core Workload (min)",
        "Ancillary Workload (min)",
        "Supporting Workload (min)",
        "Exception Workload (min)",
        "Total Workload (min)",
    ]:
        if c not in d.columns:
            d[c] = 0.0
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0)

    agg = (
        d.groupby("Segment", as_index=False)
        .agg({
            "Core Workload (min)": "sum",
            "Ancillary Workload (min)": "sum",
            "Supporting Workload (min)": "sum",
            "Exception Workload (min)": "sum",
            "Total Workload (min)": "sum",
        })
    )

    # Ensure standard business order.
    agg = (
        pd.DataFrame({"Segment": SERVICE_ORDER})
        .merge(agg, on="Segment", how="left")
        .fillna(0)
    )

    # Recalculate Total from C+A+S+E if source Total is blank/zero.
    component_total = (
        agg["Core Workload (min)"]
        + agg["Ancillary Workload (min)"]
        + agg["Supporting Workload (min)"]
        + agg["Exception Workload (min)"]
    )
    agg["Total Workload (min)"] = np.where(
        agg["Total Workload (min)"] > 0,
        agg["Total Workload (min)"],
        component_total,
    )

    grand_total = float(agg["Total Workload (min)"].sum())
    agg["Ratio"] = np.where(
        grand_total > 0,
        agg["Total Workload (min)"] / grand_total,
        0.0,
    )

    agg = agg.rename(columns={
        "Core Workload (min)": "Core Service (min)",
        "Ancillary Workload (min)": "Ancillary Service (min)",
        "Supporting Workload (min)": "Supporting Activity (min)",
        "Exception Workload (min)": "Exception Handling (min)",
    })

    return agg[cols]



def render_case_office_cards(workload_df: pd.DataFrame):
    """
    Section 5 executive cards by Office.

    IMPORTANT:
    - Cards use sheet "4. Workload by Activity" as the single source of truth.
    - C / A / S / E values are summed from the corresponding activity-volume
      columns by Office after the dashboard Month / Office filters are applied.
    - HPH is displayed as HLC to follow the dashboard's standard office naming.
    """
    if workload_df is None or workload_df.empty or "Office" not in workload_df.columns:
        return

    source_map = {
        "C": "Core Volume",
        "A": "Ancillary Volume",
        "S": "Supporting Volume",
        "E": "Exception Volume",
    }
    d = workload_df[["Office"]].copy()
    d["Office"] = workload_df["Office"].astype(str).str.strip().str.upper()
    d["Office"] = d["Office"].replace({"HPH": "HLC"})

    for activity, source_col in source_map.items():
        if source_col in workload_df.columns:
            d[activity] = pd.to_numeric(
                workload_df[source_col], errors="coerce"
            ).fillna(0.0)
        else:
            d[activity] = 0.0

    d = d[d["Office"] != ""].copy()
    if d.empty:
        return

    summary = d.groupby("Office", as_index=False)[["C", "A", "S", "E"]].sum()

    summary["Total"] = summary[["C", "A", "S", "E"]].sum(axis=1)

    present = summary["Office"].astype(str).tolist()
    offices = [o for o in STANDARD_OFFICES if o in present]
    offices += sorted([o for o in present if o not in offices])

    if not offices:
        return

    st.markdown(
        f"""
        <div style="
            color:{COLORS['navy']};
            font-size:20px;
            font-weight:700;
            margin:4px 0 10px 2px;">
            C / A / S / E Activity by Office
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Keep Office cards compact when filters return fewer than 4 offices.
    # All Offices: 4 cards fill the row as before.
    # Single Office: card uses only 1/4 row width instead of stretching full width.
    slot_count = max(4, len(offices))
    card_cols = st.columns(slot_count, gap="medium")[:len(offices)]

    activity_meta = {
        "C": ("Core", COLORS["blue"]),
        "A": ("Ancillary", COLORS["green"]),
        "S": ("Supporting", COLORS["amber"]),
        "E": ("Exception", COLORS["red"]),
    }

    for card_col, office in zip(card_cols, offices):
        row = summary.loc[summary["Office"] == office].iloc[0]

        vals = {
            activity: float(pd.to_numeric(row.get(activity, 0), errors="coerce") or 0)
            for activity in ["C", "A", "S", "E"]
        }
        total = float(sum(vals.values()))
        shares = {
            activity: safe_div(value, total)
            for activity, value in vals.items()
        }

        with card_col:
            card_html = f"""
                <div style="
                    background:#FFFFFF;
                    border:1px solid {COLORS['border']};
                    border-top:4px solid {COLORS['navy']};
                    border-radius:12px;
                    padding:14px 16px 13px;
                    min-height:180px;
                    box-sizing:border-box;
                    box-shadow:0 2px 7px rgba(0,59,112,0.045);">

                  <div style="
                      display:flex;
                      justify-content:space-between;
                      align-items:flex-start;
                      gap:10px;
                      margin-bottom:10px;">
                    <div style="
                        color:{COLORS['navy']};
                        font-size:20px;
                        font-weight:800;">
                      {html.escape(office)}
                    </div>

                    <div style="text-align:right;">
                      <div style="
                          color:#667085;
                          font-size:12px;
                          font-weight:600;">
                        TOTAL ACTIVITY
                      </div>
                      <div style="
                          color:{COLORS['navy']};
                          font-size:23px;
                          font-weight:800;
                          margin-top:2px;">
                        {total:,.0f}
                      </div>
                    </div>
                  </div>

                  <div style="
                      display:grid;
                      grid-template-columns:repeat(4,minmax(0,1fr));
                      border-top:1px solid #E7ECF1;
                      padding-top:11px;">
                    <div style="text-align:center;border-right:1px solid #E7ECF1;">
                      <div style="color:{activity_meta['C'][1]};font-size:16px;font-weight:800;">C</div>
                      <div style="color:{COLORS['navy']};font-size:18px;font-weight:750;margin-top:4px;">{vals['C']:,.0f}</div>
                    </div>
                    <div style="text-align:center;border-right:1px solid #E7ECF1;">
                      <div style="color:{activity_meta['A'][1]};font-size:16px;font-weight:800;">A</div>
                      <div style="color:{COLORS['navy']};font-size:18px;font-weight:750;margin-top:4px;">{vals['A']:,.0f}</div>
                    </div>
                    <div style="text-align:center;border-right:1px solid #E7ECF1;">
                      <div style="color:{activity_meta['S'][1]};font-size:16px;font-weight:800;">S</div>
                      <div style="color:{COLORS['navy']};font-size:18px;font-weight:750;margin-top:4px;">{vals['S']:,.0f}</div>
                    </div>
                    <div style="text-align:center;">
                      <div style="color:{activity_meta['E'][1]};font-size:16px;font-weight:800;">E</div>
                      <div style="color:{COLORS['navy']};font-size:18px;font-weight:750;margin-top:4px;">{vals['E']:,.0f}</div>
                    </div>
                  </div>

                  <div style="
                      display:grid;
                      grid-template-columns:repeat(4,minmax(0,1fr));
                      margin-top:4px;
                      color:#667085;
                      font-size:16px;
                      text-align:center;">
                    <div>{shares['C']:.1%}</div>
                    <div>{shares['A']:.1%}</div>
                    <div>{shares['S']:.1%}</div>
                    <div>{shares['E']:.1%}</div>
                  </div>
                </div>
                """
            card_html = "\n".join(line.lstrip() for line in card_html.splitlines())
            st.markdown(card_html, unsafe_allow_html=True)



def chart_case_allocation(df: pd.DataFrame):
    """C/A/S/E workload composition by Segment, displayed in hours."""
    summary = workload_breakdown_table(df)
    if summary.empty or float(summary["Total Workload (min)"].sum()) <= 0:
        st.info("No C/A/S/E workload data available for selected filters.")
        return

    pair_panel_title("Workload Composition by Activity")

    plot_df = (
        summary[summary["Total Workload (min)"] > 0]
        .copy()
        .sort_values("Total Workload (min)", ascending=True)
    )

    # Display layer only: convert minutes to hours.
    for _col in [
        "Core Service (min)",
        "Ancillary Service (min)",
        "Supporting Activity (min)",
        "Exception Handling (min)",
        "Total Workload (min)",
    ]:
        plot_df[_col] = pd.to_numeric(
            plot_df[_col], errors="coerce"
        ).fillna(0) / 60

    components = [
        ("Core Service (min)", "Core Service", COLORS["blue"]),
        ("Ancillary Service (min)", "Ancillary Service", COLORS["green"]),
        ("Supporting Activity (min)", "Supporting Activity", COLORS["amber"]),
        ("Exception Handling (min)", "Exception Handling", COLORS["navy"]),
    ]

    fig = go.Figure()

    for col, label, color in components:
        fig.add_trace(
            go.Bar(
                y=plot_df["Segment"],
                x=plot_df[col],
                name=label,
                orientation="h",
                marker_color=color,
                customdata=np.column_stack(
                    [
                        plot_df["Total Workload (min)"],
                        plot_df["Ratio"],
                    ]
                ),
                hovertemplate=(
                    f"<b>{label}</b>"
                    "<br>Segment: %{y}"
                    "<br>Workload: %{x:,.1f} hrs"
                    "<br>Segment Total: %{customdata[0]:,.1f} hrs"
                    "<br>Share of Total: %{customdata[1]:.1%}"
                    "<extra></extra>"
                ),
            )
        )

    fig.update_layout(
        barmode="stack",
        title=dict(text=""),
        xaxis_title="Workload (Hours)",
        yaxis_title="",
    )

    fig = plotly_layout(
        fig,
        350,
        show_legend=True,
        legend_position="top",
        margin_left=50,
        margin_right=35,
        margin_top=56,
        margin_bottom=40,
    )

    # Follow the same left-to-right order as the stacked bars and keep the
    # legend visually separated from the top bar.
    fig.update_layout(
        legend=dict(
            traceorder="normal",
            orientation="h",
            x=0.5,
            xanchor="center",
            y=1.10,
            yanchor="bottom",
        )
    )

    fig.update_xaxes(rangemode="tozero")

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
    )


def exception_criteria_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate filtered Exception Handling volume and share by Criteria."""
    if df is None or df.empty or "Criteria" not in df.columns or "Volume" not in df.columns:
        return pd.DataFrame(columns=["Criteria", "Volume", "Share"])

    d = df[["Criteria", "Volume"]].copy()
    d["Criteria"] = (
        d["Criteria"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )
    d["Volume"] = pd.to_numeric(d["Volume"], errors="coerce").fillna(0.0)
    d = d[
        (~d["Criteria"].str.lower().isin(["", "nan", "none"]))
        & (d["Volume"] > 0)
    ]

    if d.empty:
        return pd.DataFrame(columns=["Criteria", "Volume", "Share"])

    agg = (
        d.groupby("Criteria", as_index=False)["Volume"]
        .sum()
        .sort_values("Volume", ascending=True)
    )
    total = float(agg["Volume"].sum())
    agg["Share"] = np.where(total > 0, agg["Volume"] / total, 0.0)
    return agg


def chart_exception_by_criteria(df: pd.DataFrame):
    """Infographic-style exploded donut of Exception Handling by Criteria."""
    agg = exception_criteria_summary(df)
    if agg.empty:
        st.info("No Exception Handling criteria data available for the selected filters.")
        return

    agg = agg.sort_values("Volume", ascending=False).reset_index(drop=True)
    # Softer Yusen-inspired palette: calm for the dominant slice while smaller
    # slices remain clearly distinguishable.
    criteria_palette = [
        "#8FB7D7",  # Soft blue for the largest slice
        "#E6761B",  # Yusen orange
        "#D9A400",  # Amber
        "#269E9A",  # Teal
        "#6654C7",  # Violet
        "#347DB6",  # Corporate blue
        "#70AD47",  # Green
        "#102B55",  # Navy fallback
    ]
    pie_colors = [
        criteria_palette[i % len(criteria_palette)]
        for i in range(len(agg))
    ]

    # Center the combined arc of all smaller slices around 3 o'clock so their
    # labels appear together on the right. The angle adapts when filters change
    # the largest Criteria share.
    largest_share = float(agg.loc[0, "Share"])
    pie_rotation = 90 + 180 * max(0.0, 1.0 - largest_share)

    text_positions = [
        "inside" if float(share) >= 0.08 else "outside"
        for share in agg["Share"]
    ]
    text_templates = [
        "<b>%{label}</b><br>%{percent:.1%}"
        if float(share) >= 0.08
        else "<b>%{label}</b> %{percent:.1%}"
        for share in agg["Share"]
    ]
    # Subtle staggered pull creates the infographic effect without large gaps.
    pull_pattern = [0.035, 0.018, 0.028, 0.015, 0.025]
    pull_values = [pull_pattern[i % len(pull_pattern)] for i in range(len(agg))]
    total_volume = float(agg["Volume"].sum())

    fig = go.Figure(
        go.Pie(
            labels=agg["Criteria"],
            values=agg["Volume"],
            customdata=agg[["Share"]],
            sort=False,
            direction="clockwise",
            rotation=pie_rotation,
            # Keep space for outside labels while retaining a large donut.
            domain=dict(x=[0.06, 0.88], y=[0.08, 0.94]),
            hole=0.52,
            pull=pull_values,
            textinfo="label+percent",
            texttemplate=text_templates,
            # Match the sample: major labels inside, small labels outside.
            textposition=text_positions,
            insidetextorientation="horizontal",
            automargin=True,
            insidetextfont=dict(
                family=UI["font_family"],
                size=11,
                color=COLORS["navy"],
            ),
            outsidetextfont=dict(
                family=UI["font_family"],
                size=11,
                color=COLORS["navy"],
            ),
            marker=dict(
                colors=pie_colors,
                line=dict(color=COLORS["white"], width=2.5),
            ),
            hovertemplate=(
                "<b>%{label}</b><br>"
                "Exception Volume: %{value:,.0f}<br>"
                "Share: %{customdata[0]:.1%}<extra></extra>"
            ),
            showlegend=False,
        )
    )

    fig = plotly_layout(
        # Match the table panel height (caption + 420 px dataframe).
        fig, 455, show_legend=False,
        margin_left=38, margin_right=78, margin_top=52, margin_bottom=28,
    )
    fig.update_layout(
        title=dict(text="Exception Handling by Criteria"),
        annotations=[
            dict(
                x=0.47,
                y=0.51,
                xref="paper",
                yref="paper",
                text=(
                    "<b>EXCEPTION</b><br>"
                    "<span style='font-size:12px'>HANDLING</span><br>"
                    f"<b>{total_volume:,.0f}</b>"
                ),
                showarrow=False,
                align="center",
                font=dict(
                    family=UI["font_family"],
                    size=16,
                    color=COLORS["navy"],
                ),
            )
        ],
        # Preserve small labels rather than hiding them.
        uniformtext_minsize=8,
        uniformtext_mode="show",
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def render_workload_breakdown_table(df: pd.DataFrame):
    """C/A/S/E workload detail in hours; no TOTAL row and no Ratio column."""
    summary = workload_breakdown_table(df)
    if summary.empty:
        st.info("No workload breakdown data available for selected filters.")
        return

    pair_panel_title("Activity Breakdown")

    display = summary.copy()

    # Display layer only: convert minutes to hours.
    display["Core Service (Hours)"] = pd.to_numeric(
        display["Core Service (min)"], errors="coerce"
    ).fillna(0) / 60
    display["Ancillary Service (Hours)"] = pd.to_numeric(
        display["Ancillary Service (min)"], errors="coerce"
    ).fillna(0) / 60
    display["Supporting Activity (Hours)"] = pd.to_numeric(
        display["Supporting Activity (min)"], errors="coerce"
    ).fillna(0) / 60
    display["Exception Handling (Hours)"] = pd.to_numeric(
        display["Exception Handling (min)"], errors="coerce"
    ).fillna(0) / 60
    display["Total Workload (Hours)"] = pd.to_numeric(
        display["Total Workload (min)"], errors="coerce"
    ).fillna(0) / 60

    display = display[
        [
            "Segment",
            "Core Service (Hours)",
            "Ancillary Service (Hours)",
            "Supporting Activity (Hours)",
            "Exception Handling (Hours)",
            "Total Workload (Hours)",
        ]
    ]

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
        height=350,
        column_config={
            "Segment": st.column_config.TextColumn(
                "Segment",
                width=70,
            ),
            "Core Service (Hours)": st.column_config.NumberColumn(
                "Core Service (Hours)",
                format="%,.1f",
                width=125,
            ),
            "Ancillary Service (Hours)": st.column_config.NumberColumn(
                "Ancillary Service (Hours)",
                format="%,.1f",
                width=140,
            ),
            "Supporting Activity (Hours)": st.column_config.NumberColumn(
                "Supporting Activity (Hours)",
                format="%,.1f",
                width=150,
            ),
            "Exception Handling (Hours)": st.column_config.NumberColumn(
                "Exception Handling (Hours)",
                format="%,.1f",
                width=150,
            ),
            "Total Workload (Hours)": st.column_config.NumberColumn(
                "Total Workload (Hours)",
                format="%,.1f",
                width=135,
            ),
        },
    )



def render_activity_detail_table(
    df: pd.DataFrame,
    activity_type: str,
    table_height: Optional[int] = None,
    show_months_caption: bool = True,
    stretch_to_container: bool = False,
):
    """Detail table for one C/A/S/E source sheet."""
    if df is None or df.empty:
        st.info(f"No {activity_type} detail data available for selected filters.")
        return

    d = df.copy()

    # Defensive filter: detail table only shows months/rows with actual data.
    d["Volume"] = pd.to_numeric(d["Volume"], errors="coerce")
    d = d[d["Volume"].fillna(0) > 0].copy()

    if d.empty:
        st.info(f"No {activity_type} detail data available for selected filters.")
        return

    d["Month"] = d["MonthDate"].dt.strftime("%b-%y")

    # Criteria is carried directly from source sheet "8. E Vol.".
    # It remains visible for Exception Handling and is automatically removed
    # from C/A/S when the entire source column is blank.
    preferred = [
        "Office", "Month", "Code", "Criteria", "Code Description", "Volume"
    ]
    cols = [c for c in preferred if c in d.columns]

    # Drop descriptive columns that are completely blank.
    cols = [
        c for c in cols
        if c in ["Office", "Month", "Volume"]
        or d[c].astype(str).str.strip().replace("nan", "").ne("").any()
    ]

    # Capture months before removing MonthDate from the visible table.
    months_with_data = (
        d["MonthDate"].dropna().drop_duplicates().sort_values()
        .dt.strftime("%b-%y").tolist()
    )

    sort_source = d.copy()
    sort_cols = [
        c for c in ["Office", "Code", "Criteria", "MonthDate"]
        if c in sort_source.columns
    ]
    if sort_cols:
        sort_source = sort_source.sort_values(sort_cols)

    d = sort_source[cols].copy()
    if "Volume" in d.columns:
        d["Volume"] = pd.to_numeric(d["Volume"], errors="coerce").fillna(0)

    if show_months_caption and months_with_data:
        st.caption("Months with data: " + ", ".join(months_with_data))

    if stretch_to_container:
        # Exception table: fill the right panel and prioritize its description.
        auto_fit_config = {
            "Office": st.column_config.TextColumn("Office", width="small"),
            "Month": st.column_config.TextColumn("Month", width="small"),
            "Code": st.column_config.TextColumn("Code", width="small"),
            "Criteria": st.column_config.TextColumn("\u2003Criteria", width=90),
            "Code Description": st.column_config.TextColumn(
                "Code Description", width="large"
            ),
            "Volume": st.column_config.NumberColumn(
                "Volume", format="%,.0f", width="small"
            ),
        }
    else:
        # Core / Ancillary / Supporting tables: size columns from their content.
        auto_fit_config = {
            "Office": st.column_config.TextColumn("Office"),
            "Month": st.column_config.TextColumn("Month"),
            "Code": st.column_config.TextColumn("Code"),
            "Criteria": st.column_config.TextColumn("Criteria"),
            "Code Description": st.column_config.TextColumn("Code Description"),
            "Volume": st.column_config.NumberColumn("Volume", format="%,.0f"),
        }

    # Streamlit's grid ignores CSS text alignment for text columns. Use fixed
    # column width plus non-breaking em-space padding for reliable centering.
    display_table = d
    if "Criteria" in d.columns:
        display_table = d.copy()
        display_table["Criteria"] = (
            display_table["Criteria"]
            .fillna("")
            .astype(str)
            .str.strip()
            .map(lambda value: f"\u2003\u2003{value}" if value else "")
        )

    st.dataframe(
        display_table,
        use_container_width=stretch_to_container,
        hide_index=True,
        height=(
            table_height
            if table_height is not None
            else min(420, max(160, 38 + len(d) * 34))
        ),
        column_config={c: auto_fit_config[c] for c in d.columns if c in auto_fit_config},
    )


def prepare_resolution(df: pd.DataFrame) -> pd.DataFrame:
    """
    Sheet 'CS Resolutions Rate'
    Source structure:
        OFFICE | Month | Total abnormality/month |
        No of abnormality resolved by CS | CS Resolution rate

    Dashboard rules:
    - Only months with actual source data are retained.
    - Resolution Rate is recalculated from Resolved / Total Abnormality
      so the dashboard does not depend on Excel formula cache.
    """
    cols = [
        "Office", "MonthDate",
        "Total Abnormality", "Resolved", "Resolution Rate",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)

    d = df.copy()
    office_col = first_existing(d, ["OFFICE", "Office"])
    month_col = first_existing(d, ["Month"])
    total_col = first_existing(
        d, ["Total abnormality/month", "Total abnormality", "Total Exception Case"]
    )
    resolved_col = first_existing(
        d, ["No of abnormality resolved by CS", "Resolved", "No of Exception Case Resolved by CS"]
    )

    if not office_col or not month_col or not total_col or not resolved_col:
        return pd.DataFrame(columns=cols)

    d["Office"] = d[office_col].map(normalize_office)
    d["MonthDate"] = d[month_col].map(parse_month)

    # Preserve blanks first; do not convert future empty months to zero.
    d["Total Abnormality"] = pd.to_numeric(d[total_col], errors="coerce")
    d["Resolved"] = pd.to_numeric(d[resolved_col], errors="coerce")

    # Keep only rows/months where the source contains actual activity data.
    d = d[
        (d["Office"] != "")
        & (~d["MonthDate"].isna())
        & (
            d["Total Abnormality"].notna()
            | d["Resolved"].notna()
        )
    ].copy()

    if d.empty:
        return pd.DataFrame(columns=cols)

    d["Total Abnormality"] = d["Total Abnormality"].fillna(0)
    d["Resolved"] = d["Resolved"].fillna(0)

    d["Resolution Rate"] = np.where(
        d["Total Abnormality"] > 0,
        d["Resolved"] / d["Total Abnormality"],
        np.nan,
    )

    return d[cols].reset_index(drop=True)


@st.cache_data(show_spinner=False)
def prepare_yvf(df: pd.DataFrame) -> pd.DataFrame:
    """
    Sheet 'YVF' — source structure:
        OFFICE | [Month, if available] |
        Total YVF booking/month | Total IFF shipment/month | YVF booking ratio

    Rules:
    - Preserve the source structure; do not invent months.
    - If a Month column exists, parse it and allow Month/Year filtering.
    - Rows where both YVF Booking and IFF Shipment are blank are excluded,
      so future empty periods never appear on the dashboard.
    - Recalculate YVF Booking Ratio = YVF Booking / IFF Shipment to avoid
      stale Excel formula-cache values.
    """
    base_cols = [
        "Office", "MonthDate",
        "YVF Booking", "IFF Shipment", "YVF Booking Ratio",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=base_cols)

    d = df.copy()

    office_col = first_existing(d, ["OFFICE", "Office"])
    month_col = first_existing(d, ["Month"])
    yvf_col = first_existing(
        d, ["Total YVF booking/month", "Total YVF booking"]
    )
    iff_col = first_existing(
        d, ["Total IFF shipment/month", "Total IFF shipment", "Total IFF Booking"]
    )

    if not office_col or not yvf_col or not iff_col:
        return pd.DataFrame(columns=base_cols)

    d["Office"] = d[office_col].map(normalize_office)
    d["MonthDate"] = (
        d[month_col].map(parse_month)
        if month_col else pd.NaT
    )

    # Preserve blanks first so empty/future periods are not converted to zeros.
    d["YVF Booking"] = pd.to_numeric(d[yvf_col], errors="coerce")
    d["IFF Shipment"] = pd.to_numeric(d[iff_col], errors="coerce")

    d = d[
        (d["Office"] != "")
        & (
            d["YVF Booking"].notna()
            | d["IFF Shipment"].notna()
        )
    ].copy()

    if d.empty:
        return pd.DataFrame(columns=base_cols)

    d["YVF Booking"] = d["YVF Booking"].fillna(0)
    d["IFF Shipment"] = d["IFF Shipment"].fillna(0)

    d["YVF Booking Ratio"] = np.where(
        d["IFF Shipment"] > 0,
        d["YVF Booking"] / d["IFF Shipment"],
        np.nan,
    )

    return d[base_cols].reset_index(drop=True)



def all_periods(*dfs: pd.DataFrame) -> List[pd.Timestamp]:
    periods = []
    for df in dfs:
        if df is not None and not df.empty and "MonthDate" in df.columns:
            periods.extend(df["MonthDate"].dropna().unique().tolist())
    return sorted(pd.to_datetime(pd.Series(periods)).dropna().drop_duplicates().tolist())


def apply_filters(df: pd.DataFrame, year: str, month: str, office: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    if "MonthDate" in out.columns:
        if year != "All":
            out = out[out["MonthDate"].dt.year.astype(str) == year]
        if month != "All":
            target = parse_month(month)
            out = out[out["MonthDate"] == target]
    if office != "All Offices" and "Office" in out.columns:
        out = out[out["Office"] == office]
    return out


def filter_office_only(df: pd.DataFrame, office: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    if office != "All Offices" and "Office" in out.columns:
        out = out[out["Office"] == office]
    return out

# ============================================================
# KPI CALCULATION
# ============================================================



def calculate_active_customers(shipment_df: pd.DataFrame) -> float:
    """
    Active Customers source: Shipment volume.
    For a single selected month, this returns the sum across selected offices.
    For multi-month / All selection, it returns the average monthly active-customer total
    because the source sheet contains monthly counts rather than customer-level IDs.
    """
    if shipment_df is None or shipment_df.empty or "Active Customers" not in shipment_df.columns:
        return 0.0

    d = shipment_df[["MonthDate", "Active Customers"]].copy()
    d["Active Customers"] = pd.to_numeric(d["Active Customers"], errors="coerce")
    d = d.dropna(subset=["MonthDate", "Active Customers"])

    if d.empty:
        return 0.0

    monthly = (
        d.groupby("MonthDate", as_index=False)["Active Customers"]
        .sum(min_count=1)
        .dropna(subset=["Active Customers"])
    )
    if monthly.empty:
        return 0.0

    if len(monthly) == 1:
        return float(monthly["Active Customers"].iloc[0])

    return float(monthly["Active Customers"].mean())


def calculate_kpis(hc, workload, fte, shipment) -> Dict[str, float]:
    total_workload_hours = float(workload["Workload Hours"].sum()) if not workload.empty and "Workload Hours" in workload.columns else 0.0
    required_fte_total_period = total_workload_hours / CAPACITY_HOURS_PER_FTE if total_workload_hours else 0.0

    # Capacity is summed over selected months. Actual FTE card is average monthly FTE.
    monthly_fte = pd.DataFrame()
    if not fte.empty:
        monthly_fte = fte.groupby("MonthDate", dropna=True)["Actual FTE"].sum().reset_index()
    actual_fte_avg = float(monthly_fte["Actual FTE"].mean()) if not monthly_fte.empty else 0.0
    capacity_hours = float((monthly_fte["Actual FTE"] * CAPACITY_HOURS_PER_FTE).sum()) if not monthly_fte.empty else 0.0
    required_fte_avg = required_fte_total_period / max(len(monthly_fte), 1) if not monthly_fte.empty else required_fte_total_period

    # HC is average monthly total when multiple months are selected.
    approved_hc = weighted_period_avg(hc, "Total Approved HC") if not hc.empty and "Total Approved HC" in hc.columns else 0.0
    actual_hc = weighted_period_avg(hc, "Total Actual HC") if not hc.empty and "Total Actual HC" in hc.columns else 0.0
    total_shipment = float(shipment["Total Shipment"].sum()) if not shipment.empty and "Total Shipment" in shipment.columns else 0.0
    util = safe_div(total_workload_hours, capacity_hours)
    fte_gap = actual_fte_avg - required_fte_avg
    return {
        "Approved HC": approved_hc,
        "Actual HC": actual_hc,
        "Actual FTE": actual_fte_avg,
        "Required FTE": required_fte_avg,
        "Capacity Hours": capacity_hours,
        "Workload Hours": total_workload_hours,
        "Utilization": util,
        "FTE Gap": fte_gap,
        "Total Shipment": total_shipment,
    }


def build_reconciliation(hc, workload, fte, shipment) -> pd.DataFrame:
    kpis = calculate_kpis(hc, workload, fte, shipment)
    rows = []
    # Reference values from HC if available.
    ref_actual_hc = weighted_period_avg(hc, "Total Actual HC") if not hc.empty and "Total Actual HC" in hc.columns else np.nan
    ref_approved_hc = weighted_period_avg(hc, "Total Approved HC") if not hc.empty and "Total Approved HC" in hc.columns else np.nan
    ref_required_hc = weighted_period_avg(hc, "Total Required HC") if not hc.empty and "Total Required HC" in hc.columns else np.nan
    ref_ship = shipment["Total Shipment"].sum() if not shipment.empty and "Total Shipment" in shipment.columns else np.nan

    def add(kpi, calc, ref, status="PASS", note=""):
        diff = calc - ref if pd.notna(ref) else np.nan
        rows.append({"KPI": kpi, "Calculated Value": calc, "Reference Value": ref, "Difference": diff, "Status": status, "Note": note})

    add("Total Approved HC", kpis["Approved HC"], ref_approved_hc, "PASS" if pd.notna(ref_approved_hc) else "WARNING", "Source: HC")
    add("Total Actual HC", kpis["Actual HC"], ref_actual_hc, "PASS" if pd.notna(ref_actual_hc) else "WARNING", "Source: HC")
    add("Total Shipment", kpis["Total Shipment"], ref_ship, "PASS" if pd.notna(ref_ship) else "WARNING", "Source: Shipment volume")
    add("Required FTE", kpis["Required FTE"], ref_required_hc, "WARNING", "Dashboard calculates from Workload Hours / 167.2; HC value is reference only.")
    add("Utilization", kpis["Utilization"], np.nan, "WARNING", "Dashboard formula: Workload Hours / Capacity Hours.")
    return pd.DataFrame(rows)

# ============================================================
# CHARTS
# ============================================================



def chart_office_capacity_trend(df: pd.DataFrame):
    """3-line HC trend from sheet HC with shaded gap between Approved HC and Actual HC."""
    if df.empty:
        st.info("No HC trend data available for selected filters.")
        return

    required_cols = ["MonthDate", "Total Approved HC", "Total Actual HC", "Total Required HC"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.info("HC trend cannot be displayed because required HC columns are missing.")
        return

    trend_source = df[
        ["MonthDate", "Total Approved HC", "Total Actual HC", "Total Required HC"]
    ].copy()

    for col in ["Total Approved HC", "Total Actual HC", "Total Required HC"]:
        trend_source[col] = pd.to_numeric(trend_source[col], errors="coerce")

    # Exclude months where all three HC values are blank.
    trend_source = trend_source.dropna(
        subset=["Total Approved HC", "Total Actual HC", "Total Required HC"],
        how="all",
    )

    if trend_source.empty:
        st.info("No HC trend data available for selected filters.")
        return

    trend = (
        trend_source.groupby("MonthDate", as_index=False)[
            ["Total Approved HC", "Total Actual HC", "Total Required HC"]
        ]
        .sum(min_count=1)
        .sort_values("MonthDate")
    )
    trend["Month"] = trend["MonthDate"].dt.strftime("%b-%y")

    # Use one canonical series variable for Required HC so the line and
    # its visible labels can never reference different column names.
    if "Total Required HC" in trend.columns:
        required_values = trend["Total Required HC"]
    elif "Required" in trend.columns:
        required_values = trend["Required"]
    elif "Required HC" in trend.columns:
        required_values = trend["Required HC"]
    else:
        st.info("HC trend cannot be displayed because Required HC data is missing.")
        return

    fig = go.Figure()

    # Approved HC line
    fig.add_trace(
        go.Scatter(
            x=trend["Month"],
            y=trend["Total Approved HC"],
            mode="lines+markers",
            name="Approved HC",
            line=dict(color=BUSINESS_COLORS["approved"], width=3),
            marker=dict(size=7),
            hovertemplate="%{x}<br>Approved HC: %{y:,.1f}<extra></extra>",
        )
    )

    # Actual HC line — baseline for the shaded Actual vs Required gap.
    fig.add_trace(
        go.Scatter(
            x=trend["Month"],
            y=trend["Total Actual HC"],
            mode="lines+markers",
            name="Actual HC",
            line=dict(color=BUSINESS_COLORS["actual"], width=3),
            marker=dict(size=7),
            hovertemplate="%{x}<br>Actual HC: %{y:,.1f}<extra></extra>",
        )
    )

    # Required HC line + shaded gap to Actual HC.
    # The fill is intentionally between Actual HC and Required HC.
    fig.add_trace(
        go.Scatter(
            x=trend["Month"],
            y=required_values,
            mode="lines+markers",
            name="Required HC",
            line=dict(color=BUSINESS_COLORS["required"], width=3, dash="solid"),
            marker=dict(size=7),
            fill="tonexty",
            fillcolor="rgba(245, 158, 11, 0.14)",
            hovertemplate="%{x}<br>Required HC: %{y:,.2f}<extra></extra>",
        )
    )

    fig.update_layout(
        title="HC Capacity Trend",
        yaxis_title="HC",
        hovermode="x unified",
    )
    fig = plotly_layout(fig, UI["chart_height"], show_legend=True, legend_position="top", margin_left=56, margin_right=42, margin_top=76, margin_bottom=46)

    # Keep the HC chart proportional: Y-axis always starts from zero.
    fig.update_yaxes(rangemode="tozero")

    fig.update_xaxes(type="category", categoryorder="array", categoryarray=trend["Month"].tolist())
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})



def chart_workload_by_pic(fte_df: pd.DataFrame, selected_office: str):
    """
    PIC Workload.

    Business display rule:
    - PIC Workload (hrs) = CS FTE Factor × Available Standard Time / PIC.
    - Available Standard Time / PIC = 167.2 hrs/month.
    - Utilization = PIC Workload / 167.2 = CS FTE Factor.

    Display logic:
    - Specific Office: show all PICs with data in that office.
    - All Offices: show Top 10 PICs by Utilization across all offices.
    - Colors:
        >100%      = Red (Overload)
        >95%–100%  = Amber (High Load)
        90%–95%    = Blue (Balanced)
        <90%       = Green (Less Load)
    """
    if fte_df is None or fte_df.empty:
        st.info("No CS FTE data available for selected filters.")
        return

    d = fte_df.copy()
    d["Actual FTE"] = pd.to_numeric(d["Actual FTE"], errors="coerce")
    d = d.dropna(subset=["Office", "CS PIC", "Actual FTE"])
    d = d[(d["Office"] != "") & (d["CS PIC"] != "")]

    if d.empty:
        st.info("No CS FTE data available for selected filters.")
        return

    pic_data = (
        d.groupby(["Office", "CS PIC"], as_index=False)["Actual FTE"]
        .mean()
    )
    pic_data["Avaiable time Hours"] = CAPACITY_HOURS_PER_FTE
    pic_data["Actual Workload Hours"] = pic_data["Actual FTE"] * CAPACITY_HOURS_PER_FTE
    pic_data["Utilization"] = pic_data["Actual FTE"]

    def _status(util):
        # Standard workload color rule:
        # >100% = Overload / Red
        # >95%–100% = High Load / Amber
        # 90%–95% = Balanced / Blue
        # <90% = Less Load / Green
        if util > 1.00:
            return "Overload", COLORS["red"]
        if util > 0.95:
            return "High Load", COLORS["high_load"]
        if util >= 0.90:
            return "Balanced", COLORS["blue"]
        return "Less Load", COLORS["green"]

    mapped = pic_data["Utilization"].apply(_status)
    pic_data["Status"] = mapped.map(lambda x: x[0])
    pic_data["Bar Color"] = mapped.map(lambda x: x[1])

    total_pic = int(len(pic_data))
    overloaded_pic = int((pic_data["Utilization"] > 1.0).sum())

    if selected_office == "All Offices":
        display = (
            pic_data.sort_values(
                ["Utilization", "Actual Workload Hours"],
                ascending=[False, False],
            )
            .head(10)
            .copy()
        )
        display["PIC Label"] = display.apply(
            lambda r: f"{r['Office']} | {r['CS PIC']}",
            axis=1,
        )
        subtitle = "Top 10 PICs by Capacity Utilization – All Offices"
    else:
        display = (
            pic_data[pic_data["Office"] == selected_office]
            .sort_values(
                ["Utilization", "Actual Workload Hours"],
                ascending=[False, False],
            )
            .copy()
        )
        display["PIC Label"] = display["CS PIC"].astype(str)
        subtitle = f"All PICs – {selected_office}"

    if display.empty:
        st.info("No PIC workload data available for selected filters.")
        return

    display = display.sort_values(
        ["Utilization", "Actual Workload Hours"],
        ascending=[True, True],
    )

    display["Label"] = display.apply(
        lambda r: f"{r['Actual Workload Hours']:,.1f} | {r['Utilization']*100:.0f}%",
        axis=1,
    )

    chart_height = max(360, min(650, 43 * len(display) + 150))

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=display["Actual Workload Hours"],
            y=display["PIC Label"],
            orientation="h",
            name="PIC Workload",
            marker_color=display["Bar Color"],
            text=display["Label"],
            textposition="outside",
            cliponaxis=False,
            customdata=np.column_stack([
                display["Office"],
                display["Actual FTE"],
                display["Utilization"],
                display["Status"],
            ]),
            hovertemplate=(
                "<b>%{y}</b><br>"
                "Office: %{customdata[0]}<br>"
                "PIC Workload: %{x:,.1f} hrs<br>"
                "CS FTE Factor: %{customdata[1]:.2f}<br>"
                "Utilization: %{customdata[2]:.1%}<br>"
                "Status: %{customdata[3]}"
                "<extra></extra>"
            ),
        )
    )

    fig.add_vline(
        x=CAPACITY_HOURS_PER_FTE,
        line_width=2.5,
        line_dash="dash",
        line_color=COLORS["navy"],
        annotation_text="Avaiable time 167.2 hour",
        annotation_position="top",
        annotation_font_color=COLORS["navy"],
    )

    max_actual = float(display["Actual Workload Hours"].max())
    x_max = max(max_actual * 1.15, CAPACITY_HOURS_PER_FTE * 1.25)

    fig.update_layout(
        title=dict(
            text=(
                "PIC Workload"
                f"<br><span style='font-size:11px;color:#667085;font-weight:400'>{subtitle}</span>"
            ),
            x=0.0,
            xanchor="left",
            y=0.98,
            yanchor="top",
            font=dict(size=UI["chart_title_size"], color=COLORS["navy"]),
        ),
        height=chart_height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Arial", color=COLORS["text"]),
        margin=dict(l=125, r=80, t=76, b=48),
        bargap=0.24,
        showlegend=False,
    )

    # Keep only one compact management summary in the upper-right.
    existing_annotations = list(fig.layout.annotations) if fig.layout.annotations else []
    existing_annotations.append(
        dict(
            text=f"Overloaded PICs: <b>{overloaded_pic}</b> / Total PICs: <b>{total_pic}</b>",
            x=1,
            y=1.075,
            xref="paper",
            yref="paper",
            showarrow=False,
            xanchor="right",
            yanchor="bottom",
            font=dict(size=UI["note_size"], color=COLORS["muted"]),
        )
    )
    fig.update_layout(annotations=existing_annotations)

    fig.update_xaxes(
        title_text="Workload Hours",
        range=[0, x_max],
        gridcolor=COLORS["grid"],
        zeroline=False,
        automargin=True,
        tickfont=dict(size=UI["axis_size"]),
        title_font=dict(size=UI["axis_size"], color=COLORS["gray_dark"]),
    )
    fig.update_yaxes(
        title_text="",
        gridcolor="rgba(0,0,0,0)",
        automargin=True,
        tickfont=dict(size=UI["axis_size"]),
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
    )

    st.markdown(
        f"""
        <div style="
            display:flex;
            justify-content:flex-end;
            align-items:center;
            gap:18px;
            margin-top:4px;
            margin-bottom:2px;
            color:#667085;
            font-size:11px;
            line-height:1.2;
            white-space:normal;
            flex-wrap:wrap;">
            <span><span style="display:inline-block;width:9px;height:9px;background:{COLORS['red']};margin-right:5px;border-radius:2px;"></span>Overload &gt;100%</span>
            <span><span style="display:inline-block;width:9px;height:9px;background:{COLORS['high_load']};margin-right:5px;border-radius:2px;"></span>High Load &gt;95–100%</span>
            <span><span style="display:inline-block;width:9px;height:9px;background:{COLORS['blue']};margin-right:5px;border-radius:2px;"></span>Balanced 90–95%</span>
            <span><span style="display:inline-block;width:9px;height:9px;background:{COLORS['green']};margin-right:5px;border-radius:2px;"></span>Less Load &lt;90%</span>
        </div>
        """,
        unsafe_allow_html=True,
    )



def chart_workload_by_service(df: pd.DataFrame):
    if df.empty:
        st.info("No workload data available for selected filters.")
        return
    agg = df.groupby(["Segment", "Service Label"], as_index=False)["Workload Hours"].sum()
    total = agg["Workload Hours"].sum()
    agg["% of Total"] = agg["Workload Hours"].apply(lambda x: safe_div(x, total))
    agg["Label"] = agg.apply(lambda r: f"{r['Workload Hours']:,.1f} | {r['% of Total']*100:.1f}%", axis=1)
    agg["SortOrder"] = agg["Segment"].apply(lambda x: SERVICE_ORDER.index(x) if x in SERVICE_ORDER else 999)
    agg = agg.sort_values(["Workload Hours"], ascending=True)
    fig = px.bar(
        agg,
        x="Workload Hours",
        y="Service Label",
        orientation="h",
        text="Label",
        color_discrete_sequence=[COLORS["blue"]],
        title="Workload Breakdown by Service Type",
    )
    fig.update_traces(textposition="outside", cliponaxis=False, hovertemplate="%{y}<br>%{x:,.1f} hours<extra></extra>")
    fig = plotly_layout(fig, UI["chart_height"], show_legend=False, margin_left=110, margin_right=70, margin_top=64, margin_bottom=44)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def chart_workload_composition(df: pd.DataFrame):
    if df.empty:
        st.info("No workload composition data available.")
        return
    comp = pd.DataFrame({
        "Activity": ["Core", "Ancillary", "Supporting", "Exception"],
        "Hours": [df["Core Hours"].sum(), df["Ancillary Hours"].sum(), df["Supporting Hours"].sum(), df["Exception Hours"].sum()],
    })
    total = comp["Hours"].sum()
    comp["Share"] = comp["Hours"].apply(lambda x: safe_div(x, total))
    fig = go.Figure()
    color_map = [COLORS["blue"], COLORS["green"], COLORS["amber"], COLORS["red"]]
    for _, row in comp.iterrows():
        color = color_map[comp.index[comp["Activity"] == row["Activity"]][0]]
        fig.add_trace(go.Bar(
            y=["Total Workload"],
            x=[row["Share"]],
            name=row["Activity"],
            orientation="h",
            marker_color=color,
            text=[f"{row['Activity']}<br>{row['Share']*100:.1f}%" if row["Share"] > 0.06 else ""],
            hovertemplate=f"{row['Activity']}: {row['Hours']:,.1f} hrs ({row['Share']*100:.1f}%)<extra></extra>",
        ))
    fig.update_layout(barmode="stack", xaxis_tickformat=".0%", title="Workload Composition – C/A/S/E")
    fig = plotly_layout(fig, 320, show_legend=True, legend_position="top", margin_left=52, margin_right=40, margin_top=66, margin_bottom=44)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def chart_workload_trend(df: pd.DataFrame):
    if df.empty:
        st.info("No trend data available.")
        return
    agg = df.groupby(["MonthDate", "Office"], as_index=False)["Workload Hours"].sum().sort_values("MonthDate")
    agg["Month"] = agg["MonthDate"].dt.strftime("%b-%y")
    fig = px.line(
        agg,
        x="Month",
        y="Workload Hours",
        color="Office",
        markers=True,
        color_discrete_sequence=[COLORS["blue"], COLORS["green"], COLORS["amber"], COLORS["red"]],
        title="Monthly Workload Trend",
    )
    fig.update_traces(hovertemplate="%{fullData.name}<br>%{x}: %{y:,.1f} hrs<extra></extra>")
    fig.update_xaxes(categoryorder="array", categoryarray=agg["Month"].drop_duplicates().tolist())
    fig = plotly_layout(fig, UI["chart_height"], show_legend=True, legend_position="top", margin_left=56, margin_right=40, margin_top=66, margin_bottom=48)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def chart_capacity_trend(workload: pd.DataFrame, fte: pd.DataFrame):
    if workload.empty and fte.empty:
        st.info("No capacity data available.")
        return
    wl = workload.groupby("MonthDate", as_index=False)["Workload Hours"].sum() if not workload.empty else pd.DataFrame(columns=["MonthDate", "Workload Hours"])
    ft = fte.groupby("MonthDate", as_index=False)["Actual FTE"].sum() if not fte.empty else pd.DataFrame(columns=["MonthDate", "Actual FTE"])
    cap = pd.merge(wl, ft, on="MonthDate", how="outer").fillna(0)
    cap["Capacity Hours"] = cap["Actual FTE"] * CAPACITY_HOURS_PER_FTE
    cap["Month"] = cap["MonthDate"].dt.strftime("%b-%y")
    cap = cap.sort_values("MonthDate")
    fig = go.Figure()
    fig.add_trace(go.Bar(x=cap["Month"], y=cap["Capacity Hours"], name="Capacity Hours", marker_color=COLORS["light_blue"]))
    fig.add_trace(go.Scatter(x=cap["Month"], y=cap["Workload Hours"], name="Workload Hours", mode="lines+markers", line=dict(color=COLORS["red"], width=3)))
    fig.update_layout(title="Workload vs Capacity Trend", yaxis_title="Hours")
    fig.update_xaxes(categoryorder="array", categoryarray=cap["Month"].tolist())
    fig = plotly_layout(fig, UI["chart_height"], show_legend=True, legend_position="top", margin_left=58, margin_right=40, margin_top=66, margin_bottom=48)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def build_segment_workload(
    df: pd.DataFrame,
    mode_df: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    """
    Section 4 source table — Workload by Segment.

    Business fields:
    - Allocation Time (h): total workload hours from BU allocation.
    - Workload Share (%): Segment workload / total workload.
    - Required FTE:
        Prefer source field "Office HC Allocation Ratio" when available.
        For multiple months, calculate monthly Segment FTE and average across
        valid months because Required FTE is a monthly capacity requirement,
        not a cumulative-period quantity.
        Fallback = monthly Workload Hours / 167.2 hours/FTE.
    - Shipment Volume: Shipment volume sheet, mapped to AE/AI/OE/OI/CC/TR/WH.
    """
    base_cols = [
        "Segment", "Allocation Time (h)", "Workload Share",
        "Required FTE", "Shipment Volume"
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=base_cols)

    d = df.copy()
    if "Segment" not in d.columns or "Workload Hours" not in d.columns:
        return pd.DataFrame(columns=base_cols)

    d["Segment"] = d["Segment"].astype(str).str.strip().str.upper()
    d["Workload Hours"] = pd.to_numeric(d["Workload Hours"], errors="coerce").fillna(0)
    d = d[d["Segment"].isin(SERVICE_ORDER)].copy()

    if d.empty:
        return pd.DataFrame(columns=base_cols)

    seg = (
        d.groupby("Segment", as_index=False)["Workload Hours"]
        .sum()
        .rename(columns={"Workload Hours": "Allocation Time (h)"})
    )

    seg = (
        pd.DataFrame({"Segment": SERVICE_ORDER})
        .merge(seg, on="Segment", how="left")
        .fillna({"Allocation Time (h)": 0.0})
    )

    total_hours = float(seg["Allocation Time (h)"].sum())
    seg["Workload Share"] = np.where(
        total_hours > 0,
        seg["Allocation Time (h)"] / total_hours,
        0.0,
    )

    fte_by_segment = pd.DataFrame({"Segment": SERVICE_ORDER, "Required FTE": 0.0})

    if "MonthDate" in d.columns and d["MonthDate"].notna().any():
        monthly = d.copy()

        use_source_fte = (
            "Office HC Allocation Ratio" in monthly.columns
            and pd.to_numeric(
                monthly["Office HC Allocation Ratio"], errors="coerce"
            ).fillna(0).abs().sum() > 0
        )

        if use_source_fte:
            monthly["Required FTE Source"] = pd.to_numeric(
                monthly["Office HC Allocation Ratio"], errors="coerce"
            )
            monthly_fte = (
                monthly.dropna(subset=["MonthDate"])
                .groupby(["MonthDate", "Segment"], as_index=False)["Required FTE Source"]
                .sum(min_count=1)
                .rename(columns={"Required FTE Source": "Required FTE"})
            )
        else:
            monthly_hours = (
                monthly.dropna(subset=["MonthDate"])
                .groupby(["MonthDate", "Segment"], as_index=False)["Workload Hours"]
                .sum()
            )
            monthly_hours["Required FTE"] = (
                monthly_hours["Workload Hours"] / CAPACITY_HOURS_PER_FTE
            )
            monthly_fte = monthly_hours[["MonthDate", "Segment", "Required FTE"]]

        if not monthly_fte.empty:
            fte_by_segment = (
                monthly_fte.groupby("Segment", as_index=False)["Required FTE"]
                .mean()
            )
            fte_by_segment = (
                pd.DataFrame({"Segment": SERVICE_ORDER})
                .merge(fte_by_segment, on="Segment", how="left")
                .fillna({"Required FTE": 0.0})
            )
    else:
        fte_by_segment = seg[["Segment", "Allocation Time (h)"]].copy()
        fte_by_segment["Required FTE"] = (
            fte_by_segment["Allocation Time (h)"] / CAPACITY_HOURS_PER_FTE
        )
        fte_by_segment = fte_by_segment[["Segment", "Required FTE"]]

    seg = seg.merge(fte_by_segment, on="Segment", how="left")
    seg["Required FTE"] = pd.to_numeric(
        seg["Required FTE"], errors="coerce"
    ).fillna(0)

    seg["Shipment Volume"] = 0.0
    if (
        mode_df is not None
        and not mode_df.empty
        and {"Mode", "Volume"}.issubset(mode_df.columns)
    ):
        vol = mode_df.copy()
        vol["Mode"] = vol["Mode"].astype(str).str.strip().str.upper()
        vol["Volume"] = pd.to_numeric(vol["Volume"], errors="coerce").fillna(0)

        volume_segment_map = {
            "AE": "AE", "AI": "AI", "OE": "OE", "OI": "OI",
            "OEFCL": "OE", "OELCL": "OE", "OIFCL": "OI", "OILCL": "OI",
            "CC": "CC", "CE": "CC", "CI": "CC",
            "TR": "TR", "DM": "TR", "DE": "TR", "DI": "TR",
            "WH": "WH", "HE": "WH", "HI": "WH",
        }
        vol["Segment"] = vol["Mode"].map(volume_segment_map)

        volume_by_segment = (
            vol.dropna(subset=["Segment"])
            .groupby("Segment", as_index=False)["Volume"]
            .sum()
            .rename(columns={"Volume": "Shipment Volume Source"})
        )
        seg = seg.merge(volume_by_segment, on="Segment", how="left")
        seg["Shipment Volume"] = pd.to_numeric(
            seg["Shipment Volume Source"], errors="coerce"
        ).fillna(0)
        seg = seg.drop(columns=["Shipment Volume Source"])

    seg["Segment"] = pd.Categorical(
        seg["Segment"], categories=SERVICE_ORDER, ordered=True
    )
    seg = seg.sort_values("Segment").reset_index(drop=True)
    seg["Segment"] = seg["Segment"].astype(str)

    return seg[
        ["Segment", "Allocation Time (h)", "Workload Share", "Required FTE", "Shipment Volume"]
    ]


def chart_service_matrix(
    df: pd.DataFrame,
    mode_df: Optional[pd.DataFrame] = None,
):
    """Workload by Segment — compact pie chart."""
    seg = build_segment_workload(df, mode_df)
    if seg.empty or float(seg["Allocation Time (h)"].sum()) <= 0:
        st.info("No segment workload data available for selected filters.")
        return


    plot_df = (
        seg[seg["Allocation Time (h)"] > 0]
        .copy()
        .sort_values("Workload Share", ascending=False)
        .reset_index(drop=True)
    )

    segment_color_map = {
        svc: CORPORATE_PALETTE[i % len(CORPORATE_PALETTE)]
        for i, svc in enumerate(SERVICE_ORDER)
    }
    pie_colors = plot_df["Segment"].map(segment_color_map).fillna(COLORS["blue"])
    text_positions = [
        "inside" if float(share) >= 0.04 else "outside"
        for share in plot_df["Workload Share"]
    ]
    text_templates = [
        "<b>%{label}</b><br>%{percent:.1%}"
        if float(share) >= 0.04
        else "<b>%{label}</b> %{percent:.1%}"
        for share in plot_df["Workload Share"]
    ]
    fig = go.Figure(
        go.Pie(
            labels=plot_df["Segment"],
            values=plot_df["Allocation Time (h)"],
            customdata=np.column_stack([
                plot_df["Shipment Volume"],
                plot_df["Required FTE"],
                plot_df["Workload Share"],
            ]),
            sort=False,
            direction="clockwise",
            rotation=90,
            hole=0,
            # Keep every slice joined inside one complete circle.
            pull=0,
            textposition=text_positions,
            texttemplate=text_templates,
            insidetextorientation="horizontal",
            insidetextfont=dict(
                family=UI["font_family"], size=11, color=COLORS["white"]
            ),
            outsidetextfont=dict(
                family=UI["font_family"], size=10, color=COLORS["navy"]
            ),
            marker=dict(
                colors=pie_colors,
                line=dict(color=COLORS["white"], width=2.5),
            ),
            hovertemplate=(
                "<b>%{label}</b>"
                "<br>Shipment Volume: %{customdata[0]:,.0f}"
                "<br>Allocation Time: %{value:,.1f} hrs"
                "<br>Required FTE: %{customdata[1]:.2f}"
                "<br>Workload Share: %{customdata[2]:.1%}"
                "<extra></extra>"
            ),
            showlegend=False,
            automargin=True,
        )
    )

    fig = plotly_layout(
        fig, 350, show_legend=False,
        margin_left=42, margin_right=42, margin_top=12, margin_bottom=18,
    )
    fig.update_layout(
        title=dict(text=""),
        uniformtext_minsize=9,
        uniformtext_mode="show",
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

def segment_workload_table(df: pd.DataFrame, mode_df: pd.DataFrame):
    """Executive summary table for Section 4; no TOTAL row in detail tables."""
    seg = build_segment_workload(df, mode_df)
    if seg.empty:
        st.info("No segment workload data available for selected filters.")
        return

    pair_panel_title("Segment Workload Breakdown (including MNG)")
    display = (
        seg.copy()
        .sort_values("Workload Share", ascending=False)
        .reset_index(drop=True)
        .rename(columns={"Workload Share": "Workload Share (%)"})
    )
    display["Workload Share (%)"] = pd.to_numeric(display["Workload Share (%)"], errors="coerce").fillna(0) * 100
    display = display[["Segment", "Shipment Volume", "Allocation Time (h)", "Required FTE", "Workload Share (%)"]]

    # Fit the table to its actual row count so no empty rows appear below.
    segment_table_height = 38 + 35 * len(display)

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
        height=segment_table_height,
        column_config={
            "Segment": st.column_config.TextColumn("Segment", width=70),
            "Shipment Volume": st.column_config.NumberColumn("Volume", width="medium", format="%,.0f"),
            "Allocation Time (h)": st.column_config.NumberColumn("Actual Working Time (Hours)", width="medium", format="%,.1f"),
            "Required FTE": st.column_config.NumberColumn("FTE Allocation", width="small", format="%.2f"),
            "Workload Share (%)": st.column_config.NumberColumn("Workload Share (%)", width="medium", format="%.1f%%"),
        },
    )

def chart_shipment_modes(mode_df: pd.DataFrame):
    """Horizontal bar chart showing shipment volume and share by transportation mode."""
    if mode_df.empty:
        st.info("No shipment mode data available for selected filters.")
        return
    agg = mode_df.groupby("Mode", as_index=False)["Volume"].sum().sort_values("Volume", ascending=False).reset_index(drop=True)
    total = float(agg["Volume"].sum())
    if total <= 0:
        st.info("No shipment mode data available for selected filters.")
        return

    pair_panel_title("Shipment Volume by Transportation Mode")
    agg["Share"] = agg["Volume"] / total
    plot_df = agg.sort_values("Volume", ascending=True).copy()
    plot_df["Display Label"] = plot_df.apply(lambda r: f"{r['Volume']:,.0f} ({r['Share']:.1%})", axis=1)
    fig = go.Figure(go.Bar(
        x=plot_df["Volume"], y=plot_df["Mode"], orientation="h",
        marker=dict(color=COLORS["blue"]), text=plot_df["Display Label"], textposition="outside",
        textfont=dict(size=UI["axis_size"], color=COLORS["navy"]), cliponaxis=False,
        hovertemplate="<b>%{y}</b><br>Shipment Volume: %{x:,.0f}<br>Share: %{customdata:.1%}<extra></extra>",
        customdata=plot_df["Share"],
    ))
    fig.update_layout(title_text="", xaxis_title=None, yaxis_title="", bargap=0.26)
    fig.update_yaxes(categoryorder="array", categoryarray=plot_df["Mode"].tolist(), automargin=True, tickfont=dict(size=UI["axis_size"]))
    fig.update_xaxes(automargin=True, rangemode="tozero", tickformat=",.0f")
    fig = plotly_layout(fig, 460, show_legend=False, margin_left=58, margin_right=105, margin_top=12, margin_bottom=40)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

def mode_detail_table(mode_df: pd.DataFrame):
    """Detail table paired with the transportation-mode chart; no TOTAL row."""
    if mode_df is None or mode_df.empty:
        st.info("No shipment mode detail available for selected filters.")
        return
    detail = mode_df.groupby("Mode", as_index=False)["Volume"].sum().sort_values("Volume", ascending=False).reset_index(drop=True)
    total = float(detail["Volume"].sum())
    if total <= 0:
        st.info("No shipment mode detail available for selected filters.")
        return
    pair_panel_title("Transportation Mode Detail")
    detail["Rank"] = np.arange(1, len(detail) + 1)
    detail["Share"] = detail["Volume"] / total
    display = detail.rename(columns={"Volume": "Shipment Volume"})[["Rank", "Mode", "Shipment Volume", "Share"]].copy()

    # Compact height: only show the rows that actually exist instead of
    # reserving the full chart height and leaving blank rows underneath.
    mode_table_height = min(SHIPMENT_PAIR_HEIGHT, 38 + 35 * len(display))

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
        height=mode_table_height,
        column_config={
            "Rank": st.column_config.NumberColumn("Rank", width="small", format="%d"),
            "Mode": st.column_config.TextColumn("Mode", width="small"),
            "Shipment Volume": st.column_config.NumberColumn("Shipment Volume", width="large", format="%,.0f"),
            "Share": st.column_config.NumberColumn("Share", width="small", format="percent"),
        },
    )

def build_customer_ranking(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate and rank all customers by shipment volume for current filters."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["Rank", "Customer", "Shipment Volume"])

    ranking = (
        df.groupby("Customer", as_index=False)["Volume"]
        .sum()
        .sort_values("Volume", ascending=False)
        .reset_index(drop=True)
    )
    ranking["Rank"] = np.arange(1, len(ranking) + 1)
    ranking = ranking.rename(columns={"Volume": "Shipment Volume"})
    return ranking[["Rank", "Customer", "Shipment Volume"]]


def chart_top_customers(df: pd.DataFrame):
    if df.empty:
        st.info("No customer volume data available for selected filters.")
        return
    ranking = build_customer_ranking(df)
    top = ranking.head(15).sort_values("Shipment Volume", ascending=True)
    pair_panel_title("Top 15 Customers by Shipment Volume")
    fig = px.bar(top, x="Shipment Volume", y="Customer", orientation="h", text="Shipment Volume", color_discrete_sequence=[COLORS["blue"]])
    fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False, hovertemplate="%{y}<br>Shipment Volume: %{x:,.0f}<extra></extra>")
    fig.update_layout(title_text="", yaxis_title="", xaxis_title=None, bargap=0.18)
    fig.update_yaxes(automargin=True, tickfont=dict(size=UI["axis_size"]))
    fig.update_xaxes(automargin=True)
    fig = plotly_layout(
        fig,
        SHIPMENT_PAIR_HEIGHT,
        show_legend=False,
        margin_left=155,
        margin_right=60,
        margin_top=22,
        margin_bottom=40,
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

def customer_detail_volume_table(df: pd.DataFrame):
    """Full customer ranking paired with the Top 15 chart; scrollable and no TOTAL row."""
    ranking = build_customer_ranking(df)
    if ranking.empty:
        st.info("No customer detail data available for selected filters.")
        return

    pair_panel_title("Customer Volume Detail")
    st.dataframe(
        ranking,
        use_container_width=True,
        hide_index=True,
        height=SHIPMENT_PAIR_HEIGHT,  # keep full-height scrollable detail for all customers
        column_config={
            "Rank": st.column_config.NumberColumn("Rank", width="small", format="%d"),
            "Customer": st.column_config.TextColumn("Customer", width="large"),
            "Shipment Volume": st.column_config.NumberColumn("Shipment Volume", width="medium", format="%,.0f"),
        },
    )


def chart_resolution(df: pd.DataFrame):
    """CS Solution performance chart."""
    if df is None or df.empty:
        st.info("No CS Resolution data available for selected filters.")
        return
    pair_panel_title("CS Resolution Trend")
    agg = df.groupby("MonthDate", as_index=False).agg(**{"Total Abnormality": ("Total Abnormality", "sum"), "Resolved": ("Resolved", "sum")}).sort_values("MonthDate")
    agg["Resolution Rate"] = np.where(agg["Total Abnormality"] > 0, agg["Resolved"] / agg["Total Abnormality"], np.nan)
    agg["Month"] = agg["MonthDate"].dt.strftime("%b-%y")
    # Use numeric x positions so the resolution-rate marker stays exactly at
    # the horizontal centre of its corresponding "Resolved by CS" bar.
    month_x = np.arange(len(agg), dtype=float)
    bar_offset = 0.18
    bar_width = 0.34
    total_x = month_x - bar_offset
    resolved_x = month_x + bar_offset
    fig = go.Figure()
    fig.add_trace(go.Bar(x=total_x, width=bar_width, y=agg["Total Abnormality"], name="Total Exception Case", marker_color=BUSINESS_COLORS["supporting"], text=agg["Total Abnormality"], texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False, customdata=agg["Month"], hovertemplate="%{customdata}<br>Total Exception Case: %{y:,.0f}<extra></extra>"))
    fig.add_trace(go.Bar(x=resolved_x, width=bar_width, y=agg["Resolved"], name="Resolved by CS", marker_color=BUSINESS_COLORS["actual"], text=agg["Resolved"], texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False, customdata=agg["Month"], hovertemplate="%{customdata}<br>Resolved by CS: %{y:,.0f}<extra></extra>"))
    fig.add_trace(go.Scatter(x=resolved_x, y=agg["Resolution Rate"], name="CS Resolution Rate", mode="lines+markers+text", line=dict(color="#F97316", width=3), marker=dict(size=8, color="#F97316", line=dict(color="#FFFFFF", width=1.5)), text=agg["Resolution Rate"], texttemplate="%{text:.1%}", textposition="bottom center", yaxis="y2", customdata=agg["Month"], hovertemplate="%{customdata}<br>CS Resolution Rate: %{y:.1%}<extra></extra>"))
    fig.update_layout(title_text="", barmode="overlay", yaxis=dict(title="Cases", rangemode="tozero"), yaxis2=dict(title="Resolution Rate", overlaying="y", side="right", tickformat=".0%", range=[0, 1.20], showgrid=False))
    fig = plotly_layout(fig, 390, show_legend=True, legend_position="top", margin_left=58, margin_right=68, margin_top=38, margin_bottom=44)
    fig.update_xaxes(tickmode="array", tickvals=month_x, ticktext=agg["Month"].tolist(), range=[-0.55, len(agg) - 0.45])
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

def render_cs_solution_table(df: pd.DataFrame):
    """CS Resolution detail table; no TOTAL row."""
    if df is None or df.empty:
        st.info("No CS Resolution data available for selected filters.")
        return
    pair_panel_title("CS Resolution by Office")
    d = df.copy().sort_values(["Office", "MonthDate"])
    d["Month"] = d["MonthDate"].dt.strftime("%b-%y")
    display = d[["Office", "Month", "Total Abnormality", "Resolved", "Resolution Rate"]].copy()
    display["Resolution Rate (%)"] = (
        pd.to_numeric(display["Resolution Rate"], errors="coerce") * 100
    )
    display["Resolution Rate (%)"] = display["Resolution Rate (%)"].apply(
        lambda value: (
            ""
            if pd.isna(value)
            else f"{value:.0f}%"
            if np.isclose(value, 0) or np.isclose(value, 100)
            else f"{value:.2f}%"
        )
    )
    display = display.drop(columns=["Resolution Rate"])

    display["Total Abnormality"] = display["Total Abnormality"].apply(
        lambda value: "" if pd.isna(value) else f"{value:,.0f}"
    )
    display["Resolved"] = display["Resolved"].apply(
        lambda value: "" if pd.isna(value) else f"{value:,.0f}"
    )
    display = display.rename(
        columns={
            "Total Abnormality": "Total Exception Case",
            "Resolved": "Resolved by CS",
            "Resolution Rate (%)": "CS Resolution Rate",
        }
    )

    resolution_table_height = min(390, max(160, 38 + len(display) * 34))
    table_html = display.to_html(
        index=False,
        escape=True,
        classes="cs-resolution-table",
        border=0,
    )
    table_css = textwrap.dedent(
        f"""
        <style>
        .cs-resolution-table-wrap {{
            width: 100%;
            max-height: {resolution_table_height}px;
            overflow-y: auto;
            border: 1px solid #D8E0E8;
            border-radius: 10px;
            background: #FFFFFF;
        }}
        .cs-resolution-table {{
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
            font-family: {UI['font_family']};
            font-size: 13px;
        }}
        .cs-resolution-table th,
        .cs-resolution-table td {{
            height: 34px;
            padding: 6px 8px;
            border-right: 1px solid #E1E6EB;
            border-bottom: 1px solid #E1E6EB;
            text-align: center !important;
            vertical-align: middle !important;
            white-space: nowrap;
        }}
        .cs-resolution-table th {{
            position: sticky;
            top: 0;
            z-index: 1;
            background: #F7F9FB;
            color: #718096;
            font-weight: 400;
        }}
        .cs-resolution-table th:last-child,
        .cs-resolution-table td:last-child {{ border-right: 0; }}
        .cs-resolution-table tbody tr:last-child td {{ border-bottom: 0; }}
        </style>
        """
    ).strip()
    st.markdown(
        table_css
        + f'\n<div class="cs-resolution-table-wrap">{table_html}</div>',
        unsafe_allow_html=True,
    )

def chart_yvf(df: pd.DataFrame):
    """YVF booking share of Total IFF Bookings."""
    if df is None or df.empty:
        st.info("No YVF data available for selected filters.")
        return
    d = df.copy()
    d["YVF Booking"] = pd.to_numeric(d["YVF Booking"], errors="coerce").fillna(0)
    d["IFF Shipment"] = pd.to_numeric(d["IFF Shipment"], errors="coerce").fillna(0)
    d = d[(d["YVF Booking"] != 0) | (d["IFF Shipment"] != 0)].copy()
    if d.empty:
        st.info("No YVF data available for selected filters.")
        return
    pair_panel_title("YVF Booking Adoption")
    total_yvf = float(d["YVF Booking"].sum()); total_iff = float(d["IFF Shipment"].sum())
    remaining_iff = max(total_iff - total_yvf, 0.0); ratio = safe_div(total_yvf, total_iff)
    fig = go.Figure(
        data=[
            go.Pie(
                labels=["YVF Bookings", "IFF Bookings"],
                values=[total_yvf, remaining_iff],
                hole=0.58,
                sort=False,
                direction="clockwise",
                marker=dict(
                    colors=[BUSINESS_COLORS["actual"], COLORS["grid"]],
                    line=dict(color="white", width=2),
                ),
                textinfo="none",
                hovertemplate=(
                    "<b>%{label}</b>"
                    "<br>Shipments: %{value:,.0f}"
                    "<br>Share: %{percent:.1%}"
                    "<extra></extra>"
                ),
            )
        ]
    )
    fig.update_layout(title_text="", annotations=[dict(text=f"<b>{ratio:.1%}</b><br><span style='font-size:12px'>YVF Adoption</span><br><span style='font-size:11px'>{total_yvf:,.0f} / {total_iff:,.0f}</span>", x=0.5, y=0.5, font=dict(size=22, color=COLORS["navy"], family=UI["font_family"]), showarrow=False, align="center")])
    fig = plotly_layout(fig, 340, show_legend=True, legend_position="top", margin_left=44, margin_right=44, margin_top=34, margin_bottom=24)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

def render_yvf_table(df: pd.DataFrame):
    """YVF detail table; no TOTAL row."""
    if df is None or df.empty:
        st.info("No YVF data available for selected filters.")
        return
    d = df.copy()
    d = d[(pd.to_numeric(d["YVF Booking"], errors="coerce").fillna(0) != 0) | (pd.to_numeric(d["IFF Shipment"], errors="coerce").fillna(0) != 0)].copy()
    if d.empty:
        st.info("No YVF data available for selected filters.")
        return
    pair_panel_title("YVF Performance by Office")
    has_month = "MonthDate" in d.columns and d["MonthDate"].notna().any()
    if has_month:
        d = d.sort_values(["MonthDate", "Office"]).copy(); d["Month"] = d["MonthDate"].dt.strftime("%b-%y")
        display = d[["Office", "Month", "YVF Booking", "IFF Shipment", "YVF Booking Ratio"]].copy()
    else:
        display = d[["Office", "YVF Booking", "IFF Shipment", "YVF Booking Ratio"]].copy().sort_values(["Office"])
    column_cfg = {
        "Office": st.column_config.TextColumn("Office", width=70),
        "YVF Booking": st.column_config.NumberColumn("Total YVF Bookings", width="medium", format="%,.0f"),
        "IFF Shipment": st.column_config.NumberColumn("Total IFF Bookings", width="medium", format="%,.0f"),
        "YVF Booking Ratio": st.column_config.NumberColumn("YVF Booking Ratio", width=110, format="%.1f%%"),
    }
    if has_month:
        column_cfg["Month"] = st.column_config.TextColumn("Month", width="small")
    display["YVF Booking Ratio"] = pd.to_numeric(
        display["YVF Booking Ratio"], errors="coerce"
    ) * 100

    yvf_table_height = min(
        390,
        max(160, 38 + len(display) * 34),
    )

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
        height=yvf_table_height,
        column_config=column_cfg,
    )

# ============================================================
# COVER / WELCOME PAGE
# UI only — does not change business logic or dashboard calculations
# ============================================================

def render_cover_page() -> None:
    """Render the approved corporate cover page."""

    cover_css = """
<style>
section[data-testid="stSidebar"]{display:none!important}
[data-testid="stSidebarCollapsedControl"]{display:none!important}
header[data-testid="stHeader"]{display:none!important;height:0!important}
[data-testid="stToolbar"]{display:none!important}
[data-testid="stDecoration"]{display:none!important}
[data-testid="stStatusWidget"]{display:none!important}
#MainMenu{visibility:hidden!important}
[data-testid="stAppViewContainer"]{height:100dvh!important;overflow:hidden!important;padding:0!important;margin:0!important}
[data-testid="stAppViewBlockContainer"]{height:100dvh!important;overflow:hidden!important;padding:0!important;margin:0!important}
.main{height:100dvh!important;overflow:hidden!important;padding:0!important;margin:0!important}
footer{display:none!important}
.block-container{max-width:none!important;width:100%!important;padding:0!important;margin:0!important}
.stApp{background:#06183F!important;height:100dvh!important;min-height:0!important;overflow:hidden!important}

.cover-stage{
    position:fixed;
    inset:0;
    width:100vw;
    height:100dvh;
    min-height:0;
    overflow:hidden;
    z-index:999;
    background:linear-gradient(135deg,#031532 0%,#06183F 48%,#082759 100%);
    font-family:Inter,"Segoe UI",Arial,sans-serif;
}

.cover-panel{
    position:absolute;
    top:10px;
    bottom:14px;
    left:1.2vw;
    width:56vw;
    max-width:none;
    min-width:760px;
    height:auto;
    min-height:0;
    box-sizing:border-box;
    padding:22px 44px 24px;
    border-radius:24px;
    background:linear-gradient(145deg,#FFFFFF 0%,#FCFDFE 100%);
    border:1px solid rgba(255,255,255,.72);
    box-shadow:0 18px 42px rgba(0,0,0,.18);
    z-index:10;
}

.cover-logo-real{
    display:none !important;
}

.cover-title{
    margin:0;
    max-width:650px;
    color:#06183F;
    font-size:clamp(46px,3.65vw,66px);
    line-height:.98;
    letter-spacing:-.035em;
    font-weight:850;
    text-transform:uppercase;
}

.cover-title-accent{
    width:100%;
    max-width:100%;
    height:5px;
    border-radius:999px;
    background:#E6761B;
    margin:clamp(14px,2.2vh,22px) 0 clamp(10px,1.6vh,16px) 2px;
}

.cover-subtitle{
    color:#4F5B6A;
    font-size:clamp(18px,1.45vw,23px);
    line-height:1.24;
    font-weight:500;
    max-width:440px;
    margin-bottom:clamp(12px,1.8vh,20px);
}

.cover-separator{
    height:1px;
    background:#D8E1EA;
    margin:0 0 clamp(12px,1.8vh,18px);
}

.cover-pillars{
    display:grid;
    grid-template-columns:repeat(4,minmax(0,1fr));
    gap:0;
    margin-top:38px;
    padding:10px 12px 0;
}

.cover-pillar{
    position:relative;
    text-align:center;
    padding:0 22px;
}

.cover-pillar:not(:last-child)::after{
    content:"";
    position:absolute;
    right:0;
    top:48px;
    width:1px;
    height:92px;
    background:#DDE5EC;
}

.cover-icon{
    width:72px;
    height:72px;
    margin:0 auto 12px;
    border-radius:50%;
    display:flex;
    align-items:center;
    justify-content:center;
    color:#fff;
    box-shadow:0 5px 14px rgba(0,0,0,.10);
}

.cover-icon svg{
    width:39px;
    height:39px;
    stroke:currentColor;
    fill:none;
    stroke-width:1.9;
    stroke-linecap:round;
    stroke-linejoin:round;
}

.icon-capacity{background:#06183F}
.icon-workload{background:#0DBAEE}
.icon-productivity{background:#F57C00}
.icon-insights{background:#45B84A}

.cover-pillar-title{
    color:#06183F;
    font-size:16px;
    font-weight:800;
    line-height:1.18;
    text-transform:uppercase;
    margin-bottom:9px;
}

.cover-pillar-note{
    color:#4F5B6A;
    font-size:14px;
    line-height:1.40;
    font-weight:500;
}

.cover-cta{
    position:absolute;
    left:36px;
    bottom:32px;
    display:inline-flex;
    align-items:center;
    justify-content:space-between;
    gap:18px;
    min-width:300px;
    height:52px;
    padding:0 20px;
    box-sizing:border-box;
    border-radius:9px;
    background:linear-gradient(180deg,#FF7F0A 0%,#EE6500 100%);
    color:#fff!important;
    text-decoration:none!important;
    font-size:20px;
    font-weight:800;
    box-shadow:0 8px 18px rgba(230,118,27,.28);
}

.cover-cta-icon{
    width:30px;
    height:30px;
    border:2px solid rgba(255,255,255,.9);
    border-radius:4px;
    display:flex;
    align-items:center;
    justify-content:center;
    flex:0 0 30px;
}

.cover-cta-arrow{
    font-size:34px;
    line-height:1;
    margin-left:auto;
    font-weight:300;
}

/* SVG arc field based on the approved reference */
.cover-arc-svg{
    position:absolute;
    right:0;
    top:0;
    width:55vw;
    height:100dvh;
    z-index:2;
    pointer-events:none;
    overflow:visible;
}

.cover-right-footer{
    position:absolute;
    right:28px;
    bottom:24px;
    display:flex;
    align-items:center;
    gap:11px;
    color:#06183F;
    font-size:13px;
    font-weight:700;
    z-index:12;
    white-space:nowrap;
}

.cover-headset{
    width:30px;
    height:30px;
    display:flex;
    align-items:center;
    justify-content:center;
    color:#06183F;
}

.cover-headset svg{
    width:28px;
    height:28px;
    stroke:currentColor;
    fill:none;
    stroke-width:1.8;
}

.footer-divider{
    width:1.5px;
    height:22px;
    background:#E6761B;
}

.cover-wave-1,.cover-wave-2,.cover-wave-3{
    position:absolute;
    left:-5%;
    width:112%;
    border-radius:50%;
    pointer-events:none;
}

.cover-wave-1{
    height:90px;
    bottom:-58px;
    border-top:9px solid #0DBAEE;
    transform:rotate(-1.5deg);
    opacity:.96;
}

.cover-wave-2{
    height:125px;
    bottom:-88px;
    border-top:7px solid #E6761B;
    transform:rotate(1.7deg);
    opacity:.98;
}

.cover-wave-3{
    height:155px;
    bottom:-113px;
    border-top:8px solid #005BAC;
    transform:rotate(-.2deg);
    opacity:.86;
}


@media(max-height:820px) and (min-width:901px){
    .cover-panel{top:8px;bottom:10px;padding:18px 40px 18px;width:64vw;min-width:760px}
    .cover-logo-real{width:176px;margin-bottom:11px}
    .cover-title{font-size:clamp(44px,3.3vw,60px)}
    .cover-title-accent{margin:14px 0 11px 2px}
    .cover-subtitle{font-size:19px;margin-bottom:12px}
    .cover-separator{margin-bottom:12px}
    .cover-icon{width:64px;height:64px}
    .cover-icon svg{width:35px;height:35px}
    .cover-pillar-title{font-size:14px;margin-bottom:7px}
    .cover-pillar-note{font-size:12.5px;line-height:1.38}
    .cover-pillars{margin-top:30px;padding:8px 8px 0}
    .cover-pillar{padding:0 16px}
    .cover-pillar:not(:last-child)::after{top:42px;height:84px}
    .cover-cta{bottom:26px;height:50px;min-width:300px;font-size:19px}
    .cover-right-footer{right:24px;bottom:26px;font-size:12px}
}

@media(max-width:900px){
    .cover-stage{min-height:920px}
    .cover-panel{
        position:relative;
        top:auto;
        left:auto;
        width:calc(100% - 28px);
        min-width:0;
        height:auto;
        min-height:830px;
        margin:14px;
        padding:28px 24px 110px;
    }
    .cover-title{font-size:46px}
    .cover-pillars{grid-template-columns:repeat(2,1fr);row-gap:20px}
    .cover-pillar:nth-child(2)::after{display:none}
    .cover-cta{left:24px;right:24px;width:auto;bottom:28px}
    .cover-arc-svg,.cover-wave-1,.cover-wave-2,.cover-wave-3{display:none}.cover-right-footer{right:20px;bottom:18px;font-size:11px;gap:8px}
}

    /* ===== FINAL OVERRIDE: FTE WORKLOAD STATUS ===== */
    .workload-status-text,
    .status-badge {
        font-size: 28px !important;
        line-height: 1.05 !important;
        font-weight: 800 !important;
        min-height: 44px !important;
        padding: 6px 20px !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        text-align: center !important;
        border-radius: 999px !important;
        letter-spacing: 0 !important;
    }

    @media (max-width: 1366px) {
        .workload-status-text,
        .status-badge {
            font-size: 26px !important;
            min-height: 42px !important;
            padding: 5px 18px !important;
        }
    }

</style>
"""

    cover_html = """<div class="cover-stage">
<svg class="cover-arc-svg" viewBox="0 0 900 900" preserveAspectRatio="xMidYMid slice" aria-hidden="true">
  <path d="M560 -60 C690 120 700 300 610 470 C550 585 445 675 330 735" fill="none" stroke="#F58220" stroke-width="2.2"/>
  <path d="M640 -45 C760 130 770 320 685 495 C625 620 525 710 415 775" fill="none" stroke="#F58220" stroke-width="2.2"/>
  <path d="M715 -28 C825 145 838 345 758 525 C700 655 610 752 500 820" fill="none" stroke="#F58220" stroke-width="2.2"/>
  <path d="M785 -10 C885 165 900 375 825 555 C772 685 690 785 585 855" fill="none" stroke="#F58220" stroke-width="2.2"/>
  <path d="M850 12 C940 190 958 405 890 585 C842 710 765 815 670 885" fill="none" stroke="#F58220" stroke-width="2.2"/>
  <circle cx="603" cy="160" r="7" fill="#F58220"/>
  <circle cx="690" cy="235" r="7" fill="#F58220"/>
  <circle cx="752" cy="340" r="7" fill="#F58220"/>
  <circle cx="636" cy="510" r="7" fill="#F58220"/>
  <circle cx="780" cy="625" r="7" fill="#F58220"/>
  <circle cx="700" cy="760" r="7" fill="#F58220"/>
</svg>
<div class="cover-wave-1"></div><div class="cover-wave-2"></div><div class="cover-wave-3"></div>
<div class="cover-panel">
<img src="data:image/png;base64,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" class="cover-logo-real" alt="Yusen Logistics">
<h1 class="cover-title">CS OPERATIONS PERFORMANCE DASHBOARD</h1>
<div class="cover-title-accent"></div>
<div class="cover-subtitle">Capacity • Workload • Utilization • Performance</div>
<div class="cover-separator"></div>
<div class="cover-pillars">
<div class="cover-pillar"><div class="cover-icon icon-capacity"><svg viewBox="0 0 32 32"><circle cx="16" cy="9" r="4"></circle><circle cx="7" cy="12" r="3"></circle><circle cx="25" cy="12" r="3"></circle><path d="M9 25v-3c0-4 3-7 7-7s7 3 7 7v3"></path><path d="M2 24v-2c0-3 2-5 5-5"></path><path d="M30 24v-2c0-3-2-5-5-5"></path></svg></div><div class="cover-pillar-title">Capacity</div><div class="cover-pillar-note">HC Capacity,<br>Requirement & Gap</div></div>
<div class="cover-pillar"><div class="cover-icon icon-workload"><svg viewBox="0 0 32 32"><rect x="7" y="6" width="18" height="22" rx="2"></rect><path d="M12 6V4h8v2"></path><path d="M11 12h10"></path><path d="M11 17h10"></path><path d="M11 22h7"></path></svg></div><div class="cover-pillar-title">Workload</div><div class="cover-pillar-note">Customer, Volume,<br>Segment & Activity</div></div>
<div class="cover-pillar"><div class="cover-icon icon-productivity"><svg viewBox="0 0 32 32"><path d="M5 25h22"></path><rect x="7" y="17" width="4" height="8"></rect><rect x="14" y="12" width="4" height="13"></rect><rect x="21" y="7" width="4" height="18"></rect><path d="M6 13l6-5 5 2 8-7"></path><path d="M22 3h4v4"></path></svg></div><div class="cover-pillar-title">UTILIZATION</div><div class="cover-pillar-note">Office Workload,<br>CS Allocation</div></div>
<div class="cover-pillar"><div class="cover-icon icon-insights"><svg viewBox="0 0 32 32"><circle cx="16" cy="16" r="10"></circle><circle cx="16" cy="16" r="5"></circle><path d="M16 16l8-8"></path><path d="M22 8h4v4"></path></svg></div><div class="cover-pillar-title">PERFORMANCE</div><div class="cover-pillar-note">CS Resolution,<br>YVF Booking Adoption</div></div>
</div>
<a href="?enter=1" target="_self" class="cover-cta"><span class="cover-cta-icon">↗</span><span>VIEW DASHBOARD</span><span class="cover-cta-arrow">→</span></a>
<div class="cover-right-footer"><span class="cover-headset"><svg viewBox="0 0 32 32"><path d="M5 17v-2a11 11 0 0 1 22 0v2"></path><rect x="3" y="16" width="5" height="9" rx="2"></rect><rect x="24" y="16" width="5" height="9" rx="2"></rect><path d="M24 26c-2 3-5 3-8 3"></path></svg></span><span class="footer-divider"></span><span>CS DIVISION</span><span class="footer-divider"></span><span>FY2026</span></div>
</div>

</div>"""

    st.markdown(cover_css.strip(), unsafe_allow_html=True)
    st.markdown(cover_html.strip(), unsafe_allow_html=True)



def render_cover_gate() -> None:
    """Stop on the cover until the user selects VIEW DASHBOARD."""
    if "dashboard_entered" not in st.session_state:
        st.session_state["dashboard_entered"] = False

    # The cover CTA is a styled HTML link to ?enter=1.
    # This allows the button to remain visually inside the white cover panel.
    try:
        enter_param = st.query_params.get("enter")
    except Exception:
        logger.debug("Unable to read Streamlit query parameters", exc_info=True)
        enter_param = None

    if isinstance(enter_param, list):
        enter_param = enter_param[0] if enter_param else None

    if str(enter_param) == "1":
        st.session_state["dashboard_entered"] = True
        try:
            st.query_params.clear()
        except Exception:
            logger.debug("Unable to clear Streamlit query parameters", exc_info=True)

    if not st.session_state["dashboard_entered"]:
        render_cover_page()
        st.stop()



# ============================================================
# MAIN APP
# ============================================================


def main():
    # Cover page is displayed before any Excel loading/filtering.
    render_cover_gate()

    # Load default workbook first. Upload remains below Month / Office filters.
    file_path = DEFAULT_FILE_PATH
    cache_token = ""

    # If an uploaded workbook was already saved in this session, reuse the saved file.
    uploaded_path_cached = st.session_state.get("dashboard_uploaded_path")
    uploaded_sig_cached = st.session_state.get("dashboard_uploaded_sig")

    if uploaded_path_cached and Path(uploaded_path_cached).exists():
        file_path = Path(uploaded_path_cached)
        cache_token = uploaded_sig_cached or ""
    elif file_path.exists():
        stat = file_path.stat()
        cache_token = f"{stat.st_mtime_ns}_{stat.st_size}"

    if not file_path.exists():
        st.error(
            f"Không tìm thấy file dữ liệu: {file_path}. "
            "Vui lòng đặt file Excel cùng thư mục app.py hoặc upload file ở Sidebar."
        )
        st.stop()

    with st.spinner("Loading and validating Excel data..."):
        try:
            raw = load_data(str(file_path), cache_token)
        except WorkbookLoadError as exc:
            st.error(str(exc))
            st.stop()
        hc = prepare_hc(raw["hc"])
        workload = prepare_workload(raw["workload"])
        fte = prepare_fte(raw["fte"])
        shipment, shipment_mode = prepare_shipment(raw["shipment"])

        # New MASTER DATA SOURCE stores shipment-by-segment volume in
        # "4. Workload by Activity" (C Volume), while the old workbook stored
        # transportation-mode columns inside "Shipment volume".
        # Build the same canonical Mode/Volume table only when direct mode data is absent.
        if shipment_mode.empty and not workload.empty and "Core Volume" in workload.columns:
            _mode = workload[["Office", "MonthDate", "Segment", "Core Volume"]].copy()
            _mode["Volume"] = pd.to_numeric(_mode["Core Volume"], errors="coerce").fillna(0)
            _mode = _mode[_mode["Volume"] > 0]
            shipment_mode = _mode.rename(columns={"Segment": "Mode"})[
                ["Office", "MonthDate", "Mode", "Volume"]
            ].reset_index(drop=True)

        customer = prepare_customer(raw)
        # Section 2 customer ranking/detail uses the mapped customer-volume source: 11. Vol. by Customer.
        customer_ns = customer_wide_to_long(raw["customer_ns"])
        resolution = prepare_resolution(raw["resolution"])
        yvf = prepare_yvf(raw["yvf"])

        # Section 5 detail sources (C / A / S / E)
        core_detail = prepare_case_detail(raw["core"], "Core Service")
        ancillary_detail = prepare_case_detail(raw["ancillary"], "Ancillary Service")
        supporting_detail = prepare_case_detail(raw["supporting"], "Supporting Activity")
        exception_detail = prepare_case_detail(raw["exception"], "Exception Handling")

        # Code description lookup from sheet "Ghi chú".
        code_note_map = prepare_code_note_map(raw["notes"])

        # Add one consistent "Code Description" field to all C/A/S/E sources.
        core_detail = add_code_description(core_detail, "Core Service", code_note_map)
        ancillary_detail = add_code_description(ancillary_detail, "Ancillary Service", code_note_map)
        supporting_detail = add_code_description(supporting_detail, "Supporting Activity", code_note_map)
        exception_detail = add_code_description(exception_detail, "Exception Handling", code_note_map)

    periods = all_periods(hc, workload, fte, shipment, customer, customer_ns, resolution)
    month_options = ["All"] + [format_month(p) for p in periods]

    offices_from_data = sorted(set(
        list(hc.get("Office", pd.Series(dtype=str)).dropna().unique())
        + list(workload.get("Office", pd.Series(dtype=str)).dropna().unique())
        + list(fte.get("Office", pd.Series(dtype=str)).dropna().unique())
        + list(shipment.get("Office", pd.Series(dtype=str)).dropna().unique())
        + list(customer.get("Office", pd.Series(dtype=str)).dropna().unique())
        + list(customer_ns.get("Office", pd.Series(dtype=str)).dropna().unique())
    ))
    office_options = ["All Offices"] + sorted(set(STANDARD_OFFICES + [o for o in offices_from_data if o]))

    # Sidebar order: Month -> Office -> Upload file. No Year and no Reset button.
    # UI only: styled to match the approved Yusen executive HOME format.
    with st.sidebar:
        if st.button("HOME", icon=":material/home:", use_container_width=True, key="back_to_cover_btn"):
            st.session_state["dashboard_entered"] = False
            st.rerun()
        st.markdown('<div class="sidebar-filter-title">FILTERS</div><div class="sidebar-filter-spacer"></div>', unsafe_allow_html=True)
        month = st.selectbox("MONTH", month_options, key="month_filter")
        office = st.selectbox("OFFICE", office_options, key="office_filter")
        st.markdown('<div class="sidebar-bottom-anchor"></div>', unsafe_allow_html=True)
        st.markdown("---")
        uploaded = st.file_uploader(
            "UPLOAD EXCEL FILE",
            type=["xlsx", "xlsm"],
            help=(
                "Hỗ trợ định dạng .xlsx và .xlsm. Nếu không upload, Dashboard "
                "sẽ đọc file mặc định trong cùng thư mục với file Python."
            ),
            key="excel_uploader",
        )
        if uploaded is not None:
            new_bytes = uploaded.getvalue()
            new_sig = hashlib.sha256(new_bytes).hexdigest()

            if st.session_state.get("dashboard_uploaded_sig") != new_sig:
                upload_dir = Path(tempfile.gettempdir()) / "cs_workload_dashboard"
                upload_dir.mkdir(parents=True, exist_ok=True)
                upload_suffix = Path(uploaded.name).suffix.lower()
                tmp_path = upload_dir / f"workbook_{new_sig[:20]}{upload_suffix}"

                # Content-addressed filenames prevent different user sessions
                # from overwriting one another on the Streamlit server.
                if not tmp_path.exists() or tmp_path.stat().st_size != len(new_bytes):
                    tmp_path.write_bytes(new_bytes)

                st.session_state["dashboard_uploaded_sig"] = new_sig
                st.session_state["dashboard_uploaded_path"] = str(tmp_path)
                st.rerun()

        st.markdown(
            """
            <div class="sidebar-footer">
                <span>Version 1.0</span>
                <span class="footer-sep">|</span>
                <span>© 2026 CS Division</span>
                <span class="footer-sep">|</span>
                <span>🔒 Internal Use Only</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # Year is intentionally not exposed as a filter.
    year = "All"

    # Apply filters
    f_hc = apply_filters(hc, year, month, office)
    f_workload = apply_filters(workload, year, month, office)
    f_fte = apply_filters(fte, year, month, office)
    f_shipment = apply_filters(shipment, year, month, office)
    f_mode = apply_filters(shipment_mode, year, month, office)
    f_customer = apply_filters(customer, year, month, office)
    f_customer_ns = apply_filters(customer_ns, year, month, office)
    f_resolution = apply_filters(resolution, year, month, office)
    # YVF follows Month/Year filters only when the source sheet contains Month.
    # Current office-only source remains valid without inventing a time dimension.
    if (
        yvf is not None
        and not yvf.empty
        and "MonthDate" in yvf.columns
        and yvf["MonthDate"].notna().any()
    ):
        f_yvf = apply_filters(yvf, year, month, office)
    else:
        f_yvf = filter_office_only(yvf, office)

    f_core_detail = apply_filters(core_detail, year, month, office)
    f_ancillary_detail = apply_filters(ancillary_detail, year, month, office)
    f_supporting_detail = apply_filters(supporting_detail, year, month, office)
    f_exception_detail = apply_filters(exception_detail, year, month, office)

    st.markdown(
        f"""
        <div class="main-header">
            <div class="main-title">{APP_TITLE}</div>
            <div class="subtitle">{APP_SUBTITLE}</div>
            <div class="filter-summary-card">
                <div class="filter-summary-item">
                    <div class="filter-summary-icon">
                        <svg viewBox="0 0 24 24" aria-hidden="true">
                            <rect x="3" y="5" width="18" height="16" rx="2"></rect>
                            <path d="M8 3v4M16 3v4M3 10h18"></path>
                        </svg>
                    </div>
                    <div class="filter-summary-label">Selected Month</div>
                    <div class="filter-summary-value">{month}</div>
                </div>
                <div class="filter-summary-item">
                    <div class="filter-summary-icon">
                        <svg viewBox="0 0 24 24" aria-hidden="true">
                            <path d="M4 21V7l8-4 8 4v14"></path>
                            <path d="M8 21v-6h8v6M8 9h2M14 9h2M8 12h2M14 12h2"></path>
                        </svg>
                    </div>
                    <div class="filter-summary-label">Selected Office</div>
                    <div class="filter-summary-value">{office}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if workload.empty or fte.empty:
        st.markdown(
            """
            <div class="warning-box">
            WARNING: Một số dữ liệu workload hoặc CS FTE có thể chưa đầy đủ. Dashboard vẫn chạy dynamic và sẽ tự cập nhật khi bổ sung dữ liệu vào file nguồn.
            </div>
            """,
            unsafe_allow_html=True,
        )

    section_title("1. Workload & Capacity Utilization")

    # Section 1 uses the HC sheet as the single source of truth.
    approved_hc = weighted_period_avg(f_hc, "Total Approved HC") if not f_hc.empty else 0.0
    approved_mng = weighted_period_avg(f_hc, "Approved HC MNG") if not f_hc.empty else 0.0
    approved_pic = weighted_period_avg(f_hc, "Approved HC PIC") if not f_hc.empty else 0.0

    actual_hc = weighted_period_avg(f_hc, "Total Actual HC") if not f_hc.empty else 0.0
    actual_mng = weighted_period_avg(f_hc, "Actual HC MNG") if not f_hc.empty else 0.0
    actual_pic = weighted_period_avg(f_hc, "Actual HC PIC") if not f_hc.empty else 0.0

    required_hc = weighted_period_avg(f_hc, "Total Required HC") if not f_hc.empty else 0.0
    required_mng = weighted_period_avg(f_hc, "Required HC MNG") if not f_hc.empty else 0.0
    required_pic = weighted_period_avg(f_hc, "Required HC PIC") if not f_hc.empty else 0.0

    hc_variance = required_hc - actual_hc
    
    if hc_variance > 0:
        variance_status = ("OVERLOAD", COLORS["red"], "#FEE2E2")
    elif hc_variance < 0:
        variance_status = ("REDUNDANT", COLORS["green"], "#DCFCE7")
    else:
        variance_status = ("BALANCED", COLORS["blue"], "#E0F2FE")

    hc1, hc2, hc3, hc4 = st.columns(4, gap="medium")

    with hc1:
        hc_detail_card(
            "Approved HC",
            approved_hc,
            approved_mng,
            approved_pic,
        )

    with hc2:
        hc_detail_card(
            "Actual HC",
            actual_hc,
            actual_mng,
            actual_pic,
        )

    with hc3:
        hc_detail_card(
            "Required HC",
            required_hc,
            required_mng,
            required_pic,
        )

    with hc4:
        hc_variance_card(
            "HC Gap",
            hc_variance,
            "Required HC − Actual HC",
            variance_status[0],
            variance_status[1],
            variance_status[2],
        )

    if office == "All Offices":
        render_hc_office_comparison(f_hc)

    # KPI cards follow Month + Office filters.
    # The line chart keeps all available months so management can see the HC trend.
    hc_trend_data = filter_office_only(hc, office)
    st.markdown('<div class="chart-box" style="margin-top:12px;">', unsafe_allow_html=True)
    chart_office_capacity_trend(hc_trend_data)

    section_title("2. Shipment Volume")

    # Source for both KPIs: Shipment volume sheet
    shipment_total = (
        float(f_shipment["Total Shipment"].sum())
        if not f_shipment.empty and "Total Shipment" in f_shipment.columns
        else 0.0
    )
    active_customers = calculate_active_customers(f_shipment)

    # KPI order requested:
    # 1) Active Customers
    # 2) Total Shipment Volume
    # Keep 2 empty columns so KPI widths remain consistent with Section 1.
    sk1, sk2, sk3, sk4 = st.columns(4, gap="medium")
    with sk1:
        shipment_kpi_card(
            "ACTIVE CUSTOMERS",
            fmt_int(active_customers),
            "",
        )
    with sk2:
        shipment_kpi_card(
            "TOTAL SHIPMENT VOLUME",
            fmt_int(shipment_total),
            "",
        )
    with sk3:
        st.empty()
    with sk4:
        st.empty()

    # Customer shipment analysis:
    # Remove Transportation Mode chart/detail from the dashboard.
    # Show Top 15 Customers chart and Customer Volume Detail on the same row.
    customer_chart_col, customer_detail_col = st.columns([0.58, 0.42], gap="medium")

    with customer_chart_col:
        chart_top_customers(f_customer_ns)

    with customer_detail_col:
        customer_detail_volume_table(f_customer_ns)


    section_title("3. Workload per PIC")

    # KPI source: sheet "2. FTE Workload".
    # Single source of truth for Section 3:
    #   (i)  Total Available Time
    #   (ii) Total Actual Working Time
    #        FTE Workload = ii / i
    #        FTE Workload Status
    #
    # Month = All:
    #   Calculate monthly office/PIC totals first, then sum all months for
    #   Total Available Time and Total Actual Working Time.
    # Selected month:
    #   Show the actual total of that selected month.

    if f_fte is not None and not f_fte.empty:
        fte_kpi = f_fte.copy()

        fte_kpi["Available Time"] = pd.to_numeric(
            fte_kpi["Available Time"], errors="coerce"
        )
        fte_kpi["Actual Working Time"] = pd.to_numeric(
            fte_kpi["Actual Working Time"], errors="coerce"
        )

        monthly_fte = (
            fte_kpi.dropna(
                subset=["MonthDate", "Available Time", "Actual Working Time"]
            )
            .groupby("MonthDate", as_index=False)
            .agg(
                Total_Available_Time=("Available Time", "sum"),
                Total_Actual_Working_Time=("Actual Working Time", "sum"),
            )
        )

        if not monthly_fte.empty:
            if str(month).strip().lower() == "all":
                total_available = float(
                    monthly_fte["Total_Available_Time"].sum()
                )
                total_actual_working = float(
                    monthly_fte["Total_Actual_Working_Time"].sum()
                )
            else:
                selected_month_row = monthly_fte.sort_values(
                    "MonthDate"
                ).iloc[-1]
                total_available = float(
                    selected_month_row["Total_Available_Time"]
                )
                total_actual_working = float(
                    selected_month_row["Total_Actual_Working_Time"]
                )

            fte_workload = safe_div(
                total_actual_working,
                total_available,
            )
            fte_status = status_from_util(fte_workload)
        else:
            total_available = float("nan")
            total_actual_working = float("nan")
            fte_workload = float("nan")
            fte_status = ("NO DATA", COLORS["muted"], COLORS["light_blue"])
    else:
        total_available = float("nan")
        total_actual_working = float("nan")
        fte_workload = float("nan")
        fte_status = ("NO DATA", COLORS["muted"], COLORS["light_blue"])

    # Four management KPIs in one row.
    p1, p2, p3, p4 = st.columns(4, gap="medium")

    def section3_kpi_card(label: str, value: str, note: str = ""):
        note_html = (
            f'<div class="pic-kpi-note">{note}</div>'
            if note else ""
        )
        st.markdown(
            f"""
            <div class="pic-kpi-card">
                <div class="pic-kpi-label">{label}</div>
                <div class="pic-kpi-value">{value}</div>
                {note_html}
            </div>
            """,
            unsafe_allow_html=True,
        )

    with p1:
        section3_kpi_card(
            "Total Available Time",
            fmt_num(total_available, 1)
            if not pd.isna(total_available) else "N/A",
            "95% × 8 × 22 × PIC (hour)",
        )

    with p2:
        section3_kpi_card(
            "Total Actual Working Time",
            fmt_num(total_actual_working, 1)
            if not pd.isna(total_actual_working) else "N/A",
            "C + A + S + E (hour)",
        )

    with p3:
        fte_value = (
            f"{fte_workload * 100:,.1f}%"
            if not pd.isna(fte_workload) else "N/A"
        )
        st.markdown(
            f"""
            <div class="pic-kpi-card">
                <div class="pic-kpi-label">Average PIC Workload</div>
                <div class="pic-kpi-value" style="
                    font-size:38px !important;
                    font-weight:800 !important;
                    line-height:1.05 !important;
                ">
                    {fte_value}
                </div>
                <div class="pic-kpi-note">Actual Time vs Available Time</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with p4:
        status_text, status_color, status_bg = fte_status
        st.markdown(
            f"""
            <div class="pic-kpi-card">
                <div class="pic-kpi-label">PIC Workload Status</div>
                <div style="
                    margin-top:0;
                    min-height:52px;
                    display:flex;
                    justify-content:center;
                    align-items:center;
                ">
            <span class="status-badge"
                style="
                    color:{status_color};
                    background:{status_bg};
                    font-size:30px !important;
                    line-height:1.05 !important;
                    font-weight:800 !important;
                    padding:10px 32px !important;
                    min-width:220px;
                    min-height:44px;
                    display:inline-flex;
                    align-items:center;
                    justify-content:center;
                    text-align:center;
                    border-radius:999px;
                ">
                {status_text}
            </span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    if office == "All Offices":
        render_fte_office_comparison(f_fte, month)

    # Chart source: 2. FTE Workload
    # PIC Workload = FTE Workload factor × Available Time / PIC.
    # Available Standard Time / PIC = 95% × 8 × 22 = 167.2 hours.
    # Therefore: PIC Workload = CS FTE coefficient × 167.2 hours.
    # When All Offices is selected, only overloaded PICs/offices are displayed.
    st.markdown('<div class="chart-box" style="margin-top:8px;">', unsafe_allow_html=True)
    chart_workload_by_pic(f_fte, office)

    section_title("4. Workload Distribution by Segment")

    segment_summary = build_segment_workload(f_workload, f_mode)
    segment_total_hours = (
        float(segment_summary["Allocation Time (h)"].sum())
        if not segment_summary.empty
        else 0.0
    )

    # Executive one-row layout:
    # Left = compact Segment Summary panel
    # Right = Workload by Segment chart
    # Detail table remains full width below.
    if not segment_summary.empty:
        _seg_rank = segment_summary.sort_values(
            "Workload Share", ascending=False
        ).reset_index(drop=True)
        _top3 = _seg_rank.head(3).copy()
        _rank_colors = [COLORS["amber"], COLORS["blue"], "#8EB7D8"]
        top3_rows_html = "".join(
            f"""
            <div style="display:flex;align-items:center;justify-content:space-between;padding:7px 0;border-bottom:{'0' if idx == len(_top3) - 1 else '1px solid #EEF2F6'};">
                <div style="display:flex;align-items:center;gap:10px;">
                    <span style="width:22px;height:22px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;background:{_rank_colors[idx]};color:#FFFFFF;font-size:10px;font-weight:800;">{idx + 1}</span>
                    <span style="color:#003B70;font-size:13px;font-weight:700;">{html.escape(str(row['Segment']))}</span>
                </div>
                <span style="color:#003B70;font-size:13px;font-weight:700;">{float(row['Workload Share']):.1%}</span>
            </div>
            """
            for idx, (_, row) in enumerate(_top3.iterrows())
        )
    else:
        top3_rows_html = '<div style="color:#98A2B3;font-size:12px;padding:8px 0;">No data available</div>'

    seg_summary_col, seg_chart_col = st.columns([0.32, 0.68], gap="medium")

    with seg_summary_col:
        summary_html = f"""
<div style="height:340px;min-height:340px;box-sizing:border-box;display:flex;flex-direction:column;justify-content:center;background:#FFFFFF;border:1px solid #D8E1EA;border-radius:12px;box-shadow:0 1px 4px rgba(16,24,40,0.045);padding:24px 26px;">
    <div style="color:#667085;font-size:12px;line-height:1.25;font-weight:600;letter-spacing:0.025em;text-transform:uppercase;">Total Workload Hours</div>
  <div style="color:#003B70;font-size:34px;line-height:1.05;font-weight:700;letter-spacing:-0.02em;margin-top:8px;">{fmt_num(segment_total_hours, 1)}</div>
  <div style="color:#667085;font-size:11px;margin-top:6px;"></div>
  <div style="height:1px;background:#E6ECF2;margin:20px 0 14px 0;"></div>
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:3px;">
    <span style="color:#667085;font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:0.025em;">Top 3 Leading</span>
    <span style="color:#98A2B3;font-size:10px;font-weight:600;text-transform:uppercase;">Workload Share</span>
  </div>
  <div>{top3_rows_html}</div>
  </div>
"""
        # Remove leading indentation so Markdown does not treat nested HTML
        # as a fenced/code block and display the raw <div> tags.
        summary_html = "\n".join(
            line.lstrip() for line in summary_html.splitlines()
        )
        st.markdown(summary_html.strip(), unsafe_allow_html=True)

    with seg_chart_col:
        chart_service_matrix(f_workload, f_mode)

    # Full-width detail table below the executive row.
    segment_workload_table(f_workload, f_mode)

    
    section_title("5. Workload Breakdown by Activity & Segment")

    st.markdown(
        """
        <div style="
            color:#667085;
            font-size:12px;
            line-height:1.5;
            margin:0 0 10px 2px;">
            Workload composition: Core Service (C), Ancillary Service (A),
            Supporting Activity (S) and Exception Handling (E).
        </div>
        """,
        unsafe_allow_html=True,
    )

    # C/A/S/E summary cards by Office — same executive idea as the HC office cards.
    render_case_office_cards(f_workload)

    # Summary table + C/A/S/E allocation chart
    wb_chart, wb_table = st.columns([0.43, 0.57], gap="medium")
    with wb_chart:
        chart_case_allocation(f_workload)
    with wb_table:
        render_workload_breakdown_table(f_workload)

    # Four source-detail tables
    st.markdown(
        f"""
        <div style="
            color:{COLORS['navy']};
            font-size:{UI['chart_title_size']}px;
            font-weight:700;
            margin:14px 0 8px 2px;">
            C / A / S / E Details
        </div>
        """,
        unsafe_allow_html=True,
    )

    casetab_c, casetab_a, casetab_s, casetab_e = st.tabs([
        "C · Core Service",
        "A · Ancillary Service",
        "S · Supporting Activity",
        "E · Exception Handling",
    ])
    with casetab_c:
        render_activity_detail_table(f_core_detail, "Core Service")
    with casetab_a:
        render_activity_detail_table(f_ancillary_detail, "Ancillary Service")
    with casetab_s:
        render_activity_detail_table(f_supporting_detail, "Supporting Activity")
    with casetab_e:
        exception_months = []
        if (
            f_exception_detail is not None
            and not f_exception_detail.empty
            and "MonthDate" in f_exception_detail.columns
        ):
            exception_months = (
                pd.to_datetime(
                    f_exception_detail["MonthDate"], errors="coerce"
                )
                .dropna()
                .drop_duplicates()
                .sort_values()
                .dt.strftime("%b-%y")
                .tolist()
            )

        # Keep the month note outside the two-column row. This lets the chart
        # and dataframe start from exactly the same vertical coordinate.
        if exception_months:
            months_text = "Months with data: " + ", ".join(exception_months)
            st.markdown(
                f'<div style="margin:-12px 0 4px 41.5%; color:#98A2B3; '
                f'font-size:11px; line-height:16px;">{months_text}</div>',
                unsafe_allow_html=True,
            )

        exception_chart_col, exception_table_col = st.columns(
            [0.40, 0.60], gap="medium"
        )
        with exception_chart_col:
            chart_exception_by_criteria(f_exception_detail)
        with exception_table_col:
            render_activity_detail_table(
                f_exception_detail,
                "Exception Handling",
                table_height=455,
                show_months_caption=False,
                stretch_to_container=True,
            )

    section_title("6. Control Tower effectiveness = CS Resolutions Rate")

    # Definition note — UI only; no business logic/calculation changes.
    st.markdown(
        """
        <div style="
            background:#F8FAFC;
            border:1px solid #D5E1EA;
            border-left:4px solid #0DBAEE;
            border-radius:10px;
            padding:10px 14px;
            margin:0 0 12px 0;
            color:#C2410C;
            font-size:14px;
            line-height:1.5;
        ">
            <strong style="color:#C2410C;">Definition:</strong>
            CS Resolution Rate = % of customer requests/issues fully resolved by CS without handing off the resolution to another BU.
            Requests merely forwarded to BUs and relayed back to customers are not considered CS resolutions.
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Executive KPIs sourced from sheet "CS Resolutions Rate".
    if not f_resolution.empty:
        total_abn = float(f_resolution["Total Abnormality"].sum())
        resolved = float(f_resolution["Resolved"].sum())
        rate = safe_div(resolved, total_abn)

        cs1, cs2, cs3 = st.columns(3, gap="medium")
        with cs1:
            kpi_card(
                "Total Exception Case",
                fmt_int(total_abn),
                "",
            )
        with cs2:
            kpi_card(
                "Resolved by CS",
                fmt_int(resolved),
                "",
            )
        with cs3:
            st.markdown(
                f"""
                <div class="kpi-card">
                    <div class="kpi-label">CS Resolution Rate</div>
                    <div class="kpi-value" style="
                        font-size:38px !important;
                        font-weight:800 !important;
                        line-height:1.05 !important;
                    ">
                        {fmt_pct(rate)}
                    </div>
                    <div class="kpi-note">
                        {fmt_int(resolved)} / {fmt_int(total_abn)} cases
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    cs_chart, cs_table = st.columns([0.55, 0.45], gap="medium")
    with cs_chart:
        chart_resolution(f_resolution)
    with cs_table:
        render_cs_solution_table(f_resolution)

    section_title("7. YVF Promotion Effectiveness")

    # Show only one common message when no YVF data is available.
    # Chart and detail table are rendered only when filtered YVF data exists.
    if f_yvf.empty:
        st.info("No YVF data available for selected filters.")
    else:
        yvf_booking = float(f_yvf["YVF Booking"].sum())
        iff = float(f_yvf["IFF Shipment"].sum())
        yvf_rate = safe_div(yvf_booking, iff)

        y1, y2, y3 = st.columns(3, gap="medium")
        with y1:
            kpi_card(
                "Total YVF Bookings",
                fmt_int(yvf_booking),
                "",
            )
        with y2:
            kpi_card(
                "Total IFF Bookings",
                fmt_int(iff),
                "",
            )
        with y3:
            st.markdown(
                f"""
                <div class="kpi-card">
                    <div class="kpi-label">YVF Booking Ratio</div>
                    <div class="kpi-value" style="
                        font-size:38px !important;
                        font-weight:800 !important;
                        line-height:1.05 !important;
                    ">
                        {fmt_pct(yvf_rate)}
                    </div>
                    <div class="kpi-note">
                        {fmt_int(yvf_booking)} / {fmt_int(iff)}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        yvf_chart_col, yvf_table_col = st.columns([0.52, 0.48], gap="medium")
        with yvf_chart_col:
            chart_yvf(f_yvf)
        with yvf_table_col:
            render_yvf_table(f_yvf)



# ============================================================
# SIDEBAR MICRO-POLISH FINAL
# UI ONLY — no changes to filters, upload logic, session state,
# calculations, charts, or main dashboard layout.
# ============================================================
st.markdown(
    """
    <style>
    /* Sidebar only */
    section[data-testid="stSidebar"] > div:first-child {
        padding-top: 0.35rem !important;
        padding-left: 0.90rem !important;
        padding-right: 0.90rem !important;
        padding-bottom: 0.80rem !important;
    }

    /* HOME: lower height and tighter spacing */
    section[data-testid="stSidebar"] div[data-testid="stButton"] > button {
        min-height: 41px !important;
        height: 41px !important;
        margin: 0 0 9px 0 !important;
        border-radius: 8px !important;
        font-size: 12.5px !important;
    }

    /* FILTERS block */
    .sidebar-filter-title {
        margin: 3px 0 1px 0 !important;
        font-size: 16px !important;
        gap: 7px !important;
    }

    .sidebar-filter-title::after {
        width: 20px !important;
        height: 2px !important;
    }

    .sidebar-filter-caption {
        display: none !important;
    }

    /* Labels */
    section[data-testid="stSidebar"] label {
        font-size: 12px !important;
        margin-bottom: 4px !important;
    }

    /* Select boxes */
    section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
        min-height: 41px !important;
        height: 41px !important;
        border-radius: 8px !important;
    }

    section[data-testid="stSidebar"] [data-testid="stSelectbox"] {
        margin-bottom: 0.20rem !important;
    }

    /* Divider */
    section[data-testid="stSidebar"] hr {
        margin: 9px 0 9px 0 !important;
    }

    /* Upload title */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] {
        margin-top: 0 !important;
    }

    /* Compact upload card */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section {
        min-height: 78px !important;
        padding: 7px 8px !important;
        border-radius: 8px !important;
    }

    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section * {
        font-size: 10.8px !important;
        line-height: 1.25 !important;
    }

    section[data-testid="stSidebar"] [data-testid="stFileUploader"] button {
        min-height: 32px !important;
        height: 32px !important;
        padding: 0 12px !important;
        border-radius: 7px !important;
        font-size: 11.5px !important;
    }

    /* Remove visual clutter from uploader help icon where possible */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] [data-testid="stTooltipIcon"] {
        opacity: 0.55 !important;
        transform: scale(0.88);
    }

    /* Footer */
    .sidebar-footer {
        margin-top: 12px !important;
        padding-top: 10px !important;
        font-size: 10px !important;
        line-height: 1.25 !important;
        gap: 4px !important;
    }

    /* Laptop */
    @media (max-width:1366px) {
        section[data-testid="stSidebar"] > div:first-child {
            padding-top: 0.28rem !important;
            padding-left: 0.85rem !important;
            padding-right: 0.85rem !important;
        }

        section[data-testid="stSidebar"] div[data-testid="stButton"] > button {
            min-height: 40px !important;
            height: 40px !important;
        }

        section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
            min-height: 40px !important;
            height: 40px !important;
        }

        .sidebar-footer {
            font-size: 9.8px !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SIDEBAR POSITIONING FINAL — UI ONLY
# HOME icon / filter breathing room / lower upload & footer
# ============================================================
st.markdown(
    """
    <style>
    /* Give FILTERS a little more breathing room before MONTH */
    .sidebar-filter-spacer {
        height: 14px !important;
        min-height: 14px !important;
    }

    /* Push Upload + Version area lower on normal laptop screens.
       This is visual spacing only; upload/filter logic is unchanged. */
    .sidebar-bottom-anchor {
        height: clamp(150px, 28vh, 360px) !important;
        min-height: 150px !important;
    }

    /* Keep the lower block visually compact once it reaches the bottom area */
    section[data-testid="stSidebar"] hr {
        margin-top: 8px !important;
        margin-bottom: 12px !important;
    }

    .sidebar-footer {
        margin-top: 26px !important;
        padding-top: 12px !important;
    }

    /* On shorter screens, reduce the spacer automatically to avoid clipping */
    @media (max-height: 760px) {
        .sidebar-bottom-anchor {
            height: 95px !important;
            min-height: 95px !important;
        }
        .sidebar-footer {
            margin-top: 20px !important;
        }
    }

    @media (max-height: 650px) {
        .sidebar-bottom-anchor {
            height: 48px !important;
            min-height: 48px !important;
        }
        .sidebar-footer {
            margin-top: 14px !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


if __name__ == "__main__":
    main()

