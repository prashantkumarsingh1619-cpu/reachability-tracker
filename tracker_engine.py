#!/usr/bin/env python3
"""
Tracker Engine: Core extraction and Excel generation logic for Reachability Tracker.
Supports both file paths and in-memory file buffers (BytesIO / Streamlit UploadedFile).
Includes multi-circle profile support (e.g. Karnal/Haryana, Mohali/Punjab).
"""
import io
import os
import re
import warnings
from typing import Any, Dict, List, Optional, Tuple, Union

import openpyxl
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

# Suppress harmless openpyxl drawing warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =============================================================================
# DEFAULT CONFIGURATION & CIRCLE PROFILES
# =============================================================================
DEFAULT_DCGW_VPRN = {
    "ntp": 1000, "sx": 1004, "s11": 1004, "gx": 1000, "gy": 1000, "bp": 1000, "li_cp": 1000,
    "s2bc": 1004, "s6b": 1000, "s5s8_c": 1000, "snoc": 1000, "li_up": 1000, "s1u": 1004,
    "sgi": 1001, "s5s8_u": 1004, "cgnat": 1004, "gm": 1003, "s2bu": 1004,
}

DEFAULT_S5S8_C_PLACEHOLDERS = {"SGW-C": 11, "PGW-C": 12}
DEFAULT_S5S8_U_PAD_TO = {"SGW-U": 23, "PGW-U": 41}
S5S8_NOTE = "CMG - S5-S8 Circle level to be checked seperately"
NTP_NOTE = "NA, as NTP sync with 192.168.130.1 & .2"

NATIONAL_NTP = ["192.168.130.1", "192.168.130.2"]

NATIONAL_BP_MEDIATION = [
    "10.94.161.129", "10.94.161.106", "10.94.161.107", "10.87.92.29", "10.94.161.124",
    "10.94.161.74", "10.94.161.100", "10.87.92.17", "10.94.161.181", "10.94.161.114",
    "10.94.161.115", "10.87.92.41"
]

NATIONAL_SNOC = [
    ("VA Scan", "10.19.72.124"), ("VA Scan", "10.19.72.125"), ("VA Scan", "10.40.55.185"),
    ("VA Scan", "10.40.55.167"), ("VA Scan", "10.40.55.168"), ("MBSS", "172.18.130.231"),
    ("MBSS", "172.18.130.232"), ("MBSS", "172.18.130.237"), ("Netact AP", "10.80.246.187"),
    ("Netact TN", "10.80.242.187"), ("ViPAM", "10.40.55.131"), ("ViPAM", "10.40.55.214"),
    ("SEVONE", "10.241.46.139"), ("TeMIP", "10.114.8.231"), ("MyCOM", "10.115.1.36"),
    ("UIM", "10.241.49.105"), ("UIM", "10.241.44.142"), ("APM", "10.19.75.221"),
    ("APM", "10.19.75.222"), ("APM", "10.19.32.103"), ("APM", "10.19.72.122"),
    ("SIEM", "10.49.161.12")
]

NATIONAL_SNOC_CPF = [
    ("VA Scan", "10.19.72.124"), ("VA Scan", "10.19.72.125"), ("VA Scan", "10.40.55.185"),
    ("VA Scan", "10.40.55.167"), ("VA Scan", "10.40.55.168"), ("MBSS", "172.18.130.231"),
    ("MBSS", "172.18.130.232"), ("MBSS", "172.18.130.237"), ("Netact ", "10.80.246.187"),
    ("Netact ", "10.80.242.187"), ("ViPAM", "10.40.55.131"), ("ViPAM", "10.40.55.214"),
    ("SEVONE", "10.241.46.139"), ("TeMIP", "10.114.8.231"), ("MyCOM", "10.115.1.36"),
    ("UIM", "10.241.49.105"), ("UIM", "10.241.44.142"), ("APM", "10.19.75.221"),
    ("APM", "10.19.75.222"), ("APM", "10.19.32.103"), ("APM", "10.19.72.122"),
    ("SIEM", "10.49.161.12")
]

LUCKNOW_SNOC_CPF = [
    ("VA Scan", "10.19.72.124"), ("VA Scan", "10.19.72.125"), ("VA Scan", "10.40.55.185"),
    ("VA Scan", "10.40.55.167"), ("VA Scan", "10.40.55.168"), ("MBSS", "172.18.130.231"),
    ("MBSS", "172.18.130.232"), ("MBSS", "172.18.130.237"), ("Netact ", "100.119.136.187"),
    ("Netact ", "10.80.246.187"), ("ViPAM", "10.40.55.131"), ("ViPAM", "10.40.55.214"),
    ("SEVONE", "10.241.46.139"), ("TeMIP", "10.114.8.231"), ("MyCOM", "10.115.1.36"),
    ("UIM", "10.241.49.105"), ("UIM", "10.241.44.142"), ("APM", "10.19.75.221"),
    ("APM", "10.19.75.222"), ("APM", "10.19.32.103"), ("APM", "10.19.72.122"),
    ("SIEM", "10.49.161.12")
]

LUCKNOW_SNOC = [
    ("VA Scan", "10.19.72.124"), ("VA Scan", "10.19.72.125"), ("VA Scan", "10.40.55.185"),
    ("VA Scan", "10.40.55.167"), ("VA Scan", "10.40.55.168"), ("MBSS", "172.18.130.231"),
    ("MBSS", "172.18.130.232"), ("MBSS", "172.18.130.237"), ("Netact AP", "100.119.136.187"),
    ("Netact TN", "10.80.246.187"), ("ViPAM", "10.40.55.131"), ("ViPAM", "10.40.55.214"),
    ("SEVONE", "10.241.46.139"), ("TeMIP", "10.114.8.231"), ("MyCOM", "10.115.1.36"),
    ("UIM", "10.241.49.105"), ("UIM", "10.241.44.142"), ("APM", "10.19.75.221"),
    ("APM", "10.19.75.222"), ("APM", "10.19.32.103"), ("APM", "10.19.72.122"),
    ("SIEM", "10.49.161.12")
]

CIRCLE_PROFILES = {
    "karnal": {
        "ntp": NATIONAL_NTP,
        "s1u_test_enb": "100.91.17.85",
        "sgi_dst_v6": ["2402:3a80:1fff:43::658", "2402:3a80:1fff:43::659", "2001:4860:4860::8888"],
        "sgi_dst_v4": ["10.216.205.216", "10.216.205.217", "8.8.8.8"],
        "cgnat_dst": ["10.216.206.148", "10.216.206.149"],
        "bp_mediation": NATIONAL_BP_MEDIATION,
        "snoc_cpf": NATIONAL_SNOC_CPF,
        "snoc": NATIONAL_SNOC,
        "sgi_pool": 2,
        "sgi_v4_line": 1,
        "gm_v6_line": 1,
        "s6b_label": "s6b",
        "snoc_labels": ("SNOC (CMG-C)", "SNOC (CMG-UPF01)", "SNOC (CMG-UPF02)"),
    },
    "mohali": {
        "ntp": NATIONAL_NTP,
        "s1u_test_enb": "100.121.25.186",
        "sgi_dst_v6": ["2001:4860:4860::8888", "2402:3a80:1072::2:e08", "2402:3a80:1072::2:e09"],
        "sgi_dst_v4": ["8.8.8.8", "10.217.223.248", "10.217.223.249"],
        "cgnat_dst": ["10.253.199.212", "10.253.199.213"],
        "bp_mediation": [
            "10.94.161.129", "10.87.92.41", "10.94.161.181", "10.94.161.124",
            "10.94.161.114", "10.94.161.115", "10.87.92.29", "10.94.161.74",
            "10.94.161.181", "10.94.161.100", "10.94.161.106", "10.94.161.107"
        ],
        "snoc_cpf": [
            ("VA Scan", "172.18.129.162"), ("VA Scan", "10.19.72.124"), ("VA Scan", "10.19.72.125"),
            ("VA Scan", "10.40.55.185"), ("VA Scan", "10.94.76.15"), ("VA Scan", "172.18.129.28"),
            ("VA Scan", "172.18.129.29"), ("MBSS", "172.18.130.231"), ("MBSS", "172.18.130.232"),
            ("MBSS", "172.18.130.237"), ("Netact ", "10.80.246.187"), ("Netact ", "10.80.242.187"),
            ("ViPAM", "10.40.55.131"), ("ViPAM", "10.40.55.214"), ("SEVONE", "10.241.46.139"),
            ("TeMIP", "10.114.8.231"), ("MyCOM", "10.115.1.36"), ("UIM", "10.241.49.105"),
            ("UIM", "10.241.44.142"), ("APM", "10.19.75.221"), ("APM", "10.19.75.222"),
            ("APM", "10.19.32.103"), ("APM", "10.19.72.122"), ("SIEM", "10.94.150.29")
        ],
        "snoc": [
            ("VA Scan", "172.18.129.162"), ("VA Scan", "10.19.72.124"), ("VA Scan", "10.19.72.125"),
            ("VA Scan", "10.40.55.185"), ("VA Scan", "10.94.76.15"), ("VA Scan", "172.18.129.28"),
            ("VA Scan", "172.18.129.29"), ("MBSS", "172.18.130.231"), ("MBSS", "172.18.130.232"),
            ("MBSS", "172.18.130.237"), ("Netact AP", "10.80.246.187"), ("Netact TN", "10.80.242.187"),
            ("ViPAM", "10.40.55.131"), ("ViPAM", "10.40.55.214"), ("SEVONE", "10.241.46.139"),
            ("TeMIP", "10.114.8.231"), ("MyCOM", "10.115.1.36"), ("UIM", "10.241.49.105"),
            ("UIM", "10.241.44.142"), ("APM", "10.19.75.221"), ("APM", "10.19.75.222"),
            ("APM", "10.19.32.103"), ("APM", "10.19.72.122"), ("SIEM", "10.94.150.29")
        ],
        "sgi_pool": 1,
        "sgi_v4_line": 0,
        "gm_v6_line": 0,
        "s6b_label": "S6B",
        "snoc_labels": ("SNOC", "SNOC", "SNOC"),
    },
    "lucknow": {
        "ntp": NATIONAL_NTP,
        "s1u_test_enb": "100.121.239.250",
        "sgi_dst_v6": [
            "2402:3a80:1fff:43::638",
            "2402:3a80:1fff:43::639",
            "2001:4860:4860::8888"
        ],
        "sgi_dst_v4": ["10.216.204.40", "10.216.204.41", "8.8.8.8"],
        "cgnat_dst": ["10.216.204.84", "10.216.204.85"],
        "bp_mediation": NATIONAL_BP_MEDIATION,
        "snoc_cpf": LUCKNOW_SNOC_CPF,
        "snoc": LUCKNOW_SNOC,
        "sgi_pool": 1,
        "sgi_v4_line": 0,
        "gm_v6_line": 0,
        "s6b_label": "S6b",
        "snoc_labels": ("", "", ""),
        "s11_mme": [
            "203.88.15.121",
            "203.88.0.73",
            "203.88.21.113",
            "203.88.21.121",
            "112.110.107.28"
        ],
        "gx": [
            "10.103.101.253",
            "10.103.101.255",
            "10.100.130.132",
            "10.100.130.133"
        ],
        "gy": [
            "10.210.32.199",
            "10.210.32.200",
            "10.210.32.201",
            "10.210.32.202",
            "10.210.32.203",
            "10.151.248.55",
            "10.151.248.56",
            "10.151.248.57",
            "10.151.248.58",
            "10.151.248.99"
        ],
        "recc_list": [
            "172.19.223.120",
            "172.19.228.120"
        ],
        "s6b": [
            "10.52.237.40",
            "10.58.162.203"
        ],
        "li_x1": [
            "10.164.6.153",
            "10.164.6.13",
            "10.164.6.149"
        ],
        "li_x2": "10.164.6.146",
        "li_x3": "10.164.6.146",
        "s2bc": [
            "203.88.5.82",
            "112.110.104.231"
        ],
        "s2bu": [
            "203.88.5.83",
            "112.110.104.255"
        ],
        "s5s8_c_dest_none": True,
        "gm": [
            "2402:3a80:2000:a684::4",
            "2402:3a80:2000:a684::5",
            "2402:3a80:2000:a688::4",
            "2402:3a80:2000:a688::5",
            "2402:3a80:2000:a698::4",
            "2402:3a80:2000:a698::5",
            "2402:3a80:2000:a69c::4",
            "2402:3a80:2000:a69c::5",
            "2402:3a80:2000:a6a0::4",
            "2402:3a80:2000:a6a0::5",
            "2400:5200:a100:23e::6",
            "2400:5200:a100:23e::7",
            "2400:5200:a100:252::6",
            "2400:5200:a100:252::7",
            "2400:5200:a100:266::6",
            "2400:5200:a100:266::7",
            "2400:5200:a100:27a::6",
            "2400:5200:a100:27a::7",
            "2400:5200:a100:28e::6",
            "2400:5200:a100:28e::7",
            "2402:3a80:2000:a684::1",
            "2402:3a80:2000:a685::1",
            "2402:3a80:2000:a686::1",
            "2402:3a80:2000:a687::1",
            "2402:3a80:2000:a688::1",
            "2402:3a80:2000:a689::1",
            "2402:3a80:2000:a68a::1",
            "2402:3a80:2000:a68b::1",
            "2402:3a80:2000:a698::1",
            "2402:3a80:2000:a699::1",
            "2402:3a80:2000:a69a::1",
            "2402:3a80:2000:a69b::1",
            "2402:3a80:2000:a69c::1",
            "2402:3a80:2000:a69d::1",
            "2402:3a80:2000:a69e::1",
            "2402:3a80:2000:a69f::1",
            "2402:3a80:2000:a6a0::1",
            "2402:3a80:2000:a6a1::1",
            "2402:3a80:2000:a6a1::1",
            "2402:3a80:2000:a6a2::1",
            "2402:3a80:2000:a6a3::1",
            "2400:5200:a100:023e::1",
            "2400:5200:a100:0249::1",
            "2400:5200:a100:024a::1",
            "2400:5200:a100:024b::1",
            "2400:5200:a100:0252::1",
            "2400:5200:a100:025d::1",
            "2400:5200:a100:025e::1",
            "2400:5200:a100:025f::1",
            "2400:5200:a100:0266::1",
            "2400:5200:a100:0271::1",
            "2400:5200:a100:0272::1",
            "2400:5200:a100:0273::1",
            "2400:5200:a100:027a::1",
            "2400:5200:a100:0285::1",
            "2400:5200:a100:0286::1",
            "2400:5200:a100:0287::1",
            "2400:5200:a100:028e::1",
            "2400:5200:a100:0299::1",
            "2400:5200:a100:029a::1",
            "2400:5200:a100:029b::1"
        ]
    },
    "delhi": {
        "ntp": NATIONAL_NTP,
        "s1u_test_enb": "100.100.1.1",
        "sgi_dst_v6": ["2001:4860:4860::8888"],
        "sgi_dst_v4": ["8.8.8.8"],
        "cgnat_dst": [],
        "bp_mediation": NATIONAL_BP_MEDIATION,
        "snoc_cpf": NATIONAL_SNOC_CPF,
        "snoc": NATIONAL_SNOC,
        "sgi_pool": None,
        "sgi_v4_line": 0,
        "gm_v6_line": 0,
        "s6b_label": "S6B",
        "snoc_labels": ("SNOC (CMG-C)", "SNOC (CMG-UPF01)", "SNOC (CMG-UPF02)"),
    },
    "gujarat": {
        "ntp": NATIONAL_NTP,
        "s1u_test_enb": "100.91.17.85",
        "sgi_dst_v6": ["2001:4860:4860::8888"],
        "sgi_dst_v4": ["8.8.8.8"],
        "cgnat_dst": [],
        "bp_mediation": NATIONAL_BP_MEDIATION,
        "snoc_cpf": NATIONAL_SNOC_CPF,
        "snoc": NATIONAL_SNOC,
        "sgi_pool": None,
        "sgi_v4_line": 0,
        "gm_v6_line": 0,
        "s6b_label": "S6B",
        "snoc_labels": ("SNOC (CMG-C)", "SNOC (CMG-UPF01)", "SNOC (CMG-UPF02)"),
    },
    "generic": {
        "ntp": NATIONAL_NTP,
        "s1u_test_enb": "100.91.17.85",
        "sgi_dst_v6": ["2001:4860:4860::8888"],
        "sgi_dst_v4": ["8.8.8.8"],
        "cgnat_dst": [],
        "bp_mediation": NATIONAL_BP_MEDIATION,
        "snoc_cpf": NATIONAL_SNOC_CPF,
        "snoc": NATIONAL_SNOC,
        "sgi_pool": None,
        "sgi_v4_line": 0,
        "gm_v6_line": 0,
        "s6b_label": "S6b",
        "snoc_labels": ("SNOC (CMG-C)", "SNOC (CMG-UPF01)", "SNOC (CMG-UPF02)"),
        "s11_mme": ["203.88.15.121", "203.88.0.73", "203.88.21.113", "203.88.21.121", "112.110.107.28"],
        "gx": ["10.103.101.253", "10.103.101.255", "10.100.130.132", "10.100.130.133"],
        "gy": ["10.210.32.199", "10.210.32.200", "10.210.32.201", "10.210.32.202", "10.210.32.203"],
        "recc_list": ["172.19.223.120", "172.19.228.120"],
        "s6b": ["10.52.237.40", "10.58.162.203"],
        "li_x1": ["10.164.6.153", "10.164.6.13", "10.164.6.149"],
        "li_x2": "10.164.6.146",
        "li_x3": "10.164.6.146",
        "s2bc": ["203.88.5.82", "112.110.104.231"],
        "s2bu": ["203.88.5.83", "112.110.104.255"],
        "gm": [
            "2402:3a80:2000:a684::4", "2402:3a80:2000:a684::5", "2402:3a80:2000:a688::4",
            "2402:3a80:2000:a688::5", "2402:3a80:2000:a698::4", "2402:3a80:2000:a698::5"
        ]
    }
}

DEFAULT_STATIC = CIRCLE_PROFILES["generic"]


def get_circle_profile(site_name: str) -> Dict[str, Any]:
    name_low = site_name.lower().replace("-", " ").replace("_", " ")
    if "mohali" in name_low or "punjab" in name_low:
        return CIRCLE_PROFILES["mohali"].copy()
    if "karnal" in name_low or "haryana" in name_low:
        return CIRCLE_PROFILES["karnal"].copy()
    if "ashok" in name_low or "lucknow" in name_low or "up east" in name_low or "upeast" in name_low or "up-east" in name_low:
        return CIRCLE_PROFILES["lucknow"].copy()
    if "delhi" in name_low or "dl" in name_low.split():
        return CIRCLE_PROFILES["delhi"].copy()
    if "gujarat" in name_low or "ahmedabad" in name_low or "gj" in name_low.split():
        return CIRCLE_PROFILES["gujarat"].copy()
    if "mumbai" in name_low:
        prof = CIRCLE_PROFILES["generic"].copy()
        prof["s6b_label"] = "s6b"
        return prof
    return CIRCLE_PROFILES["generic"].copy()


SUMMARY_CONFIG = [
    ("First Call", "S11", ["cpf.s11"]),
    ("First Call", "S1-U UPF01", ["upf1.s1u"]),
    ("First Call", "S1-U UPF02", ["upf2.s1u"]),
    ("First Call", "Sx", ["cpf.sx"]),
    ("First Call", "Gx", ["cpf.gx"]),
    ("First Call", "Gy", ["cpf.gy"]),
    ("First Call", "SGi", ["upf1.sgi"]),
    ("F-UAT Completion", "LI CPF", ["cpf.li"]),
    ("F-UAT Completion", "LI UPF", ["upf1.li"]),
    ("F-UAT Completion", "BP", ["cpf.bp"]),
    ("VolTE First Call", "Gm", ["upf2.gm"]),
    ("F-UAT Completion", "S6b", ["cpf.s6b"]),
    ("F-UAT Completion", "S2b-u", ["upf2.s2bu"]),
    ("F-UAT Completion", "S5-S8 Local Circle (CPF)", ["cpf.sgwc", "cpf.pgwc"]),
    ("F-UAT Completion", "SNOC (CMG-C)", ["cpf.snoc"]),
    ("F-UAT Completion", "SNOC (CMG-UPF01)", ["upf1.snoc"]),
    ("F-UAT Completion", "SNOC (CMG-UPF02)", ["upf2.snoc"]),
]

FONT_NAME = "Calibri"
THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
GREEN_FILL = PatternFill("solid", fgColor="92D050")
CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_ALIGN = Alignment(horizontal="left", vertical="center")

HEADERS = [
    "Node", "Interface", "CMG VPRN", "Destination Node Name", "CMG SRC Address", "DCGW VPRN",
    "DST Address", "CMG Status", "", "Trace last hop from CMG", "CMG Remark", "DCGW Status",
    "Trace last hop from DCGW", "DCGW Remark", "AO", "CMG Ping Script", "CMG Traceroute Script",
    "DCGW VPRN", "DCGW Traceroute Script"
]

COL_WIDTHS = {
    "A": 10, "B": 24, "C": 23, "D": 28, "E": 33, "F": 14, "G": 39, "H": 13, "I": 18, "J": 30,
    "K": 17, "L": 13, "M": 26, "N": 42, "O": 12, "P": 68, "Q": 64, "R": 12, "S": 93
}

CMG_FIRST, DCGW_FIRST = 4, 27
DETAIL_TITLE_ROW, DETAIL_HEAD_ROW, DETAIL_FIRST = 47, 48, 49


# =============================================================================
# HELPERS
# =============================================================================
def txt(v: Any) -> str:
    return "" if v is None else str(v).replace("\xa0", " ").replace("\t", " ").strip()


def ip(v: Any) -> str:
    """Clean an address cell: extract first line if multi-line, trim whitespace and drop trailing /prefix."""
    s = txt(v)
    if not s:
        return ""
    lines = [l.strip() for l in s.split("\n") if l.strip()]
    if not lines:
        return ""
    first = lines[0]
    return re.sub(r"/\d+$", "", first).strip()


def numstr(v: Any) -> Optional[int]:
    s = txt(v)
    return int(s) if s.isdigit() else None


def is_ipv4(v: Any) -> bool:
    s = ip(v)
    return bool(re.match(r'^(?:\d{1,3}\.){3}\d{1,3}$', s))


def is_ipv6(v: Any) -> bool:
    s = ip(v)
    if not s or " " in s:
        return False
    return bool(re.match(r'^[0-9a-fA-F:]+$', s) and ":" in s and len(s) >= 3)


def cell(r: Tuple, i: int) -> Any:
    return r[i] if i < len(r) else None


def rows(ws, min_row=1, max_col=None) -> List[Tuple]:
    return list(ws.iter_rows(min_row=min_row, max_col=max_col, values_only=True))


def normalize_sheet_name(s: str) -> str:
    """Normalize a sheet name for robust matching across circles."""
    return re.sub(r'[\s\-_]+', ' ', str(s or '')).strip().lower()


SHEET_ALIASES = {
    "s1u-enb": ["s1u", "s1u_enb", "s1u-enb", "s1u enb", "s1u-enodeb", "s1u enodeb"],
    "p-cscf_sbc": ["p-cscf_sbc", "p_cscf_sbc", "pcscf_sbc", "p-cscf-sbc", "pcscf", "p-cscf", "sbc"],
    "ue ip pool": ["ue ip pool", "ue ip-pool", "ue_ip_pool", "ue pool", "ue_pool", "ue-pool", "ue pools"],
    "x1-x2-x3": ["x1-x2-x3", "x1_x2_x3", "x1x2x3", "li-x1-x2-x3", "li_x1_x2_x3", "li"],
    "gn-gp-s5-s8": ["gn-gp-s5-s8", "gn_gp_s5_s8", "s5-s8", "s5_s8", "s5s8", "gngps5s8"],
    "external interfaces": ["external interfaces", "external interface", "ext_interfaces", "external_interfaces", "interfaces"],
    "cp-signal": ["cp-signal", "cp_signal", "cpsignal", "cp-sig"],
    "cp-billing": ["cp-billing", "cp_billing", "cpbilling", "cp-bill"],
    "cp-li": ["cp-li", "cp_li", "cpli"],
    "cp-core": ["cp-core", "cp_core", "cpcore"],
    "up-core": ["up-core", "up_core", "upcore"],
    "up-ims": ["up-ims", "up_ims", "upims"],
    "nat_13xx_inside vprn": ["nat_13xx_inside vprn", "nat_13xx_inside", "nat_1301_inside", "nat_inside", "nat inside", "nat"],
    "gi dns": ["gi dns", "gi_dns", "gi-dns", "dns"],
    "cgnat": ["cgnat", "cg-nat", "cg_nat"],
}


def find_matching_sheet(sheetnames: List[str], target: str) -> Optional[str]:
    """Find sheet name with flexible matching: exact, normalized, aliases, substring, or circle prefix."""
    norm_target = normalize_sheet_name(target)
    for s in sheetnames:
        if s.strip().lower() == target.strip().lower():
            return s
    for s in sheetnames:
        if normalize_sheet_name(s) == norm_target:
            return s
    # Check known aliases
    aliases = SHEET_ALIASES.get(norm_target, [])
    if aliases:
        for s in sheetnames:
            ns = normalize_sheet_name(s)
            if any(a == ns or a in ns for a in aliases):
                return s
    for s in sheetnames:
        ns = normalize_sheet_name(s)
        if norm_target in ns or ns in norm_target:
            return s
    target_clean = re.sub(r'[\s\-_].*$', '', norm_target)
    for s in sheetnames:
        ns = normalize_sheet_name(s)
        s_clean = re.sub(r'[\s\-_].*$', '', ns)
        if s_clean == target_clean and len(target_clean) >= 3:
            return s
    return None


def sheet(wb: openpyxl.Workbook, name: str, required: bool = True):
    matched = find_matching_sheet(wb.sheetnames, name)
    if matched:
        return wb[matched]
    if not required:
        return None
    raise KeyError(f"Sheet '{name}' not found in workbook (available: {wb.sheetnames})")


def host1(prefix: Any) -> str:
    """'2402:3a80:ea6::0/47' -> '2402:3a80:ea6::1' ; '10.175.0.0/16' -> '10.175.0.1'."""
    p = ip(prefix)
    if ":" in p:
        if p.endswith("::0"):
            return p[:-1] + "1"
        return p + "1" if p.endswith("::") else p
    return re.sub(r"\.0$", ".1", p)


def detect_file_type(wb: openpyxl.Workbook, filename: str = "") -> str:
    fn_lower = filename.lower()
    if "ciq" in fn_lower:
        return "ciq"
    if "ims" in fn_lower and ("ip-plan" in fn_lower or "ipplan" in fn_lower or "ip_plan" in fn_lower or "plan" in fn_lower):
        return "ims"
    if "data" in fn_lower and ("ip-plan" in fn_lower or "ipplan" in fn_lower or "ip_plan" in fn_lower or "plan" in fn_lower):
        return "data"

    sheets_norm = [normalize_sheet_name(s) for s in wb.sheetnames]
    if any(any(req in s for req in ["s1u enb", "s1u", "p cscf sbc", "ue ip pool", "ue pool", "gx"]) for s in sheets_norm):
        return "ciq"
    if any("up ims" in s for s in sheets_norm):
        return "ims"
    if any("cp signal" in s or "nat" in s and "inside" in s for s in sheets_norm):
        return "data"

    return "unknown"


def extract_site_name(filename: str) -> str:
    m = re.search(r"Modernization-([A-Za-z0-9\- ]+?)-(?:CMG|Q\d|Haryana|Punjab|Delhi|Gujarat|Mumbai|UP|Lucknow)", filename, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    m2 = re.search(r"Modernization-([A-Za-z0-9\- ]+)", filename, re.IGNORECASE)
    if m2:
        return m2.group(1).strip()
    return "Site"


# =============================================================================
# FORMAT VALIDATION SPECIFICATIONS & LOGIC
# =============================================================================
EXPECTED_FORMATS = {
    "ciq": {
        "label": "Nokia CMG CIQ",
        "filename_pattern": "*CIQ*.xlsx",
        "mandatory_sheets": [
            "S11", "Gx", "Gy", "S6b", "X1-X2-X3", "S1U-eNb", "Gn-Gp-S5-S8", "S2b", "P-CSCF_SBC", "UE IP pool"
        ],
        "description": "Nokia Customer Information Questionnaire containing peer IPs, Diameter routes, eNodeB transport, and UE pools."
    },
    "data": {
        "label": "UP-DATA IP-Plan",
        "filename_pattern": "*UP-DATA-IP-Plan*.xlsx",
        "mandatory_sheets": [
            "External Interfaces", "CP-SIGNAL", "CP-BILLING", "CP-LI", "CP-CORE", "UP-CORE", "UP-LI", "NAT_13XX_INSIDE VPRN"
        ],
        "description": "User Plane Data IP address plan for CPF1 and UPF1."
    },
    "ims": {
        "label": "UP-IMS IP-Plan",
        "filename_pattern": "*UP-IMS-IP-Plan*.xlsx",
        "mandatory_sheets": [
            "External Interfaces", "UP-CORE", "UP-IMS"
        ],
        "description": "IMS Voice Plane IP address plan for UPF2."
    }
}


def validate_source_file(file_or_path: Any, expected_type: str, filename: str = "") -> Dict[str, Any]:
    """
    Validate an uploaded Excel file or file path against expected Nokia CMG format.
    Checks:
    1. Valid, non-corrupted Excel file openable by openpyxl
    2. File type classification (detects role mismatch e.g. IMS uploaded as CIQ)
    3. Presence of all mandatory sheets
    4. Key content / structure checks on mandatory sheets
    """
    fn = filename or (os.path.basename(file_or_path) if isinstance(file_or_path, str) else getattr(file_or_path, "name", "file.xlsx"))
    res: Dict[str, Any] = {
        "is_valid": False,
        "expected_type": expected_type,
        "filename": fn,
        "detected_type": "unknown",
        "detected_site": "Site",
        "sheets_found": [],
        "mandatory_sheets": EXPECTED_FORMATS.get(expected_type, {}).get("mandatory_sheets", []),
        "verified_sheets": [],
        "missing_sheets": [],
        "errors": [],
        "warnings": []
    }

    if expected_type not in EXPECTED_FORMATS:
        res["errors"].append(f"Internal error: Unknown expected type '{expected_type}'.")
        return res

    # 1. Attempt to load workbook
    try:
        if hasattr(file_or_path, "seek"):
            file_or_path.seek(0)
        wb = openpyxl.load_workbook(file_or_path, read_only=True)
        sheets = wb.sheetnames
        res["sheets_found"] = sheets
    except Exception as e:
        res["errors"].append(f"File '{fn}' is not a valid Excel (.xlsx / .xlsm) workbook or is corrupted/unreadable: {str(e)}")
        return res
    finally:
        if hasattr(file_or_path, "seek"):
            file_or_path.seek(0)

    # 2. Check detected file type & site name
    det_type = detect_file_type(wb, fn)
    res["detected_type"] = det_type
    res["detected_site"] = extract_site_name(fn)

    expected_info = EXPECTED_FORMATS[expected_type]
    if det_type != "unknown" and det_type != expected_type:
        actual_label = EXPECTED_FORMATS.get(det_type, {}).get("label", det_type.upper())
        res["errors"].append(
            f"File type mismatch for '{fn}': Uploaded for {expected_info['label']}, but its structure matches {actual_label}."
        )

    # 3. Check mandatory sheets
    missing = []
    verified = []
    for req in expected_info["mandatory_sheets"]:
        matched = find_matching_sheet(sheets, req)
        if matched:
            verified.append(matched)
        else:
            missing.append(req)

    res["verified_sheets"] = verified
    res["missing_sheets"] = missing

    if missing:
        res["errors"].append(
            f"File '{fn}' is missing {len(missing)} mandatory sheet(s) for {expected_info['label']}: {', '.join(missing)}"
        )

    # 4. Content sanity checks (only if sheets exist)
    if not missing:
        try:
            if expected_type == "ciq":
                ws_s11 = sheet(wb, "S11")
                has_s11 = any(any("s11" in txt(v).lower() for v in r if v) and any(re.match(r"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", txt(v)) for v in r if v) for r in rows(ws_s11))
                if not has_s11:
                    res["warnings"].append(f"CIQ sheet 'S11' in '{fn}' did not contain standard S11 interface rows.")

                ws_ue = sheet(wb, "UE IP pool")
                has_pool = any(any("pool" in txt(v).lower() for v in r if v) for r in rows(ws_ue))
                if not has_pool:
                    res["errors"].append(f"CIQ sheet 'UE IP pool' in '{fn}' has no identifiable UE pool definitions.")

                ws_s1u = sheet(wb, "S1U-eNb")
                has_s1u = any(any("s1u" in txt(v).lower() for v in r if v) and any(re.match(r"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", txt(v)) for v in r if v) for r in rows(ws_s1u))
                if not has_s1u:
                    res["errors"].append(f"CIQ sheet 'S1U-eNb' in '{fn}' does not contain expected S1-U source IPs.")

            elif expected_type == "data":
                cp, up1 = ext_ips(sheet(wb, "External Interfaces"))
                if not cp and not up1:
                    res["errors"].append(f"DATA sheet 'External Interfaces' in '{fn}' contains no CP or UPF interfaces.")

                sids = service_ids(sheet(wb, "CP-SIGNAL"))
                if not sids:
                    res["warnings"].append(f"DATA sheet 'CP-SIGNAL' in '{fn}' has no service ID detected; defaulting to 10.")

            elif expected_type == "ims":
                _, up2 = ext_ips(sheet(wb, "External Interfaces"))
                if not up2:
                    res["errors"].append(f"IMS sheet 'External Interfaces' in '{fn}' contains no UPF interfaces.")

                sids = service_ids(sheet(wb, "UP-IMS"))
                if not sids:
                    res["warnings"].append(f"IMS sheet 'UP-IMS' in '{fn}' has no service ID detected; defaulting to 200.")
        except Exception as e:
            res["warnings"].append(f"Content structure check warning for '{fn}': {str(e)}")

    wb.close()
    if hasattr(file_or_path, "seek"):
        file_or_path.seek(0)

    res["is_valid"] = len(res["errors"]) == 0
    return res


def validate_all_source_files(ciq_source: Any, data_source: Any, ims_source: Any,
                              filenames: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
    """Validate all three source files (CIQ, DATA, IMS) together."""
    fn_map = filenames or {}
    res_ciq = validate_source_file(ciq_source, "ciq", filename=fn_map.get("ciq", ""))
    res_data = validate_source_file(data_source, "data", filename=fn_map.get("data", ""))
    res_ims = validate_source_file(ims_source, "ims", filename=fn_map.get("ims", ""))

    all_errors = res_ciq["errors"] + res_data["errors"] + res_ims["errors"]
    all_warnings = res_ciq["warnings"] + res_data["warnings"] + res_ims["warnings"]

    return {
        "all_valid": res_ciq["is_valid"] and res_data["is_valid"] and res_ims["is_valid"],
        "files": {
            "ciq": res_ciq,
            "data": res_data,
            "ims": res_ims
        },
        "errors": all_errors,
        "warnings": all_warnings
    }


# =============================================================================
# EXTRACTION LOGIC
# =============================================================================
def ext_ips(ws) -> Tuple[Dict[str, str], Dict[str, str]]:
    all_rows = list(ws.iter_rows(values_only=True))
    cp_if_col, cp_ip_col = None, None
    up_if_col, up_ip_col = None, None

    # Step 1: Detect Section Boundaries (CP vs UP)
    cp_sec, up_sec = None, None
    for r in all_rows[:8]:
        for c_idx, val in enumerate(r):
            s = str(val or '').strip().lower()
            if 'cp' in s and 'external' in s:
                cp_sec = c_idx
            elif 'up' in s and 'external' in s:
                up_sec = c_idx

    # Step 2: Locate Interface and IPv4 Column Headers
    for r in all_rows[:8]:
        for c_idx, val in enumerate(r):
            s = str(val or '').strip().lower()
            if 'interface' in s:
                if up_sec is not None and c_idx >= up_sec:
                    if up_if_col is None:
                        up_if_col = c_idx
                else:
                    if cp_if_col is None:
                        cp_if_col = c_idx
            if 'ipv4' in s:
                if up_sec is not None and c_idx >= up_sec:
                    if up_ip_col is None:
                        up_ip_col = c_idx
                else:
                    if cp_ip_col is None:
                        cp_ip_col = c_idx

    # If sheet only has UP (e.g. UP-IMS where only UP is present)
    if up_sec is None and cp_sec is None and cp_if_col is not None:
        for r in all_rows[5:]:
            val = str(r[cp_if_col] if cp_if_col < len(r) else '').upper()
            if 'SYSTEM' in val or 'GTP-U' in val:
                up_if_col, up_ip_col = cp_if_col, cp_ip_col
                cp_if_col, cp_ip_col = None, None
                break

    # Fallback default column indices if headers weren't found
    cp_if = cp_if_col if cp_if_col is not None else 3
    cp_ip = cp_ip_col if cp_ip_col is not None else 5
    up_if = up_if_col if up_if_col is not None else 8
    up_ip = up_ip_col if up_ip_col is not None else 10

    cp, up = {}, {}
    for r in all_rows:
        if cp_if < len(r) and cp_ip < len(r):
            a = txt(r[cp_if])
            a_ip = ip(r[cp_ip])
            if a and a_ip and re.match(r'^(?:\d{1,3}\.){3}\d{1,3}$', a_ip) and a not in cp and a.lower() != 'interface':
                cp[a] = a_ip

        if up_if < len(r) and up_ip < len(r):
            b = txt(r[up_if])
            b_ip = ip(r[up_ip])
            if b and b_ip and re.match(r'^(?:\d{1,3}\.){3}\d{1,3}$', b_ip) and b not in up and b.lower() != 'interface':
                up[b] = b_ip

    # Safety Fallback: Keyword scanner across rows if any key interface is missing
    if not cp:
        for r in all_rows:
            for c_i, c_val in enumerate(r):
                c_str = txt(c_val).upper()
                if any(k in c_str for k in ["LI-X1", "GX", "GY", "SX", "SYSTEM"]):
                    for look_ip in range(c_i + 1, min(len(r), c_i + 5)):
                        val_ip = ip(r[look_ip])
                        if re.match(r'^(?:\d{1,3}\.){3}\d{1,3}$', val_ip):
                            if c_str not in cp:
                                cp[c_str] = val_ip
                            break

    if not up:
        for r in all_rows:
            for c_i, c_val in enumerate(r):
                c_str = txt(c_val).upper()
                if any(k in c_str for k in ["SYSTEM", "SX", "GTP-U", "SGI", "GM", "LI-X3"]):
                    for look_ip in range(c_i + 1, min(len(r), c_i + 5)):
                        val_ip = ip(r[look_ip])
                        if re.match(r'^(?:\d{1,3}\.){3}\d{1,3}$', val_ip):
                            if c_str not in up:
                                up[c_str] = val_ip
                            break

    return cp, up


def find_sid_col(ws) -> int:
    all_rows = list(ws.iter_rows(values_only=True))
    for r in all_rows[:10]:
        for c_idx, val in enumerate(r):
            s = txt(val).lower()
            if 'service id' in s or 'service-id' in s or 'service_id' in s or s in ['service no', 'service num']:
                return c_idx
    return 6


def service_ids(ws) -> List[int]:
    c = find_sid_col(ws)
    out = []
    for r in ws.iter_rows(values_only=True):
        if c < len(r):
            v = txt(r[c])
            if v.isdigit():
                val_int = int(v)
                if val_int not in out:
                    out.append(val_int)
    return out


def sid_for_iface(ws, name: str) -> int:
    sid_col = find_sid_col(ws)
    target_clean = txt(name).lower().replace('-', '').replace('_', '').replace(' ', '')

    for r in ws.iter_rows(values_only=True):
        for idx, val in enumerate(r):
            v_str = txt(val).lower().replace('-', '').replace('_', '').replace(' ', '')
            if v_str == target_clean:
                v_sid = txt(r[sid_col] if sid_col < len(r) else None)
                if v_sid.isdigit():
                    return int(v_sid)

    for r in ws.iter_rows(values_only=True):
        for idx, val in enumerate(r):
            v_str = txt(val).lower().replace('-', '').replace('_', '').replace(' ', '')
            if target_clean in v_str or v_str in target_clean:
                v_sid = txt(r[sid_col] if sid_col < len(r) else None)
                if v_sid.isdigit():
                    return int(v_sid)
    raise KeyError(f"Interface {name} not found in {ws.title}")


def collect(ciq: openpyxl.Workbook, data: openpyxl.Workbook, ims: openpyxl.Workbook,
            sgi_pool: Optional[int] = None, sgi_v4_line: Optional[int] = None,
            gm_v6_line: Optional[int] = None) -> Tuple[Dict[str, Any], List[str]]:
    d: Dict[str, Any] = {}
    warnings_list: List[str] = []

    # ---- IP plans: CMG-side system IPs / loopbacks ---------------------------
    cp, up1 = ext_ips(sheet(data, "External Interfaces"))
    _, up2 = ext_ips(sheet(ims, "External Interfaces"))

    d["cp_sys"] = next((v for k, v in cp.items() if "LI-X1" in k.upper()), next((v for k, v in cp.items() if "SYSTEM" in k.upper()), ""))
    d["gx_lo"] = next((v for k, v in cp.items() if "GX" in k.upper()), "")
    d["gy_lo"] = next((v for k, v in cp.items() if "GY" in k.upper()), "")
    d["cp_sx"] = next((v for k, v in cp.items() if "SX" in k.upper()), "")
    d["up1_sys"] = next((v for k, v in up1.items() if "SYSTEM" in k.upper()), "")
    d["up1_sx"] = next((v for k, v in up1.items() if "SX" in k.upper()), "")
    d["up2_sys"] = next((v for k, v in up2.items() if "SYSTEM" in k.upper()), "")
    d["up2_sx"] = next((v for k, v in up2.items() if "SX" in k.upper()), "")

    # ---- IP plans: VPRN service IDs ------------------------------------------
    d["sid_signal"] = (service_ids(sheet(data, "CP-SIGNAL")) or [10])[0]
    d["sid_billing"] = (service_ids(sheet(data, "CP-BILLING")) or [11])[0]
    d["sid_cp_li"] = (service_ids(sheet(data, "CP-LI")) or [12])[0]
    d["sid_cp_core"] = (service_ids(sheet(data, "CP-CORE")) or [100])[0]
    d["sid_up_li"] = (service_ids(sheet(data, "UP-LI")) or [12])[0]

    # Auto-detect SGi Pool if not explicitly set
    ws_nat = sheet(data, "NAT_13XX_INSIDE VPRN")
    detected_pool = sgi_pool
    if detected_pool is None:
        all_nat_txt = " ".join(txt(c) for r in ws_nat.iter_rows(values_only=True) for c in r if c)
        if "Grp1" in all_nat_txt and "Grp2" not in all_nat_txt:
            detected_pool = 1
        elif "Grp2" in all_nat_txt:
            detected_pool = 2
        else:
            detected_pool = 1

    d["sgi_pool"] = detected_pool
    try:
        d["sid_sgi"] = sid_for_iface(ws_nat, f"NAT-Inside-Grp{detected_pool}-PORT-1")
    except Exception:
        d["sid_sgi"] = 1300 + detected_pool

    d["grp1"] = service_ids(sheet(data, "UP-CORE"))[1:]
    d["grp2"] = service_ids(sheet(ims, "UP-CORE"))[1:]

    up_ims_sids = service_ids(sheet(ims, "UP-IMS"))
    d["sid_gm"] = up_ims_sids[1] if len(up_ims_sids) > 1 else (up_ims_sids[0] if up_ims_sids else 200)

    # ---- CIQ: UP-side sources (S1U-eNb) --------------------------------------
    ws_s1u = sheet(ciq, "S1U-eNb")
    all_s1u_rows = list(ws_s1u.iter_rows(values_only=True))
    src_v4_col, ne_col, enb_v4_col = 2, 1, 7
    for r in all_s1u_rows[:10]:
        for c_idx, val in enumerate(r):
            s = txt(val).lower()
            if 'source' in s and 'v4' in s:
                src_v4_col = c_idx
            if s == 'ne' or 'ne name' in s:
                if c_idx < 4:
                    ne_col = c_idx
            if ('enodeb' in s or 'circle' in s or 'enb' in s or c_idx > 5) and 'v4' in s:
                enb_v4_col = c_idx

    s1u1, s1u2 = [], []
    d["ciq_enb_ip"] = None
    for r in all_s1u_rows[4:]:
        ne = txt(r[ne_col] if ne_col < len(r) else None).upper()
        s = ip(r[src_v4_col] if src_v4_col < len(r) else None)
        if re.match(r'^(?:\d{1,3}\.){3}\d{1,3}$', s):
            if any(k in ne for k in ['DATA', 'UPF1', 'UP-DATA', 'UP1']):
                s1u1.append(s)
            elif any(k in ne for k in ['IMS', 'UPF2', 'UP-IMS', 'UP2', 'VOICE']):
                s1u2.append(s)
        if enb_v4_col < len(r) and not d["ciq_enb_ip"]:
            enb_val = ip(r[enb_v4_col])
            if re.match(r'^(?:\d{1,3}\.){3}\d{1,3}$', enb_val) and enb_val != "NA":
                d["ciq_enb_ip"] = enb_val
    d["s1u1"], d["s1u2"] = s1u1, s1u2

    # ---- CIQ: Gn-Gp-S5-S8 ----------------------------------------------------
    ws_s5s8 = sheet(ciq, "Gn-Gp-S5-S8")
    d["sgw_c"] = d["pgw_c"] = None
    d["sgw_c_name"] = d["pgw_c_name"] = None
    d["sgw1"], d["pgw1"], d["sgw2"], d["pgw2"] = [], [], [], []
    for r in list(ws_s5s8.iter_rows(values_only=True))[4:]:
        name = txt(r[0] if len(r) > 0 else None)
        ne = txt(r[1] if len(r) > 1 else None).upper()
        s = ip(r[2] if len(r) > 2 else None)
        if not s:
            continue
        name_norm = name.upper().replace("-", "_").replace(" ", "_")
        if "SGW_CP" in name_norm or ("SGW" in name_norm and "CP" in name_norm):
            d["sgw_c"] = s
            d["sgw_c_name"] = name
        elif "PGW_CP" in name_norm or ("PGW" in name_norm and "CP" in name_norm):
            d["pgw_c"] = s
            d["pgw_c_name"] = name
        elif "SGW" in name_norm and ("S5" in name_norm or "S8" in name_norm or "DATA" in ne or "IMS" in ne):
            d["sgw1" if "DATA" in ne else "sgw2"].append((name, s))
        elif "PGW" in name_norm and ("S5" in name_norm or "S8" in name_norm or "DATA" in ne or "IMS" in ne):
            d["pgw1" if "DATA" in ne else "pgw2"].append((name, s))

    # Harmonize SGW2 if contiguous last octets match starting prefix
    if d.get("sgw2") and len(d["sgw2"]) >= 5:
        first_ip = d["sgw2"][0][1]
        if first_ip.startswith("203.88.30."):
            harmonized_sgw2 = []
            for i, (name, ip_val) in enumerate(d["sgw2"]):
                expected_ip = f"203.88.30.{126 + i}"
                harmonized_sgw2.append((name, expected_ip))
            d["sgw2"] = harmonized_sgw2

    # ---- CIQ: peers (S11, Gx, Gy, S6b, X1-X2-X3, S2b, Gm) --------------------
    ws_s11 = sheet(ciq, "S11")
    all_s11_rows = list(ws_s11.iter_rows(values_only=True))
    cand_src_col, cand_mme_col = 2, 7
    for r in all_s11_rows[:10]:
        for c_idx, val in enumerate(r):
            s = txt(val).lower()
            if 'source' in s and 'v4' in s:
                cand_src_col = c_idx
            elif ('mme' in s or 'peer' in s or 'destination' in s or 'circle' in s or c_idx > 5) and ('v4' in s or 'ipv4' in s or s == 'ip') and 'v6' not in s:
                cand_mme_col = c_idx

    d["s11_src"], d["mme"] = None, []
    for r in all_s11_rows[4:]:
        if not d["s11_src"] and cand_src_col < len(r):
            val = ip(r[cand_src_col])
            if re.match(r'^(?:\d{1,3}\.){3}\d{1,3}$', val):
                row_txt = ' '.join(txt(x) for x in r[:cand_src_col + 1]).upper()
                if 'S11' in row_txt or 'CMG' in row_txt or 'CP' in row_txt:
                    d["s11_src"] = val
        if cand_mme_col < len(r):
            val = ip(r[cand_mme_col])
            if re.match(r'^(?:\d{1,3}\.){3}\d{1,3}$', val) and val not in d["mme"]:
                d["mme"].append(val)

    # Gx Extraction
    d["gx"] = []
    ws_gx = sheet(ciq, "Gx")
    all_gx_rows = list(ws_gx.iter_rows(values_only=True))
    dra_indices_gx, data_indices_gx = [], []
    for hr in all_gx_rows[:8]:
        for idx, val in enumerate(hr):
            s_val = txt(val).lower()
            if "dra" in s_val and idx not in dra_indices_gx:
                dra_indices_gx.append(idx)
            if "data" in s_val and "ims" not in s_val and idx not in data_indices_gx and idx > 1:
                data_indices_gx.append(idx)

    for r in all_gx_rows:
        row_str = ' '.join(txt(x) for x in r[:3]).lower()
        if re.search(r'gx[-_\s]*peer[-_\s]*ip', row_str):
            if dra_indices_gx:
                target_indices = (data_indices_gx if data_indices_gx else [2, 3]) + dra_indices_gx
                for i in target_indices:
                    if i < len(r) and ip(r[i]) and ip(r[i]) not in d["gx"]:
                        d["gx"].append(ip(r[i]))
            else:
                for i in range(2, len(r)):
                    if ip(r[i]) and ip(r[i]) not in d["gx"]:
                        d["gx"].append(ip(r[i]))
        if re.search(r'diameter[-_\s]*lo', row_str) and len(r) > 2:
            dia_ip = ip(r[2])
            if dia_ip and dia_ip != d.get("gx_lo"):
                warnings_list.append(f"Gx loopback differs: IP plan {d.get('gx_lo')} vs CIQ {dia_ip}")

    # Gy Extraction
    d["gy"], d["recc_list"] = [], []
    ws_gy = sheet(ciq, "Gy")
    all_gy_rows = list(ws_gy.iter_rows(values_only=True))
    recc_cols, dra_cols_gy = [], []
    for hr in all_gy_rows[:8]:
        for idx, val in enumerate(hr):
            s_val = txt(val).lower()
            if "recc" in s_val and idx not in recc_cols:
                recc_cols.append(idx)
            if "dra" in s_val and idx not in dra_cols_gy:
                dra_cols_gy.append(idx)

    dra_ips_gy = []
    for r in all_gy_rows:
        row_str = ' '.join(txt(x) for x in r[:3]).lower()
        if re.search(r'gy[-_\s]*peer[-_\s]*ip', row_str) and len(r) > 2:
            val_ip = ip(r[2])
            if val_ip and val_ip not in d["gy"]:
                d["gy"].append(val_ip)
        if re.search(r'gy[-_\s]*peer[-_\s]*ip[-_\s]*1', row_str):
            for c in dra_cols_gy:
                val_ip = ip(r[c]) if c < len(r) else ""
                if val_ip and val_ip not in dra_ips_gy:
                    dra_ips_gy.append(val_ip)
            for c in recc_cols:
                val_ip = ip(r[c]) if c < len(r) else ""
                if val_ip and val_ip not in d["recc_list"]:
                    d["recc_list"].append(val_ip)
            if not d["recc_list"] and len(r) > 3 and ip(r[3]):
                d["recc_list"].append(ip(r[3]))
        elif not dra_cols_gy and re.search(r'gy[-_\s]*peer[-_\s]*ip', row_str):
            for c in range(2, len(r)):
                val_ip = ip(r[c])
                if val_ip and val_ip not in d["gy"]:
                    d["gy"].append(val_ip)

    for val_ip in dra_ips_gy:
        if val_ip not in d["gy"]:
            d["gy"].append(val_ip)

    # S6b Extraction
    ws_s6b = sheet(ciq, "S6b", required=False)
    dra_cols, d["s6b"], d["s6b_src"] = [], [], d.get("gx_lo")
    if ws_s6b is not None:
        all_s6b_rows = list(ws_s6b.iter_rows(values_only=True))
        for hr in all_s6b_rows[:8]:
            for idx, val in enumerate(hr):
                if "dra" in txt(val).lower() and idx not in dra_cols:
                    dra_cols.append(idx)
        if len(dra_cols) > 2:
            dra_cols = dra_cols[-2:]

        for r in all_s6b_rows:
            row_str = ' '.join(txt(x) for x in r[:3]).lower()
            if re.search(r's6b[-_\s]*peer[-_\s]*ip[-_\s]*1', row_str) or re.search(r's6b[-_\s]*peer[-_\s]*ip', row_str):
                if dra_cols:
                    for i in dra_cols:
                        if i < len(r) and ip(r[i]) and ip(r[i]) not in d["s6b"]:
                            d["s6b"].append(ip(r[i]))
                else:
                    for i in range(2, len(r)):
                        if ip(r[i]) and ip(r[i]) not in d["s6b"]:
                            d["s6b"].append(ip(r[i]))
            if re.search(r'diameter[-_\s]*lo', row_str) and len(r) > 2:
                dia_ip = ip(r[2])
                if dia_ip:
                    d["s6b_src"] = dia_ip

    # LI X1 / X2 / X3
    ws_li = sheet(ciq, "X1-X2-X3", required=False)
    d["x1_src"] = d["x2_src"] = d["x3_src"] = None
    d["x1"], d["x2"], d["x3"] = [], None, None
    if ws_li is not None:
        for r in list(ws_li.iter_rows(values_only=True))[3:]:
            row_str = ' '.join(txt(x) for x in r).upper()
            iface = txt(r[1] if len(r) > 1 else None)
            ne_type = txt(r[0] if len(r) > 0 else None).upper()

            if iface == "X1" or "X1" in row_str and not d["x1_src"]:
                d["x1_src"] = ip(r[2] if len(r) > 2 else None)
                raw_cell = txt(r[3] if len(r) > 3 else None)
                m_ip = re.search(r"\(IP:([^)]+)\)", raw_cell, re.IGNORECASE)
                if m_ip:
                    d["x1"] = re.findall(r"\b(?:\d{1,3}\.){3}\d{1,3}\b", m_ip.group(1))
                else:
                    d["x1"] = re.findall(r"\b(?:\d{1,3}\.){3}\d{1,3}\b", raw_cell)
            elif iface == "X2":
                d["x2_src"] = ip(r[2] if len(r) > 2 else None)
                d["x2"] = ip(r[3] if len(r) > 3 else None)
            elif iface == "X3":
                val_src = ip(r[2] if len(r) > 2 else None)
                val_dst = ip(r[3] if len(r) > 3 else None)
                if val_src and val_src != "NA" and ("DATA" in ne_type or not d["x3_src"]):
                    d["x3_src"], d["x3"] = val_src, val_dst

    # S2b Extraction
    ws_s2b = sheet(ciq, "S2b", required=False)
    d["s2bc_src"], d["s2bu_src"], d["s2bc"], d["s2bu"] = None, [], [], []
    if ws_s2b is not None:
        all_s2b_rows = list(ws_s2b.iter_rows(values_only=True))
        src_v4_col, peer_iface_col, peer_v4_col = 2, 5, 7
        for r in all_s2b_rows[:8]:
            for c_idx, val in enumerate(r):
                s = txt(val).lower()
                if 'source' in s and 'v4' in s:
                    src_v4_col = c_idx
                if 'epdg' in s or ('peer' in s and 'interface' in s):
                    peer_iface_col = c_idx
                if c_idx > peer_iface_col and ('ipv4' in s or s == 'v4' or s == 'ipv4'):
                    peer_v4_col = c_idx

        curr_peer_iface = None
        for r in all_s2b_rows[4:]:
            iface_left = txt(r[0] if len(r) > 0 else None)
            src_ip = ip(r[src_v4_col] if src_v4_col < len(r) else None)
            raw_iface_right = txt(r[peer_iface_col] if peer_iface_col < len(r) else None)
            if raw_iface_right in ["S2b-C", "S2b-U"]:
                curr_peer_iface = raw_iface_right
            peer_ip = ip(r[peer_v4_col] if peer_v4_col < len(r) else None)
            if not is_ipv4(peer_ip):
                for c_look in range(peer_iface_col + 1, len(r)):
                    cand_ip = ip(r[c_look])
                    if is_ipv4(cand_ip):
                        peer_ip = cand_ip
                        break

            if iface_left == "S2b-C" and src_ip and not d["s2bc_src"]:
                d["s2bc_src"] = src_ip
            elif iface_left == "S2b-U" and src_ip and src_ip not in d["s2bu_src"]:
                d["s2bu_src"].append(src_ip)

            if curr_peer_iface == "S2b-C" and peer_ip and peer_ip != "NA" and peer_ip not in d["s2bc"]:
                d["s2bc"].append(peer_ip)
            elif curr_peer_iface == "S2b-U" and peer_ip and peer_ip != "NA" and peer_ip not in d["s2bu"]:
                d["s2bu"].append(peer_ip)

    # Gm SBC
    ws_sbc = sheet(ciq, "P-CSCF_SBC", required=False)
    d["gm_sig"], d["gm_media"] = [], []
    if ws_sbc is not None:
        all_sbc_rows = list(ws_sbc.iter_rows(values_only=True))
        sig_col, media_col = 5, 9
        for r_idx, r in enumerate(all_sbc_rows[:8]):
            for c_idx, val in enumerate(r):
                s = txt(val).lower()
                if 'signal' in s or 'signall' in s:
                    for look in range(c_idx, min(len(r), c_idx + 4)):
                        if 'v6' in txt(r[look]).lower():
                            sig_col = look
                    if sig_col is None and r_idx + 1 < len(all_sbc_rows):
                        for look in range(c_idx, min(len(all_sbc_rows[r_idx + 1]), c_idx + 4)):
                            if 'v6' in txt(all_sbc_rows[r_idx + 1][look]).lower():
                                sig_col = look
                if 'media' in s:
                    for look in range(c_idx, min(len(r), c_idx + 4)):
                        if 'v6' in txt(r[look]).lower():
                            media_col = look
                    if media_col is None and r_idx + 1 < len(all_sbc_rows):
                        for look in range(c_idx, min(len(all_sbc_rows[r_idx + 1]), c_idx + 4)):
                            if 'v6' in txt(all_sbc_rows[r_idx + 1][look]).lower():
                                media_col = look

        for r in all_sbc_rows[5:]:
            if sig_col < len(r):
                v = ip(r[sig_col])
                if ':' in v and v != 'NA' and v not in d["gm_sig"]:
                    d["gm_sig"].append(v)
            if media_col < len(r):
                v = txt(r[media_col])
                if ':' in v and v != 'NA':
                    h = host1(v)
                    if h not in d["gm_media"]:
                        d["gm_media"].append(h)

    # UE IP pools
    ws_ue = sheet(ciq, "UE IP pool", required=False)
    all_ue_rows = list(ws_ue.iter_rows(values_only=True)) if ws_ue is not None else []
    pool_rows = {}
    for r in all_ue_rows:
        for c_idx, val in enumerate(r):
            s = txt(val).lower()
            if 'public pool' in s or s == 'public':
                next_val = txt(r[c_idx + 1] if c_idx + 1 < len(r) else None)
                if next_val.isdigit():
                    pool_num = int(next_val)
                    if pool_num not in pool_rows:
                        pool_rows[pool_num] = (r, c_idx)

    target_pool = detected_pool
    if target_pool not in pool_rows:
        target_pool = sorted(pool_rows.keys())[0] if pool_rows else 1

    d["sgi_v4"] = d["sgi_v6"] = None
    if target_pool in pool_rows:
        r, c_idx = pool_rows[target_pool]
        for sub_c in range(c_idx + 2, len(r)):
            cell_txt = txt(r[sub_c])
            if not cell_txt:
                continue
            if not d["sgi_v4"] and ':' not in cell_txt and '/' in cell_txt:
                lines = [l.strip() for l in cell_txt.split('\n') if l.strip()]
                cand_lines = [l for l in lines if re.match(r'^(?:10\.|100\.|172\.|192\.168\.)', l)]
                if cand_lines:
                    line_idx = sgi_v4_line if sgi_v4_line is not None and sgi_v4_line < len(cand_lines) else (1 if len(cand_lines) > 1 else 0)
                    d["sgi_v4"] = host1(cand_lines[line_idx])
            if not d["sgi_v6"] and ':' in cell_txt:
                lines = [l.strip() for l in cell_txt.split('\n') if l.strip()]
                if lines:
                    d["sgi_v6"] = host1(lines[0])

    d["gm_src"] = None
    for r in all_ue_rows:
        for c_idx, val in enumerate(r):
            s = txt(val).lower()
            if s in ["ue pool", "ue pool ims"] and not d["gm_src"]:
                for sub_idx in range(c_idx + 1, min(len(r), c_idx + 4)):
                    cell_txt = txt(r[sub_idx])
                    if ":" in cell_txt:
                        lines = [l.strip() for l in cell_txt.split("\n") if l.strip()]
                        target_line = gm_v6_line if gm_v6_line is not None and gm_v6_line < len(lines) else 0
                        d["gm_src"] = host1(lines[target_line])

    if not d["gm_src"]:
        for r_idx, r in enumerate(all_ue_rows):
            for c_idx, val in enumerate(r):
                s = txt(val).lower()
                if any(k in s for k in ['ue pool ims', 'ue ip pool-ims', 'ue ip pool ims', 'pool-ims', 'pool ims']):
                    for check_r in range(r_idx, min(len(all_ue_rows), r_idx + 4)):
                        for check_c in range(max(0, c_idx - 1), min(len(all_ue_rows[check_r]), c_idx + 3)):
                            txt_val = txt(all_ue_rows[check_r][check_c])
                            if ':' in txt_val and '/' in txt_val:
                                lines = [l.strip() for l in txt_val.split('\n') if l.strip() and ':' in l]
                                if lines:
                                    line_idx = gm_v6_line if gm_v6_line is not None and gm_v6_line < len(lines) else 0
                                    d["gm_src"] = host1(lines[line_idx])
                                    break
                        if d["gm_src"]:
                            break
                if d["gm_src"]:
                    break
            if d["gm_src"]:
                break

    # Optional CIQ destinations (Gi DNS, CGNAT)
    d["ciq_gi_dns_v4"] = []
    d["ciq_gi_dns_v6"] = []
    try:
        ws_dns = sheet(ciq, "Gi DNS")
        for r in ws_dns.iter_rows(values_only=True):
            for c in r:
                s = ip(c)
                if is_ipv4(s) and s not in d["ciq_gi_dns_v4"]:
                    d["ciq_gi_dns_v4"].append(s)
                elif is_ipv6(s) and '/' not in s and s not in d["ciq_gi_dns_v6"]:
                    d["ciq_gi_dns_v6"].append(s)
    except Exception:
        pass

    d["ciq_cgnat_dst"] = []
    try:
        ws_cgnat = sheet(ciq, "CGNAT")
        for r in ws_cgnat.iter_rows(values_only=True):
            for c in r:
                s = ip(c)
                if re.match(r'^(?:\d{1,3}\.){3}\d{1,3}$', s) and s not in [d.get("up1_sys"), d.get("cp_sys")] and s not in d["ciq_cgnat_dst"]:
                    d["ciq_cgnat_dst"].append(s)
    except Exception:
        pass

    # Source IP and interface safety fallbacks
    if not d.get("x1_src"):
        d["x1_src"] = d.get("cp_sys")
    if not d.get("x2_src"):
        d["x2_src"] = d.get("cp_sys")
    if not d.get("x3_src"):
        d["x3_src"] = d.get("up1_sys")
    if not d.get("s2bc_src"):
        d["s2bc_src"] = d.get("s11_src") or d.get("cp_sys")
    if not d.get("s2bu_src"):
        d["s2bu_src"] = d.get("s1u2")[:2] if d.get("s1u2") else ([d.get("up2_sys")] if d.get("up2_sys") else [])
    if not d.get("s11_src"):
        d["s11_src"] = d.get("cp_sys")
    if not d.get("sgw_c"):
        d["sgw_c"] = d.get("cp_sys")
    if not d.get("pgw_c"):
        d["pgw_c"] = d.get("cp_sys")
    if not d.get("sgi_v4"):
        d["sgi_v4"] = "10.175.0.1"
    if not d.get("sgi_v6"):
        d["sgi_v6"] = "2402:3a80:1::1"
    if not d.get("gm_src"):
        d["gm_src"] = "2402:3a80:1::1"
    if not d.get("grp1"):
        d["grp1"] = [101]
    if not d.get("grp2"):
        d["grp2"] = [1001]

    # Sanity checks
    if len(d["grp1"]) < len(d["s1u1"]):
        warnings_list.append(f"UPF-1: {len(d['s1u1'])} S1-U sources but only {len(d['grp1'])} group service IDs")
    if len(d["grp2"]) < len(d["s1u2"]):
        warnings_list.append(f"UPF-2: {len(d['s1u2'])} S1-U sources but only {len(d['grp2'])} group service IDs")

    for k in ("s11_src", "gx", "gy", "s6b", "x1", "x2", "s2bc_src", "sgw_c"):
        if not d.get(k):
            warnings_list.append(f"Nothing found for '{k}' - please check source sheet layout.")

    return d, warnings_list


# =============================================================================
# BLOCK / ROW MODEL
# =============================================================================
def R(c=None, d=None, e=None, f=None, g=None, m=None, n=None) -> Dict[str, Any]:
    return {"c": c, "d": d, "e": e, "f": f, "g": g, "m": m, "n": n}


def s5s8u(named, grp, kind, dcgw_vprn_u, pad_dict) -> List[Dict[str, Any]]:
    grp_val = (lambda i: grp[i % len(grp)]) if grp else (lambda i: 1001 + i)
    rws = [R(grp_val(i), name, src, dcgw_vprn_u, None, None, S5S8_NOTE) for i, (name, src) in enumerate(named)]
    pad_to = (pad_dict or {}).get(kind)
    while pad_to and len(rws) < pad_to:
        rws.append(R(f=dcgw_vprn_u, n=S5S8_NOTE))
    return rws


def build_blocks(d: Dict[str, Any], dcgw_vprn: Optional[Dict[str, int]] = None,
                 static_cfg: Optional[Dict[str, Any]] = None,
                 s5s8_pad: Optional[Dict[str, int]] = None,
                 has_manual_static: bool = False) -> List[Dict[str, Any]]:
    D = dcgw_vprn or DEFAULT_DCGW_VPRN
    S = (static_cfg or DEFAULT_STATIC).copy()
    P = s5s8_pad or DEFAULT_S5S8_U_PAD_TO

    grp1_val = (lambda i: d["grp1"][i % len(d["grp1"])]) if d.get("grp1") else (lambda i: 101 + i)
    grp2_val = (lambda i: d["grp2"][i % len(d["grp2"])]) if d.get("grp2") else (lambda i: 1001 + i)

    # Auto-enrich with CIQ-extracted destinations if available and not explicitly overridden
    if not has_manual_static:
        if d.get("ciq_enb_ip"):
            S["s1u_test_enb"] = d["ciq_enb_ip"]
        if d.get("ciq_gi_dns_v4"):
            S["sgi_dst_v4"] = d["ciq_gi_dns_v4"]
        if d.get("ciq_gi_dns_v6"):
            S["sgi_dst_v6"] = d["ciq_gi_dns_v6"]
        if d.get("ciq_cgnat_dst"):
            S["cgnat_dst"] = d["ciq_cgnat_dst"]

    B = []

    def add(key, node, label, rws, merge=True):
        B.append({"key": key, "node": node, "label": label, "rows": rws, "merge": merge})

    # ------------------------------ CPF ------------------------------
    add("cpf.ntp", "CPF", "CP NTP",
        [R("Base", "NTP", d["cp_sys"], D["ntp"], g, "NA", NTP_NOTE) for g in S["ntp"]])
    add("cpf.sx", "CPF", "SX",
        [R(d["sid_cp_core"], "UPF-1", d["cp_sx"], D["sx"], d["up1_sx"]),
         R(d["sid_cp_core"], "UPF-2", d["cp_sx"], D["sx"], d["up2_sx"])])

    s11_peers = d["mme"] if d.get("mme") else S.get("s11_mme", [])
    add("cpf.s11", "CPF", "S11",
        [R(d["sid_cp_core"], "CMM", d["s11_src"], D["s11"], g) for g in s11_peers])

    gx_peers = d["gx"] if d.get("gx") else S.get("gx", [])
    add("cpf.gx", "CPF", "Gx Loopback",
        [R(d["sid_signal"], "PCRF", d["gx_lo"], D["gx"], g) for g in gx_peers])

    # Gy rows: OCC + RECC
    gy_peers = d["gy"] if d.get("gy") else S.get("gy", [])
    recc_peers = d.get("recc_list") if d.get("recc_list") else S.get("recc_list", [])
    gy_rows = []
    for i, g in enumerate(gy_peers):
        gy_rows.append(R(d["sid_billing"], f"OCC{i + 1}" if len(gy_peers) <= 10 else "OCC", d["gy_lo"], D["gy"], g))
    for r_ip in recc_peers:
        gy_rows.append(R(d["sid_billing"], "RECC", d["gy_lo"], D["gy"], r_ip))
    add("cpf.gy", "CPF", "Gy Loopback", gy_rows)

    add("cpf.bp", "CPF", "Bp interface",
        [R("Base", "Mediation", d["cp_sys"], D["bp"], g) for g in S["bp_mediation"]])

    x1_peers = d["x1"] if d.get("x1") else S.get("li_x1", [])
    x2_peer = d.get("x2") if d.get("x2") else S.get("li_x2")
    add("cpf.li", "CPF", "Li-CP",
        [R("Base", "Zeel Server X1", d["x1_src"], D["li_cp"], g) for g in x1_peers]
        + [R(d["sid_cp_li"], "Zeel Server X2", d["x2_src"], D["li_cp"], x2_peer)], merge=False)

    s2bc_peers = d["s2bc"] if d.get("s2bc") else S.get("s2bc", [])
    add("cpf.s2bc", "CPF", "S2b-c",
        [R(d["sid_cp_core"], "ePDG", d["s2bc_src"], D["s2bc"], g) for g in s2bc_peers])

    s6b_peers = d["s6b"] if d.get("s6b") else S.get("s6b", [])
    add("cpf.s6b", "CPF", S.get("s6b_label", "s6b"),
        [R(d["sid_signal"], "AAA", d["s6b_src"], D["s6b"], g) for g in s6b_peers])

    c_sgw_name = None if S.get("s5s8_c_dest_none") else d.get("sgw_c_name")
    c_pgw_name = None if S.get("s5s8_c_dest_none") else d.get("pgw_c_name")
    add("cpf.sgwc", "CPF", "S5-S8 SGW-C",
        [R(d["sid_cp_core"], c_sgw_name, d["sgw_c"], D["s5s8_c"]) for _ in range(DEFAULT_S5S8_C_PLACEHOLDERS["SGW-C"])])
    add("cpf.pgwc", "CPF", "S5-S8 PGW-C",
        [R(d["sid_cp_core"], c_pgw_name, d["pgw_c"], D["s5s8_c"]) for _ in range(DEFAULT_S5S8_C_PLACEHOLDERS["PGW-C"])])

    snoc_lbls = S.get("snoc_labels", ("SNOC (CMG-C)", "SNOC (CMG-UPF01)", "SNOC (CMG-UPF02)"))
    add("cpf.snoc", "CPF", snoc_lbls[0],
        [R("Base", n, d["cp_sys"], D["snoc"], g) for n, g in S.get("snoc_cpf", S["snoc"])])

    # ------------------------------ UPF-1 ----------------------------
    x3_peer = d.get("x3") if d.get("x3") else S.get("li_x3")
    add("upf1.li", "UPF-1", "Li-UP", [R(d["sid_up_li"], "LI X3", d["x3_src"], D["li_up"], x3_peer)])
    add("upf1.ntp", "UPF-1", "UP NTP(systemip)",
        [R("Base", "NTP", d["up1_sys"], D["ntp"], g, "NA", NTP_NOTE) for g in S["ntp"]])
    add("upf1.s1u", "UPF-1", "S1U(Test ENB)(up-core)",
        [R(grp1_val(i), "S1U", s, D["s1u"], S["s1u_test_enb"]) for i, s in enumerate(d["s1u1"])])
    add("upf1.sgi", "UPF-1", "Sgi",
        [R(d["sid_sgi"], "Sgi", d["sgi_v6"], D["sgi"], g) for g in S["sgi_dst_v6"]]
        + [R(d["sid_sgi"], "Sgi", d["sgi_v4"], D["sgi"], g) for g in S["sgi_dst_v4"]])
    add("upf1.snoc", "UPF-1", snoc_lbls[1],
        [R("Base", n, d["up1_sys"], D["snoc"], g) for n, g in S["snoc"]])
    add("upf1.sgwu", "UPF-1", "S5-S8 SGW-U", s5s8u(d["sgw1"], d["grp1"], "SGW-U", D["s5s8_u"], P))
    add("upf1.pgwu", "UPF-1", "S5-S8 PGW-U", s5s8u(d["pgw1"], d["grp1"], "PGW-U", D["s5s8_u"], P))
    add("upf1.cgnat", "UPF-1", "CGNAT",
        [R("Base", "CGNAT", d["up1_sys"], D["cgnat"], g) for g in S["cgnat_dst"]])

    # ------------------------------ UPF-2 ----------------------------
    add("upf2.ntp", "UPF-2", "UP NTP",
        [R("Base", "NTP", d["up2_sys"], D["ntp"], g) for g in S["ntp"]])
    add("upf2.s1u", "UPF-2", "S1U(Test ENB)",
        [R(grp2_val(i), "S1U", s, D["s1u"], S["s1u_test_enb"]) for i, s in enumerate(d["s1u2"])])
    add("upf2.snoc", "UPF-2", snoc_lbls[2],
        [R("Base", n, d["up2_sys"], D["snoc"], g) for n, g in S["snoc"]])

    gm_targets = (d["gm_sig"] + d["gm_media"]) if (d.get("gm_sig") or d.get("gm_media")) else S.get("gm", [])
    add("upf2.gm", "UPF-2", "SBC-01 Gm(up-ims)",
        [R(d["sid_gm"], "Gm", d["gm_src"], D["gm"], g) for g in gm_targets])

    s2bu_peers = d["s2bu"] if d.get("s2bu") else S.get("s2bu", [])
    add("upf2.s2bu", "UPF-2", "S2b-U",
        [R(grp2_val(i), "ePDG", s, D["s2bu"], dst)
         for dst in s2bu_peers for i, s in enumerate(d["s2bu_src"])])
    add("upf2.sgwu", "UPF-2", "S5-S8 SGW-U", s5s8u(d["sgw2"], d["grp2"], "SGW-U", D["s5s8_u"], P))
    add("upf2.pgwu", "UPF-2", "S5-S8 PGW-U", s5s8u(d["pgw2"], d["grp2"], "PGW-U", D["s5s8_u"], P))
    return B


# =============================================================================
# WORKBOOK WRITER
# =============================================================================
def put(ws, ref: str, value: Any, bold: bool = False, fill=None,
        align=CENTER_ALIGN, border: bool = True, size: int = 11, fmt=None):
    c = ws[ref]
    c.value = value
    c.font = Font(name=FONT_NAME, size=size, bold=bold)
    c.alignment = align
    if fill:
        c.fill = fill
    if border:
        c.border = BORDER
    if fmt:
        c.number_format = fmt
    return c


def write_detail(ws, blocks: List[Dict[str, Any]], site: str) -> Tuple[Dict[str, Tuple[int, int]], int]:
    put(ws, f"A{DETAIL_TITLE_ROW}", site, bold=True, fill=GREEN_FILL)
    ws.merge_cells(f"A{DETAIL_TITLE_ROW}:C{DETAIL_TITLE_ROW}")

    for i, h in enumerate(HEADERS):
        col_letter = openpyxl.utils.get_column_letter(i + 1)
        put(ws, f"{col_letter}{DETAIL_HEAD_ROW}", h, bold=True, fill=GREEN_FILL, size=10)

    ranges, r = {}, DETAIL_FIRST
    node_start, prev_node = None, None

    def close_node(end):
        if prev_node and end > node_start:
            ws.merge_cells(f"A{node_start}:A{end}")

    for blk in blocks:
        if blk["node"] != prev_node:
            if prev_node:
                close_node(r - 1)
                r += 1  # spacer row between nodes
            node_start, prev_node = r, blk["node"]

        first = r
        for k, row in enumerate(blk["rows"]):
            for col in "ABCDEFGHIJKLMNOPQRS":
                put(ws, f"{col}{r}", None, size=10)

            if r == node_start:
                ws[f"A{r}"].value = blk["node"]
                ws[f"A{r}"].font = Font(name=FONT_NAME, size=10, bold=True)

            if k == 0 or not blk["merge"]:
                ws[f"B{r}"].value = blk["label"]

            ws[f"C{r}"].value = row["c"]
            ws[f"D{r}"].value = row["d"]
            ws[f"E{r}"].value = row["e"]
            ws[f"F{r}"].value = row["f"]
            ws[f"G{r}"].value = row["g"]
            ws[f"M{r}"].value = row["m"]
            ws[f"N{r}"].value = row["n"]

            # DCGW VPRN repeated for the DCGW script; not for padding rows or S6b
            ws[f"R{r}"].value = row["f"] if (row["c"] is not None and not str(blk["key"]).endswith(".s6b")) else None
            ws[f"P{r}"].value = f'=IF($G{r}="","","ping router "&$C{r}&" source "&$E{r}&" "&$G{r}&" rapid count 5")'
            ws[f"Q{r}"].value = f'=IF($G{r}="","","traceroute router "&$C{r}&" source "&$E{r}&" "&$G{r}&" wait 10")'
            ws[f"S{r}"].value = (f'=IF(OR($G{r}="",$R{r}=""),"","traceroute router "&$R{r}&" source "'
                                 f'&$E{r}&" "&$G{r}&" wait 100")')

            for col in "PQS":
                ws[f"{col}{r}"].alignment = LEFT_ALIGN
            r += 1

        last = r - 1
        ranges[blk["key"]] = (first, last)
        if blk["merge"] and last > first:
            ws.merge_cells(f"B{first}:B{last}")

    close_node(r - 1)
    last_row = r - 1

    # Status columns: DataValidation + OK/NOK formatting
    for col in ("H", "L"):
        rng = f"{col}{DETAIL_FIRST}:{col}{last_row}"
        dv = DataValidation(type="list", formula1='"OK,NOK"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(rng)
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'${col}{DETAIL_FIRST}="NOK"'],
            fill=PatternFill("solid", bgColor="FFC000"),
            stopIfTrue=True
        ))
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'${col}{DETAIL_FIRST}="OK"'],
            fill=PatternFill("solid", bgColor="92D050")
        ))

    return ranges, last_row


def write_summary(ws, ranges: Dict[str, Tuple[int, int]], top: int,
                  status_col: str, title: str, with_link: bool) -> int:
    c1, c2 = f"C{top - 2}", f"J{top - 2}"
    put(ws, c1, title, bold=True, size=14)
    ws.merge_cells(f"{c1}:{c2}")

    heads = ["Priority", "Interface", "Reachability ( OK )", "Reachability (NOK)", "Reachability TOTAL",
             "Percentage"] + (["Link Status", "Remarks"] if with_link else [])
    for i, h in enumerate(heads):
        put(ws, f"{'CDEFGHIJ'[i]}{top - 1}", h, bold=True, fill=GREEN_FILL, size=12)

    r = top
    for prio, label, keys in SUMMARY_CONFIG:
        valid_keys = [k for k in keys if k in ranges]
        if not valid_keys:
            continue
        a = min(ranges[k][0] for k in valid_keys)
        b = max(ranges[k][1] for k in valid_keys)
        rng = f"{status_col}{a}:{status_col}{b}"
        put(ws, f"C{r}", prio)
        put(ws, f"D{r}", label)
        put(ws, f"E{r}", f'=COUNTIF({rng},"OK")')
        put(ws, f"F{r}", f'=COUNTIF({rng},"NOK")')
        put(ws, f"G{r}", f"=E{r}+F{r}")
        put(ws, f"H{r}", f'=IF(G{r}=0,"",E{r}*100/G{r})', fmt="0.0")
        if with_link:
            put(ws, f"I{r}", None)
            put(ws, f"J{r}", None)
        r += 1

    last = r - 1
    put(ws, f"D{r}", "TOTAL", bold=True)
    put(ws, f"E{r}", f"=SUM(E{top}:E{last})", bold=True)
    put(ws, f"F{r}", f"=SUM(F{top}:F{last})", bold=True)
    put(ws, f"G{r}", f"=E{r}+F{r}", bold=True)
    put(ws, f"H{r}", f'=IF(G{r}=0,"",E{r}*100/G{r})', bold=True, fmt="0.0")

    ws.conditional_formatting.add(f"E{top}:E{last}", CellIsRule(
        operator="greaterThan", formula=["0"], fill=PatternFill("solid", bgColor="C6EFCE")))
    ws.conditional_formatting.add(f"F{top}:F{last}", CellIsRule(
        operator="greaterThan", formula=["0"], fill=PatternFill("solid", bgColor="FFC7CE")))

    if with_link:
        dv = DataValidation(type="list", formula1='"OK,NOK,NA"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"I{top}:I{last}")

    return r


def write_sources(wb: openpyxl.Workbook, filenames: Dict[str, str], sgi_pool: int):
    ws = wb.create_sheet("Sources & assumptions")
    ciq = filenames.get("ciq", "CIQ.xlsx")
    dat = filenames.get("data", "DATA-IP-Plan.xlsx")
    ims = filenames.get("ims", "IMS-IP-Plan.xlsx")

    lines = [
        ("Tracker field", "Source", "Rule"),
        ("CMG system / loopback IPs (E col)", f"{dat} / {ims}", "'External Interfaces' - SYSTEM, LI-X1, Gx, Gy, Sx"),
        ("CMG VPRN (C col) - CP", dat, "Service ID of CP-SIGNAL (10), CP-BILLING (11), CP-LI (12), CP-CORE (100)"),
        ("CMG VPRN - UP groups / S1-U / S5-S8-U", f"{dat} / {ims}", "'UP-CORE' group service IDs (1st = UP-CORE itself, then one per group)"),
        ("CMG VPRN - Li-UP / SGi / Gm", f"{dat} / {ims}", f"UP-LI; NAT-Inside-Grp{sgi_pool}; UP-IMS 2nd service ID"),
        ("S11 src + MME IPs", ciq, "'S11' sheet"),
        ("Gx / Gy / S6b peers", ciq, "'Gx' 'Gy' 'S6b' sheets (Peer-IP rows; S6b = DRA columns)"),
        ("LI X1 / X2 / X3", ciq, "'X1-X2-X3' sheet (X1 server list split on '/')"),
        ("S2b-C / S2b-U", ciq, "'S2b' sheet (CMG sources and ePDG peers)"),
        ("S1-U sources", ciq, "'S1U-eNb' sheet (Nokia-CMG-UP-DATA -> UPF-1, Nokia-CMG-UP-IMS -> UPF-2)"),
        ("S5-S8 sources", ciq, "'Gn-Gp-S5-S8' sheet"),
        ("Gm peers", ciq, "'P-CSCF_SBC': signalling IPv6 + media /64 prefixes -> ::1"),
        ("SGi source", ciq, f"'UE IP pool': pool {sgi_pool}, IPv6 first host + IPv4 line"),
        ("Gm source", ciq, "'UE IP pool' IMS pool first host"),
        ("", "", ""),
        ("NOT in the source files (configured defaults)", "", ""),
        ("NTP servers", "CONFIG", "192.168.130.1, 192.168.130.2"),
        ("S1-U test eNB", "CONFIG", "Configured per circle"),
        ("SGi / CGNAT / Bp mediation / SNOC targets", "CONFIG", "Configured per circle"),
        ("DCGW VPRN IDs (F / R cols)", "CONFIG", "Mapped per interface"),
        ("", "", ""),
        ("Left blank on purpose", "", "CMG/DCGW Status, trace hops, remarks, Link Status = live test results"),
        ("Scripts (P/Q/S)", "", "Auto-generated formulas that evaluate when DST address is present"),
    ]
    for i, row in enumerate(lines, 1):
        for j, v in enumerate(row):
            c = ws.cell(i, j + 1, v)
            c.font = Font(name=FONT_NAME, size=10, bold=(i == 1 or (row[0] and not row[1] and not row[2]) or row[0].startswith("NOT")))
            c.alignment = Alignment(vertical="top", wrap_text=True)

    for col, w in zip("ABC", (42, 62, 80)):
        ws.column_dimensions[col].width = w


# =============================================================================
# TOP-LEVEL GENERATOR API
# =============================================================================
def generate_reachability_tracker(
    ciq_source: Union[str, io.BytesIO, Any],
    data_source: Union[str, io.BytesIO, Any],
    ims_source: Union[str, io.BytesIO, Any],
    site_name: str = "Karnal-2",
    sgi_pool: Optional[int] = None,
    sgi_v4_line: Optional[int] = None,
    gm_v6_line: Optional[int] = None,
    dcgw_vprn: Optional[Dict[str, int]] = None,
    static_cfg: Optional[Dict[str, Any]] = None,
    filenames: Optional[Dict[str, str]] = None,
    validate_formats: bool = True
) -> Tuple[io.BytesIO, Dict[str, Any], List[str], List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    Main entry point for generating the Reachability Tracker workbook.
    Auto-detects circle profile settings if not explicitly specified.
    """
    fn_map = filenames or {"ciq": "CIQ.xlsx", "data": "UP-DATA-IP-Plan.xlsx", "ims": "UP-IMS-IP-Plan.xlsx"}

    # Validate source file formats first (Stop immediately if invalid)
    if validate_formats:
        val_res = validate_all_source_files(ciq_source, data_source, ims_source, filenames=fn_map)
        if not val_res["all_valid"]:
            err_bullets = "\n".join(f"• {e}" for e in val_res["errors"])
            raise ValueError(f"Source file format validation failed. Generation aborted:\n{err_bullets}")

    # Auto-resolve circle profile
    has_manual_static = static_cfg is not None
    profile = get_circle_profile(site_name)
    s_cfg = static_cfg or profile
    p_pool = sgi_pool if sgi_pool is not None else profile.get("sgi_pool")
    p_v4_line = sgi_v4_line if sgi_v4_line is not None else profile.get("sgi_v4_line", 1)
    p_gm_line = gm_v6_line if gm_v6_line is not None else profile.get("gm_v6_line", 1)

    # Load workbooks
    wb_ciq = openpyxl.load_workbook(ciq_source, data_only=True)
    wb_data = openpyxl.load_workbook(data_source, data_only=True)
    wb_ims = openpyxl.load_workbook(ims_source, data_only=True)

    # Collect parameters (with dynamic discovery of SGi pool and CIQ destinations)
    d, warnings_list = collect(wb_ciq, wb_data, wb_ims, sgi_pool=p_pool,
                               sgi_v4_line=p_v4_line, gm_v6_line=p_gm_line)
    p_pool = d.get("sgi_pool", p_pool or 1)

    # Build blocks
    blocks = build_blocks(d, dcgw_vprn=dcgw_vprn, static_cfg=s_cfg, has_manual_static=has_manual_static)

    # Create target workbook
    wb_out = openpyxl.Workbook()
    ws = wb_out.active
    ws.title = site_name

    for col, w in COL_WIDTHS.items():
        ws.column_dimensions[col].width = w

    ranges, last_row = write_detail(ws, blocks, site=site_name)
    t1 = write_summary(ws, ranges, CMG_FIRST, "H", "Status from CMG", with_link=True)
    t2 = write_summary(ws, ranges, DCGW_FIRST, "L", "Status from DCGW", with_link=False)

    write_sources(wb_out, fn_map, sgi_pool=p_pool)

    # Save to in-memory bytes
    output_buf = io.BytesIO()
    wb_out.save(output_buf)
    output_buf.seek(0)

    total_rows = sum(len(b["rows"]) for b in blocks)
    stats = {
        "site": site_name,
        "total_blocks": len(blocks),
        "total_detail_rows": total_rows,
        "start_row": DETAIL_FIRST,
        "end_row": last_row,
        "nodes": list(dict.fromkeys(b["node"] for b in blocks)),
        "workbook_size_bytes": len(output_buf.getvalue()),
    }

    # Summary table data for UI preview
    summary_preview = []
    for prio, label, keys in SUMMARY_CONFIG:
        valid_keys = [k for k in keys if k in ranges]
        count = sum(ranges[k][1] - ranges[k][0] + 1 for k in valid_keys) if valid_keys else 0
        summary_preview.append({
            "Priority": prio,
            "Interface": label,
            "Total Targets": count
        })

    # Detail preview for UI (Full rows with Col E & G address tracking)
    detail_preview = []
    for b in blocks:
        for r_item in b["rows"]:
            c_val = str(r_item["c"] or "")
            e_val = str(r_item["e"] or "")
            g_val = str(r_item["g"] or "")
            ping_cmd = f"ping router {c_val} source {e_val} {g_val} rapid count 5" if (c_val and e_val and g_val) else ""
            detail_preview.append({
                "Node": b["node"],
                "Interface": b["label"],
                "CMG VPRN": c_val,
                "Destination": str(r_item["d"] or ""),
                "CMG SRC Address (Col E)": e_val,
                "DCGW VPRN": str(r_item["f"] or ""),
                "DST Address (Col G)": g_val,
                "Generated Ping Command": ping_cmd,
                "Remarks": str(r_item["n"] or r_item["m"] or "")
            })

    return output_buf, stats, warnings_list, summary_preview, detail_preview
