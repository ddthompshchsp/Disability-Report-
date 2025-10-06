# app.py — HCHSP Disability Report (GoEngage #10443)
import base64
import io
import re
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

# =========================
# Constants & Page Config
# =========================
ENROLLMENT = 2480
TARGET = int(ENROLLMENT * 0.10)

st.set_page_config(page_title="HCHSP Disability Report — GoEngage #10443", layout="wide")

# =========================
# Logo (local-only, no uploads)
# =========================
def load_logo_bytes() -> bytes | None:
    p = Path("header_logo.png")
    return p.read_bytes() if p.exists() else None

def logo_img_tag_centered(logo_bytes: bytes | None, width_px: int = 220) -> str:
    if not logo_bytes:
        return ""
    b64 = base64.b64encode(logo_bytes).decode("utf-8")
    return (
        f"<div class='hero-logo'><img src='data:image/png;base64,{b64}' "
        f"alt='HCHSP Logo' width='{width_px}'/></div>"
    )

LOGO_BYTES = load_logo_bytes()

# =========================
# Hero Header
# =========================
st.markdown(
    """
    <style>
      .hero {text-align:center; margin-top:16px; margin-bottom:8px;}
      .hero h1 {font-size: 42px; line-height: 1.15; margin: 12px 0 0 0;}
      .hero p.sub {color:#555; margin-top:6px; font-size:14px;}
      .muted {color:#666; font-size: 13px;}
      .hero-logo {display:flex; justify-content:center; margin-top:8px; margin-bottom:6px;}
    </style>
    """,
    unsafe_allow_html=True,
)
ts_page = datetime.now(ZoneInfo("America/Chicago")).strftime("%m/%d/%y %I:%M %p %Z")
st.markdown(
    "<div class='hero'>"
    f"{logo_img_tag_centered(LOGO_BYTES)}"
    "<h1>HCHSP Disability Report (2025–2026)</h1>"
    "<p class='sub'>Upload your GoEngage #10443 export (.xlsx) to generate the formatted report, center totals, and dashboard.</p>"
    f"<p class='muted'>Exported on: {ts_page}</p>"
    "</div>",
    unsafe_allow_html=True,
)

# =========================
# Upload (data file only)
# =========================
st.markdown("### Upload GoEngage Report.xlsx")
uploaded = st.file_uploader(
    "Upload the Disability Report from GoEngage (headers on row 5).",
    type=["xlsx"],
    label_visibility="collapsed",
)

# Read once -> bytes; verify #10443 before processing
if uploaded is None:
    st.info("Upload the raw GEHS Quick Report #10443 (xlsx) to begin.")
    st.stop()

file_bytes = uploaded.read()

# =========================
# Verify correct report: 10443 only
# =========================
try:
    # Look for "10443" in the title block (first few rows, no header)
    peek = pd.read_excel(io.BytesIO(file_bytes), nrows=6, header=None, dtype=str)
    if not peek.astype(str).apply(lambda col: col.str.contains("10443", na=False, case=False)).any().any():
        st.error("🚫 This file does not appear to be the correct GoEngage Disability Report (#10443). Please upload the 10443 export.")
        st.stop()
    else:
        st.caption("✅ Detected GoEngage Report #10443 — ready to process.")
except Exception:
    # If detection fails, let processing continue (the next step will still expect 10443 structure)
    st.warning("⚠️ Could not automatically verify the report number. Proceeding, but this expects #10443.")

# =========================
# Helpers
# =========================
def normalize_pid(x: object) -> str:
    s = str(x)
    digits = re.sub(r"\D", "", s)
    return digits.lstrip("0") or digits

def is_date_header(colname: str) -> bool:
    return bool(re.search(r"(date|form|valid from|valid thru)", str(colname), flags=re.I))

def pick_one_disability(cell: str) -> str:
    if not isinstance(cell, str):
        return "Unspecified"
    first = cell.split(",")[0].strip()
    return first if first else "Unspecified"

def find_col(df: pd.DataFrame, patterns: list[str], prefer: str | None = None) -> str | None:
    if prefer and prefer in df.columns:
        return prefer
    for c in df.columns:
        for p in patterns:
            if re.search(p, str(c), flags=re.I):
                return c
    return None

# =========================
# Main Processing
# =========================
@st.cache_data(show_spinner=False)
def process(file_bytes: bytes):
    # 10443 exports use headers on row 5 (index 4)
    df = pd.read_excel(io.BytesIO(file_bytes), header=4).dropna(how="all")

    # Normalize header names for key columns
    rename_exact = {
        "ST: Participant PID": "PID",
