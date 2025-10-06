# app.py — HCHSP Disability Report
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

st.set_page_config(page_title="HCHSP Disability Report", layout="wide")

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

ts = datetime.now(ZoneInfo("America/Chicago")).strftime("%m/%d/%y %I:%M %p %Z")

st.markdown(
    "<div class='hero'>"
    f"{logo_img_tag_centered(LOGO_BYTES)}"
    "<h1>HCHSP Disability Report (2025–2026)</h1>"
    "<p class='sub'>Upload your GoEngage Report (.xlsx) to generate the formatted report, center totals, and dashboard.</p>"
    f"<p class='muted'>Exported on: {ts}</p>"
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
st.divider()

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
    df = pd.read_excel(io.BytesIO(file_bytes), header=4).dropna(how="all")

    # Normalize header names for key columns
    rename_exact = {
        "ST: Participant PID": "PID",
        "ST: Participant": "Participant",
        "ST: Class Name": "Class",
        "ST: Center Name": "Center",
    }
    df = df.rename(columns=lambda c: rename_exact.get(c, c))

    # Locate key columns (accept 10415 OR 10432)
    pid_col = "PID" if "PID" in df.columns else find_col(df, [r"\bparticipant pid\b", r"\bpid\b"])
    identified_col = find_col(
        df,
        [
            r"^IEP/IFSP Dis:Identified$",
            r"iep/ifsp.*identified",
            r"^Disability Identified$",
            r"disability.*identified",
        ],
        prefer="IEP/IFSP Dis:Identified",
    )
    iep_form_col = find_col(df, [r"^IEP/IFSP:Form Date$", r"iep.*form.*date"], prefer="IEP/IFSP:Form Date")
    auth_col = find_col(df, [r"authorization.*date", r"\bauthorization\b"])
    center_col = "Center" if "Center" in df.columns else find_col(df, [r"center name|campus|site name|location"])

    # Normalize PID for dedupe
    df["PID_norm"] = (df.get(pid_col, df.iloc[:, 0])).apply(normalize_pid)

    # Trust the system report — include all rows
    df["__IncludeFlag"] = True

    # Merge duplicates: left->right, join with commas; format date-like columns
    date_like_cols = [c for c in df.columns if is_date_header(c)]

    def merge_group_ordered(g: pd.DataFrame) -> pd.Series:
        out = {}
        for c in df.columns:
            if c in ["__IncludeFlag"]:
                continue
            vals = g[c].tolist()
            vals = [v for v in vals if not (pd.isna(v) or (isinstance(v, str) and v.strip() == ""))]
            uniq, seen = [], set()
            for v in vals:
                key = str(v)
                if key not in seen:
                    seen.add(key)
                    uniq.append(v)
            if c in date_like_cols:
                fmt = []
                for v in uniq:
                    dt = pd.to_datetime(v, errors="coerce")
                    fmt.append(dt.strftime("%m/%d/%y") if pd.notna(dt) else str(v).strip())
                out[c] = ", ".join([x for x in fmt if x])
            else:
                out[c] = ", ".join([str(x).strip() for x in uniq if str(x).strip()])
        out["PID_norm"] = g["PID_norm"].iloc[0]
        out["__AnyInclude"] = g["__IncludeFlag"].any()
        return pd.Series(out)

    merged = df.groupby("PID_norm", dropna=False, as_index=False).apply(merge_group_ordered)
    clean = merged[merged["__AnyInclude"] == True].copy()

    # Authorization formatting (X for missing)
    if auth_col and auth_col in clean.columns:
        def fmt_auth(val):
            parts = [p.strip() for p in str(val).split(",") if p.strip()]
            if not parts:
                return "X"
            out = []
            for p in parts:
                dt = pd.to_datetime(p, errors="coerce")
                out.append(dt.strftime("%m/%d/%y") if pd.notna(dt) else p)
            return ", ".join(out)
        clean[auth_col] = clean[auth_col].apply(fmt_auth)

    # Column order
    front_cols = [c for c in ["PID", "Participant", "Center", "Class"] if c in clean.columns]
    the_rest = [c for c in df.columns if c not in front_cols and c not in ["__IncludeFlag"]]
    final_cols = front_cols + the_rest + [c for c in clean.columns if c not in front_cols + the_rest]

    # Drop columns R, S, and M (13th col)
    def excel_col_letter(idx_zero_based: int) -> str:
        letters, idx = "", idx_zero_based + 1
        while idx:
            idx, rem = divmod(idx - 1, 26)
            letters = chr(65 + rem) + letters
        return letters
    final_cols = [c for i, c in enumerate(final_cols) if excel_col_letter(i) not in ("R", "S")]
    if len(final_cols) >= 13:
        final_cols = [c for idx, c in enumerate(final_cols) if idx != 12]

    # Center totals
    if center_col and center_col in clean.columns:
        centers = (
            clean.groupby(center_col).size().reset_index(name="Identified").sort_values("Identified", ascending=False)
        )
        centers["% of Enrollment"] = centers["Identified"] / ENROLLMENT
        centers["% Campus Contribution to 10% Goal"] = centers["Identified"] / TARGET
    else:
        centers = pd.DataFrame(columns=["Center", "Identified", "% of Enrollment", "% Campus Contribution to 10% Goal"])

    # Disability breakdown (unique per student)
    if identified_col and identified_col in clean.columns:
        one_dis = clean[[identified_col]].copy()
        one_dis[identified_col] = one_dis[identified_col].apply(pick_one_disability).str.strip().str.title()
        disab_breakdown = one_dis[identified_col].value_counts(dropna=False).reset_index()
        disab_breakdown.columns = ["Disability Type", "Count"]
    else:
        disab_breakdown = pd.DataFrame(columns=["Disability Type", "Count"])

    return clean[final_cols], centers, disab_breakdown, (auth_col if (auth_col and auth_col in final_cols) else None)
