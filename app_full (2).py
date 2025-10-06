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

@st.cache_data(show_spinner=False)
def process(file_bytes: bytes):
    # Load with row 5 as header (index=4)
    df = pd.read_excel(io.BytesIO(file_bytes), header=4).dropna(how="all")

    # Normalize header names for key columns
    rename_exact = {
        "ST: Participant PID": "PID",
        "ST: Participant": "Participant",
        "ST: Class Name": "Class",
        "ST: Center Name": "Center",
    }
    df = df.rename(columns=lambda c: rename_exact.get(c, c))

    # Locate key columns
    pid_col = "PID" if "PID" in df.columns else find_col(df, [r"\bparticipant pid\b", r"\bpid\b"])
    identified_col = find_col(df, [r"^IEP/IFSP Dis:Identified$", r"iep/ifsp.*identified"], prefer="IEP/IFSP Dis:Identified")
    iep_form_col = find_col(df, [r"^IEP/IFSP:Form Date$", r"iep.*form.*date"], prefer="IEP/IFSP:Form Date")
    auth_col = find_col(df, [r"authorization.*date", r"\bauthorization\b"])
    center_col = "Center" if "Center" in df.columns else find_col(df, [r"center name|campus|site name|location"])

    # Normalize PID for dedupe
    df["PID_norm"] = (df.get(pid_col, df.iloc[:, 0])).apply(normalize_pid)

    # Inclusion: must have an IEP/IFSP Form Date
    has_iep_date = (
        pd.to_datetime(df[iep_form_col], errors="coerce").notna()
        if iep_form_col in df.columns
        else pd.Series(False, index=df.index)
    )
    df["__IncludeFlag"] = has_iep_date

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

    # Authorization formatting (UI shows value; Excel adds red X via conditional formatting)
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

    # Drop Excel columns R & S by position and also drop column M (index 12)
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
        # keep these two names in this order so the chart points to the 4th column (index 3)
        centers["% of Enrollment"] = centers["Identified"] / ENROLLMENT
        centers["% of 10% Target (248)"] = centers["Identified"] / TARGET
    else:
        centers = pd.DataFrame(
            columns=["Center", "Identified", "% of Enrollment", "% of 10% Target (248)"]
        )

    # Disability breakdown (unique per student)
    identified_col = find_col(clean, [r"^IEP/IFSP Dis:Identified$", r"iep/ifsp.*identified"])
    if identified_col and identified_col in clean.columns:
        one_dis = clean[[identified_col]].copy()
        one_dis[identified_col] = one_dis[identified_col].apply(pick_one_disability).str.strip().str.title()
        disab_breakdown = one_dis[identified_col].value_counts(dropna=False).reset_index()
        disab_breakdown.columns = ["Disability Type", "Count"]
    else:
        disab_breakdown = pd.DataFrame(columns=["Disability Type", "Count"])

    return clean[final_cols], centers, disab_breakdown, (auth_col if (auth_col and auth_col in final_cols) else None)

def build_excel(summary_df: pd.DataFrame, centers_df: pd.DataFrame, disab_df: pd.DataFrame,
                auth_col_name: str | None, logo_bytes: bytes | None) -> io.BytesIO:
    import xlsxwriter

    out_buf = io.BytesIO()
    with pd.ExcelWriter(out_buf, engine="xlsxwriter") as writer:
        wb = writer.book
        header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#305496", "border": 1, "align": "center"})
        bold = wb.add_format({"bold": True})
        pct_fmt = wb.add_format({"num_format": "0.00%"})
        redx_fmt = wb.add_format({"font_color": "#C00000", "bold": True})
        title_fmt = wb.add_format({"bold": True, "font_size": 16, "font_color": "white", "bg_color": "#305496", "align": "center"})
        sub_fmt = wb.add_format({"italic": True, "font_size": 9, "align": "center"})
        val_fmt = wb.add_format({"bold": True, "border": 1})

        ts_local = datetime.now(ZoneInfo("America/Chicago")).strftime("%m/%d/%y %I:%M %p %Z")

        # Sheet1: Disability Summary
        summary_df.to_excel(writer, sheet_name="Disability Summary", index=False, startrow=5)
        ws1 = writer.sheets["Disability Summary"]
        ws1.merge_range(0, 1, 1, max(1, summary_df.shape[1] - 1), "HCHSP – Disability Report", title_fmt)
        ws1.merge_range(2, 1, 2, max(1, summary_df.shape[1] - 1), f"Exported on: {ts_local}", sub_fmt)
        if logo_bytes:
            ws1.insert_image(0, 0, "logo.png", {"image_data": io.BytesIO(logo_bytes), "x_scale": 0.4, "y_scale": 0.4})
        for j, h in enumerate(summary_df.columns):
            ws1.write(5, j, h, header_fmt)
            ws1.set_column(j, j, 22)
        ws1.autofilter(5, 0, 5 + len(summary_df), max(0, summary_df.shape[1] - 1))
        ws1.freeze_panes(6, 0)
        if auth_col_name is not None and auth_col_name in summary_df.columns:
            auth_idx = list(summary_df.columns).index(auth_col_name)
            ws1.conditional_format(6, auth_idx, 6 + len(summary_df), auth_idx,
                                   {"type": "cell", "criteria": "equal to", "value": '"X"', "format": redx_fmt})
        lastrow = 6 + len(summary_df)
        ws1.write(lastrow + 1, 0, "Agency Total Identified", bold)
        ws1.write_formula(lastrow + 1, 1, f"=SUBTOTAL(3,A7:A{lastrow})", bold)
        ws1.write(lastrow + 2, 0, "Agency % of Enrollment (2480)", bold)
        ws1.write_formula(lastrow + 2, 1, f"=SUBTOTAL(3,A7:A{lastrow})/2480", pct_fmt)

        # Sheet2: Center Totals
        centers_df.to_excel(writer, sheet_name="Center Totals", index=False, startrow=5)
        ws2 = writer.sheets["Center Totals"]
        ws2.merge_range(0, 1, 1, 10, "Center Totals (Goal & Agency Share)", title_fmt)
        ws2.merge_range(2, 1, 2, 10, f"Exported on: {ts_local}", sub_fmt)
        if logo_bytes:
            ws2.insert_image(0, 0, "logo.png", {"image_data": io.BytesIO(logo_bytes), "x_scale": 0.4, "y_scale": 0.4})
        for j, h in enumerate(centers_df.columns):
            ws2.write(5, j, h, header_fmt)
            ws2.set_column(j, j, 32 if j == 0 else 24, pct_fmt if j >= 2 else None)
        centers_n = len(centers_df)
        ws2.autofilter(5, 0, 5 + centers_n, len(centers_df.columns) - 1)
        end = 6 + centers_n
        ws2.write(end + 2, 0, "AGENCY TOTAL COUNT", bold)
        ws2.write_formula(end + 2, 1, f"=SUBTOTAL(9,B7:B{end})", bold)
        ws2.write(end + 3, 0, "% of 10% Target (248)", bold)
        ws2.write_formula(end + 3, 1, f"=SUBTOTAL(9,B7:B{end})/248", pct_fmt)
        ws2.write(end + 4, 0, "% of Enrollment", bold)
        ws2.write_formula(end + 4, 1, f"=SUBTOTAL(9,B7:B{end})/2480", pct_fmt)

        # Sheet3: Dashboard
        ws3 = wb.add_worksheet("Dashboard")
        ws3.merge_range(0, 1, 1, 12, "HCHSP — Agency Dashboard", title_fmt)
        ws3.merge_range(2, 1, 2, 12, f"Exported on: {ts_local}", sub_fmt)
        if logo_bytes:
            ws3.insert_image(0, 0, "logo.png", {"image_data": io.BytesIO(logo_bytes), "x_scale": 0.4, "y_scale": 0.4})
        ws3.write(4, 0, "Program Enrollment", bold); ws3.write(4, 1, ENROLLMENT, val_fmt)
        ws3.write(5, 0, "Target (10%)", bold);       ws3.write(5, 1, TARGET, val_fmt)
        ws3.write(6, 0, "Current Identified", bold); ws3.write_formula(6, 1, f"=SUBTOTAL(3,'Disability Summary'!A7:A{lastrow})", val_fmt)
        ws3.write(7, 0, "% of Enrollment", bold);    ws3.write_formula(7, 1, f"=SUBTOTAL(3,'Disability Summary'!A7:A{lastrow})/2480", pct_fmt)

        # Chart 1: Identified vs Target
        chart1 = wb.add_chart({"type": "column"})
        chart1.add_series({"name": "Current Identified", "categories": ["Dashboard", 6, 0, 6, 0], "values": ["Dashboard", 6, 1, 6, 1], "data_labels": {"value": True, "font": {"bold": True}}})
        chart1.add_series({"name": "Target (10%)", "categories": ["Dashboard", 5, 0, 5, 0], "values": ["Dashboard", 5, 1, 5, 1], "data_labels": {"value": True, "font": {"bold": True}}})
        chart1.set_title({"name": "Identified vs Target — Agency"})
        chart1.set_legend({"position": "bottom"})
        ws3.insert_chart(4, 4, chart1, {"x_scale": 1.2, "y_scale": 1.2})

        # Chart 2: Campus Contribution toward Program 10% Goal
        centers_n = len(centers_df)
        if centers_n > 0:
            chart2 = wb.add_chart({"type": "bar"})
            chart2.add_series({
                "name": "% of 10% Target (248)",
                "categories": ["Center Totals", 6, 0, 6 + centers_n - 1, 0],
                "values": ["Center Totals", 6, 3, 6 + centers_n - 1, 3],
                "data_labels": {"value": True, "font": {"bold": True}},
            })
            chart2.set_title({"name": "Percentage of Enrolled Children w/ Disabilities by Campus"})
            chart2.set_legend({"none": True})
            ws3.insert_chart(12, 0, chart2, {"x_scale": 1.2, "y_scale": 1.2})

        # Chart 3: Disability Type Breakdown
        if not disab_df.empty:
            drow = 12 + (centers_n // 2) + 8
            disab_breakdown = disab_df.copy()
            disab_breakdown.to_excel(writer, sheet_name="Dashboard", index=False, startrow=drow, startcol=0)
            for j, h in enumerate(disab_breakdown.columns):
                ws3.write(drow, j, h, header_fmt)
            chart3 = wb.add_chart({"type": "column"})
            chart3.add_series({
                "name": "Count",
                "categories": ["Dashboard", drow + 1, 0, drow + len(disab_breakdown), 0],
                "values": ["Dashboard", drow + 1, 1, drow + len(disab_breakdown), 1],
                "data_labels": {"value": True, "font": {"bold": True}},
            })
            chart3.set_title({"name": "Disability Type Breakdown (Unique per Student)"})
            chart3.set_legend({"none": True})
            ws3.insert_chart(drow, 4, chart3, {"x_scale": 1.2, "y_scale": 1.1})

    out_buf.seek(0)
    return out_buf

# =========================
# Process + UI
# =========================
if uploaded is None:
    st.info("Upload the raw GEHS Quick Report (xlsx) to begin.")
    st.stop()

df_summary, df_centers, df_disab, auth_col_name = process(uploaded.read())

# KPIs
cA, cB, cC = st.columns(3)
with cA:
    st.metric("Current Identified", len(df_summary))
with cB:
    st.metric("Target (10%)", TARGET)
with cC:
    st.metric("% of Enrollment", f"{(len(df_summary)/ENROLLMENT):.2%}")

# Filters and Tabs
centers_list = sorted(df_summary["Center"].dropna().unique().tolist()) if "Center" in df_summary.columns else []
sel_centers = st.multiselect(
    "Filter by Center(s)",
    centers_list,
    default=centers_list,
    placeholder="Choose one or more centers...",
)
df_view = df_summary[df_summary["Center"].isin(sel_centers)].copy() if sel_centers else df_summary.copy()

tab_dash, tab_centers, tab_summary, tab_export = st.tabs([" Dashboard (Preview)", "Center Totals", " Disability Summary", "⬇️ Export"])

with tab_dash:
    st.caption("Full charts appear in the Excel export with labels and formatting.")
    st.write("**Top Centers (by identified count)**")
    st.dataframe(df_centers.head(10), use_container_width=True)
    st.write("**Disability Type Breakdown (unique)**")
    st.dataframe(df_disab, use_container_width=True)

with tab_centers:
    st.dataframe(df_centers, use_container_width=True)

with tab_summary:
    st.dataframe(df_view, use_container_width=True)

with tab_export:
    xlsx = build_excel(df_summary, df_centers, df_disab, auth_col_name, LOGO_BYTES)
    st.download_button(
        "Download Excel Export",
        data=xlsx,
        file_name=f"HCHSP_Disability_Export_{datetime.now(ZoneInfo('America/Chicago')).strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
