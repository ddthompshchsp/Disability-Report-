
import io
import re
from datetime import datetime
import pandas as pd
import streamlit as st

ENROLLMENT = 2480
TARGET = int(ENROLLMENT * 0.10)

st.set_page_config(page_title="HCHSP – Disability Report", layout="wide")

st.title("HCHSP – Disability Report")
st.caption(datetime.now().strftime("Exported on: %m/%d/%y %I:%M %p (Central Time)"))

uploaded = st.file_uploader("Upload GEHS Quick Report (raw) – headers on row 5", type=["xlsx"])
logo = st.file_uploader("Optional: Upload logo (PNG)", type=["png"])

def normalize_pid(x: object) -> str:
    s = str(x)
    digits = re.sub(r"\D", "", s)
    return digits.lstrip("0") or digits

def is_date_header(colname: str) -> bool:
    import re as _re
    return bool(_re.search(r"(date|form|valid from|valid thru)", str(colname), flags=_re.I))

def pick_one_disability(cell: str) -> str:
    if not isinstance(cell, str):
        return "Unspecified"
    first = cell.split(",")[0].strip()
    return first if first else "Unspecified"

@st.cache_data
def process(file_bytes: bytes, logo_bytes: bytes | None):
    df = pd.read_excel(io.BytesIO(file_bytes), header=4).dropna(how="all")
    rename_exact = {
        "ST: Participant PID": "PID",
        "ST: Participant": "Participant",
        "ST: Class Name": "Class",
        "ST: Center Name": "Center",
    }
    df = df.rename(columns=lambda c: rename_exact.get(c, c))

    def find_col(patterns, prefer=None):
        import re as _re
        if prefer and prefer in df.columns:
            return prefer
        for c in df.columns:
            for p in patterns:
                if _re.search(p, str(c), flags=_re.I):
                    return c
        return None

    pid_col = "PID" if "PID" in df.columns else find_col([r"\\bparticipant pid\\b", r"\\bpid\\b"])
    identified_col = find_col([r"^IEP/IFSP Dis:Identified$", r"iep/ifsp.*identified"], prefer="IEP/IFSP Dis:Identified")
    iep_form_col = find_col([r"^IEP/IFSP:Form Date$", r"iep.*form.*date"], prefer="IEP/IFSP:Form Date")
    auth_col = find_col([r"authorization.*date", r"\\bauthorization\\b"])
    center_col = "Center" if "Center" in df.columns else find_col([r"center name|campus|site name|location"])

    df["PID_norm"] = df.get(pid_col, df.iloc[:,0]).apply(normalize_pid)
    has_iep_date = pd.to_datetime(df[iep_form_col], errors="coerce").notna() if iep_form_col in df.columns else pd.Series(False, index=df.index)
    df["__IncludeFlag"] = has_iep_date

    def merge_group_ordered(g: pd.DataFrame) -> pd.Series:
        out = {}
        for c in df.columns:
            if c in ["__IncludeFlag"]:
                continue
            vals = g[c].tolist()
            norm = []
            for v in vals:
                if pd.isna(v) or (isinstance(v, str) and v.strip()==""):
                    continue
                norm.append(v)
            uniq, seen = [], set()
            for v in norm:
                key = str(v)
                if key not in seen:
                    seen.add(key)
                    uniq.append(v)
            if is_date_header(c):
                fmt_vals = []
                for v in uniq:
                    dt = pd.to_datetime(v, errors="coerce")
                    fmt_vals.append(dt.strftime("%m/%d/%y") if pd.notna(dt) else str(v).strip())
                out[c] = ", ".join([x for x in fmt_vals if x])
            else:
                out[c] = ", ".join([str(x).strip() for x in uniq if str(x).strip()])
        out["PID_norm"] = g["PID_norm"].iloc[0]
        out["__AnyInclude"] = g["__IncludeFlag"].any()
        return pd.Series(out)

    merged = df.groupby("PID_norm", dropna=False, as_index=False).apply(merge_group_ordered)
    clean = merged[merged["__AnyInclude"] == True].copy()

    if auth_col and auth_col in clean.columns:
        def fmt_auth(val):
            parts = [p.strip() for p in str(val).split(",") if p.strip()]
            if not parts: return "X"
            out = []
            for p in parts:
                dt = pd.to_datetime(p, errors="coerce")
                out.append(dt.strftime("%m/%d/%y") if pd.notna(dt) else p)
            return ", ".join(out)
        clean[auth_col] = clean[auth_col].apply(fmt_auth)

    front_cols = [c for c in ["PID","Participant","Center","Class"] if c in clean.columns]
    the_rest = [c for c in df.columns if c not in front_cols and c not in ["__IncludeFlag"]]
    final_cols = front_cols + the_rest + [c for c in clean.columns if c not in front_cols + the_rest]

    # Drop R, S, and M
    def excel_col_letter(idx_zero_based):
        letters = ""
        idx = idx_zero_based + 1
        while idx:
            idx, rem = divmod(idx - 1, 26)
            letters = chr(65 + rem) + letters
        return letters
    final_cols = [c for i, c in enumerate(final_cols) if excel_col_letter(i) not in ("R","S")]
    if len(final_cols) >= 13:
        final_cols = [c for idx, c in enumerate(final_cols) if idx != 12]

    # Center totals
    if center_col and center_col in clean.columns:
        centers = clean.groupby(center_col).size().reset_index(name="Identified").sort_values("Identified", ascending=False)
        centers["Percentage of 10% Goal per Campus (Internal Goal)"] = centers["Identified"] / ENROLLMENT
        centers["% of 10% Target (248)"] = centers["Identified"] / TARGET
    else:
        centers = pd.DataFrame(columns=["Center","Identified","Percentage of 10% Goal per Campus (Internal Goal)","% of 10% Target (248)"])

    # Disability breakdown: unique + 'Unspecified' for blanks
    if identified_col and identified_col in clean.columns:
        one_dis = clean[[identified_col]].copy()
        one_dis[identified_col] = one_dis[identified_col].apply(pick_one_disability).str.strip().str.title()
        disab_breakdown = one_dis[identified_col].value_counts(dropna=False).reset_index()
        disab_breakdown.columns = ["Disability Type","Count"]
    else:
        disab_breakdown = pd.DataFrame(columns=["Disability Type","Count"])

    return clean[final_cols], centers, disab_breakdown

if uploaded:
    df_summary, df_centers, df_disab = process(uploaded.read(), logo.read() if logo else None)

    # Filters
    st.subheader("Filters")
    centers_list = sorted(df_summary["Center"].dropna().unique().tolist()) if "Center" in df_summary.columns else []
    sel_centers = st.multiselect("Filter by Center(s)", centers_list, default=centers_list)

    if sel_centers:
        df_view = df_summary[df_summary["Center"].isin(sel_centers)].copy()
    else:
        df_view = df_summary.copy()

    # KPI
    current_identified = len(df_view)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Current Identified", current_identified)
    with col2:
        st.metric("Target (10%)", TARGET)
    with col3:
        st.metric("% of Enrollment", f"{current_identified/ENROLLMENT:.2%}")

    st.write("### Disability Summary (filtered)")
    st.dataframe(df_view, use_container_width=True)

    st.write("### Center Totals (filterable)")
    st.dataframe(df_centers, use_container_width=True)

    st.write("### Disability Type Breakdown (Unique per Student; 'Unspecified' filled)")
    st.dataframe(df_disab, use_container_width=True)

else:
    st.info("Upload the raw GEHS Quick Report (xlsx) to generate the dashboard.")
