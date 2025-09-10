import io
import math
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

from streamlit import sidebar

st.set_page_config(
    page_title="WSGT_LEED V5 Precheck Tool",
    page_icon="Pamo_Icon_White.png",
    layout="wide"
)

color_map = {
    "Integrative Process, Planning and Assessments": "#962713",
    "Location and Transportation": "#B87400",
    "Sustainable Sites": "#526325",
    "Water Efficiency": "#026772",
    "Energy and Atmosphere": "#C75227",
    "Materials and Resources": "#7C4900",
    "Indoor Environmental Quality": "#674559",
    "Project Priorities": "#7E8AA2",
}

STATUS_OPTIONS = [
    "Not Pursued",
    "Planned",
    "Under Implementation",
    "Implemented",
    "Verified",
]

EFFORT_OPTIONS = [
    "No Effort",
    "Low Effort",
    "Medium Effort",
    "High Effort",
    "Extrem Effort",
    "Not Feasible",
]

RESPONSIBLE_OPTIONS = [
    "Owner", "Operator", "Architect", "Landscape Architect", "Interior Architect",
    "MEP Engineer", "Structural Engineer", "Electrical Engineer", "Sustainability Consultant", "Building Physics",
    "Energy Engineer","Facility Manager", "Comissioning Agent", "Infrastructure Engineer",
    "Simulation Expert", "Accoustic Engineer", "Lighting Designer", "Contractor"
]
RESPONSIBLE_DEFAULTS = ["Owner", "Sustainability Consultant", "Architect"]

# =========================
# Sidebar — template download & file upload
# =========================
st.sidebar.image("Pamo_Icon_Black.png", width=80)
st.sidebar.write("## BPVis LEED V5 Precheck")
st.sidebar.write("Version 1.0.0")

st.sidebar.markdown("### Download Template")
template_path = Path("templates/LEED v5 BD+C Requirements.xlsx")
if template_path.exists():
    with open(template_path, "rb") as file:
        st.sidebar.download_button(
            label="Download Excel Template",
            data=file.read(),
            file_name="LEED v5 BD+C Requirements.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.sidebar.markdown("---")
st.sidebar.write("### Upload Data")
uploaded_file = st.sidebar.file_uploader("Upload Excel File", type="xlsx")

# =========================
# Header
# =========================
col1, col2 = st.columns(2)
with col2:
    logo_path = Path("WS_Logo.jpg")
    if logo_path.exists():
        st.image(str(logo_path), width=900)

# --- Defaults
if "project_name" not in st.session_state:
    st.session_state["project_name"] = "LEED V5 Project"
if "rating_system" not in st.session_state:
    st.session_state["rating_system"] = "BD+C: New Construction"

# ---------- helper: load Available_Points_Cat ----------
def load_available_points_cat(xls_file):
    """
    Returns a long-form DF with columns:
      Category | Rating_System | Available_Points
    Works with either a wide sheet (Category + rating system columns)
    or a pre-long sheet (Category, Rating_System, Available_Points).
    """
    try:
        ap = pd.read_excel(xls_file, sheet_name="Available_Points_Cat")
    except Exception:
        return None

    # Normalize column names for detection
    ap_cols = [str(c).strip() for c in ap.columns]
    ap.columns = ap_cols
    lower = {c.lower(): c for c in ap_cols}

    # If already long:
    if {"category", "rating_system", "available_points"}.issubset(lower.keys()):
        out = ap[[lower["category"], lower["rating_system"], lower["available_points"]]].copy()
        out.columns = ["Category", "Rating_System", "Available_Points"]
    else:
        # Assume wide: first column is Category (by name or first non-numeric)
        cat_col = None
        for c in ap_cols:
            if c.lower() == "category":
                cat_col = c
                break
        if cat_col is None:
            # fallback: use the first non-numeric column as Category
            non_num_cols = [c for c in ap_cols if not pd.api.types.is_numeric_dtype(ap[c])]
            cat_col = non_num_cols[0] if non_num_cols else ap_cols[0]

        rating_cols = [c for c in ap_cols if c != cat_col]
        out = ap.melt(
            id_vars=[cat_col],
            value_vars=rating_cols,
            var_name="Rating_System",
            value_name="Available_Points"
        )
        out.rename(columns={cat_col: "Category"}, inplace=True)

    # Clean
    out["Category"] = out["Category"].astype(str).str.replace("\u00A0", " ", regex=False).str.strip()
    out["Rating_System"] = out["Rating_System"].astype(str).str.replace("\u00A0", " ", regex=False).str.strip()
    out["Available_Points"] = pd.to_numeric(out["Available_Points"], errors="coerce").fillna(0.0)

    return out

# ---------- NEW helper: load Available_Points_Credit ----------
def load_available_points_credit(xls_file):
    """
    Returns a long-form DF with columns:
      Credit_ID | Rating_System | Max_Credit_Points
    Works with either a wide sheet (Credit_ID + rating system columns)
    or a pre-long sheet (Credit_ID, Rating_System, Max_Credit_Points).
    """
    try:
        apc = pd.read_excel(xls_file, sheet_name="Available_Points_Credit")
    except Exception:
        return None

    apc_cols = [str(c).strip() for c in apc.columns]
    apc.columns = apc_cols
    lower = {c.lower(): c for c in apc_cols}

    # Long already?
    if {"credit_id", "rating_system", "max_credit_points"}.issubset(lower.keys()):
        out = apc[[lower["credit_id"], lower["rating_system"], lower["max_credit_points"]]].copy()
        out.columns = ["Credit_ID", "Rating_System", "Max_Credit_Points"]
    else:
        # Assume wide: first column is Credit_ID (by name or first non-numeric)
        id_col = None
        for c in apc_cols:
            if c.lower() == "credit_id":
                id_col = c
                break
        if id_col is None:
            non_num_cols = [c for c in apc_cols if not pd.api.types.is_numeric_dtype(apc[c])]
            id_col = non_num_cols[0] if non_num_cols else apc_cols[0]

        rating_cols = [c for c in apc_cols if c != id_col]
        out = apc.melt(
            id_vars=[id_col],
            value_vars=rating_cols,
            var_name="Rating_System",
            value_name="Max_Credit_Points"
        )
        out.rename(columns={id_col: "Credit_ID"}, inplace=True)

    # Clean
    out["Credit_ID"] = out["Credit_ID"].astype(str).str.replace("\u00A0", " ", regex=False).str.strip()
    out["Rating_System"] = out["Rating_System"].astype(str).str.replace("\u00A0", " ", regex=False).str.strip()
    out["Max_Credit_Points"] = pd.to_numeric(out["Max_Credit_Points"], errors="coerce").fillna(0)

    return out

# ---------- Load & keep DF in session_state ----------
def load_dataframe(xls_file):
    df = pd.read_excel(xls_file, sheet_name="LEED_V5_Requirements")

    # Try loading project information (sheet 'Project_information') if present
    try:
        proj_info = pd.read_excel(xls_file, sheet_name="Project_information")
        if isinstance(proj_info, pd.DataFrame) and not proj_info.empty:
            cols_lower = {str(c).strip().lower(): c for c in proj_info.columns}
            if 'project_name' in cols_lower or 'rating_system' in cols_lower:
                if 'project_name' in cols_lower:
                    val = proj_info.iloc[0][cols_lower['project_name']]
                    if isinstance(val, str) and val.strip():
                        st.session_state['project_name'] = val
                if 'rating_system' in cols_lower:
                    val = proj_info.iloc[0][cols_lower['rating_system']]
                    if isinstance(val, str) and val.strip():
                        st.session_state['rating_system'] = val
            elif 'key' in cols_lower and 'value' in cols_lower:
                kv = {str(k).strip().lower(): v for k, v in
                      zip(proj_info[cols_lower['key']], proj_info[cols_lower['value']])}
                if 'project_name' in kv and str(kv['project_name']).strip():
                    st.session_state['project_name'] = str(kv['project_name'])
                if 'rating_system' in kv and str(kv['rating_system']).strip():
                    st.session_state['rating_system'] = str(kv['rating_system'])
    except Exception:
        pass

    # Clean up text columns (preserve NaNs!)
    text_cols = [
        'Category', 'Credit_ID', 'Credit_Name',
        'Option_Title', 'Path_Title',
        'Requirement',
        'Thresholds', 'Tresholds', 'And', 'Or',
        'Documentation', 'Referenced_Standards', 'Referenced Standards',
        'Approach', 'Status', 'Responsible', 'Effort'
    ]
    for c in text_cols:
        if c in df.columns:
            na_mask = df[c].isna()
            cleaned = (
                df[c].astype(str)
                .str.replace('\u00A0', ' ', regex=False)
                .str.replace(r'\s+', ' ', regex=True)
                .str.strip()
            )
            df[c] = cleaned
            df.loc[na_mask, c] = np.nan  # restore NaNs

    # Normalize legacy "Tresholds"
    if 'Thresholds' not in df.columns and 'Tresholds' in df.columns:
        df['Thresholds'] = df['Tresholds']

    # Ensure columns exist (defaults handled in UI)
    if 'Approach' not in df.columns:
        df['Approach'] = ""
    if 'Status' not in df.columns:
        df['Status'] = np.nan
    if 'Responsible' not in df.columns:
        df['Responsible'] = np.nan
    if 'Effort' not in df.columns:
        df['Effort'] = np.nan

    # Numeric
    if 'Max_Points' in df.columns:
        df['Max_Points'] = pd.to_numeric(df['Max_Points'], errors='coerce')

    # Pursued
    if 'Pursued' not in df.columns:
        df['Pursued'] = False
    else:
        df['Pursued'] = df['Pursued'].astype(str).str.lower().isin(['true', '1', 'yes', 'y'])

    # Planned points
    if 'Planned_Points' not in df.columns:
        df['Planned_Points'] = 0
    else:
        df['Planned_Points'] = pd.to_numeric(df['Planned_Points'], errors='coerce').fillna(0).astype(int)

    # Effective points helper
    df['_opt_points'] = (
        df.groupby(['Category', 'Credit_ID', 'Credit_Name', 'Option_Title'], dropna=False)['Max_Points']
        .transform('max')
    )
    df['Max_Points_Effective'] = df['Max_Points'].fillna(df['_opt_points'])

    return df


if uploaded_file is not None:
    if 'df' not in st.session_state or st.session_state.get('uploaded_name') != uploaded_file.name:
        st.session_state.df = load_dataframe(uploaded_file)
        st.session_state.uploaded_name = uploaded_file.name
        # Load Available_Points_Cat & Credit (for charts / bars)
        st.session_state.available_points_cat = load_available_points_cat(uploaded_file)
        st.session_state.available_points_credit = load_available_points_credit(uploaded_file)

df = st.session_state.get('df', None)

# Show project information inputs only after file is uploaded
if df is not None:
    st.sidebar.markdown("---")
    st.sidebar.write("### Project Information")

    st.sidebar.text_input("Project Name", key="project_name")
    st.sidebar.selectbox(
        "Rating System",
        ['BD+C: New Construction', 'BD+C: Core and Shell'],
        key="rating_system"
    )

# --- Title reflects current session state value
st.title(st.session_state["project_name"])
if df is not None:
    st.write("LEED V5", st.session_state["rating_system"])

if df is None:
    st.info("Upload an Excel file to begin.")
    st.stop()

# ---------- Tabs ----------
tab_select, tab_dashboard = st.tabs(["Credits Catalog", "Project Dashboard"])

# ========== CREDITS CATALOG ==========
with tab_select:
    # =========================
    # Credit Library (selectors in an expander)
    # =========================
    with st.expander("Credit Library", expanded=True):
        # NEW: checkbox to filter to only pursued
        show_only_pursued_catalog = st.checkbox(
            "Show only pursued credits/options/paths",
            value=False,
            help="Filter the dropdowns to items marked as Pursued."
        )

        # Base source df for the catalog depending on the filter
        src = df[df['Pursued'] == True] if show_only_pursued_catalog else df

        # Categories
        categories = pd.unique(src['Category'].dropna())
        if len(categories) == 0 and show_only_pursued_catalog:
            st.info("No pursued items yet. Showing all categories.")
            src = df
            categories = pd.unique(src['Category'].dropna())

        selected_category = st.selectbox("Select a LEED v5 Category", categories, key="cat")

        # Credits (ID + Name for display)
        cred_src = src.loc[src['Category'].eq(selected_category), ['Credit_ID', 'Credit_Name']].dropna().drop_duplicates()
        if cred_src.empty and show_only_pursued_catalog:
            st.info("No pursued credits for this category. Showing all credits.")
            cred_src = df.loc[df['Category'].eq(selected_category), ['Credit_ID', 'Credit_Name']].dropna().drop_duplicates()

        cred_src['display'] = cred_src['Credit_ID'].astype(str) + " " + cred_src['Credit_Name']
        selected_credit_display = st.selectbox(
            f"Select a {selected_category} credit",
            cred_src['display'],
            key="credit_disp"
        )
        selected_credit_row = cred_src.loc[cred_src['display'] == selected_credit_display].iloc[0]
        selected_credit = selected_credit_row['Credit_Name']
        selected_credit_id = selected_credit_row['Credit_ID']

        # Options
        opt_src = src.loc[
            (src['Category'].eq(selected_category)) &
            (src['Credit_ID'].eq(selected_credit_id)) &
            (src['Credit_Name'].eq(selected_credit)),
            'Option_Title'
        ].dropna().drop_duplicates()
        if opt_src.empty and show_only_pursued_catalog:
            st.info("No pursued options for this credit. Showing all options.")
            opt_src = df.loc[
                (df['Category'].eq(selected_category)) &
                (df['Credit_ID'].eq(selected_credit_id)) &
                (df['Credit_Name'].eq(selected_credit)),
                'Option_Title'
            ].dropna().drop_duplicates()

        selected_option = st.selectbox(f"Select a {selected_credit} option", opt_src, key="opt")

        # Paths
        path_src = src.loc[
            (src['Category'].eq(selected_category)) &
            (src['Credit_ID'].eq(selected_credit_id)) &
            (src['Credit_Name'].eq(selected_credit)) &
            (src['Option_Title'].eq(selected_option)),
            'Path_Title'
        ].dropna().drop_duplicates()
        if path_src.empty and show_only_pursued_catalog:
            st.info("No pursued paths for this option. Showing all paths.")
            path_src = df.loc[
                (df['Category'].eq(selected_category)) &
                (df['Credit_ID'].eq(selected_credit_id)) &
                (df['Credit_Name'].eq(selected_credit)) &
                (df['Option_Title'].eq(selected_option)),
                'Path_Title'
            ].dropna().drop_duplicates()

        selected_path = st.selectbox(f"Select a {selected_option} path", path_src, key="path")

    # --- Final mask for this exact selection ---
    mask = (
        df['Category'].eq(selected_category) &
        df['Credit_ID'].eq(selected_credit_id) &
        df['Credit_Name'].eq(selected_credit) &
        df['Option_Title'].eq(selected_option) &
        df['Path_Title'].eq(selected_path)
    )

    # Pull Requirement / Points / Thresholds / And / Or for this selection
    req_series = df.loc[mask, 'Requirement'].dropna().drop_duplicates()
    pts_series = df.loc[mask, 'Max_Points_Effective'].dropna().drop_duplicates()
    thresholds_series = df.loc[mask, 'Thresholds'].dropna().drop_duplicates() if 'Thresholds' in df.columns else pd.Series([], dtype=str)
    and_series = df.loc[mask, 'And'].dropna().drop_duplicates() if 'And' in df.columns else pd.Series([], dtype=str)
    or_series = df.loc[mask, 'Or'].dropna().drop_duplicates() if 'Or' in df.columns else pd.Series([], dtype=str)

    max_points = int(pts_series.iloc[0]) if not pts_series.empty and not math.isnan(pts_series.iloc[0]) else None

    st.markdown("---")
    colL, colR = st.columns([3, 1])

    with colL:
        st.caption("Credit")
        st.write(f"## {selected_credit_id} {selected_credit}")
        st.caption("Option")
        st.write(f"### {selected_option}")
        st.caption("Path")
        st.write(f"### {selected_path}")

    with colR:
        if max_points is not None:
            st.caption("Max Points")
            st.write(f"## {max_points}")

        # Pursued + Planned Points input
        pursued_default = bool(df.loc[mask, 'Pursued'].any())
        key_base = f"{selected_category}|{selected_credit_id}|{selected_credit}|{selected_option}|{selected_path}"
        pursued_new = st.checkbox(
            "Pursued",
            value=pursued_default,
            help="Mark this path as pursued to include in project's scorecard",
            key=f"pursued::{key_base}"
        )

        planned_default = int(df.loc[mask, 'Planned_Points'].max()) if pursued_default else 0

        if pursued_new:
            min_val = 0
            max_val = max_points if max_points is not None else 100
            if max_points is not None and max_points > 0:
                planned_new = st.number_input(
                    "Planned Points",
                    min_value=min_val,
                    max_value=int(max_val),
                    step=1,
                    value=min(int(planned_default), int(max_val)),
                    key=f"planned::{key_base}",
                    help="Enter the number of points you plan to achieve for this selection"
                )
            else:
                planned_new = 0
                if max_points is not None and max_points == 0:
                    st.caption("This is a prerequisite (0 points)")
        else:
            planned_new = 0

        if pursued_new != pursued_default or planned_new != planned_default:
            idx = st.session_state.df.loc[mask].index
            st.session_state.df.loc[idx, 'Pursued'] = pursued_new
            st.session_state.df.loc[idx, 'Planned_Points'] = int(planned_new)

        # Dependencies
        if (('And' in df.columns) and not and_series.empty) or (('Or' in df.columns) and not or_series.empty):
            st.markdown("---")
            st.caption("Dependencies")
            if not and_series.empty:
                and_text = str(and_series.iloc[0]).replace("\r\n", "<br>").replace("\n", "<br>")
                st.markdown(f"**And:**<br>{and_text}", unsafe_allow_html=True)
            if not or_series.empty:
                or_text = str(or_series.iloc[0]).replace("\r\n", "<br>").replace("\n", "<br>")
                st.markdown(f"**Or:**<br>{or_text}", unsafe_allow_html=True)

    # Requirement
    if not req_series.empty:
        requirement = str(req_series.iloc[0])
        requirement_html = requirement.replace("\r\n", "<br>").replace("\n", "<br>")
        st.markdown("---")
        st.write("### Requirement:")
        st.markdown(requirement_html, unsafe_allow_html=True)

    # Thresholds
    if 'Thresholds' in df.columns and not thresholds_series.empty:
        thresholds_html = str(thresholds_series.iloc[0]).replace("\r\n", "<br>").replace("\n", "<br>")
        st.write("### Thresholds:")
        st.markdown(thresholds_html, unsafe_allow_html=True)

    # Documentation & Referenced Standards
    doc_col = 'Documentation' if 'Documentation' in df.columns else None
    ref_col = 'Referenced_Standards' if 'Referenced_Standards' in df.columns else ('Referenced Standards' if 'Referenced Standards' in df.columns else None)

    if doc_col:
        _doc_series = df.loc[mask, doc_col].dropna().drop_duplicates()
        if not _doc_series.empty:
            st.write("### Documentation:")
            _doc_html = str(_doc_series.iloc[0]).replace("\r\n", "<br>").replace("\n", "<br>")
            st.markdown(_doc_html, unsafe_allow_html=True)

    if ref_col:
        _ref_series = df.loc[mask, ref_col].dropna().drop_duplicates()
        if not _ref_series.empty:
            st.write("### Referenced Standards:")
            _ref_html = str(_ref_series.iloc[0]).replace("\r\n", "<br>").replace("\n", "<br>")
            st.markdown(_ref_html, unsafe_allow_html=True)

    # =========================
    # Design Team Strategy + Approach + Status + Responsible + Effort
    # =========================
    st.write("### Design Team Strategy:")
    with st.expander("Approach", expanded=False):
        key_base = f"{selected_category}|{selected_credit_id}|{selected_credit}|{selected_option}|{selected_path}"
        approach_key = f"approach::{key_base}"
        status_key = f"status::{key_base}"
        resp_key = f"responsible::{key_base}"
        effort_key = f"effort::{key_base}"

        # ---- Approach (text) ----
        current_df_approach = ""
        _ser_appr = df.loc[mask, 'Approach'].dropna().astype(str)
        if not _ser_appr.empty:
            current_df_approach = _ser_appr.iloc[0]

        if approach_key not in st.session_state:
            st.session_state[approach_key] = current_df_approach

        new_text = st.text_area(
            "Project's Strategy",
            key=approach_key,
            height=180,
            help="Describe design team strategy to meet credit's requirement."
        )

        if new_text != current_df_approach:
            idx = st.session_state.df.loc[mask].index
            st.session_state.df.loc[idx, 'Approach'] = new_text
            st.caption("Approach autosaved ✓")

        # ---- Status (dropdown) ----
        current_df_status = "Not Pursued"
        _ser_status = df.loc[mask, 'Status'].dropna().astype(str)
        if not _ser_status.empty and _ser_status.iloc[0] in STATUS_OPTIONS:
            current_df_status = _ser_status.iloc[0]

        if status_key not in st.session_state:
            st.session_state[status_key] = current_df_status

        new_status = st.selectbox(
            "Status",
            STATUS_OPTIONS,
            index=STATUS_OPTIONS.index(st.session_state[status_key]) if st.session_state[status_key] in STATUS_OPTIONS else 0,
            key=status_key,
            help="Track the current progress of this credit/option/path."
        )

        if new_status != current_df_status:
            idx = st.session_state.df.loc[mask].index
            st.session_state.df.loc[idx, 'Status'] = new_status
            st.caption("Status autosaved ✓")

        # ---- Responsible (multi-select) ----
        # CSS hack to show full text in selected tags (chips)
        st.markdown("""
        <style>
        .stMultiSelect div[data-baseweb="tag"] {
          max-width: none !important;
          white-space: normal !important;
        }
        .stMultiSelect div[data-baseweb="tag"] span {
          max-width: none !important;
        }
        </style>
        """, unsafe_allow_html=True)

        def parse_responsible(cell: str) -> list:
            """Split a cell value into a list of valid stakeholders."""
            if not cell or str(cell).strip().lower() in ("nan", "none"):
                return []
            parts = [p.strip() for chunk in str(cell).split(";") for p in chunk.split(",")]
            seen, cleaned = set(), []
            for p in parts:
                if p in RESPONSIBLE_OPTIONS and p not in seen:
                    cleaned.append(p)
                    seen.add(p)
            return cleaned

        current_df_resp = RESPONSIBLE_DEFAULTS.copy()
        _ser_resp = df.loc[mask, 'Responsible'].dropna().astype(str)
        if not _ser_resp.empty:
            parsed = parse_responsible(_ser_resp.iloc[0])
            if parsed:
                current_df_resp = parsed

        if resp_key not in st.session_state:
            st.session_state[resp_key] = current_df_resp

        new_resp = st.multiselect(
            "Responsible stakeholders",
            options=RESPONSIBLE_OPTIONS,
            default=st.session_state[resp_key],
            key=resp_key,
            help="Select all stakeholders responsible for this credit/option/path."
        )

        if new_resp:
            st.markdown(" • " + "\n • ".join(new_resp))

        if sorted(new_resp) != sorted(current_df_resp):
            idx = st.session_state.df.loc[mask].index
            st.session_state.df.loc[idx, 'Responsible'] = "; ".join(new_resp)
            st.caption("Responsible autosaved ✓")

        # ---- Effort (dropdown) ----
        current_df_effort = "No Effort"
        _ser_eff = df.loc[mask, 'Effort'].dropna().astype(str)
        if not _ser_eff.empty and _ser_eff.iloc[0] in EFFORT_OPTIONS:
            current_df_effort = _ser_eff.iloc[0]

        if effort_key not in st.session_state:
            st.session_state[effort_key] = current_df_effort

        new_effort = st.selectbox(
            "Expected Effort",
            EFFORT_OPTIONS,
            index=EFFORT_OPTIONS.index(st.session_state[effort_key]) if st.session_state[effort_key] in EFFORT_OPTIONS else 0,
            key=effort_key,
            help="Estimated effort to implement this credit/option/path."
        )

        if new_effort != current_df_effort:
            idx = st.session_state.df.loc[mask].index
            st.session_state.df.loc[idx, 'Effort'] = new_effort
            st.caption("Effort autosaved ✓")

    # Download updated Excel
    st.markdown("---")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        export_df = st.session_state.df.drop(columns=['_opt_points'], errors='ignore')
        export_df.to_excel(writer, sheet_name="LEED_V5_Requirements", index=False)

        proj_out = pd.DataFrame([{
            "project_name": st.session_state.get("project_name", ""),
            "rating_system": st.session_state.get("rating_system", "")
        }])
        proj_out.to_excel(writer, sheet_name="Project_information", index=False)

        # Export Available_Points_Cat if present
        ap_cat_to_save = st.session_state.get("available_points_cat", None)
        if ap_cat_to_save is not None and not ap_cat_to_save.empty:
            ap_cat_to_save.to_excel(writer, sheet_name="Available_Points_Cat", index=False)

        # Export Available_Points_Credit if present
        ap_credit_to_save = st.session_state.get("available_points_credit", None)
        if ap_credit_to_save is not None and not ap_credit_to_save.empty:
            ap_credit_to_save.to_excel(writer, sheet_name="Available_Points_Credit", index=False)

    st.sidebar.markdown("---")
    st.sidebar.download_button(
        label="Save Project",
        use_container_width=True,
        data=buf.getvalue(),
        file_name="LEED_v5_Requirements_updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.sidebar.markdown("---")
    st.caption(f"Matched rows for this selection: {mask.sum()}")

# ========== PROJECT DASHBOARD ==========
with tab_dashboard:
    st.write("## Points by Category")

    pursued = st.session_state.df[st.session_state.df['Pursued'] == True].copy()

    # Defaults to avoid NameError
    to_sum = pd.DataFrame(columns=[
        'Category', 'Credit_ID', 'Credit_Name', 'Option_Title', 'Path_Title', 'Requirement', 'Planned_Points'
    ])
    agg = pd.DataFrame(columns=["Category", "Points"])
    total_points = 0.0

    if pursued.empty:
        st.info("No paths marked as Pursued yet.")
    else:
        # option-level row if Max_Points is NaN but we have effective points
        pursued['is_option_level'] = pursued['Max_Points'].isna() & pursued['Max_Points_Effective'].notna()

        # Path-level rows keep each pursued path
        rows_path = pursued[~pursued['is_option_level']].copy()

        # Option-level rows: one row per option with max planned points
        rows_option = (
            pursued[pursued['is_option_level']]
            .groupby(['Category', 'Credit_ID', 'Credit_Name', 'Option_Title'], dropna=False, as_index=False)
            .agg({'Planned_Points': 'max'})
        )
        rows_option['Path_Title'] = np.nan
        rows_option['Requirement'] = np.nan

        to_sum = pd.concat(
            [
                rows_path[['Category','Credit_ID','Credit_Name','Option_Title','Path_Title','Requirement','Planned_Points']],
                rows_option[['Category','Credit_ID','Credit_Name','Option_Title','Path_Title','Requirement','Planned_Points']],
            ],
            ignore_index=True
        )

        agg = (
            to_sum.groupby('Category', as_index=False)['Planned_Points']
                 .sum()
                 .rename(columns={'Planned_Points': 'Points'})
        )
        total_points = float(agg['Points'].sum())

    col1, col2 = st.columns([3, 1])

    # ---------- MAIN PIE ----------
    with col1:
        if agg.empty:
            st.warning("Nothing to plot yet — mark some items as Pursued and set Planned Points.")
        else:
            agg_sorted = agg.copy().sort_values("Category", kind="stable")
            categories_sorted_main = sorted(agg_sorted["Category"].unique().tolist())

            fig = px.pie(
                agg_sorted,
                names='Category',
                values='Points',
                hole=0.5,
                color="Category",
                color_discrete_map=color_map,
                height=800,
                category_orders={"Category": categories_sorted_main}
            )
            fig.update_traces(textposition='inside', textinfo='percent+value')
            fig.update_layout(
                annotations=[dict(text=f"{total_points:.0f}", x=0.5, y=0.5,
                                  font_size=80, showarrow=False)]
            )
            st.plotly_chart(fig, use_container_width=True)

    # ---------- Certification + KPIs ----------
    thresholds = [
        ("LEED V5 Certified", 40),
        ("LEED V5 Silver", 50),
        ("LEED V5 Gold", 60),
        ("LEED V5 Platinum", 80),
    ]

    certification_level = "Not Certified"
    current_floor = 0
    next_threshold = None
    for name, t in thresholds:
        if total_points >= t:
            certification_level = name
            current_floor = t
        elif next_threshold is None and total_points < t:
            next_threshold = t
            break

    points_to_next = 0 if next_threshold is None else max(0, next_threshold - total_points)
    safety_points = max(0.0, total_points - current_floor)
    total_percentage = (total_points / 110.0) * 100.0

    with col2:
        st.metric("Total Points", f"{total_points:.0f}")
        if certification_level == "LEED V5 Certified":
            st.image("LEED_Certified.png", width=200)
        if certification_level == "LEED V5 Silver":
            st.image("LEED_Silver.png", width=200)
        if certification_level == "LEED V5 Gold":
            st.image("LEED_Gold.png", width=200)
        if certification_level == "LEED V5 Platinum":
            st.image("LEED_Platinum.png", width=200)
        st.metric("Certification", certification_level)

        st.metric("Points to Next Level", "—" if next_threshold is None else f"{points_to_next:.0f}")
        st.metric("Safety Points", f"{safety_points:.0f}")
        st.metric("Total Percentage", f"{total_percentage:.1f}%")

    # ---------- Points by Credit (enhanced hover: stable + detailed breakdown) ----------
    st.write("## Points by Credit")

    if to_sum.empty:
        st.info("No pursued items to display per-credit bars.")
    else:
        # Exclude prerequisites
        non_prereq = to_sum.merge(
            st.session_state.df[['Category', 'Credit_ID', 'Credit_Name',
                                 'Option_Title', 'Path_Title', 'Max_Points_Effective']],
            on=['Category', 'Credit_ID', 'Credit_Name', 'Option_Title', 'Path_Title'],
            how='left'
        )
        non_prereq = non_prereq[non_prereq['Max_Points_Effective'].fillna(0) > 0]

        if non_prereq.empty:
            st.info("Only prerequisites selected — no credits to plot.")
        else:
            # Concise credit label
            non_prereq['Credit'] = non_prereq['Credit_ID'].astype(str) + " " + non_prereq['Credit_Name'].astype(str)

            # Per-credit totals (planned points)
            per_credit = (
                non_prereq.groupby(['Category', 'Credit_ID', 'Credit_Name'], as_index=False)['Planned_Points']
                .sum()
                .rename(columns={'Planned_Points': 'Points'})
            )
            per_credit['Credit'] = per_credit['Credit_ID'].astype(str) + " " + per_credit['Credit_Name'].astype(str)

            # ---- Max available per credit for current rating system ----
            ap_credit = st.session_state.get("available_points_credit", None)
            if ap_credit is not None and not ap_credit.empty:
                rs = st.session_state.get("rating_system", "BD+C: New Construction")
                ap_credit_rs = ap_credit[ap_credit["Rating_System"] == rs].copy()
                ap_credit_rs["Credit_ID"] = ap_credit_rs["Credit_ID"].astype(str).str.strip()
                per_credit["Credit_ID"] = per_credit["Credit_ID"].astype(str).str.strip()

                per_credit = per_credit.merge(
                    ap_credit_rs[["Credit_ID", "Max_Credit_Points"]],
                    on="Credit_ID",
                    how="left"
                )
            else:
                per_credit["Max_Credit_Points"] = np.nan

            # Text label "x/y" (+ flag if exceeded)
            def mk_label(x, y):
                if pd.isna(y) or y == 0:
                    return f"{int(round(x))}"
                flag = " 🚩" if float(x) > float(y) else ""
                return f"{int(round(x))}/{int(round(y))}{flag}"

            per_credit["LabelText"] = per_credit.apply(
                lambda r: mk_label(r["Points"], r["Max_Credit_Points"]), axis=1
            )

            # === Build stable breakdown per credit (Option/Path lines) ===
            # First sum per Option/Path to avoid duplicates, then assemble lines
            lines_src = (
                non_prereq.groupby(
                    ['Credit_ID', 'Credit_Name', 'Option_Title', 'Path_Title'],
                    as_index=False
                )['Planned_Points'].sum()
            )

            def line_for_row(row):
                opt = str(row['Option_Title']) if pd.notna(row['Option_Title']) else ""
                path = str(row['Path_Title']) if pd.notna(row['Path_Title']) else ""
                left_bits = [b for b in [opt, path] if b and b.lower() != 'nan']
                left = " — ".join(left_bits) if left_bits else "(no option/path)"
                return f"{left}: {int(row['Planned_Points'])} pts"

            lines_src['__line'] = lines_src.apply(line_for_row, axis=1)

            breakdown_df = (
                lines_src.sort_values(['Credit_ID', 'Option_Title', 'Path_Title'])
                         .groupby(['Credit_ID', 'Credit_Name'], as_index=False)['__line']
                         .apply(lambda s: "\n".join(s.tolist()))
                         .rename(columns={'__line': 'BreakdownText'})
            )

            per_credit = per_credit.merge(breakdown_df, on=['Credit_ID', 'Credit_Name'], how='left')
            per_credit['BreakdownText'] = per_credit['BreakdownText'].fillna("(no option/path)")
            per_credit['BreakdownHTML'] = per_credit['BreakdownText'].str.replace("\n", "<br>", regex=False)
            per_credit['CreditLabel'] = per_credit['Credit']  # already "ID Name"

            # Sort by Credit_ID descending (define final plotting order)
            per_credit = per_credit.sort_values(['Credit_ID'], ascending=False)
            y_order = per_credit['Credit'].tolist()

            bar_height = max(400, min(1000, 40 * len(per_credit) + 120))

            # Build bar with customdata so hover is fully controlled & aligned
            fig_bar = px.bar(
                per_credit,
                x='Points',
                y='Credit',
                color='Category',
                orientation='h',
                color_discrete_map=color_map,
                text='LabelText',
                custom_data=['CreditLabel', 'BreakdownHTML'],  # <— stable, aligned to bars
                hover_data={'Category': False}  # we’ll override hovertemplate
            )

            # Clean, bold title + breakdown lines
            fig_bar.update_traces(
                hovertemplate="<b>%{customdata[0]}</b><br>%{customdata[1]}<extra></extra>",
                textposition='outside',
                cliponaxis=False
            )

            fig_bar.update_layout(
                height=bar_height,
                xaxis_title="Planned Points",
                yaxis_title="Credit",
                yaxis=dict(categoryorder='array', categoryarray=y_order),
                margin=dict(l=10, r=30, t=30, b=10),
                legend_title="Category",
                legend_traceorder="reversed"
            )

            st.plotly_chart(fig_bar, use_container_width=True)

    # --- Details table in an expander (includes prerequisites and planned points) ---
    st.write("## Summary")
    if to_sum.empty:
        st.info("No pursued items yet.")
    else:
        with st.expander("Requirements List", expanded=False):
            # Optional: filter to only pursued in the summary table
            show_only_pursued_table = st.checkbox("Show only pursued credits/options/paths", value=False)

            details = to_sum.copy()

            eff_lookup = st.session_state.df[
                ['Category', 'Credit_ID', 'Credit_Name', 'Option_Title', 'Path_Title',
                 'Max_Points_Effective', 'Thresholds', 'Approach', 'Status', 'Responsible', 'Effort', 'Pursued']
            ].copy()
            details = details.merge(
                eff_lookup,
                on=['Category', 'Credit_ID', 'Credit_Name', 'Option_Title', 'Path_Title'],
                how='left'
            )

            if show_only_pursued_table:
                details = details[details['Pursued'] == True]

            details['Type'] = np.where(details['Max_Points_Effective'].fillna(0).eq(0), 'Prerequisite', 'Credit')
            details['Planned_Points'] = details['Planned_Points'].fillna(0).astype(int)

            # Format text columns for display
            details['Requirement'] = (
                details['Requirement'].astype(str)
                .replace("nan", "")
                .str.replace("<br>", "\n", regex=False)
                .str.replace("\r\n", "\n", regex=False)
                .str.replace("\n\n+", "\n", regex=True)
                .str.strip()
            )
            if 'Thresholds' in details.columns:
                details['Thresholds'] = (
                    details['Thresholds'].astype(str)
                    .replace("nan", "")
                    .str.replace("<br>", "\n", regex=False)
                    .str.replace("\r\n", "\n", regex=False)
                    .str.replace("\n\n+", "\n", regex=True)
                    .str.strip()
                )
            if 'Approach' in details.columns:
                details['Approach'] = (
                    details['Approach'].astype(str)
                    .replace("nan", "")
                    .str.replace("<br>", "\n", regex=False)
                    .str.replace("\r\n", "\n", regex=False)
                    .str.strip()
                )

            if 'Responsible' in details.columns:
                details['Responsible'] = details['Responsible'].fillna("").astype(str)
            if 'Effort' in details.columns:
                details['Effort'] = details['Effort'].fillna("No Effort").astype(str)

            details = details[[
                'Category', 'Credit_ID', 'Credit_Name',
                'Option_Title', 'Path_Title',
                'Requirement', 'Thresholds', 'Approach', 'Status', 'Responsible', 'Effort',
                'Type', 'Planned_Points'
            ]].drop_duplicates()

            details = details.sort_values(
                by=['Category', 'Credit_ID', 'Credit_Name', 'Option_Title', 'Path_Title', 'Type']
            ).reset_index(drop=True)

            # Wrap long text
            st.markdown("""
            <style>
            .stDataFrame div[data-testid="stDataFrameCell"]{
                white-space: pre-wrap;
                overflow-wrap: break-word;
            }
            </style>
            """, unsafe_allow_html=True)

            st.dataframe(details, use_container_width=True)

    # ===== Separate pies from Available_Points_Cat (LEED V5 Scorecard) =====
    st.write("## LEED V5 Scorecard")
    ap_cat = st.session_state.get("available_points_cat", None)

    # Build a global A→Z category order from the template
    all_cats_sorted = []
    if ap_cat is not None and not ap_cat.empty:
        all_cats_sorted = sorted(ap_cat["Category"].dropna().astype(str).unique().tolist())

    with st.expander("Available Points per Category", expanded=False):
        if ap_cat is None or ap_cat.empty:
            st.info("Sheet 'Available_Points_Cat' not found or empty in the uploaded Excel.")
        else:
            colA, colB = st.columns(2)

            def plot_available(rs_label, container):
                subset = ap_cat[ap_cat["Rating_System"] == rs_label].copy()
                if subset.empty:
                    container.warning(f"No data for '{rs_label}' in Available_Points_Cat.")
                    return

                # Map available points by category
                m = dict(zip(subset["Category"], subset["Available_Points"]))

                # Use ALL categories from template, sorted A→Z
                labels = all_cats_sorted if all_cats_sorted else sorted(subset["Category"].tolist())
                values = [float(m.get(lbl, 0.0)) for lbl in labels]
                colors = [color_map.get(lbl, "#CCCCCC") for lbl in labels]

                fig_av = go.Figure(go.Pie(
                    labels=labels,
                    values=values,
                    hole=0.5,
                    sort=False,  # keep our alphabetical order
                    textinfo="percent+value",
                    marker=dict(colors=colors),
                    hovertemplate="<b>%{label}</b><br>Available: %{value:.0f}<extra></extra>",
                    showlegend=False
                ))
                fig_av.update_layout(
                    title_text=rs_label,
                    height=600,
                    margin=dict(l=0, r=0, t=40, b=0),
                    annotations=[dict(text="110", x=0.5, y=0.5, font_size=80, showarrow=False)]
                )
                container.plotly_chart(fig_av, use_container_width=True)

            with colA:
                plot_available("BD+C: New Construction", st)
            with colB:
                plot_available("BD+C: Core and Shell", st)

    # ====== NEW: Available Points per Credit (two side-by-side bars) ======
    ap_credit = st.session_state.get("available_points_credit", None)
    with st.expander("Available Points per Credit", expanded=False):
        if ap_credit is None or ap_credit.empty:
            st.info("Sheet 'Available_Points_Credit' not found or empty in the uploaded Excel.")
        else:
            # Build a Credit lookup (Category, Credit_Name) from the main DF
            credit_lookup = (
                st.session_state.df[['Credit_ID', 'Credit_Name', 'Category']]
                .dropna(subset=['Credit_ID'])
                .drop_duplicates(subset=['Credit_ID'])
                .copy()
            )
            credit_lookup['Credit_ID'] = credit_lookup['Credit_ID'].astype(str).str.strip()

            def render_possible_bars(rs_label, container):
                # Filter APC by rating system
                sub = ap_credit[ap_credit['Rating_System'] == rs_label].copy()
                if sub.empty:
                    container.warning(f"No data for '{rs_label}' in Available_Points_Credit.")
                    return
                sub['Credit_ID'] = sub['Credit_ID'].astype(str).str.strip()

                # Merge category & credit name for coloring and labeling
                merged = sub.merge(credit_lookup, on='Credit_ID', how='left')

                # Label text (just the max points)
                merged['Credit'] = merged['Credit_ID'].astype(str) + " " + merged['Credit_Name'].astype(str)

                # Sort by Credit_ID descending to match main bar sorting style
                merged = merged.sort_values(['Credit_ID'], ascending=False)
                y_order = merged['Credit'].tolist()

                # Build bar
                bar_height = max(400, min(1000, 40 * len(merged) + 120))
                fig_possible = px.bar(
                    merged,
                    x='Max_Credit_Points',
                    y='Credit',
                    color='Category',
                    orientation='h',
                    color_discrete_map=color_map,
                    text='Max_Credit_Points',
                    hover_data={
                        'Category': True,
                        'Max_Credit_Points': True,
                        'Credit': False
                    },
                    title=rs_label
                )
                fig_possible.update_traces(
                    texttemplate='%{x:.0f}',
                    textposition='outside',
                    cliponaxis=False
                )
                fig_possible.update_layout(
                    height=bar_height,
                    xaxis_title="Max Points (per credit)",
                    yaxis_title="Credit",
                    yaxis=dict(categoryorder='array', categoryarray=y_order),
                    margin=dict(l=10, r=30, t=40, b=10),
                    legend_title="Category",
                    legend_traceorder="reversed",
                    showlegend=False  # hide legend on these "what's possible" bars
                )
                container.plotly_chart(fig_possible, use_container_width=True)

            c1, c2 = st.columns(2)
            with c1:
                render_possible_bars("BD+C: New Construction", st)
            with c2:
                render_possible_bars("BD+C: Core and Shell", st)

with sidebar:
    st.caption("*A product of*")
    st.image("WS_Logo.png", width=300)
    st.caption("Werner Sobek Green Technologies GmbH")
    st.caption("Fachgruppe Simulation")
    st.markdown("---")
    st.caption("*Coded by*")
    st.caption("Rodrigo Carvalho")
    st.caption("*Need help? Contact me under:*")
    st.caption("*email:* rodrigo.carvalho@wernersobek.com")
    st.caption("*Tel* +49.40.6963863-14")
    st.caption("*Mob* +49.171.964.7850")