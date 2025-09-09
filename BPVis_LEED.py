import io
import math
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
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
    "MEP Engineer", "Structural Engineer", "Electrical Engineer", "Sustainability Consultant",
    "Facility Manager", "Comissioning Agent", "Infrastructure Engineer",
    "Simulation Expert", "Accoustic Engineer", "Lighting Designer"
]
RESPONSIBLE_DEFAULTS = ["Owner", "Sustainability Consultant", "Architect"]

# =========================
# Sidebar — template download & file upload
# =========================
st.sidebar.image("Pamo_Icon_Black.png", width=80)
st.sidebar.write("## BPVis LEED V5 Precheck")
st.sidebar.write("Version 0.0.1")

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

# --- Ensure default
if "project_name" not in st.session_state:
    st.session_state["project_name"] = "LEED V5 Project"
if "rating_system" not in st.session_state:
    st.session_state["rating_system"] = "BD+C: New Construction"


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

# ========== CREDITS CATALOG (unchanged) ==========
with tab_select:
    with st.expander("Credit Library", expanded=True):
        categories = pd.unique(df['Category'].dropna())
        selected_category = st.selectbox("Select a LEED v5 Category", categories, key="cat")

        credit_rows = (
            df.loc[df['Category'].eq(selected_category), ['Credit_ID', 'Credit_Name']]
            .dropna()
            .drop_duplicates()
        )
        credit_rows['display'] = credit_rows['Credit_ID'].astype(str) + " " + credit_rows['Credit_Name']
        selected_credit_display = st.selectbox(
            f"Select a {selected_category} credit",
            credit_rows['display'],
            key="credit_disp"
        )
        selected_credit_row = credit_rows.loc[credit_rows['display'] == selected_credit_display].iloc[0]
        selected_credit = selected_credit_row['Credit_Name']
        selected_credit_id = selected_credit_row['Credit_ID']

        credit_options = pd.unique(
            df.loc[
                (df['Category'].eq(selected_category)) &
                (df['Credit_ID'].eq(selected_credit_id)) &
                (df['Credit_Name'].eq(selected_credit)),
                'Option_Title'
            ].dropna()
        )
        selected_option = st.selectbox(f"Select a {selected_credit} option", credit_options, key="opt")

        option_paths = pd.unique(
            df.loc[
                (df['Category'].eq(selected_category)) &
                (df['Credit_ID'].eq(selected_credit_id)) &
                (df['Credit_Name'].eq(selected_credit)) &
                (df['Option_Title'].eq(selected_option)),
                'Path_Title'
            ].dropna()
        )
        selected_path = st.selectbox(f"Select a {selected_option} path", option_paths, key="path")

    mask = (
        df['Category'].eq(selected_category) &
        df['Credit_ID'].eq(selected_credit_id) &
        df['Credit_Name'].eq(selected_credit) &
        df['Option_Title'].eq(selected_option) &
        df['Path_Title'].eq(selected_path)
    )

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

        if (('And' in df.columns) and not and_series.empty) or (('Or' in df.columns) and not or_series.empty):
            st.markdown("---")
            st.caption("Dependencies")
            if not and_series.empty:
                and_text = str(and_series.iloc[0]).replace("\r\n", "<br>").replace("\n", "<br>")
                st.markdown(f"**And:**<br>{and_text}", unsafe_allow_html=True)
            if not or_series.empty:
                or_text = str(or_series.iloc[0]).replace("\r\n", "<br>").replace("\n", "<br>")
                st.markdown(f"**Or:**<br>{or_text}", unsafe_allow_html=True)

    if not req_series.empty:
        requirement = str(req_series.iloc[0])
        requirement_html = requirement.replace("\r\n", "<br>").replace("\n", "<br>")
        st.markdown("---")
        st.write("### Requirement:")
        st.markdown(requirement_html, unsafe_allow_html=True)

    if 'Thresholds' in df.columns and not thresholds_series.empty:
        thresholds_html = str(thresholds_series.iloc[0]).replace("\r\n", "<br>").replace("\n", "<br>")
        st.write("### Thresholds:")
        st.markdown(thresholds_html, unsafe_allow_html=True)

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

    st.write("### Design Team Strategy:")
    with st.expander("Approach", expanded=False):
        key_base = f"{selected_category}|{selected_credit_id}|{selected_credit}|{selected_option}|{selected_path}"
        approach_key = f"approach::{key_base}"
        status_key = f"status::{key_base}"
        resp_key = f"responsible::{key_base}"
        effort_key = f"effort::{key_base}"

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

        # CSS: show full text on multiselect chips
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

    with col1:
        if agg.empty:
            st.warning("Nothing to plot yet — mark some items as Pursued and set Planned Points.")
        else:
            # Donut with total in center
            fig = px.pie(
                agg,
                names='Category',
                values='Points',
                hole=0.5,
                color="Category",
                color_discrete_map=color_map,
                height=800
            )
            fig.update_traces(textposition='inside', textinfo='percent+value')
            fig.update_layout(
                annotations=[dict(text=f"{total_points:.0f}", x=0.5, y=0.5,
                                  font_size=70, showarrow=False)]
            )
            st.plotly_chart(fig, use_container_width=True)

    # ---------- Certification + KPIs ----------
    # Thresholds (inclusive at these values)
    thresholds = [
        ("LEED V5 Certified", 40),
        ("LEED V5 Silver", 50),
        ("LEED V5 Gold", 60),
        ("LEED V5 Platinum", 80),
    ]

    # Determine current level
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
    # If already Platinum, no next level
    points_to_next = 0 if next_threshold is None else max(0, next_threshold - total_points)
    safety_points = max(0.0, total_points - current_floor)
    total_percentage = (total_points / 110.0) * 100.0

    with col2:

        if certification_level == "LEED V5 Certified":
            st.image("LEED_Certified.png", width=200)
        if certification_level == "LEED V5 Silver":
            st.image("LEED_Silver.png", width=200)
        if certification_level == "LEED V5 Gold":
            st.image("LEED_Gold.png", width=200)
        if certification_level == "LEED V5 Platinum":
            st.image("LEED_Platinum.png", width=200)
        st.metric("Certification", certification_level)

        # New KPIs
        st.metric("Total Points", f"{total_points:.0f}")
        st.metric("Points to Next Level", "—" if next_threshold is None else f"{points_to_next:.0f}")
        st.metric("Safety Points in Current Level", f"{safety_points:.0f}")
        st.metric("Total Points Percentage Achieved", f"{total_percentage:.0f}%")

    # ---------- Points by Credit (no prerequisites, sorted by Credit_ID descending, legend aligned) ----------
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
            non_prereq['Credit'] = non_prereq['Credit_ID'].astype(str) + " " + non_prereq['Credit_Name'].astype(str)

            per_credit = (
                non_prereq.groupby(['Category', 'Credit_ID', 'Credit_Name'], as_index=False)['Planned_Points']
                .sum()
                .rename(columns={'Planned_Points': 'Points'})
            )
            per_credit['Credit'] = per_credit['Credit_ID'].astype(str) + " " + per_credit['Credit_Name'].astype(str)

            # Breakdown hover
            def row_label(r):
                opt = str(r['Option_Title']) if pd.notna(r['Option_Title']) else ""
                path = str(r['Path_Title']) if pd.notna(r['Path_Title']) else ""
                bits = [b for b in [opt, path] if b and b.lower() != 'nan']
                left = " — ".join(bits) if bits else "(no option/path)"
                return f"{left}: {int(r['Planned_Points'])} pts"

            breakdown_src = non_prereq.copy()
            breakdown_src['__label'] = breakdown_src.apply(row_label, axis=1)
            breakdown = (
                breakdown_src.groupby(['Credit_ID', 'Credit_Name'], as_index=False)['__label']
                .apply(lambda s: "\n".join(sorted({x for x in s if isinstance(x, str) and x.strip()})))
                .rename(columns={'__label': 'Breakdown'})
            )
            per_credit = per_credit.merge(breakdown, on=['Credit_ID', 'Credit_Name'], how='left')

            # Sort by Credit_ID descending
            per_credit = per_credit.sort_values(['Credit_ID'], ascending=False)
            y_order = per_credit['Credit'].tolist()

            bar_height = max(400, min(1000, 40 * len(per_credit) + 120))

            fig_bar = px.bar(
                per_credit,
                x='Points',
                y='Credit',
                color='Category',
                orientation='h',
                color_discrete_map=color_map,
                text='Points',
                hover_data={'Category': True, 'Points': True, 'Credit': False, 'Breakdown': True}
            )
            fig_bar.update_traces(
                texttemplate='%{x:.0f}',
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
                legend_traceorder="reversed"  # keep legend visually aligned with bar order
            )

            st.plotly_chart(fig_bar, use_container_width=True)

    # --- Details table in an expander (includes prerequisites and planned points) ---
    st.write("## Summary")
    if to_sum.empty:
        st.info("No pursued items yet.")
    else:
        with st.expander("Requirements List", expanded=False):
            details = to_sum.copy()

            eff_lookup = st.session_state.df[
                ['Category', 'Credit_ID', 'Credit_Name', 'Option_Title', 'Path_Title',
                 'Max_Points_Effective', 'Thresholds', 'Approach', 'Status', 'Responsible', 'Effort']
            ].copy()
            details = details.merge(
                eff_lookup,
                on=['Category', 'Credit_ID', 'Credit_Name', 'Option_Title', 'Path_Title'],
                how='left'
            )
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
