import io
import math
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import re
from pathlib import Path

from streamlit import sidebar

# ==== NEW: for PDF & static chart export ====
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, Image as RLImage
from reportlab.lib import colors
from reportlab.lib.units import mm
import plotly.io as pio  # needs kaleido installed

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
    "Owner", "Project Manager", "Operator", "Architect", "Landscape Architect", "Interior Architect",
    "MEP Engineer", "Structural Engineer", "Electrical Engineer", "Sustainability Consultant", "Building Physics",
    "Energy Engineer","Facility Manager", "Comissioning Authority", "Infrastructure Engineer",
    "Simulation Expert", "Accoustic Engineer", "Lighting Designer", "Contractor"
]
RESPONSIBLE_DEFAULTS = ["Owner", "Sustainability Consultant", "Architect"]

IMPLEMENTATION_PHASES = [f"LPH{i}" for i in range(1, 10)]
DEFAULT_IMPLEMENTATION_PHASE = "LPH3"


# =======================
# Helpers
# =======================
def split_tokens(cell) -> list:
    """
    Split a delimiter-separated cell (commas/semicolons) into clean, de-duplicated tokens.
    Returns [] for empty / NaN-like values.
    """
    if cell is None:
        return []
    try:
        # pandas NaN support
        import pandas as _pd
        if _pd.isna(cell):
            return []
    except Exception:
        pass

    s = str(cell).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return []

    parts = [p.strip() for p in s.replace(";", ",").split(",")]
    out, seen = [], set()
    for p in parts:
        if not p:
            continue
        if p not in seen:
            out.append(p)
            seen.add(p)
    return out


def parse_responsible_valid(cell) -> list:
    """Return only tokens that match RESPONSIBLE_OPTIONS (order preserved)."""
    toks = split_tokens(cell)
    return [t for t in toks if t in RESPONSIBLE_OPTIONS]


def responsible_matches(cell, selected: list, mode: str = "any") -> bool:
    """
    True if the Responsible cell matches the selected stakeholders.
    mode: "any" (intersection) or "all" (subset).
    """
    if not selected:
        return True
    toks = set(split_tokens(cell))
    sel = set(selected)
    if mode == "all":
        return sel.issubset(toks)
    return bool(toks.intersection(sel))


def comment_colname(stakeholder: str) -> str:
    """Column name for stakeholder-specific comments (Excel-safe, deterministic)."""
    safe = re.sub(r'[^0-9A-Za-z]+', '_', str(stakeholder)).strip('_')
    return f"Comments_{safe}"

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
        'Approach', 'Status', 'Responsible', 'Effort', 'Implementation_Phase'
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


    # Implementation Phase
    if 'Implementation_Phase' not in df.columns:
        df['Implementation_Phase'] = DEFAULT_IMPLEMENTATION_PHASE
    df['Implementation_Phase'] = df['Implementation_Phase'].fillna('').astype(str).str.strip()
    _mask_blank_phase = (df['Implementation_Phase'].eq('') | df['Implementation_Phase'].str.lower().isin(['nan','none','null']))
    df.loc[_mask_blank_phase, 'Implementation_Phase'] = DEFAULT_IMPLEMENTATION_PHASE
    # Validate against allowed phases; fallback to default
    df.loc[~df['Implementation_Phase'].isin(IMPLEMENTATION_PHASES), 'Implementation_Phase'] = DEFAULT_IMPLEMENTATION_PHASE

    # Stakeholder comments (one column per stakeholder)
    for _stakeholder in RESPONSIBLE_OPTIONS:
        _col = comment_colname(_stakeholder)
        if _col not in df.columns:
            df[_col] = np.nan

    # Clean comment columns (preserve NaNs)
    for _col in [comment_colname(s) for s in RESPONSIBLE_OPTIONS if comment_colname(s) in df.columns]:
        _na = df[_col].isna()
        _cleaned = (
            df[_col].astype(str)
            .str.replace('\u00A0', ' ', regex=False)
            .str.replace(r'\s+', ' ', regex=True)
            .str.strip()
        )
        df[_col] = _cleaned
        df.loc[_na, _col] = np.nan

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

# ==== PDF / REPORT HELPERS ====================================================
def _htmlize(text: str) -> str:
    """Convert rich cell content into safe, simple HTML paragraphs for ReportLab."""
    if text is None:
        return ""
    s = str(text)
    s = s.replace("\r\n", "\n").replace("<BR>", "\n").replace("<br/>", "\n").replace("<br>", "\n")
    # Escape & < >
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    # restore intended breaks and convert to <br/>
    s = s.replace("&lt;br&gt;", "\n").replace("&lt;br/&gt;", "\n")
    s = s.replace("\n", "<br/>")
    return s

def _compute_dashboard_data(df_: pd.DataFrame):
    """Reproduce the same aggregation used in the dashboard (Planned points by Category, etc.)."""
    pursued = df_[df_['Pursued'] == True].copy()
    to_sum = pd.DataFrame(columns=['Category','Credit_ID','Credit_Name','Option_Title','Path_Title','Requirement','Planned_Points'])
    agg = pd.DataFrame(columns=['Category','Points'])
    total = 0.0

    if not pursued.empty:
        pursued['is_option_level'] = pursued['Max_Points'].isna() & pursued['Max_Points_Effective'].notna()
        rows_path = pursued[~pursued['is_option_level']].copy()
        rows_option = (
            pursued[pursued['is_option_level']]
            .groupby(['Category','Credit_ID','Credit_Name','Option_Title'], dropna=False, as_index=False)
            .agg({'Planned_Points':'max'})
        )
        rows_option['Path_Title'] = np.nan
        rows_option['Requirement'] = np.nan

        to_sum = pd.concat(
            [rows_path[['Category','Credit_ID','Credit_Name','Option_Title','Path_Title','Requirement','Planned_Points']],
             rows_option[['Category','Credit_ID','Credit_Name','Option_Title','Path_Title','Requirement','Planned_Points']]],
            ignore_index=True
        )
        agg = (to_sum.groupby('Category', as_index=False)['Planned_Points']
               .sum().rename(columns={'Planned_Points':'Points'}))
        total = float(agg['Points'].sum())

    # Build per-credit for the bars + labels (same as UI)
    per_credit = pd.DataFrame()
    if not to_sum.empty:
        non_prereq = to_sum.merge(
            df_[['Category','Credit_ID','Credit_Name','Option_Title','Path_Title','Max_Points_Effective']],
            on=['Category','Credit_ID','Credit_Name','Option_Title','Path_Title'],
            how='left'
        )
        non_prereq = non_prereq[non_prereq['Max_Points_Effective'].fillna(0) > 0]
        if not non_prereq.empty:
            per_credit = (
                non_prereq.groupby(['Category','Credit_ID','Credit_Name'], as_index=False)['Planned_Points']
                .sum().rename(columns={'Planned_Points':'Points'})
            )
            per_credit['Credit'] = per_credit['Credit_ID'].astype(str) + " " + per_credit['Credit_Name'].astype(str)

            ap_credit = st.session_state.get("available_points_credit", None)
            if ap_credit is not None and not ap_credit.empty:
                rs = st.session_state.get("rating_system", "BD+C: New Construction")
                ap_credit_rs = ap_credit[ap_credit["Rating_System"] == rs].copy()
                ap_credit_rs["Credit_ID"] = ap_credit_rs["Credit_ID"].astype(str).str.strip()
                per_credit["Credit_ID"] = per_credit["Credit_ID"].astype(str).str.strip()
                per_credit = per_credit.merge(
                    ap_credit_rs[['Credit_ID','Max_Credit_Points']],
                    on='Credit_ID', how='left'
                )
            else:
                per_credit['Max_Credit_Points'] = np.nan

            def mk_label(x,y):
                if pd.isna(y) or y == 0:
                    return f"{int(round(x))}"
                flag = " 🚩" if float(x) > float(y) else ""
                return f"{int(round(x))}/{int(round(y))}{flag}"
            per_credit["LabelText"] = per_credit.apply(lambda r: mk_label(r["Points"], r["Max_Credit_Points"]), axis=1)
            per_credit = per_credit.sort_values(['Credit_ID'], ascending=False)
    return to_sum, agg, total, per_credit

def _figure_bytes(fig, width=None, height=None, scale=2):
    """Return PNG bytes from a Plotly fig using kaleido; on failure return None."""
    try:
        return fig.to_image(format="png", width=width, height=height, scale=scale)
    except Exception:
        return None

def build_dashboard_page_images(df_: pd.DataFrame, color_map_: dict):
    """Create the donut + bar charts as images for the PDF."""
    to_sum, agg, total_points, per_credit = _compute_dashboard_data(df_)

    donut_png = None
    bars_png = None

    if not agg.empty:
        agg_sorted = agg.copy().sort_values("Category", kind="stable")
        categories_sorted_main = sorted(agg_sorted["Category"].unique().tolist())
        fig = px.pie(
            agg_sorted,
            names='Category',
            values='Points',
            hole=0.5,
            color="Category",
            color_discrete_map=color_map_,
            height=800,
            category_orders={"Category": categories_sorted_main}
        )
        fig.update_traces(textposition='inside', textinfo='percent+value')
        fig.update_layout(annotations=[dict(text=f"{total_points:.0f}", x=0.5, y=0.5, font_size=80, showarrow=False)])
        donut_png = _figure_bytes(fig, width=1100, height=800, scale=2)

    if not per_credit.empty:
        y_order = per_credit['Credit'].tolist()
        bar_height = max(400, min(1000, 40 * len(per_credit) + 120))
        fig_bar = px.bar(
            per_credit,
            x='Points', y='Credit', color='Category', orientation='h',
            color_discrete_map=color_map_, text='LabelText'
        )
        fig_bar.update_traces(textposition='outside', cliponaxis=False)
        fig_bar.update_layout(
            height=bar_height, xaxis_title="Planned Points", yaxis_title="Credit",
            yaxis=dict(categoryorder='array', categoryarray=y_order),
            margin=dict(l=10, r=30, t=30, b=10), legend_title="Category", legend_traceorder="reversed"
        )
        bars_png = _figure_bytes(fig_bar, width=1100, height=bar_height, scale=2)

    return donut_png, bars_png

def build_full_report_pdf(df: pd.DataFrame, project_name: str, rating_system: str, color_map_: dict) -> bytes:
    """Build a PDF: page 1 Dashboard (charts), then pursued-items detailed report."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=18*mm, rightMargin=18*mm, topMargin=18*mm, bottomMargin=18*mm)
    styles = getSampleStyleSheet()
    h1 = ParagraphStyle('h1', parent=styles['Heading1'], fontSize=18, spaceAfter=8, leading=22)
    h2 = ParagraphStyle('h2', parent=styles['Heading2'], fontSize=14, spaceAfter=6, leading=18, textColor=colors.HexColor('#333333'))
    body = ParagraphStyle('body', parent=styles['BodyText'], fontSize=10.5, leading=14, spaceAfter=6)
    small = ParagraphStyle('small', parent=styles['BodyText'], fontSize=9.5, leading=13, spaceAfter=4)

    story = []

    # ===== PAGE 1: Dashboard =====
    story.append(Paragraph(project_name, h1))
    story.append(Paragraph(f"LEED v5 {rating_system}", body))
    story.append(Spacer(1, 6*mm))

    donut_png, bars_png = build_dashboard_page_images(df, color_map_)
    page_width, page_height = A4
    max_w = page_width - (doc.leftMargin + doc.rightMargin)

    # Place donut (top)
    if donut_png:
        img1 = RLImage(io.BytesIO(donut_png))
        # 50% scale feel: use ~max_w (it will be scaled from original)
        img1._restrictSize(max_w, 110*mm)
        story.append(img1)
        story.append(Spacer(1, 4*mm))

    # Place bars (below)
    if bars_png:
        img2 = RLImage(io.BytesIO(bars_png))
        img2._restrictSize(max_w, 150*mm)
        story.append(img2)

    story.append(PageBreak())

    # ===== PAGES 2+: Pursued content =====
    pursued = df[df['Pursued'] == True].copy()
    if pursued.empty:
        story.append(Paragraph("No credits/options/paths are currently marked as <b>Pursued</b>.", body))
        doc.build(story)
        return buffer.getvalue()

    pursued = pursued.sort_values(['Category','Credit_ID','Option_Title','Path_Title'], ascending=[True, False, True, True])

    for category, cat_df in pursued.groupby('Category'):
        story.append(Paragraph(f"Category: {category}", h2))
        for (credit_id, credit_name), cred_df in cat_df.groupby(['Credit_ID','Credit_Name']):
            credit_label = f"{credit_id} {credit_name}"
            story.append(Paragraph(f"<b>{credit_label}</b>", body))

            planned_sum = int(pd.to_numeric(cred_df['Planned_Points'], errors='coerce').fillna(0).sum())
            tbl = Table([[f"Planned Points in this Credit: {planned_sum}"]], colWidths=[(page_width - doc.leftMargin - doc.rightMargin)])
            tbl.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f6f6f9')),
                ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#333333')),
                ('INNERGRID', (0,0), (-1,-1), 0.25, colors.HexColor('#dddddd')),
                ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
                ('LEFTPADDING', (0,0), (-1,-1), 6),
                ('RIGHTPADDING', (0,0), (-1,-1), 6),
                ('TOPPADDING', (0,0), (-1,-1), 4),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ]))
            story.append(tbl)
            story.append(Spacer(1, 2*mm))

            for (option_title, path_title), row_df in cred_df.groupby(['Option_Title','Path_Title']):
                story.append(Paragraph(f"<b>Option:</b> {option_title}", small))
                story.append(Paragraph(f"<b>Path:</b> {path_title}", small))

                # pull fields
                def _first(series):
                    return series.dropna().astype(str).drop_duplicates().iloc[0] if not series.dropna().empty else None

                req = _first(row_df['Requirement'])
                thr = _first(row_df['Thresholds']) if 'Thresholds' in row_df.columns else None
                and_ = _first(row_df['And']) if 'And' in row_df.columns else None
                or_  = _first(row_df['Or']) if 'Or' in row_df.columns else None
                docu = _first(row_df['Documentation']) if 'Documentation' in row_df.columns else None
                refs = None
                if 'Referenced_Standards' in row_df.columns:
                    refs = _first(row_df['Referenced_Standards'])
                elif 'Referenced Standards' in row_df.columns:
                    refs = _first(row_df['Referenced Standards'])
                max_pts_eff = row_df['Max_Points_Effective'].dropna()
                planned_pts = int(pd.to_numeric(row_df['Planned_Points'], errors='coerce').fillna(0).max())
                status = _first(row_df['Status'])
                responsible = _first(row_df['Responsible'])
                effort = _first(row_df['Effort'])
                approach = _first(row_df['Approach'])

                pts_line = f"<b>Planned Points:</b> {planned_pts}"
                if not max_pts_eff.empty:
                    try:
                        mp = int(float(max_pts_eff.max()))
                        pts_line += f" (Max: {mp})"
                    except Exception:
                        pass
                story.append(Paragraph(pts_line, small))

                if status:      story.append(Paragraph(f"<b>Status:</b> {status}", small))
                if effort:      story.append(Paragraph(f"<b>Effort:</b> {effort}", small))
                if responsible: story.append(Paragraph(f"<b>Responsible:</b> {responsible}", small))

                if approach:
                    story.append(Paragraph("<b>Approach</b>", small))
                    story.append(Paragraph(_htmlize(approach), small))

                if req:
                    story.append(Paragraph("<b>Requirement</b>", small))
                    story.append(Paragraph(_htmlize(req), small))

                if thr:
                    story.append(Paragraph("<b>Thresholds</b>", small))
                    story.append(Paragraph(_htmlize(thr), small))

                if and_:
                    story.append(Paragraph("<b>And</b>", small))
                    story.append(Paragraph(_htmlize(and_), small))
                if or_:
                    story.append(Paragraph("<b>Or</b>", small))
                    story.append(Paragraph(_htmlize(or_), small))

                if docu:
                    story.append(Paragraph("<b>Documentation</b>", small))
                    story.append(Paragraph(_htmlize(docu), small))
                if refs:
                    story.append(Paragraph("<b>Referenced Standards</b>", small))
                    story.append(Paragraph(_htmlize(refs), small))

                story.append(Spacer(1, 4*mm))
            story.append(Spacer(1, 6*mm))

        story.append(PageBreak())

    doc.build(story)
    return buffer.getvalue()


def build_filtered_catalog_report_pdf(
    df_filtered: pd.DataFrame,
    project_name: str,
    rating_system: str,
    color_map_: dict,
    filter_state: dict | None = None,
) -> bytes:
    """Build a stakeholder PDF based on the Credit Library filters (incl. Responsible filter)."""
    from datetime import datetime

    filter_state = filter_state or {}
    show_only_pursued = bool(filter_state.get("show_only_pursued", False))
    resp_filter = filter_state.get("resp_filter", []) or []
    resp_filter_mode = filter_state.get("resp_filter_mode", "Any")
    phase_filter = filter_state.get("phase_filter", []) or []


    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=18*mm, rightMargin=18*mm, topMargin=18*mm, bottomMargin=18*mm)
    styles = getSampleStyleSheet()
    h1 = ParagraphStyle('h1', parent=styles['Heading1'], fontSize=18, spaceAfter=8, leading=22)
    h2 = ParagraphStyle('h2', parent=styles['Heading2'], fontSize=14, spaceAfter=6, leading=18, textColor=colors.HexColor('#333333'))
    body = ParagraphStyle('body', parent=styles['BodyText'], fontSize=10.5, leading=14, spaceAfter=6)
    small = ParagraphStyle('small', parent=styles['BodyText'], fontSize=9.5, leading=13, spaceAfter=4)

    story = []
    story.append(Paragraph(project_name, h1))
    story.append(Paragraph(f"LEED v5 {rating_system}", body))
    story.append(Paragraph("Stakeholder Requirements Report (filtered)", body))
    story.append(Paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M')}", small))
    story.append(Spacer(1, 3*mm))

    # Filter summary
    filt_lines = []
    filt_lines.append(f"<b>Show only pursued:</b> {'Yes' if show_only_pursued else 'No'}")
    if resp_filter:
        filt_lines.append(f"<b>Responsible filter:</b> {', '.join(resp_filter)}")
        filt_lines.append(f"<b>Responsible match:</b> {resp_filter_mode}")
    else:
        filt_lines.append("<b>Responsible filter:</b> (none)")

    if phase_filter:
        filt_lines.append(f"<b>Implementation Phase filter:</b> {', '.join(phase_filter)}")
    else:
        filt_lines.append("<b>Implementation Phase filter:</b> (none)")

    story.append(Paragraph("<br/>".join(filt_lines), small))
    story.append(Spacer(1, 6*mm))

    # Page 1 charts (same style as Dashboard, but for the filtered subset)
    try:
        donut_png, bars_png = build_dashboard_page_images(df_filtered, color_map_)
    except Exception:
        donut_png, bars_png = (None, None)

    page_width, page_height = A4
    max_w = page_width - (doc.leftMargin + doc.rightMargin)

    if donut_png:
        img1 = RLImage(io.BytesIO(donut_png))
        img1._restrictSize(max_w, 110*mm)
        story.append(img1)
        story.append(Spacer(1, 4*mm))
    if bars_png:
        img2 = RLImage(io.BytesIO(bars_png))
        img2._restrictSize(max_w, 150*mm)
        story.append(img2)

    story.append(PageBreak())

    # Detail pages: all filtered items (not only pursued — unless the filter already restricted it)
    if df_filtered is None or df_filtered.empty:
        story.append(Paragraph("No credits/options/paths match the current Credit Library filters.", body))
        doc.build(story)
        return buffer.getvalue()

    df_rep = df_filtered.copy()
    # Stable sort
    for col in ['Category', 'Credit_ID', 'Credit_Name', 'Option_Title', 'Path_Title']:
        if col not in df_rep.columns:
            df_rep[col] = np.nan
    df_rep = df_rep.sort_values(['Category','Credit_ID','Option_Title','Path_Title'], ascending=[True, False, True, True])

    def _first(series):
        return series.dropna().astype(str).drop_duplicates().iloc[0] if series is not None and not series.dropna().empty else None

    for category, cat_df in df_rep.groupby('Category', dropna=False):
        cat_label = str(category) if pd.notna(category) else "(No Category)"
        story.append(Paragraph(f"Category: {cat_label}", h2))

        for (credit_id, credit_name), cred_df in cat_df.groupby(['Credit_ID','Credit_Name'], dropna=False):
            credit_label = f"{credit_id} {credit_name}".strip()
            story.append(Paragraph(f"<b>{_htmlize(credit_label)}</b>", body))

            # Credit-level summary (planned points only from pursued items)
            cred_pursued = cred_df[cred_df.get('Pursued', False) == True].copy()
            planned_sum = int(pd.to_numeric(cred_pursued.get('Planned_Points', 0), errors='coerce').fillna(0).sum()) if not cred_pursued.empty else 0
            count_items = int(len(cred_df.groupby(['Option_Title','Path_Title'], dropna=False)))
            summary_line = f"Entries in this credit (filtered): {count_items}  |  Planned points (pursued only): {planned_sum}"
            tbl = Table([[summary_line]], colWidths=[(page_width - doc.leftMargin - doc.rightMargin)])
            tbl.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f6f6f9')),
                ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#333333')),
                ('INNERGRID', (0,0), (-1,-1), 0.25, colors.HexColor('#dddddd')),
                ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#cccccc')),
                ('LEFTPADDING', (0,0), (-1,-1), 6),
                ('RIGHTPADDING', (0,0), (-1,-1), 6),
                ('TOPPADDING', (0,0), (-1,-1), 4),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ]))
            story.append(tbl)
            story.append(Spacer(1, 2*mm))

            for (option_title, path_title), row_df in cred_df.groupby(['Option_Title','Path_Title'], dropna=False):
                story.append(Paragraph(f"<b>Option:</b> {_htmlize(str(option_title))}", small))
                story.append(Paragraph(f"<b>Path:</b> {_htmlize(str(path_title))}", small))

                # Implementation phase
                impl_phase = _first(row_df.get('Implementation_Phase')) if 'Implementation_Phase' in row_df.columns else None
                if impl_phase:
                    story.append(Paragraph(f"<b>Implementation Phase:</b> {_htmlize(str(impl_phase))}", small))

                # Core fields
                req = _first(row_df.get('Requirement'))
                thr = _first(row_df.get('Thresholds')) if 'Thresholds' in row_df.columns else None
                and_ = _first(row_df.get('And')) if 'And' in row_df.columns else None
                or_  = _first(row_df.get('Or')) if 'Or' in row_df.columns else None
                docu = _first(row_df.get('Documentation')) if 'Documentation' in row_df.columns else None
                refs = None
                if 'Referenced_Standards' in row_df.columns:
                    refs = _first(row_df.get('Referenced_Standards'))
                elif 'Referenced Standards' in row_df.columns:
                    refs = _first(row_df.get('Referenced Standards'))

                max_pts_eff = row_df.get('Max_Points_Effective', pd.Series(dtype=float)).dropna()
                planned_pts = int(pd.to_numeric(row_df.get('Planned_Points', 0), errors='coerce').fillna(0).max())
                status = _first(row_df.get('Status'))
                responsible = _first(row_df.get('Responsible'))
                effort = _first(row_df.get('Effort'))
                approach = _first(row_df.get('Approach'))
                pursued_flag = bool(row_df.get('Pursued', pd.Series([False])).fillna(False).max())

                pts_line = f"<b>Pursued:</b> {'Yes' if pursued_flag else 'No'} &nbsp;&nbsp; <b>Planned Points:</b> {planned_pts}"
                if not max_pts_eff.empty:
                    try:
                        mp = int(float(max_pts_eff.max()))
                        pts_line += f" (Max: {mp})"
                    except Exception:
                        pass
                story.append(Paragraph(pts_line, small))

                if status:      story.append(Paragraph(f"<b>Status:</b> {_htmlize(status)}", small))
                if effort:      story.append(Paragraph(f"<b>Effort:</b> {_htmlize(effort)}", small))
                if responsible: story.append(Paragraph(f"<b>Responsible:</b> {_htmlize(responsible)}", small))

                # Approach
                if approach:
                    story.append(Paragraph("<b>Approach</b>", small))
                    story.append(Paragraph(_htmlize(approach), small))

                # Stakeholder comments
                if responsible:
                    assigned = [t for t in split_tokens(responsible) if t in RESPONSIBLE_OPTIONS]
                else:
                    assigned = []
                # Stakeholder comments for all assigned Responsible stakeholders
                for s in assigned:
                    col = comment_colname(s)
                    if col in row_df.columns:
                        comm = _first(row_df.get(col))
                        if comm:
                            story.append(Paragraph(f'<b>Comments "{_htmlize(s)}"</b>', small))
                            story.append(Paragraph(_htmlize(comm), small))

                # Requirement content
                if req:
                    story.append(Paragraph("<b>Requirement</b>", small))
                    story.append(Paragraph(_htmlize(req), small))
                if thr:
                    story.append(Paragraph("<b>Thresholds</b>", small))
                    story.append(Paragraph(_htmlize(thr), small))
                if and_:
                    story.append(Paragraph("<b>And</b>", small))
                    story.append(Paragraph(_htmlize(and_), small))
                if or_:
                    story.append(Paragraph("<b>Or</b>", small))
                    story.append(Paragraph(_htmlize(or_), small))
                if docu:
                    story.append(Paragraph("<b>Documentation</b>", small))
                    story.append(Paragraph(_htmlize(docu), small))
                if refs:
                    story.append(Paragraph("<b>Referenced Standards</b>", small))
                    story.append(Paragraph(_htmlize(refs), small))

                story.append(Spacer(1, 4*mm))

            story.append(Spacer(1, 6*mm))
        story.append(PageBreak())

    doc.build(story)
    return buffer.getvalue()

# ==== END REPORT HELPERS ======================================================

# ---------- Tabs ----------
tab_select, tab_dashboard = st.tabs(["Credits Library", "Project Dashboard"])

# ========== CREDITS Library ==========
with tab_select:
    # =========================
    # Credit Library (selectors in an expander)
    # =========================
    with st.expander("Credits Library", expanded=True):
        # NEW: checkbox to filter to only pursued
        show_only_pursued_catalog = st.checkbox(
            "Show only pursued credits/options/paths",
            value=False,
            help="Filter the dropdowns to items marked as Pursued."
        )

        # NEW: filter catalog by Responsible stakeholders
        c1, c2 = st.columns([5, 1])
        with c1:
            resp_filter = st.multiselect(
                "Filter catalog by Responsible",
                options=["Unassigned"] + RESPONSIBLE_OPTIONS,
                default=[],
                key="catalog_resp_filter",
                help="Show only credits/options/paths assigned to the selected stakeholders (based on the Responsible field)."
            )
        with c2:
            resp_filter_mode = st.radio(
                "Responsible match",
                options=["Any", "All"],
                index=0,
                horizontal=True,
                key="catalog_resp_filter_mode",
                help="Any: match if any selected stakeholder is assigned. All: match only if all selected stakeholders are assigned."
            )
        # NEW: filter catalog by Implementation Phase
        phase_filter = st.multiselect(
            "Filter catalog by Implementation Phase",
            options=["Unspecified"] + IMPLEMENTATION_PHASES,
            default=[],
            key="catalog_phase_filter",
            help="Show only credits/options/paths matching the selected Implementation Phase values."
        )


        # Base source df for the catalog depending on the filter
        src = df[df['Pursued'] == True] if show_only_pursued_catalog else df

        # Apply Responsible filter (if set)
        if resp_filter:
            include_unassigned = "Unassigned" in resp_filter
            selected_resp = [r for r in resp_filter if r != "Unassigned"]
            mode = "all" if resp_filter_mode == "All" else "any"

            col_resp = src["Responsible"] if "Responsible" in src.columns else pd.Series([""] * len(src), index=src.index)
            if selected_resp:
                mask_resp = col_resp.apply(lambda x: responsible_matches(x, selected_resp, mode))
            else:
                mask_resp = pd.Series([False] * len(src), index=src.index)

            if include_unassigned:
                col_norm = col_resp.fillna("").astype(str).str.strip()
                mask_unassigned = (col_norm == "") | (col_norm.str.lower().isin(["nan", "none", "null"]))
                mask_final = mask_unassigned if not selected_resp else (mask_unassigned | mask_resp)
            else:
                mask_final = mask_resp

            src = src.loc[mask_final]

            if src.empty:
                st.warning("No items match the current Responsible filter. Showing all items.")
                src = df[df['Pursued'] == True] if show_only_pursued_catalog else df


        # Apply Implementation Phase filter (if set)
        if phase_filter:
            include_unspecified = "Unspecified" in phase_filter
            selected_phases = [p for p in phase_filter if p != "Unspecified"]

            col_phase = src["Implementation_Phase"] if "Implementation_Phase" in src.columns else pd.Series([""] * len(src), index=src.index)
            col_phase_norm = col_phase.fillna("").astype(str).str.strip()

            mask_phase = pd.Series([False] * len(src), index=src.index)
            if selected_phases:
                mask_phase = col_phase_norm.isin(selected_phases)

            if include_unspecified:
                mask_unspecified = (col_phase_norm.eq("") | col_phase_norm.str.lower().isin(["nan", "none", "null"]))
                mask_phase = mask_phase | mask_unspecified

            src = src.loc[mask_phase]

            if src.empty:
                st.warning("No items match the current Implementation Phase filter. Showing all items.")
                src = df[df['Pursued'] == True] if show_only_pursued_catalog else df

        st.caption(f"Catalog entries shown: {len(src):,}")
        # Persist current catalog filters + filtered dataset for stakeholder PDF export
        st.session_state['catalog_filtered_df'] = src.copy()
        st.session_state['catalog_filter_state'] = {
            'show_only_pursued': show_only_pursued_catalog,
            'resp_filter': resp_filter,
            'resp_filter_mode': resp_filter_mode,
            'phase_filter': phase_filter,
        }

        # Categories
        categories = pd.unique(src['Category'].dropna())
        if len(categories) == 0 and show_only_pursued_catalog:
            st.info("No pursued items yet. Showing all categories.")
            src = df
            categories = pd.unique(src['Category'].dropna())

        cat_options = list(categories)
        prev_cat = st.session_state.get("cat", None)
        cat_index = cat_options.index(prev_cat) if prev_cat in cat_options else 0
        selected_category = st.selectbox("Select a LEED v5 Category", cat_options, index=cat_index, key="cat")
        # Credits (ID + Name for display)
        cred_src = src.loc[src['Category'].eq(selected_category), ['Credit_ID', 'Credit_Name']].dropna().drop_duplicates()
        if cred_src.empty and show_only_pursued_catalog:
            st.info("No pursued credits for this category. Showing all credits.")
            cred_src = df.loc[df['Category'].eq(selected_category), ['Credit_ID', 'Credit_Name']].dropna().drop_duplicates()

        cred_src['display'] = cred_src['Credit_ID'].astype(str) + " " + cred_src['Credit_Name']
        credit_options = cred_src['display'].tolist()
        prev_credit = st.session_state.get("credit_disp", None)
        credit_index = credit_options.index(prev_credit) if prev_credit in credit_options else 0
        selected_credit_display = st.selectbox(
            f"Select a {selected_category} credit",
            credit_options,
            index=credit_index,
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

        opt_options = opt_src.tolist() if hasattr(opt_src, "tolist") else list(opt_src)
        prev_opt = st.session_state.get("opt", None)
        opt_index = opt_options.index(prev_opt) if prev_opt in opt_options else 0
        selected_option = st.selectbox(f"Select a {selected_credit} option", opt_options, index=opt_index, key="opt")

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

        path_options = path_src.tolist() if hasattr(path_src, "tolist") else list(path_src)
        prev_path = st.session_state.get("path", None)
        path_index = path_options.index(prev_path) if prev_path in path_options else 0
        selected_path = st.selectbox(f"Select a {selected_option} path", path_options, index=path_index, key="path")

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
    st.markdown("---")
    st.write("## Project's Strategy:")
    st.write("### General Approach")
    with st.expander("General Approach", expanded=True):
        key_base = f"{selected_category}|{selected_credit_id}|{selected_credit}|{selected_option}|{selected_path}"
        approach_key = f"approach::{key_base}"
        status_key = f"status::{key_base}"
        resp_key = f"responsible::{key_base}"
        effort_key = f"effort::{key_base}"
        phase_key = f"phase::{key_base}"

        st.caption("#### Input Werner Sobek")
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

        # Responsible parsing handled by parse_responsible_valid() helper

        current_df_resp = RESPONSIBLE_DEFAULTS.copy()
        _ser_resp = df.loc[mask, 'Responsible'].dropna().astype(str)
        if not _ser_resp.empty:
            parsed = parse_responsible_valid(_ser_resp.iloc[0])
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



        # ---- Implementation Phase (dropdown) ----
        current_df_phase = DEFAULT_IMPLEMENTATION_PHASE
        _ser_phase = df.loc[mask, 'Implementation_Phase'].dropna().astype(str)
        if not _ser_phase.empty and _ser_phase.iloc[0] in IMPLEMENTATION_PHASES:
            current_df_phase = _ser_phase.iloc[0]

        if phase_key not in st.session_state:
            st.session_state[phase_key] = current_df_phase

        new_phase = st.selectbox(
            "Implementation Phase",
            IMPLEMENTATION_PHASES,
            index=IMPLEMENTATION_PHASES.index(st.session_state[phase_key]) if st.session_state[phase_key] in IMPLEMENTATION_PHASES else IMPLEMENTATION_PHASES.index(DEFAULT_IMPLEMENTATION_PHASE),
            key=phase_key,
            help="Planned phase of implementation across HOAI Leistungsphasen (LPH)."
        )

        if new_phase != current_df_phase:
            idx = st.session_state.df.loc[mask].index
            st.session_state.df.loc[idx, 'Implementation_Phase'] = new_phase
            st.caption("Implementation Phase autosaved ✓")

    # =========================
    # Stakeholder-specific comments (per Responsible)
    # =========================
    st.write("### Stakeholder Approach")
    with st.expander("Approach Stakeholder", expanded=True):
        key_base = f"{selected_category}|{selected_credit_id}|{selected_credit}|{selected_option}|{selected_path}"
        # Use the current Responsible selection (stored in session_state for the selected item)
        assigned_stakeholders = st.session_state.get(f"responsible::{key_base}", [])
        if not isinstance(assigned_stakeholders, list):
            assigned_stakeholders = []

        if not assigned_stakeholders:
            st.info("No Responsible stakeholders are assigned to this credit/option/path. Add them under 'Approach' to enable stakeholder comments.")
        else:
            st.caption("#### Input Stakeholders")
            for _stakeholder in assigned_stakeholders:
                _col = comment_colname(_stakeholder)
                if _col not in st.session_state.df.columns:
                    st.session_state.df[_col] = np.nan

                # Current comment from dataframe (first non-empty)
                _current = ""
                _ser = st.session_state.df.loc[mask, _col].dropna().astype(str)
                if not _ser.empty:
                    _current = _ser.iloc[0]

                _k = f"comment::{key_base}::{_col}"
                if _k not in st.session_state:
                    st.session_state[_k] = _current

                _new = st.text_area(
                    f'Comments "{_stakeholder}"',
                    key=_k,
                    height=140,
                    help=f"Input from {_stakeholder} regarding how to meet this credit/path during Design."
                )

                if _new != _current:
                    _idx = st.session_state.df.loc[mask].index
                    st.session_state.df.loc[_idx, _col] = _new
                    st.caption(f'Comment "{_stakeholder}" autosaved ✓')


    # =========================
    # Stakeholder PDF export (respects Credit Library filters)
    # =========================
    st.markdown("---")
    st.write("## Export")
    _df_filtered = st.session_state.get("catalog_filtered_df", df).copy()
    _filter_state = st.session_state.get("catalog_filter_state", {})
    st.write("Export here a report PDF file for the filtered credits")
    st.caption("Always make sure to regenerate report before export")
    if st.button("Generate Filtered Credits Report", type="primary",use_container_width=True):
        try:
            _pdf_bytes = build_filtered_catalog_report_pdf(
                df_filtered=_df_filtered,
                project_name=st.session_state.get("project_name", "Project"),
                rating_system=st.session_state.get("rating_system", ""),
                color_map_=color_map,
                filter_state=_filter_state
            )
            st.session_state["stakeholder_pdf_bytes"] = _pdf_bytes
            # Simple filename
            _resp = _filter_state.get("resp_filter", [])
            _tag = "All" if not _resp else "_".join([re.sub(r'[^0-9A-Za-z]+','', r) for r in _resp])[:60]
            st.session_state["stakeholder_pdf_name"] = f"LEED_v5_Stakeholder_Report_{_tag}.pdf"
            st.success("Stakeholder PDF generated.")
        except Exception as e:
            st.error(f"PDF generation failed: {e}")

    if st.session_state.get("stakeholder_pdf_bytes"):
        st.download_button(
            label="Download Filtered Credits Report",
            data=st.session_state["stakeholder_pdf_bytes"],
            file_name=st.session_state.get("stakeholder_pdf_name", "LEED_v5_Stakeholder_Report.pdf"),
            mime="application/pdf",
            use_container_width=True
        )

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

    # ==== EXPORT SECTION (PDF REPORT) ====
    st.markdown("---")
    st.write("## Export")
    pdf_bytes = build_full_report_pdf(
        st.session_state.df,
        project_name=st.session_state.get("project_name", "LEED v5 Project"),
        rating_system=st.session_state.get("rating_system", "BD+C: New Construction"),
        color_map_=color_map
    )
    st.download_button(
        label="Download Full Report (PDF)",
        data=pdf_bytes,
        file_name="LEED_v5_Full_Report.pdf",
        mime="application/pdf",
        use_container_width=True
    )

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
