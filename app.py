import traceback
import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import landscape, A4

st.set_page_config(layout="wide")
st.title("JJM Performance Monitoring Dashboard")

st.markdown("Upload BFM, Functionality and SO Details files")

# ================= SESSION STATE =================

if "reports_generated" not in st.session_state:
    st.session_state.reports_generated = False
if "so_group" not in st.session_state:
    st.session_state.so_group = None
if "sub_group" not in st.session_state:
    st.session_state.sub_group = None
if "div_group" not in st.session_state:
    st.session_state.div_group = None


# ================= FILE UPLOAD =================

bfm_file  = st.file_uploader("Upload BFM File",           type=["xlsx", "csv"])
func_file = st.file_uploader("Upload Functionality File", type=["xlsx", "csv"])
so_file   = st.file_uploader("Upload SO Details File",    type=["xlsx", "csv"])

generate = st.button("Generate Report")


# ================= FILE READ =================

def detect_header_row(raw_df):
    """
    Scan every row of a header=None DataFrame and return the index of the
    first row that contains 'imis' in any cell.
    Uses explicit str() per cell to guarantee no float-iteration errors.
    """
    for i in range(len(raw_df)):
        # tolist() first, then explicit str() — avoids astype edge-cases
        row = [str(v).lower() for v in raw_df.iloc[i].tolist()]
        if any("imis" in cell for cell in row):
            return i
    return None


def clean_columns(df):
    """
    Guarantee every column name is a plain stripped lowercase Python string.
    1. Explicit str() on every column name  →  float NaN → "nan"
    2. Strip BOM (\\ufeff) common in Excel-CSV exports
    3. strip + lower
    4. Drop columns named "nan" or "" (empty/unnamed columns)
    """
    new_cols = [
        str(c).replace('\ufeff', '').strip().lower()
        for c in df.columns
    ]
    df.columns = new_cols
    df = df.loc[:, [c for c in df.columns if c not in ("nan", "")]]
    return df


def read_excel_safe(file):
    file.seek(0)
    raw = pd.read_excel(file, header=None, dtype=str)   # dtype=str → no float cells
    header_row = detect_header_row(raw)
    if header_row is None:
        st.error(
            f"Could not find a header row containing 'imis' in **{file.name}**. "
            "Please check the file."
        )
        st.stop()
    file.seek(0)
    df = pd.read_excel(file, header=header_row)
    return clean_columns(df)


def read_csv_safe(file):
    # First pass: scan raw rows for header
    file.seek(0)
    raw = pd.read_csv(file, header=None, dtype=str)
    header_row = detect_header_row(raw)

    # Second pass: read with correct header, keep everything as str initially
    file.seek(0)
    if header_row is not None and header_row > 0:
        df = pd.read_csv(file, header=header_row, dtype=str)
    else:
        df = pd.read_csv(file, dtype=str)

    return clean_columns(df)


def read_file(file):
    if file.name.lower().endswith(".csv"):
        return read_csv_safe(file)
    return read_excel_safe(file)


# ================= HELPERS =================

def to_numeric_safe(series):
    """Convert a column to numeric, coercing errors, then fill NaN with 0."""
    return pd.to_numeric(series, errors='coerce').fillna(0)


def round_numeric_columns(df):
    numeric_cols = df.select_dtypes(include=['float', 'float64']).columns
    df[numeric_cols] = df[numeric_cols].round(2)
    return df


def assign_grade_by_rank(df):
    total   = len(df)
    top_cut = round(total * 0.33)
    mid_cut = round(total * 0.66)
    grades  = []
    for r in df["rank"]:
        if r <= top_cut:
            grades.append("Good")
        elif r <= mid_cut:
            grades.append("Needs Improvement")
        else:
            grades.append("Critical")
    df["grade"] = grades
    return df


def style_rank_column(df):
    def color_rank(row):
        if row["grade"] == "Good":
            return "background-color:#2ecc71; color:white"
        elif row["grade"] == "Needs Improvement":
            return "background-color:#f1c40f; color:black"
        else:
            return "background-color:#e74c3c; color:white"

    return df.style.apply(
        lambda row: [color_rank(row) if col == "rank" else "" for col in df.columns],
        axis=1
    )


# ================= PDF =================

def generate_pdf(title, df):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20
    )
    elements = []
    styles   = getSampleStyleSheet()
    elements.append(Paragraph(title, styles["Heading2"]))
    elements.append(Spacer(1, 0.3 * inch))

    data      = [df.columns.tolist()] + df.values.tolist()
    col_width = (landscape(A4)[0] - doc.leftMargin - doc.rightMargin) / len(df.columns)
    table     = Table(data, colWidths=[col_width] * len(df.columns), repeatRows=1)

    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('GRID',       (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTSIZE',   (0, 0), (-1, -1), 7),
    ])

    rank_idx = df.columns.get_loc("rank")
    total    = len(df)
    top_cut  = round(total * 0.33)
    mid_cut  = round(total * 0.66)

    for i in range(1, len(data)):
        rv = data[i][rank_idx]
        bg = colors.green if rv <= top_cut else (colors.yellow if rv <= mid_cut else colors.red)
        style.add('BACKGROUND', (rank_idx, i), (rank_idx, i), bg)

    table.setStyle(style)
    elements.append(table)
    elements.append(Spacer(1, 0.4 * inch))
    elements.append(Paragraph("Legend:",                                  styles["Heading3"]))
    elements.append(Paragraph("Green  → Top 33% (Good)",                 styles["Normal"]))
    elements.append(Paragraph("Yellow → Middle 33% (Needs Improvement)", styles["Normal"]))
    elements.append(Paragraph("Red    → Bottom 34% (Critical)",          styles["Normal"]))

    doc.build(elements)
    buffer.seek(0)
    return buffer


# ================= MAIN LOGIC =================

if generate:

    if not (bfm_file and func_file and so_file):
        st.error("Please upload all three files.")
        st.stop()

    try:
        # ---- Read files ----
        bfm  = read_file(bfm_file)
        func = read_file(func_file)
        so   = read_file(so_file)

        # ---- Debug expander: shows detected column names ----
        with st.expander("🔍 Debug: detected column names (expand to verify)"):
            st.write("**BFM columns:**",  list(bfm.columns))
            st.write("**Func columns:**", list(func.columns))
            st.write("**SO columns:**",   list(so.columns))

        # ---- Guard: required columns ----
        required_func = {'imis_id', 'work_status', 'functional_days'}
        required_so   = {'imis_id', 'so_name', 'sub_divisions', 'division'}
        required_bfm  = {'imis_id', 'no_of_days_bfm_reported'}

        missing_func = required_func - set(func.columns)
        missing_so   = required_so   - set(so.columns)
        missing_bfm  = required_bfm  - set(bfm.columns)

        if missing_func:
            st.error(f"Functionality file is missing columns: {missing_func}")
            st.stop()
        if missing_so:
            st.error(f"SO Details file is missing columns: {missing_so}")
            st.stop()
        if missing_bfm:
            st.error(f"BFM file is missing columns: {missing_bfm}")
            st.stop()

        # ---- Normalise key ID / text columns ----
        func['imis_id']     = func['imis_id'].astype(str).str.strip()
        so['imis_id']       = so['imis_id'].astype(str).str.strip()
        bfm['imis_id']      = bfm['imis_id'].astype(str).str.strip()
        func['work_status'] = func['work_status'].astype(str).str.strip().str.lower()

        # ---- Convert numeric columns BEFORE any arithmetic ----
        func['functional_days']         = to_numeric_safe(func['functional_days'])
        bfm['no_of_days_bfm_reported']  = to_numeric_safe(bfm['no_of_days_bfm_reported'])

        # ---- Filter handed-over ----
        func = func[func['work_status'] == "handed-over"].copy()

        if func.empty:
            st.warning(
                "No rows with work_status = 'handed-over' found. "
                "Check the Functionality file — the value must be exactly 'handed-over'."
            )
            st.stop()

        # ---- Drop division from func (will come from SO merge) ----
        if "division" in func.columns:
            func = func.drop(columns=["division"])

        # ---- Merge ----
        df = func.merge(
            so[['imis_id', 'so_name', 'sub_divisions', 'division']],
            on='imis_id', how='left'
        )
        df = df.merge(
            bfm[['imis_id', 'no_of_days_bfm_reported']],
            on='imis_id', how='left'
        )
        df['no_of_days_bfm_reported'] = to_numeric_safe(df['no_of_days_bfm_reported'])

        # ---- Compute scores ----
        max_days = max(
            df['functional_days'].max(),
            df['no_of_days_bfm_reported'].max()
        )
        if max_days == 0:
            st.error(
                "max_days is 0 — both functional_days and no_of_days_bfm_reported "
                "contain only zeros or blanks. Check your files."
            )
            st.stop()

        df['bfm_%']           = (df['no_of_days_bfm_reported'] / max_days) * 100
        df['functionality_%'] = (df['functional_days']          / max_days) * 100
        df['performance_%']   = 0.5 * df['bfm_%'] + 0.5 * df['functionality_%']

        # ================= SO =================
        so_group = df.groupby(['so_name', 'sub_divisions', 'division'], dropna=False).agg(
            schemes          =('imis_id',        'count'),
            bfm_pct          =('bfm_%',          'mean'),
            functionality_pct=('functionality_%', 'mean'),
            performance_pct  =('performance_%',   'mean')
        ).reset_index()
        so_group.rename(columns={
            'bfm_pct': 'bfm_%', 'functionality_pct': 'functionality_%',
            'performance_pct': 'performance_%'
        }, inplace=True)

        max_s = so_group['schemes'].max()
        so_group['workload_factor'] = 1 + (so_group['schemes'] / max_s)
        so_group['weighted_score']  = so_group['performance_%'] * so_group['workload_factor']
        so_group = so_group.sort_values("weighted_score", ascending=False).reset_index(drop=True)
        so_group['sl_no'] = range(1, len(so_group) + 1)
        so_group['rank']  = range(1, len(so_group) + 1)
        so_group = assign_grade_by_rank(so_group)
        so_group = so_group[['sl_no', 'so_name', 'sub_divisions', 'division', 'schemes',
                              'bfm_%', 'functionality_%', 'performance_%',
                              'workload_factor', 'weighted_score', 'rank', 'grade']]
        so_group = round_numeric_columns(so_group)

        # ================= SUB =================
        sub_group = df.groupby(['sub_divisions', 'division'], dropna=False).agg(
            schemes          =('imis_id',        'count'),
            bfm_pct          =('bfm_%',          'mean'),
            functionality_pct=('functionality_%', 'mean'),
            performance_pct  =('performance_%',   'mean')
        ).reset_index()
        sub_group.rename(columns={
            'bfm_pct': 'bfm_%', 'functionality_pct': 'functionality_%',
            'performance_pct': 'performance_%'
        }, inplace=True)

        max_s = sub_group['schemes'].max()
        sub_group['workload_factor'] = 1 + (sub_group['schemes'] / max_s)
        sub_group['weighted_score']  = sub_group['performance_%'] * sub_group['workload_factor']
        sub_group = sub_group.sort_values("weighted_score", ascending=False).reset_index(drop=True)
        sub_group['sl_no'] = range(1, len(sub_group) + 1)
        sub_group['rank']  = range(1, len(sub_group) + 1)
        sub_group = assign_grade_by_rank(sub_group)
        sub_group = sub_group[['sl_no', 'sub_divisions', 'division', 'schemes',
                                'bfm_%', 'functionality_%', 'performance_%',
                                'workload_factor', 'weighted_score', 'rank', 'grade']]
        sub_group = round_numeric_columns(sub_group)

        # ================= DIV =================
        div_group = df.groupby(['division'], dropna=False).agg(
            schemes          =('imis_id',        'count'),
            bfm_pct          =('bfm_%',          'mean'),
            functionality_pct=('functionality_%', 'mean'),
            performance_pct  =('performance_%',   'mean')
        ).reset_index()
        div_group.rename(columns={
            'bfm_pct': 'bfm_%', 'functionality_pct': 'functionality_%',
            'performance_pct': 'performance_%'
        }, inplace=True)

        max_s = div_group['schemes'].max()
        div_group['workload_factor'] = 1 + (div_group['schemes'] / max_s)
        div_group['weighted_score']  = div_group['performance_%'] * div_group['workload_factor']
        div_group = div_group.sort_values("weighted_score", ascending=False).reset_index(drop=True)
        div_group['sl_no'] = range(1, len(div_group) + 1)
        div_group['rank']  = range(1, len(div_group) + 1)
        div_group = assign_grade_by_rank(div_group)
        div_group = div_group[['sl_no', 'division', 'schemes',
                                'bfm_%', 'functionality_%', 'performance_%',
                                'workload_factor', 'weighted_score', 'rank', 'grade']]
        div_group = round_numeric_columns(div_group)

        st.session_state.so_group  = so_group
        st.session_state.sub_group = sub_group
        st.session_state.div_group = div_group
        st.session_state.reports_generated = True

    except Exception as e:
        st.error(f"Critical Error: {e}")
        # Full traceback shown so you can see the EXACT line that failed
        st.code(traceback.format_exc(), language="python")


# ================= DISPLAY =================

if st.session_state.reports_generated:

    st.header("SO Ranking")
    st.dataframe(style_rank_column(st.session_state.so_group))
    st.download_button("Download SO Ranking PDF",
                       generate_pdf("SO Ranking", st.session_state.so_group),
                       "SO_Ranking_Report.pdf", "application/pdf")

    st.header("Subdivision Ranking")
    st.dataframe(style_rank_column(st.session_state.sub_group))
    st.download_button("Download Subdivision Ranking PDF",
                       generate_pdf("Subdivision Ranking", st.session_state.sub_group),
                       "Subdivision_Ranking_Report.pdf", "application/pdf")

    st.header("Division Ranking")
    st.dataframe(style_rank_column(st.session_state.div_group))
    st.download_button("Download Division Ranking PDF",
                       generate_pdf("Division Ranking", st.session_state.div_group),
                       "Division_Ranking_Report.pdf", "application/pdf")
