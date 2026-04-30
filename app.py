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


# ================= FILE HELPERS =================

def normalize_id(series):
    """
    Normalise IMIS IDs so that "12345", "12345.0", 12345, 12345.0 all become "12345".
    This prevents merge failures when one file stores IDs as float and another as int.
    """
    def _fix(val):
        s = str(val).strip()
        try:
            return str(int(float(s)))
        except (ValueError, OverflowError):
            return s
    return series.map(_fix)


def detect_header_row(raw_df):
    """Return index of first row whose cells contain the substring 'imis'."""
    for i in range(len(raw_df)):
        row = [str(v).lower() for v in raw_df.iloc[i].tolist()]
        if any("imis" in cell for cell in row):
            return i
    return None


def clean_columns(df):
    """
    Make every column name a plain stripped lowercase Python string.
    Handles: float NaN names, BOM characters, leading/trailing whitespace.
    Drops columns with empty or 'nan' names.
    """
    new_cols = [
        str(c).replace('\ufeff', '').strip().lower()
        for c in df.columns
    ]
    df.columns = new_cols
    df = df.loc[:, [c for c in df.columns if c not in ("nan", "")]]
    return df


def read_excel_safe(file):
    # First pass with dtype=str just to find the header row safely
    file.seek(0)
    raw = pd.read_excel(file, header=None, dtype=str)
    header_row = detect_header_row(raw)
    if header_row is None:
        st.error(f"Could not find a header row containing 'imis' in **{file.name}**.")
        st.stop()
    # Second pass: read normally (let pandas infer numeric types)
    file.seek(0)
    df = pd.read_excel(file, header=header_row)
    return clean_columns(df)


def read_csv_safe(file):
    # First pass with dtype=str to find the header row safely
    file.seek(0)
    raw = pd.read_csv(file, header=None, dtype=str)
    header_row = detect_header_row(raw)

    # Second pass: read normally (let pandas infer numeric types)
    file.seek(0)
    if header_row is not None and header_row > 0:
        df = pd.read_csv(file, header=header_row)
    else:
        df = pd.read_csv(file)
    return clean_columns(df)


def read_file(file):
    if file.name.lower().endswith(".csv"):
        return read_csv_safe(file)
    return read_excel_safe(file)


def to_numeric_safe(series):
    return pd.to_numeric(series, errors='coerce').fillna(0)


def round_numeric_columns(df):
    numeric_cols = df.select_dtypes(include=['float', 'float64']).columns
    df[numeric_cols] = df[numeric_cols].round(2)
    return df


# ================= RANK GRADING =================

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


# ================= STYLING =================

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
        ('BACKGROUND', (0, 0), (-1,  0), colors.grey),
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
        # ---- Read files (dtype inferred — no forced str this time) ----
        bfm  = read_file(bfm_file)
        func = read_file(func_file)
        so   = read_file(so_file)

        # ---- Debug expander ----
        with st.expander("🔍 Debug: detected column names"):
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
            st.error(f"Functionality file missing columns: {missing_func}")
            st.stop()
        if missing_so:
            st.error(f"SO Details file missing columns: {missing_so}")
            st.stop()
        if missing_bfm:
            st.error(f"BFM file missing columns: {missing_bfm}")
            st.stop()

        # ---- Normalise IMIS IDs: "12345.0" → "12345" across all three files ----
        func['imis_id'] = normalize_id(func['imis_id'])
        so['imis_id']   = normalize_id(so['imis_id'])
        bfm['imis_id']  = normalize_id(bfm['imis_id'])

        # ---- Normalise work_status text ----
        func['work_status'] = func['work_status'].astype(str).str.strip().str.lower()

        # ---- Convert numeric columns ----
        func['functional_days']        = to_numeric_safe(func['functional_days'])
        bfm['no_of_days_bfm_reported'] = to_numeric_safe(bfm['no_of_days_bfm_reported'])

        # ---- Filter handed-over ----
        func = func[func['work_status'] == "handed-over"].copy()

        if func.empty:
            st.warning(
                "No rows with work_status = 'handed-over' found. "
                "Check the Functionality file — the value must be exactly 'handed-over'."
            )
            st.stop()

        # ---- Drop division from func (comes from SO) ----
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

        # ---- Warn if merge produced no matches ----
        unmatched = df['so_name'].isna().sum()
        if unmatched > 0:
            st.warning(
                f"⚠️ {unmatched} rows in the Functionality file had no matching imis_id "
                f"in the SO file. Those rows will be excluded from SO/Sub/Division ranking. "
                f"Sample unmatched IDs: {df.loc[df['so_name'].isna(), 'imis_id'].head(5).tolist()}"
            )

        # ---- Drop rows where SO info is missing (unmatched merges) ----
        df = df.dropna(subset=['so_name', 'sub_divisions', 'division'])

        if df.empty:
            st.error(
                "After merging, no rows have SO/Subdivision/Division info. "
                "Check that imis_id values match between your files."
            )
            st.stop()

        # ---- Compute scores ----
        max_days = max(
            df['functional_days'].max(),
            df['no_of_days_bfm_reported'].max()
        )
        if max_days == 0:
            st.error("max_days is 0 — functional_days and no_of_days_bfm_reported have no valid values.")
            st.stop()

        df['bfm_%']           = (df['no_of_days_bfm_reported'] / max_days) * 100
        df['functionality_%'] = (df['functional_days']          / max_days) * 100
        df['performance_%']   = 0.5 * df['bfm_%'] + 0.5 * df['functionality_%']

        # ================= SO =================
        so_group = df.groupby(['so_name', 'sub_divisions', 'division']).agg(
            schemes          =('imis_id',        'count'),
            bfm_pct          =('bfm_%',          'mean'),
            functionality_pct=('functionality_%', 'mean'),
            performance_pct  =('performance_%',   'mean')
        ).reset_index()
        so_group.rename(columns={
            'bfm_pct': 'bfm_%',
            'functionality_pct': 'functionality_%',
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
        sub_group = df.groupby(['sub_divisions', 'division']).agg(
            schemes          =('imis_id',        'count'),
            bfm_pct          =('bfm_%',          'mean'),
            functionality_pct=('functionality_%', 'mean'),
            performance_pct  =('performance_%',   'mean')
        ).reset_index()
        sub_group.rename(columns={
            'bfm_pct': 'bfm_%',
            'functionality_pct': 'functionality_%',
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
        div_group = df.groupby(['division']).agg(
            schemes          =('imis_id',        'count'),
            bfm_pct          =('bfm_%',          'mean'),
            functionality_pct=('functionality_%', 'mean'),
            performance_pct  =('performance_%',   'mean')
        ).reset_index()
        div_group.rename(columns={
            'bfm_pct': 'bfm_%',
            'functionality_pct': 'functionality_%',
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
