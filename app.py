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

bfm_file = st.file_uploader("Upload BFM File", type=["xlsx", "csv"])
func_file = st.file_uploader("Upload Functionality File", type=["xlsx", "csv"])
so_file = st.file_uploader("Upload SO Details File", type=["xlsx", "csv"])

generate = st.button("Generate Report")


# ================= FILE READ =================

def detect_header(file):
    raw = pd.read_excel(file, header=None)
    for i in range(len(raw)):
        row = raw.iloc[i].astype(str).str.lower().tolist()
        if any("imis" in cell for cell in row):
            return i
    return None

def read_excel_safe(file):
    header_row = detect_header(file)
    if header_row is None:
        st.error(f"Header row not detected in {file.name}")
        st.stop()
    return pd.read_excel(file, header=header_row)

def read_file(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file)
    return read_excel_safe(file)


# ================= ROUND NUMERIC =================

def round_numeric_columns(df):
    numeric_cols = df.select_dtypes(include=['float', 'float64']).columns
    df[numeric_cols] = df[numeric_cols].round(2)
    return df


# ================= RANK BASED GRADING =================

def assign_grade_by_rank(df):
    total = len(df)
    top_cut = round(total * 0.33)
    mid_cut = round(total * 0.66)

    grades = []
    for r in df["rank"]:
        if r <= top_cut:
            grades.append("Good")
        elif r <= mid_cut:
            grades.append("Needs Improvement")
        else:
            grades.append("Critical")

    df["grade"] = grades
    return df


# ================= STYLE RANK ONLY =================

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


# ================= PDF AUTO FIT =================

def generate_pdf(title, df):

    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=20,
        leftMargin=20,
        topMargin=20,
        bottomMargin=20
    )

    elements = []
    styles = getSampleStyleSheet()

    elements.append(Paragraph(title, styles["Heading2"]))
    elements.append(Spacer(1, 0.3 * inch))

    data = [df.columns.tolist()] + df.values.tolist()

    page_width = landscape(A4)[0] - doc.leftMargin - doc.rightMargin
    col_count = len(df.columns)
    col_width = page_width / col_count

    table = Table(data, colWidths=[col_width]*col_count, repeatRows=1)

    style = TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.grey),
        ('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('FONTSIZE',(0,0),(-1,-1),7)
    ])

    rank_index = df.columns.get_loc("rank")

    total = len(df)
    top_cut = round(total * 0.33)
    mid_cut = round(total * 0.66)

    for i in range(1, len(data)):
        rank_val = data[i][rank_index]

        if rank_val <= top_cut:
            bg = colors.green
        elif rank_val <= mid_cut:
            bg = colors.yellow
        else:
            bg = colors.red

        style.add('BACKGROUND', (rank_index, i), (rank_index, i), bg)

    table.setStyle(style)
    elements.append(table)

    elements.append(Spacer(1, 0.4 * inch))
    elements.append(Paragraph("Legend:", styles["Heading3"]))
    elements.append(Paragraph("Green → Top 33% (Good)", styles["Normal"]))
    elements.append(Paragraph("Yellow → Middle 33% (Needs Improvement)", styles["Normal"]))
    elements.append(Paragraph("Red → Bottom 34% (Critical)", styles["Normal"]))

    doc.build(elements)
    buffer.seek(0)
    return buffer


# ================= MAIN LOGIC =================

if generate:

    if not (bfm_file and func_file and so_file):
        st.error("Please upload all three files.")
        st.stop()

    try:
        bfm = read_file(bfm_file)
        func = read_file(func_file)
        so = read_file(so_file)

        bfm.columns = bfm.columns.str.strip().str.lower()
        func.columns = func.columns.str.strip().str.lower()
        so.columns = so.columns.str.strip().str.lower()

        func = func[func['work_status'].str.lower() == "handed-over"]

        if "division" in func.columns:
            func = func.drop(columns=["division"])

        df = func.merge(
            so[['imis_id', 'so_name', 'sub_divisions', 'division']],
            on='imis_id',
            how='left'
        )

        df = df.merge(
            bfm[['imis_id', 'no_of_days_bfm_reported']],
            on='imis_id',
            how='left'
        )

        df['no_of_days_bfm_reported'] = df['no_of_days_bfm_reported'].fillna(0)

        max_days = max(
            df['functional_days'].max(),
            df['no_of_days_bfm_reported'].max()
        )

        df['bfm_%'] = (df['no_of_days_bfm_reported'] / max_days) * 100
        df['functionality_%'] = (df['functional_days'] / max_days) * 100
        df['performance_%'] = 0.5 * df['bfm_%'] + 0.5 * df['functionality_%']

        # ================= SO =================
        so_group = df.groupby(['so_name', 'sub_divisions', 'division']).agg({
            'imis_id': 'count',
            'bfm_%': 'mean',
            'functionality_%': 'mean',
            'performance_%': 'mean'
        }).reset_index()

        so_group.rename(columns={'imis_id': 'schemes'}, inplace=True)

        max_schemes = so_group['schemes'].max()
        so_group['workload_factor'] = 1 + (so_group['schemes'] / max_schemes)
        so_group['weighted_score'] = so_group['performance_%'] * so_group['workload_factor']

        so_group = so_group.sort_values("weighted_score", ascending=False)
        so_group['sl_no'] = range(1, len(so_group) + 1)
        so_group['rank'] = range(1, len(so_group) + 1)
        so_group = assign_grade_by_rank(so_group)
        
        # Reorder columns: sl_no first
        cols = ['sl_no', 'so_name', 'sub_divisions', 'division', 'schemes', 'bfm_%', 
                'functionality_%', 'performance_%', 'workload_factor', 'weighted_score', 'rank', 'grade']
        so_group = so_group[cols]
        
        so_group = round_numeric_columns(so_group)

        # ================= SUB =================
        sub_group = df.groupby(['sub_divisions', 'division']).agg({
            'imis_id': 'count',
            'bfm_%': 'mean',
            'functionality_%': 'mean',
            'performance_%': 'mean'
        }).reset_index()

        sub_group.rename(columns={'imis_id': 'schemes'}, inplace=True)

        max_schemes = sub_group['schemes'].max()
        sub_group['workload_factor'] = 1 + (sub_group['schemes'] / max_schemes)
        sub_group['weighted_score'] = sub_group['performance_%'] * sub_group['workload_factor']

        sub_group = sub_group.sort_values("weighted_score", ascending=False)
        sub_group['sl_no'] = range(1, len(sub_group) + 1)
        sub_group['rank'] = range(1, len(sub_group) + 1)
        sub_group = assign_grade_by_rank(sub_group)
        
        # Reorder columns: sl_no first
        cols = ['sl_no', 'sub_divisions', 'division', 'schemes', 'bfm_%', 
                'functionality_%', 'performance_%', 'workload_factor', 'weighted_score', 'rank', 'grade']
        sub_group = sub_group[cols]
        
        sub_group = round_numeric_columns(sub_group)

        # ================= DIV =================
        div_group = df.groupby(['division']).agg({
            'imis_id': 'count',
            'bfm_%': 'mean',
            'functionality_%': 'mean',
            'performance_%': 'mean'
        }).reset_index()

        div_group.rename(columns={'imis_id': 'schemes'}, inplace=True)

        max_schemes = div_group['schemes'].max()
        div_group['workload_factor'] = 1 + (div_group['schemes'] / max_schemes)
        div_group['weighted_score'] = div_group['performance_%'] * div_group['workload_factor']

        div_group = div_group.sort_values("weighted_score", ascending=False)
        div_group['sl_no'] = range(1, len(div_group) + 1)
        div_group['rank'] = range(1, len(div_group) + 1)
        div_group = assign_grade_by_rank(div_group)
        
        # Reorder columns: sl_no first
        cols = ['sl_no', 'division', 'schemes', 'bfm_%', 
                'functionality_%', 'performance_%', 'workload_factor', 'weighted_score', 'rank', 'grade']
        div_group = div_group[cols]
        
        div_group = round_numeric_columns(div_group)

        st.session_state.so_group = so_group
        st.session_state.sub_group = sub_group
        st.session_state.div_group = div_group
        st.session_state.reports_generated = True

    except Exception as e:
        st.error(f"Critical Error: {e}")


# ================= DISPLAY PERSISTENT =================

if st.session_state.reports_generated:

    st.header("SO Ranking")
    st.dataframe(style_rank_column(st.session_state.so_group))
    st.download_button("Download SO Ranking PDF",
                       generate_pdf("SO Ranking", st.session_state.so_group),
                       "SO_Ranking_Report.pdf",
                       "application/pdf")

    st.header("Subdivision Ranking")
    st.dataframe(style_rank_column(st.session_state.sub_group))
    st.download_button("Download Subdivision Ranking PDF",
                       generate_pdf("Subdivision Ranking", st.session_state.sub_group),
                       "Subdivision_Ranking_Report.pdf",
                       "application/pdf")

    st.header("Division Ranking")
    st.dataframe(style_rank_column(st.session_state.div_group))
    st.download_button("Download Division Ranking PDF",
                       generate_pdf("Division Ranking", st.session_state.div_group),
                       "Division_Ranking_Report.pdf",
                       "application/pdf")
