import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

st.set_page_config(layout="wide")
st.title("JJM Performance Monitoring Dashboard")

st.markdown("Upload BFM, Functionality and SO Details files")

# ================= FILE UPLOAD =================
bfm_file = st.file_uploader("Upload BFM File", type=["xlsx", "csv"])
func_file = st.file_uploader("Upload Functionality File", type=["xlsx", "csv"])
so_file = st.file_uploader("Upload SO Details File", type=["xlsx", "csv"])

generate = st.button("Generate Report")


# ================= HEADER DETECTION =================

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


def calculate_grade(score):
    if score >= 90:
        return "Excellent"
    elif score >= 80:
        return "Good"
    elif score >= 70:
        return "Needs Improvement"
    else:
        return "Critical"


def apply_traffic_colors(df, column):
    top = df[column].quantile(0.67)
    bottom = df[column].quantile(0.33)

    def color(val):
        if val >= top:
            return "background-color: #28a745; color: white"
        elif val <= bottom:
            return "background-color: #dc3545; color: white"
        else:
            return "background-color: #ffc107; color: black"

    return df.style.applymap(color, subset=[column])


# ================= MAIN =================

if generate:

    if not (bfm_file and func_file and so_file):
        st.error("Please upload all three files.")
        st.stop()

    try:
        # Read files safely
        bfm = read_file(bfm_file)
        func = read_file(func_file)
        so = read_file(so_file)

        # Clean column names
        bfm.columns = bfm.columns.str.strip().str.lower()
        func.columns = func.columns.str.strip().str.lower()
        so.columns = so.columns.str.strip().str.lower()

        # Validate required columns
        required_func = ['imis_id', 'functional_days', 'work_status']
        required_so = ['imis_id', 'so_name', 'sub_divisions', 'division']
        required_bfm = ['imis_id', 'no_of_days_bfm_reported']

        for col in required_func:
            if col not in func.columns:
                st.error(f"Missing column in Functionality file: {col}")
                st.stop()

        for col in required_so:
            if col not in so.columns:
                st.error(f"Missing column in SO file: {col}")
                st.stop()

        for col in required_bfm:
            if col not in bfm.columns:
                st.error(f"Missing column in BFM file: {col}")
                st.stop()

        # Keep handed-over schemes only
        func = func[func['work_status'].str.lower() == "handed-over"]

        # Remove division from func to prevent duplication
        if "division" in func.columns:
            func = func.drop(columns=["division"])

        # Merge SO (source of division)
        df = func.merge(
            so[['imis_id', 'so_name', 'sub_divisions', 'division']],
            on='imis_id',
            how='left'
        )

        # Merge BFM
        df = df.merge(
            bfm[['imis_id', 'no_of_days_bfm_reported']],
            on='imis_id',
            how='left'
        )

        df['no_of_days_bfm_reported'] = df['no_of_days_bfm_reported'].fillna(0)

        # Hard validation of division column
        if 'division' not in df.columns:
            st.error("Division column missing after merge.")
            st.write("Available columns:", df.columns.tolist())
            st.stop()

        # Detect month days dynamically
        max_days = max(
            df['functional_days'].max(),
            df['no_of_days_bfm_reported'].max()
        )

        df['bfm_%'] = (df['no_of_days_bfm_reported'] / max_days) * 100
        df['functionality_%'] = (df['functional_days'] / max_days) * 100
        df['performance_%'] = 0.5 * df['bfm_%'] + 0.5 * df['functionality_%']

        # ================= SO LEVEL =================
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

        max_weighted = so_group['weighted_score'].max()
        so_group['final_performance_%'] = (so_group['weighted_score'] / max_weighted) * 100

        so_group['grade'] = so_group['final_performance_%'].apply(calculate_grade)
        so_group = so_group.sort_values("final_performance_%", ascending=False)
        so_group['rank'] = range(1, len(so_group) + 1)

        # ================= SUBDIVISION LEVEL =================
        sub_group = df.groupby(['sub_divisions', 'division']).agg({
            'so_name': 'nunique',
            'bfm_%': 'mean',
            'functionality_%': 'mean',
            'performance_%': 'mean'
        }).reset_index()

        sub_group.rename(columns={'so_name': 'so_count'}, inplace=True)

        max_so = sub_group['so_count'].max()
        sub_group['workload_factor'] = 1 + (sub_group['so_count'] / max_so)
        sub_group['weighted_score'] = sub_group['performance_%'] * sub_group['workload_factor']

        max_weighted_sub = sub_group['weighted_score'].max()
        sub_group['final_performance_%'] = (sub_group['weighted_score'] / max_weighted_sub) * 100

        sub_group['grade'] = sub_group['final_performance_%'].apply(calculate_grade)
        sub_group = sub_group.sort_values("final_performance_%", ascending=False)
        sub_group['rank'] = range(1, len(sub_group) + 1)

        # ================= DIVISION LEVEL =================
        div_group = df.groupby(['division']).agg({
            'sub_divisions': 'nunique',
            'bfm_%': 'mean',
            'functionality_%': 'mean',
            'performance_%': 'mean'
        }).reset_index()

        div_group.rename(columns={'sub_divisions': 'subdivision_count'}, inplace=True)

        max_sub = div_group['subdivision_count'].max()
        div_group['workload_factor'] = 1 + (div_group['subdivision_count'] / max_sub)
        div_group['weighted_score'] = div_group['performance_%'] * div_group['workload_factor']

        max_weighted_div = div_group['weighted_score'].max()
        div_group['final_performance_%'] = (div_group['weighted_score'] / max_weighted_div) * 100

        div_group['grade'] = div_group['final_performance_%'].apply(calculate_grade)
        div_group = div_group.sort_values("final_performance_%", ascending=False)
        div_group['rank'] = range(1, len(div_group) + 1)

        # ================= DISPLAY =================
        st.header("SO Ranking")
        st.dataframe(apply_traffic_colors(so_group, "final_performance_%"))

        st.header("Subdivision Ranking")
        st.dataframe(apply_traffic_colors(sub_group, "final_performance_%"))

        st.header("Division Ranking")
        st.dataframe(apply_traffic_colors(div_group, "final_performance_%"))

    except Exception as e:
        st.error(f"Critical Error: {e}")