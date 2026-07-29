import streamlit as st
import pandas as pd
from collections import Counter
import sqlite3
import datetime
import io

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Multi-Style Stock Variance Dashboard", layout="wide")
st.title("📊 Outlet Daily Multi-Style Stock Variance Dashboard")

# --- DATABASE SETUP (Free & Built-in SQLite) ---
DB_NAME = "variance_history.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS daily_variance (
            date TEXT,
            style_number TEXT,
            location TEXT,
            color_size_code TEXT,
            color_size_name TEXT,
            color TEXT,
            size TEXT,
            system_qty INTEGER,
            physical_qty INTEGER,
            variance INTEGER
        )
    ''')
    conn.commit()
    conn.close()

init_db()

# --- HELPER FUNCTIONS ---
def extract_style_number(code_or_name):
    """
    Extracts the first 8-digit style number (e.g. '25933700' from '25933700009S06')
    """
    str_val = str(code_or_name).strip()
    if len(str_val) >= 8 and str_val[:8].isdigit():
        return str_val[:8]
    return "UNKNOWN"

def parse_color_size(name):
    """
    Extracts Color and Size from 'Color Size Name'
    Example: '25933700 - WW PLEATING DRESS BEIGE-06' -> Color: BEIGE, Size: 06
    """
    name = str(name)
    if '-' in name:
        parts = name.rsplit('-', 1)
        size = parts[-1].strip()
        color_part = parts[0].split('-')[-1].strip() if len(parts[0].split('-')) > 1 else 'N/A'
        return color_part, size
    return 'N/A', 'N/A'

def process_scanned_data(file_or_text):
    """
    Reads stacked barcode strings from uploaded file or pasted text
    """
    if file_or_text is None:
        return []
    
    codes = []
    if hasattr(file_or_text, 'name'):
        filename = file_or_text.name.lower()
        if filename.endswith(('.xlsx', '.xls')):
            df_temp = pd.read_excel(file_or_text, header=None)
            codes = df_temp.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        elif filename.endswith('.csv'):
            df_temp = pd.read_csv(file_or_text, header=None)
            codes = df_temp.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        elif filename.endswith('.txt'):
            stringio = io.StringIO(file_or_text.getvalue().decode("utf-8"))
            codes = [line.strip() for line in stringio.readlines() if line.strip()]
    elif isinstance(file_or_text, str):
        codes = [line.strip() for line in file_or_text.split('\n') if line.strip()]
        
    return codes


# --- SIDEBAR INPUTS ---
st.sidebar.header("📅 Select Date")
selected_date = st.sidebar.date_input("Report Date", datetime.date.today())

st.sidebar.header("1. Upload Daily Style Files")
erp_files = st.sidebar.file_uploader(
    "Upload today's 5 Style ERP Excel Files", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True,
    key="erp_multi"
)

st.sidebar.header("2. Outlet Physical Scans")
outlets = ["CCC", "COLOMBO 03", "NUGEGODA", "ONLINE", "WATTALA"]
outlet_scans = {}

for outlet in outlets:
    with st.sidebar.expander(f"📍 {outlet} Scans"):
        tab1, tab2 = st.tabs(["📁 Upload File", "📋 Paste Codes"])
        with tab1:
            u_file = st.file_uploader(f"Upload file for {outlet}", type=["xlsx", "xls", "csv", "txt"], key=f"f_{outlet}")
        with tab2:
            p_text = st.text_area(f"Paste barcodes for {outlet}:", height=80, key=f"t_{outlet}")
        
        if u_file:
            c = process_scanned_data(u_file)
            outlet_scans[outlet] = c
            st.caption(f"✓ {len(c)} items loaded from file")
        elif p_text:
            c = process_scanned_data(p_text)
            outlet_scans[outlet] = c
            st.caption(f"✓ {len(c)} items loaded from text")


# --- MAIN ACTION: PROCESS AND STORE TODAY'S DATA ---
if erp_files:
    if st.sidebar.button("💾 Calculate & Save Daily Report"):
        # Combine all ERP Files uploaded
        all_erp_dfs = []
        for f in erp_files:
            df_temp = pd.read_excel(f)
            df_temp.columns = df_temp.columns.str.strip()
            if 'Unnamed: 0' in df_temp.columns:
                df_temp = df_temp.drop(columns=['Unnamed: 0'])
            all_erp_dfs.append(df_temp)
            
        master_erp = pd.concat(all_erp_dfs, ignore_index=True)
        master_erp.rename(columns={'Qty': 'System Qty'}, inplace=True)
        
        # Extract style, color, size
        master_erp['Style Number'] = master_erp['Color Size Code'].apply(extract_style_number)
        color_size_parsed = master_erp['Color Size Name'].apply(parse_color_size)
        master_erp['Color'] = [x[0] for x in color_size_parsed]
        master_erp['Size'] = [x[1] for x in color_size_parsed]

        # Aggregate Physical Scans
        scanned_records = []
        for loc, codes in outlet_scans.items():
            counts = Counter(codes)
            for code, count in counts.items():
                scanned_records.append({
                    'Location': loc,
                    'Color Size Code': code,
                    'Physical Qty': count
                })
        scanned_df = pd.DataFrame(scanned_records)

        # Merge ERP & Physical
        if not scanned_df.empty:
            merged = pd.merge(master_erp, scanned_df, on=['Location', 'Color Size Code'], how='outer')
        else:
            merged = master_erp.copy()
            merged['Physical Qty'] = 0

        merged['System Qty'] = merged['System Qty'].fillna(0)
        merged['Physical Qty'] = merged['Physical Qty'].fillna(0)
        merged['Variance'] = merged['Physical Qty'] - merged['System Qty']
        merged['Date'] = str(selected_date)

        # Save to SQLite Database (replace old records for same date if re-calculated)
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("DELETE FROM daily_variance WHERE date = ?", (str(selected_date),))
        
        insert_rows = []
        for _, r in merged.iterrows():
            insert_rows.append((
                str(selected_date),
                str(r.get('Style Number', extract_style_number(r['Color Size Code']))),
                str(r.get('Location', '')),
                str(r['Color Size Code']),
                str(r.get('Color Size Name', '')),
                str(r.get('Color', '')),
                str(r.get('Size', '')),
                int(r['System Qty']),
                int(r['Physical Qty']),
                int(r['Variance'])
            ))
        
        c.executemany('''
            INSERT INTO daily_variance VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', insert_rows)
        conn.commit()
        conn.close()
        
        st.success(f"✅ Data for {selected_date} successfully processed and saved to Database!")


# --- DASHBOARD & ANALYSIS SECTION ---
st.markdown("---")
st.header("🏢 Outlet-Wise Style Variance Analysis")

# Fetch dates saved in Database
conn = sqlite3.connect(DB_NAME)
available_dates = pd.read_sql_query("SELECT DISTINCT date FROM daily_variance ORDER BY date DESC", conn)

if not available_dates.empty:
    col_d, col_o = st.columns(2)
    view_date = col_d.selectbox("Select Date to View", available_dates['date'].tolist())
    view_outlet = col_o.selectbox("Select Outlet Name", outlets)

    # Query filtered data from DB
    query = f"""
        SELECT style_number as 'Style Number', 
               color_size_code as 'Color Size Code', 
               color_size_name as 'Description', 
               color as 'Color', 
               size as 'Size', 
               system_qty as 'System Qty', 
               physical_qty as 'Physical Qty', 
               variance as 'Variance'
        FROM daily_variance 
        WHERE date = '{view_date}' AND location = '{view_outlet}'
    """
    outlet_df = pd.read_sql_query(query, conn)
    conn.close()

    if not outlet_df.empty:
        # Style Summary Metrics Table
        st.subheader(f"📌 Style Summary for {view_outlet} on {view_date}")
        style_summary = outlet_df.groupby('Style Number').agg(
            System_Qty=('System Qty', 'sum'),
            Physical_Qty=('Physical Qty', 'sum'),
            Net_Variance=('Variance', 'sum')
        ).reset_index()

        st.dataframe(style_summary, use_container_width=True)

        # Style Selector for Detailed Inspection
        selected_style = st.selectbox(
            "Select a Style to inspect Color & Size Variance:", 
            ["All Styles"] + style_summary['Style Number'].tolist()
        )

        if selected_style != "All Styles":
            detailed_df = outlet_df[outlet_df['Style Number'] == selected_style]
        else:
            detailed_df = outlet_df

        st.subheader(f"📋 Detailed Color & Size Breakdown ({selected_style})")
        st.dataframe(detailed_df, use_container_width=True)

        # Download Report
        csv_out = detailed_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label=f"📥 Download Report for {view_outlet} ({view_date})",
            data=csv_out,
            file_name=f"{view_outlet}_variance_{view_date}.csv",
            mime="text/csv"
        )
    else:
        st.info(f"No records found for {view_outlet} on {view_date}.")
else:
    conn.close()
    st.info("👈 Upload daily style files and click 'Calculate & Save Daily Report' to store data.")
