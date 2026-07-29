import streamlit as st
import pandas as pd
from collections import Counter
import io

st.set_page_config(page_title="Outlet Stock Variance Dashboard", layout="wide")
st.title("📦 Outlet Daily Stock Variance Dashboard")

# --- DATA PROCESSING FUNCTION ---
def process_scanned_data(file_or_text, location_name):
    """
    Parses scanned codes from either uploaded files (.xlsx, .csv, .txt) 
    or pasted text and counts occurrences.
    """
    codes = []
    
    if file_or_text is None:
        return []
    
    # Check if input is an uploaded file
    if hasattr(file_or_text, 'name'):
        filename = file_or_text.name.lower()
        if filename.endswith(('.xlsx', '.xls')):
            # Read without header so top row code isn't lost
            df_temp = pd.read_excel(file_or_text, header=None)
            codes = df_temp.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        elif filename.endswith('.csv'):
            df_temp = pd.read_csv(file_or_text, header=None)
            codes = df_temp.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        elif filename.endswith('.txt'):
            stringio = io.StringIO(file_or_text.getvalue().decode("utf-8"))
            codes = [line.strip() for line in stringio.readlines() if line.strip()]
    elif isinstance(file_or_text, str):
        # Handle pasted raw text
        codes = [line.strip() for line in file_or_text.split('\n') if line.strip()]
        
    return codes

def parse_color_size(row):
    """
    Extracts Color and Size cleanly from 'Color Size Name'
    Example: '25933700 - WW PLEATING DRESS BEIGE-06' -> Color: BEIGE, Size: 06
    """
    name = str(row.get('Color Size Name', ''))
    if '-' in name:
        parts = name.rsplit('-', 1)
        size = parts[-1].strip()
        color_part = parts[0].split('-')[-1].strip() if len(parts[0].split('-')) > 1 else ''
        return pd.Series([color_part, size])
    return pd.Series(['N/A', 'N/A'])


# --- SIDEBAR CONTROLS ---
st.sidebar.header("1. System Master Stock File")
erp_file = st.sidebar.file_uploader("Upload Raysoft ERP File (.xlsx)", type=["xlsx", "xls"], key="erp")

st.sidebar.header("2. Outlet Scans (Upload File OR Paste)")
outlets = ["CCC", "COLOMBO 03", "NUGEGODA", "ONLINE", "WATTALA"]
outlet_scans = {}

for outlet in outlets:
    with st.sidebar.expander(f"📍 {outlet} Data Entry"):
        tab1, tab2 = st.tabs(["📁 Upload File", "📋 Paste Codes"])
        
        with tab1:
            uploaded_scans = st.file_uploader(
                f"Upload scan file for {outlet}", 
                type=["xlsx", "xls", "csv", "txt"], 
                key=f"file_{outlet}"
            )
        
        with tab2:
            pasted_text = st.text_area(
                f"Paste barcode lines for {outlet}:", 
                height=100, 
                key=f"text_{outlet}"
            )
        
        # Priority: File input if available, otherwise pasted text
        if uploaded_scans:
            codes = process_scanned_data(uploaded_scans, outlet)
            outlet_scans[outlet] = codes
            st.success(f"Loaded {len(codes)} scanned items from file.")
        elif pasted_text:
            codes = process_scanned_data(pasted_text, outlet)
            outlet_scans[outlet] = codes
            st.success(f"Loaded {len(codes)} scanned items from pasted text.")


# --- MAIN APPLICATION LOGIC ---
if erp_file:
    # 1. Read ERP Master Data
    erp_df = pd.read_excel(erp_file)
    erp_df.columns = erp_df.columns.str.strip()
    if 'Unnamed: 0' in erp_df.columns:
        erp_df = erp_df.drop(columns=['Unnamed: 0'])
    erp_df.rename(columns={'Qty': 'System Qty'}, inplace=True)

    # 2. Extract Color and Size attributes
    erp_df[['Color', 'Size']] = erp_df.apply(parse_color_size, axis=1)

    # 3. Process Physical Outlet Scans
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

    # 4. Merge ERP and Physical Scans
    if not scanned_df.empty:
        merged_df = pd.merge(erp_df, scanned_df, on=['Location', 'Color Size Code'], how='outer')
    else:
        merged_df = erp_df.copy()
        merged_df['Physical Qty'] = 0

    merged_df['System Qty'] = merged_df['System Qty'].fillna(0)
    merged_df['Physical Qty'] = merged_df['Physical Qty'].fillna(0)
    merged_df['Variance'] = merged_df['Physical Qty'] - merged_df['System Qty']

    # 5. Dashboard Summary KPI Cards
    st.subheader("📊 Key Metrics Summary")
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Total System Stock", int(merged_df['System Qty'].sum()))
    kpi2.metric("Total Physical Scanned", int(merged_df['Physical Qty'].sum()))
    kpi3.metric("Net Variance", int(merged_df['Variance'].sum()))

    st.markdown("---")

    # 6. Interactive Filters
    st.subheader("🔍 Filter Report")
    f_col1, f_col2, f_col3 = st.columns(3)
    
    selected_outlet = f_col1.selectbox("Filter Outlet", ["All Outlets"] + outlets)
    colors_available = ["All Colors"] + [c for c in merged_df['Color'].dropna().unique() if c != 'N/A']
    selected_color = f_col2.selectbox("Filter Color", colors_available)
    
    sizes_available = ["All Sizes"] + [s for s in merged_df['Size'].dropna().unique() if s != 'N/A']
    selected_size = f_col3.selectbox("Filter Size", sizes_available)

    # Apply Filters
    filtered_df = merged_df.copy()
    if selected_outlet != "All Outlets":
        filtered_df = filtered_df[filtered_df['Location'] == selected_outlet]
    if selected_color != "All Colors":
        filtered_df = filtered_df[filtered_df['Color'] == selected_color]
    if selected_size != "All Sizes":
        filtered_df = filtered_df[filtered_df['Size'] == selected_size]

    # 7. Display Main Table
    st.subheader("📋 Color, Size & Outlet-wise Variance Matrix")
    
    display_cols = ['Location', 'Color Size Code', 'Color Size Name', 'Color', 'Size', 'System Qty', 'Physical Qty', 'Variance']
    available_display_cols = [c for c in display_cols if c in filtered_df.columns]
    
    st.dataframe(filtered_df[available_display_cols], use_container_width=True)

    # 8. Export Option
    csv_bytes = filtered_df[available_display_cols].to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Detailed Report (CSV)",
        data=csv_bytes,
        file_name="outlet_color_size_variance.csv",
        mime="text/csv"
    )

else:
    st.info("👈 Please upload the Raysoft ERP stock file in the left sidebar to generate reports.")
