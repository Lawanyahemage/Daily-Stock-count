import streamlit as st
import pandas as pd
from collections import Counter

# Set up page layout
st.set_page_config(page_title="Stock Variance Automation", layout="wide")
st.title("📦 Outlet Daily Stock Variance Dashboard")

# --- 1. PROCESSING FUNCTION ---
def process_daily_variance(erp_excel_path, outlet_scans_dict):
    """
    Processes ERP stock file and scanned barcodes to calculate variance.
    """
    # Read ERP stock Excel file
    erp_df = pd.read_excel(erp_excel_path)
    
    # Clean up column names and drop extra index column if present
    erp_df.columns = erp_df.columns.str.strip()
    if 'Unnamed: 0' in erp_df.columns:
        erp_df = erp_df.drop(columns=['Unnamed: 0'])
        
    erp_df.rename(columns={'Qty': 'System Qty'}, inplace=True)

    # Convert scanned lists into counted DataFrames
    scanned_list = []
    for location, barcodes in outlet_scans_dict.items():
        counts = Counter([b.strip() for b in barcodes if b.strip()])
        for code, count in counts.items():
            scanned_list.append({
                'Location': location.strip(),
                'Color Size Code': code,
                'Physical Qty': count
            })
            
    scanned_df = pd.DataFrame(scanned_list)

    # Full outer merge to catch mismatches
    if not scanned_df.empty:
        merged_df = pd.merge(
            erp_df, 
            scanned_df, 
            on=['Location', 'Color Size Code'], 
            how='outer'
        )
    else:
        merged_df = erp_df.copy()
        merged_df['Physical Qty'] = 0

    # Fill missing quantities with 0
    merged_df['System Qty'] = merged_df['System Qty'].fillna(0)
    merged_df['Physical Qty'] = merged_df['Physical Qty'].fillna(0)

    # Compute Variance: (Physical Scanned - System Expected)
    merged_df['Variance'] = merged_df['Physical Qty'] - merged_df['System Qty']

    # Status Flag
    def categorize(var):
        if var == 0:
            return "MATCHED"
        elif var > 0:
            return "SURPLUS (+)"
        else:
            return "SHORTAGE (-)"

    merged_df['Status'] = merged_df['Variance'].apply(categorize)

    return merged_df


# --- 2. STREAMLIT USER INTERFACE ---

# Sidebar - Step 1: Upload ERP Data
st.sidebar.header("1. Upload System File")
erp_file = st.sidebar.file_uploader("Upload Raysoft Stock File (.xlsx)", type=["xlsx", "xls"])

# Sidebar - Step 2: Input Scanned Barcodes
st.sidebar.header("2. Input Scanned Barcodes")
outlets = ["CCC", "COLOMBO 03", "NUGEGODA", "ONLINE", "WATTALA"]
scanned_inputs = {}

for outlet in outlets:
    with st.sidebar.expander(f"Scans for {outlet}"):
        raw_text = st.text_area(f"Paste barcodes for {outlet}:", height=120, key=outlet)
        if raw_text:
            codes = [line.strip() for line in raw_text.split("\n") if line.strip()]
            scanned_inputs[outlet] = codes

# Main Page Display
if erp_file:
    # Run calculations
    results_df = process_daily_variance(erp_file, scanned_inputs)
    
    # Filter options
    selected_outlet = st.selectbox("Filter Results by Location", ["All Locations"] + outlets)
    if selected_outlet != "All Locations":
        view_df = results_df[results_df['Location'] == selected_outlet]
    else:
        view_df = results_df

    # Summary Metrics
    st.subheader("Summary Metrics")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total System Qty", int(view_df['System Qty'].sum()))
    col2.metric("Total Scanned Qty", int(view_df['Physical Qty'].sum()))
    col3.metric("Net Discrepancy", int(view_df['Variance'].sum()))

    # Detailed Table
    st.subheader("Variance Breakdown")
    st.dataframe(view_df, use_container_width=True)

    # Download button for export
    csv_data = view_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Variance Report (CSV)",
        data=csv_data,
        file_name="daily_variance_report.csv",
        mime="text/csv"
    )

else:
    st.info("👈 Please upload the Raysoft ERP stock file in the sidebar to get started.")
