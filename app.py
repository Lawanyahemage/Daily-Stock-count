"""
Stock Variance Dashboard
------------------------
Web app version of the variance checker. Admin uploads the day's data once;
outlet managers (and anyone else with the link) open the same URL and view
the selected styles + variances for their own outlet.

Run locally to test:   streamlit run app.py
Deploy for a free shareable link: push this repo to GitHub, then deploy on
https://share.streamlit.io (Streamlit Community Cloud) - see notes at bottom
of the chat message for step-by-step instructions.
"""

import pandas as pd
import streamlit as st
from io import StringIO

st.set_page_config(page_title="Stock Variance Dashboard", layout="wide")

# ---------------------------------------------------------------------------
# Core logic (same as the CLI script, adapted to work on uploaded files)
# ---------------------------------------------------------------------------

def load_master(uploaded_file):
    if uploaded_file is None:
        return None
    name = uploaded_file.name.lower()
    if name.endswith(".xlsx") or name.endswith(".xls"):
        df = pd.read_excel(uploaded_file, header=None, names=["code", "name"])
    else:
        df = pd.read_csv(uploaded_file, sep=None, engine="python", header=None, names=["code", "name"])
    df["code"] = df["code"].astype(str).str.strip()
    return df.set_index("code")["name"].to_dict()


def load_stock(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".xlsx") or name.endswith(".xls"):
        df = pd.read_excel(uploaded_file)
    else:
        df = pd.read_csv(uploaded_file, sep=None, engine="python")

    df.columns = [str(c).strip() for c in df.columns]

    # match columns case/spacing-insensitively so small naming differences don't break it
    rename_map = {}
    for col in df.columns:
        key = col.lower().replace(" ", "").replace("_", "")
        if key == "location":
            rename_map[col] = "location"
        elif key in ("colorsizecode", "coloursizecode"):
            rename_map[col] = "code"
        elif key in ("colorsizename", "coloursizename"):
            rename_map[col] = "name"
        elif key == "qty":
            rename_map[col] = "system_qty"
    df = df.rename(columns=rename_map)

    required = {"location", "code", "system_qty"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            f"Stock report is missing expected column(s): {', '.join(missing)}. "
            f"Columns found in your file: {list(df.columns)}"
        )

    df["code"] = df["code"].astype(str).str.strip()
    df["location"] = df["location"].astype(str).str.strip()
    if "name" not in df.columns:
        df["name"] = None
    return df


def load_scan_counts(uploaded_file):
    """
    Accepts scanned codes from a plain text list (one code per line) OR from
    an Excel file (codes scanned directly into a column, e.g. column A).
    Only the first column is used; any header row or blank row is ignored.
    """
    name = uploaded_file.name.lower()
    if name.endswith(".xlsx") or name.endswith(".xls"):
        raw = pd.read_excel(uploaded_file, header=None)
    else:
        raw = pd.read_csv(uploaded_file, header=None, sep=None, engine="python")

    codes = raw.iloc[:, 0].astype(str).str.strip()
    codes = codes[codes != ""]
    codes = codes[codes.str.contains(r"\d")]  # drops a stray header row like "Code"

    counts = codes.value_counts().rename("physical_qty").reset_index()
    counts.columns = ["code", "physical_qty"]
    return counts


def build_variance(stock_df, master_dict, scan_files):
    """scan_files: list of uploaded file objects, filename (without extension) = outlet code."""
    all_sheets = {}
    summary_rows = []
    warnings = []

    for scan_file in scan_files:
        outlet = scan_file.name.rsplit(".", 1)[0]
        physical = load_scan_counts(scan_file)

        outlet_stock = stock_df[stock_df["location"].str.upper() == outlet.upper()].copy()
        if outlet_stock.empty:
            warnings.append(
                f"No system stock rows found for outlet '{outlet}'. "
                f"Check the file name matches the Location value exactly."
            )

        merged = pd.merge(
            outlet_stock[["code", "name", "system_qty"]],
            physical,
            on="code",
            how="outer",
        )
        merged["physical_qty"] = merged["physical_qty"].fillna(0).astype(int)
        merged["system_qty"] = merged["system_qty"].fillna(0).astype(int)

        if master_dict:
            merged["name"] = merged.apply(
                lambda r: r["name"] if pd.notna(r["name"]) else master_dict.get(r["code"], "UNKNOWN CODE"),
                axis=1,
            )
        merged["name"] = merged["name"].fillna("UNKNOWN CODE")

        merged["variance"] = merged["physical_qty"] - merged["system_qty"]
        merged["status"] = merged["variance"].apply(
            lambda v: "MATCH" if v == 0 else ("SHORTAGE" if v < 0 else "EXCESS")
        )
        merged = merged.sort_values(["status", "code"]).reset_index(drop=True)

        all_sheets[outlet] = merged
        for _, r in merged.iterrows():
            summary_rows.append({
                "outlet": outlet, "code": r["code"], "name": r["name"],
                "system_qty": r["system_qty"], "physical_qty": r["physical_qty"],
                "variance": r["variance"], "status": r["status"],
            })

    summary = pd.DataFrame(summary_rows)
    return all_sheets, summary, warnings


def style_variance(df):
    def highlight(row):
        if row["status"] == "SHORTAGE":
            return ["background-color: #ffe0e0"] * len(row)
        elif row["status"] == "EXCESS":
            return ["background-color: #fff6d6"] * len(row)
        return [""] * len(row)
    return df.style.apply(highlight, axis=1)


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

st.title("📦 Daily Stock Variance Dashboard")

if "sheets" not in st.session_state:
    st.session_state.sheets = None
    st.session_state.summary = None

role = st.sidebar.radio("I am a:", ["Outlet Manager", "Admin (upload today's data)"])

# ---- ADMIN VIEW: upload & process ----
if role == "Admin (upload today's data)":
    st.sidebar.subheader("Upload today's files")
    master_file = st.sidebar.file_uploader("Master color-size list (optional)", type=["csv", "txt", "xlsx"])
    stock_file = st.sidebar.file_uploader("System stock report (required)", type=["csv", "txt", "xlsx"])
    scan_files = st.sidebar.file_uploader(
        "Outlet scan files (one per outlet, filename = outlet code, e.g. CCC.xlsx or CCC.txt)",
        type=["csv", "txt", "xlsx"], accept_multiple_files=True,
    )

    if st.sidebar.button("Process today's data", type="primary"):
        if stock_file is None or not scan_files:
            st.sidebar.error("Please upload the stock report and at least one outlet scan file.")
        else:
            try:
                master_dict = load_master(master_file)
                stock_df = load_stock(stock_file)
                sheets, summary, warnings = build_variance(stock_df, master_dict, scan_files)
                st.session_state.sheets = sheets
                st.session_state.summary = summary
                for w in warnings:
                    st.sidebar.warning(w)
                st.sidebar.success(f"Processed {len(scan_files)} outlet(s). Data is now live for everyone with the link.")
            except Exception as e:
                st.sidebar.error(f"Couldn't process the files: {e}")

    st.info("Upload the 3 file types on the left and click **Process today's data**. "
            "Once processed, outlet managers can switch to 'Outlet Manager' view to see it.")

# ---- Show results (both roles see this once data exists) ----
if st.session_state.summary is not None:
    summary = st.session_state.summary
    sheets = st.session_state.sheets

    if role == "Outlet Manager":
        outlets = sorted(sheets.keys())
        outlet = st.selectbox("Select your outlet", outlets)
        df = sheets[outlet]

        st.subheader(f"Selected styles & stock — {outlet}")
        st.caption("System stock quantity for today's selected styles, and your physical (scanned) count.")
        st.dataframe(style_variance(df), use_container_width=True)

        mismatches = df[df["status"] != "MATCH"]
        if mismatches.empty:
            st.success("No variances — physical count matches system stock for all items. ✅")
        else:
            st.warning(f"{len(mismatches)} item(s) with a variance at {outlet}.")

    else:  # Admin overview
        st.subheader("Variance summary — all outlets")
        mismatches = summary[summary["status"] != "MATCH"]
        st.metric("Total variance lines", len(mismatches), delta=None)
        st.dataframe(style_variance(mismatches), use_container_width=True)

        with st.expander("See all items (including matches) across all outlets"):
            st.dataframe(style_variance(summary), use_container_width=True)

        # Download as Excel
        from io import BytesIO
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            mismatches.to_excel(writer, sheet_name="Variance Summary", index=False)
            summary.to_excel(writer, sheet_name="All Items", index=False)
            for outlet, df in sheets.items():
                df.to_excel(writer, sheet_name=str(outlet)[:31], index=False)
        st.download_button(
            "Download full report (Excel)", data=buffer.getvalue(),
            file_name="variance_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    if role == "Outlet Manager":
        st.info("No data uploaded yet for today. Please check back once the admin has uploaded today's files.")
