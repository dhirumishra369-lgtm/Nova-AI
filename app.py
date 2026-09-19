import streamlit as st
import pandas as pd
import io
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ----------------- MAJOR CONSUMABLE MODULE EXTENSION -----------------
st.sidebar.markdown("---")
st.sidebar.header("📊 Major Consumable Upload (Optional)")
uploaded_consumable = st.sidebar.file_uploader("Upload Major Consumable Excel (.xlsx)", type=["xlsx", "xls"], key="consumable_upload")

# Tab navigation update
tab_closing, tab_master, tab_consumable = st.tabs([
    "📦 Store RM Issue & Closing Reconciliation",
    "📋 Master Directory (All Items)",
    "📊 Major Consumable & Yield Analysis"
])

with tab_consumable:
    st.subheader("📊 Major Consumable Consumption & Discrepancy Report")
    st.caption("Yahan aap apni major consumable files upload karke Summary aur Discrepancy check kar sakte hain:")

    if uploaded_consumable is not None:
        try:
            xl_cons = pd.ExcelFile(uploaded_consumable)
            sheet_list = xl_cons.sheet_names
            selected_sheet = st.selectbox("Select Month / Sheet", sheet_list)
            
            df_cons = pd.read_excel(uploaded_consumable, sheet_name=selected_sheet)
            st.success(f"Successfully loaded sheet: {selected_sheet}")
            
            # Preview dataframe
            st.dataframe(df_cons, use_container_width=True)
            
        except Exception as e:
            st.error(f"Error loading consumable file: {e}")
    else:
            st.info("💡 Kripya sidebar se apni Major Consumable Excel file upload karein (jaise Bakery ya CK major consumable).")
