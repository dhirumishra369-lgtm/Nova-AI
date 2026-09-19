import streamlit as st
import pandas as pd
import google.generativeai as genai
import io
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(
    page_title="ERP Inventory, Closing & Recipe Costing System",
    page_icon="🏭",
    layout="wide"
)

# ----------------- PASSWORD SECURITY CHECK -----------------
def check_password():
    def password_entered():
        if st.session_state["password"] == "Dhirendra@123":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input(
            "🔒 Enter Password to Access ERP System", type="password", on_change=password_entered, key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        st.text_input(
            "🔒 Enter Password to Access ERP System", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Password incorrect. Please try again.")
        return False
    else:
        return True

if not check_password():
    st.stop()

st.title("🏭 Bakery, Cafe, Sweets & Savoury Inventory & Recipe Costing System")
st.caption("100% Item-Code Driven Stock Reconciliation, Closing & Major Consumable Yield Analysis")

# ----------------- SIDEBAR MASTER UPLOADER -----------------
st.sidebar.header("📁 Upload Files")
uploaded_master = st.sidebar.file_uploader("1. Upload Items Master (.xlsx)", type=["xlsx", "xls"], key="master_file")
uploaded_consumable = st.sidebar.file_uploader("2. Upload Major Consumable File (.xlsx)", type=["xlsx", "xls"], key="consumable_file")

ITEM_MASTER = {}
if uploaded_master is not None:
    try:
        master_df = pd.read_excel(uploaded_master, sheet_name='Items')
        for _, row in master_df.dropna(subset=['No.']).iterrows():
            code = str(row['No.'].strip())
            ITEM_MASTER[code] = {
                "name": str(row.get('Description', '')),
                "dept": str(row.get('Inventory Posting Group', 'GENERAL')),
                "uom": str(row.get('Base Unit of Measure', 'PCS')),
                "price": float(row.get('Unit Cost', 0.0) or 0.0),
                "opening": float(row.get('Inventory', 0.0) or 0.0)
            }
        st.sidebar.success(f"✅ Master Loaded! ({len(ITEM_MASTER)} Items)")
    except Exception as e:
        st.sidebar.error(f"Error reading master: {e}")

CODE_OPTIONS = list(ITEM_MASTER.keys())

if "audit_items" not in st.session_state:
    st.session_state.audit_items = pd.DataFrame([
        {"Item_Code": "FGBK0003" if "FGBK0003" in CODE_OPTIONS else (CODE_OPTIONS[0] if CODE_OPTIONS else "ITEM1"), "Store_RM_Issued": 20.0, "Sales_Consumed": 250.0, "Wastage": 2.0, "Physical_Closing": 54.0},
    ])

# ----------------- EXCEL EXPORT FOR CLOSING AUDIT -----------------
def generate_closing_audit_excel(df_audit):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Closing_Audit_Report"
    ws.views.sheetView[0].showGridLines = True

    navy = "1E3A8A"
    mid_b = "2563EB"
    light_b = "DBEAFE"
    border_c = "CBD5E1"
    
    thin_border = Border(left=Side(style='thin', color=border_c), right=Side(style='thin', color=border_c), top=Side(style='thin', color=border_c), bottom=Side(style='thin', color=border_c))
    double_bottom = Border(top=Side(style='thin', color=navy), bottom=Side(style='double', color=navy))
    
    ws.merge_cells("A1:M1")
    ws["A1"] = "MASTER INVENTORY CLOSING & RECONCILIATION AUDIT REPORT"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(start_color=navy, end_color=navy, fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 34

    headers = [
        "Item Code", "Item Name", "Department", "UOM", "Unit Cost (₹)",
        "Opening Stock", "Store RM Issued", "Sales / Consumed", "Wastage / Scrap",
        "Should-Be Closing", "Physical Closing Qty", "Variance (Qty)", "Financial Impact (₹)"
    ]
    
    ws.row_dimensions[4].height = 26
    for c_idx, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=c_idx, value=h)
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=navy, end_color=navy, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")

    start_r = 5
    for idx, row in df_audit.iterrows():
        r = start_r + idx
        ws.row_dimensions[r].height = 20
        code = str(row['Item_Code'])
        info = ITEM_MASTER.get(code, {"name": "Unknown", "dept": "GENERAL", "uom": "PCS", "price": 0.0, "opening": 0.0})

        ws.cell(row=r, column=1, value=code).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=info["name"]).alignment = Alignment(horizontal="left")
        ws.cell(row=r, column=3, value=info["dept"]).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=4, value=info["uom"]).alignment = Alignment(horizontal="center")
        
        c5 = ws.cell(row=r, column=5, value=info["price"])
        c5.number_format = '₹#,##0.00'; c5.alignment = Alignment(horizontal="right")

        c6 = ws.cell(row=r, column=6, value=info["opening"])
        c6.number_format = '#,##0.00'; c6.alignment = Alignment(horizontal="right")

        c7 = ws.cell(row=r, column=7, value=float(row['Store_RM_Issued']))
        c7.number_format = '#,##0.00'; c7.alignment = Alignment(horizontal="right")

        c8 = ws.cell(row=r, column=8, value=float(row['Sales_Consumed']))
        c8.number_format = '#,##0.00'; c8.alignment = Alignment(horizontal="right")

        c9 = ws.cell(row=r, column=9, value=float(row['Wastage']))
        c9.number_format = '#,##0.00'; c9.alignment = Alignment(horizontal="right")

        c10 = ws.cell(row=r, column=10, value=f"=F{r}+G{r}-H{r}-I{r}")
        c10.number_format = '#,##0.00'; c10.alignment = Alignment(horizontal="right")

        c11 = ws.cell(row=r, column=11, value=float(row['Physical_Closing']))
        c11.number_format = '#,##0.00'; c11.alignment = Alignment(horizontal="right")

        c12 = ws.cell(row=r, column=12, value=f"=K{r}-J{r}")
        c12.number_format = '#,##0.00'; c12.alignment = Alignment(horizontal="right")
        c12.font = Font(name="Calibri", size=11, bold=True)

        c13 = ws.cell(row=r, column=13, value=f"=L{r}*E{r}")
        c13.number_format = '₹#,##0.00'; c13.alignment = Alignment(horizontal="right")
        c13.font = Font(name="Calibri", size=11, bold=True)

        bg = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
        for col_i in range(1, 14):
            c_cell = ws.cell(row=r, column=col_i)
            c_cell.border = thin_border
            c_cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")

    tot_r = start_r + len(df_audit)
    ws.row_dimensions[tot_r].height = 24
    ws.merge_cells(f"A{tot_r}:L{tot_r}")
    ws[f"A{tot_r}"] = "TOTAL NET FINANCIAL VARIANCE (₹)"
    ws[f"A{tot_r}"].font = Font(name="Calibri", size=11, bold=True)
    ws[f"A{tot_r}"].alignment = Alignment(horizontal="right", vertical="center")

    tot_fi = ws.cell(row=tot_r, column=13, value=f"=SUM(M{start_r}:M{tot_r-1})")
    tot_fi.font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    tot_fi.number_format = '₹#,##0.00'
    tot_fi.alignment = Alignment(horizontal="right", vertical="center")

    for col_i in range(1, 14):
        c_cell = ws.cell(row=tot_r, column=col_i)
        c_cell.border = double_bottom
        c_cell.fill = PatternFill(start_color=light_b, end_color=light_b, fill_type="solid")

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 10
    for col_l in ["E", "F", "G", "H", "I", "J", "K", "L"]:
        ws.column_dimensions[col_l].width = 16
    ws.column_dimensions["M"].width = 22

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ----------------- UI TABS SETUP -----------------
tab_closing, tab_master, tab_consumable, tab_ai = st.tabs([
    "📦 Daily Closing & Stock Reconciliation",
    "📋 Item Code Master Directory",
    "📊 Major Consumable & Yield Analysis",
    "🤖 AI Consultant"
])

with tab_closing:
    st.subheader("1. Item-Code Driven Store RM Issue & Closing Table")
    if len(CODE_OPTIONS) == 0:
        st.warning("⚠️ Please upload your `Items (4).xlsx` master file in the sidebar to enable full item code dropdowns.")
    
    edited_audit = st.data_editor(
        st.session_state.audit_items,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Item_Code": st.column_config.SelectboxColumn("Item Code", options=CODE_OPTIONS if CODE_OPTIONS else ["ITEM1"], required=True),
            "Store_RM_Issued": st.column_config.NumberColumn("Store RM Issued", min_value=0.0, format="%.2f"),
            "Sales_Consumed": st.column_config.NumberColumn("Sales / Consumed", min_value=0.0, format="%.2f"),
            "Wastage": st.column_config.NumberColumn("Wastage / Scrap", min_value=0.0, format="%.2f"),
            "Physical_Closing": st.column_config.NumberColumn("Physical Closing", min_value=0.0, format="%.2f"),
        }
    )

    if len(ITEM_MASTER) > 0:
        clean_audit = edited_audit.dropna(subset=['Item_Code']).copy()
        clean_audit['Item_Name'] = clean_audit['Item_Code'].map(lambda c: ITEM_MASTER.get(c, {}).get('name', ''))
        clean_audit['Department'] = clean_audit['Item_Code'].map(lambda c: ITEM_MASTER.get(c, {}).get('dept', ''))
        clean_audit['Unit_Cost'] = clean_audit['Item_Code'].map(lambda c: ITEM_MASTER.get(c, {}).get('price', 0.0))
        clean_audit['Opening_Stock'] = clean_audit['Item_Code'].map(lambda c: ITEM_MASTER.get(c, {}).get('opening', 0.0))
        
        clean_audit['Should_Be_Closing'] = clean_audit['Opening_Stock'] + clean_audit['Store_RM_Issued'] - clean_audit['Sales_Consumed'] - clean_audit['Wastage']
        clean_audit['Variance_Qty'] = clean_audit['Physical_Closing'] - clean_audit['Should_Be_Closing']
        clean_audit['Financial_Impact'] = clean_audit['Variance_Qty'] * clean_audit['Unit_Cost']

        st.markdown("### 🔍 Live Reconciliation & Variance Preview")
        st.dataframe(
            clean_audit[['Item_Code', 'Item_Name', 'Department', 'Opening_Stock', 'Store_RM_Issued', 'Sales_Consumed', 'Wastage', 'Should_Be_Closing', 'Physical_Closing', 'Variance_Qty', 'Financial_Impact']],
            use_container_width=True
        )

        st.markdown("---")
        excel_bytes = generate_closing_audit_excel(clean_audit)
        st.download_button(
            label="📥 Download Professional Closing Audit Report (.xlsx)",
            data=excel_bytes,
            file_name=f"Inventory_Closing_Audit_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

with tab_master:
    st.subheader("📋 Master Directory (All Uploaded Items)")
    if len(ITEM_MASTER) == 0:
        st.info("Please upload `Items (4).xlsx` in the sidebar.")
    else:
        master_df = pd.DataFrame([
            {"Item Code": k, "Item Name": v["name"], "Department": v["dept"], "UOM": v["uom"], "Unit Cost (₹)": v["price"], "Opening Stock": v["opening"]}
            for k, v in ITEM_MASTER.items()
        ])
        st.dataframe(master_df, use_container_width=True)

with tab_consumable:
    st.subheader("📊 Major Consumable & Yield Analysis Dashboard")
    st.caption("Upload your Major Consumable Excel file (Bakery or CK) to view Summary reconciliation and monthly production recipes:")

    if uploaded_consumable is not None:
        try:
            xl_cons = pd.ExcelFile(uploaded_consumable)
            sheets = xl_cons.sheet_names
            selected_cons_sheet = st.selectbox("Select Consumable Sheet / Month", sheets)
            
            df_cons_sheet = pd.read_excel(uploaded_consumable, sheet_name=selected_cons_sheet)
            st.success(f"Loaded Sheet: {selected_cons_sheet}")
            
            st.markdown(f"### 📋 Data Preview ({selected_cons_sheet})")
            st.dataframe(df_cons_sheet, use_container_width=True)
            
            # Download button for viewed consumable sheet
            cons_out = io.BytesIO()
            with pd.ExcelWriter(cons_out, engine='openpyxl') as writer:
                df_cons_sheet.to_excel(writer, index=False, sheet_name=selected_cons_sheet[:30])
            
            st.download_button(
                label=f"📥 Download {selected_cons_sheet} Report (.xlsx)",
                data=cons_out.getvalue(),
                file_name=f"Major_Consumable_{selected_cons_sheet}_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Error reading consumable file: {e}")
    else:
        st.info("💡 Please upload your Major Consumable file in the sidebar to view the analysis.")

with tab_ai:
    st.subheader("🤖 AI Inventory & Production Consultant")
    api_key = st.text_input("Gemini API Key", type="password")
    query = st.text_area("Your Question", placeholder="e.g., How to analyze bakery consumable discrepancies?")
    if st.button("Ask AI", type="primary"):
        if not api_key or not query:
            st.warning("Please enter API Key and Query.")
        else:
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel("gemini-1.5-flash")
                res = model.generate_content(f"You are a commercial ERP and factory production expert.\nQuestion: {query}")
                st.success("Advice:")
                st.write(res.text)
            except Exception as e:
                st.error(f"Error: {e}")
