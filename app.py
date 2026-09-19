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

st.title("🏭 Bakery, Cafe, Sweets & Savoury ERP System")
st.caption("Direct Built-in 5-Items Master & Dynamic Row Addition System (No Upload Required)")

# ----------------- BUILT-IN 5 ITEMS DEMO MASTER -----------------
BUILT_IN_MASTER = {
    "FGBK0003": {"name": "BURGER BUN SMALL", "dept": "BAKERY", "uom": "PAC", "price": 7.87, "opening": 286.0},
    "FGBK0030": {"name": "WHEAT PIZZA BASE", "dept": "BAKERY", "uom": "PAC", "price": 18.50, "opening": 150.0},
    "FGCK0007": {"name": "ALMOND & BROCCOLI SOUP", "dept": "CAFE", "uom": "PCS", "price": 120.00, "opening": 50.0},
    "FGSW0037": {"name": "MANGO KAJU CAKE", "dept": "SWEETS", "uom": "PCS", "price": 450.00, "opening": 1860.0},
    "RM0279": {"name": "MAIDA (REFINED FLOUR)", "dept": "SAVOURY", "uom": "KGS", "price": 38.00, "opening": 70.0},
}

CODE_OPTIONS = list(BUILT_IN_MASTER.keys())

# Initializing session state with 5 default items
if "demo_audit_items" not in st.session_state:
    st.session_state.demo_audit_items = pd.DataFrame([
        {"Item_Code": "FGBK0003", "Store_RM_Issued": 20.0, "Sales_Consumed": 250.0, "Wastage": 2.0, "Physical_Closing": 54.0},
        {"Item_Code": "FGBK0030", "Store_RM_Issued": 15.0, "Sales_Consumed": 100.0, "Wastage": 1.0, "Physical_Closing": 64.0},
        {"Item_Code": "FGCK0007", "Store_Issued": 10.0, "Sales_Consumed": 40.0, "Wastage": 1.0, "Physical_Closing": 19.0},
        {"Item_Code": "FGSW0037", "Store_RM_Issued": 200.0, "Sales_Consumed": 1500.0, "Wastage": 10.0, "Physical_Closing": 550.0},
        {"Item_Code": "RM0279", "Store_RM_Issued": 50.0, "Sales_Consumed": 60.0, "Wastage": 2.0, "Physical_Closing": 58.0},
    ])

# ----------------- EXCEL EXPORT FUNCTION -----------------
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
    ws["A1"] = "BUILT-IN 5 ITEMS INVENTORY CLOSING & RECONCILIATION REPORT"
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
        info = BUILT_IN_MASTER.get(code, {"name": "Unknown", "dept": "GENERAL", "uom": "PCS", "price": 0.0, "opening": 0.0})

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

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ----------------- UI TABS -----------------
tab_closing, tab_master = st.tabs([
    "📦 Built-in 5-Items Closing & Reconciliation",
    "📋 Built-in Master Directory"
])

with tab_closing:
    st.subheader("1. Interactive 5-Items Closing Audit Table (Dynamic Add Row Enabled)")
    st.caption("Aap neeচে '+' button se aur rows jod sakte hain ya existing 5 items ki values badल sakte hain:")

    edited_audit = st.data_editor(
        st.session_state.demo_audit_items,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Item_Code": st.column_config.SelectboxColumn("Item Code", options=CODE_OPTIONS, required=True),
            "Store_RM_Issued": st.column_config.NumberColumn("Store RM Issued", min_value=0.0, format="%.2f"),
            "Sales_Consumed": st.column_config.NumberColumn("Sales / Consumed", min_value=0.0, format="%.2f"),
            "Wastage": st.column_config.NumberColumn("Wastage / Scrap", min_value=0.0, format="%.2f"),
            "Physical_Closing": st.column_config.NumberColumn("Physical Closing", min_value=0.0, format="%.2f"),
        }
    )

    clean_audit = edited_audit.dropna(subset=['Item_Code']).copy()
    clean_audit['Item_Name'] = clean_audit['Item_Code'].map(lambda c: BUILT_IN_MASTER.get(c, {}).get('name', ''))
    clean_audit['Department'] = clean_audit['Item_Code'].map(lambda c: BUILT_IN_MASTER.get(c, {}).get('dept', ''))
    clean_audit['Unit_Cost'] = clean_audit['Item_Code'].map(lambda c: BUILT_IN_MASTER.get(c, {}).get('price', 0.0))
    clean_audit['Opening_Stock'] = clean_audit['Item_Code'].map(lambda c: BUILT_IN_MASTER.get(c, {}).get('opening', 0.0))
    
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
        label="📥 Download 5-Items Closing Audit Report (.xlsx)",
        data=excel_bytes,
        file_name=f"BuiltIn_5_Items_Audit_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with tab_master:
    st.subheader("📋 Built-in Demo Master Directory (5 Items)")
    master_df = pd.DataFrame([
        {"Item Code": k, "Item Name": v["name"], "Department": v["dept"], "UOM": v["uom"], "Unit Cost (₹)": v["price"], "Opening Stock": v["opening"]}
        for k, v in BUILT_IN_MASTER.items()
    ])
    st.dataframe(master_df, use_container_width=True)
