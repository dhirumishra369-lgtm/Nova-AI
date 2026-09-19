import streamlit as st
import pandas as pd
import google.generativeai as genai
import io
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import streamlit.components.v1 as components

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
st.caption("Secure Inventory Closing (20 Rows), High-Demand Recipes, Voice AI & Consultant")

# ----------------- BUILT-IN MASTER & 20 ROWS INITIALIZATION -----------------
BUILT_IN_MASTER = {
    "FGBK0003": {"name": "BURGER BUN SMALL", "dept": "BAKERY", "uom": "PAC", "price": 7.87, "opening": 286.0},
    "FGBK0030": {"name": "WHEAT PIZZA BASE", "dept": "BAKERY", "uom": "PAC", "price": 18.50, "opening": 150.0},
    "FGCK0007": {"name": "ALMOND & BROCCOLI SOUP", "dept": "CAFE", "uom": "PCS", "price": 120.00, "opening": 50.0},
    "FGSW0037": {"name": "MANGO KAJU CAKE", "dept": "SWEETS", "uom": "PCS", "price": 450.00, "opening": 1860.0},
    "RM0279": {"name": "MAIDA (REFINED FLOUR)", "dept": "SAVOURY", "uom": "KGS", "price": 38.00, "opening": 70.0},
}

CODE_OPTIONS = list(BUILT_IN_MASTER.keys())
default_code = CODE_OPTIONS[0] if CODE_OPTIONS else "ITEM1"

if "demo_audit_items" not in st.session_state:
    initial_rows = []
    for _ in range(20):
        initial_rows.append({
            "Item_Code": default_code,
            "Store_RM_Issued": 0.0,
            "Sales_Consumed": 0.0,
            "Wastage": 0.0,
            "Physical_Closing": 0.0
        })
    st.session_state.demo_audit_items = pd.DataFrame(initial_rows)

if "custom_recipes" not in st.session_state:
    st.session_state.custom_recipes = {
        "Meva Besan Laddu": {"Besan": 1.0, "Sugar": 1.0, "Ghee": 1.0},
        "Chocolate Truffle Celebration Cake": {"Dark Chocolate": 0.4, "Fresh Cream": 0.3, "Maida": 0.2, "Butter": 0.1},
        "Kaju Katli": {"Cashew Splits": 0.7, "Sugar": 0.3, "Ghee": 0.05},
        "Veg Cheese Sandwich": {"Bread Loaf": 0.5, "Paneer/Veggie Mix": 0.3, "Cheese Slice": 0.2, "Butter": 0.05},
        "Tutti Frutti Dry Cake": {"Maida": 0.4, "Sugar": 0.25, "Butter/Oil": 0.2, "Tutti Frutti & Eggs": 0.15}
    }

# ----------------- EXCEL EXPORT FUNCTION (FIRST SHEET) -----------------
def generate_closing_audit_excel(df_audit):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Closing_Audit_Report"
    ws.views.sheetView[0].showGridLines = True

    navy = "1E3A8A"
    light_b = "DBEAFE"
    border_c = "CBD5E1"
    
    thin_border = Border(left=Side(style='thin', color=border_c), right=Side(style='thin', color=border_c), top=Side(style='thin', color=border_c), bottom=Side(style='thin', color=border_c))
    double_bottom = Border(top=Side(style='thin', color=navy), bottom=Side(style='double', color=navy))
    
    ws.merge_cells("A1:M1")
    ws["A1"] = "MASTER INVENTORY CLOSING & RECONCILIATION REPORT (20 ROWS)"
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
tab_closing, tab_recipe, tab_ai = st.tabs([
    "📦 Daily Closing & Stock Reconciliation (20 Rows)",
    "🍰 High-Demand Recipe & Yield Calculator",
    "🤖 AI Assistant & Voice Command"
])

with tab_closing:
    st.subheader("1. Daily Closing & Stock Reconciliation Table (20 Rows Ready)")
    st.caption("Aapke paas 20 rows ready hain. Values update karein aur report download karein:")

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
        label="📥 Download Closing Audit Report (.xlsx)",
        data=excel_bytes,
        file_name=f"Daily_Closing_Audit_20Rows_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with tab_recipe:
    st.subheader("2. High-Demand Recipe & Yield Calculator")
    st.caption("Select any popular Bakery, Sweets or Cafe item to calculate exact raw material requirements:")

    recipe_names = list(st.session_state.custom_recipes.keys())
    selected_recipe = st.selectbox("Select High-Demand Recipe", recipe_names)

    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    with col1:
        target_prod_kg = st.number_input("Target Production (kg)", value=1000.0, min_value=1.0, step=10.0)
    with col2:
        piece_wt_gm = st.number_input("Single Piece Weight (gm)", value=30.0, min_value=1.0, step=1.0)
    with col3:
        st.write("")
        st.write("")
        add_new_recipe_expander = st.expander("➕ Add Custom Recipe")

    with add_new_recipe_expander:
        new_rec_name = st.text_input("New Recipe Name")
        new_item1 = st.text_input("Ingredient 1 Name", value="Item 1")
        new_qty1 = st.number_input("Ingredient 1 Qty", value=1.0)
        new_item2 = st.text_input("Ingredient 2 Name", value="Item 2")
        new_qty2 = st.number_input("Ingredient 2 Qty", value=1.0)
        new_item3 = st.text_input("Ingredient 3 Name", value="Item 3")
        new_qty3 = st.number_input("Ingredient 3 Qty", value=1.0)
        
        if st.button("Save New Recipe"):
            if new_rec_name:
                st.session_state.custom_recipes[new_rec_name] = {
                    new_item1: new_qty1,
                    new_item2: new_qty2,
                    new_item3: new_qty3
                }
                st.success(f"Recipe '{new_rec_name}' added successfully! Please re-select from dropdown.")
                st.rerun()

    current_recipe_ingredients = st.session_state.custom_recipes[selected_recipe]
    total_batch_ratio_wt = sum(current_recipe_ingredients.values())

    total_pieces = (target_prod_kg * 1000) / piece_wt_gm

    st.markdown(f"### 📦 Production Output Summary for **{selected_recipe}**")
    st.info(f"Target Production: **{target_prod_kg:,.2f} kg** | Total Pieces (Qty): **{total_pieces:,.0f} Pcs** (@ {piece_wt_gm}g/pc)")

    st.markdown("### 📋 Raw Material Requirement Breakdown (BOM)")
    
    breakdown_data = []
    for ing, ratio in current_recipe_ingredients.items():
        required_qty = (ratio / total_batch_ratio_wt) * target_prod_kg
        breakdown_data.append({
            "Raw Material / Ingredient": ing,
            "Ratio per Batch (kg)": f"{ratio} kg",
            f"Requirement for {target_prod_kg} kg Production": f"{required_qty:.2f} kg"
        })

    recipe_df = pd.DataFrame(breakdown_data)
    st.dataframe(recipe_df, use_container_width=True)

with tab_ai:
    st.subheader("🤖 AI Assistant & Voice Command")
    st.caption("Aap bolkar ya type karke AI se koi bhi business query pooch sakte hain:")
    
    ai_api_key = st.text_input("Enter Gemini API Key", type="password", key="ai_key")
    
    # Built-in Browser Voice Recognition HTML/JS Component
    st.markdown("### 🎙️ Live Voice Command Box")
    voice_html = """
    <div style="background-color: #f8fafc; padding: 15px; border-radius: 8px; border: 1px solid #cbd5e1;">
        <p style="font-weight: bold; color: #1e3a8a; margin-bottom: 8px;">Microphone se bolne ke liye niche diye gaye button par click karein:</p>
        <button onclick="startListening()" style="background-color: #2563eb; color: white; border: none; padding: 10px 20px; border-radius: 5px; font-weight: bold; cursor: pointer;">🎤 Start Voice Command</button>
        <p id="speechResult" style="margin-top: 10px; font-style: italic; color: #334155;"></p>
    </div>
    <script>
        function startListening() {
            const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
            if (!SpeechRecognition) {
                alert("Speech recognition is not supported in this browser. Please use Chrome.");
                return;
            }
            const recognition = new SpeechRecognition();
            recognition.lang = 'hi-IN'; // Hindi & English support
            recognition.interimResults = false;
            
            recognition.onstart = function() {
                document.getElementById("speechResult").innerText = "Listening... Kripya boliye...";
            };
            
            recognition.onresult = function(event) {
                const speechToText = event.results[0][0].transcript;
                document.getElementById("speechResult").innerText = "Aapne bola: " + speechToText;
                // We can pass this text back or display it
            };
            
            recognition.onerror = function(event) {
                document.getElementById("speechResult").innerText = "Error occurred in recognition: " + event.error;
            };
            
            recognition.start();
        }
    </script>
    """
    components.html(voice_html, height=130)

    user_query = st.text_area("Or type your query here for AI Consultant:")
    
    if st.button("Ask AI Consultant", type="primary"):
        if not ai_api_key or not user_query:
            st.warning("Please enter your Gemini API Key and query.")
        else:
            try:
                genai.configure(api_key=ai_api_key)
                model = genai.GenerativeModel("gemini-1.5-flash")
                response = model.generate_content(f"You are an expert commercial food manufacturing and ERP consultant for bakeries and sweet shops.\nQuery: {user_query}")
                st.success("AI Expert Advice:")
                st.write(response.text)
            except Exception as e:
                st.error(f"Error communicating with AI: {e}")
