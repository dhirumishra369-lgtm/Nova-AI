import streamlit as st
import pandas as pd
import google.generativeai as genai
import io
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(
    page_title="Recipe Costing & Yield Calculator",
    page_icon="🍰",
    layout="wide"
)

# ----------------- SESSION STATE & SETUP -----------------
if "ingredients" not in st.session_state:
    st.session_state.ingredients = pd.DataFrame([
        {"Ingredient": "Kaju (Cashew)", "Quantity_KG": 5.0, "Rate_Per_KG": 680.0},
        {"Ingredient": "Sugar", "Quantity_KG": 8.0, "Rate_Per_KG": 42.0},
        {"Ingredient": "Silver Vark", "Quantity_KG": 0.05, "Rate_Per_KG": 4000.0},
        {"Ingredient": "Cardamom / Ghee", "Quantity_KG": 0.2, "Rate_Per_KG": 600.0}
    ])

# ----------------- EXCEL EXPORT FUNCTION -----------------
def generate_professional_excel(recipe_name, df, loss_pct, final_yield, raw_cost, labor_cost, pack_cost, total_cost, cost_kg, margin_pct, sp_kg, profit_kg):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Costing & Yield Report"
    ws.views.sheetView[0].showGridLines = True

    # Title Banner
    ws.merge_cells("A1:E1")
    title = ws["A1"]
    title.value = "RECIPE COSTING & BATCH YIELD REPORT"
    title.font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    title.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    # Metadata Row
    ws["A2"] = "Product / Recipe:"
    ws["B2"] = recipe_name
    ws["D2"] = "Date:"
    ws["E2"] = datetime.date.today().strftime("%d-%b-%Y")
    ws["A2"].font = Font(name="Calibri", size=11, bold=True, color="4B5563")
    ws["B2"].font = Font(name="Calibri", size=11, bold=True, color="111827")
    ws["D2"].font = Font(name="Calibri", size=11, bold=True, color="4B5563")
    ws["E2"].font = Font(name="Calibri", size=11, bold=True, color="111827")
    ws.row_dimensions[2].height = 22

    # Section 1 Header: Raw Material
    ws.merge_cells("A4:E4")
    sec1 = ws["A4"]
    sec1.value = "1. RAW MATERIAL & INGREDIENT BREAKDOWN"
    sec1.font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    sec1.fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    ws.row_dimensions[4].height = 24

    headers = ["S.No.", "Ingredient Name", "Quantity (KG)", "Rate / KG (₹)", "Total Amount (₹)"]
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=5, column=col_idx, value=h)
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
        c.alignment = Alignment(horizontal="center" if col_idx != 2 else "left", vertical="center")
    ws.row_dimensions[5].height = 24

    thin_border = Border(
        left=Side(style='thin', color='E5E7EB'),
        right=Side(style='thin', color='E5E7EB'),
        top=Side(style='thin', color='E5E7EB'),
        bottom=Side(style='thin', color='E5E7EB')
    )

    start_row = 6
    for idx, row in df.iterrows():
        curr = start_row + idx
        ws.row_dimensions[curr].height = 20
        ws.cell(row=curr, column=1, value=idx+1).alignment = Alignment(horizontal="center")
        ws.cell(row=curr, column=2, value=str(row['Ingredient'])).alignment = Alignment(horizontal="left")
        
        c3 = ws.cell(row=curr, column=3, value=float(row['Quantity_KG']))
        c3.number_format = '#,##0.00'
        c3.alignment = Alignment(horizontal="right")
        
        c4 = ws.cell(row=curr, column=4, value=float(row['Rate_Per_KG']))
        c4.number_format = '₹#,##0.00'
        c4.alignment = Alignment(horizontal="right")
        
        c5 = ws.cell(row=curr, column=5, value=f"=C{curr}*D{curr}")
        c5.number_format = '₹#,##0.00'
        c5.alignment = Alignment(horizontal="right")

        bg_color = "F9FAFB" if idx % 2 == 1 else "FFFFFF"
        for c_idx in range(1, 6):
            ws.cell(row=curr, column=c_idx).border = thin_border
            ws.cell(row=curr, column=c_idx).fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")

    tot_r = start_row + len(df)
    ws.row_dimensions[tot_r].height = 22
    ws.cell(row=tot_r, column=2, value="Subtotal (Raw Materials)").font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=tot_r, column=3, value=f"=SUM(C{start_row}:C{tot_r-1})").font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=tot_r, column=3).number_format = '#,##0.00'
    ws.cell(row=tot_r, column=3).alignment = Alignment(horizontal="right")
    
    sub_amt = ws.cell(row=tot_r, column=5, value=f"=SUM(E{start_row}:E{tot_r-1})")
    sub_amt.font = Font(name="Calibri", size=11, bold=True)
    sub_amt.number_format = '₹#,##0.00'
    sub_amt.alignment = Alignment(horizontal="right")

    sum_border = Border(top=Side(style='thin', color='1E3A8A'), bottom=Side(style='double', color='1E3A8A'))
    for c_idx in range(1, 6):
        ws.cell(row=tot_r, column=c_idx).border = sum_border
        ws.cell(row=tot_r, column=c_idx).fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

    # Section 2 Header: Summary
    sum_start = tot_r + 2
    ws.merge_cells(f"A{sum_start}:E{sum_start}")
    sec2 = ws[f"A{sum_start}"]
    sec2.value = "2. YIELD LOSS, OVERHEADS & FINANCIAL SUMMARY"
    sec2.font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    sec2.fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    ws.row_dimensions[sum_start].height = 24

    metrics = [
        ("Total Raw Material Batch Weight", f"=C{tot_r}", "KG", '#,##0.00'),
        ("Process / Moisture Loss (%)", loss_pct / 100.0, "%", '0.0%'),
        ("Yield Loss Quantity (Wastage/Evaporation)", f"=E{sum_start+1}*E{sum_start+2}", "KG", '#,##0.00'),
        ("Final Net Yield Output", f"=E{sum_start+1}-E{sum_start+3}", "KG", '#,##0.00'),
        ("Raw Material Total Cost", f"=E{tot_r}", "INR", '₹#,##0.00'),
        ("Labour & Fuel Overheads", float(labor_cost), "INR", '₹#,##0.00'),
        ("Packaging & Box Cost", float(pack_cost), "INR", '₹#,##0.00'),
        ("Total Batch Production Cost", f"=SUM(E{sum_start+5}:E{sum_start+7})", "INR", '₹#,##0.00'),
        ("Final Cost Per KG (True Cost)", f"=E{sum_start+8}/E{sum_start+4}", "INR/KG", '₹#,##0.00'),
        ("Target Profit Margin (%)", margin_pct / 100.0, "%", '0.0%'),
        ("Suggested Selling Price Per KG", f"=E{sum_start+9}/(1-E{sum_start+10})", "INR/KG", '₹#,##0.00'),
        ("Net Profit Per KG", f"=E{sum_start+11}-E{sum_start+9}", "INR/KG", '₹#,##0.00'),
    ]

    for idx, (label, val, unit, num_fmt) in enumerate(metrics):
        r = sum_start + 1 + idx
        ws.row_dimensions[r].height = 21
        ws.merge_cells(f"A{r}:C{r}")
        ws[f"A{r}"] = label
        ws[f"A{r}"].alignment = Alignment(horizontal="left", vertical="center")
        
        ws[f"D{r}"] = unit
        ws[f"D{r}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"D{r}"].font = Font(name="Calibri", size=10, color="6B7280")
        
        v = ws[f"E{r}"]
        v.value = val
        v.number_format = num_fmt
        v.alignment = Alignment(horizontal="right", vertical="center")

        is_high = label in ["Total Batch Production Cost", "Final Cost Per KG (True Cost)", "Suggested Selling Price Per KG", "Net Profit Per KG"]
        fill_col = "FEF3C7" if ("Selling" in label or "Profit" in label) else ("E0F2FE" if is_high else "FFFFFF")
        for c_idx in range(1, 6):
            cell = ws.cell(row=r, column=c_idx)
            cell.border = thin_border
            cell.fill = PatternFill(start_color=fill_col, end_color=fill_col, fill_type="solid")
            if is_high:
                cell.font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 22

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ----------------- UI TABS -----------------
tab1, tab2 = st.tabs(["📊 Recipe Costing & Yield Calculator", "🤖 AI Chef & Production Consultant"])

with tab1:
    st.caption("Precise cost, batch yield, and margin analysis for commercial kitchens, confectioneries, and bakeries.")
    
    col_left, col_right = st.columns([1.1, 0.9], gap="large")

    with col_left:
        st.subheader("1. Batch & Product Details")
        recipe_name = st.text_input("Product / Recipe Name", value="Premium Kaju Katli")

        st.markdown("**Ingredients & Raw Material Rates:**")
        st.caption("Click directly inside any cell to edit ingredient names, weight (KG), and rate per KG:")
        
        edited_df = st.data_editor(
            st.session_state.ingredients,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Ingredient": st.column_config.TextColumn("Ingredient", required=True),
                "Quantity_KG": st.column_config.NumberColumn("Quantity (KG)", min_value=0.001, format="%.3f"),
                "Rate_Per_KG": st.column_config.NumberColumn("Rate / KG (₹)", min_value=0.0, format="₹%.2f"),
            }
        )

        st.subheader("2. Yield Loss & Overheads")
        c1, c2, c3 = st.columns(3)
        with c1:
            loss_percent = st.number_input("Cooking / Moisture Loss (%)", min_value=0.0, max_value=90.0, value=12.0, step=0.5)
        with c2:
            labor_gas_cost = st.number_input("Labor + Fuel Cost (₹)", min_value=0.0, value=400.0, step=50.0)
        with c3:
            packaging_cost = st.number_input("Packaging Cost (₹)", min_value=0.0, value=250.0, step=50.0)

        target_margin = st.slider("Target Gross Margin (%)", min_value=5.0, max_value=80.0, value=35.0, step=1.0)

    # ----------------- CALCULATIONS -----------------
    clean_df = edited_df.dropna(subset=['Quantity_KG', 'Rate_Per_KG']).copy()
    raw_material_weight = clean_df['Quantity_KG'].sum()
    clean_df['Cost'] = clean_df['Quantity_KG'] * clean_df['Rate_Per_KG']
    raw_material_cost = clean_df['Cost'].sum()

    yield_loss_kg = raw_material_weight * (loss_percent / 100.0)
    final_yield_kg = raw_material_weight - yield_loss_kg
    total_batch_cost = raw_material_cost + labor_gas_cost + packaging_cost

    cost_per_kg = (total_batch_cost / final_yield_kg) if final_yield_kg > 0 else 0.0
    selling_price_per_kg = (cost_per_kg / (1 - (target_margin / 100.0))) if target_margin < 100 else 0.0
    profit_per_kg = selling_price_per_kg - cost_per_kg

    with col_right:
        st.subheader("📋 Yield Loss Breakdown")
        st.markdown(f"""
        <div style="background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 12px; padding: 20px; margin-bottom: 20px;">
            <div style="display: flex; justify-content: space-between; margin-bottom: 12px;">
                <div>
                    <p style="margin:0; font-size:12px; color:#64748B;">Total Input</p>
                    <h4 style="margin:0; color:#0F172A;">{raw_material_weight:,.2f} KG</h4>
                </div>
                <div>
                    <p style="margin:0; font-size:12px; color:#DC2626;">📉 Yield Loss</p>
                    <h4 style="margin:0; color:#DC2626;">{yield_loss_kg:,.2f} KG ({loss_percent}%)</h4>
                </div>
                <div>
                    <p style="margin:0; font-size:12px; color:#16A34A;">✅ Net Yield</p>
                    <h4 style="margin:0; color:#16A34A;">{final_yield_kg:,.2f} KG</h4>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Generate Professional Excel File
        excel_file_bytes = generate_professional_excel(
            recipe_name, clean_df, loss_percent, final_yield_kg,
            raw_material_cost, labor_gas_cost, packaging_cost,
            total_batch_cost, cost_per_kg, target_margin,
            selling_price_per_kg, profit_per_kg
        )

        st.download_button(
            label="📥 Download Professional Excel Report (.xlsx)",
            data=excel_file_bytes,
            file_name=f"{recipe_name.replace(' ', '_')}_Costing_Sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

with tab2:
    st.subheader("🤖 AI Chef & Production Consultant")
    st.caption("Ask questions about recipe optimization, shelf-life improvement, process loss reduction, or packaging tips.")
    
    api_key = st.text_input("Enter Gemini API Key", type="password")
    user_query = st.text_area("Your Question", placeholder="e.g., How can I reduce moisture loss below 12% in cashew fudge? Or how do I extend shelf-life without chemical preservatives?")
    
    if st.button("Ask AI Consultant", type="primary"):
        if not api_key:
            st.warning("Please enter your Gemini API Key first.")
        elif not user_query:
            st.warning("Please type a question.")
        else:
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel("gemini-2.5-flash")
                prompt = f"""
                You are an expert commercial food technologist, chef, and bakery production consultant.
                Current Product: {recipe_name}
                Raw Material Batch Weight: {raw_material_weight:.2f} kg
                Process / Yield Loss: {loss_percent}% ({yield_loss_kg:.2f} kg lost)
                Final Output Yield: {final_yield_kg:.2f} kg

                User Query: {user_query}
                Please provide practical, accurate, and scientifically backed commercial kitchen guidance.
                """
                with st.spinner("AI is analyzing your recipe..."):
                    response = model.generate_content(prompt)
                    st.success("Consultant Recommendation:")
                    st.write(response.text)
            except Exception as e:
                st.error(f"Error: {e}")
