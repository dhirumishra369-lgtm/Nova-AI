import datetime
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Recipe Costing & Yield Calculator", page_icon="🍰", layout="wide"
)

# ----------------- SESSION STATE & SETUP -----------------
if "ingredients" not in st.session_state:
  st.session_state.ingredients = pd.DataFrame([
      {"Ingredient": "Kaju (Cashew)", "Quantity_KG": 5.0, "Rate_Per_KG": 680.0},
      {"Ingredient": "Sugar", "Quantity_KG": 8.0, "Rate_Per_KG": 42.0},
      {"Ingredient": "Silver Vark", "Quantity_KG": 0.05, "Rate_Per_KG": 4000.0},
      {"Ingredient": "Cardamom / Ghee", "Quantity_KG": 0.2, "Rate_Per_KG": 600.0},
  ])

if "recipe_title" not in st.session_state:
  st.session_state.recipe_title = "Premium Kaju Katli"

if "master_excel_data" not in st.session_state:
  st.session_state.master_excel_data = None

if "price_lookup_dict" not in st.session_state:
  st.session_state.price_lookup_dict = {}

# Standard Nutritional Database per 100g for common ingredients
NUTRIENTS_DB = {
    "kaju": {"kcal": 553, "protein": 18.2, "fat": 43.8, "carbs": 30.2},
    "cashew": {"kcal": 553, "protein": 18.2, "fat": 43.8, "carbs": 30.2},
    "sugar": {"kcal": 387, "protein": 0.0, "fat": 0.0, "carbs": 100.0},
    "ghee": {"kcal": 900, "protein": 0.0, "fat": 100.0, "carbs": 0.0},
    "butter": {"kcal": 717, "protein": 0.85, "fat": 81.0, "carbs": 0.06},
    "milk": {"kcal": 42, "protein": 3.4, "fat": 1.0, "carbs": 5.0},
    "pistachio": {"kcal": 562, "protein": 20.0, "fat": 45.0, "carbs": 28.0},
    "cardamom": {"kcal": 311, "protein": 11.0, "fat": 7.0, "carbs": 68.0},
    "silver vark": {"kcal": 0.0, "protein": 0.0, "fat": 0.0, "carbs": 0.0},
    "salt": {"kcal": 0.0, "protein": 0.0, "fat": 0.0, "carbs": 0.0},
    "vanilla essence": {"kcal": 288, "protein": 0.1, "fat": 0.06, "carbs": 12.65},
}


def get_nutrients(ingredient_name):
  name_lower = str(ingredient_name).lower()
  for key, val in NUTRIENTS_DB.items():
    if key in name_lower:
      return val
  return {"kcal": 400, "protein": 5.0, "fat": 10.0, "carbs": 70.0}


# ----------------- EXCEL EXPORT FUNCTION -----------------
def generate_professional_excel(
    recipe_name,
    df,
    loss_pct,
    final_yield,
    raw_cost,
    labor_cost,
    pack_cost,
    total_cost,
    cost_kg,
    margin_pct,
    sp_kg,
    profit_kg,
):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Costing & Yield Report"
  ws.views.sheetView[0].showGridLines = True

  # Title Banner
  ws.merge_cells("A1:E1")
  title = ws["A1"]
  title.value = "RECIPE COSTING & BATCH YIELD REPORT"
  title.font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
  title.fill = PatternFill(
      start_color="1E3A8A", end_color="1E3A8A", fill_type="solid"
  )
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
  sec1.fill = PatternFill(
      start_color="DBEAFE", end_color="DBEAFE", fill_type="solid"
  )
  ws.row_dimensions[4].height = 24

  headers = [
      "S.No.",
      "Ingredient Name",
      "Quantity (KG)",
      "Rate / KG (₹)",
      "Total Amount (₹)",
  ]
  for col_idx, h in enumerate(headers, 1):
    c = ws.cell(row=5, column=col_idx, value=h)
    c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c.fill = PatternFill(
        start_color="2563EB", end_color="2563EB", fill_type="solid"
    )
    c.alignment = Alignment(
        horizontal="center" if col_idx != 2 else "left", vertical="center"
    )
  ws.row_dimensions[5].height = 24

  thin_border = Border(
      left=Side(style="thin", color="E5E7EB"),
      right=Side(style="thin", color="E5E7EB"),
      top=Side(style="thin", color="E5E7EB"),
      bottom=Side(style="thin", color="E5E7EB"),
  )

  start_row = 6
  for idx, row in df.iterrows():
    curr = start_row + idx
    ws.row_dimensions[curr].height = 20
    ws.cell(row=curr, column=1, value=idx + 1).alignment = Alignment(
        horizontal="center"
    )
    ws.cell(row=curr, column=2, value=str(row["Ingredient"])).alignment = (
        Alignment(horizontal="left")
    )

    c3 = ws.cell(row=curr, column=3, value=float(row["Quantity_KG"]))
    c3.number_format = "#,##0.00"
    c3.alignment = Alignment(horizontal="right")

    c4 = ws.cell(row=curr, column=4, value=float(row["Rate_Per_KG"]))
    c4.number_format = "₹#,##0.00"
    c4.alignment = Alignment(horizontal="right")

    c5 = ws.cell(row=curr, column=5, value=f"=C{curr}*D{curr}")
    c5.number_format = "₹#,##0.00"
    c5.alignment = Alignment(horizontal="right")

    bg_color = "F9FAFB" if idx % 2 == 1 else "FFFFFF"
    for c_idx in range(1, 6):
      ws.cell(row=curr, column=c_idx).border = thin_border
      ws.cell(row=curr, column=c_idx).fill = PatternFill(
          start_color=bg_color, end_color=bg_color, fill_type="solid"
      )

  tot_r = start_row + len(df)
  ws.row_dimensions[tot_r].height = 22
  ws.cell(row=tot_r, column=2, value="Subtotal (Raw Materials)").font = Font(
      name="Calibri", size=11, bold=True
  )
  ws.cell(
      row=tot_r, column=3, value=f"=SUM(C{start_row}:C{tot_r-1})"
  ).font = Font(name="Calibri", size=11, bold=True)
  ws.cell(row=tot_r, column=3).number_format = "#,##0.00"
  ws.cell(row=tot_r, column=3).alignment = Alignment(horizontal="right")

  sub_amt = ws.cell(row=tot_r, column=5, value=f"=SUM(E{start_row}:E{tot_r-1})")
  sub_amt.font = Font(name="Calibri", size=11, bold=True)
  sub_amt.number_format = "₹#,##0.00"
  sub_amt.alignment = Alignment(horizontal="right")

  sum_border = Border(
      top=Side(style="thin", color="1E3A8A"),
      bottom=Side(style="double", color="1E3A8A"),
  )
  for c_idx in range(1, 6):
    ws.cell(row=tot_r, column=c_idx).border = sum_border
    ws.cell(row=tot_r, column=c_idx).fill = PatternFill(
        start_color="EFF6FF", end_color="EFF6FF", fill_type="solid"
    )

  # Section 2 Header: Summary
  sum_start = tot_r + 2
  ws.merge_cells(f"A{sum_start}:E{sum_start}")
  sec2 = ws[f"A{sum_start}"]
  sec2.value = "2. YIELD LOSS, OVERHEADS & FINANCIAL SUMMARY"
  sec2.font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
  sec2.fill = PatternFill(
      start_color="DBEAFE", end_color="DBEAFE", fill_type="solid"
  )
  ws.row_dimensions[sum_start].height = 24

  metrics = [
      ("Total Raw Material Batch Weight", f"=C{tot_r}", "KG", "#,##0.00"),
      ("Process / Moisture Loss (%)", loss_pct / 100.0, "%", "0.0%"),
      (
          "Final Net Yield Output",
          f"=E{sum_start+1}*(1-E{sum_start+2})",
          "KG",
          "#,##0.00",
      ),
      ("Raw Material Total Cost", f"=E{tot_r}", "INR", "₹#,##0.00"),
      ("Labour & Fuel Overheads", float(labor_cost), "INR", "₹#,##0.00"),
      ("Packaging & Box Cost", float(pack_cost), "INR", "₹#,##0.00"),
      (
          "Total Batch Production Cost",
          f"=SUM(E{sum_start+4}:E{sum_start+6})",
          "INR",
          "₹#,##0.00",
      ),
      (
          "Final Cost Per KG (True Cost)",
          f"=E{sum_start+7}/E{sum_start+3}",
          "INR/KG",
          "₹#,##0.00",
      ),
      ("Target Profit Margin (%)", margin_pct / 100.0, "%", "0.0%"),
      (
          "Suggested Selling Price Per KG",
          f"=E{sum_start+8}/(1-E{sum_start+9})",
          "INR/KG",
          "₹#,##0.00",
      ),
      (
          "Net Profit Per KG",
          f"=E{sum_start+10}-E{sum_start+8}",
          "INR/KG",
          "₹#,##0.00",
      ),
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

    is_high = label in [
        "Total Batch Production Cost",
        "Final Cost Per KG (True Cost)",
        "Suggested Selling Price Per KG",
        "Net Profit Per KG",
    ]
    fill_col = (
        "FEF3C7"
        if ("Selling" in label or "Profit" in label)
        else ("E0F2FE" if is_high else "FFFFFF")
    )
    for c_idx in range(1, 6):
      cell = ws.cell(row=r, column=c_idx)
      cell.border = thin_border
      cell.fill = PatternFill(
          start_color=fill_col, end_color=fill_col, fill_type="solid"
      )
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


# ----------------- UI TABS SETUP -----------------
tab1, tab2 = st.tabs([
    "📊 Recipe Costing, Yield & Pricing",
    "🥗 Calorie & Nutrition Breakdown",
])

# ================= TAB 1: CALCULATOR =================
with tab1:
  st.caption(
      "Precise cost, batch yield, and margin analysis with automatic Master"
      " Price List integration."
  )

  col_up1, col_up2 = st.columns(2)
  with col_up1:
    uploaded_master_file = st.file_uploader(
        "1. Upload Master Recipe Collection (.xlsx)",
        type=["xlsx"],
        key="master_file",
    )
    if uploaded_master_file is not None:
      st.session_state.master_excel_data = uploaded_master_file

  with col_up2:
    uploaded_price_file = st.file_uploader(
        "2. Upload Bakery Price List (.xlsx)", type=["xlsx"], key="price_file"
    )
    if uploaded_price_file is not None:
      try:
        price_df = pd.read_excel(uploaded_price_file)
        price_map = {}
        for _, row in price_df.iterrows():
          desc = str(row.get("Description", "")).strip().lower()
          cost = float(row.get("Unit Cost", 0.0))
          if desc:
            price_map[desc] = cost
        st.session_state.price_lookup_dict = price_map
        st.success(
            f"Price list loaded successfully ({len(price_map)} items mapped)!"
        )
      except Exception as e:
        st.error(f"Error loading price list: {e}")

  if st.session_state.master_excel_data is not None:
    try:
      xls = pd.ExcelFile(st.session_state.master_excel_data)
      sheet_options = [s for s in xls.sheet_names if s != "Master Summary"]


      def on_recipe_select():
        chosen_sheet = st.session_state.recipe_dropdown
        recipe_df = pd.read_excel(
            st.session_state.master_excel_data,
            sheet_name=chosen_sheet,
            header=2,
        )
        recipe_df = recipe_df.dropna(subset=["Ingredient Name"])

        def convert_to_kg(row):
          qty = float(row["Base Qty"]) if pd.notnull(row["Base Qty"]) else 0.0
          unit = str(row["Unit"]).strip().lower()
          if unit in ["g", "gram", "grams", "ml"]:
            return qty / 1000.0
          return qty

        recipe_df["Quantity_KG"] = recipe_df.apply(convert_to_kg, axis=1)
        recipe_df["Ingredient"] = recipe_df["Ingredient Name"]

        rates = []
        for ing_name in recipe_df["Ingredient Name"]:
          clean_name = str(ing_name).strip().lower()
          matched_rate = 500.0

          if clean_name == "salt":
            for p_desc, p_cost in st.session_state.price_lookup_dict.items():
              if "tata salt" in p_desc or p_desc == "salt":
                matched_rate = p_cost
                break
          else:
            for p_desc, p_cost in st.session_state.price_lookup_dict.items():
              if clean_name in p_desc or p_desc in clean_name:
                matched_rate = p_cost
                break

          rates.append(matched_rate)

        recipe_df["Rate_Per_KG"] = rates

        st.session_state.ingredients = recipe_df[
            ["Ingredient", "Quantity_KG", "Rate_Per_KG"]
        ].reset_index(drop=True)
        st.session_state.recipe_title = chosen_sheet

      selected_sheet = st.selectbox(
          "Select Recipe Sheet to Load",
          sheet_options,
          key="recipe_dropdown",
          on_change=on_recipe_select,
      )
    except Exception as e:
      st.error(f"Error reading master file: {e}")

  col_left, col_right = st.columns([1.1, 0.9], gap="large")

  with col_left:
    st.subheader("1. Batch & Product Details")
    recipe_name = st.text_input(
        "Product / Recipe Name", value=st.session_state.recipe_title
    )
    st.session_state.recipe_title = recipe_name

    st.markdown("**Ingredients & Raw Material Rates:**")
    st.caption(
        "Click directly inside any cell to edit ingredient names, weight (KG),"
        " and rate per KG:"
    )

    edited_df = st.data_editor(
        st.session_state.ingredients,
        num_rows="dynamic",
        use_container_width=True,
        key="recipe_data_editor",
        column_config={
            "Ingredient": st.column_config.TextColumn("Ingredient", required=True),
            "Quantity_KG": st.column_config.NumberColumn(
                "Quantity (KG)", min_value=0.001, format="%.3f"
            ),
            "Rate_Per_KG": st.column_config.NumberColumn(
                "Rate / KG (₹)", min_value=0.0, format="₹%.2f"
            ),
        },
    )
    st.session_state.ingredients = edited_df

    st.subheader("2. Yield Loss & Overheads")
    c1, c2, c3 = st.columns(3)
    with c1:
      loss_percent = st.number_input(
          "Cooking / Moisture Loss (%)",
          min_value=0.0,
          max_value=90.0,
          value=10.0,
          step=0.5,
      )
    with c2:
      labor_gas_cost = st.number_input(
          "Labor + Fuel Cost (₹)", min_value=0.0, value=400.0, step=50.0
      )
    with c3:
      packaging_cost = st.number_input(
          "Packaging Cost (₹)", min_value=0.0, value=250.0, step=50.0
      )

    target_margin = st.slider(
        "Target Gross Margin (%)",
        min_value=5.0,
        max_value=80.0,
        value=35.0,
        step=1.0,
    )

  # Calculations
  clean_df = edited_df.dropna(subset=["Quantity_KG", "Rate_Per_KG"]).copy()
  raw_material_weight = clean_df["Quantity_KG"].sum()
  clean_df["Total Amount (₹)"] = (
      clean_df["Quantity_KG"] * clean_df["Rate_Per_KG"]
  )
  raw_material_cost = clean_df["Total Amount (₹)"].sum()

  final_yield_kg = raw_material_weight * (1 - (loss_percent / 100.0))
  total_batch_cost = raw_material_cost + labor_gas_cost + packaging_cost

  cost_per_kg = (
      (total_batch_cost / final_yield_kg) if final_yield_kg > 0 else 0.0
  )
  selling_price_per_kg = (
      (cost_per_kg / (1 - (target_margin / 100.0)))
      if target_margin < 100
      else 0.0
  )
  profit_per_kg = selling_price_per_kg - cost_per_kg

  with col_right:
    st.subheader("📋 Output & Cost Summary")

    st.markdown("**Ingredient-wise Total Cost Breakdown:**")
    display_summary_df = clean_df[
        ["Ingredient", "Quantity_KG", "Rate_Per_KG", "Total Amount (₹)"]
    ].copy()
    display_summary_df.columns = [
        "Ingredient",
        "Qty (KG)",
        "Rate/KG (₹)",
        "Total (₹)",
    ]
    st.dataframe(display_summary_df, use_container_width=True, hide_index=True)

    st.markdown(
        f"""
        <div style="background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 12px; padding: 20px; margin-bottom: 20px;">
            <p style="margin:0; font-size:14px; color:#64748B;">Selected Recipe: <b>{recipe_name}</b></p>
            <hr style="margin: 10px 0; border: 0; border-top: 1px solid #E2E8F0;">
            <p style="margin:0; font-size:14px; color:#64748B;">Total Raw Material Cost</p>
            <h3 style="margin:0 0 10px 0; color:#0F172A; font-size:24px;">₹{raw_material_cost:,.2f}</h3>
            <p style="margin:0; font-size:14px; color:#64748B;">Total Batch Cost (with Overheads)</p>
            <h2 style="margin:0 0 15px 0; color:#1E3A8A; font-size:30px;">₹{total_batch_cost:,.2f}</h2>
            <p style="margin:0; font-size:14px; color:#64748B;">Final Net Yield</p>
            <h3 style="margin:0 0 5px 0; color:#1E293B; font-size:24px;">{final_yield_kg:,.2f} KG</h3>
            <span style="color:#DC2626; font-size:13px; font-weight:600;">↓ {loss_percent}% Process Loss</span>
            <hr style="margin: 15px 0; border: 0; border-top: 1px solid #E2E8F0;">
            <div style="display: flex; justify-content: space-between;">
                <div>
                    <p style="margin:0; font-size:13px; color:#64748B;">Cost Per KG</p>
                    <h3 style="margin:0; color:#0F172A;">₹{cost_per_kg:,.2f}</h3>
                </div>
                <div>
                    <p style="margin:0; font-size:13px; color:#64748B;">Suggested Selling Price</p>
                    <h3 style="margin:0; color:#16A34A;">₹{selling_price_per_kg:,.2f}</h3>
                </div>
            </div>
            <p style="margin:12px 0 0 0; font-size:13px; color:#2563EB; font-weight:600;">
                💡 Net Profit: ₹{profit_per_kg:,.2f} per KG ({target_margin}% Margin)
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    excel_file_bytes = generate_professional_excel(
        recipe_name,
        clean_df,
        loss_percent,
        final_yield_kg,
        raw_material_cost,
        labor_gas_cost,
        packaging_cost,
        total_batch_cost,
        cost_per_kg,
        target_margin,
        selling_price_per_kg,
        profit_per_kg,
    )

    st.download_button(
        label=f"📥 Download Report for {recipe_name} (.xlsx)",
        data=excel_file_bytes,
        file_name=f"{recipe_name.replace(' ', '_')}_Costing_Sheet.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        use_container_width=True,
    )


# ================= TAB 2: CALORIE & NUTRITION =================
with tab2:
  st.subheader(f"🥗 Nutritional & Calorie Breakdown for: {recipe_name}")
  st.caption(
      "Estimated nutritional values per ingredient based on standard food"
      " composition data (per 100g basis)."
  )

  nutri_rows = []
  total_batch_kcal = 0
  total_batch_protein = 0
  total_batch_fat = 0
  total_batch_carbs = 0

  for _, row in clean_df.iterrows():
    ing = row["Ingredient"]
    qty_kg = row["Quantity_KG"]
    qty_g = qty_kg * 1000.0

    nutrients = get_nutrients(ing)
    factor = qty_g / 100.0
    kcal = nutrients["kcal"] * factor
    protein = nutrients["protein"] * factor
    fat = nutrients["fat"] * factor
    carbs = nutrients["carbs"] * factor

    total_batch_kcal += kcal
    total_batch_protein += protein
    total_batch_fat += fat
    total_batch_carbs += carbs

    nutri_rows.append({
        "Ingredient": ing,
        "Quantity (g)": qty_g,
        "Calories (kcal)": round(kcal, 1),
        "Protein (g)": round(protein, 1),
        "Fat (g)": round(fat, 1),
        "Carbs (g)": round(carbs, 1),
    })

  nutri_df = pd.DataFrame(nutri_rows)
  st.dataframe(nutri_df, use_container_width=True)

  final_yield_g = final_yield_kg * 1000.0 if final_yield_kg > 0 else 1.0
  per_100g_factor = 100.0 / final_yield_g

  yield_kcal_100g = total_batch_kcal * per_100g_factor
  yield_protein_100g = total_batch_protein * per_100g_factor
  yield_fat_100g = total_batch_fat * per_100g_factor
  yield_carbs_100g = total_batch_carbs * per_100g_factor

  st.markdown(f"### 📊 Summary Per 100g Finished Product ({recipe_name})")
  col_n1, col_n2, col_n3, col_n4 = st.columns(4)
  with col_n1:
    st.metric(label="Calories (Per 100g)", value=f"{yield_kcal_100g:.1f} kcal")
  with col_n2:
    st.metric(label="Protein", value=f"{yield_protein_100g:.1f} g")
  with col_n3:
    st.metric(label="Total Fat", value=f"{yield_fat_100g:.1f} g")
  with col_n4:
    st.metric(label="Carbohydrates", value=f"{yield_carbs_100g:.1f} g")
