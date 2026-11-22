import datetime
import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Nova-AI | Recipe Costing & Yield",
    page_icon="🍰",
    layout="wide",
)

# ----------------- SESSION STATE -----------------
if "ingredients" not in st.session_state:
  st.session_state.ingredients = pd.DataFrame([
      {"Ingredient": "Kaju (Cashew)", "Quantity_KG": 5.0, "Rate_Per_KG": 680.0},
      {"Ingredient": "Sugar", "Quantity_KG": 8.0, "Rate_Per_KG": 42.0},
      {"Ingredient": "Silver Vark", "Quantity_KG": 0.05, "Rate_Per_KG": 4000.0},
      {"Ingredient": "Cardamom / Ghee", "Quantity_KG": 0.2, "Rate_Per_KG": 600.0},
  ])

if "recipe_title" not in st.session_state:
  st.session_state.recipe_title = "Premium Kaju Katli"

if "master_excel_files" not in st.session_state:
  st.session_state.master_excel_files = []

if "price_lookup_dict" not in st.session_state:
  st.session_state.price_lookup_dict = {}

# ----------------- NUTRITION DB -----------------
NUTRIENTS_DB = {
    "kaju": {"kcal": 553, "protein": 18.2, "fat": 43.8, "carbs": 30.2},
    "cashew": {"kcal": 553, "protein": 18.2, "fat": 43.8, "carbs": 30.2},
    "sugar": {"kcal": 387, "protein": 0.0, "fat": 0.0, "carbs": 100.0},
    "ghee": {"kcal": 900, "protein": 0.0, "fat": 100.0, "carbs": 0.0},
    "butter": {"kcal": 717, "protein": 0.85, "fat": 81.0, "carbs": 0.06},
    "milk": {"kcal": 42, "protein": 3.4, "fat": 1.0, "carbs": 5.0},
    "pistachio": {"kcal": 562, "protein": 20.0, "fat": 45.0, "carbs": 28.0},
    "almond": {"kcal": 579, "protein": 21.1, "fat": 49.9, "carbs": 21.6},
    "badam": {"kcal": 579, "protein": 21.1, "fat": 49.9, "carbs": 21.6},
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


def match_rate_from_lookup(ing_name, price_lookup):
  clean_name = str(ing_name).strip().lower()
  matched_rate = 500.0
  if not price_lookup:
    return matched_rate

  if clean_name == "salt":
    for p_desc, p_cost in price_lookup.items():
      if "tata salt" in p_desc or p_desc == "salt":
        return p_cost
  elif "badam" in clean_name or "almond" in clean_name:
    for p_desc, p_cost in price_lookup.items():
      if "almond factory" in p_desc or "badam factory" in p_desc:
        return p_cost
    for p_desc, p_cost in price_lookup.items():
      if "almond" in p_desc or "badam" in p_desc:
        return p_cost
  else:
    for p_desc, p_cost in price_lookup.items():
      if clean_name in p_desc or p_desc in clean_name:
        return p_cost
  return matched_rate


def load_recipe_from_excel(file_obj, sheet_name, price_lookup):
  file_obj.seek(0)
  recipe_df = pd.read_excel(file_obj, sheet_name=sheet_name, header=2)
  recipe_df = recipe_df.dropna(subset=["Ingredient Name"])

  def convert_to_kg(row):
    qty = float(row["Base Qty"]) if pd.notnull(row["Base Qty"]) else 0.0
    unit = str(row["Unit"]).strip().lower()
    if unit in ["g", "gram", "grams", "ml"]:
      return qty / 1000.0
    return qty

  recipe_df["Quantity_KG"] = recipe_df.apply(convert_to_kg, axis=1)
  recipe_df["Ingredient"] = recipe_df["Ingredient Name"]
  recipe_df["Rate_Per_KG"] = [
      match_rate_from_lookup(name, price_lookup)
      for name in recipe_df["Ingredient Name"]
  ]
  return recipe_df[["Ingredient", "Quantity_KG", "Rate_Per_KG"]].reset_index(
      drop=True
  )


# ----------------- EXCEL GENERATORS -----------------
def generate_recipe_excel(
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

  ws.merge_cells("A1:E1")
  title = ws["A1"]
  title.value = "RECIPE COSTING & BATCH YIELD REPORT"
  title.font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
  title.fill = PatternFill(
      start_color="1E3A8A", end_color="1E3A8A", fill_type="solid"
  )
  title.alignment = Alignment(horizontal="center", vertical="center")
  ws.row_dimensions[1].height = 35

  ws["A2"] = "Product / Recipe:"
  ws["B2"] = recipe_name
  ws["D2"] = "Date:"
  ws["E2"] = datetime.date.today().strftime("%d-%b-%Y")
  ws.row_dimensions[2].height = 22

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
    c3.number_format = "#,##0.000"
    c3.alignment = Alignment(horizontal="right")

    c4 = ws.cell(row=curr, column=4, value=float(row["Rate_Per_KG"]))
    c4.number_format = "₹#,##0.00"
    c4.alignment = Alignment(horizontal="right")

    c5 = ws.cell(row=curr, column=5, value=f"=C{curr}*D{curr}")
    c5.number_format = "₹#,##0.00"
    c5.alignment = Alignment(horizontal="right")

    for c_idx in range(1, 6):
      ws.cell(row=curr, column=c_idx).border = thin_border

  tot_r = start_row + len(df)
  ws.cell(row=tot_r, column=2, value="Subtotal (Raw Materials)").font = Font(
      bold=True
  )
  c_tot_q = ws.cell(
      row=tot_r, column=3, value=f"=SUM(C{start_row}:C{tot_r-1})"
  )
  c_tot_q.font = Font(bold=True)
  c_tot_q.number_format = "#,##0.000"

  c_tot_a = ws.cell(
      row=tot_r, column=5, value=f"=SUM(E{start_row}:E{tot_r-1})"
  )
  c_tot_a.font = Font(bold=True)
  c_tot_a.number_format = "₹#,##0.00"

  sum_start = tot_r + 2
  metrics = [
      ("Total Raw Material Batch Weight", f"=C{tot_r}", "KG", "#,##0.000"),
      ("Process / Moisture Loss (%)", loss_pct / 100.0, "%", "0.0%"),
      (
          "Final Net Yield Output",
          f"=E{sum_start+1}*(1-E{sum_start+2})",
          "KG",
          "#,##0.000",
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
          f"=IF(E{sum_start+3}>0,E{sum_start+7}/E{sum_start+3},0)",
          "INR/KG",
          "₹#,##0.00",
      ),
      ("Target Profit Margin (%)", margin_pct / 100.0, "%", "0.0%"),
      (
          "Suggested Selling Price Per KG",
          f"=IF(E{sum_start+9}<1,E{sum_start+8}/(1-E{sum_start+9}),0)",
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
    ws.merge_cells(f"A{r}:C{r}")
    ws[f"A{r}"] = label
    ws[f"D{r}"] = unit
    v = ws[f"E{r}"]
    v.value = val
    v.number_format = num_fmt

  for col_letter, width in [
      ("A", 8),
      ("B", 30),
      ("C", 16),
      ("D", 18),
      ("E", 22),
  ]:
    ws.column_dimensions[col_letter].width = width

  out = io.BytesIO()
  wb.save(out)
  return out.getvalue()


def generate_monthly_master_template():
  wb = openpyxl.Workbook()
  wb.remove(wb.active)
  header_fill = PatternFill(
      start_color="1E3A8A", end_color="1E3A8A", fill_type="solid"
  )
  header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
  title_font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")

  sheets_spec = [
      (
          "Daily RM Issue",
          "DAILY RAW MATERIAL ISSUE & CONSUMPTION LOG",
          [
              "Date",
              "Item Code",
              "Item Description",
              "Qty Issued",
              "Unit",
              "Total Cost (₹)",
              "Slip / Photo Link",
          ],
          [
              "2026-10-01",
              "RM0005",
              "ALMOND FACTORY",
              25.0,
              "KGS",
              20846.75,
              "Attach Slip Here",
          ],
      ),
      (
          "Daily Production & Dispatch",
          "DAILY FINISHED GOODS PRODUCTION & DISPATCH LOG",
          [
              "Date",
              "Item Code",
              "Item Description",
              "Produced Qty",
              "Dispatched Qty",
              "Total Cost (₹)",
              "Total MRP (₹)",
              "Production Slip Photo",
          ],
          [
              "2026-10-01",
              "FGSW0037",
              "MANGO KAJU CAKE",
              50.0,
              45.0,
              649.39,
              1000.00,
              "Attach Photo Here",
          ],
      ),
      (
          "Master_30_Day_Report",
          "30-DAY MONTHLY MASTER CONSOLIDATED REPORT",
          [
              "Item Code / Description",
              "Total Issued Qty (RM)",
              "Total Produced Qty (FG)",
              "Total Dispatched Qty (FG)",
              "Total Cost (₹)",
              "Total MRP (₹)",
          ],
          ["MANGO KAJU CAKE", "-", 50.0, 45.0, 649.39, 1000.00],
      ),
  ]

  for title_tab, heading_txt, hdrs, sample_row in sheets_spec:
    ws = wb.create_sheet(title=title_tab)
    ws.views.sheetView[0].showGridLines = True
    ws.append([heading_txt])
    ws["A1"].font = title_font
    ws.append(hdrs)
    for col_num in range(1, len(hdrs) + 1):
      cell = ws.cell(row=2, column=col_num)
      cell.fill = header_fill
      cell.font = header_font
      cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.append(sample_row)

  for sheet in wb.worksheets:
    for col in sheet.columns:
      max_len = max(len(str(cell.value or "")) for cell in col)
      col_letter = get_column_letter(col[0].column)
      sheet.column_dimensions[col_letter].width = max(max_len + 4, 15)

  out = io.BytesIO()
  wb.save(out)
  return out.getvalue()


def build_merged_excel_bytes(merged_df, sheet_tabs_dict):
  wb = openpyxl.Workbook()
  ws_all = wb.active
  ws_all.title = "Merged_All_Data"
  ws_all.views.sheetView[0].showGridLines = True

  hdr_fill = PatternFill(
      start_color="1E3A8A", end_color="1E3A8A", fill_type="solid"
  )
  hdr_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

  if not merged_df.empty:
    ws_all.append(list(merged_df.columns))
    for col_i in range(1, len(merged_df.columns) + 1):
      cell = ws_all.cell(row=1, column=col_i)
      cell.fill = hdr_fill
      cell.font = hdr_font
    for _, r_data in merged_df.iterrows():
      ws_all.append([None if pd.isna(v) else v for v in r_data.tolist()])

  for tab_title, df_part in sheet_tabs_dict.items():
    safe_title = re.sub(r"[\[\]:*?/\\]", "_", str(tab_title))[:31] or "Sheet"
    if safe_title in wb.sheetnames:
      safe_title = f"{safe_title[:26]}_{len(wb.sheetnames)}"
    ws_p = wb.create_sheet(title=safe_title)
    ws_p.views.sheetView[0].showGridLines = True
    if not df_part.empty:
      ws_p.append(list(df_part.columns))
      for col_i in range(1, len(df_part.columns) + 1):
        c = ws_p.cell(row=1, column=col_i)
        c.fill = hdr_fill
        c.font = hdr_font
      for _, r_data in df_part.iterrows():
        ws_p.append([None if pd.isna(v) else v for v in r_data.tolist()])

  for sheet in wb.worksheets:
    for col in sheet.columns:
      max_len = max(
          (len(str(cell.value or "")) for cell in list(col)[:50]), default=12
      )
      col_letter = get_column_letter(col[0].column)
      sheet.column_dimensions[col_letter].width = min(max(max_len + 4, 14), 45)

  out = io.BytesIO()
  wb.save(out)
  return out.getvalue()


# ----------------- TABS -----------------
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Recipe Costing & Yield",
    "🥗 Calorie & Nutrition",
    "📈 30-Day Master Ledger & Templates",
    "📦 Multi-Excel Batch Upload & Merge",
])

with tab1:
  col_up1, col_up2 = st.columns(2)
  with col_up1:
    uploaded_master_files = st.file_uploader(
        "1. Upload Master Recipe Files (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="master_files_multi",
    )
    if uploaded_master_files:
      st.session_state.master_excel_files = uploaded_master_files

  with col_up2:
    uploaded_price_files = st.file_uploader(
        "2. Upload Bakery Price Lists (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="price_files_multi",
    )
    if uploaded_price_files:
      try:
        price_map = {}
        for pf in uploaded_price_files:
          pf.seek(0)
          price_df = pd.read_excel(pf)
          for _, row in price_df.iterrows():
            desc = str(row.get("Description", "")).strip().lower()
            cost = float(row.get("Unit Cost", 0.0) or 0.0)
            if desc:
              price_map[desc] = cost
        st.session_state.price_lookup_dict = price_map
        st.success(f"Loaded {len(price_map)} price items.")
      except Exception as e:
        st.error(f"Price list error: {e}")

  if st.session_state.master_excel_files:
    try:
      recipe_choices = {}
      for mf in st.session_state.master_excel_files:
        mf.seek(0)
        xls = pd.ExcelFile(mf)
        for s in xls.sheet_names:
          if s != "Master Summary":
            display_key = (
                f"{s}"
                if len(st.session_state.master_excel_files) == 1
                else f"[{mf.name}] {s}"
            )
            recipe_choices[display_key] = (mf, s)

      def on_recipe_select():
        chosen_key = st.session_state.recipe_dropdown
        mf_obj, chosen_sheet = recipe_choices[chosen_key]
        st.session_state.ingredients = load_recipe_from_excel(
            mf_obj, chosen_sheet, st.session_state.price_lookup_dict
        )
        st.session_state.recipe_title = chosen_sheet

      st.selectbox(
          "Select Recipe Sheet",
          list(recipe_choices.keys()),
          key="recipe_dropdown",
          on_change=on_recipe_select,
      )
    except Exception as e:
      st.error(f"Master file error: {e}")

  col_left, col_right = st.columns([1.1, 0.9], gap="large")
  with col_left:
    recipe_name = st.text_input(
        "Product / Recipe Name", value=st.session_state.recipe_title
    )
    st.session_state.recipe_title = recipe_name

    edited_df = st.data_editor(
        st.session_state.ingredients,
        num_rows="dynamic",
        use_container_width=True,
        key="recipe_data_editor",
    )
    st.session_state.ingredients = edited_df

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
