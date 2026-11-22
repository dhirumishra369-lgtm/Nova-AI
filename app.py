def load_recipe_from_excel(file_obj, sheet_name, price_lookup):
  file_obj.seek(0)
  # Pehle bina header ke padh kar 'Ingredient' ya 'Item' wali header row dhoondein
  raw_df = pd.read_excel(file_obj, sheet_name=sheet_name, header=None)

  header_row_idx = None
  ing_col_name = None
  qty_col_name = None
  unit_col_name = None

  for r_idx, row in raw_df.iterrows():
    row_vals = [str(v).strip() for v in row.tolist() if pd.notnull(v)]
    lower_map = {str(v).strip().lower(): str(v).strip() for v in row.tolist() if pd.notnull(v)}
    if any("ingredient" in k or "item" in k for k in lower_map):
      header_row_idx = r_idx
      for k, orig in lower_map.items():
        if "ingredient" in k or ("item" in k and not ing_col_name):
          ing_col_name = orig
        if "qty" in k or "quantity" in k or "base" in k:
          qty_col_name = orig
        if "unit" in k or "uom" in k:
          unit_col_name = orig
      break

  file_obj.seek(0)
  if header_row_idx is not None and ing_col_name:
    recipe_df = pd.read_excel(file_obj, sheet_name=sheet_name, header=header_row_idx)
  else:
    recipe_df = pd.read_excel(file_obj, sheet_name=sheet_name, header=2)
    ing_col_name = "Ingredient Name" if "Ingredient Name" in recipe_df.columns else recipe_df.columns[0]

  recipe_df = recipe_df.dropna(how="all")
  if ing_col_name in recipe_df.columns:
    recipe_df = recipe_df.dropna(subset=[ing_col_name])
  else:
    return pd.DataFrame(columns=["Ingredient", "Quantity_KG", "Rate_Per_KG"])

  # Total / summary rows hataayein agar sheet ke end mein hon
  recipe_df = recipe_df[
      ~recipe_df[ing_col_name]
      .astype(str)
      .str.strip()
      .str.lower()
      .isin(["", "nan", "total", "subtotal", "grand total"])
  ]

  def convert_to_kg(row):
    raw_qty = row[qty_col_name] if (qty_col_name and qty_col_name in row) else 0.0
    try:
      qty = float(raw_qty) if pd.notnull(raw_qty) else 0.0
    except (ValueError, TypeError):
      qty = 0.0
    unit = (
        str(row[unit_col_name]).strip().lower()
        if (unit_col_name and unit_col_name in row and pd.notnull(row[unit_col_name]))
        else "kg"
    )
    if unit in ["g", "gm", "gram", "grams", "ml"]:
      return qty / 1000.0
    return qty

  recipe_df["Quantity_KG"] = recipe_df.apply(convert_to_kg, axis=1)
  recipe_df["Ingredient"] = recipe_df[ing_col_name].astype(str).str.strip()
  recipe_df["Rate_Per_KG"] = [
      match_rate_from_lookup(name, price_lookup)
      for name in recipe_df["Ingredient"]
  ]
  return recipe_df[["Ingredient", "Quantity_KG", "Rate_Per_KG"]].reset_index(
      drop=True
  )
