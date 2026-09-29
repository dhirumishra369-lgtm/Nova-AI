# File Uploader for Ingredients with Index-based fallback
  with st.expander("📁 Import Ingredients from Excel / CSV File"):
    uploaded_ingredients_file = st.file_uploader(
        "Upload ingredient sheet", type=["csv", "xlsx"], key="ing_file"
    )
    if uploaded_ingredients_file is not None:
      try:
        if uploaded_ingredients_file.name.endswith(".csv"):
          temp_df = pd.read_csv(uploaded_ingredients_file)
        else:
          temp_df = pd.read_excel(uploaded_ingredients_file)

        temp_df.columns = temp_df.columns.str.strip()
        rename_map = {}
        for col in temp_df.columns:
          col_lower = str(col).lower()
          if (
              "ingredient" in col_lower
              or "item" in col_lower
              or "name" in col_lower
              or "raw" in col_lower
          ):
            rename_map[col] = "Ingredient"
          elif (
              "quantity" in col_lower
              or "qty" in col_lower
              or "weight" in col_lower
              or "wt" in col_lower
              or "amount" in col_lower
          ):
            rename_map[col] = "Quantity_KG"
          elif (
              "rate" in col_lower
              or "price" in col_lower
              or "cost" in col_lower
              or "per" in col_lower
              or "rs" in col_lower
          ):
            rename_map[col] = "Rate_Per_KG"

        temp_df = temp_df.rename(columns=rename_map)

        # Fallback: Agar naam match na ho, toh column position (0, 1, 2) ke hisaab se utha lo
        if "Ingredient" not in temp_df.columns and len(temp_df.columns) > 0:
          temp_df = temp_df.rename(columns={temp_df.columns[0]: "Ingredient"})
        if "Quantity_KG" not in temp_df.columns and len(temp_df.columns) > 1:
          temp_df = temp_df.rename(columns={temp_df.columns[1]: "Quantity_KG"})
        if "Rate_Per_KG" not in temp_df.columns and len(temp_df.columns) > 2:
          temp_df = temp_df.rename(columns={temp_df.columns[2]: "Rate_Per_KG"})

        if all(
            col in temp_df.columns
            for col in ["Ingredient", "Quantity_KG", "Rate_Per_KG"]
        ):
          st.session_state.ingredients = temp_df[
              ["Ingredient", "Quantity_KG", "Rate_Per_KG"]
          ].dropna(how="all")
          st.success(
              "Ingredients imported successfully! Scroll down to see updated"
              " table."
          )
          st.rerun()
        else:
          st.error(
              "Could not read columns. Please ensure your file has at least 3"
              " columns (Item Name, Quantity, Rate)."
          )
      except Exception as e:
        st.error(f"Error loading file: {e}")
