uploaded_ingredients_file = st.file_uploader(
        "Upload ingredient sheet", type=["csv", "xlsx"], key="ing_file"
    )
    if uploaded_ingredients_file is not None:
      try:
        if uploaded_ingredients_file.name.endswith(".csv"):
          temp_df = pd.read_csv(uploaded_ingredients_file)
        else:
          temp_df = pd.read_excel(uploaded_ingredients_file)

        # Clean column names (remove extra spaces, lowercase)
        temp_df.columns = temp_df.columns.str.strip()

        # Map common alternate names automatically
        rename_map = {}
        for col in temp_df.columns:
          col_lower = col.lower()
          if "ingredient" in col_lower or "item" in col_lower or "name" in col_lower:
            rename_map[col] = "Ingredient"
          elif "quantity" in col_lower or "qty" in col_lower or "weight" in col_lower or "wt" in col_lower:
            rename_map[col] = "Quantity_KG"
          elif "rate" in col_lower or "price" in col_lower or "cost" in col_lower or "per" in col_lower:
            rename_map[col] = "Rate_Per_KG"

        temp_df = temp_df.rename(columns=rename_map)

        if all(
            col in temp_df.columns
            for col in ["Ingredient", "Quantity_KG", "Rate_Per_KG"]
        ):
          st.session_state.ingredients = temp_df[
              ["Ingredient", "Quantity_KG", "Rate_Per_KG"]
          ].dropna(how="all")
          st.success("Ingredients imported successfully!")
          st.rerun()
        else:
          st.error(
              "Could not automatically detect columns. Please ensure your file"
              " has columns for Ingredient, Quantity, and Rate."
          )
      except Exception as e:
        st.error(f"Error loading file: {e}")
