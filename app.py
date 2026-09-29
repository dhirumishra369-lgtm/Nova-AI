def on_recipe_select():
        chosen_sheet = st.session_state.recipe_dropdown
        recipe_df = pd.read_excel(
            st.session_state.master_excel_data,
            sheet_name=chosen_sheet,
            header=2,
        )

        # Extract Total Base Weight if mentioned in the sheet
        total_weight_row = recipe_df[
            recipe_df["Ingredient Name"]
            .astype(str)
            .str.contains("TOTAL BASE INGREDIENT WEIGHT", case=False, na=False)
        ]
        excel_total_weight = None
        if not total_weight_row.empty:
          excel_total_weight = float(total_weight_row["Base Qty"].values[0])

        # Drop non-ingredient rows
        recipe_df = recipe_df.dropna(subset=["Ingredient Name"])
        recipe_df = recipe_df[
            ~recipe_df["Ingredient Name"]
            .astype(str)
            .str.contains("TOTAL", case=False, na=False)
        ]

        def convert_to_kg(row):
          qty = float(row["Base Qty"]) if pd.notnull(row["Base Qty"]) else 0.0
          unit = str(row["Unit"]).strip().lower()
          if unit in ["g", "gram", "grams", "ml"]:
            return qty / 1000.0
          return qty

        recipe_df["Quantity_KG"] = recipe_df.apply(convert_to_kg, axis=1)
        recipe_df["Ingredient"] = recipe_df["Ingredient Name"]

        # If Excel has a defined total weight row, ensure sum matches or store it
        rates = []
        for ing_name in recipe_df["Ingredient Name"]:
          clean_name = str(ing_name).strip().lower()
          matched_rate = 500.0

          if clean_name == "salt":
            for p_desc, p_cost in st.session_state.price_lookup_dict.items():
              if "tata salt" in p_desc or p_desc == "salt":
                matched_rate = p_cost
                break
          elif "badam" in clean_name or "almond" in clean_name:
            for p_desc, p_cost in st.session_state.price_lookup_dict.items():
              if "almond factory" in p_desc or "badam factory" in p_desc:
                matched_rate = p_cost
                break
            if matched_rate == 500.0:
              for p_desc, p_cost in st.session_state.price_lookup_dict.items():
                if "almond" in p_desc or "badam" in p_desc:
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
