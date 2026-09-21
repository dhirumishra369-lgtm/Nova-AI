with tab2:
  st.subheader("🤖 AI Chef & Production Consultant")
  st.caption(
      "Ask questions about recipe optimization, shelf-life improvement, process"
      " loss reduction, or packaging tips."
  )

  user_query = st.text_area(
      "Your Question",
      placeholder=(
          "e.g., How can I reduce moisture loss below 12% in cashew fudge? Or"
          " how do I extend shelf-life without chemical preservatives?"
      ),
  )

  if st.button("Ask AI Consultant", type="primary"):
    if not user_query:
      st.warning("Please type a question.")
    else:
      try:
        # Yeh automatic Google ki tarah background se key utha lega
        genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
        model = genai.GenerativeModel("gemini-1.5-flash")

        prompt = f"""
                You are an expert commercial food technologist, chef, and bakery production consultant.
                Current Product: {recipe_name}
                Raw Material Batch Weight: {raw_material_weight:.2f} kg
                Process Loss: {loss_percent}%
                Final Output Yield: {final_yield_kg:.2f} kg
                Cost Per KG: ₹{cost_per_kg:.2f}

                User Query: {user_query}
                Please provide practical, accurate, and scientifically backed commercial kitchen guidance.
                """
        with st.spinner("AI is analyzing your recipe..."):
          response = model.generate_content(prompt)
          st.success("Consultant Recommendation:")
          st.write(response.text)
      except Exception as e:
        st.error(
            f"Configuration Error: Please ensure GOOGLE_API_KEY is set in your"
            f" Streamlit secrets. Details: {e}"
        )
