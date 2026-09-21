import streamlit as st
import pandas as pd

# Maan lijiye aapke paas ingredients ka dataframe 'df' hai
search_term = st.text_input("🔍 Search Ingredient ya Recipe:")

if search_term:
    # Yeh code user ke type kiye hue text ke hisaab se data filter kar dega
    filtered_df = df[df['Ingredient'].str.contains(search_term, case=False, na=False)]
    st.dataframe(filtered_df)
else:
    st.dataframe(df)
