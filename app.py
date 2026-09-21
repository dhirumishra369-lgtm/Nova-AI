import streamlit as st
import pandas as pd

st.title("🔍 Mera Custom Search Engine")

# 1. Data load karein (aap apni Excel ya CSV file yahan daal sakte hain)
# df = pd.read_excel('apni_file.xlsx')

# Dummy data example ke liye:
data = {'Recipe': ['Kaju Katli', 'Gulab Jamun', 'Rasgulla', 'Jodhpuri Dana'],
        'Cost': [500, 300, 250, 400]}
df = pd.DataFrame(data)

# 2. User se search input lena
query = st.text_input("Kuch bhi search karein...")

# 3. Search logic aur results dikhana
if query:
    # Yeh code query ko match karega (case-insensitive)
    results = df[df['Recipe'].str.contains(query, case=False, na=False)]
    
    if not results.empty:
        st.success(f"{len(results)} result(s) mile:")
        st.dataframe(results)
    else:
        st.warning("Koi matching data nahi mila.")
else:
    st.info("Search box mein kuch type karein.")
