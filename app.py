import streamlit as st
import google.generativeai as genai

st.set_page_config(page_title="Fast AI App", page_icon="⚡", layout="centered")

st.title("⚡ Fast AI Assistant & Search")
st.caption("Ekdum fast aur simple AI app—apni API Key daalein aur kuch bhi poochhein!")

# Password Security (Optional)
password = st.text_input("🔒 Enter Password", type="password")
if password != "Dhirendra@123":
    st.warning("Please enter the correct password to access the app.")
    st.stop()

# API Key Input
api_key = st.text_input("🔑 Enter Gemini API Key", type="password")

# User Query / Search Box
user_prompt = st.text_area("✍️ Type your question or search query here:")

if st.button("🚀 Run Fast AI", type="primary"):
    if not api_key:
        st.error("Please enter your Gemini API Key.")
    elif not user_prompt:
        st.warning("Please enter a question or prompt.")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel("gemini-1.5-flash")
            
            with st.spinner("AI is thinking... ⚡"):
                response = model.generate_content(user_prompt)
                
            st.success("✅ Result:")
            st.write(response.text)
        except Exception as e:
            st.error(f"Error: {e}")
