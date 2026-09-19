import streamlit as st
import google.generativeai as genai

st.set_page_config(page_title="Nova-AI", page_icon="⚡", layout="centered")

st.title("⚡ Nova-AI Assistant & Search")
st.caption("High-Performance Assistant | Direct Access Mode")

# Direct API Key input box so you never face secrets error again
api_key = st.text_input("🔑 Gemini API Key darj karein:", type="password")

user_prompt = st.text_area("✍️ Apana sawal ya search query yahan likhein:")

if st.button("🚀 Run Nova-AI", type="primary"):
    if not api_key:
        st.warning("Kripya pehle apni Gemini API Key darj karein.")
    elif not user_prompt:
        st.warning("Kripya koi sawal ya prompt darj karein.")
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
