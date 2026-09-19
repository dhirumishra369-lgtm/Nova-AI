import streamlit as st
import google.generativeai as genai

st.set_page_config(page_title="Nova-AI", page_icon="⚡", layout="centered")

st.title("⚡ Nova-AI Assistant & Search")
st.caption("High-Performance Assistant | Direct Access Mode")

# User Query Box (No Password, No API Key Box)
user_prompt = st.text_area("✍️ Apana sawal ya search query yahan likhein:")

if st.button("🚀 Run Nova-AI", type="primary"):
    if not user_prompt:
        st.warning("Kripya koi sawal ya prompt darj karein.")
    else:
        try:
            # Streamlit Secrets se API key automatically configure karne ke liye
            if "GEMINI_API_KEY" in st.secrets:
                genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
            else:
                st.error("⚠️ Error: Streamlit Cloud ke 'Secrets' mein GEMINI_API_KEY set nahi hai!")
                st.stop()

            model = genai.GenerativeModel("gemini-1.5-flash")
            
            with st.spinner("AI is thinking... ⚡"):
                response = model.generate_content(user_prompt)
                
            st.success("✅ Result:")
            st.write(response.text)
        except Exception as e:
            st.error(f"Error: {e}")
