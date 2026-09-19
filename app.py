import streamlit as st
import google.generativeai as genai

st.set_page_config(page_title="Fast AI Assistant", page_icon="⚡", layout="centered")

st.title("⚡ Fast AI Assistant & Search")
st.caption("High-Performance Assistant | Direct Access Mode")

api_key = st.text_input("🔑 Gemini API Key darj karein:", type="password")
user_prompt = st.text_area("✍️ Apana sawal ya search query yahan likhein:")

if st.button("🚀 Run Fast AI", type="primary"):
    if not api_key:
        st.warning("Kripya pehle apni Gemini API Key darj karein.")
    elif not user_prompt:
        st.warning("Kripya koi sawal ya prompt darj karein.")
    else:
        try:
            genai.configure(api_key=api_key)
            
            # Automatically find a working model from your account
            available_model = None
            for m in genai.list_models():
                if 'generateContent' in m.supported_generation_methods:
                    available_model = m.name
                    break
            
            if not available_model:
                available_model = "gemini-1.5-flash"
                
            model = genai.GenerativeModel(available_model)
            
            with st.spinner(f"AI is thinking using {available_model}... ⚡"):
                response = model.generate_content(user_prompt)
                
            st.success("✅ Result:")
            st.write(response.text)
        except Exception as e:
            st.error(f"Error aaya: {e}")
