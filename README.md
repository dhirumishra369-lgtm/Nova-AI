import streamlit as st
import google.generativeai as genai

st.set_page_config(page_title="Fast AI Assistant", page_icon="⚡", layout="centered")

st.title("⚡ Fast AI Assistant & Search")
st.caption("High-Performance Assistant | Direct Access Mode")

# User Query / Search Box (No Password, No API Key Required)
user_prompt = st.text_area("✍️ Apana sawal ya search query yahan likhein:")

if st.button("🚀 Run Fast AI", type="primary"):
    if not user_prompt:
        st.warning("Kripya koi sawal ya prompt darj karein.")
    else:
        try:
            # Direct response without manual key entry
            model = genai.GenerativeModel("gemini-1.5-flash")
            
            with st.spinner("AI is thinking... ⚡"):
                response = model.generate_content(user_prompt)
                
            st.success("✅ Result:")
            st.write(response.text)
        except Exception as e:
            st.error(f"Error: {e}. (Kripya sunishchit karein ki environment mein valid key set ho.)")
