import streamlit as st
from google import genai  # new SDK: pip install google-genai

st.set_page_config(page_title="Fast AI Assistant", page_icon="⚡", layout="centered")

st.title("⚡ Fast AI Assistant & Search")
st.caption("High-Performance Assistant | Direct Access Mode")

MODEL_NAME = "gemini-2.5-flash"


@st.cache_resource
def get_client():
    # API key comes from secrets, so users never see an API key box
    api_key = st.secrets.get("GEMINI_API_KEY")
    if not api_key:
        st.error("GEMINI_API_KEY is missing. Add it to .streamlit/secrets.toml or your app's Secrets settings.")
        st.stop()
    return genai.Client(api_key=api_key)


# User query box (no password, no API key box)
user_prompt = st.text_area("✍️ Apana sawal ya search query yahan likhein:")

if st.button("🚀 Run Fast AI", type="primary"):
    if not user_prompt.strip():
        st.warning("Kripya koi sawal ya prompt darj karein.")
    else:
        try:
            client = get_client()
            with st.spinner("AI is thinking... ⚡"):
                response = client.models.generate_content(
                    model=MODEL_NAME,
                    contents=user_prompt,
                )
            st.success("✅ Result:")
            st.write(response.text or "No text was returned. Try rephrasing the prompt.")
        except Exception as e:
            st.error(f"Error: {e}")
