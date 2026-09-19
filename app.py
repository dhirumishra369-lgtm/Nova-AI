import streamlit as st
import google.generativeai as genai

st.set_page_config(page_title="Nova-AI Production & Costing", page_icon="⚡", layout="wide")

st.title("⚡ Nova-AI: Recipe Costing, Yield & Production Assistant")
st.caption("Commercial Kitchen, Confectionery & Bakery Manufacturing Intelligence")

# Initialize Session State for API Key
if "api_key" not in st.session_state:
    st.session_state.api_key = ""

st.session_state.api_key = st.text_input(
    "🔑 Yahan apni Gemini API Key darj karein:", 
    type="password", 
    value=st.session_state.api_key
)

tab1, tab2 = st.tabs(["📊 Recipe Costing & Yield Calculator", "🤖 AI Chef & Production Assistant"])

with tab1:
    st.subheader("1. Batch & Product Details")
    recipe_name = st.text_input("Product / Recipe Name", "Special Bakery Item")
    
    col1, col2 = st.columns(2)
    with col1:
        loss_pct = st.number_input("Cooking / Moisture Loss (%)", value=12.0)
    with col2:
        target_margin = st.number_input("Target Gross Margin (%)", value=35.0)

with tab2:
    st.subheader("🤖 AI Production Consultant & Search")
    user_prompt = st.text_area("✍️ Apana sawal, recipe ya production query yahan likhein:")

    if st.button("🚀 Run AI Analysis", type="primary"):
        if not st.session_state.api_key:
            st.warning("⚠️ Kripya sabse upar diye gaye box mein apni Gemini API Key darj karein.")
        elif not user_prompt:
            st.warning("⚠️ Kripya koi sawal ya prompt darj karein.")
        else:
            try:
                genai.configure(api_key=st.session_state.api_key)
                model = genai.GenerativeModel("gemini-1.5-flash")
                
                with st.spinner("AI Chef is analyzing... ⚡"):
                    response = model.generate_content(user_prompt)
                    
                st.success("✅ AI Result:")
                st.write(response.text)
            except Exception as e:
                st.error(f"Error: {e}")
