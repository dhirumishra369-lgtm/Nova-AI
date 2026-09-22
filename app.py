import streamlit as st
import google.generativeai as genai

st.set_page_config(
    page_title="AI Custom Search Engine",
    page_icon="🔍",
    layout="wide"
)

# ----------------- SESSION STATE SETUP -----------------
if "search_history" not in st.session_state:
    st.session_state.search_history = []

# ----------------- UI HEADER -----------------
st.title("🔍 AI-Powered Custom Search Engine")
st.caption("Type any topic, question, or keyword below to get instant, detailed, and intelligent answers powered by Google Gemini.")

# Sidebar for API Key
with st.sidebar:
    st.header("🔑 Configuration")
    api_key = st.text_input("Enter Gemini API Key", type="password", key="search_api_key")
    
    st.markdown("---")
    st.markdown("### 💡 Tips")
    st.markdown("- Make sure your API key starts with `AIzaSy`.")
    st.markdown("- You can search for recipes, technical concepts, business ideas, or general knowledge.")
    
    if st.button("Clear Search History"):
        st.session_state.search_history = []
        st.rerun()

# ----------------- MAIN SEARCH INTERFACE -----------------
st.markdown("### 🌐 What would you like to search today?")
search_query = st.text_input("", placeholder="e.g., Explain commercial bakery yield loss formulas or latest solar panel technologies...")

col1, col2 = st.columns([1, 5])
with col1:
    search_button = st.button("Search AI 🚀", type="primary")

# ----------------- SEARCH & GENERATE LOGIC -----------------
if search_button:
    if not api_key:
        st.warning("Please enter your Gemini API Key in the sidebar first.")
    elif not search_query:
        st.warning("Please type a search query.")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel("gemini-2.5-flash")
            
            prompt = f"""
            You are an advanced, accurate, and helpful AI Search Engine and Expert Knowledge Assistant.
            User Search Query / Topic: {search_query}
            
            Please provide a comprehensive, well-structured, easy-to-read, and detailed response with clear headings, bullet points, and practical insights.
            """
            
            with st.spinner("Searching and synthesizing information..."):
                response = model.generate_content(prompt)
                answer = response.text
                
                # Save to history (latest first)
                st.session_state.search_history.insert(0, {"query": search_query, "answer": answer})
                
        except Exception as e:
            st.error(f"Error: {e}")

# ----------------- DISPLAY SEARCH RESULTS & HISTORY -----------------
if st.session_state.search_history:
    st.markdown("---")
    st.subheader("📑 Search Results & History")
    
    for item in st.session_state.search_history:
        with st.container():
            st.markdown(f"### 🔎 Q: {item['query']}")
            st.markdown(item['answer'])
            st.markdown("---")
else:
    if not search_button:
        st.info("💡 Enter your search term above and click **Search AI** to get started.")
