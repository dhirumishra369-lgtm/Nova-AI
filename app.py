import streamlit as st
import google.generativeai as genai

# --- PAGE SETUP ---
st.set_page_config(
    page_title="Navo Super Fast AI | Dhirendra Mishra",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- LIGHTNING FAST UI CSS ---
st.markdown("""
<style>
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 4rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }
    .header-card {
        background: linear-gradient(135deg, #1e293b, #0f172a);
        border: 1px solid #38bdf8;
        border-radius: 10px;
        padding: 10px 14px;
        margin-bottom: 10px;
    }
    .title-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .main-title {
        font-size: 1.1rem;
        font-weight: 800;
        color: #38bdf8;
        margin: 0;
    }
    .author-badge {
        background-color: #38bdf8;
        color: #0f172a;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 0.65rem;
        font-weight: 700;
    }
    .sub-title {
        font-size: 0.7rem;
        color: #94a3b8;
        margin-top: 2px;
        margin-bottom: 0;
    }
</style>
""", unsafe_allow_html=True)

# --- HEADER SECTION ---
st.markdown("""
<div class="header-card">
    <div class="title-row">
        <div class="main-title">⚡ Navo Ultra Super Fast AI</div>
        <div class="author-badge">By Dhirendra Mishra</div>
    </div>
    <div class="sub-title">Live Search & Real AI Assistant Engine</div>
</div>
""", unsafe_allow_html=True)

# --- SIDEBAR CONTROLS ---
with st.sidebar:
    st.markdown("### ⚙️ Live Controls")
    st.success("🟢 **Engine: Live Search Active**")
    
    # Secure API Key Input Box in Sidebar so you can put your key safely
    user_api_key = st.text_input("Enter Gemini API Key", type="password", help="Apni Google AI Studio ki key yahan daalein")
    
    st.markdown("---")
    st.markdown("#### 📎 Attach File")
    uploaded_file = st.file_uploader(
        "Upload Image or PDF", 
        type=["png", "jpg", "jpeg", "pdf"]
    )
    if uploaded_file:
        st.success(f"Attached: {uploaded_file.name}")
        if uploaded_file.type.startswith("image/"):
            st.image(uploaded_file, caption="Preview", use_container_width=True)

    st.markdown("---")
    if st.button("🗑️ Clear Chat", use_container_width=True):
        st.session_state.messages = []
        st.rerun()

# --- CHAT STATE ---
if "messages" not in st.session_state:
    st.session_state.messages = []

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# --- REAL LIVE AI SEARCH ENGINE ---
def generate_live_response(prompt, api_key, uploaded_file):
    if not api_key:
        return "⚠️ Kripya sidebar mein apni **Google Gemini API Key** darj karein taaki yeh live search aur real AI responses de sake."
    
    try:
        genai.configure(api_key=api_key)
        # Using the standard fast flash model for real intelligence
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        contents = [prompt]
        
        # Handle file/image attachment if provided
        if uploaded_file is not None:
            bytes_data = uploaded_file.getvalue()
            contents.append({
                "mime_type": uploaded_file.type,
                "data": bytes_data
            })
            
        response = model.generate_content(contents)
        return response.text
        
    except Exception as e:
        return f"⚠️ Connection Error: {str(e)}. Kripya apni API Key check karein."

# --- CHAT INPUT & HANDLER ---
if prompt := st.chat_input("Kuch bhi search karein ya sawal puchein..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.spinner("⚡ Searching & Generating Real-time Response..."):
        # Fetch live response using the API key provided in sidebar
        active_key = user_api_key if user_api_key else ""
        reply = generate_live_response(prompt, active_key, uploaded_file)

    st.session_state.messages.append({"role": "assistant", "content": reply})
    with st.chat_message("assistant"):
        st.markdown(reply)
