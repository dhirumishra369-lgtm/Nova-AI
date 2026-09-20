import streamlit as st
import google.generativeai as genai

# --- PAGE SETUP ---
st.set_page_config(
    page_title="Navo Super Fast AI | Dhirendra Mishra",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
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
    <div class="sub-title">Connected to Official Google Gemini Engine</div>
</div>
""", unsafe_allow_html=True)

# --- SIDEBAR CONFIGURATION ---
with st.sidebar:
    st.markdown("### ⚙️ Gemini API Configuration")
    
    # Direct text input for Gemini API Key in sidebar
    user_api_key = st.text_input("Enter your Gemini API Key", type="password", placeholder="AIzaSy...")
    
    if user_api_key:
        st.success("🟢 API Key Provided!")
    else:
        st.warning("⚠️ Kripya Google AI Studio se apni API key yahan daalein.")
    
    st.markdown("---")
    st.markdown("#### 📎 Attach File / Image")
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

# --- REAL GEMINI API CALL FUNCTION ---
def ask_real_gemini(prompt_text, api_key, file_obj):
    if not api_key:
        return "⚠️ Kripya sidebar mein apni valid Google Gemini API Key darj karein taaki yeh seedhe Google Gemini se jud sake."
    
    try:
        # Configure Gemini with the newly generated secure user-provided key
        genai.configure(api_key=api_key.strip())
        
        # Using gemini-2.5-flash for maximum support and reliability
        model = genai.GenerativeModel('gemini-2.5-flash')
        
        contents = [prompt_text]
        
        # If an image or file is uploaded, pass it along safely
        if file_obj is not None:
            bytes_data = file_obj.getvalue()
            contents.append({
                "mime_type": file_obj.type,
                "data": bytes_data
            })
            
        response = model.generate_content(contents)
        return response.text
        
    except Exception as err:
        return f"⚠️ Connection Error: Kripya apni API Key check karein. Details: {str(err)}"

# --- CHAT INPUT & HANDLER ---
if prompt := st.chat_input("Google Gemini se kuch bhi puchiye..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.spinner("⚡ Connecting to Google Gemini..."):
        bot_reply = ask_real_gemini(prompt, user_api_key, uploaded_file)

    st.session_state.messages.append({"role": "assistant", "content": bot_reply})
    with st.chat_message("assistant"):
        st.markdown(bot_reply)
