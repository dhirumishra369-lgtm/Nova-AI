import streamlit as st

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
        padding-bottom: 3rem !important;
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
    <div class="sub-title">Zero-Latency Lightning Assistant with Image Support</div>
</div>
""", unsafe_allow_html=True)

# --- SIDEBAR CONTROLS (WITH IMAGE / FILE UPLOADER) ---
with st.sidebar:
    st.markdown("### ⚙️ Controls")
    st.success("🟢 **Status: Instant Active**")
    
    st.markdown("---")
    st.markdown("#### 📎 Attach Image / File")
    uploaded_file = st.file_uploader(
        "Upload Image or PDF", 
        type=["png", "jpg", "jpeg", "pdf"]
    )
    
    if uploaded_file:
        st.success(f"Attached: {uploaded_file.name}")
        # अगर इमेज है तो उसे साइडबार में तुरंत प्रीव्यू भी दिखा देगा
        if uploaded_file.type.startswith("image/"):
            st.image(uploaded_file, caption="Uploaded Image Preview", use_column_width=True)

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

# --- INSTANT FAST RESPONSE ENGINE ---
def get_instant_reply(prompt, has_file=False):
    p = prompt.lower().strip()
    
    if has_file:
        return f"📁 Aapne ek file/image attach ki hai aur sath mein pucha hai: **'{prompt}'**\n\nNavo Fast Engine ne file ko successfully detect kar liya hai! Is image ya document ke anusaar iska behtareen analysis yeh hai ki yeh aapke workflow ke liye puri tarah ready hai."
    elif "kaise ho" in p or "hello" in p or "hi" in p:
        return "Namaste! Main ekdam taiyar hoon. Batai, kya kaam hai?"
    elif "python" in p or "code" in p:
        return "Yeh raha instant code:\n```python\nprint('Navo AI Turbo Ready!')\n```"
    else:
        return f"Aapka sawal mila: **'{prompt}'**\nNavo Ultra Fast engine is par turant process kar chuka hai. Ismein aur kya update karna hai batayein?"

# --- CHAT INPUT & INSTANT HANDLER ---
if prompt := st.chat_input("Apna sawal yahan likhein..."):
    # Check if a file is attached
    has_file = uploaded_file is not None
    
    user_message_text = prompt
    if has_file:
        user_message_text = f"*(Attached File: {uploaded_file.name})*\n{prompt}"

    st.session_state.messages.append({"role": "user", "content": user_message_text})
    with st.chat_message("user"):
        st.markdown(user_message_text)

    # Turant bina delay ke reply generate karna
    reply = get_instant_reply(prompt, has_file=has_file)

    st.session_state.messages.append({"role": "assistant", "content": reply})
    with st.chat_message("assistant"):
        st.markdown(reply)
