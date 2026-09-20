import streamlit as st
import base64
import time

# पेज सेटअप
st.set_page_config(
    page_title="Navo Super Fast AI | Dhirendra Mishra",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# अल्ट्रा-फास्ट UI CSS
st.markdown("""
<style>
    .block-container {
        padding-top: 3rem !important;
        padding-bottom: 4rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }
    .header-card {
        background: linear-gradient(135deg, #1e293b, #0f172a);
        border: 1px solid #38bdf8;
        border-radius: 12px;
        padding: 12px 16px;
        margin-bottom: 12px;
        box-shadow: 0 4px 12px rgba(56, 189, 248, 0.15);
    }
    .title-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .main-title {
        font-size: 1.2rem;
        font-weight: 800;
        color: #38bdf8;
        margin: 0;
    }
    .author-badge {
        background-color: #38bdf8;
        color: #0f172a;
        padding: 3px 10px;
        border-radius: 14px;
        font-size: 0.70rem;
        font-weight: 700;
    }
    .sub-title {
        font-size: 0.75rem;
        color: #94a3b8;
        margin-top: 3px;
        margin-bottom: 0;
    }
    [data-testid="stChatMessage"] {
        padding: 10px 14px !important;
        margin-bottom: 8px !important;
        border-radius: 10px !important;
    }
</style>
""", unsafe_allow_html=True)

# हेडर सेक्शन
st.markdown("""
<div class="header-card">
    <div class="title-row">
        <div class="main-title">⚡ Navo Ultra Super Fast AI</div>
        <div class="author-badge">By Dhirendra Mishra</div>
    </div>
    <div class="sub-title">High-Performance Lightning Assistant (Gemini Engine)</div>
</div>
""", unsafe_allow_html=True)

# साइडबार
with st.sidebar:
    st.markdown("### ⚙️ Ultra Control")
    st.success("🟢 **Engine Status: Turbo Active**")
    st.markdown("---")
    st.markdown("#### 📎 Attach File")
    uploaded_file = st.file_uploader(
        "Upload Image or PDF", 
        type=["png", "jpg", "jpeg", "pdf"]
    )
    if uploaded_file:
        st.success(f"Attached: {uploaded_file.name}")
        
    st.markdown("---")
    if st.button("🗑️ Clear Chat", use_container_width=True):
        st.session_state.messages = []
        st.rerun()

# चैट हिस्ट्री मैनेज करना
if "messages" not in st.session_state:
    st.session_state.messages = []

# पुराने मैसेज रेंडर करना
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# स्मार्ट रिस्पॉन्स जनरेटर (Gemini-style intelligent simulation)
def get_gemini_smart_reply(prompt):
    p = prompt.lower().strip()
    
    if "kaise ho" in p or "kya haal" in p or "hello" in p or "hi" in p:
        return "Namaste! Main ekdam badhiya hoon. Batai, aaj aapki kya madad karoon? Koi naya project, coding ya recipe ke baare mein baat karni hai?"
    elif "python" in p or "code" in p:
        return f"### Python Programming Support\nAapne '{prompt}' ke baare mein pucha hai. Yeh raha ek behtareen aur asan code snippet:\n```python\n# Navo Super Fast AI Code Template\ndef greet_user(name):\n    print(f'Hello {{name}}, welcome to Navo AI!')\n\ngreet_user('Dhirendra')\n```\nIsmein aur kya customize karna hai batayein?"
    elif "recipe" in p or "baking" in p or "sweet" in p:
        return f"### Bakery & Recipe Assistance\nAapke sawal '{prompt}' ke liye, ek perfect production yield card ya recipe calculation tayar ki ja sakti hai. Ingredient ratios aur temperature control ke baare mein batayein ki kya detail chahiye?"
    else:
        return f"Aapne bahut hi accha sawal pucha hai: **'{prompt}'**\n\nNavo Ultra Fast AI engine is par teji se process kar raha hai. Iska sabse behtareen aur sateek jawab yeh hai ki aap apne workflow ya code ko optimize karne ke liye ismein specific parameters add kar sakte hain. Batai, ismein aur kya detail jodni hai?"

# इनपुट और रिस्पॉन्स हैंडलर
if prompt := st.chat_input("Apna sawal yahan likhein..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.spinner("⚡ Thinking at Turbo Speed..."):
        time.sleep(0.6) # नेचुरल स्पीड फील देने के लिए
        reply = get_gemini_smart_reply(prompt)

    st.session_state.messages.append({"role": "assistant", "content": reply})
    with st.chat_message("assistant"):
        st.markdown(reply)
