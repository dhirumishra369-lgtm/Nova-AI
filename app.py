import streamlit as st
import requests
import json
import base64
import time

# पेज सेटअप - मोबाइल स्क्रीन के लिए बेहतर लेआउट
st.set_page_config(
    page_title="Navo Super Fast AI | Dhirendra Mishra",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# मोबाइल-फ्रेंडली CSS स्टाइलिंग
st.markdown("""
<style>
    /* हेडर टेक्स्ट साइज - मोबाइल और लैपटॉप दोनों के लिए */
    .main-title {
        font-size: clamp(1.6rem, 5vw, 2.4rem);
        font-weight: 700;
        color: #2563eb;
        line-height: 1.2;
        margin-bottom: 4px;
    }
    .sub-title {
        font-size: clamp(0.85rem, 3vw, 1.05rem);
        color: #64748b;
        margin-top: 2px;
        margin-bottom: 12px;
        display: flex;
        flex-wrap: wrap;
        align-items: center;
        gap: 8px;
    }
    .author-tag {
        background-color: #1e293b;
        color: #38bdf8;
        padding: 3px 10px;
        border-radius: 10px;
        font-size: 0.78rem;
        font-weight: 600;
        display: inline-block;
    }

    /* मोबाइल पर चैट पैडिंग सही करना */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 4rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }

    /* मैसेज बॉक्स का फॉन्ट और साइज */
    [data-testid="stChatMessage"] {
        padding: 12px 14px !important;
        margin-bottom: 10px !important;
        border-radius: 12px !important;
    }
</style>
""", unsafe_allow_html=True)

# हेडर सेक्शन
st.markdown('<div class="main-title">⚡ Navo Super Fast AI</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title"><span>High-Performance Assistant</span><span class="author-tag">By Dhirendra Mishra</span></div>', unsafe_allow_html=True)
st.divider()

API_KEY = "AQ.Ab8RN6I8IV-N4P4uqzuUsH1VD1IrxnjhzARZZZLlMdg1-G4ixA"

# साइडबार
with st.sidebar:
    st.markdown("### ⚙️ Control Center")
    st.success("🟢 **Navo Engine : Active**")
    
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

# चैट हिस्ट्री
if "messages" not in st.session_state:
    st.session_state.messages = []

# मैसेज दिखाना
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# इनपुट बॉक्स
if prompt := st.chat_input("Apna sawal yahan likhein..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    contents_payload = []
    for msg in st.session_state.messages[:-1]:
        api_role = "user" if msg["role"] == "user" else "model"
        contents_payload.append({
            "role": api_role,
            "parts": [{"text": msg["content"]}]
        })

    current_parts = [{"text": prompt}]
    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        encoded_file = base64.b64encode(file_bytes).decode("utf-8")
        current_parts.append({
            "inline_data": {
                "mime_type": uploaded_file.type,
                "data": encoded_file
            }
        })

    contents_payload.append({
        "role": "user",
        "parts": current_parts
    })

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key={API_KEY}"
    headers = {"Content-Type": "application/json"}
    payload = {"contents": contents_payload}

    reply = ""
    with st.spinner("Thinking..."):
        for attempt in range(3):
            try:
                res = requests.post(url, headers=headers, json=payload)
                data = res.json()
                if "candidates" in data:
                    reply = data["candidates"][0]["content"]["parts"][0]["text"]
                    break
                elif data.get("error", {}).get("code") == 503:
                    time.sleep(2)
                    continue
                else:
                    reply = f"API Response: {data}"
                    break
            except Exception as e:
                reply = f"Error: {e}"
                break

    if not reply:
        reply = "Server busy hai. Kripya punah prayas karein."

    st.session_state.messages.append({"role": "assistant", "content": reply})
    with st.chat_message("assistant"):
        st.markdown(reply)
