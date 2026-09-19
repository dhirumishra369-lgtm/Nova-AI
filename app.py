import streamlit as st
import requests
import json
import base64
import time

# पेज सेटअप
st.set_page_config(
    page_title="Navo Super Fast AI | Dhirendra Mishra",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# मोबाइल स्क्रीन अलाइनमेंट CSS
st.markdown("""
<style>
    /* हेडर को नेविगेशन बार के नीचे लाने के लिए सही मार्जिन */
    .block-container {
        padding-top: 3.5rem !important;
        padding-bottom: 5rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }

    /* कॉम्पैक्ट क्लीन हेडर कार्ड */
    .header-card {
        background: #f1f5f9;
        border: 1px solid #cbd5e1;
        border-radius: 12px;
        padding: 12px 14px;
        margin-bottom: 14px;
    }
    
    .title-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 6px;
    }

    .main-title {
        font-size: 1.15rem;
        font-weight: 700;
        color: #1d4ed8;
        margin: 0;
        line-height: 1.2;
    }

    .author-badge {
        background-color: #0f172a;
        color: #38bdf8;
        padding: 3px 8px;
        border-radius: 14px;
        font-size: 0.70rem;
        font-weight: 600;
        white-space: nowrap;
    }

    .sub-title {
        font-size: 0.76rem;
        color: #64748b;
        margin-top: 4px;
        margin-bottom: 0;
    }

    /* चैट मैसेज स्टाइल */
    [data-testid="stChatMessage"] {
        padding: 10px 12px !important;
        margin-bottom: 8px !important;
        border-radius: 10px !important;
    }
</style>
""", unsafe_allow_html=True)

# हेडर सेक्शन
st.markdown("""
<div class="header-card">
    <div class="title-row">
        <div class="main-title">⚡ Navo Super Fast AI</div>
        <div class="author-badge">By Dhirendra Mishra</div>
    </div>
    <div class="sub-title">High-Performance Intelligent Assistant</div>
</div>
""", unsafe_allow_html=True)

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