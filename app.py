import streamlit as st
import pandas as pd
import google.generativeai as genai
import io

st.set_page_config(
    page_title="Excel Data Analysis & AI Assistant",
    page_icon="📈",
    layout="wide"
)

# ----------------- SESSION STATE SETUP -----------------
if "messages" not in st.session_state:
    st.session_state.messages = []

# ----------------- UI HEADER -----------------
st.title("📈 Excel Data Analysis & AI Assistant")
st.caption("Upload your Excel sheet, view data insights, and chat with Google Gemini to analyze your numbers instantly.")

# Sidebar for API Key & File Upload
with st.sidebar:
    st.header("🔑 Configuration")
    api_key = st.text_input("Enter Gemini API Key", type="password", key="excel_chat_key")
    
    st.markdown("---")
    st.header("📁 Upload Excel File")
    uploaded_file = st.file_uploader("Upload .xlsx, .xls or .csv file", type=["xlsx", "xls", "csv"])

# ----------------- MAIN APP LOGIC -----------------
df = None
if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        st.success(f"Successfully loaded: **{uploaded_file.name}**")
        
        # Display Data Overview
        with st.expander("📊 Preview Uploaded Data & Summary", expanded=True):
            st.dataframe(df.head(10), use_container_width=True)
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Rows", df.shape[0])
            with col2:
                st.metric("Total Columns", df.shape[1])
            with col3:
                st.metric("Missing Values", df.isna().sum().sum())
                
            st.subheader("Statistical Summary")
            st.dataframe(df.describe(), use_container_width=True)
            
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
else:
    st.info("👈 Please upload an Excel or CSV file from the sidebar to start your data analysis.")

st.markdown("---")

# ----------------- AI CHAT CONSULTANT FOR EXCEL -----------------
st.subheader("🤖 Google Gemini Excel Data Analyst")
st.caption("Ask questions about your uploaded spreadsheet (e.g., 'Find the top 5 highest values', 'Summarize total sales', or 'Check for discrepancies').")

# Display chat history
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# Accept user input
if prompt := st.chat_input("Apne Excel data ke baare mein kuch bhi puchiye..."):
    if not api_key:
        st.warning("Please enter your Gemini API Key in the sidebar first.")
    elif df is None:
        st.warning("Please upload an Excel file first so the AI can analyze it.")
    else:
        # Add user message to chat history
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        # Generate response from Gemini using dataframe context
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel("gemini-2.5-flash")
            
            # Convert dataframe sample / info to string context for AI
            data_summary = df.head(50).to_string() # Passing first 50 rows as context
            columns_list = list(df.columns)
            
            context_prompt = f"""
            You are an expert Data Analyst and Excel expert. 
            The user has uploaded an Excel file with columns: {columns_list}
            Here is a preview of the data (up to 50 rows):
            {data_summary}

            User Query / Analysis Request: {prompt}
            Please provide a clear, accurate, professional, and data-backed analysis or solution.
            """
            
            with st.spinner("Gemini is analyzing your Excel data..."):
                response = model.generate_content(context_prompt)
                bot_reply = response.text

            # Add assistant response to chat history
            st.session_state.messages.append({"role": "assistant", "content": bot_reply})
            with st.chat_message("assistant"):
                st.markdown(bot_reply)
                
        except Exception as e:
            error_msg = f"Error: {e}"
            st.error(error_msg)
