import streamlit as st
import google.generativeai as genai
import pandas as pd

# Page Configuration
st.set_page_config(
    page_title="Smart Kitchen SaaS & Costing",
    page_icon="🍳",
    layout="wide"
)

# Safe API Setup (लोकल और क्लाउड दोनों पर बिना क्रैश हुए चलेगा)
api_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

# अगर secrets नहीं मिले तो साइडबार से API Key लेने का विकल्प
with st.sidebar:
    st.header("⚙️ सेटिंग्स")
    if not api_key:
        api_key = st.text_input("Gemini API Key डालें:", type="password")
        if api_key:
            st.success("API Key सेट हो गई!")
    else:
        st.success("✅ API Key कनेक्टेड है")

if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
else:
    model = None

# Custom Styling
st.markdown("""
    <style>
    .main-title { font-size: 28px; font-weight: bold; color: #1E3A8A; margin-bottom: 5px; }
    .sub-title { font-size: 15px; color: #4B5563; margin-bottom: 20px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🍳 Smart Recipe Costing & Inventory Yield System</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">कमर्शियल किचन, बेकरी और मिठाई उत्पादन के लिए सटीक लागत और मार्जिन विश्लेषक</div>', unsafe_allow_html=True)

# Tabs
tab1, tab2 = st.tabs(["📊 रेसिपी कॉस्टिंग & यील्ड कैलकुलेटर", "🤖 AI शेफ़ / बिज़नेस कंसल्टेंट"])

# ==================== TAB 1: CALCULATOR ====================
with tab1:
    col_left, col_right = st.columns([1.2, 0.8], gap="large")

    with col_left:
        st.subheader("1. बैच और उत्पाद विवरण")
        dish_name = st.text_input("उत्पाद / रेसिपी का नाम", value="Premium Kaju Katli")
        
        st.write("**सामग्री विवरण (Ingredients & Rates):**")
        st.caption("टेबल में सीधे क्लिक करके इंग्रीडिएंट, मात्रा (kg) और रेट प्रति kg बदल सकते हैं:")

        default_data = {
            "Ingredient": ["Kaju (Cashew)", "Sugar (चीनी)", "Silver Vark", "Cardamom/Ghee"],
            "Quantity_KG": [10.0, 8.0, 0.05, 0.20],
            "Rate_Per_KG": [680.0, 42.0, 4000.0, 600.0]
        }
        
        edited_df = st.data_editor(
            pd.DataFrame(default_data),
            num_rows="dynamic",
            use_container_width=True
        )

        st.subheader("2. यील्ड और अतिरिक्त खर्चे")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            cooking_loss_pct = st.number_input("कुकिंग / मॉइस्चर लॉस (%)", min_value=0.0, max_value=50.0, value=12.0, step=0.5)
            packaging_cost = st.number_input("पैकेजिंग खर्च (कुल बैच का ₹)", min_value=0.0, value=250.0, step=10.0)
        with col_c2:
            labor_gas_cost = st.number_input("लेबर + गैस/बिजली खर्च (₹)", min_value=0.0, value=400.0, step=20.0)
            target_margin_pct = st.number_input("वांछित मुनाफ़ा मार्जिन (%)", min_value=5.0, max_value=80.0, value=35.0, step=1.0)

    # Calculations
    edited_df["Total_Cost"] = edited_df["Quantity_KG"] * edited_df["Rate_Per_KG"]
    raw_material_cost = edited_df["Total_Cost"].sum()
    raw_material_weight = edited_df["Quantity_KG"].sum()

    # Final Yield Weight
    final_yield_weight = raw_material_weight * (1 - (cooking_loss_pct / 100))
    total_batch_cost = raw_material_cost + packaging_cost + labor_gas_cost

    cost_per_kg = (total_batch_cost / final_yield_weight) if final_yield_weight > 0 else 0
    selling_price_per_kg = cost_per_kg / (1 - (target_margin_pct / 100)) if target_margin_pct < 100 else 0
    profit_per_kg = selling_price_per_kg - cost_per_kg

    with col_right:
        st.subheader("📋 आउटपुट और कॉस्ट समरी")
        
        st.metric("कुल बैच लागत (Total Batch Cost)", f"₹{total_batch_cost:,.2f}")
        st.metric("फ़ाइनल आउटपुट वजन (Final Yield)", f"{final_yield_weight:.2f} KG", delta=f"-{cooking_loss_pct}% Process Loss", delta_color="inverse")
        
        st.divider()
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.metric("लागत प्रति KG (Cost/KG)", f"₹{cost_per_kg:,.2f}")
        with col_m2:
            st.metric("सुझाया गया विक्रय मूल्य", f"₹{selling_price_per_kg:,.2f}")

        st.info(f"💡 **शुद्ध मुनाफ़ा:** ₹{profit_per_kg:,.2f} प्रति KG ({target_margin_pct}% मार्जिन पर)")

        # Export Button
        summary_export = pd.DataFrame([{
            "Product": dish_name,
            "Raw Weight (KG)": raw_material_weight,
            "Final Yield (KG)": round(final_yield_weight, 2),
            "Raw Material Cost": round(raw_material_cost, 2),
            "Total Batch Cost": round(total_batch_cost, 2),
            "Cost Per KG": round(cost_per_kg, 2),
            "Selling Price Per KG": round(selling_price_per_kg, 2)
        }])
        
        st.download_button(
            label="📥 कॉस्टिंग रिपोर्ट डाउनलोड करें (CSV)",
            data=summary_export.to_csv(index=False),
            file_name=f"{dish_name}_costing_summary.csv",
            mime="text/csv",
            use_container_width=True
        )

# ==================== TAB 2: AI CONSULTANT ====================
with tab2:
    st.subheader("💡 AI कमर्शियल किचन असिस्टेंट")
    st.caption("रेसिपी में वेस्टेज कम करने, शेल्फ लाइफ बढ़ाने या पैकेजिंग ऑप्टिमाइज़ करने के लिए पूछें:")

    user_query = st.text_area("अपना सवाल लिखें:", placeholder="जैसे: काजू कतली में मॉइस्चर लॉस कम करने और 15 दिन शेल्फ लाइफ रखने के लिए क्या सावधानियां रखें?")
    
    if st.button("पूछें (Analyze with AI)"):
        if not model:
            st.error("कृपया बाईं तरफ (Sidebar) में अपनी Gemini API Key डालें।")
        elif not user_query.strip():
            st.warning("कृपया कोई सवाल लिखें।")
        else:
            with st.spinner("AI विश्लेषण कर रहा है..."):
                system_prompt = f"""
                तुम एक सीनियर कमर्शियल शेफ़ और प्रोडक्शन मैनेजर हो। 
                वर्तमान रेसिपी: {dish_name}
                लागत प्रति किलो: ₹{cost_per_kg:.2f}
                यूज़र का सवाल: {user_query}
                व्यावहारिक, सटीक और इंडस्ट्री-ग्रेड सलाह दो।
                """
                response = model.generate_content(system_prompt)
                st.markdown(response.text)
