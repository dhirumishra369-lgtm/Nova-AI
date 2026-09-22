import os
import google.generativeai as genai

# 1. अपनी Gemini API Key यहाँ डालें (या एनवायरनमेंट वेरिएबल सेट करें)
API_KEY = "Yahan_Apni_AIzaSy_Wali_Key_Daalein"

def run_ai_search():
    if API_KEY == "Yahan_Apni_AIzaSy_Wali_Key_Daalein" or not API_KEY:
        print("❌ Pehle apni valid Google Gemini API Key code me daalein!")
        return

    # Gemini Configure Karein
    genai.configure(api_key=API_KEY)
    
    # Model Select Karein
    model = genai.GenerativeModel("gemini-2.5-flash")

    print("==================================================")
    print(" 🔍 AI-Powered Python Search Engine (Gemini)")
    print("==================================================")
    print("Exit karne ke liye 'exit' type karein.\n")

    while True:
        query = input("🌐 Kya search karna chahte hain?: ")
        
        if query.lower() == 'exit':
            print("Dhanyawad! Alvida.")
            break
            
        if not query.strip():
            continue

        print("\n⏳ Searching and synthesizing info...")
        try:
            prompt = f"""
            You are an advanced AI Search Engine and Expert Knowledge Assistant.
            User Search Query: {query}
            
            Please provide a comprehensive, clear, detailed, and structured response.
            """
            
            response = model.generate_content(prompt)
            print("\n----------------- SEARCH RESULTS -----------------")
            print(response.text)
            print("--------------------------------------------------\n")
            
        except Exception as e:
            print(f"❌ Error aa gaya: {e}\n")

if __name__ == "__main__":
    run_ai_search()
