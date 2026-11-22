import datetime
import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Nova-AI", page_icon="🍰", layout="wide")

st.title("🍰 Nova-AI — Recipe Costing & Yield")
st.caption("App status: Online & Running")

if "df" not in st.session_state:
  st.session_state.df = pd.DataFrame([
      {"Ingredient": "Kaju", "Qty_KG": 5.0, "Rate_KG": 680.0},
      {"Ingredient": "Sugar", "Qty_KG": 8.0, "Rate_KG": 42.0},
  ])

recipe_name = st.text_input("Recipe Name", value="Kaju Katli")

edited = st.data_editor(
    st.session_state.df,
    num_rows="dynamic",
    use_container_width=True,
    key="editor_main",
)
st.session_state.df = edited

c1, c2, c3 = st.columns(3)
loss_pct = c1.number_input("Loss (%)", 0.0, 90.0, 10.0)
labor = c2.number_input("Labor + Fuel (₹)", 0.0, 100000.0, 400.0)
packing = c3.number_input("Packaging (₹)", 0.0, 100000.0, 250.0)

valid = edited.dropna(subset=["Qty_KG", "Rate_KG"]).copy()
valid["Amount"] = valid["Qty_KG"] * valid["Rate_KG"]

raw_kg = valid["Qty_KG"].sum()
raw_cost = valid["Amount"].sum()
net_kg = raw_kg * (1 - loss_pct / 100.0)
total_cost = raw_cost + labor + packing
cost_per_kg = (total_cost / net_kg) if net_kg > 0 else 0.0

m1, m2, m3, m4 = st.columns(4)
m1.metric("Raw Cost", f"₹{raw_cost:,.2f}")
m2.metric("Total Batch Cost", f"₹{total_cost:,.2f}")
m3.metric("Net Yield", f"{net_kg:,.2f} KG")
m4.metric("Cost / KG", f"₹{cost_per_kg:,.2f}")

# Excel export
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Costing"
ws.append(["Recipe", recipe_name, "Date", str(datetime.date.today())])
ws.append(["Ingredient", "Qty_KG", "Rate_KG", "Amount"])
for _, r in valid.iterrows():
  ws.append([
      str(r["Ingredient"]),
      float(r["Qty_KG"]),
      float(r["Rate_KG"]),
      float(r["Amount"]),
  ])
ws.append(["Total Batch Cost", total_cost])
ws.append(["Net Yield (KG)", net_kg])
ws.append(["Cost Per KG", cost_per_kg])

buf = io.BytesIO()
wb.save(buf)

st.download_button(
    "📥 Download Excel (.xlsx)",
    data=buf.getvalue(),
    file_name=f"{recipe_name.replace(' ', '_')}_Costing.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
