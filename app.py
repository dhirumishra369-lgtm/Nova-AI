import datetime
import io
import json
import re
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Nova-AI | THF Price List Upload, Recipe OCR & Colorful Excel Studio",
    page_icon="🎨",
    layout="wide",
)

# ----------------- SUPER COLORFUL CUSTOM CSS -----------------
st.markdown(
    """
    <style>
    .hero-banner {
        background: linear-gradient(135deg, #3B0764 0%, #6D28D9 45%, #0284C7 100%);
        padding: 18px 24px;
        border-radius: 14px;
        color: #FFFFFF;
        box-shadow: 0 6px 18px rgba(109, 40, 217, 0.25);
        margin-bottom: 14px;
    }
    .hero-banner h1 { margin: 0; font-size: 24px; font-weight: 800; }
    .hero-banner p { margin: 4px 0 0 0; font-size: 13px; color: #FEF08A; font-weight: 500; }
    .upload-card-1 {
        background: linear-gradient(135deg, #ECFDF5 0%, #D1FAE5 100%);
        border: 2px solid #059669;
        border-radius: 12px;
        padding: 12px 14px;
        margin-bottom: 8px;
    }
    .upload-card-2 {
        background: linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%);
        border: 2px solid #2563EB;
        border-radius: 12px;
        padding: 12px 14px;
        margin-bottom: 8px;
    }
    .upload-card-3 {
        background: linear-gradient(135deg, #FFFBEB 0%, #FEF3C7 100%);
        border: 2px solid #D97706;
        border-radius: 12px;
        padding: 12px 14px;
        margin-bottom: 8px;
    }
    .kpi-box {
        border-radius: 12px;
        padding: 14px 16px;
        color: #FFFFFF;
        box-shadow: 0 4px 10px rgba(0,0,0,0.08);
    }
    .kpi-lbl { font-size: 11px; font-weight: 700; text-transform: uppercase; opacity: 0.92; }
    .kpi-num { font-size: 22px; font-weight: 800; margin-top: 4px; }
    .kpi-tag { font-size: 11px; opacity: 0.88; margin-top: 2px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------- SESSION STATE -----------------
if "recipe_book" not in st.session_state:
    st.session_state.recipe_book = {
        "WALNUT BAKLAVA": {
            "yield_pcs": 215,
            "df": pd.DataFrame([
                {"Category": "BASE SHEET", "Ingredient": "Fillo Sheet", "Item_Code": "RM0165", "Quantity_KG": 0.360, "Rate_Per_KG": 521.13},
                {"Category": "MIX", "Ingredient": "Kaju 04 Pcs", "Item_Code": "RM0067", "Quantity_KG": 0.600, "Rate_Per_KG": 841.87},
                {"Category": "MIX", "Ingredient": "Walnut", "Item_Code": "RM0505", "Quantity_KG": 0.350, "Rate_Per_KG": 1150.00},
                {"Category": "MIX", "Ingredient": "(Amul) Salted Butter", "Item_Code": "RM0052", "Quantity_KG": 0.320, "Rate_Per_KG": 537.50},
                {"Category": "MIX", "Ingredient": "Milk Maid", "Item_Code": "RM0300", "Quantity_KG": 0.220, "Rate_Per_KG": 327.88},
                {"Category": "SUGAR SYRUP", "Ingredient": "Sugar (D.F)", "Item_Code": "RM0046", "Quantity_KG": 0.500, "Rate_Per_KG": 65.27},
                {"Category": "SUGAR SYRUP", "Ingredient": "Honey", "Item_Code": "RM0220", "Quantity_KG": 0.150, "Rate_Per_KG": 311.10},
            ]),
        }
    }

if "price_list_df" not in st.session_state:
    st.session_state.price_list_df = pd.DataFrame(
        columns=["No.", "Description", "Base Unit of Measure", "Unit Cost"]
    )

if "price_lookup_dict" not in st.session_state:
    st.session_state.price_lookup_dict = {}

# ----------------- NUTRITION & ALIAS MAPS -----------------
NUTRIENTS_DB = {
    "kaju": {"kcal": 553.0, "protein": 18.2, "fat": 43.8, "carbs": 30.2},
    "cashew": {"kcal": 553.0, "protein": 18.2, "fat": 43.8, "carbs": 30.2},
    "pisi kaju": {"kcal": 553.0, "protein": 18.2, "fat": 43.8, "carbs": 30.2},
    "sugar": {"kcal": 387.0, "protein": 0.0, "fat": 0.0, "carbs": 100.0},
    "suger": {"kcal": 387.0, "protein": 0.0, "fat": 0.0, "carbs": 100.0},
    "ghee": {"kcal": 900.0, "protein": 0.0, "fat": 100.0, "carbs": 0.0},
    "butter": {"kcal": 717.0, "protein": 0.85, "fat": 81.0, "carbs": 0.06},
    "milk maid": {"kcal": 321.0, "protein": 7.9, "fat": 8.7, "carbs": 54.4},
    "condence": {"kcal": 321.0, "protein": 7.9, "fat": 8.7, "carbs": 54.4},
    "milk": {"kcal": 42.0, "protein": 3.4, "fat": 1.0, "carbs": 5.0},
    "pista": {"kcal": 562.0, "protein": 20.0, "fat": 45.0, "carbs": 28.0},
    "pistachio": {"kcal": 562.0, "protein": 20.0, "fat": 45.0, "carbs": 28.0},
    "almond": {"kcal": 579.0, "protein": 21.1, "fat": 49.9, "carbs": 21.6},
    "badam": {"kcal": 579.0, "protein": 21.1, "fat": 49.9, "carbs": 21.6},
    "walnut": {"kcal": 654.0, "protein": 15.2, "fat": 65.2, "carbs": 13.7},
    "coconut": {"kcal": 660.0, "protein": 6.9, "fat": 64.5, "carbs": 23.7},
    "honey": {"kcal": 304.0, "protein": 0.3, "fat": 0.0, "carbs": 82.4},
    "maida": {"kcal": 364.0, "protein": 10.3, "fat": 1.0, "carbs": 76.3},
    "fillo": {"kcal": 300.0, "protein": 8.0, "fat": 4.0, "carbs": 60.0},
    "filow": {"kcal": 300.0, "protein": 8.0, "fat": 4.0, "carbs": 60.0},
    "cardamom": {"kcal": 311.0, "protein": 11.0, "fat": 7.0, "carbs": 68.0},
    "cinnamon": {"kcal": 247.0, "protein": 4.0, "fat": 1.2, "carbs": 80.6},
    "cinemon": {"kcal": 247.0, "protein": 4.0, "fat": 1.2, "carbs": 80.6},
    "chocolate": {"kcal": 546.0, "protein": 4.9, "fat": 31.3, "carbs": 61.0},
    "water": {"kcal": 0.0, "protein": 0.0, "fat": 0.0, "carbs": 0.0},
    "salt": {"kcal": 0.0, "protein": 0.0, "fat": 0.0, "carbs": 0.0},
    "citric": {"kcal": 0.0, "protein": 0.0, "fat": 0.0, "carbs": 0.0},
    "cytric": {"kcal": 0.0, "protein": 0.0, "fat": 0.0, "carbs": 0.0},
}

ALIAS_MAP = {
    "fillo sheet": "filow sheet",
    "kaju 04 pcs": "cashew 4 pcs",
    "kaju 08 pcs": "cashew 4 pcs",
    "pisi kaju": "cashew 4 pcs",
    "milk maid": "milkmaid big",
    "condence milk": "milkmaid big",
    "(amul) salted butter": "butter salted",
    "salted butter": "butter salted",
    "sugar (d.f)": "breakfast sugar",
    "suger (d.f)": "breakfast sugar",
    "suger": "sugar",
    "cytrick (char minar)": "citric char minar",
    "cytric (char minar)": "citric char minar",
    "cytric": "citric char minar",
    "cinemon stick": "cinnamon stick",
    "cardamom": "green cardamom",
    "badam": "almond factory",
    "icing suger": "icing sugar",
    "s.c pista": "pishori pista",
    "pista green": "pishori pista",
    "salt": "tata salt",
    "milk": "fresh milk",
}


def get_nutrients(ingredient_name):
    nl = str(ingredient_name).lower()
    for k, v in NUTRIENTS_DB.items():
        if k in nl:
            return v
    return {"kcal": 400.0, "protein": 5.0, "fat": 10.0, "carbs": 70.0}


def build_price_lookup(df_price):
    lookup = {}
    if df_price is None or df_price.empty:
        return lookup
    code_col = next((c for c in df_price.columns if "no" in str(c).lower() or "code" in str(c).lower()), None)
    desc_col = next((c for c in df_price.columns if "desc" in str(c).lower() or "item" in str(c).lower() or "name" in str(c).lower()), None)
    cost_col = next((c for c in df_price.columns if "cost" in str(c).lower() or "rate" in str(c).lower() or "price" in str(c).lower()), None)
    if not desc_col or not cost_col:
        return lookup
    for _, row in df_price.iterrows():
        desc = str(row.get(desc_col, "")).strip().lower()
        code = str(row.get(code_col, "MANUAL")).strip() if code_col else "MANUAL"
        try:
            cost = float(row.get(cost_col, 0.0) or 0.0)
        except (ValueError, TypeError):
            cost = 0.0
        if desc and desc != "nan":
            lookup[desc] = (code, cost)
    return lookup


def match_rate_and_code(ing_name, price_lookup):
    clean = str(ing_name).strip().lower()
    if not price_lookup:
        return "MANUAL", 0.0
    mapped = ALIAS_MAP.get(clean, clean)
    if mapped in price_lookup:
        return price_lookup[mapped]
    if clean in price_lookup:
        return price_lookup[clean]
    for k, v in price_lookup.items():
        if mapped in k or k in mapped:
            return v
    for k, v in price_lookup.items():
        if clean in k or k in clean:
            return v
    return "MANUAL", 0.0


def parse_price_list_file(file_obj):
    file_obj.seek(0)
    raw = pd.read_excel(file_obj, header=None)
    hdr_row = 0
    for r_i, r in raw.iterrows():
        vals = [str(x).strip().lower() for x in r.tolist() if pd.notnull(x)]
        if any("desc" in v or "item" in v for v in vals) and any("cost" in v or "rate" in v or "price" in v for v in vals):
            hdr_row = r_i
            break
    file_obj.seek(0)
    df = pd.read_excel(file_obj, header=hdr_row).dropna(how="all")
    return df


def parse_recipe_excel(file_obj, sheet_name, price_lookup):
    file_obj.seek(0)
    raw_df = pd.read_excel(file_obj, sheet_name=sheet_name, header=None)
    hdr_idx = None
    cat_col = ing_col = qty_col = unit_col = None
    yield_pcs = 100

    for r_idx, row in raw_df.iterrows():
        row_strs = [str(v).strip() for v in row.tolist() if pd.notnull(v)]
        if row_strs and "total yield" in row_strs[0].lower() and len(row_strs) > 1:
            try:
                yield_pcs = int(float(row_strs[1]))
            except Exception:
                pass
        lmap = {str(v).strip().lower(): str(v).strip() for v in row.tolist() if pd.notnull(v)}
        if any("ingredient" in k or "item" in k for k in lmap):
            hdr_idx = r_idx
            for k, orig in lmap.items():
                if "cat" in k or "section" in k:
                    cat_col = orig
                if "ingredient" in k or ("item" in k and not ing_col):
                    ing_col = orig
                if "qty" in k or "quantity" in k or "base" in k:
                    qty_col = orig
                if "unit" in k or "uom" in k:
                    unit_col = orig

    file_obj.seek(0)
    if hdr_idx is not None and ing_col:
        df = pd.read_excel(file_obj, sheet_name=sheet_name, header=hdr_idx)
    else:
        df = pd.read_excel(file_obj, sheet_name=sheet_name, header=0)
        ing_col = df.columns[0]

    df = df.dropna(subset=[ing_col]).copy()
    stops = {
        "", "nan", "total", "subtotal", "grand total", "sum of weight ingredients (gm)",
        "baking & yield specifications", "baking temperature", "baking time",
        "total yield (pieces)", "total weight", "weight per piece"
    }
    df = df[~df[ing_col].astype(str).str.strip().str.lower().isin(stops)]

    def to_kg(r):
        raw_q = r[qty_col] if (qty_col and qty_col in r) else 0.0
        try:
            q = float(raw_q) if pd.notnull(raw_q) else 0.0
        except (ValueError, TypeError):
            q = 0.0
        u = str(r[unit_col]).strip().lower() if (unit_col and unit_col in r and pd.notnull(r[unit_col])) else "kg"
        if u in ["g", "gm", "gram", "grams", "ml"]:
            return q / 1000.0
        if u in ["pc", "pcs", "piece", "pieces"]:
            return q * 0.015
        return q

    rows = []
    for _, r in df.iterrows():
        ing_nm = str(r[ing_col]).strip()
        q_kg = to_kg(r)
        if q_kg <= 0:
            continue
        cat_nm = str(r[cat_col]).strip().upper() if (cat_col and cat_col in r and pd.notnull(r[cat_col])) else "MIX"
        code, rate = match_rate_and_code(ing_nm, price_lookup)
        rows.append({
            "Category": cat_nm,
            "Ingredient": ing_nm,
            "Item_Code": code,
            "Quantity_KG": round(q_kg, 5),
            "Rate_Per_KG": rate,
        })
    return pd.DataFrame(rows), yield_pcs


def extract_recipe_from_image(image_bytes, mime_type, api_key, price_lookup):
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key)
    prompt = (
        "Read this bakery/sweet recipe image carefully. Extract the recipe title, yield pieces (if shown, else 100), "
        "and every ingredient with category (Base Sheet / Mix / Sugar Syrup) and weight converted to KG "
        "(600 gm -> 0.600, 24 pcs fillo -> 0.360). Return ONLY valid JSON: "
        '{"recipe_title": "string", "yield_pcs": 100, "ingredients": [{"Category": "MIX", "Ingredient": "string", "Quantity_KG": 0.0}]}'
    )
    resp = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[types.Part.from_bytes(data=image_bytes, mime_type=mime_type), prompt],
    )
    txt = re.sub(r"^```(?:json)?|```$", "", resp.text.strip(), flags=re.MULTILINE).strip()
    data = json.loads(txt)
    title = str(data.get("recipe_title") or "PHOTO RECIPE").strip().upper()
    ypcs = int(data.get("yield_pcs") or 100)
    rows = []
    for it in data.get("ingredients", []):
        nm = str(it.get("Ingredient", "")).strip()
        qkg = float(it.get("Quantity_KG", 0.0) or 0.0)
        cat = str(it.get("Category", "MIX")).strip().upper()
        if nm and qkg > 0:
            code, rate = match_rate_and_code(nm, price_lookup)
            rows.append({
                "Category": cat,
                "Ingredient": nm,
                "Item_Code": code,
                "Quantity_KG": round(qkg, 4),
                "Rate_Per_KG": rate,
            })
    return title, ypcs, pd.DataFrame(rows)


def generate_super_colorful_excel(recipe_book, price_df, mrp_mult=4.0, oh_pct=0.05):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    f_ban = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    f_sub = Font(name="Calibri", size=10, italic=True, color="FEF08A")
    f_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    f_dat = Font(name="Calibri", size=10, bold=False, color="0F172A")
    f_bld = Font(name="Calibri", size=10, bold=True, color="0F172A")
    f_dd = Font(name="Calibri", size=13, bold=True, color="7C2D12")
    f_kl = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    f_kv = Font(name="Calibri", size=13, bold=True, color="0F172A")

    thin = Side(border_style="thin", color="CBD5E1")
    med_gold = Side(border_style="medium", color="D97706")
    med_dark = Side(border_style="medium", color="1E1B4B")
    dbl_dark = Side(border_style="double", color="1E1B4B")
    b_c = Border(left=thin, right=thin, top=thin, bottom=thin)
    b_dd = Border(left=med_gold, right=med_gold, top=med_gold, bottom=med_gold)
    b_h = Border(left=thin, right=thin, top=med_dark, bottom=med_dark)
    b_t = Border(left=thin, right=thin, top=thin, bottom=dbl_dark)

    kpi_styles = [
        ("0284C7", "E0F2FE"),
        ("059669", "D1FAE5"),
        ("D97706", "FEF3C7"),
        ("E11D48", "FFE4E6"),
        ("7C3AED", "EDE9FE"),
    ]

    db_rows = []
    recipe_names = list(recipe_book.keys())
    for r_name in recipe_names:
        info = recipe_book[r_name]
        df_i = info["df"].dropna(subset=["Ingredient", "Quantity_KG"]).copy()
        ypcs = int(info.get("yield_pcs", 100))
        for idx, (_, row) in enumerate(df_i.iterrows(), 1):
            ing = str(row["Ingredient"]).strip().upper()
            cat = str(row.get("Category", "MIX")).strip().upper()
            code = str(row.get("Item_Code", "MANUAL")).strip().upper()
            qkg = float(row["Quantity_KG"])
            rate = float(row.get("Rate_Per_KG", 0.0) or 0.0)
            kcal100 = get_nutrients(ing)["kcal"]
            db_rows.append((r_name.upper(), idx, cat, ing, code, qkg, rate, kcal100, ypcs))

    db_end = max(len(db_rows) + 1, 2)
    lk_end = max(len(recipe_names) + 1, 2)
    rl_end = max(len(price_df) + 7, 8) if (price_df is not None and not price_df.empty) else 8

    # TAB 1: 1_Dropdown_Recipe_Viewer
    ws_sel = wb.create_sheet(title="1_Dropdown_Recipe_Viewer")
    ws_sel.sheet_properties.tabColor = "7C3AED"
    ws_sel.views.sheetView[0].showGridLines = False

    ws_sel.merge_cells("B2:J2")
    ws_sel["B2"] = "✨ THE HAZELNUT FACTORY — SUPER COLORFUL DROPDOWN RECIPE & COSTING STUDIO"
    ws_sel["B2"].font = f_ban
    ws_sel["B2"].fill = PatternFill("solid", fgColor="3B0764")
    ws_sel["B2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_sel.row_dimensions[2].height = 34

    ws_sel.merge_cells("B3:J3")
    ws_sel["B3"] = "▼ Cell D5 Dropdown Se Recipe Chunein — Ingredients, Live Rate List Cost, MRP Aur Killori Apne Aap Update Honge!"
    ws_sel["B3"].font = f_sub
    ws_sel["B3"].fill = PatternFill("solid", fgColor="581C87")
    ws_sel["B3"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_sel.row_dimensions[3].height = 20

    ws_sel.row_dimensions[5].height = 28
    ws_sel.merge_cells("B5:C5")
    ws_sel["B5"] = "👇 SELECT RECIPE:"
    ws_sel["B5"].font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    ws_sel["B5"].fill = PatternFill("solid", fgColor="C2410C")
    ws_sel["B5"].alignment = Alignment(horizontal="center", vertical="center")

    ws_sel.merge_cells("D5:F5")
    ws_sel["D5"] = recipe_names[0].upper() if recipe_names else "WALNUT BAKLAVA"
    ws_sel["D5"].font = f_dd
    ws_sel["D5"].fill = PatternFill("solid", fgColor="FEF08A")
    ws_sel["D5"].alignment = Alignment(horizontal="center", vertical="center")
    for cc in (4, 5, 6):
        ws_sel.cell(5, cc).border = b_dd
        ws_sel.cell(5, cc).fill = PatternFill("solid", fgColor="FEF08A")

    ws_sel["G5"] = "MRP Multiplier:"
    ws_sel["G5"].font = f_bld
    ws_sel["G5"].alignment = Alignment(horizontal="right", vertical="center")
    ws_sel["H5"] = float(mrp_mult)
    ws_sel["H5"].font = f_dd
    ws_sel["H5"].fill = PatternFill("solid", fgColor="FEF08A")
    ws_sel["H5"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sel["H5"].number_format = '0.0"x"'
    ws_sel["H5"].border = b_dd

    ws_sel["I5"] = "Overhead %:"
    ws_sel["I5"].font = f_bld
    ws_sel["I5"].alignment = Alignment(horizontal="right", vertical="center")
    ws_sel["J5"] = float(oh_pct)
    ws_sel["J5"].font = f_dd
    ws_sel["J5"].fill = PatternFill("solid", fgColor="FEF08A")
    ws_sel["J5"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sel["J5"].number_format = "0.0%"
    ws_sel["J5"].border = b_dd

    dv = DataValidation(type="list", formula1=f"=5_Recipe_List_Lookup!$A$2:$A${lk_end}", allow_blank=False)
    ws_sel.add_data_validation(dv)
    dv.add(ws_sel["D5"])

    tbl_hdr = 10
    ws_sel.row_dimensions[tbl_hdr].height = 26
    sel_hdrs = [
        "S.NO.", "CATEGORY", "INGREDIENT ITEM", "ITEM CODE",
        "WEIGHT (KG)", "RATE / KG (₹)", "TOTAL COST (₹)", "KILLORI / 100G", "TOTAL KILLORI (KCAL)"
    ]
    for idx_h, h in enumerate(sel_hdrs):
        ci = 2 + idx_h
        c = ws_sel.cell(tbl_hdr, ci, h)
        c.font = f_hdr
        c.fill = PatternFill("solid", fgColor="6D28D9")
        c.border = b_h
        c.alignment = Alignment(horizontal="center" if idx_h not in (1, 2) else "left", vertical="center", wrap_text=True)

    s_r = 11
    max_slots = 20
    for slot in range(1, max_slots + 1):
        rr = s_r + slot - 1
        ws_sel.row_dimensions[rr].height = 20
        k_ex = f'$D$5&"|"&{slot}'

        c_sn = ws_sel.cell(rr, 2, f'=IFERROR(INDEX(4_Recipe_Master_DB!$B$2:$B${db_end},MATCH({k_ex},4_Recipe_Master_DB!$J$2:$J${db_end},0)),"")')
        c_sn.alignment = Alignment(horizontal="center", vertical="center")
        c_ct = ws_sel.cell(rr, 3, f'=IFERROR(INDEX(4_Recipe_Master_DB!$C$2:$C${db_end},MATCH({k_ex},4_Recipe_Master_DB!$J$2:$J${db_end},0)),"")')
        c_ig = ws_sel.cell(rr, 4, f'=IFERROR(INDEX(4_Recipe_Master_DB!$D$2:$D${db_end},MATCH({k_ex},4_Recipe_Master_DB!$J$2:$J${db_end},0)),"")')
        c_cd = ws_sel.cell(rr, 5, f'=IFERROR(INDEX(4_Recipe_Master_DB!$E$2:$E${db_end},MATCH({k_ex},4_Recipe_Master_DB!$J$2:$J${db_end},0)),"")')
        c_cd.alignment = Alignment(horizontal="center", vertical="center")
        c_wt = ws_sel.cell(rr, 6, f'=IFERROR(INDEX(4_Recipe_Master_DB!$F$2:$F${db_end},MATCH({k_ex},4_Recipe_Master_DB!$J$2:$J${db_end},0)),0)')
        c_wt.number_format = "0.000"
        c_rt = ws_sel.cell(rr, 7, f'=IF(D{rr}="","",IFERROR(VLOOKUP(E{rr},3_Colorful_Rate_List!$B$8:$F${rl_end},5,FALSE),IFERROR(INDEX(4_Recipe_Master_DB!$G$2:$G${db_end},MATCH({k_ex},4_Recipe_Master_DB!$J$2:$J${db_end},0)),0)))')
        c_rt.number_format = "₹#,##0.00"
        c_tc = ws_sel.cell(rr, 8, f'=IF(F{rr}=0,0,F{rr}*G{rr})')
        c_tc.number_format = "₹#,##0.00"
        c_k1 = ws_sel.cell(rr, 9, f'=IFERROR(INDEX(4_Recipe_Master_DB!$H$2:$H${db_end},MATCH({k_ex},4_Recipe_Master_DB!$J$2:$J${db_end},0)),0)')
        c_k1.number_format = "0.0"
        c_tk = ws_sel.cell(rr, 10, f'=(F{rr}*10)*I{rr}')
        c_tk.number_format = "#,##0.0"

        for ci in range(2, 11):
            cell = ws_sel.cell(rr, ci)
            cell.font = f_dat
            cell.border = b_c
            if ci >= 6:
                cell.alignment = Alignment(horizontal="right", vertical="center")
            if slot % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F5F3FF")

    e_r = s_r + max_slots - 1
    tot_r = e_r + 1
    ws_sel.row_dimensions[tot_r].height = 24
    ws_sel.cell(tot_r, 4, "SELECTED RECIPE BATCH TOTAL")
    for ci, fmt in [(6, "0.000"), (8, "₹#,##0.00"), (10, "#,##0.0")]:
        ltr = get_column_letter(ci)
        c = ws_sel.cell(tot_r, ci, f"=SUM({ltr}{s_r}:{ltr}{e_r})")
        c.number_format = fmt
        c.alignment = Alignment(horizontal="right", vertical="center")
    for ci in range(2, 11):
        ws_sel.cell(tot_r, ci).font = f_bld
        ws_sel.cell(tot_r, ci).fill = PatternFill("solid", fgColor="DDD6FE")
        ws_sel.cell(tot_r, ci).border = b_t

    ws_sel.row_dimensions[7].height = 18
    ws_sel.row_dimensions[8].height = 26
    kpi_defs = [
        (2, 3, "BATCH YIELD (PCS)", f'=IFERROR(INDEX(5_Recipe_List_Lookup!$B$2:$B${lk_end},MATCH($D$5,5_Recipe_List_Lookup!$A$2:$A${lk_end},0)),0)', "#,##0", kpi_styles[0]),
        (4, 5, "BATCH WEIGHT (KG)", f"=F{tot_r}", "0.000", kpi_styles[1]),
        (6, 7, "TOTAL COST + OH (₹)", f"=H{tot_r}*(1+$J$5)", "₹#,##0.00", kpi_styles[2]),
        (8, 8, "AUTO MRP / PC (₹)", "=IF(B8>0,(F8/B8)*$H$5,0)", "₹#,##0.00", kpi_styles[3]),
        (9, 10, "KILLORI / 100G (KCAL)", f"=IF(D8>0,J{tot_r}/(D8*10),0)", "0.0", kpi_styles[4]),
    ]
    for cs, ce, lbl, form, fmt, (hh, bh) in kpi_defs:
        ws_sel.merge_cells(start_row=7, start_column=cs, end_row=7, end_column=ce)
        ws_sel.merge_cells(start_row=8, start_column=cs, end_row=8, end_column=ce)
        ws_sel.cell(7, cs, lbl).font = f_kl
        ws_sel.cell(7, cs).alignment = Alignment(horizontal="center", vertical="center")
        cv = ws_sel.cell(8, cs, form)
        cv.font = f_kv
        cv.number_format = fmt
        cv.alignment = Alignment(horizontal="center", vertical="center")
        for cc in range(cs, ce + 1):
            ws_sel.cell(7, cc).fill = PatternFill("solid", fgColor=hh)
            ws_sel.cell(7, cc).border = b_c
            ws_sel.cell(8, cc).fill = PatternFill("solid", fgColor=bh)
            ws_sel.cell(8, cc).border = b_c

    ws_sel.conditional_formatting.add(f"H{s_r}:H{e_r}", DataBarRule(start_type="min", end_type="max", color="7C3AED"))
    ws_sel.conditional_formatting.add(f"J{s_r}:J{e_r}", DataBarRule(start_type="min", end_type="max", color="059669"))

    ch_ing = BarChart()
    ch_ing.type = "col"
    ch_ing.style = 10
    ch_ing.title = "INGREDIENT COST BREAKDOWN (SELECTED RECIPE)"
    ch_ing.y_axis.title = "COST (₹)"
    d_ing = Reference(ws_sel, min_col=8, max_col=8, min_row=tbl_hdr, max_row=e_r)
    c_ing_ref = Reference(ws_sel, min_col=4, max_col=4, min_row=s_r, max_row=e_r)
    ch_ing.add_data(d_ing, titles_from_data=True)
    ch_ing.set_categories(c_ing_ref)
    ch_ing.legend = None
    ch_ing.width = 21
    ch_ing.height = 10
    ws_sel.add_chart(ch_ing, f"B{tot_r + 2}")

    for ci, w in enumerate([3, 8, 22, 30, 14, 14, 15, 16, 15, 18], 1):
        ws_sel.column_dimensions[get_column_letter(ci)].width = w

    # TAB 2: 2_All_Recipes_Portfolio
    ws_port = wb.create_sheet(title="2_All_Recipes_Portfolio")
    ws_port.sheet_properties.tabColor = "0284C7"
    ws_port.views.sheetView[0].showGridLines = False

    ws_port.merge_cells("B2:J2")
    ws_port["B2"] = "📊 ALL RECIPES PORTFOLIO COMPARISON — COST, AUTO MRP & KILLORI"
    ws_port["B2"].font = f_ban
    ws_port["B2"].fill = PatternFill("solid", fgColor="0C4A6E")
    ws_port["B2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_port.row_dimensions[2].height = 32

    p_hdrs = [
        "RANK", "RECIPE NAME", "YIELD (PCS)", "BATCH WT (KG)", "RM COST (₹)",
        "TOTAL COST + OH (₹)", "COST / KG (₹)", "AUTO MRP / PC (₹)", "KILLORI / 100G"
    ]
    ws_port.row_dimensions[4].height = 26
    for idx_h, h in enumerate(p_hdrs):
        ci = 2 + idx_h
        c = ws_port.cell(4, ci, h)
        c.font = f_hdr
        c.fill = PatternFill("solid", fgColor="0284C7")
        c.border = b_h
        c.alignment = Alignment(horizontal="center" if idx_h != 1 else "left", vertical="center", wrap_text=True)

    p_end = 4 + len(recipe_names)
    for idx, r_name in enumerate(recipe_names):
        rr = 5 + idx
        ws_port.row_dimensions[rr].height = 21
        info = recipe_book[r_name]
        df_i = info["df"].dropna(subset=["Ingredient", "Quantity_KG"]).copy()
        ypcs = int(info.get("yield_pcs", 100))
        tot_kg = float(df_i["Quantity_KG"].sum())
        rm_c = float((df_i["Quantity_KG"] * df_i["Rate_Per_KG"]).sum())
        tot_k = sum(float(r["Quantity_KG"]) * 10.0 * get_nutrients(r["Ingredient"])["kcal"] for _, r in df_i.iterrows())

        ws_port.cell(rr, 2, f"=RANK(H{rr},$H$5:$H${p_end})").alignment = Alignment(horizontal="center")
        ws_port.cell(rr, 3, r_name.upper()).font = f_bld
        c_y = ws_port.cell(rr, 4, ypcs)
        c_y.number_format = "#,##0"
        c_w = ws_port.cell(rr, 5, round(tot_kg, 3))
        c_w.number_format = "0.000"
        c_rm = ws_port.cell(rr, 6, round(rm_c, 2))
        c_rm.number_format = "₹#,##0.00"
        c_tc = ws_port.cell(rr, 7, f"=F{rr}*(1+'1_Dropdown_Recipe_Viewer'!$J$5)")
        c_tc.number_format = "₹#,##0.00"
        c_ck = ws_port.cell(rr, 8, f"=IF(E{rr}>0,G{rr}/E{rr},0)")
        c_ck.number_format = "₹#,##0.00"
        c_mr = ws_port.cell(rr, 9, f"=IF(D{rr}>0,(G{rr}/D{rr})*'1_Dropdown_Recipe_Viewer'!$H$5,0)")
        c_mr.number_format = "₹#,##0.00"
        c_kc = ws_port.cell(rr, 10, round(tot_k / (tot_kg * 10.0) if tot_kg > 0 else 0.0, 1))
        c_kc.number_format = "0.0"

        for ci in range(2, 11):
            cell = ws_port.cell(rr, ci)
            if ci != 3:
                cell.font = f_dat
            cell.border = b_c
            if idx % 2 == 1:
                cell.fill = PatternFill("solid", fgColor="E0F2FE")

    for ci, w in enumerate([3, 8, 25, 13, 15, 16, 18, 15, 18, 15], 1):
        ws_port.column_dimensions[get_column_letter(ci)].width = w

    # TAB 3: 3_Colorful_Rate_List
    ws_rl = wb.create_sheet(title="3_Colorful_Rate_List")
    ws_rl.sheet_properties.tabColor = "059669"
    ws_rl.views.sheetView[0].showGridLines = True

    ws_rl.merge_cells("A1:G1")
    ws_rl["A1"] = "💰 THE HAZELNUT FACTORY — SUPER COLORFUL RATE LIST MASTER"
    ws_rl["A1"].font = f_ban
    ws_rl["A1"].fill = PatternFill("solid", fgColor="064E3B")
    ws_rl["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_rl.row_dimensions[1].height = 32

    rl_hdr_r = 7
    rl_hdrs = ["S. NO.", "ITEM CODE", "CATEGORY", "DESCRIPTION", "UOM", "UNIT COST (₹)", "COST TIER"]
    ws_rl.row_dimensions[rl_hdr_r].height = 25
    for c_i, h in enumerate(rl_hdrs, 1):
        c = ws_rl.cell(rl_hdr_r, c_i, h)
        c.font = f_hdr
        c.fill = PatternFill("solid", fgColor="059669")
        c.border = b_h
        c.alignment = Alignment(horizontal="center" if c_i != 4 else "left", vertical="center")

    fl_rm = PatternFill("solid", fgColor="DBEAFE")
    fl_pm = PatternFill("solid", fgColor="FCE7F3")
    fl_prem = PatternFill("solid", fgColor="FEE2E2")
    fl_mid = PatternFill("solid", fgColor="FEF3C7")
    fl_eco = PatternFill("solid", fgColor="DCFCE7")

    rl_start = 8
    if price_df is not None and not price_df.empty:
        code_col = price_df.columns[0]
        desc_col = price_df.columns[1] if len(price_df.columns) > 1 else price_df.columns[0]
        uom_col = price_df.columns[2] if len(price_df.columns) > 2 else price_df.columns[0]
        cost_col = price_df.columns[3] if len(price_df.columns) > 3 else price_df.columns[-1]

        for idx, (_, r) in enumerate(price_df.iterrows()):
            rr = rl_start + idx
            ws_rl.row_dimensions[rr].height = 20
            code_v = str(r.get(code_col, "")).strip()
            desc_v = str(r.get(desc_col, "")).strip().upper()
            uom_v = str(r.get(uom_col, "KGS")).strip().upper()
            try:
                cost_v = float(r.get(cost_col, 0.0) or 0.0)
            except (ValueError, TypeError):
                cost_v = 0.0
            is_pm = code_v.startswith("PM")
            cat_v = "PACKAGING (PM)" if is_pm else "RAW MATERIAL (RM)"

            ws_rl.cell(rr, 1, idx + 1).alignment = Alignment(horizontal="center", vertical="center")
            ws_rl.cell(rr, 2, code_v).alignment = Alignment(horizontal="center", vertical="center")
            ccat = ws_rl.cell(rr, 3, cat_v)
            ccat.alignment = Alignment(horizontal="center", vertical="center")
            ccat.fill = fl_pm if is_pm else fl_rm

            ws_rl.cell(rr, 4, desc_v).alignment = Alignment(horizontal="left", vertical="center")
            ws_rl.cell(rr, 5, uom_v).alignment = Alignment(horizontal="center", vertical="center")
            cc = ws_rl.cell(rr, 6, cost_v)
            cc.number_format = "₹#,##0.00"
            cc.alignment = Alignment(horizontal="right", vertical="center")

            ct = ws_rl.cell(rr, 7, f'=IF(F{rr}>=1000,"PREMIUM (>₹1000)",IF(F{rr}>=300,"MID-RANGE","ECONOMY"))')
            ct.alignment = Alignment(horizontal="center", vertical="center")
            ct.fill = fl_prem if cost_v >= 1000 else (fl_mid if cost_v >= 300 else fl_eco)

            for c_i in range(1, 8):
                cell = ws_rl.cell(rr, c_i)
                cell.font = f_bld if c_i in (3, 7) else f_dat
                cell.border = b_c
                if c_i not in (3, 7) and idx % 2 == 1:
                    cell.fill = PatternFill("solid", fgColor="ECFDF5")

    for c_i, w in enumerate([10, 14, 20, 42, 10, 18, 22], 1):
        ws_rl.column_dimensions[get_column_letter(c_i)].width = w

    # TAB 4: 4_Recipe_Master_DB
    ws_db = wb.create_sheet(title="4_Recipe_Master_DB")
    ws_db.sheet_properties.tabColor = "D97706"
    db_hdrs = [
        "RECIPE NAME", "S.NO.", "CATEGORY", "INGREDIENT ITEM", "ITEM CODE",
        "WEIGHT (KG)", "FALLBACK RATE", "KCAL / 100G", "YIELD (PCS)", "LOOKUP KEY"
    ]
    for ci, h in enumerate(db_hdrs, 1):
        c = ws_db.cell(1, ci, h)
        c.font = f_hdr
        c.fill = PatternFill("solid", fgColor="B45309")
        c.border = b_h

    for idx, row_t in enumerate(db_rows, 2):
        for ci, val in enumerate(row_t, 1):
            c = ws_db.cell(idx, ci, val)
            c.border = b_c
            if idx % 2 == 1:
                c.fill = PatternFill("solid", fgColor="FFFBEB")
        ck = ws_db.cell(idx, 10, f'=A{idx}&"|"&B{idx}')
        ck.border = b_c

    for ci in range(1, 11):
        ws_db.column_dimensions[get_column_letter(ci)].width = 20

    # TAB 5: 5_Recipe_List_Lookup
    ws_lk = wb.create_sheet(title="5_Recipe_List_Lookup")
    ws_lk.sheet_properties.tabColor = "E11D48"
    for ci, h in enumerate(["RECIPE NAME", "YIELD (PCS)"], 1):
        c = ws_lk.cell(1, ci, h)
        c.font = f_hdr
        c.fill = PatternFill("solid", fgColor="BE123C")
        c.border = b_h

    for idx, r_name in enumerate(recipe_names, 2):
        ws_lk.cell(idx, 1, r_name.upper()).border = b_c
        ws_lk.cell(idx, 2, int(recipe_book[r_name].get("yield_pcs", 100))).border = b_c
    ws_lk.column_dimensions["A"].width = 28
    ws_lk.column_dimensions["B"].width = 16

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ----------------- HERO HEADER -----------------
st.markdown(
    """
    <div class="hero-banner">
        <h1>🎨 NOVA-AI — PRICE LIST UPLOAD, RECIPE PHOTO/EXCEL & SUPER COLORFUL EXCEL STUDIO</h1>
        <p>Step 1: Rate List Upload Karein • Step 2: Recipe Excel Ya Photo Upload Karein • Step 3: Dropdown Se Dekhein Aur Excel Save Karein</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ----------------- SIDEBAR + MAIN TOP UPLOAD CARDS (VISIBLE EVERYWHERE) -----------------
with st.sidebar:
    st.markdown("## 📂 QUICK UPLOAD CENTER")
    sb_price = st.file_uploader(
        "💰 1. Upload Rate List (.xlsx)",
        type=["xlsx", "xls", "csv"],
        key="sb_price_uploader",
    )
    sb_recipes = st.file_uploader(
        "📄 2. Upload Recipe Excel(s)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="sb_recipe_uploader",
    )
    sb_photo = st.file_uploader(
        "📸 3. Upload Recipe Photo",
        type=["jpg", "jpeg", "png", "webp"],
        key="sb_photo_uploader",
    )

u1, u2, u3 = st.columns(3)

with u1:
    st.markdown(
        '<div class="upload-card-1"><b>💰 STEP 1: RATE LIST UPLOAD KAREIN</b><br>'
        '<span style="font-size:12px;color:#065F46;">Apni Bakery Price List (.xlsx / .csv) yahan upload karein</span></div>',
        unsafe_allow_html=True,
    )
    main_price = st.file_uploader(
        "Choose Rate List File (.xlsx / .csv)",
        type=["xlsx", "xls", "csv"],
        key="main_price_uploader",
    )
    active_price_file = main_price or sb_price
    if active_price_file is not None:
        try:
            if active_price_file.name.lower().endswith(".csv"):
                active_price_file.seek(0)
                p_df = pd.read_csv(active_price_file)
            else:
                p_df = parse_price_list_file(active_price_file)
            st.session_state.price_list_df = p_df
            st.session_state.price_lookup_dict = build_price_lookup(p_df)
            st.success(f"✅ Rate List Loaded: {len(p_df)} rows ({len(st.session_state.price_lookup_dict)} items mapped)!")
        except Exception as e_pl:
            st.error(f"Rate List read error: {e_pl}")

with u2:
    st.markdown(
        '<div class="upload-card-2"><b>📄 STEP 2: RECIPE EXCEL UPLOAD KAREIN</b><br>'
        '<span style="font-size:12px;color:#1E40AF;">1 ya multiple Recipe Excel files upload karke auto-convert karein</span></div>',
        unsafe_allow_html=True,
    )
    main_rec_xls = st.file_uploader(
        "Choose Recipe Excel File(s) (.xlsx)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="main_recipe_uploader",
    )
    active_rec_files = main_rec_xls if main_rec_xls else sb_recipes
    if active_rec_files and st.button("⚡ Excel Recipes Load Karein", use_container_width=True):
        cnt = 0
        for rf in active_rec_files:
            rf.seek(0)
            xl = pd.ExcelFile(rf)
            for s in xl.sheet_names:
                if "price" not in s.lower() and "rate" not in s.lower():
                    df_p, y_p = parse_recipe_excel(rf, s, st.session_state.price_lookup_dict)
                    if not df_p.empty:
                        title_clean = s.replace("Recipe - ", "").strip().upper()
                        st.session_state.recipe_book[title_clean] = {"df": df_p, "yield_pcs": y_p}
                        cnt += 1
        st.success(f"🎉 {cnt} Recipe(s) Dropdown Mein Add Ho Gayi!")

with u3:
    st.markdown(
        '<div class="upload-card-3"><b>📸 STEP 3: RECIPE PHOTO UPLOAD KAREIN</b><br>'
        '<span style="font-size:12px;color:#92400E;">Photo se ingredients aur rates nikal kar Excel banayein</span></div>',
        unsafe_allow_html=True,
    )
    main_photo = st.file_uploader(
        "Choose Recipe Slip Photo (JPG/PNG)",
        type=["jpg", "jpeg", "png", "webp"],
        key="main_photo_uploader",
    )
    active_photo = main_photo or sb_photo
    def_key = ""
    try:
        def_key = st.secrets.get("GOOGLE_API_KEY", "")
    except Exception:
        def_key = ""
    api_key = st.text_input("Gemini API Key (Photo OCR ke liye)", value=def_key, type="password")
    if active_photo and st.button("✨ Photo Se Recipe Excel Banayein", use_container_width=True):
        if not api_key:
            st.warning("Photo OCR ke liye Gemini API Key enter karein.")
        else:
            with st.spinner("Photo se recipe padhi ja rahi hai..."):
                try:
                    t_img, y_img, df_img = extract_recipe_from_image(
                        active_photo.getvalue(),
                        active_photo.type or "image/jpeg",
                        api_key,
                        st.session_state.price_lookup_dict,
                    )
                    if not df_img.empty:
                        st.session_state.recipe_book[t_img] = {"df": df_img, "yield_pcs": y_img}
                        st.success(f"🎨 '{t_img}' ({len(df_img)} items) Dropdown mein jud gaya!")
                except Exception as ex:
                    st.error(f"Photo OCR Error: {ex}")

st.markdown("---")

# ----------------- MAIN INTERACTIVE DROPDOWN VIEWER -----------------
c_sel, c_m1, c_m2 = st.columns([0.45, 0.27, 0.28])
selected_recipe = c_sel.selectbox(
    "👇 DROPDOWN SE RECIPE CHUNEIN:",
    list(st.session_state.recipe_book.keys()),
    key="dd_recipe_selector",
)
mrp_mult_ui = c_m1.number_input("MRP Multiplier (x Cost)", 1.5, 8.0, 4.0, 0.5)
oh_pct_ui = c_m2.number_input("Overhead / Wastage (%)", 0.0, 30.0, 5.0, 0.5)

cur_entry = st.session_state.recipe_book[selected_recipe]
cur_df = cur_entry["df"].copy()

if st.session_state.price_lookup_dict:
    for idx, r in cur_df.iterrows():
        c_code, c_rate = match_rate_and_code(r["Ingredient"], st.session_state.price_lookup_dict)
        if c_rate > 0:
            cur_df.at[idx, "Item_Code"] = c_code
            cur_df.at[idx, "Rate_Per_KG"] = c_rate

ed_col1, ed_col2 = st.columns([0.75, 0.25])
new_r_title = ed_col1.text_input("Recipe Name", value=selected_recipe)
new_r_yield = ed_col2.number_input("Batch Yield (Pieces)", min_value=1, value=int(cur_entry.get("yield_pcs", 100)))

edited_df = st.data_editor(
    cur_df,
    num_rows="dynamic",
    use_container_width=True,
    key=f"editor_{selected_recipe}",
    column_config={
        "Category": st.column_config.TextColumn("Category"),
        "Ingredient": st.column_config.TextColumn("Ingredient", required=True),
        "Item_Code": st.column_config.TextColumn("Rate List Code"),
        "Quantity_KG": st.column_config.NumberColumn("Weight (KG)", min_value=0.0001, format="%.3f"),
        "Rate_Per_KG": st.column_config.NumberColumn("Rate / KG (₹)", min_value=0.0, format="₹%.2f"),
    },
)

if new_r_title.strip().upper() != selected_recipe:
    del st.session_state.recipe_book[selected_recipe]
    selected_recipe = new_r_title.strip().upper()
st.session_state.recipe_book[selected_recipe] = {"df": edited_df, "yield_pcs": new_r_yield}

valid_df = edited_df.dropna(subset=["Quantity_KG", "Rate_Per_KG"]).copy()
valid_df["Total_Cost"] = valid_df["Quantity_KG"] * valid_df["Rate_Per_KG"]
tot_wt = valid_df["Quantity_KG"].sum()
rm_cost = valid_df["Total_Cost"].sum()
tot_cost_oh = rm_cost * (1 + oh_pct_ui / 100.0)
cost_pc = tot_cost_oh / max(new_r_yield, 1)
auto_mrp = cost_pc * mrp_mult_ui
tot_kcal = sum(float(r["Quantity_KG"]) * 10.0 * get_nutrients(r["Ingredient"])["kcal"] for _, r in valid_df.iterrows())
kcal_100g = (tot_kcal / (tot_wt * 10.0)) if tot_wt > 0 else 0.0

k1, k2, k3, k4, k5 = st.columns(5)
k1.markdown(f'<div class="kpi-box" style="background:linear-gradient(135deg,#0284C7,#0369A1);"><div class="kpi-lbl">Batch Yield</div><div class="kpi-num">{new_r_yield:,} Pcs</div><div class="kpi-tag">Finished Output</div></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="kpi-box" style="background:linear-gradient(135deg,#059669,#047857);"><div class="kpi-lbl">Batch Weight</div><div class="kpi-num">{tot_wt:,.3f} KG</div><div class="kpi-tag">Total Net Mix</div></div>', unsafe_allow_html=True)
k3.markdown(f'<div class="kpi-box" style="background:linear-gradient(135deg,#D97706,#B45309);"><div class="kpi-lbl">Total Cost + OH</div><div class="kpi-num">₹{tot_cost_oh:,.2f}</div><div class="kpi-tag">Incl. {oh_pct_ui}% OH</div></div>', unsafe_allow_html=True)
k4.markdown(f'<div class="kpi-box" style="background:linear-gradient(135deg,#E11D48,#BE123C);"><div class="kpi-lbl">Auto MRP / Pc</div><div class="kpi-num">₹{auto_mrp:,.2f}</div><div class="kpi-tag">{mrp_mult_ui}x Multiplier</div></div>', unsafe_allow_html=True)
k5.markdown(f'<div class="kpi-box" style="background:linear-gradient(135deg,#7C3AED,#5B21B6);"><div class="kpi-lbl">Killori / 100g</div><div class="kpi-num">{kcal_100g:,.1f} kcal</div><div class="kpi-tag">Nutrition Density</div></div>', unsafe_allow_html=True)

st.markdown("")
super_excel_bytes = generate_super_colorful_excel(
    st.session_state.recipe_book,
    st.session_state.price_list_df,
    mrp_mult=mrp_mult_ui,
    oh_pct=oh_pct_ui / 100.0,
)
st.download_button(
    label="💾 SUPER COLORFUL DROPDOWN EXCEL BANAYEIN AUR SAVE KAREIN (.xlsx)",
    data=super_excel_bytes,
    file_name=f"THF_Super_Colorful_Recipe_Master_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

ch1, ch2 = st.columns(2)
with ch1:
    fig1 = px.bar(
        valid_df,
        x="Ingredient",
        y="Total_Cost",
        color="Total_Cost",
        title=f"💰 Ingredient Cost Breakdown — {selected_recipe}",
        color_continuous_scale="Turbo",
        height=350,
    )
    st.plotly_chart(fig1, use_container_width=True)

with ch2:
    st.markdown("#### 💰 Uploaded Rate List Sheet (Live Editable)")
    edited_pl_view = st.data_editor(
        st.session_state.price_list_df,
        num_rows="dynamic",
        use_container_width=True,
        height=310,
        key="bottom_rate_list_editor",
    )
    if not edited_pl_view.equals(st.session_state.price_list_df):
        st.session_state.price_list_df = edited_pl_view
        st.session_state.price_lookup_dict = build_price_lookup(edited_pl_view)
