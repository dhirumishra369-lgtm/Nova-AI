import datetime
import io
import json
import math
import re
import uuid

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
try:
    import plotly.express as px
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False
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
        border: 2px solid #059669; border-radius: 12px; padding: 12px 14px; margin-bottom: 8px;
    }
    .upload-card-2 {
        background: linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%);
        border: 2px solid #2563EB; border-radius: 12px; padding: 12px 14px; margin-bottom: 8px;
    }
    .upload-card-3 {
        background: linear-gradient(135deg, #FFFBEB 0%, #FEF3C7 100%);
        border: 2px solid #D97706; border-radius: 12px; padding: 12px 14px; margin-bottom: 8px;
    }
    .kpi-box {
        border-radius: 12px; padding: 14px 16px; color: #FFFFFF;
        box-shadow: 0 4px 10px rgba(0,0,0,0.08);
    }
    .kpi-lbl { font-size: 11px; font-weight: 700; text-transform: uppercase; opacity: 0.92; }
    .kpi-num { font-size: 22px; font-weight: 800; margin-top: 4px; }
    .kpi-tag { font-size: 11px; opacity: 0.88; margin-top: 2px; }
    </style>
    """,
    unsafe_allow_html=True,
)

RECIPE_COLS = ["Category", "Ingredient", "Item_Code", "Quantity_KG", "Rate_Per_KG"]
GEMINI_MODEL = "gemini-2.5-flash"


# ----------------- DATA HELPERS -----------------
def to_float(value, default=0.0):
    try:
        number = float(value)
        return default if math.isnan(number) else number
    except (TypeError, ValueError):
        return default


def normalize_recipe_df(df):
    """Same columns, clean types. Empty rows are kept so a newly added editor row doesn't vanish."""
    df = pd.DataFrame() if df is None else df.copy()
    for c in RECIPE_COLS:
        if c not in df.columns:
            df[c] = None
    df = df[RECIPE_COLS].reset_index(drop=True)
    df["Category"] = df["Category"].fillna("MIX").astype(str).str.strip().str.upper().replace({"": "MIX", "NAN": "MIX"})
    df["Ingredient"] = df["Ingredient"].fillna("").astype(str).str.strip().replace({"nan": "", "None": ""})
    df["Item_Code"] = df["Item_Code"].fillna("MANUAL").astype(str).str.strip().replace({"": "MANUAL", "nan": "MANUAL"})
    df["Quantity_KG"] = pd.to_numeric(df["Quantity_KG"], errors="coerce")
    df["Rate_Per_KG"] = pd.to_numeric(df["Rate_Per_KG"], errors="coerce").fillna(0.0)
    return df


def usable_rows(df):
    """Rows that count for costing: has a name and a weight above zero."""
    df = normalize_recipe_df(df)
    return df[(df["Ingredient"] != "") & (df["Quantity_KG"].fillna(0) > 0)].reset_index(drop=True)


def apply_editor_changes(df, changes):
    """Apply a st.data_editor change-set (edited/deleted/added rows) to the source DataFrame."""
    df = df.reset_index(drop=True).copy()
    for pos, cols in (changes or {}).get("edited_rows", {}).items():
        for col, val in cols.items():
            if col not in df.columns:
                continue
            try:
                df.at[int(pos), col] = val
            except (TypeError, ValueError):
                df[col] = df[col].astype(object)
                df.at[int(pos), col] = val
    deleted = (changes or {}).get("deleted_rows", [])
    if deleted:
        df = df.drop(index=[d for d in deleted if d in df.index])
    added = [{k: v for k, v in row.items() if k != "_index"} for row in (changes or {}).get("added_rows", [])]
    if added:
        df = pd.concat([df, pd.DataFrame(added)], ignore_index=True)
    return df.reset_index(drop=True)


def new_entry(df, yield_pcs, sell_price=None):
    sp = to_float(sell_price, None) if sell_price is not None else None
    return {"df": normalize_recipe_df(df), "yield_pcs": max(int(to_float(yield_pcs, 100) or 100), 1),
            "uid": uuid.uuid4().hex[:8], "sell_price": sp if sp and sp > 0 else None}


# ----------------- SAVE / LOAD (so recipes survive a page refresh) -----------------
STORE_FILE = "nova_recipe_store.json"
EMPTY_PRICE_DF = pd.DataFrame(columns=["No.", "Description", "Base Unit of Measure", "Unit Cost"])


def _records(df):
    clean = df.astype(object).where(df.notna(), None)
    return clean.to_dict("records")


def state_to_json(book, price_df):
    price_df = EMPTY_PRICE_DF if price_df is None else price_df
    return json.dumps({
        "app": "nova-ai-recipe-studio", "version": 2,
        "saved_on": datetime.datetime.now().strftime("%d-%m-%Y %H:%M"),
        "recipes": [{"name": n, "yield_pcs": e.get("yield_pcs", 100), "sell_price": e.get("sell_price"),
                     "rows": _records(normalize_recipe_df(e["df"]))} for n, e in book.items()],
        "price_list": {"columns": [str(c) for c in price_df.columns],
                       "data": [[None if pd.isna(v) else (v if isinstance(v, (int, float, str, bool)) else str(v))
                                 for v in row] for row in price_df.itertuples(index=False)]},
    }, ensure_ascii=False, indent=1)


def json_to_state(text):
    data = json.loads(text)
    book = {}
    for r in data.get("recipes", []):
        name = str(r.get("name", "")).strip().upper()
        if name:
            book[name] = new_entry(pd.DataFrame(r.get("rows", [])), r.get("yield_pcs", 100), r.get("sell_price"))
    pl = data.get("price_list") or {}
    price_df = pd.DataFrame(pl.get("data", []), columns=pl.get("columns") or EMPTY_PRICE_DF.columns)
    return book, price_df


def load_store_file():
    try:
        with open(STORE_FILE, "r", encoding="utf-8") as f:
            book, price_df = json_to_state(f.read())
        return (book, price_df) if book else None
    except (OSError, ValueError, TypeError, KeyError):
        return None


def autosave(book, price_df):
    try:
        payload = state_to_json(book, price_df)
        body = payload.split('"recipes"', 1)[1]
        if body != st.session_state.get("_saved_body"):
            with open(STORE_FILE, "w", encoding="utf-8") as f:
                f.write(payload)
            st.session_state._saved_body = body
    except Exception:
        pass  # saving is best-effort; the Backup button is the safe copy


# ----------------- SESSION STATE -----------------
_loaded = None if "recipe_book" in st.session_state else load_store_file()
if _loaded:
    st.session_state.recipe_book, st.session_state.price_list_df = _loaded
    st.session_state.flash_msg = ("success", f"💾 Pichla saved data load ho gaya: {len(_loaded[0])} recipe(s).")
if "recipe_book" not in st.session_state:
    st.session_state.recipe_book = {
        "WALNUT BAKLAVA": new_entry(pd.DataFrame([
            {"Category": "BASE SHEET", "Ingredient": "Fillo Sheet", "Item_Code": "RM0165", "Quantity_KG": 0.360, "Rate_Per_KG": 521.13},
            {"Category": "MIX", "Ingredient": "Kaju 04 Pcs", "Item_Code": "RM0067", "Quantity_KG": 0.600, "Rate_Per_KG": 841.87},
            {"Category": "MIX", "Ingredient": "Walnut", "Item_Code": "RM0505", "Quantity_KG": 0.350, "Rate_Per_KG": 1150.00},
            {"Category": "MIX", "Ingredient": "(Amul) Salted Butter", "Item_Code": "RM0052", "Quantity_KG": 0.320, "Rate_Per_KG": 537.50},
            {"Category": "MIX", "Ingredient": "Milk Maid", "Item_Code": "RM0300", "Quantity_KG": 0.220, "Rate_Per_KG": 327.88},
            {"Category": "SUGAR SYRUP", "Ingredient": "Sugar (D.F)", "Item_Code": "RM0046", "Quantity_KG": 0.500, "Rate_Per_KG": 65.27},
            {"Category": "SUGAR SYRUP", "Ingredient": "Honey", "Item_Code": "RM0220", "Quantity_KG": 0.150, "Rate_Per_KG": 311.10},
        ]), 215)
    }

defaults = {
    "price_list_df": EMPTY_PRICE_DF.copy(),
    "price_lookup_dict": {},
    "price_file_sig": None,
    "ed_ver": 0,
    "pl_ver": 0,
}
if "flash_msg" not in st.session_state:
    st.session_state.flash_msg = None
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

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


def detect_price_columns(df):
    """Find code / description / UOM / cost columns by header name, not by position."""
    cols = list(df.columns)
    low = {c: str(c).strip().lower() for c in cols}

    def first(pred):
        return next((c for c in cols if pred(low[c])), None)

    desc = first(lambda s: "desc" in s) or first(lambda s: "name" in s) \
        or first(lambda s: "item" in s and "code" not in s and "no" not in s.split())
    code = first(lambda s: "code" in s) or first(lambda s: s in ("no", "no.", "item no", "item no.", "s.no", "sku"))
    cost = first(lambda s: "unit cost" in s) or first(lambda s: "cost" in s) \
        or first(lambda s: "rate" in s) or first(lambda s: "price" in s)
    uom = first(lambda s: "uom" in s or "measure" in s) or first(lambda s: "unit" in s and "cost" not in s)
    return code, desc, uom, cost


def build_price_lookup(df_price):
    lookup = {}
    if df_price is None or df_price.empty:
        return lookup
    code_col, desc_col, _, cost_col = detect_price_columns(df_price)
    if not desc_col or not cost_col:
        return lookup
    for _, row in df_price.iterrows():
        desc = str(row.get(desc_col, "")).strip().lower()
        code = str(row.get(code_col, "MANUAL")).strip() if code_col else "MANUAL"
        if not code or code.lower() == "nan":
            code = "MANUAL"
        cost = to_float(row.get(cost_col))
        if desc and desc != "nan":
            lookup[desc] = (code, cost)
    return lookup


def match_rate_and_code(ing_name, price_lookup):
    clean = str(ing_name).strip().lower()
    if not price_lookup or not clean or clean == "nan":
        return "MANUAL", 0.0
    mapped = ALIAS_MAP.get(clean, clean)
    if mapped in price_lookup:
        return price_lookup[mapped]
    if clean in price_lookup:
        return price_lookup[clean]
    # Partial matching, ignoring very short names that would match almost anything
    for target in (mapped, clean):
        for k, v in price_lookup.items():
            if len(k) >= 3 and (target in k or k in target):
                return v
    return "MANUAL", 0.0


def apply_price_list(df, lookup):
    """Fill Item_Code and Rate_Per_KG from the rate list wherever an ingredient matches."""
    df = normalize_recipe_df(df)
    if not lookup:
        return df
    for i, r in df.iterrows():
        code, rate = match_rate_and_code(r["Ingredient"], lookup)
        if rate > 0:
            df.at[i, "Item_Code"] = code
            df.at[i, "Rate_Per_KG"] = rate
    return df


def set_price_list(df):
    """Store a new rate list and re-price every recipe once (manual rate edits after this are kept)."""
    df = df.reset_index(drop=True)
    st.session_state.price_list_df = df
    lookup = build_price_lookup(df)
    st.session_state.price_lookup_dict = lookup
    for entry in st.session_state.recipe_book.values():
        entry["df"] = apply_price_list(entry["df"], lookup)
    st.session_state.ed_ver += 1
    return lookup


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
    return pd.read_excel(file_obj, header=hdr_row).dropna(how="all")


def parse_recipe_excel(file_obj, sheet_name, price_lookup):
    file_obj.seek(0)
    raw_df = pd.read_excel(file_obj, sheet_name=sheet_name, header=None)
    hdr_idx = None
    yield_pcs = 100

    # Look for "Total Yield" anywhere in the sheet (it is often below the ingredient table)
    for _, row in raw_df.iterrows():
        row_strs = [str(v).strip() for v in row.tolist() if pd.notnull(v)]
        if row_strs and "total yield" in row_strs[0].lower() and len(row_strs) > 1:
            y = to_float(row_strs[1], None)
            if y and y > 0:
                yield_pcs = int(y)
                break

    for r_idx, row in raw_df.iterrows():
        lmap = {str(v).strip().lower() for v in row.tolist() if pd.notnull(v)}
        if any("ingredient" in k or "item" in k for k in lmap):
            hdr_idx = r_idx
            break

    file_obj.seek(0)
    df = pd.read_excel(file_obj, sheet_name=sheet_name, header=hdr_idx if hdr_idx is not None else 0)

    if df is None or df.empty or len(df.columns) == 0:
        return pd.DataFrame(columns=RECIPE_COLS), yield_pcs

    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    cat_col = next((c for kl, c in cols_lower.items() if "cat" in kl or "section" in kl), None)
    ing_col = next((c for kl, c in cols_lower.items() if "ingredient" in kl), None)
    if ing_col is None:
        ing_col = next((c for kl, c in cols_lower.items() if "item" in kl or "material" in kl or "description" in kl), None)
    if ing_col is None:
        ing_col = df.columns[0]

    qty_col = next((c for kl, c in cols_lower.items() if "qty" in kl or "quantity" in kl or "base" in kl or "wt" in kl or "weight" in kl), None)
    unit_col = next((c for kl, c in cols_lower.items() if "unit" in kl or "uom" in kl), None)

    df = df.dropna(subset=[ing_col]).copy()
    stops = {
        "", "nan", "total", "subtotal", "grand total", "sum of weight ingredients (gm)",
        "baking & yield specifications", "baking temperature", "baking time",
        "total yield (pieces)", "total weight", "weight per piece"
    }
    df = df[~df[ing_col].astype(str).str.strip().str.lower().isin(stops)]

    def to_kg(r):
        q = to_float(r[qty_col]) if (qty_col is not None and qty_col in r.index) else 0.0
        u = str(r[unit_col]).strip().lower() if (unit_col is not None and unit_col in r.index and pd.notnull(r[unit_col])) else "kg"
        if u in ["g", "gm", "gms", "gram", "grams", "ml"]:
            return q / 1000.0
        if u in ["pc", "pcs", "piece", "pieces"]:
            return q * 0.015
        return q

    rows = []
    for _, r in df.iterrows():
        ing_nm = str(r[ing_col]).strip()
        if not ing_nm or ing_nm.lower() == "nan":
            continue
        q_kg = to_kg(r)
        if q_kg <= 0:
            continue
        cat_nm = str(r[cat_col]).strip().upper() if (cat_col is not None and pd.notnull(r[cat_col])) else "MIX"
        code, rate = match_rate_and_code(ing_nm, price_lookup)
        rows.append({"Category": cat_nm, "Ingredient": ing_nm, "Item_Code": code,
                     "Quantity_KG": round(q_kg, 5), "Rate_Per_KG": rate})
    return pd.DataFrame(rows, columns=RECIPE_COLS), yield_pcs


def get_secret_api_key():
    for name in ("GEMINI_API_KEY", "GOOGLE_API_KEY"):
        try:
            value = st.secrets.get(name)
        except Exception:
            value = None
        if value:
            return value
    return ""


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
        model=GEMINI_MODEL,
        contents=[types.Part.from_bytes(data=image_bytes, mime_type=mime_type), prompt],
        config={"response_mime_type": "application/json"},
    )
    txt = re.sub(r"^```(?:json)?|```$", "", (resp.text or "").strip(), flags=re.MULTILINE).strip()
    data = json.loads(txt)
    title = str(data.get("recipe_title") or "PHOTO RECIPE").strip().upper()
    ypcs = int(to_float(data.get("yield_pcs"), 100) or 100)
    rows = []
    for it in data.get("ingredients", []):
        nm = str(it.get("Ingredient", "")).strip()
        qkg = to_float(it.get("Quantity_KG"))
        cat = str(it.get("Category", "MIX")).strip().upper()
        if nm and qkg > 0:
            code, rate = match_rate_and_code(nm, price_lookup)
            rows.append({"Category": cat, "Ingredient": nm, "Item_Code": code,
                         "Quantity_KG": round(qkg, 4), "Rate_Per_KG": rate})
    return title, ypcs, pd.DataFrame(rows, columns=RECIPE_COLS)


# ----------------- EXCEL BUILDER -----------------
# Sheet names start with a digit, so Excel needs them quoted inside formulas.
SH_VIEW = "'1_Dropdown_Recipe_Viewer'"
SH_RL = "'3_Colorful_Rate_List'"
SH_DB = "'4_Recipe_Master_DB'"
SH_LK = "'5_Recipe_List_Lookup'"


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

    kpi_styles = [("0284C7", "E0F2FE"), ("059669", "D1FAE5"), ("D97706", "FEF3C7"),
                  ("E11D48", "FFE4E6"), ("7C3AED", "EDE9FE")]

    # Clean every recipe once, so NaN never reaches the workbook
    clean_book = {name.upper(): {"df": usable_rows(info["df"]), "yield_pcs": int(info.get("yield_pcs", 100) or 100)}
                  for name, info in recipe_book.items()}
    recipe_names = list(clean_book.keys())

    db_rows = []
    for r_name in recipe_names:
        info = clean_book[r_name]
        for idx, (_, row) in enumerate(info["df"].iterrows(), 1):
            ing = row["Ingredient"].upper()
            db_rows.append((r_name, idx, row["Category"], ing, row["Item_Code"].upper(),
                            float(row["Quantity_KG"]), float(row["Rate_Per_KG"]),
                            get_nutrients(ing)["kcal"], info["yield_pcs"]))

    db_end = max(len(db_rows) + 1, 2)
    lk_end = max(len(recipe_names) + 1, 2)
    has_prices = price_df is not None and not price_df.empty
    rl_end = max(len(price_df) + 7, 8) if has_prices else 8
    max_ing = max((len(v["df"]) for v in clean_book.values()), default=0)
    max_slots = max(20, max_ing)  # never cut off long recipes

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
    ws_sel["D5"] = recipe_names[0] if recipe_names else ""
    ws_sel["D5"].font = f_dd
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

    dv = DataValidation(type="list", formula1=f"={SH_LK}!$A$2:$A${lk_end}", allow_blank=False)
    ws_sel.add_data_validation(dv)
    dv.add(ws_sel["D5"])

    tbl_hdr = 10
    ws_sel.row_dimensions[tbl_hdr].height = 26
    sel_hdrs = ["S.NO.", "CATEGORY", "INGREDIENT ITEM", "ITEM CODE", "WEIGHT (KG)",
                "RATE / KG (₹)", "TOTAL COST (₹)", "KILLORI / 100G", "TOTAL KILLORI (KCAL)"]
    for idx_h, h in enumerate(sel_hdrs):
        c = ws_sel.cell(tbl_hdr, 2 + idx_h, h)
        c.font = f_hdr
        c.fill = PatternFill("solid", fgColor="6D28D9")
        c.border = b_h
        c.alignment = Alignment(horizontal="center" if idx_h not in (1, 2) else "left", vertical="center", wrap_text=True)

    s_r = 11
    db_rng = lambda col: f"{SH_DB}!${col}$2:${col}${db_end}"  # noqa: E731
    for slot in range(1, max_slots + 1):
        rr = s_r + slot - 1
        ws_sel.row_dimensions[rr].height = 20
        match = f'MATCH($D$5&"|"&{slot},{db_rng("J")},0)'

        ws_sel.cell(rr, 2, f'=IFERROR(INDEX({db_rng("B")},{match}),"")').alignment = Alignment(horizontal="center", vertical="center")
        ws_sel.cell(rr, 3, f'=IFERROR(INDEX({db_rng("C")},{match}),"")')
        ws_sel.cell(rr, 4, f'=IFERROR(INDEX({db_rng("D")},{match}),"")')
        ws_sel.cell(rr, 5, f'=IFERROR(INDEX({db_rng("E")},{match}),"")').alignment = Alignment(horizontal="center", vertical="center")
        ws_sel.cell(rr, 6, f'=IFERROR(INDEX({db_rng("F")},{match}),0)').number_format = "0.000"
        ws_sel.cell(rr, 7, f'=IF(D{rr}="","",IFERROR(VLOOKUP(E{rr},{SH_RL}!$B$8:$F${rl_end},5,FALSE),'
                           f'IFERROR(INDEX({db_rng("G")},{match}),0)))').number_format = "₹#,##0.00"
        ws_sel.cell(rr, 8, f'=IF(F{rr}=0,0,F{rr}*G{rr})').number_format = "₹#,##0.00"
        ws_sel.cell(rr, 9, f'=IFERROR(INDEX({db_rng("H")},{match}),0)').number_format = "0.0"
        ws_sel.cell(rr, 10, f'=(F{rr}*10)*I{rr}').number_format = "#,##0.0"

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
        (2, 3, "BATCH YIELD (PCS)", f'=IFERROR(INDEX({SH_LK}!$B$2:$B${lk_end},MATCH($D$5,{SH_LK}!$A$2:$A${lk_end},0)),0)', "#,##0", kpi_styles[0]),
        (4, 5, "BATCH WEIGHT (KG)", f"=F{tot_r}", "0.000", kpi_styles[1]),
        (6, 7, "TOTAL COST + OH (₹)", f"=H{tot_r}*(1+$J$5)", "₹#,##0.00", kpi_styles[2]),
        (8, 8, "AUTO MRP / PC (₹)", "=IF(B8>0,(F8/B8)*$H$5,0)", "₹#,##0.00", kpi_styles[3]),
        (9, 10, "KILLORI / 100G (KCAL)", f"=IF(D8>0,J{tot_r}/(D8*10),0)", "0.0", kpi_styles[4]),
    ]
    for cs, ce, lbl, form, fmt, (hh, bh) in kpi_defs:
        if ce > cs:
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
    ch_ing.add_data(Reference(ws_sel, min_col=8, max_col=8, min_row=tbl_hdr, max_row=e_r), titles_from_data=True)
    ch_ing.set_categories(Reference(ws_sel, min_col=4, max_col=4, min_row=s_r, max_row=e_r))
    ch_ing.legend = None
    ch_ing.roundedCorners = False
    ch_ing.x_axis.delete = False
    ch_ing.y_axis.delete = False
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

    p_hdrs = ["RANK", "RECIPE NAME", "YIELD (PCS)", "BATCH WT (KG)", "RM COST (₹)",
              "TOTAL COST + OH (₹)", "COST / KG (₹)", "AUTO MRP / PC (₹)", "KILLORI / 100G"]
    ws_port.row_dimensions[4].height = 26
    for idx_h, h in enumerate(p_hdrs):
        c = ws_port.cell(4, 2 + idx_h, h)
        c.font = f_hdr
        c.fill = PatternFill("solid", fgColor="0284C7")
        c.border = b_h
        c.alignment = Alignment(horizontal="center" if idx_h != 1 else "left", vertical="center", wrap_text=True)

    p_end = 4 + len(recipe_names)
    for idx, r_name in enumerate(recipe_names):
        rr = 5 + idx
        ws_port.row_dimensions[rr].height = 21
        info = clean_book[r_name]
        df_i = info["df"]
        tot_kg = float(df_i["Quantity_KG"].sum())
        rm_c = float((df_i["Quantity_KG"] * df_i["Rate_Per_KG"]).sum())
        tot_k = sum(float(r["Quantity_KG"]) * 10.0 * get_nutrients(r["Ingredient"])["kcal"] for _, r in df_i.iterrows())

        ws_port.cell(rr, 2, f"=RANK(H{rr},$H$5:$H${p_end})").alignment = Alignment(horizontal="center")
        ws_port.cell(rr, 3, r_name).font = f_bld
        ws_port.cell(rr, 4, info["yield_pcs"]).number_format = "#,##0"
        ws_port.cell(rr, 5, round(tot_kg, 3)).number_format = "0.000"
        ws_port.cell(rr, 6, round(rm_c, 2)).number_format = "₹#,##0.00"
        ws_port.cell(rr, 7, f"=F{rr}*(1+{SH_VIEW}!$J$5)").number_format = "₹#,##0.00"
        ws_port.cell(rr, 8, f"=IF(E{rr}>0,G{rr}/E{rr},0)").number_format = "₹#,##0.00"
        ws_port.cell(rr, 9, f"=IF(D{rr}>0,(G{rr}/D{rr})*{SH_VIEW}!$H$5,0)").number_format = "₹#,##0.00"
        ws_port.cell(rr, 10, round(tot_k / (tot_kg * 10.0) if tot_kg > 0 else 0.0, 1)).number_format = "0.0"

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

    if has_prices:
        code_col, desc_col, uom_col, cost_col = detect_price_columns(price_df)

        def cell_text(row, col, default=""):
            if col is None:
                return default
            v = row.get(col)
            return default if v is None or (isinstance(v, float) and math.isnan(v)) else str(v).strip()

        for idx, (_, r) in enumerate(price_df.iterrows()):
            rr = 8 + idx
            ws_rl.row_dimensions[rr].height = 20
            code_v = cell_text(r, code_col)
            desc_v = cell_text(r, desc_col).upper()
            uom_v = cell_text(r, uom_col, "KGS").upper()
            cost_v = to_float(r.get(cost_col)) if cost_col is not None else 0.0
            is_pm = code_v.upper().startswith("PM")

            ws_rl.cell(rr, 1, idx + 1).alignment = Alignment(horizontal="center", vertical="center")
            ws_rl.cell(rr, 2, code_v).alignment = Alignment(horizontal="center", vertical="center")
            ccat = ws_rl.cell(rr, 3, "PACKAGING (PM)" if is_pm else "RAW MATERIAL (RM)")
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
    db_hdrs = ["RECIPE NAME", "S.NO.", "CATEGORY", "INGREDIENT ITEM", "ITEM CODE",
               "WEIGHT (KG)", "FALLBACK RATE", "KCAL / 100G", "YIELD (PCS)", "LOOKUP KEY"]
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
        ws_db.cell(idx, 10, f'=A{idx}&"|"&B{idx}').border = b_c

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
        ws_lk.cell(idx, 1, r_name).border = b_c
        ws_lk.cell(idx, 2, clean_book[r_name]["yield_pcs"]).border = b_c
    ws_lk.column_dimensions["A"].width = 28
    ws_lk.column_dimensions["B"].width = 16

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


if not st.session_state.price_lookup_dict and not st.session_state.price_list_df.empty:
    st.session_state.price_lookup_dict = build_price_lookup(st.session_state.price_list_df)


# ----------------- CALLBACKS -----------------
def on_recipe_edit(recipe_name, editor_key):
    entry = st.session_state.recipe_book.get(recipe_name)
    if entry is None:
        return
    entry["df"] = normalize_recipe_df(apply_editor_changes(entry["df"], st.session_state.get(editor_key)))
    st.session_state.ed_ver += 1  # fresh editor on the saved data, so edits are never applied twice


def on_price_edit(editor_key):
    new_df = apply_editor_changes(st.session_state.price_list_df, st.session_state.get(editor_key))
    st.session_state.pl_ver += 1
    set_price_list(new_df)


def on_rename(old_name, name_key):
    book = st.session_state.recipe_book
    new_name = str(st.session_state.get(name_key, "")).strip().upper()
    if not new_name or new_name == old_name:
        st.session_state[name_key] = old_name
        return
    if new_name in book:
        st.session_state[name_key] = old_name
        st.session_state.flash_msg = ("warning", f"'{new_name}' naam ki recipe pehle se hai. Koi doosra naam chunein.")
        return
    st.session_state.recipe_book = {(new_name if k == old_name else k): v for k, v in book.items()}
    st.session_state.dd_recipe_selector = new_name
    st.session_state[name_key] = new_name



def on_yield(recipe_name, yield_key):
    entry = st.session_state.recipe_book.get(recipe_name)
    if entry is not None:
        entry["yield_pcs"] = int(st.session_state.get(yield_key) or 1)


def go_edit(recipe_name):
    st.session_state.dd_recipe_selector = recipe_name


def unique_name(base):
    book = st.session_state.recipe_book
    name, i = base, 2
    while name in book:
        name, i = f"{base} {i}", i + 1
    return name


def on_duplicate(recipe_name):
    src_e = st.session_state.recipe_book.get(recipe_name)
    if src_e is None:
        return
    new_name = unique_name(f"{recipe_name} (COPY)")
    st.session_state.recipe_book[new_name] = new_entry(src_e["df"], src_e["yield_pcs"], src_e.get("sell_price"))
    st.session_state.dd_recipe_selector = new_name
    st.session_state.flash_msg = ("success", f"📄 '{new_name}' ban gayi. Ab ise edit kar sakte hain.")


def on_delete(recipe_name):
    book = st.session_state.recipe_book
    book.pop(recipe_name, None)
    if not book:
        book["NEW RECIPE"] = new_entry(pd.DataFrame([{}]), 100)
    st.session_state.dd_recipe_selector = next(iter(book))
    st.session_state.flash_msg = ("success", f"🗑️ '{recipe_name}' delete ho gayi.")


def on_new_recipe():
    name = str(st.session_state.get("new_recipe_name", "")).strip().upper()
    if not name:
        st.session_state.flash_msg = ("warning", "Nayi recipe ka naam likhein.")
        return
    if name in st.session_state.recipe_book:
        st.session_state.flash_msg = ("warning", f"'{name}' naam ki recipe pehle se hai.")
        return
    st.session_state.recipe_book[name] = new_entry(pd.DataFrame([{}]), st.session_state.get("new_recipe_yield", 100))
    st.session_state.dd_recipe_selector = name
    st.session_state.new_recipe_name = ""
    st.session_state.flash_msg = ("success", f"➕ '{name}' ban gayi. 'Recipe Edit' tab mein ingredients bharein.")


def on_sell_price(recipe_name, key):
    entry = st.session_state.recipe_book.get(recipe_name)
    if entry is not None:
        v = to_float(st.session_state.get(key), 0.0)
        entry["sell_price"] = v if v > 0 else None


def on_reset_price(recipe_name):
    entry = st.session_state.recipe_book.get(recipe_name)
    if entry is not None:
        entry["sell_price"] = None


def on_restore():
    up = st.session_state.get("restore_uploader")
    if up is None:
        st.session_state.flash_msg = ("warning", "Pehle backup (.json) file chunein.")
        return
    try:
        book, price_df = json_to_state(up.getvalue().decode("utf-8"))
    except Exception as e:
        st.session_state.flash_msg = ("warning", f"Backup file padhi nahi ja saki: {e}")
        return
    if not book:
        st.session_state.flash_msg = ("warning", "Backup file mein koi recipe nahi mili.")
        return
    st.session_state.recipe_book = book
    st.session_state.price_list_df = price_df
    st.session_state.price_lookup_dict = build_price_lookup(price_df)
    st.session_state.ed_ver += 1
    st.session_state.pl_ver += 1
    st.session_state.dd_recipe_selector = next(iter(book))
    st.session_state.flash_msg = ("success", f"♻️ Backup restore ho gaya: {len(book)} recipe(s).")


# ----------------- RECIPE SUMMARY HELPERS -----------------
def costed_rows(entry):
    df = usable_rows(entry["df"])
    df["Total_Cost"] = df["Quantity_KG"] * df["Rate_Per_KG"]
    return df


def recipe_summary(entry, mrp_mult, oh_pct):
    df = costed_rows(entry)
    ypcs = max(int(entry.get("yield_pcs", 100) or 1), 1)
    wt = float(df["Quantity_KG"].sum())
    rm = float(df["Total_Cost"].sum())
    cost_oh = rm * (1 + oh_pct / 100.0)
    tot = {k: 0.0 for k in ("kcal", "protein", "fat", "carbs")}
    unknown = []
    for _, r in df.iterrows():
        nut = get_nutrients(r["Ingredient"])
        if not any(k in str(r["Ingredient"]).lower() for k in NUTRIENTS_DB):
            unknown.append(r["Ingredient"])
        for k in tot:
            tot[k] += float(r["Quantity_KG"]) * 10.0 * nut[k]
    per100 = {k: (v / (wt * 10.0) if wt > 0 else 0.0) for k, v in tot.items()}
    cost_pc = cost_oh / ypcs
    mrp_pc = cost_pc * mrp_mult
    sell = entry.get("sell_price") or mrp_pc
    return {
        "yield": ypcs, "weight": wt, "rm_cost": rm, "cost_oh": cost_oh,
        "cost_pc": cost_pc, "mrp_pc": mrp_pc, "sell_pc": sell, "custom_price": bool(entry.get("sell_price")),
        "margin_pct": (sell - cost_pc) / sell * 100 if sell > 0 else 0.0,
        "kcal_100g": per100["kcal"], "per100": per100,
        "per_pc": {k: v / ypcs for k, v in tot.items()}, "g_pc": wt * 1000.0 / ypcs,
        "unknown_nutr": unknown,
        "n_ing": len(df), "missing": df[df["Rate_Per_KG"] <= 0]["Ingredient"].tolist(),
    }


KPI_COLORS = ["#0284C7,#0369A1", "#059669,#047857", "#D97706,#B45309", "#E11D48,#BE123C", "#7C3AED,#5B21B6"]


def render_kpis(s, mrp_mult, oh_pct):
    boxes = [
        ("Batch Yield", f"{s['yield']:,} Pcs", "Finished Output"),
        ("Batch Weight", f"{s['weight']:,.3f} KG", f"{s['n_ing']} Ingredients"),
        ("Total Cost + OH", f"₹{s['cost_oh']:,.2f}", f"Incl. {oh_pct}% OH"),
        ("Auto MRP / Pc", f"₹{s['mrp_pc']:,.2f}", f"Cost/Pc ₹{s['cost_pc']:,.2f} × {mrp_mult}"),
        ("Killori / 100g", f"{s['kcal_100g']:,.1f} kcal", "Nutrition Density"),
    ]
    for col, (lbl, num, tag), grad in zip(st.columns(5), boxes, KPI_COLORS):
        col.markdown(
            f'<div class="kpi-box" style="background:linear-gradient(135deg,{grad});">'
            f'<div class="kpi-lbl">{lbl}</div><div class="kpi-num">{num}</div><div class="kpi-tag">{tag}</div></div>',
            unsafe_allow_html=True,
        )
    if s["missing"]:
        st.warning("⚠️ In ingredients ka rate 0 hai (rate list mein match nahi mila): " + ", ".join(s["missing"]))


ING_VIEW_CONFIG = {
    "Ingredient": st.column_config.TextColumn("Ingredient"),
    "Item_Code": st.column_config.TextColumn("Code"),
    "Quantity_KG": st.column_config.NumberColumn("Weight (KG)", format="%.3f"),
    "Rate_Per_KG": st.column_config.NumberColumn("Rate / KG", format="₹%.2f"),
    "Total_Cost": st.column_config.NumberColumn("Cost", format="₹%.2f"),
}


def plan_to_excel(summary_df, rm_df, detail_df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        for sheet, df, fill in (("1_Production_Summary", summary_df, "6D28D9"),
                                ("2_Raw_Material_Kharid", rm_df, "059669"),
                                ("3_Recipe_Wise_Detail", detail_df, "D97706")):
            df.to_excel(xw, sheet_name=sheet, index=False)
            ws = xw.sheets[sheet]
            for cell in ws[1]:
                cell.font = Font(name="Calibri", bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor=fill)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.row_dimensions[1].height = 30
            for ci, col in enumerate(df.columns, 1):
                ws.column_dimensions[get_column_letter(ci)].width = max(12, min(40, len(str(col)) + 4))
                fmt = "₹#,##0.00" if any(t in str(col) for t in ("₹", "Cost", "Price", "Sales", "Profit", "Rate")) \
                    else ("0.000" if "KG" in str(col) else None)
                if fmt:
                    for row in range(2, len(df) + 2):
                        ws.cell(row, ci).number_format = fmt
            ws.freeze_panes = "A2"
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

# ----------------- SIDEBAR + MAIN TOP UPLOAD CARDS -----------------
with st.sidebar:
    st.markdown("## 📂 QUICK UPLOAD CENTER")
    sb_price = st.file_uploader("💰 1. Upload Rate List (.xlsx)", type=["xlsx", "xls", "csv"], key="sb_price_uploader")
    sb_recipes = st.file_uploader("📄 2. Upload Recipe Excel(s)", type=["xlsx", "xls"],
                                  accept_multiple_files=True, key="sb_recipe_uploader")
    sb_photo = st.file_uploader("📸 3. Upload Recipe Photo", type=["jpg", "jpeg", "png", "webp"], key="sb_photo_uploader")

u1, u2, u3 = st.columns(3)

with u1:
    st.markdown(
        '<div class="upload-card-1"><b>💰 STEP 1: RATE LIST UPLOAD KAREIN</b><br>'
        '<span style="font-size:12px;color:#065F46;">Apni Bakery Price List (.xlsx / .csv) yahan upload karein</span></div>',
        unsafe_allow_html=True,
    )
    main_price = st.file_uploader("Choose Rate List File (.xlsx / .csv)", type=["xlsx", "xls", "csv"], key="main_price_uploader")
    active_price_file = main_price or sb_price
    if active_price_file is not None:
        sig = (active_price_file.name, active_price_file.size)
        # Read the file only when a NEW file arrives, so edits in the rate-list table are not wiped on every click
        if sig != st.session_state.price_file_sig:
            try:
                if active_price_file.name.lower().endswith(".csv"):
                    active_price_file.seek(0)
                    p_df = pd.read_csv(active_price_file)
                else:
                    p_df = parse_price_list_file(active_price_file)
                lookup = set_price_list(p_df)
                st.session_state.price_file_sig = sig
                st.session_state.pl_ver += 1
                st.session_state.flash_msg = ("success", f"✅ Rate List Loaded: {len(p_df)} rows ({len(lookup)} items mapped)! Saari recipes ke rate update ho gaye.")
            except Exception as e_pl:
                st.error(f"Rate List read error: {e_pl}")
        else:
            st.caption(f"✅ Loaded: {active_price_file.name} ({len(st.session_state.price_lookup_dict)} items mapped)")

with u2:
    st.markdown(
        '<div class="upload-card-2"><b>📄 STEP 2: RECIPE EXCEL UPLOAD KAREIN</b><br>'
        '<span style="font-size:12px;color:#1E40AF;">1 ya multiple Recipe Excel files upload karke auto-convert karein</span></div>',
        unsafe_allow_html=True,
    )
    main_rec_xls = st.file_uploader("Choose Recipe Excel File(s) (.xlsx)", type=["xlsx", "xls"],
                                    accept_multiple_files=True, key="main_recipe_uploader")
    active_rec_files = main_rec_xls if main_rec_xls else sb_recipes
    if active_rec_files and st.button("⚡ Excel Recipes Load Karein", width="stretch"):
        cnt = 0
        for rf in active_rec_files:
            try:
                rf.seek(0)
                xl = pd.ExcelFile(rf)
                for s in xl.sheet_names:
                    if "price" in s.lower() or "rate" in s.lower():
                        continue
                    df_p, y_p = parse_recipe_excel(rf, s, st.session_state.price_lookup_dict)
                    if not df_p.empty:
                        title_clean = s.replace("Recipe - ", "").strip().upper()
                        st.session_state.recipe_book[title_clean] = new_entry(df_p, y_p)
                        cnt += 1
            except Exception as e_rx:
                st.error(f"'{rf.name}' read nahi ho payi: {e_rx}")
        st.session_state.ed_ver += 1
        st.success(f"🎉 {cnt} Recipe(s) Dropdown Mein Add Ho Gayi!")

with u3:
    st.markdown(
        '<div class="upload-card-3"><b>📸 STEP 3: RECIPE PHOTO UPLOAD KAREIN</b><br>'
        '<span style="font-size:12px;color:#92400E;">Photo se ingredients aur rates nikal kar Excel banayein</span></div>',
        unsafe_allow_html=True,
    )
    main_photo = st.file_uploader("Choose Recipe Slip Photo (JPG/PNG)", type=["jpg", "jpeg", "png", "webp"], key="main_photo_uploader")
    active_photo = main_photo or sb_photo

    # The secret key stays on the server; it is never sent to the browser
    api_key = get_secret_api_key()
    if api_key:
        st.caption("🔐 Gemini API key app ke Secrets se li ja rahi hai.")
    else:
        api_key = st.text_input("Gemini API Key (Photo OCR ke liye)", type="password")

    if active_photo and st.button("✨ Photo Se Recipe Excel Banayein", width="stretch"):
        if not api_key:
            st.warning("Photo OCR ke liye Gemini API Key enter karein.")
        else:
            with st.spinner("Photo se recipe padhi ja rahi hai..."):
                try:
                    t_img, y_img, df_img = extract_recipe_from_image(
                        active_photo.getvalue(), active_photo.type or "image/jpeg",
                        api_key, st.session_state.price_lookup_dict,
                    )
                    if df_img.empty:
                        st.warning("Photo mein koi ingredient nahi mila. Saaf photo ke saath dobara try karein.")
                    else:
                        st.session_state.recipe_book[t_img] = new_entry(df_img, y_img)
                        st.session_state.dd_recipe_selector = t_img
                        st.session_state.ed_ver += 1
                        st.success(f"🎨 '{t_img}' ({len(df_img)} items) Dropdown mein jud gaya!")
                except json.JSONDecodeError:
                    st.error("AI ka jawab samajh nahi aaya. Ek baar aur try karein.")
                except Exception as ex:
                    st.error(f"Photo OCR Error: {ex}")

if st.session_state.flash_msg:
    kind, msg = st.session_state.flash_msg
    (st.success if kind == "success" else st.warning)(msg)
    st.session_state.flash_msg = None

st.markdown("---")

# ----------------- GLOBAL SETTINGS + EXCEL DOWNLOAD -----------------
book = st.session_state.recipe_book
if not book:
    book["NEW RECIPE"] = new_entry(pd.DataFrame([{}]), 100)
recipe_options = list(book.keys())

g1, g2, g3 = st.columns([0.22, 0.22, 0.56])
mrp_mult_ui = g1.number_input("MRP Multiplier (x Cost)", 1.5, 8.0, 4.0, 0.5, key="mrp_mult")
oh_pct_ui = g2.number_input("Overhead / Wastage (%)", 0.0, 30.0, 5.0, 0.5, key="oh_pct")
with g3:
    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
    try:
        st.download_button(
            label=f"💾 SAARI {len(book)} RECIPES KA SUPER COLORFUL DROPDOWN EXCEL (.xlsx)",
            data=generate_super_colorful_excel(book, st.session_state.price_list_df,
                                               mrp_mult=mrp_mult_ui, oh_pct=oh_pct_ui / 100.0),
            file_name=f"THF_Super_Colorful_Recipe_Master_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch", type="primary",
        )
    except Exception as e_xl:
        st.error(f"Excel banane mein error: {e_xl}")

tab_view, tab_edit, tab_plan, tab_rates = st.tabs(
    ["📖 Recipe Viewer", "✏️ Recipe Edit", "🏭 Production Planner", "💰 Rate List"])

MONEY = lambda: st.column_config.NumberColumn(format="₹%.2f")  # noqa: E731

# ----------------- TAB 1: RECIPE VIEWER -----------------
ALL_OPTION = "🌟 SAARI RECIPES EK SAATH DEKHEIN"
with tab_view:
    if st.session_state.get("view_selector") not in [ALL_OPTION] + recipe_options:
        st.session_state.view_selector = ALL_OPTION
    view_choice = st.selectbox(
        f"👇 DROPDOWN SE RECIPE CHUNEIN ({len(recipe_options)} recipes):",
        [ALL_OPTION] + recipe_options, key="view_selector",
    )

    if view_choice == ALL_OPTION:
        summaries = {n: recipe_summary(book[n], mrp_mult_ui, oh_pct_ui) for n in recipe_options}
        port = pd.DataFrame([{
            "Recipe": n, "Yield (Pcs)": s["yield"], "Ingredients": s["n_ing"],
            "Batch Wt (KG)": s["weight"], "Cost + OH": s["cost_oh"], "Cost / Pc": s["cost_pc"],
            "MRP / Pc": s["mrp_pc"], "Selling Price / Pc": s["sell_pc"], "Margin %": s["margin_pct"],
            "Kcal / 100g": s["kcal_100g"],
        } for n, s in summaries.items()])

        t1, t2, t3, t4 = st.columns(4)
        t1.metric("Kul Recipes", len(port))
        t2.metric("Sabse Zyada Margin", port.loc[port["Margin %"].idxmax(), "Recipe"] if len(port) else "-")
        t3.metric("Average Margin", f"{port['Margin %'].mean():.1f}%" if len(port) else "-")
        t4.metric("Kul Batch Cost", f"₹{port['Cost + OH'].sum():,.2f}")

        f1, f2 = st.columns([0.6, 0.4])
        search = f1.text_input("🔍 Recipe dhundhein", key="view_search", placeholder="jaise: baklava")
        sort_by = f2.selectbox("↕️ Sort karein", ["Naam (A-Z)", "Margin % (kam se zyada)", "Cost / Pc (zyada se kam)",
                                                 "Selling Price / Pc (zyada se kam)"], key="view_sort")
        shown = port[port["Recipe"].str.contains(search.strip(), case=False, regex=False)] if search.strip() else port
        sort_map = {"Naam (A-Z)": ("Recipe", True), "Margin % (kam se zyada)": ("Margin %", True),
                    "Cost / Pc (zyada se kam)": ("Cost / Pc", False),
                    "Selling Price / Pc (zyada se kam)": ("Selling Price / Pc", False)}
        s_col, s_asc = sort_map[sort_by]
        shown = shown.sort_values(s_col, ascending=s_asc)

        st.dataframe(
            shown, hide_index=True, width="stretch",
            column_config={
                "Batch Wt (KG)": st.column_config.NumberColumn(format="%.3f"),
                "Cost + OH": MONEY(), "Cost / Pc": MONEY(), "MRP / Pc": MONEY(), "Selling Price / Pc": MONEY(),
                "Margin %": st.column_config.ProgressColumn("Margin %", format="%.1f%%", min_value=0, max_value=100),
                "Kcal / 100g": st.column_config.NumberColumn(format="%.1f"),
            },
        )
        low = shown[shown["Margin %"] < 50]["Recipe"].tolist()
        if low:
            st.warning("📉 In recipes ka margin 50% se kam hai, price check karein: " + ", ".join(low))

        if len(shown) and HAS_PLOTLY:
            fig_all = px.bar(shown, x="Recipe", y=["Cost / Pc", "Selling Price / Pc"], barmode="group",
                             title="💰 Har Recipe: Cost / Pc vs Selling Price / Pc", height=380,
                             color_discrete_sequence=["#D97706", "#059669"])
            fig_all.update_layout(yaxis_title="₹ per piece", legend_title="")
            st.plotly_chart(fig_all, width="stretch")
        elif len(shown):
            st.bar_chart(shown.set_index("Recipe")[["Cost / Pc", "Selling Price / Pc"]])

        st.markdown("#### 📚 Har Recipe Ki Poori Detail")
        for n in shown["Recipe"]:
            s = summaries[n]
            with st.expander(f"🍰 {n}  |  {s['yield']} pcs  |  Cost/Pc ₹{s['cost_pc']:,.2f}  |  "
                             f"Price/Pc ₹{s['sell_pc']:,.2f}  |  Margin {s['margin_pct']:.0f}%"):
                df_n = costed_rows(book[n])
                st.dataframe(df_n[["Category", "Ingredient", "Item_Code", "Quantity_KG", "Rate_Per_KG", "Total_Cost"]],
                             hide_index=True, width="stretch", column_config=ING_VIEW_CONFIG)
                st.button("✏️ Is recipe ko edit karein", key=f"goedit_{book[n]['uid']}",
                          on_click=go_edit, args=(n,), help="Phir upar '✏️ Recipe Edit' tab kholein")

    else:
        entry = book[view_choice]
        s = recipe_summary(entry, mrp_mult_ui, oh_pct_ui)
        st.markdown(f"### 🍰 {view_choice}")
        render_kpis(s, mrp_mult_ui, oh_pct_ui)
        st.markdown("")

        vdf = costed_rows(entry)
        left, right = st.columns([0.58, 0.42])
        with left:
            if vdf.empty:
                st.info("Is recipe mein abhi koi ingredient nahi hai. '✏️ Recipe Edit' tab mein bharein.")
            for cat, g in vdf.groupby("Category", sort=False):
                st.markdown(f"**🔸 {cat}** — ₹{g['Total_Cost'].sum():,.2f}  ({g['Quantity_KG'].sum():.3f} KG)")
                st.dataframe(g[["Ingredient", "Item_Code", "Quantity_KG", "Rate_Per_KG", "Total_Cost"]],
                             hide_index=True, width="stretch", column_config=ING_VIEW_CONFIG)
        with right:
            if not vdf.empty and vdf["Total_Cost"].sum() > 0:
                if HAS_PLOTLY:
                    fig_pie = px.pie(vdf, names="Ingredient", values="Total_Cost", hole=0.45,
                                     title="💰 Cost Mein Kiska Kitna Hissa", height=360,
                                     color_discrete_sequence=px.colors.qualitative.Bold)
                    fig_pie.update_traces(textposition="inside", textinfo="percent+label")
                    fig_pie.update_layout(showlegend=False)
                    st.plotly_chart(fig_pie, width="stretch")
                else:
                    st.bar_chart(vdf.set_index("Ingredient")["Total_Cost"])
            st.button("✏️ Is recipe ko edit karein", key="goedit_single",
                      on_click=go_edit, args=(view_choice,), help="Phir upar '✏️ Recipe Edit' tab kholein")

        # ---- Pricing & profit calculator ----
        st.markdown("#### 💹 Pricing & Profit Calculator")
        sp_key = f"sp_{entry['uid']}"
        if not entry.get("sell_price") or sp_key not in st.session_state:
            st.session_state[sp_key] = round(float(entry.get("sell_price") or s["mrp_pc"]), 2)
        p1, p2, p3 = st.columns([0.35, 0.35, 0.3])
        p1.number_input("Aapka Selling Price / Pc (₹)", min_value=0.0, step=1.0, key=sp_key,
                        on_change=on_sell_price, args=(view_choice, sp_key),
                        help="Badal kar Enter dabayein. Ye price save ho jayega aur Planner mein bhi use hoga.")
        target_fc = p2.slider("Target Food Cost %", 10, 60, 30, 1, key="target_fc",
                              help="Bakery mein aam taur par 25–35% food cost rakha jaata hai.")
        with p3:
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            st.button("↩️ Auto MRP par wapas", on_click=on_reset_price, args=(view_choice,),
                      disabled=not entry.get("sell_price"), width="stretch")

        sp = float(st.session_state[sp_key])
        cost_pc = s["cost_pc"]
        profit_pc = sp - cost_pc
        food_cost = cost_pc / sp * 100 if sp > 0 else 0.0
        suggested = cost_pc / (target_fc / 100.0)
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Profit / Pc", f"₹{profit_pc:,.2f}")
        m2.metric("Profit Margin", f"{(profit_pc / sp * 100 if sp > 0 else 0):.1f}%")
        m3.metric("Food Cost %", f"{food_cost:.1f}%", delta=f"{food_cost - target_fc:+.1f}% target se", delta_color="inverse")
        m4.metric(f"{target_fc}% food cost par price", f"₹{suggested:,.2f}")
        if sp <= 0:
            st.info("Selling price daalein.")
        elif profit_pc < 0:
            st.error(f"🚨 Is price par har piece par ₹{-profit_pc:,.2f} ka nuksaan hai!")
        elif food_cost <= target_fc:
            st.success(f"✅ Food cost target ({target_fc}%) ke andar hai. Batch ({s['yield']} pcs) par profit: ₹{profit_pc * s['yield']:,.2f}")
        else:
            st.warning(f"⚠️ Food cost target se zyada hai. {target_fc}% ke liye price kam se kam ₹{suggested:,.2f} rakhein.")

        # ---- Nutrition panel ----
        st.markdown("#### 🥗 Nutrition (andaza)")
        n1, n2, n3, n4, n5 = st.columns(5)
        n1.metric("Weight / Pc", f"{s['g_pc']:.1f} g")
        n2.metric("Kcal / Pc", f"{s['per_pc']['kcal']:.0f}", f"{s['per100']['kcal']:.0f} / 100g", delta_color="off")
        n3.metric("Protein / Pc", f"{s['per_pc']['protein']:.1f} g", f"{s['per100']['protein']:.1f} g / 100g", delta_color="off")
        n4.metric("Fat / Pc", f"{s['per_pc']['fat']:.1f} g", f"{s['per100']['fat']:.1f} g / 100g", delta_color="off")
        n5.metric("Carbs / Pc", f"{s['per_pc']['carbs']:.1f} g", f"{s['per100']['carbs']:.1f} g / 100g", delta_color="off")
        note = "Ye values aam ingredient data se nikale gaye andaze hain, lab test nahi."
        if s["unknown_nutr"]:
            note += " In ingredients ka data nahi mila, isliye average maana gaya: " + ", ".join(s["unknown_nutr"])
        st.caption(note)

# ----------------- TAB 2: RECIPE EDIT -----------------
with tab_edit:
    if st.session_state.get("dd_recipe_selector") not in recipe_options:
        st.session_state.dd_recipe_selector = recipe_options[0]
    selected_recipe = st.selectbox("👇 EDIT KARNE KE LIYE RECIPE CHUNEIN:", recipe_options, key="dd_recipe_selector")

    a1, a2, a3 = st.columns(3)
    with a1.popover("➕ Nayi recipe banayein", width="stretch"):
        st.session_state.setdefault("new_recipe_yield", 100)
        st.text_input("Recipe ka naam", key="new_recipe_name", placeholder="jaise: KAJU KATLI")
        st.number_input("Batch Yield (Pieces)", min_value=1, step=1, key="new_recipe_yield")
        st.button("Banayein", type="primary", on_click=on_new_recipe, key="btn_new_recipe")
    a2.button("📄 Duplicate (copy) banayein", on_click=on_duplicate, args=(selected_recipe,),
              width="stretch", help="Variation banane ke liye, jaise kam sugar wala version")
    with a3.popover("🗑️ Recipe delete karein", width="stretch"):
        st.write(f"**{selected_recipe}** hamesha ke liye hat jayegi.")
        st.button("Haan, delete karo", type="primary", on_click=on_delete, args=(selected_recipe,), key="btn_confirm_delete")

    cur_entry = book[selected_recipe]
    uid = cur_entry.setdefault("uid", uuid.uuid4().hex[:8])
    name_key, yield_key = f"name_{uid}", f"yield_{uid}"
    if st.session_state.get(name_key) != selected_recipe:
        st.session_state[name_key] = selected_recipe
    if yield_key not in st.session_state:
        st.session_state[yield_key] = int(cur_entry.get("yield_pcs", 100) or 100)

    ed_col1, ed_col2 = st.columns([0.75, 0.25])
    ed_col1.text_input("Recipe Name (badal kar Enter dabayein)", key=name_key,
                       on_change=on_rename, args=(selected_recipe, name_key))
    ed_col2.number_input("Batch Yield (Pieces)", min_value=1, step=1, key=yield_key,
                         on_change=on_yield, args=(selected_recipe, yield_key))

    editor_key = f"editor_{uid}_{st.session_state.ed_ver}"
    st.data_editor(
        cur_entry["df"],
        num_rows="dynamic", hide_index=True, width="stretch",
        key=editor_key, on_change=on_recipe_edit, args=(selected_recipe, editor_key),
        column_config={
            "Category": st.column_config.TextColumn("Category"),
            "Ingredient": st.column_config.TextColumn("Ingredient"),
            "Item_Code": st.column_config.TextColumn("Rate List Code"),
            "Quantity_KG": st.column_config.NumberColumn("Weight (KG)", min_value=0.0, format="%.3f"),
            "Rate_Per_KG": st.column_config.NumberColumn("Rate / KG (₹)", min_value=0.0, format="₹%.2f"),
        },
    )
    st.caption("Table mein badlav turant save hote hain aur Viewer, Planner aur Excel sab mein dikhte hain. "
               "Nayi row ke liye table ke neeche ➕ dabayein.")
    render_kpis(recipe_summary(cur_entry, mrp_mult_ui, oh_pct_ui), mrp_mult_ui, oh_pct_ui)

# ----------------- TAB 3: PRODUCTION PLANNER -----------------
with tab_plan:
    st.markdown("#### 🏭 Production Planner — kitne pieces banane hain?")
    st.caption("Recipes aur pieces chunein. App har recipe ki quantity scale karke poori raw material kharid list, "
               "cost aur expected profit bata dega.")
    st.session_state.plan_recipes = [r for r in st.session_state.get("plan_recipes", []) if r in recipe_options]
    plan_sel = st.multiselect("Aaj kaun si recipes banani hain?", recipe_options, key="plan_recipes",
                              placeholder="Recipes chunein...")

    if not plan_sel:
        st.info("👆 Upar se ek ya zyada recipes chunein.")
    else:
        qty = {}
        q_cols = st.columns(min(len(plan_sel), 4))
        for i, n in enumerate(plan_sel):
            e = book[n]
            k = f"plan_qty_{e['uid']}"
            st.session_state.setdefault(k, int(e["yield_pcs"]))
            qty[n] = q_cols[i % len(q_cols)].number_input(f"{n} (pcs)", min_value=0, step=10, key=k)

        summary_rows, detail_parts = [], []
        for n in plan_sel:
            e = book[n]
            s = recipe_summary(e, mrp_mult_ui, oh_pct_ui)
            factor = qty[n] / s["yield"]
            d = costed_rows(e)
            d["Quantity_KG"] = d["Quantity_KG"] * factor
            d["Total_Cost"] = d["Total_Cost"] * factor
            d.insert(0, "Recipe", n)
            detail_parts.append(d)
            cost_total = s["cost_oh"] * factor
            sales = s["sell_pc"] * qty[n]
            summary_rows.append({
                "Recipe": n, "Pieces": qty[n], "Batches": round(factor, 2),
                "Total Wt (KG)": s["weight"] * factor, "Cost + OH (₹)": cost_total,
                "Cost / Pc (₹)": s["cost_pc"], "Price / Pc (₹)": s["sell_pc"],
                "Sales Value (₹)": sales, "Profit (₹)": sales - cost_total,
            })
        plan_df = pd.DataFrame(summary_rows)
        detail = pd.concat(detail_parts, ignore_index=True) if detail_parts else pd.DataFrame()

        if detail.empty:
            rm_df = pd.DataFrame(columns=["Ingredient", "Item_Code", "Quantity_KG", "Rate_Per_KG", "Total_Cost", "Used In"])
        else:
            detail["_key"] = detail["Ingredient"].str.upper()
            rm_df = (detail.groupby("_key", as_index=False)
                     .agg(Item_Code=("Item_Code", "first"), Quantity_KG=("Quantity_KG", "sum"),
                          Total_Cost=("Total_Cost", "sum"),
                          **{"Used In": ("Recipe", lambda x: ", ".join(sorted(set(x))))})
                     .rename(columns={"_key": "Ingredient"}))
            rm_df["Rate_Per_KG"] = rm_df.apply(lambda r: r["Total_Cost"] / r["Quantity_KG"] if r["Quantity_KG"] > 0 else 0.0, axis=1)
            rm_df = rm_df[["Ingredient", "Item_Code", "Quantity_KG", "Rate_Per_KG", "Total_Cost", "Used In"]] \
                .sort_values("Total_Cost", ascending=False)
            detail = detail.drop(columns="_key")

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Kul Pieces", f"{int(plan_df['Pieces'].sum()):,}")
        k2.metric("Raw Material Kharid", f"₹{rm_df['Total_Cost'].sum():,.2f}")
        k3.metric("Expected Sales", f"₹{plan_df['Sales Value (₹)'].sum():,.2f}")
        k4.metric("Expected Profit", f"₹{plan_df['Profit (₹)'].sum():,.2f}")

        st.markdown("**📋 Production Summary**")
        st.dataframe(plan_df, hide_index=True, width="stretch", column_config={
            "Total Wt (KG)": st.column_config.NumberColumn(format="%.3f"),
            **{c: MONEY() for c in ["Cost + OH (₹)", "Cost / Pc (₹)", "Price / Pc (₹)", "Sales Value (₹)", "Profit (₹)"]},
        })

        st.markdown("**🛒 Raw Material Kharid List (saari recipes mila kar)**")
        st.dataframe(rm_df, hide_index=True, width="stretch", column_config={
            "Quantity_KG": st.column_config.NumberColumn("Quantity (KG)", format="%.3f"),
            "Rate_Per_KG": st.column_config.NumberColumn("Rate / KG", format="₹%.2f"),
            "Total_Cost": st.column_config.NumberColumn("Cost", format="₹%.2f"),
        })
        missing_plan = rm_df[rm_df["Rate_Per_KG"] <= 0]["Ingredient"].tolist()
        if missing_plan:
            st.warning("⚠️ In items ka rate 0 hai, isliye kharid cost poori nahi hai: " + ", ".join(missing_plan))

        with st.expander("👩‍🍳 Kitchen Sheet — har recipe ki scaled quantity"):
            for n in plan_sel:
                part = detail[detail["Recipe"] == n] if not detail.empty else detail
                st.markdown(f"**{n}** — {qty[n]} pcs")
                if part.empty:
                    st.caption("Koi ingredient nahi.")
                else:
                    st.dataframe(part[["Category", "Ingredient", "Quantity_KG", "Total_Cost"]], hide_index=True,
                                 width="stretch", column_config=ING_VIEW_CONFIG)

        st.download_button(
            "📥 Production Plan + Kharid List Excel Download",
            data=plan_to_excel(
                plan_df,
                rm_df.rename(columns={"Quantity_KG": "Quantity KG", "Rate_Per_KG": "Rate / KG (₹)", "Total_Cost": "Cost (₹)"}),
                detail.rename(columns={"Quantity_KG": "Quantity KG", "Rate_Per_KG": "Rate / KG (₹)", "Total_Cost": "Cost (₹)"}),
            ),
            file_name=f"THF_Production_Plan_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch", type="primary",
        )

# ----------------- TAB 4: RATE LIST -----------------
with tab_rates:
    st.markdown("#### 💰 Uploaded Rate List Sheet (Live Editable)")
    pl_key = f"rate_list_editor_{st.session_state.pl_ver}"
    st.data_editor(
        st.session_state.price_list_df,
        num_rows="dynamic", hide_index=True, width="stretch", height=420,
        key=pl_key, on_change=on_price_edit, args=(pl_key,),
    )
    st.caption("Rate list mein badlav karte hi saari recipes ke matching rates update ho jaate hain.")

# ----------------- SIDEBAR: BACKUP & RESTORE -----------------
with st.sidebar:
    st.markdown("---")
    st.markdown("## 💾 BACKUP & RESTORE")
    st.caption("Recipes apne aap save hoti hain, par app reboot hone par server ki file mit sakti hai. "
               "Isliye roz ek backup download kar lein.")
    st.download_button(
        "⬇️ Backup Download (.json)",
        data=state_to_json(book, st.session_state.price_list_df),
        file_name=f"nova_recipes_backup_{datetime.date.today().strftime('%Y%m%d')}.json",
        mime="application/json", width="stretch",
    )
    st.file_uploader("♻️ Backup file chunein", type=["json"], key="restore_uploader")
    st.button("♻️ Backup Restore Karein", on_click=on_restore, width="stretch",
              help="Abhi ki saari recipes backup wali recipes se badal jayengi.")

autosave(book, st.session_state.price_list_df)
