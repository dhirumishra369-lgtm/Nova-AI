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


def new_entry(df, yield_pcs):
    return {"df": normalize_recipe_df(df), "yield_pcs": int(yield_pcs), "uid": uuid.uuid4().hex[:8]}


#
