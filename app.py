
import re, io, unicodedata, calendar
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime as DT

st.set_page_config(page_title="DOC Hesap", layout="wide")
st.title("📦 Days of Coverage (DOC) Hesaplayıcı")

# ------------ Yardımcılar ------------
def norm_text(s: str) -> str:
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower()
    s = re.sub(r"\s+", " ", s)
    return s

KF_PATTERNS = {
    "consensus": [
        "kisit siz consensus","consensus",
        "kisit siz consensus sell in forecast / malzeme tuketim mik",
        "kisit siz consensus forecast / malzeme tuketim mik",
        "kisit siz consensus sell in forecast / malzeme tuketim mik.",
        "kısıtsız consensus sell-in forecast / malzeme tüketim mik",
        "kısıtsız consensus sell-in forecast / malzeme tüketim mik."
    ],
    "beginning_stock": ["baslangic stok","beginning stock"],
    "transport_receipt": ["transport receipt"],
    "recommended_order": ["recommended order"],
    "projected_stock": [
        "unconstrained projected stock","projected stock","unconstrainded projected stock"
    ],
    "doc": ["unconstrained days of coverage","days of coverage"]
}

def classify_kf(val):
    v = norm_text(val)
    for key, pats in KF_PATTERNS.items():
        for p in pats:
            if p in v: return key
    return None

def detect_month_columns_by_parsing(df: pd.DataFrame):
    month_cols = []
    for c in df.columns:
        s = str(c).strip()
        m = re.match(r"^(\d{4}[-/]\d{2}[-/]\d{2})", s)
        if m:
            ts = pd.to_datetime(m.group(1), errors="coerce")
            if pd.notna(ts):
                month_cols.append((c, pd.Timestamp(ts.year, ts.month, 1)))
    month_cols.sort(key=lambda x: x[1])
    return month_cols

# ------------ Dosya yükleme ------------
uploaded = st.file_uploader("Excel'i sürükleyip bırakın", type=["xlsx","xls"])

if uploaded is None:
    st.info("Başlamak için bir Excel yükleyin.")
    st.stop()

# ------------ 1: Oku + ilk görünüm ------------
df = pd.read_excel(uploaded)
st.success("Dosya okundu ✅")
st.dataframe(df.head(), use_container_width=True)

plant_col = "Plant"
kf_col    = "Key Figure"

# ------------ 2: KF sınıflandırma & hızlı kontrol ------------
months_columns = [c for c in df.columns if isinstance(c, (pd.Timestamp, DT))]
months_columns.sort()
st.write("**Months columns (ilk 6):**", months_columns[:6])
st.write("**Toplam ay kolon sayısı:**", len(months_columns))

df["_kf_class"] = df[kf_col].map(classify_kf)
df.loc[df["_kf_class"] == "consensus", plant_col] = "EIP"

st.subheader("Key Figure eşleştirme sonucu")
st.dataframe(df[["_kf_class", kf_col]].drop_duplicates(), use_container_width=True)

df["_key_figure_normalized"] = df[kf_col].map(norm_text)
st.subheader("'consensus' içeren normalized satırlar")
st.dataframe(
    df[df["_key_figure_normalized"].str.contains("consensus", na=False)][[kf_col, "_key_figure_normalized"]],
    use_container_width=True
)

# ------------ 3: Long form + DOC hesap ------------
month_cols = detect_month_columns_by_parsing(df)
if not month_cols:
    st.error("Ay kolonları bulunamadı (başlık 'YYYY-MM-DD ...' formatında olmalı).")
    st.stop()

st.write("**Bulunan ay kolon sayısı:**", len(month_cols))
st.write("**İlk 6 ay:**", month_cols[:6])

month_names = [c for c, _ in month_cols]
col_to_ts   = dict(month_cols)

df_long = df.melt(
    id_vars=[c for c in df.columns if c not in month_names],
    value_vars=month_names,
    var_name="month_col",
    value_name="value"
)
df_long["month_ts"] = df_long["month_col"].map(col_to_ts)

if "_kf_class" not in df.columns:
    df["_kf_class"] = df[kf_col].map(classify_kf)
df_long["_kf_class"] = df_long[kf_col].map(classify_kf)

is_eip         = df_long[plant_col].astype(str).str.lower().str.contains("eip", na=False)
mask_consensus = (df_long["_kf_class"] == "consensus") & is_eip
mask_projected = (df_long["_kf_class"] == "projected_stock")

df_long["value"] = pd.to_numeric(df_long["value"], errors="coerce")
df_long.loc[df_long["_kf_class"]=="consensus","value"] = (
    df_long.loc[df_long["_kf_class"]=="consensus","value"].clip(lower=0)
)

cons_month = (df_long.loc[mask_consensus]
              .groupby("month_ts", dropna=True)["value"].sum()
              .rename("monthly_consensus_eip"))
proj_month = (df_long.loc[mask_projected]
              .groupby("month_ts", dropna=True)["value"].sum()
              .rename("monthly_projected_eip_gp"))

doc_df = pd.concat([proj_month, cons_month], axis=1).sort_index()

MAX_DOC_IF_NO_RUNOUT     = 600
DAYS_PER_MONTH           = 30
CONSENSUS_UNIT_MULTIPLIER= 1

months = doc_df.index.to_list()
stock  = doc_df["monthly_projected_eip_gp"].reindex(months).fillna(0).astype(float)
dem    = (doc_df["monthly_consensus_eip"].reindex(months).fillna(0).astype(float)
          * CONSENSUS_UNIT_MULTIPLIER).clip(lower=0)

def doc_days_from_stock(stock_val, future_monthly_demand):
    if pd.isna(stock_val) or stock_val <= 0: return 0.0
    cum = 0.0; full_months = 0
    for dm in pd.Series(future_monthly_demand).fillna(0).astype(float):
        dm = max(0.0, dm)
        if dm == 0:
            full_months += 1; continue
        if cum + dm < stock_val:
            cum += dm; full_months += 1
        else:
            remaining = stock_val - cum
            frac = max(0.0, remaining)/dm
            return full_months*DAYS_PER_MONTH + frac*DAYS_PER_MONTH
    return MAX_DOC_IF_NO_RUNOUT

doc_vals = []
for i, _ in enumerate(months):
    future_dem = dem.iloc[i+1:]
    doc_vals.append(doc_days_from_stock(stock.iloc[i], future_dem))

doc_df["DOC_days"] = doc_vals

if len(months) >= 2 and dem.iloc[1] > 0:
    naive_first = stock.iloc[0] / dem.iloc[1] * DAYS_PER_MONTH
    st.caption(f"[Sanity] 1. satır (sadece bir sonraki ay) ≈ {naive_first:.2f} gün")

st.subheader("📊 DOC Sonuç Tablosu")
out_df = (doc_df[["monthly_projected_eip_gp","monthly_consensus_eip","DOC_days"]]
          .reset_index(names=["month"]))
st.dataframe(out_df, use_container_width=True)

# İndirilebilir Excel
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
    out_df.to_excel(w, sheet_name="DOC", index=False)
st.download_button(
    "Excel'i indir (DOC_summary.xlsx)",
    data=buf.getvalue(),
    file_name="DOC_summary.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)



