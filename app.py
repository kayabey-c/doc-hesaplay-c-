# app.py
import io, re, unicodedata
from datetime import datetime as DT
import numpy as np
import pandas as pd
import streamlit as st

# =============== Sayfa Ayarı ===============
st.set_page_config(page_title="DOC Hesap", layout="wide")
st.title("📦 Days of Coverage (DOC) Hesaplayıcı")
st.caption("Excel yükleyin → *projected stock* ve *consensus demand* üzerinden DOC hesaplayın.")

# =============== Yardımcılar ===============
def norm_text(s: str) -> str:
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower()
    s = re.sub(r"\s+", " ", s)
    return s

KF_PATTERNS = {
    "consensus": [
        "kisit siz consensus", "consensus",
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
            if p in v:
                return key
    return None

def detect_month_columns_flexible(df: pd.DataFrame):
    """
    1) Başlık datetime/Timestamp ise ayın ilk gününe yuvarlar.
    2) Metin başlık 'YYYY-MM-DD...' ile başlıyorsa parse eder.
    3) Metin başlık 'YYYY/MM/DD...' da kabul.
    Dönen: [(orijinal_kolon_adı, month_start_ts), ...]
    """
    month_cols = []
    for c in df.columns:
        # 1) Doğrudan Timestamp/datetime
        if isinstance(c, (pd.Timestamp, DT)):
            ts = pd.Timestamp(c)
            month_cols.append((c, pd.Timestamp(ts.year, ts.month, 1)))
            continue

        # 2) Metin başlığın başında tarih var mı?
        s = str(c).strip()
        m = re.match(r"^(\d{4}[-/]\d{2}[-/]\d{2})", s)
        if m:
            ts = pd.to_datetime(m.group(1), errors="coerce")
            if pd.notna(ts):
                month_cols.append((c, pd.Timestamp(ts.year, ts.month, 1)))

    # Tekrarları temizle + sırala
    month_cols = list(dict.fromkeys(month_cols))
    month_cols.sort(key=lambda x: x[1])
    return month_cols

def tr_thousands(n, ndigits=2):
    """Türkçe benzeri binlik biçim: 1.234.567,89"""
    try:
        if pd.isna(n):
            return ""
        s = f"{float(n):,.{ndigits}f}"  # 1,234,567.89
        s = s.replace(",", "X").replace(".", ",").replace("X", ".")
        return s
    except Exception:
        return str(n)

MAX_DOC_IF_NO_RUNOUT = 600
DAYS_PER_MONTH = 30

def doc_days_from_stock(stock_val, future_monthly_demand):
    """
    Stok (o ay başındaki proj. stok) ileri ayların talebiyle tüketilir.
    - Talep 0 ise 30 gün ekleyip sonraki aya geçer.
    - Stok bir ayın içinde biterse oransal gün eklenir.
    - Hiç bitmezse 600 gün (üst sınır).
    """
    if pd.isna(stock_val) or float(stock_val) <= 0:
        return 0.0

    stock_val = float(stock_val)
    cum = 0.0
    full_months = 0

    for dm in pd.Series(future_monthly_demand).fillna(0).astype(float):
        dm = max(0.0, dm)
        if dm == 0:
            full_months += 1
            continue
        if cum + dm < stock_val:
            cum += dm
            full_months += 1
        else:
            remaining = stock_val - cum
            frac = max(0.0, remaining) / dm
            return full_months * DAYS_PER_MONTH + frac * DAYS_PER_MONTH

    return MAX_DOC_IF_NO_RUNOUT

# =============== Kenar Çubuğu (Ayarlar) ===============
with st.sidebar:
    st.subheader("Ayarlar")
    use_tr_format = st.checkbox("Tabloda TR sayı formatı (1.234.567,89)", value=False)
    show_checks = st.checkbox("Ara kontrol tablolarını göster", value=True)
    demo = st.checkbox("Demo veriyle dene (Excel gerekmez)", value=False)

# =============== Veri Kaynağı ===============
df = None
if demo:
    # Küçük demo veri (deploy test)
    dates = pd.date_range("2025-01-01", periods=6, freq="MS")
    cols = ["Plant", "Key Figure"] + [d.strftime("%Y-%m-%d 00:00:00") for d in dates]
    rows = [
        ["EIP", "Consensus",           1000, 1200, 1100,  900, 1000, 1000],
        ["GP",  "Projected Stock",     5000, 4000, 3500, 2800, 2600, 2400],
        ["EIP", "Kısıtsız Consensus Sell-in Forecast / Malzeme Tüketim Mik.", 1000,1200,1100,900,1000,1000],
        ["GP",  "Unconstrained Projected Stock", 5000,4000,3500,2800,2600,2400],
    ]
    df = pd.DataFrame(rows, columns=cols)
else:
    uploaded = st.file_uploader("Excel'i sürükleyip bırakın", type=["xlsx", "xls"])
    if uploaded is None:
        st.info("Başlamak için bir Excel dosyası yükleyin ya da 'Demo veriyle dene' kutusunu işaretleyin.")
        st.stop()
    try:
        df = pd.read_excel(uploaded)  # openpyxl şart (requirements'ta var)
    except Exception as e:
        st.error(f"Excel okunamadı: {e}")
        st.stop()

st.success("Veri yüklendi ✅")
st.write("**Kolonlar:**", list(df.columns))
st.dataframe(df.head(), use_container_width=True)

# =============== Kolon seçimi (hata kaynağı %90 burası) ===============
all_cols = list(df.columns)
with st.sidebar:
    plant_col = st.selectbox("Plant kolonu", options=all_cols, index=all_cols.index("Plant") if "Plant" in all_cols else 0)
    kf_col = st.selectbox("Key Figure kolonu", options=all_cols, index=all_cols.index("Key Figure") if "Key Figure" in all_cols else 0)

missing = [c for c in [plant_col, kf_col] if c not in df.columns]
if missing:
    st.error(f"Seçilen kolon(lar) dataframe'de yok: {missing}")
    st.stop()

# =============== KF sınıflandırma ===============
df["_kf_class"] = df[kf_col].map(classify_kf)

# Consensus satırlarında Plant eksik olabilir → EIP'e set edelim
df.loc[df["_kf_class"] == "consensus", plant_col] = df.loc[df["_kf_class"] == "consensus", plant_col].fillna("EIP")
df.loc[df["_kf_class"] == "consensus", plant_col] = "EIP"

if show_checks:
    st.subheader("Key Figure eşleştirme sonucu (benzersiz)")
    st.dataframe(
        df[["_kf_class", kf_col]].drop_duplicates().sort_values("_kf_class", na_position="last"),
        use_container_width=True
    )
    df["_key_figure_normalized"] = df[kf_col].map(norm_text)
    st.subheader("'consensus' içeren normalized satırlar")
    st.dataframe(
        df[df["_key_figure_normalized"].str.contains("consensus", na=False)][[kf_col, "_key_figure_normalized"]],
        use_container_width=True
    )

# =============== Ay kolonları + long format ===============
month_cols = detect_month_columns_flexible(df)
st.write("**Bulunan ay kolon sayısı:**", len(month_cols))
if month_cols:
    st.write("**İlk 6 ay:**", month_cols[:6])
else:
    st.error(
        "Ay kolonları bulunamadı. Çözümler:\n"
        "- Başlıklar gerçek tarih (Excel datetime) olsun **veya** 'YYYY-MM-DD ...' ile başlasın.\n"
        "- Örn: '2025-08-01 00:00:00'.\n"
        "- Şu anki kolon adları için yukarıdaki 'Kolonlar' çıktısına bak."
    )
    st.stop()

month_names = [c for c, _ in month_cols]
col_to_ts = dict(month_cols)

# Long form
df_long = df.melt(
    id_vars=[c for c in df.columns if c not in month_names],
    value_vars=month_names,
    var_name="month_col",
    value_name="value"
)
df_long["month_ts"] = df_long["month_col"].map(col_to_ts)

# Sınıflandırma güvenliği
df_long["_kf_class"] = df_long[kf_col].map(classify_kf)

# Filtreler
is_eip = df_long[plant_col].astype(str).str.lower().str.contains("eip", na=False)
mask_consensus = (df_long["_kf_class"] == "consensus") & is_eip
mask_projected = (df_long["_kf_class"] == "projected_stock")

# Sayısala çevir ve negatif consensus'u 0'a kırp
df_long["value"] = pd.to_numeric(df_long["value"], errors="coerce")
df_long.loc[df_long["_kf_class"] == "consensus", "value"] = (
    df_long.loc[df_long["_kf_class"] == "consensus", "value"].clip(lower=0)
)

# Aylık toplama
cons_month = (
    df_long.loc[mask_consensus]
    .groupby("month_ts", dropna=True)["value"]
    .sum()
    .rename("monthly_consensus_eip")
)
proj_month = (
    df_long.loc[mask_projected]
    .groupby("month_ts", dropna=True)["value"]
    .sum()
    .rename("monthly_projected_eip_gp")
)

doc_df = pd.concat([proj_month, cons_month], axis=1).sort_index()
st.write("**Aylık özet (ilk satırlar):**")
st.dataframe(doc_df.head(), use_container_width=True)

# =============== DOC hesap ===============
CONSENSUS_UNIT_MULTIPLIER = 1.0
months = doc_df.index.to_list()

stock = doc_df["monthly_projected_eip_gp"].reindex(months).fillna(0).astype(float)
dem = (doc_df["monthly_consensus_eip"].reindex(months).fillna(0).astype(float) * CONSENSUS_UNIT_MULTIPLIER).clip(lower=0)

doc_vals = []
for i, _ in enumerate(months):
    # Aynı ay stoğunu bir SONRAKİ ay talebiyle tüket
    future_dem = dem.iloc[i + 1 :]
    doc_vals.append(doc_days_from_stock(stock.iloc[i], future_dem))

doc_df["DOC_days"] = doc_vals

# Hızlı sanity check
if len(months) >= 2 and dem.iloc[1] > 0:
    naive_first = stock.iloc[0] / dem.iloc[1] * DAYS_PER_MONTH
    st.caption(f"[Sanity] 1. satır (sadece bir sonraki ay) ≈ {naive_first:.2f} gün")

# =============== Çıktı & İndirme ===============
st.subheader("📊 DOC Sonuç Tablosu")
out_df = doc_df[["monthly_projected_eip_gp", "monthly_consensus_eip", "DOC_days"]].reset_index(names=["month"])

if out_df.empty:
    st.warning("Çıktı tablosu boş görünüyor. 'projected_stock' veya 'consensus' sınıflaması eşleşmemiş olabilir.")
    st.stop()

if use_tr_format:
    show_df = out_df.copy()
    show_df["monthly_projected_eip_gp"] = show_df["monthly_projected_eip_gp"].map(lambda x: tr_thousands(x, 2))
    show_df["monthly_consensus_eip"] = show_df["monthly_consensus_eip"].map(lambda x: tr_thousands(x, 2))
    show_df["DOC_days"] = show_df["DOC_days"].map(lambda x: tr_thousands(x, 2))
    st.dataframe(show_df, use_container_width=True)
else:
    st.dataframe(out_df, use_container_width=True)

buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="XlsxWriter") as writer:
    out_df.to_excel(writer, sheet_name="DOC", index=False)
    wb = writer.book
    ws = writer.sheets["DOC"]
    num_fmt = wb.add_format({"num_format": "#,##0.00"})
    day_fmt = wb.add_format({"num_format": "0.00"})
    ws.set_column("A:A", 12)      # month
    ws.set_column("B:C", 18, num_fmt)  # stocks/demand
    ws.set_column("D:D", 12, day_fmt)  # DOC_days

st.download_button(
    "Excel'i indir (DOC_summary.xlsx)",
    data=buf.getvalue(),
    file_name="DOC_summary.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)





