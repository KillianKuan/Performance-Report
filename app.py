import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel 銷售資料分析工具", layout="wide")
st.title("📊 Excel 銷售資料分析工具")

REQUIRED_COLS = [
    "Customer Name", "Ship Date", "QTY",
    "SALES Total AMT", "final GP(NTD,data from Financial Report)",
    "Part Number", "Category"
]
VALID_CATEGORIES = {"Tablet", "CDR", "Tablet ACC", "CDR ACC"}
GP_COL = "final GP(NTD,data from Financial Report)"
# ── DES 關鍵字分類規則（維護此處即可，同步更新頁面對照表）──
DES_RULES = {
    "CDR ACC":    ["cdr", "gemini", "evo", "sprint", "sd card", "panic button", "iosix", "uvc camera", "k220", "k245", "k265"],
    "Tablet ACC": ["tablet", "chiron", "hera", "phaeton", "surfing pro", "cradle", "f840"],
}

# ── 1. 上傳檔案 ──────────────────────────────────────────────
uploaded = st.file_uploader("上傳 .xlsx 檔案", type=["xlsx"])
if not uploaded:
    st.info("請上傳含有 'Actual' 工作表的 .xlsx 檔案。")
    st.stop()

# ── 2. 讀取與驗證 ────────────────────────────────────────────
try:
    xl = pd.ExcelFile(uploaded)
except Exception as e:
    st.error(f"無法讀取檔案：{e}")
    st.stop()

if "Actual" not in xl.sheet_names:
    st.error(f"找不到 'Actual' 工作表。現有工作表：{xl.sheet_names}")
    st.stop()

raw = xl.parse("Actual")
missing = [c for c in REQUIRED_COLS if c not in raw.columns]
if missing:
    st.error(f"缺少必要欄位：{missing}")
    st.stop()

has_des = "DES" in raw.columns
read_cols = REQUIRED_COLS + (["DES"] if has_des else [])
df = raw[read_cols].copy()
if not has_des:
    st.warning("⚠️ 找不到 'DES' 欄位，無法進行 DES 關鍵字分類，未知 Category 將歸入 Others。")

# ── 3. 日期解析 ──────────────────────────────────────────────
df["Ship Date"] = df["Ship Date"].astype(str).str.strip()
df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
nat_count = df["Ship Date"].isna().sum()
if nat_count > 0:
    st.warning(f"⚠️ 跳過 {nat_count} 筆無效或空白的 Ship Date 資料。")
df = df.dropna(subset=["Ship Date"])
df["Month"] = df["Ship Date"].dt.strftime("%Y-%m")

# ── 4. Category 正規化（含 DES 輔助判定）────────────────────
df["Category"] = df["Category"].astype(str).str.strip()
if has_des:
    df["DES"] = df["DES"].astype(str).str.strip()
def classify_by_des(des_val):
    des_lower = des_val.lower()
    return [cat for cat, kws in DES_RULES.items() if any(kw in des_lower for kw in kws)]
ambiguous_rows = []
def normalize_category(row):
    cat = row["Category"]
    if cat in VALID_CATEGORIES:
        return cat
    if not has_des:
        return "Others"
    matches = classify_by_des(str(row.get("DES", "")))
    if len(matches) == 0:
        return "Others"
    elif len(matches) == 1:
        return matches[0]
    else:
        ambiguous_rows.append({
            "Part Number": row.get("Part Number", ""),
            "DES": row.get("DES", ""),
            "Original Category": cat,
            "命中分類": " / ".join(matches),
            "暫定分類": matches[0],
        })
        return matches[0]
df["Category"] = df.apply(normalize_category, axis=1)
if ambiguous_rows:
    st.warning(
        f"⚠️ 發現 {len(ambiguous_rows)} 筆 DES 同時命中多個分類，"
        f"暫以第一匹配（{list(DES_RULES.keys())[0]}）處理，請人工確認："
    )
    st.dataframe(pd.DataFrame(ambiguous_rows), use_container_width=True)

# ── 5. 欄位型別 ──────────────────────────────────────────────
df["QTY"] = pd.to_numeric(df["QTY"], errors="coerce").fillna(0)
df["SALES Total AMT"] = pd.to_numeric(df["SALES Total AMT"], errors="coerce").fillna(0)
df[GP_COL] = pd.to_numeric(df[GP_COL], errors="coerce").fillna(0)
df["Customer Name"] = df["Customer Name"].astype(str).str.strip()
df["Part Number"] = df["Part Number"].astype(str).str.strip()

# ── 6. Customer 搜尋與選擇（checkbox 列表）───────────────────
st.subheader("🔍 Customer 篩選")
cust_query = st.text_input("輸入 Customer Name 關鍵字（substring, 不分大小寫）")

all_customers = sorted(df["Customer Name"].unique())

if not cust_query.strip():
    st.info("請輸入關鍵字以搜尋客戶。")
    st.stop()

matched = [c for c in all_customers if cust_query.strip().lower() in c.lower()]

if not matched:
    st.warning("無符合客戶，顯示 0 rows。")
    st.stop()

st.markdown(f"**符合客戶（共 {len(matched)} 筆），請勾選：**")

selected_customers = []
for cust in matched:
    key = f"cust__{cust}"
    if key not in st.session_state:
        st.session_state[key] = True
    checked = st.checkbox(cust, value=st.session_state[key], key=key)
    if checked:
        selected_customers.append(cust)

# ── 7. Part Number 篩選 ──────────────────────────────────────
st.subheader("🔩 Part Number 篩選（可選）")
use_pn_filter = st.checkbox("啟用 Part Number 篩選（只影響 QTY Filtered）")
selected_pns = []
if use_pn_filter:
    pn_query = st.text_input("輸入 Part Number 關鍵字（substring, 不分大小寫）")
    all_pns = sorted(df["Part Number"].unique())
    matched_pns = (
        [p for p in all_pns if pn_query.strip().lower() in p.lower()]
        if pn_query.strip() else all_pns
    )
    selected_pns = st.multiselect("選擇 Part Number（可多選）", options=matched_pns)

# ── 8. QTY 只計 Tablet & CDR ────────────────────────────────
use_tablet_cdr_only = st.checkbox("QTY 只加總 Category = Tablet & CDR（排除 ACC）")

# ── 9. Category 拆分 ─────────────────────────────────────────
use_cat_split = st.checkbox("依 Category 拆分報表")

# ── 10. 彙整函式 ─────────────────────────────────────────────
def build_long(grp_df, qty_df, group_cols):
    agg = grp_df.groupby(group_cols, sort=True).agg(
        **{
            "SALES Total AMT": ("SALES Total AMT", "sum"),
            "final GP(NTD)": (GP_COL, "sum"),
        }
    ).reset_index()

    qty_all = (
        qty_df.groupby(group_cols)["QTY"].sum()
        .reset_index().rename(columns={"QTY": "QTY (All)"})
    )
    agg = agg.merge(qty_all, on=group_cols, how="left")
    agg["QTY (All)"] = agg["QTY (All)"].fillna(0)

    if use_pn_filter:
        if selected_pns:
            qty_f = qty_df[qty_df["Part Number"].isin(selected_pns)]
            qty_filtered = (
                qty_f.groupby(group_cols)["QTY"].sum()
                .reset_index().rename(columns={"QTY": "QTY (Filtered)"})
            )
            agg = agg.merge(qty_filtered, on=group_cols, how="left")
            agg["QTY (Filtered)"] = agg["QTY (Filtered)"].fillna(0)
        else:
            agg["QTY (Filtered)"] = agg["QTY (All)"]

    return agg


def to_wide(long_df, group_cols, add_total=False):
    month_col = "Month"
    extra_cols = [c for c in group_cols if c != month_col]

    metrics = ["QTY (All)", "SALES Total AMT", "final GP(NTD)"]
    if use_pn_filter:
        metrics.append("QTY (Filtered)")

    id_vars = extra_cols + [month_col]
    melted = long_df.melt(id_vars=id_vars, value_vars=metrics,
                          var_name="Metric", value_name="Value")

    if extra_cols:
        melted["Row"] = melted[extra_cols[0]] + " | " + melted["Metric"]
        pivot = melted.pivot_table(index="Row", columns=month_col,
                                   values="Value", aggfunc="sum")
    else:
        pivot = melted.pivot_table(index="Metric", columns=month_col,
                                   values="Value", aggfunc="sum")
        pivot = pivot.reindex(metrics)

    pivot.columns.name = None

    if add_total:
        month_cols = list(pivot.columns)
        pivot["Total"] = pivot[month_cols].sum(axis=1)

    pivot = pivot.reset_index()
    return pivot


def format_wide(df):
    """對所有數值欄位套用千分位格式（無小數）"""
    label_col = df.columns[0]
    num_cols = [c for c in df.columns if c != label_col]
    return df.style.format(
        formatter="{:,.0f}",
        subset=num_cols,
        na_rep="-"
    )
# ── 11. 產生報表 ─────────────────────────────────────────────
if st.button("▶ Run"):
    if not selected_customers:
        st.warning("請至少勾選一位客戶。")
        st.stop()

    base = df[df["Customer Name"].isin(selected_customers)].copy()
    if base.empty:
        st.warning("選定客戶無資料，顯示 0 rows。")
        st.stop()

    qty_base = base[base["Category"].isin({"Tablet", "CDR"})] if use_tablet_cdr_only else base

    long_summary = build_long(base, qty_base, ["Month"])
    wide_summary = to_wide(long_summary, ["Month"], add_total=True)
    st.subheader("📋 Summary（橫式月報）")
    st.dataframe(format_wide(wide_summary), use_container_width=True)

    wide_bycat = pd.DataFrame()
    if use_cat_split:
        long_bycat = build_long(base, qty_base, ["Month", "Category"])
        wide_bycat = to_wide(long_bycat, ["Month", "Category"], add_total=False)
        st.subheader("📋 ByCategory（橫式月報 × Category）")
        st.dataframe(format_wide(wide_bycat), use_container_width=True)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        wide_summary.to_excel(writer, sheet_name="Summary", index=False)
        if use_cat_split and not wide_bycat.empty:
            wide_bycat.to_excel(writer, sheet_name="ByCategory", index=False)
    buf.seek(0)

    st.download_button(
        label="⬇️ 下載 Excel 報表",
        data=buf,
        file_name="sales_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
