"""
🎯 台灣權證篩選系統 — Web 版
從 TWSE 即時抓取權證行情，篩選低價＋長天期的權證
"""

import json, time, re, warnings
from datetime import datetime, timedelta
from typing import Optional

import pandas as pd
import requests
import streamlit as st
from dateutil.relativedelta import relativedelta

warnings.filterwarnings("ignore")

# ─── Page Config ───
st.set_page_config(
    page_title="台灣權證篩選器",
    page_icon="🎯",
    layout="wide",
)

# ─── Custom CSS ───
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap');

    .stApp { font-family: 'Noto Sans TC', sans-serif; }

    .main-header {
        background: linear-gradient(135deg, #0f0c29 0%, #302b63 50%, #24243e 100%);
        padding: 2rem 2.5rem;
        border-radius: 16px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .main-header h1 {
        font-size: 2rem;
        margin: 0 0 0.3rem 0;
        font-weight: 700;
    }
    .main-header p {
        margin: 0;
        opacity: 0.8;
        font-size: 0.95rem;
    }

    .metric-card {
        background: #f8f9fa;
        border-radius: 12px;
        padding: 1.2rem;
        text-align: center;
        border: 1px solid #e9ecef;
    }
    .metric-card .number {
        font-size: 2rem;
        font-weight: 700;
        color: #302b63;
    }
    .metric-card .label {
        font-size: 0.85rem;
        color: #6c757d;
        margin-top: 0.2rem;
    }

    div[data-testid="stDataFrame"] {
        border-radius: 12px;
        overflow: hidden;
    }
</style>
""", unsafe_allow_html=True)

# ─── Header ───
st.markdown("""
<div class="main-header">
    <h1>🎯 台灣權證篩選系統</h1>
    <p>即時從證交所抓取權證行情，篩選低價 + 長天期的權證</p>
</div>
""", unsafe_allow_html=True)

H = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 Chrome/120.0 Safari/537.36"
}


# ━━━━━━━━━━━━━━━━━ 工具函式 ━━━━━━━━━━━━━━━━━
def parse_price(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v) if float(v) != 0 else None
    s = str(v).strip().replace(",", "")
    if s in ("", "-", "--", "N/A", "0", "0.00", "X"):
        return None
    try:
        f = float(s)
        return f if f != 0 else None
    except:
        return None


def extract_expiry_from_name(name):
    """
    權證命名: {標的}{券商}{年末碼}{月份}{購/售}{流水號}
    """
    if not name or not isinstance(name, str):
        return None
    m = re.search(r"(\d{2,3})(購|售)(\d+)\s*$", name.strip())
    if not m:
        return None

    num = m.group(1)
    if len(num) == 2:
        yd, mo = int(num[0]), int(num[1])
    elif len(num) == 3:
        yd, mo = int(num[0]), int(num[1:])
    else:
        return None
    if mo < 1 or mo > 12:
        return None

    cur_roc = datetime.now().year - 1911
    dec = cur_roc // 10
    best = min(
        [(dec - 1) * 10 + yd, dec * 10 + yd, (dec + 1) * 10 + yd],
        key=lambda y: abs(y - cur_roc) if 2020 <= y + 1911 <= 2040 else 9999,
    )
    if best + 1911 < 2020 or best + 1911 > 2040:
        return None
    y = best + 1911
    return (
        datetime(y, mo + 1, 1) - timedelta(days=1)
        if mo < 12
        else datetime(y, 12, 31)
    )


def safe_get(url, label=""):
    try:
        r = requests.get(url, headers=H, timeout=30)
        return r if r.status_code == 200 and len(r.content) > 50 else None
    except:
        return None


# ━━━━━━━━━━━━━━━━━ 資料抓取 ━━━━━━━━━━━━━━━━━
def fetch_twse_warrants(date_str=None, _depth=0):
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")
    url = (
        f"https://www.twse.com.tw/exchangeReport/MI_INDEX"
        f"?response=json&date={date_str}&type=0999"
    )
    resp = safe_get(url, f"TWSE ({date_str})")
    if resp is None:
        return pd.DataFrame()
    try:
        data = resp.json()
    except:
        return pd.DataFrame()
    if data.get("stat", "").upper() != "OK":
        if _depth < 5:
            prev = (
                datetime.strptime(date_str, "%Y%m%d") - timedelta(days=1)
            ).strftime("%Y%m%d")
            time.sleep(2)
            return fetch_twse_warrants(prev, _depth + 1)
        return pd.DataFrame()
    for t in data.get("tables", []):
        if isinstance(t, dict) and len(t.get("data", [])) > 100:
            f = t.get("fields", [])
            rows = t["data"]
            return (
                pd.DataFrame(rows, columns=f)
                if f and len(f) == len(rows[0])
                else pd.DataFrame(rows)
            )
    return pd.DataFrame()


def normalize(df):
    m, used = {}, set()
    for col in df.columns:
        c = str(col).strip()
        t = None
        if ("證券代號" in c or c == "代號") and "代號" not in used:
            t = "代號"
        elif ("證券名稱" in c or c == "名稱") and "名稱" not in used:
            t = "名稱"
        elif "收盤" in c and "收盤價" not in used:
            t = "收盤價"
        elif "成交股數" in c and "成交量" not in used:
            t = "成交量"
        elif "開盤" in c and "開盤價" not in used:
            t = "開盤價"
        elif "最高" in c and "最高價" not in used:
            t = "最高價"
        elif "最低" in c and "最低價" not in used:
            t = "最低價"
        if t:
            m[col] = t
            used.add(t)
    return df.rename(columns=m) if m else df


# ━━━━━━━━━━━━━━━━━ 主篩選 ━━━━━━━━━━━━━━━━━
def screen_warrants(max_price, min_months, warrant_type, progress_bar):
    today = datetime.now()
    cutoff = today + relativedelta(months=min_months)

    progress_bar.progress(10, "📡 從證交所抓取行情...")
    df = normalize(fetch_twse_warrants())
    if df.empty:
        return pd.DataFrame(), 0

    total_raw = len(df)
    progress_bar.progress(50, "🔧 解析權證名稱中的到期月份...")
    df["到期日"] = df["名稱"].apply(extract_expiry_from_name)
    df["收盤價_v"] = df["收盤價"].apply(parse_price)
    df["距到期天數"] = df["到期日"].apply(
        lambda d: (d - today).days if isinstance(d, datetime) else None
    )

    progress_bar.progress(80, "🎯 篩選中...")
    mask = (
        df["收盤價_v"].notna()
        & (df["收盤價_v"] > 0)
        & (df["收盤價_v"] <= max_price)
        & df["到期日"].notna()
        & (df["到期日"] >= cutoff)
    )
    if warrant_type == "call":
        mask &= df["名稱"].str.contains("購", na=False)
    elif warrant_type == "put":
        mask &= df["名稱"].str.contains("售", na=False)

    result = df[mask].sort_values(
        ["距到期天數", "收盤價_v"], ascending=[False, True]
    ).reset_index(drop=True)

    progress_bar.progress(100, "✅ 完成！")
    return result, total_raw


# ━━━━━━━━━━━━━━━━━ Sidebar 篩選條件 ━━━━━━━━━━━━━━━━━
with st.sidebar:
    st.markdown("### ⚙️ 篩選條件")

    max_price = st.number_input(
        "收盤價上限（元）",
        min_value=0.01,
        max_value=100.0,
        value=0.5,
        step=0.1,
        format="%.2f",
    )
    min_months = st.slider("到期日至少在幾個月之後", 1, 24, 6)

    wt_options = {"全部": None, "認購 (Call)": "call", "認售 (Put)": "put"}
    wt_label = st.radio("權證類型", list(wt_options.keys()))
    wt_code = wt_options[wt_label]

    st.markdown("---")
    run_btn = st.button("🔍 開始篩選", use_container_width=True, type="primary")

    st.markdown("---")
    st.markdown(
        """
        <div style="font-size:0.8rem; color:#888;">
        <b>命名規則</b><br>
        <code>{標的}{券商}{年末碼}{月}{購/售}{流水號}</code><br>
        例：大立光凱基76購06<br>→ 民國107年6月 認購 第06檔
        </div>
        """,
        unsafe_allow_html=True,
    )

# ━━━━━━━━━━━━━━━━━ 主頁面 ━━━━━━━━━━━━━━━━━
if run_btn:
    progress = st.progress(0, "準備中...")
    df_result, total_raw = screen_warrants(max_price, min_months, wt_code, progress)
    time.sleep(0.3)
    progress.empty()

    if df_result.empty:
        st.warning("⚠️ 無符合條件的結果，請調整篩選條件後再試")
    else:
        # Metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(
                f'<div class="metric-card"><div class="number">{total_raw:,}</div>'
                f'<div class="label">全部權證</div></div>',
                unsafe_allow_html=True,
            )
        with col2:
            st.markdown(
                f'<div class="metric-card"><div class="number">{len(df_result):,}</div>'
                f'<div class="label">符合條件</div></div>',
                unsafe_allow_html=True,
            )
        with col3:
            avg_price = df_result["收盤價_v"].mean()
            st.markdown(
                f'<div class="metric-card"><div class="number">{avg_price:.3f}</div>'
                f'<div class="label">平均收盤價</div></div>',
                unsafe_allow_html=True,
            )
        with col4:
            avg_days = df_result["距到期天數"].mean()
            st.markdown(
                f'<div class="metric-card"><div class="number">{avg_days:.0f}</div>'
                f'<div class="label">平均距到期天數</div></div>',
                unsafe_allow_html=True,
            )

        st.markdown("<br>", unsafe_allow_html=True)

        # 準備顯示欄位
        show_cols = [
            c
            for c in [
                "代號", "名稱", "收盤價", "到期月份", "距到期",
                "開盤價", "最高價", "最低價", "成交量",
            ]
            if c in df_result.columns or c in ("到期月份", "距到期")
        ]

        dd = df_result.copy()
        dd["到期月份"] = dd["到期日"].apply(
            lambda d: d.strftime("%Y/%m") if isinstance(d, datetime) else "-"
        )
        dd["距到期"] = dd["距到期天數"].apply(
            lambda x: f"{int(x)}天" if pd.notna(x) else "-"
        )

        display_cols = [
            c for c in [
                "代號", "名稱", "收盤價", "到期月份", "距到期",
                "開盤價", "最高價", "最低價", "成交量",
            ]
            if c in dd.columns
        ]

        st.dataframe(
            dd[display_cols],
            use_container_width=True,
            height=600,
        )

        # 搜尋特定標的
        st.markdown("---")
        keyword = st.text_input("🔎 搜尋特定標的（如：台積電、鴻海）")
        if keyword:
            filtered = dd[dd["名稱"].str.contains(keyword, na=False)]
            if filtered.empty:
                st.info(f'找不到包含「{keyword}」的權證')
            else:
                st.success(f'找到 {len(filtered)} 筆包含「{keyword}」的權證')
                st.dataframe(filtered[display_cols], use_container_width=True)

        # CSV 下載
        st.markdown("---")
        csv_data = dd[display_cols].to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="💾 下載 CSV",
            data=csv_data,
            file_name=f"warrant_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True,
        )
else:
    st.info("👈 請在左側設定篩選條件，然後點「🔍 開始篩選」")

    # 使用說明
    st.markdown("---")
    st.markdown("""
    ### 📖 使用說明

    1. **設定篩選條件**：左側可設定價格上限、到期月數門檻、權證類型
    2. **點擊篩選**：系統會即時從證交所抓取最新行情
    3. **查看結果**：表格支援排序，也可搜尋特定標的
    4. **下載 CSV**：一鍵匯出篩選結果

    > **注意**：資料來自證交所公開 API，僅在交易日有更新。若為假日會自動回溯到最近交易日。
    """)
