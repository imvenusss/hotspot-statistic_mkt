#原版本 + Hotspot Statistic 
# -*- coding: utf-8 -*-
import io
import re
from collections import Counter

import pandas as pd
import plotly.express as px
import streamlit as st

# =========================
# 基本設定
# =========================
st.set_page_config(page_title="WiFi AP 統計面板", layout="wide")
st.title("📶 WiFi AP 統計面板")
st.caption(
    "上傳本月 / 上月資料，所有總數與差異統計只採計六類（CTM/Managed/Mixed/Bus/Ferry/Limo）。\n\n"
    "本版本已支援『欄位名稱模糊識別』（忽略大小寫、空白、底線、各式連字號與非英數字元），例如："
    "`Wi-Fi Technology`、`WiFi Technology`、`wifi technology`、`Wi‑Fi-Technology` 皆可自動對應。\n"
    "新增支援輸入與輸出 **Hotspot Name (Chinese)**：在差異與明細中會同時列出 Site Code 與對應中文名稱（上月/本月）。"
)

# =========================
# 常數與分類規則
# =========================
CTM_WIFI_TYPES = [
    'CTM WiFi', 'CTM Internal', 'CTM Wifi (strategic - Convenience Store)',
    'Partnership Wifi', 'Partnership wifi', 'CTM WiFi (ifood)',
    'CTM Wifi (strategic - Chain Supermarket)', 'Free Hotspots',
    'CTM Wifi', 'CTM Wifi ', 'FreeWiFi.MO by CTM', 'CTM WiFi (Ports)',
    'Wifi Street', 'SOC_TEST', 'CTM WiFi (Partnership)'
]
MANAGED_WIFI_TYPES = ['Managed Wi-Fi', 'FreeWiFi.MO by CityU', ' FreeWiFi.MO by IAM']
BUS_WIFI_TYPES = ['Bus Wifi']
FERRY_WIFI_TYPES = ['Ferry Wifi']
LIMO_WIFI_TYPES = ['Limo / Shuttle Wi-Fi']

CATEGORY_ORDER = ["CTM WiFi", "Managed WiFi", "Mixed Site", "Bus WiFi", "Ferry WiFi", "Limo WiFi"]
CATEGORY_SET = set(CATEGORY_ORDER)

# 圖表顯示名稱
CATEGORY_DISPLAY_NAMES = {
    "CTM WiFi": "CTM Wi-Fi Hotspot",
    "Managed WiFi": "Managed Wi-Fi Hotspot",
    "Mixed Site": "Mixed Site(CTM Wi-Fi &amp; Managed Wi-Fi)",
    "Bus WiFi": "Public Bus Wi-Fi(CTM Wi-Fi Hotspot)",
    "Ferry WiFi": "Ferry(Managed Wi-Fi)",
    "Limo WiFi": "Limo/Shuttle(Managed Wi-Fi)"
}

# Managed Wi‑Fi 組合（四類）
MANAGED_ORIGINAL_CATEGORIES = ["Managed WiFi", "Mixed Site", "Ferry WiFi", "Limo WiFi"]
MANAGED_GROUP_NAMES = {
    "Managed WiFi": "Managed Wi-Fi",
    "Mixed Site": "Mixed Site",
    "Ferry WiFi": "Ferry",
    "Limo WiFi": "Limo/Shuttle"
}
MANAGED_GROUP_ORDER = ["Managed Wi‑Fi", "Mixed Site", "Ferry", "Limo/Shuttle"]

WIFI_LEVELS_DISPLAY = ["Wi‑Fi 4", "Wi‑Fi 5", "Wi‑Fi 6", "Wi‑Fi 7"]

COLOR_MAP = {
    "Wi‑Fi 4": "#4F81BD",
    "Wi‑Fi 5": "#C0504D",
    "Wi‑Fi 6": "#9BBB59",
    "Wi‑Fi 7": "#8064A2",
    "Unknown": "#BDBDBD",
    "Huawei": "#C0504D",
    "Ruckus": "#4F81BD"
}

# =========================
# 欄位名稱模糊解析：別名與工具
# =========================
COLUMN_ALIASES = {
    "service_type": ["service type","servicetype","svc type","svctype","type of service"],
    "ssid1": ["ssid 1","ssid1","ssid-1","ssid_1","ssid","primary ssid"],
    "site_code": ["site code","sitecode","site id","siteid","site","location code","location id"],
    "wifi_technology": ["wifi technology","wi fi technology","wi-fi technology","wi‑fi technology","wifi tech","wifi standard","wifi gen","wireless standard"],
    "ap_model": ["ap model","ap-model","apmodel","model","device model","equipment model","hardware model"],
    "hotspot_name_cn": [
        "hotspot name (chinese)","hotspot name chinese","hotspot chinese name",
        "hotspotname(chinese)","hotspotname chinese","hotspot_cn","hotspot cn name",
        "chinese hotspot name","chinese name","名稱(中文)","hotspot 名稱(中文)",
        "hotspot名稱(中文)","hotspot中文名稱","hotspot chinese","hotspot name zh",
        "hotspot zh name","hotspot name (zh)"
    ],
}

def _norm_key(s: str) -> str:
    if not isinstance(s, str): s = str(s)
    s = s.lower().replace("‑","-").replace("–","-").replace("—","-")
    s = re.sub(r'[\s\-_]+','', s)
    s = re.sub(r'[^a-z0-9]','', s)
    return s

def _tokens(s: str) -> set:
    if not isinstance(s, str): s = str(s)
    s = s.lower().replace("‑"," ").replace("–"," ").replace("—"," ").replace("-"," ")
    s = re.sub(r'[^a-z0-9\s]',' ', s)
    toks = [t for t in re.split(r'\s+', s) if t]
    return set(toks)

def _best_match_column(df_columns, alias_list):
    if not df_columns: return None
    info = [{"raw":c,"norm":_norm_key(c),"tok":_tokens(c)} for c in df_columns]
    for alias in alias_list:
        an, at = _norm_key(alias), _tokens(alias)
        for i in info:
            if at and at.issubset(i["tok"]): return i["raw"]
        for i in info:
            if an and an in i["norm"]: return i["raw"]
    return None

def resolve_columns(df: pd.DataFrame,
                    required_min=("service_type","ssid1","site_code"),
                    optional=("wifi_technology","ap_model","hotspot_name_cn")) -> dict:
    mapping, cols, missing = {}, list(df.columns), []
    for key in required_min:
        found = _best_match_column(cols, COLUMN_ALIASES.get(key,[key]))
        if found is None: missing.append(key)
        mapping[key] = found
    if missing:
        readable = {"service_type":"Service Type","ssid1":"SSID 1","site_code":"Site Code"}
        raise ValueError(f"缺少欄位：{[readable.get(m,m) for m in missing]}")
    for key in optional:
        mapping[key] = _best_match_column(cols, COLUMN_ALIASES.get(key,[key]))
    return mapping

# =========================
# 輔助函式
# =========================
def normalize_wifi_tech(val: str) -> str:
    if not isinstance(val, str): return "Unknown"
    s = val.strip().lower().replace(" ","").replace("-","")
    mapping = {"wifi4":"Wi‑Fi 4","wi‑fi4":"Wi‑Fi 4","wi-fi4":"Wi‑Fi 4",
               "wifi5":"Wi‑Fi 5","wi‑fi5":"Wi‑Fi 5","wi-fi5":"Wi‑Fi 5",
               "wifi6":"Wi‑Fi 6","wi‑fi6":"Wi‑Fi 6","wi-fi6":"Wi‑Fi 6",
               "wifi7":"Wi‑Fi 7","wi‑fi7":"Wi‑Fi 7","wi-fi7":"Wi‑Fi 7"}
    return mapping.get(s, val.strip() if val.strip() else "Unknown")

def classify_vendor(model):
    if isinstance(model, str) and "airengine" in model.lower(): return "Huawei"
    return "Ruckus"

# === 新增：安全字串工具 + 特殊站點判斷（Event/Idle 放行用於異常列排除） ===
SPECIAL_SITE_CODES = {"event", "idle"}

def _to_str(x) -> str:
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x)

def _safe_strip(x) -> str:
    return _to_str(x).strip()

def _safe_upper(x) -> str:
    return _safe_strip(x).upper()

def _is_special_site(site_val) -> bool:
    try:
        return _safe_strip(site_val).lower() in SPECIAL_SITE_CODES
    except Exception:
        return False
# === 新增結束 ===

def assign_category_row(service_type, ssid1):
    # === 修補（最小改動）：避免 ssid1 為 float/NaN 觸發 strip 錯誤 ===
    s = _safe_upper(ssid1)  # 原：(ssid1 or "").strip().upper()
    if service_type in MANAGED_WIFI_TYPES:
        return "Mixed Site" if s == "CTM-WIFI" else "Managed WiFi"
    if service_type in CTM_WIFI_TYPES: return "CTM WiFi"
    if service_type in BUS_WIFI_TYPES: return "Bus WiFi"
    if service_type in FERRY_WIFI_TYPES: return "Ferry WiFi"
    if service_type in LIMO_WIFI_TYPES: return "Limo WiFi"
    return "Other"

def read_upload(uploaded):
    if uploaded is None: return None, "尚未上傳"
    try:
        if uploaded.name.lower().endswith(".csv"):
            return pd.read_csv(uploaded), None
        return pd.read_excel(uploaded, engine="openpyxl"), None
    except Exception as e:
        return None, str(e)

def prepare_df(df: pd.DataFrame, allow_missing_wifi_vendor: bool = False):
    d = df.copy()
    colmap = resolve_columns(d)
    col_service, col_ssid1, col_site = colmap["service_type"], colmap["ssid1"], colmap["site_code"]
    col_wifi, col_model, col_hotspot_cn = colmap.get("wifi_technology"), colmap.get("ap_model"), colmap.get("hotspot_name_cn")

    if col_wifi is not None:
        d["Wifi Technology (norm)"] = d[col_wifi].apply(normalize_wifi_tech)
    else:
        if allow_missing_wifi_vendor: d["Wifi Technology (norm)"] = "Unknown"
        else: raise ValueError("缺少欄位：Wifi Technology")

    if col_model is not None:
        d["Vendor"] = d[col_model].apply(classify_vendor)
    else:
        if allow_missing_wifi_vendor: d["Vendor"] = "Unknown"
        else: raise ValueError("缺少欄位：AP Model")

    d["Category"] = d.apply(lambda r: assign_category_row(r[col_service], r[col_ssid1]), axis=1)
    d["Site Code"] = d[col_site].astype(str)
    d["Hotspot Name (Chinese)"] = d[col_hotspot_cn].astype(str).fillna("").str.strip() if col_hotspot_cn else ""
    d6 = d[d["Category"].isin(CATEGORY_SET)].copy()
    return d, d6

def count_wifi_tech_series(df, col="Wifi Technology (norm)"):
    vc = df[col].value_counts()
    return vc.reindex(WIFI_LEVELS_DISPLAY, fill_value=0)

def series_to_df_for_chart(series, name_col, value_col, include_unknown=False, full_series=None):
    df = series.rename_axis(name_col).reset_index(name=value_col)
    if include_unknown and full_series is not None:
        unknown = int(full_series.sum() - series.sum())
        if unknown > 0:
            df = pd.concat([df, pd.DataFrame([{name_col:"Unknown", value_col:unknown}])], ignore_index=True)
    return df

def site_category_majority(df: pd.DataFrame) -> dict:
    res = {}
    for site, sub in df.groupby("Site Code"):
        cnt = Counter(sub["Category"])
        res[site] = cnt.most_common(1)[0][0]
    return res

# =========================
# Plotly config：移除相機、全屏、Plotly logo
# =========================
PLOTLY_CONFIG = dict(
    displayModeBar=False,
    displaylogo=False,
    responsive=True
)

# =========================
# 非餅圖的統一樣式（apply_clean_layout）
# =========================
def apply_clean_layout(
    fig, title_text: str,
    remove_y_title: bool = True,
    percent_axis: bool = False,
    x_domain_end: float = 0.90,
    bar_thickness: float = 0.55,
    bar_gap: float = 0.10,
    show_grid: bool = True,
    legend_x: float = 0.96,
    right_margin: int = 180
):
    x_domain_end = max(0.75, min(x_domain_end, 0.95))
    legend_x = max(0.90, min(legend_x, 1.0))  # （後面會再以 update_layout 推到 >1）
    right_margin = max(80, min(int(right_margin), 320))  # （後面會再以 update_layout 擴大）

    fig.update_layout(
        title={
            "text": title_text,
            "x": 0.49, "xanchor": "center",
            "y": 0.97, "yanchor": "top",
            "font": {"size": 28, "color": "#111"},
            "pad": {"b": 0}
        },
        margin=dict(t=42, r=right_margin, l=10, b=42),
        legend=dict(
            orientation="v",
            x=legend_x, y=0.5, xanchor="left", yanchor="middle",
            bgcolor="rgba(255,255,255,0.6)",
            font=dict(size=24, color="#000")
        ),
        legend_title_text=None,
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(color="#000", size=15),
        dragmode=False,
        bargap=bar_gap
    )

    xaxis_common = dict(
        title=None,
        title_font=dict(size=18, color="#000"),
        tickfont=dict(size=18, color="#000"),
        domain=[0.0, x_domain_end],
        showgrid=show_grid,
        gridcolor="#E0E0E0",
        gridwidth=1,
        zeroline=False,
        ticksuffix="%" if percent_axis else None
    )
    if percent_axis:
        fig.update_xaxes(range=[0, 100], tick0=0, dtick=10, **xaxis_common)
    else:
        fig.update_xaxes(**xaxis_common)

    fig.update_yaxes(
        title=None if remove_y_title else fig.layout.yaxis.title,
        title_font=dict(size=18, color="#000"),
        tickfont=dict(size=24, color="#000"),
        showgrid=False
    )

    fig.update_traces(width=max(0.20, min(bar_thickness, 0.95)))
    fig.update_traces(text=None, texttemplate=None, textposition=None)

    return fig

# =========================
# 12cm 正方形：拆成「全網 AP」與「Managed Wi‑Fi」兩套配置
# =========================
def _make_square_pie_12cm_base(
    df: pd.DataFrame,
    names: str,
    values: str,
    *,
    title_text: str,
    color_col: str = None,
    color_discrete_map: dict = None,
    show_value_and_percent: bool = False,
    small_pct_threshold: float = 0.05,
    # ↓↓↓ 可調參數（用於區分兩套配置）
    title_y: float = 0.99,
    pie_y1: float = 0.95,
    margin_t: int = 6,
    # === 新增：可覆寫扇區文字 ===
    text_values=None
):
    """
    共有邏輯：12cm 畫布、主圓直徑 8cm（= 8/12）、右側垂直圖例 24px、扇區文字 16px、
    小比例外放（帶指引線）、整圖標題（annotation, 粗體 24, 對齊整張圖）。
    """
    # 12cm → px
    PX_PER_CM = 96 / 2.54
    FIG_SIZE = int(round(12 * PX_PER_CM))  # ≈ 454

    # 主圓 domain：直徑比例 8/12=0.6667；水平置中；垂直由 pie_y1 控制上緣
    domain_size = 8.0 / 12.0               # ≈ 0.6667
    y1 = pie_y1
    y0 = y1 - domain_size
    x0 = 0.5 - domain_size/2.0
    x1 = 0.5 + domain_size/2.0

    # 基礎餅圖
    fig = px.pie(
        df,
        names=names,
        values=values,
        title=None,  # 用 annotation 畫整圖標題
        color=(color_col or names),
        color_discrete_map=color_discrete_map
    )
    fig.update_traces(sort=False, direction="clockwise", rotation=0)

    total = float(df[values].sum()) if len(df) else 0.0
    fracs = (df[values] / total).fillna(0).tolist() if total > 0 else [0.0] * len(df)

    # 小比例外放（帶指引線）
    text_positions = ["inside" if f >= small_pct_threshold else "outside" for f in fracs]

    # 文字模板（百分比 → 整數 %）
    if show_value_and_percent:
        texttemplate = "%{value}，%{percent:.0%}"   # 數量 + 整數百分比
    else:
        texttemplate = "%{percent:.0%}"             # 僅整數百分比

    # 應用 trace 設定 —— 關閉 automargin，避免被外部文字再“頂下去”
    if text_values is not None:
        # 直接使用外部給定的文字（例如「123，45%」或「45%」），避免 Plotly 再自行四捨五入
        fig.update_traces(
            text=text_values,
            textposition=text_positions,
            texttemplate="%{text}",
            textfont=dict(size=16, color="#000"),
            insidetextfont=dict(size=16, color="#000"),
            outsidetextfont=dict(size=16, color="#000"),
            hovertemplate="%{label}<br>%{text}",
            automargin=False,
            showlegend=True,
            domain=dict(x=[x0, x1], y=[y0, y1])
        )
    else:
        # 保持原行為（使用 Plotly 的 %{percent}）
        fig.update_traces(
            textposition=text_positions,
            texttemplate=texttemplate,
            textfont=dict(size=16, color="#000"),
            insidetextfont=dict(size=16, color="#000"),
            outsidetextfont=dict(size=16, color="#000"),
            hovertemplate="%{label}: %{value} (%{percent:.0%})",
            automargin=False,
            showlegend=True,
            domain=dict(x=[x0, x1], y=[y0, y1])
        )

    # 佈局與圖例（頂部邊距由配置控制）
    fig.update_layout(
        width=FIG_SIZE,
        height=FIG_SIZE,
        margin=dict(t=margin_t, r=22, b=16, l=22),
        legend=dict(
            orientation="v",
            x=0.92,  # 右側垂直
            y=0.57, xanchor="left", yanchor="middle",
            font=dict(size=22, color="#000"),
            bgcolor="rgba(255,255,255,0)"
        ),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(color="#000", size=15),
        dragmode=False
    )

    # 整圖標題（annotation）：對齊整張圖，位置由配置控制
    fig.update_layout(
        annotations=[
            dict(
                text=f"<b>{title_text}</b>",
                x=0.5, y=title_y,
                xref="paper", yref="paper",
                xanchor="center", yanchor="top",
                showarrow=False,
                font=dict(size=24, color="#111")
            )
        ]
    )
    return fig

def make_square_pie_12cm_overall(
    df: pd.DataFrame,
    names: str,
    values: str,
    title_text: str,
    *,
    color_col: str = None,
    color_discrete_map: dict = None,
    show_value_and_percent: bool = False,
    small_pct_threshold: float = 0.05,
    # === 新增：可覆寫扇區文字 ===
    text_values=None
):
    """全網 AP：標題更靠上，主圓下移以拉開距離（pie_y1=0.92）"""
    return _make_square_pie_12cm_base(
        df, names, values,
        title_text=title_text,
        color_col=color_col,
        color_discrete_map=color_discrete_map,
        show_value_and_percent=show_value_and_percent,
        small_pct_threshold=small_pct_threshold,
        title_y=0.994,     # 標題更靠上
        pie_y1=0.92,       # ← 下移
        margin_t=4,        # 頂部極小緩衝
        text_values=text_values
    )

def make_square_pie_12cm_managed(
    df: pd.DataFrame,
    names: str,
    values: str,
    title_text: str,
    *,
    color_col: str = None,
    color_discrete_map: dict = None,
    show_value_and_percent: bool = False,
    small_pct_threshold: float = 0.05,
    # === 新增：可覆寫扇區文字 ===
    text_values=None
):
    """Managed Wi‑Fi：兩行標題，主圓下移避免重疊（pie_y1=0.83）"""
    return _make_square_pie_12cm_base(
        df, names, values,
        title_text=title_text,
        color_col=color_col,
        color_discrete_map=color_discrete_map,
        show_value_and_percent=show_value_and_percent,
        small_pct_threshold=small_pct_threshold,
        title_y=0.988,
        pie_y1=0.83,
        margin_t=8,
        text_values=text_values
    )

# === 新增：將每個分組的 Count 轉整數百分比（合計=100）的工具函式（堆疊條） ===
import math
from decimal import Decimal, getcontext
getcontext().prec = 28  # 高精度，避免浮點殘差

def assign_integer_percent(
    df: pd.DataFrame,
    group_col: str,
    count_col: str,
    out_col: str = "PercentInt"
) -> pd.DataFrame:
    """
    對於 df 中每一個 group_col（例如一列分類）內的 count_col（例如各技術的數量），
    依照最大餘數法計算整數百分比，並寫到 out_col。
    - 每個分組的 out_col 合計必定為 100。
    - 對 count 做非負化與 NaN 轉 0，避免異常。
    - 使用 Decimal 提升精度，並在每組分配時以穩定排序分派剩餘份數（保留原始順序）。
    """
    df = df.copy()
    df[count_col] = pd.to_numeric(df[count_col], errors="coerce").fillna(0).clip(lower=0)

    percent_int = [0] * len(df)
    for g, idx in df.groupby(group_col, sort=False).groups.items():
        idx_list = list(idx)
        counts = [Decimal(str(x)) for x in df.loc[idx_list, count_col].tolist()]
        total = sum(counts)

        if total <= 0:
            for k in idx_list:
                percent_int[df.index.get_loc(k)] = 0
            continue

        raw = [ (c * Decimal(100)) / total for c in counts ]
        floors = [ int(x.to_integral_value(rounding="ROUND_FLOOR")) for x in raw ]
        remainder = 100 - sum(floors)
        remainder = int(max(0, min(100, remainder)))

        fracs = [(raw[i] - Decimal(floors[i]), i) for i in range(len(raw))]
        fracs.sort(key=lambda t: t[0], reverse=True)

        ints = floors[:]
        for j in range(remainder):
            if j < len(fracs):
                ints[fracs[j][1]] += 1

        for loc_pos, v in zip(idx_list, ints):
            percent_int[df.index.get_loc(loc_pos)] = int(v)

    df[out_col] = percent_int
    return df

# === 新增：餅圖的整數百分比分配（合計=100） ===
def compute_pie_integer_percent(df: pd.DataFrame, value_col: str, out_col: str = "PercentInt") -> pd.DataFrame:
    """
    對單一餅圖的資料表 df（每列一個扇區），依 value_col 的數值
    以「最大餘數法」計算整數百分比，寫入 out_col（整數，合計=100）。
    不改動原來 df 的其它欄位與列順序。
    """
    from decimal import Decimal, getcontext
    getcontext().prec = 28

    df = df.copy()
    vals = pd.to_numeric(df[value_col], errors="coerce").fillna(0).clip(lower=0).tolist()
    dec_vals = [Decimal(str(v)) for v in vals]
    total = sum(dec_vals)

    if total <= 0:
        df[out_col] = 0
        return df

    raw = [(v * Decimal(100)) / total for v in dec_vals]
    floors = [int(x.to_integral_value(rounding="ROUND_FLOOR")) for x in raw]
    remainder = 100 - sum(floors)
    remainder = int(max(0, min(100, remainder)))

    fracs = [(raw[i] - Decimal(floors[i]), i) for i in range(len(raw))]
    fracs.sort(key=lambda t: t[0], reverse=True)  # 穩定排序：小數部份由大到小

    ints = floors[:]
    for j in range(remainder):
        if j < len(fracs):
            ints[fracs[j][1]] += 1

    df[out_col] = ints
    return df

# =========================
# 側邊欄
# =========================
with st.sidebar:
    st.header("操作說明")
    st.markdown(
        "- **本月**必要欄位（支持模糊識別）：Service Type / SSID 1 / Site Code / Wifi Technology / AP Model\n"
        "- **上月**最小欄位：Service Type / SSID 1 / Site Code（Wifi/Model 可缺省）\n"
        "- 若提供 **Hotspot Name (Chinese)**，差異與明細將同時顯示、並輸出至 Excel。\n"
        "- 所有總數與差異只計入 **六類（CTM/Managed/Mixed/Bus/Ferry/Limo）**。\n"
    )
    st.divider()
    show_unknown = st.checkbox("圖表/表格顯示 Unknown Wi‑Fi Technology", value=False)

# =========================
# 檔案上傳
# =========================
st.subheader("📁 上傳資料")
c1, c2 = st.columns(2)
with c1:
    uploaded_curr = st.file_uploader("本月資料（CSV 或 Excel）", type=["csv","xlsx","xls"], key="curr")
with c2:
    uploaded_prev = st.file_uploader("上月資料（CSV 或 Excel，可缺 Wifi/Model）", type=["csv","xlsx","xls"], key="prev")

if uploaded_curr is None:
    st.info("請至少上傳 **本月資料**。"); st.stop()

# 讀本月檔案
df_curr_raw, err1 = read_upload(uploaded_curr)
if err1:
    st.error(f"本月檔案讀取失敗：{err1}"); st.stop()

try:
    colmap_curr = resolve_columns(df_curr_raw)
    with st.expander("🧭 本月欄位對照（模糊識別結果）", expanded=False):
        st.write({
            "Service Type": colmap_curr.get("service_type"),
            "SSID 1": colmap_curr.get("ssid1"),
            "Site Code": colmap_curr.get("site_code"),
            "Wifi Technology": colmap_curr.get("wifi_technology"),
            "AP Model": colmap_curr.get("ap_model"),
            "Hotspot Name (Chinese)": colmap_curr.get("hotspot_name_cn"),
        })
except Exception as e:
    st.warning(f"本月欄位解析警告：{e}")

try:
    df_curr, df_curr6 = prepare_df(df_curr_raw, allow_missing_wifi_vendor=False)
    # 保險：正規化 Site Code
    df_curr6["Site Code"] = df_curr6["Site Code"].astype(str).str.strip()
except Exception as e:
    st.error(f"本月資料準備失敗：{e}"); st.stop()

# 讀上月檔案（可缺 Wifi/Model）
df_prev_raw, err2 = read_upload(uploaded_prev) if uploaded_prev is not None else (None, "未上傳")
if uploaded_prev is None:
    # 若未上傳上月檔案，建立空的 df_prev6 以避免後續 NameError
    df_prev_raw = None
    df_prev6 = pd.DataFrame(columns=df_curr6.columns)
else:
    if err2:
        st.warning(f"上月檔案讀取警告：{err2}")
        df_prev6 = pd.DataFrame(columns=df_curr6.columns)
    else:
        try:
            # allow_missing_wifi_vendor=True 因為上月可缺 Wifi/Model
            df_prev, df_prev6 = prepare_df(df_prev_raw, allow_missing_wifi_vendor=True)
            # 正規化 Site Code（避免 '001' vs '1' 或前後空白導致比對失敗）
            df_prev6["Site Code"] = df_prev6["Site Code"].astype(str).str.strip()
        except Exception as e:
            st.warning(f"上月資料準備失敗：{e}")
            df_prev6 = pd.DataFrame(columns=df_curr6.columns)

# === 新增：本月資料質檢（逐列定位可疑資料） + 異常列輸出 ===
def _collect_row_errors_current(df_all: pd.DataFrame, original_df: pd.DataFrame, colmap: dict) -> pd.DataFrame:
    errs = []
    col_service, col_ssid1, col_site = colmap["service_type"], colmap["ssid1"], colmap["site_code"]
    col_wifi, col_model = colmap.get("wifi_technology"), colmap.get("ap_model")

    for idx, r in df_all.iterrows():
        site_disp = r.get("Site Code", _to_str(original_df.at[idx, col_site]) if col_site else "")
        if _is_special_site(site_disp):  # Event/Idle 放行
            continue

        issues = []

        # Service Type 空白
        stype_raw = original_df.at[idx, col_service] if col_service else None
        if _safe_strip(stype_raw) == "":
            issues.append("Service Type 空白")

        # Wifi Technology 空白 or 無法識別（Unknown）
        if col_wifi is not None:
            wifi_raw = original_df.at[idx, col_wifi]
            wifi_norm = r.get("Wifi Technology (norm)")
            if _safe_strip(wifi_raw) == "":
                issues.append("Wifi Technology 空白")
            elif wifi_norm == "Unknown":
                issues.append(f"Wifi Technology 無法識別（原值: {wifi_raw}）")

        # AP Model 空白
        if col_model is not None:
            model_raw = original_df.at[idx, col_model]
            if _safe_strip(model_raw) == "":
                issues.append("AP Model 空白")

        # Managed Wi‑Fi 需要 SSID 1 協助判斷 Mixed/Managed
        if _safe_strip(stype_raw) in MANAGED_WIFI_TYPES:
            ssid_raw = original_df.at[idx, col_ssid1] if col_ssid1 else None
            if _safe_strip(ssid_raw) == "":
                issues.append("Managed Wi‑Fi 站點 SSID 1 空白，可能影響 Mixed/Managed 判定")

        # Category=Other（多為未知 Service Type 或不匹配）
        if r.get("Category") == "Other" and _safe_strip(stype_raw) != "":
            issues.append("Category=Other（可能為未知的 Service Type）")

        if issues:
            errs.append({
                "DataFrame Index": int(idx),
                "Row (1-based, 含標題假設)": int(idx) + 2,
                "Site Code": site_disp,
                "Service Type": stype_raw,
                "SSID 1": (original_df.at[idx, col_ssid1] if col_ssid1 else None),
                "Wifi Technology (raw)": (original_df.at[idx, col_wifi] if col_wifi else None),
                "AP Model (raw)": (original_df.at[idx, col_model] if col_model else None),
                "Category": r.get("Category"),
                "Error": "；".join(issues)
            })

    return pd.DataFrame(errs)

try:
    _row_errors_curr_df = _collect_row_errors_current(df_curr, df_curr_raw, colmap_curr)
    with st.expander("❗資料異常列（本月）", expanded=False):
        if not _row_errors_curr_df.empty:
            st.dataframe(_row_errors_curr_df, use_container_width=True)
            st.download_button(
                "下載異常列 CSV（本月）",
                _row_errors_curr_df.to_csv(index=False).encode("utf-8-sig"),
                "current_row_errors.csv",
                "text/csv"
            )
        else:
            st.info("本月未偵測到可疑資料列（非 Event/Idle 站點）。")
except Exception as _e:
    st.warning(f"異常列產生時發生問題：{_e}")
# === 新增結束 ===


# -------------------------
# Hotspot Statistic（放在本月六類合併統計之前）
# -------------------------
st.markdown("## 🔎 Hotspot Statistic")

# 使用者手動輸入月份顯示文字（可改為 selectbox）
col_m1, col_m2 = st.columns(2)
with col_m1:
    month_curr_input = st.text_input("本月（顯示文字，例如：二月-26）", value="")
with col_m2:
    month_prev_input = st.text_input("上月（顯示文字，例如：一月-26）", value="")

# 確保 df_curr6 / df_prev6 的 Site Code 為字串且去空白（避免比較錯誤）
if 'df_curr6' in globals() and isinstance(df_curr6, pd.DataFrame):
    df_curr6["Site Code"] = df_curr6["Site Code"].astype(str).str.strip()
if 'df_prev6' in globals() and isinstance(df_prev6, pd.DataFrame):
    df_prev6["Site Code"] = df_prev6["Site Code"].astype(str).str.strip()

prev_available = 'df_prev6' in globals() and isinstance(df_prev6, pd.DataFrame) and not df_prev6.empty
if not prev_available:
    st.info("若要完整 Hotspot Statistic（含上月比較），請上傳上月資料。")

try:
    # site-level AP counts（本月 / 上月）
    curr_site_ap = df_curr6.groupby("Site Code").size().rename("AP_Count_Curr")
    prev_site_ap = df_prev6.groupby("Site Code").size().rename("AP_Count_Prev") if prev_available else pd.Series(dtype=int)

    sites_curr = set(curr_site_ap.index.tolist())
    sites_prev = set(prev_site_ap.index.tolist())

    new_sites = sites_curr - sites_prev
    ceased_sites = sites_prev - sites_curr
    common_sites = sites_curr & sites_prev

    # New / Cessation AP 計算
    new_ap_from_new_sites = int(curr_site_ap.loc[list(new_sites)].sum()) if new_sites else 0
    ap_diff_common = (curr_site_ap.reindex(list(common_sites), fill_value=0) - prev_site_ap.reindex(list(common_sites), fill_value=0))
    new_ap_from_increase = int(ap_diff_common[ap_diff_common > 0].sum()) if not ap_diff_common.empty else 0
    new_ap_total = new_ap_from_new_sites + new_ap_from_increase

    ceased_ap_from_sites = int(prev_site_ap.loc[list(ceased_sites)].sum()) if ceased_sites else 0
    ceased_ap_from_decrease = int((-ap_diff_common[ap_diff_common < 0]).sum()) if not ap_diff_common.empty else 0
    ceased_ap_total = ceased_ap_from_sites + ceased_ap_from_decrease

    new_site_count = len(new_sites)
    ceased_site_count = len(ceased_sites)

    # 分類顯示名稱對應（合併為 Fixed Hotspot / Transportation）
    category_map = {
        "CTM WiFi": ("Fixed Hotspot", "Fixed Wi-Fi (CTM Wi-Fi Hotspot)"),
        "Managed WiFi": ("Fixed Hotspot", "Fixed Wi-Fi (Managed Wi-Fi)"),
        "Mixed Site": ("Fixed Hotspot", "Fixed Wi-Fi (CTM Wi-Fi Hotspot & Managed Wi-Fi)"),
        "Bus WiFi": ("Transportation", "Public Bus Wi-Fi (CTM Wi-Fi Hotspot)"),
        "Ferry WiFi": ("Transportation", "Ferry (Managed Wi-Fi)"),
        "Limo WiFi": ("Transportation", "Limo / Shuttle (Managed Wi-Fi)")
    }

    # 聚合函式（回傳 Category 順序的 Site 與 AP）
    def agg_by_category(df6):
        if df6 is None or df6.empty:
            return pd.DataFrame({
                "Category": CATEGORY_ORDER,
                "Site_Count": [0]*len(CATEGORY_ORDER),
                "AP_Count": [0]*len(CATEGORY_ORDER)
            })
        site_counts = df6.groupby("Category")["Site Code"].nunique().reindex(CATEGORY_ORDER, fill_value=0)
        ap_counts = df6.groupby("Category").size().reindex(CATEGORY_ORDER, fill_value=0)
        return pd.DataFrame({
            "Category": CATEGORY_ORDER,
            "Site_Count": site_counts.values,
            "AP_Count": ap_counts.values
        })

    agg_curr = agg_by_category(df_curr6)
    agg_prev = agg_by_category(df_prev6) if prev_available else agg_by_category(pd.DataFrame(columns=df_curr6.columns))

    total_site_by_cat = agg_curr.set_index("Category")["Site_Count"].reindex(CATEGORY_ORDER, fill_value=0).astype(int)
    total_ap_by_cat = agg_curr.set_index("Category")["AP_Count"].reindex(CATEGORY_ORDER, fill_value=0).astype(int)
    prev_site_by_cat = agg_prev.set_index("Category")["Site_Count"].reindex(CATEGORY_ORDER, fill_value=0).astype(int)
    prev_ap_by_cat = agg_prev.set_index("Category")["AP_Count"].reindex(CATEGORY_ORDER, fill_value=0).astype(int)

    # 構建輸出資料（Prev 在 Curr 前）
    rows = []
    # New Installation
    rows.append({"Section":"New Installation","Category":"No. of Site","Prev": "", "Curr": new_site_count, "vs": ""})
    rows.append({"Section":"New Installation","Category":"No. of AP","Prev": "", "Curr": new_ap_total, "vs": ""})
    # Cessation
    rows.append({"Section":"Cessation","Category":"No. of Site","Prev": "", "Curr": ceased_site_count, "vs": ""})
    rows.append({"Section":"Cessation","Category":"No. of AP","Prev": "", "Curr": ceased_ap_total, "vs": ""})

    # Total of Site (顯示 vs)
    total_site_curr = int(total_site_by_cat.sum())
    total_site_prev = int(prev_site_by_cat.sum())
    rows.append({"Section":"Total","Category":"Total of Site","Prev": total_site_prev, "Curr": total_site_curr, "vs": int(total_site_curr - total_site_prev)})

    # Per-category Site rows (Prev shown, vs empty) — will be grouped into Fixed/Transportation in HTML
    for cat in CATEGORY_ORDER:
        sect, label = category_map.get(cat, ("Other", cat))
        rows.append({
            "Section":"Total",
            "Category": label,
            "Prev": int(prev_site_by_cat[cat]),
            "Curr": int(total_site_by_cat[cat]),
            "vs": ""
        })

    # Total of AP (顯示 vs)
    total_ap_curr = int(total_ap_by_cat.sum())
    total_ap_prev = int(prev_ap_by_cat.sum())
    rows.append({"Section":"Total_AP","Category":"Total of AP","Prev": total_ap_prev, "Curr": total_ap_curr, "vs": int(total_ap_curr - total_ap_prev)})

    # Per-category AP rows
    for cat in CATEGORY_ORDER:
        sect, label = category_map.get(cat, ("Other", cat))
        rows.append({
            "Section":"Total_AP",
            "Category": label,
            "Prev": int(prev_ap_by_cat[cat]),
            "Curr": int(total_ap_by_cat[cat]),
            "vs": ""
        })

    hotspot_stat_df = pd.DataFrame(rows)

    # 轉成 HTML table 以支援合併單元格與三塊分隔（Prev 在 Curr 前）
    def build_html_table(df, month_prev_label, month_curr_label):
        css = """
        <style>
        table.hotstat {border-collapse: collapse; width:100%; font-size:14px;}
        table.hotstat th, table.hotstat td {border:1px solid #ddd; padding:6px; text-align:center;}
        table.hotstat th {background:#f7f7f7; font-weight:700;}
        .section-title {font-weight:700; text-align:left; padding:8px 0;}
        .sep {border-top:3px solid #333;}
        .bold-row td {font-weight:700; background:#fafafa;}
        .group-cell {font-weight:700; background:#fff; vertical-align:middle;}
        </style>
        """
        html = [css, "<table class='hotstat'>"]
        # header (Prev before Curr)
        html.append(f"<tr><th>Section</th><th>Category</th><th>{month_prev_label}</th><th>{month_curr_label}</th><th>vs Previous Month</th></tr>")

        # 1) New Installation & Cessation block (first 4 rows)
        # New Installation (2 rows)
        for i in range(0,2):
            r = df.iloc[i]
            if i == 0:
                html.append(f"<tr><td rowspan='2'>New Installation</td><td>{r['Category']}</td><td>{r['Prev']}</td><td>{r['Curr']}</td><td>{r['vs']}</td></tr>")
            else:
                html.append(f"<tr><td>{r['Category']}</td><td>{r['Prev']}</td><td>{r['Curr']}</td><td>{r['vs']}</td></tr>")

        # Cessation (2 rows)
        html.append("<tr class='sep'></tr>")  # thin visual separator between New and Cessation
        for i in range(2,4):
            r = df.iloc[i]
            if i == 2:
                html.append(f"<tr><td rowspan='2'>Cessation</td><td>{r['Category']}</td><td>{r['Prev']}</td><td>{r['Curr']}</td><td>{r['vs']}</td></tr>")
            else:
                html.append(f"<tr><td>{r['Category']}</td><td>{r['Prev']}</td><td>{r['Curr']}</td><td>{r['vs']}</td></tr>")

        # 大分隔線：結束 New/Cessation，開始 Total of Site block
        html.append("<tr class='sep'></tr>")

        # 2) Total of Site block (Total of Site row + grouped Fixed Hotspot / Transportation site rows)
        r_tot_site = df[df['Category']=="Total of Site"].iloc[0]
        html.append(f"<tr class='bold-row'><td> </td><td>{r_tot_site['Category']}</td><td>{r_tot_site['Prev']}</td><td>{r_tot_site['Curr']}</td><td>{r_tot_site['vs']}</td></tr>")

        # prepare labels
        fixed_labels = [category_map[c][1] for c in ["CTM WiFi","Managed WiFi","Mixed Site"]]
        trans_labels = [category_map[c][1] for c in ["Bus WiFi","Ferry WiFi","Limo WiFi"]]

        site_rows = df[df['Section']=="Total"].set_index("Category")

        # Fixed Hotspot group (merge group cell)
        html.append(f"<tr><td class='group-cell' rowspan='{len(fixed_labels)}'>Fixed Hotspot</td><td>{fixed_labels[0]}</td><td>{site_rows.loc[fixed_labels[0],'Prev']}</td><td>{site_rows.loc[fixed_labels[0],'Curr']}</td><td></td></tr>")
        for lbl in fixed_labels[1:]:
            html.append(f"<tr><td>{lbl}</td><td>{site_rows.loc[lbl,'Prev']}</td><td>{site_rows.loc[lbl,'Curr']}</td><td></td></tr>")

        # Transportation group
        html.append(f"<tr><td class='group-cell' rowspan='{len(trans_labels)}'>Transportation</td><td>{trans_labels[0]}</td><td>{site_rows.loc[trans_labels[0],'Prev']}</td><td>{site_rows.loc[trans_labels[0],'Curr']}</td><td></td></tr>")
        for lbl in trans_labels[1:]:
            html.append(f"<tr><td>{lbl}</td><td>{site_rows.loc[lbl,'Prev']}</td><td>{site_rows.loc[lbl,'Curr']}</td><td></td></tr>")

        # 大分隔線：結束 Site block，開始 AP block
        html.append("<tr class='sep'></tr>")

        # 3) Total of AP block (Total of AP row + grouped AP rows)
        r_tot_ap = df[df['Category']=="Total of AP"].iloc[0]
        html.append(f"<tr class='bold-row'><td> </td><td>{r_tot_ap['Category']}</td><td>{r_tot_ap['Prev']}</td><td>{r_tot_ap['Curr']}</td><td>{r_tot_ap['vs']}</td></tr>")

        ap_rows = df[df['Section']=="Total_AP"].set_index("Category")

        # Fixed Hotspot AP rows
        html.append(f"<tr><td class='group-cell' rowspan='{len(fixed_labels)}'>Fixed Hotspot</td><td>{fixed_labels[0]}</td><td>{ap_rows.loc[fixed_labels[0],'Prev']}</td><td>{ap_rows.loc[fixed_labels[0],'Curr']}</td><td></td></tr>")
        for lbl in fixed_labels[1:]:
            html.append(f"<tr><td>{lbl}</td><td>{ap_rows.loc[lbl,'Prev']}</td><td>{ap_rows.loc[lbl,'Curr']}</td><td></td></tr>")

        # Transportation AP rows
        html.append(f"<tr><td class='group-cell' rowspan='{len(trans_labels)}'>Transportation</td><td>{trans_labels[0]}</td><td>{ap_rows.loc[trans_labels[0],'Prev']}</td><td>{ap_rows.loc[trans_labels[0],'Curr']}</td><td></td></tr>")
        for lbl in trans_labels[1:]:
            html.append(f"<tr><td>{lbl}</td><td>{ap_rows.loc[lbl,'Prev']}</td><td>{ap_rows.loc[lbl,'Curr']}</td><td></td></tr>")

        html.append("</table>")
        return "\n".join(html)

    month_prev_label = month_prev_input or "Previous"
    month_curr_label = month_curr_input or "Current"
    html_table = build_html_table(hotspot_stat_df, month_prev_label, month_curr_label)
    st.markdown(html_table, unsafe_allow_html=True)

    # 下載：以 DataFrame（Prev 在前）匯出
    export_df = hotspot_stat_df.copy()[["Category","Prev","Curr","vs"]]
    export_df = export_df.rename(columns={"Prev": month_prev_label, "Curr": month_curr_label, "vs":"vs Previous Month"})
    csv_bytes = export_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button("下載 Hotspot Statistic CSV", csv_bytes, file_name="hotspot_statistic.csv", mime="text/csv")

    # Excel
    try:
        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
            export_df.to_excel(writer, index=False, sheet_name="HotspotStatistic")
        # context manager 已經寫入並關閉 writer
        towrite.seek(0)
        excel_bytes = towrite.getvalue()
        st.download_button(
            "下載 Hotspot Statistic Excel",
            excel_bytes,
            file_name="hotspot_statistic.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as _e:
        st.warning(f"Excel 下載暫時不可用：{_e}")

except Exception as e:
    st.error(f"產生 Hotspot Statistic 時發生錯誤：{e}")

# =========================
# 本月（僅六類）統計與視覺化（重構：先預計算，再顯示）
# =========================
st.subheader("📊 本月六類合併統計")
category_dfs_curr = {cat: df_curr6[df_curr6["Category"] == cat].copy() for cat in CATEGORY_ORDER}
# 安全 concat（若所有子表都為空，回傳空 DataFrame）
vals = [v for v in category_dfs_curr.values() if not v.empty]
combined_df = pd.concat(vals, ignore_index=True) if vals else pd.DataFrame(columns=df_curr6.columns)

m1, m2, m3, m4 = st.columns(4)
m1.metric("本月：總 AP（六類）", len(combined_df))
m2.metric("本月：總 Site（六類去重）", combined_df["Site Code"].nunique() if not combined_df.empty else 0)
m3.metric("本月：Huawei（六類）", int((combined_df["Vendor"]=="Huawei").sum()) if not combined_df.empty else 0)
m4.metric("本月：Ruckus（六類）", int((combined_df["Vendor"]=="Ruckus").sum()) if not combined_df.empty else 0)

# ---------- 預計算：建立 summary_rows / per_category tables（不做 expanders） ----------
summary_rows = []
per_category_wifi_tables = {}
per_category_vendor_tables = {}
pct_rows = []
site_count_map = {}

for name in CATEGORY_ORDER:
    sub_df = category_dfs_curr.get(name, pd.DataFrame(columns=df_curr6.columns))
    ap_count = len(sub_df)
    site_count = sub_df['Site Code'].nunique() if not sub_df.empty else 0
    site_count_map[name] = site_count

    wifi_series = count_wifi_tech_series(sub_df) if not sub_df.empty else pd.Series(0, index=WIFI_LEVELS_DISPLAY)
    vendor_series = sub_df["Vendor"].value_counts().reindex(["Huawei","Ruckus"], fill_value=0) if not sub_df.empty else pd.Series([0,0], index=["Huawei","Ruckus"])

    # 準備顯示用表格
    wifi_full = sub_df["Wifi Technology (norm)"].value_counts() if not sub_df.empty else pd.Series(dtype=int)
    wifi_df = series_to_df_for_chart(wifi_series, "Wi‑Fi Technology", "Count", show_unknown, wifi_full)
    per_category_wifi_tables[name] = wifi_df.copy()

    vendor_df = vendor_series.rename_axis("Vendor").reset_index(name="Count")
    per_category_vendor_tables[name] = vendor_df.copy()

    total_c = wifi_df["Count"].sum()
    if total_c == 0:
        for tech in wifi_df["Wi‑Fi Technology"].tolist():
            pct_rows.append({"Category": name, "Wi‑Fi Technology": tech, "Count": 0, "Percent": 0.0})
    else:
        for _, r in wifi_df.iterrows():
            pct_rows.append({
                "Category": name,
                "Wi‑Fi Technology": r["Wi‑Fi Technology"],
                "Count": int(r["Count"]),
                "Percent": round(100.0*r["Count"]/total_c, 6)
            })

    summary_rows.append({
        "Category": name,
        "AP Count": ap_count,
        "Site Count": site_count,
        "Huawei": int(vendor_series.get("Huawei",0)),
        "Ruckus": int(vendor_series.get("Ruckus",0)),
        "Wi‑Fi 4": int(wifi_series.get("Wi‑Fi 4",0)),
        "Wi‑Fi 5": int(wifi_series.get("Wi‑Fi 5",0)),
        "Wi‑Fi 6": int(wifi_series.get("Wi‑Fi 6",0)),
        "Wi‑Fi 7": int(wifi_series.get("Wi‑Fi 7",0)),
    })
# ---------- 預計算結束 ----------

# -------------------------
# 1) Managed/CTM Hotspot 對全網 AP 的占比（顯示）
# -------------------------
st.markdown("### 📈 Managed/CTM Hotspot 對全網 AP 的占比")

# 全網 AP（六類）
total_ap_all = len(combined_df)

# Managed Wi‑Fi（四類）
managed_df = df_curr6[df_curr6["Category"].isin(MANAGED_ORIGINAL_CATEGORIES)].copy()
managed_ap = len(managed_df)

# CTM Hotspot（CTM WiFi + Bus WiFi）
ctm_hotspot_ap = len(df_curr6[df_curr6["Category"].isin(["CTM WiFi", "Bus WiFi"])])

# 分母安全處理：回傳整數百分比字串（例如 "57%"）
def _pct_int_str(part: int, whole: int) -> str:
    if not whole:
        return "0%"
    return f"{int(round(part / whole * 100, 0))}%"

managed_pct_str = _pct_int_str(managed_ap, total_ap_all)
ctm_pct_str = _pct_int_str(ctm_hotspot_ap, total_ap_all)

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    st.markdown(
        f"**Managed Wi‑Fi（四類）**  \n"
        f"{managed_ap:,} AP（{managed_pct_str}）"
    )
with col2:
    st.markdown(
        f"**CTM Hotspot（CTM WiFi + Bus WiFi）**  \n"
        f"{ctm_hotspot_ap:,} AP（{ctm_pct_str}）"
    )
with col3:
    st.markdown(
        f"**全網 AP（六類）**  \n"
        f"{total_ap_all:,} AP"
    )

share_rows = [
    {"組別": "Managed Wi‑Fi（四類）", "AP 數": managed_ap, "占全網比例": managed_pct_str},
    {"組別": "CTM Hotspot（CTM WiFi + Bus WiFi）", "AP 數": ctm_hotspot_ap, "占全網比例": ctm_pct_str},
]
df_share = pd.DataFrame(share_rows)
st.dataframe(df_share, use_container_width=True)

# -------------------------
# 2) 本月 - 各分類彙總（六類） 顯示（使用預計算的 summary_rows）
# -------------------------
# 加上 Total 行（確保只加一次）
if not any(r.get("Category") == "Total" for r in summary_rows):
    summary_rows.append({
        "Category": "Total",
        "Site Count": combined_df["Site Code"].nunique() if not combined_df.empty else 0,
        "Wi‑Fi 4": int((combined_df["Wifi Technology (norm)"]=="Wi‑Fi 4").sum()) if not combined_df.size else 0,
        "Wi‑Fi 5": int((combined_df["Wifi Technology (norm)"]=="Wi‑Fi 5").sum()) if combined_df.size else 0,
        "Wi‑Fi 6": int((combined_df["Wifi Technology (norm)"]=="Wi‑Fi 6").sum()) if combined_df.size else 0,
        "Wi‑Fi 7": int((combined_df["Wifi Technology (norm)"]=="Wi‑Fi 7").sum()) if combined_df.size else 0,
        "Ruckus": int((combined_df["Vendor"]=="Ruckus").sum()) if not combined_df.empty else 0,
        "Huawei": int((combined_df["Vendor"]=="Huawei").sum()) if not combined_df.empty else 0,
        "AP Count": combined_df.shape[0]
    })

summary_df = pd.DataFrame(summary_rows)

COL_ORDER = [
    "Category",
    "Site Count",
    "Wi‑Fi 4",
    "Wi‑Fi 5",
    "Wi‑Fi 6",
    "Wi‑Fi 7",
    "Ruckus",
    "Huawei",
    "AP Count"
]

summary_df = summary_df[[c for c in COL_ORDER if c in summary_df.columns]]

st.markdown("### 📑 本月 - 各分類彙總（六類）")
st.dataframe(summary_df, use_container_width=True)

# -------------------------
# 3) 本月 - Managed Wi‑Fi 各分類彙總（四類） 顯示（使用獨立計算）
# -------------------------
st.subheader("📊 本月 - Managed Wi‑Fi 各分類彙總（四類）")

managed_dfs_curr = {cat: df_curr6[df_curr6["Category"] == cat].copy() for cat in MANAGED_ORIGINAL_CATEGORIES}
vals_man = [v for v in managed_dfs_curr.values() if not v.empty]
combined_managed_df = pd.concat(vals_man, ignore_index=True) if vals_man else pd.DataFrame(columns=df_curr6.columns)

summary_rows_managed = []
for name in MANAGED_ORIGINAL_CATEGORIES:
    sub_df = managed_dfs_curr.get(name, pd.DataFrame(columns=df_curr6.columns))
    ap_count = len(sub_df)
    site_count = sub_df['Site Code'].nunique() if not sub_df.empty else 0
    wifi_series = count_wifi_tech_series(sub_df) if not sub_df.empty else pd.Series(0, index=WIFI_LEVELS_DISPLAY)
    vendor_series = sub_df["Vendor"].value_counts().reindex(["Huawei","Ruckus"], fill_value=0) if not sub_df.empty else pd.Series([0,0], index=["Huawei","Ruckus"])

    summary_rows_managed.append({
        "Category": MANAGED_GROUP_NAMES[name],
        "Site Count": site_count,
        "Wi‑Fi 4": int(wifi_series.get("Wi‑Fi 4",0)),
        "Wi‑Fi 5": int(wifi_series.get("Wi‑Fi 5",0)),
        "Wi‑Fi 6": int(wifi_series.get("Wi‑Fi 6",0)),
        "Wi‑Fi 7": int(wifi_series.get("Wi‑Fi 7",0)),
        "Ruckus": int(vendor_series.get("Ruckus",0)),
        "Huawei": int(vendor_series.get("Huawei",0)),
        "AP Count": ap_count
    })

# 加上 Total 行（確保只加一次）
if not any(r.get("Category") == "Total" for r in summary_rows_managed):
    summary_rows_managed.append({
        "Category": "Total",
        "AP Count": combined_managed_df.shape[0],
        "Site Count": combined_managed_df["Site Code"].nunique() if not combined_managed_df.empty else 0,
        "Huawei": int((combined_managed_df["Vendor"]=="Huawei").sum()) if not combined_managed_df.empty else 0,
        "Ruckus": int((combined_managed_df["Vendor"]=="Ruckus").sum()) if not combined_managed_df.empty else 0,
        "Wi‑Fi 4": int((combined_managed_df["Wifi Technology (norm)"]=="Wi‑Fi 4").sum()) if combined_managed_df.size else 0,
        "Wi‑Fi 5": int((combined_managed_df["Wifi Technology (norm)"]=="Wi‑Fi 5").sum()) if combined_managed_df.size else 0,
        "Wi‑Fi 6": int((combined_managed_df["Wifi Technology (norm)"]=="Wi‑Fi 6").sum()) if combined_managed_df.size else 0,
        "Wi‑Fi 7": int((combined_managed_df["Wifi Technology (norm)"]=="Wi‑Fi 7").sum()) if combined_managed_df.size else 0
    })

summary_df_managed = pd.DataFrame(summary_rows_managed)

COL_ORDER = [
    "Category",
    "Site Count",
    "Wi‑Fi 4",
    "Wi‑Fi 5",
    "Wi‑Fi 6",
    "Wi‑Fi 7",
    "Ruckus",
    "Huawei",
    "AP Count"
]

summary_df_managed = summary_df_managed[
    [c for c in COL_ORDER if c in summary_df_managed.columns]
]

st.dataframe(summary_df_managed, use_container_width=True)

# -------------------------
# 4) 最後顯示每分類的 expanders（使用預計算的 per_category_* 表格）
# 占比圖（分類單行；Sites 放右側縱向；圖例右移到圖外避免重疊，legend.x=1.10）
st.subheader("🧩 總 AP（六類）的 Wi‑Fi Technology 占比圖")
df_pct = pd.DataFrame(pct_rows)
if not df_pct.empty:
    # 取得每個分類的 Sites 數
    def _sites_of(cat: str) -> int:
        m = summary_df.loc[summary_df["Category"] == cat, "Site Count"]
        return int(m.values[0]) if len(m) else 0

    # 分類顯示為單行（不把 Sites 混入）
    display_single_line_map = {
        c: f"{CATEGORY_DISPLAY_NAMES.get(c, c)}"
        for c in CATEGORY_ORDER
    }
    df_pct["Category_Display"] = df_pct["Category"].map(display_single_line_map)
    ordered_display = [display_single_line_map[c] for c in CATEGORY_ORDER]
    tech_order = WIFI_LEVELS_DISPLAY + (["Unknown"] if show_unknown else [])

    # === 先做每列×技術去重聚合 → 整數分配 ===
    df_pct_base = (
        df_pct
        .groupby(["Category_Display", "Wi‑Fi Technology"], as_index=False, sort=False)
        .agg(Count=("Count", "sum"))
    )
    df_pct_int = assign_integer_percent(
        df_pct_base,
        group_col="Category_Display",
        count_col="Count",
        out_col="PercentInt"
    )

    fig_stacked = px.bar(
        df_pct_int, x="PercentInt", y="Category_Display",
        color="Wi‑Fi Technology",
        category_orders={"Category_Display": ordered_display, "Wi‑Fi Technology": tech_order},
        orientation="h", barmode="stack", color_discrete_map=COLOR_MAP,
        title=None
    )
    fig_stacked.update_traces(texttemplate="%{x}%", textposition="inside", insidetextanchor="middle")

    # 套用統一樣式 —— 收一點繪圖區、加寬右側空間
    fig_stacked = apply_clean_layout(
        fig_stacked,
        "Distribution of Wi‑Fi Technology",
        remove_y_title=True,
        percent_axis=True,
        x_domain_end=0.86,   # 往左收（可視覺微調 0.84~0.90）
        right_margin=300,    # 預留 Sites + Legend 空間
        legend_x=0.99        # 先貼圖內最右，再補丁推出圖外
    )

    # 圖例移到圖外右側（不與右側 Sites 重疊）
    fig_stacked.update_layout(legend=dict(
        x=1.10,               # 指定：>1.0 在繪圖區外（你要求的 1.10）
        xanchor="left",
        y=0.5, yanchor="middle",
        bgcolor="rgba(255,255,255,0)"
    ))
    fig_stacked.update_layout(margin=dict(r=340))  # 右邊距加大以避免裁切

    # 在右側貼上每個分類的 Sites 數（視覺上像右側縱軸）
    annotations = list(fig_stacked.layout.annotations) if fig_stacked.layout.annotations else []
    for cat in CATEGORY_ORDER:
        y_label = display_single_line_map[cat]
        sites_n = _sites_of(cat)
        annotations.append(dict(
            xref="x", yref="y",
            x=100,                 # 百分比軸最右端
            y=y_label,             # 對應這一列的 y 類別
            text=f"Sites: {sites_n}",
            showarrow=False,
            xanchor="left",
            align="left",
            font=dict(size=22, color="#000"),  # 右側 Sites 字號 22
            xshift=10              # 往右挪一點，避免貼在線上
        ))
    fig_stacked.update_layout(annotations=annotations)

    # 類別刻度字號（若要統一 22）
    fig_stacked.update_yaxes(tickfont=dict(size=22))

    st.plotly_chart(fig_stacked, use_container_width=True, config=PLOTLY_CONFIG)

st.markdown("") 

# Managed Wi‑Fi（四類）— 圖表 + 統計表（右側 Sites + 圖例外推，legend.x=1.10）
# =========================
st.subheader("🧩 Managed Wi‑Fi（四類）的 Wi‑Fi Technology 占比圖")
managed_df = df_curr6[df_curr6["Category"].isin(MANAGED_ORIGINAL_CATEGORIES)].copy()
if managed_df.empty:
    st.info("目前六類資料中沒有 Managed/Mixed/Ferry/Limo 類別的 AP。")
else:
    # 標準化 Category(Eng)：去除首尾空格，統一連字符為 ASCII '-'，避免鍵不一致
    managed_df["Category(Eng)"] = managed_df["Category"].map(MANAGED_GROUP_NAMES).astype(str)
    managed_df["Category(Eng)"] = (
        managed_df["Category(Eng)"]
        .str.strip()
        .str.replace("‑", "-", regex=False)
        .str.replace("–", "-", regex=False)
        .str.replace("—", "-", regex=False)
    )
    
    # === Managed 與 CTM Hotspot 對全網 AP 的占比（百分比字串、四捨五入為整數） ===
    # 全網 AP（六類）
    total_ap_all = len(combined_df)  # 你前面已經組好的六類合併 df

    # Managed Wi‑Fi（四類）：Managed/Mixed/Ferry/Limo
    managed_ap = len(managed_df)

    # CTM Hotspot：CTM WiFi + Bus WiFi（仍僅限六類的範圍內）
    ctm_hotspot_ap = len(df_curr6[df_curr6["Category"].isin(["CTM WiFi", "Bus WiFi"])])

    # 分母安全處理：回傳整數百分比字串（例如 "57%"）
    def _pct_int_str(part: int, whole: int) -> str:
        if not whole:
            return "0%"
        return f"{int(round(part / whole * 100, 0))}%"

    managed_pct_str = _pct_int_str(managed_ap, total_ap_all)
    ctm_pct_str     = _pct_int_str(ctm_hotspot_ap, total_ap_all)

    # =========================
    # 本月 Managed Wi‑Fi 四類彙總
    # =========================

    # 1) 占比（橫條堆疊，顯示 Sites 在右側）
    rows = []
    display_order = []
    for cat_eng in MANAGED_GROUP_ORDER:
        cat_key = (
            str(cat_eng).strip()
            .replace("‑", "-").replace("–", "-").replace("—", "-")
        )
        sub = managed_df[managed_df["Category(Eng)"] == cat_key].copy()
        sites_this = sub["Site Code"].nunique()

        wifi_series = count_wifi_tech_series(sub)
        for tech in WIFI_LEVELS_DISPLAY:
            rows.append({
                "Category": f"{cat_key}（Sites: {sites_this}）",
                "Wi‑Fi Technology": tech,
                "Count": int(wifi_series.get(tech, 0))
            })
        display_order.append(f"{cat_key}（Sites: {sites_this}）")

    df_mgd = pd.DataFrame(rows)

    if df_mgd["Count"].sum() > 0:
        df_mgd["Percent"] = df_mgd["Count"] / df_mgd.groupby("Category")["Count"].transform("sum") * 100

        # 準備四類的 Sites 數（直接用 managed_df 保證鍵一致）
        sites_map_mgd = (
            managed_df.groupby("Category(Eng)")["Site Code"]
            .nunique()
            .to_dict()
        )

        # 類別鍵的標準順序（全部轉為 ASCII '-'）
        ordered_keys_mgd = [
            str(k).strip().replace("‑","-").replace("–","-").replace("—","-")
            for k in MANAGED_GROUP_ORDER
        ]
        display_single_line_map_mgd = {k: k for k in ordered_keys_mgd}

        # 將 df_mgd 由『名稱（Sites: N）』轉回『純鍵』
        df_mgd_plot = df_mgd.copy()
        df_mgd_plot["Category_Key"] = df_mgd_plot["Category"].apply(
            lambda s: re.sub(r"（.*?）$", "", str(s)).strip()
        )
        df_mgd_plot["Category_Key"] = df_mgd_plot["Category_Key"].apply(
            lambda s: re.sub(r"\(.*?\)$", "", str(s)).strip()
        )

        # 顯示欄位（單行）
        df_mgd_plot["Category_Display"] = df_mgd_plot["Category_Key"].map(display_single_line_map_mgd)

        # 類別順序（單行顯示）
        ordered_display_mgd = [display_single_line_map_mgd[c] for c in ordered_keys_mgd]

        # === 先做每列×技術去重聚合 → 整數分配 ===
        df_mgd_plot_base = (
            df_mgd_plot
            .groupby(["Category_Display", "Wi‑Fi Technology"], as_index=False, sort=False)
            .agg(Count=("Count", "sum"))
        )
        df_mgd_plot_int = assign_integer_percent(
            df_mgd_plot_base,
            group_col="Category_Display",
            count_col="Count",
            out_col="PercentInt"
        )

        # 繪圖（水平整數百分比堆疊條）
        fig_mgd = px.bar(
            df_mgd_plot_int,
            x="PercentInt",
            y="Category_Display",
            color="Wi‑Fi Technology",
            category_orders={"Category_Display": ordered_display_mgd, "Wi‑Fi Technology": WIFI_LEVELS_DISPLAY},
            orientation="h",
            barmode="stack",
            color_discrete_map=COLOR_MAP,
            title=None,
            text="PercentInt"
        )
        fig_mgd.update_traces(texttemplate="%{x}%", textposition="inside", insidetextanchor="middle")

        # 套用統一樣式 —— 收一點繪圖區、加寬右側
        fig_mgd = apply_clean_layout(
            fig_mgd,
            "Distribution of Wi‑Fi Technology for Managed Wi‑Fi",
            remove_y_title=True,
            percent_axis=True,
            x_domain_end=0.86,   # 往左收，空出右側位置
            right_margin=300,    # 預留 Sites + Legend 空間
            legend_x=0.99        # 先貼圖內最右，再補丁推出圖外
        )
        
        # 將標題向左移（0.0=最左，0.5=置中）
        fig_mgd.update_layout(title=dict(x=0.15, xanchor="left"))
        # 也可以用簡寫：fig_mgd.update_layout(title_x=0.15)

        # 圖例移到圖外右側（不與右側 Sites 重疊）
        fig_mgd.update_layout(legend=dict(
            x=1.01,               # 指定：>1.0 在繪圖區外（與全網一致）
            xanchor="left",
            y=0.5, yanchor="middle",
            bgcolor="rgba(255,255,255,0)"
        ))
        fig_mgd.update_layout(margin=dict(r=340))  # 右邊距加大以避免裁切

        # 在右側貼上每個類別的 Sites 數
        annotations = list(fig_mgd.layout.annotations) if fig_mgd.layout.annotations else []
        for cat_key in ordered_keys_mgd:
            y_label = display_single_line_map_mgd[cat_key]
            sites_n = int(sites_map_mgd.get(cat_key, 0))
            annotations.append(dict(
                xref="x", yref="y",
                x=100,                 # 百分比軸最右端
                y=y_label,             # 對應這一列 y 類別
                text=f"Sites: {sites_n}",
                showarrow=False,
                xanchor="left",
                align="left",
                font=dict(size=22, color="#000"),
                xshift=10
            ))
        fig_mgd.update_layout(annotations=annotations)

        # 類別刻度字號（可統一 22）
        fig_mgd.update_yaxes(tickfont=dict(size=22))

        st.plotly_chart(fig_mgd, use_container_width=True, config=PLOTLY_CONFIG)


st.markdown("") 
st.markdown("") 
# =========================
# 總 AP（六類）Wi‑Fi Technology：表 + 「全網 AP」餅圖（數量，整數百分比）
# =========================
st.markdown("### 🔹 本月總 AP（六類）的 Wi‑Fi Technology 統計")

# 原本的統計邏輯
total_wifi_tech = count_wifi_tech_series(combined_df)
total_wifi_full = combined_df["Wifi Technology (norm)"].value_counts()
total_wifi_df = series_to_df_for_chart(
    total_wifi_tech,
    "Wi‑Fi Technology",
    "Count",
    show_unknown,
    total_wifi_full
)

# ✅ 新增：百分比（整數、四捨五入、字串格式 xx%）— 表格
_total = int(total_wifi_df["Count"].sum())
if _total > 0:
    total_wifi_df["Percent"] = (
        (total_wifi_df["Count"] / _total * 100)
        .round(0)
        .astype(int)
        .astype(str)
        + "%"
    )
else:
    total_wifi_df["Percent"] = "0%"

# （可選）整理欄位順序
total_wifi_df = total_wifi_df[["Wi‑Fi Technology", "Count", "Percent"]]

# 顯示表格
st.dataframe(total_wifi_df, use_container_width=True)

# === 新增：餅圖用整數百分比（合計=100） + 覆寫文字 ===
_total_pie_df = compute_pie_integer_percent(total_wifi_df, value_col="Count", out_col="PercentInt")
_wifi_pie_text = [f"{int(c)}，{int(p)}%" for c, p in zip(_total_pie_df["Count"], _total_pie_df["PercentInt"])]

# 餅圖沿用 Count，不需修改
fig_wifi_pie = make_square_pie_12cm_overall(
    df=total_wifi_df,
    names="Wi‑Fi Technology",
    values="Count",
    title_text="Total number of APs",
    color_col="Wi‑Fi Technology",
    color_discrete_map=COLOR_MAP,
    show_value_and_percent=True,
    text_values=_wifi_pie_text   # ← 使用預先分配好的整數百分比文字
)
center_l, center_c, center_r = st.columns([1, 2, 1])
with center_c:
    st.plotly_chart(fig_wifi_pie, use_container_width=False, config=PLOTLY_CONFIG)

# =========================
# 總 AP（六類）Vendor：表 + 「全網 AP」餅圖（僅整數百分比）
# =========================
st.markdown("### 🔹 本月總 AP（六類）的 Vendor 統計")

total_vendor = (
    combined_df["Vendor"]
    .value_counts()
    .reindex(["Huawei", "Ruckus"], fill_value=0)
)

total_vendor_df = (
    total_vendor
    .rename_axis("Vendor")
    .reset_index(name="Count")
)

# ✅ 計算百分比
_total = int(total_vendor_df["Count"].sum())
if _total > 0:
    total_vendor_df["Percent"] = (
        (total_vendor_df["Count"] / _total * 100)
        .round(0)
        .astype(int)
        .astype(str)
        + "%"
    )
else:
    total_vendor_df["Percent"] = "0%"

# 顯示表格
st.dataframe(
    total_vendor_df[["Vendor", "Count", "Percent"]],
    use_container_width=True
)

# === 新增：餅圖用整數百分比（合計=100） + 覆寫文字（僅百分比） ===
_total_vendor_pie_df = compute_pie_integer_percent(total_vendor_df, value_col="Count", out_col="PercentInt")
_vendor_pie_text = [f"{int(p)}%" for p in _total_vendor_pie_df["PercentInt"]]

fig_vendor_pie = make_square_pie_12cm_overall(
    df=total_vendor_df,
    names="Vendor",
    values="Count",
    title_text="Brand Distribution of APs",
    color_col="Vendor",
    color_discrete_map=COLOR_MAP,
    show_value_and_percent=False,  # 僅整數百分比
    text_values=_vendor_pie_text   # ← 使用預先分配好的整數百分比文字
)
center_l, center_c, center_r = st.columns([1, 2, 1])
with center_c:
    st.plotly_chart(fig_vendor_pie, use_container_width=False, config=PLOTLY_CONFIG)

# =========================
# 按 Wi‑Fi Technology 分的品牌分佈（表 + 堆疊條）
# =========================
st.subheader("📘 各 Wi‑Fi Technology 的品牌分佈（Huawei / Ruckus）")
tech_vendor_rows = []
for tech in WIFI_LEVELS_DISPLAY:
    sub = combined_df[combined_df["Wifi Technology (norm)"] == tech]
    if sub.empty: continue
    vc = sub["Vendor"].value_counts().reindex(["Huawei","Ruckus"], fill_value=0)
    tech_vendor_rows.append({"Wi‑Fi Technology":tech, "Huawei":int(vc["Huawei"]), "Ruckus":int(vc["Ruckus"]), "Total":int(vc.sum())})
df_tech_vendor = pd.DataFrame(tech_vendor_rows)
if df_tech_vendor.empty:
    st.info("目前六類資料中沒有 Wi‑Fi 4/5/6/7 的 AP。")
else:
    st.dataframe(df_tech_vendor, use_container_width=True)
    df_tv_plot = df_tech_vendor.melt(id_vars=["Wi‑Fi Technology","Total"], value_vars=["Huawei","Ruckus"],
                                     var_name="Vendor", value_name="Count")
    fig_tv = px.bar(df_tv_plot, x="Count", y="Wi‑Fi Technology", color="Vendor",
                    orientation="h", barmode="stack", color_discrete_map=COLOR_MAP,
                    title=None, text="Count")
    fig_tv.update_traces(textposition="inside")
    fig_tv = apply_clean_layout(fig_tv, "AP Vendor Breakdown within each Wi‑Fi Technology", remove_y_title=False, percent_axis=False)
    st.plotly_chart(fig_tv, use_container_width=True, config=PLOTLY_CONFIG)


    st.markdown("") 
    st.markdown("") 
# =========================

    # 1) Managed — 統計表
    st.markdown("### 📑 Managed Wi‑Fi 統計表")

    wifi_series_mgd = count_wifi_tech_series(managed_df)
    df_wifi_mgd = (
        wifi_series_mgd
        .rename_axis("Wi‑Fi Technology")
        .reset_index(name="Count")
    )

    # ✅ 計算百分比
    _wifi_total = int(df_wifi_mgd["Count"].sum())
    if _wifi_total > 0:
        df_wifi_mgd["Percent"] = (
            (df_wifi_mgd["Count"] / _wifi_total * 100)
            .round(0)
            .astype(int)
            .astype(str)
            + "%"
        )
    else:
        df_wifi_mgd["Percent"] = "0%"

    st.markdown("**Wi‑Fi Technology 統計**")
    st.dataframe(
        df_wifi_mgd[["Wi‑Fi Technology", "Count", "Percent"]],
        use_container_width=True
    )

    vendor_series_mgd = (
        managed_df["Vendor"]
        .value_counts()
        .reindex(["Huawei", "Ruckus"], fill_value=0)
    )

    df_vendor_mgd = (
        vendor_series_mgd
        .rename_axis("Vendor")
        .reset_index(name="Count")
    )

    # ✅ 計算百分比
    _vendor_total = int(df_vendor_mgd["Count"].sum())
    if _vendor_total > 0:
        df_vendor_mgd["Percent"] = (
            (df_vendor_mgd["Count"] / _vendor_total * 100)
            .round(0)
            .astype(int)
            .astype(str)
            + "%"
        )
    else:
        df_vendor_mgd["Percent"] = "0%"

    st.markdown("**Vendor 統計**")
    st.dataframe(
        df_vendor_mgd[["Vendor", "Count", "Percent"]],
        use_container_width=True
    )
    
    # 2) Managed — Wi‑Fi Technology 餅圖（數量，整數百分比）
    # === 新增：餅圖用整數百分比（合計=100） + 覆寫文字 ===
    _mgd_pie_df = compute_pie_integer_percent(df_wifi_mgd, value_col="Count", out_col="PercentInt")
    _mgd_wifi_text = [f"{int(c)}，{int(p)}%" for c, p in zip(_mgd_pie_df["Count"], _mgd_pie_df["PercentInt"])]

    fig_mgd_pie = make_square_pie_12cm_managed(
        df=df_wifi_mgd,
        names="Wi‑Fi Technology",
        values="Count",
        title_text="Total number of APs<br>for Managed Wi‑Fi",
        color_col="Wi‑Fi Technology",
        color_discrete_map=COLOR_MAP,
        show_value_and_percent=True,
        text_values=_mgd_wifi_text   # ← 使用預先分配好的整數百分比文字
    )
    center_l, center_c, center_r = st.columns([1, 2, 1])
    with center_c:
        st.plotly_chart(fig_mgd_pie, use_container_width=False, config=PLOTLY_CONFIG)

    # 3) Managed — Vendor 餅圖（僅整數百分比）
    # === 新增：餅圖用整數百分比（合計=100） + 覆寫文字（僅百分比） ===
    _mgd_vendor_pie_df = compute_pie_integer_percent(df_vendor_mgd, value_col="Count", out_col="PercentInt")
    _mgd_vendor_text = [f"{int(p)}%" for p in _mgd_vendor_pie_df["PercentInt"]]

    fig_vendor_mgd = make_square_pie_12cm_managed(
        df=df_vendor_mgd,
        names="Vendor",
        values="Count",
        title_text="Brand Distribution of APs<br>for Managed Wi‑Fi",
        color_col="Vendor",
        color_discrete_map=COLOR_MAP,
        show_value_and_percent=False,
        text_values=_mgd_vendor_text   # ← 使用預先分配好的整數百分比文字
    )
    center_l, center_c, center_r = st.columns([1, 2, 1])
    with center_c:
        st.plotly_chart(fig_vendor_mgd, use_container_width=False, config=PLOTLY_CONFIG)

# =========================
# 月度差異（本月 vs 上月）— 保留原功能
# =========================
if uploaded_prev is not None:
    df_prev_raw, err2 = read_upload(uploaded_prev)
    if err2:
        st.error(f"上月檔案讀取失敗：{err2}")
    else:
        try:
            colmap_prev_min = resolve_columns(df_prev_raw, required_min=("service_type","ssid1","site_code"), optional=())
            with st.expander("🧭 上月欄位對照（模糊識別結果）", expanded=False):
                st.write({
                    "Service Type": colmap_prev_min.get("service_type"),
                    "SSID 1": colmap_prev_min.get("ssid1"),
                    "Site Code": colmap_prev_min.get("site_code"),
                    "Wifi Technology": _best_match_column(list(df_prev_raw.columns), COLUMN_ALIASES.get("wifi_technology", [])),
                    "AP Model": _best_match_column(list(df_prev_raw.columns), COLUMN_ALIASES.get("ap_model", [])),
                    "Hotspot Name (Chinese)": _best_match_column(list(df_prev_raw.columns), COLUMN_ALIASES.get("hotspot_name_cn", [])),
                })
        except Exception as e:
            st.error(f"上月資料缺少必要欄位（最小）：{e}")
            colmap_prev_min = None

        if colmap_prev_min is not None:
            try:
                df_prev, df_prev6 = prepare_df(df_prev_raw, allow_missing_wifi_vendor=True)
            except Exception as e:
                st.error(f"上月資料準備失敗：{e}")
                df_prev6 = None

            if df_prev6 is not None:
                sites_curr = set(df_curr6["Site Code"])
                sites_prev = set(df_prev6["Site Code"])
                added_sites_all = sorted(list(sites_curr - sites_prev))
                removed_sites_all = sorted(list(sites_prev - sites_curr))

                ap_curr_by_site = df_curr6.groupby("Site Code").size()
                ap_prev_by_site = df_prev6.groupby("Site Code").size()

                cat_curr_by_site = site_category_majority(df_curr6)
                cat_prev_by_site = site_category_majority(df_prev6)

                name_curr_by_site = df_curr6.groupby("Site Code")["Hotspot Name (Chinese)"].first().to_dict()
                name_prev_by_site = df_prev6.groupby("Site Code")["Hotspot Name (Chinese)"].first().to_dict()

                common_sites = sites_curr & sites_prev
                changed_sites_all = sorted([s for s in common_sites if int(ap_curr_by_site.get(s,0)) != int(ap_prev_by_site.get(s,0))])
                moved_sites = sorted([s for s in common_sites if cat_curr_by_site.get(s) != cat_prev_by_site.get(s)])

                def delta_rows_for_sites(site_list):
                    rows = []
                    for s in site_list:
                        prev_ap = int(ap_prev_by_site.get(s,0))
                        curr_ap = int(ap_curr_by_site.get(s,0))
                        rows.append({
                            "Site Code": s,
                            "Hotspot Name (Chinese)（上月）": name_prev_by_site.get(s,"-") if s in sites_prev else "-",
                            "Hotspot Name (Chinese)（本月）": name_curr_by_site.get(s,"-") if s in sites_curr else "-",
                            "上月 AP 數（六類）": prev_ap,
                            "本月 AP 數（六類）": curr_ap,
                            "Δ AP": curr_ap - prev_ap,
                            "上月類型": cat_prev_by_site.get(s,"-"),
                            "本月類型": cat_curr_by_site.get(s,"-")
                        })
                    return pd.DataFrame(rows)

                added_sites_df = delta_rows_for_sites(added_sites_all)
                removed_sites_df = delta_rows_for_sites(removed_sites_all)
                changed_sites_df = delta_rows_for_sites(changed_sites_all)

                moved_rows = []
                for s in moved_sites:
                    moved_rows.append({
                        "Site Code": s,
                        "Hotspot Name (Chinese)（上月）": name_prev_by_site.get(s,"-"),
                        "Hotspot Name (Chinese)（本月）": name_curr_by_site.get(s,"-"),
                        "上月類型": cat_prev_by_site.get(s,"-"),
                        "本月類型": cat_curr_by_site.get(s,"-"),
                        "上月 AP 數（六類）": int(ap_prev_by_site.get(s,0)),
                        "本月 AP 數（六類）": int(ap_curr_by_site.get(s,0)),
                        "Δ AP": int(ap_curr_by_site.get(s,0) - ap_prev_by_site.get(s,0))
                    })
                moved_sites_df = pd.DataFrame(moved_rows)

                st.subheader("📦 月度差異（本月 vs 上月，僅六類）")
                c1,c2,c3,c4,c5,c6 = st.columns(6)
                c1.metric("本月 Site（六類）", len(sites_curr))
                c2.metric("上月 Site（六類）", len(sites_prev))
                c3.metric("新增站點（六類）", len(added_sites_all))
                c4.metric("移除站點（六類）", len(removed_sites_all))
                c5.metric("AP 變動站點（六類）", len(changed_sites_all))
                c6.metric("類型變更（六類）", len(moved_sites))

                st.subheader("📂 各分類 Site 差異（六類；新增 / 移除 / AP 變動；以本月類別分組）")

                def attach_delta(df_sites):
                    if df_sites.empty: return df_sites
                    df_sites = df_sites.copy()
                    df_sites["上月 AP 數（六類）"] = df_sites["Site Code"].map(lambda s: int(ap_prev_by_site.get(s,0)))
                    df_sites["本月 AP 數（六類）"] = df_sites["Site Code"].map(lambda s: int(ap_curr_by_site.get(s,0)))
                    df_sites["Δ AP"] = df_sites["本月 AP 數（六類）"] - df_sites["上月 AP 數（六類）"]
                    df_sites["上月類型"] = df_sites["Site Code"].map(lambda s: cat_prev_by_site.get(s,"-"))
                    df_sites["本月類型"] = df_sites["Site Code"].map(lambda s: cat_curr_by_site.get(s,"-"))
                    df_sites["Hotspot Name (Chinese)（上月）"] = df_sites["Site Code"].map(lambda s: name_prev_by_site.get(s,"-"))
                    df_sites["Hotspot Name (Chinese)（本月）"] = df_sites["Site Code"].map(lambda s: name_curr_by_site.get(s,"-"))
                    return df_sites

                percat_added, percat_removed, percat_changed = {}, {}, {}
                for cat in CATEGORY_ORDER:
                    sites_curr_cat = {s for s in sites_curr if cat_curr_by_site.get(s)==cat}
                    sites_prev_cat = {s for s in sites_prev if cat_prev_by_site.get(s)==cat}
                    added_cat = sorted(list(sites_curr_cat - sites_prev))
                    removed_cat = sorted(list(sites_prev_cat - sites_curr))
                    changed_cat = sorted([s for s in (sites_curr_cat & sites_prev) if int(ap_curr_by_site.get(s,0)) != int(ap_prev_by_site.get(s,0))])
                    percat_added[cat], percat_removed[cat], percat_changed[cat] = added_cat, removed_cat, changed_cat

                    with st.expander(f"{cat}：新增 {len(added_cat)} / 移除 {len(removed_cat)} / AP變動 {len(changed_cat)}（六類）", expanded=False):
                        A,B,C = st.columns(3)
                        with A:
                            st.write("**新增站點（含 ΔAP）**")
                            st.dataframe(attach_delta(pd.DataFrame({"Site Code": added_cat})), use_container_width=True)
                        with B:
                            st.write("**移除站點（含 ΔAP）**")
                            st.dataframe(attach_delta(pd.DataFrame({"Site Code": removed_cat})), use_container_width=True)
                        with C:
                            st.write("**AP 變動站點（含 ΔP）**")
                            st.dataframe(attach_delta(pd.DataFrame({"Site Code": changed_cat})), use_container_width=True)

                st.markdown("### ➕ 新增站點（六類；本月有、上月無）")
                st.dataframe(added_sites_df, use_container_width=True)
                st.markdown(f"- 新增站點 **AP 總增量（六類）**：{int(added_sites_df['本月 AP 數（六類）'].sum())}")

                st.markdown("### ➖ 移除站點（六類；上月有、本月無）")
                st.dataframe(removed_sites_df, use_container_width=True)
                st.markdown(f"- 移除站點 **AP 總減量（六類）**：{int(removed_sites_df['上月 AP 數（六類）'].sum())}")

                st.markdown("### 🔁 AP 變動站點（六類；兩月皆存在，但 AP 數不同）")
                st.dataframe(changed_sites_df, use_container_width=True)
                if not changed_sites_df.empty:
                    st.markdown(f"- **ΔAP 合計（六類）**：{int(changed_sites_df['Δ AP'].sum())}（同一 Site 的 AP 增減合計）")

                if not moved_sites_df.empty:
                    st.markdown("### 🔄 類型變更站點（六類）")
                    st.dataframe(moved_sites_df, use_container_width=True)

                st.subheader("⬇️ 差異結果下載（只含六類）")
                excel_bio = io.BytesIO()
                with pd.ExcelWriter(excel_bio, engine="openpyxl") as writer:
                    summary_df.to_excel(writer, index=False, sheet_name="本月_各分類彙總_六類")
                    total_wifi_df.to_excel(writer, index=False, sheet_name="本月_total_wifi_六類")
                    total_vendor_df.to_excel(writer, index=False, sheet_name="本月_total_vendor_六類")
                    for name in CATEGORY_ORDER:
                        per_category_wifi_tables[name].to_excel(writer, index=False, sheet_name=f"本月_{name[:24]}_wifi")
                        per_category_vendor_tables[name].to_excel(writer, index=False, sheet_name=f"本月_{name[:24]}_vendor")
                    added_sites_df.to_excel(writer, index=False, sheet_name="差異_新增站點_六類")
                    removed_sites_df.to_excel(writer, index=False, sheet_name="差異_移除站點_六類")
                    changed_sites_df.to_excel(writer, index=False, sheet_name="差異_AP變動_六類")
                    if not moved_sites_df.empty:
                        moved_sites_df.to_excel(writer, index=False, sheet_name="差異_類型變更站點_六類")
                    safe_cols = [c for c in ["Site Code","Hotspot Name (Chinese)","Wifi Technology (norm)","Vendor","Category"] if c in df_curr6.columns]
                    df_curr6[safe_cols].to_excel(writer, index=False, sheet_name="本月_明細_六類")
                    df_prev6[safe_cols].to_excel(writer, index=False, sheet_name="上月_明細_六類")
                excel_bio.seek(0)
                st.download_button(
                    "下載【月度差異 + 本月彙總】Excel（只含六類）",
                    excel_bio, "wifi_ap_monthly_diff_6cats.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# =========================
# 單檔輸出：本月彙總（六類）
# =========================
st.subheader("⬇️ 本月彙總下載（六類）")
excel_bio_single = io.BytesIO()
with pd.ExcelWriter(excel_bio_single, engine="openpyxl") as writer:
    summary_df.to_excel(writer, index=False, sheet_name="summary_六類")
    total_wifi_df.to_excel(writer, index=False, sheet_name="total_wifi_六類")
    total_vendor_df.to_excel(writer, index=False, sheet_name="total_vendor_六類")
excel_bio_single.seek(0)
st.download_button(
    "下載本月彙總 Excel（六類）",
    excel_bio_single, "wifi_ap_current_summary_6cats.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("完成！")
