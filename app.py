"""
SharkNinja 데일리 리포트 — Summary / 상품별 / 매체별
ROAS = 매출 ÷ 비용 × 100 (%)
"""
from __future__ import annotations

import json
import os
import re
import tempfile
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from drive_report import download_drive_file, list_drive_reports
from madup_api import (
    DEFAULT_BASE as MADUP_API_DEFAULT_BASE,
    download_madup_file,
    find_latest_madup_path,
    madup_report_filenames,
)

DEFAULT_REPORT_DIR = Path(
    r"C:\Users\MADUP\주식회사매드업 Dropbox\광고사업부\4. 광고주\샤크닌자\07. 리포트"
)
REPORT_FILE_RE = re.compile(
    r"^Madup_Sharkninja_Daily[- ]Report_(\d{6})\.xlsx$", re.IGNORECASE
)

# ── 열 이름 ──────────────────────────────────────────────────────────────────
COL_IMP      = "노출"
COL_CLICK    = "클릭"
COL_COST     = "비용(markup+)"
COL_REV_SS   = "(스마트스토어센터기준) 전체 매출"
COL_REV_DB   = "(대시보드기준) 전체 매출"
COL_BUY_SS   = "(스마트스토어센터기준) 전체 구매"
COL_BUY_DB   = "(대시보드기준) 전체 구매"
COL_COUP_REV = "쿠팡_매출"
COL_GM_REV   = "g마켓_매출"
COL_PRODUCT  = "실 전환 발생 상품"
COL_PROMO    = "프로모션명"

# ── KPI 목표 ─────────────────────────────────────────────────────────────────
KPI = {
    "네이버 브랜드스토어": 2_143_260_060,
    "쿠팡":               835_011_825,
    "G마켓":              58_287_235,
}
HIER_COLS = ["매체상세", "캠페인", "그룹", "소재/키워드", COL_PRODUCT]


# ── 파일 유틸 ─────────────────────────────────────────────────────────────────
def find_latest_report(folder: Path) -> Path | None:
    best: tuple[str, Path] | None = None
    if not folder.is_dir():
        return None
    for p in folder.iterdir():
        if not p.is_file():
            continue
        m = REPORT_FILE_RE.match(p.name)
        if m:
            tag = m.group(1)
            if best is None or tag > best[0]:
                best = (tag, p)
    return best[1] if best else None


@st.cache_data(ttl=120, show_spinner=False)
def load_raw(report_path: str) -> pd.DataFrame:
    df = pd.read_excel(Path(report_path), sheet_name="raw", engine="openpyxl")
    df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce")
    return df


def _drive_secrets_ok() -> bool:
    try:
        fid = st.secrets["GOOGLE_DRIVE_FOLDER_ID"]
        js = st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"]
        if not str(fid).strip():
            return False
        if isinstance(js, dict):
            return bool(js)
        return bool(str(js).strip())
    except Exception:
        return False


@st.cache_data(ttl=120, show_spinner=False)
def _drive_cached_bytes(file_id: str, sa_json: str) -> bytes:
    return download_drive_file(file_id, sa_json)


def _server_os_cannot_use_embedded_windows_path() -> bool:
    """Windows가 아닌 OS(Streamlit Cloud·Linux·macOS 등)에서는 C:\\ 기본 경로를 열 수 없음.

    로컬 Windows(`os.name == "nt"`)만 예외. WSL은 보통 /mnt/c 로 `Path(...).is_dir()` 가 True가 되어 여기서 막히지 않음.
    """
    if os.name == "nt":
        return False
    s = str(DEFAULT_REPORT_DIR)
    if not (len(s) >= 2 and s[1] == ":" and s[0].isalpha()):
        return False
    return not Path(DEFAULT_REPORT_DIR).is_dir()


def _secret_has(name: str) -> bool:
    try:
        v = st.secrets.get(name)
        if v is None:
            return False
        return bool(str(v).strip())
    except Exception:
        return False


def _madup_secrets_ok() -> bool:
    try:
        key = str(st.secrets.get("MADUP_API_KEY", "") or "").strip()
        if not key:
            return False
        fd = str(st.secrets.get("MADUP_DROPBOX_FOLDER", "") or "").strip()
        sp = str(
            st.secrets.get("MADUP_DROPBOX_PATH", "")
            or st.secrets.get("MADUP_DROPBOX_FILE_PATH", "")
            or ""
        ).strip()
        return bool(fd or sp)
    except Exception:
        return False


@st.cache_data(ttl=120, show_spinner=False)
def _madup_bytes_cached(api_key: str, path: str, base: str) -> bytes:
    return download_madup_file(api_key, path, base)


@st.cache_data(ttl=120, show_spinner=False)
def _madup_latest_pair_cached(api_key: str, folder: str, base: str) -> tuple[str, bytes]:
    return find_latest_madup_path(api_key, folder, base)


# ── 포매터 ───────────────────────────────────────────────────────────────────
def fm(v: float) -> str:
    return f"{int(round(v)):,}"

def fp(v: float, d: int = 1) -> str:
    return f"{v:.{d}f}%"


# ── 기간 집계 (SS + DB 모두) ──────────────────────────────────────────────────
def agg_range(raw: pd.DataFrame, start: pd.Timestamp, end: pd.Timestamp) -> dict:
    s = pd.Timestamp(start).normalize()
    e = pd.Timestamp(end).normalize()
    m = (raw["날짜"].dt.normalize() >= s) & (raw["날짜"].dt.normalize() <= e)
    sub = raw.loc[m]
    노출  = float(sub[COL_IMP].sum())   if COL_IMP   in sub.columns else 0.0
    클릭  = float(sub[COL_CLICK].sum()) if COL_CLICK  in sub.columns else 0.0
    cost  = float(sub[COL_COST].sum())
    ss_rev = float(sub[COL_REV_SS].sum())
    db_rev = float(sub[COL_REV_DB].sum())
    ss_buy = float(sub[COL_BUY_SS].sum()) if COL_BUY_SS in sub.columns else 0.0
    db_buy = float(sub[COL_BUY_DB].sum()) if COL_BUY_DB in sub.columns else 0.0
    cpc        = cost / 클릭  if 클릭  else 0.0
    ctr_pct    = 클릭 / 노출 * 100 if 노출 else 0.0
    roas_ss    = ss_rev / cost * 100 if cost else 0.0
    roas_db    = db_rev / cost * 100 if cost else 0.0
    return {
        "노출": 노출, "클릭": 클릭, "CPC": cpc, "CTR(%)": ctr_pct,
        "비용": cost,
        "SS_구매": ss_buy, "SS_매출": ss_rev, "ROAS_SS": roas_ss,
        "DB_구매": db_buy, "DB_매출": db_rev, "ROAS_DB": roas_db,
    }


def agg_monthly(raw: pd.DataFrame, d_max: pd.Timestamp) -> dict:
    first = pd.Timestamp(d_max.year, d_max.month, 1)
    return agg_range(raw, first, d_max)


# ── KPI 달성률 ────────────────────────────────────────────────────────────────
def kpi_achievement(raw: pd.DataFrame, d_max: pd.Timestamp) -> dict[str, tuple[float, int]]:
    first = pd.Timestamp(d_max.year, d_max.month, 1)
    mask  = (raw["날짜"].dt.normalize() >= first) & (raw["날짜"].dt.normalize() <= pd.Timestamp(d_max).normalize())
    sub   = raw.loc[mask]
    naver_mask = sub["채널"].astype(str).str.contains("네이버", na=False, case=False)
    naver = float(sub.loc[naver_mask, COL_REV_SS].sum())
    coup  = float(sub[COL_COUP_REV].sum()) if COL_COUP_REV in sub.columns else 0.0
    gm    = float(sub[COL_GM_REV].sum())   if COL_GM_REV   in sub.columns else 0.0
    return {
        "네이버 브랜드스토어": (naver, KPI["네이버 브랜드스토어"]),
        "쿠팡":               (coup,  KPI["쿠팡"]),
        "G마켓":              (gm,    KPI["G마켓"]),
    }


def kpi_bar_html(name: str, achieved: float, target: int) -> str:
    rate = min(achieved / target, 1.0) if target else 0.0
    pct  = rate * 100
    if pct >= 80:
        bar_color, bg_light = "#27AE60", "#D5F5E3"
    elif pct >= 50:
        bar_color, bg_light = "#F39C12", "#FEF9E7"
    else:
        bar_color, bg_light = "#E74C3C", "#FDEDEC"
    return f"""
<div style="background:{bg_light};border-radius:14px;padding:18px 20px;
            box-shadow:0 3px 10px rgba(0,0,0,0.08);margin-bottom:4px;">
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
    <span style="font-weight:700;font-size:15px;color:#2C3E50;">{name}</span>
    <span style="font-size:24px;font-weight:900;color:{bar_color};">{pct:.1f}%</span>
  </div>
  <div style="background:#ECF0F1;border-radius:10px;height:22px;
              overflow:hidden;box-shadow:inset 0 2px 4px rgba(0,0,0,0.12);">
    <div style="width:{pct:.1f}%;height:100%;
                background:linear-gradient(90deg,{bar_color}aa,{bar_color});
                border-radius:10px;box-shadow:0 2px 6px {bar_color}55;
                display:flex;align-items:center;justify-content:flex-end;padding-right:8px;">
      <span style="color:white;font-weight:700;font-size:12px;">{pct:.1f}%</span>
    </div>
  </div>
  <div style="display:flex;justify-content:space-between;margin-top:10px;font-size:13px;color:#555;">
    <span>달성 <strong style="color:{bar_color};font-size:14px;">{fm(achieved)}원</strong></span>
    <span style="color:#888;">목표 {fm(target)}원</span>
  </div>
</div>"""


# ── ROAS/비용 이중축 차트 (입체적·색상 리치) ──────────────────────────────────
def make_roas_cost_chart(raw: pd.DataFrame, last_day: pd.Timestamp, n: int = 7) -> go.Figure:
    end  = pd.Timestamp(last_day).normalize()
    ds   = sorted(raw["날짜"].dt.normalize().dropna().unique())
    use  = [d for d in ds if d <= end][-n:]
    labels, roas_vals, cost_vals = [], [], []
    for d in use:
        a = agg_range(raw, d, d)
        labels.append(pd.Timestamp(d).strftime("%m.%d"))
        roas_vals.append(round(a["ROAS_SS"], 1))
        cost_vals.append(round(a["비용"] / 10_000, 1))

    avg_roas = sum(roas_vals) / len(roas_vals) if roas_vals else 0

    fig = go.Figure()

    fig.add_hline(
        y=avg_roas, line_dash="dash", line_color="#C8C8C8", line_width=1.2,
        annotation_text=f"평균 {avg_roas:,.0f}%",
        annotation_position="top left",
        annotation_font=dict(size=10, color="#999"),
    )

    fig.add_trace(go.Bar(
        x=labels, y=roas_vals, name="ROAS (%)",
        marker=dict(
            color=roas_vals,
            colorscale=[[0, "rgba(200,180,240,0.5)"], [0.5, "rgba(150,110,210,0.65)"], [1, "rgba(110,60,180,0.8)"]],
            line=dict(color="rgba(100,60,170,0.5)", width=1),
        ),
        text=[f"{v:,.1f}%" for v in roas_vals],
        textposition="inside",
        textangle=0,
        textfont=dict(size=11, color="#5B3A8C", family="Arial"),
        yaxis="y1",
    ))

    fig.add_trace(go.Scatter(
        x=labels, y=cost_vals, name="비용 (만원)",
        mode="lines+markers",
        line=dict(color="#6BBF6B", width=2.5, shape="spline"),
        marker=dict(size=8, color="white",
                    line=dict(color="#6BBF6B", width=2),
                    symbol="circle"),
        yaxis="y2",
    ))

    roas_max = max(roas_vals) if roas_vals else 1000
    cost_max = max(cost_vals) if cost_vals else 100

    fig.update_layout(
        yaxis=dict(
            title=None,
            showticklabels=False,
            showgrid=True, gridcolor="rgba(200,200,200,0.3)", gridwidth=1,
            zeroline=False,
            range=[0, roas_max * 1.3],
        ),
        yaxis2=dict(
            title=None,
            showticklabels=False,
            overlaying="y", side="right", showgrid=False, zeroline=False,
            range=[0, cost_max * 1.6],
        ),
        xaxis=dict(
            tickfont=dict(size=12, color="#555", family="Arial"),
            showgrid=False,
            showline=True, linecolor="#ddd", linewidth=1,
        ),
        legend=dict(
            orientation="h", y=1.12, x=0.5, xanchor="center",
            bgcolor="rgba(255,255,255,0)", borderwidth=0,
            font=dict(size=11, color="#666"),
        ),
        height=350,
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(l=10, r=10, t=50, b=30),
        bargap=0.25,
        hoverlabel=dict(bgcolor="white", font_size=12, bordercolor="#ccc"),
    )
    return fig


# ── 다중 비교 기간 생성 ───────────────────────────────────────────────────────
def multi_period_compare(
    raw: pd.DataFrame, sel_s: pd.Timestamp, sel_e: pd.Timestamp
) -> list[tuple[str, dict]]:
    """선택 기간 기준 전일/동요일/2일전/3일전 비교 딕셔너리 리스트."""
    n = int((sel_e - sel_s).days) + 1
    comparisons: list[tuple[str, dict]] = []
    # 전일 (직전 동일 기간)
    p1_e = sel_s - timedelta(days=1)
    p1_s = p1_e  - timedelta(days=n - 1)
    comparisons.append(("전일 대비", agg_range(raw, p1_s, p1_e)))
    # 전주 동요일 (7일 전)
    comparisons.append(("동요일 대비", agg_range(raw, sel_s - timedelta(days=7), sel_e - timedelta(days=7))))
    # 2일 전
    p2_e = sel_s - timedelta(days=2)
    p2_s = p2_e  - timedelta(days=n - 1)
    comparisons.append(("2일 전 대비", agg_range(raw, p2_s, p2_e)))
    # 3일 전
    p3_e = sel_s - timedelta(days=3)
    p3_s = p3_e  - timedelta(days=n - 1)
    comparisons.append(("3일 전 대비", agg_range(raw, p3_s, p3_e)))
    return comparisons


def build_period_table_html(cur: dict, comparisons: list[tuple[str, dict]], basis: str) -> str:
    """기간별 성과 비교 표 HTML (SS 또는 DB 기준)."""
    if basis == "SS":
        metrics = [
            ("비용 (원)",  "비용",     False),
            ("구매 (건)",  "SS_구매",  False),
            ("매출 (원)",  "SS_매출",  False),
            ("ROAS (%)",   "ROAS_SS",  True),
        ]
    else:
        metrics = [
            ("비용 (원)",  "비용",     False),
            ("구매 (건)",  "DB_구매",  False),
            ("매출 (원)",  "DB_매출",  False),
            ("ROAS (%)",   "ROAS_DB",  True),
        ]

    def _delta_badge(d: float, is_pct: bool) -> str:
        if is_pct:
            txt = f"{d:+.1f}p"
        else:
            txt = f"{d:+,.0f}"
        if d > 0:
            color = "#27AE60"
            arrow = "▲"
        elif d < 0:
            color = "#E74C3C"
            arrow = "▼"
        else:
            color = "#888"
            arrow = "−"
        return (
            f'<span style="color:{color};font-weight:600;font-size:12px;">'
            f'{arrow} {txt}</span>'
        )

    # 헤더
    cols = ["지표", "선택 기간"] + [c[0] for c in comparisons]
    hdr = "".join(f'<th style="padding:8px 12px;background:#F8F9FA;font-size:13px;'
                  f'border-bottom:2px solid #DEE2E6;text-align:center;">{c}</th>' for c in cols)

    rows_html = ""
    for label, key, is_pct in metrics:
        v = cur[key]
        val_str = fp(v) if is_pct else fm(v)
        cells = f'<td style="padding:8px 12px;font-weight:600;font-size:13px;border-bottom:1px solid #EEE;">{label}</td>'
        cells += f'<td style="padding:8px 12px;text-align:right;font-weight:700;font-size:14px;border-bottom:1px solid #EEE;">{val_str}</td>'
        for _, comp in comparisons:
            delta = v - comp[key]
            cells += (
                f'<td style="padding:8px 12px;text-align:right;border-bottom:1px solid #EEE;">'
                f'{_delta_badge(delta, is_pct)}</td>'
            )
        rows_html += f"<tr>{cells}</tr>"

    return f"""
<div style="border:1px solid #DEE2E6;border-radius:12px;overflow:hidden;
            box-shadow:0 2px 8px rgba(0,0,0,0.06);margin-bottom:8px;">
<table style="width:100%;border-collapse:collapse;">
<thead><tr>{hdr}</tr></thead>
<tbody>{rows_html}</tbody>
</table>
</div>"""


# ── 전일 대비 그룹 비교 ───────────────────────────────────────────────────────
def two_day_compare(
    raw: pd.DataFrame, group_cols: list[str],
    d_last: pd.Timestamp, d_prev: pd.Timestamp,
) -> pd.DataFrame:
    r = raw.copy()
    r["d"] = r["날짜"].dt.normalize()
    dl = pd.Timestamp(d_last).normalize()
    dp = pd.Timestamp(d_prev).normalize()
    for c in group_cols:
        if c in r.columns:
            r[c] = r[c].fillna("(미입력)").astype(str)
    agg_cols = [c for c in [COL_REV_DB, COL_COST, COL_BUY_DB] if c in r.columns]

    def one(d: pd.Timestamp) -> pd.DataFrame:
        return r[r["d"] == d].groupby(group_cols, dropna=False)[agg_cols].sum()

    try:
        a = one(dl).add_suffix("_전일")
        b = one(dp).add_suffix("_전전일")
    except Exception:
        return pd.DataFrame()
    m = a.join(b, how="outer").fillna(0)
    if f"{COL_REV_DB}_전일" not in m.columns:
        return pd.DataFrame()
    m["매출_증감"] = m[f"{COL_REV_DB}_전일"] - m[f"{COL_REV_DB}_전전일"]
    m["비용_증감"] = m[f"{COL_COST}_전일"]   - m[f"{COL_COST}_전전일"]
    bl = f"{COL_BUY_DB}_전일"; bp = f"{COL_BUY_DB}_전전일"
    m["구매_증감"] = m[bl] - m[bp] if bl in m.columns and bp in m.columns else 0
    def _ro(rev: float, cost: float) -> float:
        return rev / cost * 100 if cost else 0.0
    m["ROAS%_전일"]   = m.apply(lambda r: _ro(r[f"{COL_REV_DB}_전일"],   r[f"{COL_COST}_전일"]),   axis=1)
    m["ROAS%_전전일"] = m.apply(lambda r: _ro(r[f"{COL_REV_DB}_전전일"], r[f"{COL_COST}_전전일"]), axis=1)
    m["ROAS%p_차이"]  = m["ROAS%_전일"] - m["ROAS%_전전일"]
    m = m.reset_index()
    m["영향도"] = m["매출_증감"].abs()
    return m.sort_values("영향도", ascending=False)


def two_day_compare_ss(
    raw: pd.DataFrame, group_cols: list[str],
    d_last: pd.Timestamp, d_prev: pd.Timestamp,
) -> pd.DataFrame:
    """SS(스마트스토어) 기준 매출/구매 + 비용/ROAS 비교"""
    r = raw.copy()
    r["d"] = r["날짜"].dt.normalize()
    dl = pd.Timestamp(d_last).normalize()
    dp = pd.Timestamp(d_prev).normalize()
    for c in group_cols:
        if c in r.columns:
            r[c] = r[c].fillna("(미입력)").astype(str)
    agg_cols = [c for c in [COL_REV_SS, COL_BUY_SS, COL_COST] if c in r.columns]

    def one(d: pd.Timestamp) -> pd.DataFrame:
        return r[r["d"] == d].groupby(group_cols, dropna=False)[agg_cols].sum()

    try:
        a = one(dl).add_suffix("_전일")
        b = one(dp).add_suffix("_전전일")
    except Exception:
        return pd.DataFrame()
    m = a.join(b, how="outer").fillna(0)
    rev_col = f"{COL_REV_SS}_전일"
    if rev_col not in m.columns:
        return pd.DataFrame()
    m["매출_증감"] = m[f"{COL_REV_SS}_전일"] - m[f"{COL_REV_SS}_전전일"]
    bl = f"{COL_BUY_SS}_전일"; bp = f"{COL_BUY_SS}_전전일"
    m["구매_증감"] = m[bl] - m[bp] if bl in m.columns and bp in m.columns else 0
    cl = f"{COL_COST}_전일"; cp = f"{COL_COST}_전전일"
    m["비용_증감"] = m[cl] - m[cp] if cl in m.columns and cp in m.columns else 0
    def _ro(rev: float, cost: float) -> float:
        return rev / cost * 100 if cost else 0.0
    m["ROAS%_전일"]   = m.apply(lambda r: _ro(r[f"{COL_REV_SS}_전일"],   r.get(cl, 0)), axis=1)
    m["ROAS%_전전일"] = m.apply(lambda r: _ro(r[f"{COL_REV_SS}_전전일"], r.get(cp, 0)), axis=1)
    m["ROAS%p_차이"]  = m["ROAS%_전일"] - m["ROAS%_전전일"]
    m = m.reset_index()
    m["영향도"] = m["매출_증감"].abs()
    return m.sort_values("영향도", ascending=False)


# ── 뎁스 코멘트 (매체상세별 드릴다운 서술형) ──────────────────────────────────
def summary_depth_comment(raw: pd.DataFrame, d_last: pd.Timestamp, d_prev: pd.Timestamp) -> str:
    if d_prev is None:
        return "_전전일 데이터가 없어 비교 불가합니다._"

    cur  = agg_range(raw, d_last, d_last)
    prev = agg_range(raw, d_prev, d_prev)
    swd_ts = pd.Timestamp(d_last) - timedelta(days=7)
    swd    = agg_range(raw, swd_ts, swd_ts)
    last_s = pd.Timestamp(d_last).strftime("%m/%d")
    prev_s = pd.Timestamp(d_prev).strftime("%m/%d")
    swd_s  = swd_ts.strftime("%m/%d")

    lines: list[str] = []

    # ── 1) 전일 성과 요약 ──
    lines.append(f"#### 📌 전일 성과  ({last_s})")
    lines.append(f"| 지표 | {last_s} | 전일 대비 ({prev_s}) | 동요일 대비 ({swd_s}) |")
    lines.append("|:--|--:|--:|--:|")
    for label, key, is_pct in [
        ("구매 (SS)", "SS_구매", False), ("구매 (DB)", "DB_구매", False),
        ("매출 (SS)", "SS_매출", False), ("매출 (DB)", "DB_매출", False),
        ("ROAS (SS)", "ROAS_SS", True),  ("ROAS (DB)", "ROAS_DB", True),
    ]:
        v = cur[key]; dv = v - prev[key]; sv = v - swd[key]
        if is_pct:
            lines.append(f"| {label} | {v:.1f}% | {dv:+.1f}p | {sv:+.1f}p |")
        else:
            lines.append(f"| {label} | {v:,.0f} | {dv:+,.0f} | {sv:+,.0f} |")

    # ── 2) 전체 방향 판단 ──
    rev_ss_d = cur["SS_매출"] - prev["SS_매출"]
    dir_ss = "증가" if rev_ss_d >= 0 else "감소"
    emoji_ss = "📈" if rev_ss_d >= 0 else "📉"
    lines.append(
        f"\n> {emoji_ss} SS 기준 매출: 전전일 대비 **{rev_ss_d:+,.0f}원** ({dir_ss})  \n"
        f"> ROAS(SS): **{cur['ROAS_SS']:.1f}%** ({cur['ROAS_SS']-prev['ROAS_SS']:+.1f}p)"
    )

    # ── 3) 매체상세별 데일리 코멘트 (SS 기준) ──
    m_media_ss = two_day_compare_ss(raw, ["매체상세"], d_last, d_prev)
    if m_media_ss.empty:
        return "\n".join(lines)

    top_medias = m_media_ss.head(3)["매체상세"].tolist()

    r_copy = raw.copy()
    r_copy["d"] = r_copy["날짜"].dt.normalize()

    is_coupang = lambda name: "쿠팡" in str(name)

    lines.append(f"\n---\n#### 📋 데일리 코멘트  ({last_s} vs {prev_s})")
    lines.append(f"매출 영향도 큰 매체상세 (SS 기준): **{', '.join(top_medias)}**\n")

    for rank, media in enumerate(top_medias, 1):
        mr = m_media_ss[m_media_ss["매체상세"] == media].iloc[0]
        delta_dir = "증가" if mr["매출_증감"] >= 0 else "감소"
        ss_rev = mr[f"{COL_REV_SS}_전일"]
        ss_roas = mr["ROAS%_전일"]
        ss_roas_d = mr["ROAS%p_차이"]
        cost_val = mr.get(f"{COL_COST}_전일", 0)
        cost_d = mr.get("비용_증감", 0)

        lines.append(
            f"##### {rank}. {media}  ({delta_dir})\n"
            f"**[SS 기준]** 매출 **{ss_rev:,.0f}원** ({mr['매출_증감']:+,.0f}원) · "
            f"비용 {cost_val:,.0f}원 ({cost_d:+,.0f}원) · "
            f"ROAS {ss_roas:.1f}% ({ss_roas_d:+.1f}p)"
        )

        sub = r_copy[r_copy["매체상세"].astype(str) == str(media)]

        # 실전환 광고 상품 (쿠팡 제외, SS 기준)
        if not is_coupang(media):
            prod_sub = sub[
                sub[COL_PRODUCT].notna() &
                (sub[COL_PRODUCT].astype(str).str.strip() != "") &
                (sub[COL_PRODUCT].astype(str).str.strip() != "-") &
                (sub[COL_PRODUCT].astype(str) != "(미입력)")
            ]
            if not prod_sub.empty:
                m_prod = two_day_compare_ss(prod_sub, [COL_PRODUCT], d_last, d_prev)
                if not m_prod.empty:
                    top_pr = m_prod.head(3)
                    lines.append("\n**실전환 광고 상품 (SS 기준):**")
                    for _, rr in top_pr.iterrows():
                        pname = str(rr.get(COL_PRODUCT, ""))
                        if len(pname) > 50:
                            pname = pname[:50] + "…"
                        pr_dir = "▲" if rr["매출_증감"] >= 0 else "▼"
                        lines.append(
                            f"- {pr_dir} {pname} : 매출 **{rr['매출_증감']:+,.0f}원**, "
                            f"구매 {rr.get('구매_증감', 0):+,.0f}건"
                        )

        # 캠페인/그룹 (DB 기준, 참고)
        m_cg = two_day_compare(sub, ["캠페인", "그룹"], d_last, d_prev)
        if not m_cg.empty:
            top_cg = m_cg.head(3)
            lines.append("\n**참고 – 캠페인/그룹 영향 (DB 기준):**")
            for _, rr in top_cg.iterrows():
                camp = str(rr.get("캠페인", ""))
                grp  = str(rr.get("그룹", ""))
                label_parts = [p for p in [camp, grp] if p and p not in ("(미입력)", "nan")]
                lbl = " > ".join(label_parts) if label_parts else "(기타)"
                cg_dir = "▲" if rr["매출_증감"] >= 0 else "▼"
                lines.append(
                    f"- {cg_dir} {lbl} : 매출 **{rr['매출_증감']:+,.0f}원**, "
                    f"ROAS {rr['ROAS%_전일']:.1f}% ({rr['ROAS%p_차이']:+.1f}p)"
                )

        lines.append("")

    # ── 4) SUMMARY 표 (SS 기준 매출/ROAS + DB 기준 캠페인 참고) ──
    lines.append(f"---\n#### 📊 SUMMARY  ({last_s} vs {prev_s} 대비)")

    m_media_db = two_day_compare(raw, ["매체상세"], d_last, d_prev)

    summary_html = (
        '<div style="overflow-x:auto;margin-bottom:12px;">'
        '<table style="width:100%;border-collapse:collapse;font-size:13px;">'
        '<thead><tr style="background:#f8f9fa;border-bottom:2px solid #dee2e6;">'
        f'<th style="padding:8px 10px;text-align:left;">매체상세</th>'
        f'<th style="padding:8px 10px;text-align:right;">SS 매출<br><span style="font-size:11px;color:#888;">{last_s}</span></th>'
        f'<th style="padding:8px 10px;text-align:right;">SS 매출 증감<br><span style="font-size:11px;color:#888;">{last_s} vs {prev_s}</span></th>'
        f'<th style="padding:8px 10px;text-align:right;">비용 증감<br><span style="font-size:11px;color:#888;">{last_s} vs {prev_s}</span></th>'
        f'<th style="padding:8px 10px;text-align:right;">SS ROAS<br><span style="font-size:11px;color:#888;">{last_s}</span></th>'
        f'<th style="padding:8px 10px;text-align:right;">SS ROAS 증감<br><span style="font-size:11px;color:#888;">{last_s} vs {prev_s}</span></th>'
        f'<th style="padding:8px 10px;text-align:left;min-width:220px;">참고: 핵심 캠페인 (DB 기준)</th>'
        '</tr></thead><tbody>'
    )

    for i, (_, row) in enumerate(m_media_ss.iterrows()):
        media = row["매체상세"]
        rev   = row[f"{COL_REV_SS}_전일"]
        rd    = row["매출_증감"]
        cd    = row.get("비용_증감", 0)
        roas  = row["ROAS%_전일"]
        rp    = row["ROAS%p_차이"]

        sub = r_copy[r_copy["매체상세"].astype(str) == str(media)]
        m_cg = two_day_compare(sub, ["캠페인", "그룹"], d_last, d_prev)
        factor = "-"
        if not m_cg.empty:
            top1 = m_cg.iloc[0]
            camp = str(top1.get("캠페인", ""))
            grp  = str(top1.get("그룹", ""))
            parts = [p for p in [camp, grp] if p and p not in ("(미입력)", "nan")]
            factor = " > ".join(parts) if parts else "-"

        rd_color = "#d63031" if rd < 0 else "#00b894"
        cd_color = "#d63031" if cd < 0 else "#00b894"
        rp_color = "#d63031" if rp < 0 else "#00b894"
        bg = "#fff" if i % 2 == 0 else "#f9f9fb"

        summary_html += (
            f'<tr style="background:{bg};border-bottom:1px solid #eee;">'
            f'<td style="padding:7px 10px;font-weight:600;">{media}</td>'
            f'<td style="padding:7px 10px;text-align:right;">{rev:,.0f}</td>'
            f'<td style="padding:7px 10px;text-align:right;color:{rd_color};font-weight:600;">{rd:+,.0f}</td>'
            f'<td style="padding:7px 10px;text-align:right;color:{cd_color};">{cd:+,.0f}</td>'
            f'<td style="padding:7px 10px;text-align:right;">{roas:.1f}%</td>'
            f'<td style="padding:7px 10px;text-align:right;color:{rp_color};font-weight:600;">{rp:+.1f}p</td>'
            f'<td style="padding:7px 10px;font-size:12px;word-break:break-all;">{factor}</td>'
            '</tr>'
        )

    summary_html += '</tbody></table></div>'

    return "\n".join(lines) + "\n" + summary_html


# ── 상시/프로모션 ─────────────────────────────────────────────────────────────
def classify_promo(name: object) -> str:
    s = str(name).strip() if pd.notna(name) else ""
    return "상시" if "상시" in s else ("프로모션" if s else "미분류")


def promo_split_comment(raw: pd.DataFrame, media: str, d_last: pd.Timestamp, d_prev: pd.Timestamp) -> str:
    r = raw.copy()
    r["d"] = r["날짜"].dt.normalize()
    r = r[r["매체상세"].astype(str) == str(media)]
    r["_분류"] = r[COL_PROMO].map(classify_promo)
    dl = pd.Timestamp(d_last).normalize()
    dp = pd.Timestamp(d_prev).normalize()
    lines = []
    for label in ["상시", "프로모션"]:
        sub = r[r["_분류"] == label]
        if sub.empty:
            continue
        def _agg(d: pd.Timestamp) -> tuple[float, float, float]:
            s2 = sub[sub["d"] == d]
            c = float(s2[COL_COST].sum())
            rv = float(s2[COL_REV_DB].sum())
            return c, rv, rv / c * 100 if c else 0.0
        c1, r1, o1 = _agg(dl)
        c0, r0, o0 = _agg(dp)
        lines.append(
            f"- **{label}**: 매출 {fm(r1)}원 ({r1-r0:+,.0f}원), "
            f"비용 {fm(c1)}원 ({c1-c0:+,.0f}원), ROAS **{o1:.1f}%** ({o1-o0:+.1f}p)"
        )
    return "\n".join(lines) if lines else "(데이터 없음)"


# ── 상품 Top10 ────────────────────────────────────────────────────────────────
def product_top10_table(
    raw: pd.DataFrame, media: str, d_last: pd.Timestamp, d_prev: pd.Timestamp
) -> tuple[pd.DataFrame, str]:
    dl = pd.Timestamp(d_last).normalize()
    dp = pd.Timestamp(d_prev).normalize()
    r = raw.copy()
    r["d"] = r["날짜"].dt.normalize()
    r = r[r["매체상세"].astype(str) == str(media)]
    r[COL_PRODUCT] = r[COL_PRODUCT].fillna("(미입력)")
    last = r[r["d"] == dl].groupby(COL_PRODUCT)[[COL_REV_DB, COL_COST, COL_BUY_DB]].sum()
    last = last.sort_values(COL_REV_DB, ascending=False).head(10)
    prev = r[r["d"] == dp].groupby(COL_PRODUCT)[[COL_REV_DB, COL_COST, COL_BUY_DB]].sum()
    rows = []
    for prod in last.index:
        rv1 = float(last.loc[prod, COL_REV_DB]); c1 = float(last.loc[prod, COL_COST]); b1 = float(last.loc[prod, COL_BUY_DB])
        rv0 = float(prev.loc[prod, COL_REV_DB]) if prod in prev.index else 0.0
        c0  = float(prev.loc[prod, COL_COST])   if prod in prev.index else 0.0
        b0  = float(prev.loc[prod, COL_BUY_DB]) if prod in prev.index else 0.0
        rows.append({
            COL_PRODUCT:   prod,
            "구매_전일":   b1,  "구매_전전일": b0,  "구매_증감":   b1-b0,
            "매출_전일":   rv1, "매출_전전일": rv0, "매출_증감":   rv1-rv0,
            "ROAS%_전일":  round(rv1/c1*100 if c1 else 0, 1),
            "ROAS%_전전일": round(rv0/c0*100 if c0 else 0, 1),
            "ROAS%p_증감": round((rv1/c1-rv0/c0)*100 if c1 and c0 else 0, 1),
        })
    out = pd.DataFrame(rows)
    if out.empty:
        return out, ""
    up = int((out["매출_증감"] > 0).sum()); dn = int((out["매출_증감"] < 0).sum())
    top_up = out.sort_values("매출_증감", ascending=False).head(2)[COL_PRODUCT].tolist()
    return out, f"Top10 중 매출 증가 **{up}**개 · 감소 **{dn}**개. 증가 폭 큰 상품: {', '.join(top_up) if top_up else '—'}."


# ── 매체 피벗 ─────────────────────────────────────────────────────────────────
def media_daily_pivot(raw: pd.DataFrame) -> pd.DataFrame:
    r = raw.copy()
    r["d"] = r["날짜"].dt.normalize()
    g = r.groupby(["매체상세", "d"], dropna=False)[COL_REV_DB].sum().reset_index()
    p = g.pivot(index="매체상세", columns="d", values=COL_REV_DB)
    p = p.reindex(sorted(p.columns), axis=1)
    p.columns = [pd.Timestamp(c).strftime("%m-%d") for c in p.columns]
    return p.fillna(0)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
def main() -> None:
    st.set_page_config(page_title="SharkNinja Daily", layout="wide", initial_sidebar_state="expanded")

    with st.sidebar:
        st.subheader("리포트 파일")
        path: Path

        if _madup_secrets_ok():
            st.caption("🔗 **Madup API** (Dropbox 다운로드)")
            api_key = str(st.secrets["MADUP_API_KEY"]).strip()
            base = str(st.secrets.get("MADUP_API_BASE") or MADUP_API_DEFAULT_BASE).strip()
            folder = str(st.secrets.get("MADUP_DROPBOX_FOLDER") or "").strip()
            single_path = str(
                st.secrets.get("MADUP_DROPBOX_PATH")
                or st.secrets.get("MADUP_DROPBOX_FILE_PATH")
                or ""
            ).strip()

            if folder:
                use_m = st.radio("파일", ["최신 자동", "날짜 선택"], horizontal=True)
                if use_m == "최신 자동":
                    try:
                        dp, data = _madup_latest_pair_cached(api_key, folder, base)
                    except Exception as e:
                        st.error(f"Madup API: {e}")
                        st.stop()
                    path = Path(tempfile.gettempdir()) / "sn_madup_latest.xlsx"
                    path.write_bytes(data)
                    st.caption(f"열림: `{dp}`")
                else:
                    tags = [
                        (datetime.now() - timedelta(days=i)).strftime("%y%m%d")
                        for i in range(120)
                    ]
                    pick = st.selectbox(
                        "리포트 날짜 (파일명 YYMMDD)",
                        tags,
                        format_func=lambda t: f"{t[:2]}/{t[2:4]}/{t[4:6]} ({t})",
                    )
                    data = None
                    dp = ""
                    for fname in madup_report_filenames(pick):
                        cand = f"{folder.rstrip('/')}/{fname}"
                        try:
                            data = _madup_bytes_cached(api_key, cand, base)
                            dp = cand
                            break
                        except Exception:
                            continue
                    if data is None:
                        st.error("해당 날짜 파일이 없거나 API 오류(파일명 공백/하이픈 형식 모두 시도함).")
                        st.stop()
                    path = Path(tempfile.gettempdir()) / f"sn_madup_{pick}.xlsx"
                    path.write_bytes(data)
                    st.caption(f"열림: `{dp}`")
            elif single_path:
                try:
                    data = _madup_bytes_cached(api_key, single_path, base)
                except Exception as e:
                    st.error(f"Madup API: {e}")
                    st.stop()
                path = Path(tempfile.gettempdir()) / "sn_madup_single.xlsx"
                path.write_bytes(data)
                st.caption(f"열림: `{single_path}`")
            else:
                st.error("MADUP_DROPBOX_FOLDER 또는 MADUP_DROPBOX_PATH 가 Secrets에 필요합니다.")
                st.stop()

        elif _drive_secrets_ok():
            st.caption("📁 **구글 드라이브** (Streamlit Secrets)")
            folder_id = str(st.secrets["GOOGLE_DRIVE_FOLDER_ID"]).strip()
            sa_raw = st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"]
            sa_json = json.dumps(sa_raw) if isinstance(sa_raw, dict) else str(sa_raw)
            try:
                listing = list_drive_reports(folder_id, sa_json)
            except Exception as e:
                st.error(f"Google Drive 연결 실패: {e}")
                st.stop()
            if not listing:
                st.error("드라이브 폴더에 `Madup_Sharkninja_Daily Report_YYMMDD.xlsx` 파일이 없습니다.")
                st.stop()

            use_file = st.radio("파일", ["최신 자동", "직접 선택"], horizontal=True)
            if use_file == "직접 선택":
                _tag, file_id, name = st.selectbox(
                    "파일 선택", listing, format_func=lambda x: x[2]
                )
            else:
                _tag, file_id, name = listing[0]

            data = _drive_cached_bytes(file_id, sa_json)
            path = Path(tempfile.gettempdir()) / f"sn_drive_{file_id}.xlsx"
            path.write_bytes(data)
            st.caption(f"열림: `{name}`")
        else:
            st.caption("💾 **로컬 폴더** (이 PC 경로)")
            ak = _secret_has("MADUP_API_KEY")
            df = _secret_has("MADUP_DROPBOX_FOLDER")
            sp = _secret_has("MADUP_DROPBOX_PATH") or _secret_has("MADUP_DROPBOX_FILE_PATH")
            cloudish = _server_os_cannot_use_embedded_windows_path()

            if cloudish:
                st.warning(
                    "**Madup Secrets 인식 여부** (하나라도 ❌면 지금 같은 로컬 폴더 화면만 뜹니다): "
                    f"`MADUP_API_KEY` {'✅' if ak else '❌'} · "
                    f"`MADUP_DROPBOX_FOLDER` {'✅' if df else '❌'} · "
                    f"`MADUP_DROPBOX_PATH` {'✅' if sp else '❌'}"
                )
                st.caption(
                    "❌이면 **키 값이 틀린 것보다**, Secrets **이름 오타·들여쓰기·`[섹션]` 안에 넣기**를 먼저 의심하세요. "
                    "아래처럼 **맨 위 평평한 두 줄**이어야 합니다. **Save → Reboot app**."
                )
                st.code(
                    'MADUP_API_KEY = "mk_..."\n'
                    'MADUP_DROPBOX_FOLDER = "/광고사업부/4. 광고주/샤크닌자/07. 리포트"',
                    language="toml",
                )
                st.error(
                    "Streamlit Cloud 서버에는 **내 PC `C:\\` 폴더가 없습니다.** "
                    "위 항목이 모두 ✅가 되어야 Madup으로 xlsx를 받습니다."
                )
                st.stop()

            report_dir = st.text_input("리포트 폴더", value=str(DEFAULT_REPORT_DIR))
            folder = Path(report_dir.strip())
            latest = find_latest_report(folder)
            if latest is None:
                st.error("리포트 xlsx를 찾을 수 없습니다.")
                st.info(
                    "**Streamlit Cloud** 배포 시: Secrets에 Madup 또는 드라이브 설정이 없으면 "
                    "위 오류가 납니다. 로컬과 동일한 키를 Cloud 앱 Secrets에 추가했는지 확인하세요."
                )
                st.stop()

            use_file = st.radio("파일", ["최신 자동", "직접 선택"], horizontal=True)
            if use_file == "직접 선택":
                opts = sorted(
                    [p for p in folder.glob("*.xlsx") if REPORT_FILE_RE.match(p.name)],
                    key=lambda p: REPORT_FILE_RE.match(p.name).group(1),  # type: ignore
                    reverse=True,
                )
                if not opts:
                    st.stop()
                path = st.selectbox("파일", opts, format_func=lambda p: p.name)
            else:
                path = latest

            st.caption(f"열림: `{path.name}`")
        page = st.radio("메뉴", ["Summary", "상품별", "매체별"], label_visibility="collapsed")
        if st.button("데이터 새로고침"):
            st.cache_data.clear()
            st.rerun()

    try:
        raw_df = load_raw(str(path))
    except Exception as e:
        st.exception(e)
        st.stop()

    dates = sorted(raw_df["날짜"].dt.normalize().dropna().unique())
    if not dates:
        st.error("날짜 데이터가 없습니다.")
        st.stop()

    d_max       = dates[-1]
    d_prev_file = dates[-2] if len(dates) >= 2 else d_max

    COL_SN = "샤크/닌자구분"
    sn_options = []
    if COL_SN in raw_df.columns:
        sn_options = sorted(raw_df[COL_SN].dropna().astype(str).unique())

    def _apply_sn(df: pd.DataFrame, val: str) -> pd.DataFrame:
        if val == "전체" or COL_SN not in df.columns:
            return df
        return df[df[COL_SN].astype(str) == val]

    # ══════════════════════════════════════════════════════════════════════════
    if page == "Summary":
        st.title("Summary")

        sn_summary = st.radio("브랜드 필터", ["전체"] + sn_options, horizontal=True, key="sn_summary")
        raw_s = _apply_sn(raw_df, sn_summary)

        # ── 1. 월 누계 TOTAL (2줄) ───────────────────────────────────────────
        mtot = agg_monthly(raw_s, d_max)
        first_of_m = pd.Timestamp(d_max.year, d_max.month, 1)
        st.subheader(
            f"📋 {d_max.month}월 누계 TOTAL  "
            f"({first_of_m.strftime('%m/%d')} ~ {d_max.strftime('%m/%d')})"
        )

        tbl_data = {
            "기준":       ["스마트스토어", "대시보드"],
            "노출":       [fm(mtot["노출"]),  fm(mtot["노출"])],
            "클릭":       [fm(mtot["클릭"]),  fm(mtot["클릭"])],
            "CPC (원)":   [fm(mtot["CPC"]),   fm(mtot["CPC"])],
            "CTR (%)":    [fp(mtot["CTR(%)"], 3), fp(mtot["CTR(%)"], 3)],
            "비용 (원)":  [fm(mtot["비용"]),  fm(mtot["비용"])],
            "구매 (건)":  [fm(mtot["SS_구매"]),  fm(mtot["DB_구매"])],
            "매출 (원)":  [fm(mtot["SS_매출"]),  fm(mtot["DB_매출"])],
            "ROAS (%)":   [fp(mtot["ROAS_SS"]), fp(mtot["ROAS_DB"])],
        }
        st.dataframe(
            pd.DataFrame(tbl_data).set_index("기준"),
            use_container_width=True, height=110,
        )

        # ── 일자별 상세 (스마트스토어 기준) ──
        with st.expander("📅 일자별 상세 보기", expanded=False):
            month_dates = [d for d in dates if pd.Timestamp(d).month == d_max.month]
            daily_rows = []
            for d in month_dates:
                a = agg_range(raw_s, d, d)
                daily_rows.append({
                    "날짜": pd.Timestamp(d).strftime("%Y-%m-%d"),
                    "노출": fm(a["노출"]),
                    "클릭": fm(a["클릭"]),
                    "CPC": fm(a["CPC"]),
                    "CTR": fp(a["CTR(%)"], 3) + "%",
                    "비용": fm(a["비용"]),
                    "SS 구매": fm(a["SS_구매"]),
                    "SS 매출": fm(a["SS_매출"]),
                    "SS ROAS": fp(a["ROAS_SS"]) + "%",
                    "DB 구매": fm(a["DB_구매"]),
                    "DB 매출": fm(a["DB_매출"]),
                    "DB ROAS": fp(a["ROAS_DB"]) + "%",
                })
            daily_df = pd.DataFrame(daily_rows).set_index("날짜")
            st.dataframe(daily_df, use_container_width=True, height=min(len(daily_rows) * 36 + 40, 600))

        st.divider()

        # ── 2. KPI 달성률 ────────────────────────────────────────────────────
        st.subheader(f"🎯 KPI 달성률  ({d_max.month}월 누계)")
        kpi_data = kpi_achievement(raw_df, d_max)
        k1, k2, k3 = st.columns(3)
        for col, (name, (achieved, target)) in zip([k1, k2, k3], kpi_data.items()):
            col.markdown(kpi_bar_html(name, achieved, target), unsafe_allow_html=True)

        st.divider()

        # ── 3. 기간별 성과 요약 ─────────────────────────────────────────────
        st.subheader("📅 기간별 성과 요약")

        col_base, col_comp = st.columns(2)
        with col_base:
            sel_base = st.date_input(
                "기준일 선택 날짜",
                value=(d_max.to_pydatetime().date(), d_max.to_pydatetime().date()),
                min_value=dates[0].to_pydatetime().date(),
                max_value=d_max.to_pydatetime().date(),
                key="base_dates",
            )
        with col_comp:
            d_prev_default = (d_max - timedelta(days=1)).to_pydatetime().date()
            sel_comp = st.date_input(
                "비교 기간 날짜",
                value=(d_prev_default, d_prev_default),
                min_value=dates[0].to_pydatetime().date(),
                max_value=d_max.to_pydatetime().date(),
                key="comp_dates",
            )

        def _parse_range(sel):
            if isinstance(sel, (list, tuple)) and len(sel) == 2:
                return pd.Timestamp(sel[0]).normalize(), pd.Timestamp(sel[1]).normalize()
            if isinstance(sel, (list, tuple)) and len(sel) == 1:
                t = pd.Timestamp(sel[0]).normalize()
                return t, t
            t = pd.Timestamp(sel).normalize()
            return t, t

        base_s, base_e = _parse_range(sel_base)
        comp_s, comp_e = _parse_range(sel_comp)

        cur  = agg_range(raw_s, base_s, base_e)
        comp = agg_range(raw_s, comp_s, comp_e)

        base_label = base_s.strftime('%m/%d') + (" – " + base_e.strftime('%m/%d') if base_s != base_e else "")
        comp_label = comp_s.strftime('%m/%d') + (" – " + comp_e.strftime('%m/%d') if comp_s != comp_e else "")

        # 데이터 기간 표 (4월 누계 TOTAL과 동일 형식)
        st.caption(f"데이터 기간: **{base_label}**")
        base_tbl = {
            "기준":      ["스마트스토어", "대시보드"],
            "노출":      [fm(cur["노출"]),  fm(cur["노출"])],
            "클릭":      [fm(cur["클릭"]),  fm(cur["클릭"])],
            "CPC (원)":  [fm(cur["CPC"]),   fm(cur["CPC"])],
            "CTR (%)":   [fp(cur["CTR(%)"], 3), fp(cur["CTR(%)"], 3)],
            "비용 (원)": [fm(cur["비용"]),  fm(cur["비용"])],
            "구매 (건)": [fm(cur["SS_구매"]), fm(cur["DB_구매"])],
            "매출 (원)": [fm(cur["SS_매출"]), fm(cur["DB_매출"])],
            "ROAS (%)":  [fp(cur["ROAS_SS"]), fp(cur["ROAS_DB"])],
        }
        st.dataframe(
            pd.DataFrame(base_tbl).set_index("기준"),
            use_container_width=True, height=110,
        )

        # 비교 기간 증감 표
        st.caption(f"비교 기간 증감  ({base_label} vs {comp_label})")

        def _delta_str(cur_v: float, comp_v: float, is_pct: bool = False) -> str:
            d = cur_v - comp_v
            if is_pct:
                return f"{d:+.1f}p"
            return f"{d:+,.0f}"

        def _color_val(cur_v: float, comp_v: float) -> str:
            d = cur_v - comp_v
            if d > 0:
                return "🔺"
            elif d < 0:
                return "🔻"
            return "➖"

        delta_tbl = {
            "기준":      ["스마트스토어", "대시보드"],
            "노출":      [f"{_color_val(cur['노출'],comp['노출'])} {_delta_str(cur['노출'],comp['노출'])}" for _ in range(1)] * 2,
            "클릭":      [f"{_color_val(cur['클릭'],comp['클릭'])} {_delta_str(cur['클릭'],comp['클릭'])}" for _ in range(1)] * 2,
            "CPC (원)":  [f"{_color_val(cur['CPC'],comp['CPC'])} {_delta_str(cur['CPC'],comp['CPC'])}" for _ in range(1)] * 2,
            "CTR (%)":   [f"{_color_val(cur['CTR(%)'],comp['CTR(%)'])} {_delta_str(cur['CTR(%)'],comp['CTR(%)'],True)}" for _ in range(1)] * 2,
            "비용 (원)": [f"{_color_val(cur['비용'],comp['비용'])} {_delta_str(cur['비용'],comp['비용'])}" for _ in range(1)] * 2,
            "구매 (건)": [
                f"{_color_val(cur['SS_구매'],comp['SS_구매'])} {_delta_str(cur['SS_구매'],comp['SS_구매'])}",
                f"{_color_val(cur['DB_구매'],comp['DB_구매'])} {_delta_str(cur['DB_구매'],comp['DB_구매'])}",
            ],
            "매출 (원)": [
                f"{_color_val(cur['SS_매출'],comp['SS_매출'])} {_delta_str(cur['SS_매출'],comp['SS_매출'])}",
                f"{_color_val(cur['DB_매출'],comp['DB_매출'])} {_delta_str(cur['DB_매출'],comp['DB_매출'])}",
            ],
            "ROAS (%)": [
                f"{_color_val(cur['ROAS_SS'],comp['ROAS_SS'])} {_delta_str(cur['ROAS_SS'],comp['ROAS_SS'],True)}",
                f"{_color_val(cur['ROAS_DB'],comp['ROAS_DB'])} {_delta_str(cur['ROAS_DB'],comp['ROAS_DB'],True)}",
            ],
        }
        st.dataframe(
            pd.DataFrame(delta_tbl).set_index("기준"),
            use_container_width=True, height=110,
        )

        st.divider()

        # ── 4. ROAS/비용 추이 (이중축 차트) ──────────────────────────────────
        chart_n = 7
        chart_dates = sorted(raw_df["날짜"].dt.normalize().dropna().unique())
        chart_use = [d for d in chart_dates if d <= d_max][-chart_n:]
        chart_from = pd.Timestamp(chart_use[0]).strftime("%m/%d") if chart_use else ""
        chart_to   = pd.Timestamp(chart_use[-1]).strftime("%m/%d") if chart_use else ""
        st.subheader(f"📊 스마트스토어 ROAS & 비용 추이  ({chart_from} – {chart_to}, 최근 {chart_n}일)")
        st.caption("막대: ROAS (%)  |  선: 비용 (만원)")
        fig = make_roas_cost_chart(raw_df, d_max, chart_n)
        st.plotly_chart(fig, use_container_width=True)

        st.divider()

        # ── 5. 뎁스 분석 코멘트 ──────────────────────────────────────────────
        st.subheader(
            f"🔍 데일리 코멘트  "
            f"({d_max.strftime('%m/%d')} vs {d_prev_file.strftime('%m/%d')})"
        )
        with st.spinner("코멘트 생성 중…"):
            comment = summary_depth_comment(raw_df, d_max, d_prev_file)
        st.markdown(comment, unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════════
    elif page == "상품별":
        st.title("상품별")
        st.caption("채널: 네이버브랜드스토어 · 실전환 광고 상품 (SS 기준) · 쿠팡/G마켓 제외")

        # 필터: 네이버브랜드스토어 + 실전환 상품 유효 값 + 쿠팡/G마켓 제외
        prod_df = raw_df.copy()
        prod_df["d"] = prod_df["날짜"].dt.normalize()
        if "채널" in prod_df.columns:
            prod_df = prod_df[prod_df["채널"].astype(str).str.strip() == "네이버브랜드스토어"]
        prod_df = prod_df[
            prod_df[COL_PRODUCT].notna() &
            (prod_df[COL_PRODUCT].astype(str).str.strip() != "") &
            (prod_df[COL_PRODUCT].astype(str).str.strip() != "-")
        ]
        prod_df = prod_df[~prod_df["매체상세"].astype(str).str.contains("쿠팡|G마켓", na=False)]

        # ── 샤크/닌자 필터 ──
        COL_SN = "샤크/닌자구분"
        if COL_SN in prod_df.columns:
            sn_vals = sorted(prod_df[COL_SN].dropna().astype(str).unique())
            sn_filter = st.radio("브랜드 필터", ["전체"] + sn_vals, horizontal=True, key="sn_filter")
            if sn_filter != "전체":
                prod_df = prod_df[prod_df[COL_SN].astype(str) == sn_filter]

        if prod_df.empty:
            st.warning("조건에 맞는 실전환 상품 데이터가 없습니다.")
        else:
            d_min_s = prod_df["d"].min().strftime("%m/%d")
            d_max_s = d_max.strftime("%m/%d")
            prev_s  = d_prev_file.strftime("%m/%d")

            # ── 1. 전체 기간 브랜드스토어 상품 매출 누적 TOP10 ──
            st.subheader("📦 전체 기간 네이버 브랜드스토어 상품 매출 TOP10 (SS 기준)")
            st.caption(f"기간: {d_min_s} – {d_max_s} 누적")

            total_prod = (
                prod_df.groupby(COL_PRODUCT, dropna=False)
                .agg({COL_REV_SS: "sum", COL_BUY_SS: "sum"})
                .sort_values(COL_REV_SS, ascending=False)
                .head(10)
                .reset_index()
            )
            total_prod.rename(columns={COL_REV_SS: "매출 (원)", COL_BUY_SS: "구매 (건)", COL_PRODUCT: "상품명"}, inplace=True)
            total_prod["short"] = total_prod["상품명"].apply(lambda x: x[:35] + "…" if len(str(x)) > 35 else str(x))

            fig_prod = go.Figure()
            fig_prod.add_trace(go.Bar(
                y=total_prod["short"][::-1],
                x=total_prod["매출 (원)"][::-1],
                orientation="h",
                marker=dict(
                    color=total_prod["매출 (원)"][::-1],
                    colorscale=[[0, "rgba(200,180,240,0.5)"], [0.5, "rgba(150,110,210,0.65)"], [1, "rgba(110,60,180,0.8)"]],
                    line=dict(color="rgba(100,60,170,0.5)", width=1),
                ),
                text=[f"{v:,.0f}원 / {b:,.0f}건" for v, b in zip(total_prod["매출 (원)"][::-1], total_prod["구매 (건)"][::-1])],
                textposition="outside",
                textfont=dict(size=11, color="#333"),
                constraintext="none",
            ))
            max_rev = total_prod["매출 (원)"].max() if not total_prod.empty else 1
            fig_prod.update_layout(
                height=420,
                plot_bgcolor="white", paper_bgcolor="white",
                xaxis=dict(title="매출 (원)", showgrid=True, gridcolor="rgba(200,200,200,0.3)",
                           range=[0, max_rev * 1.35]),
                yaxis=dict(title=None, tickfont=dict(size=11)),
                margin=dict(l=10, r=120, t=10, b=40),
            )
            st.plotly_chart(fig_prod, use_container_width=True)

            disp_total = total_prod[["상품명", "매출 (원)", "구매 (건)"]].copy()
            disp_total["매출 (원)"] = disp_total["매출 (원)"].map(lambda x: fm(float(x)))
            disp_total["구매 (건)"] = disp_total["구매 (건)"].map(lambda x: fm(float(x)))
            disp_total.index = range(1, len(disp_total) + 1)
            disp_total.index.name = "순위"
            st.dataframe(disp_total, use_container_width=True, height=390)

            st.divider()

            # ── 2. 매체상세별 누적 TOP10 + 전일 TOP10 ──
            st.subheader("📊 매체상세별 상품 매출 TOP10 (SS 기준)")
            naver_medias = sorted(prod_df["매체상세"].dropna().astype(str).unique())

            for media in naver_medias:
                m_sub = prod_df[prod_df["매체상세"].astype(str) == media]
                if m_sub.empty:
                    continue

                st.markdown(f"### {media}")

                if COL_SN in m_sub.columns:
                    sn_media_vals = sorted(m_sub[COL_SN].dropna().astype(str).unique())
                    sn_media_f = st.radio(
                        "브랜드", ["전체"] + sn_media_vals,
                        horizontal=True, key=f"sn_{media}",
                    )
                    if sn_media_f != "전체":
                        m_sub = m_sub[m_sub[COL_SN].astype(str) == sn_media_f]

                col_cum, col_daily = st.columns(2)

                with col_cum:
                    st.markdown(f"**누적 매출 TOP10** ({d_min_s} – {d_max_s})")
                    cum_top = (
                        m_sub.groupby(COL_PRODUCT, dropna=False)
                        .agg({COL_REV_SS: "sum", COL_BUY_SS: "sum"})
                        .sort_values(COL_REV_SS, ascending=False)
                        .head(10)
                        .reset_index()
                    )
                    cum_disp = cum_top.rename(columns={COL_PRODUCT: "상품명", COL_REV_SS: "매출 (원)", COL_BUY_SS: "구매 (건)"})
                    cum_disp["매출 (원)"] = cum_disp["매출 (원)"].map(lambda x: fm(float(x)))
                    cum_disp["구매 (건)"] = cum_disp["구매 (건)"].map(lambda x: fm(float(x)))
                    cum_disp.index = range(1, len(cum_disp) + 1)
                    cum_disp.index.name = "순위"
                    st.dataframe(cum_disp, use_container_width=True, height=390)

                with col_daily:
                    st.markdown(f"**전일 매출 TOP10** ({d_max_s})")
                    day_sub = m_sub[m_sub["d"] == d_max]
                    day_top = (
                        day_sub.groupby(COL_PRODUCT, dropna=False)
                        .agg({COL_REV_SS: "sum", COL_BUY_SS: "sum"})
                        .sort_values(COL_REV_SS, ascending=False)
                        .head(10)
                        .reset_index()
                    )
                    day_disp = day_top.rename(columns={COL_PRODUCT: "상품명", COL_REV_SS: "매출 (원)", COL_BUY_SS: "구매 (건)"})
                    day_disp["매출 (원)"] = day_disp["매출 (원)"].map(lambda x: fm(float(x)))
                    day_disp["구매 (건)"] = day_disp["구매 (건)"].map(lambda x: fm(float(x)))
                    day_disp.index = range(1, len(day_disp) + 1)
                    day_disp.index.name = "순위"
                    st.dataframe(day_disp, use_container_width=True, height=390)

                # ── 전일 대비 상승/하락 TOP10 ──
                st.markdown(f"**전일 대비 상품 변동** ({d_max_s} vs {prev_s})")
                m_prod_cmp = two_day_compare_ss(m_sub, [COL_PRODUCT], d_max, d_prev_file)
                if not m_prod_cmp.empty:
                    m_prod_cmp = m_prod_cmp[
                        (m_prod_cmp[COL_PRODUCT] != "(미입력)") &
                        (m_prod_cmp[COL_PRODUCT].astype(str).str.strip() != "-")
                    ]
                    col_up, col_dn = st.columns(2)

                    with col_up:
                        st.markdown("🔺 **매출 상승 TOP10**")
                        up = m_prod_cmp[m_prod_cmp["매출_증감"] > 0].head(10).copy()
                        if up.empty:
                            st.caption("상승 상품 없음")
                        else:
                            up_disp = up[[COL_PRODUCT, f"{COL_REV_SS}_전일", "매출_증감", "구매_증감"]].copy()
                            up_disp.columns = ["상품명", f"매출 ({d_max_s})", "매출 증감", "구매 증감"]
                            up_disp[f"매출 ({d_max_s})"] = up_disp[f"매출 ({d_max_s})"].map(lambda x: fm(float(x)))
                            up_disp["매출 증감"] = up_disp["매출 증감"].map(lambda x: f"+{fm(float(x))}")
                            up_disp["구매 증감"] = up_disp["구매 증감"].map(lambda x: f"{float(x):+,.0f}")
                            up_disp.index = range(1, len(up_disp) + 1)
                            st.dataframe(up_disp, use_container_width=True, height=390)

                    with col_dn:
                        st.markdown("🔻 **매출 하락 TOP10**")
                        dn = m_prod_cmp[m_prod_cmp["매출_증감"] < 0].sort_values("매출_증감").head(10).copy()
                        if dn.empty:
                            st.caption("하락 상품 없음")
                        else:
                            dn_disp = dn[[COL_PRODUCT, f"{COL_REV_SS}_전일", "매출_증감", "구매_증감"]].copy()
                            dn_disp.columns = ["상품명", f"매출 ({d_max_s})", "매출 증감", "구매 증감"]
                            dn_disp[f"매출 ({d_max_s})"] = dn_disp[f"매출 ({d_max_s})"].map(lambda x: fm(float(x)))
                            dn_disp["매출 증감"] = dn_disp["매출 증감"].map(lambda x: fm(float(x)))
                            dn_disp["구매 증감"] = dn_disp["구매 증감"].map(lambda x: f"{float(x):+,.0f}")
                            dn_disp.index = range(1, len(dn_disp) + 1)
                            st.dataframe(dn_disp, use_container_width=True, height=390)

                st.divider()

    # ══════════════════════════════════════════════════════════════════════════
    else:
        st.title("매체별")

        r_media = raw_df.copy()
        r_media["d"] = r_media["날짜"].dt.normalize()
        first_of_m = pd.Timestamp(d_max.year, d_max.month, 1)
        m_dates = [d for d in dates if d >= first_of_m]
        cum_label = f"{first_of_m.strftime('%m/%d')} – {d_max.strftime('%m/%d')}"
        prev_label = d_prev_file.strftime("%m/%d")
        last_label = d_max.strftime("%m/%d")

        # ── 1) 채널-매체상세 누적 성과 ──
        st.subheader(f"📊 채널-매체상세 누적 성과 ({cum_label})")

        sn_media_top = st.radio("브랜드 필터", ["전체"] + sn_options, horizontal=True, key="sn_media_top")
        r_media_f = _apply_sn(r_media, sn_media_top)

        st.caption("스마트스토어 기준 · 대시보드 기준 둘 다 표시")

        agg_cols_media = [COL_IMP, COL_CLICK, COL_COST, COL_REV_SS, COL_BUY_SS, COL_REV_DB, COL_BUY_DB]
        agg_cols_exist = [c for c in agg_cols_media if c in r_media_f.columns]
        m_cum = r_media_f[r_media_f["d"] >= first_of_m]

        ch_col = "채널" if "채널" in m_cum.columns else None
        group_keys = ([ch_col, "매체상세"] if ch_col else ["매체상세"])
        media_cum = m_cum.groupby(group_keys, dropna=False)[agg_cols_exist].sum().reset_index()
        media_cum["CPC"] = (media_cum[COL_COST] / media_cum[COL_CLICK].replace(0, float("nan"))).fillna(0)
        media_cum["CTR"] = (media_cum[COL_CLICK] / media_cum[COL_IMP].replace(0, float("nan")) * 100).fillna(0)
        media_cum["SS ROAS"] = (media_cum[COL_REV_SS] / media_cum[COL_COST].replace(0, float("nan")) * 100).fillna(0)
        media_cum["DB ROAS"] = (media_cum[COL_REV_DB] / media_cum[COL_COST].replace(0, float("nan")) * 100).fillna(0)
        media_cum = media_cum.sort_values(COL_REV_SS, ascending=False)

        disp_mcum = media_cum.copy()
        rename_map = {COL_IMP: "노출", COL_CLICK: "클릭", COL_COST: "비용",
                      COL_REV_SS: "SS 매출", COL_BUY_SS: "SS 구매",
                      COL_REV_DB: "DB 매출", COL_BUY_DB: "DB 구매"}
        disp_mcum.rename(columns=rename_map, inplace=True)
        for c in ["노출", "클릭", "비용", "SS 매출", "SS 구매", "DB 매출", "DB 구매", "CPC"]:
            if c in disp_mcum.columns:
                disp_mcum[c] = disp_mcum[c].map(lambda x: fm(float(x)))
        for c in ["CTR", "SS ROAS", "DB ROAS"]:
            if c in disp_mcum.columns:
                disp_mcum[c] = disp_mcum[c].map(lambda x: f"{float(x):.1f}%")
        st.dataframe(disp_mcum, use_container_width=True, height=min(len(disp_mcum) * 36 + 40, 600))

        st.divider()

        # ── 2~4) 매체상세별 상세 (캠페인/그룹 + 키워드/소재) ──
        st.subheader("📋 매체상세별 상세 성과")

        all_medias = sorted(r_media["매체상세"].dropna().astype(str).unique())

        KEYWORD_MEDIAS = {"네이버 파워링크", "네이버 쇼핑검색", "네이버 브랜드검색", "네이버 신제품검색"}
        CREATIVE_MEDIAS = {"네이버 쇼핑프로모션", "네이버 스마트채널", "네이버 메인배너"}
        COL_SS_KW = "스마트스토어센터_키워드"

        for media in all_medias:
            msub = r_media[r_media["매체상세"].astype(str) == media]
            if msub.empty:
                continue

            st.markdown(f"### {media}")

            if COL_SN in msub.columns:
                sn_m_vals = sorted(msub[COL_SN].dropna().astype(str).unique())
                sn_m_f = st.radio("브랜드", ["전체"] + sn_m_vals, horizontal=True, key=f"sn_media_{media}")
                msub = _apply_sn(msub, sn_m_f)

            # ── 캠페인/그룹 누적 성과 ──
            cg_cum = (
                msub[msub["d"] >= first_of_m]
                .groupby(["캠페인", "그룹"], dropna=False)[agg_cols_exist].sum()
                .reset_index()
            )
            if not cg_cum.empty:
                cg_cum["SS ROAS"] = (cg_cum[COL_REV_SS] / cg_cum[COL_COST].replace(0, float("nan")) * 100).fillna(0)
                cg_cum["DB ROAS"] = (cg_cum[COL_REV_DB] / cg_cum[COL_COST].replace(0, float("nan")) * 100).fillna(0)
                cg_cum = cg_cum.sort_values(COL_REV_SS, ascending=False)

                with st.expander(f"캠페인/그룹 누적 성과 ({cum_label})", expanded=False):
                    cg_d = cg_cum.head(30).copy()
                    cg_d.rename(columns=rename_map, inplace=True)
                    for c in ["노출", "클릭", "비용", "SS 매출", "SS 구매", "DB 매출", "DB 구매"]:
                        if c in cg_d.columns:
                            cg_d[c] = cg_d[c].map(lambda x: fm(float(x)))
                    for c in ["SS ROAS", "DB ROAS"]:
                        if c in cg_d.columns:
                            cg_d[c] = cg_d[c].map(lambda x: f"{float(x):.1f}%")
                    cg_d.index = range(1, len(cg_d) + 1)
                    st.dataframe(cg_d, use_container_width=True, height=min(len(cg_d) * 36 + 40, 500))

            # ── 캠페인/그룹 전일 대비 증감 ──
            cg_cmp = two_day_compare(msub, ["캠페인", "그룹"], d_max, d_prev_file)
            if not cg_cmp.empty:
                with st.expander(f"캠페인/그룹 전일 대비 증감 ({last_label} vs {prev_label}, DB 기준)", expanded=False):
                    cg_s = cg_cmp.head(20).copy()
                    show = ["캠페인", "그룹"]
                    for c in [f"{COL_REV_DB}_전일", "매출_증감", f"{COL_COST}_전일", "비용_증감"]:
                        if c in cg_s.columns:
                            show.append(c)
                            cg_s[c] = cg_s[c].map(lambda x: fm(float(x)))
                    for c in ["ROAS%_전일", "ROAS%p_차이"]:
                        if c in cg_s.columns:
                            show.append(c)
                            cg_s[c] = cg_s[c].map(lambda x: f"{float(x):.1f}%")
                    cg_s.index = range(1, len(cg_s) + 1)
                    st.dataframe(cg_s[show], use_container_width=True, height=min(len(cg_s) * 36 + 40, 500))

            # ── 키워드별 성과 (파워링크/쇼핑검색/브랜드검색/신제품검색) ──
            if media in KEYWORD_MEDIAS and COL_SS_KW in msub.columns:
                kw_sub = msub[msub[COL_SS_KW].notna() & (msub[COL_SS_KW].astype(str).str.strip() != "") & (msub[COL_SS_KW].astype(str).str.strip() != "-")]
                if not kw_sub.empty:
                    kw_cum = (
                        kw_sub[kw_sub["d"] >= first_of_m]
                        .groupby(COL_SS_KW, dropna=False)[[COL_REV_SS, COL_BUY_SS, COL_COST]].sum()
                        .reset_index()
                    )
                    kw_cum["SS ROAS"] = (kw_cum[COL_REV_SS] / kw_cum[COL_COST].replace(0, float("nan")) * 100).fillna(0)
                    kw_cum = kw_cum.sort_values(COL_REV_SS, ascending=False)

                    with st.expander(f"스마트스토어센터 키워드별 성과 ({cum_label}, SS 기준)", expanded=False):
                        col_kw_cum, col_kw_chg = st.columns(2)
                        with col_kw_cum:
                            st.markdown("**누적 매출 TOP20**")
                            kd = kw_cum.head(20).copy()
                            kd.rename(columns={COL_SS_KW: "키워드", COL_REV_SS: "SS 매출", COL_BUY_SS: "SS 구매", COL_COST: "비용"}, inplace=True)
                            for c in ["SS 매출", "SS 구매", "비용"]:
                                kd[c] = kd[c].map(lambda x: fm(float(x)))
                            kd["SS ROAS"] = kd["SS ROAS"].map(lambda x: f"{float(x):.1f}%")
                            kd.index = range(1, len(kd) + 1)
                            st.dataframe(kd, use_container_width=True, height=min(len(kd) * 36 + 40, 500))
                        with col_kw_chg:
                            st.markdown(f"**전일 대비 증감** ({last_label} vs {prev_label})")
                            kw_cmp = two_day_compare_ss(kw_sub, [COL_SS_KW], d_max, d_prev_file)
                            if not kw_cmp.empty:
                                kw_c = kw_cmp.head(20).copy()
                                kw_show = kw_c[[COL_SS_KW, f"{COL_REV_SS}_전일", "매출_증감", "구매_증감"]].copy()
                                kw_show.columns = ["키워드", f"SS 매출 ({last_label})", "매출 증감", "구매 증감"]
                                kw_show[f"SS 매출 ({last_label})"] = kw_show[f"SS 매출 ({last_label})"].map(lambda x: fm(float(x)))
                                kw_show["매출 증감"] = kw_show["매출 증감"].map(lambda x: f"{float(x):+,.0f}")
                                kw_show["구매 증감"] = kw_show["구매 증감"].map(lambda x: f"{float(x):+,.0f}")
                                kw_show.index = range(1, len(kw_show) + 1)
                                st.dataframe(kw_show, use_container_width=True, height=min(len(kw_show) * 36 + 40, 500))
                            else:
                                st.caption("증감 데이터 없음")

            # ── 소재/키워드별 성과 (쇼핑프로모션/스마트채널/메인배너) ──
            if media in CREATIVE_MEDIAS and "소재/키워드" in msub.columns:
                cr_sub = msub[msub["소재/키워드"].notna() & (msub["소재/키워드"].astype(str).str.strip() != "") & (msub["소재/키워드"].astype(str).str.strip() != "-")]
                if not cr_sub.empty:
                    with st.expander(f"소재/키워드별 성과 ({cum_label})", expanded=False):
                        # 필터: 캠페인 / 그룹
                        cr_camps = sorted(cr_sub["캠페인"].dropna().astype(str).unique())
                        cr_camp_f = st.selectbox("캠페인", ["전체"] + cr_camps, key=f"cr_camp_{media}")
                        cr_filtered = cr_sub if cr_camp_f == "전체" else cr_sub[cr_sub["캠페인"].astype(str) == cr_camp_f]

                        cr_grps = sorted(cr_filtered["그룹"].dropna().astype(str).unique())
                        cr_grp_f = st.selectbox("그룹", ["전체"] + cr_grps, key=f"cr_grp_{media}")
                        if cr_grp_f != "전체":
                            cr_filtered = cr_filtered[cr_filtered["그룹"].astype(str) == cr_grp_f]

                        cr_cum = (
                            cr_filtered[cr_filtered["d"] >= first_of_m]
                            .groupby("소재/키워드", dropna=False)[[COL_REV_SS, COL_BUY_SS, COL_COST, COL_REV_DB, COL_BUY_DB]].sum()
                            .reset_index()
                        )
                        cr_cum["SS ROAS"] = (cr_cum[COL_REV_SS] / cr_cum[COL_COST].replace(0, float("nan")) * 100).fillna(0)
                        cr_cum["DB ROAS"] = (cr_cum[COL_REV_DB] / cr_cum[COL_COST].replace(0, float("nan")) * 100).fillna(0)
                        cr_cum = cr_cum.sort_values(COL_REV_DB, ascending=False)

                        col_cr_cum, col_cr_chg = st.columns(2)
                        with col_cr_cum:
                            st.markdown("**누적 TOP20**")
                            cd = cr_cum.head(20).copy()
                            cd.rename(columns={COL_REV_SS: "SS 매출", COL_BUY_SS: "SS 구매", COL_COST: "비용", COL_REV_DB: "DB 매출", COL_BUY_DB: "DB 구매"}, inplace=True)
                            for c in ["SS 매출", "SS 구매", "비용", "DB 매출", "DB 구매"]:
                                if c in cd.columns:
                                    cd[c] = cd[c].map(lambda x: fm(float(x)))
                            for c in ["SS ROAS", "DB ROAS"]:
                                cd[c] = cd[c].map(lambda x: f"{float(x):.1f}%")
                            cd.index = range(1, len(cd) + 1)
                            st.dataframe(cd, use_container_width=True, height=min(len(cd) * 36 + 40, 500))
                        with col_cr_chg:
                            st.markdown(f"**전일 대비 증감** ({last_label} vs {prev_label}, DB 기준)")
                            cr_cmp = two_day_compare(cr_filtered, ["소재/키워드"], d_max, d_prev_file)
                            if not cr_cmp.empty:
                                cr_c = cr_cmp.head(20).copy()
                                cr_show = cr_c[["소재/키워드", f"{COL_REV_DB}_전일", "매출_증감", "구매_증감"]].copy()
                                cr_show.columns = ["소재/키워드", f"DB 매출 ({last_label})", "매출 증감", "구매 증감"]
                                cr_show[f"DB 매출 ({last_label})"] = cr_show[f"DB 매출 ({last_label})"].map(lambda x: fm(float(x)))
                                cr_show["매출 증감"] = cr_show["매출 증감"].map(lambda x: f"{float(x):+,.0f}")
                                cr_show["구매 증감"] = cr_show["구매 증감"].map(lambda x: f"{float(x):+,.0f}")
                                cr_show.index = range(1, len(cr_show) + 1)
                                st.dataframe(cr_show, use_container_width=True, height=min(len(cr_show) * 36 + 40, 500))
                            else:
                                st.caption("증감 데이터 없음")

            st.divider()


if __name__ == "__main__":
    main()
