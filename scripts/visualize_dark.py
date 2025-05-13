import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

# === 컬러·폰트·테마 팔레트 ===
BG   = "#021321"
PANEL= "#06263A"
C1   = "#00D2FF"   # Primary
C2   = "#25F4EE"   # Secondary
FONT = "Inter"
FONT_CDN = '<link href="https://fonts.googleapis.com/css?family=Inter:400,700&display=swap" rel="stylesheet">'

# === 데이터 로드 ===
# 실적/예측 데이터
hist_path = "output/logistics_mapping.xlsx"
fcst_path = "output/forecast_dashboard.html"

# STEP_FLOW 시트에서 ATA, SITE, 전체 리드타임
df = pd.read_excel(hist_path, sheet_name="STEP_FLOW")
series = df[["ATA", "SITE", "전체 리드타임"]].rename(columns={"ATA": "ds", "전체 리드타임": "y"})
series = series.dropna(subset=["ds", "y", "SITE"])
series = series.sort_values("ds")

# 최근 1년치 실적
df_hist = series.tail(365)

# 예측 데이터: forecast_dashboard.html에서 추출 불가하므로, fcast.py에서 df_fcst를 output/forecast.csv로 저장했다고 가정
fcst_csv = "output/forecast.csv"
if os.path.exists(fcst_csv):
    df_fcst = pd.read_csv(fcst_csv, parse_dates=["ds"])
else:
    # 임시: 실적 마지막값을 180일 복제
    import numpy as np
    last_date = df_hist["ds"].max()
    future_dates = pd.date_range(last_date, periods=180, freq="D")
    df_fcst = pd.DataFrame({"ds": future_dates, "y": np.nan})

# KPI 계산
avg_lt = df_hist["y"].tail(30).mean()
rmse = 0  # 실제 RMSE 값은 fcast.py에서 전달 필요

# 히트맵 데이터: SITE×월별 평균 리드타임
series["month"] = series["ds"].dt.to_period("M")
heat_df = (series.groupby(["SITE", "month"])["y"]
           .mean().reset_index()
           .rename(columns={"y":"mean_delay", "month":"month"}))

# === 대시보드 생성 ===
def dashboard_style(fig):
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor=BG,
        plot_bgcolor=PANEL,
        font=dict(family=FONT, size=12, color="#B7CADB"),
        margin=dict(t=60, l=20, r=20, b=20),
        hovermode="x unified"
    )
    return fig

def build_dashboard(df_hist, df_fcst, heat_df, avg_lt, rmse, out="output/dark_dash.html"):
    fig = make_subplots(
        rows=3, cols=3,
        specs=[[{"type":"indicator"},{"type":"indicator"},{"type":"indicator"}],
               [{"colspan":3}, None, None],
               [{"type":"heatmap"},{"type":"table","colspan":2}, None]],
        column_widths=[0.33,0.33,0.34],
        row_heights=[0.20,0.25,0.55],
        vertical_spacing=0.04, horizontal_spacing=0.02
    )

    # ── KPI 카드들 ──────────────────────
    fig.add_trace(go.Indicator(
        mode="number", value=avg_lt,
        number={"suffix":" d","font":{"size":36,"color":C1}},
        title={"text":"<b>최근 30일 평균 LT</b>"},
        ), row=1, col=1)

    fig.add_trace(go.Indicator(
        mode="number", value=rmse,
        number={"suffix":" d","font":{"size":36,"color":C2}},
        title={"text":"<b>예측 RMSE</b>"},
        ), row=1, col=2)

    fig.add_trace(go.Indicator(
        mode="number+delta",
        value=(df_hist["y"].iloc[-1]),
        delta=dict(reference=df_hist["y"].iloc[-30]),
        number={"suffix":" d","font":{"size":36}},
        title={"text":"<b>최근 vs 30일 전</b>"},
        ), row=1, col=3)

    # ── 실적 vs 예측 라인 ──────────────
    fig.add_trace(go.Scatter(
        x=df_hist["ds"], y=df_hist["y"],
        mode="lines", name="Actual",
        line=dict(color=C1, width=2)), row=2, col=1)

    fig.add_trace(go.Scatter(
        x=df_fcst["ds"], y=df_fcst["y"],
        mode="lines", name="Forecast",
        line=dict(color=C2, width=2, dash="dot")), row=2, col=1)

    # ── SITE × 월 히트맵 ───────────────
    fig.add_trace(go.Heatmap(
        z=heat_df["mean_delay"],
        x=heat_df["month"].astype(str),
        y=heat_df["SITE"],
        colorscale=[[0,"#013A55"],[1,C1]],
        colorbar=dict(title="Days")), row=3, col=1)

    # ── Ranking Table (Top5 지연) ──────
    top5 = (df_hist.sort_values("y", ascending=False)
                  .head(5)[["SITE","y"]]
                  .rename(columns={"y":"Delay"}))
    fig.add_trace(go.Table(
        header=dict(values=["<b>SITE</b>","<b>Delay(d)</b>"],
                    fill_color=PANEL, font=dict(color=C1)),
        cells=dict(values=[top5.SITE, top5.Delay],
                   fill_color=BG, font=dict(color="#FFFFFF"))),
        row=3, col=2)

    # ── 스타일 통합 ────────────────────
    dashboard_style(fig)
    fig.update_layout(title="<b>HVDC Logistics Delay Dashboard</b>",
                      title_font_size=18)
    # HTML 저장 + Inter 폰트 적용
    html = fig.to_html(include_plotlyjs="cdn", full_html=True)
    html = html.replace("<head>", f"<head>{FONT_CDN}")
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    print("🔹 Saved:", out)

if __name__ == "__main__":
    build_dashboard(df_hist, df_fcst, heat_df, avg_lt, rmse) 