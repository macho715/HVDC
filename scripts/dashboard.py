import pandas as pd
from pathlib import Path
import numpy as np
import xlsxwriter

# === 0. Load workbook ===
src_path = Path("output/logistics_mapping.xlsx")
wb = pd.ExcelFile(src_path)

step_flow = wb.parse("STEP_FLOW")
sla_exceed = wb.parse("SLA_Exceed")
container_sum = wb.parse("Container_Summary")

# === 1. Data cleansing ===
step_flow["ATA"] = pd.to_datetime(step_flow["ATA"], errors="coerce")
step_flow["전체 리드타임"] = pd.to_numeric(step_flow.get("전체 리드타임"), errors="coerce")

# 1‑1) Trend by month
step_flow["ATA_Month"] = step_flow["ATA"].dt.to_period("M").astype(str)
trend = (
    step_flow.groupby("ATA_Month")
    .agg(건수=("NO.", "count"), 평균_LT=("전체 리드타임", "mean"))
    .reset_index()
    .sort_values("ATA_Month")
)
trend["평균_LT"] = trend["평균_LT"].fillna(0).round(1)
trend["목표_LT"] = 30  # 목표 리드타임 컬럼 추가

# 1‑2) Leadtime by step
lead_by_step = (
    step_flow.groupby("STEP_NAME")["전체 리드타임"]
    .mean()
    .reset_index()
    .rename(columns={"전체 리드타임": "평균_LT"})
    .fillna(0)
    .round(1)
    .sort_values("평균_LT", ascending=False)
)

# 1‑3) SLA exceed by site
sla_by_site = (
    sla_exceed.groupby("SITE")["SITE"]
    .count()
    .reset_index(name="SLA_초과_건수")
    .sort_values("SLA_초과_건수", ascending=False)
)

# 1‑4) Container type share
total_20 = container_sum["20ft Q'TY"].sum() if "20ft Q'TY" in container_sum.columns else 0
total_40 = container_sum["40ft Q'TY"].sum() if "40ft Q'TY" in container_sum.columns else 0
type_share = pd.DataFrame({
    "컨테이너타입": ["20FT", "40FT"],
    "개수": [total_20, total_40]
})

# 1‑5) Top 10 shipments
if "TOTAL Q'TY" in container_sum.columns:
    top10 = (
        container_sum.nlargest(10, "TOTAL Q'TY")
        [["SCT SHIP NO.", "TOTAL Q'TY"]]
        .rename(columns={"TOTAL Q'TY": "컨테이너_총량"})
    )
else:
    top10 = pd.DataFrame(columns=["SCT SHIP NO.", "컨테이너_총량"])

# === 2. Write dashboard ===
out_path = "output/HVDC_KPI_Dashboard.xlsx"
workbook = xlsxwriter.Workbook(out_path)
fmt_title = workbook.add_format({
    "bold": True, "font_size": 16, "font_color": "#FFFFFF",
    "align": "center", "valign": "vcenter", "bg_color": "#1D2433"
})
fmt_bg = workbook.add_format({"bg_color": "#F2EDF3"})
fmt_hdr = workbook.add_format({"bold": True, "bg_color": "#D9D9D9", "border": 1})
fmt_red = workbook.add_format({"bg_color": "#FFC7CE"})

data_ws = workbook.add_worksheet("Data")
data_ws.hide()

def write_df(ws, df, start_row):
    ws.write_row(start_row, 0, df.columns.tolist(), fmt_hdr)
    df_filled = df.fillna('')
    for r, row_data in df_filled.iterrows():
        ws.write_row(start_row + 1 + r, 0, row_data.tolist())
    return int(start_row + 1 + len(df_filled) + 2)

row_idx = 0
trend_start = row_idx
row_idx = write_df(data_ws, trend, row_idx)
lead_start = row_idx
row_idx = write_df(data_ws, lead_by_step, row_idx)
sla_start = row_idx
row_idx = write_df(data_ws, sla_by_site, row_idx)
share_start = row_idx
row_idx = write_df(data_ws, type_share, row_idx)
top_start = row_idx
row_idx = write_df(data_ws, top10, row_idx)

# 누적 % 계산 및 추가
total = top10["컨테이너_총량"].sum()
top10["누적%"] = (top10["컨테이너_총량"].cumsum() / total * 100).astype(float).round(1)
row_idx = write_df(data_ws, top10[["SCT SHIP NO.", "누적%"]], row_idx)

line5 = workbook.add_chart({"type": "line"})
line5.add_series({
    "name": "누적 %",
    "categories": ["Data", row_idx - len(top10), 0, row_idx - 1, 0],
    "values": ["Data", row_idx - len(top10), 1, row_idx - 1, 1],
    "y2_axis": True,
    "line": {"color": "#FF0000"}
})

dash = workbook.add_worksheet("대시보드")
dash.set_tab_color("#1D2433")
dash.set_default_row(20)
dash.set_column("A:Q", 16, fmt_bg)
dash.merge_range("A1:Q2", "HVDC 물류 KPI 대시보드", fmt_title)

positions = ["A4", "I4", "A20", "I20", "A36"]

# Chart 1: 월별 처리량 & 평균 리드타임
chart1 = workbook.add_chart({"type": "column"})
chart1.add_series({
    "name": "건수",
    "categories": ["Data", trend_start + 1, 0, trend_start + len(trend), 0],
    "values": ["Data", trend_start + 1, 1, trend_start + len(trend), 1],
})
chart1.set_y_axis({"name": "건수"})
chart1.set_x_axis({"name": "월"})
chart1.set_title({"name": "월별 처리량 & 평균 리드타임"})

line1 = workbook.add_chart({"type": "line"})
line1.add_series({
    "name": "평균 LT",
    "categories": ["Data", trend_start + 1, 0, trend_start + len(trend), 0],
    "values": ["Data", trend_start + 1, 2, trend_start + len(trend), 2],
    "y2_axis": True,
    "line": {"color": "#FF0000"}
})
chart1.combine(line1)

# 목표 리드타임 라인 추가
# trend 컬럼 인덱스: 0=ATA_Month, 1=건수, 2=평균_LT, 3=목표_LT
목표LT_col_idx = 3
target_line = workbook.add_chart({"type": "line"})
target_line.add_series({
    "name": "목표 LT",
    "categories": ["Data", trend_start + 1, 0, trend_start + len(trend), 0],
    "values": ["Data", trend_start + 1, 목표LT_col_idx, trend_start + len(trend), 목표LT_col_idx],
    "y2_axis": True,
    "line": {"color": "#FF0000", "dash_type": "dash"}
})
chart1.combine(target_line)

dash.insert_chart(positions[0], chart1, {"x_offset": 5, "y_offset": 5})

# Chart 2: 공정별 평균 리드타임
chart2 = workbook.add_chart({"type": "bar"})
chart2.add_series({
    "categories": ["Data", lead_start + 1, 0, lead_start + len(lead_by_step), 0],
    "values": ["Data", lead_start + 1, 1, lead_start + len(lead_by_step), 1],
    "fill": {"color": "#4472C4"},
    "data_labels": {"value": True, "position": "outside_end"}
})
chart2.set_title({"name": "공정별 평균 리드타임"})
chart2.set_x_axis({"name": "일"})

# SLA 초과 조건부 서식
for i, row in enumerate(lead_by_step.itertuples(), start=lead_start + 2):
    if row.평균_LT > 30:
        data_ws.conditional_format(f"B{row}", {"type": "cell", "criteria": ">", "value": 30, "format": fmt_red})

dash.insert_chart(positions[1], chart2, {"x_offset": 5, "y_offset": 5})

# Chart 3: SITE별 SLA 초과 건수
chart3 = workbook.add_chart({"type": "bar"})
chart3.add_series({
    "categories": ["Data", sla_start + 1, 0, sla_start + len(sla_by_site), 0],
    "values": ["Data", sla_start + 1, 1, sla_start + len(sla_by_site), 1],
    "data_labels": {"value": True, "position": "outside_end"},
    "fill": {"color": "#ED7D31"}
})
chart3.set_title({"name": "SITE별 SLA 초과 건수"})
chart3.set_x_axis({"name": "건수"})
dash.insert_chart(positions[2], chart3, {"x_offset": 5, "y_offset": 5})

# Chart 4: 컨테이너 타입 비중
chart4 = workbook.add_chart({"type": "doughnut"})
chart4.add_series({
    "categories": ["Data", share_start + 1, 0, share_start + len(type_share), 0],
    "values": ["Data", share_start + 1, 1, share_start + len(type_share), 1],
    "data_labels": {"percentage": True, "leader_lines": True},
    "points": [
        {"fill": {"color": "#4472C4"}},
        {"fill": {"color": "#ED7D31"}}
    ]
})
chart4.set_title({"name": "컨테이너 타입 비중"})
chart4.set_rotation(90)
dash.insert_chart(positions[3], chart4, {"x_offset": 5, "y_offset": 5})

# Chart 5: Top10 SCT SHIP NO. 컨테이너
chart5 = workbook.add_chart({"type": "column"})
chart5.add_series({
    "categories": ["Data", top_start + 1, 0, top_start + len(top10), 0],
    "values": ["Data", top_start + 1, 1, top_start + len(top10), 1],
    "fill": {"color": "#4472C4"},
    "data_labels": {"value": True}
})

line5 = workbook.add_chart({"type": "line"})
line5.add_series({
    "name": "누적 %",
    "categories": ["Data", row_idx - len(top10), 0, row_idx - 1, 0],
    "values": ["Data", row_idx - len(top10), 1, row_idx - 1, 1],
    "y2_axis": True,
    "line": {"color": "#FF0000"}
})
chart5.combine(line5)

chart5.set_title({"name": "Top10 SCT SHIP NO. 컨테이너"})
chart5.set_x_axis({"name": "SCT SHIP NO."})
chart5.set_y_axis({"name": "컨테이너 수"})
chart5.set_y2_axis({"name": "누적 %"})

# 상위 3개 강조
for i in range(3):
    chart5.add_series({
        "categories": ["Data", top_start + 1, 0, top_start + 1, 0],
        "values": ["Data", top_start + 1 + i, 1, top_start + 1 + i, 1],
        "fill": {"color": "#ED7D31"}
    })

dash.insert_chart(positions[4], chart5, {"x_offset": 5, "y_offset": 5})

workbook.close()

print(f"[Download KPI Dashboard](sandbox:{out_path})") 