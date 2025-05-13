import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

# ... 기존 시각화 함수 ...
def create_route_delay_dashboard(input_file='output/logistics_mapping.xlsx', output_file='output/route_delay_dashboard.html'):
    df = pd.read_excel(input_file, sheet_name='STEP_FLOW')
    if 'DELAY_FLAG' not in df.columns:
        print('DELAY_FLAG 컬럼이 없습니다. analyze_data.py를 먼저 실행하세요.')
        return
    # 히트맵: SITE별 입항→통관 지연일수
    if '입항→통관' in df.columns:
        heatmap_data = df.pivot_table(index='SITE', columns='STEP_NAME', values='입항→통관', aggfunc='mean')
        fig1 = px.imshow(heatmap_data, text_auto=True, color_continuous_scale='Reds',
                        labels=dict(color='평균 지연일수'), title='SITE별/STEP별 평균 입항→통관 지연일수')
    else:
        fig1 = go.Figure()
    # Top-N 지연 리스트
    top_delay = df.sort_values('입항→통관', ascending=False).head(10)
    table = go.Figure(data=[go.Table(
        header=dict(values=list(top_delay[['NO.', 'SITE', '입항→통관']].columns), fill_color='paleturquoise', align='left'),
        cells=dict(values=[top_delay['NO.'], top_delay['SITE'], top_delay['입항→통관']], fill_color='lavender', align='left')
    )])
    table.update_layout(title='Top 10 입항→통관 지연 리스트')
    # HTML로 저장
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(fig1.to_html(full_html=False, include_plotlyjs='cdn'))
        f.write('<hr>')
        f.write(table.to_html(full_html=False, include_plotlyjs='cdn'))

def create_forecast_dashboard(df_hist, df_fcst, out_html):
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df_hist["ds"], y=df_hist["y"],
        mode="lines+markers", name="Actual",
        line=dict(width=2)))
    fig.add_trace(go.Scatter(
        x=df_fcst["ds"], y=df_fcst["y"],
        mode="lines", name="Forecast", line=dict(dash="dot")))
    fig.update_layout(
        template="plotly_white",
        title="리드타임 예측(6개월)",
        yaxis_title="Days",
        legend=dict(orientation="h"))
    fig.write_html(out_html, include_plotlyjs="cdn")

if __name__ == "__main__":
    create_route_delay_dashboard() 