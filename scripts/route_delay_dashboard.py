import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import argparse
from pathlib import Path

PROCESS_WEIGHTS = {"AGI": 5, "DAS": 5, "SHU": 0, "MIR": 0}

def calc_allowed_delay(site, base_delay):
    return base_delay + PROCESS_WEIGHTS.get(site, 0)

def flag_delay(row, delay_col, base_delay=3):
    allowed = calc_allowed_delay(row['SITE'], base_delay)
    return row[delay_col] > allowed

def route_delay_summary(df, delay_flag_col='DELAY_FLAG'):
    return df.groupby('SITE').agg(
        delay_count=(delay_flag_col, 'sum'),
        total=('NO.', 'count'),
        delay_ratio=(delay_flag_col, 'mean')
    ).reset_index()

def main(file, export='html'):
    df = pd.read_excel(file, sheet_name='STEP_FLOW')
    df['DELAY_FLAG'] = df.apply(lambda row: flag_delay(row, '입항→통관', base_delay=3), axis=1)
    summary = route_delay_summary(df, 'DELAY_FLAG')

    # 히트맵
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

    # 결과 저장
    if export == 'html':
        out_path = 'output/route_delay_dashboard.html'
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(fig1.to_html(full_html=False, include_plotlyjs='cdn'))
            f.write('<hr>')
            f.write(table.to_html(full_html=False, include_plotlyjs='cdn'))
        print(f"HTML 대시보드 저장 완료: {out_path}")
    elif export == 'xlsx':
        out_path = 'output/route_delay_report.xlsx'
        with pd.ExcelWriter(out_path) as writer:
            df.to_excel(writer, sheet_name='Detail', index=False)
            summary.to_excel(writer, sheet_name='Route_Delay_Summary', index=False)
        print(f"Excel 리포트 저장 완료: {out_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--file', type=str, default='output/logistics_mapping.xlsx', help='분석 대상 파일')
    parser.add_argument('--export', type=str, default='html', help='결과 내보내기 형식(html/xlsx)')
    args = parser.parse_args()
    main(args.file, args.export) 