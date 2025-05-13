import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

# 파일 경로
file_path = 'data/HVDC-STATUS-cleaned.xlsx'

# Excel 파일 읽기
df = pd.read_excel(file_path)

# 1. 물류 현황 대시보드
def create_logistics_dashboard():
    # 1.1 컨테이너 타입별 현황
    container_types = ['20DC', '40DC', 'LCL']
    container_data = {}
    for col in container_types:
        container_data[col] = pd.to_numeric(df[col], errors='coerce').sum()
    
    fig1 = px.bar(
        x=list(container_data.keys()),
        y=list(container_data.values()),
        title='컨테이너 타입별 현황',
        labels={'x': '컨테이너 타입', 'y': '수량'}
    )
    fig1.write_html('output/container_types.html')

    # 1.2 벤더별 물품 현황
    vendor_counts = df['VENDOR'].value_counts().head(10)
    fig2 = px.bar(
        x=vendor_counts.index,
        y=vendor_counts.values,
        title='상위 10개 벤더별 물품 현황',
        labels={'x': '벤더', 'y': '물품 수'}
    )
    fig2.write_html('output/vendor_distribution.html')

    # 1.3 카테고리별 분포
    category_counts = df['CATEGORY'].value_counts()
    fig3 = px.pie(
        values=category_counts.values,
        names=category_counts.index,
        title='카테고리별 분포'
    )
    fig3.write_html('output/category_distribution.html')

# 2. 시간별 추이 분석
def create_time_analysis():
    # 2.1 월별 물품 수송 현황
    df['ETD'] = pd.to_datetime(df['ETD'], errors='coerce')
    monthly_shipments = df.groupby(df['ETD'].dt.strftime('%Y-%m')).size()
    
    fig4 = px.line(
        x=monthly_shipments.index,
        y=monthly_shipments.values,
        title='월별 물품 수송 현황',
        labels={'x': '월', 'y': '수송 건수'}
    )
    fig4.write_html('output/monthly_shipments.html')

    # 2.2 운송선사별 현황
    shipping_line_counts = df['SHIPPING LINE'].value_counts().head(10)
    fig5 = px.bar(
        x=shipping_line_counts.index,
        y=shipping_line_counts.values,
        title='상위 10개 운송선사별 현황',
        labels={'x': '운송선사', 'y': '건수'}
    )
    fig5.write_html('output/shipping_line_distribution.html')

# 3. 물품 특성 분석
def create_item_analysis():
    # 3.1 중량 분포
    weight_data = pd.to_numeric(df['GWT\n (KG)'], errors='coerce')
    weight_data = weight_data[weight_data.notna()]  # 결측치 제거
    
    fig6 = px.histogram(
        x=weight_data,
        title='물품 중량 분포',
        labels={'x': '중량 (KG)'},
        nbins=30
    )
    fig6.write_html('output/weight_distribution.html')

    # 3.2 부피 분포
    volume_data = pd.to_numeric(df['CBM'], errors='coerce')
    volume_data = volume_data[volume_data.notna()]  # 결측치 제거
    
    fig7 = px.histogram(
        x=volume_data,
        title='물품 부피 분포',
        labels={'x': '부피 (CBM)'},
        nbins=30
    )
    fig7.write_html('output/volume_distribution.html')

# 4. 통합 대시보드
def create_integrated_dashboard():
    # 4.1 주요 지표 요약
    total_items = len(df)
    total_weight = pd.to_numeric(df['GWT\n (KG)'], errors='coerce').sum()
    total_volume = pd.to_numeric(df['CBM'], errors='coerce').sum()
    
    # 배송 기간 계산
    df['ETD'] = pd.to_datetime(df['ETD'], errors='coerce')
    df['ATA'] = pd.to_datetime(df['ATA'], errors='coerce')
    delivery_times = (df['ATA'] - df['ETD']).dt.days
    avg_delivery_time = delivery_times.mean()

    # HTML 형식의 대시보드 생성
    dashboard_html = f"""
    <html>
    <head>
        <title>HVDC 물류 현황 대시보드</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 20px; }}
            .container {{ display: flex; flex-wrap: wrap; gap: 20px; }}
            .card {{ 
                background: #f5f5f5;
                padding: 20px;
                border-radius: 10px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
                flex: 1;
                min-width: 200px;
            }}
            .card h3 {{ margin-top: 0; }}
            .card p {{ font-size: 24px; margin: 10px 0; }}
        </style>
    </head>
    <body>
        <h1>HVDC 물류 현황 대시보드</h1>
        <div class="container">
            <div class="card">
                <h3>총 물품 수</h3>
                <p>{total_items:,}건</p>
            </div>
            <div class="card">
                <h3>총 중량</h3>
                <p>{total_weight:,.0f} KG</p>
            </div>
            <div class="card">
                <h3>총 부피</h3>
                <p>{total_volume:,.0f} CBM</p>
            </div>
            <div class="card">
                <h3>평균 배송 기간</h3>
                <p>{avg_delivery_time:.1f}일</p>
            </div>
        </div>
    </body>
    </html>
    """

    with open('output/dashboard.html', 'w', encoding='utf-8') as f:
        f.write(dashboard_html)

# 출력 디렉토리 생성
os.makedirs('output', exist_ok=True)

# 시각화 실행
print("시각화를 시작합니다...")
create_logistics_dashboard()
create_time_analysis()
create_item_analysis()
create_integrated_dashboard()
print("시각화가 완료되었습니다. output 디렉토리에서 결과를 확인하세요.") 