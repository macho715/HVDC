import plotly.express as px
import pandas as pd

print('plotly, pandas import 완료')

df = pd.DataFrame({'x': [1,2,3], 'y': [4,5,6]})
fig = px.line(df, x='x', y='y', title='Test Dashboard Plot')
try:
    fig.write_image('test_dashboard_plot.png')
    print("프로젝트 루트에 test_dashboard_plot.png 생성 성공")
except Exception as e:
    print(f"프로젝트 루트에 이미지 생성 실패: {e}")
    import traceback
    traceback.print_exc() 