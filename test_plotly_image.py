import plotly.express as px
import pandas as pd

print('plotly, pandas import 완료')

df = pd.DataFrame({'x': [1,2,3], 'y': [4,5,6]})
fig = px.line(df, x='x', y='y', title='Test')
try:
    fig.write_image("test_plot.png")
    print("test_plot.png 파일이 성공적으로 생성되었습니다.")
except Exception as e:
    print(f"plotly 이미지 내보내기 중 오류 발생: {e}")
    import traceback
    traceback.print_exc() 