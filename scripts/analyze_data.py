import pandas as pd
from pathlib import Path

# 1. 가중치 정의
PROCESS_WEIGHTS = {"AGI": 5, "DAS": 5, "SHU": 0, "MIR": 0}

# 2. 허용 지연 계산 함수
def calc_allowed_delay(site, base_delay):
    return base_delay + PROCESS_WEIGHTS.get(site, 0)

# 3. 지연 플래그 함수
def flag_delay(row, delay_col, base_delay=3):
    allowed = calc_allowed_delay(row['SITE'], base_delay)
    return row[delay_col] > allowed

# 4. SITE별 지연 통계 요약 함수
def route_delay_summary(df, delay_flag_col='DELAY_FLAG'):
    return df.groupby('SITE').agg(
        delay_count=(delay_flag_col, 'sum'),
        total=('NO.', 'count'),
        delay_ratio=(delay_flag_col, 'mean')
    ).reset_index()

# 5. 메인 분석 함수 예시
def main(input_file='output/logistics_mapping.xlsx', output_file='output/route_delay_report.xlsx'):
    df = pd.read_excel(input_file, sheet_name='STEP_FLOW')
    # 예시: '입항→통관' 컬럼에 대해 지연 플래그 생성
    df['DELAY_FLAG'] = df.apply(lambda row: flag_delay(row, '입항→통관', base_delay=3), axis=1)
    summary = route_delay_summary(df, 'DELAY_FLAG')
    # 결과 저장
    with pd.ExcelWriter(output_file) as writer:
        df.to_excel(writer, sheet_name='Detail', index=False)
        summary.to_excel(writer, sheet_name='Route_Delay_Summary', index=False)

if __name__ == "__main__":
    main()

df = pd.read_excel("data/HVDC-STATUS(20250513).xlsx", sheet_name="STATUS")
print(df.columns.tolist()) 