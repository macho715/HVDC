import pandas as pd
import os
import numpy as np

# 파일 경로
file_path = 'data/HVDC-STATUS.xlsx'

# Excel 파일 읽기
df = pd.read_excel(file_path)

# 결측치 처리 전략
def handle_missing_values(df):
    # 1. MR# 결측치 처리 (19개)
    df['MR#'] = df['MR#'].fillna('Unknown')
    
    # 2. CATEGORY 결측치 처리 (6개)
    df['CATEGORY'] = df['CATEGORY'].fillna('Uncategorized')
    
    # 3. MAIN DESCRIPTION 결측치 처리 (19개)
    df['MAIN DESCRIPTION (PO)'] = df['MAIN DESCRIPTION (PO)'].fillna('No Description')
    
    # 4. SHIPPING LINE 결측치 처리 (160개)
    df['SHIPPING LINE'] = df['SHIPPING LINE'].fillna('Not Assigned')
    
    # 5. FORWARDER 결측치 처리 (83개)
    df['FORWARDER'] = df['FORWARDER'].fillna('Not Assigned')
    
    # 6. 날짜 관련 결측치 처리
    date_columns = ['ATA', 'Attestation\n Date', 'DO Collection', 'Customs\n Start']
    for col in date_columns:
        df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # 7. 창고 관련 결측치 처리
    warehouse_columns = ['SHU', 'DAS', 'MIR', 'AGI', 'SHU.1', 'MIR.1', 'DAS.1', 'AGI.1',
                        'DSV\n Indoor', 'DSV\n Outdoor', 'DSV\n MZD', 'JDN\n MZD',
                        'JDN\n Waterfront', 'MOSB', 'AAA Storage', 'ZENER (WH)',
                        'Hauler DG Storage', 'Vijay Tanks']
    
    for col in warehouse_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    return df

# 결측치 처리 실행
df_cleaned = handle_missing_values(df)

# 처리 결과 확인
print('=== 결측치 처리 전 ===')
print(df.isnull().sum())
print('\n=== 결측치 처리 후 ===')
print(df_cleaned.isnull().sum())

# 처리된 데이터 저장
output_path = 'data/HVDC-STATUS-cleaned.xlsx'
df_cleaned.to_excel(output_path, index=False)
print(f'\n처리된 데이터가 {output_path}에 저장되었습니다.')

# 기본 정보 출력
print('데이터 크기:', df_cleaned.shape)
print('\n컬럼 목록:', list(df_cleaned.columns))
print('\n데이터 미리보기:')
print(df_cleaned.head())

# 기본 통계 정보
print('\n기본 통계 정보:')
print(df_cleaned.describe()) 