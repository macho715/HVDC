# HVDC Status Analysis and Forecasting

이 프로젝트는 HVDC(High Voltage Direct Current) 물류 데이터를 분석하고 예측하는 도구입니다.

## 주요 기능

1. 데이터 분석
   - 결측치 처리
   - 기본 통계 분석
   - 데이터 시각화

2. 시계열 예측
   - ARIMA 모델
   - Prophet 모델
   - 자동 모델 선택 및 비교

## 설치 방법

1. 필요한 패키지 설치:
```bash
pip install -r requirements.txt
```

2. xlwings Excel 애드인 설치:
```bash
xlwings addin install
```

## 사용 방법

1. 데이터 분석:
```bash
python analyze_data.py
```

2. 데이터 시각화:
```bash
python visualize_data.py
```

3. 예측 실행:
```bash
python fcast.py
```

## 프로젝트 구조

- `analyze_data.py`: 데이터 분석 스크립트
- `visualize_data.py`: 데이터 시각화 스크립트
- `fcast.py`: 시계열 예측 스크립트
- `data/`: 데이터 파일 디렉토리
- `output/`: 시각화 결과 저장 디렉토리

## 의존성

- Python 3.8+
- pandas
- numpy
- statsmodels
- prophet
- xlwings
- plotly
- scikit-learn

## Troubleshooting
- 매크로 차단: 신뢰할 수 있는 위치에 저장
- xlwings not found: 가상환경 활성화 후 pip 재설치
- PQ 오류: PQ 편집기에서 열/시트명 동기화