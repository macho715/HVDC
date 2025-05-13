# HVDC Status Analysis and Forecasting

HVDC 물류 데이터 분석 및 예측 시스템

## 주요 기능

- 데이터 분석 및 시각화
- 시계열 예측 (ARIMA vs Prophet)
- KPI 대시보드 생성
- 품질 점검 및 리포트

## 프로젝트 구조

```
HVDC/
├── data/               # 데이터 파일
├── output/            # 출력 파일 (대시보드, 리포트)
├── scripts/           # Python 스크립트
│   ├── automation/    # 자동화 스크립트
│   ├── core/         # 핵심 유틸리티
│   ├── data/         # 데이터 처리
│   ├── logistics/    # 물류 분석
│   └── reporting/    # 리포트 생성
└── tests/            # 테스트 코드
```

## 설치 및 실행 방법

### 0) 가상환경 설정 및 의존성 설치 (선택사항)

```bash
# 가상환경 생성 및 활성화
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows

# 필요한 패키지 설치
pip install pandas openpyxl xlsxwriter matplotlib seaborn plotly xlwings statsmodels prophet
```

### 1) 프로젝트 구조 설정

```bash
python setup_project.py  # data/, output/, scripts/ 디렉토리 생성 및 검증
```

### 2-A) 실제 데이터로 실행

```bash
# 1. 데이터 파일 복사
cp <신규파일>.xlsx data/HVDC-STATUS.xlsx

# 2. 데이터 처리 및 분석 실행
python scripts/logistics_mapper.py  # 매핑 및 리포트 생성
python scripts/quality_check.py     # 품질 점검
python scripts/dashboard.py         # KPI 대시보드 생성
```

### 2-B) 데모 데이터로 실행

```bash
python scripts/create_sample_data.py
```

### 3) 자동 파이프라인 실행

```bash
python scripts/auto_pipeline.py <신규파일>.xlsx
```

## 주요 파일 설명

- `fcast.py`: 시계열 예측 모델 (ARIMA vs Prophet)
- `analyze_data.py`: 데이터 분석 스크립트
- `visualize_data.py`: 데이터 시각화 스크립트
- `mPython.bas`: Excel VBA Python 연동 모듈
- `mChart.bas`: 차트 생성 VBA 모듈
- `mPivot.bas`: 피벗 테이블 처리 VBA 모듈

## 출력 파일

- `output/dashboard.html`: 통합 대시보드
- `output/HVDC_KPI_Dashboard.xlsx`: KPI 대시보드
- `output/quality_check/`: 품질 점검 리포트
- `output/logistics_mapping.xlsx`: 물류 매핑 결과

## 라이선스

MIT License