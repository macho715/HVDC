import numpy as np, pandas as pd, joblib, pathlib
import pmdarima as pm                    # auto_arima
from prophet import Prophet
from sklearn.model_selection import TimeSeriesSplit
from sklearn.metrics import mean_absolute_error, mean_squared_error
from visualize_data import create_forecast_dashboard

# ── NEW: 파라미터
BASE_THRESHOLD  = 10
PROCESS_WEIGHTS = {"AGI":5,"DAS":5,"SHU":0,"MIR":0}
CV_SPLITS       = 5
MODEL_DIR       = pathlib.Path("output/model_cache")
MODEL_DIR.mkdir(exist_ok=True)

# ── NEW: 공통 가중치 회귀변수
def add_site_weight(df):
    df = df.copy()
    df["site_weight"] = df["SITE"].map(PROCESS_WEIGHTS).fillna(0)
    return df

# ── ARIMA → auto_arima
def fit_arima(y):
    return pm.auto_arima(
        y, seasonal=False, information_criterion="aic",
        suppress_warnings=True, error_action='ignore'
    )

# ── Prophet 튜닝
def fit_prophet(df):
    m = Prophet(
        growth="linear", yearly_seasonality="auto",
        weekly_seasonality="auto", changepoint_range=0.9,
        changepoint_prior_scale=0.05
    )
    m.add_country_holidays("AE")
    m.add_regressor("site_weight")
    df = add_site_weight(df)
    m.fit(df.rename(columns={"ds":"ds","y":"y"}))
    return m

# ── NEW: Time-series CV
def cv_score(model_type, df, splits=CV_SPLITS):
    tscv = TimeSeriesSplit(n_splits=splits)
    maes, rmses = [], []
    for tr, ts in tscv.split(df):
        train, test = df.iloc[tr], df.iloc[ts]
        if model_type=="ARIMA":
            mdl = fit_arima(train["y"])
            pred = mdl.predict(n_periods=len(test))
        else:                  # Prophet
            mdl = fit_prophet(train.rename(columns={"ds":"ds","y":"y"}))
            fut  = mdl.make_future_dataframe(len(test))
            fut["site_weight"] = add_site_weight(train).iloc[-1]["site_weight"]
            pred = mdl.predict(fut).iloc[-len(test):]["yhat"].values
        maes.append(mean_absolute_error(test["y"], pred))
        rmses.append(mean_squared_error(test["y"], pred, squared=False))
    return np.mean(maes), np.mean(rmses)

# ── MAIN 변경: 모델 비교·캐싱·시트 기록
# Load data and prepare series
# Use output/logistics_mapping.xlsx STEP_FLOW sheet for forecasting
data_path = pathlib.Path("output/logistics_mapping.xlsx")
df = pd.read_excel(data_path, sheet_name="STEP_FLOW")
series = df[["ATA", "SITE", "전체 리드타임"]].rename(columns={"ATA": "ds", "전체 리드타임": "y"})
series = series.dropna(subset=["ds", "y", "SITE"])  # 결측치 제거
value_col = "y"

arima_mae, arima_rmse = cv_score("ARIMA", series)
prop_mae, prop_rmse   = cv_score("Prophet", series)

best = "Prophet" if prop_mae < arima_mae else "ARIMA"
best_model = fit_prophet(series) if best=="Prophet" else fit_arima(series["y"])
joblib.dump(best_model, MODEL_DIR/f"{best}_{value_col}.pkl")

# 예측 6개월
if best == "ARIMA":
    forecast_values = best_model.predict(180)
    last_date = series["ds"].max()
    future_dates = pd.date_range(last_date, periods=180, freq="D")
    df_fcst = pd.DataFrame({"ds": future_dates, "y": forecast_values})
else:
    future = best_model.make_future_dataframe(180)
    future["site_weight"] = series["site_weight"].iloc[-1] if "site_weight" in series.columns else 0
    forecast_df = best_model.predict(future)
    df_fcst = forecast_df.iloc[-180:][["ds", "yhat"]].rename(columns={"yhat": "y"})

# 최근 1년치 실적 + 예측 시각화
import pandas as pd
recent_hist = series.sort_values("ds").tail(365)
create_forecast_dashboard(recent_hist, df_fcst, "output/forecast_dashboard.html")

# → Excel 시트 'Model_CV' & 'Forecast' 추가 … 