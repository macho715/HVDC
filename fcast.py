import pandas as pd
import numpy as np
from statsmodels.tsa.arima.model import ARIMA
from prophet import Prophet
from sklearn.metrics import mean_absolute_error, mean_squared_error
import xlwings as xw
from datetime import datetime, timedelta
import traceback

def prepare_data(df, date_col, value_col):
    """데이터 전처리 함수"""
    try:
        # 날짜 컬럼을 datetime으로 변환
        df[date_col] = pd.to_datetime(df[date_col])
        # 값 컬럼을 숫자로 변환
        df[value_col] = pd.to_numeric(df[value_col], errors='coerce')
        # 날짜별 평균값 계산
        daily_data = df.groupby(date_col)[value_col].mean().reset_index()
        daily_data.columns = ['ds', 'y']
        # 결측치 처리
        daily_data = daily_data.dropna()
        return daily_data
    except Exception as e:
        print(f"Error in prepare_data: {str(e)}")
        print(f"Available columns: {df.columns.tolist()}")
        raise

def fit_arima(data, order=(1,1,1)):
    """ARIMA 모델 학습"""
    try:
        model = ARIMA(data['y'], order=order)
        results = model.fit()
        return results
    except Exception as e:
        print(f"Error in fit_arima: {str(e)}")
        raise

def fit_prophet(data):
    """Prophet 모델 학습"""
    try:
        model = Prophet(yearly_seasonality=True, 
                       weekly_seasonality=True,
                       daily_seasonality=False)
        model.fit(data)
        return model
    except Exception as e:
        print(f"Error in fit_prophet: {str(e)}")
        raise

def evaluate_models(train_data, test_data, arima_model, prophet_model):
    """모델 성능 평가"""
    try:
        # ARIMA 예측
        arima_pred = arima_model.forecast(steps=len(test_data))
        arima_mae = mean_absolute_error(test_data['y'], arima_pred)
        arima_rmse = np.sqrt(mean_squared_error(test_data['y'], arima_pred))
        
        # Prophet 예측
        future = prophet_model.make_future_dataframe(periods=len(test_data))
        prophet_pred = prophet_model.predict(future)
        prophet_pred = prophet_pred.iloc[-len(test_data):]['yhat']
        prophet_mae = mean_absolute_error(test_data['y'], prophet_pred)
        prophet_rmse = np.sqrt(mean_squared_error(test_data['y'], prophet_pred))
        
        return {
            'arima': {'mae': arima_mae, 'rmse': arima_rmse},
            'prophet': {'mae': prophet_mae, 'rmse': prophet_rmse}
        }
    except Exception as e:
        print(f"Error in evaluate_models: {str(e)}")
        raise

def run_forecast(file_path, sheet=0, date_col="ATA", value_col="전체 리드타임", horizon=3):
    """메인 예측 실행 함수"""
    try:
        print(f"Opening file: {file_path}")
        # Excel 파일 읽기
        wb = xw.Book(file_path)
        print(f"Available sheets: {wb.sheets}")
        
        # 시트 인덱스 또는 이름으로 접근
        ws = wb.sheets[sheet]
        print(f"Reading data from sheet: {ws.name}")
        df = pd.DataFrame(ws.range('A1').expand().value)
        print(f"Data shape: {df.shape}")
        print(f"Columns: {df.columns.tolist()}")
        
        df.columns = df.iloc[0]
        df = df.iloc[1:]
        
        # 데이터 전처리
        print("Preparing data...")
        data = prepare_data(df, date_col, value_col)
        print(f"Prepared data shape: {data.shape}")
        
        # 학습/테스트 데이터 분할
        train_size = int(len(data) * 0.8)
        train_data = data.iloc[:train_size]
        test_data = data.iloc[train_size:]
        
        # 모델 학습
        print("Fitting ARIMA model...")
        arima_model = fit_arima(train_data)
        print("Fitting Prophet model...")
        prophet_model = fit_prophet(train_data)
        
        # 모델 평가
        print("Evaluating models...")
        metrics = evaluate_models(train_data, test_data, arima_model, prophet_model)
        print(f"Model metrics: {metrics}")
        
        # 더 나은 모델 선택
        if metrics['arima']['mae'] < metrics['prophet']['mae']:
            best_model = 'ARIMA'
            print("Selected ARIMA model")
            # ARIMA 예측
            forecast = arima_model.forecast(steps=horizon)
            conf_int = arima_model.get_forecast(steps=horizon).conf_int(alpha=0.2)
        else:
            best_model = 'Prophet'
            print("Selected Prophet model")
            # Prophet 예측
            future = prophet_model.make_future_dataframe(periods=horizon)
            forecast_result = prophet_model.predict(future)
            forecast = forecast_result.iloc[-horizon:]['yhat']
            conf_int = forecast_result.iloc[-horizon:][['yhat_lower', 'yhat_upper']]
        
        # 예측 결과 준비
        last_date = data['ds'].max()
        forecast_dates = [last_date + timedelta(days=30*i) for i in range(1, horizon+1)]
        
        results = pd.DataFrame({
            'Forecast_Month': forecast_dates,
            'Pred_LT': forecast,
            'Low80': conf_int.iloc[:, 0] if best_model == 'ARIMA' else conf_int['yhat_lower'],
            'High80': conf_int.iloc[:, 1] if best_model == 'ARIMA' else conf_int['yhat_upper'],
            'Model': [f"{best_model}(1,1,1)" if best_model == 'ARIMA' else best_model] * horizon
        })
        
        # 결과를 Excel에 저장
        print("Saving results...")
        try:
            wb.sheets.add('DataModel')
        except:
            pass
        
        sheet_dm = wb.sheets['DataModel']
        sheet_dm.clear_contents()
        sheet_dm.range('A1').value = results
        
        # 차트 생성
        print("Creating chart...")
        chart = sheet_dm.charts.add()
        chart.set_source_data(sheet_dm.range('A1').expand())
        chart.chart_type = 'line'
        chart.top = sheet_dm.range('A1').top
        chart.left = sheet_dm.range('A1').left + 400
        
        print("Forecast completed successfully!")
        return True
        
    except Exception as e:
        print(f"Error in run_forecast: {str(e)}")
        print("Traceback:")
        print(traceback.format_exc())
        return False

if __name__ == "__main__":
    # 테스트 실행
    run_forecast("data/HVDC-STATUS.xlsx")