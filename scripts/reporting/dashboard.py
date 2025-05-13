"""
Dashboard Generator Module
=======================

This module provides functionality for generating HVDC data analysis dashboards.
"""

import pandas as pd
import numpy as np
from pathlib import Path
from typing import Dict, List, Optional, Union
from datetime import datetime

import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from scripts.core.base import HVDCBase
from scripts.core.utils import save_excel, save_json

import xlsxwriter

class HVDCDashboardGenerator(HVDCBase):
    """
    HVDC 데이터 분석 대시보드를 생성하는 클래스입니다.
    """
    def __init__(self):
        """HVDCDashboardGenerator를 초기화합니다."""
        super().__init__()
        self.logger.info("HVDCDashboardGenerator 초기화 완료")

    def _create_lead_time_analysis(self, df: pd.DataFrame) -> Dict[str, go.Figure]:
        """
        리드타임 분석 차트를 생성합니다.

        Args:
            df (pd.DataFrame): 분석할 데이터프레임

        Returns:
            Dict[str, go.Figure]: 리드타임 분석 차트들
        """
        charts = {}
        
        # 1. 리드타임 분포 히스토그램
        fig_hist = px.histogram(
            df, 
            x='리드타임(일)',
            nbins=30,
            title='리드타임 분포',
            labels={'리드타임(일)': '리드타임 (일)', 'count': '건수'},
            color_discrete_sequence=['#1f77b4']
        )
        fig_hist.update_layout(
            showlegend=False,
            xaxis_title='리드타임 (일)',
            yaxis_title='건수'
        )
        charts['lead_time_histogram'] = fig_hist

        # 2. 공정단계별 리드타임 박스플롯
        fig_box = px.box(
            df,
            x='공정단계_HVDC_Label',
            y='리드타임(일)',
            title='공정단계별 리드타임 분포',
            labels={
                '공정단계_HVDC_Label': '공정단계',
                '리드타임(일)': '리드타임 (일)'
            }
        )
        fig_box.update_layout(
            xaxis_title='공정단계',
            yaxis_title='리드타임 (일)',
            xaxis={'tickangle': 45}
        )
        charts['lead_time_by_process'] = fig_box

        # 3. 리드타임 상태별 파이 차트
        status_counts = df['리드타임 상태'].value_counts()
        fig_pie = px.pie(
            values=status_counts.values,
            names=status_counts.index,
            title='리드타임 상태 분포',
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        charts['lead_time_status'] = fig_pie

        return charts

    def _create_process_analysis(self, df: pd.DataFrame) -> Dict[str, go.Figure]:
        """
        공정단계 분석 차트를 생성합니다.

        Args:
            df (pd.DataFrame): 분석할 데이터프레임

        Returns:
            Dict[str, go.Figure]: 공정단계 분석 차트들
        """
        charts = {}
        
        # 1. 공정단계별 건수 바 차트
        process_counts = df['공정단계_HVDC_Label'].value_counts()
        fig_bar = px.bar(
            x=process_counts.index,
            y=process_counts.values,
            title='공정단계별 건수',
            labels={'x': '공정단계', 'y': '건수'},
            color_discrete_sequence=['#2ca02c']
        )
        fig_bar.update_layout(
            xaxis_title='공정단계',
            yaxis_title='건수',
            xaxis={'tickangle': 45}
        )
        charts['process_counts'] = fig_bar

        # 2. 공정단계별 위험도 분포
        risk_by_process = pd.crosstab(
            df['공정단계_HVDC_Label'],
            df['위험도']
        )
        fig_stacked = px.bar(
            risk_by_process,
            title='공정단계별 위험도 분포',
            labels={'value': '건수', '위험도': '위험도'},
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        fig_stacked.update_layout(
            xaxis_title='공정단계',
            yaxis_title='건수',
            xaxis={'tickangle': 45}
        )
        charts['risk_by_process'] = fig_stacked

        return charts

    def _create_vendor_analysis(self, df: pd.DataFrame) -> Dict[str, go.Figure]:
        """
        벤더 분석 차트를 생성합니다.

        Args:
            df (pd.DataFrame): 분석할 데이터프레임

        Returns:
            Dict[str, go.Figure]: 벤더 분석 차트들
        """
        charts = {}
        
        # 1. 벤더별 건수 바 차트
        vendor_counts = df['VENDOR'].value_counts().head(10)  # 상위 10개 벤더만 표시
        fig_bar = px.bar(
            x=vendor_counts.index,
            y=vendor_counts.values,
            title='벤더별 건수 (상위 10개)',
            labels={'x': '벤더', 'y': '건수'},
            color_discrete_sequence=['#ff7f0e']
        )
        fig_bar.update_layout(
            xaxis_title='벤더',
            yaxis_title='건수',
            xaxis={'tickangle': 45}
        )
        charts['vendor_counts'] = fig_bar

        # 2. 벤더별 평균 리드타임
        vendor_lead_time = df.groupby('VENDOR')['리드타임(일)'].mean().sort_values(ascending=False).head(10)
        fig_lead_time = px.bar(
            x=vendor_lead_time.index,
            y=vendor_lead_time.values,
            title='벤더별 평균 리드타임 (상위 10개)',
            labels={'x': '벤더', 'y': '평균 리드타임 (일)'},
            color_discrete_sequence=['#d62728']
        )
        fig_lead_time.update_layout(
            xaxis_title='벤더',
            yaxis_title='평균 리드타임 (일)',
            xaxis={'tickangle': 45}
        )
        charts['vendor_lead_time'] = fig_lead_time

        return charts

    def _create_risk_analysis(self, df: pd.DataFrame) -> Dict[str, go.Figure]:
        """
        위험도 분석 차트를 생성합니다.

        Args:
            df (pd.DataFrame): 분석할 데이터프레임

        Returns:
            Dict[str, go.Figure]: 위험도 분석 차트들
        """
        charts = {}
        
        # 1. 위험도 분포 파이 차트
        risk_counts = df['위험도'].value_counts()
        fig_pie = px.pie(
            values=risk_counts.values,
            names=risk_counts.index,
            title='위험도 분포',
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        charts['risk_distribution'] = fig_pie

        # 2. 위험도별 평균 리드타임
        risk_lead_time = df.groupby('위험도')['리드타임(일)'].mean()
        fig_lead_time = px.bar(
            x=risk_lead_time.index,
            y=risk_lead_time.values,
            title='위험도별 평균 리드타임',
            labels={'x': '위험도', 'y': '평균 리드타임 (일)'},
            color_discrete_sequence=['#9467bd']
        )
        fig_lead_time.update_layout(
            xaxis_title='위험도',
            yaxis_title='평균 리드타임 (일)'
        )
        charts['risk_lead_time'] = fig_lead_time

        return charts

    def _create_summary_statistics(self, df: pd.DataFrame) -> Dict[str, float]:
        """
        데이터의 주요 통계량을 계산합니다.

        Args:
            df (pd.DataFrame): 분석할 데이터프레임

        Returns:
            Dict[str, float]: 주요 통계량
        """
        stats = {
            'total_items': len(df),
            'avg_lead_time': df['리드타임(일)'].mean(),
            'max_lead_time': df['리드타임(일)'].max(),
            'min_lead_time': df['리드타임(일)'].min(),
            'std_lead_time': df['리드타임(일)'].std(),
            'high_risk_items': len(df[df['위험도'] == 'High']),
            'medium_risk_items': len(df[df['위험도'] == 'Medium']),
            'low_risk_items': len(df[df['위험도'] == 'Low']),
            'unique_vendors': df['VENDOR'].nunique(),
            'unique_processes': df['공정단계_HVDC_Label'].nunique()
        }
        return stats

    def _prepare_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        대시보드 생성을 위한 데이터 전처리를 수행합니다.

        Args:
            df (pd.DataFrame): 원본 데이터프레임

        Returns:
            pd.DataFrame: 전처리된 데이터프레임
        """
        # 데이터 복사
        processed_df = df.copy()
        
        # 필요한 컬럼이 없는 경우 기본값 설정
        required_columns = ['NO.', 'VENDOR', '공정단계_HVDC_Label', '리드타임(일)', '위험도']
        for col in required_columns:
            if col not in processed_df.columns:
                if col == '리드타임(일)':
                    processed_df[col] = 0
                elif col == '위험도':
                    processed_df[col] = 'N/A'
                else:
                    processed_df[col] = ''
        
        # 숫자형 컬럼 변환
        numeric_columns = ['NO.', '리드타임(일)']
        for col in numeric_columns:
            if col in processed_df.columns:
                processed_df[col] = pd.to_numeric(processed_df[col], errors='coerce')
        
        return processed_df

    def create_dashboard(self, filename: str, source_data: pd.DataFrame) -> Path:
        """
        HVDC 데이터 분석 대시보드를 생성합니다.

        Args:
            filename (str): 생성할 대시보드 파일명
            source_data (pd.DataFrame): 분석할 데이터프레임

        Returns:
            Path: 생성된 대시보드 파일 경로
        """
        self.logger.info("대시보드 생성 시작...")
        df = self._prepare_data(source_data)
        stats = self._create_summary_statistics(df)
        lead_time_charts = self._create_lead_time_analysis(df)
        process_charts = self._create_process_analysis(df)
        vendor_charts = self._create_vendor_analysis(df)
        risk_charts = self._create_risk_analysis(df)

        output_filepath = self.output_dir / filename
        self.logger.info(f"Excel 대시보드 파일 생성 시작: {output_filepath}")
        with pd.ExcelWriter(output_filepath, engine='xlsxwriter') as writer:
            self.logger.info("ExcelWriter 객체가 생성되었습니다.")
            workbook = writer.book
            self.logger.info("Workbook 객체가 성공적으로 초기화되었습니다.")

            # 데이터 시트 (숨김)
            self.logger.info("데이터 시트 생성 시작...")
            df.to_excel(writer, sheet_name='Data', index=False)
            worksheet = writer.sheets['Data']
            worksheet.hide()
            self.logger.info("데이터 시트가 생성되고 숨겨졌습니다.")

            # 대시보드 시트
            self.logger.info("대시보드 시트 생성 시작...")
            dashboard_sheet = workbook.add_worksheet('대시보드')
            self.logger.info("대시보드 시트가 생성되었습니다.")

            # 통계 정보 추가
            self.logger.info("통계 정보 추가 시작...")
            row = 0
            dashboard_sheet.write(row, 0, "총 아이템 수")
            dashboard_sheet.write(row, 1, stats['total_items'])
            row += 1
            dashboard_sheet.write(row, 0, "평균 리드타임")
            dashboard_sheet.write(row, 1, f"{stats['avg_lead_time']:.1f}일")
            row += 1
            dashboard_sheet.write(row, 0, "고위험 아이템")
            dashboard_sheet.write(row, 1, stats['high_risk_items'])
            row += 2
            self.logger.info("통계 정보가 추가되었습니다.")

            # 차트 추가
            self.logger.info("차트 추가 시작...")
            for section, charts in [
                ("리드타임 분석", lead_time_charts),
                ("공정단계 분석", process_charts),
                ("벤더 분석", vendor_charts),
                ("위험도 분석", risk_charts)
            ]:
                self.logger.info(f"{section} 차트 추가 중...")
                dashboard_sheet.write(row, 0, section)
                row += 1
                for name, fig in charts.items():
                    try:
                        self.logger.info(f"{name} 차트 이미지 생성 중...")
                        chart_path = self.output_dir / f"{name}.png"
                        fig.write_image(chart_path)
                        self.logger.info(f"{name} 차트 이미지 생성 완료: {chart_path}")
                        dashboard_sheet.insert_image(row, 0, str(chart_path))
                        row += 20  # 차트 간 간격
                        self.logger.info(f"{name} 차트가 추가되었습니다.")
                    except Exception as e:
                        self.logger.error(f"{name} 차트 생성/삽입 중 오류 발생: {e}")
            self.logger.info("모든 차트가 추가되었습니다.")

        self.logger.info(f"대시보드 생성 완료: {output_filepath}")
        return output_filepath

if __name__ == '__main__':
    print("HVDCDashboardGenerator 테스트 시작...")
    
    # 테스트용 DataFrame 생성
    test_data = {
        'NO.': range(1, 101),
        'VENDOR': [f'Vendor_{i%10+1}' for i in range(100)],
        '공정단계_HVDC_Label': np.random.choice(
            ['Converter', 'Transmission', 'Filter/Reactor', 'Control/Protection', 'Grounding', 'Spare/Maintenance', '기타'],
            size=100
        ),
        '리드타임(일)': np.random.normal(30, 10, 100).clip(0),
        '위험도': np.random.choice(['High', 'Medium', 'Low'], size=100, p=[0.2, 0.5, 0.3])
    }
    test_df = pd.DataFrame(test_data)

    # 대시보드 생성
    generator = HVDCDashboardGenerator()
    result = generator.create_dashboard("test_dashboard.xlsx", test_df)
    
    print("\n생성된 파일:")
    print(f"대시보드: {result}")
    
    print("\nHVDCDashboardGenerator 테스트 종료.") 