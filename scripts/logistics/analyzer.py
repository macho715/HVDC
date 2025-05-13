"""
HVDC 물류 데이터 분석 모듈입니다.
"""

from pathlib import Path
from typing import Dict, Any, Optional, List
import pandas as pd
import numpy as np

from scripts.core.base import HVDCBase

class HVDCLogisticsAnalyzer(HVDCBase):
    """
    HVDC 물류 데이터 분석 클래스입니다.
    """
    def __init__(self):
        """HVDCLogisticsAnalyzer를 초기화합니다."""
        super().__init__()
        self.logger.info("HVDCLogisticsAnalyzer 초기화 완료.")

    def analyze_lead_time(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        리드타임 분석을 수행합니다.

        Args:
            df (pd.DataFrame): 분석할 데이터프레임

        Returns:
            Dict[str, Any]: 리드타임 분석 결과
        """
        self.logger.info("리드타임 분석 시작")
        
        try:
            # 기본 통계량 계산
            lead_time_stats = {
                "mean": df["리드타임(일)"].mean(),
                "median": df["리드타임(일)"].median(),
                "std": df["리드타임(일)"].std(),
                "min": df["리드타임(일)"].min(),
                "max": df["리드타임(일)"].max()
            }

            # 상태별 통계
            status_stats = df.groupby("리드타임 상태")["리드타임(일)"].agg([
                "count", "mean", "median", "std"
            ]).to_dict()

            # 공정단계별 통계
            process_stats = df.groupby("공정단계_HVDC_Label")["리드타임(일)"].agg([
                "count", "mean", "median", "std"
            ]).to_dict()

            self.logger.info("리드타임 분석 완료")
            return {
                "lead_time_stats": lead_time_stats,
                "status_stats": status_stats,
                "process_stats": process_stats
            }

        except Exception as e:
            self.logger.error(f"리드타임 분석 중 오류 발생: {e}")
            return {}

    def analyze_vendor_performance(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        벤더별 성과 분석을 수행합니다.

        Args:
            df (pd.DataFrame): 분석할 데이터프레임

        Returns:
            Dict[str, Any]: 벤더별 성과 분석 결과
        """
        self.logger.info("벤더별 성과 분석 시작")
        
        try:
            # 벤더별 기본 통계
            vendor_stats = df.groupby("VENDOR").agg({
                "리드타임(일)": ["count", "mean", "median", "std"],
                "NO.": "count"
            }).to_dict()

            # 벤더별 상태 분포
            vendor_status = pd.crosstab(
                df["VENDOR"], 
                df["리드타임 상태"]
            ).to_dict()

            self.logger.info("벤더별 성과 분석 완료")
            return {
                "vendor_stats": vendor_stats,
                "vendor_status": vendor_status
            }

        except Exception as e:
            self.logger.error(f"벤더별 성과 분석 중 오류 발생: {e}")
            return {}

    def generate_summary_report(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        전체 데이터에 대한 요약 보고서를 생성합니다.

        Args:
            df (pd.DataFrame): 분석할 데이터프레임

        Returns:
            Dict[str, Any]: 요약 보고서 데이터
        """
        self.logger.info("요약 보고서 생성 시작")
        
        try:
            # 기본 통계
            summary = {
                "total_items": len(df),
                "unique_vendors": df["VENDOR"].nunique(),
                "unique_processes": df["공정단계_HVDC_Label"].nunique(),
                "date_range": {
                    "ata_min": df["ATA"].min(),
                    "ata_max": df["ATA"].max(),
                    "mosb_min": df["MOSB"].min(),
                    "mosb_max": df["MOSB"].max()
                }
            }

            # 상태 분포
            status_distribution = df["리드타임 상태"].value_counts().to_dict()
            summary["status_distribution"] = status_distribution

            # 공정단계 분포
            process_distribution = df["공정단계_HVDC_Label"].value_counts().to_dict()
            summary["process_distribution"] = process_distribution

            self.logger.info("요약 보고서 생성 완료")
            return summary

        except Exception as e:
            self.logger.error(f"요약 보고서 생성 중 오류 발생: {e}")
            return {}

# 테스트 코드
if __name__ == "__main__":
    # 테스트용 데이터 생성
    test_data = {
        "NO.": range(1, 11),
        "VENDOR": ["Vendor A"] * 5 + ["Vendor B"] * 5,
        "ATA": pd.date_range(start="2024-01-01", periods=10),
        "MOSB": pd.date_range(start="2024-02-01", periods=10),
        "리드타임(일)": np.random.randint(10, 50, 10),
        "공정단계_HVDC_Label": ["Converter"] * 3 + ["Transmission"] * 3 + ["Filter/Reactor"] * 4,
        "리드타임 상태": ["양호"] * 4 + ["주의"] * 3 + ["지연"] * 3
    }
    test_df = pd.DataFrame(test_data)

    # 분석기 인스턴스 생성
    analyzer = HVDCLogisticsAnalyzer()

    # 각 분석 메소드 테스트
    lead_time_analysis = analyzer.analyze_lead_time(test_df)
    vendor_analysis = analyzer.analyze_vendor_performance(test_df)
    summary_report = analyzer.generate_summary_report(test_df)

    # 결과 출력
    print("\n=== 리드타임 분석 결과 ===")
    print(lead_time_analysis)
    
    print("\n=== 벤더별 성과 분석 결과 ===")
    print(vendor_analysis)
    
    print("\n=== 요약 보고서 ===")
    print(summary_report) 