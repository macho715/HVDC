"""
Logistics Mapper Module
=====================

This module provides functionality for mapping and transforming HVDC logistics data.
"""

import pandas as pd
from pathlib import Path
from typing import Dict, Any

from scripts.core.base import HVDCBase
from scripts.core.utils import load_excel, format_date, parse_date

class HVDCLogisticsMapper(HVDCBase):
    """
    HVDC 물류 데이터를 매핑하고 기본적인 변환을 수행하는 클래스입니다.
    """
    def __init__(self, input_filename: str = "HVDC-STATUS.xlsx",
                 max_no_filter: int = 560):
        """
        HVDCLogisticsMapper를 초기화합니다.

        Args:
            input_filename (str): self.data_dir 내에서 읽어올 입력 Excel 파일 이름입니다.
            max_no_filter (int): 'NO.' 컬럼 기준으로 필터링할 최대값입니다.
        """
        super().__init__()
        self.input_filepath: Path = self.data_dir / input_filename
        self.max_no_filter: int = max_no_filter
        self.logger.info(f"HVDCLogisticsMapper 초기화 완료. 입력 파일: {self.input_filepath}")

    def _load_data(self) -> pd.DataFrame:
        """입력 Excel 파일을 로드합니다."""
        if not self.input_filepath.exists():
            self.logger.error(f"입력 파일을 찾을 수 없습니다: {self.input_filepath}")
            raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {self.input_filepath}")
        
        try:
            df = load_excel(self.input_filepath)
            self.logger.info(f"원본 데이터 로드 완료: {self.input_filepath} ({len(df)} 행)")
            return df
        except Exception as e:
            self.logger.error(f"Excel 파일 로드 중 오류 발생: {self.input_filepath}, 오류: {e}")
            raise

    def _calculate_lead_time(self, df: pd.DataFrame) -> pd.DataFrame:
        """리드타임을 계산합니다."""
        self.logger.debug("리드타임 계산 시작...")
        df["ATA_dt"] = pd.to_datetime(df["ATA"], errors='coerce')
        df["MOSB_dt"] = pd.to_datetime(df["MOSB"], errors='coerce')
        df["리드타임(일)"] = (df["MOSB_dt"] - df["ATA_dt"]).dt.days
        self.logger.debug("리드타임 계산 완료.")
        return df

    def _classify_hvdc_step(self, df: pd.DataFrame) -> pd.DataFrame:
        """SUB DESCRIPTION을 기반으로 HVDC 공정 단계를 분류합니다."""
        self.logger.debug("HVDC 공정 단계 분류 시작...")
        
        def updated_hvdc_step_logic(desc: Any) -> int:
            if pd.isna(desc):
                return 99
            desc_str = str(desc).lower()
            if any(k in desc_str for k in ["converter transformer", "valve", "thyristor", "igbt"]):
                return 1
            elif any(k in desc_str for k in ["dc cable", "submarine", "overhead", "transmission"]):
                return 2
            elif any(k in desc_str for k in ["filter", "reactor", "capacitor", "harmonic"]):
                return 3
            elif any(k in desc_str for k in ["scada", "control", "protection", "monitoring"]):
                return 4
            elif any(k in desc_str for k in ["grounding", "electrode", "earth"]):
                return 5
            elif any(k in desc_str for k in ["spare", "repair"]):
                return 6
            else:
                return 99

        df["공정단계_HVDC"] = df["SUB DESCRIPTION"].apply(updated_hvdc_step_logic)
        
        step_labels: Dict[int, str] = {
            1: "Converter", 2: "Transmission", 3: "Filter/Reactor",
            4: "Control/Protection", 5: "Grounding", 6: "Spare/Maintenance",
            99: "기타"
        }
        df["공정단계_HVDC_Label"] = df["공정단계_HVDC"].map(step_labels)
        self.logger.debug("HVDC 공정 단계 분류 완료.")
        return df

    def _classify_lead_time_status(self, df: pd.DataFrame) -> pd.DataFrame:
        """리드타임(일)을 기반으로 리드타임 상태를 분류합니다."""
        self.logger.debug("리드타임 상태 분류 시작...")
        
        def classify_status(lead_time: float) -> str:
            if pd.isna(lead_time):
                return "미도착"
            elif lead_time <= 30:
                return "양호"
            elif lead_time <= 60:
                return "주의"
            else:
                return "지연"
        
        df["리드타임 상태"] = df["리드타임(일)"].apply(classify_status)
        self.logger.debug("리드타임 상태 분류 완료.")
        return df

    def _add_risk_level(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        리드타임 상태를 기반으로 '위험도' 컬럼을 생성합니다.
        """
        self.logger.debug("위험도 컬럼 생성 시작...")
        if "리드타임 상태" not in df.columns:
            self.logger.warning("'리드타임 상태' 컬럼이 없어 '위험도'를 'N/A'로 설정합니다.")
            df["위험도"] = "N/A"
            return df

        def assign_risk(status: str) -> str:
            if status == "지연":
                return "High"
            elif status == "주의":
                return "Medium"
            elif status in ["양호", "미도착"]:
                return "Low"
            else:
                return "N/A"

        df["위험도"] = df["리드타임 상태"].apply(assign_risk)
        self.logger.debug("위험도 컬럼 생성 완료.")
        return df

    def _select_and_filter_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """최종 컬럼 선택, 정렬 및 필터링을 수행합니다."""
        self.logger.debug("최종 컬럼 선택, 정렬 및 필터링 시작...")
        final_columns = [
            "NO.", "VENDOR", "MAIN DESCRIPTION (PO)", "SUB DESCRIPTION",
            "공정단계_HVDC_Label", "리드타임(일)", "리드타임 상태", "위험도",
            "INCOTERMS", "DG 분류", "섬 운송 여부", "ATA", "MOSB"
        ]
        
        # 누락된 최종 컬럼이 있다면 NA 값으로 추가 (오류 방지)
        for col in final_columns:
            if col not in df.columns:
                df[col] = pd.NA
                self.logger.warning(f"최종 컬럼 '{col}'이(가) 없어 NA로 추가합니다.")
        
        # NO. 컬럼이 숫자가 아닐 경우를 대비한 변환
        df["NO."] = pd.to_numeric(df["NO."], errors='coerce')
        
        # 필터링 및 정렬
        final_df = df[df["NO."] <= self.max_no_filter][final_columns].sort_values("NO.").copy()
        self.logger.debug("최종 컬럼 선택, 정렬 및 필터링 완료.")
        return final_df

    def process_data(self) -> pd.DataFrame:
        """
        전체 데이터 처리 파이프라인을 실행하고 처리된 DataFrame을 반환합니다.
        """
        self.logger.info("물류 데이터 매핑 처리 시작...")
        
        raw_df = self._load_data()
        df_with_lead_time = self._calculate_lead_time(raw_df)
        df_with_steps = self._classify_hvdc_step(df_with_lead_time)
        df_with_status = self._classify_lead_time_status(df_with_steps)
        df_with_risk = self._add_risk_level(df_with_status)
        final_df = self._select_and_filter_columns(df_with_risk)
        
        self.logger.info(f"최종 매핑 데이터 생성 완료: {len(final_df)} 행")
        return final_df

if __name__ == '__main__':
    print("HVDCLogisticsMapper 테스트 시작...")
    try:
        mapper = HVDCLogisticsMapper(input_filename="HVDC-STATUS-Sample-Test.xlsx", max_no_filter=50)
        processed_dataframe = mapper.process_data()
        
        print("\n데이터 처리 성공!")
        print(f"처리된 데이터 총 행 수: {len(processed_dataframe)}")
        if not processed_dataframe.empty:
            print(f"처리된 데이터 첫 5행:\n{processed_dataframe.head()}")
            print(f"\n처리된 데이터 컬럼:\n{processed_dataframe.columns.tolist()}")
            if "공정단계_HVDC_Label" in processed_dataframe.columns:
                print(f"\n공정단계별 건수:\n{processed_dataframe['공정단계_HVDC_Label'].value_counts()}")
            if "리드타임 상태" in processed_dataframe.columns:
                print(f"\n리드타임 상태별 건수:\n{processed_dataframe['리드타임 상태'].value_counts()}")
        else:
            print("처리된 데이터가 비어있습니다 (필터 조건 등을 확인하세요).")

    except FileNotFoundError:
        print(f"테스트 오류: 입력 파일을 찾을 수 없습니다.")
        print(f"팁: scripts.data.generator 모듈을 실행하여 '{mapper.input_filepath.name}' 샘플 파일을 먼저 생성하세요.")
    except Exception as e:
        print(f"테스트 중 예상치 못한 오류 발생: {e}")
        import traceback
        traceback.print_exc()
    print("\nHVDCLogisticsMapper 테스트 종료.") 