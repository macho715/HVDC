"""
Data Validator Module
===================

This module provides functionality for validating HVDC data quality.
"""

import pandas as pd
import numpy as np
from datetime import datetime
from typing import List, Dict, Any, Union, Optional
from pathlib import Path

from scripts.core.base import HVDCBase
from scripts.core.utils import parse_date, format_date

class HVDCQualityChecker(HVDCBase):
    """
    HVDC 데이터의 품질을 검증하는 클래스입니다.
    """
    def __init__(self):
        """HVDCQualityChecker를 초기화합니다."""
        super().__init__()
        self.validation_results: List[Dict[str, Any]] = []
        self.logger.info("HVDCQualityChecker 초기화 완료")

    def _add_result(self, check_name: str, status: str, message: str, details: Any = None):
        """
        검증 결과를 리스트에 추가합니다.

        Args:
            check_name (str): 검증 항목 이름
            status (str): 검증 상태 ("PASS", "FAIL", "WARN")
            message (str): 검증 결과 메시지
            details (Any, optional): 추가 상세 정보
        """
        self.validation_results.append({
            "check_name": check_name,
            "status": status,
            "message": message,
            "details": details
        })
        if status == "FAIL":
            self.logger.error(message)
        elif status == "WARN":
            self.logger.warning(message)
        else:
            self.logger.info(message)

    def check_required_columns(self, df: pd.DataFrame, required_columns: List[str]) -> bool:
        """
        필수 컬럼의 존재 여부를 확인합니다.

        Args:
            df (pd.DataFrame): 검증할 데이터프레임
            required_columns (List[str]): 필수 컬럼 목록

        Returns:
            bool: 모든 필수 컬럼이 존재하면 True, 아니면 False
        """
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            message = f"필수 컬럼 누락: {', '.join(missing_columns)}"
            self._add_result("필수 컬럼 확인", "FAIL", message, {"missing_columns": missing_columns})
            return False
        
        message = "모든 필수 컬럼이 존재합니다."
        self._add_result("필수 컬럼 확인", "PASS", message)
        return True

    def check_missing_values(self, df: pd.DataFrame, column_name: str, threshold_percent: float = 0.0) -> bool:
        """
        특정 컬럼의 결측치 비율을 확인합니다.

        Args:
            df (pd.DataFrame): 검증할 데이터프레임
            column_name (str): 검증할 컬럼 이름
            threshold_percent (float): 허용 가능한 결측치 비율 (%)

        Returns:
            bool: 결측치 비율이 임계값 이하면 True, 아니면 False
        """
        if column_name not in df.columns:
            message = f"'{column_name}' 컬럼이 존재하지 않아 결측치 검증을 건너뜁니다."
            self._add_result("결측치 확인", "WARN", message, {"column": column_name})
            return True

        missing_count = df[column_name].isnull().sum()
        total_count = len(df)
        missing_percent = (missing_count / total_count) * 100 if total_count > 0 else 0
        
        status = "PASS"
        message = f"'{column_name}' 컬럼 결측치: {missing_count}개 ({missing_percent:.2f}%)"
        
        if missing_percent > threshold_percent:
            status = "FAIL" if threshold_percent == 0.0 else "WARN"
            message += f" - 허용 임계치({threshold_percent}%) 초과"
        
        self._add_result("결측치 확인", status, message, {
            "column": column_name,
            "missing_count": missing_count,
            "missing_percent": missing_percent,
            "threshold": threshold_percent
        })
        
        return status == "PASS" or (status == "WARN" and threshold_percent > 0.0)

    def check_data_types(self, df: pd.DataFrame, column_types: Dict[str, List[type]]) -> bool:
        """
        정의된 컬럼들의 데이터 타입을 확인합니다.

        Args:
            df (pd.DataFrame): 검증할 데이터프레임
            column_types (Dict[str, List[type]]): 컬럼별 예상 데이터 타입 목록

        Returns:
            bool: 모든 컬럼의 데이터 타입이 예상과 일치하면 True, 아니면 False
        """
        all_passed = True
        for col, expected_types in column_types.items():
            if col not in df.columns:
                message = f"'{col}' 컬럼이 존재하지 않아 타입 검증을 건너뜁니다."
                self._add_result("데이터 타입 확인", "WARN", message, {"column": col})
                continue

            actual_type = df[col].dtype
            # 실제 타입이 예상 타입 목록 중 하나라도 만족하는지 확인
            if not any(isinstance(df[col].iloc[0] if len(df[col]) > 0 and not pd.isna(df[col].iloc[0]) else None, t) for t in expected_types) and \
               not any(pd.api.types.is_dtype_equal(actual_type, t) for t in expected_types):
                message = f"'{col}' 컬럼 타입 불일치. 예상: {expected_types}, 실제: {actual_type}"
                self._add_result("데이터 타입 확인", "FAIL", message, {
                    "column": col,
                    "expected": str(expected_types),
                    "actual": str(actual_type)
                })
                all_passed = False
            else:
                message = f"'{col}' 컬럼 타입 일치: {actual_type}"
                self._add_result("데이터 타입 확인", "PASS", message, {"column": col})
        
        return all_passed

    def check_value_ranges(self, df: pd.DataFrame, range_checks: Dict[str, Dict[str, float]]) -> bool:
        """
        숫자형 컬럼의 값 범위를 확인합니다.

        Args:
            df (pd.DataFrame): 검증할 데이터프레임
            range_checks (Dict[str, Dict[str, float]]): 컬럼별 최소/최대값 설정

        Returns:
            bool: 모든 값이 범위 내에 있으면 True, 아니면 False
        """
        all_passed = True
        for col, ranges in range_checks.items():
            if col not in df.columns:
                message = f"'{col}' 컬럼이 존재하지 않아 범위 검증을 건너뜁니다."
                self._add_result("값 범위 확인", "WARN", message, {"column": col})
                continue

            if not pd.api.types.is_numeric_dtype(df[col]):
                message = f"'{col}' 컬럼이 숫자형이 아닙니다."
                self._add_result("값 범위 확인", "FAIL", message, {"column": col})
                all_passed = False
                continue

            min_val = ranges.get("min", float("-inf"))
            max_val = ranges.get("max", float("inf"))
            
            out_of_range = df[col][~df[col].isna()].apply(lambda x: x < min_val or x > max_val)
            out_of_range_count = out_of_range.sum()
            
            if out_of_range_count > 0:
                message = f"'{col}' 컬럼의 {out_of_range_count}개 값이 범위({min_val} ~ {max_val})를 벗어납니다."
                self._add_result("값 범위 확인", "FAIL", message, {
                    "column": col,
                    "min": min_val,
                    "max": max_val,
                    "out_of_range_count": out_of_range_count
                })
                all_passed = False
            else:
                message = f"'{col}' 컬럼의 모든 값이 범위({min_val} ~ {max_val}) 내에 있습니다."
                self._add_result("값 범위 확인", "PASS", message, {"column": col})
        
        return all_passed

    def check_unique_values(self, df: pd.DataFrame, unique_columns: List[str]) -> bool:
        """
        특정 컬럼의 값이 고유한지 확인합니다.

        Args:
            df (pd.DataFrame): 검증할 데이터프레임
            unique_columns (List[str]): 고유해야 하는 컬럼 목록

        Returns:
            bool: 모든 컬럼이 고유하면 True, 아니면 False
        """
        all_passed = True
        for col in unique_columns:
            if col not in df.columns:
                message = f"'{col}' 컬럼이 존재하지 않아 고유값 검증을 건너뜁니다."
                self._add_result("고유값 확인", "WARN", message, {"column": col})
                continue

            duplicates = df[col].duplicated()
            duplicate_count = duplicates.sum()
            
            if duplicate_count > 0:
                message = f"'{col}' 컬럼에 {duplicate_count}개의 중복값이 있습니다."
                self._add_result("고유값 확인", "FAIL", message, {
                    "column": col,
                    "duplicate_count": duplicate_count,
                    "duplicate_values": df[col][duplicates].unique().tolist()
                })
                all_passed = False
            else:
                message = f"'{col}' 컬럼의 모든 값이 고유합니다."
                self._add_result("고유값 확인", "PASS", message, {"column": col})
        
        return all_passed

    def check_date_sequence(self, df: pd.DataFrame, date_columns: Dict[str, str]) -> bool:
        """
        날짜 컬럼들의 순서를 확인합니다.

        Args:
            df (pd.DataFrame): 검증할 데이터프레임
            date_columns (Dict[str, str]): {이전_날짜_컬럼: 이후_날짜_컬럼} 형태의 딕셔너리

        Returns:
            bool: 모든 날짜 순서가 올바르면 True, 아니면 False
        """
        all_passed = True
        for before_col, after_col in date_columns.items():
            if before_col not in df.columns or after_col not in df.columns:
                message = f"날짜 컬럼 '{before_col}' 또는 '{after_col}'이 존재하지 않아 순서 검증을 건너뜁니다."
                self._add_result("날짜 순서 확인", "WARN", message, {
                    "before_column": before_col,
                    "after_column": after_col
                })
                continue

            # 날짜 컬럼을 datetime으로 변환
            before_dates = pd.to_datetime(df[before_col], errors='coerce')
            after_dates = pd.to_datetime(df[after_col], errors='coerce')
            
            # 둘 다 유효한 날짜인 경우만 검사
            valid_dates = ~before_dates.isna() & ~after_dates.isna()
            invalid_sequence = before_dates[valid_dates] > after_dates[valid_dates]
            
            if invalid_sequence.any():
                invalid_count = invalid_sequence.sum()
                message = f"'{before_col}'이 '{after_col}'보다 늦은 경우가 {invalid_count}건 있습니다."
                self._add_result("날짜 순서 확인", "FAIL", message, {
                    "before_column": before_col,
                    "after_column": after_col,
                    "invalid_count": invalid_count
                })
                all_passed = False
            else:
                message = f"'{before_col}'과 '{after_col}'의 날짜 순서가 올바릅니다."
                self._add_result("날짜 순서 확인", "PASS", message, {
                    "before_column": before_col,
                    "after_column": after_col
                })
        
        return all_passed

    def check_categorical_values(self, df: pd.DataFrame, category_checks: Dict[str, List[str]]) -> bool:
        """
        카테고리형 컬럼의 값이 허용된 목록에 포함되는지 확인합니다.

        Args:
            df (pd.DataFrame): 검증할 데이터프레임
            category_checks (Dict[str, List[str]]): 컬럼별 허용된 값 목록

        Returns:
            bool: 모든 값이 허용된 목록에 포함되면 True, 아니면 False
        """
        all_passed = True
        for col, allowed_values in category_checks.items():
            if col not in df.columns:
                message = f"'{col}' 컬럼이 존재하지 않아 카테고리 값 검증을 건너뜁니다."
                self._add_result("카테고리 값 확인", "WARN", message, {"column": col})
                continue

            invalid_values = df[col][~df[col].isna()].apply(lambda x: x not in allowed_values)
            invalid_count = invalid_values.sum()
            
            if invalid_count > 0:
                message = f"'{col}' 컬럼에 {invalid_count}개의 허용되지 않은 값이 있습니다."
                self._add_result("카테고리 값 확인", "FAIL", message, {
                    "column": col,
                    "invalid_count": invalid_count,
                    "invalid_values": df[col][invalid_values].unique().tolist(),
                    "allowed_values": allowed_values
                })
                all_passed = False
            else:
                message = f"'{col}' 컬럼의 모든 값이 허용된 목록에 포함됩니다."
                self._add_result("카테고리 값 확인", "PASS", message, {"column": col})
        
        return all_passed

    def validate_dataframe(self, df: pd.DataFrame, config: Dict[str, Any]) -> pd.DataFrame:
        """
        주어진 DataFrame에 대해 정의된 검증 규칙들을 실행합니다.

        Args:
            df (pd.DataFrame): 검증할 데이터프레임
            config (Dict[str, Any]): 검증 설정을 담은 딕셔너리

        Returns:
            pd.DataFrame: 검증 결과 요약 DataFrame
        """
        self.logger.info("데이터 품질 검증 시작...")
        self.validation_results = []  # 이전 결과 초기화

        # 필수 컬럼 검증
        if "required_columns" in config:
            self.check_required_columns(df, config["required_columns"])

        # 결측치 검증
        if "missing_value_checks" in config:
            for check_config in config["missing_value_checks"]:
                self.check_missing_values(df, check_config["column"], check_config.get("threshold", 0.0))

        # 데이터 타입 검증
        if "data_type_checks" in config:
            self.check_data_types(df, config["data_type_checks"])

        # 값 범위 검증
        if "range_checks" in config:
            self.check_value_ranges(df, config["range_checks"])

        # 고유값 검증
        if "unique_columns" in config:
            self.check_unique_values(df, config["unique_columns"])

        # 날짜 순서 검증
        if "date_sequence_checks" in config:
            self.check_date_sequence(df, config["date_sequence_checks"])

        # 카테고리 값 검증
        if "category_checks" in config:
            self.check_categorical_values(df, config["category_checks"])

        self.logger.info("데이터 품질 검증 완료.")
        return pd.DataFrame(self.validation_results)

if __name__ == '__main__':
    print("HVDCQualityChecker 테스트 시작...")
    
    # 테스트용 DataFrame 생성
    sample_df_data = {
        'NO.': [1, 2, 3, 4, None],
        'ATA': ['2023-01-01', '2023-01-05', '2023-02-10', None, '2023-03-15'],
        'MOSB': ['2023-01-15', '2023-01-20', '2023-02-05', '2023-03-01', '2023-03-10'],
        '리드타임(일)': [10, 15, None, 20, 25],
        'VENDOR': ['A', 'B', 'A', 'C', 'D'],
        '공정단계_HVDC_Label': ['Converter', 'Transmission', 'Filter/Reactor', 'Control/Protection', '기타']
    }
    test_df = pd.DataFrame(sample_df_data)
    
    # ATA, MOSB 컬럼을 datetime으로 변환
    test_df['ATA_dt'] = pd.to_datetime(test_df['ATA'], errors='coerce')
    test_df['MOSB_dt'] = pd.to_datetime(test_df['MOSB'], errors='coerce')

    checker = HVDCQualityChecker()
    
    validation_config = {
        "required_columns": ["NO.", "ATA", "MOSB", "리드타임(일)", "VENDOR", "공정단계_HVDC_Label"],
        "missing_value_checks": [
            {"column": "NO.", "threshold": 0.0},
            {"column": "ATA", "threshold": 5.0},
            {"column": "리드타임(일)", "threshold": 25.0}
        ],
        "data_type_checks": {
            "NO.": [int, float, np.int64, np.float64],
            "ATA_dt": [pd.Timestamp, datetime],
            "VENDOR": [str]
        },
        "range_checks": {
            "리드타임(일)": {"min": 0, "max": 365}
        },
        "unique_columns": ["NO."],
        "date_sequence_checks": {
            "ATA_dt": "MOSB_dt"
        },
        "category_checks": {
            "공정단계_HVDC_Label": ["Converter", "Transmission", "Filter/Reactor", 
                                "Control/Protection", "Grounding", "Spare/Maintenance", "기타"]
        }
    }
    
    results_df = checker.validate_dataframe(test_df, validation_config)
    
    print("\n검증 결과 요약:")
    print(results_df)
    
    print("\n세부 결과 메시지:")
    for index, row in results_df.iterrows():
        print(f"- {row['check_name']} ({row['details'].get('column', 'N/A')}): {row['status']} - {row['message']}")

    print("\nHVDCQualityChecker 테스트 종료.") 