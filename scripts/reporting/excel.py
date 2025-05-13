"""
HVDC Excel 리포트 생성 모듈입니다.
"""

from pathlib import Path
from typing import Optional
import pandas as pd
from scripts.core.base import HVDCBase

class HVDCExcelReporter(HVDCBase):
    """
    HVDC 데이터프레임을 Excel 파일로 저장하는 리포터 클래스입니다.
    """
    def __init__(self):
        super().__init__()
        self.logger.info("HVDCExcelReporter 초기화 완료.")

    def create_report(self, df: pd.DataFrame, filename: Optional[str] = None) -> Optional[Path]:
        """
        DataFrame을 Excel 파일로 저장합니다.

        Args:
            df (pd.DataFrame): 저장할 데이터프레임
            filename (Optional[str]): 저장할 파일명 (output 디렉토리 기준)

        Returns:
            Optional[Path]: 저장된 파일의 경로 (성공 시), 실패 시 None
        """
        if filename is None:
            filename = "final_mapping.xlsx"
        output_path = self.output_dir / filename
        try:
            df.to_excel(output_path, index=False)
            self.logger.info(f"Excel 리포트 저장 완료: {output_path}")
            return output_path
        except Exception as e:
            self.logger.error(f"Excel 리포트 저장 중 오류 발생: {e}")
            return None

# 테스트 코드
if __name__ == "__main__":
    import numpy as np
    # 테스트용 데이터프레임 생성
    test_df = pd.DataFrame({
        "NO.": range(1, 6),
        "VENDOR": ["A", "B", "C", "A", "B"],
        "리드타임(일)": np.random.randint(10, 50, 5)
    })
    reporter = HVDCExcelReporter()
    saved_path = reporter.create_report(test_df, filename="test_report.xlsx")
    print(f"저장된 파일 경로: {saved_path}") 