"""
Data Generator Module
===================

This module provides functionality for generating sample HVDC project data.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from pathlib import Path
from scripts.core.base import HVDCBase

class HVDCDataGenerator(HVDCBase):
    """
    샘플 HVDC 프로젝트 상태 데이터를 생성하는 클래스입니다.
    HVDCBase를 상속받아 로깅 및 경로 설정을 활용합니다.
    """

    def __init__(self, num_samples: int = 560, seed: int = 42):
        """
        HVDCDataGenerator를 초기화합니다.

        Args:
            num_samples (int): 생성할 샘플 데이터의 수입니다.
            seed (int): 재현 가능한 결과를 위한 NumPy 랜덤 시드 값입니다.
        """
        super().__init__()  # 부모 HVDCBase 클래스의 __init__ 호출
        self.num_samples = num_samples
        self.seed = seed

        # data_dir이 HVDCBase에서 Path 객체로 설정되었다고 가정합니다.
        # 해당 디렉토리가 없으면 생성합니다.
        try:
            self.data_dir.mkdir(parents=True, exist_ok=True)
            self.logger.info(f"데이터 디렉토리 확인/생성 완료: {self.data_dir}")
        except AttributeError as e:
            self.logger.warning(f"HVDCBase에서 self.data_dir 또는 self.logger가 초기화되지 않았을 수 있습니다. ({e})")
            self.logger.warning("기본 경로 'data'를 사용합니다.")
            if not hasattr(self, 'data_dir'):
                self.data_dir = Path("data")
                self.data_dir.mkdir(parents=True, exist_ok=True)

    def generate_excel(self, filename: str = "HVDC-STATUS.xlsx") -> Path:
        """
        샘플 데이터를 생성하여 지정된 파일명으로 data_dir에 Excel 파일을 저장합니다.

        Args:
            filename (str): 저장할 Excel 파일의 이름입니다.

        Returns:
            Path: 생성된 Excel 파일의 전체 경로입니다.
        """
        self.logger.info(f"'{filename}' 이름으로 샘플 데이터 생성 시작 (총 {self.num_samples}개)...")
        np.random.seed(self.seed)

        data = {
            "NO.": range(1, self.num_samples + 1),
            "VENDOR": np.random.choice(["Vendor A", "Vendor B", "Vendor C", "Vendor D", "Vendor E"], self.num_samples),
            "MAIN DESCRIPTION (PO)": [f"PO Item {i:03d} - Main Component" for i in range(1, self.num_samples + 1)],
            "SUB DESCRIPTION": np.random.choice([
                "Converter Transformer Assembly",
                "Thyristor Valve Set",
                "DC XLPE Cable - Submarine",
                "AC Harmonic Filter Bank",
                "Smoothing Reactor Unit",
                "Control & Protection System Panel",
                "SCADA Interface Module",
                "Grounding Electrode System",
                "Spare Parts Kit - Critical",
                "Optical Fiber Cable for DCS",
                "PLC Module - Series X",
                "HVDC Bushing Assembly",
                "Surge Arrester - Station Class"
            ], self.num_samples),
            "INCOTERMS": np.random.choice(["CIF", "FOB", "EXW", "DAP", "DDP"], self.num_samples),
            "DG 분류": np.random.choice(["DG Class 3", "DG Class 8", "Non-DG", "DG Class 9"], self.num_samples, p=[0.05, 0.05, 0.8, 0.1]),
            "섬 운송 여부": np.random.choice(["Yes", "No"], self.num_samples, p=[0.25, 0.75])
        }

        # 날짜 생성 로직 (현재 날짜 기준)
        current_date = datetime.now()
        # ATA: 최근 1년 전부터 현재까지 랜덤 생성
        data["ATA"] = [current_date - timedelta(days=np.random.randint(0, 365)) for _ in range(self.num_samples)]
        # MOSB: ATA 날짜로부터 30일에서 200일 후로 랜덤 생성
        data["MOSB"] = [ata_date + timedelta(days=np.random.randint(30, 201)) for ata_date in data["ATA"]]

        df = pd.DataFrame(data)

        # 날짜 형식 변환 (Excel에서 날짜로 인식되도록)
        df["ATA"] = pd.to_datetime(df["ATA"]).dt.strftime('%Y-%m-%d')
        df["MOSB"] = pd.to_datetime(df["MOSB"]).dt.strftime('%Y-%m-%d')

        # 파일 저장 경로
        output_file_path = self.data_dir / filename
        
        try:
            df.to_excel(output_file_path, index=False, engine='openpyxl')
            self.logger.info(f"샘플 데이터 생성 완료: {output_file_path}")
        except Exception as e:
            self.logger.error(f"Excel 파일 저장 중 오류 발생: {output_file_path}, 오류: {e}")
            raise

        return output_file_path

if __name__ == '__main__':
    print("HVDCDataGenerator 테스트 시작...")
    
    try:
        generator = HVDCDataGenerator(num_samples=50, seed=101)  # 테스트용으로 샘플 수 줄임
        generated_file = generator.generate_excel(filename="HVDC-STATUS-Sample-Test.xlsx")
        
        if generated_file.exists():
            print(f"테스트용 샘플 파일이 성공적으로 생성되었습니다: {generated_file}")
        else:
            print(f"오류: 테스트용 샘플 파일이 생성되지 않았습니다.")

    except Exception as e:
        print(f"테스트 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()
    
    print("HVDCDataGenerator 테스트 종료.") 