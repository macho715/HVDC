import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from pathlib import Path
import logging
import sys

class HVDCDataGenerator:
    def __init__(self):
        self.data_dir = Path('data')
        self.output_dir = Path('output')
        self.log_dir = Path('logs')
        
        # 디렉토리 생성
        for dir_path in [self.data_dir, self.output_dir, self.log_dir]:
            dir_path.mkdir(parents=True, exist_ok=True)
        
        # 로깅 설정
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(self.log_dir / 'data_generator.log', encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def generate_sample_data(self, n_samples=560, seed=42):
        """샘플 데이터 생성"""
        self.logger.info(f"샘플 데이터 생성 시작 (n_samples={n_samples})")
        np.random.seed(seed)
        
        try:
            # 기본 데이터
            data = {
                "NO.": range(1, n_samples + 1),
                "VENDOR": np.random.choice(["Vendor A", "Vendor B", "Vendor C"], n_samples),
                "MAIN DESCRIPTION (PO)": [f"Item {i}" for i in range(1, n_samples + 1)],
                "SUB DESCRIPTION": np.random.choice([
                    "Converter Transformer",
                    "DC Cable",
                    "Filter Reactor",
                    "Control System",
                    "Grounding Equipment",
                    "Spare Parts",
                    "Other Components"
                ], n_samples),
                "INCOTERMS": np.random.choice(["CIF", "FOB", "EXW", "DAP"], n_samples),
                "DG 분류": np.random.choice(["DG", "Non-DG"], n_samples, p=[0.1, 0.9]),
                "섬 운송 여부": np.random.choice(["Yes", "No"], n_samples, p=[0.3, 0.7])
            }
            
            # 날짜 생성
            base_date = datetime(2024, 1, 1)
            data["ATA"] = [base_date + timedelta(days=np.random.randint(0, 30)) for _ in range(n_samples)]
            data["Customs Close"] = [ata + timedelta(days=np.random.randint(1, 5)) for ata in data["ATA"]]
            data["DSV Out"] = [cc + timedelta(days=np.random.randint(1, 3)) for cc in data["Customs Close"]]
            data["MOSB"] = [dsv + timedelta(days=np.random.randint(1, 7)) for dsv in data["DSV Out"]]
            
            # 컨테이너 데이터 추가
            data["20ft Q'TY"] = np.random.randint(0, 5, n_samples)
            data["40ft Q'TY"] = np.random.randint(0, 3, n_samples)
            data["TOTAL Q'TY"] = data["20ft Q'TY"] + data["40ft Q'TY"]
            data["SCT SHIP NO."] = [f"SCT{np.random.randint(1000, 9999)}" for _ in range(n_samples)]
            
            # 현장 데이터 추가
            data["MIR"] = np.random.choice([None, "MIR"], n_samples, p=[0.7, 0.3])
            data["SHU"] = np.random.choice([None, "SHU"], n_samples, p=[0.7, 0.3])
            data["DAS"] = np.random.choice([None, "DAS"], n_samples, p=[0.7, 0.3])
            data["AGI"] = np.random.choice([None, "AGI"], n_samples, p=[0.7, 0.3])
            
            # DataFrame 생성
            df = pd.DataFrame(data)
            
            # 엑셀 파일로 저장
            output_file = self.data_dir / 'HVDC-STATUS.xlsx'
            df.to_excel(output_file, index=False)
            self.logger.info(f"샘플 데이터 생성 완료: {output_file}")
            
            return df
            
        except Exception as e:
            self.logger.error(f"샘플 데이터 생성 중 오류 발생: {str(e)}")
            raise
    
    def process_data(self, df):
        """데이터 전처리 및 매핑"""
        self.logger.info("데이터 전처리 및 매핑 시작")
        
        try:
            # 1. 리드타임 계산
            df["ATA"] = pd.to_datetime(df["ATA"], errors="coerce")
            df["MOSB"] = pd.to_datetime(df["MOSB"], errors="coerce")
            df["리드타임(일)"] = (df["MOSB"] - df["ATA"]).dt.days
            self.logger.info("리드타임 계산 완료")
            
            # 2. 공정 분류
            def updated_hvdc_step(desc):
                if pd.isna(desc):
                    return 99
                desc = str(desc).lower()
                if any(k in desc for k in ["converter transformer", "valve", "thyristor", "igbt"]):
                    return 1
                elif any(k in desc for k in ["dc cable", "submarine", "overhead", "transmission"]):
                    return 2
                elif any(k in desc for k in ["filter", "reactor", "capacitor", "harmonic"]):
                    return 3
                elif any(k in desc for k in ["scada", "control", "protection", "monitoring"]):
                    return 4
                elif any(k in desc for k in ["grounding", "electrode", "earth"]):
                    return 5
                elif any(k in desc for k in ["spare", "repair"]):
                    return 6
                else:
                    return 99
            
            df["공정단계_HVDC"] = df["SUB DESCRIPTION"].apply(updated_hvdc_step)
            df["공정단계_HVDC_Label"] = df["공정단계_HVDC"].map({
                1: "Converter",
                2: "Transmission",
                3: "Filter/Reactor",
                4: "Control/Protection",
                5: "Grounding",
                6: "Spare/Maintenance",
                99: "기타"
            })
            self.logger.info("공정 분류 완료")
            
            # 3. 리드타임 상태 분류
            df["리드타임 상태"] = df["리드타임(일)"].apply(
                lambda x: "미도착" if pd.isna(x) else ("양호" if x <= 30 else ("주의" if x <= 90 else "지연"))
            )
            self.logger.info("리드타임 상태 분류 완료")
            
            # 4. 최종 자재 리스트 정리
            final_table = df[df["NO."] <= 560][[
                "NO.", "VENDOR", "MAIN DESCRIPTION (PO)", "SUB DESCRIPTION",
                "공정단계_HVDC_Label", "리드타임(일)", "리드타임 상태",
                "INCOTERMS", "DG 분류", "섬 운송 여부", "ATA", "MOSB"
            ]].sort_values("NO.")
            
            self.logger.info(f"최종 매핑 테이블 생성 완료: {len(final_table)} 행")
            return final_table
            
        except Exception as e:
            self.logger.error(f"데이터 처리 중 오류 발생: {str(e)}")
            raise
    
    def save_results(self, final_table):
        """결과 저장"""
        try:
            # 엑셀 파일로 저장
            output_file = self.output_dir / 'final_mapping.xlsx'
            final_table.to_excel(output_file, index=False)
            self.logger.info(f"매핑 결과 저장 완료: {output_file}")
            
            # 통계 정보 저장
            stats_file = self.output_dir / 'mapping_stats.txt'
            with open(stats_file, 'w', encoding='utf-8') as f:
                f.write("=== HVDC 매핑 통계 ===\n\n")
                f.write(f"총 자재 수: {len(final_table)}\n")
                f.write(f"공정별 분포:\n{final_table['공정단계_HVDC_Label'].value_counts()}\n\n")
                f.write(f"리드타임 상태 분포:\n{final_table['리드타임 상태'].value_counts()}\n\n")
                f.write(f"평균 리드타임: {final_table['리드타임(일)'].mean():.1f}일\n")
                f.write(f"최대 리드타임: {final_table['리드타임(일)'].max():.1f}일\n")
                f.write(f"최소 리드타임: {final_table['리드타임(일)'].min():.1f}일\n")
            
            self.logger.info(f"통계 정보 저장 완료: {stats_file}")
            
        except Exception as e:
            self.logger.error(f"결과 저장 중 오류 발생: {str(e)}")
            raise
    
    def run_pipeline(self, n_samples=560):
        """전체 파이프라인 실행"""
        try:
            # 1. 샘플 데이터 생성
            df = self.generate_sample_data(n_samples)
            
            # 2. 데이터 처리
            final_table = self.process_data(df)
            
            # 3. 결과 저장
            self.save_results(final_table)
            
            self.logger.info("=== 전체 파이프라인 완료 ===")
            return True
            
        except Exception as e:
            self.logger.error(f"파이프라인 실행 중 오류 발생: {str(e)}")
            return False

def main():
    generator = HVDCDataGenerator()
    if not generator.run_pipeline():
        sys.exit(1)

if __name__ == "__main__":
    main() 