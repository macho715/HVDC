import sys
import subprocess
from pathlib import Path
import shutil
from datetime import datetime
import pandas as pd
import logging
import argparse

class HVDCPipeline:
    def __init__(self):
        self.script_dir = Path(__file__).parent
        self.root_dir = self.script_dir.parent
        self.data_dir = self.root_dir / 'data'
        self.output_dir = self.root_dir / 'output'
        self.log_dir = self.root_dir / 'logs'
        
        # 로그 디렉토리 생성
        self.log_dir.mkdir(parents=True, exist_ok=True)
        
        # 로깅 설정
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(self.log_dir / 'pipeline.log', encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def upload_new_data(self, new_file_name):
        """신규 데이터 업로드 및 백업"""
        if str(new_file_name).startswith("data/") or str(new_file_name).startswith("data\\"):
            new_file_path = Path(new_file_name)
        else:
            new_file_path = self.data_dir / new_file_name
        main_file_path = self.data_dir / 'HVDC-STATUS.xlsx'
        log_file = self.data_dir / 'upload_log.txt'
        
        if not new_file_path.exists():
            self.logger.error(f"신규 파일이 존재하지 않습니다: {new_file_path}")
            return False
        
        try:
            # 기존 파일 백업
            if main_file_path.exists():
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                backup_path = self.data_dir / f"HVDC-STATUS_백업_{timestamp}.xlsx"
                shutil.copy2(main_file_path, backup_path)
                self.logger.info(f"기존 파일 백업: {backup_path}")
            
            # 신규 파일을 메인 파일로 복사
            shutil.copy2(new_file_path, main_file_path)
            self.logger.info(f"신규 파일을 HVDC-STATUS.xlsx로 교체 완료")
            
            # 업로드 로그 기록
            with open(log_file, 'a', encoding='utf-8') as f:
                f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {new_file_name} → HVDC-STATUS.xlsx\n")
            self.logger.info(f"업로드 이력 기록: {log_file}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"데이터 업로드 중 오류 발생: {str(e)}")
            return False
    
    def run_step(self, description, command):
        """단계별 실행 및 로깅"""
        self.logger.info(f"=== {description} ===")
        try:
            result = subprocess.run(command, shell=True, capture_output=True, text=True)
            if result.returncode == 0:
                self.logger.info(f"[성공] {description}")
                if result.stdout:
                    self.logger.debug(f"출력: {result.stdout}")
                return True
            else:
                self.logger.error(f"[실패] {description}")
                if result.stderr:
                    self.logger.error(f"오류: {result.stderr}")
                return False
        except Exception as e:
            self.logger.error(f"명령 실행 중 오류 발생: {str(e)}")
            return False
    
    def validate_data(self):
        """데이터 유효성 검증"""
        try:
            data_path = self.data_dir / 'HVDC-STATUS.xlsx'
            df = pd.read_excel(data_path)
            
            # 실제 컬럼명 출력
            self.logger.info(f"실제 컬럼명 목록: {list(df.columns)}")
            
            # 필수 컬럼 확인
            required_columns = ['NO.', 'SCT SHIP NO.', 'VENDOR', 'MAIN DESCRIPTION (PO)', 'SUB DESCRIPTION']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                self.logger.error(f"필수 컬럼 누락: {missing_columns}")
                return False
            
            # 숫자형 검증 제거 (SCT SHIP NO.는 문자열로 처리)
            
            self.logger.info("데이터 유효성 검증 완료")
            return True
            
        except Exception as e:
            self.logger.error(f"데이터 검증 중 오류 발생: {str(e)}")
            return False
    
    def run_pipeline(self, new_file_name, weighted_delay=False, route_delay_report=False, export=None):
        self.logger.info("=== HVDC 데이터 파이프라인 시작 ===")
        # 1. 신규 데이터 업로드
        if not self.upload_new_data(new_file_name):
            self.logger.error("데이터 업로드 실패")
            return False
        # 2. 데이터 유효성 검증
        if not self.validate_data():
            self.logger.error("데이터 유효성 검증 실패")
            return False
        # 3. 매핑/전처리 실행
        if not self.run_step(
            "최신 데이터로 매핑/전처리 실행",
            f"python {self.script_dir / 'logistics_mapper.py'}"
        ):
            return False
        # 4. 품질점검 실행
        if not self.run_step(
            "품질점검 자동 실행",
            f"python {self.script_dir / 'quality_check.py'}"
        ):
            return False
        # 5. 가중치 기반 지연 분석 및 경로별 리포트
        if route_delay_report:
            analyze_cmd = f"python {self.script_dir / 'analyze_data.py'} --route-delay-report"
            if weighted_delay:
                analyze_cmd += " --weighted-delay"
            if export:
                analyze_cmd += f" --export {export}"
            if not self.run_step(
                "경로별 지연 리포트 생성",
                analyze_cmd
            ):
                return False
        self.logger.info("=== 전체 자동화 파이프라인 완료 ===")
        self.logger.info("결과: output/ 및 output/quality_check/ 폴더를 확인하세요.")
        return True

def main():
    parser = argparse.ArgumentParser(description="HVDC 자동 파이프라인")
    parser.add_argument("new_file_name", help="신규 엑셀 파일명")
    parser.add_argument("--weighted-delay", action="store_true", help="가중치 기반 지연 분석 실행")
    parser.add_argument("--route-delay-report", action="store_true", help="경로별 지연 리포트 생성")
    parser.add_argument("--export", type=str, default=None, help="결과 내보내기 형식(html 등)")
    parser.add_argument("--forecast", action="store_true",
                        help="예측 모델 학습·예측 실행")
    args = parser.parse_args()

    pipeline = HVDCPipeline()
    if not pipeline.run_pipeline(
        args.new_file_name,
        weighted_delay=args.weighted_delay,
        route_delay_report=args.route_delay_report,
        export=args.export
    ):
        sys.exit(1)

    if args.forecast:
        subprocess.run([
            "python", "scripts/fcast.py",
            "--file", args.new_file_name,
            "--value-col", "전체 리드타임",
            "--cv", "--horizon", "6"
        ], check=True)

if __name__ == "__main__":
    main() 