import shutil
from pathlib import Path
import pandas as pd
import os

class ProjectSetup:
    def __init__(self):
        self.root_dir = Path.cwd()
        self.data_dir = self.root_dir / 'data'
        self.output_dir = self.root_dir / 'output'
        self.scripts_dir = self.root_dir / 'scripts'
        
    def create_directories(self):
        """필요한 디렉토리 생성"""
        for dir_path in [self.data_dir, self.output_dir, self.scripts_dir]:
            dir_path.mkdir(exist_ok=True)
            print(f"✓ 디렉토리 생성/확인: {dir_path}")
    
    def find_excel_files(self):
        """프로젝트 내 엑셀 파일 검색"""
        excel_files = []
        for root, _, files in os.walk(self.root_dir):
            for file in files:
                if file.endswith(('.xlsx', '.xls')):
                    excel_files.append(Path(root) / file)
        return excel_files
    
    def organize_files(self):
        """파일 정리 및 이동"""
        # 1. 엑셀 파일 검색
        excel_files = self.find_excel_files()
        
        # 2. 파일 이동
        for file_path in excel_files:
            if file_path.name == 'HVDC-STATUS.xlsx':
                # 원본 데이터 파일은 data/ 폴더로
                target_path = self.data_dir / file_path.name
                if file_path != target_path:
                    try:
                        if target_path.exists():
                            # 기존 파일 백업
                            backup_path = target_path.with_suffix('.xlsx.bak')
                            shutil.move(str(target_path), str(backup_path))
                            print(f"✓ 기존 파일 백업: {target_path} → {backup_path}")
                        shutil.move(str(file_path), str(target_path))
                        print(f"✓ 원본 데이터 이동: {file_path} → {target_path}")
                    except Exception as e:
                        print(f"! 파일 이동 실패: {e}")
            elif file_path.name == 'final_mapping.xlsx':
                # 분석 결과는 output/ 폴더에 유지
                if file_path.parent != self.output_dir:
                    target_path = self.output_dir / file_path.name
                    try:
                        if target_path.exists():
                            # 기존 파일 백업
                            backup_path = target_path.with_suffix('.xlsx.bak')
                            shutil.move(str(target_path), str(backup_path))
                            print(f"✓ 기존 파일 백업: {target_path} → {backup_path}")
                        shutil.move(str(file_path), str(target_path))
                        print(f"✓ 분석 결과 이동: {file_path} → {target_path}")
                    except Exception as e:
                        print(f"! 파일 이동 실패: {e}")
    
    def verify_setup(self):
        """설정 검증"""
        # 1. 원본 데이터 확인
        original_data = self.data_dir / 'HVDC-STATUS.xlsx'
        if original_data.exists():
            try:
                df = pd.read_excel(original_data)
                print(f"✓ 원본 데이터 검증 성공: {len(df)} 행")
            except Exception as e:
                print(f"! 원본 데이터 검증 실패: {e}")
        else:
            print("! 원본 데이터 파일을 찾을 수 없습니다.")
        
        # 2. 분석 결과 확인
        mapping_result = self.output_dir / 'final_mapping.xlsx'
        if mapping_result.exists():
            try:
                df = pd.read_excel(mapping_result)
                print(f"✓ 분석 결과 검증 성공: {len(df)} 행")
            except Exception as e:
                print(f"! 분석 결과 검증 실패: {e}")
        else:
            print("! 분석 결과 파일을 찾을 수 없습니다.")

def main():
    print("=== HVDC Analysis 프로젝트 설정 시작 ===")
    setup = ProjectSetup()
    
    # 1. 디렉토리 생성
    setup.create_directories()
    
    # 2. 파일 정리
    setup.organize_files()
    
    # 3. 설정 검증
    setup.verify_setup()
    
    print("\n=== 프로젝트 설정 완료 ===")
    print("\n프로젝트 구조:")
    print(f"📁 data/     - 원본 데이터")
    print(f"📁 output/   - 분석 결과")
    print(f"📁 scripts/  - 분석 스크립트")

if __name__ == "__main__":
    main() 