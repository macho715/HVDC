"""
HVDC 데이터 처리 및 리포팅 자동화 파이프라인 모듈입니다.
"""

from pathlib import Path
from typing import Dict, Any, Optional
import datetime
import pandas as pd

from scripts.core.base import HVDCBase
from scripts.data.generator import HVDCDataGenerator
from scripts.logistics.mapper import HVDCLogisticsMapper
from scripts.data.validator import HVDCQualityChecker
from scripts.reporting.excel import HVDCExcelReporter
from scripts.reporting.dashboard import HVDCDashboardGenerator

class HVDCPipeline(HVDCBase):
    """
    HVDC 데이터 처리 및 리포팅 자동화 파이프라인 클래스입니다.
    """
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        """
        HVDCPipeline을 초기화합니다.

        Args:
            config (Optional[Dict[str, Any]]): 파이프라인 실행을 위한 설정값.
                                                향후 config.py에서 로드될 수 있습니다.
        """
        super().__init__()
        self.config = config if config is not None else {} # 기본 빈 딕셔너리로 설정
        self.logger.info("HVDCPipeline 초기화 완료.")

    def run_pipeline(self, generate_sample_data: bool = False) -> bool:
        """
        전체 데이터 처리 및 리포팅 파이프라인을 실행합니다.

        Args:
            generate_sample_data (bool): True이면 샘플 데이터를 생성합니다.
                                         False이면 기존 데이터를 사용하려고 시도합니다.
        Returns:
            bool: 파이프라인 전체 실행 성공 여부.
        """
        self.logger.info("===== HVDC 자동화 파이프라인 시작 =====")
        
        try:
            # --- 1. (선택적) 샘플 데이터 생성 ---
            raw_data_filename = self.config.get("input_data_filename", "HVDC-STATUS.xlsx")
            if generate_sample_data:
                self.logger.info("--- 단계 1: 샘플 데이터 생성 시작 ---")
                data_generator = HVDCDataGenerator(
                    num_samples=self.config.get("num_samples_to_generate", 560)
                )
                generated_raw_data_path = data_generator.generate_excel(filename=raw_data_filename)
                if not generated_raw_data_path or not generated_raw_data_path.exists():
                    self.logger.error("샘플 데이터 생성에 실패하여 파이프라인을 중단합니다.")
                    return False
                self.logger.info(f"샘플 데이터 생성 완료: {generated_raw_data_path}")
            else:
                self.logger.info(f"--- 단계 1: 기존 데이터 사용 예정 ({raw_data_filename}) ---")

            # --- 2. 데이터 매핑 및 변환 ---
            self.logger.info("--- 단계 2: 데이터 매핑 및 변환 시작 ---")
            mapper = HVDCLogisticsMapper(
                input_filename=raw_data_filename,
                max_no_filter=self.config.get("mapper_max_no_filter", 560)
            )
            processed_df = mapper.process_data()
            if processed_df is None or processed_df.empty:
                self.logger.error("데이터 매핑 결과가 비어있어 파이프라인을 중단합니다.")
                return False
            self.logger.info(f"데이터 매핑 및 변환 완료. 처리된 행 수: {len(processed_df)}")

            # --- 3. 데이터 품질 검증 ---
            self.logger.info("--- 단계 3: 데이터 품질 검증 시작 ---")
            quality_checker = HVDCQualityChecker()
            validation_config = self.config.get("validation_rules", self._get_default_validation_config())
            validation_results_df = quality_checker.validate_dataframe(processed_df, config=validation_config)
            
            if not validation_results_df.empty and 'status' in validation_results_df.columns:
                if "FAIL" in validation_results_df["status"].unique():
                    self.logger.error("데이터 품질 검증 실패 항목이 발견되었습니다. 상세 내용은 검증 결과를 확인하세요.")
                else:
                    self.logger.info("데이터 품질 검증 통과 (또는 WARN 항목만 존재).")
            else:
                self.logger.warning("데이터 품질 검증 결과가 없거나 'status' 컬럼이 없습니다.")

            # --- 4. 처리된 데이터 Excel 파일로 저장 ---
            self.logger.info("--- 단계 4: 처리된 데이터 Excel 파일 저장 시작 ---")
            excel_reporter = HVDCExcelReporter()
            processed_data_filename = self.config.get("processed_data_filename", "final_mapping.xlsx")
            processed_excel_path = excel_reporter.create_report(processed_df, filename=processed_data_filename)
            if not processed_excel_path or not processed_excel_path.exists():
                self.logger.error("처리된 데이터를 Excel 파일로 저장하는 데 실패하여 파이프라인을 중단합니다.")
                return False
            self.logger.info(f"처리된 데이터 Excel 파일 저장 완료: {processed_excel_path}")

            # --- 5. KPI 대시보드 생성 ---
            self.logger.info("--- 단계 5: KPI 대시보드 생성 시작 ---")
            dashboard_generator = HVDCDashboardGenerator()
            dashboard_filename_from_config = self.config.get("dashboard_filename", "HVDC_KPI_Dashboard.xlsx")
            try:
                dashboard_path = dashboard_generator.create_dashboard(
                    filename=dashboard_filename_from_config,
                    source_data=processed_df
                )
                self.logger.info(f"HVDCDashboardGenerator.create_dashboard 반환 값: {dashboard_path} (타입: {type(dashboard_path)})")
                
                if not dashboard_path or not dashboard_path.exists():
                    self.logger.warning("KPI 대시보드 생성에 실패했거나 파일이 생성되지 않았습니다.")
                else:
                    self.logger.info(f"KPI 대시보드 생성 완료: {dashboard_path}")
            except Exception as e:
                self.logger.error(f"KPI 대시보드 생성 중 오류 발생: {e}")
                raise

            # --- 6. 추가 데이터 분석 (HVDCLogisticsAnalyzer 사용) ---
            self.logger.info("--- 단계 6: 추가 데이터 분석 시작 ---")
            from scripts.logistics.analyzer import HVDCLogisticsAnalyzer
            analyzer = HVDCLogisticsAnalyzer()
            lead_time_analysis = analyzer.analyze_lead_time(processed_df)
            vendor_performance = analyzer.analyze_vendor_performance(processed_df)
            summary_report = analyzer.generate_summary_report(processed_df)
            self.logger.info(f"리드타임 분석 결과: {str(lead_time_analysis)[:200]}...")
            self.logger.info(f"벤더 성과 분석 결과: {str(vendor_performance)[:200]}...")
            self.logger.info(f"요약 보고서 결과: {str(summary_report)[:200]}...")
            # (선택) 분석 결과를 Excel로 저장
            try:
                with pd.ExcelWriter(self.output_dir / 'analysis_results.xlsx') as writer:
                    if 'lead_time_stats' in lead_time_analysis:
                        pd.DataFrame([lead_time_analysis['lead_time_stats']]).to_excel(writer, sheet_name='LeadTimeStats')
                    if 'vendor_stats' in vendor_performance:
                        pd.DataFrame(vendor_performance['vendor_stats']).to_excel(writer, sheet_name='VendorStats')
                    if summary_report:
                        pd.DataFrame([summary_report]).to_excel(writer, sheet_name='Summary')
                self.logger.info('분석 결과가 analysis_results.xlsx로 저장되었습니다.')
            except Exception as e:
                self.logger.error(f'분석 결과 저장 중 오류: {e}')

            self.logger.info("run_pipeline 메소드의 거의 마지막 지점 도달.")
            self.logger.info("===== HVDC 자동화 파이프라인 성공적으로 완료 =====")
            return True

        except FileNotFoundError as e:
            self.logger.error(f"파일 접근 오류로 파이프라인 중단: {e}")
            return False
        except ValueError as e:
            self.logger.error(f"데이터 값 오류로 파이프라인 중단: {e}")
            return False
        except Exception as e:
            self.logger.critical(f"예상치 못한 오류로 파이프라인 중단: {e}", exc_info=True)
            return False

    def _get_default_validation_config(self) -> Dict[str, Any]:
        """HVDCQualityChecker를 위한 기본 검증 규칙 설정을 반환합니다."""
        self.logger.info("기본 검증 규칙 설정 사용.")
        
        hvdc_step_labels_list = [
            "Converter", "Transmission", "Filter/Reactor", "Control/Protection", 
            "Grounding", "Spare/Maintenance", "기타"
        ]

        return {
            "missing_value_checks": [
                {"column": "NO.", "threshold_percent": 0.0},
                {"column": "ATA", "threshold_percent": 5.0},
                {"column": "MOSB", "threshold_percent": 5.0},
                {"column": "리드타임(일)", "threshold_percent": 10.0},
                {"column": "공정단계_HVDC_Label", "threshold_percent": 0.0}
            ],
            "data_type_checks": {
                "NO.": ["int64", "float64", int, float],
                "ATA": ["datetime64[ns]", pd.Timestamp, datetime.datetime],
                "MOSB": ["datetime64[ns]", pd.Timestamp, datetime.datetime],
                "리드타임(일)": ["float64", "int64", float, int],
                "공정단계_HVDC_Label": [str],
                "리드타임 상태": [str],
                "VENDOR": [str]
            },
            "categorical_checks": [
                {
                    "column": "공정단계_HVDC_Label",
                    "allowed_values": hvdc_step_labels_list
                },
                {
                    "column": "리드타임 상태",
                    "allowed_values": ["양호", "주의", "지연", "미도착"]
                }
            ]
        }

# 스크립트를 직접 실행하여 파이프라인을 테스트할 경우
if __name__ == '__main__':
    print("HVDCPipeline 테스트 실행...")
    
    # 테스트용 설정
    test_pipeline_config = {
        "input_data_filename": "HVDC-STATUS-Sample-Test.xlsx",
        "processed_data_filename": "Test-final_mapping_output.xlsx",
        "dashboard_filename": "Test-HVDC_KPI_Dashboard_output.xlsx",
        "num_samples_to_generate": 50,
        "mapper_max_no_filter": 50,
        "validation_rules": {}
    }

    # HVDCBase 모의 객체 설정
    if not hasattr(HVDCBase, '_is_mocked_for_test_pipeline'):
        class MockHVDCBasePipeline:
            _is_mocked_for_test_pipeline = True
            def __init__(self):
                self.project_root = Path(__file__).resolve().parent.parent.parent
                self.data_dir = self.project_root / "data"
                self.output_dir = self.project_root / "output"
                import logging
                logging.basicConfig(level=logging.DEBUG, 
                                    format='%(asctime)s - %(name)s - %(levelname)s - [%(funcName)s:%(lineno)d] - %(message)s')
                self.logger = logging.getLogger(self.__class__.__name__)
        
        OriginalHVDCBase_Pipeline = HVDCBase 
        HVDCBase = MockHVDCBasePipeline
        print(f"테스트: HVDCBase 모의 객체 (Pipeline용) 사용. Project Root: {HVDCBase().project_root}")

    pipeline = HVDCPipeline(config=test_pipeline_config)
    success = pipeline.run_pipeline(generate_sample_data=True) 

    if success:
        print("\n파이프라인 실행 성공!")
    else:
        print("\n파이프라인 실행 중 오류 발생 또는 실패.")

    # 테스트 완료 후 HVDCBase 원상 복구
    if 'OriginalHVDCBase_Pipeline' in locals():
        HVDCBase = OriginalHVDCBase_Pipeline
        print("테스트: HVDCBase 원본으로 복원됨.")
        
    print("HVDCPipeline 테스트 종료.") 