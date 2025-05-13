import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
from datetime import datetime
import plotly.figure_factory as ff
import plotly.express as px
import matplotlib.font_manager as fm

class HVDCQualityCheck:
    def __init__(self, data_path):
        # 한글 폰트 설정
        plt.rcParams['font.family'] = 'Malgun Gothic'  # 윈도우 기본 한글 폰트
        plt.rcParams['axes.unicode_minus'] = False
        
        # 데이터 로드
        if not Path(data_path).exists():
            raise FileNotFoundError(f"데이터 파일을 찾을 수 없습니다: {data_path}")
        self.data = pd.read_excel(data_path)
        
        # 출력 디렉토리 설정
        self.output_dir = Path('output/quality_check')
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
    def check_columns(self):
        """컬럼 구조 점검"""
        column_info = []
        for i, col in enumerate(self.data.columns):
            column_info.append({
                '순번': i + 1,
                '컬럼명': col,
                '타입': str(type(col)),
                '데이터 타입': str(self.data[col].dtype),
                '결측치 수': self.data[col].isnull().sum(),
                '고유값 수': self.data[col].nunique()
            })
        return pd.DataFrame(column_info)
    
    def check_data_completeness(self):
        """기본 데이터 완성도 점검"""
        checks = {
            "총 행 개수": len(self.data),
            "Material_ID 범위": f"{self.data['NO.'].min()} ~ {self.data['NO.'].max()}",
            "중복 Material_ID": self.data['NO.'].duplicated().sum(),
            "누락된 ATA": self.data['ATA'].isnull().sum(),
            "누락된 MOSB": self.data['MOSB'].isnull().sum()
        }
        return pd.Series(checks)
    
    def check_leadtime(self):
        """리드타임 이상값 점검"""
        leadtime = self.data['리드타임(일)']
        checks = {
            "음수 리드타임": len(leadtime[leadtime < 0]),
            "180일 초과": len(leadtime[leadtime > 180]),
            "NaN 리드타임": leadtime.isnull().sum(),
            "평균 리드타임": leadtime.mean(),
            "중앙값 리드타임": leadtime.median()
        }
        return pd.Series(checks)
    
    def check_process_classification(self):
        """공정 분류 점검"""
        process_counts = self.data['공정단계_HVDC_Label'].value_counts()
        process_ratio = process_counts / len(self.data) * 100
        
        checks = {
            "총 공정 단계 수": len(process_counts),
            "99(기타) 비율": process_ratio.get("기타", 0),
            "가장 많은 공정": process_counts.index[0],
            "가장 적은 공정": process_counts.index[-1]
        }
        return pd.Series(checks)
    
    def check_incoterms(self):
        """INCOTERMS 분포 점검"""
        incoterms = self.data['INCOTERMS'].value_counts()
        return incoterms
    
    def check_dangerous_goods(self):
        """위험물 분류 점검"""
        dg_count = (self.data['DG 분류'] == 'DG').sum()
        return {
            "위험물 총 개수": dg_count,
            "위험물 비율": (dg_count / len(self.data)) * 100
        }
    
    def create_visualizations(self):
        """시각화 생성"""
        # 1. 리드타임 히스토그램
        plt.figure(figsize=(10, 6))
        sns.histplot(data=self.data, x='리드타임(일)', bins=30)
        plt.title('리드타임 분포')
        plt.savefig(self.output_dir / 'leadtime_distribution.png')
        plt.close()
        
        # 2. 공정별 리드타임 박스플롯
        plt.figure(figsize=(12, 6))
        sns.boxplot(data=self.data, x='공정단계_HVDC_Label', y='리드타임(일)')
        plt.xticks(rotation=45)
        plt.title('공정별 리드타임 분포')
        plt.tight_layout()
        plt.savefig(self.output_dir / 'process_leadtime_boxplot.png')
        plt.close()
        
        # 3. Gantt Chart (Plotly)
        gantt_data = []
        for _, row in self.data.iterrows():
            if pd.notnull(row['ATA']) and pd.notnull(row['MOSB']):
                gantt_data.append(dict(
                    Task=f"Material {row['NO.']}",
                    Start=row['ATA'],
                    Finish=row['MOSB'],
                    Resource=row['공정단계_HVDC_Label']
                ))
        
        fig = ff.create_gantt(gantt_data, 
                            index_col='Resource',
                            show_colorbar=True,
                            group_tasks=True)
        fig.write_html(str(self.output_dir / 'gantt_chart.html'))
        
        # 4. 컬럼별 결측치 시각화
        plt.figure(figsize=(12, 6))
        missing_data = self.data.isnull().sum() / len(self.data) * 100
        missing_data = missing_data[missing_data > 0].sort_values(ascending=False)
        sns.barplot(x=missing_data.index, y=missing_data.values)
        plt.xticks(rotation=45)
        plt.title('컬럼별 결측치 비율')
        plt.ylabel('결측치 비율 (%)')
        plt.tight_layout()
        plt.savefig(self.output_dir / 'missing_data_analysis.png')
        plt.close()
    
    def generate_report(self):
        """종합 품질 점검 보고서 생성"""
        report = {
            "1. 컬럼 구조 점검": self.check_columns(),
            "2. 데이터 완성도": self.check_data_completeness(),
            "3. 리드타임 점검": self.check_leadtime(),
            "4. 공정 분류 점검": self.check_process_classification(),
            "5. INCOTERMS 분포": self.check_incoterms(),
            "6. 위험물 점검": self.check_dangerous_goods()
        }
        
        # 보고서 저장
        with open(self.output_dir / 'quality_check_report.txt', 'w', encoding='utf-8') as f:
            f.write("=== HVDC 데이터 품질 점검 보고서 ===\n")
            f.write(f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            
            for section, data in report.items():
                f.write(f"\n{section}\n")
                f.write("-" * 50 + "\n")
                if isinstance(data, (pd.Series, pd.DataFrame)):
                    f.write(data.to_string())
                else:
                    f.write(str(data))
                f.write("\n")
        
        # 시각화 생성
        self.create_visualizations()
        
        print(f"품질 점검 보고서가 {self.output_dir}에 생성되었습니다.")

if __name__ == "__main__":
    try:
        # 품질 점검 실행
        checker = HVDCQualityCheck('output/final_mapping.xlsx')
        checker.generate_report()
    except Exception as e:
        print(f"Error: {e}") 