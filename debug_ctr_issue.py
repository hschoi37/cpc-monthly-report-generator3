import pandas as pd
from openpyxl import load_workbook

def analyze_ctr_data(file_path):
    """CTR 데이터 추출 오류 분석"""
    
    wb = load_workbook(file_path, data_only=True)
    
    # 문제가 있는 시트들 분석
    problem_sheets = ['일편등심명동_0818', '육목원_0817']
    
    for sheet_name in problem_sheets:
        print(f"\n{'='*60}")
        print(f"시트명: {sheet_name}")
        print(f"{'='*60}")
        
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        
        # 주차별 데이터 찾기
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell_value = str(df.iloc[i, j])
                if '주차' in cell_value and ('1주차' in cell_value or '2주차' in cell_value):
                    print(f"\n주차별 데이터 발견: 행 {i}, '{cell_value}'")
                    
                    # 해당 행과 다음 행 출력
                    for k in range(max(0, i-1), min(i+3, len(df))):
                        row_data = []
                        for col in range(min(8, len(df.columns))):
                            value = df.iloc[k, col]
                            if pd.isna(value):
                                row_data.append("NaN")
                            else:
                                row_data.append(str(value))
                        print(f"행 {k:2d}: {row_data}")
                    
                    # CTR 데이터 분석
                    current_row = df.iloc[i].tolist()
                    print(f"\n=== CTR 분석 ===")
                    print(f"전체 행 데이터: {current_row}")
                    
                    # 노출수, 클릭수, CTR 위치 확인
                    for idx, val in enumerate(current_row):
                        if val and not pd.isna(val):
                            try:
                                num_val = float(val)
                                if 0 < num_val < 1:  # CTR 같은 비율 값
                                    print(f"  CTR 후보 (위치 {idx}): {val} ({num_val * 100:.2f}%)")
                                elif 1000 <= num_val <= 10000:  # 노출수 같은 값
                                    print(f"  노출수 후보 (위치 {idx}): {val}")
                                elif 1 <= num_val <= 1000:  # 클릭수 같은 값
                                    print(f"  클릭수 후보 (위치 {idx}): {val}")
                            except:
                                pass
                    print("-" * 40)

if __name__ == "__main__":
    file_path = "/Users/hongsukchoi/Desktop/cursor_project/01.live/cpc_monthly_report_웹통합_250817/0818_HTAG_CPC운영현황 보고서.xlsx"
    analyze_ctr_data(file_path)
