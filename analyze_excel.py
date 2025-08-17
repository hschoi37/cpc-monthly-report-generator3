#!/usr/bin/env python3
import pandas as pd
import json
import sys

def analyze_excel_file(filename):
    print(f"\n=== {filename} 분석 ===")
    
    try:
        # 엑셀 파일 읽기
        df = pd.read_excel(filename, header=None)
        
        print(f"파일 크기: {df.shape[0]}행 x {df.shape[1]}열")
        print("\n처음 20행:")
        print(df.head(20).to_string())
        
        # 월 통계 관련 행 찾기
        monthly_stats = []
        for idx, row in df.iterrows():
            if any(str(cell).find('월 통계') != -1 for cell in row if pd.notna(cell)):
                monthly_stats.append({
                    'row_index': idx,
                    'data': row.tolist()
                })
                # 다음 5행도 확인
                for i in range(1, 6):
                    if idx + i < len(df):
                        next_row = df.iloc[idx + i]
                        monthly_stats.append({
                            'row_index': idx + i,
                            'data': next_row.tolist(),
                            'type': 'following'
                        })
        
        print(f"\n월 통계 관련 행들 ({len(monthly_stats)}개):")
        for stat in monthly_stats:
            print(f"행 {stat['row_index']}: {stat['data']}")
        
        # 주차별 데이터 찾기
        weekly_data = []
        for idx, row in df.iterrows():
            if any(str(cell).find('주차') != -1 for cell in row if pd.notna(cell)):
                weekly_data.append({
                    'row_index': idx,
                    'data': row.tolist()
                })
        
        print(f"\n주차별 데이터 ({len(weekly_data)}개):")
        for week in weekly_data:
            print(f"행 {week['row_index']}: {week['data']}")
        
        # CPC 관련 데이터 찾기
        cpc_data = []
        for idx, row in df.iterrows():
            if any(str(cell).find('CPC') != -1 for cell in row if pd.notna(cell)):
                cpc_data.append({
                    'row_index': idx,
                    'data': row.tolist()
                })
        
        print(f"\nCPC 관련 데이터 ({len(cpc_data)}개):")
        for cpc in cpc_data:
            print(f"행 {cpc['row_index']}: {cpc['data']}")
        
    except Exception as e:
        print(f"오류: {e}")

if __name__ == "__main__":
    files = ['dt_test.xlsx', 'ip_cpc.xlsx']
    
    for file in files:
        try:
            analyze_excel_file(file)
        except FileNotFoundError:
            print(f"파일을 찾을 수 없습니다: {file}")
        except Exception as e:
            print(f"파일 분석 중 오류 발생: {file} - {e}") 