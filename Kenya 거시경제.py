import os
import pandas as pd
from datetime import datetime

print("🚀 [2단계] 데이터 병합 및 마스터 파일 생성 프로세스를 시작합니다...")

base_path = os.path.dirname(os.path.abspath(__file__))

# 1. 1단계에서 수집한 일별 Raw 데이터 찾기
file_list = []
if os.path.exists(base_path):
    for f in os.listdir(base_path):
        if f.startswith("nairobi_raw_") and f.endswith(".csv"):
            file_list.append(os.path.join(base_path, f))

if not file_list:
    print("❌ 수집된 부동산 raw 데이터가 없습니다. 1단계 크롤러를 먼저 실행해주세요.")
    exit()

# 모든 csv 파일을 읽어서 하나의 거대한 데이터프레임으로 결합
df_list = [pd.read_csv(file) for file in file_list]
df_raw_combined = pd.concat(df_list, ignore_index=True)

# 2. 중복 매물 완벽 제거
if 'Property_ID' in df_raw_combined.columns:
    df_raw_clean = df_raw_combined.drop_duplicates(subset=['Property_ID'], keep='last')
else:
    df_raw_clean = df_raw_combined.drop_duplicates(keep='last')

df_raw_clean['Scraped_Date'] = pd.to_datetime(df_raw_clean['Scraped_Date'])
df_raw_clean = df_raw_clean.sort_values('Scraped_Date').reset_index(drop=True)

print(f"✅ 총 {len(file_list)}일 치 부동산 데이터를 병합했습니다. (순수 매물: {len(df_raw_clean)}건)")

# 3. 거시경제 데이터 불러오기 및 결합 (xlsx 우선, csv 폴백)
macro_xlsx = os.path.join(base_path, "kenya_macro_history.xlsx")
macro_csv = os.path.join(base_path, "kenya_macro.csv")

if os.path.exists(macro_xlsx):
    print("🔗 거시경제 엑셀(kenya_macro_history.xlsx)과 부동산 데이터를 결합 중...")
    df_macro = pd.read_excel(macro_xlsx, engine='openpyxl')
    df_macro['Date'] = pd.to_datetime(df_macro['Date'])
    df_macro = df_macro.sort_values('Date')
    
    df_master = pd.merge_asof(
        df_raw_clean, 
        df_macro, 
        left_on='Scraped_Date', 
        right_on='Date', 
        direction='backward'
    )
    print(f"  ✅ 거시경제 지표 결합 완료! (환율, 금리 등 {len(df_macro.columns)-1}개 지표)")
elif os.path.exists(macro_csv):
    print("🔗 거시경제 CSV(kenya_macro.csv)와 부동산 데이터를 결합 중...")
    df_macro = pd.read_csv(macro_csv)
    df_macro['Date'] = pd.to_datetime(df_macro['Date'])
    df_macro = df_macro.sort_values('Date')
    
    df_master = pd.merge_asof(
        df_raw_clean, 
        df_macro, 
        left_on='Scraped_Date', 
        right_on='Date', 
        direction='backward'
    )
    print(f"  ✅ 거시경제 지표 결합 완료!")
else:
    print(f"⚠️ 거시경제 파일을 찾을 수 없습니다. 부동산 데이터만으로 마스터를 생성합니다.")
    df_master = df_raw_clean

# 4. 최종 마스터 데이터 저장
today_date = datetime.now().strftime("%Y%m%d")
master_filename = f"nairobi_master_data_{today_date}.csv"
master_path = os.path.join(base_path, master_filename)

df_master.to_csv(master_path, index=False, encoding="utf-8-sig")
print(f"🎉 [최종 완료] 마스터 데이터가 '{master_filename}' 이름으로 저장되었습니다!")

# 마스터 데이터에 거시경제 컬럼 존재 여부 확인
macro_cols = ['USD_KES_Rate', 'CBR_Rate', 'CPI_Index']
found = [c for c in macro_cols if c in df_master.columns]
print(f"  📊 거시경제 컬럼 확인: {found if found else '없음 (거시경제 파일 누락)'}")