"""
[신규] fetch_macro.py - 거시경제 데이터 실시간 수집 및 단일 엑셀 축적
매주 파이프라인 실행 시 가장 먼저 실행됩니다.
"""
import os
import pandas as pd
import requests
from datetime import datetime

def fetch_and_accumulate_macro():
    print("📊 [0단계] 거시경제 데이터 실시간 수집 및 축적을 시작합니다...")
    
    base_path = os.path.dirname(os.path.abspath(__file__))
    xlsx_path = os.path.join(base_path, "kenya_macro_history.xlsx")
    today = datetime.now().strftime("%Y-%m-%d")
    
    # 1. 기존 엑셀 파일 읽기 (없으면 빈 DataFrame 생성)
    if os.path.exists(xlsx_path):
        df_macro = pd.read_excel(xlsx_path, engine='openpyxl')
        print(f"  📂 기존 데이터 로드 완료: {len(df_macro)}행")
    else:
        df_macro = pd.DataFrame(columns=['Date', 'USD_KES_Rate', 'CBR_Rate', 'Remittance_M_USD', 'CPI_Index'])
        print("  📂 기존 데이터 없음. 새 엑셀 파일을 생성합니다.")
    
    # 2. 이번 주 데이터가 이미 있으면 중복 방지 (같은 날짜 스킵)
    if 'Date' in df_macro.columns and not df_macro.empty:
        df_macro['Date'] = pd.to_datetime(df_macro['Date']).dt.strftime('%Y-%m-%d')
        if today in df_macro['Date'].values:
            print(f"  ⏭️ 오늘({today}) 데이터가 이미 존재합니다. 수집을 건너뜁니다.")
            return
    
    # 3. 무료 환율 API에서 USD/KES 실시간 수집 (인증키 불필요)
    usd_kes_rate = None
    print("  🌐 환율 API(open.er-api.com)에서 USD/KES 환율을 수집 중...")
    try:
        r = requests.get("https://open.er-api.com/v6/latest/USD", timeout=15)
        r.raise_for_status()
        rates = r.json().get("rates", {})
        usd_kes_rate = rates.get("KES")
        if usd_kes_rate:
            print(f"  ✅ 실시간 환율 수집 성공: 1 USD = {usd_kes_rate} KES")
        else:
            print("  ⚠️ KES 환율을 찾을 수 없습니다.")
    except Exception as e:
        print(f"  ⚠️ 환율 API 호출 실패: {e}")
    
    # 4. CBK 기준금리 (Central Bank Rate)
    # CBK는 공개 API가 없어 최신 발표 기준으로 유지합니다.
    # 2026년 4월 기준 CBK Rate: 10.00% (2024년 하반기부터 연속 인하 추세)
    cbr_rate = 10.0
    print(f"  📌 CBK 기준금리: {cbr_rate}% (최신 발표 기준)")
    
    # 5. CPI 및 송금 데이터 (직전 값 유지)
    cpi_index = None
    remittance = None
    if not df_macro.empty:
        last_row = df_macro.iloc[-1]
        cpi_index = last_row.get('CPI_Index', None)
        remittance = last_row.get('Remittance_M_USD', None)
    
    # 6. 새 행 추가
    new_row = {
        'Date': today,
        'USD_KES_Rate': usd_kes_rate,
        'CBR_Rate': cbr_rate,
        'Remittance_M_USD': remittance,
        'CPI_Index': cpi_index
    }
    
    df_macro = pd.concat([df_macro, pd.DataFrame([new_row])], ignore_index=True)
    
    # 7. 같은 파일에 덮어쓰기 (단일 파일 유지)
    df_macro.to_excel(xlsx_path, index=False, engine='openpyxl')
    print(f"  ✅ 거시경제 데이터 축적 완료! 현재 총 {len(df_macro)}행 ({xlsx_path})")
    print(f"     최신 행: {new_row}")

if __name__ == "__main__":
    fetch_and_accumulate_macro()
