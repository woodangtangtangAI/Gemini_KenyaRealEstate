import sys
try: sys.stdout.reconfigure(encoding='utf-8')
except: pass

"""
[개편] fetch_macro.py - 거시경제 데이터 실시간 수집 및 단일 엑셀 축적
케냐 중앙은행(CBK)에서 기준금리(CBR) 및 해외송금액(Remittances)을 동적으로 수집합니다.
"""
import os
import re
import pandas as pd
import requests
import urllib.request
from bs4 import BeautifulSoup
from datetime import datetime

def fetch_and_accumulate_macro():
    print("📊 [0단계] 거시경제 데이터 실시간 수집 및 축적을 시작합니다...")
    
    base_path = os.path.dirname(os.path.abspath(__file__))
    xlsx_dir = os.path.join(base_path, "분석 결과")
    os.makedirs(xlsx_dir, exist_ok=True)
    xlsx_path = os.path.join(xlsx_dir, "kenya_macro_history.xlsx")
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
            
    # 3. KES/USD 실시간 환율 수집
    usd_kes_rate = None
    print("  🌐 환율 API(open.er-api.com)에서 USD/KES 환율을 수집 중...")
    try:
        r = requests.get("https://open.er-api.com/v6/latest/USD", timeout=15)
        r.raise_for_status()
        rates = r.json().get("rates", {})
        usd_kes_rate = rates.get("KES")
        if usd_kes_rate:
            print(f"  ✅ 실시간 환율 수집 성공: 1 USD = {usd_kes_rate:.2f} KES")
        else:
            print("  ⚠️ KES 환율을 찾을 수 없습니다.")
    except Exception as e:
        print(f"  ⚠️ 환율 API 호출 실패: {e}")
        
    # 4. CBK 기준금리 (CBR) 동적 수집
    cbr_rate = 10.0
    print("  🌐 케냐 중앙은행(CBK) 홈페이지에서 실시간 기준금리(CBR) 조회 중...")
    try:
        req = urllib.request.Request('https://www.centralbank.go.ke/', headers={'User-Agent': 'Mozilla/5.0'})
        html = urllib.request.urlopen(req, timeout=15).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        text = soup.get_text()
        
        cbr_match = re.search(r'CBR\s+at\s+(\d+(?:\.\d+)?)\s*percent', text, re.IGNORECASE)
        if not cbr_match:
            cbr_match = re.search(r'Central\s+Bank\s+Rate\s*\(CBR\)\s*at\s+(\d+(?:\.\d+)?)\s*percent', text, re.IGNORECASE)
        if not cbr_match:
            cbr_match = re.search(r'CBR\s*:\s*(\d+(?:\.\d+)?)\s*%', text, re.IGNORECASE)
        if not cbr_match:
            cbr_match = re.search(r'Central\s+Bank\s+Rate\s*:\s*(\d+(?:\.\d+)?)\s*%', text, re.IGNORECASE)
            
        if cbr_match:
            cbr_rate = float(cbr_match.group(1))
            print(f"  ✅ 기준금리 수집 성공: {cbr_rate}%")
        else:
            print("  ⚠️ 기준금리 매칭 실패, 이전 값 또는 기본값을 유지합니다.")
            if not df_macro.empty and 'CBR_Rate' in df_macro.columns:
                cbr_rate = df_macro.iloc[-1]['CBR_Rate']
    except Exception as e:
        print(f"  ⚠️ 기준금리 수집 실패: {e}")
        if not df_macro.empty and 'CBR_Rate' in df_macro.columns:
            cbr_rate = df_macro.iloc[-1]['CBR_Rate']
            
    # 5. Diaspora Remittances (해외송금액) 수집
    remittance_m_usd = None
    print("  🌐 케냐 중앙은행(CBK)에서 해외 송금액 데이터 수집 중...")
    try:
        req = urllib.request.Request('https://www.centralbank.go.ke/diaspora-remittances/', headers={'User-Agent': 'Mozilla/5.0'})
        html = urllib.request.urlopen(req, timeout=15).read()
        soup = BeautifulSoup(html, 'html.parser')
        table = soup.find('table')
        rows = table.find_all('tr')
        
        data_rows = []
        for row in rows:
            cells = [td.get_text(strip=True) for td in row.find_all(['td', 'th'])]
            if cells and len(cells) >= 6 and cells[0] != 'Year':
                data_rows.append(cells)
                
        if data_rows:
            latest_row = data_rows[-1]
            total_str = latest_row[5].replace(',', '')
            total_val = float(total_str)
            remittance_m_usd = total_val / 1000.0  # Convert thousands to Millions
            print(f"  ✅ 해외 송금액 수집 성공: {remittance_m_usd:.2f} Million USD ({latest_row[0]}-{latest_row[1]})")
        else:
            print("  ⚠️ 송금액 데이터 파싱 실패, 이전 값을 사용합니다.")
            if not df_macro.empty and 'Remittance_M_USD' in df_macro.columns:
                remittance_m_usd = df_macro.iloc[-1]['Remittance_M_USD']
    except Exception as e:
        print(f"  ⚠️ 해외 송금액 수집 실패: {e}")
        if not df_macro.empty and 'Remittance_M_USD' in df_macro.columns:
            remittance_m_usd = df_macro.iloc[-1]['Remittance_M_USD']
            
    # 6. CPI 인덱스 (이전 값 유지)
    cpi_index = None
    if not df_macro.empty and 'CPI_Index' in df_macro.columns:
        cpi_index = df_macro.iloc[-1]['CPI_Index']
        
    # 7. 새 행 추가 및 저장
    new_row = {
        'Date': today,
        'USD_KES_Rate': usd_kes_rate,
        'CBR_Rate': cbr_rate,
        'Remittance_M_USD': remittance_m_usd,
        'CPI_Index': cpi_index
    }
    
    df_macro = pd.concat([df_macro, pd.DataFrame([new_row])], ignore_index=True)
    df_macro.to_excel(xlsx_path, index=False, engine='openpyxl')
    print(f"  ✅ 거시경제 데이터 누적 완료: {xlsx_path}")
    print(f"     금주 데이터: {new_row}")

if __name__ == "__main__":
    fetch_and_accumulate_macro()
