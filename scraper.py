import sys
try: sys.stdout.reconfigure(encoding='utf-8')
except: pass

import time
import random
import pandas as pd
from bs4 import BeautifulSoup
import os
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

def get_nairobi_data_pagination(max_pages=265): 
    print(f"🚀 [System] 페이지 자동 이동(1~{max_pages}페이지) 수집을 시작합니다...")

    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(options=chrome_options)
    
    raw_data_list = []
    scraped_date = pd.Timestamp.now().strftime('%Y-%m-%d')

    try:
        for page in range(1, max_pages + 1):
            print(f"📄 [진행중] {page}페이지로 이동 중...")
            url = f"https://www.buyrentkenya.com/flats-apartments-for-sale/nairobi?page={page}"
            driver.get(url)
            time.sleep(random.uniform(4.0, 7.0))
            
            html_source = driver.page_source
            soup = BeautifulSoup(html_source, "html.parser")

            # BuyRentKenya 구조 변경에 대비해 <a> 태그(링크) 안의 내용도 추출 대상으로 포함
            listings = soup.find_all(["div", "article", "a"], class_=re.compile("listing|card|relative|property"))
            print(f"   🔎 {page}페이지에서 매물 {len(listings)}개 발견 및 추출 중...")

            for item in listings:
                # 매물별 고유 링크(URL)를 추출하여 Property_ID(고유 식별자)로 사용 (중복 제거의 핵심)
                href = item.get('href', '')
                if href and href.startswith('/'):
                    property_id = "https://www.buyrentkenya.com" + href
                elif href and href.startswith('http'):
                    property_id = href
                else:
                    # 링크가 없으면 글자 전체를 합쳐서 가상의 ID로 만듦
                    property_id = item.get_text(strip=True)[:50]

                text_content = item.get_text(separator='|', strip=True)
                parts = [p.strip() for p in text_content.split('|') if len(p) > 1]
                if len(parts) < 3: continue

                price, loc, bed, bath, size = "NaN", "Nairobi", "NaN", "NaN", "NaN"

                for p in parts:
                    if 'KSh' in p or 'KES' in p:
                        price_numeric = re.sub(r'[^0-9]', '', p)
                        price = price_numeric if price_numeric else p
                    elif 'Bed' in p: bed = re.sub(r'[^0-9]', '', p)
                    elif 'Bath' in p: bath = re.sub(r'[^0-9]', '', p)
                    elif 'm²' in p or 'sqm' in p: size = re.sub(r'[^0-9.]', '', p)
                    elif any(area in p for area in [
                        'Westlands', 'Riverside', 'Kilimani', 'Lavington', 'Kileleshwa',
                        'Karen', 'Runda', 'Muthaiga', 'Spring Valley', 'Loresho',
                        'Langata', 'South B', 'South C', 'Nairobi West', 'Upper Hill',
                        'Upperhill', 'Embakasi', 'Kasarani', 'Roysambu', 'Ruaka',
                        'Parklands', 'Ngara', 'Eastleigh', 'Donholm', 'Umoja',
                        'Kahawa', 'Syokimau', 'Athi River', 'Kitengela', 'Ongata Rongai',
                        'Gigiri', 'Garden Estate', 'Thika Road', 'Hurlingham',
                        'Ngong Road', 'Dagoretti', 'Zimmerman', 'Mihango',
                        'Ruiru', 'Juja', 'Kikuyu', 'Mlolongo', 'Pipeline'
                    ]):
                        loc = p

                if price != "NaN":
                    raw_data_list.append({
                        "Property_ID": property_id, # 나중에 2단계에서 중복을 완벽히 걸러낼 고유 번호
                        "Price_KES": price,
                        "Location": loc,
                        "Bedrooms": bed,
                        "Bathrooms": bath,
                        "Size_sqm": size,
                        "Scraped_Date": scraped_date
                    })

    except Exception as e:
        print(f"❌ 에러 발생: {e}")

    # 결과 저장 로직 (일별 적재 방식으로 변경)
    df = pd.DataFrame(raw_data_list)
    
    # 1차적으로 오늘 수집된 데이터 안에서도 중복 매물(URL 기준) 제거
    if not df.empty and 'Property_ID' in df.columns:
        df = df.drop_duplicates(subset=['Property_ID'])

    base_path = os.path.dirname(os.path.abspath(__file__))
    save_path = os.path.join(base_path, "raw_data")
    os.makedirs(save_path, exist_ok=True) # 폴더가 없으면 에러 안 내고 생성
    
    # 초 단위 지저분한 파일명 대신, 깔끔하게 'nairobi_raw_20260426.csv' 형태로 고정
    today_date = datetime.now().strftime("%Y%m%d")
    filename = f"nairobi_raw_{today_date}.csv"
    full_path = os.path.join(save_path, filename)

    if not df.empty:
        df.to_csv(full_path, index=False, encoding="utf-8-sig")
        print(f"\n✅ [성공] 총 {len(df)}개의 순수 매물을 안전하게 적재했습니다!")
        print(f"📁 생성/덮어쓰기 된 파일: {filename}")
    else:
        print("\n⚠️ 수집된 데이터가 없습니다.")

if __name__ == "__main__":
    get_nairobi_data_pagination(max_pages=265)