import sys
try: sys.stdout.reconfigure(encoding='utf-8')
except: pass

import time
import random
import pandas as pd
from bs4 import BeautifulSoup
import os
import re
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

def scrape_listings_by_type(driver, listing_type, max_pages):
    print(f"🚀 [Scraper] {listing_type} 수집을 시작합니다 (최대 {max_pages}페이지)...")
    raw_data_list = []
    scraped_date = pd.Timestamp.now().strftime('%Y-%m-%d')

    try:
        for page in range(1, max_pages + 1):
            print(f"📄 [{listing_type}] {page}페이지로 이동 중...")
            url = f"https://www.buyrentkenya.com/{listing_type}/nairobi?page={page}"
            driver.get(url)
            time.sleep(random.uniform(4.0, 7.0))
            
            html_source = driver.page_source
            soup = BeautifulSoup(html_source, "html.parser")

            listings = soup.find_all(["div", "article", "a"], class_=re.compile("listing|card|relative|property"))
            print(f"   🔎 {page}페이지에서 매물 {len(listings)}개 발견 및 추출 중...")

            for item in listings:
                href = item.get('href', '')
                if href and href.startswith('/'):
                    property_id = "https://www.buyrentkenya.com" + href
                elif href and href.startswith('http'):
                    property_id = href
                else:
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
                    # Convert values to clean types
                    try:
                        price_val = float(price)
                    except:
                        price_val = None
                        
                    try:
                        size_val = float(size)
                    except:
                        size_val = None
                        
                    price_per_sqm = None
                    if price_val and size_val and size_val > 0:
                        price_per_sqm = price_val / size_val

                    raw_data_list.append({
                        "Property_ID": property_id,
                        "Listing_Type": 'sale' if 'sale' in listing_type else 'rent',
                        "Price_KES": price_val,
                        "Location": loc,
                        "Bedrooms": float(bed) if bed != "NaN" else None,
                        "Bathrooms": float(bath) if bath != "NaN" else None,
                        "Size_sqm": size_val,
                        "Price_per_sqm": price_per_sqm,
                        "Scraped_Date": scraped_date
                    })
    except Exception as e:
        print(f"❌ 에러 발생 ({listing_type}): {e}")

    return raw_data_list

def scrape_kenya_real_estate_news():
    print("📰 [News] Google News RSS에서 케냐 부동산/개발 정책 관련 뉴스 수집 중...")
    import json
    
    base_path = os.path.dirname(os.path.abspath(__file__))
    results_dir = os.path.join(base_path, "분석 결과")
    os.makedirs(results_dir, exist_ok=True)
    news_json_path = os.path.join(results_dir, "scraped_news.json")
    
    try:
        # Search query on Google News (localized to Kenya)
        query = 'Nairobi+real+estate+OR+Nairobi+zoning+OR+Nairobi+infrastructure'
        url = f'https://news.google.com/rss/search?q={query}&hl=en-KE&gl=KE&ceid=KE:en'
        
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        xml_data = urllib.request.urlopen(req, timeout=15).read()
        
        root = ET.fromstring(xml_data)
        news_items = []
        
        for item in root.findall('.//item')[:10]: # Collect top 10 articles
            title = item.find('title').text
            link = item.find('link').text
            pub_date = item.find('pubDate').text
            source = item.find('source').text if item.find('source') is not None else "Unknown"
            
            news_items.append({
                "title": title,
                "link": link,
                "pubDate": pub_date,
                "source": source
            })
            
        with open(news_json_path, 'w', encoding='utf-8') as f:
            json.dump(news_items, f, ensure_ascii=False, indent=2)
        print(f"  ✅ 실시간 기사 {len(news_items)}개 수집 및 저장 완료: scraped_news.json")
        
    except Exception as e:
        print(f"  ⚠️ 뉴스 수집 실패: {e}")
        # Save an empty list or keep existing
        if not os.path.exists(news_json_path):
            with open(news_json_path, 'w', encoding='utf-8') as f:
                json.dump([], f)

def get_nairobi_data_pagination(max_pages_sale=100, max_pages_rent=30): 
    print(f"🚀 [System] 케냐 부동산 수집 통합 파이프라인을 기동합니다...")
    
    # 1. 뉴스 데이터 수집
    scrape_kenya_real_estate_news()

    # 2. 부동산 매물 수집 (Selenium headless 드라이버 설정)
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(options=chrome_options)
    
    combined_data = []
    
    try:
        # 매매(Sale) 수집
        sale_data = scrape_listings_by_type(driver, "flats-apartments-for-sale", max_pages_sale)
        combined_data.extend(sale_data)
        
        # 임대(Rent) 수집
        rent_data = scrape_listings_by_type(driver, "flats-apartments-for-rent", max_pages_rent)
        combined_data.extend(rent_data)
        
    finally:
        driver.quit()
        
    # 결과 처리 및 저장
    df = pd.DataFrame(combined_data)
    
    # 중복 제거 (Property_ID 기준)
    if not df.empty and 'Property_ID' in df.columns:
        df = df.drop_duplicates(subset=['Property_ID'])

    base_path = os.path.dirname(os.path.abspath(__file__))
    save_path = os.path.join(base_path, "raw_data")
    os.makedirs(save_path, exist_ok=True)
    
    today_date = datetime.now().strftime("%Y%m%d")
    filename = f"nairobi_raw_{today_date}.csv"
    full_path = os.path.join(save_path, filename)

    if not df.empty:
        df.to_csv(full_path, index=False, encoding="utf-8-sig")
        sales_cnt = len(df[df['Listing_Type'] == 'sale'])
        rents_cnt = len(df[df['Listing_Type'] == 'rent'])
        print(f"\n✅ [성공] 총 {len(df)}개의 순수 매물을 수집 완료했습니다! (매매: {sales_cnt}건, 월세: {rents_cnt}건)")
        print(f"📁 생성된 파일: {filename} ({full_path})")
    else:
        print("\n⚠️ 수집된 부동산 매물 데이터가 없습니다.")

if __name__ == "__main__":
    # GitHub Actions 등 클라우드나 로컬 스케줄러 환경에서는 실행 시간 제약을 방지하기 위해 
    # 매매 100페이지, 월세 30페이지 수준으로 한정해 안전하고 빠르게 구동시킵니다.
    get_nairobi_data_pagination(max_pages_sale=100, max_pages_rent=30)