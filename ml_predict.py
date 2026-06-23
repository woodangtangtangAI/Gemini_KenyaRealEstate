import sys
try: sys.stdout.reconfigure(encoding='utf-8')
except: pass

import pandas as pd
import numpy as np
import os
import json
import xgboost as xgb
import matplotlib
matplotlib.use('Agg')  # 클라우드 환경 대응
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error, r2_score
from sklearn.preprocessing import LabelEncoder

def train_real_estate_model():
    print("🚀 [System] 3단계: 나이로비 부동산 시장 다각도 분석 및 XGBoost 예측 프로세스를 가동합니다...")

    base_path = os.path.dirname(os.path.abspath(__file__))
    results_dir = os.path.join(base_path, "분석 결과")
    os.makedirs(results_dir, exist_ok=True)
    
    # 1. 마스터 데이터 자동 탐색
    file_list = []
    if os.path.exists(results_dir):
        for file in os.listdir(results_dir):
            if file.startswith("nairobi_master_data_") and file.endswith(".csv"):
                file_list.append(os.path.join(results_dir, file))
    
    if not file_list:
        print("❌ [오류] 마스터 데이터를 찾을 수 없습니다. 2단계 병합이 완료되었는지 확인하세요.")
        return
        
    latest_file = max(file_list, key=os.path.getmtime)
    print(f"📁 타겟 마스터 데이터: '{os.path.basename(latest_file)}'")
    
    df = pd.read_csv(latest_file)
    if df.empty:
        print("⚠️ 데이터가 비어 있습니다.")
        return

    # 2. 매매(Sale)와 월세(Rent) 분리 및 전처리
    if 'Listing_Type' not in df.columns:
        df['Listing_Type'] = 'sale'  # 폴백
        
    df['Scraped_Date'] = pd.to_datetime(df['Scraped_Date'])
    df = df.sort_values('Scraped_Date').reset_index(drop=True)

    # 2.1 시계열 특징 피처(Lags) 엔지니어링 (거시경제 변수 결합용)
    print("📈 [Feature Engineering] 거시경제 이동평균 및 변화율 피처 생성 중...")
    macro_xlsx = os.path.join(results_dir, "kenya_macro_history.xlsx")
    if os.path.exists(macro_xlsx):
        try:
            df_macro = pd.read_excel(macro_xlsx, engine='openpyxl')
            df_macro['Date'] = pd.to_datetime(df_macro['Date'])
            df_macro = df_macro.sort_values('Date').reset_index(drop=True)
            
            # 이동평균 및 변화율 계산 (기존 데이터가 충분히 쌓인 경우 유의미)
            df_macro['USD_KES_Rate_MA3'] = df_macro['USD_KES_Rate'].rolling(window=3, min_periods=1).mean()
            df_macro['USD_KES_Pct_Change'] = df_macro['USD_KES_Rate'].pct_change(fill_method=None).fillna(0)
            
            if 'Remittance_M_USD' in df_macro.columns:
                df_macro['Remittance_MA3'] = df_macro['Remittance_M_USD'].rolling(window=3, min_periods=1).mean()
                df_macro['Remittance_Pct_Change'] = df_macro['Remittance_M_USD'].pct_change(fill_method=None).fillna(0)
                
            # 피처들을 마스터 데이터에 재결합 (Date 기준)
            df = df.drop(columns=[col for col in ['USD_KES_Rate_MA3', 'USD_KES_Pct_Change', 'Remittance_MA3', 'Remittance_Pct_Change'] if col in df.columns])
            df = pd.merge_asof(
                df, 
                df_macro[['Date', 'USD_KES_Rate_MA3', 'USD_KES_Pct_Change', 'Remittance_MA3', 'Remittance_Pct_Change'] if 'Remittance_M_USD' in df_macro.columns else ['Date', 'USD_KES_Rate_MA3', 'USD_KES_Pct_Change']], 
                left_on='Scraped_Date', 
                right_on='Date', 
                direction='backward'
            )
            print("  ✅ 시계열 매크로 Lags 피처 결합 완료!")
        except Exception as e:
            print(f"  ⚠️ 매크로 Lags 피처 생성 실패: {e}")

    # 숫자형 컬럼 변환
    numeric_cols = ['Price_KES', 'Bedrooms', 'Bathrooms', 'Size_sqm', 'Price_per_sqm']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # 매물 분리
    df_sale = df[df['Listing_Type'] == 'sale'].copy()
    df_rent = df[df['Listing_Type'] == 'rent'].copy()

    print(f"📊 [매물 분류] 매매 매물: {len(df_sale)}건, 임대 매물: {len(df_rent)}건")

    # 3. 매매 데이터 통계 및 렌트 수익률 계산
    regional_dict = {"sales": {}, "rentals": {}, "yields": {}}
    
    if not df_sale.empty and 'Location' in df_sale.columns:
        # IQR 이상치 제거 (매매 전용)
        q1_s = df_sale['Price_KES'].quantile(0.25)
        q3_s = df_sale['Price_KES'].quantile(0.75)
        iqr_s = q3_s - q1_s
        df_sale_clean = df_sale[(df_sale['Price_KES'] >= (q1_s - 1.5 * iqr_s)) & (df_sale['Price_KES'] <= (q3_s + 1.5 * iqr_s))].copy()
        
        # ㎡당 단가 결측치 제거
        df_sale_clean = df_sale_clean.dropna(subset=['Price_per_sqm', 'Size_sqm'])
        df_sale_clean = df_sale_clean[df_sale_clean['Price_per_sqm'] > 0]
        
        sales_grouped = df_sale_clean.groupby('Location').agg(
            count=('Price_KES', 'count'),
            avg_price=('Price_KES', 'mean'),
            median_price=('Price_KES', 'median'),
            avg_price_sqm=('Price_per_sqm', 'mean'),
            median_price_sqm=('Price_per_sqm', 'median')
        )
        
        # 유효 지역 선정 (최소 매물 5건 이상)
        valid_sales = sales_grouped[sales_grouped['count'] >= 5].sort_values('avg_price_sqm', ascending=False)
        for loc, row in valid_sales.head(15).iterrows():
            regional_dict["sales"][loc] = {
                "count": int(row['count']),
                "avg_price": int(row['avg_price']),
                "median_price": int(row['median_price']),
                "avg_price_sqm": int(row['avg_price_sqm']),
                "median_price_sqm": int(row['median_price_sqm'])
            }

    if not df_rent.empty and 'Location' in df_rent.columns:
        # IQR 이상치 제거 (월세 전용)
        q1_r = df_rent['Price_KES'].quantile(0.25)
        q3_r = df_rent['Price_KES'].quantile(0.75)
        iqr_r = q3_r - q1_r
        df_rent_clean = df_rent[(df_rent['Price_KES'] >= (q1_r - 1.5 * iqr_r)) & (df_rent['Price_KES'] <= (q3_r + 1.5 * iqr_r))].copy()
        
        df_rent_clean = df_rent_clean.dropna(subset=['Price_KES'])
        
        rent_grouped = df_rent_clean.groupby('Location').agg(
            count=('Price_KES', 'count'),
            avg_rent=('Price_KES', 'mean'),
            median_rent=('Price_KES', 'median')
        )
        
        valid_rents = rent_grouped[rent_grouped['count'] >= 5]
        for loc, row in valid_rents.iterrows():
            regional_dict["rentals"][loc] = {
                "count": int(row['count']),
                "avg_rent": int(row['avg_rent']),
                "median_rent": int(row['median_rent'])
            }
            
        # 임대 수익률 산출 (매매 평균 가격과 연간 월세 수입 비교)
        if not valid_sales.empty:
            for loc in valid_sales.index:
                if loc in valid_rents.index:
                    avg_sale = valid_sales.loc[loc, 'avg_price']
                    avg_rent = valid_rents.loc[loc, 'avg_rent']
                    if avg_sale > 0:
                        annual_yield = (avg_rent * 12) / avg_sale
                        regional_dict["yields"][loc] = round(annual_yield, 4)

    # JSON 저장
    regional_path = os.path.join(results_dir, "regional_analysis.json")
    with open(regional_path, 'w', encoding='utf-8') as f:
        json.dump(regional_dict, f, ensure_ascii=False, indent=2)
    print("  ✅ 지역별 매매/월세/수익률 데이터 분석 결과 저장: regional_analysis.json")

    # 지역별 ㎡당 매매 가격 차트 시각화
    if "sales" in regional_dict and len(regional_dict["sales"]) >= 3:
        top_regions = pd.DataFrame.from_dict(regional_dict["sales"], orient='index').head(10)
        fig, ax = plt.subplots(figsize=(12, 6))
        sns.barplot(x=top_regions['avg_price_sqm'], y=top_regions.index, hue=top_regions.index, palette='viridis', legend=False, ax=ax)
        ax.set_xlabel('Average Price per SQM (KES)', fontsize=12)
        ax.set_title('Top 10 Nairobi Suburbs by Average Price per SQM', fontsize=14)
        ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x:,.0f}'))
        plt.tight_layout()
        
        regional_img_path = os.path.join(results_dir, "regional_price_comparison.png")
        plt.savefig(regional_img_path, dpi=300)
        plt.close()
        print("  ✅ 지역별 ㎡당 가격 비교 차트 저장: regional_price_comparison.png")

    # 4. XGBoost 모델링: 매매 ㎡당 가격 예측 모델링
    if not df_sale.empty and len(df_sale) > 10:
        print("\n🧠 [XGBoost 모델 구축] 매매 ㎡당 단가(Price_per_sqm) 예측 모델 학습을 개시합니다...")
        df_sale_model = df_sale.dropna(subset=['Price_per_sqm', 'Size_sqm', 'Location']).copy()
        
        drop_cols = ['Price_KES', 'Price_per_sqm', 'Scraped_Date', 'Property_ID', 'Listing_Type', 'Date']
        X = df_sale_model.drop(columns=[col for col in drop_cols if col in df_sale_model.columns])
        y = df_sale_model['Price_per_sqm']

        if 'Location' in X.columns:
            le = LabelEncoder()
            X['Location'] = le.fit_transform(X['Location'].astype(str))

        # 결측값 채우기
        X = X.fillna(0)
        
        # 시간순 분할
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, shuffle=False)

        model = xgb.XGBRegressor(
            n_estimators=150,
            learning_rate=0.08,
            max_depth=5,
            random_state=42,
            objective='reg:squarederror'
        )
        model.fit(X_train, y_train)

        pred = model.predict(X_test)
        rmse = np.sqrt(mean_squared_error(y_test, pred))
        r2 = r2_score(y_test, pred)

        print(f"  📊 매매 ㎡당 가격 예측 모델 평가:")
        print(f"   - RMSE: {rmse:,.0f} KES/sqm")
        print(f"   - R² Score: {r2:.4f}")

        # 변수 중요도 추출 및 저장
        importances = model.feature_importances_
        features = X.columns
        importance_df = pd.DataFrame({'Feature': features, 'Importance': importances}).sort_values('Importance', ascending=False)
        
        top_features_text = ", ".join([f"{row['Feature']} ({row['Importance']:.3f})" for idx, row in importance_df.head(5).iterrows()])
        with open(os.path.join(results_dir, "top_features.txt"), "w", encoding="utf-8") as f:
            f.write(top_features_text)
            
        # 변수 중요도 차트 그리기
        plt.figure(figsize=(10, 6))
        sns.barplot(x='Importance', y='Feature', data=importance_df.head(10), hue='Feature', palette='magma', legend=False)
        plt.title('Top Feature Importances (Price per SQM Model)', fontsize=14)
        plt.xlabel('Importance', fontsize=12)
        plt.tight_layout()
        
        img_path = os.path.join(results_dir, "feature_importance.png")
        plt.savefig(img_path, dpi=300)
        plt.close()
        print("  ✅ 모델 변수 중요도 차트 저장: feature_importance.png")
    else:
        print("\n⚠️ 매매 데이터가 부족하여 XGBoost 모델 학습을 생략합니다.")

if __name__ == "__main__":
    train_real_estate_model()