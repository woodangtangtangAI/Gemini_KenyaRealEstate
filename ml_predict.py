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
    print("🚀 [System] 3단계: 나이로비 부동산 가격 예측(XGBoost) 및 요인 분석을 시작합니다...")

    base_path = os.path.dirname(os.path.abspath(__file__))
    
    # 1. 마스터 데이터 자동 탐색
    file_list = []
    if os.path.exists(base_path):
        for file in os.listdir(base_path):
            if file.startswith("nairobi_master_data_") and file.endswith(".csv"):
                file_list.append(os.path.join(base_path, file))
    
    if not file_list:
        print("❌ [오류] 마스터 데이터를 찾을 수 없습니다. 2단계 병합이 완료되었는지 확인하세요.")
        return
        
    latest_file = max(file_list, key=os.path.getmtime)
    print(f"📁 타겟 확인: '{os.path.basename(latest_file)}' 데이터로 학습합니다.")
    
    # 데이터 로드
    df = pd.read_csv(latest_file)

    # 2. 데이터 전처리 (결측치 및 타입 정리)
    df['Scraped_Date'] = pd.to_datetime(df['Scraped_Date'])
    df = df.sort_values('Scraped_Date').reset_index(drop=True)

    numeric_cols = ['Price_KES', 'Bedrooms', 'Bathrooms', 'Size_sqm']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    df = df.dropna(subset=['Price_KES'])
    df = df.fillna(0)

    # ============================================================
    # 2.5 [신규] IQR 기반 아웃라이어 제거
    # ============================================================
    before_count = len(df)
    Q1 = df['Price_KES'].quantile(0.25)
    Q3 = df['Price_KES'].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    
    df = df[(df['Price_KES'] >= lower_bound) & (df['Price_KES'] <= upper_bound)]
    after_count = len(df)
    removed = before_count - after_count
    
    print(f"🧹 [아웃라이어 제거] IQR 방식 적용")
    print(f"   정상 범위: {lower_bound:,.0f} ~ {upper_bound:,.0f} KES")
    print(f"   제거 전: {before_count:,}건 → 제거 후: {after_count:,}건 (이상치 {removed:,}건 제외)")

    # ============================================================
    # 2.7 [신규] 지역별 통계 분석
    # ============================================================
    print("\n📍 [지역 간 분석] 나이로비 지역별 부동산 가격 통계...")
    
    if 'Location' in df.columns:
        regional = df.groupby('Location').agg(
            매물수=('Price_KES', 'count'),
            평균가=('Price_KES', 'mean'),
            중앙값=('Price_KES', 'median'),
            최저가=('Price_KES', 'min'),
            최고가=('Price_KES', 'max')
        ).sort_values('평균가', ascending=False)
        
        # 매물 5건 이상인 지역만 분석 대상
        regional_valid = regional[regional['매물수'] >= 5]
        
        print(f"   총 {len(regional_valid)}개 지역 분석 완료 (매물 5건 이상)")
        print(regional_valid.head(10).to_string())
        
        # 지역별 분석 결과를 JSON으로 저장 (send_report.py에서 사용)
        regional_dict = {}
        for loc, row in regional_valid.head(15).iterrows():
            regional_dict[loc] = {
                'count': int(row['매물수']),
                'avg_price': int(row['평균가']),
                'median_price': int(row['중앙값']),
                'min_price': int(row['최저가']),
                'max_price': int(row['최고가'])
            }
        
        regional_path = os.path.join(base_path, "regional_analysis.json")
        with open(regional_path, 'w', encoding='utf-8') as f:
            json.dump(regional_dict, f, ensure_ascii=False, indent=2)
        print(f"   ✅ 지역별 분석 결과 저장: regional_analysis.json")
        
        # 지역별 평균가 차트 생성
        if len(regional_valid) >= 3:
            top_regions = regional_valid.head(10)
            fig, ax = plt.subplots(figsize=(12, 6))
            bars = ax.barh(top_regions.index, top_regions['평균가'], color=sns.color_palette('viridis', len(top_regions)))
            ax.set_xlabel('Average Price (KES)', fontsize=12)
            ax.set_title('Top 10 Nairobi Neighborhoods by Average Property Price', fontsize=14)
            ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x:,.0f}'))
            plt.tight_layout()
            
            regional_img_path = os.path.join(base_path, "regional_price_comparison.png")
            plt.savefig(regional_img_path, dpi=300)
            plt.close()
            print(f"   ✅ 지역별 가격 비교 차트 저장: regional_price_comparison.png")

    # 3. 특성(X)과 타겟(y) 분리
    drop_cols = ['Price_KES', 'Scraped_Date', 'Property_ID']
    X = df.drop(columns=[col for col in drop_cols if col in df.columns])
    y = df['Price_KES']

    if 'Location' in X.columns:
        le = LabelEncoder()
        X['Location'] = le.fit_transform(X['Location'].astype(str))

    # Date 컬럼이 merge_asof로 생겼을 수 있으므로 제거
    if 'Date' in X.columns:
        X = X.drop(columns=['Date'])

    y_log = np.log1p(y)

    # 4. 시계열 데이터 분할
    print("\n⏳ 데이터 누수 방지를 위해 시간순(Chronological)으로 Train/Test 셋을 분리합니다...")
    X_train, X_test, y_train_log, y_test_log = train_test_split(
        X, y_log, test_size=0.2, shuffle=False
    )

    # 5. XGBoost 모델 학습
    print("🧠 XGBoost Regressor 모델 학습 중...")
    model = xgb.XGBRegressor(
        n_estimators=200,
        learning_rate=0.05,
        max_depth=6,
        random_state=42,
        objective='reg:squarederror'
    )
    
    model.fit(X_train, y_train_log)

    # 6. 예측 및 성능 평가
    pred_log = model.predict(X_test)
    y_test_actual = np.expm1(y_test_log)
    pred_actual = np.expm1(pred_log)

    rmse = np.sqrt(mean_squared_error(y_test_actual, pred_actual))
    r2 = r2_score(y_test_actual, pred_actual)

    print("\n📊 [모델 성능 평가 결과]")
    print(f" - RMSE (평균 오차): {rmse:,.0f} KES")
    print(f" - R² Score (설명력): {r2:.4f} (1.0에 가까울수록 완벽한 예측)")
    
    # 7. 변수 중요도 시각화
    print("\n📉 변수 중요도 분석 및 시각화 이미지 생성 중...")
    
    importances = model.feature_importances_
    features = X.columns
    
    importance_df = pd.DataFrame({'Feature': features, 'Importance': importances})
    importance_df = importance_df.sort_values(by='Importance', ascending=False).head(10)

    plt.figure(figsize=(10, 6))
    sns.barplot(x='Importance', y='Feature', data=importance_df, hue='Feature', palette='viridis', legend=False)
    plt.title('Top 10 Feature Importances for Nairobi Real Estate Prediction', fontsize=14)
    plt.xlabel('Relative Importance (XGBoost)', fontsize=12)
    plt.ylabel('Features', fontsize=12)
    plt.tight_layout()

    # 상위 5개 중요 변수를 텍스트 파일로 저장 (send_report.py 에서 프롬프트 주입용)
    top_features_text = ", ".join([f"{row['Feature']} ({row['Importance']:.3f})" for idx, row in importance_df.head(5).iterrows()])
    with open(os.path.join(base_path, "top_features.txt"), "w", encoding="utf-8") as f:
        f.write(top_features_text)

    img_filename = "feature_importance.png"
    img_path = os.path.join(base_path, img_filename)
    plt.savefig(img_path, dpi=300)
    plt.close()

    print(f"✅ [완료] 변수 중요도 그래프가 '{img_filename}'로 저장되었습니다.")

if __name__ == "__main__":
    train_real_estate_model()