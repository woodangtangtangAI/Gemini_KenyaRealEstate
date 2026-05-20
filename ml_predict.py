import pandas as pd
import numpy as np
import os
import xgboost as xgb
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error, r2_score
from sklearn.preprocessing import LabelEncoder

def train_real_estate_model():
    print("🚀 [System] 3단계: 나이로비 부동산 가격 예측(XGBoost) 및 요인 분석을 시작합니다...")

    base_path = os.path.dirname(os.path.abspath(__file__))
    
    # 1. [수정됨] 대괄호 폴더명을 지원하는 마스터 데이터 자동 탐색 로직
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

    # 3. 특성(X)과 타겟(y) 분리
    drop_cols = ['Price_KES', 'Scraped_Date', 'Property_ID']
    X = df.drop(columns=[col for col in drop_cols if col in df.columns])
    y = df['Price_KES']

    if 'Location' in X.columns:
        le = LabelEncoder()
        X['Location'] = le.fit_transform(X['Location'].astype(str))

    y_log = np.log1p(y)

    # 4. 시계열 데이터 분할
    print("⏳ 데이터 누수 방지를 위해 시간순(Chronological)으로 Train/Test 셋을 분리합니다...")
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
    sns.barplot(x='Importance', y='Feature', data=importance_df, palette='viridis')
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

    print(f"✅ [완료] 변수 중요도 그래프가 '{img_filename}'로 구글 드라이브에 저장되었습니다.")

if __name__ == "__main__":
    train_real_estate_model()