import os
import json
import smtplib
import pandas as pd
import requests
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
except ImportError:
    pass
from dotenv import load_dotenv

# 1. 환경 변수 로드
base_path = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(base_path, ".env"))

GENAI_API_KEY = os.getenv("GEMINI_API_KEY")
SENDER_EMAIL = os.getenv("MY_EMAIL")
SENDER_PW = os.getenv("MY_PASSWORD")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL")

def generate_and_send_report():
    print("🚀 AI 주간 리포트 '즉시 발행' 프로세스를 시작합니다...")
    
    base_path = os.path.dirname(os.path.abspath(__file__))
    
    # 2. 최신 마스터 데이터 로드
    file_list = [os.path.join(base_path, f) for f in os.listdir(base_path) if f.startswith("nairobi_master_data_") and f.endswith(".csv")]
    if not file_list:
        print("❌ 마스터 데이터를 찾을 수 없습니다.")
        return
        
    latest_file = max(file_list, key=os.path.getmtime)
    df = pd.read_csv(latest_file)
    latest_macro = df.iloc[-1]
    
    # 거시경제 데이터 안전하게 추출
    usd_rate = latest_macro.get('USD_KES_Rate', '데이터 없음') if 'USD_KES_Rate' in latest_macro.index else '데이터 없음'
    cbr_rate = latest_macro.get('CBR_Rate', '데이터 없음') if 'CBR_Rate' in latest_macro.index else '데이터 없음'
    
    # 거시경제 추세 분석 (엑셀에서 최근 데이터 읽기)
    macro_trend_text = ""
    macro_xlsx = os.path.join(base_path, "kenya_macro_history.xlsx")
    if os.path.exists(macro_xlsx):
        try:
            df_macro = pd.read_excel(macro_xlsx, engine='openpyxl')
            if len(df_macro) >= 2 and 'USD_KES_Rate' in df_macro.columns:
                recent = df_macro.tail(5)  # 최근 5주 데이터
                rates = recent['USD_KES_Rate'].dropna().tolist()
                dates = recent['Date'].astype(str).tolist()
                trend_pairs = [f"{d}: {r} KES" for d, r in zip(dates, rates)]
                macro_trend_text = f"\n    [환율 추세 (최근 {len(rates)}주)]\n    " + " → ".join(trend_pairs)
                
                if len(rates) >= 2:
                    change = rates[-1] - rates[0]
                    direction = "상승" if change > 0 else "하락"
                    macro_trend_text += f"\n    → 추세: {abs(change):.1f} KES {direction}"
        except Exception as e:
            print(f"  ⚠️ 거시경제 추세 분석 실패: {e}")
    
    # 머신러닝 상위 변수 가져오기
    top_features_path = os.path.join(base_path, "top_features.txt")
    top_features_data = "추출된 변수 데이터 없음"
    if os.path.exists(top_features_path):
        with open(top_features_path, "r", encoding="utf-8") as f:
            top_features_data = f.read().strip()
    
    # [신규] 지역별 분석 데이터 가져오기
    regional_text = ""
    regional_path = os.path.join(base_path, "regional_analysis.json")
    if os.path.exists(regional_path):
        try:
            with open(regional_path, 'r', encoding='utf-8') as f:
                regional_data = json.load(f)
            regional_lines = []
            for loc, stats in regional_data.items():
                regional_lines.append(
                    f"    - {loc}: 매물 {stats['count']}건, "
                    f"평균 {stats['avg_price']:,} KES, "
                    f"중앙값 {stats['median_price']:,} KES"
                )
            regional_text = "\n    [나이로비 지역별 가격 비교 (상위 15개 지역)]\n" + "\n".join(regional_lines)
        except Exception as e:
            print(f"  ⚠️ 지역별 분석 데이터 로드 실패: {e}")
            
    # 3. [고도화된 프롬프트] 퀀트 애널리스트 버전
    prompt = f"""
    당신은 철저하게 데이터에 기반하여 객관적 통찰을 제공하는 퀀트 애널리스트입니다.
    아래의 정량적 데이터를 분석하여 나이로비 부동산 주간 브리핑을 작성하세요.

    [정량적 시장 데이터]
    - 분석 대상 매물 수: {len(df):,}건
    - 시장 평균 호가: {df['Price_KES'].mean():,.0f} KES
    - 시장 중앙값: {df['Price_KES'].median():,.0f} KES
    - 시장 최고가: {df['Price_KES'].max():,.0f} KES / 최저가: {df['Price_KES'].min():,.0f} KES
    - 현재 환율: {usd_rate} KES/USD
    - 현재 기준금리(CBR): {cbr_rate}%
    - 집값 영향 핵심 변수 (XGBoost 분석 상위 5개): {top_features_data}
    {macro_trend_text}
    {regional_text}

    [객관적 분석 지시사항]
    1. 위 [정량적 시장 데이터]의 '구체적인 숫자'와 특히 **핵심 변수 5가지**를 적극 인용하여 어느 지역, 어떤 조건이 가격을 방어/주도하는지 깊이 있게 분석할 것.
    2. **지역별 가격 비교 데이터**를 활용하여 고가 지역 vs 저가 지역의 특성을 비교 분석하고, 투자 관점에서의 시사점을 제시할 것.
    3. 최근 거시경제(금리, 환율) 흐름과 결합했을 때의 Most Likely 시나리오를 투명하게 서술할 것.
    4. 다음 주 가장 주의 깊게 모니터링해야 할 지표와 그 이유를 제시할 것.
    5. 전체 내용을 서론-본론-결론의 명확하고 전문적인 리포트 형태로 작성할 것. (워드 문서로 출력될 예정이므로 소제목 등을 깔끔하게 구성)
    """
    
    # 4. 내 API 키로 사용 가능한 모델 자동 탐색
    print("🔍 내 API 키로 사용 가능한 AI 모델을 검색 중입니다...")
    target_model = "models/gemini-2.5-flash"
    
    try:
        models_url = f"https://generativelanguage.googleapis.com/v1beta/models?key={GENAI_API_KEY}"
        model_res = requests.get(models_url, timeout=10).json()
        
        available_models = [m['name'] for m in model_res.get('models', []) if 'generateContent' in m.get('supportedGenerationMethods', [])]
        
        # 우선순위: 2.5-flash > 1.5-pro > 아무 flash > 1.0-pro
        preferred = ["models/gemini-2.5-flash", "models/gemini-1.5-pro", "models/gemini-1.5-pro-latest"]
        for pm in preferred:
            if pm in available_models:
                target_model = pm
                break
        else:
            flash_models = [m for m in available_models if 'flash' in m]
            if flash_models:
                target_model = flash_models[0]
            
        print(f"✅ 사용 가능한 모델 확인 완료: {target_model}")
    except Exception as e:
        print(f"⚠️ 모델 검색 실패, 기본 모델로 시도합니다. ({e})")

    # 5. 선택된 모델과 다이렉트 통신
    print(f"🧠 {target_model} 모델과 직접 통신하여 분석을 요청합니다...")
    
    url = f"https://generativelanguage.googleapis.com/v1beta/{target_model}:generateContent?key={GENAI_API_KEY}"
    headers = {'Content-Type': 'application/json'}
    data = {
        "contents": [{"parts": [{"text": prompt}]}]
    }
    
    try:
        response = requests.post(url, headers=headers, json=data, timeout=120)
        response.raise_for_status()
        result = response.json()
        ai_insight = result['candidates'][0]['content']['parts'][0]['text']
        print("✅ 숫자 기반의 AI 분석을 성공적으로 받아왔습니다!")
    except Exception as e:
        resp_text = ''
        if 'response' in locals():
            resp_text = response.text[:500]
        print(f"❌ AI 분석 실패: {e}")
        print(f"   상세 응답: {resp_text}")
        ai_insight = f"[AI 자동 분석 일시 장애] 다음 주에 자동 복구됩니다.\n오류 내용: {e}"

    # 5.5 Word 파일 생성 및 저장
    current_date = datetime.now().strftime('%Y%m%d')
    docx_filename = f"나이로비_부동산_주간_브리핑_{current_date}.docx"
    docx_path = os.path.join(base_path, docx_filename)
    
    print(f"📝 워드 문서 생성을 시작합니다: {docx_filename}")
    try:
        doc = Document()
        
        # 제목
        title = doc.add_heading(f"나이로비 부동산 퀀트 분석 브리핑 ({datetime.now().strftime('%Y-%m-%d')})", 0)
        title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER if 'WD_PARAGRAPH_ALIGNMENT' in globals() else 1
        
        # 기본 지표 추가
        doc.add_heading('시장 주요 지표 요약', level=1)
        p1 = doc.add_paragraph()
        p1.add_run(f"분석 대상 매물 수: {len(df):,}건\n").bold = True
        p1.add_run(f"시장 평균 호가: {df['Price_KES'].mean():,.0f} KES\n")
        p1.add_run(f"시장 중앙값: {df['Price_KES'].median():,.0f} KES\n")
        p1.add_run(f"환율: {usd_rate} KES/USD | 금리: {cbr_rate}%\n")
        
        # 지역별 분석 섹션
        if regional_text:
            doc.add_heading('나이로비 지역별 가격 분석', level=1)
            if os.path.exists(regional_path):
                for loc, stats in regional_data.items():
                    p = doc.add_paragraph()
                    p.add_run(f"{loc}").bold = True
                    p.add_run(f": 매물 {stats['count']}건, 평균 {stats['avg_price']:,} KES, 중앙값 {stats['median_price']:,} KES")
        
        # AI 분석 추가
        doc.add_heading('AI 퀀트 애널리스트 Insight', level=1)
        for line in ai_insight.split('\n'):
            if line.strip().startswith('#'):
                doc.add_heading(line.replace('#', '').strip(), level=2)
            elif line.strip():
                doc.add_paragraph(line.strip())
                
        # 중요 변수
        doc.add_heading('주요 변수 영향도 (XGBoost)', level=1)
        doc.add_paragraph(f"핵심 변수: {top_features_data}")
        
        doc.save(docx_path)
        print(f"✅ 워드 리포트 저장 완료: {docx_path}")
    except Exception as e:
        print(f"⚠️ 워드 리포트 생성 실패 (python-docx가 설치되어 있는지 확인하세요): {e}")

    print("📧 이메일을 발송합니다...")
    
    # 6. 이메일 템플릿
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    # 지역별 분석 HTML 테이블 생성
    regional_html = ""
    if os.path.exists(regional_path):
        try:
            regional_html = '<table style="width:100%; border-collapse:collapse; margin:15px 0;">'
            regional_html += '<tr style="background:#004a99; color:white;"><th style="padding:8px; border:1px solid #ddd;">지역</th><th style="padding:8px; border:1px solid #ddd;">매물수</th><th style="padding:8px; border:1px solid #ddd;">평균가 (KES)</th><th style="padding:8px; border:1px solid #ddd;">중앙값 (KES)</th></tr>'
            for i, (loc, stats) in enumerate(regional_data.items()):
                bg = '#f9f9f9' if i % 2 == 0 else '#ffffff'
                regional_html += f'<tr style="background:{bg};"><td style="padding:6px; border:1px solid #ddd;">{loc}</td><td style="padding:6px; border:1px solid #ddd; text-align:center;">{stats["count"]}</td><td style="padding:6px; border:1px solid #ddd; text-align:right;">{stats["avg_price"]:,}</td><td style="padding:6px; border:1px solid #ddd; text-align:right;">{stats["median_price"]:,}</td></tr>'
            regional_html += '</table>'
        except:
            regional_html = ""
    
    html_body = f"""
    <html>
    <body style="font-family: 'Malgun Gothic', sans-serif; color: #333; line-height: 1.6;">
        <h2 style="color: #004a99; border-bottom: 2px solid #004a99; padding-bottom: 5px;">📊 나이로비 부동산 시장 AI 브리핑</h2>
        <p><strong>발행일시:</strong> {current_time}</p>
        <div style="background: #f9f9f9; padding: 15px; border-radius: 5px; margin-bottom: 20px;">
            <p style="margin: 0;"><strong>데이터 기준 환율:</strong> {usd_rate} KES/USD | <strong>금리:</strong> {cbr_rate}%</p>
            <p style="margin: 5px 0 0 0;"><strong>분석 매물:</strong> {len(df):,}건 | <strong>평균 호가:</strong> {df['Price_KES'].mean():,.0f} KES</p>
        </div>
        <h3 style="margin-bottom: 10px;">📍 지역별 가격 비교</h3>
        {regional_html}
        <h3 style="margin-bottom: 10px;">🧠 퀀트 애널리스트 Insight</h3>
        <div style="white-space: pre-wrap; background: #ffffff; border-left: 4px solid #004a99; padding: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">{ai_insight}</div>
        <h3 style="margin-top: 30px;">🔍 주요 변수 영향도 (XGBoost 모델, 하단 첨부파일 참조)</h3>
    </body>
    </html>
    """

    msg = MIMEMultipart()
    msg['Subject'] = f"[{datetime.now().strftime('%m/%d %H:%M:%S')}] 나이로비 부동산 퀀트 분석 리포트"
    msg['From'] = SENDER_EMAIL
    
    recipients = [RECEIVER_EMAIL, "donghyun1.kwon@lge.com"]
    msg['To'] = ", ".join([r for r in recipients if r])
    
    msg.attach(MIMEText(html_body, 'html'))

    # 7. 이미지 첨부 (변수 중요도 + 지역별 비교)
    for img_name in ["feature_importance.png", "regional_price_comparison.png"]:
        img_path = os.path.join(base_path, img_name)
        if os.path.exists(img_path):
            with open(img_path, 'rb') as f:
                msg.attach(MIMEImage(f.read(), name=img_name))
            
    # Word 파일 첨부
    if 'docx_path' in locals() and os.path.exists(docx_path):
        with open(docx_path, 'rb') as f:
            part = MIMEApplication(f.read(), Name=docx_filename)
            part['Content-Disposition'] = f'attachment; filename="{docx_filename}"'
            msg.attach(part)

    # 8. 이메일 발송
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(SENDER_EMAIL, SENDER_PW)
            server.send_message(msg, from_addr=SENDER_EMAIL, to_addrs=recipients)
        print("✅ [최종 성공] 이메일 발송 완료! 메일함을 확인해 주세요.")
    except Exception as e:
        print(f"❌ [실패] 메일 서버 에러: {e}")
        
    try:
        input("\n엔터 키를 누르면 창이 닫힙니다...")
    except EOFError:
        pass

if __name__ == "__main__":
    generate_and_send_report()