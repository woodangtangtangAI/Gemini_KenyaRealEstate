import os
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
        input("엔터 키를 누르면 종료됩니다...")
        return
        
    latest_file = max(file_list, key=os.path.getmtime)
    df = pd.read_csv(latest_file)
    latest_macro = df.iloc[-1]
    
    usd_rate = latest_macro.get('USD_KES_Rate', '데이터 없음') if 'USD_KES_Rate' in latest_macro.index else '데이터 없음'
    cbr_rate = latest_macro.get('CBR_Rate', '데이터 없음') if 'CBR_Rate' in latest_macro.index else '데이터 없음'
    
    # [추가] 머신러닝 상위 변수 가져오기
    top_features_path = os.path.join(base_path, "top_features.txt")
    top_features_data = "추출된 변수 데이터 없음"
    if os.path.exists(top_features_path):
        with open(top_features_path, "r", encoding="utf-8") as f:
            top_features_data = f.read().strip()
            
    # 3. [고도화된 프롬프트] 퀀트 애널리스트 버전 (구체적 수치 강제 적용)
    prompt = f"""
    당신은 철저하게 데이터에 기반하여 객관적 통찰을 제공하는 퀀트 애널리스트입니다.
    아래의 정량적 데이터를 분석하여 나이로비 부동산 주간 브리핑을 작성하세요.

    [정량적 시장 데이터]
    - 분석 대상 매물 수: {len(df):,}건
    - 시장 평균 호가: {df['Price_KES'].mean():,.0f} KES
    - 시장 최고가: {df['Price_KES'].max():,.0f} KES / 최저가: {df['Price_KES'].min():,.0f} KES
    - 현재 환율: {usd_rate}
    - 현재 기준금리: {cbr_rate}
    - 집값 영향 핵심 변수 (XGBoost 분석 상위 5개): {top_features_data}

    [객관적 분석 지시사항]
    1. 위 [정량적 시장 데이터]의 '구체적인 숫자'와 특히 **핵심 변수 5가지**를 적극 인용하여 어느 지역, 어떤 조건이 가격을 방어/주도하는지 깊이 있게 분석할 것.
    2. 최근 거시경제(금리, 환율) 흐름과 결합했을 때의 Most Likely 시나리오를 투명하게 서술할 것.
    3. 다음 주 가장 주의 깊게 모니터링해야 할 지표와 그 이유를 제시할 것.
    4. 전체 내용을 서론-본론-결론의 명확하고 전문적인 리포트 형태로 작성할 것. (워드 문서로 출력될 예정이므로 소제목 등을 깔끔하게 구성)
    """
    
    # 4. 내 API 키로 사용 가능한 모델 자동 탐색
    print("🔍 내 API 키로 사용 가능한 AI 모델을 검색 중입니다...")
    target_model = "models/gemini-1.5-pro" # 1.5 Pro 기본값 상향
    
    try:
        models_url = f"https://generativelanguage.googleapis.com/v1beta/models?key={GENAI_API_KEY}"
        model_res = requests.get(models_url).json()
        
        available_models = [m['name'] for m in model_res.get('models', []) if 'generateContent' in m.get('supportedGenerationMethods', [])]
        if "models/gemini-1.5-pro" in available_models or "models/gemini-1.5-pro-latest" in available_models:
            target_model = "models/gemini-1.5-pro"
        elif any('flash' in m for m in available_models):
            target_model = [m for m in available_models if 'flash' in m][0]
        else:
            target_model = "models/gemini-1.0-pro"
            
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
        p1.add_run(f"환율: {usd_rate} | 금리: {cbr_rate}\n")
        
        # AI 분석 추가
        doc.add_heading('AI 퀀트 애널리스트 Insight', level=1)
        for line in ai_insight.split('\n'):
            if line.strip().startswith('#'):
                doc.add_heading(line.replace('#', '').strip(), level=2)
            else:
                doc.add_paragraph(line.strip())
                
        # 중요 변수
        doc.add_heading('주요 변수 영향도 (XGBoost)', level=1)
        doc.add_paragraph(f"핵심 변수: {top_features_data}")
        
        doc.save(docx_path)
        print(f"✅ 워드 리포트 저장 완료: {docx_path}")
    except Exception as e:
        print(f"⚠️ 워드 리포트 생성 실패 (python-docx가 설치되어 있는지 확인하세요): {e}")

    print("📧 이메일을 발송합니다...")
    
    # 6. 이메일 템플릿 (초 단위까지 넣어 지메일 겹침 방지)
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    html_body = f"""
    <html>
    <body style="font-family: 'Malgun Gothic', sans-serif; color: #333; line-height: 1.6;">
        <h2 style="color: #004a99; border-bottom: 2px solid #004a99; padding-bottom: 5px;">📊 나이로비 부동산 시장 AI 브리핑</h2>
        <p><strong>발행일시:</strong> {current_time}</p>
        <div style="background: #f9f9f9; padding: 15px; border-radius: 5px; margin-bottom: 20px;">
            <p style="margin: 0;"><strong>데이터 기준 환율:</strong> {usd_rate} | <strong>금리:</strong> {cbr_rate}</p>
        </div>
        <h3 style="margin-bottom: 10px;">🧠 퀀트 애널리스트 Insight</h3>
        <div style="white-space: pre-wrap; background: #ffffff; border-left: 4px solid #004a99; padding: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">{ai_insight}</div>
        <h3 style="margin-top: 30px;">🔍 주요 변수 영향도 (XGBoost 모델, 하단 첨부파일 참조)</h3>
    </body>
    </html>
    """

    msg = MIMEMultipart()
    # 이메일 제목에도 초 단위를 추가하여 지메일이 점 3개(...)로 묶어버리는 현상 완벽 차단
    msg['Subject'] = f"[{datetime.now().strftime('%m/%d %H:%M:%S')}] 나이로비 부동산 퀀트 분석 리포트"
    msg['From'] = SENDER_EMAIL
    
    # 수신자에 donghyun1.kwon@lge.com 추가
    recipients = [RECEIVER_EMAIL, "donghyun1.kwon@lge.com"]
    msg['To'] = ", ".join([r for r in recipients if r])
    
    msg.attach(MIMEText(html_body, 'html'))

    # 7. 이미지 첨부
    img_path = os.path.join(base_path, "feature_importance.png")
    if os.path.exists(img_path):
        with open(img_path, 'rb') as f:
            msg.attach(MIMEImage(f.read(), name="feature_importance.png"))
            
    # [추가] Word 파일 첨부
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
        pass  # 클라우드(GitHub Actions) 환경에서는 키보드 입력이 없으므로 무시

if __name__ == "__main__":
    generate_and_send_report()