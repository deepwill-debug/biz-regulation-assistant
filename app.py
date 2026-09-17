import os
import sys
import io
import requests
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.prompts import ChatPromptTemplate

# Windows 콘솔 UTF-8 출력 보장
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

# 국가법령 API 클라이언트 임포트
from law_api_client import LawApiClient


# 0. Gemini response content 텍스트 안전 추출 헬퍼 함수
def extract_text(content):
    if isinstance(content, str):
        return content
    elif isinstance(content, list):
        texts = []
        for item in content:
            if isinstance(item, str):
                texts.append(item)
            elif isinstance(item, dict) and "text" in item:
                texts.append(item["text"])
            else:
                texts.append(str(item))
        return "\n".join(texts)
    return str(content)


# 1. 페이지 설정
st.set_page_config(
    page_title="인천상공회의소 기업애로 진단 및 건의서 자동생성 지원시스템",
    page_icon="🏛️",
    layout="wide"
)

# 2. 사이드바 API 설정 (secrets.toml -> 환경변수 -> 직접 입력)
st.sidebar.title("⚙️ 시스템 설정")

gemini_api_key = ""
try:
    gemini_api_key = st.secrets.get("GEMINI_API_KEY", "")
except Exception:
    pass

if not gemini_api_key:
    gemini_api_key = os.environ.get("GEMINI_API_KEY", "")

if gemini_api_key:
    st.sidebar.success("🔑 Gemini API Key 자동 로드 완료!")
else:
    st.sidebar.caption("Google AI Studio(aistudio.google.com)에서 무료 발급")
    gemini_api_key = st.sidebar.text_input("Gemini API Key 입력", type="password", help="AIzaSy... 로 시작하는 키 입력")

# 국가법령 API OC ID (secrets.toml -> 환경변수 -> incham)
law_oc_id = ""
try:
    law_oc_id = st.secrets.get("LAW_OC_ID", "")
except Exception:
    pass

if not law_oc_id:
    law_oc_id = os.environ.get("LAW_OC_ID", "incham")

law_oc_id = st.sidebar.text_input("국가법령 API OC ID", value=law_oc_id, help="open.law.go.kr 발급 ID")


# 3. Google Gemini API 키 기준 지원 모델 동적 조회 (list_models)
@st.cache_data(ttl=1800)
def fetch_available_gemini_models(api_key: str):
    """API 키에서 generateContent를 지원하는 유효 모델 목록 조회"""
    default_models = ["gemini-3.6-flash", "gemini-2.5-flash", "gemini-1.5-flash", "gemini-2.0-flash"]
    if not api_key.strip():
        return default_models

    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key.strip()}"
        res = requests.get(url, timeout=5)
        if res.status_code == 200:
            data = res.json()
            valid_models = []
            for m in data.get("models", []):
                methods = m.get("supportedGenerationMethods", [])
                if "generateContent" in methods:
                    name = m.get("name", "").replace("models/", "")
                    if any(prefix in name for prefix in ["gemini-3.6", "gemini-2.5", "gemini-2.0", "gemini-1.5", "gemini-pro"]):
                        valid_models.append(name)
            
            if valid_models:
                preferred_order = [
                    "gemini-3.6-flash",
                    "gemini-2.5-flash",
                    "gemini-1.5-flash",
                    "gemini-2.0-flash",
                    "gemini-1.5-pro",
                    "gemini-2.0-flash-exp",
                    "gemini-1.5-flash-8b",
                ]
                sorted_models = [m for m in preferred_order if m in valid_models] + [
                    m for m in valid_models if m not in preferred_order
                ]
                return sorted_models
    except Exception:
        pass

    return default_models


available_models = fetch_available_gemini_models(gemini_api_key)
model_choice = st.sidebar.selectbox(
    "Gemini 모델 선택 (API 조회 결과)",
    available_models,
    index=0,
    help="선택된 API 키에서 generateContent를 지원하는 유효 모델 목록입니다."
)

st.sidebar.markdown("---")
st.sidebar.info(
    "💡 **Gemini + 국가법령 API 연동 중**\n\n"
    f"현재 적용된 OC ID: `{law_oc_id}`\n\n"
    f"선택된 모델: `{model_choice}`\n\n"
    "기업 애로 제출 시 `law.go.kr`에서 실시간 조문을 검색하여 정확한 법령 근거를 반영합니다."
)

# 4. 메인 타이틀
st.title("🏛️ 인천상공회의소 기업애로 진단 및 건의서 자동생성 지원시스템")
st.caption("Google Gemini 모델을 기반으로 기업 애로를 분석하여 수용 가능성을 진단하고, 실시간 국가법령 조문을 반영한 정책건의서 초안을 작성합니다.")

# 5. 입력 폼 구성
with st.container():
    col1, col2 = st.columns([2, 1])
    with col1:
        complaint_text = st.text_area(
            "기업 접수 애로사항 (상담 일지, 메모, 구두 접수 내용 등)",
            placeholder="예시: 지식산업센터에 입주한 제조업체인데, 공장 등록 기준 규제가 너무 까다로워서 증설을 못 하고 있습니다. 현행 면적 및 부대시설 기준 완화가 필요합니다.",
            height=180
        )
    with col2:
        category = st.selectbox(
            "소관 분야 (추정)",
            ["선택 안 함 (AI 자동분류)", "입지/공장등록/산업단지", "환경/유해화학물질/ESG", "노동/인력/비자", "안전/소방/건축", "조세/세제/금융", "기타 규제애로"]
        )
        search_keyword_input = st.text_input(
            "법령 검색 직접 키워드 (선택)",
            placeholder="예: 산업집적활성화, 공장등록",
            help="비워두면 AI가 내용에서 키워드를 자동 추출합니다."
        )
        submit_btn = st.button("🚀 애로 진단 및 건의서 생성", type="primary", use_container_width=True)


# 6. 법령 검색 헬퍼 함수 (LawApiClient RAG)
def get_relevant_laws(complaint: str, custom_kw: str, category_name: str, oc_id: str):
    client = LawApiClient(oc_id=oc_id)
    
    keywords = []
    if custom_kw.strip():
        keywords = [k.strip() for k in custom_kw.split(",") if k.strip()]
    else:
        cat_map = {
            "입지/공장등록/산업단지": ["산업집적활성화", "국토계획법"],
            "환경/유해화학물질/ESG": ["화학물질관리법", "환경영향평가법"],
            "노동/인력/비자": ["근로기준법", "파견근로자"],
            "안전/소방/건축": ["소방시설법", "산업안전보건법", "건축법"],
            "조세/세제/금융": ["조세특례제한법", "지방세특례제한법"],
        }
        if category_name in cat_map:
            keywords.extend(cat_map[category_name])
            
        if "ESS" in complaint or "배터리" in complaint or "전기" in complaint:
            keywords.append("전기사업법")
        if "공장" in complaint or "증설" in complaint or "면적" in complaint:
            keywords.append("산업집적활성화")
        if "화학" in complaint or "유해" in complaint:
            keywords.append("화학물질관리법")

    if not keywords:
        keywords = ["산업집적활성화", "전기사업법"]

    keywords = list(dict.fromkeys(keywords))[:3]
    
    law_summary_blocks = []
    raw_articles_list = []

    for kw in keywords:
        laws = client.search_laws(kw, count=1)
        for law in laws:
            articles = client.get_law_articles(law.mst)
            if not articles:
                continue

            art_snippets = []
            for art in articles[:5]:
                art_snippets.append(f"  - [{art.article_no} {art.title}] {art.content[:200]}...")
                raw_articles_list.append({
                    "law_name": law.name,
                    "law_type": law.law_type,
                    "article_no": art.article_no,
                    "title": art.title,
                    "content": art.content
                })

            block = f"■ [실시간 검토 법령] {law.name} ({law.law_type}, MST: {law.mst})\n" + "\n".join(art_snippets)
            law_summary_blocks.append(block)

    context_str = "\n\n".join(law_summary_blocks) if law_summary_blocks else "관련 법령 실시간 검색 결과 없음"
    return context_str, raw_articles_list


# 7. 워드(.docx) 문서 생성 서식 헬퍼 함수
def create_docx(result_text, complaint: str = "", category: str = ""):
    text_str = extract_text(result_text)
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    title_p = doc.add_paragraph()
    title_run = title_p.add_run("기업 규제애로 진단 보고서 및 정책 건의서")
    title_run.font.size = Pt(20)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(16, 44, 87)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub_p = doc.add_paragraph()
    sub_run = sub_p.add_run(f"분야: {category} | 인천상공회의소 규제개혁 지원시스템 (Gemini)")
    sub_run.font.size = Pt(10)
    sub_run.font.italic = True
    sub_run.font.color.rgb = RGBColor(100, 100, 100)
    sub_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_paragraph("=" * 60)

    if complaint:
        p_orig = doc.add_paragraph()
        r_orig_title = p_orig.add_run("📌 접수 애로사항:\n")
        r_orig_title.font.bold = True
        r_orig_body = p_orig.add_run(f'"{complaint}"')
        r_orig_body.font.italic = True
        doc.add_paragraph("-" * 60)

    for line in text_str.splitlines():
        line_str = line.strip()
        if line_str.startswith("###") or line_str.startswith("PART"):
            p = doc.add_paragraph()
            r = p.add_run(line_str.replace("#", "").strip())
            r.font.size = Pt(14)
            r.font.bold = True
            r.font.color.rgb = RGBColor(27, 89, 150)
        elif line_str.startswith("- **") or line_str.startswith("1.") or line_str.startswith("2.") or line_str.startswith("3."):
            p = doc.add_paragraph()
            r = p.add_run(line_str)
            r.font.size = Pt(11)
            r.font.bold = True
        else:
            p = doc.add_paragraph(line_str)
            p.style.font.size = Pt(10.5)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# 8. 분석 및 생성 로직
if submit_btn:
    if not gemini_api_key.strip():
        st.error("🔑 Google Gemini API Key를 입력하거나 secrets.toml / 환경변수(GEMINI_API_KEY)로 등록해 주세요.")
    elif not complaint_text.strip():
        st.warning("⚠️ 분석할 기업애로 내용을 입력해 주세요.")
    else:
        status_box = st.status("🔍 기업애로 분석 및 법령 연동 진행 중...", expanded=True)
        
        try:
            # 1단계: 실시간 법령 조회
            status_box.write("📜 [1/3] 국가법령정보센터(law.go.kr) 실시간 조문 검색 중...")
            law_context, raw_articles = get_relevant_laws(
                complaint_text, search_keyword_input, category, law_oc_id
            )

            # 2단계: Gemini 모델 호출
            status_box.write(f"🤖 [2/3] Google Gemini ({model_choice}) 기반 수용성 진단 및 건의서 작성 중...")
            llm = ChatGoogleGenerativeAI(
                model=model_choice,
                temperature=0.2,
                google_api_key=gemini_api_key.strip()
            )
            
            system_prompt = """당신은 대한민국의 규제개혁 및 기업애로 정책건의 전문 행정위원입니다.
기업의 비구조화된 단순 애로사항과 [국가법령정보센터에서 실시간 조회된 법령 조문 컨텍스트]를 검토하여
다음 두 가지 파트로 전문적인 결과물을 작성하세요.

---
### PART 1. 정부/지자체 수용 가능성 사전 진단
1. **수용 가능성 지수**: [높음(75% 이상) / 보통(40~74%) / 낮음(40% 미만)] 및 예상 확률(%)
2. **주관 부처 및 대상 규제 법령**: 관련 법률, 시행령, 시행규칙 또는 지자체 조례 등 조항 명시 (실시간 조문 컨텍스트 적극 인용)
3. **부처 거절/불수용 예상 사유 (3대 리스크 점검)**:
   - 안전·환경·국민 건강 등 공익 침해 우려
   - 타 기업·업종 간 형평성 및 특혜 시비
   - 상위법 저촉 및 세수 감소/재정 부담 여부
4. **수용률 제고를 위한 전략적 조언**: 법률 전면 개정 대신 시행령/규칙 개정, 규제샌드박스, 부관(조건부 완화) 등 현실적 대안 제시

---
### PART 2. 부처 제출용 공식 정책건의서 초안
- **[건의 제목]**: (단순 완화 요구가 아닌, 규제 합리화를 통한 경제 활성화·공익 기여형 제목)
- **1. 현황 및 문제점**:
  - 관련 법령 조항 및 현행 기준 명시
  - 현행 규제로 인한 기업 현장의 경영 병목 및 애로 구체화
- **2. 개선 방안**:
  - 법률 개정(국회) 지양, 시행령·시행규칙 또는 고시 개정안 중심 제시
  - 안전장치(부관) 마련을 통한 공무원 면책 논거 포함
- **3. 기대 효과**:
  - 개별 기업 혜택 배제, 지역 산업 경쟁력 강화, 투자 유발, 고용 창출 등 공익적 파급효과
"""

            user_prompt = f"""[분야] {category}

[기업 접수 애로내용]
{complaint_text}

[국가법령정보센터 API 실시간 조문 컨텍스트]
{law_context}
"""
            
            prompt = ChatPromptTemplate.from_messages([
                ("system", system_prompt),
                ("user", user_prompt)
            ])
            
            response = (prompt | llm).invoke({})
            result_content = extract_text(response.content)
            
            status_box.update(label="✅ 분석 및 정책건의서 생성이 완료되었습니다!", state="complete", expanded=False)
            
            # 3단계: 결과 화면 출력 (탭 구성)
            tab1, tab2, tab3 = st.tabs(["📊 진단 보고서 & 건의서", "📜 참조 실시간 법령 조문", "📥 Word 문서 다운로드"])
            
            with tab1:
                st.markdown(result_content)
                
            with tab2:
                st.subheader("📜 실시간 검색된 관련 법령 조문 목록")
                if raw_articles:
                    for idx, item in enumerate(raw_articles, 1):
                        with st.expander(f"{idx}. {item['law_name']} ({item['law_type']}) - {item['article_no']} {item['title']}"):
                            st.write(item["content"])
                else:
                    st.info("검색된 법령 조문이 없습니다.")

            with tab3:
                st.subheader("📄 제출용 Word (.docx) 문서 다운로드")
                st.write("작성된 진단 보고서와 정책건의서를 서식이 적용된 MS Word 문서 파일로 다운로드합니다.")
                docx_file = create_docx(result_content, complaint_text, category)
                st.download_button(
                    label="📄 Word(.docx) 파일 다운로드",
                    data=docx_file,
                    file_name="기업규제애로_정책건의서.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )
        except Exception as e:
            status_box.update(label="❌ 처리 중 오류 발생", state="error", expanded=True)
            st.error(f"오류가 발생했습니다: {e}")
