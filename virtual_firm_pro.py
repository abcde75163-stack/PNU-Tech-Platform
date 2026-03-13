import streamlit as st
import fitz
from google import genai
import io
import re
import time
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# [설정 - API]
try:
    MY_API_KEY = st.secrets["GEMINI_API_KEY"].strip()
    client = genai.Client(api_key=MY_API_KEY)
except Exception as e:
    st.error("API 키 로드 실패. Secrets 설정을 확인하세요.")

# 1. 폰트 설정
def set_font(run, font_name, size, bold=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold

# 2. 파일 텍스트 추출 (PDF/Docx)
def extract_text_from_file(uploaded_file):
    if uploaded_file is None: return ""
    file_name = uploaded_file.name.lower()
    uploaded_file.seek(0)
    try:
        if file_name.endswith('.pdf'):
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            return "".join([page.get_text() for page in doc])[:15000]
        elif file_name.endswith('.docx'):
            doc = Document(uploaded_file)
            return "\n".join([p.text for p in doc.paragraphs])[:15000]
    except Exception: return ""
    return ""

# 3. 스타일 콘텐츠 삽입 (표 없이 개조식 텍스트 최적화)
def add_styled_content(doc, text):
    lines = str(text).split('\n')
    for line in lines:
        line_stripped = line.strip()
        if not line_stripped: continue
        
        if line_stripped.startswith('## '): # 소제목
            p = doc.add_paragraph()
            run = p.add_run(line_stripped.replace('## ', ''))
            set_font(run, "KoPub돋움체_Pro Medium", 13, bold=True)
            p.paragraph_format.space_before = Pt(12)
            continue
            
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.6
        p.paragraph_format.space_after = Pt(10)
        
        # 볼드체(**) 처리
        parts = re.split(r'(\*\*.*?\*\*)', line_stripped)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                run = p.add_run(part.replace('**', ''))
                set_font(run, "KoPub돋움체_Pro Medium", 11, bold=True)
            else:
                run = p.add_run(part)
                set_font(run, "KoPub돋움체_Pro Light", 11)

# 💡 태그 추출 방어 코드
def extract_tag(text, tag_name):
    pattern = f"<{tag_name}>(.*?)(?:</{tag_name}>|$)"
    match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    return match.group(1).strip() if match else ""

# 5. 메인 실행 함수
def run_virtual_firm(spec_file, doc_template, target_corp, ir_data, business_status):
    if not spec_file:
        st.error("특허 명세서 파일이 필요합니다.")
        return

    st.subheader(f"🏢 {target_corp if target_corp else 'Virtual Firm'} 심층 보고서 생성")
    tech_text = extract_text_from_file(spec_file)
    
    if len(tech_text.strip()) < 50:
        st.error("❌ 파일에서 텍스트를 읽을 수 없습니다. (이미지 스캔본 여부 확인)")
        return

    with st.spinner("🚀 환각 방지 및 심층 텍스트 분석 중... (약 20~30초 소요)"):
        try:
            ir_text = extract_text_from_file(ir_data) if ir_data else ""
            context = f"타겟 기업: {target_corp}\nIR 데이터: {ir_text}\n사업 현황: {business_status}"
            
            # [프롬프트] 요약 금지 및 심층 서술 명령 극대화
            prompt = f"""당신은 국내 최고의 기술사업화 전문가이자 전략 컨설턴트입니다.
            제공된 [특허 명세서]의 실제 기술 내용을 기반으로, 절대로 요약하지 말고 아주 상세하고 전문적인 '심층 분석 보고서'를 작성하세요.

            [특허 명세서 원본 데이터]
            {tech_text}

            [작성 지침 - 절대 엄수]
            1. 모든 내용은 소제목(##)과 개조식(-, *) 텍스트로만 '글로 길게' 작성하세요. (표 사용 절대 금지)
            2. 응답은 반드시 <tech_title>, <section_1>, <section_2>, <section_3>, <section_4> 태그 형식을 유지하세요.
            3. 각 섹션은 반드시 A4 용지 반 페이지 이상이 채워질 만큼 **문단별로 깊이 있게(5~10문단 이상)** 서술하세요.

            [섹션별 필수 포함 내용]
            - <tech_title>: 기술의 핵심을 관통하는 전문적인 기술 비즈니스 명칭
            - <section_1>: 기술의 메커니즘, 작동 원리, 기존 기술 대비 차별성을 명세서 기반으로 아주 상세히 분석
            - <section_2>: 전방 산업 트렌드 및 해결하려는 페인 포인트(Pain point) 심층 분석
            - <section_3>: 단계별 스케일업 전략. '수익접근법' 기술가치평가 수식을 제시하고, 예상 매출 도출의 근거(가정)를 상세히 글로 나열
            - <section_4>: 3C/SWOT 정밀 분석. 최종적으로 '기술이전' vs '직접 창업' 중 하나를 선택하고 그 근거를 매우 논리적으로 서술

            [기타 정보]
            {context}"""

            response = client.models.generate_content(
                model="models/gemini-2.5-flash-lite", # 가장 안정적인 모델로 변경
                contents=prompt
            )
            raw_response = response.text.strip()

            ai_data = {t: extract_tag(raw_response, t) for t in ["tech_title", "section_1", "section_2", "section_3", "section_4"]}
            
            # 💡 [핵심] 템플릿 방식 탈피 - 새 문서로 조립
            doc = Document()
            
            # 제목 섹션
            title_p = doc.add_paragraph("\n\n")
            title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            clean_title = ai_data.get("tech_title", "심층 사업화 보고서").replace('#', '').replace('*', '').strip()
            title_run = title_p.add_run(f"Virtual Firm 심층 사업화 전략 보고서\n\n[{clean_title}]")
            set_font(title_run, "KoPub돋움체_Pro Bold", 18, bold=True)
            
            doc.add_page_break()

            # 본문 섹션들
            sections = [
                ("Ⅰ. 기술 개요 및 메커니즘 분석", "section_1"),
                ("Ⅱ. 시장 트렌드 및 과제 해결 분석", "section_2"),
                ("Ⅲ. Scale-up 및 심층 재무 로드맵", "section_3"),
                ("Ⅳ. 최종 사업화 제안 (이전 vs 창업)", "section_4")
            ]

            for title_text, key in sections:
                h_p = doc.add_paragraph()
                set_font(h_p.add_run(f"\n{title_text}"), "KoPub돋움체_Pro Bold", 15, bold=True)
                h_p.paragraph_format.space_before = Pt(15)
                
                content = ai_data.get(key, "")
                if content:
                    add_styled_content(doc, content)
                else:
                    doc.add_paragraph("생성된 내용이 없습니다.")

            doc_io = io.BytesIO()
            doc.save(doc_io)
            
            st.success("✅ 고품질 심층 보고서 작성이 완료되었습니다!")
            st.download_button(label="📥 최종 심층 보고서 다운로드", data=doc_io.getvalue(), 
                               file_name=f"VF_Full_Report_{target_corp}.docx")
        except Exception as e:
            st.error(f"오류 발생: {str(e)}")



