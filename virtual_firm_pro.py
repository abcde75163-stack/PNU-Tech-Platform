import streamlit as st
import fitz
from google import genai
import io
import re
import os
import time
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

# [설정 - API]
MY_API_KEY = st.secrets["GEMINI_API_KEY"].strip() 
client = genai.Client(api_key=MY_API_KEY)

# 고정 템플릿 파일명 설정
DEFAULT_WORD_TEMPLATE = "default_vf_template.docx"

# 1. 폰트 설정 함수
def set_font(run, font_name, size, bold=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold

# 2. 파일 텍스트 추출 (PDF 및 Word 지원) 
def extract_text_from_file(uploaded_file):
    if uploaded_file is None: return ""
    file_name = uploaded_file.name.lower()
    uploaded_file.seek(0)
    
    if file_name.endswith('.pdf'):
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        return "".join([page.get_text() for page in doc])[:15000]
    elif file_name.endswith('.docx'):
        doc = Document(uploaded_file)
        return "\n".join([p.text for p in doc.paragraphs])[:15000]
    return ""

# 3. 레이아웃 밀림 방지형 콘텐츠 삽입 함수 
def add_styled_content_at(target_p, text):
    lines = [line.strip() for line in str(text).split('\n') if line.strip()]
    if not lines: return target_p

    current_p = target_p
    for i, line in enumerate(lines):
        if i == 0:
            # 템플릿의 {{tag}} 줄을 삭제하고 그 자리에 첫 줄 삽입 (밀림 방지 핵심) [cite: 378]
            current_p.text = "" 
            p_to_style = current_p
        else:
            new_p_xml = OxmlElement('w:p')
            current_p._p.addnext(new_p_xml)
            p_to_style = Paragraph(new_p_xml, current_p._parent)
            current_p = p_to_style
            
        if line.startswith('## '):
            run = p_to_style.add_run(line.replace('## ', ''))
            set_font(run, "KoPub돋움체_Pro Medium", 12, bold=True)
        else:
            p_to_style.paragraph_format.line_spacing = 1.6
            p_to_style.paragraph_format.space_after = Pt(12)
            parts = re.split(r'(\*\*.*?\*\*)', line)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    run = p_to_style.add_run(part.replace('**', ''))
                    set_font(run, "KoPub돋움체_Pro Medium", 11, bold=True)
                else:
                    run = p_to_style.add_run(part)
                    set_font(run, "KoPub돋움체_Pro Light", 11)
    return current_p

# 4. 템플릿 치환 함수 (표 내부 [[태그]] 및 일반 {{태그}} 대응) [cite: 363, 372]
def replace_placeholder(doc, placeholder, content, is_inline=False, font_name=None, font_size=None, is_bold=False):
    def process_paragraph(p):
        if placeholder in p.text:
            if is_inline:
                # 제목 치환: [[tech_title]] 대응 및 폰트 고정 
                p.text = p.text.replace(placeholder, content)
                if font_name:
                    for run in p.runs:
                        set_font(run, font_name, font_size, bold=is_bold)
            else:
                # 본문 삽입: {{section_x}} 대응 [cite: 378]
                add_styled_content_at(p, content)
            return True
        return False

    # 1) 일반 단락에서 찾기
    for p in doc.paragraphs:
        if process_paragraph(p): return True
            
    # 2) 모든 표(Table) 내부에서 찾기 (템플릿 제목 대응) 
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if process_paragraph(p): return True
    return False

def extract_tag(text, tag_name):
    pattern = f"<{tag_name}>(.*?)(?:</{tag_name}>|$)"
    match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    return match.group(1).strip() if match else ""

# 5. 스마트 AI 생성기 (가치평가 및 전략 수립) [cite: 72, 133, 322]
def generate_one_shot(prompt):
    models = ["gemini-2.0-flash", "gemini-2.0-flash-lite"]
    for model_name in models:
        try:
            response = client.models.generate_content(
                model=model_name, contents=prompt, config={"temperature": 0.2}
            )
            if response.text: return response.text.strip(), model_name
        except: continue
    return None, "생성 실패"

# 6. 메인 실행 함수
def run_virtual_firm(spec_file, doc_template, target_corp, ir_data, business_status):
    if not spec_file:
        st.error("분석을 위한 기술 명세서 파일이 필요합니다. [cite: 17]")
        return

    st.subheader(f"🏢 {target_corp if target_corp else 'Virtual Firm'} 보고서 생성 중... [cite: 361]")
    
    with st.spinner("🚀 기술가치평가 및 비즈니스 모델링을 수행 중입니다... [cite: 133]"):
        try:
            tech_text = extract_text_from_file(spec_file) # PDF/DOCX 자동 감지 [cite: 19]
            ir_text = extract_text_from_file(ir_data) if ir_data else ""
            context = f"대상기업: {target_corp}\n기업IR: {ir_text}\n사업현황: {business_status}"
            
            prompt = f"""당신은 부산대학교 기술지주회사의 최고급 비즈니스 아키텍트입니다. [cite: 362]
            아래 정보를 바탕으로 <tech_title>, <section_1>, <section_2>, <section_3>, <section_4> 태그 내에 보고서 내용을 작성하세요.
            반드시 Ⅲ장에서는 '수익접근법'을 통한 기술가치평가 산출 과정을 포함해야 합니다. [cite: 133, 147]
            
            [분석 데이터]
            {context}
            명세서: {tech_text}
            """

            raw_response, used_model = generate_one_shot(prompt)
            if not raw_response:
                st.error("AI 생성 중 오류가 발생했습니다.")
                return

            # 태그별 데이터 추출
            ai_data = {t: extract_tag(raw_response, t) for t in ["tech_title", "section_1", "section_2", "section_3", "section_4"]}
            
            # 문서 로드 (사용자 템플릿 -> 기본 템플릿 -> 백지 순) [cite: 360]
            if doc_template:
                doc = Document(doc_template)
            elif os.path.exists(DEFAULT_WORD_TEMPLATE):
                doc = Document(DEFAULT_WORD_TEMPLATE)
            else:
                doc = Document()

            # 템플릿 치환 로직 [cite: 363, 378, 382, 385, 388]
            # 1. 제목 치환 ([[tech_title]] 대응, 18pt Bold) 
            replace_placeholder(doc, "[[tech_title]]", ai_data['tech_title'], is_inline=True, 
                                font_name="KoPub돋움체_Pro Bold", font_size=18, is_bold=True)
            
            # 2. 본문 섹션 치환 ({{section_1}} ~ {{section_4}} 대응) [cite: 378, 381, 384, 387]
            for i in range(1, 5):
                replace_placeholder(doc, f"{{{{section_{i}}}}}}", ai_data[f'section_{i}'])

            # 결과물 저장 및 다운로드
            doc_io = io.BytesIO()
            doc.save(doc_io)
            st.success("✅ 보고서 생성이 완료되었습니다! [cite: 354]")
            st.download_button(
                label="📥 심층 Virtual Firm 보고서 다운로드", 
                data=doc_io.getvalue(), 
                file_name=f"Virtual_Firm_Report_{target_corp}.docx", 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"실행 중 오류 발생: {str(e)}")
