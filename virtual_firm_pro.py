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
    # 빈 줄 제거 및 텍스트 정제
    lines = [line.strip() for line in str(text).split('\n') if line.strip()]
    if not lines: return target_p

    current_p = target_p
    for i, line in enumerate(lines):
        if i == 0:
            current_p.text = "" # 템플릿의 {{tag}} 줄을 재활용하여 밀림 방지
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

# 4. 템플릿 치환 함수 (표 내부 [[태그]] 및 일반 {{태그}} 대응)
def replace_placeholder(doc, placeholder, content, is_inline=False, font_name=None, font_size=None, is_bold=False):
    def process_paragraph(p):
        if placeholder in p.text:
            if is_inline:
                p.text = p.text.replace(placeholder, content)
                if font_name:
                    for run in p.runs:
                        set_font(run, font_name, font_size, bold=is_bold)
            else:
                add_styled_content_at(p, content)
            return True
        return False

    # 본문 검색
    for p in doc.paragraphs:
        if process_paragraph(p): return True
            
    # 표 내부 검색 (템플릿의 tech_title이 표 안에 있음)
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

# 5. 스마트 AI 생성기
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
        st.error("분석을 위한 명세서 파일이 필요합니다.")
        return

    st.subheader(f"🏢 {target_corp if target_corp else 'Virtual Firm'} 보고서 생성 중...")
    
    with st.spinner("🚀 기술가치평가 및 비즈니스 모델링을 수행 중입니다..."):
        tech_text = extract_text_from_file(spec_file)
        ir_text = extract_text_from_file(ir_data) if ir_data else ""
        context = f"대상기업: {target_corp}\n기업IR: {ir_text}\n사업현황: {business_status}"
        
        prompt = f"""당신은 부산대학교 기술지주회사의 가치평가 전문가입니다. 
        <tech_title>, <section_1>, <section_2>, <section_3>, <section_4> 태그를 사용하여 
        기술 개요, 시장 분석, 기술가치평가 산출 과정, 최종 창업 전략을 상세히 작성하세요. 
        마크다운 표는 절대 금지하며 소제목(##)과 개조식을 사용하세요.
        [대상 정보] {context}\n[명세서] {tech_text}"""

        raw_response, used_model = generate_one_shot(prompt)
        if not raw_response: return

        ai_data = {t: extract_tag(raw_response, t) for t in ["tech_title", "section_1", "section_2", "section_3", "section_4"]}
        
        # 문서 로드
        if doc_template:
            doc = Document(doc_template)
        elif os.path.exists(DEFAULT_WORD_TEMPLATE):
            doc = Document(DEFAULT_WORD_TEMPLATE)
        else:
            doc = Document()

        # 템플릿 치환 (태그 형태에 맞춰 수정)
        # 1. 제목: [[tech_title]] 대응 (표 내부 포함)
        replace_placeholder(doc, "[[tech_title]]", ai_data['tech_title'], is_inline=True, 
                            font_name="KoPub돋움체_Pro Bold", font_size=18, is_bold=True)
        
        # 2. 각 섹션: {{section_x}} 대응
        for i in range(1, 5):
            replace_placeholder(doc, f"{{{{section_{i}}}}}}", ai_data[f'section_{i}'])

        # 결과 다운로드
        doc_io = io.BytesIO()
        doc.save(doc_io)
        st.download_button(label="📥 가상기업 보고서 다운로드", data=doc_io.getvalue(), 
                           file_name="Virtual_Firm_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
