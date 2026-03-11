import streamlit as st
import fitz
from google import genai
import io
import re
import os
import time
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

# [설정 - API]
MY_API_KEY = st.secrets["GEMINI_API_KEY"].strip() 
client = genai.Client(api_key=MY_API_KEY)

# 고정 템플릿 파일명 설정
DEFAULT_WORD_TEMPLATE = "default_vf_template.docx"

# 1. 폰트 설정 (한글/영문 글꼴 지정)
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

# 3. 레이아웃 밀림 및 여백 보존형 삽입 함수 (여백 실종 방지 핵심)
def add_styled_content_at(target_p, text):
    lines = [line.strip() for line in str(text).split('\n') if line.strip()]
    if not lines: return target_p

    # 템플릿의 기존 여백 및 정렬 정보 상속
    orig_align = target_p.alignment
    orig_left = target_p.paragraph_format.left_indent
    orig_right = target_p.paragraph_format.right_indent

    current_p = target_p
    for i, line in enumerate(lines):
        if i == 0:
            current_p.text = "" 
            p_to_style = current_p
        else:
            new_p_xml = OxmlElement('w:p')
            current_p._p.addnext(new_p_xml)
            p_to_style = Paragraph(new_p_xml, current_p._parent)
            # 여백 상속
            p_to_style.alignment = orig_align
            p_to_style.paragraph_format.left_indent = orig_left
            p_to_style.paragraph_format.right_indent = orig_right
            current_p = p_to_style
            
        if line.startswith('## '):
            run = p_to_style.add_run(line.replace('## ', ''))
            set_font(run, "KoPub돋움체_Pro Medium", 12, bold=True)
        else:
            p_to_style.paragraph_format.line_spacing = 1.6
            p_to_style.paragraph_format.space_after = Pt(10)
            parts = re.split(r'(\*\*.*?\*\*)', line)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    run = p_to_style.add_run(part.replace('**', ''))
                    set_font(run, "KoPub돋움체_Pro Medium", 11, bold=True)
                else:
                    run = p_to_style.add_run(part)
                    set_font(run, "KoPub돋움체_Pro Light", 11)
    return current_p

# 4. 템플릿 치환 함수 (표 내부 [[tech_title]] 필터링 포함)
def replace_placeholder(doc, placeholder, content, is_inline=False, font_name=None, font_size=None, is_bold=False):
    def process_p(p):
        if placeholder in p.text:
            if is_inline:
                p.text = p.text.replace(placeholder, str(content))
                for run in p.runs: set_font(run, font_name, font_size, bold=is_bold)
            else:
                add_styled_content_at(p, content)
            return True
        return False

    for p in doc.paragraphs:
        if process_p(p): return True
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if process_p(p): return True
    return False

def extract_tag(text, tag_name):
    pattern = f"<{tag_name}>(.*?)(?:</{tag_name}>|$)"
    match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    return match.group(1).strip() if match else ""

# 5. 실행 함수 (분석 항목 100% 유지)
def run_virtual_firm(spec_file, doc_template, target_corp, ir_data, business_status):
    if not spec_file: return
    
    st.subheader(f"🏢 {target_corp if target_corp else 'Virtual Firm'} 심층 보고서")
    
    with st.spinner("🚀 3C, SWOT, 가치평가를 포함한 심층 분석 중..."):
        try:
            tech_text = extract_text_from_file(spec_file)
            ir_text = extract_text_from_file(ir_data) if ir_data else ""
            context = f"기업: {target_corp}\nIR: {ir_text}\n상태: {business_status}"
            
            # [프롬프트 - 모든 분석 항목 강제 지정]
            prompt = f"""당신은 가치평가사입니다. 아래 내용을 <tech_title>, <section_1>~<section_4> 태그로 작성하세요.
            - <section_3>: 사업화 로드맵, 예상매출액, 수익접근법 기술가치평가(산출 수식 상세 기술) 포함.
            - <section_4>: 3C 분석, SWOT 분석, Lean Canvas, 최종 제안(기술이전 vs 창업) 포함.
            데이터: {context}\n명세서: {tech_text}"""

            response = client.models.generate_content(model="gemini-2.0-flash", contents=prompt)
            raw_response = response.text.strip()

            ai_data = {t: extract_tag(raw_response, t) for t in ["tech_title", "section_1", "section_2", "section_3", "section_4"]}
            
            doc = Document(doc_template if doc_template else DEFAULT_WORD_TEMPLATE)

            # 제목 치환 (표 내부 [[tech_title]], 18pt Bold)
            replace_placeholder(doc, "[[tech_title]]", ai_data['tech_title'], is_inline=True, 
                                font_name="KoPub돋움체_Pro Bold", font_size=18, is_bold=True)
            
            # 섹션 치환 (여백 보존)
            for i in range(1, 5):
                replace_placeholder(doc, "{{" + f"section_{i}" + "}}", ai_data[f'section_{i}'])

            doc_io = io.BytesIO()
            doc.save(doc_io)
            st.success("✅ 모든 분석이 포함된 보고서가 완료되었습니다!")
            st.download_button(label="📥 보고서 다운로드", data=doc_io.getvalue(), 
                               file_name="Virtual_Firm_Master_Report.docx", 
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"오류: {str(e)}"
