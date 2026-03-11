import streamlit as st
import fitz
from google import genai
import io
import re
import os
import time
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

# [설정 - API]
MY_API_KEY = st.secrets["GEMINI_API_KEY"].strip() 
client = genai.Client(api_key=MY_API_KEY)

# 고정 템플릿 파일명 설정
DEFAULT_WORD_TEMPLATE = "default_vf_template.docx"

# 1. 폰트 및 스타일 설정
def set_font(run, font_name, size, bold=False, color=None):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)

# 2. PDF 텍스트 추출
def extract_text_from_pdf(uploaded_file):
    if uploaded_file is None: return ""
    uploaded_file.seek(0)
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    return "".join([page.get_text() for page in doc])[:15000]

# 3. [C안 구현] 핵심 지표 요약 박스 생성 함수
def add_summary_box(doc, summary_text):
    # 단일 셀 표를 만들어 박스 효과 구현
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Light List Accent 1' # 강조 스타일 적용
    cell = table.cell(0, 0)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 박스 내부 텍스트 스타일링
    lines = summary_text.split('\n')
    for line in lines:
        if not line.strip(): continue
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(line.strip())
        set_font(run, "KoPub돋움체_Pro Medium", 11, bold=True)
    
    doc.add_paragraph() # 박스 뒤 간격 추가

# 4. [B안 구현] 텍스트를 표로 변환하는 함수 (SWOT, 3C용)
def add_styled_table_from_text(doc, section_title, items_text):
    doc.add_paragraph().add_run(f"■ {section_title} 상세 분석").bold = True
    
    # 데이터를 리스트화
    items = [item.strip() for item in items_text.split('\n') if item.strip()]
    
    table = doc.add_table(rows=len(items), cols=1)
    table.style = 'Table Grid'
    
    for i, item in enumerate(items):
        cell = table.cell(i, 0)
        p = cell.paragraphs[0]
        run = p.add_run(item)
        set_font(run, "KoPub돋움체_Pro Light", 10)
    
    doc.add_paragraph()

# 5. 콘텐츠 삽입 및 스타일링
def add_styled_content_at(target_p, text):
    lines = str(text).split('\n')
    current_p = target_p
    for line in lines:
        line_stripped = line.strip()
        if not line_stripped: continue
        
        new_p_xml = OxmlElement('w:p')
        current_p._p.addnext(new_p_xml)
        new_p = Paragraph(new_p_xml, current_p._parent)
        
        if line_stripped.startswith('## '):
            run = new_p.add_run(line_stripped.replace('## ', ''))
            set_font(run, "KoPub돋움체_Pro Medium", 12, bold=True, color=(0, 51, 153)) # 파란색 제목
        else:
            new_p.paragraph_format.line_spacing = 1.6
            parts = re.split(r'(\*\*.*?\*\*)', line_stripped)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    run = new_p.add_run(part.replace('**', ''))
                    set_font(run, "KoPub돋움체_Pro Medium", 11, bold=True)
                else:
                    run = new_p.add_run(part)
                    set_font(run, "KoPub돋움체_Pro Light", 11)
        current_p = new_p
    return current_p

def replace_placeholder(doc, placeholder, content, is_inline=False):
    for p in doc.paragraphs:
        if placeholder in p.text:
            if is_inline:
                p.text = p.text.replace(placeholder, content)
            else:
                p.text = p.text.replace(placeholder, "") 
                if content: add_styled_content_at(p, content)    
            return True
    return False

def extract_tag(text, tag_name):
    pattern = f"<{tag_name}>(.*?)(?:</{tag_name}>|$)"
    match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    return match.group(1).strip() if match else ""

def generate_one_shot(prompt):
    try:
        response = client.models.generate_content(
            model="gemini-2.0-flash", # 고성능 모델 지정
            contents=prompt,
            config={"temperature": 0.2} 
        )
        return response.text.strip(), "gemini-2.0-flash"
    except Exception as e:
        return None, str(e)

# ★ 메인 실행 함수
def run_virtual_firm(spec_file, doc_template, target_corp, ir_data, business_status):
    if not spec_file:
        st.error("분석을 위한 기술 명세서(PDF)가 필요합니다.")
        return

    st.subheader(f"🏢 {target_corp if target_corp else 'Virtual Firm'} 프리미엄 전략 보고서 생성 중...")
    
    with st.spinner("📊 핵심 지표 요약 박스 및 비즈니스 표를 구성하고 있습니다..."):
        try:
            tech_text = extract_text_from_pdf(spec_file)
            ir_text = extract_text_from_pdf(ir_data) if ir_data else ""
            context = f"대상기업: {target_corp}\n기업IR자료: {ir_text}\n사업현황: {business_status}"
            
            prompt = f"""당신은 최고급 비즈니스 아키텍트입니다. 텍스트 위주의 보고서를 탈피하기 위해 요약 박스와 분석 데이터를 구조화하여 응답하세요.

            <tech_title>핵심 비즈니스 명칭</tech_title>

            <summary_box>
            보고서 최상단에 배치될 핵심 요약입니다. (반드시 아래 수치를 포함)
            1. 최종 기술가치평가액: [금액]
            2. 2028년 예상 SOM(시장규모): [금액]
            3. 핵심 경쟁력 지표: [예: 효율 40% 향상 등]
            </summary_box>
            
            <section_1>Ⅰ. 기술 개요 (상세)</section_1>
            <section_2>Ⅱ. 문제점 및 해결 방안 (상세)</section_2>
            <section_3>Ⅲ. Scale-up 및 기술가치평가 (수식과 근거 포함 상세 기술)</section_3>
            
            <section_4>
            Ⅳ. Virtual Firm 활용 (최종 제안)
            반드시 3C, SWOT, Lean Canvas 내용을 포함하되, 각 분석의 핵심 포인트는 표로 정리될 수 있게 개조식으로 작성하세요.
            </section_4>

            데이터: {tech_text}
            {context}"""

            raw_response, used_model = generate_one_shot(prompt)
            if not raw_response: return

            ai_data = {
                "tech_title": extract_tag(raw_response, "tech_title"),
                "summary": extract_tag(raw_response, "summary_box"),
                "section_1": extract_tag(raw_response, "section_1"),
                "section_2": extract_tag(raw_response, "section_2"),
                "section_3": extract_tag(raw_response, "section_3"),
                "section_4": extract_tag(raw_response, "section_4"),
            }

            # 문서 생성 및 시각화 적용
            doc = Document(DEFAULT_WORD_TEMPLATE) if os.path.exists(DEFAULT_WORD_TEMPLATE) else Document()
            
            # [C안] 요약 박스 삽입
            if ai_data["summary"]:
                add_summary_box(doc, ai_data["summary"])

            # 각 섹션 치환
            replace_placeholder(doc, "{{tech_title}}", ai_data["tech_title"], is_inline=True)
            replace_placeholder(doc, "{{section_1}}", ai_data["section_1"])
            replace_placeholder(doc, "{{section_2}}", ai_data["section_2"])
            replace_placeholder(doc, "{{section_3}}", ai_data["section_3"])
            replace_placeholder(doc, "{{section_4}}", ai_data["section_4"])

            doc_io = io.BytesIO()
            doc.save(doc_io)
            st.download_button("📥 고도화된 보고서 다운로드", doc_io.getvalue(), "Virtual_Firm_Premium_Report.docx")

        except Exception as e:
            st.error(f"오류 발생: {e}")
