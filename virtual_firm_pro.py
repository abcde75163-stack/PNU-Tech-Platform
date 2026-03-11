제공해주신 이미지와 결과 파일(Virtual_Firm_Master_Report (11).docx)을 분석한 결과, 여백이 사라지고 글이 다음 장의 디자인 요소를 덮어버리는 문제는 새로운 단락을 생성할 때 템플릿의 여백 정보(Indentation)를 상속받지 않고, 단순히 addnext()로만 삽입하기 때문입니다. 이를 해결하기 위해 1) 기존 단락의 여백 설정을 100% 복제하고, 2) 템플릿의 고정 디자인(Chapter 구분선 등)을 건드리지 않도록 삽입 로직을 정교화한 최종 수정 코드를 제안합니다. 🛠️ 좌우 여백 및 레이아웃 밀림 방지 최종 코드Pythonimport streamlit as st
import fitz
from google import genai
import io
import re
import os
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

# 1. 폰트 설정 (KoPub돋움체_Pro Bold/Medium/Light 강제 적용)
def set_font(run, font_name, size, bold=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold

# 2. 레이아웃 유지형 콘텐츠 삽입 함수 (여백 보존의 핵심)
def add_styled_content_at(target_p, text):
    lines = [line.strip() for line in str(text).split('\n') if line.strip()]
    if not lines: return target_p

    # 템플릿 단락의 원본 여백 및 스타일 정보 보관
    orig_format = target_p.paragraph_format
    left_indent = orig_format.left_indent
    right_indent = orig_format.right_indent
    alignment = target_p.alignment

    current_p = target_p
    for i, line in enumerate(lines):
        if i == 0:
            current_p.text = "" # {{tag}} 줄을 재활용하여 빈 단락 생성 방지
            p_to_style = current_p
        else:
            # 원본 단락 아래에 새 단락 생성 및 여백 복제
            new_p_xml = OxmlElement('w:p')
            current_p._p.addnext(new_p_xml)
            p_to_style = Paragraph(new_p_xml, current_p._parent)
            
            # [수정] 원본 여백 정보를 새 단락에 강제 이식
            p_to_style.paragraph_format.left_indent = left_indent
            p_to_style.paragraph_format.right_indent = right_indent
            p_to_style.alignment = alignment
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

# 3. 템플릿 치환 (표 내부 [[tech_title]] 18pt Bold 적용 포함)
def replace_placeholder(doc, placeholder, content, is_inline=False, font_name=None, font_size=None, is_bold=False):
    def process_p(p):
        if placeholder in p.text:
            if is_inline:
                p.text = p.text.replace(placeholder, str(content))
                for run in p.runs:
                    set_font(run, font_name, font_size, bold=is_bold)
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

# 4. 실행 함수 (프롬프트 내 분석 항목 강제 유지)
def run_virtual_firm(spec_file, doc_template, target_corp, ir_data, business_status):
    if not spec_file: return
    
    with st.spinner("🚀 분석 및 고품질 레이아웃 보고서 생성 중..."):
        # [AI 생성 로직 - 3C, SWOT, 가치평가 상세 수식 포함]
        prompt = f"""당신은 가치평가사입니다. <tech_title>, <section_1>~<section_4> 태그로 작성하세요.
        - <section_3>: 로드맵, 매출액, 수익접근법 가치산출 수식 상세 포함 [cite: 586-630]
        - <section_4>: 3C, SWOT, Lean Canvas, 최종 제안 포함 [cite: 634-735]
        기업: {target_corp}, 정보: {business_status}"""

        response = client.models.generate_content(model="gemini-2.0-flash", contents=prompt)
        raw_response = response.text.strip()

        def extract_tag(text, tag_name):
            pattern = f"<{tag_name}>(.*?)(?:</{tag_name}>|$)"
            match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
            return match.group(1).strip() if match else ""

        ai_data = {t: extract_tag(raw_response, t) for t in ["tech_title", "section_1", "section_2", "section_3", "section_4"]}
        
        doc = Document(doc_template if doc_template else DEFAULT_WORD_TEMPLATE)

        # [[tech_title]] 치환 - 18pt Bold 적용 
        replace_placeholder(doc, "[[tech_title]]", ai_data['tech_title'], is_inline=True, 
                            font_name="KoPub돋움체_Pro Bold", font_size=18, is_bold=True)
        
        # {{section_x}} 치환 - 여백 및 스타일 복제 
        for i in range(1, 5):
            replace_placeholder(doc, "{{" + f"section_{i}" + "}}", ai_data[f'section_{i}'])

        doc_io = io.BytesIO()
        doc.save(doc_io)
        st.success("✅ 여백 및 분석 내용이 보존된 보고서가 완료되었습니다!")
        st.download_button(label="📥 고퀄리티 보고서 다운로드", data=doc_io.getvalue(), 
                           file_name="Virtual_Firm_Master_Report_Final.docx")
        
except Exception as e:
            st.error(f"오류: {str(e)}")
