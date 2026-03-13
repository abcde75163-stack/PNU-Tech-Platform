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
try:
    MY_API_KEY = st.secrets["GEMINI_API_KEY"].strip()
    client = genai.Client(api_key=MY_API_KEY)
except Exception as e:
    st.error("API 키 로드 실패. Secrets 설정을 확인하세요.")

# 고정 템플릿 파일명 설정
DEFAULT_WORD_TEMPLATE = "default_vf_template.docx"

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

# 3. 레이아웃 보존형 삽입 (줄바꿈 및 스타일 적용)
def add_styled_content_at(target_p, text):
    lines = [line.strip() for line in str(text).split('\n') if line.strip()]
    if not lines: return target_p
    orig_format = target_p.paragraph_format
    left_indent = orig_format.left_indent
    right_indent = orig_format.right_indent
    alignment = target_p.alignment
    current_p = target_p
    for i, line in enumerate(lines):
        if i == 0:
            current_p.text = "" 
            p_to_style = current_p
        else:
            new_p_xml = OxmlElement('w:p')
            current_p._p.addnext(new_p_xml)
            p_to_style = Paragraph(new_p_xml, current_p._parent)
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

# 4. 템플릿 치환
def replace_placeholder(doc, placeholder, content, is_inline=False, font_name=None, font_size=None, is_bold=False):
    def process_p(p):
        if placeholder in p.text:
            if is_inline:
                p.text = p.text.replace(placeholder, str(content))
                for run in p.runs: set_font(run, font_name, font_size, bold=is_bold)
            else: add_styled_content_at(p, content)
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

# 5. 메인 실행 함수
def run_virtual_firm(spec_file, doc_template, target_corp, ir_data, business_status):
    if not spec_file:
        st.error("특허 명세서 파일이 필요합니다.")
        return

    st.subheader(f"🏢 {target_corp if target_corp else 'Virtual Firm'} 심층 보고서 생성")
    tech_text = extract_text_from_file(spec_file)
    if len(tech_text.strip()) < 50:
        st.error("❌ 파일에서 텍스트를 읽을 수 없습니다. (스캔본 여부 확인)")
        return

    with st.spinner("🚀 고퀄리티 심층 분석 및 텍스트 나열형 보고서 작성 중..."):
        try:
            ir_text = extract_text_from_file(ir_data) if ir_data else ""
            context = f"타겟 기업: {target_corp}\nIR 데이터: {ir_text}\n사업 현황: {business_status}"
            
            # 💡 [프롬프트 고도화] 표 작성 금지 및 글로 나열 지시
            prompt = f"""당신은 국내 최고의 기술사업화 전략가이자 가치평가 전문가입니다.
            반드시 제공된 [특허 명세서]의 실제 공학적 내용을 바탕으로 '글로 나열'하는 방식의 심층 보고서를 작성하세요.

            [특허 명세서]
            {tech_text}

            [작성 지침 - 절대 엄수]
            1. 마크다운 표(|---| 등)를 절대 사용하지 마세요. 모든 정보는 소제목(##)과 개조식(Bullet points) 텍스트로만 나열하세요.
            2. 응답은 반드시 <tech_title>, <section_1>, <section_2>, <section_3>, <section_4> 태그로 감싸세요.
            3. 각 섹션은 최소 5~8개 이상의 상세 문단(또는 긴 개조식 목록)으로 매우 풍성하게 작성하세요.
            
            [섹션별 구성 가이드]
            - <section_3>: 단계별 스케일업(Scale-up) 로드맵을 연도별로 상세히 글로 나열하고, 기술가치평가(수익접근법) 수식과 함께 해당 수치가 도출된 가정(단가, 판매량, 점유율 등)을 구체적 근거와 함께 서술하세요.
            - <section_4>: 3C 및 SWOT 분석 내용을 개조식으로 상세 기술하고, 최종적으로 이 기술을 '기술이전'할지 '직접 창업'할지 명확히 선택하여 그 이유를 정량적/정성적 근거를 들어 논리적으로 제안하세요.

            [기타 정보]
            {context}"""

            response = client.models.generate_content(
                model="models/gemini-2.5-flash-lite", 
                contents=prompt
            )
            raw_response = response.text.strip()

            def extract_tag(text, tag_name):
                pattern = f"<{tag_name}>(.*?)(?:</{tag_name}>|$)"
                match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
                return match.group(1).strip() if match else ""

            ai_data = {t: extract_tag(raw_response, t) for t in ["tech_title", "section_1", "section_2", "section_3", "section_4"]}
            doc = Document(doc_template if doc_template else DEFAULT_WORD_TEMPLATE)

            replace_placeholder(doc, "[[tech_title]]", ai_data['tech_title'], is_inline=True, 
                                font_name="KoPub돋움체_Pro Bold", font_size=18, is_bold=True)
            for i in range(1, 5):
                replace_placeholder(doc, "{{" + f"section_{i}" + "}}", ai_data[f'section_{i}'])

            doc_io = io.BytesIO()
            doc.save(doc_io)
            st.success("✅ 심층 보고서 작성이 완료되었습니다!")
            st.download_button(label="📥 심층 보고서 다운로드", data=doc_io.getvalue(), 
                               file_name=f"VF_Strategic_Report_{target_corp}.docx")
        except Exception as e:
            st.error(f"오류 발생: {str(e)}")



