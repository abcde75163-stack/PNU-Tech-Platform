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
# st.secrets 사용 시 키 이름이 정확한지 확인하세요.
try:
    MY_API_KEY = st.secrets["GEMINI_API_KEY"].strip()
    client = genai.Client(api_key=MY_API_KEY)
except Exception as e:
    st.error("API 키를 로드할 수 없습니다. Secrets 설정을 확인해주세요.")

# 고정 템플릿 파일명 설정
DEFAULT_WORD_TEMPLATE = "default_vf_template.docx"

# 1. 폰트 설정 함수
def set_font(run, font_name, size, bold=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold

# 2. 파일 텍스트 추출 (호환성 및 추출 확인 기능 강화)
def extract_text_from_file(uploaded_file):
    if uploaded_file is None: return ""
    file_name = uploaded_file.name.lower()
    uploaded_file.seek(0)
    try:
        if file_name.endswith('.pdf'):
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            text = "".join([page.get_text() for page in doc])
            return text[:15000] # 토큰 한계 고려
        elif file_name.endswith('.docx'):
            doc = Document(uploaded_file)
            return "\n".join([p.text for p in doc.paragraphs])[:15000]
    except Exception as e:
        st.error(f"파일 추출 오류: {str(e)}")
        return ""
    return ""

# 3. 레이아웃 보존형 삽입 함수 (수정 없음 - 유지)
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

# 4. 템플릿 치환 함수 (수정 없음 - 유지)
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

# 5. 메인 실행 함수 (모델명 및 프롬프트 대폭 강화)
def run_virtual_firm(spec_file, doc_template, target_corp, ir_data, business_status):
    if not spec_file:
        st.error("분석을 위한 기술 명세서 파일이 필요합니다.")
        return

    st.subheader(f"🏢 {target_corp if target_corp else 'Virtual Firm'} 심층 보고서 생성")
    
    # 명세서 텍스트 미리 추출 및 확인
    tech_text = extract_text_from_file(spec_file)
    if len(tech_text.strip()) < 100:
        st.error("파일에서 텍스트를 충분히 읽어오지 못했습니다. 스캔된 이미지 PDF인지 확인해주세요.")
        return

    with st.spinner("🚀 재무 계획 및 심층 로드맵 분석 중..."):
        try:
            ir_text = extract_text_from_file(ir_data) if ir_data else ""
            context = f"타겟 기업: {target_corp if target_corp else '미정'}\n사업 현황: {business_status}"
            
            # 💡 환각 방지 및 재무 근거 강화를 위한 프롬프트 최적화
            prompt = f"""당신은 최고급 기술가치평가 전문가입니다. 
            반드시 아래 제공된 [특허 명세서]의 실제 기술 내용만을 바탕으로 분석을 수행하세요. 절대 다른 기술을 지어내지 마세요.

            [특허 명세서]
            {tech_text}

            [작성 지침]
            - 응답은 반드시 <tech_title>, <section_1>, <section_2>, <section_3>, <section_4> 태그 형식을 유지하세요.
            - <section_3>: 사업화 로드맵과 5개년 예상 매출액을 제시하고, '수익접근법'에 기반한 기술가치 산출 수식(할인율, 기술기여도 등)과 그 근거를 구체적 수치로 상세히 기술하세요.
            - <section_4>: 3C 분석, SWOT 분석을 텍스트 기반 개조식으로 작성하고, 최종적으로 '기술이전'과 '직접 창업' 중 최적안을 선택하여 그 이유를 논리적으로 제안하세요.

            [기타 정보]
            {context}
            기업 IR 참고: {ir_text}"""

            # 💡 [핵심 수정] 모델명 형식을 소문자 표준 규격으로 변경
            # 만약 2.5 Flash Lite가 지속적으로 400 에러를 낸다면 "gemini-1.5-flash"로 교체해 보세요.
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
            
            # 템플릿 로드 실패 방지
            template_path = doc_template if doc_template else DEFAULT_WORD_TEMPLATE
            if not os.path.exists(template_path):
                # 템플릿이 없을 경우 새 문서 생성 (에러 방지용)
                doc = Document()
                doc.add_paragraph("[[tech_title]]")
                for i in range(1, 5): doc.add_paragraph(f"{{{{section_{i}}}}}")
            else:
                doc = Document(template_path)

            # 데이터 치환
            replace_placeholder(doc, "[[tech_title]]", ai_data['tech_title'], is_inline=True, 
                                font_name="KoPub돋움체_Pro Bold", font_size=18, is_bold=True)
            
            for i in range(1, 5):
                tag_str = "{{" + f"section_{i}" + "}}"
                replace_placeholder(doc, tag_str, ai_data[f'section_{i}'])

            doc_io = io.BytesIO()
            doc.save(doc_io)
            st.success("✅ 심층 보고서 생성이 완료되었습니다!")
            st.download_button(label="📥 보고서 다운로드", data=doc_io.getvalue(), 
                               file_name=f"Virtual_Firm_Report_{target_corp}.docx")
        except Exception as e:
            st.error(f"보고서 생성 중 오류 발생: {str(e)}")



