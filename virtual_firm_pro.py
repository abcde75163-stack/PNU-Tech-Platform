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

def set_font(run, font_name, size, bold=False, color=None):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)

def extract_text_from_pdf(uploaded_file):
    if uploaded_file is None: return ""
    uploaded_file.seek(0)
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    return "".join([page.get_text() for page in doc])[:15000]

def add_styled_content_at(target_p, text, doc=None):
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
            set_font(run, "KoPub돋움체_Pro Medium", 13, bold=True, color=(0, 51, 153))
            new_p.paragraph_format.space_before = Pt(12)
        elif line_stripped.startswith('### '):
            run = new_p.add_run(line_stripped.replace('### ', ''))
            set_font(run, "KoPub돋움체_Pro Medium", 12, bold=True)
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

# ★ 수정됨: is_inline 대체 시 폰트와 크기를 강제로 주입할 수 있도록 개선
def replace_placeholder(doc, placeholder, content, is_inline=False, font_name=None, font_size=None, is_bold=False):
    # 1. 일반 단락 검사
    for p in doc.paragraphs:
        if placeholder in p.text:
            if is_inline:
                text_parts = p.text.split(placeholder)
                p.clear() # 단락의 정렬 속성은 유지하면서 기존 텍스트(런)만 삭제
                for i, part in enumerate(text_parts):
                    if part:
                        p.add_run(part)
                    if i < len(text_parts) - 1:
                        run = p.add_run(content)
                        if font_name and font_size:
                            set_font(run, font_name, font_size, is_bold)
            else:
                p.text = p.text.replace(placeholder, "") 
                if content: add_styled_content_at(p, content, doc)    
            return True
            
    # 2. 표 내부 단락 검사
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if placeholder in p.text:
                        if is_inline:
                            text_parts = p.text.split(placeholder)
                            p.clear()
                            for i, part in enumerate(text_parts):
                                if part:
                                    p.add_run(part)
                                if i < len(text_parts) - 1:
                                    run = p.add_run(content)
                                    if font_name and font_size:
                                        set_font(run, font_name, font_size, is_bold)
                        else:
                            p.text = p.text.replace(placeholder, "")
                            if content: add_styled_content_at(p, content, doc)
                        return True
    return False

def extract_tag(text, tag_name):
    pattern = f"<{tag_name}>(.*?)(?:</{tag_name}>|$)"
    match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    return match.group(1).strip() if match else ""

def generate_one_shot(prompt):
    fallback_models = ["gemini-2.5-flash", "gemini-1.5-pro", "gemini-1.5-flash"]
    last_error = ""
    for model_name in fallback_models:
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                    config={"temperature": 0.2} 
                )
                if response and response.text:
                    return response.text.strip(), model_name
            except Exception as e:
                error_msg = str(e)
                last_error = error_msg
                if "429" in error_msg or "RESOURCE_EXHAUSTED" in error_msg:
                    time.sleep(3) 
                    continue
                else:
                    time.sleep(1)
                    continue
    return None, f"생성 실패 (에러: {last_error[:100]})"

def run_virtual_firm(spec_file, doc_template, target_corp, ir_data, business_status):
    if not spec_file:
        st.error("분석을 위한 기술 명세서(PDF)가 필요합니다.")
        return

    st.subheader(f"🏢 {target_corp if target_corp else 'Virtual Firm'} 전략 보고서 생성 중...")
    
    with st.spinner("🚀 완벽한 비즈니스 이전을 위한 심층 사업 타당성 및 기술가치평가 산출을 진행 중입니다... (약 20~30초 소요)"):
        try:
            tech_text = extract_text_from_pdf(spec_file)
            ir_text = extract_text_from_pdf(ir_data) if ir_data else ""
            context = f"대상기업: {target_corp}\n기업IR자료: {ir_text}\n사업현황: {business_status}"
            
            prompt = f"""당신은 부산대학교 기술지주회사의 최고급 비즈니스 아키텍트이자 공인 기술가치평가사입니다.
            
            [작성 규칙 - 절대 엄수]
            1. 마크다운 표(| 구분 | 내용 | 등)를 절대 사용하지 마세요.
            2. 비교 분석, 재무 수치, 로드맵 등은 반드시 소제목(##, ###)과 개조식(-, *)을 활용한 텍스트로만 서술하세요.
            3. 응답은 반드시 아래 5개의 <태그>를 모두 열고 닫아야 합니다.

            <tech_title>이 기술을 기반으로 한 가상 기업의 비즈니스 명칭 (20자 내외 한 줄)</tech_title>
            
            <section_1>
            ## Ⅰ. 기술 개요
            (1,000자 이상 심층 작성. 기술의 작동 원리와 압도적 경쟁력, 비즈니스적 해자(Moat)를 상세히 설명하세요.)
            </section_1>
            
            <section_2>
            ## Ⅱ. 문제점 및 해결 방안
            (800자 이상 심층 작성. 타겟 시장의 치명적인 한계점과 혁신적 해결책을 상세히 설명하세요.)
            </section_2>
            
            <section_3>
            ## Ⅲ. Scale-up 및 기술가치평가
            (1,500자 이상 아주 상세히 작성)
            ### 1. 사업화 로드맵
            ### 2. 시장 규모 추정 및 예상 매출액
            ### 3. 기술가치평가 산출
            - 평가 방법론 명시 및 예상매출액 × 로열티율 × 기술기여도 등 실제 계산식을 직접 풀어서 전개하고 산출 근거를 상세히 쓰세요.
            </section_3>
            
            <section_4>
            ## Ⅳ. Virtual Firm 활용 (최종 제안)
            (1,500자 이상 아주 상세히 작성)
            ### 1. 3C 분석
            ### 2. SWOT 분석
            ### 3. Lean Canvas
            ### 4. 최종 창업 및 이전 전략 제안
            </section_4>

            데이터: {tech_text}
            {context}"""

            raw_response, used_model = generate_one_shot(prompt)
            
            if not raw_response:
                st.error(f"⚠️ AI 응답 생성에 실패했습니다. 사유: {used_model}")
                return

            st.toast(f"✅ 심층 비즈니스 분석 완료! (사용 모델: {used_model})")

            ai_data = {
                "tech_title": extract_tag(raw_response, "tech_title"),
                "section_1": extract_tag(raw_response, "section_1"),
                "section_2": extract_tag(raw_response, "section_2"),
                "section_3": extract_tag(raw_response, "section_3"),
                "section_4": extract_tag(raw_response, "section_4"),
            }

            doc = Document(DEFAULT_WORD_TEMPLATE) if os.path.exists(DEFAULT_WORD_TEMPLATE) else Document()
            
            # ★ 수정됨: tech_title 삽입 시 "KoPub돋움체_Pro Bold", 18pt 강제 지정
            replace_placeholder(
                doc, 
                "{{tech_title}}", 
                ai_data["tech_title"], 
                is_inline=True, 
                font_name="KoPub돋움체_Pro Bold", 
                font_size=18, 
                is_bold=True
            )
            
            replace_placeholder(doc, "{{section_1}}", ai_data["section_1"])
            replace_placeholder(doc, "{{section_2}}", ai_data["section_2"])
            replace_placeholder(doc, "{{section_3}}", ai_data["section_3"])
            replace_placeholder(doc, "{{section_4}}", ai_data["section_4"])

            doc_io = io.BytesIO()
            doc.save(doc_io)
            st.download_button(
                label="📥 전략 보고서 다운로드 (안정화 버전)", 
                data=doc_io.getvalue(), 
                file_name="Virtual_Firm_Master_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"문서 생성 중 오류 발생: {e}")
