아이고, 이 부분은 제 불찰입니다! 🙏

에러 메시지("no style with name 'Light List Accent 1'")가 발생한 이유는, 제가 요약 박스를 예쁘게 꾸미기 위해 지정한 'Light List Accent 1'이라는 표 스타일이 위원님께서 사용 중인 기본 워드 템플릿(default_vf_template.docx) 안에 등록되어 있지 않기 때문입니다.

워드 파일마다 내장된 테마(스타일) 이름이 달라서 발생하는 흔한 충돌입니다. 이를 해결하기 위해, 어느 워드 파일에서나 100% 작동하는 가장 기본 스타일인 'Table Grid'로 변경하고, 파이썬 코드로 직접 **'옅은 회색 배경'**을 칠해 요약 박스 느낌을 내도록 코드를 수정했습니다.

아래 최종 수정된 코드로 다시 한번 전체 덮어쓰기를 부탁드립니다!

💻 완벽 수정된 virtual_firm_pro.py (표 스타일 에러 해결)
Python
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

# 3. [C안 구현] 핵심 지표 요약 박스 생성 함수 (★ 스타일 에러 수정)
def add_summary_box(doc, summary_text):
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid' # 모든 워드에서 지원하는 기본 스타일로 변경
    
    cell = table.cell(0, 0)
    
    # 옅은 회색 배경 직접 추가
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), 'F2F2F2')
    cell._tc.get_or_add_tcPr().append(shading_elm)

    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    lines = summary_text.split('\n')
    for line in lines:
        if not line.strip(): continue
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(line.strip())
        set_font(run, "KoPub돋움체_Pro Medium", 11, bold=True)
    
    doc.add_paragraph()

# 4. 콘텐츠 삽입 및 스타일링
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

def replace_placeholder(doc, placeholder, content, is_inline=False):
    for p in doc.paragraphs:
        if placeholder in p.text:
            if is_inline:
                p.text = p.text.replace(placeholder, content)
            else:
                p.text = p.text.replace(placeholder, "") 
                if content: add_styled_content_at(p, content, doc)    
            return True
    return False

def extract_tag(text, tag_name):
    pattern = f"<{tag_name}>(.*?)(?:</{tag_name}>|$)"
    match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    return match.group(1).strip() if match else ""

# 5. 스마트 라우터
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

# ★ 메인 실행 함수
def run_virtual_firm(spec_file, doc_template, target_corp, ir_data, business_status):
    if not spec_file:
        st.error("분석을 위한 기술 명세서(PDF)가 필요합니다.")
        return

    st.subheader(f"🏢 {target_corp if target_corp else 'Virtual Firm'} 프리미엄 전략 보고서 생성 중...")
    
    with st.spinner("📊 핵심 지표 요약 박스 및 비즈니스 표를 구성하고 있습니다... (약 20~30초 소요)"):
        try:
            tech_text = extract_text_from_pdf(spec_file)
            ir_text = extract_text_from_pdf(ir_data) if ir_data else ""
            context = f"대상기업: {target_corp}\n기업IR자료: {ir_text}\n사업현황: {business_status}"
            
            prompt = f"""당신은 최고급 비즈니스 아키텍트이자 공인 기술가치평가사입니다.
            텍스트 위주의 보고서를 탈피하기 위해 핵심 요약 박스와 체계적인 구조로 응답하세요.

            [작성 규칙]
            1. 마크다운 표(| | |)는 에러를 유발하므로 절대 사용하지 말고, 소제목(##)과 개조식(-)으로 깔끔하게 구조화하세요.
            2. 응답은 반드시 5개의 <태그>를 모두 열고 닫아야 합니다.

            <tech_title>핵심 비즈니스 명칭 (20자 내외)</tech_title>

            <summary_box>
            보고서 최상단에 배치될 핵심 요약입니다. (반드시 아래 3가지를 포함하여 짧고 강렬하게 작성)
            1. 예상 기술가치평가액: [금액]
            2. 타겟 시장 규모(TAM/SOM): [금액]
            3. 핵심 경쟁력 지표: [예: 효율 40% 향상 등]
            </summary_box>
            
            <section_1>
            Ⅰ. 기술 개요
            - 기술의 작동 원리와 압도적 경쟁력
            - 비즈니스적 해자(Moat)
            </section_1>
            
            <section_2>
            Ⅱ. 문제점 및 해결 방안
            - 타겟 시장의 치명적인 한계점
            - 본 기술의 혁신적 해결책
            </section_2>
            
            <section_3>
            Ⅲ. Scale-up 및 기술가치평가
            ## 1. 사업화 로드맵
            - 단계별 타임라인
            ## 2. 시장 규모 및 예상 매출액
            - 추정치 및 산출 근거
            ## 3. 기술가치평가 산출 (수식 상세 전개)
            - 수익접근법 등 평가 방법론 명시
            - 산출 과정: 예상매출액 × 로열티율 × 기술기여도 등 실제 계산식을 직접 풀어서 상세히 서술
            </section_3>
            
            <section_4>
            Ⅳ. Virtual Firm 활용 (최종 제안)
            ## 1. 3C 분석 (자사/경쟁사/고객)
            ## 2. SWOT 분석 (강점/약점/기회/위협)
            ## 3. Lean Canvas (9대 핵심 요소)
            ## 4. 최종 창업 및 이전 전략 제안
            </section_4>

            데이터: {tech_text}
            {context}"""

            raw_response, used_model = generate_one_shot(prompt)
            
            if not raw_response:
                st.error(f"⚠️ AI 응답 생성에 실패했습니다. 사유: {used_model}")
                return

            st.toast(f"✅ 프리미엄 보고서 생성 완료! (사용 모델: {used_model})")

            ai_data = {
                "tech_title": extract_tag(raw_response, "tech_title"),
                "summary": extract_tag(raw_response, "summary_box"),
                "section_1": extract_tag(raw_response, "section_1"),
                "section_2": extract_tag(raw_response, "section_2"),
                "section_3": extract_tag(raw_response, "section_3"),
                "section_4": extract_tag(raw_response, "section_4"),
            }

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
            st.download_button(
                label="📥 프리미엄 전략 보고서 다운로드", 
                data=doc_io.getvalue(), 
                file_name="Virtual_Firm_Premium_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"문서 생성 중 오류 발생: {e}")
