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

def set_font(run, font_name, size, bold=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold

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

def add_styled_content(doc, text):
    lines = str(text).split('\n')
    for line in lines:
        line_stripped = line.strip()
        
        # 만약 AI가 실수로 구분자 기호를 출력했다면 화면에 보이지 않게 청소
        line_stripped = re.sub(r'\[SECTION_\d\]|\[TECH_TITLE\]', '', line_stripped).strip()
        if not line_stripped: continue

        if line_stripped.startswith('## '): 
            p = doc.add_paragraph()
            run = p.add_run(line_stripped.replace('## ', ''))
            set_font(run, "KoPub돋움체_Pro Medium", 13, bold=True)
            p.paragraph_format.space_before = Pt(12)
            continue
            
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.6
        p.paragraph_format.space_after = Pt(10)
        
        parts = re.split(r'(\*\*.*?\*\*)', line_stripped)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                run = p.add_run(part.replace('**', ''))
                set_font(run, "KoPub돋움체_Pro Medium", 11, bold=True)
            else:
                run = p.add_run(part)
                set_font(run, "KoPub돋움체_Pro Light", 11)

# 💡 [핵심] 중복 복제를 완벽하게 차단하는 스마트 파싱 함수
def parse_ai_response(text):
    data = {"tech_title": "", "section_1": "", "section_2": "", "section_3": "", "section_4": ""}
    
    # [마커] 부터 다음 [마커] 직전까지만 정확히 잘라내어 중복을 0%로 만듦
    pattern = r'\[(TECH_TITLE|SECTION_1|SECTION_2|SECTION_3|SECTION_4)\](.*?)(?=\[(?:TECH_TITLE|SECTION_1|SECTION_2|SECTION_3|SECTION_4)\]|$)'
    matches = re.finditer(pattern, text, re.DOTALL | re.IGNORECASE)
    
    for match in matches:
        key = match.group(1).lower()
        content = match.group(2).strip()
        if key in data:
            data[key] = content
            
    return data

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

    with st.spinner("🚀 데이터 중복 오류를 차단하며 완벽한 문서를 조립 중입니다... (약 15초 소요)"):
        try:
            ir_text = extract_text_from_file(ir_data) if ir_data else ""
            context = f"타겟 기업: {target_corp}\nIR 데이터: {ir_text}\n사업 현황: {business_status}"
            
            # 💡 [프롬프트] 태그 형식을 가장 안전한 대괄호[ ] 구분자로 변경
            prompt = f"""당신은 국내 최고의 기술사업화 전략가이자 가치평가 전문가입니다.
            제공된 [특허 명세서]의 실제 기술 내용을 기반으로 아주 상세하고 전문적인 '심층 분석 보고서'를 작성하세요.

            [특허 명세서 원본 데이터]
            {tech_text}

            [작성 지침 - 절대 엄수]
            1. 마크다운 표(|---| 등)를 절대 사용하지 마세요. 모든 정보는 소제목(##)과 개조식(-, *) 텍스트로만 깔끔하게 작성하세요.
            2. 각 섹션은 3~5개의 핵심 문단(또는 개조식 목록)으로 짜임새 있게 작성하세요.
            3. 응답은 반드시 아래 제공된 5개의 [구분자]를 사용하여 섹션을 명확히 나누어야 합니다. (구분자 형태를 절대 변경하지 마세요)

            [TECH_TITLE]
            (여기에 기술의 핵심을 관통하는 20자 내외의 전문적인 사업화 명칭 작성)
            
            [SECTION_1]
            (여기에 기술의 메커니즘, 작동 원리, 차별성을 아주 상세히 작성)
            
            [SECTION_2]
            (여기에 전방 산업 트렌드 및 해결하려는 과제(Pain point) 작성)
            
            [SECTION_3]
            (여기에 단계별 스케일업 전략 및 '수익접근법' 예상 매출 도출 근거 상세 작성)
            
            [SECTION_4]
            (여기에 3C/SWOT 정밀 분석 및 '기술이전' vs '직접 창업'에 대한 최종 제안 작성)

            [기타 정보]
            {context}"""

            max_retries = 2
            raw_response = ""
            for attempt in range(max_retries):
                try:
                    response = client.models.generate_content(
                        model="models/gemini-2.5-flash-lite", 
                        contents=prompt
                    )
                    raw_response = response.text.strip()
                    break
                except Exception as api_e:
                    if attempt == max_retries - 1:
                        raise api_e
                    time.sleep(3)

            # 💡 [해결] 여기서 문서가 30페이지씩 폭주하는 것을 원천 차단
            ai_data = parse_ai_response(raw_response)
            
            doc = Document()
            for section in doc.sections:
                section.top_margin = Pt(72)
                section.bottom_margin = Pt(72)
                section.left_margin = Pt(72)
                section.right_margin = Pt(72)

            doc.add_paragraph("\n\n")
            title_p = doc.add_paragraph()
            title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            clean_title = ai_data.get("tech_title", "심층 사업화 보고서").replace('#', '').replace('*', '').strip()
            if not clean_title: clean_title = "심층 사업화 보고서"
            
            title_run = title_p.add_run(f"Virtual Firm 심층 사업화 전략 보고서\n\n[{clean_title}]")
            set_font(title_run, "KoPub돋움체_Pro Bold", 18, bold=True)
            
            doc.add_paragraph("\n\n\n")
            info_p = doc.add_paragraph()
            info_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            target_display = target_corp if target_corp else "잠재 수요기업"
            info_run = info_p.add_run(f"분석 대상 기업: {target_display}\n작성 기관: 부산대학교 산학협력단")
            set_font(info_run, "KoPub돋움체_Pro Light", 11)
            
            doc.add_page_break()

            sections_list = [
                ("Ⅰ. 기술 개요 및 메커니즘 분석", "section_1"),
                ("Ⅱ. 시장 트렌드 및 과제 해결 분석", "section_2"),
                ("Ⅲ. Scale-up 및 심층 재무 로드맵", "section_3"),
                ("Ⅳ. 최종 사업화 제안 (이전 vs 창업)", "section_4")
            ]

            for title_text, key in sections_list:
                h_p = doc.add_paragraph()
                set_font(h_p.add_run(f"{title_text}"), "KoPub돋움체_Pro Bold", 15, bold=True)
                h_p.paragraph_format.space_before = Pt(20)
                h_p.paragraph_format.space_after = Pt(10)
                
                content = ai_data.get(key, "")
                if content:
                    add_styled_content(doc, content)
                else:
                    p = doc.add_paragraph()
                    set_font(p.add_run("생성된 내용이 없습니다."), "KoPub돋움체_Pro Light", 11)

            doc_io = io.BytesIO()
            doc.save(doc_io)
            
            st.success("✅ 오류 없이 완벽한 비율의 심층 보고서 작성이 완료되었습니다!")
            st.download_button(label="📥 최종 심층 보고서 다운로드 (클릭)", data=doc_io.getvalue(), 
                               file_name=f"VF_Master_Report_{target_corp}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"보고서 생성 중 오류 발생: {str(e)}")




