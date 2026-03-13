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
        line_stripped = re.sub(r'\[SECTION_\d\]|\[TECH_TITLE\]', '', line_stripped).strip()
        if not line_stripped: continue

        if line_stripped.startswith('## '): 
            p = doc.add_paragraph()
            run = p.add_run(line_stripped.replace('## ', ''))
            set_font(run, "KoPub돋움체_Pro Medium", 13, bold=True)
            p.paragraph_format.space_before = Pt(18)
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

def parse_ai_response(text):
    data = {"tech_title": "", "section_1": "", "section_2": "", "section_3": "", "section_4": ""}
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

    with st.spinner("🚀 VC 심사역 수준의 방대한 심층 텍스트를 생성 중입니다... (데이터 중복 차단 완비)"):
        try:
            ir_text = extract_text_from_file(ir_data) if ir_data else ""
            context = f"타겟 기업: {target_corp}\nIR 데이터: {ir_text}\n사업 현황: {business_status}"
            
            # 💡 [프롬프트] 섹션 2에 TAM-SAM-SOM 및 기술 동향 필수 포함 지시 추가
            prompt = f"""당신은 국내 최고의 기술사업화 전략가이자 벤처캐피탈(VC) 수석 심사역입니다.
            제공된 [특허 명세서]를 바탕으로 실제 대규모 투자 유치에 사용될 '초정밀 심층 비즈니스 플랜'을 작성해야 합니다.

            [특허 명세서 원본 데이터]
            {tech_text}

            [작성 지침 - 분량 및 깊이 절대 엄수]
            1. 절대로 요약하거나 짧게 끝내지 마세요. 각 [SECTION]마다 다양한 소제목(##)을 최소 4~5개 이상 사용하고, 하위에 상세한 개조식(-, *) 설명과 분석을 덧붙여 각 섹션이 방대한 분량이 되도록 길고 깊게 작성하세요.
            2. 마크다운 표(|---|)는 에러를 유발하므로 절대 사용하지 마세요. 모든 수치, 표, 분석 결과는 '글로 길게 풀어서' 상세히 나열하세요.
            3. 명세서에 데이터가 부족하더라도, 합리적인 가설과 추론을 통해 내용을 매우 구체적으로 채워 넣으세요.
            4. 응답은 반드시 아래의 5개 [구분자]를 사용하여 나누어야 합니다.

            [TECH_TITLE]
            (기술의 핵심 가치를 보여주는 20자 내외의 임팩트 있는 비즈니스 명칭 한 줄만 작성)
            
            [SECTION_1]
            (기술의 근본적 메커니즘, 화학적/물리적 작동 원리, 기존 기술의 한계점과 본 기술의 혁신적 돌파구, 그리고 확장 가능한 파생 기술 가능성까지 아주 상세히 쪼개어 분석하세요.)
            
            [SECTION_2]
            (글로벌 전방 산업의 메가 트렌드 및 '관련 기술 동향'을 심층 분석하세요. 타겟 고객군의 치명적 페인 포인트(Pain point)를 분석하고, 특히 시장 규모를 'TAM(전체 시장) - SAM(유효 시장) - SOM(수익 시장)' 프레임워크를 적용하여 구체적인 추정 수치와 산출 근거를 텍스트로 아주 상세히 서술하세요.)
            
            [SECTION_3]
            (단기/중기/장기 스케일업 마일스톤을 연도별로 아주 길게 서술하세요. '수익접근법' 기술가치평가 시 제품 단가, 예상 점유율, 초기 투자 비용 등의 가설을 구체적인 숫자로 제시하고 산출 근거를 텍스트로 증명하세요. 마지막에는 반드시 '사업화 중장기 로드맵'을 단계별로 명확하게 텍스트로 정리하여 포함하세요.)
            
            [SECTION_4]
            (내부 역량(3C), 외부 환경(SWOT) 분석과 함께 '린 캔버스(Lean Canvas)'의 9가지 핵심 블록(문제, 고객군, 고유 가치 제안, 솔루션, 채널, 수익원, 비용 구조, 핵심 지표, 경쟁 우위)을 소제목과 개조식 텍스트로 아주 상세히 서술하세요. 최종적으로 이 기술의 특성을 고려하여 '기술이전(Licensing)'과 '직접 창업(Spin-off)' 중 하나를 명확히 선택하고, 투자자를 설득할 수 있는 정량적/정성적 근거를 아주 길고 논리적으로 서술하세요.)

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

            for i, (title_text, key) in enumerate(sections_list):
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
                
                # 마지막 섹션(Ⅳ)이 아닐 경우에만 페이지 넘김을 실행
                if i < len(sections_list) - 1:
                    doc.add_page_break()

            doc_io = io.BytesIO()
            doc.save(doc_io)
            
            st.success("✅ 오류 없이 방대한 분량의 심층 보고서 작성이 완료되었습니다!")
            st.download_button(label="📥 최종 심층 보고서 다운로드 (클릭)", data=doc_io.getvalue(), 
                               file_name=f"VF_Master_Report_{target_corp}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"보고서 생성 중 오류 발생: {str(e)}")




