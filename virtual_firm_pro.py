import streamlit as st
import fitz
from google import genai
import io
import re
import time
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# [설정 - API]
try:
    MY_API_KEY = st.secrets["GEMINI_API_KEY"].strip()
    client = genai.Client(api_key=MY_API_KEY)
except Exception as e:
    st.error("API 키 로드 실패. Secrets 설정을 확인하세요.")

def set_font(run, font_name, size, bold=False, color=None):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)

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

# 💡 [핵심] 텍스트와 표를 구분하여 삽입하는 지능형 함수
defdef add_smart_content(doc, text):
    # 대소문자 구분 없이, 그리고 공백이 섞여도 잘 찾도록 정규식 보강
    table_pattern = r'\[\s*TABLE_START\s*\](.*?)\[\s*TABLE_END\s*\]'
    segments = re.split(table_pattern, text, flags=re.DOTALL | re.IGNORECASE)
    
    for i, seg in enumerate(segments):
        seg = seg.strip()
        if not seg: continue
        
        if i % 2 == 0:
            # 일반 텍스트 처리 로직 (기존과 동일)
            lines = seg.split('\n')
            for line in lines:
                line_stripped = line.strip()
                if not line_stripped: continue
                
                if line_stripped.startswith('## '):
                    p = doc.add_paragraph()
                    run = p.add_run(line_stripped.replace('## ', ''))
                    set_font(run, "KoPub돋움체_Pro Medium", 13, bold=True)
                    p.paragraph_format.space_before = Pt(15)
                else:
                    p = doc.add_paragraph()
                    p.paragraph_format.line_spacing = 1.6
                    parts = re.split(r'(\*\*.*?\*\*)', line_stripped)
                    for part in parts:
                        if part.startswith('**') and part.endswith('**'):
                            run = p.add_run(part.replace('**', ''))
                            set_font(run, "KoPub돋움체_Pro Medium", 11, bold=True)
                        else:
                            run = p.add_run(part)
                            set_font(run, "KoPub돋움체_Pro Light", 11)
        else:
            # 표 생성 로직 보강
            rows_data = [line.strip() for line in seg.split('\n') if '|' in line]
            if len(rows_data) < 1: continue
            
            grid = []
            for r in rows_data:
                cells = [c.strip() for c in r.split('|') if c.strip()]
                if cells: grid.append(cells)
            
            if not grid: continue
            
            # 열 개수가 불일치할 경우를 대비한 안전장치
            max_cols = max(len(row) for row in grid)
            table = doc.add_table(rows=len(grid), cols=max_cols)
            table.style = 'Table Grid'
            
            for r_idx, row_content in enumerate(grid):
                for c_idx, cell_value in enumerate(row_content):
                    if c_idx < max_cols:
                        cell = table.cell(r_idx, c_idx)
                        cell.text = cell_value
                    # 표 내부 폰트 설정
                    for para in cell.paragraphs:
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in para.runs:
                            set_font(run, "KoPub돋움체_Pro Light", 10, bold=(r_idx == 0))

def parse_ai_response(text):
    data = {"tech_title": "", "section_1": "", "section_2": "", "section_3": "", "section_4": ""}
    pattern = r'\[(TECH_TITLE|SECTION_1|SECTION_2|SECTION_3|SECTION_4)\](.*?)(?=\[(?:TECH_TITLE|SECTION_1|SECTION_2|SECTION_3|SECTION_4)\]|$)'
    matches = re.finditer(pattern, text, re.DOTALL | re.IGNORECASE)
    for match in matches:
        key = match.group(1).lower()
        content = match.group(2).strip()
        if key in data: data[key] = content
    return data

def run_virtual_firm(spec_file, doc_template, target_corp, ir_data, business_status):
    if not spec_file:
        st.error("특허 명세서 파일이 필요합니다.")
        return

    st.subheader(f"🏢 {target_corp if target_corp else 'Virtual Firm'} 심층 보고서 생성")
    tech_text = extract_text_from_file(spec_file)
    
    with st.spinner("🚀 제미나이가 표와 텍스트를 자율적으로 구성하여 분석 중입니다..."):
        try:
            ir_text = extract_text_from_file(ir_data) if ir_data else ""
            context = f"타겟 기업: {target_corp}\nIR 데이터: {ir_text}\n사업 현황: {business_status}"
            
            prompt = f"""당신은 국내 최고의 기술사업화 전문가입니다. 
            제공된 [특허 명세서]를 기반으로 투자용 심층 보고서를 작성하세요.

            [특허 명세서]
            {tech_text}

            [작성 지침 - 표(Table) 활용 가이드]
            1. 텍스트로만 설명하기 복잡한 데이터(시장규모 TAM-SAM-SOM, 성능 비교, 재무 수치, SWOT, Lean Canvas 등)는 반드시 아래 형식을 사용하여 '표'로 구성하세요.
            2. 표 시작은 [TABLE_START], 끝은 [TABLE_END]로 표시하고 내부는 | 기호로 구분하세요.
               예: [TABLE_START]
                   구분 | 내용 | 비고
                   TAM | 1,000억 | 글로벌 시장
                   [TABLE_END]
            3. 각 섹션은 매우 상세하게 작성하되, 섹션당 최소 1개 이상의 표를 자율적으로 설계하여 포함하세요.

            [SECTION 구분자]
            [TECH_TITLE] (임팩트 있는 명칭)
            [SECTION_1] (기술 메커니즘 분석 - 기존 기술과의 성능 비교 표 포함)
            [SECTION_2] (시장 분석 - TAM-SAM-SOM 추정 수치 표 포함)
            [SECTION_3] (스케일업 및 재무 로드맵 - 연도별 재무 추정 표 포함)
            [SECTION_4] (전략 제안 - SWOT 및 Lean Canvas 표 포함)

            [기타 정보]
            {context}"""

            response = client.models.generate_content(model="models/gemini-2.5-flash-lite", contents=prompt)
            ai_data = parse_ai_response(response.text.strip())
            
            doc = Document()
            # 제목 및 표지 생성 (생략 - 기존 로직과 동일)
            title_p = doc.add_paragraph("\n\n")
            title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_run = title_p.add_run(f"Virtual Firm 심층 사업화 전략 보고서\n\n[{ai_data['tech_title']}]")
            set_font(title_run, "KoPub돋움체_Pro Bold", 18, bold=True)
            doc.add_page_break()

            sections_list = [
                ("Ⅰ. 기술 개요 및 메커니즘 분석", "section_1"),
                ("Ⅱ. 시장 트렌드 및 TAM-SAM-SOM 분석", "section_2"),
                ("Ⅲ. Scale-up 및 심층 재무 로드맵", "section_3"),
                ("Ⅳ. 최종 사업화 제안 (Lean Canvas 포함)", "section_4")
            ]

            for i, (title_text, key) in enumerate(sections_list):
                h_p = doc.add_paragraph()
                set_font(h_p.add_run(f"{title_text}"), "KoPub돋움체_Pro Bold", 15, bold=True)
                
                content = ai_data.get(key, "")
                if content:
                    add_smart_content(doc, content)
                
                if i < len(sections_list) - 1:
                    doc.add_page_break()

            doc_io = io.BytesIO()
            doc.save(doc_io)
            st.success("✅ 표와 텍스트가 조화된 심층 보고서 작성이 완료되었습니다!")
            st.download_button(label="📥 최종 보고서 다운로드", data=doc_io.getvalue(), file_name=f"VF_Master_Report_{target_corp}.docx")
        except Exception as e:
            st.error(f"오류 발생: {str(e)}")
