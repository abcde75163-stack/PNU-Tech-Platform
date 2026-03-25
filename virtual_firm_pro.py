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

# 📊 워드 정식 표 생성 공통 함수
def create_word_table(doc, table_text):
    rows_data = [line.strip() for line in table_text.split('\n') if '|' in line]
    grid = []
    for r in rows_data:
        cells = [c.strip() for c in r.split('|') if c.strip()]
        if cells: grid.append(cells)
    
    if not grid: return
    
    max_cols = max(len(row) for row in grid)
    table = doc.add_table(rows=len(grid), cols=max_cols)
    table.style = 'Table Grid'
    
    for r_idx, row_content in enumerate(grid):
        for c_idx, cell_value in enumerate(row_content):
            if c_idx < max_cols:
                cell = table.cell(r_idx, c_idx)
                cell.text = cell_value
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in para.runs:
                        set_font(run, "KoPub돋움체_Pro Light", 10, bold=(r_idx == 0))

# 💡 텍스트와 표를 지능적으로 분리하여 삽입
def add_smart_content(doc, text):
    table_pattern = r'\[\s*TABLE_START\s*\](.*?)\[\s*TABLE_END\s*\]'
    segments = re.split(table_pattern, text, flags=re.DOTALL | re.IGNORECASE)
    
    for i, seg in enumerate(segments):
        seg = seg.strip()
        if not seg: continue
        
        if i % 2 == 0:
            lines = seg.split('\n')
            table_buffer = []
            for line in lines:
                if line.count('|') >= 2:
                    table_buffer.append(line)
                else:
                    if table_buffer:
                        create_word_table(doc, "\n".join(table_buffer))
                        table_buffer = []
                    
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
                        run = p.add_run(line_stripped)
                        set_font(run, "KoPub돋움체_Pro Light", 11)
            if table_buffer:
                create_word_table(doc, "\n".join(table_buffer))
        else:
            create_word_table(doc, seg)

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

    st.subheader(f"🏢 {target_corp if target_corp else 'Virtual Firm'} 심층 보고서")
    tech_text = extract_text_from_file(spec_file)
    
    with st.spinner("🚀 서버 부하를 체크하며 심층 보고서를 생성 중입니다..."):
        try:
            ir_text = extract_text_from_file(ir_data) if ir_data else ""
            context = f"타겟 기업: {target_corp}\nIR 데이터: {ir_text}\n사업 현황: {business_status}"
            
            prompt = f"""당신은 국내 최고의 기술사업화 전략가입니다. 제공된 [특허 명세서]를 기반으로 투자용 보고서를 작성하세요.
            [특허 명세서]: {tech_text}
            
            [핵심 지시]
            1. 모든 데이터 분석(성능 비교, TAM-SAM-SOM, 재무 추정, Lean Canvas)은 반드시 표(|---|) 형식으로 작성하세요.
            2. 표 앞뒤에는 반드시 [TABLE_START]와 [TABLE_END] 마커를 붙이세요.
            3. 각 섹션은 매우 방대한 텍스트 분석(최소 8문단 이상)을 포함해야 합니다.
            
            [구조]
            [TECH_TITLE], [SECTION_1], [SECTION_2], [SECTION_3], [SECTION_4] 마커를 반드시 유지하세요."""

            max_retries = 3
            raw_response = ""
            for attempt in range(max_retries):
                try:
                    response = client.models.generate_content(model="models/gemini-2.5-flash-lite", contents=prompt)
                    raw_response = response.text.strip()
                    break
                except Exception:
                    if attempt == max_retries - 1: raise
                    time.sleep((attempt + 1) * 5)

            ai_data = parse_ai_response(raw_response)
            doc = Document()
            for s in doc.sections:
                s.top_margin = Pt(72); s.bottom_margin = Pt(72); s.left_margin = Pt(72); s.right_margin = Pt(72)

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
                h_p.paragraph_format.space_before = Pt(20)
                
                content = ai_data.get(key, "")
                if content:
                    add_smart_content(doc, content)
                
                if i < len(sections_list) - 1:
                    doc.add_page_break()

            doc_io = io.BytesIO()
            doc.save(doc_io)
            st.success("✅ 보고서 작성이 완료되었습니다!")
            st.download_button(label="📥 최종 보고서 다운로드", data=doc_io.getvalue(), file_name=f"VF_Master_Report.docx")
        except Exception as e:
            st.error(f"오류 발생: {str(e)}")
