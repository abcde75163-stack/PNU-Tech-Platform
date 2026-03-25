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

def add_smart_content(doc, text):
    # 표 마커 기반 정밀 분리
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
                    # 소제목 처리
                    if line_stripped.startswith('## ') or line_stripped.startswith('### '):
                        p = doc.add_paragraph()
                        title_clean = line_stripped.replace('## ', '').replace('### ', '').strip()
                        run = p.add_run(title_clean)
                        set_font(run, "KoPub돋움체_Pro Medium", 13, bold=True)
                        p.paragraph_format.space_before = Pt(12)
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
    # 섹션 마커를 더 엄격하게 파싱 (중첩 방지)
    sections = ["TECH_TITLE", "SECTION_1", "SECTION_2", "SECTION_3", "SECTION_4"]
    data = {s.lower(): "" for s in sections}
    
    for s in sections:
        # 해당 마커부터 다음 마커가 나오기 전까지의 모든 텍스트를 캡처
        pattern = rf"\[{s}\](.*?)\[(?:{'|'.join(sections)})\]"
        match = re.search(pattern, text + "[END]", re.DOTALL | re.IGNORECASE)
        if match:
            data[s.lower()] = match.group(1).strip()
            
    return data

def run_virtual_firm(spec_file, doc_template, target_corp, ir_data, business_status):
    if not spec_file:
        st.error("파일이 필요합니다.")
        return

    st.subheader(f"🏢 {target_corp if target_corp else 'Virtual Firm'} 보고서 고도화 생성")
    tech_text = extract_text_from_file(spec_file)
    
    with st.spinner("🚀 챕터별 내용을 격리하여 정밀 분석 중입니다..."):
        try:
            prompt = f"""당신은 국내 최고의 기술사업화 전문가입니다. 제공된 [특허 명세서]를 기반으로 투자용 보고서를 작성하세요.
            [특허 명세서]: {tech_text}
            
            [절대 준수 지침]
            1. 응답은 반드시 아래 5개의 [마커]로 시작하여 다음 마커가 나오기 전까지 해당 내용만 작성하세요. 
            2. 섹션 간의 내용 섞임을 방지하기 위해 각 [SECTION]의 끝에는 반드시 구분선을 넣지 마세요.
            3. 분석은 매우 방대하게 작성하되, 표가 필요한 데이터는 반드시 [TABLE_START]와 [TABLE_END]를 사용하세요.

            [응답 구조]
            [TECH_TITLE] 기술 비즈니스 명칭 한 줄
            [SECTION_1] 기술 메커니즘 및 성능 분석 (기존 기술 대비 우위 표 포함)
            [SECTION_2] 시장 트렌드 및 TAM-SAM-SOM 분석 (시장 규모 추정 표 포함)
            [SECTION_3] 스케일업 및 재무 로드맵 (연도별 추정 및 로드맵 표 포함)
            [SECTION_4] 전략 제안 (SWOT 및 Lean Canvas 9개 블록 표 포함)"""

            max_retries = 3
            raw_response = ""
            for attempt in range(max_retries):
                try:
                    response = client.models.generate_content(model="models/gemini-2.5-flash-lite", contents=prompt)
                    raw_response = response.text.strip()
                    if "[SECTION_4]" in raw_response: break # 정상 응답 확인
                except:
                    time.sleep(5)

            # 섹션별 텍스트 분리 로직 강화
            ai_data = parse_ai_response(raw_response)
            
            doc = Document()
            # 표지 로직 생략 (기존과 동일)
            title_p = doc.add_paragraph("\n\n")
            title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_text = ai_data.get("tech_title", "심층 분석 보고서").split('\n')[0]
            title_run = title_p.add_run(f"Virtual Firm 심층 전략 보고서\n\n[{title_text}]")
            set_font(title_run, "KoPub돋움체_Pro Bold", 18, bold=True)
            doc.add_page_break()

            sections_map = [
                ("Ⅰ. 기술 개요 및 메커니즘 분석", "section_1"),
                ("Ⅱ. 시장 트렌드 및 TAM-SAM-SOM 분석", "section_2"),
                ("Ⅲ. Scale-up 및 심층 재무 로드맵", "section_3"),
                ("Ⅳ. 최종 사업화 제안 (Lean Canvas 포함)", "section_4")
            ]

            for i, (title_text, key) in enumerate(sections_map):
                h_p = doc.add_paragraph()
                set_font(h_p.add_run(f"{title_text}"), "KoPub돋움체_Pro Bold", 15, bold=True)
                
                content = ai_data.get(key, "")
                if content:
                    add_smart_content(doc, content)
                else:
                    doc.add_paragraph("해당 섹션의 데이터를 생성하는 중 오류가 발생했습니다.")
                
                if i < len(sections_map) - 1:
                    doc.add_page_break()

            doc_io = io.BytesIO()
            doc.save(doc_io)
            st.success("✅ 챕터별 내용이 격리된 정밀 보고서 작성이 완료되었습니다!")
            st.download_button(label="📥 최종 보고서 다운로드", data=doc_io.getvalue(), file_name=f"VF_Refined_Report.docx")
        except Exception as e:
            st.error(f"오류 발생: {str(e)}")
