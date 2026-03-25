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
    # AI 응답에서 불필요한 마커 제거
    text = re.sub(r'\[SECTION_\d\]|\[TECH_TITLE\]|\[TABLE_START\]|\[TABLE_END\]', '', text, flags=re.IGNORECASE)
    
    lines = text.split('\n')
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
            if line_stripped.startswith('## ') or line_stripped.startswith('### '):
                p = doc.add_paragraph()
                run = p.add_run(line_stripped.replace('#', '').strip())
                set_font(run, "KoPub돋움체_Pro Medium", 13, bold=True)
            else:
                p = doc.add_paragraph()
                p.paragraph_format.line_spacing = 1.6
                run = p.add_run(line_stripped)
                set_font(run, "KoPub돋움체_Pro Light", 11)
    if table_buffer:
        create_word_table(doc, "\n".join(table_buffer))

def run_virtual_firm(spec_file, doc_template, target_corp, ir_data, business_status):
    if not spec_file:
        st.error("파일이 필요합니다.")
        return

    st.subheader(f"🏢 {target_corp if target_corp else 'Virtual Firm'} 무결점 심층 보고서")
    tech_text = extract_text_from_file(spec_file)
    ir_text = extract_text_from_file(ir_data) if ir_data else ""
    context = f"타겟 기업: {target_corp}\nIR 데이터: {ir_text}\n사업 현황: {business_status}"

    sections_info = [
        ("TECH_TITLE", "기술의 핵심 가치를 보여주는 20자 내외의 전문적인 사업화 명칭"),
        ("Ⅰ. 기술 개요 및 메커니즘 분석", "기술의 근본 메커니즘, 작동 원리, 기존 기술 대비 차별성을 분석한 표와 상세 설명"),
        ("Ⅱ. 시장 트렌드 및 TAM-SAM-SOM 분석", "글로벌 산업 트렌드 및 TAM-SAM-SOM 시장 규모 추정 표와 상세 분석"),
        ("Ⅲ. Scale-up 및 심층 재무 로드맵", "연도별 스케일업 마일스톤 및 5개년 재무 추정 표와 중장기 로드맵"),
        ("Ⅳ. 최종 사업화 제안 (Lean Canvas 포함)", "SWOT 분석, 린 캔버스 9개 블록 표 및 최종 사업화 전략 제언")
    ]

    doc = Document()
    progress_bar = st.progress(0)
    
    try:
        report_data = {}
        for i, (title, mission) in enumerate(sections_info):
            with st.spinner(f"⏳ {title} 생성 중... (최대 분량 확보 중)"):
                prompt = f"""당신은 최고의 기술사업화 전문가입니다. [특허 명세서]를 기반으로 아래 미션을 수행하세요.
                [특허 명세서]: {tech_text}
                [추가 정보]: {context}
                
                [미션]: {mission}
                [지침]: 
                1. 절대 요약하지 말고 가능한 아주 길고 상세하게(A4 2페이지 이상 분량) 작성하세요.
                2. 데이터는 반드시 표(|---|) 형식을 사용하여 정리하세요.
                3. 소제목은 ## 를 사용하세요."""
                
                # 섹션별 개별 생성 (토큰 문제 해결)
                response = client.models.generate_content(model="models/gemini-2.5-flash-lite", contents=prompt)
                report_data[title] = response.text.strip()
                progress_bar.progress((i + 1) / len(sections_info))

        # 워드 조립
        # 1. 제목
        title_p = doc.add_paragraph("\n\n")
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_p.add_run(f"Virtual Firm 심층 사업화 전략 보고서\n\n[{report_data['TECH_TITLE']}]")
        set_font(title_run, "KoPub돋움체_Pro Bold", 18, bold=True)
        doc.add_page_break()

        # 2. 본문 섹션들
        for title, _ in sections_info[1:]:
            h_p = doc.add_paragraph()
            set_font(h_p.add_run(title), "KoPub돋움체_Pro Bold", 15, bold=True)
            add_smart_content(doc, report_data[title])
            doc.add_page_break()

        doc_io = io.BytesIO()
        doc.save(doc_io)
        st.success("✅ 모든 섹션이 완벽하게 포함된 심층 보고서가 완성되었습니다!")
        st.download_button(label="📥 최종 무결점 보고서 다운로드", data=doc_io.getvalue(), file_name=f"VF_Full_Master_Report.docx")
        
    except Exception as e:
        st.error(f"오류 발생: {str(e)}")
