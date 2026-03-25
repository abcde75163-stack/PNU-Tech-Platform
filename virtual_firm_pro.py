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
    st.error("API 키 로드 실패. .streamlit/secrets.toml 설정을 확인하세요.")

# 1. 폰트 설정 함수
def set_font(run, font_name, size, bold=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold

# 2. 파일 텍스트 추출 함수
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

# 3. 워드 정식 표 생성 함수
def create_word_table(doc, table_text):
    rows_data = [line.strip() for line in table_text.split('\n') if '|' in line]
    grid = []
    for r in rows_data:
        cells = [c.strip() for c in r.split('|') if c.strip()]
        if cells: grid.append(cells)
    
    if not grid: return
    
    # 열 개수 맞추기
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

# 4. 텍스트 및 표 지능형 삽입 함수 (글자 크기 고정 로직 포함)
def add_smart_content(doc, text):
    # 마크다운 샵(#) 제거하여 글자 크기 폭주 방지
    text = re.sub(r'#+', '', text)
    lines = text.split('\n')
    table_buffer = []
    
    for line in lines:
        # 표 데이터 수집
        if '|' in line:
            table_buffer.append(line)
            continue
        
        # 표 버퍼가 있고 일반 텍스트가 나오면 표 생성
        if table_buffer:
            create_word_table(doc, "\n".join(table_buffer))
            table_buffer = []
        
        line_stripped = line.strip()
        if not line_stripped: continue
        
        # 불필요한 마커 제거
        line_stripped = re.sub(r'\[SECTION_\d\]|\[TECH_TITLE\]|\[TABLE_START\]|\[TABLE_END\]', '', line_stripped, flags=re.IGNORECASE).strip()
        if not line_stripped: continue

        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5
        
        # 문장 길이나 별표(*) 여부로 소제목 판단 (글자 크기 제한)
        is_subheading = line_stripped.startswith('*') or (len(line_stripped) < 40 and not line_stripped.endswith('.'))
        clean_text = line_stripped.replace('*', '').strip()
        
        run = p.add_run(clean_text)
        if is_subheading:
            set_font(run, "KoPub돋움체_Pro Medium", 12, bold=True)
            p.paragraph_format.space_before = Pt(10)
        else:
            set_font(run, "KoPub돋움체_Pro Light", 11)

    # 남은 표 처리
    if table_buffer:
        create_word_table(doc, "\n".join(table_buffer))

# 5. 메인 실행 함수 (섹션별 순차 생성 방식)
def run_virtual_firm(spec_file, doc_template, target_corp, ir_data, business_status):
    if not spec_file:
        st.error("특허 명세서 파일이 필요합니다.")
        return

    st.subheader(f"🏢 {target_corp if target_corp else 'Virtual Firm'} 심층 전략 보고서")
    tech_text = extract_text_from_file(spec_file)
    ir_text = extract_text_from_file(ir_data) if ir_data else ""
    context = f"타겟 기업: {target_corp}\nIR 데이터: {ir_text}\n사업 현황: {business_status}"

    # 섹션 정의 및 미션 (중복 차단 가이드 포함)
    sections_info = [
        ("TECH_TITLE", "기술의 핵심 가치를 보여주는 20자 내외의 전문적인 사업화 명칭 한 줄만."),
        ("Ⅰ. 기술 개요 및 메커니즘 분석", "기술의 근본 메커니즘, 작동 원리, 기존 기술 대비 차별성 분석 표와 상세 설명. 인사말 없이 바로 본론 작성."),
        ("Ⅱ. 시장 트렌드 및 TAM-SAM-SOM 분석", "글로벌 산업 트렌드 및 TAM-SAM-SOM 시장 규모 추정 표와 산출 근거 상세 분석. 앞 섹션 내용 중복 금지."),
        ("Ⅲ. Scale-up 및 심층 재무 로드맵", "연도별 매출 추정 표와 투자 유치 전략. 시장 배경은 생략하고 수치와 마일스톤 중심으로 상세 서술."),
        ("Ⅳ. 최종 사업화 제안 (Lean Canvas 포함)", "SWOT 분석 표, 린 캔버스 9개 블록 상세 표, 그리고 최종 전략 제언. 앞 내용 요약하지 말 것.")
    ]

    doc = Document()
    # 문서 전체 여백 설정
    for s in doc.sections:
        s.top_margin = Pt(72); s.bottom_margin = Pt(72); s.left_margin = Pt(72); s.right_margin = Pt(72)

    progress_bar = st.progress(0)
    report_data = {}
    
    try:
        # 각 섹션을 개별적으로 생성 (토큰 소모가 있지만 품질과 분량이 보장됨)
        for i, (title, mission) in enumerate(sections_info):
            with st.spinner(f"⏳ {title} 분석 및 작성 중..."):
                prompt = f"""당신은 최고의 기술사업화 전략가이자 VC 심사역입니다. 
                [특허 명세서]: {tech_text[:6000]} 
                [추가 정보]: {context}
                
                [미션]: {mission}
                
                [지침]:
                1. 절대로 요약하지 말고 A4 2페이지 분량이 나올 만큼 아주 길고 전문적으로 서술하세요.
                2. 모든 핵심 데이터는 반드시 표(|---|) 형식으로 포함하세요.
                3. 다른 섹션의 내용을 반복하지 말고 오직 부여된 미션에만 집중하세요."""
                
                # 모델은 1.5 Pro (권장) 또는 2.5 Flash Lite 중 선택
                response = client.models.generate_content(model="models/gemini-1.5-pro", contents=prompt)
                report_data[title] = response.text.strip()
                progress_bar.progress((i + 1) / len(sections_info))

        # --- 워드 문서 조립 ---
        # 1. 표지/제목
        title_p = doc.add_paragraph("\n\n\n")
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # 제목에서 불필요한 텍스트 청소
        clean_title = re.sub(r'#|\*|제목:', '', report_data['TECH_TITLE']).strip()
        title_run = title_p.add_run(f"Virtual Firm 심층 사업화 전략 보고서\n\n[{clean_title}]")
        set_font(title_run, "KoPub돋움체_Pro Bold", 18, bold=True)
        doc.add_page_break()

        # 2. 본문 섹션 삽입
        for title, _ in sections_info[1:]:
            h_p = doc.add_paragraph()
            set_font(h_p.add_run(title), "KoPub돋움체_Pro Bold", 15, bold=True)
            h_p.paragraph_format.space_after = Pt(12)
            
            add_smart_content(doc, report_data[title])
            doc.add_page_break()

        # 3. 파일 저장 및 다운로드
        doc_io = io.BytesIO()
        doc.save(doc_io)
        st.success("✅ 모든 챕터가 독립적으로 구성된 심층 보고서 작성이 완료되었습니다!")
        st.download_button(label="📥 최종 고도화 보고서 다운로드", data=doc_io.getvalue(), 
                           file_name=f"VF_Final_Strategic_Report.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        
    except Exception as e:
        st.error(f"보고서 생성 중 오류 발생: {str(e)}")
