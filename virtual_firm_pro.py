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
    for r_idx, row in enumerate(grid):
        for c_idx, val in enumerate(row):
            if c_idx < max_cols:
                cell = table.cell(r_idx, c_idx)
                cell.text = val
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        set_font(r, "KoPub돋움체_Pro Light", 10, bold=(r_idx == 0))

def add_smart_content(doc, text):
    # 💡 글자 크기 폭주 방지: 모든 마크다운 기호(#, *) 제거 및 정리
    text = re.sub(r'#+', '', text)
    lines = text.split('\n')
    table_buffer = []
    
    for line in lines:
        if '|' in line:
            table_buffer.append(line)
            continue
        
        if table_buffer:
            create_word_table(doc, "\n".join(table_buffer))
            table_buffer = []
        
        line_stripped = line.strip()
        if not line_stripped: continue
        
        # 💡 내용 중복 패턴(서론, 요약 등) 강제 필터링
        if any(x in line_stripped for x in ["본 보고서는", "요약하자면", "결론적으로", "서론", "개요"]):
            if len(line_stripped) < 60: continue

        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5
        
        # 💡 강제 스타일 제어: 본문 11pt, 소제목 12pt 고정
        is_subheading = (len(line_stripped) < 50 and not line_stripped.endswith('.'))
        clean_text = line_stripped.replace('*', '').strip()
        
        run = p.add_run(clean_text)
        if is_subheading:
            set_font(run, "KoPub돋움체_Pro Medium", 12, bold=True)
            p.paragraph_format.space_before = Pt(12)
        else:
            set_font(run, "KoPub돋움체_Pro Light", 11)

    if table_buffer:
        create_word_table(doc, "\n".join(table_buffer))

def run_virtual_firm(spec_file, doc_template, target_corp, ir_data, business_status):
    if not spec_file: return
    
    tech_text = extract_text_from_file(spec_file)
    context = f"타겟 기업: {target_corp}\n사업 현황: {business_status}"
    
    # 💡 요청하신 gemini-2.5-flash 모델로 고정
    MODEL_NAME = "gemini-2.5-flash" 

    sections = [
        ("TECH_TITLE", "기술의 핵심 가치를 담은 전문적인 사업화 명칭 한 줄만."),
        ("Ⅰ. 기술 개요 및 메커니즘 분석", "기술의 작동 원리와 차별성 분석 표 포함. 다른 챕터 언급 없이 기술 본론만 상세히."),
        ("Ⅱ. 시장 트렌드 및 TAM-SAM-SOM 분석", "글로벌 시장 트렌드와 TAM-SAM-SOM 추정 수치 표. Ⅰ번 내용 반복 금지, 오직 시장 분석에만 집중."),
        ("Ⅲ. Scale-up 및 심층 재무 로드맵", "연도별 매출 추정 및 투자 유치 계획 표. 서론 생략하고 수치와 단계별 마일스톤 중심으로 상세 서술."),
        ("Ⅳ. 최종 사업화 제안 (Lean Canvas 포함)", "SWOT 분석 표와 린캔버스 9개 항목 상세 표. 앞선 내용 요약하지 말고 최종 전략 제언만 상세히.")
    ]

    doc = Document()
    for s in doc.sections:
        s.top_margin = Pt(72); s.bottom_margin = Pt(72); s.left_margin = Pt(72); s.right_margin = Pt(72)

    progress_bar = st.progress(0)
    results = {}
    
    try:
        for i, (title, mission) in enumerate(sections):
            with st.spinner(f"⏳ {title} 생성 중 (gemini-2.5-flash)..."):
                # 💡 중복 방지를 위한 강력한 네거티브 프롬프트
                prompt = f"전문가로서 [특허:{tech_text[:5000]}] 기반 [미션:{mission}]을 수행하세요.\n" \
                         f"지침: 1. 절대로 이전 섹션의 내용을 반복하거나 서론을 쓰지 마세요. 2. 인사말 없이 바로 표와 상세 분석으로 들어가세요. 3. 가능한 아주 길게 작성하세요."
                
                for _ in range(3): # 재시도 로직
                    try:
                        resp = client.models.generate_content(model=MODEL_NAME, contents=prompt)
                        results[title] = resp.text.strip()
                        break
                    except Exception:
                        time.sleep(3)
                
                progress_bar.progress((i + 1) / len(sections))

        # 문서 조립
        title_p = doc.add_paragraph("\n\n\n")
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        clean_title = re.sub(r'#|\*|제목:|명칭:', '', results.get('TECH_TITLE', '심층 보고서')).strip()
        title_run = title_p.add_run(f"Virtual Firm 심층 사업화 전략 보고서\n\n[{clean_title}]")
        set_font(title_run, "KoPub돋움체_Pro Bold", 18, bold=True)
        doc.add_page_break()

        for title, _ in sections[1:]:
            h = doc.add_paragraph()
            set_font(h.add_run(title), "KoPub돋움체_Pro Bold", 15, bold=True)
            add_smart_content(doc, results.get(title, "내용 생성 실패"))
            doc.add_page_break()

        doc_io = io.BytesIO()
        doc.save(doc_io)
        st.success("✅ gemini-2.5-flash 모델 적용 및 최적화 완료!")
        st.download_button(label="📥 최종 보고서 다운로드", data=doc_io.getvalue(), file_name=f"VF_2.5_Flash_Report.docx")
        
    except Exception as e:
        st.error(f"오류 발생: {str(e)}")
