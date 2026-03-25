import streamlit as st
import fitz
from google import genai
import io
import re
import time
import matplotlib.pyplot as plt  # 📊 차트 그리기용 라이브러리 추가
from docx import Document
from docx.shared import Pt, Inches  # 이미지 크기 조절을 위해 Inches 추가
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

# 📊 [신규 기능] 파이썬 네이티브 차트 생성 함수
def create_tam_sam_som_chart(tam, sam, som, unit):
    fig, ax = plt.subplots(figsize=(7, 4.5))
    labels = ['TAM\n(Total Addressable Market)', 'SAM\n(Serviceable Available Market)', 'SOM\n(Serviceable Obtainable Market)']
    values = [tam, sam, som]
    colors = ['#4F81BD', '#C0504D', '#9BBB59']

    bars = ax.bar(labels, values, color=colors, width=0.5)

    # 막대그래프 위에 숫자 달아주기
    for bar in bars:
        yval = bar.get_height()
        # 숫자에 콤마 찍어서 표시 (예: 1,000)
        ax.text(bar.get_x() + bar.get_width()/2, yval + (max(values)*0.02), f'{yval:,.1f}', ha='center', va='bottom', fontweight='bold')

    ax.set_ylabel(f'Market Size ({unit})', fontweight='bold')
    ax.set_title('Market Size Estimation (TAM - SAM - SOM)', fontweight='bold', pad=20)
    
    # 테두리 깔끔하게 지우기
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    img_stream = io.BytesIO()
    plt.tight_layout()
    plt.savefig(img_stream, format='png', dpi=300)
    plt.close(fig)
    img_stream.seek(0)
    return img_stream

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

    with st.spinner("🚀 시장 규모 데이터를 추정하여 텍스트 분석 및 차트를 자동 생성 중입니다... (약 20~30초 소요)"):
        try:
            ir_text = extract_text_from_file(ir_data) if ir_data else ""
            context = f"타겟 기업: {target_corp}\nIR 데이터: {ir_text}\n사업 현황: {business_status}"
            
            # 💡 [프롬프트] 차트를 그리기 위한 비밀 데이터 마커 지시 추가
            prompt = f"""당신은 국내 최고의 기술사업화 전략가이자 벤처캐피탈(VC) 수석 심사역입니다.
            제공된 [특허 명세서]를 바탕으로 실제 대규모 투자 유치에 사용될 '초정밀 심층 비즈니스 플랜'을 작성해야 합니다.

            [특허 명세서 원본 데이터]
            {tech_text}

            [작성 지침 - 분량 및 깊이 절대 엄수]
            1. 절대로 요약하거나 짧게 끝내지 마세요. 하위에 상세한 개조식(-, *) 설명과 분석을 덧붙여 각 섹션이 방대한 분량이 되도록 길고 깊게 작성하세요.
            2. 마크다운 표(|---|)는 에러를 유발하므로 절대 사용하지 마세요. 모든 수치, 표, 분석 결과는 '글로 길게 풀어서' 상세히 나열하세요.
            3. 응답은 반드시 아래의 5개 [구분자]를 사용하여 나누어야 합니다.

            [TECH_TITLE]
            (기술의 핵심 가치를 보여주는 20자 내외의 임팩트 있는 비즈니스 명칭 한 줄만 작성)
            
            [SECTION_1]
            (기술의 근본적 메커니즘, 화학적/물리적 작동 원리, 기존 기술의 한계점과 본 기술의 혁신적 돌파구를 상세히 작성)
            
            [SECTION_2]
            (글로벌 전방 산업의 메가 트렌드 및 기술 동향을 심층 분석하세요. 시장 규모를 'TAM-SAM-SOM' 프레임워크를 적용하여 글로 아주 상세히 서술하세요. 
            ★ 중요: 파이썬이 차트를 그릴 수 있도록, [SECTION_2]의 맨 마지막 줄에 반드시 아래 형식의 데이터를 추가하세요. 숫자는 쉼표 없이 적어주세요.
            [TAM_VALUE: 숫자]
            [SAM_VALUE: 숫자]
            [SOM_VALUE: 숫자]
            [UNIT: 단위(예: 억 원, 백만 달러 등)])
            
            [SECTION_3]
            (단/중/장기 스케일업 로드맵을 연도별로 아주 길게 서술하세요. '수익접근법' 기술가치평가 가설 수치와 산출 근거를 서술하세요.)
            
            [SECTION_4]
            (내부 역량(3C), 외부 환경(SWOT) 분석, '린 캔버스(Lean Canvas)'의 9가지 핵심 블록을 소제목과 텍스트로 아주 상세히 서술하세요. 최종적으로 '기술이전'과 '직접 창업' 중 하나를 명확히 선택하고 그 근거를 서술하세요.)

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
                    # 💡 [핵심] 섹션 2번일 경우 차트 데이터를 빼내고 그림 삽입
                    if key == "section_2":
                        tam_val = re.search(r'\[TAM_VALUE:\s*([\d\.]+)\]', content, re.IGNORECASE)
                        sam_val = re.search(r'\[SAM_VALUE:\s*([\d\.]+)\]', content, re.IGNORECASE)
                        som_val = re.search(r'\[SOM_VALUE:\s*([\d\.]+)\]', content, re.IGNORECASE)
                        unit_val = re.search(r'\[UNIT:\s*(.*?)\]', content, re.IGNORECASE)
                        
                        # 비밀 마커 지우기 (보고서에 노출 방지)
                        clean_content = re.sub(r'\[(?:TAM|SAM|SOM)_VALUE:.*?\]', '', content, flags=re.IGNORECASE)
                        clean_content = re.sub(r'\[UNIT:.*?\]', '', clean_content, flags=re.IGNORECASE)
                        add_styled_content(doc, clean_content.strip())
                        
                        # 값이 모두 추출되었다면 차트를 그려서 삽입
                        if tam_val and sam_val and som_val:
                            try:
                                t = float(tam_val.group(1))
                                s = float(sam_val.group(1))
                                m = float(som_val.group(1))
                                u = unit_val.group(1).strip() if unit_val else "단위"
                                
                                # 차트 생성 및 워드 삽입
                                chart_stream = create_tam_sam_som_chart(t, s, m, u)
                                pic_p = doc.add_paragraph()
                                pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                pic_run = pic_p.add_run()
                                pic_run.add_picture(chart_stream, width=Inches(5.5))
                            except Exception as e:
                                pass # 파싱 실패 시 차트만 생략 (에러 방지)
                    else:
                        add_styled_content(doc, content)
                else:
                    p = doc.add_paragraph()
                    set_font(p.add_run("생성된 내용이 없습니다."), "KoPub돋움체_Pro Light", 11)
                
                if i < len(sections_list) - 1:
                    doc.add_page_break()

            doc_io = io.BytesIO()
            doc.save(doc_io)
            
            st.success("✅ 시장 규모 차트(그래프)가 포함된 심층 보고서 작성이 완료되었습니다!")
            st.download_button(label="📥 인포그래픽 심층 보고서 다운로드 (클릭)", data=doc_io.getvalue(), 
                               file_name=f"VF_Master_Report_{target_corp}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"보고서 생성 중 오류 발생: {str(e)}")




