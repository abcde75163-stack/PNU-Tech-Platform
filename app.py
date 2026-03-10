import streamlit as st

# 페이지 설정 (가장 위에 위치해야 합니다)
st.set_page_config(page_title="PNU 기술지주 자동화 시스템", layout="wide", page_icon="🚀")

# 각 모듈에서 실행 함수 불러오기
try:
    from main1 import run_smk
    from proposal_maker import run_proposal
    from virtual_firm_pro import run_virtual_firm
except ImportError as e:
    st.error(f"⚠️ 라이브러리 또는 모듈 로드 오류: {e}\n\n(main1.py, proposal_maker.py, virtual_firm_pro.py 파일이 같은 폴더에 있는지 확인하세요.)")

st.title("🚀 PNU 기술사업화 통합 플랫폼")
st.markdown("---")

# ==========================================
# 1. 정보 입력 및 파일 업로드 섹션
# ==========================================
st.subheader("📋 정보 입력 및 파일 업로드")

col_text1, col_text2 = st.columns(2)
with col_text1:
    target_corp = st.text_input("🏢 3. 타겟 기업명 (선택)", placeholder="예: (주)부산테크")
with col_text2:
    business_status = st.text_area("📊 5. 주요 사업 현황 (선택)", placeholder="기업의 사업 내용을 입력하면 더 정교한 분석이 가능합니다.", height=68)

st.markdown("---")

col_file1, col_file2, col_file3 = st.columns(3)
with col_file1:
    spec_file = st.file_uploader("📄 1. 특허 명세서 (필수)", type=['pdf', 'docx', 'txt'])
with col_file2:
    # ✅ 수정 포인트: 워드(.docx) 확장자 허용 및 안내 문구 직관적으로 변경
    template_file = st.file_uploader("🎨 2. 템플릿 양식 (선택 - 미업로드 시 기본양식 자동적용)", type=['pptx', 'docx'])
with col_file3:
    ir_data = st.file_uploader("📂 4. IR 자료 (선택)", type=['pdf', 'pptx'])

st.markdown("---")

# ==========================================
# 2. 보고서 실행 섹션
# ==========================================
st.subheader("⚙️ 생성할 보고서 선택")
btn_smk, btn_prop, btn_vf = st.columns(3)

# --- SMK 생성 ---
with btn_smk:
    if st.button("📑 SMK 생성", use_container_width=True):
        if spec_file:
            run_smk(spec_file, template_file, target_corp, ir_data, business_status)
        else:
            st.warning("⚠️ SMK 생성을 위해 '특허 명세서'를 업로드해주세요.")

# --- 전략보고서 생성 ---
with btn_prop:
    if st.button("📊 전략보고서 생성", use_container_width=True):
        if spec_file:
            try:
                run_proposal(spec_file, template_file, target_corp, ir_data, business_status)
            except NameError:
                st.info("🛠️ 현재 proposal_maker 모듈이 준비되지 않았습니다.")
        else:
            st.warning("⚠️ 전략보고서 생성을 위해 '특허 명세서'를 업로드해주세요.")

# --- Virtual Firm 생성 ---
with btn_vf:
    if st.button("🏢 Virtual Firm 생성", use_container_width=True):
        if spec_file:
            try:
                run_virtual_firm(spec_file, template_file, target_corp, ir_data, business_status)
            except NameError:
                st.info("🛠️ 현재 virtual_firm_pro 모듈이 준비되지 않았습니다.")
        else:
            st.warning("⚠️ Virtual Firm 분석을 위해 '특허 명세서'를 업로드해주세요.")

# 하단 정보
st.markdown("---")
st.caption("© 2026 Pusan National University Technology Holdings. 필수 항목 외에는 업로드하지 않아도 범용 모드로 생성이 가능합니다.")