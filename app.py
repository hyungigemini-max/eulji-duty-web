import calendar
import contextlib
import datetime
import importlib.util
import io
import os
import openpyxl
import streamlit as st


def load_module_from_path(module_name, file_path):
  spec = importlib.util.spec_from_file_location(module_name, file_path)
  mod = importlib.util.module_from_spec(spec)
  spec.loader.exec_module(mod)
  return mod


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
template_mod = load_module_from_path(
    "template_v15", os.path.join(BASE_DIR, "1_create_template_final.py")
)
solver_mod = load_module_from_path(
    "solver_v15", os.path.join(BASE_DIR, "2_make_duty_final.py")
)


def inject_custom_css():
  """대학병원 임상 포털 스타일의 커스텀 CSS 및 폰트 적용"""
  st.markdown(
      """
      <style>
      @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;600;700&display=swap');

      html, body, [class*="css"] {
          font-family: 'Noto Sans KR', -apple-system, BlinkMacSystemFont, sans-serif;
      }

      /* 전체 배경 톤 정돈 */
      .stApp {
          background-color: #F4F7FA;
      }

      /* 상단 불필요한 기본 데코레이션 및 푸터 숨김 */
      header[data-testid="stHeader"] {
          background: transparent;
      }
      #MainMenu {visibility: hidden;}
      footer {visibility: hidden;}

      /* 상단 히어로 배너 카드 */
      .hero-banner {
          background: linear-gradient(135deg, #0F2C59 0%, #1B4965 55%, #2B7A78 100%);
          color: #FFFFFF;
          padding: 2rem 2.2rem;
          border-radius: 16px;
          box-shadow: 0 10px 25px rgba(15, 44, 89, 0.15);
          margin-bottom: 1.5rem;
      }
      .hero-badge {
          display: inline-block;
          background: rgba(255, 255, 255, 0.18);
          color: #E0F2F1;
          font-size: 0.78rem;
          font-weight: 600;
          padding: 0.25rem 0.75rem;
          border-radius: 999px;
          letter-spacing: 0.03em;
          margin-bottom: 0.7rem;
      }
      .hero-title {
          font-size: 1.75rem;
          font-weight: 700;
          margin: 0 0 0.4rem 0;
          line-height: 1.3;
          color: #FFFFFF;
      }
      .hero-sub {
          font-size: 0.95rem;
          color: #DCEEF2;
          margin: 0;
          line-height: 1.5;
      }

      /* 로그인 카드 스타일 */
      .login-header-card {
          background: linear-gradient(135deg, #0F2C59 0%, #1E517B 100%);
          color: white;
          padding: 2.2rem 2rem;
          border-radius: 18px;
          text-align: center;
          box-shadow: 0 12px 30px rgba(15, 44, 89, 0.18);
          margin-top: 1.5rem;
          margin-bottom: 1.2rem;
      }
      .login-header-card h2 {
          margin: 0.5rem 0 0.4rem 0;
          font-size: 1.55rem;
          font-weight: 700;
          color: #FFFFFF;
      }
      .login-header-card p {
          margin: 0;
          font-size: 0.9rem;
          color: #CFE4F5;
      }

      /* 폼(로그인 박스) 화이트 카드화 */
      div[data-testid="stForm"] {
          background-color: #FFFFFF;
          padding: 1.8rem;
          border-radius: 16px;
          border: 1px solid #E2E8F0;
          box-shadow: 0 6px 20px rgba(15, 23, 42, 0.05);
      }

      /* 버튼 디자인 개선 */
      .stButton > button[kind="primary"], div[data-testid="stFormSubmitButton"] > button {
          background: linear-gradient(90deg, #0F2C59 0%, #1D5B79 100%);
          color: #FFFFFF;
          border: none;
          border-radius: 10px;
          padding: 0.65rem 1.2rem;
          font-weight: 600;
          font-size: 0.98rem;
          transition: all 0.2s ease;
          box-shadow: 0 4px 12px rgba(15, 44, 89, 0.2);
      }
      .stButton > button[kind="primary"]:hover, div[data-testid="stFormSubmitButton"] > button:hover {
          transform: translateY(-1px);
          box-shadow: 0 6px 16px rgba(15, 44, 89, 0.28);
      }

      /* 탭 메뉴 세련되게 정돈 */
      button[data-baseweb="tab"] {
          font-weight: 600;
          font-size: 0.98rem;
          padding: 0.7rem 1.2rem;
      }

      /* 보안 안내 배너 */
      .security-notice {
          background-color: #FFFBEB;
          border-left: 4px solid #D97706;
          color: #78350F;
          padding: 0.9rem 1.1rem;
          border-radius: 8px;
          font-size: 0.88rem;
          margin-bottom: 1.2rem;
          line-height: 1.5;
      }

      /* 하단 크레딧 푸터 */
      .custom-footer {
          text-align: center;
          color: #64748B;
          font-size: 0.83rem;
          margin-top: 2.5rem;
          padding-top: 1rem;
          border-top: 1px solid #E2E8F0;
      }
      </style>
      """,
      unsafe_allow_html=True,
  )


def check_password():
  """원내 부서 전용 접근 비밀번호 잠금 화면"""
  expected_pw = "micu2026"
  try:
    if "APP_PASSWORD" in st.secrets:
      expected_pw = str(st.secrets["APP_PASSWORD"])
  except Exception:
    pass

  if st.session_state.get("authenticated", False):
    return True

  st.set_page_config(
      page_title="3교대 간호사 근무표 자동화 포털",
      page_icon="🏥",
      layout="centered",
  )
  inject_custom_css()

  st.markdown(
      """
      <div class="login-header-card">
          <div style="font-size: 2.4rem;">🏥</div>
          <div class="hero-badge">NOWON EULJI MEDICAL CENTER · NURSING SCHEDULER</div>
          <h2>3교대 간호사 근무표 자동화 포털</h2>
          <p>노원을지대학교병원 내과중환자실(MICU) 개발 · 원내 부서 공용 시스템 (v15)</p>
      </div>
      """,
      unsafe_allow_html=True,
  )

  with st.form("login_form"):
    st.markdown("#### 🔒 부서 전용 인증")
    st.caption("원내 간호부서 관계자만 이용할 수 있습니다. 공유받은 접속 코드를 입력하세요.")
    pw_input = st.text_input(
        "접속 비밀번호",
        type="password",
        placeholder="비밀번호를 입력하세요",
        label_visibility="collapsed",
    )
    submitted = st.form_submit_button("시스템 접속하기", use_container_width=True)
    if submitted:
      if pw_input == expected_pw:
        st.session_state["authenticated"] = True
        st.rerun()
      else:
        st.error("비밀번호가 올바르지 않습니다. 담당 파트장에게 확인해 주세요.")

  st.markdown(
      '<div class="custom-footer">Developed by 노원을지대학교병원 내과중환자실(MICU) 파트장 · Powered by OR-Tools Optimization Engine</div>',
      unsafe_allow_html=True,
  )
  return False


def generate_template_bytes(year, month, prev_result_bytes=None, prev_input_bytes=None):
  work_dir = os.path.join(BASE_DIR, "_web_tmp")
  os.makedirs(work_dir, exist_ok=True)
  orig_cwd = os.getcwd()
  try:
    os.chdir(work_dir)
    for fn in os.listdir(work_dir):
      try:
        os.remove(os.path.join(work_dir, fn))
      except Exception:
        pass

    prev_y, prev_m = template_mod.get_prev_year_month(year, month)
    if prev_result_bytes is not None:
      prev_fn = f"duty_result_{prev_y}_{prev_m:02d}_수정.xlsx"
      with open(prev_fn, "wb") as f:
        f.write(prev_result_bytes)

    if prev_input_bytes is not None:
      with open("duty_input.xlsx", "wb") as f:
        f.write(prev_input_bytes)

    orig_ask = template_mod.ask_year_month
    template_mod.ask_year_month = lambda: (year, month)
    try:
      template_mod.create_template_final()
    finally:
      template_mod.ask_year_month = orig_ask

    with open("duty_input.xlsx", "rb") as f:
      return f.read()
  finally:
    os.chdir(orig_cwd)


def run_solver_bytes(input_xlsx_bytes):
  work_dir = os.path.join(BASE_DIR, "_web_tmp")
  os.makedirs(work_dir, exist_ok=True)
  orig_cwd = os.getcwd()
  log_buf = io.StringIO()

  try:
    os.chdir(work_dir)
    for fn in os.listdir(work_dir):
      try:
        os.remove(os.path.join(work_dir, fn))
      except Exception:
        pass

    with open("duty_input.xlsx", "wb") as f:
      f.write(input_xlsx_bytes)

    wb_chk = openpyxl.load_workbook("duty_input.xlsx", data_only=True)
    ws_rule = wb_chk["설정"]
    year = int(ws_rule["B1"].value)
    month = int(ws_rule["B2"].value)
    expected_out = f"duty_result_{year}_{month:02d}.xlsx"

    with contextlib.redirect_stdout(log_buf):
      solver_mod.solve_duty_final()

    log_text = log_buf.getvalue()
    if os.path.exists(expected_out):
      with open(expected_out, "rb") as f:
        return f.read(), expected_out, log_text
    return None, None, log_text
  finally:
    os.chdir(orig_cwd)


def main():
  if not check_password():
    return

  st.set_page_config(
      page_title="3교대 간호사 근무표 자동화 포털 (v15)",
      page_icon="🏥",
      layout="wide",
  )
  inject_custom_css()

  st.markdown(
      """
      <div class="hero-banner">
          <div class="hero-badge">SMART NURSE SCHEDULING SYSTEM · CORE ENGINE v15</div>
          <div class="hero-title">🏥 3교대 간호사 근무표 자동화 포털</div>
          <p class="hero-sub">
              수리 최적화 알고리즘(Google OR-Tools) 기반으로 병동 내규, 공휴일 수당, 개인별 근무선호 및 원티드를 자동 반영합니다.
          </p>
      </div>
      <div class="security-notice">
          <strong>🛡️ 개인정보 및 데이터 보안 안내:</strong> 업로드된 엑셀 파일은 근무표 연산 직후 서버 메모리에서 즉시 파기되며 별도로 저장되지 않습니다.
          원내 정보보안 지침에 따라 엑셀에는 <strong>성명(또는 사번·이니셜)</strong>만 기입해 주세요.
      </div>
      """,
      unsafe_allow_html=True,
  )

  tab1, tab2, tab3 = st.tabs([
      "🚀 근무표 자동 생성",
      "📄 다음 달 입력 양식 만들기",
      "📘 부서별 설정 및 사용 가이드",
  ])

  with tab1:
    with st.container(border=True):
      st.subheader("Step 1. 작성 완료된 `duty_input.xlsx` 업로드")
      st.caption("원티드(O, 연차, 교육)와 [설정] 시트 입력이 완료된 엑셀 파일을 선택하세요.")
      uploaded_input = st.file_uploader(
          "duty_input.xlsx 파일 선택",
          type=["xlsx"],
          key="input_uploader",
          label_visibility="collapsed",
      )

      if uploaded_input is not None:
        st.markdown("---")
        if st.button("✨ 최적 근무표 자동 생성 실행 (최대 50초 소요)", type="primary"):
          with st.spinner("병동 규칙과 원티드를 분석하여 최적의 근무표를 산출하고 있습니다..."):
            res_bytes, out_name, logs = run_solver_bytes(uploaded_input.getvalue())

          if res_bytes is not None:
            st.success(f"🎉 `{out_name}` 근무표 생성이 완료되었습니다!")
            st.download_button(
                label=f"📥 완성된 근무표 엑셀 다운로드 ({out_name})",
                data=res_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
          else:
            st.error("근무표 생성에 실패했습니다. 아래 실행 로그에서 충돌 원인을 확인해 주세요.")

          with st.expander("📋 상세 실행 진단 및 알림 로그", expanded=True):
            st.code(logs, language="text")

  with tab2:
    with st.container(border=True):
      st.subheader("새로운 달의 `duty_input.xlsx` 양식 생성")
      today = datetime.date.today()
      def_y = today.year + 1 if today.month == 12 else today.year
      def_m = 1 if today.month == 12 else today.month + 1

      col1, col2 = st.columns(2)
      with col1:
        sel_year = st.number_input("대상 연도", min_value=2025, max_value=2035, value=def_y, step=1)
      with col2:
        sel_month = st.number_input("대상 월", min_value=1, max_value=12, value=def_m, step=1)

      st.markdown("##### 🔄 전월 데이터 자동 이월 (선택 사항)")
      st.caption("직전 달 파일을 함께 올리면 간호사 명단, 운영 설정, 전월 말 3일 근무가 자동으로 채워집니다.")
      c_up1, c_up2 = st.columns(2)
      with c_up1:
        prev_res_file = st.file_uploader(
            "직전 달 확정 결과 파일 (duty_result_YYYY_MM.xlsx)",
            type=["xlsx"],
            key="prev_res",
        )
      with c_up2:
        prev_inp_file = st.file_uploader(
            "기존 입력 양식 파일 (duty_input.xlsx - 근무선호/설정 유지용)",
            type=["xlsx"],
            key="prev_inp",
        )

      if st.button("📄 새 입력 양식(duty_input.xlsx) 생성하기", type="primary"):
        tpl_bytes = generate_template_bytes(
            int(sel_year),
            int(sel_month),
            prev_result_bytes=prev_res_file.getvalue() if prev_res_file else None,
            prev_input_bytes=prev_inp_file.getvalue() if prev_inp_file else None,
        )
        st.success(f"✅ {int(sel_year)}년 {int(sel_month)}월 입력 양식이 준비되었습니다.")
        st.download_button(
            label="📥 duty_input.xlsx 내려받기",
            data=tpl_bytes,
            file_name="duty_input.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

  with tab3:
    with st.container(border=True):
      st.markdown(
          """
### 🏥 타 부서 맞춤 운영 설정 (`[설정]` 시트 B7~B13)
타 병동·부서에서 사용하실 때는 엑셀의 **`[설정]` 시트** 숫자만 부서 기준에 맞게 수정하시면 됩니다.
- **일일 나이트(N) 고정 인원:** 매일 고정 배치할 나이트 인원 (예: 일반 병동 `3` 또는 `4`, 중환자실 `5`)
- **일반일 D·E 최소 필요 인원:** 평일·주말 최소 인원 (입력값 ~ `입력값+2명` 범위에서 자동 배정)
- **공휴일(수당일) D·E 고정 인원:** 공휴일 근무 고정 인원
- **바쁜 요일 (우선 배치):** 인원이 여유로울 때 먼저 채울 요일 (예: `일, 월, 화, 수` 또는 `수, 목, 금`)
- **최대 연속 근무일수 / 나이트 간 최소 휴식일:** 부서 내규에 맞게 숫자 조정

### 🎯 개인별 근무선호 및 비율 입력 (`[근무표_입력]` E열)
- **비율 지정:** `6:4`, `7:3`, `8:2` 등 (앞 숫자=D 비중, 뒤 숫자=E 비중)
- **특정 듀티 위주:** `DN`(데이·나이트 위주), `EN`(이브닝·나이트 위주), `DE`(데이·이브닝 위주)
- 비워두면 일반 3교대(D/E 1:1 균등)로 자동 배정됩니다.
"""
      )

  st.markdown(
      '<div class="custom-footer">Developed by 노원을지대학교병원 내과중환자실(MICU) 파트장 · Core Engine v15</div>',
      unsafe_allow_html=True,
  )


if __name__ == "__main__":
  main()
