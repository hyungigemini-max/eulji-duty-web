import calendar
import datetime
import io
import os
import openpyxl
import streamlit as st

# 기존 V15 엔진 모듈 임포트 (같은 폴더에 위치한 1_create_template_final.py, 2_make_duty_final.py 활용)
import importlib.util


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


def check_password():
  """원내 부서 전용 접근 비밀번호 잠금 기능"""
  # Streamlit Cloud의 Secrets에 APP_PASSWORD가 설정되어 있으면 그 값을 우선 사용하고,
  # 없으면 아래 기본 비밀번호("micu2026")를 사용합니다.
  expected_pw = "micu2026"
  try:
    if "APP_PASSWORD" in st.secrets:
      expected_pw = str(st.secrets["APP_PASSWORD"])
  except Exception:
    pass

  if st.session_state.get("authenticated", False):
    return True

  st.set_page_config(
      page_title="3교대 간호사 근무표 자동 생성기 (v15)",
      page_icon="🏥",
      layout="centered",
  )
  st.title("🏥 3교대 간호사 근무표 자동 생성기")
  st.caption("노원을지대학교병원 내과중환자실(MICU) 개발 · 원내 부서 공용 웹 시스템 (v15)")
  st.info("🔒 원내 부서 전용 프로그램입니다. 공유받으신 접속 비밀번호를 입력해 주세요.")

  with st.form("login_form"):
    pw_input = st.text_input("접속 비밀번호", type="password", placeholder="비밀번호 입력")
    submitted = st.form_submit_button("접속하기", use_container_width=True)
    if submitted:
      if pw_input == expected_pw:
        st.session_state["authenticated"] = True
        st.rerun()
      else:
        st.error("❌ 비밀번호가 일치하지 않습니다. 담당 파트장에게 문의해 주세요.")
  return False


def generate_template_bytes(year, month, prev_result_bytes=None, prev_input_bytes=None):
  """웹에서 새 양식(duty_input.xlsx)을 메모리 상에서 즉시 생성해 바이트로 반환합니다."""
  work_dir = os.path.join(BASE_DIR, "_web_tmp")
  os.makedirs(work_dir, exist_ok=True)
  orig_cwd = os.getcwd()
  try:
    os.chdir(work_dir)
    # 임시 폴더 초기화
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

    # input() 프롬프트를 대체하여 선택한 연·월로 자동 실행
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
  """업로드된 duty_input.xlsx 바이트를 V15 엔진으로 계산해 결과 엑셀 바이트와 콘솔 로그를 반환합니다."""
  import contextlib

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

    # 입력 파일에서 연·월 확인
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
      page_title="3교대 간호사 근무표 자동 생성기 (v15)",
      page_icon="🏥",
      layout="wide",
  )

  st.title("🏥 3교대 간호사 근무표 자동 생성기 (v15)")
  st.markdown(
      "별도 프로그램 설치 없이 **엑셀 양식 생성 → 원티드·규칙 입력 → 자동 근무표"
      " 생성**까지 웹에서 바로 실행할 수 있습니다."
  )
  st.warning(
      "⚠️ **개인정보 보호 안내:** 업로드하신 파일은 계산 직후 서버에서 즉시 삭제되며"
      " 저장되지 않습니다. 원내 보안 규정을 위해 엑셀에는 **성명(또는 사번·이니셜)**만"
      " 입력하시고, 주민등록번호·연락처 등 민감 개인정보는 절대 포함하지 마세요."
  )

  tab1, tab2, tab3 = st.tabs([
      "1️⃣ 근무표 자동 생성하기",
      "2️⃣ 다음 달 입력 양식(duty_input.xlsx) 만들기",
      "📖 사용 가이드",
  ])

  # -----------------------------------------------------------------------
  # [탭 1] 근무표 자동 생성
  # -----------------------------------------------------------------------
  with tab1:
    st.subheader("작성 완료된 `duty_input.xlsx` 업로드")
    uploaded_input = st.file_uploader(
        "원티드 및 설정 입력이 끝난 duty_input.xlsx 파일을 선택하세요.",
        type=["xlsx"],
        key="input_uploader",
    )

    if uploaded_input is not None:
      if st.button("🚀 최적 근무표 자동 생성 시작 (최대 50초 소요)", type="primary"):
        with st.spinner("OR-Tools 최적화 엔진이 최적의 근무표를 계산하고 있습니다..."):
          res_bytes, out_name, logs = run_solver_bytes(uploaded_input.getvalue())

        with st.expander("📋 실행 진단 및 알림 로그 보기", expanded=True):
          st.code(logs, language="text")

        if res_bytes is not None:
          st.success(f"✅ `{out_name}` 생성이 완료되었습니다! 아래 버튼을 눌러 내려받으세요.")
          st.download_button(
              label=f"📥 결과 근무표 다운로드 ({out_name})",
              data=res_bytes,
              file_name=out_name,
              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              type="primary",
          )
        else:
          st.error("❌ 근무표 생성에 실패했습니다. 위 로그의 안내 메시지를 확인해 주세요.")

  # -----------------------------------------------------------------------
  # [탭 2] 새 양식 만들기
  # -----------------------------------------------------------------------
  with tab2:
    st.subheader("새로운 달의 `duty_input.xlsx` 양식 생성")
    today = datetime.date.today()
    def_y = today.year + 1 if today.month == 12 else today.year
    def_m = 1 if today.month == 12 else today.month + 1

    col1, col2 = st.columns(2)
    with col1:
      sel_year = st.number_input("연도", min_value=2025, max_value=2035, value=def_y, step=1)
    with col2:
      sel_month = st.number_input("월", min_value=1, max_value=12, value=def_m, step=1)

    st.markdown("**(선택) 직전 달 파일 첨부 시 명단·설정·전월 말 3일 근무가 자동 이월됩니다:**")
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

    if st.button("📄 새 입력 양식(duty_input.xlsx) 생성"):
      tpl_bytes = generate_template_bytes(
          int(sel_year),
          int(sel_month),
          prev_result_bytes=prev_res_file.getvalue() if prev_res_file else None,
          prev_input_bytes=prev_inp_file.getvalue() if prev_inp_file else None,
      )
      st.success(f"✅ {int(sel_year)}년 {int(sel_month)}월 입력 양식이 생성되었습니다!")
      st.download_button(
          label="📥 duty_input.xlsx 다운로드",
          data=tpl_bytes,
          file_name="duty_input.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          type="primary",
      )

  # -----------------------------------------------------------------------
  # [탭 3] 사용 가이드
  # -----------------------------------------------------------------------
  with tab3:
    st.markdown(
        """
### 📌 부서별 맞춤 설정 방법 (`[설정]` 시트)
타 부서에서 사용하실 때는 `duty_input.xlsx`의 **`[설정]` 시트 B7~B13** 값만 부서 기준에 맞게 수정하시면 됩니다.
- **일일 나이트(N) 고정 인원:** 매일 고정 배치할 나이트 인원 (예: 일반 병동 `3` 또는 `4`, 중환자실 `5`)
- **일반일 D·E 최소 필요 인원:** 평일·주말 최소 인원 (입력값 ~ `입력값+2명` 범위에서 자동 배정)
- **공휴일(수당일) D·E 고정 인원:** 공휴일 근무 고정 인원
- **바쁜 요일 (우선 배치):** 인원이 여유로울 때 먼저 채울 요일 (예: `일, 월, 화, 수` 또는 `수, 목, 금`)
- **최대 연속 근무일수 / 나이트 간 최소 휴식일:** 부서 내규에 맞게 숫자 조정

### 📌 개인별 근무선호/비율 입력 (`[근무표_입력]` E열)
- **비율 지정:** `6:4`, `7:3`, `8:2` 등 (앞 숫자=D 비중, 뒤 숫자=E 비중)
- **특정 듀티 위주:** `DN`(데이·나이트 위주), `EN`(이브닝·나이트 위주), `DE`(데이·이브닝 위주)
- 비워두면 일반 3교대(D/E 1:1 균등)로 자동 배정됩니다.
"""
    )

  st.divider()
  st.caption("Developed by 노원을지대학교병원 내과중환자실(MICU) 파트장 · Core Engine v15")


if __name__ == "__main__":
  main()
