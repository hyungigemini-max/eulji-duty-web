import calendar
import contextlib
import datetime
import importlib.util
import io
import os
import openpyxl
import streamlit as st

# ==========================================
# 1. 페이지 기본 설정 (앱 최상단에서 단 1회만 실행)
# ==========================================
st.set_page_config(
    page_title="3교대 간호사 근무표 자동화 포털 (v15)",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="collapsed",
)


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
  """상업용 SaaS & 대학병원 임상 포털 스타일 커스텀 CSS"""
  st.markdown(
      """
      <style>
      @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
      @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;600;700&display=swap');

      html, body, [class*="css"] {
          font-family: 'Pretendard', 'Noto Sans KR', -apple-system, BlinkMacSystemFont, sans-serif;
      }

      /* 전체 앱 배경: 세련된 쿨그레이-아이스블루 그라데이션 */
      [data-testid="stAppViewContainer"] {
          background: linear-gradient(135deg, #EEF2F6 0%, #E2E8F0 100%);
      }

      /* 상단 기본 헤더 투명화 및 본문 너비/여백 최적화 */
      header[data-testid="stHeader"] {
          background-color: rgba(0, 0, 0, 0);
      }
      #MainMenu {visibility: hidden;}
      footer {visibility: hidden;}

      .block-container {
          padding-top: 2rem;
          padding-bottom: 3rem;
          max-width: 1180px;
      }

      /* 상단 히어로 배너 카드 */
      .hero-banner {
          background: linear-gradient(135deg, #0F2C59 0%, #1B4965 55%, #2B7A78 100%);
          color: #FFFFFF;
          padding: 2.2rem 2.5rem;
          border-radius: 18px;
          box-shadow: 0 12px 28px rgba(15, 44, 89, 0.16);
          margin-bottom: 1.4rem;
      }
      .hero-badge {
          display: inline-block;
          background: rgba(255, 255, 255, 0.18);
          color: #E0F2F1;
          font-size: 0.78rem;
          font-weight: 600;
          padding: 0.28rem 0.85rem;
          border-radius: 999px;
          letter-spacing: 0.04em;
          margin-bottom: 0.75rem;
      }
      .hero-title {
          font-size: 1.85rem;
          font-weight: 700;
          margin: 0 0 0.45rem 0;
          line-height: 1.3;
          color: #FFFFFF;
          letter-spacing: -0.02em;
      }
      .hero-sub {
          font-size: 0.96rem;
          color: #DCEEF2;
          margin: 0;
          line-height: 1.55;
      }

      /* 로그인 헤더 카드 */
      .login-header-card {
          background: linear-gradient(135deg, #0F2C59 0%, #1E517B 100%);
          color: white;
          padding: 2.4rem 2rem;
          border-radius: 18px;
          text-align: center;
          box-shadow: 0 14px 32px rgba(15, 44, 89, 0.18);
          margin-top: 1.5rem;
          margin-bottom: 1.2rem;
      }
      .login-header-card h2 {
          margin: 0.5rem 0 0.4rem 0;
          font-size: 1.6rem;
          font-weight: 700;
          color: #FFFFFF;
      }
      .login-header-card p {
          margin: 0;
          font-size: 0.9rem;
          color: #CFE4F5;
      }

      /* st.container(border=True) 및 폼을 상업용 화이트 입체 카드로 변환 */
      [data-testid="stVerticalBlockBorderWrapper"], div[data-testid="stForm"] {
          background-color: #FFFFFF;
          border-radius: 16px;
          border: 1px solid #E2E8F0;
          box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
          padding: 1.6rem 1.8rem;
      }

      /* 버튼 디자인 개선 */
      .stButton > button[kind="primary"],
      div[data-testid="stFormSubmitButton"] > button,
      .stDownloadButton > button[kind="primary"] {
          background
