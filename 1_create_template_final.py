import calendar
import datetime
import glob
import os
import re
import shutil
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# =========================================================================
# [법정 공휴일 자동 달력 (2026~2028년)]
# - 일요일 공휴일은 '일요일 우선(기본오프)'으로 자동 제외됩니다.
# - 토요일 공휴일은 '수당일'로 분류되고 필수기본오프에서 자동 차감됩니다.
# - 임시공휴일 추가나 변경은 엑셀 [설정] 시트 B5, B6 칸에서 직접 입력 가능합니다.
# =========================================================================
KR_PUBLIC_HOLIDAYS = {
    # 2026년
    (2026, 1): [1],
    (2026, 2): [16, 17, 18],
    (2026, 3): [1, 2],
    (2026, 5): [5, 24, 25],
    (2026, 6): [3, 6],
    (2026, 8): [15, 17],
    (2026, 9): [24, 25, 26],
    (2026, 10): [3, 5, 9],
    (2026, 12): [25],
    # 2027년
    (2027, 1): [1],
    (2027, 2): [6, 7, 8, 9],
    (2027, 3): [1],
    (2027, 5): [5, 13],
    (2027, 6): [6],
    (2027, 8): [15, 16],
    (2027, 9): [14, 15, 16],
    (2027, 10): [3, 4, 9, 11],
    (2027, 12): [25],
    # 2028년
    (2028, 1): [1, 26, 27, 28],
    (2028, 3): [1],
    (2028, 4): [12],
    (2028, 5): [2, 5],
    (2028, 6): [6],
    (2028, 8): [15],
    (2028, 10): [2, 3, 4, 5, 9],
    (2028, 12): [25],
}

# 기본 간호사 명단: (성명, 그룹, 추가오프, 최대N, 근무선호/비율)
# - 전월 결과 파일이나 기존 입력 파일이 없을 때만 사용됩니다.
DEFAULT_NURSE_LIST = [
    ("김재만", "A", 0, 5, "6:4"),
    ("이환", "A", 0, 6, ""),
    ("손경혜", "A", 0, 6, ""),
    ("이정은", "A", 0, 7, ""),
    ("이규리", "A", 0, 7, ""),
    ("한정민", "B", 0, 7, ""),
    ("김선민", "B", 0, 7, ""),
    ("김시아", "B", 0, 7, ""),
    ("유경빈", "B", 0, 6, ""),
    ("김상률", "C", 0, 7, ""),
    ("지수민", "C", 0, 7, ""),
    ("조정임", "C", 0, 7, ""),
    ("박완", "C", 0, 7, ""),
    ("정다희", "C", 0, 7, ""),
    ("김솔찬뫼", "D", 0, 7, ""),
    ("신수하", "D", 0, 7, ""),
    ("이경표", "D", 0, 7, ""),
    ("박여준", "D", 0, 7, ""),
    ("설지희", "D", 0, 7, ""),
    ("김은지", "E", 0, 7, ""),
    ("김동민", "E", 0, 7, ""),
    ("김민선", "E", 0, 7, ""),
    ("이서영", "E", 0, 7, ""),
]

# 기본 운영 내규 및 트렌드 설정 (기존 duty_input.xlsx에 설정값이 있으면 자동 유지됨)
DEFAULT_SETTINGS = {
    "daily_n_req": 5,
    "min_de_req": 5,
    "hol_de_req": 5,
    "busy_days": "일, 월, 화, 수",
    "max_consec_work": 5,
    "min_n_gap": 4,
    "three_n_strength": "보통",
}


def get_next_month_default():
  today = datetime.date.today()
  if today.month == 12:
    return today.year + 1, 1
  return today.year, today.month + 1


def get_prev_year_month(year, month):
  if month == 1:
    return year - 1, 12
  return year, month - 1


def normalize_prev_duty(val):
  v = str(val or "O").strip().upper()
  if not v or v.startswith("연") or v == "OFF":
    return "O"
  if v.startswith("교"):
    return "교"
  if v in ("D", "E", "N", "O"):
    return v
  return "O"


def find_prev_result_file(prev_y, prev_m):
  """같은 폴더에서 직전 달 근무표 결과 파일(수정본 최우선)을 자동으로 찾습니다."""
  patterns = [
      f"*duty_result*_{prev_y}_{prev_m:02d}*.xlsx",
      f"*duty_result*_{prev_y}_{prev_m}*.xlsx",
      f"*{prev_y}년*{prev_m}월*근무표*.xlsx",
  ]
  candidates = []
  for pat in patterns:
    for path in glob.glob(pat):
      if os.path.basename(path).startswith("~$"):
        continue
      if path not in candidates:
        candidates.append(path)

  if not candidates:
    return None

  candidates.sort(
      key=lambda p: (
          1 if ("수정" in os.path.basename(p) or "최종" in os.path.basename(p)) else 0,
          os.path.getmtime(p),
      ),
      reverse=True,
  )
  return candidates[0]


def load_carryover_from_result(filepath, prev_y, prev_m, existing_pref_map):
  """직전 달 결과 엑셀에서 (성명, 그룹, 0, N상한, 근무선호, [전월-3, 전월-2, 전월-1]) 목록을 추출합니다."""
  _, prev_num_days = calendar.monthrange(prev_y, prev_m)
  wb = openpyxl.load_workbook(filepath, data_only=True)
  ws = wb.active

  day1_col = 6
  n_cap_col = None
  for c in range(1, ws.max_column + 1):
    h = str(ws.cell(row=1, column=c).value or "").strip()
    if h.startswith("1\n") or h == "1":
      day1_col = c
      break
  for c in range(day1_col + prev_num_days, ws.max_column + 1):
    h = str(ws.cell(row=1, column=c).value or "").strip()
    if "N상한" in h or "최대N" in h:
      n_cap_col = c
      break

  last3_cols = [
      day1_col + prev_num_days - 3,
      day1_col + prev_num_days - 2,
      day1_col + prev_num_days - 1,
  ]

  roster = []
  r = 2
  while r <= ws.max_row:
    name = ws.cell(row=r, column=1).value
    if name is None or not str(name).strip():
      break
    name_str = str(name).strip()
    if any(k in name_str for k in ("인원", "합계", "총")):
      break
    grp = str(ws.cell(row=r, column=2).value or "A").strip().upper()
    max_n = 7
    if n_cap_col is not None:
      try:
        max_n = int(ws.cell(row=r, column=n_cap_col).value)
      except (TypeError, ValueError):
        max_n = 7

    # 기존 입력 파일에 저장된 개인별 근무선호/비율이 있으면 자동 유지
    if name_str in existing_pref_map:
      pref = existing_pref_map[name_str]
    elif name_str.rstrip("AB") == "김재만" and not existing_pref_map:
      pref = "6:4"
    else:
      pref = ""

    prev3 = [
        normalize_prev_duty(ws.cell(row=r, column=c).value) for c in last3_cols
    ]
    roster.append((name_str, grp, 0, max_n, pref, prev3))
    r += 1

  return roster if roster else None


def load_existing_input_info():
  """기존 duty_input.xlsx에서 명단(최대N·근무선호 포함), 근무분리_설정, [설정] 운영 규칙을 불러옵니다."""
  settings = dict(DEFAULT_SETTINGS)
  if not os.path.exists("duty_input.xlsx"):
    return None, {}, [], settings
  try:
    wb = openpyxl.load_workbook("duty_input.xlsx", data_only=True)
  except Exception:
    return None, {}, [], settings

  roster = []
  pref_map = {}
  if "근무표_입력" in wb.sheetnames:
    ws = wb["근무표_입력"]
    col4_h = str(ws.cell(row=1, column=4).value or "").strip()
    col5_h = str(ws.cell(row=1, column=5).value or "").strip()
    has_max_n = "N" in col4_h.upper() or "나이트" in col4_h
    has_pref = "선호" in col5_h or "비율" in col5_h or "전담" in col5_h

    r = 2
    while r <= ws.max_row:
      name = ws.cell(row=r, column=1).value
      if name is None or not str(name).strip():
        break
      name_str = str(name).strip()
      grp = str(ws.cell(row=r, column=2).value or "A").strip().upper()
      max_n = 7
      if has_max_n:
        try:
          max_n = int(ws.cell(row=r, column=4).value)
        except (TypeError, ValueError):
          max_n = 7
      pref = ""
      if has_pref:
        pref = str(ws.cell(row=r, column=5).value or "").strip()
      elif name_str.rstrip("AB") == "김재만":
        pref = "6:4"

      pref_map[name_str] = pref
      roster.append((name_str, grp, 0, max_n, pref, ["D", "D", "O"]))
      r += 1

  sep_rows = []
  if "근무분리_설정" in wb.sheetnames:
    ws_sep = wb["근무분리_설정"]
    r = 2
    while r <= ws_sep.max_row:
      n1 = str(ws_sep.cell(row=r, column=1).value or "").strip()
      n2 = str(ws_sep.cell(row=r, column=2).value or "").strip()
      tp = str(ws_sep.cell(row=r, column=3).value or "").strip()
      note = str(ws_sep.cell(row=r, column=4).value or "").strip()
      if n1 and n2:
        sep_rows.append((n1, n2, tp or "전부", note))
      r += 1

  if "설정" in wb.sheetnames:
    ws_rule = wb["설정"]
    for r in range(1, 25):
      k = str(ws_rule.cell(row=r, column=1).value or "").strip()
      v = ws_rule.cell(row=r, column=2).value
      if v is None or str(v).strip() == "":
        continue
      v_str = str(v).strip()
      try:
        if "나이트(N) 고정" in k or "나이트 고정" in k:
          settings["daily_n_req"] = int(float(v_str))
        elif "일반일 D" in k or "D·E 최소" in k:
          settings["min_de_req"] = int(float(v_str))
        elif "공휴일(수당일) D" in k or "수당일 D" in k:
          settings["hol_de_req"] = int(float(v_str))
        elif "바쁜 요일" in k:
          settings["busy_days"] = v_str
        elif "최대 연속 근무" in k:
          settings["max_consec_work"] = int(float(v_str))
        elif "최소 휴식일" in k or "나이트 간 최소" in k:
          settings["min_n_gap"] = int(float(v_str))
        elif "3연속 나이트" in k:
          if v_str in ("약함", "보통", "강함"):
            settings["three_n_strength"] = v_str
      except ValueError:
        pass

  return (roster if roster else None), pref_map, sep_rows, settings


def ask_year_month():
  def_y, def_m = get_next_month_default()
  print("=" * 66)
  print("  📋 중환자실(MICU) 근무표 입력 양식(duty_input.xlsx) 자동 생성기 v15")
  print("=" * 66)
  print(f"▶ 기본 추천 월: {def_y}년 {def_m}월")
  print("  - 그대로 만들려면 그냥 [Enter] 키를 누르세요.")
  print("  - 다른 달을 만들려면 '연도 월'을 띄어쓰기로 입력하세요 (예: 2027 2)")
  print("-" * 66)

  while True:
    raw = input(f"입력 (Enter = {def_y}년 {def_m}월): ").strip()
    if not raw:
      return def_y, def_m
    nums = [int(x) for x in re.findall(r"\d+", raw)]
    if len(nums) == 2 and 2020 <= nums[0] <= 2050 and 1 <= nums[1] <= 12:
      return nums[0], nums[1]
    elif len(nums) == 1 and 1 <= nums[0] <= 12:
      return def_y, nums[0]
    print("⚠️ 형식이 올바르지 않습니다. 예: '2027 1' 또는 그냥 Enter를 눌러주세요.")


def create_template_final():
  year, month = ask_year_month()
  _, num_days = calendar.monthrange(year, month)

  # 1. 법정 공휴일 및 일요일 우선 / 토요일 공휴일 차감 자동 계산
  if (year, month) in KR_PUBLIC_HOLIDAYS:
    raw_holidays = KR_PUBLIC_HOLIDAYS[(year, month)]
  else:
    fixed_solar = {
        1: [1],
        3: [1],
        5: [5],
        6: [6],
        8: [15],
        10: [3, 9],
        12: [25],
    }
    raw_holidays = fixed_solar.get(month, [])

  paid_holidays = sorted(
      d
      for d in set(raw_holidays)
      if 1 <= d <= num_days and calendar.weekday(year, month, d) != 6
  )
  paid_holiday_set = set(paid_holidays)

  base_off_count = sum(
      1
      for d in range(1, num_days + 1)
      if calendar.weekday(year, month, d) == 6
      or (calendar.weekday(year, month, d) == 5 and d not in paid_holiday_set)
  )

  # 2. 직전 달 결과 파일 탐색 -> 전월 말 3일 근무 자동 이월 + 기존 설정/근무선호 유지!
  prev_y, prev_m = get_prev_year_month(year, month)
  prev_file = find_prev_result_file(prev_y, prev_m)
  existing_roster, existing_pref_map, existing_sep_rows, kept_settings = (
      load_existing_input_info()
  )

  roster_data = None
  carryover_msg = ""

  if prev_file:
    try:
      roster_data = load_carryover_from_result(
          prev_file, prev_y, prev_m, existing_pref_map
      )
      if roster_data:
        carryover_msg = (
            f"🔄 직전 달 파일 '{os.path.basename(prev_file)}'에서"
            f" {len(roster_data)}명의 마지막 3일 근무를 자동 이월했습니다!"
        )
    except Exception as e:
      print(f"⚠️ 직전 달 파일({prev_file}) 읽기 중 오류 발생(기본값 사용): {e}")

  if not roster_data:
    if existing_roster:
      roster_data = existing_roster
      carryover_msg = (
          f"ℹ️ 직전 달({prev_y}년 {prev_m}월) 결과 파일이 없어, 기존"
          f" duty_input.xlsx의 명단({len(roster_data)}명)을 불러왔습니다."
          " (전월-3~1 칸은 직접 확인해 주세요)"
      )
    else:
      roster_data = [
          (n, g, e, m, pref, ["D", "D", "O"])
          for (n, g, e, m, pref) in DEFAULT_NURSE_LIST
      ]
      carryover_msg = (
          f"ℹ️ 기본 간호사 명단({len(roster_data)}명)으로 양식을 생성했습니다."
          " (전월-3~1 칸을 꼭 입력해 주세요)"
      )

  if os.path.exists("duty_input.xlsx"):
    try:
      shutil.copy2("duty_input.xlsx", "duty_input_이전백업.xlsx")
    except Exception:
      pass

  # 3. 엑셀 양식 생성
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "근무표_입력"

  header_fill = PatternFill(
      start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
  )
  group_fill = PatternFill(
      start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
  )
  pref_fill = PatternFill(
      start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
  )
  prev_fill = PatternFill(
      start_color="EDEDED", end_color="EDEDED", fill_type="solid"
  )
  weekend_fill = PatternFill(
      start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
  )
  thin_border = Border(
      left=Side(style="thin", color="BFBFBF"),
      right=Side(style="thin", color="BFBFBF"),
      top=Side(style="thin", color="BFBFBF"),
      bottom=Side(style="thin", color="BFBFBF"),
  )
  thick_right = Border(
      left=Side(style="thin", color="BFBFBF"),
      right=Side(style="medium", color="595959"),
      top=Side(style="thin", color="BFBFBF"),
      bottom=Side(style="thin", color="BFBFBF"),
  )
  bold_font = Font(bold=True)
  center_align = Alignment(horizontal="center", vertical="center")

  headers = [
      (1, f"{year}년 {month}월", header_fill),
      (2, "그룹(조)", group_fill),
      (3, "추가오프", group_fill),
      (4, "최대N(기본7)", group_fill),
      (5, "근무선호/비율", pref_fill),
      (6, "전월-3", prev_fill),
      (7, "전월-2", prev_fill),
      (8, "전월-1", prev_fill),
  ]
  for col_idx, text, fill in headers:
    c = ws.cell(row=1, column=col_idx, value=text)
    c.fill = fill
    c.font = bold_font
    c.alignment = center_align
    c.border = thick_right if col_idx == 8 else thin_border

  weekdays_kr = ["월", "화", "수", "목", "금", "토", "일"]
  for day in range(1, num_days + 1):
    col = day + 8
    wd_idx = calendar.weekday(year, month, day)
    is_hol = day in paid_holiday_set
    cell = ws.cell(row=1, column=col, value=f"{day}\n({weekdays_kr[wd_idx]})")
    cell.font = Font(bold=True, color="C00000" if is_hol else "000000")
    cell.alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    cell.border = thin_border
    cell.fill = weekend_fill if (wd_idx >= 5 or is_hol) else header_fill

  for idx, (name, grp, extra_off, max_n, pref, prev3) in enumerate(
      roster_data, start=2
  ):
    c_name = ws.cell(row=idx, column=1, value=name)
    c_name.font = bold_font
    c_name.alignment = center_align
    c_name.border = thin_border

    c_grp = ws.cell(row=idx, column=2, value=grp)
    c_grp.alignment = center_align
    c_grp.fill = group_fill
    c_grp.border = thin_border

    c_ext = ws.cell(row=idx, column=3, value=extra_off)
    c_ext.alignment = center_align
    c_ext.border = thin_border

    c_maxn = ws.cell(row=idx, column=4, value=max_n)
    c_maxn.alignment = center_align
    c_maxn.fill = group_fill
    c_maxn.border = thin_border

    c_pref = ws.cell(row=idx, column=5, value=pref)
    c_pref.alignment = center_align
    c_pref.fill = pref_fill
    c_pref.border = thin_border

    for p_offset, val in enumerate(prev3):
      p_col = 6 + p_offset
      c_prev = ws.cell(row=idx, column=p_col, value=val)
      c_prev.fill = prev_fill
      c_prev.alignment = center_align
      c_prev.border = thick_right if p_col == 8 else thin_border

    for day in range(1, num_days + 1):
      c_day = ws.cell(row=idx, column=day + 8, value="")
      c_day.alignment = center_align
      c_day.border = thin_border

  ws.column_dimensions["A"].width = 13
  ws.column_dimensions["B"].width = 9
  ws.column_dimensions["C"].width = 10
  ws.column_dimensions["D"].width = 12
  ws.column_dimensions["E"].width = 14
  for col_idx in range(6, 9):
    ws.column_dimensions[get_column_letter(col_idx)].width = 7.5
  for col_idx in range(9, num_days + 9):
    ws.column_dimensions[get_column_letter(col_idx)].width = 5.5
  ws.row_dimensions[1].height = 28
  ws.freeze_panes = "I2"

  # [시트 2] 근무분리_설정 (기존 설정 자동 유지)
  ws_sep = wb.create_sheet(title="근무분리_설정")
  ws_sep["A1"] = "간호사 1 성명"
  ws_sep["B1"] = "간호사 2 성명"
  ws_sep["C1"] = "분리할 듀티 (D,E,N 또는 전부)"
  ws_sep["D1"] = "비고/사유"
  for col in ["A1", "B1", "C1", "D1"]:
    ws_sep[col].font = bold_font
    ws_sep[col].fill = header_fill
    ws_sep[col].alignment = center_align
    ws_sep[col].border = thin_border

  if existing_sep_rows:
    for r_idx, (n1, n2, tp, note) in enumerate(existing_sep_rows, start=2):
      ws_sep.cell(row=r_idx, column=1, value=n1).alignment = center_align
      ws_sep.cell(row=r_idx, column=2, value=n2).alignment = center_align
      ws_sep.cell(row=r_idx, column=3, value=tp).alignment = center_align
      ws_sep.cell(row=r_idx, column=4, value=note)
  else:
    ws_sep["D2"] = (
        "※ 김재만-이환 나이트 분리는 자동 내장됨. 추가 분리 인원만 아래에"
        " 입력하세요."
    )

  for col_letter, width in [("A", 16), ("B", 16), ("C", 26), ("D", 55)]:
    ws_sep.column_dimensions[col_letter].width = width

  # [시트 3] 설정 및 운영 규칙 안내
  ws_rule = wb.create_sheet(title="설정")
  rule_rows = [
      ("연도", year, "달력 기준 연도 (자동 계산)", header_fill),
      ("월", month, "달력 기준 월 (자동 계산)", header_fill),
      (
          "필수기본오프(자동계산)",
          base_off_count,
          "일요일 전체 + 공휴일 아닌 토요일",
          header_fill,
      ),
      (
          "자동인식 수당일(일요일제외)",
          ", ".join(str(d) for d in paid_holidays) if paid_holidays else "없음",
          "일요일 공휴일은 일반 기본오프 처리",
          header_fill,
      ),
      (
          "수당일 수동 추가(임시공휴일)",
          "",
          "예: 15 또는 2, 15 (쉼표 구분)",
          group_fill,
      ),
      (
          "수당일 수동 제외(선택)",
          "",
          "자동인식 수당일 중 제외할 날짜만 입력",
          group_fill,
      ),
      (
          "일일 나이트(N) 고정 인원",
          kept_settings["daily_n_req"],
          "매일 고정 배정할 N 인원 (MICU 기본 5, 타 병동 변경 가능)",
          pref_fill,
      ),
      (
          "일반일 D·E 최소 필요 인원",
          kept_settings["min_de_req"],
          "평일·주말 최소 인원 (최소+2명까지 일반 배정, +3명 안전밸브)",
          pref_fill,
      ),
      (
          "공휴일(수당일) D·E 고정 인원",
          kept_settings["hol_de_req"],
          "수당일 D·E 고정 인원 (MICU 기본 5)",
          pref_fill,
      ),
      (
          "바쁜 요일 (우선 배치)",
          kept_settings["busy_days"],
          "인원이 여유로울 때 먼저 6명 이상 채울 요일 (예: 수, 목, 금)",
          pref_fill,
      ),
      (
          "최대 연속 근무일수",
          kept_settings["max_consec_work"],
          "교육 포함 연속 근무 상한 (기본 5일)",
          pref_fill,
      ),
      (
          "나이트 간 최소 휴식일",
          kept_settings["min_n_gap"],
          "N 블록 종료 후 다음 N 시작 전 최소 비나이트 간격 (기본 4일)",
          pref_fill,
      ),
      (
          "3연속 나이트 억제 강도",
          kept_settings["three_n_strength"],
          "약함 / 보통 / 강함 중 선택 (강함일수록 2연속 위주 유도)",
          pref_fill,
      ),
  ]

  for r_idx, (k, v, desc, fill) in enumerate(rule_rows, start=1):
    ca = ws_rule.cell(row=r_idx, column=1, value=k)
    cb = ws_rule.cell(row=r_idx, column=2, value=v)
    cc = ws_rule.cell(row=r_idx, column=3, value=desc)
    ca.font = bold_font
    ca.fill = fill
    ca.border = thin_border
    cb.alignment = center_align
    cb.border = thin_border
    cc.border = thin_border
    cc.font = Font(color="595959")

  guides = [
      "💡 자동화 근무표(v15) 사용 가이드",
      (
          "1. [설정] 시트 B7~B13에서 일일 N 인원, 바쁜 요일, 연속근무/휴식일,"
          " 3연속 나이트 강도를 코드 수정 없이 바꿀 수 있으며 다음 달에도 자동"
          " 유지됩니다."
      ),
      (
          "2. [근무표_입력] E열(근무선호/비율) 입력 방법 (비워두면 일반 3교대"
          " D/E 균등):"
      ),
      (
          "   • 비율 지정: '6:4', '7:3', '8:2', '3:7' 등 (앞 숫자=D 비중, 뒤"
          " 숫자=E 비중)"
      ),
      (
          "   • 특정 듀티 위주: 'DN'(데이·나이트 위주), 'EN'(이브닝·나이트"
          " 위주), 'DE'(데이·이브닝 위주)"
      ),
      (
          "   • 전담 유도: 'D', 'E', 'N' 입력 시 해당 듀티 외 나머지 듀티를"
          " 강력 억제 (완전 차단 시 최대N=0 병행)"
      ),
      (
          "3. [8명 안전밸브 내장]: 일반일 D·E는 '최소인원+2명'(기본 7명)까지"
          " 우선 배정되며, 27명 완편·기본오프 8일 등 불가피한 초과 시에만"
          " '최소인원+3명'(8명)까지 자동 흡수해 초과 오프를 방지합니다."
      ),
      (
          "4. 인력 변동 시: 퇴사자는 '행 전체 삭제'(중간 빈 줄 금지), 신규"
          " 입사자는 맨 아래 줄에 성명·그룹·추가오프·최대N·근무선호·전월근무"
          "(O,O,O)를 추가하면 됩니다."
      ),
      (
          "5. 원티드 입력: 희망 오프 'O', 연차 '연1'/'연3', 교육 공가 '교',"
          " 고정 근무 'D'/'E'/'N' 입력 후 저장하고 '2_근무표_생성하기.bat'을"
          " 실행하세요."
      ),
  ]
  for idx, text in enumerate(guides, start=15):
    ws_rule[f"A{idx}"] = text
    if idx == 15:
      ws_rule[f"A{idx}"].font = bold_font

  ws_rule.column_dimensions["A"].width = 30
  ws_rule.column_dimensions["B"].width = 22
  ws_rule.column_dimensions["C"].width = 58

  try:
    wb.save("duty_input.xlsx")
  except PermissionError:
    print(
        "\n❌ 오류: 현재 'duty_input.xlsx' 파일이 엑셀에서 열려 있습니다!"
    )
    print("👉 열려 있는 엑셀 창을 닫은 뒤 다시 실행해 주세요.")
    return

  print("\n" + "=" * 66)
  print(f"✅ 'duty_input.xlsx' 생성 완료! ({year}년 {month}월, 총 {num_days}일)")
  print(
      f"   • 필수 기본오프: {base_off_count}개 | 확정 수당일(일요일 제외):"
      f" {paid_holidays if paid_holidays else '없음'}"
  )
  print(
      f"   • 운영 설정: N={kept_settings['daily_n_req']}명 / 일반 최소"
      f" D·E={kept_settings['min_de_req']}명 / 바쁜요일="
      f"[{kept_settings['busy_days']}] / 연속근무≤"
      f"{kept_settings['max_consec_work']}일 / N휴식≥"
      f"{kept_settings['min_n_gap']}일"
  )
  print(f"   • {carryover_msg}")
  print("=" * 66)


if __name__ == "__main__":
  create_template_final()
