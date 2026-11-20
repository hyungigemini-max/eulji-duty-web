import calendar
from collections import defaultdict
import datetime
import math
import re
import openpyxl
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from ortools.sat.python import cp_model

# =========================================================================
# [기본 예비 설정]
# - 엑셀(duty_input.xlsx)의 D열 '최대N', E열 '근무선호/비율', [설정] 시트 값이 항상 최우선입니다.
# - 아래 CUSTOM_MAX_N은 D열이 없는 구버전 엑셀 파일을 열 때만 쓰입니다.
DEFAULT_MAX_N = 7
CUSTOM_MAX_N = {
    "김재만": 5,
    "이환": 6,
    "손경혜": 6,
    "유경빈": 6,
}

# [법정 공휴일 자동 달력 (2026~2028년)]
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

WEEKDAY_CHAR_TO_IDX = {
    "월": 0,
    "화": 1,
    "수": 2,
    "목": 3,
    "금": 4,
    "토": 5,
    "일": 6,
}
IDX_TO_WEEKDAY_CHAR = ["월", "화", "수", "목", "금", "토", "일"]
# =========================================================================


def parse_busy_weekdays(raw_val):
  """'일, 월, 화, 수' 또는 '수목금' 문자열에서 요일 인덱스 집합을 추출합니다."""
  if raw_val is None or not str(raw_val).strip():
    return {6, 0, 1, 2}
  text = str(raw_val).strip()
  if text in ("없음", "NONE", "0", "-"):
    return set()
  found = {
      idx for ch, idx in WEEKDAY_CHAR_TO_IDX.items() if ch in text
  }
  return found if found else {6, 0, 1, 2}


def load_operational_settings(ws_rule):
  """[설정] 시트에서 병동 운영 내규 및 트렌드 설정을 읽어옵니다 (구버전 엑셀 자동 호환)."""
  cfg = {
      "daily_n_req": 5,
      "min_de_req": 5,
      "hol_de_req": 5,
      "busy_weekdays": {6, 0, 1, 2},
      "max_consec_work": 5,
      "min_n_gap": 4,
      "three_n_strength": "보통",
  }

  for r in range(1, 25):
    k = str(ws_rule.cell(row=r, column=1).value or "").strip()
    v = ws_rule.cell(row=r, column=2).value
    if not k or v is None or str(v).strip() == "":
      continue
    v_str = str(v).strip()
    try:
      if "나이트(N) 고정" in k or "나이트 고정" in k:
        cfg["daily_n_req"] = max(1, min(15, int(float(v_str))))
      elif "일반일 D" in k or "D·E 최소" in k:
        cfg["min_de_req"] = max(1, min(15, int(float(v_str))))
      elif "공휴일(수당일) D" in k or "수당일 D" in k:
        cfg["hol_de_req"] = max(1, min(15, int(float(v_str))))
      elif "바쁜 요일" in k:
        cfg["busy_weekdays"] = parse_busy_weekdays(v_str)
      elif "최대 연속 근무" in k:
        cfg["max_consec_work"] = max(3, min(7, int(float(v_str))))
      elif "최소 휴식일" in k or "나이트 간 최소" in k:
        cfg["min_n_gap"] = max(2, min(7, int(float(v_str))))
      elif "3연속 나이트" in k:
        if v_str in ("약함", "보통", "강함"):
          cfg["three_n_strength"] = v_str
    except ValueError:
      pass

  return cfg


def parse_nurse_preference(raw_str):
  """
  개인별 근무선호/비율 문자열을 해석합니다.
  - '7:3', '6:4', '8:2' 등 -> ('ratio', d_weight, e_weight)
  - 'DN', 'EN', 'DE', 'D', 'E', 'N' 등 -> ('shifts', allowed_shift_indices_set)
  - 빈칸 또는 미인식 -> None (일반 3교대 D/E 1:1 균형)
  """
  if not raw_str:
    return None
  s = str(raw_str).strip().upper().replace(" ", "")
  if not s:
    return None

  # 1. 비율 표기 (예: 7:3, 6:4, 7/3)
  m = re.match(r"^(\d+(?:\.\d+)?)[:/](\d+(?:\.\d+)?)$", s)
  if m:
    d_val = float(m.group(1))
    e_val = float(m.group(2))
    if d_val > 0 and e_val > 0:
      d_int = max(1, int(round(d_val * 10)))
      e_int = max(1, int(round(e_val * 10)))
      g = math.gcd(d_int, e_int)
      return ("ratio", d_int // g, e_int // g)
    elif d_val > 0 and e_val == 0:
      return ("shifts", {0, 2})
    elif e_val > 0 and d_val == 0:
      return ("shifts", {1, 2})

  # 2. 듀티 조합 표기 (예: DN, EN, DE, D전담, E전담, N전담)
  s_clean = s.replace("전담", "").replace("위주", "").replace(",", "").replace("-", "")
  chars = {ch for ch in s_clean if ch in ("D", "E", "N")}
  if chars and len(chars) < 3:
    c2i = {"D": 0, "E": 1, "N": 2}
    return ("shifts", {c2i[ch] for ch in chars})

  return None


def get_month_holidays_and_base_off(ws_rule, year, month, num_days):
  """[설정] 시트의 연·월 및 수동 추가/제외 기준으로 수당일과 필수기본오프를 계산합니다."""
  if (year, month) in KR_PUBLIC_HOLIDAYS:
    raw_holidays = set(KR_PUBLIC_HOLIDAYS[(year, month)])
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
    raw_holidays = set(fixed_solar.get(month, []))

  has_new_override_rows = False
  add_days = set()
  remove_days = set()
  legacy_days = set()

  for r in range(1, 15):
    k = str(ws_rule.cell(row=r, column=1).value or "").strip()
    v = ws_rule.cell(row=r, column=2).value
    nums = (
        [int(x) for x in re.findall(r"\d+", str(v))]
        if v is not None
        else []
    )
    if "수동 추가" in k or "임시공휴일" in k:
      has_new_override_rows = True
      add_days.update(d for d in nums if 1 <= d <= num_days)
    elif "수동 제외" in k or "공휴일 제외" in k:
      has_new_override_rows = True
      remove_days.update(d for d in nums if 1 <= d <= num_days)
    elif (
        ("공휴일" in k or "수당일" in k)
        and "자동인식" not in k
        and "D·E" not in k
        and "고정" not in k
        and nums
    ):
      legacy_days.update(d for d in nums if 1 <= d <= num_days)

  if has_new_override_rows:
    raw_holidays.update(add_days)
    raw_holidays.difference_update(remove_days)
  elif legacy_days:
    raw_holidays = legacy_days

  paid_holidays = sorted(
      d
      for d in raw_holidays
      if 1 <= d <= num_days and calendar.weekday(year, month, d) != 6
  )
  paid_holiday_set = set(paid_holidays)

  base_off_days = sum(
      1
      for d in range(1, num_days + 1)
      if calendar.weekday(year, month, d) == 6
      or (calendar.weekday(year, month, d) == 5 and d not in paid_holiday_set)
  )

  return paid_holidays, base_off_days


def make_off_count_formula_for_cells(col_letters, row_idx):
  if not col_letters:
    return "=0"
  terms = []
  for cl in col_letters:
    cell_ref = f"{cl}{row_idx}"
    terms.append(
        f'COUNTIF({cell_ref},"O")+COUNTIF({cell_ref},"OFF")+COUNTIF({cell_ref},"연*")'
    )
  return "=" + "+".join(terms)


def solve_duty_final():
  print("=" * 68)
  print("  🏥 중환자실(MICU) 자동 근무표 생성기 (확장 완성판 v15)")
  print("=" * 68)

  try:
    wb = openpyxl.load_workbook("duty_input.xlsx", data_only=True)
  except FileNotFoundError:
    print(
        "❌ 'duty_input.xlsx' 파일이 없습니다. 먼저 '1_새양식_만들기.bat'을"
        " 실행해 주세요."
    )
    return
  except PermissionError:
    print(
        "❌ 'duty_input.xlsx' 파일이 엑셀에 열려 있어 읽을 수 없습니다. 엑셀을"
        " 닫고 다시 실행해 주세요."
    )
    return

  ws = wb["근무표_입력"]
  ws_rule = wb["설정"]

  try:
    year = int(ws_rule["B1"].value)
    month = int(ws_rule["B2"].value)
  except (TypeError, ValueError):
    print("❌ [설정] 시트의 B1(연도), B2(월)에 올바른 숫자가 입력되어 있는지 확인해 주세요.")
    return

  _, num_days = calendar.monthrange(year, month)
  holiday_list, base_off_days = get_month_holidays_and_base_off(
      ws_rule, year, month, num_days
  )
  holiday_set = set(holiday_list)

  # [v15 신규] 엑셀 [설정] 시트에서 운영 내규 및 트렌드 설정 로드
  cfg = load_operational_settings(ws_rule)
  daily_n_req = cfg["daily_n_req"]
  min_de_req = cfg["min_de_req"]
  hol_de_req = cfg["hol_de_req"]
  busy_weekdays = cfg["busy_weekdays"]
  max_consec_work = cfg["max_consec_work"]
  min_n_gap = cfg["min_n_gap"]
  three_n_strength = cfg["three_n_strength"]

  three_n_pen_map = {"약함": 80, "보통": 350, "강함": 1200}
  three_n_penalty = three_n_pen_map.get(three_n_strength, 350)

  busy_str = (
      ", ".join(IDX_TO_WEEKDAY_CHAR[i] for i in [6, 0, 1, 2, 3, 4, 5] if i in busy_weekdays)
      if busy_weekdays
      else "없음"
  )

  raw_wknd = sum(
      1
      for d in range(1, num_days + 1)
      if calendar.weekday(year, month, d) in (5, 6)
  )
  sat_hols = [d for d in holiday_list if calendar.weekday(year, month, d) == 5]

  print(
      f"📅 대상 월: {year}년 {month}월 (총 {num_days}일) | 단순 토+일:"
      f" {raw_wknd}개 (토요일 공휴일: {sat_hols if sat_hols else '없음'})"
  )
  print(
      f"   => 확정 필수기본오프: {base_off_days}개 | 확정 수당일(일요일 제외):"
      f" {holiday_list if holiday_list else '없음'}"
  )
  print(
      f"⚙️ 운영 설정: 일일 N={daily_n_req}명 | 일반 D·E={min_de_req}~{min_de_req+2}명"
      f"(안전밸브 {min_de_req+3}명) | 수당일 D·E={hol_de_req}명"
  )
  print(
      f"   => 바쁜 요일=[{busy_str}] | 연속근무≤{max_consec_work}일 |"
      f" N휴식≥{min_n_gap}일 | 3연속N 억제=[{three_n_strength}]"
  )

  # 열 구조 자동 판별 (v15 / v14 / 구버전 모두 호환)
  col4_header = str(ws.cell(row=1, column=4).value or "").strip()
  col5_header = str(ws.cell(row=1, column=5).value or "").strip()
  has_max_n_col = "N" in col4_header.upper() or "나이트" in col4_header
  has_pref_col = (
      has_max_n_col
      and ("선호" in col5_header or "비율" in col5_header or "전담" in col5_header)
  )

  if has_pref_col:
    prev_start_col = 6
    day_start_col = 9
  elif has_max_n_col:
    prev_start_col = 5
    day_start_col = 8
  else:
    prev_start_col = 4
    day_start_col = 7

  nurses = []
  nurse_groups = []
  extra_offs = []
  max_n_limits = []
  nurse_prefs_raw = []
  nurse_prefs = []
  prev_duties = []

  duty_map = {"D": 0, "E": 1, "N": 2, "O": 3, "OFF": 3, "교": 4}
  rev_map = {0: "D", 1: "E", 2: "N", 3: "O", 4: "교"}

  fixed_requests = {}
  raw_request_labels = {}

  row = 2
  while ws.cell(row=row, column=1).value is not None:
    raw_name = str(ws.cell(row=row, column=1).value).strip()
    if not raw_name:
      break
    n_idx = len(nurses)
    nurses.append(raw_name)

    grp = str(ws.cell(row=row, column=2).value or "기본").strip().upper()
    nurse_groups.append(grp)

    try:
      ext = int(ws.cell(row=row, column=3).value or 0)
    except ValueError:
      ext = 0
    extra_offs.append(ext)

    if has_max_n_col and ws.cell(row=row, column=4).value is not None:
      try:
        n_lim = int(ws.cell(row=row, column=4).value)
      except ValueError:
        n_lim = DEFAULT_MAX_N
    else:
      n_lim = DEFAULT_MAX_N

    clean_name = raw_name.rstrip("AB")
    if raw_name in CUSTOM_MAX_N and not has_max_n_col:
      n_lim = CUSTOM_MAX_N[raw_name]
    elif clean_name in CUSTOM_MAX_N and not has_max_n_col:
      n_lim = CUSTOM_MAX_N[clean_name]
    max_n_limits.append(n_lim)

    if has_pref_col:
      p_raw = str(ws.cell(row=row, column=5).value or "").strip()
    else:
      p_raw = ""
    nurse_prefs_raw.append(p_raw)
    nurse_prefs.append(parse_nurse_preference(p_raw))

    p_list = []
    for p_col in range(prev_start_col, prev_start_col + 3):
      val = str(ws.cell(row=row, column=p_col).value or "O").strip().upper()
      if val.startswith("연") or val == "OFF":
        p_list.append(3)
      elif val.startswith("교"):
        p_list.append(4)
      else:
        p_list.append(duty_map.get(val, 3))
    prev_duties.append(p_list)

    for d in range(num_days):
      val = ws.cell(row=row, column=day_start_col + d).value
      if val is not None:
        v_str = str(val).strip().upper()
        if not v_str:
          continue
        if v_str.startswith("교"):
          fixed_requests[(n_idx, d)] = 4
          raw_request_labels[(n_idx, d)] = v_str
        elif v_str.startswith("연"):
          fixed_requests[(n_idx, d)] = 3
          raw_request_labels[(n_idx, d)] = v_str
        elif v_str in duty_map:
          fixed_requests[(n_idx, d)] = duty_map[v_str]
          raw_request_labels[(n_idx, d)] = (
              "O" if duty_map[v_str] == 3 else v_str
          )
    row += 1

  num_nurses = len(nurses)
  if num_nurses == 0:
    print("❌ 간호사 명단이 비어 있습니다. [근무표_입력] 시트를 확인해 주세요.")
    return

  # =========================================================================
  # [사전 진단 안전장치] 인력 변동 및 내규 설정 시 수학적 가능 여부 사전 점검!
  # =========================================================================
  total_required_n = num_days * daily_n_req
  total_allowed_n = sum(max_n_limits)
  strict_n_cap = total_allowed_n >= total_required_n

  group_dict = defaultdict(list)
  for idx, grp in enumerate(nurse_groups):
    group_dict[grp].append(idx)

  grp_A = group_dict.get("A", [])
  grp_B = group_dict.get("B", [])
  grp_D = group_dict.get("D", [])
  grp_E = group_dict.get("E", [])
  grp_DE = grp_D + grp_E
  non_de_indices = [i for i in range(num_nurses) if i not in grp_DE]

  non_de_n_cap = sum(max_n_limits[i] for i in non_de_indices)
  required_non_de_n = num_days * max(0, daily_n_req - 2)
  strict_de_n_cap = (not grp_DE) or (non_de_n_cap >= required_non_de_n)

  pref_summary = [
      f"{nurses[i]}({nurse_prefs_raw[i]})"
      for i in range(num_nurses)
      if nurse_prefs[i] is not None
  ]

  print(
      f"👥 로드된 간호사: 총 {num_nurses}명 | 개인별 N상한 합계:"
      f" {total_allowed_n}개 (당월 필요 N: {total_required_n}개)"
  )
  if pref_summary:
    print(f"🎯 개인별 근무선호/비율 적용: {', '.join(pref_summary)}")

  if not strict_n_cap:
    print(
        f"⚠️ [인원/나이트 총량 부족 경고] 현재 명단의 N상한 합계({total_allowed_n}개)가"
        f" 당월 필요량({total_required_n}개)보다 {total_required_n - total_allowed_n}개"
        " 부족합니다!"
    )
    print(
        "   👉 일부 간호사가 N상한을 초과 배정될 수 있습니다. (엑셀 D열 최대N 또는"
        " 일일 N 인원 조정을 검토하세요)"
    )
  if not strict_de_n_cap:
    print(
        f"⚠️ [경력자 N 총량 알림] A·B·C그룹의 N상한 합계({non_de_n_cap}개)가"
        f" 기준({required_non_de_n}개)보다 부족하여, 나이트 D+E 3명 제한을"
        " '최소화 모드'로 자동 전환합니다."
    )

  name_to_idx = {name: i for i, name in enumerate(nurses)}
  clean_name_to_idx = {name.rstrip("AB"): i for i, name in enumerate(nurses)}

  incompatible_pairs = []
  if "근무분리_설정" in wb.sheetnames:
    ws_sep = wb["근무분리_설정"]
    s_row = 2
    while ws_sep.cell(row=s_row, column=1).value is not None:
      n1_name = str(ws_sep.cell(row=s_row, column=1).value or "").strip()
      n2_name = str(ws_sep.cell(row=s_row, column=2).value or "").strip()
      sep_type = (
          str(ws_sep.cell(row=s_row, column=3).value or "전부").strip().upper()
      )
      if (
          n1_name in name_to_idx
          and n2_name in name_to_idx
          and n1_name != n2_name
      ):
        if "전부" in sep_type or "ALL" in sep_type:
          target_shifts = [0, 1, 2]
        else:
          target_shifts = [
              duty_map[ch] for ch in ("D", "E", "N") if ch in sep_type
          ]
        if target_shifts:
          incompatible_pairs.append(
              (name_to_idx[n1_name], name_to_idx[n2_name], target_shifts)
          )
      s_row += 1

  total_days = 3 + num_days
  model = cp_model.CpModel()

  work = {}
  for n in range(num_nurses):
    for t in range(total_days):
      for s in range(5):
        work[(n, t, s)] = model.NewBoolVar(f"w_{n}_{t}_{s}")

  # 1. 하루 1개 상태 배정 및 전월 말 3일 고정
  for n in range(num_nurses):
    for t in range(total_days):
      model.Add(sum(work[(n, t, s)] for s in range(5)) == 1)
    for p_t in range(3):
      model.Add(work[(n, p_t, prev_duties[n][p_t])] == 1)
      if prev_duties[n][p_t] != 4:
        model.Add(work[(n, p_t, 4)] == 0)
    for d in range(num_days):
      t = d + 3
      if fixed_requests.get((n, d)) != 4:
        model.Add(work[(n, t, 4)] == 0)

  # 2. 일자별 인원 및 안전밸브(최소+3명) 스마트 자동 활성화 판단
  #    - 일반 상한(기본 7명)으로 당월 근무 총량을 충분히 수용 가능하면 7명 상한을 Hard로 잠가
  #      23~25명 체계에서 8명 배치가 절대 발생하지 않게 보호합니다.
  #    - 27명 완편·기본오프 8일 등 총량이 상한에 근접/초과할 때만 8명 안전밸브를 자동 활성화합니다.
  normal_de_cap = min_de_req + 2
  valve_de_cap = min_de_req + 3

  reg_days_cnt = num_days - len(holiday_set)
  reg_edu_cnt = sum(
      1
      for (_, d_idx), s_req in fixed_requests.items()
      if s_req == 4 and (d_idx + 1) not in holiday_set
  )
  required_reg_work = (
      num_nurses * reg_days_cnt
      - sum(base_off_days + extra_offs[n] for n in range(num_nurses))
      - reg_edu_cnt
  )
  normal_reg_cap = reg_days_cnt * (2 * normal_de_cap + daily_n_req)
  has_high_wanted = any(
      sum(
          1
          for n in range(num_nurses)
          if fixed_requests.get((n, d_idx)) == s_chk
      )
      > normal_de_cap
      for d_idx in range(num_days)
      if (d_idx + 1) not in holiday_set
      for s_chk in (0, 1)
  )
  use_valve_cap = (required_reg_work > normal_reg_cap - 6) or has_high_wanted
  active_de_cap = valve_de_cap if use_valve_cap else normal_de_cap

  if use_valve_cap:
    print(
        f"🛡️ [안전밸브 자동 활성화] 당월 일반일 필요 근무({required_reg_work}일)가"
        f" {normal_de_cap}명 상한({normal_reg_cap}일)에 근접/초과하여"
        f" {valve_de_cap}명 안전밸브를 켭니다."
    )

  for d in range(num_days):
    t = d + 3
    day_num = d + 1
    d_sum = sum(work[(n, t, 0)] for n in range(num_nurses))
    e_sum = sum(work[(n, t, 1)] for n in range(num_nurses))
    n_sum = sum(work[(n, t, 2)] for n in range(num_nurses))

    if day_num in holiday_set:
      model.Add(d_sum == hol_de_req)
      model.Add(e_sum == hol_de_req)
      model.Add(n_sum == daily_n_req)
    else:
      model.Add(d_sum >= min_de_req)
      model.Add(d_sum <= active_de_cap)
      model.Add(e_sum >= min_de_req)
      model.Add(e_sum <= active_de_cap)
      model.Add(n_sum == daily_n_req)

  # 3. 기본 근무 분리 시트 조건
  for n1, n2, shifts in incompatible_pairs:
    for t in range(3, total_days):
      for s in shifts:
        model.Add(work[(n1, t, s)] + work[(n2, t, s)] <= 1)

  penalties = []

  # [규칙 1] 김재만(주임) & 이환(동기) 분리 + 구버전 호환 비율 처리
  idx_jm = clean_name_to_idx.get("김재만")
  idx_lh = clean_name_to_idx.get("이환")

  if idx_jm is not None and idx_lh is not None:
    for t in range(3, total_days):
      model.Add(work[(idx_jm, t, 2)] + work[(idx_lh, t, 2)] <= 1)
      for s in [0, 1]:
        both_de = model.NewBoolVar(f"jmlh_{t}_{s}")
        model.Add(work[(idx_jm, t, s)] + work[(idx_lh, t, s)] - 1 <= both_de)
        penalties.append(both_de * 180)

  # 구버전(E열 없음) 엑셀로 실행할 때만 기존 김재만 주임 D/E 비율 규칙 자동 유지
  if (not has_pref_col) and (idx_jm is not None):
    jm_d = sum(work[(idx_jm, t, 0)] for t in range(3, total_days))
    jm_e = sum(work[(idx_jm, t, 1)] for t in range(3, total_days))
    model.Add(10 * jm_d >= 13 * jm_e)
    jm_short_50 = model.NewIntVar(0, num_days * 3, "jm_short_50")
    model.Add(jm_short_50 >= 3 * jm_e - 2 * jm_d)
    penalties.append(jm_short_50 * 400)

  # [규칙 1-2 (v15 신규)] 개인별 근무선호/비율 (E열) 반영
  for n in range(num_nurses):
    pref = nurse_prefs[n]
    if pref is None:
      continue

    if pref[0] == "ratio":
      _, ra, rb = pref
      d_cnt_n = sum(work[(n, t, 0)] for t in range(3, total_days))
      e_cnt_n = sum(work[(n, t, 1)] for t in range(3, total_days))
      ratio_diff = model.NewIntVar(
          0, num_days * max(ra, rb), f"ratio_diff_{n}"
      )
      model.AddAbsEquality(ratio_diff, rb * d_cnt_n - ra * e_cnt_n)
      # 비율 숫자 크기에 관계없이 일정한 강도 유지 (6:4 -> 3:2 -> 가중치 400)
      weight = max(120, int(round(2000 / (ra + rb))))
      penalties.append(ratio_diff * weight)

    elif pref[0] == "shifts":
      allowed_shifts = pref[1]
      for s in (0, 1, 2):
        if s not in allowed_shifts:
          # 완전 배제(Hard) 대신 초강력 억제(Soft 3,000)로 인력 비상 시에만 예외 허용
          avoid_cnt = sum(work[(n, t, s)] for t in range(3, total_days))
          penalties.append(avoid_cnt * 3000)

  # [규칙 2] 그룹별(A~E) 팀차지·백업·주니어 제한
  for t in range(3, total_days):
    for s in [0, 1, 2]:
      a_cnt = sum(work[(n, t, s)] for n in grp_A)
      b_cnt = sum(work[(n, t, s)] for n in grp_B)
      e_cnt = sum(work[(n, t, s)] for n in grp_E)
      de_cnt = sum(work[(n, t, s)] for n in grp_DE)
      shift_total = sum(work[(n, t, s)] for n in range(num_nurses))

      if grp_B:
        model.Add(b_cnt <= 2)
      if grp_E:
        model.Add(e_cnt <= 2)
      if grp_A and grp_B:
        model.Add(a_cnt + b_cnt >= 1)

      if grp_A:
        a_zero = model.NewBoolVar(f"a_zero_{t}_{s}")
        model.Add(a_cnt == 0).OnlyEnforceIf(a_zero)
        model.Add(a_cnt >= 1).OnlyEnforceIf(a_zero.Not())
        penalties.append(a_zero * 300)

        a_over = model.NewIntVar(0, len(grp_A), f"a_over_{t}_{s}")
        model.Add(a_over >= a_cnt - 1)
        penalties.append(a_over * 250)

      if grp_DE:
        if s == 2:
          if strict_de_n_cap:
            model.Add(de_cnt <= 2)
          else:
            model.Add(de_cnt <= 3)
            de_3_n = model.NewBoolVar(f"de3_n_{t}")
            model.Add(de_cnt - 2 <= de_3_n)
            penalties.append(de_3_n * 5000)
        else:
          is_min_staff = model.NewBoolVar(f"is_min_{t}_{s}")
          model.Add(shift_total <= min_de_req).OnlyEnforceIf(is_min_staff)
          model.Add(shift_total >= min_de_req + 1).OnlyEnforceIf(
              is_min_staff.Not()
          )

          model.Add(de_cnt <= 3).OnlyEnforceIf(is_min_staff)

          de_ge_3 = model.NewBoolVar(f"de_ge3_{t}_{s}")
          model.Add(de_cnt >= 3).OnlyEnforceIf(de_ge_3)
          model.Add(de_cnt <= 2).OnlyEnforceIf(de_ge_3.Not())

          bad_min_de3 = model.NewBoolVar(f"bad_min_de3_{t}_{s}")
          model.AddBoolAnd([is_min_staff, de_ge_3]).OnlyEnforceIf(bad_min_de3)
          model.AddBoolOr([is_min_staff.Not(), de_ge_3.Not()]).OnlyEnforceIf(
              bad_min_de3.Not()
          )
          penalties.append(bad_min_de3 * 1200)

      for g_name, members in group_dict.items():
        if len(members) < 2:
          continue
        g_in_s = sum(work[(n, t, s)] for n in members)
        g_over = model.NewIntVar(0, len(members), f"g_over_{g_name}_{t}_{s}")
        model.Add(g_over >= g_in_s - 1)

        if s == 2:
          penalties.append(g_over * 450)
          g_zero_n = model.NewBoolVar(f"g_zero_n_{g_name}_{t}")
          model.Add(g_in_s == 0).OnlyEnforceIf(g_zero_n)
          model.Add(g_in_s >= 1).OnlyEnforceIf(g_zero_n.Not())
          penalties.append(g_zero_n * 250)
        else:
          penalties.append(g_over * 45)
          if len(members) >= 3:
            g_zero = model.NewBoolVar(f"g_zero_{g_name}_{t}_{s}")
            model.Add(g_in_s == 0).OnlyEnforceIf(g_zero)
            model.Add(g_in_s >= 1).OnlyEnforceIf(g_zero.Not())
            penalties.append(g_zero * 30)

  # [규칙 3] 개인별 근무 규칙 (최대 연속 근무일수·나이트 휴식일·3연속 강도 동적 반영)
  consec_window = max_consec_work + 1

  for n in range(num_nurses):
    for t in range(2, total_days - 1):
      model.Add(work[(n, t, 2)] + work[(n, t + 1, 0)] <= 1)  # N -> D 금지
      model.Add(work[(n, t, 2)] + work[(n, t + 1, 1)] <= 1)  # N -> E 금지
      model.Add(work[(n, t, 2)] + work[(n, t + 1, 4)] <= 1)  # N -> 교 절대 금지
      model.Add(work[(n, t, 1)] + work[(n, t + 1, 0)] <= 1)  # E -> D 금지
      model.Add(work[(n, t, 1)] + work[(n, t + 1, 4)] <= 1)  # E -> 교 절대 금지

    for t in range(1, total_days - 2):
      model.Add(
          work[(n, t, 2)] + work[(n, t + 1, 3)] + work[(n, t + 2, 0)] <= 2
      )
      model.Add(
          work[(n, t, 2)] + work[(n, t + 1, 3)] + work[(n, t + 2, 4)] <= 2
      )

    for d in range(num_days):
      lbl = raw_request_labels.get((n, d), "")
      if lbl.startswith("연"):
        t_yeon = d + 3
        model.Add(work[(n, t_yeon - 1, 2)] == 0)

    # 나이트 블록 종료 후 다음 나이트까지 최소 min_n_gap일 비나이트 간격
    for t in range(0, total_days - 1):
      for gap in range(2, min_n_gap + 1):
        if t + gap < total_days:
          model.Add(
              work[(n, t, 2)] - work[(n, t + 1, 2)] + work[(n, t + gap, 2)] <= 1
          )

    # 최대 연속 근무일수 제한 (교육 포함)
    for t in range(0, total_days - max_consec_work):
      model.Add(sum(work[(n, t + k, 3)] for k in range(consec_window)) >= 1)

    # 최대 3연속 나이트 제한
    for t in range(0, total_days - 3):
      model.Add(sum(work[(n, t + k, 2)] for k in range(4)) <= 3)

    if (
        prev_duties[n][0] == 2
        and prev_duties[n][1] == 2
        and prev_duties[n][2] == 2
    ):
      model.Add(work[(n, 3, 3)] == 1)

    if max_n_limits[n] >= 2:
      for t in range(3, total_days - 1):
        model.Add(work[(n, t - 1, 2)] + work[(n, t + 1, 2)] >= work[(n, t, 2)])

      if (
          prev_duties[n][1] != 2
          and prev_duties[n][2] == 2
          and max_n_limits[n] >= 1
      ):
        model.Add(work[(n, 3, 2)] == 1)

    n_starts = [work[(n, 3, 2)]]
    for t in range(4, total_days):
      st_var = model.NewBoolVar(f"n_start_{n}_{t}")
      model.Add(work[(n, t, 2)] - work[(n, t - 1, 2)] <= st_var)
      model.Add(st_var <= work[(n, t, 2)])
      model.Add(st_var <= 1 - work[(n, t - 1, 2)])
      n_starts.append(st_var)
    model.Add(sum(n_starts) <= 3)

    for t in range(1, total_days - 2):
      is_3n = model.NewBoolVar(f"three_n_{n}_{t}")
      model.Add(
          work[(n, t, 2)] + work[(n, t + 1, 2)] + work[(n, t + 2, 2)] - 2
          <= is_3n
      )
      penalties.append(is_3n * three_n_penalty)

    single_last_n = model.NewBoolVar(f"single_last_n_{n}")
    model.Add(
        work[(n, total_days - 1, 2)] - work[(n, total_days - 2, 2)]
        <= single_last_n
    )
    penalties.append(single_last_n * 120)

    for t in range(1, total_days - 3):
      is_noeo = model.NewBoolVar(f"noeo_{n}_{t}")
      model.Add(
          work[(n, t, 2)]
          + work[(n, t + 1, 3)]
          + work[(n, t + 2, 1)]
          + work[(n, t + 3, 3)]
          - 3
          <= is_noeo
      )
      penalties.append(is_noeo * 1500)

  # [규칙 5] 수당일(일요일 제외 공휴일)은 필수(목표) 오프수에서 제외하여 계산
  non_holiday_t_list = [
      d + 3 for d in range(num_days) if (d + 1) not in holiday_set
  ]

  n_counts = []
  n_balance_candidates = []
  for n in range(num_nurses):
    target_off = base_off_days + extra_offs[n]

    c_reg_off = model.NewIntVar(0, num_days, f"reg_off_{n}")
    model.Add(c_reg_off == sum(work[(n, t, 3)] for t in non_holiday_t_list))
    model.Add(c_reg_off >= target_off)

    excess_off = model.NewIntVar(0, num_days, f"excess_off_{n}")
    model.Add(excess_off == c_reg_off - target_off)
    penalties.append(excess_off * 25000)

    c_n = model.NewIntVar(0, num_days, f"n_{n}")
    model.Add(c_n == sum(work[(n, t, 2)] for t in range(3, total_days)))
    n_counts.append(c_n)

    # 나이트 제외 전담(DE/D/E 또는 최대N=0)은 전체 N 편차 계산에서 제외
    is_n_excluded = (max_n_limits[n] == 0) or (
        nurse_prefs[n] is not None
        and nurse_prefs[n][0] == "shifts"
        and 2 not in nurse_prefs[n][1]
    )
    if not is_n_excluded:
      n_balance_candidates.append(c_n)

    if strict_n_cap:
      model.Add(c_n <= max_n_limits[n])
    else:
      over_n = model.NewIntVar(0, num_days, f"over_n_{n}")
      model.Add(over_n >= c_n - max_n_limits[n])
      penalties.append(over_n * 500000)

  # 6. 원티드 최우선 반영
  unmet_requests = {}
  for (n, d), req_s in fixed_requests.items():
    if req_s == 4:
      model.Add(work[(n, d + 3, 4)] == 1)
    else:
      unmet = model.NewBoolVar(f"unmet_{n}_{d}")
      model.Add(work[(n, d + 3, req_s)] == 1).OnlyEnforceIf(unmet.Not())
      model.Add(work[(n, d + 3, req_s)] == 0).OnlyEnforceIf(unmet)
      unmet_requests[(n, d, req_s)] = unmet
      penalties.append(unmet * 100000)

  # [규칙 6-2] 설날·추석 등 3일 이상 연휴 수당 근무 최대 균등 배분
  if len(holiday_list) >= 3:
    holiday_t_list = [d + 2 for d in holiday_list]
    total_hol_slots = len(holiday_list) * (hol_de_req * 2 + daily_n_req)
    avail_by_nurse = []
    for n in range(num_nurses):
      wanted_non_work = sum(
          1
          for d in holiday_list
          if fixed_requests.get((n, d - 1)) in (3, 4)
      )
      avail_by_nurse.append(max(0, len(holiday_list) - wanted_non_work))

    avail_nurses_cnt = sum(1 for a in avail_by_nurse if a >= 1)
    fair_cap = (
        max(2, math.ceil(total_hol_slots / avail_nurses_cnt))
        if avail_nurses_cnt > 0
        else len(holiday_list)
    )
    can_cap_fair = (
        sum(min(fair_cap, a) for a in avail_by_nurse) >= total_hol_slots
    )
    can_guarantee_min1 = (
        avail_nurses_cnt <= total_hol_slots
    )

    for n in range(num_nurses):
      avail_h = avail_by_nurse[n]
      hol_work_cnt = sum(
          work[(n, t, 0)] + work[(n, t, 1)] + work[(n, t, 2)]
          for t in holiday_t_list
      )
      if avail_h >= 1 and can_guarantee_min1:
        model.Add(hol_work_cnt >= 1)

      if can_cap_fair:
        model.Add(hol_work_cnt <= fair_cap)
      else:
        over_hol = model.NewIntVar(0, len(holiday_list), f"over_hol_{n}")
        model.Add(over_hol >= hol_work_cnt - fair_cap)
        penalties.append(over_hol * 5000)

      target_min_hol = min(fair_cap, avail_h)
      if target_min_hol > 0:
        short_hol = model.NewIntVar(0, len(holiday_list), f"short_hol_{n}")
        model.Add(short_hol >= target_min_hol - hol_work_cnt)
        penalties.append(short_hol * 6000)

  # 7. 개인별 D/E 균형 (별도 선호/비율 미지정 인원 대상) 및 퐁당 오프 억제
  for n in range(num_nurses):
    skip_default_de_balance = (nurse_prefs[n] is not None) or (
        (not has_pref_col) and (n == idx_jm)
    )
    if not skip_default_de_balance:
      d_cnt_n = sum(work[(n, t, 0)] for t in range(3, total_days))
      e_cnt_n = sum(work[(n, t, 1)] for t in range(3, total_days))
      diff_de_person = model.NewIntVar(0, num_days, f"p_de_diff_{n}")
      model.AddAbsEquality(diff_de_person, d_cnt_n - e_cnt_n)
      penalties.append(diff_de_person * 12)

    for t in range(2, total_days - 2):
      is_pongdang = model.NewBoolVar(f"pong_{n}_{t}")
      model.Add(
          (1 - work[(n, t, 3)])
          + work[(n, t + 1, 3)]
          + (1 - work[(n, t + 2, 3)])
          - 2
          <= is_pongdang
      )
      penalties.append(is_pongdang * 6)

    for t in range(2, total_days - 1):
      is_de_switch = model.NewBoolVar(f"de_sw_{n}_{t}")
      model.Add(work[(n, t, 0)] + work[(n, t + 1, 1)] - 1 <= is_de_switch)
      penalties.append(is_de_switch * 2)

  # [규칙 4] 안전밸브(8명=15,000) + 7명 쏠림 억제(700) + 바쁜 요일 우선(420) + D=E 동수(350)
  for d in range(num_days):
    t = d + 3
    day_num = d + 1
    wd_idx = calendar.weekday(year, month, day_num)
    is_busy_day = wd_idx in busy_weekdays

    d_cnt = sum(work[(n, t, 0)] for n in range(num_nurses))
    e_cnt = sum(work[(n, t, 1)] for n in range(num_nurses))

    diff_de_day = model.NewIntVar(0, num_nurses, f"d_de_diff_{t}")
    model.AddAbsEquality(diff_de_day, d_cnt - e_cnt)
    penalties.append(diff_de_day * 350)

    d_gt_e = model.NewIntVar(0, num_nurses, f"d_gt_e_{t}")
    model.Add(d_gt_e >= d_cnt - e_cnt)
    penalties.append(d_gt_e * 90)

    if day_num not in holiday_set:
      for s_idx, s_cnt in [(0, d_cnt), (1, e_cnt)]:
        # 선형 슬랙 변수로 표현하여 CP-SAT 탐색 속도 및 수렴 품질 극대화
        if use_valve_cap:
          over_valve = model.NewIntVar(0, 1, f"over_valve_{t}_{s_idx}")
          model.Add(over_valve >= s_cnt - normal_de_cap)
          penalties.append(over_valve * 15000)

        # 일반 상한 (기본 7명 = min_de_req + 2) 쏠림 억제 (벌점 700)
        over_high = model.NewIntVar(0, 2, f"over_high_{t}_{s_idx}")
        model.Add(over_high >= s_cnt - (normal_de_cap - 1))
        penalties.append(over_high * 700)

        # 최소 인원 (기본 5명 = min_de_req) 배정 시 바쁜 요일에 10배 이상 페널티
        under_target = model.NewIntVar(0, 1, f"under_target_{t}_{s_idx}")
        model.Add(under_target >= (min_de_req + 1) - s_cnt)
        if is_busy_day:
          penalties.append(under_target * 420)
        else:
          penalties.append(under_target * 40)

  # 9. 전체 간호사 간 N 개수 편차 최소화
  if len(n_balance_candidates) >= 2:
    max_n = model.NewIntVar(0, num_days, "max_n")
    min_n = model.NewIntVar(0, num_days, "min_n")
    model.AddMaxEquality(max_n, n_balance_candidates)
    model.AddMinEquality(min_n, n_balance_candidates)
    n_spread_term = (max_n - min_n) * 300
  else:
    n_spread_term = 0

  model.Minimize(n_spread_term + sum(penalties))

  print("⏳ 최적의 근무표를 계산 중입니다 (최대 50초 소요)...")
  solver = cp_model.CpSolver()
  solver.parameters.max_time_in_seconds = 50.0
  solver.parameters.num_search_workers = 8
  status = solver.Solve(model)

  if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
    print("❌ 조건을 만족하는 근무표를 찾지 못했습니다.")
    print(
        "   👉 특정 날짜에 원티드 오프가 너무 몰렸거나, 개인별 최대N 합계·[설정] 시트"
        " 인원수·[근무분리_설정]이 충돌하는지 확인해 주세요."
    )
    return

  for (n, d, req_s), unmet_var in unmet_requests.items():
    if solver.Value(unmet_var) == 1:
      print(
          f"⚠️ [원티드 충돌 알림] {nurses[n]}님의 {d+1}일"
          f" '{raw_request_labels.get((n, d), rev_map[req_s])}' 요청이 필수"
          " 규칙과 충돌하여 변경되었습니다."
      )

  # 안전밸브(8명) 작동 여부 콘솔 안내
  valve_days = []
  for d in range(num_days):
    t = d + 3
    d_val = sum(solver.Value(work[(n, t, 0)]) for n in range(num_nurses))
    e_val = sum(solver.Value(work[(n, t, 1)]) for n in range(num_nurses))
    if d_val == valve_de_cap:
      valve_days.append(f"{d+1}일(D {d_val}명)")
    if e_val == valve_de_cap:
      valve_days.append(f"{d+1}일(E {e_val}명)")
  if valve_days:
    print(
        f"ℹ️ [안전밸브 작동 알림] 총 근무량 초과의 오프 쏠림을 막기 위해"
        f" {', '.join(valve_days)}에 {valve_de_cap}명이 배정되었습니다."
    )

  # =========================================================================
  # 결과 엑셀 파일 생성 (진한 회색 원티드 + 자동 색상 조건부 서식 + 자동 합계 수식)
  # =========================================================================
  out_wb = openpyxl.Workbook()
  out_ws = out_wb.active
  out_ws.title = f"{year}년_{month}월_근무표"

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

  fills = {
      "D": PatternFill(
          start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
      ),
      "E": PatternFill(
          start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
      ),
      "N": PatternFill(
          start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
      ),
      "O": PatternFill(
          start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
      ),
      "교": PatternFill(
          start_color="A6A6A6", end_color="A6A6A6", fill_type="solid"
      ),
      "WANTED": PatternFill(
          start_color="A6A6A6", end_color="A6A6A6", fill_type="solid"
      ),
      "PREV": PatternFill(
          start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
      ),
      "GRP": PatternFill(
          start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
      ),
  }
  center = Alignment(horizontal="center", vertical="center")
  bold = Font(bold=True)

  base_headers = ["성명", "그룹", "전월-3", "전월-2", "전월-1"]
  for idx, h_text in enumerate(base_headers, start=1):
    c = out_ws.cell(row=1, column=idx, value=h_text)
    c.font = bold
    c.alignment = center
    c.fill = (
        fills["GRP"]
        if idx == 2
        else (fills["PREV"] if idx >= 3 else PatternFill(fill_type=None))
    )
    c.border = thick_right if idx == 5 else thin_border

  weekdays_kr = ["월", "화", "수", "목", "금", "토", "일"]
  holiday_col_letters = []
  weekend_col_letters = []

  for d in range(num_days):
    col = d + 6
    col_letter = get_column_letter(col)
    day_num = d + 1
    wd_idx = calendar.weekday(year, month, day_num)
    is_hol = day_num in holiday_set
    if is_hol:
      holiday_col_letters.append(col_letter)
    if wd_idx >= 5:
      weekend_col_letters.append(col_letter)

    cell = out_ws.cell(
        row=1, column=col, value=f"{day_num}\n({weekdays_kr[wd_idx]})"
    )
    cell.font = Font(bold=True, color="C00000" if is_hol else "000000")
    cell.alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    cell.border = thin_border
    if wd_idx >= 5 or is_hol:
      cell.fill = PatternFill(
          start_color="F8CBAD", end_color="F8CBAD", fill_type="solid"
      )

  stat_headers = [
      "D",
      "E",
      "N",
      "실제총Off",
      "일반Off(수당일제외)",
      "수당일Off",
      "목표Off(일반)",
      "교육(교)",
      "주말Off",
      "N상한",
      "근무선호",
  ]
  for idx, h in enumerate(stat_headers):
    c = out_ws.cell(row=1, column=num_days + 6 + idx, value=h)
    c.font = bold
    c.alignment = center
    c.border = thin_border
    c.fill = PatternFill(
        start_color="EDEDED", end_color="EDEDED", fill_type="solid"
    )

  first_day_col = get_column_letter(6)
  last_day_col = get_column_letter(5 + num_days)
  tot_off_col_letter = get_column_letter(num_days + 6 + 3)
  hol_off_col_letter = get_column_letter(num_days + 6 + 5)

  cf_segments = []

  for n in range(num_nurses):
    r = n + 2
    c_name = out_ws.cell(row=r, column=1, value=nurses[n])
    c_name.font = bold
    c_name.alignment = center
    c_name.border = thin_border

    c_grp = out_ws.cell(row=r, column=2, value=nurse_groups[n])
    c_grp.alignment = center
    c_grp.fill = fills["GRP"]
    c_grp.border = thin_border

    for p_t in range(3):
      pc = out_ws.cell(
          row=r, column=p_t + 3, value=rev_map[prev_duties[n][p_t]]
      )
      pc.alignment = center
      pc.fill = fills["PREV"]
      pc.font = Font(color="7F7F7F")
      pc.border = thick_right if p_t == 2 else thin_border

    seg_start_col = None

    for d in range(num_days):
      t = d + 3
      col = d + 6
      assigned_s = 3
      for s in range(5):
        if solver.Value(work[(n, t, s)]) == 1:
          assigned_s = s
          break
      assigned_key = rev_map[assigned_s]

      is_req = (n, d) in fixed_requests and fixed_requests[
          (n, d)
      ] == assigned_s
      cell_text = (
          raw_request_labels[(n, d)]
          if is_req and (n, d) in raw_request_labels
          else assigned_key
      )

      c = out_ws.cell(row=r, column=col, value=cell_text)
      c.alignment = center
      c.border = thin_border
      c.fill = fills["WANTED"] if is_req else fills[assigned_key]

      if assigned_s == 3:
        c.font = Font(
            color="9C0006" if is_req else "C00000",
            bold=True,
            underline="single" if is_req else None,
        )
      elif assigned_s == 2:
        c.font = Font(
            color="002060" if is_req else "1F4E78",
            bold=True,
            underline="single" if is_req else None,
        )
      elif assigned_s == 4:
        c.font = Font(color="1F4E23", bold=True, underline="single")
      elif is_req:
        c.font = Font(color="000000", bold=True, underline="single")

      if not is_req:
        if seg_start_col is None:
          seg_start_col = col
      else:
        if seg_start_col is not None:
          s_let = get_column_letter(seg_start_col)
          e_let = get_column_letter(col - 1)
          cf_segments.append((f"{s_let}{r}:{e_let}{r}", f"{s_let}{r}"))
          seg_start_col = None

    if seg_start_col is not None:
      s_let = get_column_letter(seg_start_col)
      e_let = get_column_letter(5 + num_days)
      cf_segments.append((f"{s_let}{r}:{e_let}{r}", f"{s_let}{r}"))

    row_range = f"{first_day_col}{r}:{last_day_col}{r}"
    target_off = base_off_days + extra_offs[n]

    stat_formulas = [
        f'=COUNTIF({row_range},"D")',
        f'=COUNTIF({row_range},"E")',
        f'=COUNTIF({row_range},"N")',
        f'=COUNTIF({row_range},"O")+COUNTIF({row_range},"OFF")+COUNTIF({row_range},"연*")',
        f"={tot_off_col_letter}{r}-{hol_off_col_letter}{r}",
        make_off_count_formula_for_cells(holiday_col_letters, r),
        target_off,
        f'=COUNTIF({row_range},"교*")',
        make_off_count_formula_for_cells(weekend_col_letters, r),
        max_n_limits[n],
        nurse_prefs_raw[n],
    ]
    for idx, val in enumerate(stat_formulas):
      sc = out_ws.cell(row=r, column=num_days + 6 + idx, value=val)
      sc.alignment = center
      sc.border = thin_border
      if idx in (2, 3, 4):
        sc.font = bold

  cf_font_d = Font(color="000000", bold=False)
  cf_font_e = Font(color="000000", bold=False)
  cf_font_n = Font(color="1F4E78", bold=True)
  cf_font_o = Font(color="C00000", bold=True)
  cf_font_kyo = Font(color="1F4E23", bold=True)

  for rng_str, top_left in cf_segments:
    out_ws.conditional_formatting.add(
        rng_str,
        CellIsRule(
            operator="equal",
            formula=['"D"'],
            fill=fills["D"],
            font=cf_font_d,
        ),
    )
    out_ws.conditional_formatting.add(
        rng_str,
        CellIsRule(
            operator="equal",
            formula=['"E"'],
            fill=fills["E"],
            font=cf_font_e,
        ),
    )
    out_ws.conditional_formatting.add(
        rng_str,
        CellIsRule(
            operator="equal",
            formula=['"N"'],
            fill=fills["N"],
            font=cf_font_n,
        ),
    )
    out_ws.conditional_formatting.add(
        rng_str,
        FormulaRule(
            formula=[
                f'OR(UPPER({top_left})="O",UPPER({top_left})="OFF",LEFT({top_left},1)="연")'
            ],
            fill=fills["O"],
            font=cf_font_o,
        ),
    )
    out_ws.conditional_formatting.add(
        rng_str,
        FormulaRule(
            formula=[f'LEFT({top_left},1)="교"'],
            fill=fills["교"],
            font=cf_font_kyo,
        ),
    )

  last_nurse_row = num_nurses + 1
  summary_start = num_nurses + 3
  summary_specs = [
      ("D 인원", lambda rng: f'=COUNTIF({rng},"D")'),
      ("E 인원", lambda rng: f'=COUNTIF({rng},"E")'),
      ("N 인원", lambda rng: f'=COUNTIF({rng},"N")'),
      (
          "O(연차포함)",
          lambda rng: f'=COUNTIF({rng},"O")+COUNTIF({rng},"OFF")+COUNTIF({rng},"연*")',
      ),
      ("교육(교)", lambda rng: f'=COUNTIF({rng},"교*")'),
  ]

  for idx, (label, formula_fn) in enumerate(summary_specs):
    r = summary_start + idx
    lbl = out_ws.cell(row=r, column=1, value=label)
    lbl.font = bold
    lbl.alignment = center
    lbl.border = thin_border

    for d in range(num_days):
      col = d + 6
      col_letter = get_column_letter(col)
      col_range = f"{col_letter}2:{col_letter}{last_nurse_row}"
      c = out_ws.cell(row=r, column=col, value=formula_fn(col_range))
      c.alignment = center
      c.border = thin_border

  out_ws.column_dimensions["A"].width = 13
  out_ws.column_dimensions["B"].width = 7
  for col_idx in range(3, 6):
    out_ws.column_dimensions[get_column_letter(col_idx)].width = 6.5
  for col_idx in range(6, num_days + 6):
    out_ws.column_dimensions[get_column_letter(col_idx)].width = 4.8
  for col_idx in range(num_days + 6, num_days + 17):
    out_ws.column_dimensions[get_column_letter(col_idx)].width = 12.5
  out_ws.row_dimensions[1].height = 28
  out_ws.freeze_panes = "F2"

  out_filename = f"duty_result_{year}_{month:02d}.xlsx"
  try:
    out_wb.save(out_filename)
  except PermissionError:
    ts = datetime.datetime.now().strftime("%H%M%S")
    out_filename = f"duty_result_{year}_{month:02d}_{ts}.xlsx"
    out_wb.save(out_filename)
    print(
        f"⚠️ 기존 결과 파일이 엑셀에 열려 있어 새 이름('{out_filename}')으로"
        " 저장했습니다!"
    )

  print("=" * 68)
  print(f"✅ [완료] '{out_filename}' 근무표가 성공적으로 생성되었습니다!")
  print("   • 수기 수정 시 우측/하단 숫자와 일반 근무 색상이 자동으로 바뀝니다.")
  print("   • 다음 달 양식 생성 시 이 파일의 마지막 3일 근무가 자동 이월됩니다.")
  print("=" * 68)


if __name__ == "__main__":
  solve_duty_final()
