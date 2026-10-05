# -*- coding: utf-8 -*-
"""시연용 '실제처럼 보이는' 가상 상담일지 생성기 (모든 이름·번호는 무작위로 만든 가짜 데이터).
사용법: python make_demo.py  →  test-data/demo_hospital_2026.xlsx
실무 엑셀에 흔한 지저분함을 일부러 섞음: 1행 제목/3행 헤더, 월별 소계·맨 아래 합계, 텍스트 날짜 혼합,
환자번호 누락, 날짜 오류, 번호-이름 오타, 전년도(2025-12) 행, 같은 날 같은 환자 2회."""
import random
from datetime import date, datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

random.seed(2026)
SUR = list("김이박최정강조윤장임한오서신권황안송류홍")
GIV = "민준 서연 도윤 지우 하준 서윤 시우 지민 예준 수빈 주원 하윤 지호 채원 건우 유진 현우 소율 준서 다은 은우 아린 승현 나연 재윤 보람 태양 세아".split()
STAFF = [("김지수", 0.34), ("이도현", 0.28), ("박서연", 0.22), ("최민준", 0.16)]
TYPES = [("재진", 0.46), ("수납상담", 0.18), ("보험상담", 0.14), ("전화상담", 0.12)]   # 초진은 환자의 첫 방문에 자동 부여
HOLIDAYS = {date(2026, *md) for md in [(1, 1), (2, 16), (2, 17), (2, 18), (3, 2), (5, 5), (5, 25), (6, 3), (8, 17), (9, 24), (9, 25), (9, 26), (10, 5)]}
pick = lambda pairs: random.choices([p[0] for p in pairs], [p[1] for p in pairs])[0]

# 환자 풀(번호, 이름) — 이름이 겹치는 동명이인도 일부 생김
pool = []
for i in range(420):
    pool.append(("%d%05d" % (random.choice([2019, 2021, 2023, 2025, 2026]), 100 + i * 7), random.choice(SUR) + random.choice(GIV)))
seen = set(); rows = []         # (datetime, 번호, 이름, 담당, 유형, 비고)
def add(d, hh, mm, pt, ty=None, note=None):
    first = pt[0] not in seen; seen.add(pt[0])
    rows.append([datetime(d.year, d.month, d.day, hh, mm), pt[0], pt[1], pick(STAFF), ty or ("초진" if first else pick(TYPES)), note])

# 전년도 12월 일부 (이전 파일에서 딸려온 행 → 제외되어야 함)
for k in range(12):
    add(date(2025, 12, 15 + k % 14), 9 + k % 8, 0, random.choice(pool))
seen.clear()
day = date(2026, 1, 1); last = date(2026, 10, 2)
while day <= last:
    if day.weekday() < 6 and day not in HOLIDAYS:
        n = max(0, int(random.gauss(7.5 if day.weekday() < 5 else 3.5, 2.2)))
        if day.month in (7, 8): n = int(n * 0.8)            # 여름 비수기
        for _ in range(n):
            pt = random.choice(pool[:int(60 + 360 * (day.timetuple().tm_yday / 275))])   # 시간이 갈수록 새 환자 증가
            add(day, random.randint(9, 17), random.choice([0, 10, 20, 30, 40, 50]), pt)
            if random.random() < 0.04: add(day, random.randint(9, 17), 5, pt)      # 같은 날 같은 환자 2회
    day += timedelta(1)
rows.sort(key=lambda r: r[0])
for r in rows[12:]:
    if r[0].year == 2026 and random.random() < 0.012: r[1] = ""                    # 환자번호 누락(이름만)
bad = random.sample(range(12, len(rows)), 2)                                       # 번호-이름 오타 2건
for i in bad: rows[i][2] = rows[i][2] + "님"          # 이름 뒤에 "님"이 붙은 입력 오타

wb = Workbook(); ws = wb.active; ws.title = "2026 상담일지"
ws["A1"] = "2026년 원무팀 상담일지 (내부용)"; ws["A1"].font = Font(size=14, bold=True)
ws["A2"] = "※ 월별 소계는 수기 입력 / 담당: 원무팀"
head = ["No", "상담일자", "환자번호", "환자명", "담당자", "상담유형", "비고"]
ws.append([]); ws.delete_rows(3); ws.append(head)                                  # 3행이 헤더
for c in ws[3]: c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="DDE6F0"); c.alignment = Alignment(horizontal="center")
no = 0; cur = None; cnt = 0
def subtotal():
    ws.append(["", "소계", "", "", "", cnt, ""])
text_idx = set(random.sample(range(len(rows)), 140)); err_idx = set(random.sample(range(12, len(rows)), 4))
for i, (dt, pno, nm, st, ty, note) in enumerate(rows):
    if cur is not None and (dt.year, dt.month) != cur and dt.year == 2026 and cur[0] == 2026:
        subtotal(); cnt = 0
    cur = (dt.year, dt.month) if dt.year == 2026 else (2025, 12)
    no += 1; cnt += 1
    if i in err_idx: val = "미정"
    elif i in text_idx: val = random.choice([dt.strftime("%Y.%m.%d"), dt.strftime("%Y-%m-%d %H:%M"), dt.strftime("%Y/%m/%d")])
    else: val = dt
    ws.append([no, val, pno or None, nm, st, ty, note])
subtotal(); ws.append(["", "합계", "", "", "", len(rows), ""])
for r in range(4, ws.max_row + 1):
    if isinstance(ws.cell(r, 2).value, datetime): ws.cell(r, 2).number_format = "yyyy-mm-dd hh:mm"
for c, w in zip("ABCDEFG", (6, 19, 12, 10, 10, 10, 14)): ws.column_dimensions[c].width = w
wb.save("test-data/demo_hospital_2026.xlsx")
print("행 수:", len(rows), "→ test-data/demo_hospital_2026.xlsx")
