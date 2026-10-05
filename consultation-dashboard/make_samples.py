# -*- coding: utf-8 -*-
"""테스트 샘플 엑셀 3종 + answer_key.json 생성기.
정답지의 핵심 숫자(환자수/상담건수)는 '설계한 값'을 손으로 적은 것입니다(코드로 계산한 값 아님)."""
import json
from datetime import datetime, date
from openpyxl import Workbook

OUT = "test-data/"
EPOCH = date(1899, 12, 30)
serial = lambda d: (d - EPOCH).days          # 엑셀 시리얼 숫자
D = lambda s: datetime.strptime(s, "%Y-%m-%d %H:%M")

# ---------- 샘플1 : 정상 ----------
# (일시, 환자번호, 환자명, 담당자)
rows1 = [
    ("2025-12-20 10:00", "P0099", "백지원", "김상담"),   # T9 전년도 -> 제외
    ("2025-12-29 11:00", "P0100", "한서준", "김상담"),   # T9 전년도 -> 제외
    ("2026-01-05 09:30", "P001", "이민지", "박상담"),    # T4 Week 0
    ("2026-09-24 10:00", "P002", "최유나", "김상담"),    # T2 같은 날 같은 환자 2회
    ("2026-09-24 15:30", "P002", "최유나", "김상담"),
    ("2026-09-30 09:00", "P003", "정하늘", "박상담"),
    ("2026-10-02 10:00", "P003", "정하늘", "박상담"),
    ("2026-10-02 14:00", "P004", "강도윤", "김상담"),
    ("2026-10-05 13:00", "P005", "윤서아", "박상담"),
    ("2026-10-06 16:00", "P004", "강도윤", "김상담"),    # T3 화요일(이번주 마지막 날)
    ("2026-10-07 09:00", "P004", "강도윤", "김상담"),    # T3 수요일(다음 주 첫날)
    ("2026-10-08 11:00", "P007", "임재현", "박상담"),
    ("2026-10-15 10:00", "P008", "오시우", "김상담"),
    ("2026-10-16 10:30", "P008", "오시우", "김상담"),
]
def ex(this, prev, ytd, weeks, loaded, after, prevyear, nopno, summ=0, nodate=0, derr=0):
    counted = loaded - after - prevyear - summ - nodate - derr
    return {"this": {"patients": this[0], "consults": this[1]},
            "prev": {"patients": prev[0], "consults": prev[1]},
            "ytd": {"patients": ytd[0], "consults": ytd[1]},
            "weeks_nonzero": weeks,   # {"주차": [환자수, 상담건수]}
            "counts": {"loaded": loaded, "sum": summ, "no_date": nodate, "date_error": derr,
                       "prev_year": prevyear, "after": after, "counted": counted, "no_patient_no": nopno}}
# 주차: 2026 첫 수요일=01-07 -> 01-07~13이 Week1. 09-24는 Week 38, 09-30은 Week 39, 10-07은 Week 40, 10-14는 Week 41
S1 = [
 {"meeting": "2026-10-07", "note": "T1 기획 예시(+T2,T4,T9)", "expect": dict(ex((3,5),(1,2),(5,8),{"0":[1,1],"38":[1,2],"39":[3,5]},14,4,2,0), staff={"김상담":{"this":[1,2],"prev":[1,2],"ytd":[2,4]},"박상담":{"this":[2,3],"prev":[0,0],"ytd":[3,4]}}, months={"1":[1,1],"9":[2,3],"10":[3,4]})},
 {"meeting": "2026-10-14", "note": "T3 화/수 경계: 10-06과 10-07은 다른 주", "expect": dict(ex((2,2),(3,5),(7,10),{"0":[1,1],"38":[1,2],"39":[3,5],"40":[2,2]},14,2,2,0), months={"1":[1,1],"9":[2,3],"10":[4,6]})},
 {"meeting": "2026-10-21", "expect": ex((1,2),(2,2),(8,12),{"0":[1,1],"38":[1,2],"39":[3,5],"40":[2,2],"41":[1,2]},14,0,2,0)},
 {"meeting": "2026-01-14", "note": "T4 Week 0가 전주, 12월 제외", "expect": ex((0,0),(1,1),(1,1),{"0":[1,1]},14,11,2,0)},
 {"meeting": "2026-01-07", "note": "T4 이번주=12/31~1/6(Week 0)", "expect": ex((1,1),(0,0),(1,1),{"0":[1,1]},14,11,2,0)},
]

wb = Workbook(); ws = wb.active; ws.title = "상담내역"
ws.append(["상담일자", "환자번호", "환자명", "담당자"])
for dt, no, nm, st in rows1:
    ws.append([D(dt), no, nm, st])
for r in range(2, ws.max_row + 1): ws.cell(r, 1).number_format = "yyyy-mm-dd hh:mm"
for c, w in zip("ABCD", (20, 12, 12, 12)): ws.column_dimensions[c].width = w
wb.save(OUT + "sample1_normal.xlsx")

# ---------- 샘플2 : 불규칙 (1행 제목, 3행 헤더, 텍스트/숫자/날짜 혼합, 시트 2개) ----------
wb = Workbook()
memo = wb.active; memo.title = "안내"
memo["A1"] = "이 시트는 메모입니다. 상담 데이터는 두 번째 시트에 있습니다."
ws = wb.create_sheet("2026 상담현황")
ws["A1"] = "2026년 병원 상담 통계 (내부용)"
ws["A2"] = None
ws.append(["No", "상담 일자", "환자번호", "환자명", "담당자", "비고"])   # 3행이 헤더
fmt = [
    lambda d: d.strftime("%Y.%m.%d"),                 # 텍스트 2026.10.02
    lambda d: d.strftime("%Y-%m-%d %H:%M"),           # 텍스트 2026-10-02 14:30
    lambda d: d,                                      # 진짜 날짜
    lambda d: d.strftime("%Y/%m/%d"),                 # 텍스트 2026/10/02
]
for i, (dt, no, nm, st) in enumerate(rows1):
    d = D(dt)
    pno = int("".join(ch for ch in no if ch.isdigit())) if i % 2 else no + " "   # 숫자/뒤공백 혼합
    # 같은 환자는 같은 번호여야 하므로 번호 규칙을 환자별로 고정
    pno = no if i % 3 else no + " "
    ws.append([i + 1, fmt[i % 4](d), pno, nm, st, "메모" if i % 5 == 0 else None])
ws.cell(4, 2).number_format = "yyyy-mm-dd"
wb.save(OUT + "sample2_irregular.xlsx")
S2 = S1   # 같은 내용이므로 정답도 동일

# ---------- 샘플4 : 자동 감지 실패용(영문 제목). 컬럼 매핑 '기억' 테스트에만 사용, 내용은 샘플1과 같음 ----------
wb = Workbook(); ws = wb.active; ws.title = "log"
ws.append(["Day", "Pt No", "Pt Name", "Staff"])
for dt, no, nm, st in rows1: ws.append([D(dt), no, nm, st])
wb.save(OUT + "sample4_custom_headers.xlsx")

# ---------- 샘플3 : 엣지 (시리얼 날짜, 합계/소계 행, 환자번호 컬럼 없음, 날짜 오류/없음) ----------
wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
ws.append(["상담일자", "환자명", "담당자"])
S = lambda s: serial(datetime.strptime(s, "%Y-%m-%d").date())
data3 = [
    (S("2026-01-02"), "김하나", "A"),          # Week 0
    (S("2026-09-25"), "이둘", "A"),
    (S("2026-09-25"), "이둘", "A"),            # 같은 날 같은 환자
    ("2026.10.01", "이둘", "B"),               # 텍스트 날짜
    ("2026-10-02 14:30", "박셋", "B"),
    (S("2026-10-04"), "최넷", "A"),
    ("날짜미정", "정다섯", "A"),               # 날짜 오류
    (None, "윤여섯", "A"),                     # 날짜 없음
    (S("2025-12-31"), "오일곱", "B"),          # 전년도
    ("합계", None, None),                      # 합계 행
    ("소계", 6, None),                         # 소계 행
    (S("2026-10-08"), "한여덟", "A"),          # 10-07 회의 기준으로는 '이후'
]
for r in data3: ws.append(list(r))
for r in range(2, ws.max_row + 1):
    c = ws.cell(r, 1)
    if isinstance(c.value, int): c.number_format = "General"   # 시리얼 숫자 그대로 보이게
wb.save(OUT + "sample3_edge.xlsx")
S3 = [
 {"meeting": "2026-10-07", "note": "T6 시리얼/T7 합계행/T8 이름기준+경고/T9", "expect": dict(ex((3,3),(1,2),(5,6),{"0":[1,1],"38":[1,2],"39":[3,3]},12,1,1,6,summ=2,nodate=1,derr=1), staff={"A":{"this":[1,1],"prev":[1,2],"ytd":[3,4]},"B":{"this":[2,2],"prev":[0,0],"ytd":[2,2]}}, months={"1":[1,1],"9":[1,2],"10":[3,3]})},
 {"meeting": "2026-10-14", "expect": ex((1,1),(3,3),(6,7),{"0":[1,1],"38":[1,2],"39":[3,3],"40":[1,1]},12,0,1,7,summ=2,nodate=1,derr=1)},
]
key = {"_설명": "week=이번주, prev=전주, ytd=1/1~기준주 끝. weeks_nonzero의 주차 0=첫 수요일 이전(Week 0).",
       "samples": {"sample1_normal.xlsx": {"scenarios": S1},
                   "sample2_irregular.xlsx": {"scenarios": S2},
                   "sample3_edge.xlsx": {"scenarios": S3}}}
json.dump(key, open(OUT + "answer_key.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("완료")
