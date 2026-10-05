# -*- coding: utf-8 -*-
"""독립 검증 스크립트.
 ① pandas로 '처음부터' 같은 규칙을 다시 구현해 answer_key.json과 비교  (dashboard.html의 JS 코드를 보지 않음)
 ② dashboard.html을 헤드리스 브라우저(인터넷 차단 상태)로 열어 같은 샘플을 넣고 결과를 answer_key와 비교
    (②는 playwright가 있을 때만: pip install playwright)
사용법:  python verify.py"""
import json, sys, base64, pathlib
from datetime import date, datetime, timedelta
import pandas as pd

ROOT = pathlib.Path(__file__).parent
KEY = json.load(open(ROOT / "test-data" / "answer_key.json", encoding="utf-8"))
EPOCH = pd.Timestamp(1899, 12, 30)

# ---------------- ① pandas 독립 재계산 ----------------
def to_date(v):
    """셀 하나를 date로. 실패하면 'empty' 또는 'error'."""
    if v is None or (isinstance(v, float) and pd.isna(v)) or (isinstance(v, str) and not v.strip()):
        return "empty"
    if isinstance(v, (datetime, pd.Timestamp)): return v.date()
    if isinstance(v, (int, float)):
        return (EPOCH + pd.Timedelta(days=int(v))).date()   # 엑셀 시리얼
    import re
    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", str(v))
    if m:
        try: return date(int(m[1]), int(m[2]), int(m[3]))
        except ValueError: pass
    return "error"

def load(path):
    """헤더 행을 찾고(상위 20행) 데이터프레임으로 만든다. 시트가 여럿이면 헤더가 있는 시트."""
    for sheet, raw in pd.read_excel(path, sheet_name=None, header=None, dtype=object).items():
        for i in range(min(20, len(raw))):
            cells = [str(x).replace(" ", "") for x in raw.iloc[i].tolist() if pd.notna(x)]
            if any(("상담일" in c or c == "일자") for c in cells) and any("환자" in c for c in cells):
                df = raw.iloc[i + 1:].copy(); df.columns = [str(x).replace(" ", "") if pd.notna(x) else f"c{j}" for j, x in enumerate(raw.iloc[i])]
                return df.dropna(how="all")
    raise RuntimeError("헤더를 찾지 못함: " + str(path))

def compute(path, meeting):
    df = load(path)
    datecol = next(c for c in df.columns if "상담일" in c)
    nocol = next((c for c in df.columns if "환자번호" in c), None)
    namecol = next((c for c in df.columns if c == "환자명"), None)
    staffcol = next((c for c in df.columns if "담당" in c), None)
    W = datetime.strptime(meeting, "%Y-%m-%d").date()
    assert W.weekday() == 2, "회의일은 수요일이어야 함"
    this_s, this_e = W - timedelta(7), W - timedelta(1)
    prev_s, prev_e = W - timedelta(14), W - timedelta(8)
    Y = this_e.year
    jan1 = date(Y, 1, 1)
    first_wed = jan1 + timedelta((2 - jan1.weekday()) % 7)           # 파이썬: 월=0, 수=2
    week_of = lambda d: 0 if d < first_wed else (d - first_wed).days // 7 + 1

    cnt = dict(loaded=len(df), sum=0, no_date=0, date_error=0, prev_year=0, after=0, counted=0, no_patient_no=0)
    recs = []
    for _, r in df.iterrows():
        cells = [str(x).replace(" ", "") for x in r.tolist() if pd.notna(x)]
        if any(c in ("합계", "소계", "총계") for c in cells): cnt["sum"] += 1; continue
        d = to_date(r[datecol])
        if d == "empty": cnt["no_date"] += 1; continue
        if d == "error": cnt["date_error"] += 1; continue
        if d < jan1: cnt["prev_year"] += 1; continue
        if d > this_e: cnt["after"] += 1; continue
        cnt["counted"] += 1
        no = str(r[nocol]).strip() if nocol and pd.notna(r[nocol]) and str(r[nocol]).strip() else ""
        nm = str(r[namecol]).strip() if namecol and pd.notna(r[namecol]) else ""
        if not no: cnt["no_patient_no"] += 1
        recs.append(dict(week=week_of(d), pid=("N:" + no) if no else "M:" + nm, d=d,
                         staff=(str(r[staffcol]).strip() if staffcol and pd.notna(r[staffcol]) and str(r[staffcol]).strip() else "(미지정)")))
    t = pd.DataFrame(recs, columns=["week", "pid", "d", "staff"])
    per = t.groupby("week").agg(patients=("pid", "nunique"), consults=("pid", "size"))   # 주별 중복제거 / 행 수
    kt = week_of(this_e)
    get = lambda k: (int(per.loc[k, "patients"]), int(per.loc[k, "consults"])) if k in per.index else (0, 0)
    ytd_p = int(per.loc[per.index <= kt, "patients"].sum()); ytd_c = int(per.loc[per.index <= kt, "consults"].sum())
    staff = {}
    for sname, g in t.groupby("staff"):                                   # 직원별: 같은 규칙을 담당자 단위로
        sp = g.groupby("week").agg(p=("pid", "nunique"), c=("pid", "size"))
        sg = lambda k: [int(sp.loc[k, "p"]), int(sp.loc[k, "c"])] if k in sp.index else [0, 0]
        staff[sname] = {"this": sg(kt), "prev": sg(kt - 1),
                        "ytd": [int(sp.loc[sp.index <= kt, "p"].sum()), int(sp.loc[sp.index <= kt, "c"].sum())]}
    return {"staff": staff, "this": dict(zip(("patients", "consults"), get(kt))), "prev": dict(zip(("patients", "consults"), get(kt - 1))),
            "ytd": {"patients": ytd_p, "consults": ytd_c},
            "weeks_nonzero": {str(k): [int(v.patients), int(v.consults)] for k, v in per.iterrows()}, "counts": cnt}

# ---------------- 비교/출력 ----------------
fails = 0; total = 0
def check(label, got, want):
    global fails, total
    total += 1
    ok = got == want
    if not ok: fails += 1
    print(("  PASS " if ok else "  FAIL ") + label + ("" if ok else f"\n        기대={want}\n        실제={got}"))

def compare(tag, got, exp):
    for k in ("this", "prev", "ytd"): check(f"{tag} {k}", got[k], exp[k])
    check(f"{tag} 주별표(0 아닌 주)", got["weeks_nonzero"], exp["weeks_nonzero"])
    check(f"{tag} 품질 카운트", got["counts"], exp["counts"])
    if "staff" in exp: check(f"{tag} 직원별", got["staff"], exp["staff"])

print("=== ① pandas 독립 재계산  vs  answer_key ===")
for f, spec in KEY["samples"].items():
    print(f"[{f}]")
    for sc in spec["scenarios"]:
        compare(f"회의일 {sc['meeting']}", compute(ROOT / "test-data" / f, sc["meeting"]), sc["expect"])

print("\n=== ② dashboard.html (헤드리스 브라우저, 인터넷 차단)  vs  answer_key ===")
try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("  SKIP playwright 미설치 (pip install playwright) — ②는 건너뜀"); sync_playwright = None
if sync_playwright:
    exe = next(iter(sorted(pathlib.Path("/opt/pw-browsers").glob("chromium-*/chrome-linux/chrome"))), None)
    with sync_playwright() as p:
        b = p.chromium.launch(**({"executable_path": str(exe)} if exe else {}))
        ctx = b.new_context(); ctx.set_offline(True)            # T10: 인터넷 차단
        reqs = []; ctx.on("request", lambda r: reqs.append(r.url) if not r.url.startswith(("file:", "data:", "blob:")) else None)
        page = ctx.new_page(); errs = []; page.on("pageerror", lambda e: errs.append(str(e)))
        page.goto((ROOT / "dashboard.html").resolve().as_uri())
        check("T10 외부 요청 0건", reqs, []); check("페이지 JS 오류 없음", errs, [])
        check("내장 자가진단 전부 통과", page.inner_text("#foot").split("통과")[0].strip().replace("내장 자가진단 ", "").split("/")[0] ==
              page.inner_text("#foot").split("통과")[0].strip().split("/")[1], True)
        for f, spec in KEY["samples"].items():
            print(f"[{f}]")
            b64 = base64.b64encode((ROOT / "test-data" / f).read_bytes()).decode()
            for sc in spec["scenarios"]:
                got = page.evaluate("""([b64, m]) => { const bin = atob(b64); const u = new Uint8Array(bin.length);
                    for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
                    return window.DashboardAPI.analyzeWorkbook(u.buffer, m); }""", [b64, sc["meeting"]])
                compare(f"회의일 {sc['meeting']}", got, sc["expect"])
        # 실제 화면(업로드 → KPI 숫자) 확인
        page.set_input_files("#fileInput", str(ROOT / "test-data" / "sample1_normal.xlsx"))
        page.fill("#meetingDate", "2026-10-07"); page.dispatch_event("#meetingDate", "change")
        nums = page.eval_on_selector_all(".kpi .num", "els => els.map(e => parseInt(e.textContent))")
        check("화면 KPI 4개 (환자3·상담5·YTD 5·8)", nums, [3, 5, 5, 8])
        check("화면 주간 표 행 존재", page.locator("#weekTable tr.cur").count(), 1)
        b.close()

print(f"\n결과: {total - fails}/{total} PASS" + ("  ✅ 전부 통과" if fails == 0 else f"  ❌ {fails}개 FAIL"))
sys.exit(1 if fails else 0)
