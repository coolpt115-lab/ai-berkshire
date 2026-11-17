#!/usr/bin/env python3
"""
krstock_screen.py — KOSPI 전 종목 불량기업 걸러내기 (skills/quality-screen.md 기준의 1단계 자동화)

데이터
  - 종목 목록·업종·상장일: KRX KIND 상장법인목록 (키 불필요)
  - 재무제표 10년·주식 수·최대주주 지분·회사분할: OpenDART API (키 필요)

API 키
  환경변수 DART_API_KEY, 없으면 ~/.dart_api_key 파일(키 한 줄)을 읽는다. 키를 코드나 저장소에 적지 않는다.

사용법
  python tools/krstock_screen.py fetch  [--limit N] [--codes 005930,000660]   # DART 자료 내려받기(캐시)
  python tools/krstock_screen.py screen [--limit N] [--codes ...] [--out 경로] # 캐시로 판정, CSV·MD 출력
  python tools/krstock_screen.py run    [...]                                 # fetch + screen

한계 (결과 해석 시 반드시 감안)
  - DART 표준계정 기준 근사치다. 일회성 손익·리스·영업외 항목은 다듬지 않는다.
  - 감가상각비가 현금흐름표 본문에 없는 회사가 많다 → 그 경우 순부채/EBITDA는 '미확인'으로 두고 이자보상배율만 본다.
  - 대주주 매도·주식담보·내부거래 플래그와 예외 C(사업모델 판단)는 자동화하지 않는다 → /investment-checklist 단계에서 확인.
  - 인수합병에 따른 주식 수 증가를 구분하지 못한다 → 7번 탈락은 ⚠️로 표시해 수동 확인.
"""

import argparse
import csv
import io
import json
import os
import re
import statistics
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from datetime import date, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE = os.path.join(ROOT, "data", "kr_dart_cache")
DART = "https://opendart.fss.or.kr/api/"
KIND_URL = ("https://kind.krx.co.kr/corpgeneral/corpList.do"
            "?method=download&searchType=13&marketType=stockMkt")

LAST_FY = 2025            # 최근 사업연도 (사업보고서는 이듬해 3월 제출)
FIN_YEARS = 10            # 2016~2025
SLEEP = 0.12              # DART 호출 간격(초)

FIN_INDUSTRIES = ("은행", "보험", "금융 지원 서비스업", "신탁업", "집합투자", "저축기관", "기타 금융업")
CYCLICAL_KEYS = ("반도체 제조업", "기초 화학물질", "합성고무", "화학섬유", "1차 철강", "1차 비철금속",
                 "석유 정제", "해상 운송", "선박", "시멘트", "펄프", "항공", "건물 건설업", "토목 건설업")
NEUTRAL_SHARE_EVENTS = ("주식분할", "무상", "주식배당", "액면병합", "주식병합")
DILUTIVE_SHARE_EVENTS = ("유상", "전환권", "신주인수권", "주식매수선택권", "교환")


# ─────────────────────────── 공통 유틸 ───────────────────────────

def api_key():
    k = os.environ.get("DART_API_KEY", "").strip()
    if not k:
        p = os.path.join(os.path.expanduser("~"), ".dart_api_key")
        if os.path.exists(p):
            k = open(p, encoding="utf-8-sig").read().strip()
    if not k:
        sys.exit("DART API 키가 없습니다. 환경변수 DART_API_KEY 또는 ~/.dart_api_key 파일에 키를 넣어 주세요.")
    return k


def http_get(url, timeout=30, retries=3):
    for i in range(retries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            return urllib.request.urlopen(req, timeout=timeout).read()
        except (urllib.error.URLError, TimeoutError) as e:
            if i == retries - 1:
                raise
            time.sleep(2 * (i + 1))


class DartLimit(Exception):
    pass


def dart(endpoint, **params):
    params["crtfc_key"] = api_key()
    url = DART + endpoint + "?" + urllib.parse.urlencode(params)
    time.sleep(SLEEP)
    data = json.loads(http_get(url).decode("utf-8"))
    st = data.get("status")
    if st == "000":
        return data.get("list", [])
    if st == "013":            # 조회된 데이터 없음
        return []
    if st == "020":
        raise DartLimit("DART 일일 호출 한도 초과(020). 내일 다시 실행하면 캐시된 회사는 건너뜁니다.")
    if st in ("010", "011", "012", "901"):
        sys.exit(f"DART 키 오류({st}): {data.get('message')}")
    raise RuntimeError(f"DART {endpoint} 오류 {st}: {data.get('message')}")


def num(s):
    if s is None:
        return None
    s = str(s).replace(",", "").strip()
    if s in ("", "-"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def cache_path(*parts):
    p = os.path.join(CACHE, *parts)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    return p


def load_json(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def save_json(p, obj):
    tmp = p + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False)
    os.replace(tmp, p)


# ─────────────────────────── 종목 목록 ───────────────────────────

def kospi_list(refresh=False):
    p = cache_path("kind_kospi.json")
    if os.path.exists(p) and not refresh:
        return load_json(p)
    html = http_get(KIND_URL).decode("cp949", errors="replace")
    rows = [[re.sub(r"<.*?>", "", c).strip() for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", r, re.S)]
            for r in re.findall(r"<tr>(.*?)</tr>", html, re.S)]
    head = rows[0]
    out = [dict(zip(head, r)) for r in rows[1:] if len(r) == len(head)]
    save_json(p, out)
    return out


def corp_codes():
    p = cache_path("corpcode.json")
    if os.path.exists(p):
        return load_json(p)
    raw = http_get(DART + "corpCode.xml?crtfc_key=" + api_key(), timeout=60)
    xml = zipfile.ZipFile(io.BytesIO(raw)).read("CORPCODE.xml").decode("utf-8")
    m = {}
    for blk in re.findall(r"<list>(.*?)</list>", xml, re.S):
        sc = re.search(r"<stock_code>(.*?)</stock_code>", blk).group(1).strip()
        if sc:
            m[sc] = re.search(r"<corp_code>(.*?)</corp_code>", blk).group(1).strip()
    save_json(p, m)
    return m


def select(companies, codes=None, limit=None):
    if codes:
        want = [c.strip() for c in codes.split(",") if c.strip()]
        companies = [c for c in companies if c["종목코드"] in want]
    if limit:
        companies = companies[:limit]
    return companies


# ─────────────────────────── 내려받기 ───────────────────────────

def fetch_company(corp):
    """한 회사의 DART 자료를 캐시에 저장. 이미 있으면 건너뛴다."""
    p = cache_path("corp", corp + ".json")
    if os.path.exists(p):
        rec = load_json(p)
        if "irds" not in rec:          # 2026-11-15 추가 항목 보충
            rec["irds"] = fetch_irds(corp)
            save_json(p, rec)
            return True
        return False
    rec = {"fetched": datetime.now().isoformat(timespec="seconds")}
    info = dart_company(corp)
    rec["company"] = info
    # 재무제표: 격년 호출, 당기·전기 값을 함께 쓴다 (2025,2023,...,2017 → 2016~2025)
    fin = {}
    for y in range(LAST_FY, LAST_FY - FIN_YEARS, -2):
        rows, div = [], None
        for fs in ("CFS", "OFS"):
            rows = dart("fnlttSinglAcntAll.json", corp_code=corp, bsns_year=y, reprt_code="11011", fs_div=fs)
            if rows:
                div = fs
                break
        keep = [{k: r.get(k) for k in ("sj_div", "account_id", "account_nm", "thstrm_amount", "frmtrm_amount")}
                for r in rows if r.get("sj_div") in ("BS", "IS", "CIS", "CF")]
        fin[str(y)] = {"fs_div": div, "rows": keep}
    rec["fin"] = fin
    rec["shares"] = {str(y): dart("stockTotqySttus.json", corp_code=corp, bsns_year=y, reprt_code="11011")
                     for y in (LAST_FY, LAST_FY - 5)}
    rec["holders"] = dart("hyslrSttus.json", corp_code=corp, bsns_year=LAST_FY, reprt_code="11011")
    rec["irds"] = fetch_irds(corp)
    try:
        rec["splits"] = dart("cmpDvDecsn.json", corp_code=corp,
                             bgn_de=f"{LAST_FY - 4}0101", end_de=date.today().strftime("%Y%m%d"))
    except RuntimeError as e:
        rec["splits"] = None
        rec["splits_error"] = str(e)
    save_json(p, rec)
    return True


def fetch_irds(corp):
    """증자·감자 내역. 사업보고서마다 약 3년치가 실리므로 격년 3개를 받아 합친다."""
    ev = {}
    for y in (LAST_FY, LAST_FY - 2, LAST_FY - 4):
        for r in dart("irdsSttus.json", corp_code=corp, bsns_year=y, reprt_code="11011"):
            if r.get("isu_dcrs_de") in (None, "-"):
                continue
            key = (r.get("isu_dcrs_de"), r.get("isu_dcrs_stle"), r.get("isu_dcrs_stock_knd"), r.get("isu_dcrs_qy"))
            ev[key] = {"date": key[0], "type": key[1], "kind": key[2], "qty": key[3]}
    return list(ev.values())


def dart_company(corp):
    d = {"crtfc_key": api_key(), "corp_code": corp}
    time.sleep(SLEEP)
    data = json.loads(http_get(DART + "company.json?" + urllib.parse.urlencode(d)).decode("utf-8"))
    if data.get("status") == "020":
        raise DartLimit("DART 일일 호출 한도 초과(020).")
    return {k: data.get(k) for k in ("corp_name", "induty_code", "est_dt", "acc_mt")}


def cmd_fetch(args):
    comps = select(kospi_list(args.refresh_list), args.codes, args.limit)
    cmap = corp_codes()
    done = new = miss = 0
    t0 = time.time()
    for i, c in enumerate(comps, 1):
        corp = cmap.get(c["종목코드"])
        if not corp:
            miss += 1
            continue
        try:
            if fetch_company(corp):
                new += 1
        except DartLimit as e:
            print(f"\n{e}  (진행 {i - 1}/{len(comps)})")
            break
        except Exception as e:
            print(f"\n[{c['회사명']}] 내려받기 실패: {e}")
            continue
        done += 1
        if i % 10 == 0 or i == len(comps):
            print(f"  {i}/{len(comps)}  신규 {new}  경과 {int(time.time() - t0)}초", flush=True)
    print(f"완료: 대상 {len(comps)}  처리 {done}  신규 {new}  DART 코드 없음 {miss}")


# ─────────────────────────── 계정 추출 ───────────────────────────

ACC = {
    "rev":   (["ifrs-full_Revenue", "ifrs_Revenue"], r"^(매출액|수익\(매출액\)|영업수익|매출|수익)$", ("IS", "CIS")),
    "cogs":  (["ifrs-full_CostOfSales", "ifrs_CostOfSales"], r"^(매출원가|영업원가)$", ("IS", "CIS")),
    "gp":    (["ifrs-full_GrossProfit", "ifrs_GrossProfit"], r"^매출총이익", ("IS", "CIS")),
    "op":    (["dart_OperatingIncomeLoss"], r"^영업이익", ("IS", "CIS")),
    "ni":    (["ifrs-full_ProfitLoss", "ifrs_ProfitLoss"], r"^(당기순이익|당기순손익|연결당기순이익)", ("IS", "CIS")),
    "pbt":   (["ifrs-full_ProfitLossBeforeTax", "ifrs_ProfitLossBeforeTax"], r"법인세비용차감전", ("IS", "CIS")),
    "tax":   (["ifrs-full_IncomeTaxExpenseContinuingOperations", "ifrs_IncomeTaxExpenseContinuingOperations"],
              r"^법인세비용", ("IS", "CIS")),
    "fincost": (["ifrs-full_FinanceCosts", "ifrs_FinanceCosts"], r"^(금융비용|금융원가)$", ("IS", "CIS")),
    "equity": (["ifrs-full_Equity", "ifrs_Equity"], r"^자본총계$", ("BS",)),
    "eq_own": (["ifrs-full_EquityAttributableToOwnersOfParent"], r"^지배기업.*(소유주|주주)", ("BS",)),
    "ni_own": (["ifrs-full_ProfitLossAttributableToOwnersOfParent"], r"^지배기업.*(소유주|주주)", ("IS", "CIS")),
    "cash":  (["ifrs-full_CashAndCashEquivalents", "ifrs_CashAndCashEquivalents"], r"^현금및현금성자산$", ("BS",)),
    "stfin": (["dart_ShortTermDepositsNotClassifiedAsCashEquivalents"], r"^단기금융상품$", ("BS",)),
    "ocf":   (["ifrs-full_CashFlowsFromUsedInOperatingActivities", "ifrs_CashFlowsFromUsedInOperatingActivities"],
              r"^영업활동.*현금흐름", ("CF",)),
}
SUM_ACC = {   # 여러 줄을 더하는 계정
    "capex": (["ifrs-full_PurchaseOfPropertyPlantAndEquipment", "ifrs-full_PurchaseOfIntangibleAssets",
               "ifrs_PurchaseOfPropertyPlantAndEquipment", "ifrs_PurchaseOfIntangibleAssets"],
              r"^(유형자산|무형자산)의?\s*(취득|증가|구입)", ("CF",)),
    "da":    (["ifrs-full_AdjustmentsForDepreciationExpense", "ifrs-full_AdjustmentsForAmortisationExpense",
               "ifrs-full_DepreciationAndAmortisationExpense"],
              r"^(감가상각비|무형자산상각비|사용권자산(감가)?상각비|감가상각비와\s*상각비)$", ("CF",)),
    "intpaid": (["ifrs-full_InterestPaidClassifiedAsOperatingActivities",
                 "ifrs-full_InterestPaidClassifiedAsFinancingActivities"], r"^이자(의)?\s*지급", ("CF",)),
    "sbc":   ([], r"주식보상비용|주식기준보상", ("CF", "IS", "CIS")),
}
# 2026-11-17: 설비투자를 자산 종류별로 공시하는 회사(약 100곳)를 읽도록 추가. 투자부동산·국고보조금·선수금은 제외.
CAPEX_ITEM_RE = re.compile(
    r"^(토지|건물|건물부속설비|구축물|구축물등|기계장치|미착기계|차량운반구|선박|항공기|금형|시설장치|시설물|"
    r"공구와기구|공구과기구|공구와기|공구기구|공구기구비품|공구기구와비품|공기구비품|기구비품|비품|사무용비품|집기비품|"
    r"기타유형자산|건설중인자산|건설중인유형자산|건설중인무형자산|건설중인자산유형자산|건설중인자산무형자산|"
    r"건설중인자산\((무형자산|무형)\)|개발비|소프트웨어|산업재산권|기타무형자산)의?(취득|증가|구입)$")
# 2026-11-17: '단기차입부채'·'장기차입부채'·'리스유동부채' 같은 계정명을 읽도록 추가
DEBT_RE = re.compile(r"(차입금|차입부채|단기차입$|유동성장기차입|사채|리스부채|리스(유동|비유동)부채|유동성장기부채)")
DEBT_EXCL = re.compile(r"(할인|할증|전환권|조정|상환권|신주인수권|이자|미지급|총계|합계)")
DEBT_TOTAL_RE = re.compile(r"^([IVX]+\.)?차입부채$")   # 세부 차입 계정과 함께 나오면 합계일 수 있음


def pick(rows, ids, pat, sjs, col):
    for r in rows:
        if r["sj_div"] in sjs and r.get("account_id") in ids:
            v = num(r.get(col))
            if v is not None:
                return v
    rx = re.compile(pat)
    for r in rows:
        if r["sj_div"] in sjs and rx.search((r.get("account_nm") or "").replace(" ", "")):
            v = num(r.get(col))
            if v is not None:
                return v
    return None


def pick_sum(rows, ids, pat, sjs, col):
    rx = re.compile(pat)
    vals, seen = [], set()
    for r in rows:
        if r["sj_div"] not in sjs:
            continue
        nm = (r.get("account_nm") or "").replace(" ", "")
        if r.get("account_id") in ids or rx.search(nm):
            v = num(r.get(col))
            key = (nm, v)
            if v is not None and key not in seen:
                seen.add(key)
                vals.append(abs(v))
    return sum(vals) if vals else None


def capex_items(rows, col, total):
    """자산 종류별 취득 줄의 합. 합계 줄(유형·무형자산의 취득)과 1% 안에서 같으면 내역으로 보고 0을 돌려준다."""
    vals, seen = [], set()
    for r in rows:
        if r["sj_div"] != "CF":
            continue
        nm = (r.get("account_nm") or "").replace(" ", "")
        if CAPEX_ITEM_RE.search(nm):
            v = num(r.get(col))
            if v is not None and (nm, v) not in seen:
                seen.add((nm, v))
                vals.append(abs(v))
    s = sum(vals)
    if total and s and abs(total - s) <= 0.01 * total:
        return 0.0
    return s


def debt(rows, col):
    vals, totals = [], []
    for r in rows:
        if r["sj_div"] != "BS":
            continue
        nm = (r.get("account_nm") or "").replace(" ", "")
        if DEBT_RE.search(nm) and not DEBT_EXCL.search(nm):
            v = num(r.get(col))
            if v is not None and v > 0:
                (totals if DEBT_TOTAL_RE.search(nm) else vals).append(v)
    s = sum(vals)
    if totals:   # '차입부채' 합계 줄: 세부 줄이 없으면 그대로, 있으면 둘 중 큰 값
        s = max(s, sum(totals))
    return s if (vals or totals) else 0.0


def yearly(rec):
    """{연도: {계정: 값}} — 격년 보고서의 당기·전기 값을 펼친다. 같은 연도는 최신 보고서 값을 우선."""
    out = {}
    for y_str, blk in sorted(rec["fin"].items(), key=lambda kv: -int(kv[0])):
        rows = blk["rows"]
        if not rows:
            continue
        for col, yy in (("thstrm_amount", int(y_str)), ("frmtrm_amount", int(y_str) - 1)):
            if yy in out:
                continue
            d = {k: pick(rows, *v, col) for k, v in ACC.items()}
            d.update({k: pick_sum(rows, *v, col) for k, v in SUM_ACC.items()})
            items = capex_items(rows, col, d["capex"])
            if items:
                d["capex"] = (d["capex"] or 0) + items
            d["debt"] = debt(rows, col)
            if d["gp"] is None and d["rev"] is not None and d["cogs"] is not None:
                d["gp"] = d["rev"] - d["cogs"]
            # 매출원가가 없는 회사(성격별 분류)는 '매출총이익 = 매출' 또는 '= 영업이익'으로 들어온다 → 무효
            if d["gp"] is not None and d["rev"]:
                tol = abs(d["rev"]) * 0.001
                if abs(d["gp"] - d["rev"]) <= tol or (d["op"] is not None and abs(d["gp"] - d["op"]) <= tol):
                    d["gp"] = None
            bs_names = " ".join((r.get("account_nm") or "") for r in rows if r["sj_div"] == "BS")
            d["fin_sub"] = bool(re.search(r"금융업채권|금융업부채|금융업차입", bs_names))
            if any(d[k] is not None for k in ("rev", "op", "ni")):
                d["fs_div"] = blk["fs_div"]
                out[yy] = d
    return out


# ─────────────────────────── 지표 계산 ───────────────────────────

def mean(xs):
    xs = [x for x in xs if x is not None]
    return statistics.mean(xs) if xs else None


def last_n(ys, n):
    return [y for y in sorted(ys, reverse=True)[:n]]


def metrics(rec):
    Y = yearly(rec)
    if not Y:
        return None
    yrs = sorted(Y, reverse=True)
    m = {"years": len(yrs), "first_year": min(yrs), "last_year": max(yrs),
         "fs_div": Y[yrs[0]].get("fs_div")}

    def ratio(a, b):
        return a / b if (a is not None and b not in (None, 0)) else None

    # 1. ROIC (7년)
    roics, roes = [], []
    for y in last_n(yrs, 7):
        d = Y[y]
        if d["op"] is None or d["equity"] is None:
            continue
        tr = ratio(d["tax"], d["pbt"]) if (d["pbt"] or 0) > 0 else None
        tr = 0.22 if tr is None else min(max(tr, 0.0), 0.35)
        ic = d["equity"] + d["debt"] - (d["cash"] or 0) - (d["stfin"] or 0)
        if ic > 0:
            roics.append(d["op"] * (1 - tr) / ic)
        else:
            roics.append(None)   # 순현금이 자본보다 큼 → 계산 불가(무자본에 가까움)
    for y in last_n(yrs, 5):
        d = Y[y]
        ni = d["ni_own"] if d["ni_own"] is not None else d["ni"]
        eq = d["eq_own"] if d["eq_own"] else d["equity"]
        r = ratio(ni, eq)
        if r is not None and eq > 0:
            roes.append(r)
    m["roic7"] = mean(roics)
    m["roic_neg_ic"] = sum(1 for r in roics if r is None)
    m["roe5"] = mean(roes)

    # 2. FCF (5년)
    f5 = last_n(yrs, 5)
    fcf = [(Y[y]["ocf"] - (Y[y]["capex"] or 0)) if Y[y]["ocf"] is not None else None for y in f5]
    if all(x is not None for x in fcf) and fcf:
        m["fcf5"] = sum(fcf)
        ni5 = sum((Y[y]["ni"] or 0) for y in f5)
        m["fcf_ni"] = m["fcf5"] / ni5 if ni5 > 0 else None
        m["ni5"] = ni5
    else:
        m["fcf5"] = m["fcf_ni"] = m["ni5"] = None

    # 3. 재무 안전 (최근 연도)
    d = Y[yrs[0]]
    netdebt = d["debt"] - (d["cash"] or 0) - (d["stfin"] or 0)
    interest = d["intpaid"] if d["intpaid"] else d["fincost"]
    m["icr"] = (d["op"] / interest) if (d["op"] is not None and interest) else (None if d["debt"] > 0 else 999)
    m["icr_src"] = "이자지급" if d["intpaid"] else ("금융비용" if d["fincost"] else "-")
    ebitda = (d["op"] + d["da"]) if (d["op"] is not None and d["da"]) else None
    m["netdebt"] = netdebt
    m["nd_ebitda"] = (netdebt / ebitda) if (ebitda and ebitda > 0) else (None if netdebt > 0 else -1)
    m["ebitda_ok"] = ebitda is not None
    m["fin_sub"] = bool(d.get("fin_sub"))

    # 4. 매출총이익률
    gms = [ratio(Y[y]["gp"], Y[y]["rev"]) for y in f5]
    m["gm5"] = mean(gms)
    gm_valid = [(y, ratio(Y[y]["gp"], Y[y]["rev"])) for y in f5 if ratio(Y[y]["gp"], Y[y]["rev"]) is not None]
    m["gm_drop"] = (gm_valid[-1][1] - gm_valid[0][1]) if len(gm_valid) >= 3 else None  # 5년 전 − 최근

    # 5. 적자 연도 (10년)
    t10 = last_n(yrs, 10)
    m["loss_years"] = sum(1 for y in t10 if (Y[y]["ni"] or 0) < 0)
    m["ni_years"] = sum(1 for y in t10 if Y[y]["ni"] is not None)

    # 6. 영업이익률
    oms = {y: ratio(Y[y]["op"], Y[y]["rev"]) for y in f5}
    m["om5"] = mean(oms.values())
    o3 = [oms.get(y) for y in sorted(f5, reverse=True)[:3]]
    m["om_last2"] = mean(o3[:2])
    m["om_rising"] = all(o is not None for o in o3) and len(o3) == 3 and o3[0] > o3[1] > o3[2]
    m["npm5"] = mean([ratio(Y[y]["ni"], Y[y]["rev"]) for y in f5])

    # 예외 A용: 최근 2년 영업현금흐름
    m["ocf_pos2"] = all((Y[y]["ocf"] or -1) > 0 for y in yrs[:2])

    # 7. 희석
    def total_shares(rows):
        for r in rows or []:
            if (r.get("se") or "").replace(" ", "") in ("합계", "계"):
                return num(r.get("istc_totqy"))
        vals = [num(r.get("istc_totqy")) for r in rows or [] if num(r.get("istc_totqy"))]
        return sum(vals) if vals else None
    s_now = total_shares(rec["shares"].get(str(LAST_FY)))
    s_old = total_shares(rec["shares"].get(str(LAST_FY - 5)))
    m["share_cagr"], m["share_note"] = share_growth(s_now, s_old, rec.get("irds"))
    sbc = Y[yrs[0]]["sbc"]
    m["sbc_rev"] = ratio(sbc, Y[yrs[0]]["rev"]) if sbc else None

    # 지배구조
    m["holder_pct"] = holder_pct(rec.get("holders"))
    sp = rec.get("splits")
    m["phys_split"] = None if sp is None else any("물적" in json.dumps(s, ensure_ascii=False) for s in sp)
    return m


def share_growth(s_now, s_old, events):
    """5년 주식 수 연평균 증가율. 주식분할·무상증자·주식배당은 희석이 아니므로 뺀다. 합병 신주도 뺀다(규칙상 제외)."""
    cutoff = f"{LAST_FY - 5}.12.31"
    ev = [e for e in (events or []) if (e.get("date") or "").replace("-", ".") > cutoff]
    neutral = [e for e in ev if any(k in (e["type"] or "") for k in NEUTRAL_SHARE_EVENTS)]
    merger = sum(num(e["qty"]) or 0 for e in ev if "합병" in (e["type"] or ""))
    dil = sum(num(e["qty"]) or 0 for e in ev if any(k in (e["type"] or "") for k in DILUTIVE_SHARE_EVENTS))
    note = []
    if merger:
        note.append("합병 신주 제외")
    if s_now and s_old and not (0.2 < s_now / s_old < 20):
        return None, "주식 수 원자료 단위 오류 의심"
    if neutral and s_now:
        base = s_now - dil - merger
        note.append("분할·무상증자 보정")
        g = ((s_now - merger) / base) ** (1 / 5) - 1 if base > 0 else None
        return g, ", ".join(note)
    if s_now and s_old:
        return ((s_now - merger) / s_old) ** (1 / 5) - 1, ", ".join(note)
    return None, ", ".join(note)


def holder_pct(rows):
    if not rows:
        return None
    total = None
    for r in rows:
        if (r.get("nm") or "").replace(" ", "") in ("계", "합계"):
            kind = r.get("stock_knd") or ""
            if "우선" in kind:
                continue
            v = num(r.get("trmend_posesn_stock_qota_rt"))
            if v is not None:
                total = v if total is None else max(total, v)
    if total is None:
        vals = [num(r.get("trmend_posesn_stock_qota_rt")) for r in rows
                if "우선" not in (r.get("stock_knd") or "")]
        vals = [v for v in vals if v is not None]
        total = sum(vals) if vals else None
    return total


# ─────────────────────────── 판정 ───────────────────────────

def pctile(sorted_vals, q):
    if not sorted_vals:
        return None
    k = (len(sorted_vals) - 1) * q
    lo, hi = int(k), min(int(k) + 1, len(sorted_vals) - 1)
    return sorted_vals[lo] + (sorted_vals[hi] - sorted_vals[lo]) * (k - lo)


def peer_stats(rows, key, field):
    groups = {}
    for r in rows:
        v = r["m"].get(field) if r["m"] else None
        if v is not None:
            groups.setdefault(r[key], []).append(v)
    return {g: sorted(v) for g, v in groups.items() if len(v) >= 5}


def classify(c):
    ind, prod, name = c["업종"], c.get("주요제품", ""), c["회사명"]
    if "리츠" in name or "부동산투자" in prod or "위탁관리" in prod:
        return "리츠"
    if "인프라" in name or ind == "신탁업 및 집합투자업" and "리얼티" in name:
        return "펀드"
    if "비금융" in prod:
        return "지주"
    if ind == "기타 금융업":
        # KRX가 비금융 지주회사도 '기타 금융업'으로 분류한다 → 이름에 '금융'이 있거나 여신업일 때만 금융
        if "금융" in name or re.search(r"신용카드|캐피탈|여신|리스", prod):
            return "금융"
        return "지주"
    if "지주" in prod or "지주회사" in ind:
        return "지주"
    if any(k in ind for k in FIN_INDUSTRIES):
        return "금융"
    return "일반"


def judge(r, peers):
    m, c = r["m"], r["c"]
    res, notes, flags = {}, [], []
    kind = r["kind"]
    cyc = any(k in c["업종"] for k in CYCLICAL_KEYS)

    def peer(field):
        for key in ("ind", "ksic2"):
            st = peers[key][field].get(r[key])
            if st:
                return st, key
        return None, None

    # 1
    if kind == "금융":
        v = m["roe5"]
        res[1] = ("?" if v is None else ("✅" if v >= 0.09 else "❌"), v)
    else:
        v = m["roic7"]
        if v is None and m["roic_neg_ic"]:
            res[1] = ("✅", None); notes.append("투하자본 음수(순현금 과다)로 ROIC 계산 불가 → 통과 처리")
        else:
            res[1] = ("?" if v is None else ("✅" if v >= 0.09 else "❌"), v)
    # 2
    if kind == "금융":
        res[2] = ("—", None)
    elif m["fcf5"] is None:
        res[2] = ("?", None)
    else:
        bad = m["fcf5"] < 0 or (m["fcf_ni"] is not None and m["fcf_ni"] < 0.5)
        res[2] = ("❌" if bad else "✅", m["fcf_ni"])
    # 3
    if kind == "금융":
        res[3] = ("—", None)
    else:
        icr, nde = m["icr"], m["nd_ebitda"]
        if icr is None and nde is None:
            res[3] = ("?", None)
        else:
            bad = (icr is not None and icr < 3.5) or (nde is not None and nde > 3)
            res[3] = ("❌" if bad else "✅", icr)
            if not m["ebitda_ok"] and m["netdebt"] > 0:
                notes.append("감가상각비 미확인 → 순부채/EBITDA 미판정")
            if m["icr_src"] == "금융비용":
                notes.append("이자비용 대신 금융비용 사용(보수적)")
    # 4 · 6
    gm_st, gm_key = peer("gm5")
    om_st, om_key = peer("om5")
    r["gm_top3"] = bool(gm_st and m["gm5"] is not None and m["gm5"] >= pctile(gm_st, 2 / 3) and m["gm5"] >= 0.20)
    if kind == "금융":
        res[4] = res[6] = ("—", None)
    else:
        if m["gm5"] is None:
            res[4] = ("?", None)
        elif gm_st:
            bad = m["gm5"] < pctile(gm_st, 1 / 3) or (m["gm_drop"] is not None and m["gm_drop"] > 0.05)
            res[4] = ("❌" if bad else "✅", m["gm5"])
        else:
            bad = m["gm5"] < 0.15
            res[4] = ("❌" if bad else "✅", m["gm5"]); notes.append("4번 절대 기준 대체")
        if m["om5"] is None:
            res[6] = ("?", None)
        elif om_st:
            res[6] = ("❌" if m["om5"] < pctile(om_st, 1 / 3) else "✅", m["om5"])
        else:
            v = m["npm5"]
            res[6] = ("?" if v is None else ("❌" if v < 0.05 else "✅"), v, "순"); notes.append("6번 절대 기준(순이익률 5%) 대체")
    # 5
    lim = 4 if cyc else 3
    res[5] = ("?" if m["ni_years"] < 3 else ("❌" if m["loss_years"] >= lim else "✅"), m["loss_years"])
    # 7
    sc, sb = m["share_cagr"], m["sbc_rev"]
    if sc is None and sb is None:
        res[7] = ("?", None)
    else:
        bad = (sc is not None and sc > 0.03) or (sb is not None and sb > 0.08)
        res[7] = ("⚠️❌" if bad else "✅", sc)
        if bad:
            notes.append("7번: 증가 원인 수동 확인 필요")
        if m.get("share_note"):
            notes.append("7번 " + m["share_note"])

    # 예외
    exempt = []
    listed = c.get("상장일", "")
    young = bool(listed) and (date.today() - datetime.strptime(listed, "%Y-%m-%d").date()).days < 3652
    if res[1][0] == "❌" and young and r["gm_top3"] and m["ocf_pos2"]:
        res[1] = ("⚠️→✅", res[1][1]); exempt.append("A")
    if res[6][0] == "❌" and r["gm_top3"]:
        med = pctile(om_st, 0.5) if om_st else None
        if (med is not None and m["om_last2"] is not None and m["om_last2"] >= med) or m["om_rising"]:
            res[6] = ("⚠️→✅", res[6][1]); exempt.append("B")
    c_cand = (res[4][0] == "❌" or res[6][0] == "❌") and (m["roic7"] or 0) > 0.17 and (m["fcf_ni"] or 0) > 1.0

    # 플래그
    hp = m["holder_pct"]
    if hp is not None and hp < 20:
        flags.append("⚑지분과소")
    if hp is not None and hp >= 70:
        flags.append("⚑지분과다")
    if m["phys_split"]:
        flags.append("⚑쪼개기")
    if kind == "지주":
        flags.append("⚑지주")

    if kind == "금융":
        notes.append("금융사: DART 전체 재무제표가 2022년분부터만 제공")
    if m["fin_sub"]:
        notes.append("금융자회사 연결: 1·2·3번 왜곡 가능")
    if m["years"] < 3:
        notes.append(f"데이터 창 부족({m['years']}년)")
    fails = [k for k, v in res.items() if v[0] in ("❌",)]
    soft = [k for k, v in res.items() if v[0] == "⚠️❌"]
    unknown = [k for k, v in res.items() if v[0] == "?"]
    if fails and m["fin_sub"] and set(fails) <= {1, 2, 3}:
        verdict = "보류(금융자회사)"
    elif fails and c_cand and set(fails) <= {4, 6}:
        verdict = "보류(예외C 확인)"
    elif fails:
        verdict = "탈락"
    elif len(unknown) >= 3:
        verdict = "데이터 부족"
    elif soft:
        verdict = "보류(7번 확인)"
    elif exempt:
        verdict = "예외 통과"
    else:
        verdict = "통과"
    return res, verdict, exempt, flags, notes, cyc


def cmd_screen(args):
    comps = select(kospi_list(), args.codes, args.limit)
    cmap = corp_codes()
    rows, skipped = [], []
    for c in comps:
        corp = cmap.get(c["종목코드"])
        p = os.path.join(CACHE, "corp", f"{corp}.json") if corp else None
        if not p or not os.path.exists(p):
            skipped.append(c["회사명"])
            continue
        rec = load_json(p)
        kind = classify(c)
        m = metrics(rec) if kind not in ("리츠", "펀드") else None
        ksic = (rec.get("company") or {}).get("induty_code") or ""
        rows.append({"c": c, "corp": corp, "kind": kind, "m": m, "ind": c["업종"], "ksic2": ksic[:2]})
    peers = {key: {f: peer_stats([r for r in rows if r["kind"] in ("일반", "지주")], key, f)
                   for f in ("gm5", "om5")} for key in ("ind", "ksic2")}

    out = []
    for r in rows:
        c = r["c"]
        if r["kind"] in ("리츠", "펀드"):
            out.append({"회사명": c["회사명"], "종목코드": c["종목코드"], "업종": c["업종"], "구분": r["kind"],
                        "결과": "제외(리츠·인프라펀드)"})
            continue
        if not r["m"]:
            out.append({"회사명": c["회사명"], "종목코드": c["종목코드"], "업종": c["업종"], "구분": r["kind"],
                        "결과": "데이터 부족"})
            continue
        res, verdict, ex, flags, notes, cyc = judge(r, peers)
        m = r["m"]
        pct = lambda v: "" if v is None else f"{v * 100:.1f}%"
        out.append({
            "회사명": c["회사명"], "종목코드": c["종목코드"], "업종": c["업종"], "구분": r["kind"] + ("·경기민감" if cyc else ""),
            "결과": verdict,
            "①자본효율": f"{res[1][0]} {pct(res[1][1])}".strip(),
            "②FCF/순이익": f"{res[2][0]} {pct(res[2][1])}".strip(),
            "③이자보상": f"{res[3][0]} " + ("" if res[3][1] is None else ("무차입" if res[3][1] == 999 else f"{res[3][1]:.1f}배")),
            "④매출총이익률": f"{res[4][0]} {pct(res[4][1])}".strip(),
            "⑤적자연도": f"{res[5][0]} {res[5][1]}년",
            "⑥영업이익률": f"{res[6][0]} {res[6][2] if len(res[6]) > 2 else ''}{pct(res[6][1])}".strip(),
            "⑦주식수연증가": f"{res[7][0]} {pct(res[7][1])}".strip(),
            "순부채/EBITDA": "" if m["nd_ebitda"] is None else ("순현금" if m["nd_ebitda"] < 0 else f"{m['nd_ebitda']:.1f}배"),
            "최대주주지분": "" if m["holder_pct"] is None else f"{m['holder_pct']:.1f}%",
            "플래그": " ".join(flags), "예외": ",".join(ex), "비고": "; ".join(notes),
            "자료연도": f"{m['first_year']}~{m['last_year']}", "재무제표": "연결" if m["fs_div"] == "CFS" else "별도",
        })

    order = {"통과": 0, "예외 통과": 1, "보류(예외C 확인)": 2, "보류(금융자회사)": 3, "보류(7번 확인)": 4,
             "탈락": 5, "데이터 부족": 6}
    out.sort(key=lambda x: (order.get(x["결과"], 9), x["업종"], x["회사명"]))
    base = args.out or os.path.join(ROOT, "reports", f"KOSPI-불량기업걸러내기-{date.today():%Y%m%d}")
    os.makedirs(os.path.dirname(base), exist_ok=True)
    cols = []
    for o in out:
        for k in o:
            if k not in cols:
                cols.append(k)
    with open(base + ".csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(out)
    from collections import Counter
    cnt = Counter(o["결과"] for o in out)
    print(f"판정 {len(out)}곳 (캐시 없음 {len(skipped)}곳)")
    for k, v in cnt.most_common():
        print(f"  {k}: {v}")
    print("CSV:", base + ".csv")
    return out


def main():
    ap = argparse.ArgumentParser(description="KOSPI 불량기업 걸러내기 (DART 기반)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name in ("fetch", "screen", "run"):
        s = sub.add_parser(name)
        s.add_argument("--codes", help="종목코드 쉼표 구분")
        s.add_argument("--limit", type=int)
        s.add_argument("--out", help="출력 경로(확장자 제외)")
        s.add_argument("--refresh-list", action="store_true", help="KIND 종목 목록 새로 받기")
    a = ap.parse_args()
    if a.cmd in ("fetch", "run"):
        cmd_fetch(a)
    if a.cmd in ("screen", "run"):
        cmd_screen(a)


if __name__ == "__main__":
    main()
