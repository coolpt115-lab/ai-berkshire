import sys, csv
sys.path.insert(0, sys.argv[1])
from 종합자료 import rows, Q, NAME
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

CSV, OUT = sys.argv[2], sys.argv[3]
F = "맑은 고딕"
fb = Font(name=F, size=10); fbold = Font(name=F, size=10, bold=True)
fh = Font(name=F, size=10, bold=True, color="FFFFFF"); ft = Font(name=F, size=14, bold=True)
fn = Font(name=F, size=9, color="595959")
fill_h = PatternFill("solid", fgColor="1F3864")
fills = {"✅": PatternFill("solid", fgColor="C6EFCE"), "❓": PatternFill("solid", fgColor="FFEB9C"),
         "⏳": PatternFill("solid", fgColor="DDEBF7"), "❌": PatternFill("solid", fgColor="FFC7CE")}
thin = Side(style="thin", color="A6A6A6"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
C = Alignment(horizontal="center", vertical="center", wrap_text=True)
Lf = Alignment(horizontal="left", vertical="center", wrap_text=True)

def head(ws, r, cols):
    for i, h in enumerate(cols, 1):
        c = ws.cell(row=r, column=i, value=h); c.font = fh; c.fill = fill_h; c.alignment = C; c.border = box
def cell(ws, r, col, v, fmt=None, al=C, font=fb):
    c = ws.cell(row=r, column=col, value=v); c.font = font; c.alignment = al; c.border = box
    if fmt: c.number_format = fmt
    return c
def widths(ws, ws_w):
    for i, x in enumerate(ws_w, 1): ws.column_dimensions[get_column_letter(i)].width = x

R = rows()
wb = Workbook()

# 시트 2 먼저 만들어 두고 요약에서 참조
t = wb.active; t.title = "37곳 판정표"
t["A1"] = "정밀 체크리스트 37곳 판정표 (판정 기준 통일 후)"; t["A1"].font = ft
t["A2"] = "주가 2026-11-13 종가 · 파란 글씨는 보고서에서 옮긴 입력값, 검은 글씨는 수식 · 별점 순서: 능력범위·좋은 사업·해자·경영진·안전마진·결정 규율"; t["A2"].font = fn
cols = ["종목","구분","통일 판정","기존 판정","1관","2관","3관","4관","5관","6관","품질 결함(★2) 수","품질 ★1 여부","현재가(원)","중립 내재가치(원)","현재가 위치(할인율)","필요 안전마진","매수 검토가(원)","가격 관문","매수 검토가까지 하락폭","핵심 쟁점","보고서"]
head(t, 4, cols)
fin = Font(name=F, size=10, color="0000FF")
for i, r in enumerate(R):
    row = 5 + i
    st = [int(x) for x in r['별점'].split('·')]
    cell(t, row, 1, r['종목'], font=fbold)
    cell(t, row, 2, r['구분'])
    c = cell(t, row, 3, NAME[r['통일']], font=fbold); c.fill = fills[r['통일']]
    cell(t, row, 4, NAME[r['기존']].replace("❌ 부결", "❌ 미통과"))
    for j, s in enumerate(st): cell(t, row, 5 + j, s, font=fin)
    cell(t, row, 11, f"=COUNTIF(F{row}:H{row},2)+COUNTIF(E{row},2)")
    cell(t, row, 12, f'=IF(MIN(E{row}:H{row})=1,"있음","")')
    cell(t, row, 13, r['현재가'], "#,##0", font=fin)
    cell(t, row, 14, r['중립'], "#,##0", font=fin)
    cell(t, row, 15, f"=1-M{row}/N{row}", "0.0%")
    cell(t, row, 16, r['필요'] / 100, "0%", font=fin)
    cell(t, row, 17, r['매수'], "#,##0", font=fin)
    cell(t, row, 18, f'=IF(I{row}>=4,"합격","불합격")')
    cell(t, row, 19, f"=MAX(0,1-Q{row}/M{row})", "0%")
    cell(t, row, 20, r['쟁점'], al=Lf)
    cell(t, row, 21, "reports/" + r['경로'], al=Lf, font=fn)
last = 4 + len(R)
widths(t, [16,6,13,12,5,5,5,5,5,5,9,8,12,13,11,9,12,8,11,60,40])
t.freeze_panes = "B5"; t.auto_filter.ref = f"A4:U{last}"
t.conditional_formatting.add(f"O5:O{last}", CellIsRule(operator="lessThan", formula=["0"], font=Font(color="C00000")))

# 요약
s = wb.create_sheet("판정 요약", 0)
s["A1"] = "KOSPI 1단계 통과 53곳 버핏 체크리스트 판정 요약"; s["A1"].font = ft
s["A2"] = "작성 2026-11-17 · 판정 기준 통일(reports/판정기준-통일-20261117.md) 적용 · 숫자는 다른 시트에서 수식으로 셈"; s["A2"].font = fn
head(s, 4, ["단계", "구분", "기업 수", "비중", "기준"])
J = f"'37곳 판정표'!$C$5:$C${last}"
steps = [
 ("0단계","KOSPI 상장법인 전체","=COUNTA('1단계 전체'!$A:$A)-1",None,"KRX KIND 상장법인 목록"),
 ("0단계","제외: 리츠·인프라펀드","=COUNTIF('1단계 전체'!$E:$E,\"제외*\")","=C6/C5","별도 평가 기준 필요"),
 ("1단계","판정 대상","=C5-C6","=C7/C5","불량기업 걸러내기 7개 지표"),
 ("1단계","통과(예외 통과 포함)","=COUNTIF('1단계 전체'!$E:$E,\"통과\")+COUNTIF('1단계 전체'!$E:$E,\"예외 통과\")","=C8/C7","7개 지표 통과 또는 예외 충족"),
 ("2단계","정밀 체크리스트",f"=COUNTA('37곳 판정표'!$A$5:$A${last})","=C9/C8","6관 버핏 체크리스트"),
 ("2단계","빠른 판정만(가격 미달)",f"=COUNTA('16곳 가격 미달'!$A$5:$A$20)","=C10/C8","현재가가 중립 내재가치의 10% 할인 안쪽이거나 더 비쌈"),
 ("판정","✅ 통과",f'=COUNTIF({J},"✅*")',"=C11/$C$9","품질 모두 ★3 이상 + 가격 합격 + 부결 0 + 거울 테스트 완성"),
 ("판정","❓ 회색지대",f'=COUNTIF({J},"❓*")',"=C12/$C$9","품질 결함 1개 + 가격 합격, 또는 거울 테스트 한 문장 미완성"),
 ("판정","⏳ 가격 대기",f'=COUNTIF({J},"⏳*")',"=C13/$C$9","품질 합격, 가격 불합격"),
 ("판정","❌ 부결",f'=COUNTIF({J},"❌*")',"=C14/$C$9","빠른 부결, ★1, 결함 2개 이상, 또는 결함 + 가격 불합격"),
]
for i, x in enumerate(steps):
    r = 5 + i
    for j, v in enumerate(x, 1):
        c = cell(s, r, j, v, al=Lf if j in (2, 5) else C)
    s.cell(row=r, column=3).number_format = "#,##0"; s.cell(row=r, column=4).number_format = "0.0%"
    if x[0] == "판정": s.cell(row=r, column=2).fill = fills[x[1][0]]
s["A16"] = "검산: 판정 합계"; s["A16"].font = fbold
s["C16"] = "=SUM(C11:C14)"; s["D16"] = '=IF(C16=C9,"정밀 37곳과 일치","불일치")'
s["A17"] = "검산: 정밀 + 빠른 판정"; s["A17"].font = fbold
s["C17"] = "=C9+C10"; s["D17"] = '=IF(C17=C8,"1단계 통과 수와 일치","불일치")'
s["A19"] = "✅ 통과 4곳은 모두 조건부 통과다. 공통 약점은 해자 추세가 불명확하다는 점이며, 다음 단계는 심층 연구다."; s["A19"].font = fn
widths(s, [10, 26, 10, 9, 60])

# 16곳
q = wb.create_sheet("16곳 가격 미달")
q["A1"] = "정밀 체크리스트를 하지 않은 16곳 (빠른 판정 기준)"; q["A1"].font = ft
q["A2"] = "사업·경영진 미평가. 빠른 판정은 krstock_screen.py 결함 수정 전 값이라 재판정 필요"; q["A2"].font = fn
head(q, 4, ["종목", "구분", "현재가(원)", "빠른 판정 중립 내재가치(원)", "현재가 위치(할인율)", "빠른 판정 분류"])
for i, (n, g, p, iv, c) in enumerate(Q):
    r = 5 + i
    cell(q, r, 1, n, font=fbold); cell(q, r, 2, g); cell(q, r, 3, p, "#,##0", font=fin)
    cell(q, r, 4, iv if iv else "–", "#,##0", font=fin)
    cell(q, r, 5, f"=IF(ISNUMBER(D{r}),1-C{r}/D{r},\"–\")", "0.0%"); cell(q, r, 6, c, al=Lf)
q.conditional_formatting.add("E5:E20", CellIsRule(operator="lessThan", formula=["0"], font=Font(color="C00000")))
widths(q, [18, 6, 12, 16, 12, 22])

# 규칙
g = wb.create_sheet("판정 규칙")
g["A1"] = "판정 규칙 (2026-11-17 통일)"; g["A1"].font = ft
head(g, 3, ["판정", "조건", "뜻"])
rules = [("✅ 통과","품질 관문(1~4관) 모두 ★3 이상, 가격 관문(5관 ★4 이상 = 필요 안전마진 충족) 합격, 빠른 부결 0개, 거울 테스트 5문장 완성(조건부 포함)","심층 연구로 넘어간다"),
 ("❓ 회색지대","품질 결함(★2) 1개뿐이고 가격 합격, 또는 모든 관문 합격인데 거울 테스트 한 문장을 쓸 수 없음","싼 가격이 결함을 보상하는지 투자자가 판단"),
 ("⏳ 가격 대기","품질 관문 모두 ★3 이상인데 가격 불합격","매수 검토가까지 내려오면 다시 본다"),
 ("❌ 부결","빠른 부결 확정 1개 이상, 품질 관문 ★1, 품질 결함 2개 이상, 또는 품질 결함 + 가격 불합격","지금 가격에서는 보지 않는다")]
for i, x in enumerate(rules):
    for j, v in enumerate(x, 1):
        c = cell(g, 4 + i, j, v, al=Lf);
    g.cell(row=4 + i, column=1).fill = fills[x[0][0]]
g["A9"] = "6관(결정 규율)은 투자자의 감정 상태를 보는 관문이라 판정에 넣지 않고 경고로만 쓴다. 필요 안전마진: 기본 33%(1급 회사 25%), 대주주 감점 2개 이상이면 +5%p."; g["A9"].font = fn
widths(g, [14, 80, 36])
for r in range(4, 8): g.row_dimensions[r].height = 36

# 1단계 전체
a = wb.create_sheet("1단계 전체")
data = list(csv.reader(open(CSV, encoding="utf-8-sig")))
for r, line in enumerate(data, 1):
    for c_, v in enumerate(line, 1):
        x = a.cell(row=r, column=c_, value=v); x.font = fh if r == 1 else fb
        if r == 1: x.fill = fill_h
a.freeze_panes = "B2"; a.auto_filter.ref = f"A1:{get_column_letter(len(data[0]))}{len(data)}"
widths(a, [18, 9, 24, 8, 10] + [11] * 14)

wb.save(OUT); print("저장", OUT)
