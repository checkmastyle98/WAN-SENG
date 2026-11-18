"""가상 회사 '가온페이퍼(가상)'의 파일 서버 데이터를 만든다.

실제 회사 자료는 쓰지 않는다. 회사명, 사람, 자재, 단가, PO 번호는 모두 지어낸 값이다.
실행: python make_data.py  ->  data/ 폴더가 새로 만들어진다.
"""
import json
import os
import random
import shutil
from datetime import date, datetime

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill

random.seed(7)
ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
FS = os.path.join(ROOT, "fileserver")
BASE = os.path.join(FS, "생산부", "조성파트")

# 모든 파일 시각을 '파일 서버에 저장된 날짜'로 맞춘다 (전임자가 언제 만들었는지 보여주기 위함)
META = {}


def touch(path, d):
    """파일 시각을 맞추고, git에 올려도 날짜가 남도록 meta에도 기록한다."""
    ts = datetime(d.year, d.month, d.day, 17, 30).timestamp()
    os.utime(path, (ts, ts))
    META[os.path.relpath(path, ROOT).replace(os.sep, "/")] = d.isoformat()


def write_text(rel, text, d):
    path = os.path.join(FS, rel)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text.strip() + "\n")
    touch(path, d)


# ---------------------------------------------------------------- 자재 마스터
MATERIALS = [
    # 자재코드, 자재명, 단위, 공급사, 계약단가(원), VMI여부
    ("RM-2101", "AKD 사이즈제", "kg", "한결케미칼", 1850, False),
    ("RM-2102", "양성전분", "kg", "누리전분", 920, False),
    ("RM-2103", "보류제 A", "kg", "한결케미칼", 3100, False),
    ("RM-2104", "탄산칼슘(충전제)", "ton", "백산광업", 142000, True),
    ("RM-2105", "형광증백제", "kg", "다온컬러", 4300, False),
    ("RM-2106", "소포제", "kg", "한결케미칼", 2650, True),
    ("SP-3301", "펌프 메카니컬씰", "EA", "동진기계", 385000, False),
    ("SP-3302", "와이어 세척 노즐", "EA", "동진기계", 42000, False),
    ("SP-3303", "베어링 6312", "EA", "세원베어링", 68000, False),
    ("SP-3304", "백업필터 카트리지", "EA", "청류필터", 27000, True),
]
PRICE_2025 = {m[0]: int(m[4] * 0.95 // 10 * 10) for m in MATERIALS}  # 2025-11 단가 변경 전


def po_rows(year, month, n=12, urgent=0):
    rows = []
    for i in range(n):
        m = random.choice([x for x in MATERIALS if not x[5]])
        qty = random.choice([50, 100, 200, 300, 500]) if m[2] == "kg" else random.choice([2, 4, 6, 10])
        recv = random.choice([1.0, 1.0, 1.0, 0.95, 0.8, 0.6])
        note = ""
        if i < urgent:
            note = "긴급 - 설비 고장 대응"
        rows.append([
            f"45{year % 100:02d}{month:02d}{i + 1:03d}",
            f"{year}-{month:02d}-{random.randint(1, 25):02d}",
            m[0], m[1], qty, m[2], m[3], None, None,
            f"{year}-{month:02d}-{random.randint(15, 28):02d}",
            round(qty * recv) if m[2] != "EA" else int(qty * recv),
            None, note,
        ])
    return rows


def vmi_rows(year, month):
    rows = []
    for m in [x for x in MATERIALS if x[5]]:
        used = random.choice([12, 18, 25]) if m[2] == "ton" else random.choice([40, 80, 120])
        rows.append([m[0], m[1], m[3], used, m[2], f"{year}-{month:02d}-말일", "월말 사용량 기준 정산"])
    return rows


REV_LOG = [
    ("Rev.1", "2025-03-02", "이전담당(이도현)", "최초 작성. SAP 구매오더 조회 결과를 월별로 정리하는 양식 신설", ""),
    ("Rev.2", "2025-07-15", "이전담당(이도현)", "VMI 품목을 'VMI정산' 시트로 분리. PO 목록에서 VMI 품목 제외", "2025-07-14 구매프로세스 개선회의"),
    ("Rev.3", "2025-11-24", "이전담당(이도현)", "단가 컬럼을 '계약단가' 시트에서 VLOOKUP으로 자동 연동", "2025-11-20 공급사 계약단가 변경 메일"),
    ("Rev.4", "2026-02-04", "이전담당(이도현)", "비고란 규칙 추가: 긴급발주는 '긴급 - 사유' 형식으로 기재", "2026-02-03 월간 생산회의"),
    ("Rev.5", "2026-04-08", "이전담당(이도현)", "'입고율' 컬럼 추가, 90% 미만이면 빨간색 표시", ""),
]


def build_po_workbook(path, year, month, rev_upto, rows=None, vmi=True, d=None, extra_note=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "PO목록"
    headers = ["PO번호", "발주일", "자재코드", "자재명", "발주수량", "단위", "공급사", "단가", "금액", "납기일", "입고수량", "입고율", "비고"]
    has_vmi = rev_upto >= 2 and vmi
    has_price_link = rev_upto >= 3
    has_rate = rev_upto >= 5
    if not has_rate:
        headers.remove("입고율")
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
    rows = rows or []
    for r_i, r in enumerate(rows, start=2):
        r = list(r)
        if not has_rate:
            r.pop(11)
        ws.append(r)
        col = {h: i + 1 for i, h in enumerate(headers)}
        code_cell = ws.cell(r_i, col["자재코드"]).coordinate
        qty_cell = ws.cell(r_i, col["발주수량"]).coordinate
        price = ws.cell(r_i, col["단가"])
        if has_price_link:
            price.value = f"=VLOOKUP({code_cell},계약단가!$A:$C,3,FALSE)"
        else:
            price.value = PRICE_2025[r[2]]
        ws.cell(r_i, col["금액"]).value = f"={qty_cell}*{price.coordinate}"
        if has_rate:
            recv_cell = ws.cell(r_i, col["입고수량"]).coordinate
            ws.cell(r_i, col["입고율"]).value = f"=IFERROR({recv_cell}/{qty_cell},0)"
            ws.cell(r_i, col["입고율"]).number_format = "0%"
    if has_rate and rows:
        rate_col = ws.cell(1, headers.index("입고율") + 1).column_letter
        ws.conditional_formatting.add(
            f"{rate_col}2:{rate_col}{len(rows) + 1}",
            CellIsRule(operator="lessThan", formula=["0.9"], font=Font(color="C00000", bold=True)),
        )
    if extra_note:
        ws.cell(len(rows) + 3, 1).value = extra_note

    if has_vmi:
        v = wb.create_sheet("VMI정산")
        v.append(["자재코드", "자재명", "공급사", "월 사용량", "단위", "정산기준일", "비고"])
        for r in (vmi_rows(year, month) if rows else []):
            v.append(r)
    if has_price_link:
        p = wb.create_sheet("계약단가")
        p.append(["자재코드", "자재명", "계약단가(원)", "적용시작일", "공급사"])
        for m in MATERIALS:
            p.append([m[0], m[1], m[4], "2025-12-01", m[3]])
    log = wb.create_sheet("개정이력")
    log.append(["개정번호", "개정일", "작성자", "개정 내용", "근거 자료"])
    for rev in REV_LOG[:rev_upto]:
        log.append(list(rev))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    if d:
        touch(path, d)


def build_generic_xlsx(path, sheet, headers, rows, d):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(headers)
    for r in rows:
        ws.append(r)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    touch(path, d)


def main():
    if os.path.exists(ROOT):
        shutil.rmtree(ROOT)
    po_dir = os.path.join(BASE, "구매", "PO정리")

    # ----- 월별 PO정리 파일 (전임자가 만든 이력) -----
    for (y, m, rev, name, d) in [
        (2025, 9, 2, "2025-09_오더PO정리.xlsx", date(2025, 9, 29)),
        (2025, 10, 2, "2025-10_오더PO정리.xlsx", date(2025, 10, 30)),
        (2025, 12, 3, "2025-12_오더PO정리.xlsx", date(2025, 12, 29)),
        (2026, 3, 4, "2026-03_오더PO정리.xlsx", date(2026, 3, 30)),
        (2026, 8, 5, "2026-08_오더PO정리.xlsx", date(2026, 8, 28)),
        (2026, 9, 5, "2026-09_오더PO정리_최종.xlsx", date(2026, 9, 28)),
    ]:
        build_po_workbook(os.path.join(po_dir, str(y), name), y, m, rev, po_rows(y, m, urgent=1 if rev >= 4 else 0), d=d)

    # 10월: '최종'과 '최종_수정'이 함께 있어 헷갈리는 상황
    oct_rows = po_rows(2026, 10, urgent=1)
    build_po_workbook(os.path.join(po_dir, "2026", "2026-10_오더PO정리_최종.xlsx"), 2026, 10, 5, oct_rows, d=date(2026, 10, 28))
    oct_rows2 = [list(r) for r in oct_rows]
    oct_rows2[3][10] = oct_rows2[3][4]  # 입고수량 정정
    build_po_workbook(os.path.join(po_dir, "2026", "2026-10_오더PO정리_최종_수정.xlsx"), 2026, 10, 5, oct_rows2,
                      d=date(2026, 10, 30), extra_note="※ 10/30 구매팀 요청으로 PO 452610004 입고수량 정정 후 재송부")

    # ----- 전임자 개인 폴더: 실제로 참고하던 파일이 여기 숨어 있음 -----
    build_generic_xlsx(
        os.path.join(BASE, "개인폴더", "이도현", "새 폴더 (2)", "PO정리_참고_체크리스트.xlsx"),
        "체크리스트",
        ["순서", "할 일", "참고", "메모"],
        [
            [1, "SAP ME2N에서 당월 구매오더 조회 후 엑셀 다운로드", "구매요청 업무표준 Rev.3", "플랜트 2100, 구매그룹 P21"],
            [2, "VMI 품목(탄산칼슘, 소포제, 백업필터)은 PO목록에서 빼고 VMI정산 시트로", "2025-07 개선회의", "VMI는 PO가 아니라 사용량으로 정산"],
            [3, "단가는 직접 입력하지 말 것. 계약단가 시트만 갱신", "공급사 단가 변경 메일", "단가 바뀌면 계약단가 시트 한 줄만 수정"],
            [4, "입고수량은 MES 입고실적과 대조", "", "입고율 90% 미만은 구매팀에 따로 말해줄 것"],
            [5, "말일 2영업일 전까지 구매팀 박지현 과장에게 메일 송부", "구매요청 업무표준 Rev.3", "파일명: YYYY-MM_오더PO정리.xlsx"],
        ],
        date(2026, 5, 12),
    )
    write_text(os.path.join("생산부", "조성파트", "개인폴더", "이도현", "메모.txt"), """
- 10월 PO정리: 구매팀이 입고수량 하나 틀렸다고 해서 '최종_수정'으로 다시 보냄. 최종본은 _수정 파일임
- 11월부터는 후임자가 할 예정. 인수인계서 작성하다 말았음 (인수인계 폴더)
""", date(2026, 10, 31))

    # ----- 표준서 -----
    write_text(os.path.join("생산부", "조성파트", "표준서", "구매요청_업무표준_Rev3.txt"), """
[가온페이퍼(가상) 업무표준] 구매요청 및 월간 오더 PO 정리
문서번호: PRD-STK-007 / 개정: Rev.3 / 개정일: 2026-02-10 / 주관: 생산부 조성파트

1. 목적: 조성파트가 발주한 원부자재·설비부품의 월간 구매오더(PO) 현황을 정리해 구매팀의 월말 정산에 쓴다.
2. 주기와 기한: 매월 1회, 말일 기준 2영업일 전까지 구매팀에 송부한다.
3. 절차
  3.1 SAP(ME2N)에서 당월 발주 건을 조회한다. (플랜트 2100, 구매그룹 P21)
  3.2 '오더PO정리' 양식의 PO목록 시트에 붙여넣는다.
  3.3 VMI 계약 품목은 PO목록에 넣지 않고 VMI정산 시트에 월 사용량으로 기재한다.
  3.4 단가는 계약단가 시트에서 자동으로 불러온다. 단가를 직접 입력하지 않는다.
  3.5 입고수량은 MES 입고실적과 대조한다.
  3.6 긴급발주 건은 비고에 '긴급 - 사유'를 적는다.
4. 송부처: 구매팀 월간 정산 담당
5. 개정 이력: Rev.1(2025-03) 제정 / Rev.2(2025-07) VMI 분리 반영 / Rev.3(2026-02) 긴급발주 비고 규칙 반영
""", date(2026, 2, 10))

    # ----- AI 회의록 (녹음 → AI 정리본 형태) -----
    write_text(os.path.join("생산부", "조성파트", "회의록", "2025-07-14_구매프로세스_개선회의_AI정리.txt"), """
[AI 회의록 자동 정리] 구매 프로세스 개선 회의
일시: 2025-07-14 14:00 / 참석: 생산부 조성파트 정민수 차장, 이도현 사원, 구매팀 박지현 과장

■ 논의 요약
- 탄산칼슘·소포제·백업필터는 VMI(공급사 재고관리) 계약으로 바뀌어 PO 없이 사용량으로 정산됨.
- 그런데 월간 PO정리 파일에 VMI 품목이 섞여 있어 구매팀 정산 때 이중 계상이 2건 발생함.

■ 결정 사항
- [결정] 오더PO정리 파일에서 VMI 품목을 빼고 별도 시트(VMI정산)로 분리한다. 담당: 이도현
- [결정] VMI 품목은 월말 사용량 기준으로 정산한다.

■ 할 일
- 이도현: 7월분부터 양식 변경 (7/15까지)
""", date(2025, 7, 14))

    write_text(os.path.join("생산부", "조성파트", "회의록", "2026-02-03_월간생산회의_AI정리.txt"), """
[AI 회의록 자동 정리] 2월 월간 생산회의
일시: 2026-02-03 09:30 / 참석: 생산부장, 조성파트 정민수 차장, 이도현 사원, 품질팀

■ 논의 요약
- 1월 PM2 펌프 고장 때 긴급발주 3건이 있었는데, 구매팀이 사유를 몰라 정산이 늦어짐.

■ 결정 사항
- [결정] 긴급발주 건은 오더PO정리 비고란에 '긴급 - 사유' 형식으로 적는다.
- [결정] 업무표준 PRD-STK-007에도 반영한다. (Rev.3)

■ 기타
- 원단위 개선 건은 다음 회의로 이월
""", date(2026, 2, 3))

    write_text(os.path.join("생산부", "조성파트", "회의록", "2026-04-07_설비보전_주간회의_AI정리.txt"), """
[AI 회의록 자동 정리] 설비보전 주간회의
일시: 2026-04-07 / 참석: 보전팀, 조성파트

■ 논의 요약
- 2호기 정선기 진동 증가, 베어링 교체 검토
- 와이어 세척 노즐 막힘 빈도 증가

■ 결정 사항
- [결정] 베어링 6312 예비품 2EA 추가 발주
""", date(2026, 4, 7))

    # ----- 메일 보관 -----
    write_text(os.path.join("생산부", "조성파트", "메일보관", "2025-11-20_한결케미칼_계약단가변경안내.txt"), """
보낸사람: 한결케미칼 영업팀
받는사람: 가온페이퍼 구매팀 박지현 과장 / 참조: 생산부 이도현
제목: [안내] 2025년 12월 1일부 계약단가 변경

12월 1일부터 AKD 사이즈제, 보류제 A, 소포제의 계약단가가 약 5% 인상됩니다. 세부 단가표를 첨부합니다.

--- 회신 (이도현) ---
단가가 바뀔 때마다 월별 파일마다 고치다 보니 틀리는 일이 생겨서,
앞으로는 PO정리 파일에 '계약단가' 시트를 두고 VLOOKUP으로 불러오게 바꾸겠습니다.
""", date(2025, 11, 20))

    write_text(os.path.join("생산부", "조성파트", "메일보관", "2026-10-30_구매팀_10월PO_입고수량정정요청.txt"), """
보낸사람: 구매팀 박지현 과장
받는사람: 생산부 이도현
제목: 10월 오더PO정리 입고수량 확인 부탁드립니다

PO 452610004 입고수량이 MES 입고실적과 다릅니다. 확인 후 다시 보내주세요.

--- 회신 (이도현) ---
확인했습니다. 수정본(2026-10_오더PO정리_최종_수정.xlsx)으로 다시 보냅니다.
""", date(2026, 10, 30))

    # ----- 인수인계 폴더: 작성하다 만 인수인계서 -----
    write_text(os.path.join("생산부", "조성파트", "인수인계", "인수인계서_이도현_작성중.txt"), """
인수인계서 (작성 중)
인계자: 이도현 / 인수자: (미정)

1. 담당 업무
  - 월간 오더 PO정리 (매월 말)
  - 용수 사용량 집계
  - (작성 예정)
2. 참고 파일 위치
  - (작성 예정)
""", date(2026, 10, 31))

    # ----- 업무와 무관한 파일들 (파일 서버가 방대하다는 걸 보여주는 잡음) -----
    noise = [
        (os.path.join("안전", "2026_상반기_위험성평가.xlsx"), "위험성평가", ["공정", "유해위험요인", "위험성", "감소대책"],
         [["조성", "약품 투입 시 비산", "중", "보안경 착용"], ["초지", "와이어 교체 중 협착", "상", "LOTO 실시"]], date(2026, 6, 20)),
        (os.path.join("원단위", "2026_원단위_실적.xlsx"), "원단위", ["월", "지종", "펄프 원단위", "약품 원단위"],
         [["2026-01", "백상지", 0.82, 0.031], ["2026-02", "백상지", 0.81, 0.030]], date(2026, 3, 5)),
        (os.path.join("용수", "2026_용폐수_집계.xlsx"), "용수", ["월", "용수(톤)", "폐수(톤)"],
         [["2026-01", 182000, 151000], ["2026-02", 176000, 149500]], date(2026, 3, 3)),
        (os.path.join("설비", "PM2_정선기_점검일지.xlsx"), "점검", ["일자", "설비", "점검내용", "조치"],
         [["2026-04-02", "정선기", "진동 증가", "베어링 교체 예정"]], date(2026, 4, 2)),
        (os.path.join("구매", "견적서", "2026-05_동진기계_견적.xlsx"), "견적", ["품목", "수량", "견적가"],
         [["펌프 메카니컬씰", 2, 390000]], date(2026, 5, 14)),
        (os.path.join("구매", "PO정리", "2026", "복사본 - 2026-08_오더PO정리.xlsx"), "Sheet1", ["(빈 파일)"], [], date(2026, 8, 28)),
        (os.path.join("교육", "신입사원_안전교육_자료.xlsx"), "교육", ["항목", "내용"],
         [["보호구", "안전모·안전화 필수"]], date(2026, 1, 15)),
    ]
    for rel, sheet, h, rows, d in noise:
        build_generic_xlsx(os.path.join(BASE, rel), sheet, h, rows, d)
    write_text(os.path.join("생산부", "조성파트", "기타", "새 텍스트 문서.txt"), "test", date(2025, 6, 1))
    write_text(os.path.join("생산부", "초지파트", "운전일지", "2026-10_PM1_운전일지_요약.txt"), "10월 PM1 운전 요약: 단선 3회", date(2026, 10, 31))
    write_text(os.path.join("품질팀", "고객클레임", "2026-09_클레임_대응.txt"), "9월 클레임 2건 대응 완료", date(2026, 9, 30))

    # ----- 쪽지에 첨부된 이번 달 파일 (신입이 받은 파일) -----
    inbox = os.path.join(ROOT, "inbox")
    build_po_workbook(os.path.join(inbox, "2026-11_오더PO정리.xlsx"), 2026, 11, 5, rows=[], d=date(2026, 11, 18))
    with open(os.path.join(ROOT, "file_meta.json"), "w", encoding="utf-8") as f:
        json.dump(META, f, ensure_ascii=False, indent=1)
    print("완료:", ROOT, len(META), "files")


if __name__ == "__main__":
    main()
