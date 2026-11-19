"""완생 에이전트의 두뇌: 파일 서버 색인, AI가 쓰는 도구, 도구 호출 루프.

Streamlit 없이도 돌아가게 분리해 두었다. (테스트·로컬 AI 교체가 쉬움)
"""
import json
import os
import re
from dataclasses import dataclass, field

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
FS_ROOT = os.path.join(DATA, "fileserver")
SERVER = r"\\gaon-fs01"
CONFIRM_DIR = SERVER + r"\생산부\조성파트\확인답변"
UPLOAD_PREFIX = "업로드/"

# data 폴더가 GitHub에 안 올라갔어도 동작하도록, 없으면 가상 데이터를 그 자리에서 만든다
_REQUIRED = [os.path.join(DATA, "file_meta.json"),
             os.path.join(DATA, "inbox", "2026-11_오더PO정리.xlsx"),
             os.path.join(FS_ROOT, "생산부", "조성파트", "구매", "PO정리"),
             # v2 데이터(업무 흐름 메일)가 없으면 예전 data 폴더이므로 새로 만든다
             os.path.join(FS_ROOT, "생산부", "조성파트", "메일보관", "2026-10-06_생산계획팀_PO정리_활용안내.txt")]
if not all(os.path.exists(p) for p in _REQUIRED):  # 일부만 올라간 경우도 통째로 다시 만든다
    import make_data
    make_data.main()

try:
    with open(os.path.join(DATA, "file_meta.json"), encoding="utf-8") as _f:
        FILE_META = json.load(_f)
except FileNotFoundError:
    FILE_META = {}


def to_vpath(rel):
    return SERVER + "\\" + rel.replace("/", "\\")


# ------------------------------------------------------------------ 파일 읽기
def describe_xlsx(path, max_rows=8):
    """엑셀 구조를 사람이 읽을 수 있는 글로 바꾼다: 시트, 헤더, 수식, 조건부서식, 개정이력."""
    wb = load_workbook(path)  # 수식 문자열을 보기 위해 data_only=False
    out = [f"시트 {len(wb.sheetnames)}개: {', '.join(wb.sheetnames)}"]
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        rows = [r for r in rows if any(v is not None for v in r)]
        out.append(f"\n[시트: {ws.title}] 데이터 {max(len(rows) - 1, 0)}행")
        if not rows:
            continue
        headers = [str(h) for h in rows[0] if h is not None]
        out.append("열: " + " | ".join(headers))
        limit = len(rows) if ws.title in ("개정이력", "체크리스트", "계약단가") else max_rows
        for r in rows[1:limit + 1]:
            out.append("  - " + " | ".join("" if v is None else str(v) for v in r))
        if len(rows) - 1 > limit:
            out.append(f"  ... (이하 {len(rows) - 1 - limit}행 생략)")
        formulas = {}
        for row in ws.iter_rows(min_row=2, max_row=min(ws.max_row, 3)):
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    head = ws.cell(1, c.column).value
                    formulas.setdefault(head, c.value)
        if formulas:
            out.append("수식: " + "; ".join(f"{k} = {v}" for k, v in formulas.items()))
        for rng, rules in ws.conditional_formatting._cf_rules.items():
            for rule in rules:
                out.append(f"조건부서식: {rng.sqref} 범위, {rule.operator} {rule.formula} 이면 강조 표시")
    return "\n".join(out)


def read_text(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


@dataclass
class Doc:
    vpath: str
    name: str
    saved: str
    kind: str
    text: str


def load_doc(abs_path, vpath, saved):
    kind = "엑셀" if abs_path.endswith(".xlsx") else "문서"
    try:
        text = describe_xlsx(abs_path) if kind == "엑셀" else read_text(abs_path)
    except Exception as e:  # 깨진 파일도 목록에는 남긴다
        text = f"(읽기 실패: {e})"
    return Doc(vpath, os.path.basename(abs_path), saved, kind, text)


def build_index():
    docs = []
    for dirpath, _, files in os.walk(FS_ROOT):
        for fn in sorted(files):
            ap = os.path.join(dirpath, fn)
            rel = os.path.relpath(ap, FS_ROOT).replace(os.sep, "/")
            saved = FILE_META.get("fileserver/" + rel, "")
            docs.append(load_doc(ap, to_vpath(rel), saved))
    return docs


# ------------------------------------------------------------------ 세션 상태
MEETING_DIR = SERVER + r"\생산부\조성파트\회의록"


@dataclass
class Workspace:
    """심사위원 한 명의 세션. 다른 사람의 확인 답변과 섞이지 않게 세션마다 따로 둔다."""
    docs: list
    uploads: dict = field(default_factory=dict)        # 이름 -> Doc
    upload_paths: dict = field(default_factory=dict)   # 이름 -> 세션 전용 엑셀 사본 경로
    confirmations: list = field(default_factory=list)  # {"id","question","reason","status","answer","kind"}
    updates: list = field(default_factory=list)        # 회의록 → 개정이력 반영 기록 (신입에게 보여줄 '업무 업데이트')

    def all_docs(self):
        return self.docs + list(self.uploads.values())

    def ensure_copy(self, name, abs_path):
        """원본을 건드리지 않도록 세션 전용 사본을 만든다. (개정이력 반영도 이 사본에만 한다)"""
        import shutil
        import tempfile
        if name not in self.upload_paths:
            tmpdir = tempfile.mkdtemp(prefix="wanseng_")
            copy_path = os.path.join(tmpdir, name)
            shutil.copy(abs_path, copy_path)
            self.upload_paths[name] = copy_path
        return self.upload_paths[name]

    def add_upload(self, name, abs_path):
        self.ensure_copy(name, abs_path)
        d = load_doc(self.upload_paths[name], UPLOAD_PREFIX + name, "방금 받음")
        self.uploads[name] = d
        return d

    def revision_rows(self, name):
        """업로드한 엑셀의 개정이력 시트를 [{개정번호,...}] 로 돌려준다."""
        p = self.upload_paths.get(name)
        if not p:
            return []
        wb = load_workbook(p)
        if "개정이력" not in wb.sheetnames:
            return []
        rows = [r for r in wb["개정이력"].iter_rows(values_only=True) if any(v is not None for v in r)]
        if not rows:
            return []
        head = [str(h) for h in rows[0]]
        return [dict(zip(head, r)) for r in rows[1:]]

    def apply_meeting(self, name, title, meeting_date, minutes, decisions, approver):
        """선임이 승인한 회의 결정을 ①회의록 파일로 저장하고 ②업로드 엑셀 개정이력에 Rev를 추가하고
        ③신입에게 보여줄 업무 업데이트로 남긴다. decisions: [{"change","reason"}]"""
        slug = re.sub(r"[^0-9A-Za-z가-힣]+", "_", title).strip("_")[:30]
        mdoc = Doc(MEETING_DIR + f"\\{meeting_date}_{slug}_AI정리.txt", f"{meeting_date}_{slug}_AI정리.txt",
                   meeting_date, "문서", minutes.strip())
        self.docs = [d for d in self.docs if d.vpath != mdoc.vpath] + [mdoc]
        added = []
        p = self.upload_paths.get(name)
        if p:
            wb = load_workbook(p)
            log = wb["개정이력"] if "개정이력" in wb.sheetnames else wb.create_sheet("개정이력")
            if log.max_row < 1 or log.cell(1, 1).value is None:
                log.append(["개정번호", "개정일", "작성자", "개정 내용", "근거 자료"])
            nums = [int(m.group(1)) for r in log.iter_rows(min_row=2, values_only=True)
                    for m in [re.match(r"Rev\.(\d+)", str(r[0] or ""))] if m]
            n = max(nums or [0])
            for dcs in decisions:
                n += 1
                row = [f"Rev.{n}", meeting_date, f"{approver} 승인",
                       dcs["change"].strip(), f"{meeting_date} {title} (이유: {dcs.get('reason', '').strip() or '회의록 참조'})"]
                log.append(row)
                added.append(dict(zip(["개정번호", "개정일", "작성자", "개정 내용", "근거 자료"], row)))
            wb.save(p)
            if name in self.uploads:  # 신입이 이미 넣은 파일이면 최신 개정이력으로 다시 읽는다
                self.uploads[name] = load_doc(p, UPLOAD_PREFIX + name, "방금 받음")
        self.updates.append({"date": meeting_date, "title": title, "approver": approver,
                             "revisions": added, "minutes_path": mdoc.vpath})
        return added, mdoc

    def add_review_request(self, item, question, why):
        """위험도가 높은 항목을 선임 검토 요청으로 보낸다. (근거가 있어도 확인받는 항목)"""
        for c in self.confirmations:
            if c["question"] == question:
                return c
        c = {"id": len(self.confirmations) + 1, "question": question, "reason": why,
             "status": "답변 대기", "answer": "", "kind": f"검토 · {item}"}
        self.confirmations.append(c)
        return c

    def add_confirmation_answer(self, item, answer, who):
        item["status"] = "답변 완료"
        item["answer"] = answer
        text = (f"[담당자 확인 답변]\n질문: {item['question']}\n답변자: {who}\n답변: {answer}\n"
                f"(완생 에이전트의 확인 요청으로 기록됨)")
        slug = re.sub(r"[^0-9A-Za-z가-힣]+", "_", item["question"])[:30]
        d = Doc(CONFIRM_DIR + f"\\2026-11-18_{slug}.txt", f"{slug}.txt", "2026-11-18", "문서", text)
        self.docs.append(d)
        return d


# ------------------------------------------------------------------ 도구
def _tokens(q):
    return [t for t in re.split(r"[\s,./\\_\-()\[\]·:'\"?]+", q.lower()) if len(t) >= 2]


def search_files(ws, query, top_k=8):
    """키워드 검색. 여러 파일에 흔한 단어는 가중치를 낮춰(IDF) 드문 근거 자료가 묻히지 않게 한다."""
    import math
    toks = _tokens(query)
    docs = ws.all_docs()
    n = len(docs)
    blobs = [(d, d.name.lower(), d.vpath.lower(), d.text.lower()) for d in docs]
    idf = {}
    for t in toks:
        df = sum(1 for _, nm, pa, bo in blobs if t in nm or t in pa or t in bo)
        idf[t] = math.log((n + 1) / (df + 1)) + 0.3
    scored = []
    for d, name, path, body in blobs:
        s = 0.0
        for t in toks:
            hit = 0
            if t in name:
                hit += 3
            elif t in path:
                hit += 2
            if t in body:
                hit += min(body.count(t), 3)
            s += hit * idf[t]
        if s:
            scored.append((s, d))
    # 같은 폴더의 비슷한 월별 파일이 결과를 다 차지하지 않도록, 같은 폴더 2번째부터 점수를 깎는다
    scored.sort(key=lambda x: (-x[0], x[1].vpath))
    seen, ranked = {}, []
    for sc, d in scored:
        folder = d.vpath.rsplit("\\", 1)[0]
        k = seen.get(folder, 0)
        seen[folder] = k + 1
        ranked.append((sc * (0.45 ** k), d))
    ranked.sort(key=lambda x: (-x[0], x[1].vpath))
    hits = []
    for s, d in ranked[:top_k]:
        first = next((ln.strip() for ln in d.text.splitlines() if ln.strip()), "")
        hits.append({"경로": d.vpath, "저장일": d.saved, "종류": d.kind, "미리보기": first[:80]})
    return {"검색어": query, "결과수": len(scored), "상위결과": hits}


def read_file(ws, path):
    p = path.strip().strip('"')
    for d in ws.all_docs():
        if d.vpath == p or d.vpath.endswith(p) or d.name == p:
            return {"경로": d.vpath, "저장일": d.saved, "내용": d.text[:6000]}
    return {"오류": f"'{path}' 파일을 찾지 못했습니다. search_files로 정확한 경로를 먼저 찾으세요."}


def list_folder(ws, path):
    p = path.rstrip("\\")
    children = set()
    for d in ws.docs:
        if d.vpath.startswith(p + "\\"):
            rest = d.vpath[len(p) + 1:]
            children.add(rest.split("\\")[0] + ("\\" if "\\" in rest else ""))
    if not children:
        return {"오류": f"'{path}' 폴더가 비어 있거나 없습니다."}
    return {"폴더": p, "항목": sorted(children)}


def request_confirmation(ws, question, reason):
    for c in ws.confirmations:
        if c["question"] == question:
            return {"상태": "이미 요청됨", "요청번호": c["id"]}
    item = {"id": len(ws.confirmations) + 1, "question": question, "reason": reason,
            "status": "답변 대기", "answer": "", "kind": "근거 없음 · 확인"}
    ws.confirmations.append(item)
    return {"상태": "담당자 확인 요청 등록", "요청번호": item["id"]}


TOOLS = [
    {"type": "function", "function": {
        "name": "search_files",
        "description": "회사 파일 서버(\\\\gaon-fs01)에서 파일 이름·경로·내용을 키워드로 검색한다. 한국어 핵심 단어 2~4개를 공백으로 구분해 넣는다.",
        "parameters": {"type": "object", "properties": {"query": {"type": "string"}}, "required": ["query"]}}},
    {"type": "function", "function": {
        "name": "read_file",
        "description": "파일 하나의 내용을 읽는다. 엑셀은 시트·열·수식·조건부서식·개정이력을 글로 풀어서 준다. 업로드된 파일은 '업로드/파일명' 경로.",
        "parameters": {"type": "object", "properties": {"path": {"type": "string"}}, "required": ["path"]}}},
    {"type": "function", "function": {
        "name": "list_folder",
        "description": "파일 서버 폴더 안의 하위 폴더와 파일 목록을 본다. 예: \\\\gaon-fs01\\생산부\\조성파트",
        "parameters": {"type": "object", "properties": {"path": {"type": "string"}}, "required": ["path"]}}},
    {"type": "function", "function": {
        "name": "request_confirmation",
        "description": "파일 서버 어디에서도 근거를 찾지 못한 '왜' 질문을 담당자에게 확인 요청으로 등록한다. 근거가 있으면 쓰지 않는다.",
        "parameters": {"type": "object", "properties": {
            "question": {"type": "string", "description": "담당자에게 물을 질문 한 문장"},
            "reason": {"type": "string", "description": "어떤 자료를 찾아봤는데 근거가 없었는지"}},
            "required": ["question", "reason"]}}},
]

FUNCS = {"search_files": search_files, "read_file": read_file,
         "list_folder": list_folder, "request_confirmation": request_confirmation}

_LEVEL = {"type": "string", "enum": ["근거 있음", "추정", "알 수 없음"]}
_STEP = {"type": "object", "properties": {
    "who": {"type": "string", "description": "부서·사람 (예: 구매팀 박지현 과장)"},
    "what": {"type": "string", "description": "그 단계에서 하는 일 한 줄"},
    "source": {"type": "string", "description": "근거 파일 경로. 없으면 빈 문자열"},
    "level": _LEVEL}, "required": ["who", "what", "level"]}

SUBMIT_TOOL = {"type": "function", "function": {
    "name": "submit_onboarding_report",
    "description": "조사가 끝나면 마지막에 딱 한 번 호출해 신입에게 보여줄 업무 안내서를 제출한다. 모든 내용은 도구로 확인한 것만 쓴다.",
    "parameters": {"type": "object", "properties": {
        "file_summary": {"type": "string", "description": "받은 파일이 무엇인지 1~2문장"},
        "flow": {"type": "object", "description": "이 일의 전후 사정: 앞 단계 → 나 → 뒤 단계", "properties": {
            "before": {"type": "array", "items": _STEP, "description": "이 일이 시작되기 전 단계(누가 무엇을 넘겨주나)"},
            "me": {"type": "object", "properties": {
                "what": {"type": "string"}, "due": {"type": "string"}}, "required": ["what"]},
            "after": {"type": "array", "items": _STEP, "description": "결과물을 받는 곳과 그 쓰임새"},
            "purpose": {"type": "string", "description": "그래서 이 일을 왜 하는지 한 문장"}},
            "required": ["before", "me", "after", "purpose"]},
        "risks": {"type": "array", "description": "틀리면 어떻게 되나: 항목별 영향", "items": {"type": "object", "properties": {
            "item": {"type": "string", "description": "칸·시트·작업 이름 (예: 입고수량)"},
            "impact": {"type": "string", "description": "틀리면 어디서 무슨 문제가 생기나"},
            "risk": {"type": "string", "enum": ["높음", "보통", "낮음"]},
            "source": {"type": "string"}, "level": _LEVEL},
            "required": ["item", "impact", "risk", "level"]}},
        "reasons": {"type": "array", "description": "왜 이렇게 되어 있나: 개정마다 하나", "items": {"type": "object", "properties": {
            "rev": {"type": "string"}, "change": {"type": "string"}, "why": {"type": "string"},
            "source": {"type": "string"}, "level": _LEVEL}, "required": ["rev", "change", "why", "level"]}},
        "todos": {"type": "array", "description": "무엇을 하면 되나: 순서대로", "items": {"type": "object", "properties": {
            "task": {"type": "string"}, "due": {"type": "string"}, "source": {"type": "string"}}, "required": ["task"]}},
        "reviews": {"type": "array", "description": "무엇을 확인받아야 하나: 위험도 '높음' 항목과 근거가 없는 '왜'", "items": {"type": "object", "properties": {
            "item": {"type": "string"}, "question": {"type": "string", "description": "선임에게 보낼 질문 한 문장"},
            "why": {"type": "string", "description": "왜 확인받아야 하는지"}}, "required": ["item", "question", "why"]}},
        "previous_files": {"type": "array", "items": {"type": "string"}, "description": "전임자의 이전 파일 경로"},
    }, "required": ["file_summary", "flow", "risks", "reasons", "todos", "reviews"]}}}

SYSTEM_PROMPT = """당신은 '완생'입니다. 제조회사 가온페이퍼(가상)의 사내 AI 에이전트로, 인수인계 없이 업무를 받은 신입사원이
'이 일을 왜 하는지'를 알고 일하도록 돕습니다. AI가 일을 대신하지 않고, 신입이 빠르게 파악하고 중요한 부분은 선임에게 검토받게 합니다.
오늘은 2026-11-18입니다. 회사 파일 서버 주소는 \\\\gaon-fs01 입니다.

[일하는 방식]
1. 답하기 전에 반드시 도구로 확인한다. 업로드 파일이 있으면 먼저 read_file로 읽고, search_files로 관련 파일·회의록·메일·표준서를 찾고, 중요한 파일은 read_file로 직접 연다.
2. 경로·날짜·사람·수치는 도구 결과에 나온 것만 쓴다. 지어내지 않는다.
3. '왜 이렇게 되어 있는지'는 개정이력의 '근거 자료' 칸, 회의록의 [결정], 메일, 표준서에서 찾는다.
4. 모든 주장에 근거 수준을 붙인다.
   - [근거 있음] 파일에 직접 적혀 있음 → (출처: 파일 경로)
   - [추정] 직접 기록은 없지만 여러 자료로 미루어 짐작함 → 무엇을 보고 짐작했는지 밝힘
   - [알 수 없음] 찾아봤지만 기록이 없음 → request_confirmation 도구로 담당자 확인 요청을 등록한다
   - 이전에 등록된 '확인답변' 파일이 있으면 그것을 근거로 [근거 있음]으로 답한다
5. 같은 달 파일이 여러 개면(예: 최종, 최종_수정) 메모·메일로 어느 것이 최종본인지 확인한다.
6. 위험도: 틀리면 정산·발주·원가·품질에 영향을 주는 칸은 '높음'이다. '높음' 항목은 근거가 있어도 선임 검토 대상(reviews)에 넣는다.

후속 질문에는 짧게 답하고, 문장마다 [근거 있음]/[추정]/[알 수 없음]을 붙인다. 한국어로, 신입이 바로 따라 할 수 있게 쓴다."""

ONBOARD_RULE = """
[이번 질문은 '처음 업무를 받은 상황'이다]
조사 순서: ①업로드 파일 읽기 ②표준서·메일에서 이 일의 앞 단계(누가 넘겨주나)와 뒤 단계(누가 받아 무엇에 쓰나) 찾기
③개정이력의 근거 자료(회의록·메일) 열기 ④전임자 개인 폴더·메모 확인.
조사가 끝나면 일반 답변 대신 submit_onboarding_report 도구를 꼭 한 번 호출해 결과를 제출한다.
flow.after에는 결과물을 받는 모든 곳을 넣고, risks는 3~6개, reviews에는 위험도 '높음' 항목과 근거가 없는 질문을 넣는다."""


def run_agent(ws, history, user_msg, chat_fn, max_steps=10, onboarding=False):
    """chat_fn(messages, tools) -> {"content": str|None, "tool_calls": [{"id","name","arguments"}]}
    반환: (최종 답변 글, 도구 호출 기록, 갱신된 history, 업무 안내서 dict 또는 None)"""
    system = SYSTEM_PROMPT + (ONBOARD_RULE if onboarding else "")
    tools = TOOLS + ([SUBMIT_TOOL] if onboarding else [])
    messages = [{"role": "system", "content": system}] + history + [{"role": "user", "content": user_msg}]
    trace = []
    for step in range(max_steps):
        msg = chat_fn(messages, tools)
        calls = msg.get("tool_calls") or []
        if not calls:
            answer = msg.get("content") or "(답변을 만들지 못했습니다)"
            new_history = history + [{"role": "user", "content": user_msg},
                                     {"role": "assistant", "content": answer}]
            return answer, trace, new_history, None
        # Gemini 3 계열은 도구 호출에 붙은 서명(thought signature)을 다시 돌려줘야 하므로, 받은 그대로(raw) 되돌려 보낸다
        messages.append({"role": "assistant", "content": msg.get("content") or "",
                         "tool_calls": [c.get("raw") or {"id": c["id"], "type": "function",
                                         "function": {"name": c["name"], "arguments": c["arguments"]}} for c in calls]})
        for c in calls:
            if c["name"] == "submit_onboarding_report":
                try:
                    report = json.loads(c["arguments"] or "{}")
                except json.JSONDecodeError:
                    report = None
                if report:
                    summary = report_to_text(report)
                    new_history = history + [{"role": "user", "content": user_msg},
                                             {"role": "assistant", "content": summary}]
                    return summary, trace, new_history, report
                messages.append({"role": "tool", "tool_call_id": c["id"], "content": "형식 오류: JSON으로 다시 제출하세요."})
                continue
            try:
                args = json.loads(c["arguments"] or "{}")
                result = FUNCS[c["name"]](ws, **args)
            except Exception as e:
                args, result = {}, {"오류": str(e)}
            trace.append({"tool": c["name"], "args": args, "result": result})
            messages.append({"role": "tool", "tool_call_id": c["id"],
                             "content": json.dumps(result, ensure_ascii=False)})
        if onboarding and step == max_steps - 2:
            messages.append({"role": "user", "content": "이제 조사를 멈추고 submit_onboarding_report로 지금까지 확인한 내용을 제출하세요."})
    return "도구 호출이 너무 길어져 멈췄습니다. 질문을 조금 더 좁혀 주세요.", trace, history, None


def report_to_text(r):
    """대화 기록(후속 질문의 맥락)용으로 업무 안내서를 짧은 글로 바꾼다."""
    f = r.get("flow", {})
    lines = [f"[업무 안내서] {r.get('file_summary', '')}",
             "흐름: " + " / ".join(f"{s.get('who')}: {s.get('what')}" for s in f.get("before", []))
             + f" → 나: {f.get('me', {}).get('what', '')} → "
             + " / ".join(f"{s.get('who')}: {s.get('what')}" for s in f.get("after", [])),
             "목적: " + f.get("purpose", "")]
    lines += [f"위험 {x.get('item')}({x.get('risk')}): {x.get('impact')}" for x in r.get("risks", [])]
    lines += [f"{x.get('rev')} {x.get('change')} - 이유: {x.get('why')} [{x.get('level')}]" for x in r.get("reasons", [])]
    lines += [f"할 일: {x.get('task')} ({x.get('due', '')})" for x in r.get("todos", [])]
    lines += [f"검토: {x.get('item')} - {x.get('question')}" for x in r.get("reviews", [])]
    return "\n".join(lines)


# ------------------------------------------------------------------ 회의록 → 결정 사항 추출
def extract_decisions_rule(minutes):
    """AI 없이도 동작하는 기본 추출: '[결정]'으로 시작하는 줄을 뽑는다."""
    out = []
    for ln in minutes.splitlines():
        m = re.search(r"\[결정\]\s*(.+)", ln)
        if m:
            out.append({"change": m.group(1).strip(), "reason": ""})
    return out


def extract_decisions_ai(minutes, client, model):
    """회의록에서 업무 양식·기준을 바꾸는 결정과 그 이유를 뽑는다. 실패하면 규칙 기반으로 돌아간다."""
    prompt = ("다음은 AI로 정리한 회의록이다. 월간 '오더PO정리' 업무의 양식·기준·절차를 바꾸는 결정만 골라라. "
              "각 결정마다 change(무엇이 바뀌는지 한 문장)와 reason(회의록에 적힌 이유, 없으면 빈 문자열)을 쓴다. "
              "회의록에 없는 내용은 지어내지 않는다. JSON 객체 {\"decisions\": [...]} 로만 답한다.\n\n" + minutes)
    try:
        r = client.chat.completions.create(model=model, temperature=0,
                                           response_format={"type": "json_object"},
                                           messages=[{"role": "user", "content": prompt}])
        data = json.loads(r.choices[0].message.content or "{}")
        dec = [d for d in data.get("decisions", []) if d.get("change")]
        return dec or extract_decisions_rule(minutes), "AI"
    except Exception:
        return extract_decisions_rule(minutes), "규칙"


def openai_chat_fn(client, model):
    """OpenAI 호환 API(OpenAI, Gemini 호환 주소, 로컬 Ollama 등) 어디든 같은 코드로 붙는다."""
    def _chat(messages, tools):
        r = client.chat.completions.create(model=model, messages=messages, tools=tools, temperature=0.2)
        m = r.choices[0].message
        calls = [{"id": t.id, "name": t.function.name, "arguments": t.function.arguments,
                  "raw": t.model_dump(exclude_none=True)} for t in (m.tool_calls or [])]
        return {"content": m.content, "tool_calls": calls}
    return _chat
