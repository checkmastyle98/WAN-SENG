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
             os.path.join(FS_ROOT, "생산부", "조성파트", "구매", "PO정리")]
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
@dataclass
class Workspace:
    """심사위원 한 명의 세션. 다른 사람의 확인 답변과 섞이지 않게 세션마다 따로 둔다."""
    docs: list
    uploads: dict = field(default_factory=dict)        # 이름 -> Doc
    confirmations: list = field(default_factory=list)  # {"id","question","reason","status","answer"}

    def all_docs(self):
        return self.docs + list(self.uploads.values())

    def add_upload(self, name, abs_path):
        d = load_doc(abs_path, UPLOAD_PREFIX + name, "방금 받음")
        self.uploads[name] = d
        return d

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
            "status": "답변 대기", "answer": ""}
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

SYSTEM_PROMPT = """당신은 '완생'입니다. 제조회사 가온페이퍼(가상)의 사내 AI 에이전트로, 인수인계 없이 업무를 받은 신입사원이 전임자의 업무 맥락을 되짚도록 돕습니다.
오늘은 2026-11-18입니다. 회사 파일 서버 주소는 \\\\gaon-fs01 입니다.

[일하는 방식]
1. 답하기 전에 반드시 도구로 확인한다. 업로드 파일이 있으면 먼저 read_file로 읽고, search_files로 관련 파일·회의록·메일·표준서를 찾고, 중요한 파일은 read_file로 직접 연다.
2. 경로·날짜·사람·수치는 도구 결과에 나온 것만 쓴다. 지어내지 않는다.
3. '왜 이렇게 되어 있는지'는 개정이력의 '근거 자료' 칸, 회의록의 [결정], 메일, 표준서에서 찾는다.
4. 모든 주장 앞에 근거 수준을 붙인다.
   - [근거 있음] 파일에 직접 적혀 있음 → 문장 끝에 (출처: 파일 경로)
   - [추정] 직접 기록은 없지만 여러 자료로 미루어 짐작함 → 무엇을 보고 짐작했는지 밝힘
   - [알 수 없음] 찾아봤지만 기록이 없음 → request_confirmation 도구로 담당자 확인 요청을 등록하고 그 사실을 알림
   - [추정]인데 업무 판단에 영향이 크면 역시 request_confirmation으로 확인 요청을 등록한다
   - 이전에 등록된 '확인답변' 파일이 있으면 그것을 근거로 [근거 있음]으로 답한다
5. 같은 달 파일이 여러 개면(예: 최종, 최종_수정) 메모·메일로 어느 것이 최종본인지 확인한다.

[신입이 처음 파일을 받았을 때 답변 형식]
### 1. 이 파일은 무엇인가요
### 2. 전임자의 이전 파일은 어디에 있나요 (경로 목록)
### 3. 개정 이력과 그 이유 (개정마다 근거 수준 표시)
### 4. 이번에 할 일 (기한 포함 체크리스트)

후속 질문에는 위 형식 없이 짧게 답한다. 한국어로, 신입이 바로 따라 할 수 있게 구체적으로 쓴다."""


# ------------------------------------------------------------------ 에이전트 루프
def run_agent(ws, history, user_msg, chat_fn, max_steps=8):
    """chat_fn(messages, tools) -> {"content": str|None, "tool_calls": [{"id","name","arguments"}]}
    반환: (최종 답변, 도구 호출 기록, 갱신된 history)"""
    messages = [{"role": "system", "content": SYSTEM_PROMPT}] + history + [{"role": "user", "content": user_msg}]
    trace = []
    for _ in range(max_steps):
        msg = chat_fn(messages, TOOLS)
        calls = msg.get("tool_calls") or []
        if not calls:
            answer = msg.get("content") or "(답변을 만들지 못했습니다)"
            new_history = history + [{"role": "user", "content": user_msg},
                                     {"role": "assistant", "content": answer}]
            return answer, trace, new_history
        # Gemini 3 계열은 도구 호출에 붙은 서명(thought signature)을 다시 돌려줘야 하므로, 받은 그대로(raw) 되돌려 보낸다
        messages.append({"role": "assistant", "content": msg.get("content") or "",
                         "tool_calls": [c.get("raw") or {"id": c["id"], "type": "function",
                                         "function": {"name": c["name"], "arguments": c["arguments"]}} for c in calls]})
        for c in calls:
            try:
                args = json.loads(c["arguments"] or "{}")
                result = FUNCS[c["name"]](ws, **args)
            except Exception as e:
                args, result = {}, {"오류": str(e)}
            trace.append({"tool": c["name"], "args": args, "result": result})
            messages.append({"role": "tool", "tool_call_id": c["id"],
                             "content": json.dumps(result, ensure_ascii=False)})
    return "도구 호출이 너무 길어져 멈췄습니다. 질문을 조금 더 좁혀 주세요.", trace, history


def openai_chat_fn(client, model):
    """OpenAI 호환 API(OpenAI, Gemini 호환 주소, 로컬 Ollama 등) 어디든 같은 코드로 붙는다."""
    def _chat(messages, tools):
        r = client.chat.completions.create(model=model, messages=messages, tools=tools, temperature=0.2)
        m = r.choices[0].message
        calls = [{"id": t.id, "name": t.function.name, "arguments": t.function.arguments,
                  "raw": t.model_dump(exclude_none=True)} for t in (m.tool_calls or [])]
        return {"content": m.content, "tool_calls": calls}
    return _chat
