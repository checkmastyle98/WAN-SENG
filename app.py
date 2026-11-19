"""완생 — 인수인계 없는 신입이 '이 일을 왜 하는지' 알고 일하도록 돕는 AI 에이전트 (해커톤 시연판)

실행: streamlit run app.py
API 설정: .streamlit/secrets.toml (README 참고)
"""
import html
import os
import tempfile
from datetime import date

import streamlit as st

import engine

st.set_page_config(page_title="완생 | 업무 맥락 복원 AI", page_icon="🧭", layout="wide")

MAX_QUESTIONS = 25  # 세션당 질문 수 제한 (API 비용 보호)
INBOX_FILE = os.path.join(engine.DATA, "inbox", "2026-11_오더PO정리.xlsx")
NOTE = {
    "from": "구매팀 박지현 과장",
    "date": "2026-11-18 (수) 09:12",
    "text": ("반갑습니다 김신입 씨, 이도현 씨 이후로 들어온 신입사원이시죠? "
             "전임자가 인수인계를 했을지는 모르겠는데, 이번 달 오더 PO정리를 해주셔야 합니다. "
             "기한은 이틀 뒤(11/20)까지입니다."),
    "file": "2026-11_오더PO정리.xlsx",
}
FIRST_QUESTION = (
    f"구매팀 박지현 과장에게 이런 쪽지를 받았어: \"{NOTE['text']}\" "
    f"첨부 파일은 업로드/{NOTE['file']} 이야. "
    "이 일이 앞뒤로 누구와 이어지고 왜 하는 건지, 틀리면 어디서 문제가 생기는지, "
    "양식이 왜 이렇게 되어 있는지, 이틀 안에 뭘 해야 하고 무엇을 선임에게 확인받아야 하는지 알려줘."
)
FOLLOW_UPS = [
    "10월 파일이 '최종'이랑 '최종_수정' 두 개인데 어느 게 진짜 최종본이야?",
    "VMI정산 시트는 왜 따로 분리돼 있어?",
    "입고율 90% 미만이면 빨간색 표시되는 이유가 뭐야?",
    "전임자가 남긴 인수인계서나 참고 자료는 없어?",
]
EXAMPLE_ANSWER = "구매팀 월말 정산 때 입고율 90% 미만 건은 공급사에 미입고 사유서를 받아야 해서, 2026년 4월부터 눈에 띄게 표시하기로 함"
EXAMPLE_MINUTES = """[AI 회의록 자동 정리] 구매·생산계획 협의
일시: 2026-11-19 10:00 / 참석: 조성파트 정민수 차장, 구매팀 박지현 과장, 생산계획팀 최윤아 대리

■ 논의 요약
- 생산계획팀이 미입고분을 다음 달로 이월할 때, 납기 지연 사유를 몰라 구매팀에 매번 따로 물어보느라 반나절이 걸림.
- 입고율 90% 미만 건을 구매팀만 알고 생산계획팀은 늦게 알게 되어 원료 발주량 계획이 늦어짐.

■ 결정 사항
- [결정] 11월분부터 오더PO정리 PO목록에 '납기 지연 사유' 칸을 추가해, 입고율 90% 미만 건의 사유를 적는다.
- [결정] 입고율 90% 미만 건은 마감 메일 본문에 따로 표시해 구매팀과 생산계획팀에 함께 알린다.

■ 기타
- 다음 회의: 12월 첫째 주"""
APPROVERS = ["조성파트 정민수 차장", "구매팀 박지현 과장"]

CSS = """
<style>
.ws-flow{display:flex;align-items:stretch;gap:10px;margin:6px 0 4px}
.ws-col{flex:1;display:flex;flex-direction:column;gap:8px}
.ws-cap{font-size:13px;font-weight:700;color:#6b7280;letter-spacing:.02em}
.ws-box{border:1px solid #d6dbe4;border-radius:12px;padding:12px 14px;background:#fbfbfd}
.ws-me{border:2px solid #e0602e;background:#fff4ee}
.ws-who{font-weight:700;font-size:15px;color:#1f2a44}
.ws-what{font-size:14px;color:#374151;margin-top:4px;line-height:1.45}
.ws-src{font-size:12px;color:#6b7280;margin-top:6px;word-break:break-all}
.ws-arrow{display:flex;align-items:center;font-size:26px;color:#9aa3b2;padding-top:18px}
.ws-purpose{margin:10px 0 2px;padding:10px 14px;border-radius:10px;background:#1f2a44;color:#f6f5f1;font-size:15px}
.lv{display:inline-block;font-size:12px;font-weight:700;border-radius:6px;padding:1px 8px;margin-left:6px}
.lv-ok{background:#d7efdf;color:#1e5b33}.lv-guess{background:#fbe3cc;color:#7a3d06}.lv-unk{background:#f6d5d5;color:#7d1f1f}
.rk{display:inline-block;font-size:12px;font-weight:700;border-radius:6px;padding:1px 8px}
.rk-h{background:#e0602e;color:#fff}.rk-m{background:#fbe3cc;color:#7a3d06}.rk-l{background:#e5e7eb;color:#374151}
@media (max-width:800px){.ws-flow{flex-direction:column}.ws-arrow{justify-content:center;padding:0;transform:rotate(90deg)}}
</style>
"""


# ------------------------------------------------------------------ 설정 / 세션
def get_secret(key, default=""):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return os.environ.get(key, default)


@st.cache_resource
def base_index():
    return engine.build_index()


# 안전장치: 서버가 예전 engine을 기억하고 있거나 data가 예전 버전이어도, 앱이 직접 가상 데이터를 다시 만든다
FLOW_MAIL = os.path.join(engine.FS_ROOT, "생산부", "조성파트", "메일보관", "2026-10-06_생산계획팀_PO정리_활용안내.txt")
if not os.path.exists(INBOX_FILE) or not os.path.exists(FLOW_MAIL):
    import json
    import make_data
    make_data.main()
    with open(os.path.join(engine.DATA, "file_meta.json"), encoding="utf-8") as _f:
        engine.FILE_META.update(json.load(_f))
    base_index.clear()


def new_workspace():
    return engine.Workspace(docs=list(base_index()))


def init_state():
    ss = st.session_state
    ss.setdefault("ws", new_workspace())
    ss.setdefault("history", [])   # 모델에 넘기는 대화
    ss.setdefault("chat", [])      # 화면에 그리는 대화 (trace, report 포함)
    ss.setdefault("asked", 0)
    ss.setdefault("pending", None)
    ss.setdefault("extracted", None)   # 회의록에서 뽑은 결정 사항 (선임 승인 전)


@st.cache_resource
def get_client(api_key, base_url):
    from openai import OpenAI
    return OpenAI(api_key=api_key, base_url=base_url or None)


def client_or_none():
    key = get_secret("API_KEY")
    return get_client(key, get_secret("BASE_URL")) if key else None


def ask(question, onboarding=False):
    ss = st.session_state
    if ss.asked >= MAX_QUESTIONS:
        st.warning(f"시연판은 세션당 질문 {MAX_QUESTIONS}개까지입니다. 왼쪽 '처음부터 다시'를 눌러 주세요.")
        return
    client = client_or_none()
    if client is None:
        st.error("API 키가 설정되지 않았습니다. (운영자: Streamlit Secrets에 API_KEY를 넣어 주세요)")
        return
    chat_fn = engine.openai_chat_fn(client, get_secret("MODEL", "gpt-4o-mini"))
    ss.chat.append({"role": "user", "content": question})
    report = None
    with st.spinner("완생이 파일 서버에서 이 일의 앞뒤 사정을 찾아보는 중입니다..."):
        try:
            answer, trace, ss.history, report = engine.run_agent(ss.ws, ss.history, question, chat_fn,
                                                                 onboarding=onboarding)
        except Exception as e:
            answer, trace = f"AI 호출 중 오류가 났습니다. 잠시 뒤 다시 시도해 주세요. ({type(e).__name__})", []
    ss.chat.append({"role": "assistant", "content": answer, "trace": trace, "report": report,
                    "updates_seen": len(ss.ws.updates)})
    ss.asked += 1


def badge(text):
    return (text.replace("[근거 있음]", ":green-background[**근거 있음**]")
                .replace("[추정]", ":orange-background[**추정**]")
                .replace("[알 수 없음]", ":red-background[**알 수 없음**]"))


def lv(level):
    cls = {"근거 있음": "lv-ok", "추정": "lv-guess"}.get(level, "lv-unk")
    return f'<span class="lv {cls}">{html.escape(level or "알 수 없음")}</span>'


def rk(risk):
    cls = {"높음": "rk-h", "보통": "rk-m"}.get(risk, "rk-l")
    return f'<span class="rk {cls}">위험도 {html.escape(risk or "-")}</span>'


def e(x):
    return html.escape(str(x or ""))


def short(path):
    """경로가 길면 화면에는 파일 이름만 보여준다."""
    return e(str(path or "").replace("/", "\\").split("\\")[-1])


TOOL_LABEL = {"search_files": "🔍 검색", "read_file": "📄 파일 열기",
              "list_folder": "📁 폴더 보기", "request_confirmation": "🙋 담당자 확인 요청"}


def show_trace(trace):
    if not trace:
        return
    with st.expander(f"완생이 찾아본 과정 보기 ({len(trace)}단계)"):
        for i, t in enumerate(trace, 1):
            arg = next(iter(t["args"].values()), "") if t["args"] else ""
            st.markdown(f"**{i}. {TOOL_LABEL.get(t['tool'], t['tool'])}** `{arg}`")
            r = t["result"]
            if t["tool"] == "search_files" and "상위결과" in r:
                st.caption(f"전체 {r['결과수']}건 중 상위 {len(r['상위결과'])}건")
                for h in r["상위결과"]:
                    st.caption(f"· {h['경로']}  ({h['저장일']})")
            elif t["tool"] == "read_file" and "경로" in r:
                st.caption(f"· {r['경로']}  ({r['저장일']})")
            else:
                st.caption(str(r)[:300])


# ------------------------------------------------------------------ 업무 안내서 화면
def flow_html(flow):
    def box(s, me=False):
        src = f'<div class="ws-src">📄 {short(s.get("source"))}</div>' if s.get("source") else ""
        level = "" if me else lv(s.get("level"))
        who = "나 (김신입)" if me else s.get("who")
        due = f' · 기한 {e(s.get("due"))}' if me and s.get("due") else ""
        return (f'<div class="ws-box{" ws-me" if me else ""}"><div class="ws-who">{e(who)}{level}</div>'
                f'<div class="ws-what">{e(s.get("what"))}{due}</div>{src}</div>')

    before = "".join(box(s) for s in flow.get("before", [])) or '<div class="ws-box"><div class="ws-what">기록 없음</div></div>'
    after = "".join(box(s) for s in flow.get("after", [])) or '<div class="ws-box"><div class="ws-what">기록 없음</div></div>'
    return (CSS + '<div class="ws-flow">'
            f'<div class="ws-col"><div class="ws-cap">전 · 누가 넘겨주나</div>{before}</div>'
            '<div class="ws-arrow">➜</div>'
            f'<div class="ws-col"><div class="ws-cap">지금 · 내가 하는 일</div>{box(flow.get("me", {}), me=True)}</div>'
            '<div class="ws-arrow">➜</div>'
            f'<div class="ws-col"><div class="ws-cap">후 · 누가 받아 무엇에 쓰나</div>{after}</div>'
            '</div>'
            + (f'<div class="ws-purpose">🎯 <b>이 일을 왜 하나</b> — {e(flow.get("purpose"))}</div>' if flow.get("purpose") else ""))


def render_report(r, idx):
    ss = st.session_state
    st.markdown(f"**받은 파일** · {r.get('file_summary', '')}")

    st.markdown("#### 🔗 이 일의 흐름")
    st.markdown(flow_html(r.get("flow", {})), unsafe_allow_html=True)

    st.markdown("#### ⚠️ 틀리면 어떻게 되나")
    risks = sorted(r.get("risks", []), key=lambda x: {"높음": 0, "보통": 1}.get(x.get("risk"), 2))
    for x in risks:
        src = f'<br><span class="ws-src">📄 {short(x.get("source"))}</span>' if x.get("source") else ""
        st.markdown(f'{rk(x.get("risk"))} &nbsp;<b>{e(x.get("item"))}</b> — {e(x.get("impact"))}{lv(x.get("level"))}{src}',
                    unsafe_allow_html=True)
    if any(x.get("risk") == "높음" for x in risks):
        st.caption("위험도 '높음' 항목은 근거가 있어도 아래 🙋 선임 검토 대상으로 넘어갑니다.")

    st.markdown("#### 📜 왜 이렇게 되어 있나")
    for x in r.get("reasons", []):
        src = f'<br><span class="ws-src">📄 {short(x.get("source"))}</span>' if x.get("source") else ""
        st.markdown(f'<b>{e(x.get("rev"))}</b> {e(x.get("change"))}{lv(x.get("level"))}<br>'
                    f'&nbsp;&nbsp;→ 이유: {e(x.get("why"))}{src}', unsafe_allow_html=True)

    st.markdown("#### ✅ 무엇을 하면 되나")
    for j, x in enumerate(r.get("todos", [])):
        due = f" (기한 {x['due']})" if x.get("due") else ""
        st.checkbox(f"{x.get('task')}{due}", key=f"todo{idx}_{j}")

    st.markdown("#### 🙋 무엇을 확인받아야 하나")
    reviews = r.get("reviews", [])
    for x in reviews:
        st.markdown(f"- **{x.get('item')}** — {x.get('question')}  \n  <span class='ws-src'>왜: {e(x.get('why'))}</span>",
                    unsafe_allow_html=True)
    if reviews:
        sent = all(any(c["question"] == x.get("question") for c in ss.ws.confirmations) for x in reviews)
        if st.button("선임에게 검토 요청 보내기" if not sent else "검토 요청 보냄 ✓", key=f"rev{idx}",
                     type="primary", disabled=sent):
            for x in reviews:
                ss.ws.add_review_request(x.get("item", ""), x.get("question", ""), x.get("why", ""))
            st.rerun()


# ------------------------------------------------------------------ 화면
init_state()
ss = st.session_state

with st.sidebar:
    st.subheader("심사위원 체험 가이드 (3분)")
    st.markdown(
        "1. **① 탭**에서 첨부파일을 넣고 **쪽지 내용으로 물어보기**를 누릅니다.\n"
        "2. 맨 위 **업무 흐름 그림**(전 → 나 → 후)으로 이 일을 왜 하는지 봅니다.\n"
        "3. **틀리면 어떻게 되나 → 왜 이렇게 되어 있나 → 할 일 → 확인받을 것** 순서로 읽고, 검토 요청을 보냅니다.\n"
        "4. **③ 탭**에서 선임이 되어 새 회의록을 승인하면, 개정 이력에 반영되고 신입 화면에 업데이트로 뜹니다.\n"
        "5. 더 궁금한 건 맨 아래 입력창에 물어봅니다."
    )
    st.divider()
    pend = sum(1 for c in ss.ws.confirmations if c["status"] == "답변 대기")
    st.metric("선임 확인 대기", f"{pend}건")
    st.metric("업무 업데이트", f"{len(ss.ws.updates)}건")
    st.caption(f"이번 세션 질문 {ss.asked}/{MAX_QUESTIONS}")
    if st.button("처음부터 다시"):
        for k in ["ws", "history", "chat", "asked", "pending", "extracted"]:
            ss.pop(k, None)
        st.rerun()
    st.divider()
    st.caption("모든 회사·사람·파일은 시연용 가상 데이터입니다. 실제 개인정보나 회사 자료는 쓰지 않았습니다.")
    st.caption("AI 연결은 OpenAI 호환 방식이라, 본선에서는 접속 주소만 바꿔 사내 로컬 AI로 전환할 수 있습니다.")

st.title("🧭 완생")
st.markdown("**신입이 '이 일을 왜 하는지' 알고 일하게 돕는 AI** · 이 일의 앞뒤 흐름, 틀리면 생기는 문제, 양식의 이유를 "
            "근거와 함께 보여주고, 중요한 부분은 선임에게 검토받게 합니다.")

tab_chat, tab_confirm, tab_meet, tab_fs = st.tabs(
    ["① 쪽지 받기 · 업무 안내", "② 선임 확인 요청함", "③ 회의록 업데이트 (선임용)", "④ 회사 파일 서버 보기"])

with tab_chat:
    # 회의록으로 바뀐 내용이 있으면 신입에게 먼저 알린다
    for u in reversed(ss.ws.updates):
        with st.container(border=True):
            st.markdown(f"🆕 **업무 업데이트** · {u['date']} {u['title']} · {u['approver']} 승인")
            for rv in u["revisions"]:
                st.markdown(f"- **{rv['개정번호']}** {rv['개정 내용']}")
            st.caption(f"개정이력 시트에 반영됨 · 회의록: {u['minutes_path']}")
    last_report = next((m for m in reversed(ss.chat) if m.get("report")), None)
    if last_report and last_report.get("updates_seen", 0) < len(ss.ws.updates):
        if st.button("🔄 바뀐 내용을 반영해 업무 안내 다시 받기"):
            ss.pending = ("onboarding", FIRST_QUESTION + " 최근 회의로 바뀐 개정 이력도 반영해줘.")

    col_note, col_act = st.columns([3, 2])
    with col_note:
        with st.container(border=True):
            st.markdown(f"**📨 사내 쪽지** · {NOTE['from']} · {NOTE['date']}")
            st.write(NOTE["text"])
            path = ss.ws.upload_paths.get(NOTE["file"], INBOX_FILE)
            with open(path, "rb") as f:
                st.download_button(f"📎 {NOTE['file']} 내려받기", f.read(), file_name=NOTE["file"])
    with col_act:
        st.markdown("**첨부파일을 완생에 넣기**")
        if st.button("📎 쪽지 첨부파일 바로 넣기", type="primary"):
            ss.ws.add_upload(NOTE["file"], INBOX_FILE)
        up = st.file_uploader("또는 엑셀 파일 직접 올리기", type=["xlsx"])
        if up is not None and up.name not in ss.ws.uploads:
            tmp = os.path.join(tempfile.gettempdir(), up.name)
            with open(tmp, "wb") as f:
                f.write(up.getbuffer())
            ss.ws.add_upload(up.name, tmp)
        if ss.ws.uploads:
            st.success("넣은 파일: " + ", ".join(ss.ws.uploads))

    if ss.ws.uploads and not ss.chat:
        if st.button("💬 쪽지 내용으로 물어보기"):
            ss.pending = ("onboarding", FIRST_QUESTION)

    for i, m in enumerate(ss.chat):
        with st.chat_message(m["role"], avatar="🧑‍💼" if m["role"] == "user" else "🧭"):
            if m["role"] == "assistant" and m.get("report"):
                render_report(m["report"], i)
            else:
                st.markdown(badge(m["content"]))
            if m["role"] == "assistant":
                show_trace(m.get("trace"))

    if ss.chat:
        st.caption("더 궁금한 것 · 추천 질문")
        cols = st.columns(2)
        for i, q in enumerate(FOLLOW_UPS):
            if cols[i % 2].button(q, key=f"fu{i}"):
                ss.pending = ("chat", q)

    typed = st.chat_input("더 궁금한 건 완생에게 물어보세요 (예: 단가는 왜 직접 입력하면 안 돼?)")
    if typed:
        ss.pending = ("chat", typed)
    if ss.pending:
        mode, q = ss.pending
        ss.pending = None
        ask(q, onboarding=(mode == "onboarding"))
        st.rerun()

with tab_confirm:
    st.markdown("두 종류의 요청이 쌓입니다. **검토**는 틀리면 영향이 큰 항목이라 근거가 있어도 확인받는 것이고, "
                "**확인**은 기록이 없어 AI가 답할 수 없는 '왜'입니다. 선임이 답하면 **확인답변** 기록으로 저장되어 다음부터 근거로 쓰입니다.")
    if not ss.ws.confirmations:
        st.info("아직 요청이 없습니다. 업무 안내의 '선임에게 검토 요청 보내기'나 '입고율 90%' 추천 질문을 해 보세요.")
    for item in ss.ws.confirmations:
        with st.container(border=True):
            st.markdown(f"**요청 #{item['id']}** · `{item.get('kind', '확인')}` · {item['status']}")
            st.write("질문: " + item["question"])
            st.caption("사유: " + item["reason"])
            if item["status"] == "답변 대기":
                key = f"ans{item['id']}"
                if st.button("시연용 예시 답변 채우기", key=f"ex{item['id']}"):
                    st.session_state[key] = EXAMPLE_ANSWER if "입고율" in item["question"] else "확인했습니다. 그대로 진행하면 됩니다."
                ans = st.text_area("선임 답변 (시연에서는 직접 선임 역할을 해 보세요)", key=key)
                who = st.selectbox("답변자", APPROVERS, key=f"who{item['id']}")
                if st.button("답변 저장", key=f"save{item['id']}", disabled=not ans.strip()):
                    d = ss.ws.add_confirmation_answer(item, ans.strip(), who)
                    st.success(f"기록 저장: {d.vpath} → 이제 ① 탭에서 같은 질문을 다시 해 보세요.")
                    st.rerun()
            else:
                st.success("답변: " + item["answer"])

with tab_meet:
    st.markdown("선임이 **AI로 정리한 회의록**을 올리면, 완생이 업무 양식·기준을 바꾸는 **[결정]**을 뽑아 개정 이력 초안을 만듭니다. "
                "**선임이 승인한 것만** 개정 이력에 반영되고, 신입 화면에 '업무 업데이트'로 표시됩니다.")
    c1, c2 = st.columns([1, 1])
    m_title = c1.text_input("회의명", "구매·생산계획 협의")
    m_date = c2.date_input("회의일", date(2026, 11, 19))
    if st.button("시연용 예시 회의록 채우기"):
        st.session_state["minutes"] = EXAMPLE_MINUTES
    minutes = st.text_area("회의록 (녹음 → AI 정리본을 붙여넣기)", key="minutes", height=260)
    if st.button("🔎 결정 사항 뽑기", disabled=not minutes.strip()):
        client = client_or_none()
        if client:
            with st.spinner("회의록에서 업무를 바꾸는 결정을 찾는 중입니다..."):
                dec, how = engine.extract_decisions_ai(minutes, client, get_secret("MODEL", "gpt-4o-mini"))
        else:
            dec, how = engine.extract_decisions_rule(minutes), "규칙"
        ss.extracted = {"how": how, "rows": [{"반영": True, "바뀌는 내용": d["change"], "이유": d.get("reason", "")} for d in dec]}
    if ss.extracted is not None:
        if not ss.extracted["rows"]:
            st.warning("업무 양식·기준을 바꾸는 [결정]을 찾지 못했습니다.")
        else:
            st.caption(f"추출 방식: {ss.extracted['how']} · 내용을 고치거나 '반영'을 끄고 승인하세요. AI가 뽑은 결정은 승인 전까지 업무 기준이 되지 않습니다.")
            import pandas as pd
            edited = st.data_editor(pd.DataFrame(ss.extracted["rows"]), key="dec_editor", hide_index=True)
            approver = st.selectbox("승인자", APPROVERS, key="approver")
            chosen = [r for r in edited.to_dict("records") if r.get("반영") and str(r.get("바뀌는 내용", "")).strip()]
            if st.button(f"✅ 승인하고 개정 이력에 반영 ({len(chosen)}건)", type="primary", disabled=not chosen):
                ss.ws.ensure_copy(NOTE["file"], INBOX_FILE)
                added, mdoc = ss.ws.apply_meeting(
                    NOTE["file"], m_title, m_date.isoformat(), minutes,
                    [{"change": r["바뀌는 내용"], "reason": r.get("이유", "")} for r in chosen], approver)
                ss.extracted = None
                st.success(f"{', '.join(a['개정번호'] for a in added)} 반영 완료 · 회의록 저장: {mdoc.vpath}")
                st.rerun()
    st.divider()
    st.markdown(f"**현재 `{NOTE['file']}` 개정 이력**")
    ss.ws.ensure_copy(NOTE["file"], INBOX_FILE)
    st.dataframe(ss.ws.revision_rows(NOTE["file"]), hide_index=True)

with tab_fs:
    docs = ss.ws.docs
    st.markdown(f"가상 회사 파일 서버 `\\\\gaon-fs01` · 파일 **{len(docs)}개**. "
                "신입이 직접 찾으면 '최종', '최종_수정', '복사본', '새 폴더 (2)' 사이에서 길을 잃는 구조를 그대로 재현했습니다.")
    rows = [{"경로": d.vpath.replace(engine.SERVER + "\\", ""), "저장일": d.saved, "종류": d.kind} for d in docs]
    st.dataframe(rows, hide_index=True, height=520)
