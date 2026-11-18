"""복기 — 인수인계 없는 신입을 위한 업무 맥락 복원 AI 에이전트 (해커톤 시연판)

실행: streamlit run app.py
API 설정: .streamlit/secrets.toml (README 참고)
"""
import os
import tempfile

import streamlit as st

import engine

st.set_page_config(page_title="복기 | 업무 맥락 복원 AI", page_icon="🧭", layout="wide")

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
    "이게 무슨 파일인지, 전임자는 이전 파일들을 어디에 저장해 놨는지, "
    "개정 이력과 히스토리는 어떻게 되는지, 내가 이틀 안에 뭘 해야 하는지 알려줘."
)
FOLLOW_UPS = [
    "10월 파일이 '최종'이랑 '최종_수정' 두 개인데 어느 게 진짜 최종본이야?",
    "VMI정산 시트는 왜 따로 분리돼 있어?",
    "입고율 90% 미만이면 빨간색 표시되는 이유가 뭐야?",
    "전임자가 남긴 인수인계서나 참고 자료는 없어?",
]
EXAMPLE_ANSWER = "구매팀 월말 정산 때 입고율 90% 미만 건은 공급사에 미입고 사유서를 받아야 해서, 2026년 4월부터 눈에 띄게 표시하기로 함"


# ------------------------------------------------------------------ 설정 / 세션
def get_secret(key, default=""):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return os.environ.get(key, default)


@st.cache_resource
def base_index():
    return engine.build_index()


def new_workspace():
    return engine.Workspace(docs=list(base_index()))


def init_state():
    ss = st.session_state
    ss.setdefault("ws", new_workspace())
    ss.setdefault("history", [])   # 모델에 넘기는 대화
    ss.setdefault("chat", [])      # 화면에 그리는 대화 (trace 포함)
    ss.setdefault("asked", 0)
    ss.setdefault("pending", None)


@st.cache_resource
def get_client(api_key, base_url):
    from openai import OpenAI
    return OpenAI(api_key=api_key, base_url=base_url or None)


def ask(question):
    ss = st.session_state
    if ss.asked >= MAX_QUESTIONS:
        st.warning(f"시연판은 세션당 질문 {MAX_QUESTIONS}개까지입니다. 왼쪽 '처음부터 다시'를 눌러 주세요.")
        return
    api_key = get_secret("API_KEY")
    if not api_key:
        st.error("API 키가 설정되지 않았습니다. (운영자: Streamlit Secrets에 API_KEY를 넣어 주세요)")
        return
    chat_fn = engine.openai_chat_fn(get_client(api_key, get_secret("BASE_URL")), get_secret("MODEL", "gpt-4o-mini"))
    ss.chat.append({"role": "user", "content": question})
    with st.spinner("복기가 파일 서버를 찾아보는 중입니다..."):
        try:
            answer, trace, ss.history = engine.run_agent(ss.ws, ss.history, question, chat_fn)
        except Exception as e:
            answer, trace = f"AI 호출 중 오류가 났습니다. 잠시 뒤 다시 시도해 주세요. ({type(e).__name__})", []
    ss.chat.append({"role": "assistant", "content": answer, "trace": trace})
    ss.asked += 1


def badge(text):
    return (text.replace("[근거 있음]", ":green-background[**근거 있음**]")
                .replace("[추정]", ":orange-background[**추정**]")
                .replace("[알 수 없음]", ":red-background[**알 수 없음**]"))


TOOL_LABEL = {"search_files": "🔍 검색", "read_file": "📄 파일 열기",
              "list_folder": "📁 폴더 보기", "request_confirmation": "🙋 담당자 확인 요청"}


def show_trace(trace):
    if not trace:
        return
    with st.expander(f"복기가 찾아본 과정 보기 ({len(trace)}단계)"):
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


# ------------------------------------------------------------------ 화면
init_state()
ss = st.session_state

with st.sidebar:
    st.subheader("심사위원 체험 가이드 (3분)")
    st.markdown(
        "1. **쪽지 받기** 탭에서 첨부파일을 복기에 넣습니다.\n"
        "2. **쪽지 내용으로 물어보기**를 누릅니다.\n"
        "3. 답변의 **근거 있음 / 추정 / 알 수 없음** 표시와 출처 경로를 확인합니다.\n"
        "4. 추천 질문 중 **입고율 90%** 질문을 눌러 봅니다.\n"
        "5. **확인 요청함** 탭에서 담당자 답변을 입력한 뒤, 같은 질문을 다시 해 봅니다."
    )
    st.divider()
    pend = sum(1 for c in ss.ws.confirmations if c["status"] == "답변 대기")
    st.metric("담당자 확인 대기", f"{pend}건")
    st.caption(f"이번 세션 질문 {ss.asked}/{MAX_QUESTIONS}")
    if st.button("처음부터 다시"):
        for k in ["ws", "history", "chat", "asked", "pending"]:
            ss.pop(k, None)
        st.rerun()
    st.divider()
    st.caption("모든 회사·사람·파일은 시연용 가상 데이터입니다. 실제 개인정보나 회사 자료는 쓰지 않았습니다.")
    st.caption("AI 연결은 OpenAI 호환 방식이라, 본선에서는 접속 주소만 바꿔 사내 로컬 AI로 전환할 수 있습니다.")

st.title("🧭 복기")
st.markdown("**인수인계 없는 신입을 위한 업무 맥락 복원 AI 에이전트** · 전임자의 파일·회의록·메일을 엮어 \"왜 이렇게 되어 있는지\"를 근거와 함께 알려줍니다.")

tab_chat, tab_confirm, tab_fs = st.tabs(["① 쪽지 받기 · 질문하기", "② 확인 요청함", "③ 회사 파일 서버 보기"])

with tab_chat:
    col_note, col_act = st.columns([3, 2])
    with col_note:
        with st.container(border=True):
            st.markdown(f"**📨 사내 쪽지** · {NOTE['from']} · {NOTE['date']}")
            st.write(NOTE["text"])
            with open(INBOX_FILE, "rb") as f:
                st.download_button(f"📎 {NOTE['file']} 내려받기", f.read(), file_name=NOTE["file"])
    with col_act:
        st.markdown("**첨부파일을 복기에 넣기**")
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
            ss.pending = FIRST_QUESTION

    for m in ss.chat:
        with st.chat_message(m["role"], avatar="🧑‍💼" if m["role"] == "user" else "🧭"):
            st.markdown(badge(m["content"]))
            if m["role"] == "assistant":
                show_trace(m.get("trace"))

    if ss.chat:
        st.caption("추천 질문")
        cols = st.columns(2)
        for i, q in enumerate(FOLLOW_UPS):
            if cols[i % 2].button(q, key=f"fu{i}"):
                ss.pending = q

    typed = st.chat_input("복기에게 물어보기 (예: 단가는 왜 직접 입력하면 안 돼?)")
    if typed:
        ss.pending = typed
    if ss.pending:
        q, ss.pending = ss.pending, None
        ask(q)
        st.rerun()

with tab_confirm:
    st.markdown("근거를 찾지 못한 질문은 담당자에게 확인 요청으로 쌓입니다. "
                "담당자가 답하면 그 답이 **확인답변** 기록으로 파일 서버에 저장되고, 다음부터는 근거로 쓰입니다.")
    if not ss.ws.confirmations:
        st.info("아직 확인 요청이 없습니다. '입고율 90%' 추천 질문을 해 보세요.")
    for item in ss.ws.confirmations:
        with st.container(border=True):
            st.markdown(f"**요청 #{item['id']}** · {item['status']}")
            st.write("질문: " + item["question"])
            st.caption("찾아본 내용: " + item["reason"])
            if item["status"] == "답변 대기":
                key = f"ans{item['id']}"
                if st.button("시연용 예시 답변 채우기", key=f"ex{item['id']}"):
                    st.session_state[key] = EXAMPLE_ANSWER
                ans = st.text_area("담당자 답변 (시연에서는 직접 담당자 역할을 해 보세요)", key=key)
                who = st.selectbox("답변자", ["조성파트 정민수 차장", "구매팀 박지현 과장"], key=f"who{item['id']}")
                if st.button("답변 저장", key=f"save{item['id']}", disabled=not ans.strip()):
                    d = ss.ws.add_confirmation_answer(item, ans.strip(), who)
                    st.success(f"기록 저장: {d.vpath} → 이제 ① 탭에서 같은 질문을 다시 해 보세요.")
                    st.rerun()
            else:
                st.success("답변: " + item["answer"])

with tab_fs:
    docs = ss.ws.docs
    st.markdown(f"가상 회사 파일 서버 `\\\\gaon-fs01` · 파일 **{len(docs)}개**. "
                "신입이 직접 찾으면 '최종', '최종_수정', '복사본', '새 폴더 (2)' 사이에서 길을 잃는 구조를 그대로 재현했습니다.")
    rows = [{"경로": d.vpath.replace(engine.SERVER + "\\", ""), "저장일": d.saved, "종류": d.kind} for d in docs]
    st.dataframe(rows, hide_index=True, height=520)
