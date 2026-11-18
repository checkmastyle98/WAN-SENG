# 완생 — 인수인계 없는 신입을 위한 업무 맥락 복원 AI 에이전트

2026 대한민국 청년 CAREER FAIR AI 해커톤 예선 제출용 프로토타입입니다.
모든 회사·사람·파일은 가상 데이터입니다.

## 폴더 구성
| 파일 | 역할 |
|---|---|
| `app.py` | 웹 화면 (Streamlit) |
| `engine.py` | 파일 서버 색인, AI가 쓰는 도구 4개, 도구 호출 루프 |
| `make_data.py` | 가상 회사 파일 서버(`data/`)를 만드는 스크립트 |
| `data/` | 만들어진 가상 데이터 (반드시 함께 올릴 것) |

## 배포 (Streamlit Community Cloud)
1. 이 폴더 전체를 GitHub 공개 저장소에 올린다. `data/` 폴더도 포함한다.
2. share.streamlit.io → Create app → 저장소 선택 → Main file: `app.py`
3. Advanced settings → Secrets에 `.streamlit/secrets.toml.example` 내용을 참고해 키를 넣는다.
4. Deploy. 나온 주소를 제출한다.

## AI 제공사 지출 한도
제출 전에 API 콘솔에서 월 지출 한도를 꼭 걸어 둔다. 앱은 세션당 질문 25개로 제한되어 있다.

## 가상 데이터 다시 만들기
`python make_data.py`
