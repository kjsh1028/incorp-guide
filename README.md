# 법인 설립 가이드 생성기

누가, 어디에, 어떤 법인을 세우는지와 기본 정보를 넣으면 절차, 서류, 신청 기관, 비용 안내를 만들어 주는 내부 도구.

- `index.html`: 화면. 사건 관리(홈, 사건, 새 사건)와 가이드. 사건 자료는 더보기에서 연결한 비공개 저장소의 `cases.json` 에 저장한다(이 저장소에는 넣지 않는다).
- `tools/docgen.py`: 고객 문서 생성(자료요청서, 서명 및 공증 서류 안내, 진행 보고)과 자료요청서 회신 판독.
- `tools/templates/`: 표준 템플릿(국문·중문). 고객별 값은 `{{...}}` 빈칸.
- `tools/firm_profile.json`: 발신 변호사와 원본 수령인 정보.
- `data/tracks.json`: 첫 화면의 트랙 목록.
- `data/tracks/`: 트랙별 규칙 데이터(JSON). 화면은 이 파일들을 읽어 그린다.
- `data/bundle.js`: 위 JSON을 묶은 파일(자동 생성). 화면을 파일로 열 때 쓴다.
- `tools/build_data.py`: JSON을 고친 뒤 실행해 `data/bundle.js` 를 다시 만든다.

## 쓰는 법

    python tools/build_data.py          규칙 JSON을 고친 뒤 실행

    python tools/docgen.py request  tools/sample_case.json out/
    python tools/docgen.py signing  tools/sample_case.json out/
    python tools/docgen.py progress tools/sample_case.json out/
    python tools/docgen.py read     tools/sample_case.json 회신본.docx > 사건.case.json

회신 판독에는 python-docx 가 필요하다. 읽은 결과의 `program_inputs`를 화면의 "고객 회신에서 읽은 값 불러오기"에 붙여넣는다.

## 절대 올리지 않는 것

고객명, 여권·신분증 정보, 실제 사건 파일, 채워진 회신본. `.gitignore`가 `cases/`, `out/`을 막고 있다.
