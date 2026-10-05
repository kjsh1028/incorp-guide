# 규칙 데이터

트랙별 JSON 파일을 `tracks/<트랙ID>/` 아래에 둔다.

KR-JSC 의 6개 파일(inputs, rules, process, documents, agencies, costs)은 드라이브
`법인 설립 프로그램 > 01_규칙데이터 > tracks > KR-JSC` 에 있다. 저장소를 만든 뒤 그 폴더의
파일을 `data/tracks/KR-JSC/` 에 넣는다.

트랙 목록(첫 화면 선택지)은 `tracks.json` 에 있다. 새 트랙을 열려면 폴더에 6개 파일을 넣고 `ready` 를 true로 바꾼다.

JSON을 고친 뒤에는 `python tools/build_data.py` 를 실행한다. 화면을 파일로 열 때 쓰는 `bundle.js` 가 다시 만들어진다.

조건식(when, applies, expr, formula, warn_if)에 쓸 수 있는 것: 입력·파생 항목 키, 숫자, '문자열', true/false,
[목록], in, !, 사칙연산, 비교(== != < > <= >=), && ||, 조건 ? 값 : 값, 괄호. 비용 항목 ID(TAX-1 등)는 앞의 계산값을 가리킨다.
