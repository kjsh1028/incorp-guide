"""data/ 의 JSON을 묶어 data/bundle.js 를 만든다.

index.html 은 웹 주소(http/https)로 열면 JSON을 직접 읽고, 파일을 더블클릭해 열면
(file://) 브라우저가 JSON 읽기를 막으므로 이 묶음 파일을 쓴다.
JSON을 고친 뒤에는 반드시 다시 실행한다.

    python tools/build_data.py          묶음 파일 다시 만들기
    python tools/build_data.py --check  묶음 파일이 JSON과 같은지 확인(다르면 종료 코드 1)
"""
import json, os, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")
OUT = os.path.join(DATA, "bundle.js")
FILES = ["inputs", "rules", "process", "documents", "agencies", "costs", "outputs"]


def load(path):
    with open(path, encoding="utf-8") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError as e:
            sys.exit(f"JSON 오류: {os.path.relpath(path, ROOT)} {e.lineno}행 {e.colno}열: {e.msg}")


def build():
    tracks = load(os.path.join(DATA, "tracks.json"))
    out = {"tracks": tracks, "data": {}}
    for t in tracks["tracks"]:
        if t.get("ready"):
            d = os.path.join(DATA, "tracks", t["id"])
            out["data"][t["id"]] = {f: load(os.path.join(d, f + ".json")) for f in FILES}
    return ("/* 자동 생성 파일. 직접 고치지 말고 data/ 의 JSON을 고친 뒤 python tools/build_data.py 를 실행한다. */\n"
            "window.TRACK_DATA=" + json.dumps(out, ensure_ascii=False, separators=(",", ":")) + ";\n")


if __name__ == "__main__":
    text = build()
    old = open(OUT, encoding="utf-8").read() if os.path.exists(OUT) else ""
    if "--check" in sys.argv:
        if text != old:
            sys.exit("data/bundle.js 가 JSON과 다릅니다. python tools/build_data.py 를 실행하세요.")
        print("data/bundle.js 최신")
    else:
        with open(OUT, "w", encoding="utf-8", newline="\n") as f:
            f.write(text)
        print("data/bundle.js", "갱신" if text != old else "변경 없음")
