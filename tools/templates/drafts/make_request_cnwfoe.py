# -*- coding: utf-8 -*-
"""한국 법인의 중국 독자법인(CN-WFOE) 설립 자료요청서 초안 틀을 만든다.
확정된 request_ko.docx(법무지원팀 확인본과 같은 서식)를 복사해 내용만 바꾼다(서식을 새로 만들지 않음).
  - 노란 {{빈칸}}: 사건별로 채우는 값(docgen.py, 화면이 채움)
  - 녹색: Claude 초안 중 법무지원팀 확인이 필요한 부분(문서를 만들면 표시는 지워짐)
법무지원팀 확인본을 받으면 이 스크립트 대신 확인본에서 틀을 다시 만든다.
  python tools/templates/drafts/make_request_cnwfoe.py
"""
import copy, os, docx
from docx.enum.text import WD_COLOR_INDEX
HERE = os.path.dirname(os.path.abspath(__file__)); TPL = os.path.dirname(HERE)
d = docx.Document(os.path.join(TPL, "request_ko.docx"))
G = "green"

def hl(run, color):
    run.font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN if color == G else (WD_COLOR_INDEX.YELLOW if color == "y" else None)
def set_runs(p, parts):
    """p 의 마지막 run 서식을 본으로 parts=[(글, 색)] 를 넣는다. 번호 run(앞쪽)은 그대로 둔다."""
    keep = p.runs[:-1] if len(p.runs) >= 3 else []          # "1." + 탭 + 본문 꼴이면 앞 두 run 유지
    proto = p.runs[-1]
    for r in p.runs[len(keep):]:
        if r._element is not proto._element: r._element.getparent().remove(r._element)
    first = True
    for text, color in parts:
        r = proto if first else p.add_run()
        if not first: r._element.insert(0, copy.deepcopy(proto._element.rPr)) if proto._element.rPr is not None else None
        r.text = text; hl(r, None); 
        if color: hl(r, color)
        first = False
def P(i, *parts): set_runs(pp[i], [(x, None) if isinstance(x, str) else x for x in parts])   # pp: 처음 읽은 단락 목록(지운 뒤에도 번호가 밀리지 않게)
def drop(el): el.getparent().remove(el)
def set_cell(cell, lines):
    """lines: 줄마다 [(글, 색)...]. 칸의 단락 수를 줄 수에 맞춘다."""
    ps = cell.paragraphs
    while len(ps) < len(lines):
        ps[-1]._element.addnext(copy.deepcopy(ps[-1]._element)); ps = cell.paragraphs
    for p in ps[len(lines):]: drop(p._element)
    for p, parts in zip(cell.paragraphs, lines):
        parts = [(x, None) if isinstance(x, str) else x for x in parts]
        if not p.runs: p.add_run("")
        proto = p.runs[0]
        for r in p.runs[1:]: drop(r._element)
        proto.text = parts[0][0]; hl(proto, parts[0][1])
        for text, color in parts[1:]:
            r = p.add_run(text)
            if proto._element.rPr is not None: r._element.insert(0, copy.deepcopy(proto._element.rPr))
            hl(r, color)
def row(t, i, *cells):
    for c, v in zip(t.rows[i].cells, cells):
        if v is None: continue
        set_cell(c, v if isinstance(v, list) else [[v]])

pp = list(d.paragraphs); tb = list(d.tables)
# 표지·머리
P(0, "[초안 v0.1 (한국 법인의 중국 독자법인 설립) — 노란 부분을 사건별로 기입, 녹색 부분은 법무지원팀 확인 전 초안, 발송 전 이 줄 삭제]")
for r in pp[8].runs: r.text = r.text.replace("한국 자회사", "중국 자회사")
for r in pp[10].runs: r.text = r.text.replace("한국 자회사", "중국 자회사")
# 1. 회사 정보
t = tb[0]
row(t, 2, "★ 2. 상호", [["중문: 후보1 ________　후보2 ________　후보3 ________"], ["영문: ________________"], [("※ 중국 상호는 등기기관의 사전 확인을 거치므로 후보를 여러 개 받습니다. 지역명·업종 표기는 율촌이 맞춰 회신합니다.", G)]])
row(t, 3, "★ 3. 등록자본", [["________ (통화: □ 미화　□ 위안　□ 원)"], [("※ 납입 기한과 방식은 정관에 정하며, 법정 기한은 율촌이 확인해 안내합니다.", G)]])
row(t, 4, "4. 출자 방식", [["□ 현금　□ 기타 ________"]])
row(t, 5, "★ 5. 경영범위", [[""], [""], [("※ 하려는 사업을 모두 적어 주시면, 율촌이 외상투자 제한 여부를 검토하고 중국 표준 표기로 정리해 회신합니다.", G)]])
row(t, 6, "6. 등록 주소", [["지역(도시·구): ________"], ["사무실: □ 임차 예정　□ 확보함　□ 미정"], [("※ 등록 주소에는 실제 사무실의 임대차계약이 필요하며, 공유 주소 사용 가능 여부는 지역마다 다릅니다.", G)]])
row(t, 7, "7. 조직 구성", [["이사: □ 이사 1인　□ 이사회"], ["감사: □ 감사 1인　□ 두지 않음"], [("※ 구성 가능 여부는 율촌이 확인해 안내합니다.", G)]])
row(t, 8, "8. 해외직접투자 신고(한국)", [["□ 완료　□ 진행 중　□ 미착수　　완료 예정: ________"]])
row(t, 9, "9. 은행", [["송금 은행(한국): ________"], ["중국 개설 희망 은행: □ 있음: ________　□ 없음 (율촌이 주선)"]])
# 임원 정보
P(14, "임원 정보 (여권 기재대로)")
t = tb[1]
row(t, 0, None, "법정대표인", "이사(해당 시)", "감사(해당 시)")
row(t, 2, "한글 성명")
row(t, 6, "중국 상주 여부")
P(16, ("법정대표인은 이사(또는 경리) 중에서 정합니다. 재무책임자와 등기 연락원(중국 측 직원도 가능)의 성명·연락처도 따로 알려 주십시오.", G))
P(17, "★ 임원마다 여권 사본이 필요합니다. ", ("중국 등기 시스템의 본인 확인을 위해 휴대전화 번호와 실명 인증을 요청드릴 수 있습니다.", G))
# 2. 준비하실 서류 A
t = tb[2]
row(t, 0, None, None, None, None, None, None, "중문 번역", "비고")
A = [("투자 결정서(주주 결정)", "귀사 대표이사 서명+법인인감", "2부", "○", "○", "율촌", ""),
     ("정관(公司章程)", "귀사 대표이사 서명+법인인감", "2부", "—", "—", "율촌", ""),
     ("법정대표인·이사·감사 선임 서류", "귀사 대표이사 서명+법인인감", "1부", "○", "○", "율촌", ""),
     ("설립등기 위임장", "귀사 대표이사 서명+법인인감", "1부", "○", "○", "율촌", ""),
     ("임원 서명 확인서 (여권 사본 첨부, 해당 시)", "각 임원", "각 1부", "○", "○", "율촌", ""),
     ("은행 계좌 개설 서류 (해당 시)", "법정대표인", "1부", "—", "—", "—", "은행 양식")]
for i, a in enumerate(A, 1): row(t, i, None, *[[[(x, G)]] for x in a])
# B
P(21, "B. 귀사에 대한 서류 — 공증·아포스티유를 받아 원본을 보내 주실 서류")
t = tb[3]
B = ["법인 등기사항전부증명서 (공증+아포스티유)", "사업자등록증 사본", "대표이사 여권 사본 (공증+아포스티유)", "은행 잔고증명서 또는 자신(资信)증명"]
for i, b in enumerate(B, 1): row(t, i, None, [[(b, G)]])
# 진행 요령
P(24, "지금은 B 서류 준비를 시작해 주시고, A 서류는 율촌의 서명 서류를 받은 뒤 한꺼번에 공증해 주십시오.")
P(25, ("공증은 한국 공증사무소에서, 아포스티유는 외교부에서 받습니다.", G))
P(26, "중국 기관에 내는 중문 번역과 관련 절차는 율촌이 처리하므로 준비하지 않으셔도 됩니다.")
P(27, "서류마다 공증서를 따로 받고, 원본 부수는 위 표대로 받아 주십시오.")
P(28, "공증 신청 시 용도를 \"중국 회사 설립 등기용\"으로 적어 주십시오.")
P(29, "서류는 발급 후 ", ("일정 기간 안에 중국 기관에 제출해야 하므로(기간은 율촌이 확인해 안내)", G), ", 받으시는 대로 바로 발송해 주십시오.")
# 3. 원본 발송
P(31, "전체 스캔본을 먼저 율촌 이메일로 보내 주시고, 율촌이 확인한 뒤 원본을 아래 주소로 보내 주십시오.")
P(36, "등기우편이나 퀵서비스 등 추적 가능한 방법을 이용해 주십시오.")
# 4. 진행 순서 및 기간
t = tb[4]
S = [("정보 회신, B 서류 준비 시작", "귀사", None),
     ("A 서명 서류 발송", "율촌", "회신 후 약 1주"),
     ("서명·공증·아포스티유 (A·B)", "귀사", "—"),
     ("원본 발송", "귀사", "—"),
     ("해외직접투자 신고 (한국 외국환은행)", "귀사", "—"),
     ("중국 설립등기 (영업집조 발급)", "율촌", "—"),
     ("인감 제작·계좌 개설·세무 등록·자본금 송금", "율촌·귀사", "—")]
for i, (a, b, c) in enumerate(S, 1): row(t, i, None, [[(a, G)]], b, None if c is None else [[(c, G) if c == "—" else c]])
# 5. 미리 알아 두실 사항
P(40, "해외직접투자 신고")
P(41, ("한국 법인이 중국 회사에 출자하려면 자본금을 보내기 전에 외국환은행에 해외직접투자 신고를 해야 합니다.", G))
P(42, "신고가 끝나기 전에는 자본금을 송금할 수 없으므로, 거래 은행과 미리 협의해 주십시오.")
P(43, "신고에 필요한 서류는 은행마다 다르므로 율촌이 은행과 확인해 안내합니다.")
drop(tb[5]._element); drop(pp[45]._element)                 # 등록세 표와 그 주석
P(47, ("중국 설립등기에는 등록 주소의 임대차계약서와 소유권 증명이 필요합니다.", G))
P(48, "비용 구성: 보증금·월세·관리비·중개수수료 등이며 지역 관행에 따라 다릅니다.")
P(49, "설립 전에는 회사 명의로 계약할 수 없으므로, 계약 방식을 미리 정해 주십시오", ("(귀사 명의로 계약한 뒤 회사 명의로 바꾸는 방식 등)", G), ".")
P(50, ("공유 등록 주소(가상 사무실)는 지역에 따라 허용 여부가 다르므로 율촌이 확인해 안내합니다.", G))
P(51, "설립등기 전 비용")
P(52, ("자본금은 중국 회사의 자본금 계좌가 열린 뒤에 송금합니다.", G), " 사무실 보증금, 대행 수수료 등 등기 전 비용은 지급 방식을 미리 정해 주십시오.")
set_cell(tb[6].rows[0].cells[0], [["□ 율촌이 먼저 대납하고 등기 후 잔금과 함께 청구 (소액만)"], ["□ 귀사가 예상 비용을 율촌에 미리 송금"], ["□ 귀사의 중국 관계사가 직접 지급"], ["□ 기타 ________"]])
P(55, "사무실 보증금 등 큰 금액은 귀사의 사전 송금이나 중국 관계사 지급을 권합니다.")
P(56, "귀사가 한국에서 부담하는 비용")
P(57, "공증료 (공증사무소 요금 기준)")
P(58, "아포스티유 수수료")
P(59, "국내 발송비")
P(60, ("중국에서 발생하는 실비(등기 대행, 인감 제작 등)는 율촌이 실제 금액으로 청구합니다.", G))
P(62, "해외직접투자 신고가 끝나기 전에는 자본금을 송금할 수 없습니다.")
P(63, ("임원이 중국에 상주하려면 중국 취업 비자와 거류 허가가 필요합니다.", G))
for i in range(70, 64, -1): drop(pp[i]._element)            # 별첨(과밀억제권역) 삭제
out = os.path.join(TPL, "request_cnwfoe_ko.docx"); d.save(out); print(out)
