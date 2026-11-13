# 정리_상담원장.py — ../messy/상담원장_2025.xlsx (시트 «원장») → 정리본/
#
# 결정 기록: my-wiki/decisions/d-006-상담원장-정리-다섯-가지.md
#   d-006 1. «종결»은 «완료»로, «대기»는 «보류»로 합친다
#   d-006 2. 처리시간 음수는 지우지 않고 값을 그대로 둔다. `_표시` 칸에 이유를 적는다
#   d-006 3. 만족도 6~10점은 지우지 않고 `_표시` 칸에 이유를 적는다. 평균 만족도에서는 표시한 줄을 뺀다
#   d-006 4. 칸 이름은 원본 그대로 둔다 (`_표시`는 맨 끝에 하나 더할 뿐)
#   d-006 5. 고객번호 앞뒤 공백을 없앤다
# 위 다섯 가지 밖의 처리는 첫 정리(정리본/상담원장_정리_기록.md)의 방식을 그대로 따른다.
#
# 원본은 읽기만 한다. 쓰는 파일: 정리본/상담원장_정리.csv · 정리본/지운줄.csv · 정리본/변경요약.md
# 돌리기: python 정리/정리_상담원장.py   (cleaner 폴더에서)

import collections
import csv
import re
import sys
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
SRC = HERE.parent.parent / "messy" / "상담원장_2025.xlsx"
OUT = HERE.parent / "정리본"
OUT_CLEAN = OUT / "상담원장_정리.csv"
OUT_REMOVED = OUT / "지운줄.csv"
OUT_SUMMARY = OUT / "변경요약.md"

CHANNEL = {"유선": "전화", "전화상담": "전화", "TEL": "전화",
           "chatbot": "챗봇", "Chatbot": "챗봇", "챗 봇": "챗봇",
           "app": "앱", "APP": "앱", "모바일앱": "앱",
           "e-mail": "이메일", "E-mail": "이메일", "메일": "이메일"}
REGION = {"경기도": "경기", "서울시": "서울", "서울특별시": "서울"}
STATUS = {"처리완료": "완료", "종결": "완료",   # d-006 1 (종결)
          "처리중": "진행중",
          "대기": "보류"}                         # d-006 1 (대기)
REASK = {"y": "Y", "예": "Y", "TRUE": "Y", "n": "N", "아니오": "N", "FALSE": "N"}
FLAG_NEG = "처리시간 음수"            # d-006 2
FLAG_SAT = "만족도 1~5점 밖"          # d-006 3
SEC_OVER = 180          # 상담시간이 이 값을 넘으면 초로 적힌 것으로 보고 분으로 바꾼다


def norm_date(s):
    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
    if not m:
        sys.exit(f"알 수 없는 날짜 꼴: {s!r}")
    return "%s-%02d-%02d" % (m[1], int(m[2]), int(m[3]))


def main():
    wb = openpyxl.load_workbook(SRC, read_only=True)
    rows = list(wb["원장"].iter_rows(values_only=True))
    wb.close()
    header = list(rows[0])
    raw = rows[1:]

    recs = []
    for i, t in enumerate(raw):
        r = dict(zip(header, t))
        r["_line"] = i + 2      # 엑셀 행 번호 (머리글이 1행)
        r["_raw"] = t
        r["_flag"] = []
        r["_why"] = collections.Counter()
        recs.append(r)

    def change(r, col, new, why):
        if r[col] != new:
            r[col] = new
            r["_why"][why] += 1

    # 앞뒤 공백 — 고객번호는 d-006 5, 나머지 칸은 첫 정리 방식
    for r in recs:
        for c in header:
            v = r[c]
            if isinstance(v, str) and v != v.strip():
                change(r, c, v.strip(), "고객번호 앞뒤 공백 제거 (d-006 5)" if c == "고객번호" else f"{c} 앞뒤 공백 제거")

    removed = []
    kept = []
    for r in recs:
        if r["상담원ID"] == "AG999":
            removed.append((r, "테스트 접수 (상담원ID AG999)"))
        else:
            kept.append(r)

    for r in kept:
        change(r, "접수일", norm_date(r["접수일"]), "접수일 꼴 통일")
        change(r, "접수채널", CHANNEL.get(r["접수채널"], r["접수채널"]), "접수채널 이름 통일")
        change(r, "상담원팀", r["상담원팀"].replace(" ", ""), "상담원팀 띄어쓰기 통일")
        change(r, "지역", REGION.get(r["지역"], r["지역"]), "지역 이름 통일")
        if isinstance(r["월정액"], str):
            change(r, "월정액", int(re.sub(r"\D", "", r["월정액"])), "월정액 «원»·쉼표 떼고 숫자로")
        st = r["처리상태"]
        why = {"종결": "처리상태 종결 → 완료 (d-006 1)", "대기": "처리상태 대기 → 보류 (d-006 1)",
               "처리완료": "처리상태 처리완료 → 완료", "처리중": "처리상태 처리중 → 진행중"}.get(st)
        if why:
            change(r, "처리상태", STATUS[st], why)
        if r["재문의"] is not None:
            change(r, "재문의", REASK.get(r["재문의"], r["재문의"]), "재문의 Y/N 통일")

        s = r["만족도"]
        if s in (0, "-", "미응답"):
            change(r, "만족도", None, "만족도 미응답(0·-·미응답) → 빈칸")
        elif isinstance(s, str) and s.endswith("점"):
            change(r, "만족도", int(s[:-1]), "만족도 «점» 떼고 숫자로")
        elif isinstance(s, int) and not 1 <= s <= 5:
            r["_flag"].append(FLAG_SAT)                     # d-006 3: 값은 그대로
            r["_why"]["만족도 1~5점 밖 → 값 그대로, _표시 (d-006 3)"] += 1

        t = r["상담시간"]
        if t == 9999:
            change(r, "상담시간", None, "상담시간 9999 → 빈칸")
        elif t is not None and t > SEC_OVER:
            change(r, "상담시간", round(t / 60, 2), "상담시간 초 → 분 (180 넘는 값)")

        if r["처리시간"] is not None and r["처리시간"] < 0:
            r["_flag"].append(FLAG_NEG)                     # d-006 2: 값은 그대로
            r["_why"]["처리시간 음수 → 값 그대로, _표시 (d-006 2)"] += 1

    # 완전 중복 (정리 뒤 원래 24칸이 모두 같음)
    first = {}
    uniq = []
    for r in kept:
        k = tuple(r[c] for c in header)
        if k in first:
            removed.append((r, f"완전 중복 (정리 뒤 {first[k]}줄과 모든 칸 같음)"))
        else:
            first[k] = r["_line"]
            uniq.append(r)

    # 같은 접수번호 — 최종수정일이 가장 늦은 줄만
    groups = collections.defaultdict(list)
    for r in uniq:
        groups[r["접수번호"]].append(r)
    final = []
    for v in groups.values():
        m = max(r["최종수정일"] for r in v)
        top = [r for r in v if r["최종수정일"] == m]
        if len(top) > 1:
            sys.exit(f"최종수정일이 같은 줄이 둘 이상: {[r['_line'] for r in top]}")
        win = top[0]
        final.append(win)
        for r in v:
            if r is not win:
                removed.append((r, f"옛 상태 줄 (같은 접수번호의 {win['_line']}줄이 최종수정일 더 늦음)"))
    final.sort(key=lambda r: r["_line"])
    removed.sort(key=lambda x: x[0]["_line"])
    assert len(final) + len(removed) == len(raw)

    OUT.mkdir(exist_ok=True)
    cell = lambda v: "" if v is None else v
    with open(OUT_CLEAN, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(header + ["_표시"])                       # d-006 4
        for r in final:
            w.writerow([cell(r[c]) for c in header] + [" · ".join(r["_flag"])])
    with open(OUT_REMOVED, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(header + ["원본줄", "지운 이유"])
        for r, why in removed:
            w.writerow([cell(v) for v in r["_raw"]] + [r["_line"], why])

    # 다섯 숫자 + 고객 수
    n = len(final)
    ct = [r["상담시간"] for r in final if r["상담시간"] is not None]
    rq = [r["재문의"] for r in final if r["접수일"] >= "2025-03-01" and r["재문의"] in ("Y", "N")]
    sat = [r["만족도"] for r in final if r["만족도"] is not None and FLAG_SAT not in r["_flag"]]
    done = sum(r["처리상태"] == "완료" for r in final)
    customers = len({r["고객번호"] for r in final})
    status = collections.Counter(r["처리상태"] for r in final)

    why_rows = collections.Counter()
    for r in final:
        for k in r["_why"]:
            why_rows[k] += 1
    why_removed = collections.Counter(re.sub(r" \(.*", "", w) for _, w in removed)

    L = []
    L.append("# 상담원장 변경요약\n")
    L.append(f"- 원본: `../messy/상담원장_2025.xlsx` 시트 `원장` (읽기만 함)")
    L.append(f"- 스크립트: `정리/정리_상담원장.py` · 결정 기록: `my-wiki/decisions/d-006-상담원장-정리-다섯-가지.md`")
    L.append(f"- 결과: `정리본/상담원장_정리.csv` (원래 24칸 + 맨 끝 `_표시`) · `정리본/지운줄.csv`\n")
    L.append("## 줄 수\n")
    L.append("| | 줄 |\n|---|---:|")
    L.append(f"| 원본 | {len(raw):,} |")
    for k in ["테스트 접수", "완전 중복", "옛 상태 줄"]:
        L.append(f"| − {k} | −{why_removed[k]:,} |")
    L.append(f"| **정리본** | **{n:,}** |\n")
    L.append("## 바꾼 것 — 정리본에 남은 줄 기준, 그 변경이 일어난 줄 수\n")
    L.append("| 바꾼 것 | 줄 |\n|---|---:|")
    for k, v in sorted(why_rows.items(), key=lambda x: (-x[1], x[0])):
        L.append(f"| {k} | {v:,} |")
    L.append("")
    L.append(f"- **상담시간 칸**: 값이 180을 넘는 {why_rows['상담시간 초 → 분 (180 넘는 값)']:,}줄은 "
             f"원본에 초 단위로 적혀 있어 60으로 나눠 **분 단위로 바꿨다** (소수 둘째 자리까지). 정리본의 상담시간은 모두 분이다")
    L.append(f"- `_표시` 칸: `{FLAG_NEG}` {sum(FLAG_NEG in r['_flag'] for r in final)}줄 · "
             f"`{FLAG_SAT}` {sum(FLAG_SAT in r['_flag'] for r in final)}줄. 값은 원본 그대로 두었다\n")
    L.append("## 다섯 숫자 (정리본 기준)\n")
    L.append("| 지표 | 값 | 셈 |\n|---|---:|---|")
    L.append(f"| 상담 건수 | {n:,}건 | 정리본 줄 수 |")
    L.append(f"| 평균 상담시간 | {sum(ct)/len(ct):.2f}분 | {len(ct):,}건 평균 |")
    L.append(f"| 재문의율 | {rq.count('Y')/len(rq)*100:.2f}% | Y {rq.count('Y'):,} / {len(rq):,} — 2025-03-01 이후 접수분 |")
    L.append(f"| 평균 만족도 | {sum(sat)/len(sat):.3f}점 | {len(sat):,}건 (1~5점, `_표시` 줄 뺌) |")
    L.append(f"| 완료율 | {done/n*100:.2f}% | 완료 {done:,} / {n:,} |")
    L.append(f"\n- 처리상태: " + " · ".join(f"{k} {v:,}" for k, v in status.most_common()))
    L.append(f"- 고객 수(서로 다른 고객번호): {customers:,}명")
    OUT_SUMMARY.write_text("\n".join(L) + "\n", encoding="utf-8")

    print(f"정리본 {n}줄 · 지운줄 {len(removed)}줄 · 고객 {customers}명")


if __name__ == "__main__":
    main()
