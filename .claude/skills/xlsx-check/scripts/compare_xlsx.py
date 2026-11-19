"""원본(서식/앞 날짜 파일)과 생성된 xlsx의 '구조'를 비교한다.

셀 값은 출력하지 않는다 (이름·주민등록번호 보호). 바뀐 셀은 주소만 보여 준다.
표준 라이브러리만 사용: python -I compare_xlsx.py <원본.xlsx> <생성.xlsx>
종료 코드: 0 = 구조 동일, 1 = 차이 있음, 2 = 파일이 열리지 않음
"""
import posixpath
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}
R_ID = "{%s}id" % NS["r"]


def rels_of(z, part):
    folder, name = posixpath.split(part)
    path = posixpath.join(folder, "_rels", name + ".rels")
    if path not in z.namelist():
        return {}
    out = {}
    for rel in ET.fromstring(z.read(path)).findall("rel:Relationship", NS):
        target = rel.get("Target")
        full = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
        out[rel.get("Id")] = (rel.get("Type").rsplit("/", 1)[-1], full)
    return out


def anchors(z, drawing_part):
    root = ET.fromstring(z.read(drawing_part))
    result = []
    for tag in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
        for a in root.findall("xdr:" + tag, NS):
            kind = "pic" if a.find("xdr:pic", NS) is not None else "shape"
            f = a.find("xdr:from", NS)
            pos = (f.findtext("xdr:col", "", NS), f.findtext("xdr:row", "", NS)) if f is not None else ("", "")
            result.append((kind, tag, pos))
    return result


def cells(sheet_root):
    out = {}
    for c in sheet_root.iter("{%s}c" % NS["m"]):
        has_formula = c.find("m:f", NS) is not None
        value = c.findtext("m:v", None, NS)
        inline = "".join(t.text or "" for t in c.iter("{%s}t" % NS["m"]))
        out[c.get("r")] = (has_formula, c.get("t"), value, inline)
    return out


def summarize(path):
    z = zipfile.ZipFile(path)
    bad = z.testzip()
    if bad:
        raise ValueError(f"손상된 파트: {bad}")
    for name in z.namelist():
        if name.endswith((".xml", ".rels")):
            ET.fromstring(z.read(name))  # XML 문법 오류면 여기서 예외
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    wb_rels = rels_of(z, "xl/workbook.xml")
    calc = wb.find("m:calcPr", NS)
    info = {
        "media": sorted(n for n in z.namelist() if n.startswith("xl/media/")),
        "fullCalcOnLoad": calc is not None and calc.get("fullCalcOnLoad") in ("1", "true"),
        "definedNames": sorted((d.get("name"), d.text) for d in wb.findall("m:definedNames/m:definedName", NS)),
        "sheets": {},
    }
    for s in wb.findall("m:sheets/m:sheet", NS):
        part = wb_rels[s.get(R_ID)][1]
        root = ET.fromstring(z.read(part))
        drawing = []
        for kind, target in rels_of(z, part).values():
            if kind == "drawing":
                drawing += anchors(z, target)
        info["sheets"][s.get("name")] = {
            "merged": sorted(m.get("ref") for m in root.findall("m:mergeCells/m:mergeCell", NS)),
            "cols": [ET.tostring(c) for c in root.findall("m:cols/m:col", NS)],
            "anchors": drawing,
            "cells": cells(root),
        }
    return info


def compare(a, b):
    diffs, notes = [], []
    if len(a["media"]) != len(b["media"]):
        notes.append(f"이미지 파일 수 {len(a['media'])} → {len(b['media'])} (사진 교체 시 정상일 수 있음)")
    if a["definedNames"] != b["definedNames"]:
        diffs.append("정의된 이름/인쇄 영역이 다름")
    if not b["fullCalcOnLoad"]:
        diffs.append("생성 파일에 fullCalcOnLoad 꺼짐 (수식 재계산 안 됨)")
    if list(a["sheets"]) != list(b["sheets"]):
        diffs.append(f"시트 목록 다름: {list(a['sheets'])} → {list(b['sheets'])}")
    for name in a["sheets"]:
        if name not in b["sheets"]:
            continue
        sa, sb = a["sheets"][name], b["sheets"][name]
        for key, label in (("merged", "병합"), ("cols", "열 너비")):
            if sa[key] != sb[key]:
                diffs.append(f"[{name}] {label} 다름")
        pa = [x for x in sa["anchors"] if x[0] == "pic"]
        pb = [x for x in sb["anchors"] if x[0] == "pic"]
        if len(pa) != len(pb):
            diffs.append(f"[{name}] 그림 개수 {len(pa)} → {len(pb)}")
        elif [x[2] for x in pa] != [x[2] for x in pb]:
            notes.append(f"[{name}] 그림 시작 위치 바뀜 (그림 위치 필드를 쓴 경우 정상)")
        if len(sa["anchors"]) - len(pa) != len(sb["anchors"]) - len(pb):
            diffs.append(f"[{name}] 도형 개수 다름")
        formula_lost = [r for r, v in sa["cells"].items() if v[0] and not sb["cells"].get(r, (False,))[0]]
        if formula_lost:
            diffs.append(f"[{name}] 수식이 사라진 칸: {', '.join(formula_lost[:20])}")
        changed = sorted(r for r in set(sa["cells"]) | set(sb["cells"])
                         if sa["cells"].get(r, (None,) * 4)[1:] != sb["cells"].get(r, (None,) * 4)[1:]
                         and not sb["cells"].get(r, (False,))[0])
        notes.append(f"[{name}] 값이 바뀐 칸 {len(changed)}개: {', '.join(changed[:40])}{' …' if len(changed) > 40 else ''}")
    return diffs, notes


def main():
    if len(sys.argv) != 3:
        print("사용법: python -I compare_xlsx.py <원본.xlsx> <생성.xlsx>")
        return 2
    try:
        a, b = summarize(sys.argv[1]), summarize(sys.argv[2])
    except Exception as e:  # 열리지 않는 파일은 그대로 보고
        print(f"열기 실패: {e}")
        return 2
    print("열기: 두 파일 모두 zip·XML 정상")
    diffs, notes = compare(a, b)
    for n in notes:
        print("  참고 ", n)
    for d in diffs:
        print("  차이 ", d)
    print("결과:", "구조 동일" if not diffs else f"차이 {len(diffs)}건")
    return 1 if diffs else 0


if __name__ == "__main__":
    sys.exit(main())
