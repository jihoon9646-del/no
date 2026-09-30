import sys, openpyxl
REQ = {"상품코드": ["상품코드","style no","style","품번","상품번호","item code"],
       "색상": ["색상","color","colour","컬러"],
       "사이즈": ["사이즈","size","치수"],
       "개수": ["개수","수량","q'ty","qty","quantity","pcs"]}
def norm(v): return str(v).strip().lower() if v is not None else ""
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
for ws in wb.worksheets:
    rows = list(ws.iter_rows(values_only=True))
    hdr_i, cmap = None, {}
    for i, r in enumerate(rows[:20]):
        m = {}
        for j, c in enumerate(r):
            for k, syn in REQ.items():
                if norm(c) in syn: m[k] = (j, c)
        if len(m) >= 2: hdr_i, cmap = i, m; break
    print(f"[시트] {ws.title}")
    if hdr_i is None: print("  헤더를 찾지 못함"); continue
    print(f"  헤더 위치: {hdr_i+1}행")
    for k in REQ: print(f"  {k:4} ← " + (f"'{cmap[k][1]}' 열" if k in cmap else "❌ 열 없음"))
    data = [r for r in rows[hdr_i+1:] if any(v not in (None,"") for v in r)]
    print(f"  데이터 행 수: {len(data)}")
    issues = 0
    for n, r in enumerate(data):
        excel_row = hdr_i + 2 + n
        probs = []
        for k,(j,_) in cmap.items():
            v = r[j] if j < len(r) else None
            if v in (None, ""): probs.append(f"{k} 비어있음")
            elif k == "개수" and not isinstance(v,(int,float)): probs.append(f"개수 숫자 아님('{v}')")
        if probs:
            issues += 1
            key = r[cmap["상품코드"][0]] if "상품코드" in cmap else ""
            print(f"  - 엑셀 {excel_row}행 [{key or '코드없음'}]: " + ", ".join(probs))
    print(f"  → 문제 행 {issues}건 / 전체 {len(data)}건")
