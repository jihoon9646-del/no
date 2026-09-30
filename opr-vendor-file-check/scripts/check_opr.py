#!/usr/bin/env python3
"""OPR 업체 자료 누락 점검기.

사용법:
  python check_opr.py scan <폴더> [--config config.json]
      폴더 "바로 아래"의 새 파일만 점검하고 결과를 JSON으로 출력한다. 파일은 옮기지 않는다.
  python check_opr.py finalize <폴더> <파일명> <결과> <문제건수>
      결과 메일을 보낸 뒤 호출한다. 파일을 처리완료/ 또는 열수없음/ 으로 옮기고 처리기록.csv에 남긴다.
      <결과>: 누락있음 | 누락없음 | 열수없음
"""
import csv, datetime, json, os, re, shutil, sys

HERE = os.path.dirname(os.path.abspath(__file__))
LOG_NAME = "처리기록.csv"
DONE_DIR, FAIL_DIR = "처리완료", "열수없음"
TABLE_EXT = {".xlsx", ".xlsm", ".xls", ".csv"}
LOG_HEADER = ["처리시각", "파일명", "크기", "수정시각", "결과", "문제건수"]


def load_config(path=None):
    path = path or os.path.join(HERE, "..", "config.json")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def norm(v):
    return re.sub(r"\s+", " ", str(v)).strip().lower() if v is not None else ""


def file_key(path):
    st = os.stat(path)
    return str(st.st_size), datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S")


def read_log(folder):
    p = os.path.join(folder, LOG_NAME)
    if not os.path.exists(p):
        return set()
    with open(p, encoding="utf-8-sig") as f:
        return {(r["파일명"], r["크기"], r["수정시각"]) for r in csv.DictReader(f)}


def sniff_protection(path):
    """암호/DRM 여부를 파일 앞부분으로 판단한다. 해제는 시도하지 않는다."""
    with open(path, "rb") as f:
        head = f.read(8)
    ext = os.path.splitext(path)[1].lower()
    if head.startswith(b"SCDSA"):
        return "문서보안(DRM) 암호화 파일입니다 (SCDSA 표식, SoftCamp 계열로 추정)."
    if ext in {".xlsx", ".xlsm"} and head.startswith(b"\xd0\xcf\x11\xe0"):
        return "엑셀 비밀번호(열기 암호)가 걸린 파일입니다."
    if ext in {".xlsx", ".xlsm"} and not head.startswith(b"PK"):
        return "확장자는 엑셀이지만 엑셀 형식이 아닙니다 (보안 프로그램 암호화 또는 손상 가능성)."
    return None


def load_rows(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        for enc in ("utf-8-sig", "cp949", "euc-kr"):
            try:
                with open(path, encoding=enc, newline="") as f:
                    return [("CSV", [tuple(r) for r in csv.reader(f)])]
            except UnicodeDecodeError:
                continue
        raise ValueError("CSV 인코딩을 인식하지 못했습니다.")
    if ext == ".xls":
        try:
            import xlrd
        except ImportError:
            raise ValueError("옛 엑셀 형식(.xls)은 읽기 도구가 없습니다. 엑셀에서 .xlsx로 저장해 달라고 요청하세요.")
        wb = xlrd.open_workbook(path)
        return [(s.name, [tuple(s.row_values(i)) for i in range(s.nrows)]) for s in wb.sheets()]
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    return [(ws.title, list(ws.iter_rows(values_only=True))) for ws in wb.worksheets]


def to_number(v):
    if isinstance(v, (int, float)):
        return v
    s = str(v).replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def check_sheet(title, rows, cfg):
    fields = cfg["required_fields"]
    syn = {k: {norm(x) for x in v} for k, v in fields.items()}
    hdr_i, cmap = None, {}
    for i, r in enumerate(rows[:30]):
        m = {}
        for j, c in enumerate(r):
            for k, s in syn.items():
                if k not in m and norm(c) in s:
                    m[k] = (j, str(c).strip())
        if len(m) >= 2:
            hdr_i, cmap = i, m
            break
    if hdr_i is None:
        return {"sheet": title, "header_found": False}
    data = [(hdr_i + 2 + n, r) for n, r in enumerate(rows[hdr_i + 1:]) if any(v not in (None, "") for v in r)]
    missing_cols = [k for k in fields if k not in cmap]
    issues = []
    for excel_row, r in data:
        probs = []
        for k, (j, _) in cmap.items():
            v = r[j] if j < len(r) else None
            if v is None or str(v).strip() == "":
                probs.append(f"{k} 비어있음")
            elif k in cfg.get("numeric_fields", []):
                n = to_number(v)
                if n is None:
                    probs.append(f"{k} 숫자 아님('{v}')")
                elif n <= 0:
                    probs.append(f"{k} 0 이하({v})")
        if probs:
            code = r[cmap["상품코드"][0]] if "상품코드" in cmap and cmap["상품코드"][0] < len(r) else None
            ctx = [str(x) for x in r if x not in (None, "")][:4]
            issues.append({"row": excel_row, "code": code or "", "problems": probs, "row_preview": ctx})
    return {
        "sheet": title, "header_found": True, "header_row": hdr_i + 1,
        "mapping": {k: v[1] for k, v in cmap.items()},
        "missing_columns": missing_cols, "data_rows": len(data), "issues": issues,
    }


def scan(folder, cfg):
    done = read_log(folder)
    out = []
    for name in sorted(os.listdir(folder)):
        p = os.path.join(folder, name)
        if not os.path.isfile(p) or name == LOG_NAME or name.startswith(("~$", ".")):
            continue
        ext = os.path.splitext(name)[1].lower()
        size, mtime = file_key(p)
        if (name, size, mtime) in done:
            continue
        if ext not in TABLE_EXT:
            continue  # 서명 이미지 등은 점검 대상이 아니므로 결과에 넣지 않는다
        rec = {"file": name, "size": int(size)}
        prot = sniff_protection(p)
        if prot:
            rec.update(status="열수없음", reason=prot)
            out.append(rec); continue
        try:
            sheets = [check_sheet(t, r, cfg) for t, r in load_rows(p)]
        except Exception as e:
            rec.update(status="열수없음", reason=f"파일을 읽지 못했습니다: {e}")
            out.append(rec); continue
        usable = [s for s in sheets if s["header_found"]]
        if not usable:
            rec.update(status="열수없음", reason="필수 항목 열 제목을 찾지 못했습니다 (양식 확인 필요).", sheets=sheets)
        else:
            n = sum(len(s["issues"]) for s in usable) + sum(1 for s in usable if s["missing_columns"])
            rec.update(status="누락있음" if n else "누락없음", issue_count=n, sheets=sheets)
        out.append(rec)
    return out


def finalize(folder, name, result, count):
    src = os.path.join(folder, name)
    size, mtime = file_key(src)
    dest_dir = os.path.join(folder, FAIL_DIR if result == "열수없음" else DONE_DIR)
    os.makedirs(dest_dir, exist_ok=True)
    dest = os.path.join(dest_dir, name)
    if os.path.exists(dest):
        base, ext = os.path.splitext(name)
        dest = os.path.join(dest_dir, f"{base}_{datetime.datetime.now():%H%M%S}{ext}")
    shutil.move(src, dest)
    log = os.path.join(folder, LOG_NAME)
    new = not os.path.exists(log)
    with open(log, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        if new:
            w.writerow(LOG_HEADER)
        w.writerow([f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S}", name, size, mtime, result, count])
    print(json.dumps({"moved_to": dest}, ensure_ascii=False))


if __name__ == "__main__":
    a = sys.argv[1:]
    if len(a) >= 2 and a[0] == "scan":
        cfg = load_config(a[a.index("--config") + 1] if "--config" in a else None)
        print(json.dumps(scan(a[1], cfg), ensure_ascii=False, indent=1, default=str))
    elif len(a) == 5 and a[0] == "finalize":
        finalize(a[1], a[2], a[3], a[4])
    else:
        print(__doc__); sys.exit(1)
