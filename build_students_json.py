# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl"]
# ///
"""將三個年度的「學生資料一覽表」Excel 轉成 students.json，
並依「全校學生資料報部身份一覽表」加上報部身份、
依「原住民族別學生資料一覽表」加上原住民族別。

執行：uv run build_students_json.py
"""
import glob
import json
import re
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl

BASE = Path(__file__).parent
CURRENT_SCHOOL_YEAR = 2026  # 目前學年度（西元）＝115 學年度

# 西元入學年 -> 年級
GRADE_NAMES = {1: "高一", 2: "高二", 3: "高三"}

COHORT_GLOB = "高?學生資料一覽表_20*.xlsx"
IDENTITY_FILE = "全校學生資料報部身份一覽表.xlsx"
TRIBE_GLOB = "原住民族別學生資料一覽表*.xlsx"  # 有多份時取檔名排序最後（日期最新）的一份

# 各年級的特殊班（依班序），其餘為普通班
SPECIAL_CLASSES = {10: "音樂班", 11: "數理資優班", 12: "體育班"}


def split_name(full: str):
    """原住民學生姓名常為「漢名+族名」，以第一個拉丁字母切開。"""
    full = full.strip()
    m = re.search(r"[A-Za-z]", full)
    if not m:
        return full, None
    zh = full[: m.start()].strip(" ･·")
    return zh, full[m.start():].strip()


def parse_school(raw: str):
    code, _, name = raw.strip().partition(" ")
    county = name[:3]
    if "私立" in name or "財團法人" in name:
        ownership = "私立"
    elif "國立" in name:
        ownership = "國立"
    else:
        ownership = "公立"
    return {"code": code, "name": name, "county": county, "ownership": ownership}


def load_identities(path: Path):
    """讀取報部身份一覽表（分頁報表：每頁先有班級代號列，再列座號/學號/姓名/報部身份）。

    回傳 {學號: {"class_code", "seat_no", "name", "identity"}}。
    """
    ws = openpyxl.load_workbook(path, data_only=True).active
    result, cls = {}, None
    for row in ws.iter_rows(values_only=True):
        b, c, e, g = (str(v).strip() if v is not None else None for v in (row[1], row[2], row[4], row[6]))
        if b and re.fullmatch(r"\d{3}", b) and not c:
            cls = b
        elif c and re.fullmatch(r"\d{6}", c):
            result[c] = {"class_code": cls, "seat_no": int(b), "name": e, "identity": g}
    return result


def load_tribes(path: Path):
    """讀取原住民族別學生資料一覽表（欄位：學號、座號、姓名、原住民族別）。

    回傳 {學號: {"seat_no", "name", "tribe"}}，非原住民學生 tribe 為 None。
    """
    ws = openpyxl.load_workbook(path, data_only=True).active
    header = [c.value for c in ws[1]]
    result = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        rec = dict(zip(header, row))
        if not rec.get("學號"):
            continue
        tribe = str(rec["原住民族別"]).strip() if rec.get("原住民族別") else None
        result[str(rec["學號"]).strip()] = {
            "seat_no": int(rec["座號"]),
            "name": str(rec["姓名"]).strip(),
            "tribe": tribe,
        }
    return result


def load_file(path: Path):
    entry_year = int(re.search(r"(20\d\d)", path.name).group(1))
    grade = CURRENT_SCHOOL_YEAR - entry_year + 1
    ws = openpyxl.load_workbook(path, data_only=True).active
    header = [c.value for c in ws[1]]
    students, warnings = [], []
    class_no, prev_seat = 0, None
    for row in ws.iter_rows(min_row=2, values_only=True):
        rec = dict(zip(header, row))
        if not rec.get("學號"):
            continue
        seat = int(rec["座號"])
        # 檔案未提供班級欄位：依座號重新從 01 開始判定為下一班
        if prev_seat is None or seat <= prev_seat:
            class_no += 1
        elif seat != prev_seat + 1:
            warnings.append(f"{path.name}: 第 {class_no} 班座號 {prev_seat} → {seat} 不連續")
        prev_seat = seat

        name_zh, name_ind = split_name(str(rec["姓名"]))
        notes = [rec[k] for k in ("身份註記", "身份註記1", "身份註記2") if rec.get(k)]
        students.append({
            "student_id": str(rec["學號"]).strip(),
            "entry_year": entry_year,
            "entry_year_roc": entry_year - 1911,
            "grade": grade,
            "grade_name": GRADE_NAMES[grade],
            "class_no": class_no,
            "class_code": f"{grade}{class_no:02d}",
            "class_type": SPECIAL_CLASSES.get(class_no, "普通班"),
            "is_special_class": class_no in SPECIAL_CLASSES,
            "seat_no": seat,
            "name": str(rec["姓名"]).strip(),
            "name_zh": name_zh,
            "name_indigenous": name_ind,
            "junior_high": parse_school(str(rec["入學學校"])),
            "identity_notes": notes,
        })
    return entry_year, grade, students, warnings


def main():
    all_students, cohorts, all_warnings, sources = [], [], [], []
    for f in sorted(BASE.glob(COHORT_GLOB)):
        entry_year, grade, students, warnings = load_file(f)
        sources.append(f.name)
        all_warnings += warnings
        all_students += students
        by_class = Counter(s["class_code"] for s in students)
        cohorts.append({
            "entry_year": entry_year,
            "entry_year_roc": entry_year - 1911,
            "grade": grade,
            "grade_name": GRADE_NAMES[grade],
            "student_count": len(students),
            "class_count": len(by_class),
            "classes": [
                {"class_code": k, "class_type": SPECIAL_CLASSES.get(int(k[1:]), "普通班"), "student_count": v}
                for k, v in sorted(by_class.items())
            ],
        })

    all_students.sort(key=lambda s: (-s["grade"], s["class_no"], s["seat_no"]))

    # 報部身份：以學號對應，並用報表上的正式班級代號、座號、姓名交叉檢查
    identity_path = BASE / IDENTITY_FILE
    identities = load_identities(identity_path) if identity_path.exists() else {}
    if identities:
        sources.append(IDENTITY_FILE)
    for s in all_students:
        info = identities.get(s["student_id"])
        s["moe_identity"] = info["identity"] if info else None
        if not identities:
            continue
        if not info:
            all_warnings.append(f"{s['student_id']} 在報部身份一覽表中找不到")
            continue
        for key in ("class_code", "seat_no", "name"):
            if info[key] != s[key]:
                all_warnings.append(f"{s['student_id']} 的 {key} 不一致：學生資料一覽表為 {s[key]}，報部身份一覽表為 {info[key]}")
    known = {s["student_id"] for s in all_students}
    for sid in identities.keys() - known:
        all_warnings.append(f"報部身份一覽表的 {sid} 不在三個年級的學生資料一覽表中")

    # 原住民族別：以學號對應，並檢查座號、姓名，以及與報部身份「原住民生」是否一致
    tribe_files = sorted(BASE.glob(TRIBE_GLOB))
    tribes = load_tribes(tribe_files[-1]) if tribe_files else {}
    if tribes:
        sources.append(tribe_files[-1].name)
    for s in all_students:
        info = tribes.get(s["student_id"])
        s["indigenous_tribe"] = info["tribe"] if info else None
        if not tribes:
            continue
        if not info:
            all_warnings.append(f"{s['student_id']} 在原住民族別學生資料一覽表中找不到")
            continue
        for key in ("seat_no", "name"):
            if info[key] != s[key]:
                all_warnings.append(f"{s['student_id']} 的 {key} 不一致：學生資料一覽表為 {s[key]}，原住民族別表為 {info[key]}")
        if s["moe_identity"] is not None and (info["tribe"] is not None) != (s["moe_identity"] == "原住民生"):
            all_warnings.append(f"{s['student_id']} 報部身份為 {s['moe_identity']}，但原住民族別為 {info['tribe'] or '空白'}")
    for sid in tribes.keys() - known:
        all_warnings.append(f"原住民族別表的 {sid} 不在三個年級的學生資料一覽表中")

    cohorts.sort(key=lambda c: c["entry_year"])

    data = {
        "meta": {
            "school": "國立花蓮高級中學",
            "school_year": CURRENT_SCHOOL_YEAR,
            "school_year_roc": CURRENT_SCHOOL_YEAR - 1911,
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "source_files": sources,
            "total_students": len(all_students),
            "class_types": {str(k): v for k, v in SPECIAL_CLASSES.items()},
            "moe_identity_counts": dict(Counter(s["moe_identity"] for s in all_students).most_common()),
            "indigenous_tribe_counts": dict(
                Counter(s["indigenous_tribe"] for s in all_students if s["indigenous_tribe"]).most_common()
            ),
            "notes": [
                "原始檔無班級欄位，class_no 依座號重新從 01 開始推定（依檔案列序）。",
                "每個年級第 10 班為音樂班、第 11 班為數理資優班、第 12 班為體育班（class_type），其餘為普通班。",
                "class_code = 年級 + 兩位數班序，例如 305 表示高三第 5 班；已與報部身份一覽表上的正式班級代號比對一致。",
                "moe_identity 為報部身份（一般生、原住民生、身心障礙生等），來自全校學生資料報部身份一覽表，以學號對應。",
                "indigenous_tribe 為原住民族別（阿美族、太魯閣族等），來自原住民族別學生資料一覽表，以學號對應；非原住民學生為 null。",
                "name_zh / name_indigenous 由姓名欄以第一個拉丁字母切分，僅供參考。",
            ],
            "warnings": all_warnings,
        },
        "cohorts": cohorts,
        "students": all_students,
    }
    out = BASE / "students.json"
    out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"寫出 {out}，共 {len(all_students)} 名學生")
    for w in all_warnings:
        print("警告：", w)


if __name__ == "__main__":
    main()
