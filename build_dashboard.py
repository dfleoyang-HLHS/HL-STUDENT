"""讀取 students.json，產生國中來源分析網頁 index.html。

網頁只嵌入「年級、班級、國中、報部身份、原住民族別」五個欄位，不含學號與姓名。
執行：uv run build_dashboard.py
"""
import json
from pathlib import Path

BASE = Path(__file__).parent
data = json.loads((BASE / "students.json").read_text(encoding="utf-8"))

schools = {}
records = []
# 報部身份依人數多寡排序，records 內以索引表示
identities = list(data["meta"].get("moe_identity_counts", {}).keys())
# 原住民族別依人數多寡排序，records 內以索引表示，非原住民為 -1
tribes = list(data["meta"].get("indigenous_tribe_counts", {}).keys())
for s in data["students"]:
    jh = s["junior_high"]
    schools[jh["code"]] = {"name": jh["name"], "county": jh["county"], "ownership": jh["ownership"]}
    records.append([s["grade"], s["class_no"], jh["code"], identities.index(s["moe_identity"]),
                    tribes.index(s["indigenous_tribe"]) if s.get("indigenous_tribe") else -1])

payload = {
    "school": data["meta"]["school"],
    "schoolYearRoc": data["meta"]["school_year_roc"],
    "cohorts": [{"grade": c["grade"], "name": c["grade_name"], "entryYear": c["entry_year"]} for c in data["cohorts"]],
    "classTypes": data["meta"].get("class_types", {}),
    "identities": identities,
    "tribes": tribes,
    "schools": schools,
    "records": records,
}
template = (BASE / "index.template.html").read_text(encoding="utf-8")
html = template.replace("__DATA__", json.dumps(payload, ensure_ascii=False, separators=(",", ":")))
(BASE / "index.html").write_text(html, encoding="utf-8")
print(f"寫出 index.html（{len(records)} 筆、{len(schools)} 所國中）")
