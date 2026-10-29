#!/usr/bin/env python3
"""Generate only synthetic fixtures and execute the packaged filler.

Never consumes customer data. For a developer or platform smoke test, not tax advice.
"""
import argparse
import csv
import hashlib
import json
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TAX_ID = "DEMO-NOT-A-REAL-TAXPAYER"


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def generate(out):
    out.mkdir(parents=True, exist_ok=False)
    template = out / "示例模板.xlsx"
    shutil.copyfile(ROOT / "assets" / "example-template.xlsx", template)
    current = [
        ["项目", "金额", "", "纳税人识别号", TAX_ID],
        ["已确认销售额", "150000.30", "", "所属期开始", "2026-07-01"],
        ["已确认应纳税额", "1500.00", "", "所属期结束", "2026-09-30"],
        ["已确认已缴税额", "200.00", "", "纳税人名称", "合成演示贸易有限公司"],
        ["已确认应补税额", "1300.00", "", "所属期文字", "2026-07-01 至 2026-09-30"],
    ]
    ytd = [
        ["项目", "金额", "", "纳税人识别号", TAX_ID],
        ["已确认销售额", "600000.80", "", "所属期开始", "2026-01-01"],
        ["已确认应纳税额", "6000.00", "", "所属期结束", "2026-09-30"],
        ["已确认已缴税额", "4500.00", "", "纳税人名称", "合成演示贸易有限公司"],
        ["已确认应补税额", "1500.00", "", "用途", "合成测试数据，非税法计算案例"],
    ]
    ledger = [
        ["业务所属日期", "纳税人识别号", "业务行号", "已确认不含税销售额"],
        ["2026-07-15", TAX_ID, "SYN-001", "100000.10"],
        ["2026-08-12", TAX_ID, "SYN-002", "49999.90"],
        ["2026-09-20", TAX_ID, "SYN-003", "0.30"],
    ]
    for name, rows in [("本期确认底稿.csv", current), ("累计确认底稿.csv", ytd), ("收入台账.csv", ledger)]:
        with (out / name).open("w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f).writerows(rows)

    context = {"taxpayer_id": TAX_ID, "period_start": "2026-07-01", "period_end": "2026-09-30",
               "period_kind": "quarter", "form_name": "合成测试申报底稿（非官方表样）", "form_version": "synthetic-v1",
               "region": "合成测试区域", "currency": "CNY", "amount_unit": "元"}
    sources = []
    for sid, name, start in [("current", "本期确认底稿.csv", "2026-07-01"),
                             ("ytd", "累计确认底稿.csv", "2026-01-01"),
                             ("ledger", "收入台账.csv", "2026-07-01")]:
        sc = {k: context[k] for k in ("taxpayer_id", "period_start", "period_end", "currency", "amount_unit")}
        sc["period_start"] = start
        source = {"id": sid, "path": name, "sha256": sha(out / name), "context": sc, "metadata_confirmed": True}
        if sid != "ledger":
            source["identity_anchors"] = [{"sheet": "CSV", "cell": "E1", "expected": TAX_ID},
                                           {"sheet": "CSV", "cell": "E2", "expected": start},
                                           {"sheet": "CSV", "cell": "E3", "expected": context["period_end"]}]
        sources.append(source)
    def ref(sid, cell):
        return {"op": "ref", "source": sid, "sheet": "CSV", "cell": cell}
    fields = []
    for fid, label, target, anchor, sourcecell in [
        ("taxpayer", "纳税人识别号", "B2", "A2", "E1"),
        ("name", "纳税人名称", "B3", "A3", "E4"),
        ("period_start", "所属期开始", "D2", "C2", "E2"),
        ("period_end", "所属期结束", "D3", "C3", "E3"),
    ]:
        fields.append({"id": fid, "label": label, "kind": "text", "required": True,
                       "target": {"sheet": "申报底稿", "cell": target, "anchor": {"cell": anchor, "text": label}},
                       "expression": ref("current", sourcecell)})
    labels = {6: "已确认销售额", 7: "已确认应纳税额", 8: "已确认已缴税额", 9: "已确认应补税额"}
    for basis, col, sid in [("current", "B", "current"), ("ytd", "C", "ytd")]:
        for row in range(6, 10):
            fid = f"{sid}_{row}"
            expression = ref(sid, f"B{row - 4}")
            if sid == "current" and row == 6:
                expression = {"op": "table_sum", "source": "ledger", "sheet": "CSV", "header_row": 1,
                              "amount_column": "D", "date_column": "A", "taxpayer_column": "B", "unique_by": ["C"],
                              "headers": {"A": "业务所属日期", "B": "纳税人识别号", "C": "业务行号", "D": "已确认不含税销售额"}}
            fields.append({"id": fid, "label": labels[row], "kind": "money", "basis": basis, "required": True,
                           "target": {"sheet": "申报底稿", "cell": f"{col}{row}", "anchor": {"cell": f"A{row}", "text": labels[row]}},
                           "expression": expression})
    checks = [{"id": "current-revenue", "left": {"op": "field", "id": "current_6"}, "right": ref("current", "B2"), "tolerance": "0.00"}]
    for sid in ("current", "ytd"):
        checks.append({"id": f"{sid}-due", "basis": "ytd" if sid == "ytd" else "current",
                       "left": {"op": "sub", "args": [{"op": "field", "id": f"{sid}_7"}, {"op": "field", "id": f"{sid}_8"}]},
                       "right": {"op": "field", "id": f"{sid}_9"}, "tolerance": "0.00"})
    job = {"schema_version": 1, "context": context, "confirmations": {"context_confirmed": True, "materials_complete": True},
           "sources": sources, "template": {"path": template.name, "sha256": sha(template),
               "context_bindings": {"taxpayer_id": {"sheet":"申报底稿","cell":"B2"},
                                    "period_start": {"sheet":"申报底稿","cell":"D2"},
                                    "period_end": {"sheet":"申报底稿","cell":"D3"}}}, "fields": fields,
           "mandatory_targets": [x["target"] | {} for x in fields], "checks": checks}
    (out / "job.json").write_text(json.dumps(job, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return out / "job.json"


def main():
    parser = argparse.ArgumentParser(description="生成合成数据并完成一次填写演示，不处理真实税务资料")
    parser.add_argument("--out", required=True, type=Path)
    args = parser.parse_args()
    job = generate(args.out.resolve())
    return subprocess.call([sys.executable, str(ROOT / "scripts" / "tax_fill.py"), "run", str(job), "--out", str(job.parent / "result")])


if __name__ == "__main__":
    sys.exit(main())
