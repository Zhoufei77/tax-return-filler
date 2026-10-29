#!/usr/bin/env python3
"""Generate and run synthetic raw-input business demos, never customer facts.

The only inputs copied by this generator are frozen, labelled examples in the
package. It does not accept a real customer file or set confirmations for one.
"""
from __future__ import annotations
import argparse
import csv
import json
from pathlib import Path
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
NOTE = "合成业务演示的给定事实，不代表真实企业信息或用户已授权真实申报"


def ev(value):
    return {"value": value, "evidence": NOTE}


def write_payment(path, taxid, rows):
    matrix = [["项目", "金额", "金额单位：元", "纳税人识别号", taxid]]
    matrix.extend([[label, value, "", "", ""] for label, value in rows])
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        csv.writer(handle).writerows(matrix)


def bound_money(filename, row, label):
    return {"source": {"path": filename, "sheet": "CSV", "cell": f"B{row}", "kind": "money", "amount_unit": "元",
                       "anchor": {"cell": f"A{row}", "text": label},
                       "identity_anchors": [{"cell": "E1", "context_key": "taxpayer_id"}]}}


def make_eit(folder, form_id):
    prepaid = form_id == "A200000"
    inputs = folder / "input"
    inputs.mkdir(parents=True)
    shutil.copyfile(ROOT / "assets" / ("income-statement-quarter.xlsx" if prepaid else "income-statement-annual.xlsx"), inputs / "利润表.xlsx")
    meta = json.loads((ROOT / "assets" / f"{form_id}.json").read_text(encoding="utf-8"))
    start, end, kind = ("2026-07-01", "2026-09-30", "quarter") if prepaid else ("2025-01-01", "2025-12-31", "year")
    context = {"taxpayer_id": "DEMO-EIT", "taxpayer_name": "合成演示居民企业", "period_start": start, "period_end": end,
               "period_kind": kind, "region": "合成演示地区", "currency": "CNY", "amount_unit": "元",
               "form_name": meta["form_name"], "form_version": meta["form_version"]}
    label, paid = ("本年累计已预缴所得税额", "1750000.00") if prepaid else ("本年累计预缴所得税额", "3250000.00")
    payment = "上期预缴记录.csv"
    write_payment(inputs / payment, context["taxpayer_id"], [(label, paid)])
    zeros = ["19","20","21","22","23","24","29","31","39"] if prepaid else ["19","20","21","22","23","25","26","27","32","34","35","42","43","44"]
    facts = {"tax_lines": {"30" if prepaid else "37": bound_money(payment, 2, label)},
             "confirmed_zero_tax_lines": zeros, "confirmation_evidence": NOTE,
             "tax_treatment": "standard_25", "tax_treatment_evidence": NOTE,
             "scope": {key: ev(value) for key,value in {"resident_enterprise":True,"collection_method":"actual_profit",
                       "independent_taxpayer":True,"no_foreign_business":True,"no_special_business":True}.items()}}
    return {"schema_version":2,"return_id":form_id,"context":context,"input_dir":"input",
            "income_statement":{"path":"利润表.xlsx","sheet":"利润表","accounting_standard":"asbe","presentation":"software_export",
                                "source_period_start":f"{end[:4]}-01-01","source_period_end":end,"period_evidence":NOTE},
            "facts":facts,"confirmation_evidence":NOTE}


def make_vat(folder):
    inputs = folder / "input"
    inputs.mkdir(parents=True)
    shutil.copyfile(ROOT / "assets" / "vat-income.csv", inputs / "收入台账.csv")
    meta = json.loads((ROOT / "assets" / "VAT_SMALL_2026.json").read_text(encoding="utf-8"))
    context = {"taxpayer_id":"DEMO-VAT","taxpayer_name":"合成演示纺纱企业","period_start":"2026-01-01","period_end":"2026-03-31",
               "period_kind":"quarter","region":"合成演示地区","currency":"CNY","amount_unit":"元",
               "form_name":meta["form_name"],"form_version":"legacy_small_2026"}
    labels=["货物及劳务预缴增值税","服务等预缴增值税","城市维护建设税本期已缴","教育费附加本期已缴","地方教育附加本期已缴"]
    filename="预缴与已缴记录.csv"
    write_payment(inputs / filename, context["taxpayer_id"], [(name,"0.00") for name in labels])
    profile = {"taxpayer_type":"organization",
               "confirmations":{key:ev(True) for key in ("small_scale","normal_filing","ordinary_domestic_only","income_complete_including_unbilled","tax_dates_and_rates_confirmed","other_vat_reliefs_absent")},
               "prepaid_vat":{key:bound_money(filename,row,label) for key,row,label in [("goods_labor",2,labels[0]),("services",3,labels[1])]},
               "first_period_zero_ytd":ev(True),
               "surcharges":{"urban_rate":ev("0.07"),"education_rate":ev("0.03"),"local_education_rate":ev("0.02"),"other_reliefs_absent":ev(True),
                             "paid":{key:bound_money(filename,row,label) for key,row,label in [("urban",4,labels[2]),("education",5,labels[3]),("local",6,labels[4])]}}}
    return {"context":context,"input_dir":"input","confirmation_evidence":NOTE,"profile":profile,
            "ledger":{"path":"收入台账.csv","source_period_start":"2026-01-01","source_period_end":"2026-03-31","period_evidence":NOTE}}


def main(argv=None):
    parser=argparse.ArgumentParser(description="从明确标记的合成原始资料演示税表业务链路")
    parser.add_argument("--out",required=True,type=Path)
    parser.add_argument("--return-id",choices=["A200000","A100000","VAT_SMALL_2026","all"],default="all")
    args=parser.parse_args(argv)
    out=args.out.expanduser().resolve()
    out.mkdir(parents=True,exist_ok=False)
    ids=["A200000","A100000","VAT_SMALL_2026"] if args.return_id=="all" else [args.return_id]
    results=[]
    for form_id in ids:
        folder=out/form_id
        folder.mkdir()
        case=make_vat(folder) if form_id=="VAT_SMALL_2026" else make_eit(folder,form_id)
        task=folder/"case.json"
        task.write_text(json.dumps(case,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
        script=ROOT/"scripts"/("vat_workflow.py" if form_id=="VAT_SMALL_2026" else "tax_workflow.py")
        run=subprocess.run([sys.executable,str(script),"run",str(task),"--out",str(folder/"result")],capture_output=True,text=True)
        results.append({"return_id":form_id,"exit_code":run.returncode,"stdout":run.stdout.strip(),"stderr":run.stderr.strip()})
    (out/"demo-results.json").write_text(json.dumps(results,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    print(json.dumps({"synthetic_only":True,"out":str(out),"results":results},ensure_ascii=False))
    return 0 if all(r["exit_code"]==0 for r in results) else 2


if __name__=="__main__":
    sys.exit(main())
