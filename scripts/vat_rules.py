#!/usr/bin/env python3
"""Evidence-bound 2026 legacy small-scale VAT calculation; never files a return.

All amounts are Decimal CNY yuan. This module supports ordinary domestic
3% / reduced 1% transactions only. Its output is a draft and remains BLOCKED
when required evidence, cumulative figures, surcharge facts or relief codes
are missing. Pilot forms are deliberately a different, unsupported adapter.
"""
from __future__ import annotations

import argparse
import calendar
import json
import posixpath
import re
import sys
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
from pathlib import Path

VERSION = "0.1.0"
STATUS = "DRAFT_REVIEW_REQUIRED"
QUANTUM = Decimal("0.01")
MAX_MONEY = Decimal("9999999999999.99")
ROUNDING_ALGORITHM = "round_line15_3pct_minus_round_line16_2pct_v1"
CATEGORIES = ("goods_labor", "services")
NUMERIC = re.compile(r"^[+-]?[0-9]+(?:\.[0-9]+)?$")
RULE_SOURCES = [
    {"id": "vat_law", "url": "https://fgk.chinatax.gov.cn/zcfgk/c100009/c5237365/content.html", "verified_on": "2026-10-26"},
    {"id": "2026_10", "url": "https://fgk.chinatax.gov.cn/zcfgk/c102416/c5247434/content.html", "verified_on": "2026-10-26"},
    {"id": "2026_4", "url": "https://shanghai.chinatax.gov.cn/zcfw/zcfgk/zzs/202602/t479286.html", "verified_on": "2026-10-26"},
    {"id": "legacy_2026_instructions", "url": "https://shanghai.chinatax.gov.cn/tax/bsfw/xzzx/bgxz/sbzsl/202107/t459066.html", "verified_on": "2026-10-26"},
    {"id": "legacy_relief_instructions", "url": "https://shanghai.chinatax.gov.cn/bsfw/xzzx/bgxz/sbzsl/202502/t475112.html", "verified_on": "2026-10-26"},
    {"id": "half_reduction", "url": "https://shanghai.chinatax.gov.cn/zcfw/zcfgk/grsds/202308/t468162.html", "verified_on": "2026-10-26"},
    {"id": "education_exemption", "url": "https://tianjin.chinatax.gov.cn/11200000000/0300/030005/20260330152833612.shtml", "verified_on": "2026-10-26"},
    {"id": "urban_rates", "url": "https://shanghai.chinatax.gov.cn/tax/zcfw/zcfgk/ykgf/202008/t454992.html", "verified_on": "2026-10-26"},
    {"id": "standard_fee_rates", "url": "https://jiangsu.chinatax.gov.cn/art/2021/9/10/art_16876_1324.html", "verified_on": "2026-10-26"},
    {"id": "official_oracles", "url": "https://fgk.chinatax.gov.cn/zcfgk/c100015/c5247510/content.html", "verified_on": "2026-10-26"},
]
URLS = {item["id"]: item["url"] for item in RULE_SOURCES}


class InputError(ValueError):
    pass


def money(value, label, *, allow_negative=False):
    if not isinstance(value, str) or len(value) > 64 or not NUMERIC.fullmatch(value):
        raise InputError(f"{label} 必须是明确的十进制字符串金额，不能缺失、布尔、科学计数或含单位")
    try:
        number = Decimal(value)
    except InvalidOperation:
        raise InputError(f"{label} 数值无效")
    if not number.is_finite() or abs(number) > MAX_MONEY or (number < 0 and not allow_negative):
        raise InputError(f"{label} 非有限、超限或含未支持的负数调整")
    return number


def rounded(value):
    return value.quantize(QUANTUM, rounding=ROUND_HALF_UP)


def formatted(value):
    return format(rounded(value), ".2f")


def iso(value, label):
    if not isinstance(value, str) or not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", value):
        raise InputError(f"{label} 必须是 ISO 日期 YYYY-MM-DD")
    try:
        return date.fromisoformat(value)
    except ValueError:
        raise InputError(f"{label} 日期无效")


def has_evidence(item):
    return isinstance(item, dict) and isinstance(item.get("evidence"), str) and bool(item["evidence"].strip())


def source_identity(source):
    """Identify the amount's physical origin, never all rows of one document.

    A supplied hash is an identity token here; the workflow owns actual file
    hash verification. Missing row/cell metadata is explicitly unverified.
    """
    raw_path = source.get("path") or source.get("file")
    path = posixpath.normpath(raw_path.replace("\\", "/")) if isinstance(raw_path, str) and raw_path.strip() else None
    supplied_hash = source.get("sha256")
    stable_hash = supplied_hash.lower() if isinstance(supplied_hash, str) and re.fullmatch(r"[A-Fa-f0-9]{64}", supplied_hash) else None
    document = ("provided_sha256", stable_hash) if stable_hash else ("provided_path", path) if path else None
    sheet = source.get("sheet") if isinstance(source.get("sheet"), str) and source["sheet"].strip() else None
    candidates = []
    cells = source.get("cells", {})
    for candidate in ((cells.get("net_amount") if isinstance(cells, dict) else None), source.get("amount_cell"), source.get("cell")):
        if isinstance(candidate, str) and candidate.strip():
            raw = candidate.strip()
            if "!" in raw:
                cell_sheet, raw = raw.rsplit("!", 1)
                cell_sheet = cell_sheet.strip("'").replace("''", "'")
                if sheet is not None and sheet != cell_sheet:
                    raise InputError("来源金额单元格中的工作表与 source.sheet 冲突")
                sheet = cell_sheet
            raw = raw.replace("$", "").upper()
            if re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]{0,6}", raw):
                candidates.append(raw)
    if len(set(candidates)) > 1:
        raise InputError("source.cells.net_amount/amount_cell/cell 指向不同金额格，来源有冲突")
    cell = candidates[0] if candidates else None
    # CSV/TSV have one implicit sheet. Workbook/PDF locations without a sheet
    # or explicit page remain ambiguous and must not collapse unrelated rows.
    if sheet is None and path and Path(path).suffix.lower() in (".csv", ".tsv"):
        sheet = "CSV"
    if sheet is None and source.get("page") is not None:
        sheet = f"page:{source['page']}"
    row = source.get("row")
    if isinstance(row, str) and re.fullmatch(r"[1-9][0-9]{0,6}", row):
        row = int(row)
    if type(row) is not int or row < 1:
        row = None
    locator = ("amount_cell", cell) if cell else ("row", row) if row else None
    if document and sheet is not None and locator:
        key = document + (sheet,) + locator
        return key, {"kind": f"{document[0]}+sheet+{locator[0]}", "physical_key_available": True,
                     "document_identity": document[1], "provided_path": path, "sheet": sheet, locator[0]: locator[1],
                     "hash_verified_by_this_module": False}
    return None, {"kind": "document_reference_only" if document else "unstructured_reference",
                  "physical_key_available": False, "hash_verified_by_this_module": False,
                  "note": "缺少可唯一定位金额的文件及工作表/行格信息，仅凭交易id查重，不能声称已排除物理来源重复"}


def business_identity(item):
    """Only use an explicitly unique business key; never invent an invoice row."""
    source = item["source"]
    explicit = item.get("business_key") or source.get("business_key")
    if isinstance(explicit, str) and explicit.strip():
        return ("declared_business_key", explicit.strip()), "declared_business_key"
    number = item.get("invoice_number", item.get("invoice_no", source.get("invoice_number", source.get("invoice_no"))))
    code = item.get("invoice_code", source.get("invoice_code", ""))
    line = item.get("invoice_line_number", item.get("invoice_line", item.get("line_number",
                    source.get("invoice_line_number", source.get("invoice_line", source.get("line_number"))))))
    if isinstance(number, str) and number.strip():
        if isinstance(line, (str, int)) and not isinstance(line, bool) and str(line).strip():
            return ("invoice_line", str(code).strip(), number.strip(), str(line).strip()), "invoice_number_and_explicit_line"
        if item.get("invoice_record_scope", source.get("invoice_record_scope")) == "whole_invoice":
            return ("whole_invoice", str(code).strip(), number.strip()), "explicit_whole_invoice"
        return None, "invoice_number_only_not_assumed_unique_for_multiline_invoice"
    return None, "business_key_not_provided"


class Calculation:
    def __init__(self, transactions, profile):
        self.transactions, self.profile = transactions, profile
        self.current = {str(row): {category: Decimal(0) for category in CATEGORIES} for row in range(1, 23)}
        self.result = {
            "schema_version": 1, "engine_version": VERSION, "status": STATUS,
            "outcome": "BLOCKED", "lines": {}, "issues": [], "unresolved_fields": [],
            "transaction_trace": [], "line_trace": {}, "surcharge_details": {},
            "relief_details": {}, "rule_sources": [dict(item) for item in RULE_SOURCES],
            "rounding_comparisons": [],
            "human_review_required": True, "filing_performed": False,
            "rounding_policy": "Decimal CNY; each monetary return line rounded HALF_UP to cents; subtraction uses rounded return lines",
        }

    def issue(self, code, message, field=None, transaction_id=None):
        item = {"code": code, "message": message}
        if field is not None:
            item["field"] = field
            if field not in self.result["unresolved_fields"]:
                self.result["unresolved_fields"].append(field)
        if transaction_id is not None:
            item["transaction_id"] = transaction_id
        self.result["issues"].append(item)

    def evidence_value(self, container, key, label, *, required_true=False, numeric=False):
        item = container.get(key) if isinstance(container, dict) else None
        if not has_evidence(item) or "value" not in item:
            self.issue("EVIDENCE_REQUIRED", f"{label} 缺少明确值及来源证据，不能默认零或自动确认", label)
            return None
        value = item["value"]
        if value is None or (isinstance(value, str) and not value.strip()):
            self.issue("EVIDENCE_VALUE_MISSING", f"{label} 有来源说明但没有明确值，不能视为完成", label)
            return None
        if required_true and value is not True:
            self.issue("CONFIRMATION_REQUIRED", f"{label} 必须有明确的 true 确认及证据", label)
            return None
        if numeric:
            try:
                return money(value, label)
            except InputError as exc:
                self.issue("AMOUNT_INVALID", str(exc), label)
                return None
        return value

    def validate_profile(self):
        p = self.profile
        if not isinstance(p, dict):
            raise InputError("profile 必须为对象")
        if p.get("form_version") != "legacy_small_2026":
            self.issue("FORM_VERSION_UNSUPPORTED", "仅支持 legacy_small_2026；试点新表不能套用旧表行次和减免算法", "profile.form_version")
            return False
        if not isinstance(p.get("taxpayer_id"), str) or not p["taxpayer_id"].strip() or p["taxpayer_id"] != p["taxpayer_id"].strip():
            raise InputError("taxpayer_id 必须为明确文本税号")
        if p.get("taxpayer_type") not in ("organization", "individual"):
            raise InputError("taxpayer_type 仅支持 organization 或已登记的 individual 个体工商户；不支持自然人按次")
        self.start = iso(p.get("period_start"), "period_start")
        self.end = iso(p.get("period_end"), "period_end")
        if self.start.year != 2026 or self.end.year != 2026 or self.start > self.end:
            raise InputError("本适配器只核实了 2026 年旧表，起止必须在 2026 年且顺序正确")
        if self.start.day != 1 or self.end.day != calendar.monthrange(self.end.year, self.end.month)[1]:
            raise InputError("当前只支持完整自然月或自然季度")
        kind = p.get("period_kind")
        if kind == "month":
            if self.start.month != self.end.month:
                raise InputError("month 必须为同一完整自然月")
            self.threshold = Decimal("100000")
        elif kind == "quarter":
            if self.start.month not in (1, 4, 7, 10) or self.end.month != self.start.month + 2:
                raise InputError("quarter 必须为一个完整自然季度")
            self.threshold = Decimal("300000")
        else:
            raise InputError("period_kind 仅支持 month 或 quarter；VAT 年报不属于此分支")
        if self.end > date.today():
            raise InputError("所属期尚未结束，不能确认完整期间申报数据")
        confirmations = p.get("confirmations", {})
        for key in ("small_scale", "normal_filing", "ordinary_domestic_only", "income_complete_including_unbilled", "tax_dates_and_rates_confirmed", "other_vat_reliefs_absent"):
            self.evidence_value(confirmations, key, f"confirmations.{key}", required_true=True)
        self.result["profile_context"] = {key: p.get(key) for key in ("taxpayer_id", "taxpayer_type", "period_start", "period_end", "period_kind", "form_version")}
        self.result["confirmation_evidence"] = confirmations
        return not self.result["issues"]

    def normalize_transactions(self):
        if not isinstance(self.transactions, list):
            raise InputError("transactions 必须为列表")
        if len(self.transactions) > 50000:
            raise InputError("transactions 超过 50000 行，请按本期拆分且保留完整性依据")
        seen, source_seen, source_alias_seen, business_seen, normalized = set(), {}, {}, {}, []
        for index, item in enumerate(self.transactions):
            tid = item.get("id") if isinstance(item, dict) else None
            try:
                if not isinstance(item, dict):
                    raise InputError("交易必须为对象")
                if not isinstance(tid, str) or not tid.strip() or tid != tid.strip() or tid in seen:
                    raise InputError("交易 id 缺失或重复；不能静默去重")
                seen.add(tid)
                if item.get("taxpayer_id") != self.profile["taxpayer_id"]:
                    raise InputError("交易主体与 profile 不一致")
                dt = iso(item.get("tax_date"), "tax_date")
                if not self.start <= dt <= self.end:
                    raise InputError("交易税务所属日不在本期；不自动挪期或净额抵销")
                category, invoice_kind, status = item.get("category"), item.get("invoice_kind"), item.get("status")
                if category not in CATEGORIES:
                    raise InputError("category 必须明确为 goods_labor 或 services")
                if invoice_kind not in ("special", "ordinary", "unbilled"):
                    raise InputError("invoice_kind 仅支持 special、ordinary、unbilled")
                if status not in ("valid", "void", "red"):
                    raise InputError("status 必须明确为 valid、void 或 red")
                if not isinstance(item.get("source"), dict) or not item["source"]:
                    raise InputError("交易必须有可回查的 source 记录")
                physical_key, physical_evidence = source_identity(item["source"])
                business_key, business_evidence = business_identity(item)
                if status == "valid":
                    previous_source = source_seen.get(physical_key) if physical_key is not None else None
                    path_alias = None
                    if physical_key is not None and physical_evidence.get("provided_path"):
                        path_alias = ("provided_path", physical_evidence["provided_path"]) + physical_key[2:]
                        aliases = source_alias_seen.get(path_alias, {})
                        # A row with no supplied hash must still match the same
                        # path/location seen with a hash. Distinct valid hashes
                        # remain distinct documents even if basenames coincide.
                        previous_source = previous_source or aliases.get("any_id" if physical_key[0] == "provided_path" else "unhashed_id")
                    if previous_source is not None:
                        self.issue("SOURCE_DUPLICATE", f"同一原始金额来源重复映射（首次交易 {previous_source}）；更换id或备注不能重复计税", f"transactions[{index}].source", tid)
                        continue
                    if business_key is not None and business_key in business_seen:
                        self.issue("BUSINESS_DUPLICATE", f"明确业务主键重复（首次交易 {business_seen[business_key]}），不能重复计税", f"transactions[{index}].business_key", tid)
                        continue
                    if physical_key is not None:
                        source_seen[physical_key] = tid
                    if path_alias is not None:
                        alias = source_alias_seen.setdefault(path_alias, {"any_id": tid})
                        if physical_key[0] == "provided_path":
                            alias["unhashed_id"] = tid
                    if business_key is not None:
                        business_seen[business_key] = tid
                if status == "red":
                    self.result["transaction_trace"].append({"id": tid, "included": False, "reason": "红字／冲销业务未适配，不进行跨期净额抵销", "source": item["source"]})
                    raise InputError("红字、退回及跨期冲销尚未适配，需要独立税务处理")
                net = money(item.get("net_amount"), "net_amount")
                if item.get("applicable_rate") not in ("0.01", "0.03"):
                    raise InputError("仅支持已确认的 1% 或 3% 普通内销业务；不得猜其他征收率")
                trace = {
                    "id": tid, "taxpayer_id": item["taxpayer_id"], "tax_date": dt.isoformat(),
                    "category": category, "invoice_kind": invoice_kind, "status": status,
                    "net_amount": str(net), "applicable_rate": item["applicable_rate"],
                    "source": item["source"], "included": status == "valid",
                    "source_identity_evidence": physical_evidence, "business_identity_evidence": business_evidence,
                    "reason": "有效交易，待按全户门槛分类" if status == "valid" else "已明确作废，排除且保留追溯",
                    "rule_sources": [URLS["vat_law"], URLS["2026_4"]],
                }
                self.result["transaction_trace"].append(trace)
                if status == "valid":
                    normalized.append({**item, "net": net, "trace": trace})
            except InputError as exc:
                self.issue("TRANSACTION_INVALID", str(exc), f"transactions[{index}]", tid)
        return normalized

    def calculate_current(self, transactions):
        self.total_sales = sum((item["net"] for item in transactions), Decimal(0))
        if self.total_sales > MAX_MONEY:
            raise InputError("全户销售额超出金额可靠精度范围")
        below = self.total_sales < self.threshold
        self.result["threshold_test"] = {
            "sales_all_categories_yuan": format(self.total_sales, "f"), "threshold_yuan": formatted(self.threshold),
            "below_threshold": below, "comparison": "strict_lt", "includes_special_ordinary_unbilled": True,
            "rule_sources": [URLS["vat_law"], URLS["2026_10"]],
        }
        taxable_one = {category: Decimal(0) for category in CATEGORIES}
        membership = {str(row): {category: [] for category in CATEGORIES} for row in range(1, 21)}
        exemption_row = "10" if self.profile["taxpayer_type"] == "organization" else "11"
        for item in transactions:
            category, net = item["category"], item["net"]
            taxable = not below or item["invoice_kind"] == "special"
            if taxable:
                self.current["1"][category] += net
                membership["1"][category].append(item["id"])
                if item["invoice_kind"] == "special":
                    self.current["2"][category] += net
                    membership["2"][category].append(item["id"])
                elif item["invoice_kind"] == "ordinary":
                    self.current["3"][category] += net
                    membership["3"][category].append(item["id"])
                if item["applicable_rate"] == "0.01":
                    taxable_one[category] += net
                    membership["16"][category].append(item["id"])
                item["trace"].update({"treatment": "taxable_1pct" if item["applicable_rate"] == "0.01" else "taxable_3pct", "legacy_lines": ["1"] + (["2"] if item["invoice_kind"] == "special" else ["3"] if item["invoice_kind"] == "ordinary" else []), "reason": "达到全户起征点，全额计税" if not below else "未达起征点但专票交易放弃免税"})
            else:
                self.current[exemption_row][category] += net
                membership[exemption_row][category].append(item["id"])
                item["trace"].update({"treatment": "threshold_exempt", "legacy_lines": ["9", exemption_row], "reason": "全户未达到起征点，非专票普通业务享受门槛免税"})
        for category in CATEGORIES:
            for row in range(1, 15):
                self.current[str(row)][category] = rounded(self.current[str(row)][category])
            self.current["9"][category] = self.current["10"][category] + self.current["11"][category]
            self.current["15"][category] = rounded(self.current["1"][category] * Decimal("0.03"))
            self.current["16"][category] = rounded(taxable_one[category] * Decimal("0.02"))
            self.current["17"][category] = rounded(self.current["9"][category] * Decimal("0.03"))
            self.current["18"][category] = rounded(self.current["10"][category] * Decimal("0.03"))
            self.current["19"][category] = rounded(self.current["11"][category] * Decimal("0.03"))
            self.current["20"][category] = self.current["15"][category] - self.current["16"][category]
            direct = rounded(taxable_one[category] * Decimal("0.01") + (self.current["1"][category] - taxable_one[category]) * Decimal("0.03"))
            difference = self.current["20"][category] - direct
            self.result["rounding_comparisons"].append({
                "category": category, "line15_minus_line16": formatted(self.current["20"][category]),
                "direct_reduced_rate": formatted(direct), "difference": formatted(difference),
                "note": "辅助比较两种舍入路径；存在差异不代表已核实税局采用其中一种",
            })
            if difference:
                confirmation = self.profile.get("rounding_policy_confirmation")
                region = self.profile.get("region")
                binding_ok = has_evidence(confirmation) and confirmation.get("value") == "legacy_line_rounding" and isinstance(region, str) and bool(region.strip()) and confirmation.get("region") == region and confirmation.get("form_version") == self.profile["form_version"] and confirmation.get("algorithm") == ROUNDING_ALGORITHM
                if not binding_ok:
                    self.issue("ROUNDING_REVIEW_REQUIRED", "旧表3%应纳减2%减征与直接实际征收率计算存在分位差异；需按地区、表版和算法绑定实际税表舍入确认，不自动调平", "rounding_policy_confirmation")
                else:
                    self.result["rounding_policy_binding"] = dict(confirmation)
        formulas = {
            "1": "sum(taxable net sales)", "2": "sum(taxable special-invoice net sales)",
            "3": "sum(taxable ordinary-invoice net sales); unbilled remains only in line 1",
            "9": "10+11+12", "10": "organization threshold-exempt net sales", "11": "individual threshold-exempt net sales",
            "15": "round(1*3%)", "16": "round(taxable sales confirmed at 1% * 2%)",
            "17": "round(9*3%)", "18": "round(10*3%)", "19": "round(11*3%)", "20": "15-16",
        }
        for row in range(1, 21):
            rid = str(row)
            self.result["lines"][rid] = {category: formatted(self.current[rid][category]) for category in CATEGORIES}
            self.result["line_trace"][rid] = {
                "formula": formulas.get(rid, "zero only under confirmed ordinary domestic scope and absence of other reliefs"),
                "transaction_ids": membership[rid],
                "rule_sources": [URLS["legacy_2026_instructions"], URLS["2026_10"]],
            }
            if rid not in formulas:
                self.result["line_trace"][rid]["scope_evidence"] = self.profile["confirmations"]["ordinary_domestic_only"]
        prepaid = self.profile.get("prepaid_vat", {})
        for rid in ("21", "22"):
            self.result["lines"][rid] = {}
        for category in CATEGORIES:
            amount = self.evidence_value(prepaid, category, f"prepaid_vat.{category}", numeric=True)
            if amount is None:
                self.result["unresolved_fields"].append(f"lines.22.{category}")
                continue
            self.current["21"][category] = rounded(amount)
            self.current["22"][category] = self.current["20"][category] - rounded(amount)
            self.result["lines"]["21"][category] = formatted(amount)
            self.result["lines"]["22"][category] = formatted(self.current["22"][category])
        self.result["line_trace"]["21"] = {"formula": "confirmed actual prepaid VAT; excludes audit-assessed tax", "evidence": prepaid, "rule_sources": [URLS["legacy_2026_instructions"]]}
        self.result["line_trace"]["22"] = {"formula": "20-21; no cumulative cells on lines 21/22", "rule_sources": [URLS["legacy_2026_instructions"]]}

    def calculate_ytd(self):
        p = self.profile
        zero_flag = p.get("first_period_zero_ytd")
        confirmed_zero = has_evidence(zero_flag) and zero_flag.get("value") is True
        prior = p.get("prior_ytd_lines")
        if prior is None and confirmed_zero:
            prior = {}
        if not isinstance(prior, dict):
            self.issue("PRIOR_YTD_REQUIRED", "累计列需要上一期已申报的逐栏累计；首次零累计须明确确认并有证据", "prior_ytd_lines")
            return
        if not confirmed_zero and (not isinstance(p.get("prior_ytd_evidence"), str) or not p["prior_ytd_evidence"].strip()):
            self.issue("PRIOR_YTD_EVIDENCE", "上一期累计缺少已申报来源证据", "prior_ytd_evidence")
            return
        parsed = {}
        before = len(self.result["issues"])
        for row in range(1, 21):
            rid = str(row)
            raw = prior.get(rid, {} if confirmed_zero else None)
            if not isinstance(raw, dict):
                self.issue("PRIOR_YTD_MISSING", f"上期累计缺少第 {rid} 栏，不能补零", f"prior_ytd_lines.{rid}")
                continue
            parsed[rid] = {}
            for category in CATEGORIES:
                value = raw.get(category, "0.00" if confirmed_zero else None)
                try:
                    amount = money(value, f"prior_ytd_lines.{rid}.{category}")
                    if amount != rounded(amount):
                        raise InputError("已申报累计应精确到分，不能静默舍入历史值")
                    if confirmed_zero and amount != 0:
                        raise InputError("首次零累计确认与已提供的非零上期累计冲突")
                    parsed[rid][category] = amount
                except InputError as exc:
                    self.issue("PRIOR_YTD_INVALID", str(exc), f"prior_ytd_lines.{rid}.{category}")
        if len(self.result["issues"]) != before:
            return
        for category in CATEGORIES:
            if parsed["9"][category] != parsed["10"][category] + parsed["11"][category] + parsed["12"][category]:
                self.issue("PRIOR_YTD_RECONCILIATION", "上期累计第9栏不等于10+11+12", f"prior_ytd_lines.9.{category}")
            if parsed["20"][category] != parsed["15"][category] - parsed["16"][category]:
                self.issue("PRIOR_YTD_RECONCILIATION", "上期累计第20栏不等于15-16", f"prior_ytd_lines.20.{category}")
            if parsed["1"][category] < parsed["2"][category] + parsed["3"][category]:
                self.issue("PRIOR_YTD_RECONCILIATION", "上期累计第1栏小于其中专票及其他票销售额之和", f"prior_ytd_lines.1.{category}")
            for parent, children in (("4", ("5", "6")), ("7", ("8",)), ("13", ("14",)), ("17", ("18", "19"))):
                if parsed[parent][category] < sum((parsed[child][category] for child in children), Decimal(0)):
                    self.issue("PRIOR_YTD_RECONCILIATION", f"上期累计第{parent}栏小于其子项{'＋'.join(children)}栏合计，不能作为已核对累计继续结转", f"prior_ytd_lines.{parent}.{category}")
            if parsed["15"][category] < parsed["16"][category]:
                self.issue("PRIOR_YTD_RECONCILIATION", "上期累计第16栏减征额超过第15栏应纳税额", f"prior_ytd_lines.16.{category}")
        if len(self.result["issues"]) != before:
            return
        for row in range(1, 21):
            rid = str(row)
            for category in CATEGORIES:
                amount = parsed[rid][category] + self.current[rid][category]
                if amount > MAX_MONEY:
                    self.issue("YTD_RANGE", "累计金额超出可靠精度范围", f"lines.{rid}.ytd_{category}")
                    continue
                self.result["lines"][rid][f"ytd_{category}"] = formatted(amount)
            self.result["line_trace"][rid]["ytd_formula"] = "previous filed same-line YTD + current same-line; never reapply this period threshold to annual sales"
            self.result["line_trace"][rid]["prior_evidence"] = zero_flag.get("evidence") if confirmed_zero else p["prior_ytd_evidence"]

    def calculate_relief(self):
        reduction = sum((self.current["16"][category] for category in CATEGORIES), Decimal(0))
        if reduction == 0:
            self.result["relief_details"] = {"required": False, "reason": "当前已确认分支的主表12与16均无数据，不需旧版减免明细", "rule_sources": [URLS["legacy_relief_instructions"]]}
            return
        code = self.evidence_value(self.profile, "vat_reduction_code", "vat_reduction_code")
        if code is not None and (not isinstance(code, str) or not code.strip()):
            self.issue("RELIEF_CODE_INVALID", "减税性质代码须为有来源的非空文本", "vat_reduction_code")
            code = None
        self.result["relief_details"] = {
            "required": True, "reduction_1pct": {
                "code": code, "opening": "0.00", "arising": formatted(reduction),
                "available": formatted(reduction), "used": formatted(reduction), "closing": "0.00",
                "columns": {"1": "0.00", "2": formatted(reduction), "3": formatted(reduction), "4": formatted(reduction), "5": "0.00"},
                "code_evidence": self.profile.get("vat_reduction_code", {}).get("evidence"),
            },
            "threshold_exemption_section_required": False,
            "rule_sources": [URLS["legacy_2026_instructions"], URLS["legacy_relief_instructions"]],
        }

    def calculate_surcharges(self):
        for rid in ("23", "24", "25"):
            self.result["lines"][rid] = {}
            self.result["line_trace"][rid] = {"formula": "附列资料（二）对应税费第9栏；整户本期税费，不伪拆业务列或填累计灰格", "rule_sources": [URLS["legacy_2026_instructions"], URLS["half_reduction"], URLS["education_exemption"]]}
        if any(category not in self.result["lines"]["22"] for category in CATEGORIES):
            self.issue("SURCHARGE_BASE_UNRESOLVED", "预缴未核实，附加计税依据尚不能确定", "surcharges")
            return
        if any(self.current["22"][category] < 0 for category in CATEGORIES):
            self.issue("SURCHARGE_REFUND_UNSUPPORTED", "出现预缴超过本期应纳的退税，相关附加处理未核实，保留主表并停止附加计算", "surcharges")
            return
        s = self.profile.get("surcharges", {})
        rates = {}
        for key in ("urban_rate", "education_rate", "local_education_rate"):
            number = self.evidence_value(s, key, f"surcharges.{key}", numeric=True)
            supported = {"urban_rate": {Decimal("0.01"), Decimal("0.05"), Decimal("0.07")}, "education_rate": {Decimal("0.03")}, "local_education_rate": {Decimal("0.02")}}[key]
            if number is not None and number not in supported:
                self.issue("SURCHARGE_RATE_INVALID", "当前只适配城建1%/5%/7%、教育3%、地方教育2%的普通分支；其他费率须另核规则，不能把7%写成7", f"surcharges.{key}")
                number = None
            rates[key] = number
        self.evidence_value(s, "other_reliefs_absent", "surcharges.other_reliefs_absent", required_true=True)
        # Derive eligibility from confirmed factual scope and the official rule.
        # The accountant is not asked to calculate the result in advance.
        half = True
        education_exempt = self.total_sales <= self.threshold
        if isinstance(s, dict) and "half_reduction_confirmed" in s:
            half = self.evidence_value(s, "half_reduction_confirmed", "surcharges.half_reduction_confirmed", required_true=True)
        if isinstance(s, dict) and "education_exemption_policy_applies" in s:
            if self.evidence_value(s, "education_exemption_policy_applies", "surcharges.education_exemption_policy_applies", required_true=True) is None:
                education_exempt = None
        if isinstance(s, dict) and "education_exemption_confirmed" in s:
            self.issue("DERIVED_RESULT_AS_INPUT", "不要预先提供当期教育免征计算结果；由完整销售额和<=门槛政策自动判断。额外资格事实使用 education_exemption_policy_applies", "surcharges.education_exemption_confirmed")
            education_exempt = None
        paid = s.get("paid", {}) if isinstance(s, dict) else {}
        tax_base = sum((self.current["22"][category] for category in CATEGORIES), Decimal(0))
        for name, rate_key, rid in (("urban", "urban_rate", "23"), ("education", "education_rate", "24"), ("local_education", "local_education_rate", "25")):
            paid_key = "local_education" if name == "local_education" and isinstance(paid, dict) and "local_education" in paid else "local" if name == "local_education" else name
            already_paid = self.evidence_value(paid, paid_key, f"surcharges.paid.{paid_key}", numeric=True)
            rate = rates[rate_key]
            if rate is None or already_paid is None or half is None or not (isinstance(s, dict) and has_evidence(s.get("other_reliefs_absent")) and s["other_reliefs_absent"].get("value") is True):
                continue
            if name != "urban" and education_exempt is None:
                continue
            gross = rounded(tax_base * rate)
            other = gross if name != "urban" and education_exempt else Decimal(0)
            code = ""
            if other != 0:
                code_key = "education_exemption_code" if name == "education" else "local_education_exemption_code"
                code = self.evidence_value(s, code_key, f"surcharges.{code_key}")
                if code is not None and (not isinstance(code, str) or not code.strip()):
                    self.issue("RELIEF_CODE_INVALID", "附加减免性质代码须为有来源的非空文本", f"surcharges.{code_key}")
                    code = None
            half_amount = rounded((gross - other) * Decimal("0.50"))
            net_due = gross - other - half_amount - rounded(already_paid)
            if net_due < 0:
                self.issue("SURCHARGE_PAID_REFUND_UNSUPPORTED", "本期已缴超过附加应纳，退抵处理未核实，不输出该行应补退额", f"lines.{rid}.total")
                continue
            self.result["lines"][rid]["total"] = formatted(net_due)
            self.result["surcharge_details"][name] = {
                "columns": {"1": formatted(tax_base), "2": str(rate), "3": formatted(gross), "4": code,
                            "5": formatted(other), "6": "0.50", "7": formatted(half_amount), "8": formatted(already_paid), "9": formatted(net_due)},
                "formula": "3=round(1*2); 7=round((3-5)*50%); 9=3-5-7-8",
                "column_units": {"1": "CNY_yuan", "2": "fraction", "3": "CNY_yuan", "5": "CNY_yuan", "6": "fraction", "7": "CNY_yuan", "8": "CNY_yuan", "9": "CNY_yuan"},
                "education_exemption_sales_test": "sales <= monthly 100000 / quarterly 300000",
                "education_exemption_derived": self.total_sales <= self.threshold,
                "half_reduction_derived": "0.50 from confirmed small-scale status and 2023-12 policy; not an amount supplied by the accountant",
                "rule_sources": [URLS["legacy_2026_instructions"], URLS["half_reduction"], URLS["education_exemption"], URLS["urban_rates"], URLS["standard_fee_rates"]],
            }

    def run(self):
        try:
            if not self.validate_profile():
                return self.result
            transactions = self.normalize_transactions()
            if self.result["issues"]:
                return self.result
            self.calculate_current(transactions)
            self.calculate_ytd()
            self.calculate_relief()
            self.calculate_surcharges()
            if not self.result["issues"]:
                self.result["outcome"] = "GENERATED_FOR_REVIEW"
        except InputError as exc:
            self.issue("INPUT_INVALID", str(exc), "profile_or_transactions")
        except (InvalidOperation, ArithmeticError) as exc:
            self.issue("DECIMAL_ERROR", f"金额计算异常，未生成完整套表: {type(exc).__name__}", "calculation")
        except (TypeError, AttributeError, KeyError) as exc:
            self.issue("INPUT_STRUCTURE", f"输入结构不符合当前契约: {type(exc).__name__}", "profile_or_transactions")
        self.result["unresolved_fields"] = list(dict.fromkeys(self.result["unresolved_fields"]))
        return self.result


def calculate(transactions, profile):
    """Return known draft lines plus explicit blockers; never infer missing facts."""
    with localcontext() as context:
        context.prec = 50
        return Calculation(transactions, profile).run()


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("case", type=Path, help="JSON with transactions and profile")
    parser.add_argument("--out", type=Path, help="new JSON path; existing files are never overwritten")
    args = parser.parse_args(argv)
    try:
        if args.out and args.out.exists():
            raise InputError("--out 文件已存在，拒绝覆盖")
        if args.case.stat().st_size > 25 * 1024 * 1024:
            raise InputError("case JSON 超过25 MiB")
        payload = json.loads(args.case.read_text(encoding="utf-8-sig"))
        if not isinstance(payload, dict):
            raise InputError("case JSON 顶层必须为对象")
        result = calculate(payload.get("transactions"), payload.get("profile"))
        output = json.dumps(result, ensure_ascii=False, indent=2, allow_nan=False) + "\n"
        if args.out:
            args.out.parent.mkdir(parents=True, exist_ok=True)
            with args.out.open("x", encoding="utf-8") as handle:
                handle.write(output)
        else:
            sys.stdout.write(output)
        return 0 if result["outcome"] == "GENERATED_FOR_REVIEW" else 2
    except (InputError, OSError, ValueError, TypeError) as exc:
        print(json.dumps({"status": STATUS, "outcome": "BLOCKED", "issues": [{"code": "CLI_INPUT", "message": str(exc)}]}, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
