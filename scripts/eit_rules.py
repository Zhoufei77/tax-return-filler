#!/usr/bin/env python3
"""Evidence-bound Chinese resident enterprise income-tax draft calculations.

Only Python's standard library is used. This module never files a return,
interprets arbitrary expressions, reads a network, or invents missing amounts.
All monetary inputs and outputs are CNY yuan. See references/eit-2025-adapter.md.
"""
from __future__ import annotations

from calendar import monthrange
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
from typing import Any

VERSION = "0.1.0"
RULESET_VERSION = "cn-eit-A200000-2025-17_A100000-2025-1"
STATUS = "DRAFT_REVIEW_REQUIRED"
CENT = Decimal("0.01")
ZERO = Decimal("0")
RATE = Decimal("0.25")
PREPAID_URL = "https://fgk.chinatax.gov.cn/zcfgk/c100012/c5241820/content.html"
ANNUAL_URL = "https://fgk.chinatax.gov.cn/zcfgk/c100012/c5237968/content.html"
PREPAID_FORM_URL = "https://12366.chinatax.gov.cn/bzds/059/059.html"
ANNUAL_FORM_URL = "https://12366.chinatax.gov.cn/bzds/056/056.html"
SMALL_URL = "https://www.chinatax.gov.cn/chinatax/n810356/n3010387/c5211219/content.html"
PREPAID_ORACLE_URL = "https://fgk.chinatax.gov.cn/zcfgk/c100015/c5241823/content.html"

# Order and signs follow the 2025 A200000 and 2025 A100000 main forms.
FINANCIAL_FIELDS = (
    ("operating_revenue", "营业收入", 1),
    ("operating_cost", "营业成本", -1),
    ("taxes_surcharges", "税金及附加", -1),
    ("selling_expense", "销售费用", -1),
    ("administrative_expense", "管理费用", -1),
    ("rd_expense", "研发费用", -1),
    ("finance_expense", "财务费用", -1),
    ("other_income", "其他收益", 1),
    ("investment_income", "投资收益", 1),
    ("hedging_gain", "净敞口套期收益", 1),
    ("fair_value_gain", "公允价值变动收益", 1),
    ("credit_impairment", "信用减值损失", 1),
    ("asset_impairment", "资产减值损失", 1),
    ("asset_disposal", "资产处置收益", 1),
    ("operating_profit", "营业利润", None),
    ("nonoperating_income", "营业外收入", 1),
    ("nonoperating_expense", "营业外支出", -1),
    ("profit_total", "利润总额", None),
)
PREPAID_LABELS = {
    "19": "加：特定业务计算的应纳税所得额", "20": "减：不征税收入",
    "21": "减：资产加速折旧、摊销（扣除）调减额", "22": "减：免税收入、减计收入、加计扣除",
    "23": "减：所得减免", "24": "减：弥补以前年度亏损", "25": "实际利润额",
    "26": "税率（25%）", "27": "应纳所得税额", "28": "减：减免所得税额",
    "29": "减：抵免所得税额", "30": "减：本年累计已预缴所得税额",
    "31": "减：特定业务预缴（征）所得税额", "32": "本期应补（退）所得税额",
    "33": "总机构本期分摊应补（退）所得税额", "34": "其中：总机构分摊应补（退）所得税额",
    "35": "财政集中分配应补（退）所得税额", "36": "总机构主体生产经营部门分摊应补（退）所得税额",
    "37": "分支机构本期分摊比例", "38": "分支机构本期分摊应补（退）所得税额",
    "39": "减：民族自治地区企业所得税地方分享部分", "40": "实际应补（退）所得税额",
}
ANNUAL_LABELS = {
    "19": "减：境外所得", "20": "加：纳税调整增加额", "21": "减：纳税调整减少额",
    "22": "减：免税、减计收入及加计扣除", "23": "加：境外应税所得抵减境内亏损",
    "24": "纳税调整后所得", "25": "减：所得减免", "26": "减：弥补以前年度亏损",
    "27": "减：抵扣应纳税所得额", "28": "应纳税所得额", "29": "税率（25%）",
    "30": "应纳所得税额", "31": "减：减免所得税额", "32": "减：抵免所得税额",
    "33": "应纳税额", "34": "加：境外所得应纳所得税额", "35": "减：境外所得抵免所得税额",
    "36": "实际应纳所得税额", "37": "减：本年累计预缴所得税额", "38": "本年应补（退）所得税额",
    "39": "其中：总机构分摊本年应补（退）所得税额", "40": "财政集中分配本年应补（退）所得税额",
    "41": "总机构主体生产经营部门分摊本年应补（退）所得税额",
    "42": "减：民族自治地区企业所得税地方分享部分", "43": "减：稽查查补（退）所得税额",
    "44": "减：特别纳税调整补（退）所得税额", "45": "本年实际应补（退）所得税额",
}


def _money(value: Decimal) -> Decimal:
    return value.quantize(CENT, rounding=ROUND_HALF_UP)


def _serial(value: Any) -> Any:
    if isinstance(value, Decimal):
        return format(value, "f")
    if isinstance(value, (list, tuple)):
        return [_serial(item) for item in value]
    if isinstance(value, dict):
        return {str(k): _serial(v) for k, v in value.items()}
    return value


class Rules:
    def __init__(self, return_id, financials, facts, context):
        self.return_id, self.financials, self.facts, self.context = return_id, financials, facts, context
        self.prepaid = return_id == "A200000"
        self.labels = (PREPAID_LABELS if self.prepaid else ANNUAL_LABELS).copy()
        self.url = PREPAID_FORM_URL if self.prepaid else ANNUAL_FORM_URL
        self.values, self.lines, self.issues, self.unresolved = {}, {}, [], set()
        self.financial_values, self.financial_evidence = {}, {}
        self.tax_values, self.tax_evidence = {}, {}
        self.scope_evidence = {}
        self.qualification = None

    def issue(self, code, field, message):
        self.issues.append({"code": code, "field": field, "message": message})
        self.unresolved.add(field)

    def number(self, value, field, nonnegative=False):
        if isinstance(value, bool) or value is None or value == "":
            self.issue("MISSING_OR_INVALID_AMOUNT", field, "需要明确数值，空白不能当零")
            return None
        try:
            # Floats are accepted only by their visible decimal representation;
            # caller should pass strings/Decimal from the statement adapter.
            result = Decimal(str(value))
            if (not result.is_finite() or abs(result) > Decimal("1e18")
                    or result.as_tuple().exponent < -12 or len(result.as_tuple().digits) > 40):
                raise InvalidOperation
        except (InvalidOperation, ValueError, TypeError):
            self.issue("INVALID_AMOUNT", field, "仅支持有限十进制数值，不能包含货币符号、公式或单位")
            return None
        if nonnegative and result < 0:
            self.issue("NEGATIVE_TAX_INPUT", field, "该专项金额必须为非负数")
            return None
        return result

    def fact(self, obj, field, expected=None, numeric=False, nonnegative=False):
        if not isinstance(obj, dict) or "value" not in obj or not isinstance(obj.get("evidence"), str) or not obj["evidence"].strip():
            self.issue("FACT_EVIDENCE_REQUIRED", field, "需要 value 和可复核的 evidence，不接受自动 true 或无来源金额")
            return None
        value = obj["value"]
        if expected is not None and (value != expected or type(value) is not type(expected)):
            self.issue("UNSUPPORTED_SCOPE", field, f"当前分支要求明确确认 {expected!r}，实际为 {value!r}，需补充专项适配")
            return None
        return self.number(value, field, nonnegative) if numeric else value

    def line(self, number, value, calculation, source_fields, rule_url=None, unit="元"):
        number = str(number)
        self.values[number] = value if unit == "rate" else _money(value)
        self.lines[number] = {"label": self.labels.get(number, ""), "value": format(self.values[number], "f"),
            "unit": unit, "calculation": calculation, "source_fields": _serial(source_fields),
            "rule_url": rule_url or self.url}

    def validate_context(self):
        if self.return_id not in ("A200000", "A100000"):
            self.issue("RETURN_UNSUPPORTED", "return_id", "只适配 A200000（2025年第17号）及 A100000（2025年第1号）")
            return
        if self.context.get("currency", "CNY") != "CNY":
            self.issue("CURRENCY_UNSUPPORTED", "context.currency", "EIT 计算模块只接受人民币元，不做换汇")
        if self.context.get("amount_unit", "元") != "元":
            self.issue("UNIT_NOT_YUAN", "context.amount_unit", "调用前须把所有来源金额显式归一化为元")
        try:
            self.start = date.fromisoformat(self.context["period_start"])
            self.end = date.fromisoformat(self.context["period_end"])
            if self.start > self.end or self.start.year != self.end.year:
                raise ValueError
        except (KeyError, ValueError, TypeError):
            self.issue("CONTEXT_PERIOD", "context.period", "需要同一年度的明确 ISO 起止期")
            return
        if not self.prepaid:
            if self.start != date(self.start.year, 1, 1) or self.end != date(self.start.year, 12, 31):
                self.issue("ANNUAL_PERIOD", "context.period", "年度路径仅适配完整自然年，中途开停业及清算需专项处理")
            if not 2024 <= self.end.year <= 2027:
                self.issue("ADAPTER_REVIEW_WINDOW", "context.period", "年度适配已核验窗口为2024—2027；窗口外需复核表版与规则")
        else:
            kind = self.context.get("period_kind")
            valid_month = self.start.day == 1 and self.end.day == monthrange(self.end.year, self.end.month)[1]
            valid_month = valid_month and self.start.month == self.end.month
            valid_quarter = self.start.day == 1 and self.start.month in (1, 4, 7, 10)
            valid_quarter = valid_quarter and self.end.month == self.start.month + 2 and self.end.day == monthrange(self.end.year, self.end.month)[1]
            if not ((kind == "month" and valid_month) or (kind == "quarter" and valid_quarter)):
                self.issue("PREPAID_PERIOD", "context.period", "预缴路径需要完整自然月或自然季度；新设/终止经营首期需专项处理")
            if not (date(2026, 1, 1) <= self.start <= self.end <= date(2027, 12, 31)):
                self.issue("ADAPTER_REVIEW_WINDOW", "context.period", "预缴适配已核验窗口为2026—2027；2025及更早所属期应按当期版本另行处理")
        scope = self.facts.get("scope", {})
        if not isinstance(scope, dict):
            self.issue("SCOPE_REQUIRED", "facts.scope", "需要明确企业类型和业务范围证据")
            return
        expected = {"resident_enterprise": True, "collection_method": "actual_profit", "independent_taxpayer": True,
                    "no_foreign_business": True, "no_special_business": True}
        for key, value in expected.items():
            evidence = scope.get(key)
            self.fact(evidence, f"facts.scope.{key}", expected=value)
            if isinstance(evidence, dict):
                self.scope_evidence[key] = evidence

    def read_financials(self):
        zeros = self.facts.get("confirmed_zero_financial_fields", [])
        valid_zero_evidence = isinstance(self.facts.get("confirmation_evidence"), str) and bool(self.facts["confirmation_evidence"].strip())
        if not isinstance(zeros, list) or any(not isinstance(key, str) for key in zeros):
            self.issue("ZERO_LIST_INVALID", "facts.confirmed_zero_financial_fields", "应为逐项字段名列表")
            zeros = []
        known = {field for field, _, _ in FINANCIAL_FIELDS}
        for unknown in set(zeros) - known:
            self.issue("ZERO_FIELD_UNKNOWN", f"financials.{unknown}", "无效财务字段名，不能静默忽略")
        for key, label, _ in FINANCIAL_FIELDS:
            raw = self.financials.get(key)
            if raw is None:
                if key in zeros and valid_zero_evidence:
                    value, evidence = ZERO, {"confirmation": self.facts["confirmation_evidence"], "explicit_zero_field": key}
                else:
                    self.issue("FINANCIAL_MISSING", f"financials.{key}", f"缺少{label}；须从利润表提取或明确逐项确认无此事项")
                    continue
            else:
                payload = raw.get("value") if isinstance(raw, dict) else raw
                value = self.number(payload, f"financials.{key}")
                evidence = raw.get("source", raw.get("evidence", {"field": key, "coordinates_provided": False})) if isinstance(raw, dict) else {"field": key, "coordinates_provided": False}
                if value is None:
                    continue
                if key in zeros and value != 0:
                    self.issue("ZERO_CONFIRMATION_CONFLICT", f"financials.{key}", "来源金额非零，与无事项确认冲突")
            self.financial_values[key] = _money(value)
            self.financial_evidence[key] = {"field": key, "value_yuan": format(value, "f"), "source": evidence}
        if len(self.financial_values) != len(FINANCIAL_FIELDS):
            return False
        vals = self.financial_values
        operating = sum((vals[key] * sign for key, _, sign in FINANCIAL_FIELDS[:14]), ZERO)
        total = operating + vals["nonoperating_income"] - vals["nonoperating_expense"]
        if operating != vals["operating_profit"]:
            self.issue("FINANCIAL_RECONCILIATION", "financials.operating_profit", f"项目重算营业利润{operating}元与来源{vals['operating_profit']}元不一致；检查减值符号、费用重复或科目映射")
        if total != vals["profit_total"]:
            self.issue("FINANCIAL_RECONCILIATION", "financials.profit_total", f"项目重算利润总额{total}元与来源{vals['profit_total']}元不一致")
        for index, (key, label, _) in enumerate(FINANCIAL_FIELDS, 1):
            self.labels[str(index)] = label
            if key == "operating_profit":
                calc = "1-2-3-4-5-6-7+8+9+10+11+12+13+14；独立核对利润表营业利润"
                value, sources = operating, [self.financial_evidence[k] for k, _, _ in FINANCIAL_FIELDS[:15]]
            elif key == "profit_total":
                calc = "15+16-17；独立核对利润表利润总额"
                value, sources = total, [self.financial_evidence[k] for k in ("operating_profit", "nonoperating_income", "nonoperating_expense", "profit_total")]
            else:
                calc, value, sources = "从利润表本年累计金额取数，保留原带符号值", vals[key], [self.financial_evidence[key]]
            self.line(str(index), value, calc, sources)
        return True

    def tax_fact(self, number, allow_negative=False):
        number = str(number)
        field = f"facts.tax_lines.{number}"
        items = self.facts.get("tax_lines", {})
        if not isinstance(items, dict):
            self.issue("TAX_LINES_INVALID", "facts.tax_lines", "应为以税表行号为键的事实对象")
            return None
        item = items.get(number)
        zeros = self.facts.get("confirmed_zero_tax_lines", [])
        if not isinstance(zeros, list) or any(not isinstance(key, str) for key in zeros):
            self.issue("ZERO_LIST_INVALID", "facts.confirmed_zero_tax_lines", "应为字符串行号逐项列表")
            zeros = []
        if item is not None:
            value = self.fact(item, field, numeric=True, nonnegative=not allow_negative)
            if value is None:
                return None
            if number in zeros and value != 0:
                self.issue("ZERO_CONFIRMATION_CONFLICT", field, "专项非零金额与明确无事项确认冲突")
            evidence = {"tax_line": number, "evidence": item["evidence"], "input_value": format(value, "f")}
        elif number in zeros and isinstance(self.facts.get("confirmation_evidence"), str) and self.facts["confirmation_evidence"].strip():
            value = ZERO
            evidence = {"tax_line": number, "explicit_zero": True, "evidence": self.facts["confirmation_evidence"]}
        else:
            self.issue("TAX_FACT_MISSING", field, f"需要第{number}行{self.labels.get(number, '')}的来源金额，或逐项确认无事项")
            return None
        value = _money(value)
        self.tax_values[number], self.tax_evidence[number] = value, evidence
        self.line(number, value, "专项底稿已确认金额" if item is not None else "用户逐项明确确认该行无事项", [evidence])
        return value

    def cap(self, value, available, field):
        if value > max(available, ZERO):
            self.issue("DEDUCTION_EXCEEDS_BASE", field, f"减免/扣除{value}元超过可用金额{max(available, ZERO)}元，不能自动封顶掩盖事实错误")

    def small_qualification(self, taxable):
        details = self.facts.get("small_low_profit", {})
        if not isinstance(details, dict):
            details = {}
        self.fact(details.get("non_restricted_industry"), "facts.small_low_profit.non_restricted_industry", expected=True)
        expected_basis = "ytd_quarter_average" if self.prepaid else "annual_quarter_average"
        self.fact(details.get("average_basis"), "facts.small_low_profit.average_basis", expected=expected_basis)
        employees = assets = None
        quarters = details.get("quarters")
        if quarters is not None:
            required = (self.end.month - 1) // 3 + 1 if self.prepaid else 4
            if not isinstance(quarters, list) or len(quarters) != required:
                self.issue("SMALL_QUARTERS_MISSING", "facts.small_low_profit.quarters", f"需第1至{required}季度连续期初/末人数和资产快照")
            else:
                avgs, seen = [], set()
                for row in quarters:
                    if not isinstance(row, dict) or type(row.get("quarter")) is not int or row["quarter"] not in range(1, required + 1) or row["quarter"] in seen:
                        self.issue("SMALL_QUARTER_SEQUENCE", "facts.small_low_profit.quarters", "季度编号必须连续且不重复")
                        continue
                    seen.add(row["quarter"])
                    parsed = [self.fact(row.get(key), f"facts.small_low_profit.quarters.{row['quarter']}.{key}", numeric=True, nonnegative=True)
                              for key in ("employee_start", "employee_end", "asset_start_yuan", "asset_end_yuan")]
                    if all(v is not None for v in parsed):
                        if any(v != v.to_integral_value() for v in parsed[:2]):
                            self.issue("EMPLOYEE_COUNT", "facts.small_low_profit.quarters", "期初和期末从业人数必须为整数，包括接受的劳务派遣用工")
                        avgs.append(((parsed[0] + parsed[1]) / 2, (parsed[2] + parsed[3]) / 2))
                if len(avgs) == required:
                    employees = sum((a[0] for a in avgs), ZERO) / required
                    assets = sum((a[1] for a in avgs), ZERO) / required
                    for name, computed in (("employee_average", employees), ("asset_average_yuan", assets)):
                        if name in details:
                            supplied = self.fact(details[name], f"facts.small_low_profit.{name}", numeric=True, nonnegative=True)
                            if supplied is not None and supplied != computed:
                                self.issue("SMALL_AVERAGE_CONFLICT", f"facts.small_low_profit.{name}", "已填平均值与季度快照计算结果冲突，不能静默选用其中一个")
        else:
            employees = self.fact(details.get("employee_average"), "facts.small_low_profit.employee_average", numeric=True, nonnegative=True)
            assets = self.fact(details.get("asset_average_yuan"), "facts.small_low_profit.asset_average_yuan", numeric=True, nonnegative=True)
        if self.end > date(2027, 12, 31):
            self.issue("SMALL_POLICY_EXPIRED", "facts.tax_treatment", "当前已核实小型微利企业优惠有效至2027-12-31，之后需更新政策")
        if employees is not None and employees > 300:
            self.issue("SMALL_NOT_ELIGIBLE", "facts.small_low_profit.employee_average", "从业人数季度平均值超过300，不能按小型微利企业优惠")
        if assets is not None and assets > Decimal("50000000"):
            self.issue("SMALL_NOT_ELIGIBLE", "facts.small_low_profit.asset_average_yuan", "资产总额季度平均值超过5000万元，不能按小型微利企业优惠")
        if taxable > Decimal("3000000"):
            self.issue("SMALL_NOT_ELIGIBLE", "facts.tax_treatment", "应纳税所得额/预缴实际利润额超过300万元，不能对超出部分继续套用优惠")
        self.qualification = {"employee_average": _serial(employees), "asset_average_yuan": _serial(assets),
                              "taxable_income_yuan": _serial(taxable), "average_basis": expected_basis, "rule_url": SMALL_URL}

    def relief(self, taxable, gross_tax, number):
        treatment = self.facts.get("tax_treatment")
        evidence = self.facts.get("tax_treatment_evidence")
        if treatment not in ("standard_25", "small_low_profit", "confirmed_relief") or not isinstance(evidence, str) or not evidence.strip():
            self.issue("TAX_TREATMENT_REQUIRED", "facts.tax_treatment", "须明确选择 standard_25/small_low_profit/confirmed_relief 并提供 tax_treatment_evidence，不能默认为无优惠")
            return None
        if treatment == "confirmed_relief":
            if not isinstance(self.facts.get("tax_lines", {}).get(number), dict):
                self.issue("RELIEF_EVIDENCE_REQUIRED", f"facts.tax_lines.{number}", "confirmed_relief 必须提供具体减免金额和来源，不能只列无事项")
                return None
            value = self.tax_fact(number)
            if value is None:
                return None
        else:
            if treatment == "small_low_profit":
                self.small_qualification(taxable)
                net_tax = _money(max(taxable, ZERO) * Decimal("0.25") * Decimal("0.20"))
                value = gross_tax - net_tax
                calc, url = "名义税率25%；小微有效应纳税额=应纳税所得额×25%×20%；减免=名义应纳税额-优惠后税额", SMALL_URL
            else:
                value, calc, url = ZERO, "已明确选择普通25%且不享受减免税额待遇，本行=0", self.url
            item = self.facts.get("tax_lines", {}).get(number)
            if item is not None:
                stated = self.fact(item, f"facts.tax_lines.{number}", numeric=True, nonnegative=True)
                if stated is not None and _money(stated) != value:
                    self.issue("RELIEF_CONFLICT", f"facts.tax_lines.{number}", "申报减免源值与所选择税收待遇计算结果冲突")
            if number in self.facts.get("confirmed_zero_tax_lines", []) and value != 0:
                self.issue("RELIEF_CONFLICT", f"facts.tax_lines.{number}", "减免计算结果非零，与该行无事项确认冲突")
            self.line(number, value, calc, [{"tax_treatment": treatment, "evidence": evidence}, self.qualification] if self.qualification else [{"tax_treatment": treatment, "evidence": evidence}], rule_url=url)
        self.cap(value, gross_tax, f"facts.tax_lines.{number}")
        return value

    def read_tax_inputs(self):
        rows = ("19", "20", "21", "22", "23", "24", "29", "30", "31", "39") if self.prepaid else (
            "19", "20", "21", "22", "23", "25", "26", "27", "32", "34", "35", "37", "42", "43", "44")
        allowed = set(rows) | ({"28"} if self.prepaid else {"31"})
        item_map = self.facts.get("tax_lines", {})
        if isinstance(item_map, dict):
            for unknown in set(item_map) - allowed:
                self.issue("TAX_LINE_UNSUPPORTED", f"facts.tax_lines.{unknown}", "该行由公式生成或不属于当前表版的专项输入行，不能静默忽略")
        zeros = self.facts.get("confirmed_zero_tax_lines", [])
        if isinstance(zeros, list) and all(isinstance(item, str) for item in zeros):
            for unknown in set(zeros) - allowed:
                self.issue("TAX_LINE_UNSUPPORTED", f"facts.confirmed_zero_tax_lines.{unknown}", "无事项清单包含非专项输入行或错误版行号")
        for row in rows:
            self.tax_fact(row, allow_negative=(not self.prepaid and row in ("42", "43", "44")))
        if self.prepaid:
            for row in ("19", "31"):
                if self.tax_values.get(row, ZERO) != 0:
                    self.issue("SPECIAL_BUSINESS_REQUIRED", f"facts.tax_lines.{row}", "该非零特定业务金额超出当前普通独立企业路径，需接入对应专项申报分支")
        else:
            for row in ("19", "23", "34", "35"):
                if self.tax_values.get(row, ZERO) != 0:
                    self.issue("FOREIGN_BUSINESS_REQUIRED", f"facts.tax_lines.{row}", "非零境外项目与 no_foreign_business 确认冲突，需要境外所得专项适配")
        return len(self.tax_values) == len(rows)

    def compute(self):
        self.validate_context()
        financial_ok = self.read_financials()
        tax_ok = self.read_tax_inputs()
        if not financial_ok or not tax_ok or not hasattr(self, "end"):
            return
        f, t = self.financial_values, self.tax_values
        if self.prepaid:
            if f["investment_income"] != 0:
                items = self.facts.get("investment_income_items")
                if not isinstance(items, list) or not items:
                    self.issue("INVESTMENT_DETAIL_REQUIRED", "facts.investment_income_items", "新版预缴第9行非零时需投资收益事项明细，不能只填总额而遗漏9.1等事项")
                else:
                    amounts = []
                    for index, item in enumerate(items, 1):
                        row = f"9.{index}"
                        if not isinstance(item, dict) or not isinstance(item.get("name"), str) or not item["name"].strip():
                            self.issue("INVESTMENT_DETAIL_REQUIRED", f"facts.investment_income_items.{index}", "投资收益明细需要具体事项名称")
                            continue
                        value = self.fact(item, f"facts.investment_income_items.{index}", numeric=True)
                        if value is not None:
                            amounts.append(_money(value))
                            self.labels[row] = item["name"]
                            self.line(row, value, "来源已确认的投资收益事项，保留损失负号", [{"name": item["name"], "evidence": item["evidence"]}])
                    if len(amounts) != len(items) or sum(amounts, ZERO) != f["investment_income"]:
                        self.issue("INVESTMENT_RECONCILIATION", "facts.investment_income_items", "投资事项明细合计与利润表投资收益不一致")
            before_loss = f["profit_total"] + t["19"] - t["20"] - t["21"] - t["22"] - t["23"]
            self.cap(t["24"], before_loss, "facts.tax_lines.24")
            taxable = before_loss - t["24"]
            self.line("25", taxable, "18+19-20-21-22-23-24，保留实际利润额公式结果的负号", ["18", "19", "20", "21", "22", "23", "24"])
            self.line("26", RATE, "企业所得税主表名义税率25%；小微优惠在第28行反映", [], unit="rate")
            gross = _money(max(taxable * RATE, ZERO))
            self.line("27", gross, "max(0,25×26)，实际利润为负时不产生负所得税", ["25", "26"])
            relief = self.relief(taxable, gross, "28")
            if relief is None:
                return
            self.cap(t["29"], gross - relief, "facts.tax_lines.29")
            due = max(gross - relief - t["29"] - t["30"] - t["31"], ZERO)
            self.line("32", due, "max(0,27-28-29-30-31)；预缴多缴不在本期直接报负退款", ["27", "28", "29", "30", "31"], rule_url=PREPAID_ORACLE_URL)
            for row in ("33", "34", "35", "36", "37", "38"):
                self.line(row, ZERO, "已明确确认独立纳税，非汇总纳税总分机构，本行不适用", [self.scope_evidence.get("independent_taxpayer")], unit="rate" if row == "37" else "元")
            self.cap(t["39"], due, "facts.tax_lines.39")
            self.line("40", max(due - t["39"], ZERO), "max(0,32-39)，适用于当前独立纳税路径", ["32", "39"])
        else:
            adjusted = f["profit_total"] - t["19"] + t["20"] - t["21"] - t["22"] + t["23"]
            self.line("24", adjusted, "18-19+20-21-22+23", ["18", "19", "20", "21", "22", "23"])
            self.cap(t["25"], adjusted, "facts.tax_lines.25")
            self.cap(t["26"], adjusted - t["25"], "facts.tax_lines.26")
            self.cap(t["27"], adjusted - t["25"] - t["26"], "facts.tax_lines.27")
            taxable = adjusted - t["25"] - t["26"] - t["27"]
            self.line("28", taxable, "24-25-26-27；负差场景的2025版栏次显示规则尚未确认时阻断交付", ["24", "25", "26", "27"])
            if taxable < 0:
                self.issue("ANNUAL_NEGATIVE_BASE_DISPLAY_UNVERIFIED", "lines.28",
                           "年度第28行计算差额为负，尚未取得2025年第1号现行填表说明的负值显示原句；保留原始差额，仅供复核，不生成可填报结果")
            self.line("29", RATE, "企业所得税主表名义税率25%；小微优惠在第31行反映", [], unit="rate")
            gross = _money(max(taxable * RATE, ZERO))
            self.line("30", gross, "max(0,28×29)；年度负基础场景同时受第28行待核规则阻断", ["28", "29"])
            relief = self.relief(taxable, gross, "31")
            if relief is None:
                return
            self.cap(t["32"], gross - relief, "facts.tax_lines.32")
            due = gross - relief - t["32"]
            self.line("33", due, "30-31-32", ["30", "31", "32"])
            actual = due + t["34"] - t["35"]
            self.line("36", actual, "33+34-35", ["33", "34", "35"])
            balance = actual - t["37"]
            self.line("38", balance, "36-37；年度预缴超过全年税额时允许负数退税", ["36", "37"])
            for row in ("39", "40", "41"):
                self.line(row, ZERO, "已明确确认独立纳税，非汇总纳税总机构，本行不适用", [self.scope_evidence.get("independent_taxpayer")])
            if t["42"] > 0:
                self.cap(t["42"], actual, "facts.tax_lines.42")
            self.line("45", balance - t["42"] - t["43"] - t["44"], "38-42-43-44；负数表示待复核的年度退税额", ["38", "42", "43", "44"])

    def result(self):
        return {"status": STATUS, "outcome": "BLOCKED" if self.issues else "GENERATED_FOR_REVIEW",
            "return_id": self.return_id, "version": VERSION, "ruleset_version": RULESET_VERSION,
            "form_version": "国家税务总局公告2025年第17号" if self.prepaid else "国家税务总局公告2025年第1号",
            "form_rule_url": PREPAID_URL if self.prepaid else ANNUAL_URL,
            "adapter_review_window": {"start": "2026-01-01" if self.prepaid else "2024-01-01", "end": "2027-12-31",
                                      "meaning": "实现复核窗口，不代表表版或普通25%税率在窗口终止时失效"},
            "small_policy_valid_until": "2027-12-31", "qualification": self.qualification,
            "lines": self.lines, "issues": self.issues, "unresolved_fields": sorted(self.unresolved),
            "human_review_required": True, "filing_performed": False}


def calculate(return_id: str, financials: dict, facts: dict, context: dict) -> dict:
    """Return source-bound draft lines; missing evidence always yields BLOCKED.

    Inputs are detailed in the bundled reference. No input dictionaries are
    mutated. All return values are JSON-serializable (Decimal values as strings).
    """
    if not all(isinstance(obj, dict) for obj in (financials, facts, context)):
        return {"status": STATUS, "outcome": "BLOCKED", "return_id": return_id, "version": VERSION,
                "lines": {}, "issues": [{"code": "INPUT_SCHEMA", "field": "input", "message": "financials/facts/context 必须为字典"}],
                "unresolved_fields": ["input"], "human_review_required": True, "filing_performed": False}
    engine = Rules(return_id, financials, facts, context)
    try:
        with localcontext() as ctx:
            ctx.prec = 50
            engine.compute()
    except Exception as error:
        engine.issue("RULES_INPUT_ERROR", "input", f"{type(error).__name__}: {error}")
    return _serial(engine.result())
