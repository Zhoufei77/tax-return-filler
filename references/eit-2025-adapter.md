# 企业所得税真实业务适配：A200000 / A100000

本适配把利润表本年累计项目和必要专项事实转换成逐行计算结果，供填入有版本绑定的申报工作底稿。经办人不必先手算营业利润、计税基础、减免额或应补退税。

## 表版与实现窗口

| 路径 | 已核实官方表版 | 当前实现接收的所属期 |
|---|---|---|
| 月度、季度预缴 A200000 | 国家税务总局公告2025年第17号；2025-10-01施行 | 2026—2027完整自然月/季度 |
| 年度 A100000 | 国家税务总局公告2025年第1号，2024年度及以后汇算使用 | 2024—2027完整自然年 |

以上结束日期是本实现的**维护复核窗口**，不表示普通25%税率或表单在2027年底失效。小型微利企业优惠另有已核实的2027-12-31政策期限。2025年预缴、新设/中途终止经营、清算、境外业务、汇总纳税总分机构等进入明确未适配分支，不套用本实现。

权威入口：

- [预缴2025年第17号公告](https://fgk.chinatax.gov.cn/zcfgk/c100012/c5241820/content.html)、[现行预缴表样](https://12366.chinatax.gov.cn/bzds/059/059.html)。
- [年度2025年第1号公告](https://fgk.chinatax.gov.cn/zcfgk/c100012/c5237968/content.html)、[现行年度表样](https://12366.chinatax.gov.cn/bzds/056/056.html)。
- [2026年上海年度汇缴提示仍指定2025年第1号版](https://shanghai.chinatax.gov.cn/xwdt/ztzl/zcgll/hsqjzl/bslcqy/202603/t479726.html)。

当前年度第28行公式差额为负的情形尚未锁定2025年第1号现行说明的负值显示原句：模块保留原始负差并返回 `ANNUAL_NEGATIVE_BASE_DISPLAY_UNVERIFIED` / `BLOCKED`。不得把旧表同名项目规则冒充新表已验证规则。**年度计税基础为正、因预缴过多产生最终负退税额的路径正常支持。**

## API 与单位

```python
from eit_rules import calculate

result = calculate("A200000", financials, facts, context)
if result["outcome"] != "GENERATED_FOR_REVIEW":
    # 展示本轮合并缺口，补事实后重新运行；不能交付已填表。
    print(result["issues"], result["unresolved_fields"])
```

`context` 至少含 `period_start`、`period_end`、`period_kind`（`month` / `quarter` / `year`）。编排器同时传入主体和表单身份，复用主填表引擎的身份、SHA256及目标锚点检查。

所有财务和税务金额须在调用前明确换算为 **CNY元**。`context.amount_unit` 为 `元`，`currency` 为 `CNY`；人数不换算，资产资格输入也是元。主表之外展示资产总额万元时，由输出映射显式除以10000。

内部采用 `Decimal`；申报金额行按 `ROUND_HALF_UP` 保留两位。税率行使用字符串 `0.25`、`unit="rate"`，由模板百分比格式显示25%，不能把小微有效5%写入名义税率行。输入不得带公式、货币符号或单位。

## 从利润表自动提取的18项

`financials` 的键与 `statement_adapter` 一致。每值可以为数值字符串或 `Decimal`；生产交付优先使用 `{ "value": "12345.67", "source": {"path": "利润表.xlsx", "sheet": "利润表", "cell": "D7", "sha256": "..."} }`，以保留原文件坐标。`source` 可为字符串或可JSON序列化对象。只传裸金额不等于机器核实了来源。

| 行 | 财务键 | 对应项目 | 营业利润计算符号 |
|---|---|---|---|
| 1 | operating_revenue | 营业收入 | + |
| 2 | operating_cost | 营业成本 | − |
| 3 | taxes_surcharges | 税金及附加 | − |
| 4 | selling_expense | 销售费用 | − |
| 5 | administrative_expense | 管理费用 | − |
| 6 | rd_expense | 研发费用 | − |
| 7 | finance_expense | 财务费用 | − |
| 8 | other_income | 其他收益 | + |
| 9 | investment_income | 投资收益 | +，保留损失负号 |
| 10 | hedging_gain | 净敞口套期收益 | +，保留损失负号 |
| 11 | fair_value_gain | 公允价值变动收益 | +，保留损失负号 |
| 12 | credit_impairment | 信用减值损失 | +，保留原负号 |
| 13 | asset_impairment | 资产减值损失 | +，保留原负号 |
| 14 | asset_disposal | 资产处置收益 | +，保留损失负号 |
| 15 | operating_profit | 营业利润 | 用源表金额作独立校验 |
| 16 | nonoperating_income | 营业外收入 | 加到第15行 |
| 17 | nonoperating_expense | 营业外支出 | 从第15行减除 |
| 18 | profit_total | 利润总额 | 用源表金额作独立校验 |

两表前18行按同一财务结构取**本年累计金额**。第15行独立按第1—14行重算，第18行按15+16−17重算，与源表15/18逐分核对。不能用收入减成本代替利润总额。

18项必须有值。明确不发生的项目可以列入 `facts.confirmed_zero_financial_fields`，同时提供 `facts.confirmation_evidence`，不得因为没搜到行而自动填0。小企业会计准则表中研发可能已计入管理费用；未完成拆分确认前不能重复扣除。不能把净利润或现金流金额当利润总额。

预缴第9行非零时，另需 `facts.investment_income_items=[{"name":"具体收益事项", "value":"金额", "evidence":"明细来源"}]`，明细合计须等于投资收益。模块输出9.1等明细，防止只填总额遗漏新版事项。

## 必要专项事实

```json
{
  "tax_lines": {
    "30": {"value": "10000.00", "evidence": "本年已缴税款记录.xlsx!B5"}
  },
  "confirmed_zero_tax_lines": ["19", "20", "21", "22", "23", "24", "29", "31", "39"],
  "confirmation_evidence": "用户逐项明确确认清单中的业务未发生；会话或记录位置",
  "tax_treatment": "standard_25",
  "tax_treatment_evidence": "用户明确选择普通25%且不享受减免税额待遇；记录位置"
}
```

每条专项金额须有 `value` 和非空 `evidence`。无事项清单只能逐项列出，`true`、空对象及不具备确认依据的全零清单无效。自动计算行不能作为手填专项行混入。

| 路径 | 必需的专项输入行 |
|---|---|
| A200000 | 19、20、21、22、23、24、29、30、31、39 |
| A100000 | 19、20、21、22、23、25、26、27、32、34、35、37、42、43、44 |

这些金额分别涉及特定业务、税会调整、不征税或优惠收入、加速折旧/研发等专项、亏损弥补、抵免和历史已缴税款，不能从一张普通利润表猜出。若已有专项台账，应自动提取；只有缺少的事实才向用户合并追问。年度43/44、年度地方分享优惠调整42可含已确认的负退款或调整金额，不能把这类已缴税额调整和应税所得调整混淆。

范围也须真实确认：

```json
{
  "scope": {
    "resident_enterprise": {"value": true, "evidence": "企业类型来源或用户确认"},
    "collection_method": {"value": "actual_profit", "evidence": "居民查账及预缴方式来源"},
    "independent_taxpayer": {"value": true, "evidence": "非汇总纳税总分机构的来源"},
    "no_foreign_business": {"value": true, "evidence": "无境外所得及境外抵免业务的确认"},
    "no_special_business": {"value": true, "evidence": "无清算、房地产预售、代理出口等未适配特殊业务的确认"}
  }
}
```

这些字段表示当前算法分支，不能由Agent替用户自动设成true。出现对应业务时返回具体未覆盖字段。独立纳税确认后，A200000第33—38行和A100000第39—41行作不适用0；不表示实现了总分机构算法。

## 计算与优惠

预缴25=18+19−20−21−22−23−24，保留实际利润额负号；27=max(25×25%,0)。32=max(27−28−29−30−31,0)，40=max(32−39,0)。实际利润为负不产生负应纳税额；已有多缴税款不会直接在本次预缴输出负退款。[现行表公式](https://12366.chinatax.gov.cn/bzds/059/059.html)、[官方预缴金额算例](https://fgk.chinatax.gov.cn/zcfgk/c100015/c5241823/content.html)、[官方既有负利润税额非负说明](https://shanghai.chinatax.gov.cn/bsfw/xzzx/bgxz/sbzsl/201905/t445436.html)。最后一项是历史官方说明，未将其旧行号冒充现行行号。

年度24=18−19+20−21−22+23；28=24−25−26−27（负差场景目前阻断）；30=28×25%；33=30−31−32；36=33+34−35；38=36−37；45=38−42−43−44。年度38/45可为负数，表示待复核退税额。[现行年度表公式](https://12366.chinatax.gov.cn/bzds/056/056.html)

`tax_treatment` 必须选择一种，均须 `tax_treatment_evidence`：

- `standard_25`：明确选择无减免税额待遇后，预缴28/年度31为0，名义25%计税；不会默认替用户放弃优惠。
- `small_low_profit`：核对真实资格后，以应税所得×25%×20%计算优惠后税额，差额记在预缴28/年度31。
- `confirmed_relief`：预缴28/年度31必须有具体来源金额；本模块核对其上限，但不替代对应高新、地区等优惠资格及计算底稿。

减免不得超过名义应纳税额；抵免不得超过减免后的应纳税额；亏损弥补不得超过扣除前可用所得。异常会阻断，不能自动压低输入掩盖差错。

小微资格需要：非限制禁止行业、正确口径的平均从业人数不超过300、平均资产不超过5000万元、年度应税所得/预缴实际利润额不超过300万元。超过门槛时整个小微分支阻断，不能对超出部分继续使用5%。[官方资格及计算案例](https://www.chinatax.gov.cn/chinatax/n810356/n3010387/c5211219/content.html)

```json
{
  "small_low_profit": {
    "non_restricted_industry": {"value": true, "evidence": "行业事实来源"},
    "average_basis": {"value": "ytd_quarter_average", "evidence": "截至本期季度平均口径"},
    "employee_average": {"value": "160", "evidence": "人员快照计算记录"},
    "asset_average_yuan": {"value": "30000000", "evidence": "季度资产快照计算记录"}
  }
}
```

年度 `average_basis` 使用 `annual_quarter_average`。有快照时可改用 `quarters` 列表自动算平均：每项含 `quarter`（1至截至本期季度，年度1至4）、`employee_start`、`employee_end`、`asset_start_yuan`、`asset_end_yuan`，后四项各为 `{value,evidence}`。先算各季度（期初+期末）÷2，再算季度平均值的平均；不把全年年初/年末简单平均冒充全年季度平均。人数须包含劳务派遣用工。新设不足完整年度场景不在当前分支。

## 输出与验收

返回 `status="DRAFT_REVIEW_REQUIRED"`，`outcome` 为 `GENERATED_FOR_REVIEW` 或 `BLOCKED`。每条 `lines[行号]` 包含 `label`、数值字符串 `value`、`unit`、`calculation`、`source_fields` 和 `rule_url`，可以直接供工作底稿映射及HTML来源展示。缺失或未覆盖事项通过 `issues` 与 `unresolved_fields` 一次性列出。

`lines` 在阻断时可能含已完成的中间计算，编排器必须检查 `outcome`，不能把这些中间值当成已通过的申报结果。模块不读取文件、不联网、不提交申报；文件读取、原表回填和样式/hash保护由已有填表引擎完成。

运行：

```sh
python3 -m unittest discover -s tests -p test_eit_rules.py -v
```

官方数值验收：

1. [2025预缴官方解读](https://fgk.chinatax.gov.cn/zcfgk/c100015/c5241823/content.html)：累计实际利润400/700/1300/1200万元、此前已缴0/100/175/325万元，25%路径的本期应补为100/75/150/0万元。本测试将该金额序列重放于2026所属期，不宣称已验证案例中的分支机构分摊。
2. [总局小微例](https://www.chinatax.gov.cn/chinatax/n810356/n3010387/c5211219/content.html)：人数120/200、资产2000/4000万元、应税所得190万元，平均160人/3000万元，所得税9.5万元。测试中的完整利润表其余金额为合成数据，核验的是官方资格及税额计算。
3. 正所得年度用合成底稿核验预缴过多的负退税路径；[现行年度解读](https://shanghai.chinatax.gov.cn/zcfw/zcjd/202501/t474914.html)同时提供了年度负退税的真实官方示例。测试不把合成企业说成已办理真实申报。

商业验收仍需脱敏真实利润表、专项底稿和经办会计结果逐行对照，以及WorkBuddy预览环境实测。单元测试通过不等于税务鉴证或全部税种、全部附表均已适配。
