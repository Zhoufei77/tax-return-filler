# 从原始资料到企业所得税主表填写稿

本文供 Agent 执行。用户提供资料和需求，由 Agent 建立任务；**不要让用户编辑 JSON、逐格指定坐标或先手算税额。**

当前已能实际运行的链路是：现有 CSV/XLSX 利润表 → 自动识别18项财务主行及累计列 → 企业所得税规则 → A200000／A100000 主表工作底稿 → 逐行来源和问题报告。会计仍需确认无法从普通利润表推断的业务事实、优惠资格和专项底稿。

年度路径完成的是有依据的主表取数与计算；纳税调整、亏损、抵免等专项须有来源，不等于自动完成全部汇算清缴附表。表版、维护窗口、专项范围详见 [EIT 规则适配](eit-2025-adapter.md)。

## 1. 最少操作的执行顺序

1. 读取本次用户指定的原始利润表、已有专项底稿、上期申报及企业资料。先扫描，区分能够直接取得的信息和真正缺口。不要要求用户先改成固定表头、拆出累计列或重新计算利润总额。
2. 根据实际需求选择 A200000 月／季预缴或 A100000 年度主表。把资料里已经找到的主体、所属期、会计准则、单位和必要缺口汇总成一轮确认；保留用户真实回复或仍适用的已有确认位置。
3. 在客户工作区生成 `case.json`，运行下列 driver。出现阻断时合并解释问题和最少补充动作；补齐后使用新输出目录。成功后打开填写稿和复核报告，检查身份、期间、完整主表行次和金额，再交付会计复核。

```text
<可用Python> "<技能目录>/scripts/statement_adapter.py" scan "<本次资料目录>" --out "<新扫描JSON>" --accounting-standard asbe --presentation software_export

<可用Python> "<技能目录>/scripts/tax_workflow.py" run "<客户工作区>/case.json" --out "<新的交付目录>"
```

第一条用于准备映射和一次性整理缺口；第二条会重新扫描并验证源文件。参数中的准则和报表来源语境必须从资料或确认得到，不能为了识别某列而改参数。Python 和依赖由 Agent 按 [运行说明](runtime.md) 处理，用户不用理解命令行。

## 2. 累计列、单位和原表元数据

两种所得税主表的前18项都取年初至本期末累计金额。不能拿九月本月数或三季度发生额冒充一至九月累计，也不能把净利润当利润总额。

| 真实表头及语境 | 适配行为 |
| --- | --- |
| 本年累计／本年累计金额／年初至报告期末金额 | 明确识别为 YTD，不依赖列顺序 |
| 本月金额／本季金额 | 识别为本月／本季，不能直接用于所得税累计字段 |
| `tax_filing` + 明确执行企业会计准则 `asbe` 的“本期金额” | 按税务财报语境识别为年初至期末累计 |
| `tax_filing` + 明确执行小企业会计准则 `small_asbe` 的“本期金额” | 识别为本期；另找本年累计列 |
| `software_export` 的“本期金额”，或报表语境／准则不明 | 保留歧义，不将所有 ERP 的同名列武断认作累计 |

同名表头、重复财务主行和互相冲突的期间标签不能任选一个。只匹配财务主行，不用“其中”明细、主营业务收入替代总营业收入。借贷发生额的科目余额表也不能直接当利润表；当前链路需要可以核对的利润表，不能自行创造科目归并规则。

金额保留原符号；括号负数会留下明确转换记录。减值损失等项目不因标签含“损失”而再次翻转符号。源公式不读取缓存、不猜算；要求从原财务软件导出可核对的值版，或由用户明确允许的本地表格软件完成另行核验。

来源单位支持元／千元／万元，driver 显式换算为人民币元。缺少单位时，只有真实 `amount_unit` 与 `unit_evidence` 才能解除这一项缺口；原表写元而任务说万元时直接阻断。币种 USD、美元、EUR、欧元、港币等明确外币声明也会阻断，不因任务写 CNY 就假称人民币。

`source_period_start` 必须为该年1月1日，`source_period_end` 必须为目标期末，另保存 `period_evidence`。这仍包含人工声明边界；可见年月、完整 ISO 日期、企业名称及明确身份锚点用于冲突检查，不把声明当成已经审计原表期间或主体。

有真实原格时，建立 `identity_anchors`。每项包含 `cell`、可选 `sheet` 和 `expected_context_key`；后者仅为 `taxpayer_id`、`taxpayer_name`、`period_end`。锚点格必须实际匹配对应上下文；税号按文本保留前导零。找不到锚点时在来源记录中明确 `human_declaration_only`，不要伪造隐藏身份格。

## 3. 首次一轮确认怎么问

先用资料形成摘要，再只询问尚缺且会改变当前分支的信息。例如：

> 已找到这家公司的本期利润表、累计列和上期申报。请一次确认：本次主体、所属期和表版是否正确；是否为居民企业查账征收、独立纳税，是否涉及境外、汇总纳税总分机构或本工具未覆盖的特殊业务。尚未找到的专项是【按真实缺口列出】；没有发生的请明确确认，发生的请给对应底稿。税收待遇准备按普通25%、小型微利优惠，还是已有专项减免底稿处理？若采用小微优惠，请提供已确认的季度平均人数、资产资料或季度期初／期末快照。我据资料计算，不需要你先算应缴税额。

这段是问法示例，不是固定问卷。已从资料取得并有确认依据的内容不要再问。把程序的 `facts.tax_lines.24` 翻成“以前年度可弥补亏损及其底稿”，不要让会计逐格报税额或勾选看不懂的字段名。

| 首轮需要核实的内容 | 用户提供什么 | Agent 做什么 |
| --- | --- | --- |
| 主体、期间、准则、完整性 | 现有资料及对缺口的确认 | 自动识别、核对可见冲突和来源窗口 |
| 业务范围 | 企业类型、查账／实际利润额、独立纳税、境外及特殊业务事实 | 决定能否进入当前 EIT 分支 |
| 专项与税会差异 | 现成调整底稿、优惠收入／加计扣除、折旧、亏损和抵免资料 | 从真实资料提取专项；不会从普通利润表猜出这些金额 |
| 已预缴税款 | 上期申报、税款记录或已确认台账 | 取得本年累计已预缴额，计算当期应补或年度应退 |
| 非零投资收益 | 具体事项名称、金额及明细来源 | 核对事项合计与主行，生成9.1等事项视图 |
| 小微待遇 | 行业事实、正确口径的季度平均人数和资产，或季度快照 | 计算平均数、核对门槛并计算减免；不要求会计先算优惠税额 |

普通25%也要明确选择，不能默认替用户放弃优惠。小微资格中的“资产”输入为元，不把5000万元阈值写成5000元。预缴使用截至本期的季度平均口径，年度使用全年季度平均口径；人数包括劳务派遣用工。具体 `facts` 结构见下节及 [EIT 参考](eit-2025-adapter.md)。

无事项必须得到逐项明确确认。缺失整行财务项目时，只有规则模块认可的 `confirmed_zero_financial_fields` 和真实确认记录才能给零；已经存在的源格为空、横线、公式、冲突或重复时，不能用“无事项”绕过源错误。财务软件若把研发费用并入管理费用，应先确认拆分口径，避免重复扣除。

## 4. case 契约：由 Agent 建立

`case.json` 相对路径以任务文件所在目录为基准；`income_statement.path` 相对 `input_dir`，且必须位于该目录内。输出目录必须不存在。以下展示完整普通预缴分支的数据结构，**null 是待取证的缺口，不能直接运行或替换为虚构 evidence**。已知税额不需要用户写进结构；Agent 从文件取数和确认后填入。

```json
{
  "return_id": "A200000",
  "context": {
    "taxpayer_id": null,
    "taxpayer_name": null,
    "period_start": null,
    "period_end": null,
    "period_kind": null,
    "form_name": "中华人民共和国企业所得税月（季）度预缴纳税申报表（A类）",
    "form_version": "国家税务总局公告2025年第17号",
    "region": null,
    "currency": "CNY",
    "amount_unit": "元"
  },
  "input_dir": null,
  "income_statement": {
    "path": null,
    "sheet": null,
    "accounting_standard": null,
    "presentation": null,
    "source_period_start": null,
    "source_period_end": null,
    "period_evidence": null,
    "identity_anchors": []
  },
  "facts": {
    "scope": {
      "resident_enterprise": {"value": null, "evidence": null},
      "collection_method": {"value": null, "evidence": null},
      "independent_taxpayer": {"value": null, "evidence": null},
      "no_foreign_business": {"value": null, "evidence": null},
      "no_special_business": {"value": null, "evidence": null}
    },
    "tax_lines": {
      "19": {"value": null, "evidence": null},
      "20": {"value": null, "evidence": null},
      "21": {"value": null, "evidence": null},
      "22": {"value": null, "evidence": null},
      "23": {"value": null, "evidence": null},
      "24": {"value": null, "evidence": null},
      "29": {"value": null, "evidence": null},
      "30": {"value": null, "evidence": null},
      "31": {"value": null, "evidence": null},
      "39": {"value": null, "evidence": null}
    },
    "confirmed_zero_tax_lines": [],
    "confirmed_zero_financial_fields": [],
    "confirmation_evidence": null,
    "tax_treatment": null,
    "tax_treatment_evidence": null
  },
  "confirmation_evidence": null
}
```

| 字段 | 可接受内容及限制 |
| --- | --- |
| `return_id` | `A200000` 或 `A100000` |
| `context.period_kind` | 预缴 `month`／`quarter`；年度 `year`，均为已结束完整自然期间 |
| `context.form_name/form_version` | 从对应 `assets/<return_id>.json` 读取实际值；任务、规则和模板版本须相同。年度为2025年第1号版 |
| `accounting_standard` | `asbe`／`small_asbe`，有资料或确认依据 |
| `presentation` | `software_export`／`tax_filing`，表示真实来源语境，不是为了选列随意设置的开关 |
| 可选 `income_statement.amount_unit/unit_evidence` | 只有明确单位和真实依据时才写这两个键；未知时省略，不能写 null 后绕过校验 |
| `identity_anchors` | 例如实际税号在某原格时写 `{"cell":"真实坐标","expected_context_key":"taxpayer_id"}`；坐标必须来自扫描，不复制示例坐标 |
| `scope` | 当前分支要求居民企业 true、实际利润额 `actual_profit`、独立纳税 true、无境外及未适配特殊业务 true；每项均须真实 evidence |
| `tax_lines` | 必需专项金额为元的十进制字符串及来源；对应无事项已明确确认时，删除该项空对象，把字符串行号加入零事项清单并保留确认记录 |
| `confirmed_zero_financial_fields` | 仅填写18个财务键中明确无该事项且源表确实没有该主行的键；不默认把全部缺行填零 |
| `tax_treatment` | `standard_25`、`small_low_profit` 或 `confirmed_relief`；均有真实 `tax_treatment_evidence` |
| `confirmation_evidence` | 顶层记录本次身份、期间、资料与口径确认；facts 内记录逐项无事项等事实确认。不能用“已确认”三个字代替真实回复／材料位置 |

`financials` 不属于用户输入；driver 从真实累计列建立并换算。不要在 case 里另塞一份手工算好的18项金额来跳过原表。

年度 case 改为 `A100000`、完整自然年和对应模板名／2025年第1号版本。年度必需专项行为19、20、21、22、23、25、26、27、32、34、35、37、42、43、44；不要沿用预缴的行号含义。预缴已缴额是30行，年度是37行。年度地方分享调整42、查补43、特别调整44的正负及来源按 [EIT 契约](eit-2025-adapter.md) 核实。

### 条件性事实

- **非零预缴投资收益**：增加 `facts.investment_income_items`，每项为 `{"name":真实事项名,"value":元金额字符串,"evidence":真实明细来源}`。合计须与利润表投资收益相符。driver 只为规则已生成的 `9.正整数` 行建立私有附页；其他未知行仍阻断。
- **小型微利待遇**：增加 `facts.small_low_profit`，包括 `non_restricted_industry`、`average_basis`、`employee_average`、`asset_average_yuan`，每项为 `{value,evidence}`。预缴 basis 为 `ytd_quarter_average`，年度为 `annual_quarter_average`。也可提供 `quarters`：逐季 `quarter` 加期初／末人数和资产四项 `{value,evidence}`，让规则算平均数，不让用户先手算减免税额。
- **已核实的其他减免**：选择 `confirmed_relief`，提供预缴28／年度31行具体减免底稿及资格依据。当前模块核对金额上限，不替代对应优惠的专项资格鉴定。

## 5. 什么时候停止

| 情形 | 交付动作 |
| --- | --- |
| 缺主体、窗口、完整性或税务专项证据 | 给合并缺口清单，指出可从什么资料补齐；不虚构 evidence，不补零 |
| 累计歧义、重复主行／表头、源公式、空格／横线金额 | 定位原文件、工作表和单元格；先解决源问题，不能从其他列随意取代 |
| 单位、币种、名称、税号、可见期间或锚点冲突 | 阻断，保留冲突双方和最少修改动作 |
| 利润表项目重算与营业利润／利润总额不符 | 查符号、费用重复、拆分或源表口径，不能改公式、丢行或放宽容差凑平 |
| 年度计税基础第28行公式差额为负 | 当前规则返回 `ANNUAL_NEGATIVE_BASE_DISPLAY_UNVERIFIED`，保留原差并停止，不以旧表规则代替核验 |
| 预缴实际利润为负 | 按当前已适配规则保留实际利润负号、应纳所得税不产生负数；已有多缴不直接输出本期负退款 |
| 年度正计税基础，但预缴税款多于全年税额 | 支持最终负退税额，仍为待会计复核稿 |
| 新设／中途终止、清算、境外、汇总纳税或其他未适配分支 | 说明具体未覆盖事项，保留资料，不套用独立普通企业分支 |
| 规则 BLOCKED，或运行后源哈希变化 | 不交付有效 filled.xlsx；只交计算／问题报告。已生成后发现变化时填写稿会被撤下 |

## 6. 成功后交什么

driver 的 `workflow_result.json` 列出本轮实际存在的输出；`workflow_report.html` 提供文件链接。不要仅看退出码，应打开主表和报告回读。

| 文件 | 含义 |
| --- | --- |
| `fill-result/filled.xlsx` | 对应 A200000／A100000 的完整主表工作底稿；固定比率行与模板核对保留，其余规则行均有目标 |
| `fill-result/review.html`、`fill-result/audit.json` | 填写引擎的逐字段来源、目标、哈希及复核状态 |
| `calculation.json` | 所得税逐行值、计算公式、专项缺口、规则版本及政策链接 |
| `lineage.json` | 每行到财报原始文件哈希、sheet/cell、原值、单位、负号转换、确认事实和政策依据的追溯 |
| `scan.json` | 原始结构、选列候选、歧义与发票双粒度告警；不能把未使用的候选说成已经核实 |
| `job.json`、`return_values.csv`、`context_values.csv` | Agent 自动建立的填写任务及有来源的派生值；用户不用维护这些中间文件 |
| 条件性 `resolved_template.xlsx/.json` | 非零投资收益时的私有模板副本，附“投资收益明细”事项视图；原基础模板不改，变换哈希记录在 lineage |

这些表是**按真实官方行次重建的工作底稿，不是下载的官方原 Excel，也未证明可直接导入税局**。投资收益附页是事项视图，不冒称税局正式附表布局。状态始终为待会计复核；没有提交申报、缴款或生成受理回执。

复用时保留来源识别方式和已确认口径，但每期重新绑定真实文件、哈希、主体和窗口。金额、已缴税款、季度快照、年度专项及失效资格不能沿用旧值。不要通过只改日期把上一期金额带到本期。

## 7. 增值税路由

小规模增值税套表、门槛、发票／未开票完整性、历史累计、附加及分位路径见 [2026 VAT 适配](vat-2026-adapter.md)。只有在相应 driver、模板及本轮完整套表结果已实测通过时，才能按该链路交付；本参考中的 EIT 命令不适用于 VAT，不能把“已查到规则”说成“业务 driver 已跑通”。
