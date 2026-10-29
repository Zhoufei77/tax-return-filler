# 数据与任务契约 v1

供 Agent 生成任务文件；不要让终端用户编辑 JSON。示例生成器 `scripts/demo.py` 仅生成标记清楚的合成数据，不接收客户资料。

## 扫描与绑定

`scan <目录> --output <新JSON>` 递归扫描指定目录中的 CSV/XLSX，返回文件名、SHA-256、工作表及非空单元格。CSV 的逻辑工作表名固定为 `CSV`，坐标仍为 A1。支持 UTF-8（含 BOM）及 GB18030；无效字符、NUL、过大的文件和范围会被拒绝。

扫描值只用于建立映射，不是税务完整性证明。公式会标记为不可直接取数。扫描 JSON 可能含客户财务值，保存在客户工作区，不进入分发包或公共日志。

文件上限 25 MiB、工作簿解压合计 150 MiB；扫描最多 100 个资料文件，工作簿／CSV 最多 200000 个单元格范围及 50000 行。超限时明确说明，需要用户缩小为本次相关资料，不自行截断后声称完整。

## 顶层

```json
{
  "schema_version": 1,
  "context": {
    "taxpayer_id": "源资料中的真实文本税号",
    "period_start": "2026-07-01",
    "period_end": "2026-09-30",
    "period_kind": "quarter",
    "form_name": "用户实际模板名称",
    "form_version": "用户实际表版标识",
    "region": "用户实际地区",
    "currency": "CNY",
    "amount_unit": "元"
  },
  "confirmations": {"context_confirmed": true, "materials_complete": true},
  "sources": [],
  "template": {
    "path": "模板.xlsx", "sha256": "扫描取得的64位十六进制哈希",
    "context_bindings": {
      "taxpayer_id": {"sheet":"申报底稿","cell":"B2"},
      "period_start": {"sheet":"申报底稿","cell":"D2"},
      "period_end": {"sheet":"申报底稿","cell":"D3"}
    }
  },
  "fields": [],
  "mandatory_targets": [],
  "checks": [],
  "prefill_comparisons": []
}
```

示例中的说明字符串不是可执行值。相对路径以 job 文件所在目录为基准；绝对路径也可用。任务文件应独立于技能目录。输出目录必须尚不存在。`period_kind` 仅 `month`、`quarter`、`year`，当前版本要求已结束的完整自然月、季、年；短期设立、注销、清算、更正等其他期间不支持自动执行。

`context_confirmed` 表示用户明确确认了主体、期间及实际表版。`materials_complete` 表示本次约定填表范围所需资料已确认齐备，不能扩大为企业所有涉税材料已经审计。两者均为用户声明，程序不能证明声明真假。

`template.context_bindings` 必须绑定真实模板的税号与期间。有两种互斥模式：`taxpayer_id + period_start + period_end` 三格，或 `taxpayer_id + period` 两格。不同角色的格子不得重复。单格 period 例：`{"sheet":"申报底稿","cell":"D2","format":"iso_range"}`，仅支持 `iso_range` 的 `2026-07-01 至 2026-09-30`，及 `cn_range` 的 `2026年7月1日至2026年9月30日`，预期文字严格从 context 生成，不能自行给 expected。原生 date/datetime 起止格按 ISO 日期核对，不猜测本地字符串日期。

原值非空时须严格等于对应 context；原值空时，必须有对应 text 字段从来源填写准确值，并纳入 mandatory_targets。完全没有身份区或日期格式不支持时停止自动填入，说明具体限制；不能伪造隐藏身份格通过检查。历史模板中其他主体或其他期间值会阻断，请使用空白或当期模板。表版以实际模板哈希及 context 的明确声明绑定。

## sources

每个来源对象包含：

| 字段 | 含义 |
| --- | --- |
| `id` | 任务内唯一来源名 |
| `path`、`sha256` | 实际源文件及当前内容哈希；同一文件／内容重复注册会阻断 |
| `context` | `taxpayer_id,period_start,period_end,currency,amount_unit`；币种仅 CNY，单位仅元／千元／万元 |
| `metadata_confirmed` | 用户对来源主体、数据窗口、单位的明确确认；不可自动猜为 true |
| `identity_anchors` | 可选的真实单元格核对：`[{"sheet":"CSV","cell":"E1","expected":"实际税号"}]`；存在就必须匹配 |

源窗口可以比本期更宽。金额 `ref` 默认必须与字段所选 current/ytd 的窗口完全一致；宽窗口明细使用 `table_sum` 按日期筛选。一张表同时放本期与累计列时，使用下面的 `ref.period_scope` 明确列口径，保留原文件坐标；不要求用户拆表，也不能把原始金额复制为 job 常量绕过来源。

没有单元格锚点时，审计记录明确为 `human_declaration_only`，不能说机器已核对了源文件主体和期间。`identity_anchors` 应尽量覆盖真实税号及期间，不能拿文件名代替。

## fields

每项必须有唯一 `id`、`label`、`required:true`、`target`、`expression`。`kind` 为 `money`（默认）或 `text`；`basis` 为 `current`（默认）或 `ytd`。ytd 从该年 1 月 1 日至本期结束。`scope` 可附与 context 相同的键值，再次锁定口径。

目标示例：`{"sheet":"申报底稿","cell":"B6","anchor":{"cell":"A6","text":"已确认销售额"}}`。锚点必须是同表相邻标签，距离不超过 4 格；不可就是目标格。目标必须为空或等于即将写入值，若用户授权更新已有历史值，需提供准确 `expected_existing`；不能绕过公式保护。

所有实际写入目标与 `mandatory_targets:[{"sheet":"申报底稿","cell":"B6"}]` 必须一一相符；任一缺失、重复或范围外写入都阻断。**完整清单仍由 Agent 按用户实际范围识别，程序无法知道一张陌生税表所有业务必填项。**

金额以 Decimal 在人民币元中运算，目标按 context 单位舍入两位，使用 ROUND_HALF_UP。文本字段仅支持来源文本或日期的 `ref`，数字型税号可能已经失真，会被拒绝。金额写入 Excel 数字，不把金额当文本。

## expression

| 操作 | 结构与限制 |
| --- | --- |
| `ref` | `{"op":"ref","source":"来源id","sheet":"CSV","cell":"B2"}`；可有明确 `amount_unit` 和 `period_scope`。空、公式、错误、模糊数字不接受 |
| `field` | `{"op":"field","id":"已计算金额字段id","precision":"display"}`；默认 display 读取最终写入金额，raw 读取该字段舍入前的元值。表内已列示值的勾稽选 display；要求先合并原值再舍入时显式选 raw。不能前向引用或循环 |
| `const` | `{"op":"const","value":"0.01","reason":"具体已核实的无量纲系数及来源"}`；可有明确金额单位。禁止拿常量替代未知原始金额 |
| 算术 | `{"op":"add|sum|sub|mul|div","args":[表达式,表达式]}`；sub/div 恰好两项。mul/div 用于已确认的无量纲系数或明确换算，不能把两个金额相乘。无百分号自动解析、无 eval、无税法推断 |
| `table_sum` | 见下例。逐行记录日期、主键、原值、单位和是否入选；无入选记录不等于零申报 |

同表本期与累计列可在各自 ref 上附 `period_scope:{"period_start":"2026-07-01","period_end":"2026-09-30","anchor":{"cell":"B1","text":"本期金额"}}`。范围必须严格等于字段 current/ytd 所需期间，来源数据窗口必须覆盖，anchor 必须是该源 sheet 真实期间列表头的准确原文。列口径仍须会计确认，不能拿任意无关标签代替；锚点与范围记录在 audit 和 HTML 中。无 period_scope 时继续执行严格窗口匹配。

```json
{
  "op": "table_sum", "source": "ledger", "sheet": "CSV", "header_row": 1,
  "amount_column": "D", "date_column": "A", "taxpayer_column": "B",
  "unique_by": ["C"],
  "headers": {"A":"业务所属日期","B":"纳税人识别号","C":"业务行号","D":"已确认不含税销售额"}
}
```

日期须明确 ISO 日期或 XLSX 日期类型。所有非空行都校验：主体一致、日期处于声明源窗口、主键唯一、金额有效，再判断是否进入 current/ytd。表头、合计行、跨主体或坏行不能被悄悄忽略。多行发票请使用发票号加行号等业务唯一键，不能仅凭票号误报重复。

`table_sum` 可声明 `record_namespace`，默认为 `default`。同一金额表达式内（含通过 field 引用继承的记录），同业务域同主键重复入选会阻断，覆盖不同源文件和同一明细被相加两次。不同结果字段重复展示同一来源不因此报错。业务域用于区分真正不同类型的记录，不得通过改业务域隐藏相同交易。不同业务键描述同一交易仍需 Agent 和会计核对，程序不自动识别语义重复。

真实无业务时用会计确认的零额源单元格 `ref`，不要制造虚构交易行。负数在通用转填中可表达，但红字、冲销和更正的税务所属期需会计确认。

## checks 与预填对比

`checks` 项：`{"id":"名称","basis":"current","left":表达式,"right":表达式,"tolerance":"0.00"}`。容差按目标单位表示，默认零；不匹配阻断。`prefill_comparisons` 项：`{"id":"预填对比","field":"字段id","expression":预填来源表达式,"tolerance":"0.00"}`。默认继承被比字段的 current/ytd 口径；显式 basis 与字段不一致时阻断。不一致时停止填写，保留具体差值供解释。

没有配置的勾稽不会自动发生。表达式只能验证算术与声明口径，不能证明经济实质。不得通过删检查或扩大容差隐藏真实差异。

## 输出

退出码 0：`filled.xlsx`、`audit.json`、`review.html`；退出码 2：报告阻断，**没有填写稿**。失败到无法创建目录时只返回终端错误。状态始终为 `DRAFT_REVIEW_REQUIRED`，结果分别为 `GENERATED_FOR_REVIEW` 或 `BLOCKED`；`filing_performed` 始终 false。

`audit.json` 保存来源路径与哈希、身份证据类型、逐字段取数和运算、明细行、勾稽结果、问题及输出哈希。`review.html` 供人阅读。HTML 不运行客户脚本。JSON 可能包含敏感数据，应按客户要求保管。

模板内非目标公式保留并设置打开后完整重算，但引擎未运行 Excel 公式计算。源文件或模板在扫描后变化会阻断。模板副本只能叫填写稿，不能叫已申报。
