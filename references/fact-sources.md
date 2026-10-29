# 专项与已缴事实的原格绑定

财报项目和收入明细由对应适配器读取；已预缴、纳税调整、资格人数及资产等其他事实也应尽量直接绑定源格。用户不需要抄金额到配置，Agent 建立绑定后程序自动读值。

```json
{
  "source": {
    "path": "上期预缴记录.csv", "sheet": "CSV", "cell": "B2",
    "kind": "money", "amount_unit": "元",
    "anchor": {"cell": "A2", "text": "本年累计已预缴所得税额"},
    "identity_anchors": [{"cell": "E1", "context_key": "taxpayer_id"}]
  }
}
```

此对象可放在 `facts.tax_lines["30"]` 等规则输入值的位置；年度已预缴对应行37，VAT预缴和附加已缴按其 profile 契约放置。`value` 可省略，程序从原格取得并补充实际来源说明。若 Agent 同时写了 `value`，必须与源格归一后的金额一致，不能随意覆盖错误值。

`kind` 支持 `money`、`number`、`text`。金额必须明确元／千元／万元并归一到元；人数等用 number，不进行金额单位换算。资产资格字段仍是金额，不能把万元当元。支持会计括号负数，转换有记录。文本不是可以执行的公式。

路径限于本次资料目录，不读取别的客户目录。可给 SHA-256，文件变更会阻断；即使未给，也记录实际读值时的哈希并在返回和工作流结束时复查。空值、公式、错误、外币、单位矛盾、来源标签或身份锚点冲突均不能继续填写。

没有文件来源、仅由用户明确提供的 `value + evidence` 可以作为事实声明；审计标记 `explicit_declaration`、`file_verified=false`。有实际来源格的记录标记 `machine_read_source`，同时保存原格、单位、换算和锚点。两者不得混称“全部机器核验”。

金额单位未在源表显示时，绑定中的显式单位是声明，审计会标记 `unit_declaration_only`；不能假装从表头识别出来。来源原格读到了数字，也不等于经济实质、税收资格和完整性已审计。

Python API：`resolve_sources(facts, input_dir, context)` 返回已解析副本、来源文件清单和逐事实 trace；异常使用 `tax_fill.Blocked`。工作流必须把来源清单纳入最终哈希检查，并把 trace 写入 lineage。
