# Stage 2 C1.1 检查点（2026-09-11）

## 已完成内容与边界

只补强 C1 Parser 的风险元数据：区分缺少 NO 与结构异常，并记录同一输入单元格内部重复号码。候选提取、排除、长度、顺序、按 Digits 去重及首次原文位置均保持 C1 语义；未重写 Parser。

本模块只提供风险来源，不决定是否人工审核、风险处置或自动接受。重复仅指本次输入中已成功解析的 Digits 重复，不表示不同 B 行引用同一 A 记录。没有 C2 匹配、金额、颜色、分组、baseline 或 UI。

## 接口与 Flags

函数和结果类型保持不变：

```vb
VATStage2ParseBInvoices(ByVal rawText As String) As VATS2ParseResult
VATStage2ParserFlagNames(ByVal flags As Long) As String
```

`VATS2ParseResult.Flags` 及 `References(i).Flags` 仍为 Long 位掩码。记录级 Flags 传播给所有引用，因此引用上的重复标志表示该输入存在重复，并非逐引用重复计数。

| 常量 | 数值 | 名称及触发条件 |
|---|---:|---|
| `VATS2_FLAG_NONE` | 0 | 无本模块定义的风险，不代表已确认发票真实性 |
| `VATS2_NO_MARKER_MISSING` | 4 | `NO_MARKER_MISSING`：无 NO 且已成功提取末尾候选；本身不表示结构异常 |
| `VATS2_STRUCTURE_REVIEW` | 1 | `STRUCTURE_REVIEW`：多个 NO、括号错误、空分隔、非法拼接、小数/单横线、未知尾部等结构问题 |
| `VATS2_NUMERIC_TRAILER_REVIEW` | 2 | `NUMERIC_TRAILER_REVIEW`：保持原 `--` 纯数字尾注规则，尾注不进入候选 |
| `VATS2_DUPLICATE_REF_REVIEW` | 8 | `DUPLICATE_REF_REVIEW`：同一输入内已接受的 Digits 再次出现；候选只保留首次引用 |

四种标志可以组合。无 NO 且解析失败时，只报告实际结构问题，不添加 NO_MARKER_MISSING；无 NO 有候选且另有异常时，两者同时存在。括号或尾注里已被排除的数字不算重复引用。

名称函数固定按 `NO_MARKER_MISSING|STRUCTURE_REVIEW|NUMERIC_TRAILER_REVIEW|DUPLICATE_REF_REVIEW` 的顺序输出实际存在的项，无风险为空字符串。

旧常量 `VATS2_PARSE_REVIEW` 保留为组合掩码 **5**（STRUCTURE_REVIEW Or NO_MARKER_MISSING），不是独立风险位。旧的 `(Flags And VATS2_PARSE_REVIEW) <> 0` 宽泛检查仍可使用；新调用方应分别检查具体风险位，不使用 `Flags = VATS2_PARSE_REVIEW` 判断单一来源，也不沿用旧数值 1 代表所有旧含义。名称函数不再输出宽泛的 PARSE_REVIEW。

```vb
Dim parsed As VATS2ParseResult
parsed = VATStage2ParseBInvoices("NO.12345678/12345678")
'ReferenceCount=1；Digits=12345678；StartIndex=4；Flags=8。
Debug.Print VATStage2ParserFlagNames(parsed.Flags) 'DUPLICATE_REF_REVIEW
```

## 修改/新增文件

- 修改 `src/modVATStage2Parser.bas`：风险位、名称转换、风险来源赋值和重复提示；提取规则未改。
- 修改 `tests/modVATStage2ParserTests.bas`：原 60 项仅适配风险期望及名称；增加 10 项直接相关回归。
- 更新 `tests/stage2-parser-test-results.txt`、`tests/excel-test-results.txt`。
- 更新 `STAGE2_SPEC.md` 的 C1.1 元数据补充、`STAGE2_CHECKPOINT_C1.md` 的历史入口提示。
- 新增本文件 `STAGE2_CHECKPOINT_C1_1.md`。

Stage 1 两个业务源码、原 89 项测试源码、`tools/build.ps1` 和四个 release 文件的 SHA-256 均与本轮开始前一致。没有提交 commit，没有接入 release。

## 测试结果

- Parser：**70/70 cases PASS**。原 59 个普通用例的输入及候选期望逐项核对未变，原接口用例保留；根据新接口调整精确 Flags 期望，没有放宽候选、位置、上下文或去重断言。
- 新增 10 项涵盖：两次/多次重复、不同字母前缀下相同 Digits、无 NO＋重复、空分隔＋重复、四标志并存、排除的括号/尾注不误报重复、无 NO 失败不误报缺标记、首次位置保留及兼容掩码。
- 原用例继续精确验证：正常显式 NO 无多余风险；无 NO 正常候选只有 NO_MARKER_MISSING；多个 NO、括号异常、空分隔等为 STRUCTURE_REVIEW；纯数字尾注及与缺少 NO 的组合。
- Stage 1：**89/89 assertions PASS**，原测试源码未动。

复现命令（使用现有、已允许 VBA 工程访问的 Excel 开发环境）：

```powershell
./tools/build.ps1 -ParserTestOnly
./tools/build.ps1 -TestOnly
```

两次均已真实运行并更新结果文件；未改变安全设置或执行策略。

## 限制与接手入口

候选真实性、无标记数字是否为日期或业务编码等限制继续沿用 C1；本次只让调用方看见风险来源，不作处置决策。不提供重复次数或所有重复位置，仅按本轮要求保留首次引用并设置记录级重复标志。无已知阻塞问题。

后续先读本文件、`STAGE2_SPEC.md` 和 `STAGE2_CHECKPOINT_C1.md`，再读 Parser、独立测试及结果报告。C1.1 已完成并停止，C2 仍需用户另行授权。
