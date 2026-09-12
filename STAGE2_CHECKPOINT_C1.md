# Stage 2 C1 检查点（2026-09-11）

> C1 历史检查点。风险元数据已在 C1.1 细分，最新 Flags 和回归结果请先读 `STAGE2_CHECKPOINT_C1_1.md`；下文 60 项及旧 PARSE_REVIEW 描述保留为历史。

## C1 实际完成内容

独立 VBA 字符串 Parser 已完成。支持用户提供的 12 个真实样例、NO 的中英文标点/括号、无 NO 的末尾号码组、`/` / `，` / `、` 混合分隔、英文字母前缀、普通中英文括号内容排除及两类 `--` 尾注规则。候选保序、按 Digits 去重、保留前导零及原文定位信息。

代码只读取函数参数并返回结构，不访问 Excel 对象、业务工作簿或文件，没有金额、颜色、匹配、状态持久化或 UI 功能。C1 未接入正式 release；未开始 C2。

网络中断后已核对磁盘：源码与测试完整、两份结果报告已生成；没有未完成代码。本次续跑只补齐本检查点，没有重写或重复测试已通过的代码。

## 修改/新增文件

新增：

- `STAGE2_SPEC.md`：以用户本轮 C1 指令及后续确认的结构解析规则为准。
- `src/modVATStage2Parser.bas`：纯字符串解析、结果类型、Flags 常量和名称转换。
- `tests/modVATStage2ParserTests.bas`：独立确定性测试。
- `tests/stage2-parser-test-results.txt`：原生 Excel 中运行 Parser 测试的结果。
- `STAGE2_CHECKPOINT_C1.md`：本检查点。

修改：

- `tools/build.ps1`：新增 `-ParserTestOnly` 分支，仅在临时宿主加载 Parser 和独立测试；原非限定 `Run("VATCheck_SelfTest")` 保留。
- `tests/excel-test-results.txt`：重新运行原 89 项后的报告。

Stage 1 的 `src/modVATCheck.bas`、`src/frmVATCheck.vba`、`tests/modVATTests.bas` 及四个 release 文件均与 C1 开始前 SHA-256 一致，旧测试没有删除、修改或放宽。未提交 commit。

## Parser 接口

```vb
Public Function VATStage2ParseBInvoices(ByVal rawText As String) As VATS2ParseResult
Public Function VATStage2ParserFlagNames(ByVal flags As Long) As String
```

结果字段：

| VBA 字段 | 含义 |
|---|---|
| `OriginalText` | 原始输入，即需求中的 RawText，逐字保留 |
| `ReferenceCount` | 去重后引用数量 |
| `References()` | 需求中的 InvoiceRefs[]，一基数组，仅使用 1～ReferenceCount |
| `Flags` | 记录级风险位集合，允许多个标志同时存在 |

每个 `VATS2InvoiceReference` 包含：`Digits As String`、`Length As Long`、`RawFragment As String`、`StartIndex As Long`、`Context As String`、`Flags As Long`。

- `Digits` 全程不做数值转换，保留前导零；英文字母前缀保留在 RawFragment 中，不进入 Digits。
- `StartIndex` 为原字符串中的一基位置，遵循 VBA 字符串索引；`Mid$(OriginalText, StartIndex, Len(RawFragment))` 等于 RawFragment。
- Context 保留片段及前后最多各 16 个 VBA 字符位置的上下文。
- ReferenceCount 为 0 时 References 数组未分配；调用方按计数遍历，不直接对空数组调用 UBound。
- 相同 Digits 保留第一次出现的片段和位置；每个引用携带记录级 Flags。

```vb
Dim parsed As VATS2ParseResult, i As Long
parsed = VATStage2ParseBInvoices("供应商（NO.）00123456/98765432")
For i = 1 To parsed.ReferenceCount
    Debug.Print parsed.References(i).Digits, parsed.References(i).Length
Next i
Debug.Print VATStage2ParserFlagNames(parsed.Flags)
```

## Flags

| 常量 | 值 | 含义 |
|---|---:|---|
| `VATS2_FLAG_NONE` | 0 | 未发现本 Parser 定义的结构风险；不代表号码已验证 |
| `VATS2_PARSE_REVIEW` | 1 | 无 NO 或结构不确定，需要复核 |
| `VATS2_NUMERIC_TRAILER_REVIEW` | 2 | `--` 后为纯数字；尾注不产生候选，前面正常号码继续解析 |

Flags 用位掩码表达需求中的风险集合，可用 `And` 判断；名称函数按固定顺序返回 `PARSE_REVIEW`、`NUMERIC_TRAILER_REVIEW` 或以 `|` 连接的两者，无风险返回空字符串。

按用户最后确认：**取消 7–20 / 18–20 位硬门槛**；长度仅记录。无 NO 的末尾候选统一带 PARSE_REVIEW，不因短而直接丢弃，也不因长而自动认为可靠。

## Parser 测试数量及结果

**60/60 cases PASS**，报告为 `tests/stage2-parser-test-results.txt`，2026-09-11 10:07:42 生成。

包含全部 12 个真实样例及空值、无数字、括号/尾注排除、NO 标点、字母前缀、三类分隔符及混合、前导零、去重保序、无 NO 风险、取消长度门槛、数字不得跨括号/空格拼接、小数/单横线/全角数字、未配对括号、多个 NO 等。每个普通用例还同时检查原文、数量、Digits、长度、原片段位置、上下文、风险及去重；另有接口确定性检查。这里 60 指测试用例数，不把每个字段检查重复计数。

在项目根目录、Excel 已允许 VBA 工程访问的环境中复现：

```powershell
./tools/build.ps1 -ParserTestOnly
```

## 原 89 项回归结果

**89/89 assertions PASS**，报告为 `tests/excel-test-results.txt`，2026-09-11 10:10:18 生成。原测试源码哈希未变。

```powershell
./tools/build.ps1 -TestOnly
```

两次运行都使用独立临时测试宿主，没有重建 release；本轮没有改变执行策略或 VBA 安全设置。

## 已知限制

- 只做候选解析，不能确认号码存在、票种或真实性，也不能仅靠数字区分日期、业务编码和发票号码；无 NO 候选必须保留 PARSE_REVIEW，不应直接当作确认结果。
- 无 NO 仅考虑末尾号码组；若此前仍有未排除的数字，则不擅自选择某组。号码后的未括起中文说明不自动当尾注。
- 多个独立 NO 标记不擅自选取，输出 PARSE_REVIEW。未知/缺损结构可以保留已成功解析的项，但会给整条记录及所有引用加风险。
- 只接受 ASCII 数字，不自动转换全角数字，不将空格、小数点或单横线两侧数字拼接；分隔符限于本轮约定的三类。
- 参数必须是 String；上游从 Excel 读取错误值、空值等的处理不在 C1 内。References 是 VBA UDT 数组，尚未提供对外序列化或 UI。
- 无已知阻塞问题。上述限制属于当前保守解析边界。

## C2 接手需要读取的文件

先读本检查点及 `STAGE2_SPEC.md`，再读 `src/modVATStage2Parser.bas`、`tests/modVATStage2ParserTests.bas`、两份测试结果和 `tools/build.ps1` 的 `-ParserTestOnly` 分支。C2 具体需求需要用户另行授权；本轮没有设计或实现 C2。
