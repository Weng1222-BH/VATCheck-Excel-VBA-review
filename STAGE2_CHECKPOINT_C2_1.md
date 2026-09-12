# Stage 2 C2.1 检查点（2026-09-11）

## Matcher 接口

已完成独立、确定性的单 B 引用候选 Matcher，无 Excel/Parser 依赖，无文件读写。

```vb
VATStage2MatchReference(ByVal referenceDigits As String, _
    ByRef aInvoiceDigits() As String, ByVal aCount As Long) As VATS2ReferenceMatchResult
```

- 输入号码全程 String，不 Trim、不转数值、不改前导零；A 由调用层保证完整且为 ASCII 数字。
- A 为一维 String 数组，从实际 LBound 开始读取 aCount 个元素；支持非 1 下界。索引运算不改变号码。
- 结果：ReferenceDigits、ReferenceLength、MatchKind、CandidateCount、Candidates()、Flags。
- 每个 `VATS2MatchCandidate` 含 AIndex（输入数组实际下标）、InvoiceDigits；结果数组从 1 开始，零候选未分配，调用方用 CandidateCount 遍历。
- 先以二进制字符串比较收集所有 exact；有任意 exact 就返回，不执行 suffix。否则按 B 的实际长度匹配尾号，不设最低长度。
- 保留 A 输入顺序和重复项各自的索引，不作去重、选首项或消歧。

## MatchKind 与 Flags

`VATS2MatchKind`：

| 枚举 | 值 | 含义 |
|---|---:|---|
| VATS2_NOT_FOUND | 0 | 无候选 |
| VATS2_EXACT_UNIQUE | 1 | 一个 exact |
| VATS2_EXACT_MULTIPLE | 2 | 多个 exact |
| VATS2_SUFFIX_UNIQUE | 3 | 一个 suffix |
| VATS2_SUFFIX_MULTIPLE | 4 | 多个 suffix |
| VATS2_INVALID_REFERENCE | 5 | B 为空或非纯 ASCII 数字 |
| VATS2_INVALID_A_RANGE | 6 | aCount 为负、越界，或未分配数组且 aCount>0 |

B 合法性先检查；合法 B 加 aCount=0 返回 NOT_FOUND，允许 A 未分配。

独立 `VATS2MatcherFlags`：`VATS2_MATCH_FLAG_NONE=0`，`VATS2_SHORT_SUFFIX_REVIEW=1`。

当非空 A 没有 exact、实际执行 suffix 搜索且 B 长度 <=6 时设置 SHORT_SUFFIX_REVIEW；即使该搜索最终 NOT_FOUND 也记录路径提示。空 A 不执行 suffix，不设标志；所有 exact 不设标志。该位掩码与 Parser Flags 分属不同接口，不混用。本阶段不决定人工审核、豁免或冻结。

## 修改/新增文件

- 新增 `src/modVATStage2Matcher.bas`。
- 新增 `tests/modVATStage2MatcherTests.bas`。
- 新增 `tests/stage2-matcher-test-results.txt`。
- 修改 `tools/build.ps1`：新增独立 `-MatcherTestOnly` 分支及发布文件存在检查豁免，不重写其他流程。
- 在 `STAGE2_SPEC.md` 末尾追加 C2.1，前文 C1/C1.1 保持原文。
- 新增本检查点。
- 重新运行并写出 `tests/stage2-parser-test-results.txt`、`tests/excel-test-results.txt`（后者含本次运行耗时）。

Parser 及其测试、Stage 1 业务源码及其测试均未修改。四个 release 文件未重建或修改。未提交 commit。

## 测试结果

本轮已在本机真实 Excel 中依次执行，均 exit 0：

```powershell
./tools/build.ps1 -MatcherTestOnly
./tools/build.ps1 -ParserTestOnly
./tools/build.ps1 -TestOnly
```

- Matcher：**38/38 cases PASS**。每例同时验证 MatchKind、Flags、数量、号码、顺序、原 A 索引、引用原文/长度、数组边界，以及重复调用结果相同、A 输入不变。
- 覆盖所有指定场景：exact 唯一/多个及严格优先、suffix 唯一/多个、1/5/6/7/9 位边界、未找到、B 比 A 长、前导零、空 A。补充部分有效数组、0/正/负下界、重复 A、非法 B、非法计数及长字符串不丢精度。
- Parser：**70/70 cases PASS**，原测试未改。
- Stage 1：**89/89 assertions PASS**，原测试未改。
- 没有更改 Excel 安全设置或执行策略，没有重建 release。

## 已知限制与 C2.2 接手入口

无已知阻塞问题。Matcher 依赖上游提供合法 A 完整号码，不检查 A 全表质量，不负责实际发票真实性。本阶段线性两遍扫描单个引用，不建立全表索引；不提供多引用调度、跨 B 冲突、金额、颜色、分组、baseline、UI 或 release 集成。

后续获得授权后先读取本检查点、`STAGE2_SPEC.md`、`src/modVATStage2Matcher.bas`、`tests/modVATStage2MatcherTests.bas` 和对应结果；需要 Parser 接口时读取 `STAGE2_CHECKPOINT_C1_1.md`。测试接线见 `tools/build.ps1`。C2.2 的重复 A 分类、跨引用关系等尚未实现，不能将本次 MULTIPLE 擅自当成已消歧结果。

C2.1 已完成并停止，未开始 C2.2。
