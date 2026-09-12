# Stage 2 C5.2 检查点（2026-09-12）

## 接口与职责

```vb
VATStage2EvaluateEffectiveAmounts(ByRef a As VATS2ASnapshotResult, _
    ByRef b As VATS2BSnapshotResult, _
    ByRef effective As VATS2EffectiveRelationsResult) As VATS2EffectiveAmountResult
```

独立模块 `src/modVATStage2EffectiveAmount.bas`。完整原 AIndex 映射 AmountRaw，保持 Variant，不转换、压缩或按颜色筛选。仅处理 EffectiveBInScope=True 的完整 B 行。输入不修改；不访问Excel或文件，不调度Parser/Matcher，不生成新关系。

先校验所有输入，再检查每个 B 的质量标志。有 RELATION_A_INVALID / RELATION_A_DUPLICATE 则独立返回 RELATION_QUALITY_BLOCKED 并保留所有 RelationIssues；即使金额相等也不会调用 GroupAmount。其它行只调用冻结 VATStage2EvaluateBGroupAmount，完整保留 GroupAmount/Amount 的数学结果、风险和错误。

缺色A是正常金额成员，缺色B是正常完整金额组；不截取成功引用作部分求和。INCOMPLETE/NO_REFERENCES/SAME由冻结层阻断；仅CROSS可诊断比较。没有任何最终审核或风险豁免。

## 返回值

- FullARecordCount、FullBRecordCount；完整1..BCount数组 BInScope、BRowIds、RowStates、GroupAmounts、RelationQualityFlags；0记录时未分配。
- RelationIssueCount/RelationIssues完整复制原 BIndex、ExcelRow、ReferenceIndex、AIndex、AExcelRow、Flags。
- EffectiveBCount、ComparedCount、EqualCount、MismatchCount、NotComparableCount、AmountErrorCount、RelationQualityBlockedCount、InvalidGroupCount。仅统计诊断，不是审核结论。
- 主Status：VATS2_EFFECTIVE_AMOUNT_OK=0、INVALID_INPUT=1、INVALID_CONTRACT=2、GROUP_ERROR=3（后面三个也使用相同前缀）。
- RowStates：VATS2_EA_NOT_COMPARABLE=0、EA_COMPARED=1、EA_AMOUNT_ERROR=2、EA_INVALID_INPUT=3（EA均带VATS2_前缀），VATS2_RELATION_QUALITY_BLOCKED=4、VATS2_EA_OUT_OF_SCOPE=5。
- 前四种行状态和冻结GroupState值对齐；GroupAmounts原样传播，不换算金额。范围外/质量阻断行未执行GroupAmount，必须先看RowStates，不能使用默认嵌套Status判断成功。
- GROUP_INVALID_INPUT完整保留在GroupAmounts，主Status=GROUP_ERROR且InvalidGroupCount递增。它是接口错误，不是AMOUNT_MISMATCH。

## 契约检查

拒绝非OK上游、计数/边界/原索引/行号错误、范围外非空关系、完成B被移出范围、B原文错位、候选AIndex或完整号码与Snapshot不符、关系引用/状态/计数损坏。

AIntegrity结果与A摘要交叉核对；有问题的UNIQUE引用和RelationIssues按BIndex/ReferenceIndex逐项一致，RowIssueFlags必须等于该行所有具体问题的OR。未知flags、缺失、重复、多余、越界的质量元数据不能绕过门控。无关A质量问题保留于调用方Effective结果，不作全局金额阻断。

EffectiveConflicts必须OK。为拒绝遗漏/错位冲突，使用冻结VATStage2ScanBConflicts作契约对照并逐字段比较，随后传入原EffectiveConflicts给GroupAmount；不另写冲突算法，不替换输入。契约错误清空部分结果并保留ErrorSide/Index/ReferenceIndex/Reason。

## 修改文件

新增：
- `src/modVATStage2EffectiveAmount.bas`
- `tests/modVATStage2EffAmountTests.bas`（VBA组件名控制在31字符内）
- `tests/stage2-effective-amount-test-results.txt`
- 本检查点

修改：
- `tools/build.ps1`：仅新增-EffectiveAmountTestOnly及测试豁免接线，保留原UTF-8 BOM和旧分支。
- `STAGE2_SPEC.md`：追加C5.2。
- `STATUS.md`：更新阶段/结果/入口。
- `tests/excel-test-results.txt`：原测试运行更新耗时证据；没有改变旧测试代码或断言。

## 实际原生 Excel 验证

| build开关 | 结果 |
|---|---:|
| -EffectiveAmountTestOnly | 200/200 assertions PASS |
| -EffectiveRelationsTestOnly | 135/135 PASS |
| -FullFallbackTestOnly | 132/132 PASS |
| -CompletedScopeTestOnly | 73/73 PASS |
| -SnapshotTestOnly | 81/81 PASS |
| -GroupAmountTestOnly | 44/44 PASS |
| -AmountTestOnly | 48/48 PASS |
| -BConflictTestOnly | 41/41 PASS |
| -BRowMatchTestOnly | 34/34 PASS |
| -AIntegrityTestOnly | 29/29 PASS |
| -MatcherTestOnly | 38/38 PASS |
| -ParserTestOnly | 70/70 PASS |
| -TestOnly | 89/89 PASS |

合计1014项全部通过。新测试覆盖正/负差额、1A/多A相等与不等、缺色fallback完整组、INCOMPLETE/SAME/CROSS、关系质量门控及组合风险、无关A异常、五类非法A金额和非法B金额、Decimal 0.1+0.2=0.3、空输入、原Index/ExcelRow、逐字段传播、全部输入不变与重复调用确定，以及61项损坏契约拒绝断言。

冻结src、旧测试与release文件执行前后SHA-256一致；未更改Excel安全设置或执行策略，未重建release。开发推送分别使用私有主仓库及独立公开mirror历史；公开副本按既定白名单同步，仅机器环境路径允许脱敏，源码/测试不脱敏。

## 已知限制与后续入口

无已知阻塞。输入须来自同一批快照与关系结果；本层验证结构和关联一致性，不重新做A号码质量全表扫描、Parser/Matcher或批次fingerprint。GROUP_ERROR为冻结GroupAmount意外返回接口错误时的保留路径；合法且完整校验通过的夹具未触发该分支，损坏契约已在比较前拒绝。

下一阶段先读 `STATUS.md` → `STAGE2_SPEC.md` 的C5.2 → 本检查点 → 新模块/测试；关系来源见C5.1，金额语义见C3.2。C5.2完成后停止，未进入C5.3或总体审核、short suffix豁免、baseline、UI、release。
