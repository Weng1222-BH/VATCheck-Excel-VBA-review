# Stage 2 C4.2 检查点（2026-09-12）

## 接口

```vb
VATStage2MatchCompletedScope(ByRef a As VATS2ASnapshotResult, _
    ByRef b As VATS2BSnapshotResult) As VATS2CompletedScopeResult

VATStage2RemapScopeRow(ByRef source As VATS2BRowMatchResult, _
    ByRef originalAIndexes() As Long, ByVal scopeCount As Long, _
    ByRef remapped As VATS2BRowMatchResult) As VATS2ScopeStatus
```

新模块 `src/modVATStage2CompletedScope.bas`。纯内存，只组合冻结BRowMatch/BConflict与快照结构，不访问Excel，不复制Snapshot全量业务内容，不处理金额或再次读取颜色。

主结果保存Status、FullARecordCount、FullBRecordCount、CompletedACount、CompletedBCount、CompletedAIndexes()、CompletedAExcelRows()、AIdentityIssueCount/Issues()、BInScope()、BRowIds()、BSourceIssueCount/Issues()、BMatches()、Conflicts及ErrorSide/Index/Reason。调用方继续保留原Snapshot。

## 原索引与完成范围

按完整A顺序提取IsCompleted=True，形成临时String号码数组以及ScopePosition→OriginalAIndex映射。例如scope1/2/3对应原A1/3/5，返回的所有Candidate.AIndex必须是1/3/5。

只对完成B调用冻结BRowMatch。BMatches/BInScope/BRowIds分配完整1..BSnapshot.RecordCount，记录内BIndex必须等于数组位置，BRowIds保存该原B的ExcelRow。未完成B直接使用默认零引用NO_REFERENCES占位，不调用Parser/Matcher。必须先看BInScope，不能把占位当已匹配诊断。

回映函数先验证整个映射严格正数递增、数组1-based且数量一致，以及每个候选的scope位置；全部通过才复制源行并修改Candidate.AIndex。失败返回SCOPE_INVALID_CONTRACT，输出行清空，没有混合新旧索引的部分结果。source与remapped使用不同变量。除AIndex外不改任何字段，不重跑Matcher。

## 源值Issue

- 完成A的InvoiceDigitsRaw是String：原样复制。非String：临时号码位置为安全空String，仍属于scope，记录OriginalAIndex、ExcelRow、RawVarType。
- 完成B的SupplierTextRaw是String：原样传入BRowMatch。非String：用空String得到NO_REFERENCES，同时记录OriginalBIndex、ExcelRow、RawVarType。
- Issue顺序保留原记录顺序，非String不令整个scope失败。不CStr、不Trim、不转数字。未完成记录不产生这两类Issue。空String是String，不误标类型Issue；一般字符串质量不在本层清洗。
- 带BSourceIssue的NO_REFERENCES不能解释为真实缺票；所有Flags仅保留诊断意义，不自动消歧或免审核。

## 冲突与状态

对完整BMatches/BRowIds调用冻结VATStage2ScanBConflicts。Association.BIndex为原BIndex，BRowId为ExcelRow，AIndex为已回映的原AIndex。SAME/CROSS仍用冻结逻辑；未完成占位无association。Conflicts.BRecordCount是完整B数量；CompletedBCount另行保存。

| 状态（VATS2_前缀） | 值 | 含义 |
|---|---:|---|
| SCOPE_OK | 0 | 完成范围号码关系已建立，不代表业务审核通过 |
| SCOPE_INVALID_INPUT | 1 | 任一Snapshot.Status非OK |
| SCOPE_INVALID_CONTRACT | 2 | 快照数组/索引/行序损坏或候选回映失败 |
| SCOPE_CONFLICT_ERROR | 3 | 冻结冲突扫描未成功 |

快照契约要求HeaderRow>=1、非负计数、数组精确为1..RecordCount、原Index=数组位置、ExcelRow在表头后严格递增；零计数时数组未分配。即使损坏行未完成，也不能掩盖接口问题。

主接口回映异常清空部分scope结果并保留错误定位。冲突返回非OK时保留嵌套Conflicts错误和已经完成的B行诊断，主Status明确失败。运行时错误按构建/冲突阶段返回非OK及原始错误说明。必须首先检查主Status；未执行路径中嵌套UDT默认Status=0不是成功证据。

## 文件与实际验证

新增：

- `src/modVATStage2CompletedScope.bas`
- `tests/modVATStage2CompletedScopeTests.bas`
- `tests/stage2-completed-scope-test-results.txt`
- 本检查点。

修改：`tools/build.ps1`最小增加CompletedScopeTestOnly开关、分支及已有release测试豁免；`STAGE2_SPEC.md`仅追加C4.2；`STATUS.md`最小更新。九份旧测试结果日志通过原入口重新生成，旧测试代码未改。

真实Excel顺序执行，全部exit0：

| `./tools/build.ps1`开关 | 结果 |
|---|---:|
| -CompletedScopeTestOnly | **73/73 assertions PASS** |
| -SnapshotTestOnly | **81/81 PASS** |
| -GroupAmountTestOnly | **44/44 PASS** |
| -AmountTestOnly | **48/48 PASS** |
| -BConflictTestOnly | **41/41 PASS** |
| -BRowMatchTestOnly | **34/34 PASS** |
| -AIntegrityTestOnly | **29/29 PASS** |
| -MatcherTestOnly | **38/38 PASS** |
| -ParserTestOnly | **70/70 PASS** |
| -TestOnly | **89/89 PASS** |

测试覆盖本轮要求的完成A/B exact/suffix、1/3/5回映、multiple保序及exact优先、原B索引、无完成A/B、前导零、非String各类型Issue、Flags/SHORT_SUFFIX、SAME/CROSS索引及ExcelRow、重复A、输入不变、重复调用、非法Snapshot、候选回映越界。逐字段对照冻结BRowMatch，确认只改变AIndex。未完成B异常文本不被解析；完成A的数值12345678不被转换制造匹配。金额字段测试放入错误/文本，协调层不使用金额。

23个已有源码、旧测试和release文件执行前后SHA-256一致。未更改冻结C1～C4.1模块或旧测试，未改安全设置，未重建release，未提交commit。当前Git无提交基线，已用哈希核对而非依赖空git diff。

## 已知限制与C4.3入口

无已知阻塞问题。Scope只表示双方已完成色范围内的号码关系，NOT_FOUND不解释为全表缺失或漏色，不判A_ONLY/B_ONLY。A的String值遵循传入契约不清洗，其完整数字质量仍由AIntegrity及调用层负责；两个Snapshot属于同一业务批次的责任由调用方承担，本轮无fingerprint。

冲突错误传播分支已实现；冻结函数在本层生成的合法结果中不会自然产生冲突接口错误，没有为制造此错误篡改冻结模块或增加注入框架。候选越界通过公开回映接口直接验证。

C4.3接手先读 `STATUS.md` → `STAGE2_SPEC.md` → 本检查点 → 新模块/新测试；快照来源见C4.1，行结果/冲突结构见C2.2b/C2.2c。测试复验命令见上表。

C4.2完成并停止，尚未开始full-table fallback、颜色缺失分类、金额、baseline、最终审核、UI或release接入。
