# Stage 2 C5.1 检查点（2026-09-12）

## 接口与结果

```vb
VATStage2BuildEffectiveRelations(ByRef a As VATS2ASnapshotResult, _
    ByRef b As VATS2BSnapshotResult, _
    ByRef completed As VATS2CompletedScopeResult, _
    ByRef fallback As VATS2FullFallbackResult) As VATS2EffectiveRelationsResult
```

新模块 `src/modVATStage2EffectiveRelations.bas`。纯内存；不访问Excel，不读金额，不调用GroupAmount，不重新调用Parser/Matcher，不修改四个输入或冻结模块。

结果保存Status、FullARecordCount、FullBRecordCount、EffectiveBCount、EffectiveBInScope()、EffectiveBMatches()、BRowIds()、AIntegrity、AInvalidValue()、ADuplicateGroupIndex()、RowIssueFlags()、RelationIssueCount/Issues()、EffectiveConflicts和ErrorSide/Index/ReferenceIndex/Reason。

数组保持原完整索引空间。B侧数组为1..FullBRecordCount，A质量摘要为1..FullARecordCount；零元素未分配。原Snapshot、completed与fallback继续由调用方保留，缺色/blocker/MissingEvidence仍从fallback读取，不复制成新的来源。

## Effective关系来源

1. 所有completed B进入范围，行结果从C4.2复制。原EXACT_UNIQUE、SUFFIX_UNIQUE、EXACT_MULTIPLE、SUFFIX_MULTIPLE保持第一层结果；即使full table另有exact，也不能替换。
2. 仅原NOT_FOUND按BIndex+ReferenceIndex查找唯一BToAFallback。数量及逐项一一对应均检查；原成功/歧义不能被patch。只复制FullAMatch的MatchKind、Flags、CandidateCount、全部Candidates；原AIndex不再回映。Parser字段和行ParserFlags原样保留。
3. 补充后重新汇总整行UniqueMatchCount、NotFoundCount、MultipleMatchCount、InvalidMatchCount、MatchState、MatcherFlags。多个NOT_FOUND分别补全，multiple/notfound仍保留诊断状态；旧短suffix标志可随实际full Matcher结果新增或清除。
4. C4.3反向结果里确有UNIQUE命中且为B_COLOR_MISSING的未完成B，按原BIndex去重后直接复制FullBMatches(BIndex)整行。命中的合并行必须保留其它成功/歧义/未找到refs。仅multiple命中不纳入。
5. 其它未完成B是合法NO_REFERENCES占位；不参与association。BRowIds(BIndex)始终为该Snapshot.ExcelRow。

反向Hit逐一与FullBMatches实际候选关系对照，验证BIndex/ExcelRow/IsCompleted及unique/multiple计数；记录缺失、多余、完成B误标缺色或FullBMatches原文错位都拒绝。此过程只读既有候选，不运行号码匹配算法。

## AIntegrity与关联风险

完整A身份String数组位置=原AIndex，String原样，非String用空String，不CStr、不Trim。只调用冻结VATStage2ScanAIntegrity；本层不重写判重或ASCII数字校验。

完整结果保存在AIntegrity。摘要：

- `AInvalidValue(AIndex)`：是否属于InvalidAIndexes。
- `ADuplicateGroupIndex(AIndex)`：所属DuplicateGroups序号；0表示不在重复组。

逐Effective UNIQUE candidate检查摘要：

- `VATS2_RELATION_A_INVALID = 1`
- `VATS2_RELATION_A_DUPLICATE = 2`

RowIssueFlags(BIndex)为该行所有问题关联的OR。RelationIssues额外保存每个问题UNIQUE的BIndex、ExcelRow、ReferenceIndex、AIndex、AExcelRow、Flags，便于后续精确定位。

本层保留问题关系和原MatchKind，不选择其它候选；无关A的数据质量问题只保存在AIntegrity，不把其它关系全局作废。真实测试覆盖：`XX123456`被冻结Matcher按suffix命中但AIntegrity判非法；一个完成A与另一个未完成A号码相同，第一层仍unique，但全表AIntegrity标重复成员。

**后续不得只看UNIQUE或EFFECTIVE_OK就进入可信金额核对。必须检查RowIssueFlags/RelationIssues；RELATION_A_INVALID和RELATION_A_DUPLICATE是阻碍可信金额关系的风险。** C5.1不执行金额门控，不改冻结GroupAmount。

## post-fallback冲突重扫

调用冻结VATStage2ScanBConflicts(EffectiveBMatches, BRowIds, FullBRecordCount)，完整结果保存EffectiveConflicts。没有另写冲突算法；有A质量风险的关系仍保留用于冲突诊断。

测试已真实发现：两个fallback B最终指向同A产生CROSS；同一B的完整号与尾号fallback后指同A产生SAME；completed B与整行纳入的缺色merged B复用同A产生CROSS。所有Association继续保存原AIndex、BIndex及BRowId=ExcelRow。

## 状态与契约

| 状态（VATS2_前缀） | 值 | 含义 |
|---|---:|---|
| EFFECTIVE_OK | 0 | 关系整合、完整性扫描和冲突重扫完成；不是审核通过 |
| EFFECTIVE_INVALID_INPUT | 1 | 任一Snapshot、CompletedScope或FullFallback非OK |
| EFFECTIVE_INVALID_CONTRACT | 2 | 数组、来源映射、引用/候选或计数契约错误 |
| EFFECTIVE_ENGINE_ERROR | 3 | AIntegrity或BConflict执行异常/非OK |

检查完整记录数量、原Index、ExcelRow严格递增、completed映射和范围、行原文、四类计数/行状态/MatcherFlags、候选范围及完整号码来源。fallback的原B行快照必须与completed一致；FullAMatch的引用Digits/Length必须与原reference对应。INVALID_REFERENCE/INVALID_A_RANGE等无效Matcher状态不能夹在OK来源中继续整合。

失败不发布部分Effective关系；清空输出，保留ErrorSide、ErrorIndex、ErrorReferenceIndex和错误说明。嵌套UDT默认Status=0不表示执行成功，调用方先检查主Status。

## 文件与测试

新增：

- `src/modVATStage2EffectiveRelations.bas`
- `tests/modVATStage2EffectiveTests.bas`
- `tests/stage2-effective-relations-test-results.txt`
- 本检查点。

修改：`tools/build.ps1`仅增加EffectiveRelationsTestOnly测试开关/分支和已有发布文件测试豁免；`STAGE2_SPEC.md`追加C5.1；`STATUS.md`最小更新。十一组旧日志由原测试入口重写，旧测试代码未修改。

原生Excel实际验证，所有最终运行exit0：

| `./tools/build.ps1`开关 | 结果 |
|---|---:|
| -EffectiveRelationsTestOnly | **135/135 assertions PASS** |
| -FullFallbackTestOnly | **132/132 PASS** |
| -CompletedScopeTestOnly | **73/73 PASS** |
| -SnapshotTestOnly | **81/81 PASS** |
| -GroupAmountTestOnly | **44/44 PASS** |
| -AmountTestOnly | **48/48 PASS** |
| -BConflictTestOnly | **41/41 PASS** |
| -BRowMatchTestOnly | **34/34 PASS** |
| -AIntegrityTestOnly | **29/29 PASS** |
| -MatcherTestOnly | **38/38 PASS** |
| -ParserTestOnly | **70/70 PASS** |
| -TestOnly | **89/89 PASS** |

新测试包含第一层unique/multiple优先、merged补全、逐项patch、计数刷新、短suffix flags新增/清除、Parser字段及非零风险保留、缺色B整行和去重、所有新增冲突、A数据质量关联、原索引及前导零、零记录，以及缺失/重复/越界patch、错位FullBMatches、错误Hit和无效Matcher状态拒绝。每个正常场景都验证Snapshot/completed/fallback所有字段不变，并重复调用比较完整Effective输出。

27个已有源码、旧测试及release文件执行前后SHA-256一致。未修改C1～C4.3、旧测试或安全设置；未重建release；未提交commit。Git无提交基线，使用哈希核对冻结文件。

## 已知限制与C5.2接手入口

无已知阻塞问题。输入必须来自同一业务批次；本层校验使用的关系结构和来源，不重新执行所有上游算法，也未引入fingerprint。有效关系和冲突均为诊断数据，不输出最终审核结论。AIntegrity/冲突的运行失败分支已保留，合法自建输入未触发这些引擎失败。

C5.2接手先读 `STATUS.md` → `STAGE2_SPEC.md` → 本检查点 → 新模块/新测试，再按需读C3.2金额协调器检查点及C4.2/C4.3来源结构。必须保留第一层优先、完整BIndex、整组refs和关联质量风险；不要直接把已标问题的UNIQUE当可信金额关系。

C5.1已完成并停止。未进入C5.2，未调用GroupAmount，未实施金额比较、最终A_ONLY/B_ONLY、总体金额快速门、baseline、冻结、最终审核、UI或release接入。
