# Stage 2 C4.3 检查点（2026-09-12）

## 接口与职责

```vb
VATStage2RunFullFallback(ByRef a As VATS2ASnapshotResult, _
    ByRef b As VATS2BSnapshotResult, _
    ByRef completed As VATS2CompletedScopeResult) As VATS2FullFallbackResult
```

独立纯内存模块 `src/modVATStage2FullFallback.bas`。只接收既有快照和C4.2结果，不读Excel、不复制完整Snapshot、不处理金额、不修改第一层结果。身份判断只调用冻结Matcher/BRowMatch；不实现第二套匹配或解析算法。

主结果包含Status、BToAFallbackCount/Fallbacks、AToBFallbackCount/Fallbacks、AColorMissingCount、BColorMissingCount、CompletedACoverage、FullBScanned/FullBMatches、ABlockerCount/Blockers、BBlockerCount/Blockers及ErrorSide/Index/ReferenceIndex/Reason。

## 第一层优先与B→A

只搜索completed.BInScope=True且引用MatchKind=NOT_FOUND的引用。完整A临时String数组为1..RecordCount，位置就是原AIndex。String原样传入；非String安全空占位并记录blocker，不CStr、不Trim、不修复。

`VATS2BToAFallback`保存原BIndex、ExcelRow、ReferenceIndex、OriginalBMatch（完整第一层B行）、FullAMatch（冻结Matcher完整结果）、全部Candidates、Evidence、MissingEvidence。候选保存原AIndex、ExcelRow、InvoiceDigits、IsCompleted。

- full unique未完成A：A_COLOR_MISSING证据。
- full multiple：保留全部原候选及顺序，不消歧。
- full无候选：结合A blockers区分是否可作为可靠missing evidence。
- full unique已完成A：与第一层NOT_FOUND矛盾，主Status=INVALID_CONTRACT、Evidence=FULL_UNIQUE_COMPLETED，并保留错误BIndex/ReferenceIndex及现场。不得接受为成功匹配。

合并B只对失败引用fallback，OriginalBMatch仍保留整组。绝不拿已成功的局部A金额比较完整B金额。

## A→B与完整合并行

CompletedACoverage以完整AIndex空间保存：NONE=0、UNIQUE=1、MULTIPLE_ONLY=2（前缀VATS2_COVERAGE_）。存在任一completed B UNIQUE即覆盖；否则仅在MULTIPLE候选出现也属于已有歧义。两者都不反向搜索、不判B缺色。

只有完成A完全未在第一层候选出现才搜索完整B。存在至少一个此类目标时，所有B按原序调用冻结BRowMatch，使用完整A数组，存为FullBMatches(OriginalBIndex)。无反向目标时FullBScanned=False，不执行完整B搜索。

`VATS2AToBFallback`保存AIndex、ExcelRow、Evidence、HitCount/Hits、MissingEvidence、SourceBlocked。Hit保存BIndex、ExcelRow、IsCompleted、UniqueReferenceCount、MultipleReferenceCount。通过 `FullBMatches(Hits(i).BIndex)` 获取整个合并B的全部refs、成功/失败和全部候选；不是截取单个命中。

任何未完成B中存在UNIQUE reference指向目标A都保留为缺色证据。多个B全部保留；同一B同时存在unique与multiple也保留两个计数，仍须读取全部refs。只在multiple出现则FULL_AMBIGUOUS，不判B_COLOR_MISSING。full中已完成B唯一引用原先完全未覆盖的A，也判INVALID_CONTRACT。

FullBMatches是全表诊断，不能整体覆盖completed.BMatches。测试证明：即使full新增exact导致某已成立的第一层suffix结果不同，第一层覆盖及原结果仍保持。

AColorMissingCount与BColorMissingCount分别按原AIndex和BIndex去重计数，均为缺色记录数量。BToAFallbacks保留每个失败引用，AToBFallbacks保留每个待搜索完成A及全部命中；计数去重不删除证据，不提前处理跨行冲突。

## Blocker和状态

`VATS2FallbackBlocker`：OriginalIndex、ExcelRow、RawVarType、Reasons、ParserFlags。OriginalIndex按所在ABlocker/BBlocker数组解释为AIndex/BIndex，顺序为源记录顺序。

Reasons独立位（VATS2_BLOCK_前缀）：NON_STRING=1、STRUCTURE=2、NUMERIC_TRAILER=4、NO_REFERENCES=8、INVALID_A_STRING=16，可组合。

- A非String必须遮蔽；空或非ASCII数字String也保守记录blocker，但原字符串仍不清洗、不改变Matcher输入。
- B非String、Parser结构异常、纯数字尾注、零引用均可能遮蔽引用，保留blocker；正常缺NO或Parser去重提示不单独作为blocker。
- 无候选时，B→A看ABlockers；A→B看BBlockers及源A的SourceBlocked。任何可能遮蔽时MissingEvidence=False。
- blocker不抹去已找到的候选，也不改写冻结ParserFlags/MatcherFlags。即使找到缺色证据，后续仍可检查全局blockers。
- FullBScanned=False时未收集B blockers，BBlockerCount=0不能解释为已确认无风险。

Evidence均带VATS2_前缀：

| Evidence | 值 | 含义 |
|---|---:|---|
| FULL_UNIQUE_COMPLETED | 1 | 与第一层矛盾的已完成唯一命中，主状态失败 |
| FULL_UNIQUE_UNCOMPLETED | 2 | 存在未完成记录的唯一reference命中；反向可有多个B |
| FULL_MULTIPLE | 3 | B→A有多个A候选，不选择 |
| FULL_NOT_FOUND | 4 | 无候选且无已知遮蔽，MissingEvidence=True |
| FULL_NOT_FOUND_WITH_BLOCKERS | 5 | 无候选但搜索可能被遮蔽，不能判确定缺失 |
| FULL_AMBIGUOUS | 6 | A→B只有MULTIPLE候选，没有UNIQUE证据 |

主Status均带VATS2_前缀：FALLBACK_OK=0、FALLBACK_INVALID_INPUT=1、FALLBACK_INVALID_CONTRACT=2、FALLBACK_ENGINE_ERROR=3。

非OK输入直接拒绝。验证使用的数组、原Index、ExcelRow顺序、完成标记与映射、B来源原文、引用计数/状态/候选范围及号码来源。矛盾路径保留部分现场，仅供诊断；必须先看主Status，不能将预分配但未处理的条目当结果。运行时错误不伪装成缺失，返回ENGINE_ERROR及错误说明。

## 文件与原生Excel验证

新增：`src/modVATStage2FullFallback.bas`、`tests/modVATStage2FullFallbackTests.bas`、`tests/stage2-full-fallback-test-results.txt`、本检查点。

修改：`tools/build.ps1`仅增加FullFallbackTestOnly测试开关/分支及既有发布文件测试豁免；`STAGE2_SPEC.md`追加C4.3；`STATUS.md`更新入口。十组旧日志由原入口重新生成，旧测试源码未修改。

| `./tools/build.ps1`开关 | 实际结果 |
|---|---:|
| -FullFallbackTestOnly | **132/132 assertions PASS** |
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

所有最终测试运行exit0。首轮5个新测试失败来自夹具中额外A号码意外匹配，修正夹具后通过；无冻结算法修订。最终补充缺色A按行去重计数后复验新模块132项，冻结模块哈希保持不变。

覆盖双方缺色exact/suffix、多个候选、全表无候选、有/无blocker、第一层unique/multiple禁止fallback、矛盾已完成候选、merged B整行关系、多个B全保留、缺色A/B计数、非String多种类型、Parser结构风险、前导零、原索引行号。每次正常场景都比较所有输入Snapshot字段及完整completed结果不变，并重复调用对比完整输出确定性。

25个已有源码/旧测试/release文件执行前后SHA-256一致。未修改冻结C1～C4.2，未重建release，未改安全设置，未提交commit。仓库无提交基线，使用哈希补充git状态检查。

## 已知限制及接手入口

无已知阻塞问题。MissingEvidence仅表示当前快照及已知blocker规则下的缺失证据，不是最终A_ONLY/B_ONLY或审核结论。源数据中可解析但语义错误的号码不由本层推测；调用方必须提供同一批快照与completed结果。输入检查不替代完整重新匹配，也未引入fingerprint。

反向搜索按目标A遍历已缓存的全B匹配结果，规模很大时仍有内存/扫描开销；本轮不额外建立复杂索引。只保存值，不含Workbook/Worksheet/Range对象。

下一阶段先读 `STATUS.md` → `STAGE2_SPEC.md` → 本检查点 → 新模块/新测试；来源接口见C4.1/C4.2，匹配/冲突接口见C2检查点。复验命令见上表。

C4.3完成并停止。未开始金额比较、总体金额快速门、短suffix免审核、baseline/fingerprint、冻结、最终审核、UI或release接入。
