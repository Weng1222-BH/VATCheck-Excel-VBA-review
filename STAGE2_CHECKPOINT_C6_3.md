# Stage 2 C6.3 检查点（2026-09-13）

## 接口与职责

`VATStage2BuildFinalDecision(aggregate, amounts, findings) As VATS2FinalDecisionResult`

三个输入分别为冻结 `VATS2AggregateGateResult`、`VATS2EffectiveAmountResult`、`VATS2AuditFindingsResult`，均为只读ByRef。新模块 `src/modVATStage2FinalDecision.bas` 不访问Excel、不重算金额、不读取/重分析逐条Finding或逐行金额数组、不调用上游引擎、不生成用户报告文本。

## Status与Verdict

Status：`VATS2_FINAL_DECISION_OK=0`、`VATS2_FINAL_INVALID_INPUT=1`、`VATS2_FINAL_INVALID_CONTRACT=2`。

Verdict：`VATS2_FINAL_UNAVAILABLE=0`（安全默认）、`VATS2_FINAL_VERIFIED=1`、`VATS2_FINAL_REVIEW_REQUIRED=2`。

- 合法Aggregate非OK（pending、scan blocked、arithmetic error）：UNAVAILABLE，RequiresManualReview=False。
- Aggregate OK、TotalsEqual=True、两侧WarningCount=0、Audit OK、FindingCount=0：VERIFIED。
- Aggregate OK且任一总体差异/warning/Finding存在：REVIEW_REQUIRED，RequiresManualReview=True。
- EffectiveAmount/AuditFindings合法但非OK：Status=INVALID_INPUT、Verdict=UNAVAILABLE，不把失败上游视为零异常。未知主状态或损坏必要契约：INVALID_CONTRACT、UNAVAILABLE。上游状态及ErrorSource/ErrorReason保留，不发布部分摘要。

RequiresManualReview仅对应REVIEW_REQUIRED。UNAVAILABLE不是人工复核业务异常，不能当作通过或不通过。

## 返回摘要

- 原AggregateStatus、EffectiveAmountStatus、AuditStatus。
- TotalA、TotalB、Difference、TotalsEqual：前三项直接复制原Variant/Decimal，不转值、不再相减；TotalsEqual保留Variant/Boolean。
- A/BCompletedCount、A/BIncludedCount、A/BWarningCount。
- EffectiveBCount、ComparedCount、EqualCount、MismatchCount、NotComparableCount、AmountErrorCount、RelationQualityBlockedCount。
- FindingCount、CodeCounts(1 To 18)、ShortSuffixWaivedCount；不复制原Findings数组。
- HasAggregateMismatch、HasAggregateWarnings、HasFindings、NoCompletedRecords、RequiresManualReview。

合法总体门非OK时仍可保留已验证的诊断计数与HasFindings/HasAggregateWarnings，但金额及TotalsEqual全部Empty，HasAggregateMismatch=False、NoCompletedRecords=False、RequiresManualReview=False；即使传入残留金额也不发布。

仅可靠Aggregate OK时按双方CompletedCount都为0设置NoCompletedRecords。零完成且满足可靠相等条件仍允许VERIFIED；pending/失败返回的默认零计数不被误当作实际零完成记录。

总体warning不生成新Finding，总体不等不伪造逐组AMOUNT_MISMATCH；逐组equal不能覆盖总体不等，总体相等也不能覆盖任何已有Finding。ShortSuffixWaivedCount只统计，不改变Verdict。各Code无严重程度优先级，不合并、不删除。

## 必要契约

只检查主状态范围；金额层及事项层必须OK；所有摘要计数以及相关原计数非负；Aggregate OK时TotalsEqual必须是Boolean。

EffectiveAmount：ComparedCount=EqualCount+MismatchCount；EffectiveBCount=ComparedCount+NotComparableCount+AmountErrorCount+RelationQualityBlockedCount+InvalidGroupCount，且不超过FullBRecordCount；OK时InvalidGroupCount=0。

FindingCount=18个CodeCounts总和，ShortSuffixWaivedCount非负。计数求和用Double承载Long计数之和以避免Long溢出；这些计数范围内整数精确，不涉及任何财务金额计算。

不重新检查金额一致性、不检查逐行关系或Finding细节；调用方仍须提供同一批真实上游结果，不引入fingerprint。ErrorReason仅为接口诊断说明，不是最终用户报告。

## 文件与验证

新增：本检查点、`src/modVATStage2FinalDecision.bas`、`tests/modVATStage2FinalTests.bas`、`tests/stage2-final-decision-test-results.txt`。

修改：`tools/build.ps1`仅加入`-FinalDecisionTestOnly`和分支/测试豁免，`STAGE2_SPEC.md`追加C6.3，`STATUS.md`更新阶段入口；旧测试日志按原入口重新生成。AggregateGate类型依赖Stage1模块/窗体，因此测试宿主加载冻结窗体供编译，未显示UI或调用VATScan。

新测试覆盖全部八种mismatch/finding/warning组合、B侧warning、三种不可用状态及残留金额隔离、零记录与单侧零记录、short豁免不抵销事项、总体/逐组分层、18个Code统计、高精度Decimal原样复制、全部摘要传播、三个上游全字段及嵌套数组不变、重复调用确定，以及39项非法状态/类型/计数/溢出边界拒绝。

| 测试入口 | 原生Excel结果 |
|---|---:|
| FinalDecisionTestOnly | 172/172 PASS |
| AuditFindingsTestOnly | 194/194 PASS |
| AggregateGateTestOnly | 170/170 PASS |
| ShortSuffixPolicyTestOnly | 139/139 PASS |
| EffectiveAmountTestOnly | 200/200 PASS |
| EffectiveRelationsTestOnly | 135/135 PASS |
| FullFallbackTestOnly | 132/132 PASS |
| CompletedScopeTestOnly | 73/73 PASS |
| SnapshotTestOnly | 81/81 PASS |
| GroupAmountTestOnly | 44/44 PASS |
| AmountTestOnly | 48/48 PASS |
| BConflictTestOnly | 41/41 PASS |
| BRowMatchTestOnly | 34/34 PASS |
| AIntegrityTestOnly | 29/29 PASS |
| MatcherTestOnly | 38/38 PASS |
| ParserTestOnly | 70/70 PASS |
| TestOnly（Stage1） | 89/89 PASS |

合计1689项全部通过。37个冻结源码、旧测试和release文件SHA-256保持不变，未修改Excel安全/计算设置，未重建release。无已知功能阻塞。

## 接手边界

后续先读STATUS.md → STAGE2_SPEC.md的C6.3 → 本检查点 → 新模块/测试。结果须先看Status、再看Verdict；不可用结果不解释成业务人工审核。未开始fingerprint、持久化、freeze/unfreeze、UI、report、release或下一阶段。
