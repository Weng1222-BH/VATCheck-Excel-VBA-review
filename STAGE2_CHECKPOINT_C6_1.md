# Stage 2 C6.1 检查点（2026-09-12）

## 接口与职责

```vb
VATStage2EvaluateAggregateGate(ByVal headerA As Range, _
    ByVal headerB As Range, _
    ByVal completedColor As Long) As VATS2AggregateGateResult
```

新模块 `src/modVATStage2AggregateGate.bas`，只读Excel。分别直接调用冻结VATScan两次，用冻结VAT_FIELD_A/VAT_FIELD_B和共用完成色；不修改任何冻结业务模块，不调用VATResult，不解析结果文本，不重新逐格扫描或复制金额类型/空白/颜色/累加规则。

在扫描前检查Application.CalculationState，双侧成功后再次检查；非xlDone返回CALCULATION_PENDING，不发布金额结论。不触发计算、不改变全局计算策略、不等待/监控。单侧扫描失败仍执行另一侧扫描，以保留完整诊断证据。

## 状态与字段

| Status | 值 | 含义 |
|---|---:|---|
| VATS2_AGGREGATE_OK | 0 | 双侧扫描可靠成功、计算完成且Decimal差额成功 |
| VATS2_CALCULATION_PENDING | 1 | CalculationState非xlDone |
| VATS2_SCAN_BLOCKED | 2 | 任一VATScan失败/致命Issue或读取调用异常 |
| VATS2_AGGREGATE_ARITHMETIC_ERROR | 3 | 双侧扫描成功但两个Decimal总额相减失败 |

返回值还保存：
- CalculationState、AScanSucceeded、BScanSucceeded。
- TotalA、TotalB、Difference：Variant/Decimal；Difference方向始终A-B。
- TotalsEqual：Variant，AGGREGATE_OK时为Boolean；其它状态为Empty。
- ReliableEqualForShortSuffix：Boolean。
- ACompletedCount/BCompletedCount、AIncludedCount/BIncludedCount、AEmptyWarningCount/BEmptyWarningCount。
- AIssueCount/AIssues()、BIssueCount/BIssues()、AWarningCount/AWarnings()、BWarningCount/BWarnings()；String数组1-based、零项未分配。逐字保留VATScan原始文本和原序，不混合两侧。
- ErrorReason只记录本模块自己的读取/算术错误，不伪造VATScan问题。

非OK时TotalA/TotalB/Difference/TotalsEqual全部Empty，资格False；不泄露部分有效总额作为总体结论。扫描已取得的计数及证据仍保留。双側成功不代表减法必定成功，独立算术失败状态防止Decimal溢出被误当作金额不等。

## 与C5.3的连接契约

仅当Status=AGGREGATE_OK、双侧扫描成功且无致命Issue、两侧warning为0、Decimal差额精确为0，ReliableEqualForShortSuffix=True。真正空白产生warning时TotalsEqual仍可True，但资格必须False。

行数不决定一致性；零完成记录的可靠零相等按上述条件提供True。本层只产生资格，不调用C5.3，也不执行最终审核或最终PASS/FAIL。

## 新增/修改文件

新增：
- src/modVATStage2AggregateGate.bas
- tests/modVATStage2AggregateTests.bas
- tests/stage2-aggregate-gate-test-results.txt
- 本检查点

修改tools/build.ps1，最小加入-AggregateGateTestOnly及测试豁免；加载原frmVATCheck仅为满足modVATCheck编译依赖，不显示界面、不导出release。STAGE2_SPEC.md追加C6.1，STATUS.md更新入口。既有测试日志由原入口重写，未改变旧测试语义。

## 实际原生验证

| build.ps1开关 | 结果 |
|---|---:|
| -AggregateGateTestOnly | 170/170 assertions PASS |
| -ShortSuffixPolicyTestOnly | 139/139 PASS |
| -EffectiveAmountTestOnly | 200/200 PASS |
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

合计1323项全部通过。新测试覆盖正负差额、多A少B、0.1+0.2=0.3、零完成、单/双侧真正空白warning、仅空白、公式空串/数字文本/错误值/布尔、两侧多问题与警告共存、合并金额、转换失败、失效/Nothing表头及两个可靠总额相减溢出。

每次比较同时对照直接调用冻结VATScan的状态、六项计数及两侧Issue/Warning完整文本；验证输入值/公式/颜色/合并/格式、工作簿Saved与数量、Application.Calculation不变，并逐字段检查重复调用确定。

33个冻结源码、旧测试及release文件执行前后SHA-256一致。未修改Excel安全/计算策略，未重建release。最初测试调用的VBA标签歧义已在新测试中用显式Call修正；最终全部测试实际通过。

## 已知限制与接手入口

无已知功能阻塞。xlDone通路原生验证；未稳定制造xlPending/xlCalculating，非Done分支只按用户允许做代码级检查：扫描前及双侧成功后均有guard并提前返回，不触发扫描/金额结论。未将该分支计作实际运行测试。

C6.2接手先读STATUS.md → STAGE2_SPEC.md的C6.1 → 本检查点 → 新模块/测试；资格消费者见C5.3检查点。C6.1完成即停止，未开始最终审核分类、A_ONLY/B_ONLY汇总、baseline/fingerprint、UI、release或C6.2。
