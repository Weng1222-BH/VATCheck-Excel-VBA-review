# Stage 2 C5.3 检查点（2026-09-12）

## 接口与范围

```vb
VATStage2EvaluateShortSuffixPolicy(ByRef a As VATS2ASnapshotResult, _
    ByRef b As VATS2BSnapshotResult, _
    ByRef effective As VATS2EffectiveRelationsResult, _
    ByRef amounts As VATS2EffectiveAmountResult, _
    ByVal aggregateReliableEqual As Boolean) As VATS2ShortSuffixPolicyResult
```

`src/modVATStage2ShortSuffixPolicy.bas`是独立只读策略层，没有调用Parser、Matcher、金额或冲突扫描引擎。aggregateReliableEqual由调用方提供，表示Stage 1完成金额可靠、无空金额warning/致命异常、Decimal总体精确相等；本层不验证或计算这项总体资格。False统一阻止豁免，不细分其原因。

只为原引用实际携带SHORT_SUFFIX_REVIEW的条目生成决策。exact即使短于7位也不生成；长suffix不生成；冻结Matcher的short NOT_FOUND如果带该flag则生成REVIEW_REQUIRED。无实际flag绝不因长度短而新增决策。

## 豁免条件和返回值

只有总体资格成立、所在B整行全部UNIQUE且完整、C5.2实际比较成功且AMOUNT_EQUAL、关系质量和Parser风险为零、没有SAME/CROSS、B完成色及全部实际唯一匹配A完成色、没有其它身份结构风险时，返回WAIVED_THIS_RUN。

否则REVIEW_REQUIRED并组合原因。缺色的是组内另一条exact/长suffix成员，也会阻止本组short豁免。未知Matcher风险位保守记为OTHER_STRUCTURE_RISK，不清除任何原标志。

扁平Decisions()按原BIndex、ReferenceIndex顺序，保存BIndex、ExcelRow、ReferenceIndex、AIndex、ReferenceDigits、Decision、ReasonFlags。未找到/多候选引用AIndex=0，不擅自选择候选；来源完整候选继续保留在调用方的EffectiveRelations。

Decision：VATS2_REVIEW_REQUIRED=0，VATS2_WAIVED_THIS_RUN=1。

| ReasonFlags（VATS2_前缀） | 值 |
|---|---:|
| AGGREGATE_NOT_ELIGIBLE | 1 |
| ROW_NOT_COMPLETE_UNIQUE | 2 |
| GROUP_NOT_EQUAL | 4 |
| PARSER_RISK | 8 |
| RELATION_QUALITY_RISK | 16 |
| CONFLICT_RISK | 32 |
| COLOR_MISSING | 64 |
| OTHER_STRUCTURE_RISK | 128 |

结果包含Status、AggregateReliableEqual、DecisionCount、WaivedCount、ReviewRequiredCount及错误B/引用索引、原因。Status：VATS2_SHORT_POLICY_OK=0、VATS2_SHORT_POLICY_INVALID_INPUT=1、VATS2_SHORT_POLICY_INVALID_CONTRACT=2。失败不发布部分豁免；空结果Decisions未分配。

## 必要契约与限制

检查依赖的Snapshot/Effective/Amount/冲突状态、记录数及范围一致性、原B行号、唯一A身份/索引、质量标志一致性、组来源/实际比较状态和金额AIndex对应。不检查不使用的金额原值、不重做C5.1/C5.2完整契约、不引入指纹或全局完整性重扫。

冲突风险读取冻结组的SAME/CROSS及原EffectiveConflicts关联是否涉及该B；不计算新的冲突组。四份输入及所有MatcherFlags始终保持原样。

无已知功能阻塞。调用方必须提供同一批数据的上游结果和正确的aggregateReliableEqual；本层不证明总体金额资格、也不检测所有跨批次/输入同时被篡改的情形。豁免只适用于当前运行的短suffix人工复核，不代表组最终PASS。

## 文件与实际验证

新增：
- src/modVATStage2ShortSuffixPolicy.bas
- tests/modVATStage2ShortPolicyTests.bas
- tests/stage2-short-suffix-policy-test-results.txt
- 本检查点

修改tools/build.ps1（最小增加开关和独立分支）、STAGE2_SPEC.md（追加C5.3）、STATUS.md（当前入口）。旧日志由原测试重新产生，仅Stage 1耗时记录可能变化。

| build.ps1开关 | 实际Excel结果 |
|---|---:|
| -ShortSuffixPolicyTestOnly | 139/139 assertions PASS |
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

合计1153项全部PASS。新测试覆盖全部豁免条件、每种主要原因、相等与不等/金额错误、不完整/多候选/未找到、SAME/CROSS、两侧缺色及非short成员缺色、exact/长suffix/无short、多个引用原序、69个决策跨容量边界、未知身份风险、资格切换、原风险不变、四份输入逐字段不变及重复调用确定，并包括23项依赖契约拒绝测试。

31个冻结源码、旧测试及release文件前后SHA-256一致。未更改Excel安全设置或执行策略，未重建release。Private和Public使用各自独立历史提交；公开同步沿用白名单、仅机器环境信息脱敏及安全扫描流程。

## 后续接手入口

先读STATUS.md → STAGE2_SPEC.md的C5.3 → 本检查点 → 新模块和新测试。需要关系/金额语义时再读C5.1/C5.2检查点。本轮未开始最终审核、总体报告、baseline/fingerprint、UI、release或下一阶段。
