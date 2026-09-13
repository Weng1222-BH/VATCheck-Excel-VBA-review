# Stage 2 C6.2 检查点（2026-09-13）

## 接口与边界

`VATStage2BuildAuditFindings(a, b, fallback, effective, amounts, shortPolicy) As VATS2AuditFindingsResult`

六份输入分别为冻结 A/B Snapshot、FullFallback、EffectiveRelations、EffectiveAmount、ShortSuffixPolicy，均为只读 ByRef。新模块 `src/modVATStage2AuditFindings.bas` 只整理已有证据，不访问 Excel/文件，不调用 Parser、Matcher、AIntegrity、冲突、fallback、金额或短尾号引擎。AggregateGate 不作为 Finding 来源。

Status：`VATS2_AUDIT_OK=0`、`VATS2_AUDIT_INVALID_INPUT=1`、`VATS2_AUDIT_INVALID_CONTRACT=2`。OK 只表示证据整理成功，绝不是批次审核通过。失败清空全部部分 Findings/计数，保留 ErrorSide、ErrorIndex、ErrorReferenceIndex、ErrorReason。

## Finding 接口

结果：FindingCount、1-based Findings()、固定 CodeCounts(1 To 18)、ShortSuffixWaivedCount；零事项 Findings 未分配。

`VATS2AuditFinding` 保存 Code、Side、AIndex/AExcelRow、BIndex/BExcelRow、ReferenceIndex/ReferenceDigits、SourceFlags、ReasonFlags、ParserFlags、MatcherFlags、MatchKind、GroupIndex、MemberCount/Members()、Difference、GroupState、AmountStatus、AmountErrorSide/AmountErrorIndex、RawVarType、Detail。0索引表示该层级不适用，不能把歧义项的0当作选定了某个候选。

Members() 是 `VATS2FindingLink` 数组，逐成员保存 AIndex/AExcelRow、BIndex/BExcelRow、ReferenceIndex/ReferenceDigits。重复A保留全组A；歧义保留全部候选；CROSS保留整组所有关联；SAME按冲突组中的实际B保留其全部重复关联；金额差异保留冻结金额结果的完整A成员顺序。没有截取首个歧义候选。

Code 使用 `VATS2_FIND_` 前缀，数值直接索引 CodeCounts：

| 值 | 名称 | 来源 |
|---:|---|---|
| 1 | A_COLOR_MISSING | 完成B的B→A FULL_UNIQUE_UNCOMPLETED，逐关系定位 |
| 2 | B_COLOR_MISSING | A→B唯一Hit，读取FullBMatches中的实际唯一引用，逐A-B-ref定位 |
| 3 | B_ONLY | 完成B引用全A无候选、MissingEvidence=True、无A blocker |
| 4 | A_ONLY | 完成A全B无候选、MissingEvidence=True、无B blocker且SourceBlocked=False |
| 5 | AMBIGUOUS_MATCH | Effective EXACT/SUFFIX_MULTIPLE、FULL_MULTIPLE、反向FULL_AMBIGUOUS及Hit中的multiple |
| 6 | AMOUNT_MISMATCH | C5.2实际GROUP_COMPARED且AMOUNT_MISMATCH；Difference原样保留Decimal A−B |
| 7 | AMOUNT_ERROR | GROUP_AMOUNT_ERROR；保留原金额引擎状态、错误位置和说明 |
| 8 | INCOMPLETE_GROUP | Effective范围内BROW_INCOMPLETE或NO_REFERENCES；完整关系仅质量阻断不误标 |
| 9 | DUPLICATE_A_INVOICE | AIntegrity每个DuplicateGroup一项，全组成员完整保留 |
| 10 | A_DATA_REVIEW | AIntegrity每个InvalidAIndex，不按是否参与关系过滤 |
| 11 | DUPLICATE_REFERENCE | Parser重复提示与SAME来源独立；SAME不伪装为NOT_FOUND/关系不完整 |
| 12 | CROSS_B_ROW_REUSE | 原EffectiveConflicts冲突组及全部B关联，留给后续人工审核 |
| 13 | STRUCTURE_REVIEW | 原ParserFlags结构风险 |
| 14 | NUMERIC_TRAILER_REVIEW | 原ParserFlags数字尾注风险 |
| 15 | NO_MARKER_MISSING | 原ParserFlags无NO提示 |
| 16 | SHORT_SUFFIX_REVIEW | 仅C5.3 REVIEW_REQUIRED；原ReasonFlags保留 |
| 17 | SEARCH_BLOCKER | 原A/B blocker的side、原索引/行号、Reasons、ParserFlags、RawVarType |
| 18 | RELATION_QUALITY_REVIEW | 原RelationIssues逐关联定位，保留质量阻断的具体来源 |

SourceFlags 使用 `VATS2_FIND_FROM_` 前缀：EFFECTIVE=1、B_TO_A=2、A_TO_B=4、INTEGRITY=8、PARSER=16、CONFLICT=32、AMOUNT=64、SHORT_POLICY=128、BLOCKER=256、RELATION_QUALITY=512。各原始 Flags 保持各自位空间；ReasonFlags须按Code/Source解读，不混用枚举。

Parser风险整理Effective行；若已执行FullBScanned，范围外B中已发现的Parser风险也保留。没有运行full B时不猜测未扫描风险。Parser重复只用DUPLICATE_REFERENCE表达，不再产生重复的Parser重复事项。

C5.3 WAIVED_THIS_RUN不产生异常Finding，仅递增ShortSuffixWaivedCount；不重判豁免、不清除MatcherFlags。独立问题不互相覆盖，缺色与金额差异可共存。

## 稳定性与契约

去重键包含Code、适用位置、引用、MatchKind、组号及有序成员；同一歧义证据同时来自effective/fallback时仅OR合并SourceFlags。Dictionary只查键，不依赖枚举顺序。

稳定顺序：A独立事项在前按AIndex；B事项按BIndex、ReferenceIndex、Code；并列保留上游证据原序。冲突组/成员原始关联顺序保留，所有数组1-based。

检查所用上游主状态及AIntegrity/冲突状态、完整A/B计数、原Index/ExcelRow、有效范围、实际引用/候选、金额来源状态及成员位置、blocker/SourceBlocked/MissingEvidence与Evidence一致性；每个C5.3决策一一映射回实际short flag、原Digits/位置/AIndex，拒绝重复或遗漏。未重跑上游算法或复制上游完整防御性校验。

## 文件与真实验证

新增：本检查点、`src/modVATStage2AuditFindings.bas`、`tests/modVATStage2AuditTests.bas`、`tests/stage2-audit-findings-test-results.txt`。

修改：`tools/build.ps1`仅新增`-AuditFindingsTestOnly`分支/参数及测试豁免；`STAGE2_SPEC.md`追加C6.2；`STATUS.md`更新当前入口；Stage1测试日志重跑后仅扫描耗时变化。

| 测试入口 | 原生Excel结果 |
|---|---:|
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

合计1517项全部通过。新测试覆盖各Code、缺色与金额并存、同B多风险、反向歧义证据去重、组成员/错误位置、blocker禁止虚假missing、短尾号豁免计数、空输入、顺序与重复调用全输出确定；逐字段序列化六份输入验证无修改，并覆盖32个损坏/矛盾契约拒绝场景。35个冻结源码、旧测试、release文件SHA-256保持不变。

首轮新测试存在无参数Prepare调用被VBA当重复标签的问题，已用显式Call修正。最终测试实际通过；没有修改安全/计算设置、冻结业务模块或release。

## 已知限制与后续入口

无已知功能阻塞。调用方须提供同一批数据的冻结上游结果；本层验证实际使用的来源，不证明整个上游证据未经共同篡改，不引入fingerprint。Finding是诊断事实，不给整个批次最终PASS/FAIL，不输出中文用户报告。

接手先读STATUS.md → STAGE2_SPEC.md的C6.2 → 本检查点 → 新模块及测试。后续总体门来源见C6.1；本轮没有进入最终审核结论、baseline、UI、report、release或下一阶段。
