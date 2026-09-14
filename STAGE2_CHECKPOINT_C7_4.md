# Stage 2 C7.4 检查点：Bound Full Audit Run（2026-09-15）

## 恢复现场与架构调整

恢复时仓库clean、HEAD为C7.3（8fe6756）；C7.4尚停在规格及接口核对，没有源码、测试、接线或未完成日志。未删除、覆盖或撤销已有工作。

冻结C6.3 FinalDecision不包含批次身份，裸FinalDecision与任意fp事后拼接不能证明同批。C6.3继续冻结；原计划的Verified Seal推迟至C7.5。本轮通过新的受控调用消除这种API误用路径，不回改旧模块。

## 唯一绑定入口

`VATStage2RunBoundFullAudit(aInvoiceHeader As Range, aAmountHeader As Range, bSupplierHeader As Range, bAmountHeader As Range, completedColor As Long) As VATS2BoundAuditRunResult`

五个参数均ByVal；不能传入外部FinalDecision、Fingerprint、Aggregate、Findings或Effective结果。四个表头须为当前Excel实例中有效且未合并的单个单元格；同侧同Worksheet、同一行、身份与金额不同列。身份表头须非空文本，金额表头严格保持A=`有效抵扣税额*`、B=`税额`的Value2二进制匹配。身份字段不另设固定列或重写自动定位。工作簿/Sheet必须仍然有效。

Worksheet、HeaderRow、身份列、金额列完全从这四个Range推导，不接受独立列号。A/B可处于同一或不同工作簿，各自表头行可不同。

## 内部完整链与绑定

读取A/B Snapshot → BuildFingerprints → AggregateGate（原金额header和共享颜色）→ 复读确认Snapshot及Fingerprint → CompletedScope → FullFallback → EffectiveRelations → EffectiveAmount → ShortSuffixPolicy → AuditFindings → FinalDecision → BuildBaseline(fp)仅获取摘要。

所有步骤直接调用冻结入口，不复制业务算法。short policy的资格只能来自本次Aggregate.ReliableEqualForShortSuffix。确认后关系、金额和Findings链使用同一份原内存Snapshot；最后摘要也来自这份Snapshot生成的原fp。调用者没有注入另一份审核结果的入口。

只有整条链成功才一次性发布BindingProtocol=`VATRUN1`、FingerprintProtocol=`VATFP1`、DigestProtocol=`SHA256-UTF8`、A/BRecordCount、A/BBatchDigest、CombinedDigest，以及本次FinalDecision、完整AuditFindings、Aggregate。其它大体积中间对象只存于局部变量，不复制到返回值。

**BoundRun OK不等于Final VERIFIED。** 合法Aggregate CALCULATION_PENDING、SCAN_BLOCKED、AGGREGATE_ARITHMETIC_ERROR不直接中止其它可继续的阶段；冻结C6.3可产生UNAVAILABLE，且该技术上OK的绑定仍不能用于未来Verified Seal。业务Finding或总额不等可以产生OK + REVIEW_REQUIRED。

## 状态与当前运行稳定性

Status（`VATS2_BOUND_RUN_`前缀）：OK=0、INVALID_INPUT=1、STAGE_FAILED=2、SOURCE_CHANGED=3、INVALID_CONTRACT=4。只有OK时可读取绑定digest、FinalDecision及AuditFindings。

非OK清空协议、digest、记录计数和部分事项；FinalDecision明确INVALID_INPUT/UNAVAILABLE，AuditFindings明确INVALID_INPUT，避免默认零值伪装成成功。ErrorStage和ErrorReason定位失败，不保存前一批结果。

阶段定位覆盖INPUT、SNAPSHOT_A、SNAPSHOT_B、FINGERPRINT、AGGREGATE、SOURCE_STABILITY、COMPLETED_SCOPE、FULL_FALLBACK、EFFECTIVE_RELATIONS、EFFECTIVE_AMOUNT、SHORT_SUFFIX、AUDIT_FINDINGS、FINAL_DECISION、BASELINE_DIGEST。未知Aggregate状态为INVALID_CONTRACT；合法非OK业务状态保留给FinalDecision。

第一次Snapshot/Fingerprint之后、Aggregate之后重新验证同四个header的位置并复读A/B。纯内存辅助`VATStage2BoundSourcesStable(a,b,fp,confirmA,confirmB,confirmFP) As Boolean`检查成功状态、记录数量/必要边界、每个当前位置的record fingerprint及ExcelRow。任何变化拒绝绑定；复读阶段自身读取失败则为STAGE_FAILED/SOURCE_STABILITY。

辅助入口只能返回Boolean，不能产生绑定结果。它允许确定性测试，不形成另一个可注入外部审核结果的主入口。前后比较用于当次稳定性，未改变C7.1忽略长期位置的规则。独立运行之间正常移动/排序仍按C7.1保持长期batch一致。

此保护是规范要求的前后观测检查，不是Excel事务锁或恶意并发认证；不承诺检测两次观测之间发生后又恢复的变化，也不证明返回之后源文件仍未改变。未修改Excel事件/计算策略、未调用DoEvents、timer或后台线程。

## 测试与文件

- 新增：`src/modVATStage2BoundAuditRun.bas`、`tests/modVATStage2BoundTests.bas`、`tests/stage2-bound-audit-run-test-results.txt`、本检查点。
- 修改：`tools/build.ps1`只新增`-BoundAuditRunTestOnly`参数、加载冻结依赖的分支及release存在检查豁免；STATUS.md、STAGE2_SPEC.md更新；Stage1测试日志按本次回归生成。
- C7.4：**171/171 PASS**，真实Excel执行。覆盖VERIFIED/REVIEW/UNAVAILABLE、3类真实扫描阻断、同次short资格及warning失效、fallback补全、两个不同VERIFIED批次摘要分离、全部绑定digest/数量与冻结独立构建一致、移动/排序、零记录、13类header错误及已关闭Workbook、自然触发SHA阶段失败不发布部分绑定。
- 稳定性辅助覆盖同值、金额/身份/完成状态/formula标志变化、ExcelRow、数量、顺序及非法输入；同批重排即使长期Combined相同也被拒绝。生产SOURCE_CHANGED并发窗口未人为制造，按规格记录为代码级guard，纯内存比较已确定性验证。
- 只读检查覆盖值、公式、填色、Saved=True/False、工作簿数量/保持打开及无保存路径、Calculation/EnableEvents不变。重复调用的全部FinalDecision字段、Findings及成员结构和digest确定；无外部FinalDecision/Fingerprint主入口；无baseline/seal写入或历史读取。
- 二十组冻结回归：IncrementalDelta288、BaselineStore139、Fingerprint383、FinalDecision172、AuditFindings194、AggregateGate170、ShortSuffixPolicy139、EffectiveAmount200、EffectiveRelations135、FullFallback132、CompletedScope73、Snapshot81、GroupAmount44、Amount48、BConflict41、BRowMatch34、AIntegrity29、Matcher38、Parser70、Stage1 89，**2499/2499 PASS**。合计**2670/2670 PASS**。
- 冻结业务模块与旧测试源码未改；release四文件SHA-256保持，未重建或接入release。测试只创建并关闭自己的未保存工作簿。

## 后续与限制

C7.5 Seal只能接受受控入口产生的BoundRun，并进一步检查OK + FINAL_VERIFIED，不能重新接受裸FinalDecision + 任意fp。当前结果不是密码学认证，不防恶意修改VBA内存UDT；目标是消除API层面错误拼接。

本轮没有Verified Seal、baseline/seal保存、reuse/skip执行、历史FinalDecision返回、部分freeze、dependency graph、UI、report或release。C7.5未开始，无已知功能阻塞。

接手顺序：STATUS.md → STAGE2_SPEC.md C7.4 → 本检查点 → 新模块/测试；按需查看C6.3和C7.1/C7.2冻结协议，不重新开发旧链。
