# Stage 2 C7.5 检查点：Verified Seal / Whole-Batch Reuse Eligibility

## 接口与绑定边界

新增独立模块 `src/modVATStage2VerificationSeal.bas`，不修改任何冻结模块。

- `VATStage2BuildVerificationSeal(ByRef run As VATS2BoundAuditRunResult) As VATS2VerificationSealState`
- `VATStage2SaveVerificationSeal(ByRef baseline As VATS2VerificationSealState, Optional rootOverride As String = "") As VATS2VerificationSealIOResult`。参数类型只接受Seal，文件与baseline独立。
- `VATStage2LoadVerificationSeal(Optional rootOverride As String = "") As VATS2VerificationSealLoadResult`
- `VATStage2EvaluateReuseEligibility(ByRef delta As VATS2IncrementalDeltaResult, ByRef baselineLoad As VATS2BaselineLoadResult, ByRef sealLoad As VATS2VerificationSealLoadResult) As VATS2ReuseEligibilityResult`

**Seal唯一构建入口只接受C7.4 BoundRun，禁止裸FinalDecision、裸fp或两者事后拼接。** 检查BoundRun OK、协议、数量、三项合法摘要、Final VERIFIED且无人工复核/差额/warning/Finding标志、Aggregate OK与Boolean True、两侧扫描成功且无问题/警告，以及Final与Aggregate/Audit必要摘要一致。REVIEW_REQUIRED、UNAVAILABLE或内部矛盾均拒绝，失败清空凭证字段。只核实金额摘要一致，不重新计算业务金额。

正常API路径防止跨批误拼接；不为恶意手工伪造VBA UDT提供密码学认证。调用方必须判断返回Status，不能把默认初始化的UDT当有效凭证。C6.3保持冻结。

## Seal协议、状态与文件

五个固定协议：`VATVERIFY1` / `VATRUN1` / `VATSTAGE2-POLICY1` / `VATFP1` / `SHA256-UTF8`。PolicyProtocol显式绑定审核规则；后续不兼容规则修改必须升级policy，不能使用Git SHA代替。

State仅含Status、五项协议、A/BRecordCount、VerifiedUtc、A/BBatchDigest、CombinedDigest。VerifiedUtc是生成凭证时的Windows UTC；同一state的重复保存/读取保持不变，独立Build调用时间可不同。

默认文件 `%LOCALAPPDATA%\VATCheck\verified-v1.dat`，独立pending为 `verified-v1.pending.tmp`；测试只用注入的系统临时子目录。严格ASCII/LF，依次11行正文：五协议、A数量、B数量、UTC、A摘要、B摘要、Combined摘要；第12行是冻结`VATStage2SHA256`对完整正文含LF的校验和，再终止LF。无后续字节。数量为无前导零非负Long，UTC为合法公历日期时间，摘要严格64位小写hex。

`VATS2_VERIFY_`状态：OK=0、NOT_FOUND=1、INVALID_INPUT=2、CORRUPT=3、IO_ERROR=4、UNSUPPORTED_VERSION=5。Load失败无部分可信Seal；UnsupportedField为SCHEMA/BINDING/POLICY/FINGERPRINT/DIGEST，用于结构化区分不支持的协议，不解析ErrorReason。

保存：validate → 禁止覆盖创建同目录pending → 写入关闭 → 完整复读校验及字节内容一致 → MoveFileExW同卷替换，绝不先删除旧文件。已有pending不抢占不删除；失败只清理本次拥有的pending，旧文件保持。只复制必要存储辅助，未改C7.2私有函数或摘要算法。

不保存canonical、号码、供应商、金额、行号、Findings、FinalDecision全文或Workbook路径。checksum是损坏检测，不是认证签名、加密或防恶意篡改机制。未模拟断电/磁盘硬件故障或32位Office；继承C7.2同环境指纹及本地Windows文件系统边界。

## 整批资格

`VATS2_REUSE_`状态：OK=0、INVALID_INPUT=1、INVALID_CONTRACT=2；Decision安全默认 `VATS2_FULL_VALIDATION_REQUIRED=0`，唯一允许值 `VATS2_REUSE_ELIGIBLE=1`。

**baseline相同不足以reuse，必须delta + baseline + seal三方一致。** delta来源必须对应传入baseline状态；正常OK时计数守恒、批次标志与变化数量一致。只有没有任何新增/变化或旧剩余、三项BatchUnchanged全True、baseline与seal均有效、五协议受支持、两侧数量及三项摘要完全一致时才允许资格。授予前进一步检查delta当前记录索引/状态/边界及摘要完整消费传入baseline的二进制多重集，不把另一份UNCHANGED delta误接到该baseline。保留重复次数，不比较Excel行号。

可组合ReasonFlags（`VATS2_REUSE_`前缀）：FIRST_RUN=1、BASELINE_UNAVAILABLE=2、SEAL_NOT_FOUND=4、SEAL_UNAVAILABLE=8、DATA_NEW_OR_CHANGED=16、DATA_REMOVED_OR_CHANGED=32、BATCH_CHANGED=64、SEAL_BASELINE_MISMATCH=128、POLICY_VERSION_MISMATCH=256。不支持POLICY通过UnsupportedField辨识，附带SEAL_UNAVAILABLE；这些不是Audit Finding。

首次/无Seal/损坏旧状态是需要完整验证的正常结果；无效接口契约明确非OK且仍FULL。任意一条变化均要求整批重验，不允许部分freeze。入口纯内存，不重新跑Parser/Matcher/金额或差异算法，不执行skip、不返回历史审核结果、不自动更新/删除baseline或seal。本轮**尚未解决双文件事务提交**，留待后续规格。

## 测试及文件

- 新增：本检查点、新模块、`tests/modVATStage2SealTests.bas`、`tests/stage2-verification-seal-test-results.txt`。
- 修改：`tools/build.ps1`最小新增`-VerificationSealTestOnly`及模块加载分支；`.gitignore`明确排除两个verified文件；STATUS与SPEC更新；Stage1日志随全回归重生。
- 测试只在合成Workbook和独立临时目录运行；直接读取保存文件检查业务canary、raw fingerprint、诊断文本及Workbook路径未落盘；baseline哨兵内容不变。
- 专用测试覆盖真实VERIFIED/REVIEW/UNAVAILABLE、38种Build契约破坏、全部版本与文件格式拒绝、pending争用、文件锁导致替换失败且旧Seal保持、完整输入不变性、排序/移动、变化/增减/重复次数及三方不一致。
- C7.5真实Excel原生测试：**217/217 PASS**。二十一组冻结回归：BoundAuditRun171、IncrementalDelta288、BaselineStore139、Fingerprint383、FinalDecision172、AuditFindings194、AggregateGate170、ShortSuffixPolicy139、EffectiveAmount200、EffectiveRelations135、FullFallback132、CompletedScope73、Snapshot81、GroupAmount44、Amount48、BConflict41、BRowMatch34、AIntegrity29、Matcher38、Parser70、Stage1 89，**2670/2670 PASS**。合计**2887/2887 PASS**。旧源码/旧测试无差异，release四文件SHA-256保持，未重建。

下一位接手：STATUS → STAGE2_SPEC C7.5 → 本检查点 → 新模块与测试；需要时再读C7.4受控入口、C7.2存储和C7.3delta契约。本阶段完成后停止，不开始后续orchestrator、skip、双文件联动、依赖图、UI、report或release。
