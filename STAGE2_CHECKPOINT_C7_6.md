# Stage 2 C7.6：Verified State Commit Coordinator

## 接口与范围

新模块 `src/modVATStage2VerifiedStateCommit.bas`，只协调冻结构建与存储入口，不修改C7.2、C7.4、C7.5或其它冻结模块。

- `VATStage2BuildBaselineForBoundRun(ByRef run As VATS2BoundAuditRunResult, ByRef fp As VATS2FingerprintResult) As VATS2BoundBaselineResult`
- `VATStage2CommitVerifiedState(ByRef run As VATS2BoundAuditRunResult, ByRef fp As VATS2FingerprintResult, Optional rootOverride As String = "") As VATS2VerifiedStateCommitResult`

辅助入口只在内存中复用`VATStage2BuildBaseline(fp)`，核实fp成功、BoundRun成功、VATRUN1及指纹/摘要协议一致，然后逐项二进制比较A/BRecordCount和A/B/CombinedDigest。跨批fp立即INVALID_CONTRACT；无效主状态INVALID_INPUT。失败不返回部分baseline。该辅助只证明批次绑定，不授予VERIFIED、不写文件。

Commit首先要求BoundRun OK、FinalDecision OK且VERIFIED；再调用绑定辅助与冻结`VATStage2BuildVerificationSeal(run)`，两者均成功才允许开始写入。不能传入裸FinalDecision或未经此入口内部检查的baseline/seal。REVIEW_REQUIRED、UNAVAILABLE、绑定错误及Seal构建失败均不调用任何Save。

## 顺序、结果及失败语义

只按 `SaveBaseline → SaveVerificationSeal` 顺序调用。baseline失败立即返回，seal存储入口不调用；baseline成功后seal失败，保留成功写入的新baseline，不回滚、不删除旧seal、不伪造新seal、不增加撤销文件或隐式失效状态。没有直接文件I/O，也不先Load旧状态。

`VATS2_COMMIT_`状态：OK=0、INVALID_INPUT=1、BASELINE_FAILED=2、SEAL_FAILED_PARTIAL=3、INVALID_CONTRACT=4。

返回Status、BaselineStatus、SealStatus、BaselineWritten、SealWritten、ErrorStage、ErrorReason。Written默认False。BaselineStatus/SealStatus为Long，只记录各冻结Save的返回状态；`VATS2_COMMIT_NOT_ATTEMPTED=-1`表示没有调用该Save，不把尚未执行伪装为OK。输入/构建失败时两个存储状态都是-1。

ErrorStage：INPUT、BASELINE_BUILD、BASELINE_BINDING、SEAL_BUILD、BASELINE_SAVE、SEAL_SAVE。正常成功清空错误；保存失败保留冻结存储的原错误文本，构建失败仅返回固定不含业务原文的说明。意外异常保留已成功写入事实，不尝试补救写入。

## 用户确认的同批例外

**PARTIAL_COMMIT是“本次写入不完整”，不是“可信状态撤销”。** baseline成功、seal失败始终返回SEAL_FAILED_PARTIAL、BaselineWritten=True、SealWritten=False。是否仍具备复用资格完全由冻结C7.5根据磁盘最终baseline、seal与当前delta三方一致性决定，C7.6不调用资格层。

- 原seal属于X，新baseline是Y：C7.5因SEAL_BASELINE_MISMATCH要求FULL_VALIDATION_REQUIRED。
- seal不存在：C7.5因SEAL_NOT_FOUND要求FULL_VALIDATION_REQUIRED。
- 原baseline/seal已经合法证明X，再次完整VERIFIED的仍是相同X：baseline重写成功、新seal失败，旧seal仍完全匹配时，C7.5允许REUSE_ELIGIBLE。C7.6依然报告PARTIAL，不人为撤销原证明。

因此不需要跨两个文件的“伪原子事务”。单文件pending与同卷替换继续由冻结Save负责；双文件不一致会被C7.5拒绝，仍一致且有效的历史证明可以保留。OK仅表示本次两次Save均成功，不提供跨进程双文件隔离、回滚或硬件断电事务保证。

## 隐私与限制

Coordinator不访问Workbook、不保存canonical、号码、供应商或金额，不增加第三个状态文件，不自己序列化或计算SHA。只把核实过的state交给冻结存储；rootOverride原样传给两次Save。正常API路径验证绑定，继承C7.4/C7.5不认证恶意伪造VBA UDT的边界。

本轮没有运行时reuse分流、skip执行、历史审核结果返回、部分freeze、依赖图、UI、report或release。显式Commit可能写入本地两文件；不是后台自动持久化。不改变任何业务核对规则。

## 测试与文件

- 新增：本检查点、新模块、`tests/modVATStage2CommitTests.bas`、`tests/stage2-verified-state-commit-test-results.txt`。
- 修改：`tools/build.ps1`只新增`-VerifiedStateCommitTestOnly`参数及专用加载/测试分支；STATUS、SPEC更新。原baseline/seal的gitignore排除继续有效。
- C7.6真实Excel原生测试：**191/191 PASS**。包括真实VERIFIED提交、第二次不同批覆盖、重复提交、零记录、17类拒绝输入/契约、真实X运行+Y指纹、真实REVIEW/UNAVAILABLE不落盘、baseline pending失败不调用seal、不同批/同批/缺seal三种Partial、seal文件锁失败及显式重试。
- 两种关键Partial均通过真实Save/Load、冻结Delta及C7.5资格链验证；检查旧seal/他人pending保持、新baseline不回滚。直接读取两个文件排除业务canary及canonical，文件清单无第三个状态；输入全字段、Workbook值/公式/颜色及Saved状态不改。
- 二十二组冻结回归：VerificationSeal217、BoundAuditRun171、IncrementalDelta288、BaselineStore139、Fingerprint383、FinalDecision172、AuditFindings194、AggregateGate170、ShortSuffixPolicy139、EffectiveAmount200、EffectiveRelations135、FullFallback132、CompletedScope73、Snapshot81、GroupAmount44、Amount48、BConflict41、BRowMatch34、AIntegrity29、Matcher38、Parser70、Stage1 89，**2887/2887 PASS**。合计**3078/3078 PASS**。旧源码和测试语义未改，release四文件SHA-256保持。Stage1结果文件由本轮回归更新计时。

后续入口：STATUS → STAGE2_SPEC C7.6 → 本检查点 → 新模块/测试；按需读取C7.2、C7.4、C7.5。C7.6完成后停止，后续阶段需新的用户规格。

## Git收尾结果

- Private正式C7.6提交：`51fccfc9a070d085ebd05facb98cb03d2f6a184c`，`stage2: complete C7.6 verified state commit`。已push至origin/main；完成核验时为main、工作区clean、HEAD=origin/main。
- Public正式C7.6提交：`59c83106ce812bfe15dbce5248f3be556d92ecae`，`stage2: mirror C7.6 verified state commit`。已push至origin/main；完成核验时为main、工作区clean、HEAD=origin/main。
- Public二次安全扫描通过：102个白名单文件；路径已脱敏；无baseline/seal状态文件、凭据、业务数据或private Git历史。
- C7.6实现、测试、镜像和Git收尾全部完成。无已知阻塞。下一阶段尚未开始。
