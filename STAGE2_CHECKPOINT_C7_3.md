# Stage 2 C7.3 检查点（2026-09-14）

## 接口与安全边界

`VATStage2CompareBaselineDelta(ByRef fp As VATS2FingerprintResult, ByRef baselineLoad As VATS2BaselineLoadResult) As VATS2IncrementalDeltaResult`

独立模块`src/modVATStage2IncrementalDelta.bas`，纯内存只读输入。不访问Excel或文件，不调用业务匹配。文件读取由调用方在本层外使用冻结C7.2入口完成。本层逐条调用冻结`VATStage2SHA256`取得当前record digest，调用冻结`VATStage2BuildBaseline`仅取得当前三项批次摘要，不保存该内存state。

**digest-only baseline不能可靠区分new/change/remove。** 当前只能是UNCHANGED或NEW_OR_CHANGED；旧剩余只能解释为REMOVED_OR_CHANGED。没有旧新记录配对、业务身份猜测，也没有freeze或skip能力。UNCHANGED只证明摘要出现次数可在旧多重集中消费，不证明该条或其关联关系可以跳过验证。

## 结果与状态

Status（`VATS2_DELTA_`前缀）：OK=0、FIRST_RUN=1、BASELINE_UNAVAILABLE=2、INVALID_INPUT=3、INVALID_CONTRACT=4。

- 来源BASELINE_OK：验证必要内存契约后比较。
- 来源NOT_FOUND：FIRST_RUN，当前全部NEW_OR_CHANGED、residual为0、三项BatchUnchanged均False。这是没有旧证明的正常首次运行，不是业务异常，不自动保存baseline。
- 来源CORRUPT/IO_ERROR/UNSUPPORTED_VERSION：BASELINE_UNAVAILABLE，保留BaselineSourceStatus与来源ErrorReason（空原因使用固定说明），不返回记录分类数组或计数。不能冒充首次运行；后续调用方需要完整验证，本层不执行。
- 来源INVALID_INPUT、未知来源状态或当前Fingerprint非OK：INVALID_INPUT。
- 当前/旧数据必要契约损坏、摘要编码失败、批次与多重集矛盾：INVALID_CONTRACT。失败清空部分分类、剩余及批次相同标志，保留ErrorSide/ErrorIndex/ErrorReason和BaselineSourceStatus。

结果保存A/BRecordCount、ACurrent()/BCurrent()、A/BUnchangedCount、A/BNewOrChangedCount、A/BResidualCount、A/BResidualGroupCount、AResiduals()/BResiduals()、ABatchUnchanged、BBatchUnchanged、CombinedUnchanged及错误定位。所有有效输出数组严格1-based，零数量未分配。

Current项：CurrentIndex、Digest、State。State为VATS2_DELTA_UNCHANGED=0或VATS2_DELTA_NEW_OR_CHANGED=1；必须先判断主Status才使用分类。保持当前FP原顺序；不保存旧Index/ExcelRow/地址。

Residual项：Digest、OccurrenceCount；语义统一REMOVED_OR_CHANGED。ResidualCount是剩余**出现次数总数**，ResidualGroupCount是不同digest的组数；数组长度为组数，按二进制digest升序排列，不带旧记录位置。

## 多重集与契约

A只消费旧A，B只消费旧B。旧数组每段相同digest建立独立剩余次数预算，当前记录按原序二分查找相同段，每次最多消费一次。预算耗尽后的重复记录标为NEW_OR_CHANGED。旧剩余按digest聚合，保留完整multiplicity；不Set去重。额外内存O(N+M)，预算建立O(N)，当前检索O(M log N)，不计冻结SHA/批次构建开销。

必要契约：当前Fingerprint OK、计数非负、数组精确1..count或零时未分配、record/batch/combined签名非空。旧来源OK时内层state必须OK，版本必须VATBASE1/VATFP1/SHA256-UTF8，计数/数组边界合法，record digest严格64位小写hex且已排序，三项批次摘要合法。不重写文件解析器、时间戳验证、校验和或VATFP1内部解析。

三项BatchUnchanged使用当前与旧对应digest的二进制完全比较。各侧BatchUnchanged必须等价于该侧NEW_OR_CHANGED=0且residual=0；CombinedUnchanged必须等价于两侧均完全相同。正反任一矛盾均拒绝，不发布局部结果。

## 文件与测试

新增：本检查点、`src/modVATStage2IncrementalDelta.bas`、`tests/modVATStage2DeltaTests.bas`、`tests/stage2-incremental-delta-test-results.txt`。

修改：`tools/build.ps1`最小新增`-IncrementalDeltaTestOnly`参数/测试分支及release存在检查豁免；`STAGE2_SPEC.md`追加C7.3；`STATUS.md`更新入口；Stage1日志按原生回归重新生成。未修改冻结源码、旧测试语义、C7.1/C7.2协议或release。

- C7.3真实Excel原生测试：**288/288 PASS**。
- 覆盖完全相同、当前映射换序/位置改变、10类业务内容和标志变化、A/B增减、空侧/空批次、重复2→3及3→2/0、X/Y/X换序、X预算耗尽但Y仍匹配、跨侧相同digest不消费、三种不可用来源、FIRST_RUN、28类必要契约破坏、6类双向批次矛盾、全部输入字段不改、重复调用完整结果确定。
- 不创建文件：Delta入口调用前后只读检查项目/默认baseline等相关目录文件清单、大小和修改时间不变；同时检查新业务模块无文件/Excel访问及Save/Load调用。测试未读业务文件内容或保存任何baseline。测试报告由既有build链在入口返回后生成。
- 十九组冻结回归：BaselineStore139、Fingerprint383、FinalDecision172、AuditFindings194、AggregateGate170、ShortSuffixPolicy139、EffectiveAmount200、EffectiveRelations135、FullFallback132、CompletedScope73、Snapshot81、GroupAmount44、Amount48、BConflict41、BRowMatch34、AIntegrity29、Matcher38、Parser70、Stage1 89，**2211/2211 PASS**。
- 总计**2499/2499 PASS**；44个冻结源码、旧测试和release文件SHA-256保持。未修改Excel全局安全或计算策略，未重建release。

## 接手入口与限制

下一阶段按需读取STATUS.md → STAGE2_SPEC.md C7.3 → 本检查点 → 新模块和测试；摘要/存储契约见C7.2，canonical语义见C7.1。

无已知功能阻塞。摘要相等不构成业务验证或依赖关系不变的证明，仍受冻结VATFP1同环境稳定性与SHA摘要语义限制。本层不判断freeze eligibility、不跳过Parser/Matcher/金额、不自动保存baseline、不做依赖传播、orchestrator、UI、report或release。**C7.4未开始。**
