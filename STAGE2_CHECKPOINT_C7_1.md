# Stage 2 C7.1 检查点（2026-09-13）

## 接口与范围

`VATStage2BuildFingerprints(a, b) As VATS2FingerprintResult`

输入为冻结A/B Snapshot，只读ByRef；新模块 `src/modVATStage2Fingerprint.bas` 纯内存，不调用Excel、业务引擎、文件系统、SHA、WinAPI、外部程序或网络。没有baseline、持久化、跳过、freeze/unfreeze、orchestrator、UI、report或release。

结果保存Status、ARecordCount/BRecordCount、AFingerprints()/BFingerprints()、ABatchFingerprint/BBatchFingerprint/CombinedFingerprint、ErrorSide/ErrorIndex/ErrorReason。记录数组保持当前输入1-based原顺序，零记录数组未分配。

Status（VATS2_FINGERPRINT_前缀）：OK=0、INVALID_INPUT=1（Snapshot非OK）、INVALID_CONTRACT=2（数量/数组/当前位置契约）、UNSUPPORTED_VALUE=3（未支持的Variant类型）、ENCODING_ERROR=4（编码/分配等错误）。失败清空所有部分记录及批次签名，只保留错误定位。

## 冻结编码协议 VATFP1

完整canonical signature，不是散列或加密。内容可被解析还原，后续持久化时仍属于业务数据，不可因名为Fingerprint就视为已脱敏。

`Pack(text) = CStr(Len(text)) & ":" & text`。长度是VBA UTF-16代码单元数量，不是字节数；字符串原样保留NUL、换行、冒号、分隔符、空格、大小写、前导零及Unicode。空内容仍编码为`0:`。

- Variant编码：`Pack(VarType的十进制整数) + Pack(payload)`。
- Record：依次Pack版本`VATFP1-RECORD`、Side、已编码身份Variant、已编码金额Variant、已编码完成Boolean、已编码身份HasFormula Boolean、已编码金额HasFormula Boolean。
- A身份为InvoiceDigitsRaw；B身份为SupplierTextRaw。仅这六项业务内容（含Side）进入签名。
- 明确排除AIndex/BIndex、ExcelRow、身份地址、金额地址，以及Snapshot表名/表头行等位置元数据。它们仍留在输入供当次映射。
- Batch：依次Pack版本`VATFP1-BATCH`、Side、记录数量，再逐项Pack排序后的完整record签名。
- Combined：依次Pack版本`VATFP1-COMBINED`、完整ABatch、完整BBatch，固定A在前B在后。

排序采用显式vbBinaryCompare稳定归并排序，O(n log n)，只操作字符串副本。相等指纹全部保留，不作Set去重；排序不改变AFingerprints/BFingerprints或输入记录。

| Variant类型 | payload |
|---|---|
| Empty / Null | 空payload，靠各自VarType区分 |
| String | 原字符原序 |
| Byte / Integer / Long | VBA原生整数十进制文本 |
| Boolean | True为1，False为0 |
| Single | 固定4字节小端十六进制（8字符） |
| Double | 固定8字节小端十六进制（16字符） |
| Currency | 原四位定点整数的8字节小端十六进制 |
| Date | CDbl保留底层日期序列值后按8字节编码，不按日期显示或秒数格式化 |
| Decimal | CStr直接转换完整十进制精度，不经过Double |
| Error | CStr(Error)原生错误编号文本，另带vbError类型，不当字符串/合法金额 |

二进制复制仅使用VBA LSet及固定纯数值UDT，不使用指针或Variant内存布局。Decimal/Error的原生文本依赖环境区域设置；稳定性承诺为同一Windows/VBA及区域设置，不承诺跨区域格式协议。若后续需要跨环境协议，须显式版本升级，不悄悄改变VATFP1。

对象/数组及未列出的类型（例如LongLong）明确拒绝，绝不读对象默认属性、不清洗或猜值。指纹层不判断金额是否合法，错误、文本、超财务范围数值也按各自类型编码。

## 必要契约

Snapshot.Status=OK；RecordCount非负；非空Records数组精确1..RecordCount，零计数必须未分配；当前AIndex/BIndex等于数组位置；ExcelRow为正且当前顺序严格递增。只验证当次映射，不把这些位置加入身份，也不复制整个Snapshot防御层。

行移动/排序后，调用方提供新的合法当前位置映射；业务内容不变的record签名保持相同，批次排序后仍相同。金额、身份原值/类型、完成状态及任何指定HasFormula标志改变都会参与变化检测。只保留HasFormula Boolean，不额外读取Snapshot没有提供的公式正文。

## 文件与测试

新增：本检查点、`src/modVATStage2Fingerprint.bas`、`tests/modVATStage2FingerprintTests.bas`、`tests/stage2-fingerprint-test-results.txt`。

修改：`tools/build.ps1`最小新增`-FingerprintTestOnly`参数/分支及测试豁免，仅加载Snapshot类型所在模块、Fingerprint和新测试；`STAGE2_SPEC.md`追加C7.1，`STATUS.md`更新入口；旧测试日志按原入口重新生成。

新测试覆盖位置/地址/表名改变不变、两侧换序映射、14类内容及标志变更、无歧义拼接、Unicode/NUL、15种Variant值两两区分、Double/Single最低有效位、Decimal最大值末位/第28位、Currency极值末位、Date低位、超财务范围、重复multiplicity、A/B区别、空输入、257条逆序、完整输入不变和重复调用确定、16类错误输入拒绝。额外固定VATFP1记录/批次及数值字节协议样本，避免未来无意改变协议。

| 测试入口 | 原生Excel结果 |
|---|---:|
| FingerprintTestOnly | 383/383 PASS |
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

合计2072项全部通过。39个冻结源码、旧测试和release文件SHA-256一致；没有改Excel安全或计算设置，没有重建release。无已知功能阻塞。

## 后续入口与停止边界

先读STATUS.md → STAGE2_SPEC.md的C7.1 → 本检查点 → 新模块及测试；Snapshot字段语义见C4.1。C7.2尚未开始，本轮未决定任何持久化格式、存储位置、跳过策略或首次运行规则。
