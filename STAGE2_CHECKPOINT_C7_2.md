# Stage 2 C7.2 检查点（2026-09-14）

## 范围与接口

- `VATStage2SHA256(text As String) As String`：纯VBA FIPS 180-4 SHA-256，UTF-8 **without BOM**，固定64位小写hex。支持中文、补充平面Unicode、NUL；孤立UTF-16代理项明确报错，不替换字符。无COM、WinAPI、网络、外部程序或临时文件依赖。
- `VATStage2BuildBaseline(fp As VATS2FingerprintResult) As VATS2BaselineState`：只读冻结C7.1结果，将每条record及三个batch/combined canonical分别摘要。记录摘要稳定归并排序，保留multiplicity；不保存索引、ExcelRow或地址。
- `VATStage2SaveBaseline(baseline, Optional rootOverride As String = "") As VATS2BaselineIOResult`。
- `VATStage2LoadBaseline(Optional rootOverride As String = "") As VATS2BaselineLoadResult`：含Status、Baseline、ErrorReason。

Baseline状态（`VATS2_BASELINE_`前缀）：OK=0、NOT_FOUND=1、INVALID_INPUT=2、CORRUPT=3、IO_ERROR=4、UNSUPPORTED_VERSION=5。错误时不返回部分baseline；不存在文件返回NOT_FOUND，是正常首次运行状态，不代表损坏，也不决定首次业务流程。

State含Status、SchemaVersion、FingerprintProtocol、DigestProtocol、A/BRecordCount、A/BRecordDigests()、A/BBatchDigest、CombinedDigest、CreatedUtc、UpdatedUtc。数组严格1-based，零计数未分配。UTC由Windows GetSystemTime取得，创建时两个时间相同；保存保留调用方提供的合法时间，不自动推进业务状态。

## 隐私与存储协议

**baseline保存SHA-256 digest，绝不保存原VATFP1 canonical。VATFP1原文永不落盘。** 构建时只在内存中计算摘要；保存接口仅接受版本、数量、排序后的64位小写hex及合法UTC时间，不能夹带业务原文。

默认路径：`%LOCALAPPDATA%\VATCheck\baseline-v1.dat`。仅接受本地绝对目录，缺少有效默认目录时返回错误，不回退到项目、工作簿或Personal.xlsb目录。rootOverride用于测试或显式本地目录；测试全部使用系统临时目录中新建的独立子目录并清理，未访问正式baseline。

版本固定：schema `VATBASE1`、fingerprint `VATFP1`、digest `SHA256-UTF8`。文件为ASCII（即UTF-8无BOM子集）、严格LF换行；字段不允许自由文本，故逐行编码无歧义。顺序：

1. VATBASE1、VATFP1、SHA256-UTF8，各一行。
2. ARecordCount、BRecordCount，各一行；非负十进制，无前导零。
3. CreatedUtc、UpdatedUtc，各一行，`YYYY-MM-DDTHH:MM:SSZ`，有效日期且Updated不早于Created。
4. ABatchDigest、BBatchDigest、CombinedDigest，各一行。
5. 按序写ARecordCount个A digest，再写BRecordCount个B digest；各一行，重复保留。
6. 上述完整正文（包含末尾LF）的SHA-256校验和，再一个LF；不得存在后续字节或空行。

严格验证版本、数量与实际行数、数组、摘要字符/长度/排序、UTC、校验和及EOF。可识别VATBASE其他版本或不支持的protocol返回UNSUPPORTED_VERSION；错误magic及其他损坏返回CORRUPT。不“尽量读取”。校验和用于发现损坏，不是认证签名。

## 安全保存

校验state → 序列化digest正文 → 自动创建目录 → 同目录`baseline-v1.pending.tmp`以禁止覆盖方式创建 → 写入并关闭 → 重新完整读取、校验并比较内容 → Windows `MoveFileExW`（REPLACE_EXISTING | WRITE_THROUGH）同卷rename替换。

不先删除正式文件，不使用跨卷copy/delete。临时创建/写入、复读或替换失败返回明确错误，旧baseline保留；只清理本次拥有的临时文件。已有pending文件阻止新写入且保留该文件，避免清除其他写者或崩溃遗留文件；正式baseline仍可读取。没有自动删除未知pending、锁超时、并发合并或恢复决策。

Windows API仅在存储模块负责UTC与文件替换，摘要算法仍完全纯VBA。以本地Windows文件系统为目标；未模拟断电、磁盘硬件故障或32位Office运行环境，不能把文件替换当作硬件级持久性保证。摘要不是加密/认证，不提供防恶意篡改或低熵原文枚举保护。

参考：[FIPS 180-4](https://csrc.nist.gov/pubs/fips/180-4/upd1/final)、[NIST公开向量](https://www.nist.gov/itl/ai/ai-standards-and-guidelines-group/nsrl-test-data)、[MoveFileExW](https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-movefileexw)。

## 文件与验证

新增：`src/modVATStage2Digest.bas`、`src/modVATStage2BaselineStore.bas`、`tests/modVATStage2BaselineTests.bas`、`tests/stage2-baseline-store-test-results.txt`、本检查点。

修改：`tools/build.ps1`仅新增`-BaselineStoreTestOnly`参数/分支及release存在检查豁免；`.gitignore`排除baseline文件；`STAGE2_SPEC.md`追加本阶段、`STATUS.md`更新入口。冻结业务模块、旧测试源码和release未改。

- C7.2：**139/139 PASS**，真实Excel原生执行。含空/abc/quick brown fox/百万a公开向量，独立.NET预期的中文/emoji/NUL及55/56/64/65字节边界；非法代理项；build/往返/重复次数/排序/位置/业务标志变化；临时创建和替换失败、读锁；17种结构破坏、截断/尾部垃圾/校验和、12种非法state；文件直接搜索全部原始业务canary及record/batch/combined原文，均不存在。临时目录已清理。
- 十八组冻结回归：Fingerprint383、FinalDecision172、AuditFindings194、AggregateGate170、ShortSuffixPolicy139、EffectiveAmount200、EffectiveRelations135、FullFallback132、CompletedScope73、Snapshot81、GroupAmount44、Amount48、BConflict41、BRowMatch34、AIntegrity29、Matcher38、Parser70、Stage1 89，**2072/2072 PASS**。
- 共**2211/2211 PASS**。未更改Excel全局安全或计算设置，未重建release。

## 接手与停止

C7.2不包含任何skip/freeze、首次运行业务、增量传播、orchestrator、UI、report或release。C7.3未开始。

下一阶段按需读取STATUS.md → STAGE2_SPEC.md C7.2 → 本检查点 → 两个新模块及新测试；canonical协议参见C7.1检查点。没有已知功能阻塞；上述崩溃pending及环境限制需保留，不得默认为已有自动恢复能力。
