# Stage 2 C2.2a 检查点（2026-09-11）

## 接口与重复组

```vb
VATStage2ScanAIntegrity(ByRef aInvoiceDigits() As String, _
    ByVal aCount As Long) As VATS2AIntegrityResult
```

独立纯内存模块，不依赖 Excel、Parser 或 Matcher，不读写文件。使用 Windows 自带的 `Scripting.Dictionary`（晚绑定，无需勾选引用）及 Collection 保存索引；字典 CompareMode=vbBinaryCompare，只按完整字符串相等分组。号码不 Trim、不标准化、不转数值、不 suffix。

输入契约与 C2.1 相同：一维 String 数组，从实际 LBound 开始读取 aCount 个元素，支持任意下界。aCount=0 返回正常空结果，允许数组未分配。负数、超过容量或未分配且 aCount>0 返回 INVALID_A_RANGE。

`VATS2AIntegrityResult`：

- RecordCount：范围有效时等于 aCount，包含非法值；范围无效时为 0。
- DuplicateGroupCount / DuplicateGroups()。
- InvalidValueCount / InvalidAIndexes()。
- Status / Flags。

`VATS2ADuplicateGroup`：InvoiceDigits、OccurrenceCount、AIndexes()。只包含出现至少两次的合法 ASCII 数字号码；保留全部原数组索引。组按号码首次出现顺序，组内按 A 输入顺序。不会按第二次出现时间排序，也不依赖字典枚举顺序。

所有返回数组为 1-based；相应数量为 0 时数组未分配，调用方必须依据 Count 遍历。输入不会删除、合并或修改。

## 状态与 Flags

| 接口 | 常量 | 值 | 含义 |
|---|---|---:|---|
| Status | VATS2_A_SCAN_OK | 0 | 范围有效且扫描完成，数据问题另看 Flags |
| Status | VATS2_A_INVALID_RANGE | 1 | 无效数组范围，未扫描 |
| Flags | VATS2_A_FLAG_NONE | 0 | 无本模块报告的问题 |
| Flags | VATS2_DUPLICATE_A_INVOICE | 1 | 存在完整发票号码重复组 |
| Flags | VATS2_A_DATA_REVIEW | 2 | 存在空串或非 ASCII 数字，原索引在 InvalidAIndexes 中 |

两种 Flags 可以组合。非法值不猜测、不修复，也不作为发票号码进入重复组；相同非法文本重复仍只报告每个非法索引。Status=SCAN_OK 不代表无重复或数据有效，调用方必须同时读取 Flags。

## 新增/修改文件

- 新增 `src/modVATStage2AIntegrity.bas`。
- 新增 `tests/modVATStage2AIntegrityTests.bas`。
- 新增 `tests/stage2-a-integrity-test-results.txt`。
- 修改 `tools/build.ps1`，仅增加 `-AIntegrityTestOnly` 独立加载及日志分支、发布文件存在检查的测试豁免。
- `STAGE2_SPEC.md` 末尾追加 C2.2a，未重写冻结章节。
- 新增本检查点。
- 重跑并写出 `tests/stage2-matcher-test-results.txt`、`tests/stage2-parser-test-results.txt`、`tests/excel-test-results.txt`。

冻结的 Parser、Matcher、Stage 1 业务源码及全部旧测试源码未改。四个 release 文件未重建或修改。未提交 commit。

## 实际测试结果

本轮真实 Excel 中依次运行，全部 exit 0：

```powershell
./tools/build.ps1 -AIntegrityTestOnly
./tools/build.ps1 -MatcherTestOnly
./tools/build.ps1 -ParserTestOnly
./tools/build.ps1 -TestOnly
```

- AIntegrity：**29/29 cases PASS**。
- Matcher：**38/38 cases PASS**。
- Parser：**70/70 cases PASS**。
- Stage 1：**89/89 assertions PASS**。

新测试覆盖无重复、两次/三次、不同重复组、不连续、前导零、相同尾号不重复、任意下界、部分有效范围、空/非法数组范围、首次出现与第二次出现顺序不同、非法值与重复共存、长号码及控制字符。每例都核对完整组内容和数量、原索引顺序、数组容量、所有状态与 Flags、重复调用确定性和输入未修改。

运行未修改 Excel 安全设置或执行策略，未重建 release。

## 已知限制与 C2.2b 接手入口

无已知阻塞问题。这里只检查 ASCII 数字格式及完整字符串重复，不验证真实发票存在性、固定长度、校验码或业务有效性。需要 Windows Scripting.Dictionary；本机原生测试已验证可用。所有扫描数据都在内存中。

后续获得 C2.2b 授权后先读取本检查点、`STAGE2_SPEC.md`、`src/modVATStage2AIntegrity.bas`、`tests/modVATStage2AIntegrityTests.bas` 及其结果日志；需要 Matcher 时读取 `STAGE2_CHECKPOINT_C2_1.md`。测试入口见 `tools/build.ps1`。

未实现 B 调度、跨 B 检测、B 多引用处理、金额、颜色、分组、fallback、baseline/fingerprint、UI 或 release 接入。C2.2a 已完成并停止，未开始 C2.2b。
