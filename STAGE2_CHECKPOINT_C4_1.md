# Stage 2 C4.1 检查点（2026-09-11）

## 接口与记录

```vb
VATStage2ReadASnapshot(ByVal sheet As Worksheet, ByVal headerRow As Long, _
    ByVal invoiceColumn As Long, ByVal amountColumn As Long, _
    ByVal completedColor As Long) As VATS2ASnapshotResult

VATStage2ReadBSnapshot(ByVal sheet As Worksheet, ByVal headerRow As Long, _
    ByVal supplierColumn As Long, ByVal amountColumn As Long, _
    ByVal completedColor As Long) As VATS2BSnapshotResult
```

调用层提供已定位的 Worksheet、表头行、业务字段列和共享完成色。A金额列为有效抵扣税额*，B金额列为税额，本层不验证表头文本、不重新定位。

两结果均含 SheetName、HeaderRow、RecordCount、Records()、Status、ErrorRow、ErrorReason。

- A记录：AIndex、ExcelRow、InvoiceDigitsRaw、AmountRaw、IsCompleted、InvoiceCellAddress、AmountCellAddress、InvoiceHasFormula、AmountHasFormula。
- B记录：BIndex、ExcelRow、SupplierTextRaw、AmountRaw、IsCompleted、SupplierCellAddress、AmountCellAddress、SupplierHasFormula、AmountHasFormula。
- 两原值字段为 Variant，直接保存 Value2。无 Trim、数值转换、文本修复、Parser 或金额引擎调用。其他字段均为 Long/String/Boolean，不含 Excel 对象。地址为无美元符号的当前 Sheet 内地址，例如 I15；调用方仍需保留所属工作簿上下文。
- Status：VATS2_SNAPSHOT_OK=0、INVALID_SHEET=1、INVALID_HEADER=2、INVALID_COLUMN=3、READ_ERROR=4（后四项同样带 VATS2_SNAPSHOT_ 前缀）。先检查 Status；错误不返回部分记录，读取错误保留 ErrorReason 及 ErrorRow（区域读取失败时可为0）。合法空结果 RecordCount=0，Records 未分配。

## 范围、Index 与空行

从 HeaderRow+1 到 UsedRange 末行，与 Stage 1 行界一致。只读取两个业务字段和金额显示填色，不根据隐藏或筛选状态跳过。只有两个 Value2 均为 Empty、均无公式且金额没有完成状态时才跳过。

身份有值但金额空、身份空但金额有值、公式空串、完成色但两字段均空都保留。无关列内容/填色或非完成格式残留不制造记录。数组按256条增长，最后收缩为精确记录数量。

每个保留记录按原行序取得1..RecordCount索引，ExcelRow保存真实行号。例如 Index1/2/3 对应 Excel10/11/15，完成色为1与3时仍为1与3。模块不构造完成子集。已有记录改变颜色后索引不变；全空无公式行加上完成色会进入记录集，此类集合变化不承诺跨次索引稳定。

## 颜色与只读保证

Stage 1 没有单独公开的完成色谓词，因此新 adapter 实现相同条件：金额单元格 DisplayFormat.Interior，Pattern非无填充、ColorIndex非无填充、Color等于参数。严格区分无填充/白色，读取条件格式实际显示色。图案符合颜色时 IsCompleted=True，与 VATScan 的完成色计数相同；Stage 1 随后报图案金额异常，而本层只保留完成属性，不作金额合法性判断。

生产模块没有写入、保存、关闭、筛选、排序、强制重算或 UI 操作。当前未保存值可读，Value2错误/Boolean/文本/空串原样保存。所有 Excel 引用只存在于函数局部，关闭测试源工作簿后快照仍独立可读。

## 文件与测试

新增：

- `src/modVATStage2ExcelSnapshot.bas`
- `tests/modVATStage2SnapshotTests.bas`
- `tests/stage2-snapshot-test-results.txt`
- 本检查点。

修改：`tools/build.ps1` 仅增加 SnapshotTestOnly 开关、测试豁免及分支；`STAGE2_SPEC.md` 仅末尾追加 C4.1；`STATUS.md` 更新入口。八份冻结测试日志由原入口重新生成。

原生 Excel 实际运行，全部 exit 0：

| 命令 `./tools/build.ps1` 后的开关 | 结果 |
|---|---:|
| -SnapshotTestOnly | **81/81 assertions PASS** |
| -GroupAmountTestOnly | **44/44 PASS** |
| -AmountTestOnly | **48/48 PASS** |
| -BConflictTestOnly | **41/41 PASS** |
| -BRowMatchTestOnly | **34/34 PASS** |
| -AIntegrityTestOnly | **29/29 PASS** |
| -MatcherTestOnly | **38/38 PASS** |
| -ParserTestOnly | **70/70 PASS** |
| -TestOnly | **89/89 assertions PASS** |

Snapshot测试覆盖A/B完整记录、非连续行映射、变色不重编号、隐藏/筛选、原始类型及公式、空行/空金额/格式尾行、当前未保存修改、重复调用、256条数组扩容、非法和已关闭Worksheet、边界行列，以及源值/公式/格式/隐藏/筛选/Saved不变。6种颜色模式对照冻结VATScan，含条件格式与图案填充。测试仅使用自身创建的临时工作簿，不读真实业务文件。

21个已有源码/旧测试/release文件的执行前后 SHA-256 一致。没有更改安全设置、执行策略或冻结业务模块；没有重建 release、提交 commit。仓库目前文件仍为untracked，不能以普通git diff为空证明未变，已使用哈希核对。

## 已知限制与C4.2入口

无已知阻塞问题。索引只在完整快照内稳定，不是跨业务编辑的永久ID。调用层负责表头语义和工作簿上下文；本层不清洗数据、不判合并金额/图案金额异常、不刷新公式计算结果。读取当前Value2与DisplayFormat，调用方负责所需计算时机。

为保留尾部完成色空行，需要遍历 UsedRange 末行；异常膨胀的 UsedRange 仍可能增加读取时间，未通过固定上限截断。未增加后台监控、作用域匹配或任何后续阶段功能。

C4.2接手先读 `STATUS.md` → `STAGE2_SPEC.md` → 本检查点 → `src/modVATStage2ExcelSnapshot.bas` / `tests/modVATStage2SnapshotTests.bas`，再按实际授权查看C2/C3接口检查点。复验命令见上表，当前结果在 `tests/stage2-snapshot-test-results.txt`。

C4.1完成并停止；C4.2尚未开始。
