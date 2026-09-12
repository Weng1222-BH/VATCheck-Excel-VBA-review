# Stage 2 C3.1 检查点（2026-09-11）

## 接口与 Decimal 策略

```vb
VATStage2CompareAmounts(ByRef aIndexes() As Long, ByRef aAmounts() As Variant, _
    ByVal aCount As Long, ByVal bAmount As Variant) As VATS2AmountResult
```

独立纯内存金额引擎，无 Excel 对象、文件读写或号码模块依赖。上层决定哪些 A 属于本组，并保证 A 金额来自 `有效抵扣税额*`、B 来自 `税额`；本层不查字段、不自动找 A、不判断组完整性。

两个数组分别从实际 LBound 开始读取 aCount 个元素，按相对位置对应。原 AIndex 的0/负/正值、重复及顺序全部保留，不重排或去重。

沿用 Stage 1 `VATTryAmount` 类型白名单及运算方式：Byte/Integer/Long/Single/Double/Currency/Decimal 输入经 CDec，随后用 Variant/Decimal 累计和相减。所有成功输出金额均为 vbDecimal，不使用 Double 作为最终比较依据，不设 epsilon、分位容差或自动舍入。

**Difference = AAmountSum - BAmount**；直接与 `CDec(0)` 比较。返回 `VATS2AmountResult`：ACount、AIndexes()、AAmounts()、AAmountSum、BAmount、Difference、AmountState、Status、ErrorSide、ErrorIndex、ErrorAIndex、ErrorReason。

成功数组均为1-based，AAmounts 为逐项转换后的 Decimal。成功时清空错误字段。调用方必须先检查 Status 再使用金额结果。

## AmountState 与输入异常

| AmountState 常量 | 值 | 含义 |
|---|---:|---|
| VATS2_AMOUNT_NOT_COMPARED | 0 | 未形成有效比较 |
| VATS2_AMOUNT_EQUAL | 1 | Decimal 差额精确为0 |
| VATS2_AMOUNT_MISMATCH | 2 | Decimal 差额非0 |

EQUAL 只代表数学相等，不代表号码匹配、组完整性、颜色、冻结条件或审核通过。

| Status 常量 | 值 | 含义 |
|---|---:|---|
| VATS2_AMOUNT_OK | 0 | 比较完成 |
| VATS2_AMOUNT_EMPTY_GROUP | 1 | aCount=0，不比较 |
| VATS2_AMOUNT_INVALID_RANGE | 2 | 负数、任一数组容量不足或非零时未分配 |
| VATS2_AMOUNT_INVALID_AMOUNT | 3 | 非法类型或 CDec 转换失败 |
| VATS2_AMOUNT_ARITHMETIC_ERROR | 4 | Decimal 累加或差额计算失败 |

拒绝所有文本（包括数字文本、空字符串）、Empty、Null、Boolean、Excel Error、Date 及其他不在白名单的类型。空金额不跳过、不当零；这是纯金额引擎的输入错误。

按 A 输入顺序处理，再处理 B，最后计算差额；首个错误定位规则：

- A 转换/累加失败：ErrorSide=`A`，ErrorIndex 为 aAmounts 实际下标，ErrorAIndex 为对应原 AIndex。
- B 转换失败：ErrorSide=`B`，两个索引字段为0。
- 差额失败：ErrorSide=`DIFFERENCE`，两个索引字段为0。
- ErrorReason 保留类型/转换或运算错误原因。索引为0本身不是无错误标识，必须结合 Status 和 ErrorSide。

任何非 OK 结果都不发布部分金额：ACount=0、两个结果数组未分配、三个汇总金额为 Empty、AmountState=NOT_COMPARED。范围/空组状态不额外推断金额有效性。

## 新增/修改文件

- 新增 `src/modVATStage2Amount.bas`、`tests/modVATStage2AmountTests.bas`、`tests/stage2-amount-test-results.txt`。
- 修改 `tools/build.ps1`：仅新增 `-AmountTestOnly` 独立加载分支及发布文件存在检查的测试豁免。
- `STAGE2_SPEC.md` 末尾追加 C3.1；本检查点新增；`STATUS.md` 最小更新进度、最新检查点及 C3.2 待授权。
- 六组冻结回归日志由原测试链重新写出；旧测试源码不变。

C1/C2 所有模块、Stage 1 源码及旧测试未改；release 未重建或接入。未提交 commit。

## 实际测试结果

本机真实 Excel 顺序执行，均 exit 0：

```powershell
./tools/build.ps1 -AmountTestOnly
./tools/build.ps1 -BConflictTestOnly
./tools/build.ps1 -BRowMatchTestOnly
./tools/build.ps1 -AIntegrityTestOnly
./tools/build.ps1 -MatcherTestOnly
./tools/build.ps1 -ParserTestOnly
./tools/build.ps1 -TestOnly
```

- Amount：**48/48 cases PASS**。
- BConflict：**41/41 cases PASS**。
- BRowMatch：**34/34 cases PASS**。
- AIntegrity：**29/29 cases PASS**。
- Matcher：**38/38 cases PASS**。
- Parser：**70/70 cases PASS**。
- Stage 1：**89/89 assertions PASS**。

新增测试覆盖1→1/多A→1B、正负零差额、0.1+0.2=0.3、最大 Decimal、大数差1、最小1e-28及非零差额、分以下精度、负数/零、所有允许类型、顺序与重复索引、不同下界和部分范围、空组/非法范围、A/B各种非法类型、转换/累加/差额溢出、中间非法 A 的原下标和 AIndex。每例重复调用并验证结果、Decimal 子类型以及输入值/类型未修改。

测试中为构造极大/极小 Decimal 使用 CDec 常量；这不放宽引擎拒绝数字文本的输入规则。未更改安全设置或执行策略。

## 已知限制与 C3.2 接手入口

无已知阻塞问题。Decimal 具有有限范围和精度，转换及运算能力与 Stage 1 相同，不能恢复 Excel/调用方此前已丢失的精度。没有任意精度扩展，也不验证参数对应的业务字段来源。运算失败整次不比较，不输出部分成功数据。

C3.2 接手先读 `STATUS.md` → `STAGE2_SPEC.md` → 本检查点，再读 `src/modVATStage2Amount.bas`、`tests/modVATStage2AmountTests.bas` 及新结果日志；号码关系接口按需读取 C2.2b/C2.2c 检查点。测试入口在 `tools/build.ps1`。

C3.1 已完成并停止；C3.2 待用户授权，未实现金额门控、号码组完整性、颜色、fallback、baseline 或 UI。
