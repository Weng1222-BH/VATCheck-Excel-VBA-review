# Stage 2 C3.2 检查点（2026-09-11）

## 接口及结果

```vb
VATStage2EvaluateBGroupAmount(ByRef bRow As VATS2BRowMatchResult, ByVal bIndex As Long, _
    ByRef conflicts As VATS2BConflictResult, ByRef aAmounts() As Variant, _
    ByVal bAmount As Variant) As VATS2GroupAmountResult
```

独立纯内存协调器，无 Excel 对象、文件读写或 UI。检查当前关系及冲突，满足条件才组装金额调用冻结 `VATStage2CompareAmounts`；不复制 Parser、Matcher 或 Decimal 运算。

结果保留 BIndex、SourceMatchState、SourceConflictStatus、ReferenceCount 及四类引用计数摘要、ParserFlags、MatcherFlags、SameBRowReuse、CrossBRowReuse、GroupState、Reasons、ErrorReferenceIndex、ErrorAIndex、AmountEvaluated，以及嵌套完整 `Amount As VATS2AmountResult`。

不复制原候选生成第二套真源；调用方应继续保留传入的 bRow，按 BIndex 查看成功/未找到/多候选的完整来源。输入始终不修改。

## GroupState 与不可比较原因

| 常量 | 值 | 含义 |
|---|---:|---|
| VATS2_GROUP_NOT_COMPARABLE | 0 | 无引用、不完整关系或本行 A 重复关联，未调用 Amount |
| VATS2_GROUP_COMPARED | 1 | 调用 Amount 且 Status=OK，只有数学诊断结果 |
| VATS2_GROUP_AMOUNT_ERROR | 2 | Amount 返回非 OK，完整保留其错误 |
| VATS2_GROUP_INVALID_INPUT | 3 | 关系/冲突接口异常或 AIndex 金额映射无效，未比较 |

Reasons 独立位空间：NO_REFERENCES=1、INCOMPLETE_RELATION=2、INVALID_RELATION=4、SAME_ROW_A_REUSE=8、INVALID_CONFLICT=16、INVALID_A_MAPPING=32（VBA 名称均以 `VATS2_GROUP_` 开头）。允许组合，例如不完整且本行复用为10；接口异常优先返回 INVALID_INPUT。

NO_REFERENCES/INCOMPLETE 禁止使用部分 A 金额对完整 B 比较。检查来源计数、引用/候选数组、逐引用 MatchKind 和行状态的一致性；损坏或 BRowMatch.INVALID_INPUT 不作数学比较。

未调用引擎时 `AmountEvaluated=False`、嵌套 AmountState=NOT_COMPARED、Difference=Empty。嵌套 UDT 默认 Status=0 并不是执行成功证据，调用方必须先检查 GroupState 和 AmountEvaluated。

## 当前行冲突及风险传播

只检查 ConflictGroup.Associations 中与当前 BIndex 对应的关联：同组出现>=2次为当前行 SameBRowReuse；出现且组 DistinctBRowCount>=2为当前行 CrossBRowReuse。不能把组级 SAME 标志套给所有成员。检验当前关联与 bRow 的引用索引、匹配类型、AIndex 和完整号码一致。

同组 B1 有两个关联而 B2 只有一个：B1 same+cross 被阻断，B2仅cross可做诊断比较。其他 B 行的 SAME 不阻断当前行。另在本行唯一候选中检查 AIndex 重复，防止冲突组意外缺失时重复计数，不删除或合并引用。

ParserFlags 与 MatcherFlags 原样独立保存，不与冲突风险混为同一掩码。缺NO、短尾号、结构异常、数字尾注、Parser单格重复提示和跨B复用不会单独阻断诊断金额，且不因 AMOUNT_EQUAL 消失。本层不免除审核或自动通过。

## AIndex 映射与 Amount 调用

全部引用 UNIQUE、非空且本行无重复A时，按 Parser 原始顺序取得每个 Candidate.AIndex，直接读取 `aAmounts(AIndex)`。支持0、负数及任意数组下界；不重编号、排序、去重或按金额找记录。

先检查金额数组已分配及每个 AIndex 在真实边界内；越界返回 INVALID_INPUT，Reasons=INVALID_A_MAPPING，并保留 ErrorReferenceIndex / ErrorAIndex。未分配数组无具体引用位置，错误位置字段为0。

临时组数组为1-based，但其中 AIndexes 的值保持原标识。只调用一次冻结 Amount，完整保留合计、逐项金额、索引、差额方向（ΣA−B）、AmountState 或错误定位。Amount 的 ErrorIndex 是临时组内金额位置，ErrorAIndex 是原业务 AIndex；不擅自改写错误字段。

调用方负责同批关系/冲突快照及同一 AIndex 空间，金额 A=`有效抵扣税额*`、B=`税额`。本层不处理 AIntegrity 的全局 A_DATA_REVIEW，也不重扫 A 重复。

## 文件与实际验证

新增：`src/modVATStage2GroupAmount.bas`、`tests/modVATStage2GroupAmountTests.bas`、`tests/stage2-group-amount-test-results.txt`、本检查点。

修改：`tools/build.ps1` 最小增加 `-GroupAmountTestOnly` 分支及已有发布文件的测试豁免；`STAGE2_SPEC.md` 末尾追加 C3.2；`STATUS.md` 仅更新当前阶段/结果、最新检查点和下一阶段待授权。七组旧日志由原回归入口重新写出。

真实 Excel 顺序运行，全部 exit 0：

```powershell
./tools/build.ps1 -GroupAmountTestOnly
./tools/build.ps1 -AmountTestOnly
./tools/build.ps1 -BConflictTestOnly
./tools/build.ps1 -BRowMatchTestOnly
./tools/build.ps1 -AIntegrityTestOnly
./tools/build.ps1 -MatcherTestOnly
./tools/build.ps1 -ParserTestOnly
./tools/build.ps1 -TestOnly
```

- GroupAmount **44/44 cases PASS**。
- Amount **48/48 cases PASS**。
- BConflict **41/41 cases PASS**。
- BRowMatch **34/34 cases PASS**。
- AIntegrity **29/29 cases PASS**。
- Matcher **38/38 cases PASS**。
- Parser **70/70 cases PASS**。
- Stage 1 **89/89 assertions PASS**。

新测试覆盖本轮所有必测项，包括同组不同B门控差异、全部风险保留、阻断部分核算、金额错误原样传播、负/0/15索引、不同引用顺序、映射越界、损坏契约、缺失冲突组的本行重复防护、0.1+0.2=0.3、重复调用和所有输入不变。对比冻结 Amount 的完整输出及独立差额期望，确认未复制金额算法或改变 Difference 方向。

未修改冻结模块或旧测试；未更改安全设置或执行策略；未重建 release，未提交 commit。

## 已知限制与下一阶段入口

无已知阻塞问题。诊断金额不等于审核结论。调用方仍须保证冲突结果确实覆盖当前 bIndex 并与 bRow 属于同一批输入；接口没有批次指纹，本轮不新增 fingerprint。保留 Decimal 的既有范围/精度限制。

下一阶段等待用户授权。接手先读 `STATUS.md` → `STAGE2_SPEC.md` → 本检查点，再按需读新模块和新测试；金额错误契约见 C3.1，来源关系及冲突见 C2.2b/C2.2c 检查点。

C3.2 已完成并停止，未开始颜色、总体金额快速门、baseline、冻结、UI 或后续阶段。
