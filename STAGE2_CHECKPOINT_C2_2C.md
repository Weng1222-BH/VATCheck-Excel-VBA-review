# Stage 2 C2.2c 检查点（2026-09-11）

## 接口与结构

```vb
VATStage2ScanBConflicts(ByRef bRows() As VATS2BRowMatchResult, _
    ByRef bRowIds() As Long, ByVal bCount As Long) As VATS2BConflictResult
```

仅扫描已产生的行结果，无 Parser/Matcher 调用，无 Excel 对象或文件读写。内存分组使用晚绑定 Scripting.Dictionary 和 Collection；无需额外 VBA 引用。

两个输入数组可各有不同下界，从各自 LBound 开始按相对位置读取 bCount 个元素。BIndex 是 bRows 原下标，BRowId 是对应外部标识；同一/不同 B 输入记录仅按 BIndex 区分，重复 RowId 不隐藏冲突。

`VATS2BAssociation`：BIndex、BRowId、ReferenceIndex、ReferenceDigits、MatchKind、AIndex、AInvoiceDigits。仅 EXACT_UNIQUE / SUFFIX_UNIQUE 产生关联，其他结果不参与。INCOMPLETE 等行状态和所有风险 Flags 均不阻断逐引用唯一关联。

`VATS2BConflictGroup`：AIndex、AInvoiceDigits、AssociationCount、DistinctBRowCount、Associations()、Flags。按 AIndex 分组，只输出至少两个关联的组；保留全部关联。组按 AIndex 首次遇到顺序，关联按 B 输入顺序及行内引用顺序。

`VATS2BConflictResult`：BRecordCount、UniqueAssociationCount、ConflictGroupCount、ConflictGroups()、Status、ErrorBIndex、ErrorReferenceIndex。

UniqueAssociationCount 包括单次、无冲突的唯一关联；这些单次关联不另设输出表。结果数组1-based，零组未分配，按 Count 遍历。

## Flags 与状态

组 Flags（独立位空间）：

- `VATS2_BCONFLICT_NONE=0`。
- `VATS2_SAME_B_ROW_REUSE=1`：同一 BIndex 至少两个引用唯一指向同一 AIndex。
- `VATS2_CROSS_B_ROW_REUSE=2`：至少两个不同 BIndex 唯一指向同一 AIndex。
- 两种情况并存时 Flags=3，不消歧、不决定最终审核。Parser 已去重的单个引用不因重复风险而产生第二个关联。

Status：

- `VATS2_BCONFLICT_OK=0`：扫描完成；须查看 ConflictGroupCount，OK 不代表无冲突或业务审核通过。
- `VATS2_BCONFLICT_INVALID_RANGE=1`：负 bCount、任一数组容量不足或非零但未分配。BRecordCount=0。
- `VATS2_BCONFLICT_INVALID_CONTRACT=2`：引用数组契约错误；UNIQUE 候选数量不为1、数组未分配/边界错误；或同一 AIndex 的完整号码不一致。范围有效时 BRecordCount=bCount，首个异常定位到 ErrorBIndex / ErrorReferenceIndex；引用容器异常的 ReferenceIndex=0。

bCount=0 正常返回0关联、0冲突，无需分配输入数组。接口异常清空关联及冲突计数，不让部分结果冒充完整扫描。非契约运行时错误不会被吞掉。

## 新增/修改文件

- 新增 `src/modVATStage2BConflict.bas`。
- 新增 `tests/modVATStage2BConflictTests.bas`。
- 新增 `tests/stage2-b-conflict-test-results.txt`。
- 修改 `tools/build.ps1`：仅增加 `-BConflictTestOnly` 独立测试分支和发布文件存在检查的测试豁免。
- `STAGE2_SPEC.md` 末尾追加 C2.2c。
- `STATUS.md` 最小更新完成阶段、最新检查点及下一阶段待授权。
- 新增本检查点；五组旧回归结果 txt 由原测试链重新写出。

四个冻结 Stage 2 模块、Stage 1 源码、旧测试源码均未修改；release 未重建或接入。未提交 commit。

## 实际测试结果

本机 Excel 原生依次运行，全部 exit 0：

```powershell
./tools/build.ps1 -BConflictTestOnly
./tools/build.ps1 -BRowMatchTestOnly
./tools/build.ps1 -AIntegrityTestOnly
./tools/build.ps1 -MatcherTestOnly
./tools/build.ps1 -ParserTestOnly
./tools/build.ps1 -TestOnly
```

- BConflict：**41/41 cases PASS**。
- BRowMatch：**34/34 cases PASS**。
- AIntegrity：**29/29 cases PASS**。
- Matcher：**38/38 cases PASS**。
- Parser：**70/70 cases PASS**。
- Stage 1：**89/89 assertions PASS**。

新测试覆盖所有指定场景：两种复用及并存、三 B 共用、多组首次顺序、组内顺序、全部非 UNIQUE 排除、不完整行及风险不阻断、Parser 去重、独立数组下界、部分范围、空/非法输入、相同 RowId 不合并、所有关联字段保留。另测接口候选/引用数组异常、晚发异常不保留部分计数、同 AIndex 号码不一致、不同 AIndex 不按号码合并及负 AIndex。每例重复扫描并检查完整结果；输入全部字段及 ID 的测试快照保持不变。

测试输入主体由真实 BRowMatch 产生；接口异常和索引映射边界仅修改自建测试数据，不修改冻结模块。未更改安全设置或执行策略。

## 已知限制及后续入口

无已知阻塞问题。调用方须提供来自同一 A 索引空间的 BRowMatch 结果，A 数据完整性仍由 AIntegrity 负责。本模块不重验号码匹配算法、不解释外部 RowId、不分类 A 自身重复、不从 MULTIPLE 选取候选。冲突结构留给后续处理，不实现金额、颜色、fallback、最终分类、baseline、审核流程或 UI。

接手先读 `STATUS.md` → `STAGE2_SPEC.md` → 本检查点，再按需读新模块、新测试及其结果。输入契约见 `STAGE2_CHECKPOINT_C2_2B.md`，测试接线见 `tools/build.ps1`。不读取归档交接替代当前规格。

C2.2c 已完成并停止，未开始 C3；下一阶段等待用户授权。
