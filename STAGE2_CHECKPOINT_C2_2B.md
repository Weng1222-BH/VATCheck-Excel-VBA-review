# Stage 2 C2.2b 检查点（2026-09-11）

## BRowMatch 接口及结构

```vb
VATStage2MatchBRow(ByVal rawText As String, _
    ByRef aInvoiceDigits() As String, ByVal aCount As Long) As VATS2BRowMatchResult
```

纯内存组合层：先调用一次冻结 Parser，再按引用顺序逐个调用冻结 Matcher。不实现解析或匹配算法，无 Excel 对象、文件读写或 UI。A 数组和 aCount 原样传递，索引和有效范围契约沿用 C2.1。

`VATS2BRowMatchResult`：OriginalText、ParserFlags、ReferenceCount、References()、UniqueMatchCount、NotFoundCount、MultipleMatchCount、InvalidMatchCount、MatcherFlags、MatchState。

`VATS2BRowReference`：Digits、Length、RawFragment、StartIndex、Context、ParserFlags、MatchKind、MatcherFlags、CandidateCount、Candidates()。候选直接使用 `VATS2MatchCandidate`（AIndex、InvoiceDigits）。

引用按 Parser 原始顺序，候选按 A 原始顺序，保留全部候选及原数组下标；不挑选、不丢弃失败引用、不重新生成 Parser 已去重的引用。References 和 Candidates 均为 1-based，零数量未分配，按 Count 遍历。

## MatchState 和风险传播

| 常量 | 值 | 含义 |
|---|---:|---|
| VATS2_BROW_NO_REFERENCES | 0 | Parser 零引用，仍保留其风险；不调用 Matcher |
| VATS2_BROW_ALL_UNIQUE | 1 | 非零引用，全部 EXACT_UNIQUE / SUFFIX_UNIQUE |
| VATS2_BROW_INCOMPLETE | 2 | 有 NOT_FOUND 或 EXACT_MULTIPLE / SUFFIX_MULTIPLE，且没有接口错误 |
| VATS2_BROW_INVALID_INPUT | 3 | 有 INVALID_REFERENCE / INVALID_A_RANGE 或未知 Matcher 状态；优先于其他状态 |

四类计数对应 unique、not found、multiple、invalid，合计始终等于 ReferenceCount。未找到或歧义不会丢弃同一行其他引用的成功信息；exact 优先和 MULTIPLE 全候选由原 Matcher 决定。

- 行 ParserFlags 原样保留，逐引用 ParserFlags 对应 Parser 原返回值。
- 行 MatcherFlags 是逐引用 MatcherFlags 的 OR。
- 两个 Flags 位空间独立，不能合成一个风险掩码。
- 四种 Parser 风险及短尾号风险只传播，不改变唯一匹配的行状态。ALL_UNIQUE 不代表金额正确、自动通过或免审核。

## 新增/修改文件

- 新增 `src/modVATStage2BRowMatch.bas`。
- 新增 `tests/modVATStage2BRowMatchTests.bas`。
- 新增 `tests/stage2-b-row-match-test-results.txt`。
- 修改 `tools/build.ps1`：仅新增 `-BRowMatchTestOnly`，加载 Parser、Matcher、BRowMatch 及独立测试；增加发布文件存在检查的测试豁免。
- `STAGE2_SPEC.md` 末尾追加 C2.2b。
- `STATUS.md` 仅更新当前进度、下一阶段和最新检查点。
- 新增本检查点。
- 四组冻结回归的原结果 txt 由原测试链重新写出，旧测试源码未动。

Parser、Matcher、AIntegrity、Stage 1 业务源码及全部旧测试源码未修改；release 未重建或接入。未提交 commit。

## 实际测试

本机真实 Excel 依次执行，均 exit 0：

```powershell
./tools/build.ps1 -BRowMatchTestOnly
./tools/build.ps1 -AIntegrityTestOnly
./tools/build.ps1 -MatcherTestOnly
./tools/build.ps1 -ParserTestOnly
./tools/build.ps1 -TestOnly
```

- BRowMatch：**34/34 cases PASS**。
- AIntegrity：**29/29 cases PASS**。
- Matcher：**38/38 cases PASS**。
- Parser：**70/70 cases PASS**。
- Stage 1：**89/89 assertions PASS**。

新增用例涵盖六个用户示例及所有必测项，包括四种行状态、全部风险及组合、去重后传播、exact 优先、候选保序、非 1 下界、前导零、有效 A 范围、空输入和无效范围。每例均断言独立期望的逐引用类型/候选索引、四类计数守恒、完整原文、输入不变及重复调用确定性；并逐字段对照冻结 Parser/Matcher 输出，验证上下文、位置、号码和 Flags 未丢失。

无效 A 范围已真实验证为 INVALID_INPUT。INVALID_REFERENCE 在当前 Parser 的纯数字输出契约下不可达，因此未篡改冻结模块或增加注入入口来制造该状态；代码保留其 invalid 分类。运行时异常不被吞掉或伪装为 NOT_FOUND。

## 已知限制与下一阶段入口

无已知阻塞问题。正式调用方应先运行 AIntegrity 并处理 A_DATA_REVIEW，本层不校验 A 数据质量，也不正式分类 DUPLICATE_A_INVOICE。零引用时不调用 Matcher，因此不报告 A 范围错误；NO_REFERENCES 不表示 A 已通过完整性检查。

ALL_UNIQUE 仅逐引用判断，未验证不同引用是否指向同一 A；更不检测跨 B 行冲突。未实现金额、分组、颜色、fallback、baseline/fingerprint、审核决策、UI 或 release 接入。

下一阶段范围等待用户另行定义与授权。接手先读 `STATUS.md` → `STAGE2_SPEC.md` → 本检查点，再按需读取新模块和新测试；依赖接口见 C1.1、C2.1、C2.2a 检查点。不要读取归档旧交接来替代当前规格。

C2.2b 已完成并停止，未开始后续阶段。
