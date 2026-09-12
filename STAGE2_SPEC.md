# STAGE2_SPEC v1.0 — C1：B 表发票号码解析器

来源：2026-09-11 用户本轮指令。用户已确认以本条消息的 C1 规则为准；没有另行提供完整 Stage 2 文档。本文件不补写 C2 或后续功能规范。

## C1.1 风险元数据补充（优先于下文 C1 的旧 PARSE_REVIEW 表述）

仅细分风险来源并记录单个 B 单元格内部重复，候选提取、排除、保序、去重及首次位置不变，不决定人工审核方式：

- `NO_MARKER_MISSING`：无 NO 且已经根据末尾号码组规则成功提取候选；本身不表示结构异常。
- `STRUCTURE_REVIEW`：多个 NO、错误/未闭合括号、空分隔、非法拼接、小数/单横线、未知尾部结构等原结构问题。无 NO 但解析失败时不增加 NO_MARKER_MISSING。
- `NUMERIC_TRAILER_REVIEW`：保持原纯数字 `--` 尾注规则。
- `DUPLICATE_REF_REVIEW`：同一输入中成功解析的某个 Digits 重复出现；References 仍只保留第一次。被括号或尾注规则排除的数字不计为重复，不涉及跨 B 行引用关系。
- 四种 Flags 可组合，记录级标志继续传播给引用；接口和具体位值见 `STAGE2_CHECKPOINT_C1_1.md`。旧 `VATS2_PARSE_REVIEW` 仅保留为缺少 NO 与结构异常的组合兼容掩码。

## 范围与输入输出

实现独立、可单元测试、尽量纯函数化的 VBA Parser。输入为 B 表“供应商名称”单元格的原始 String。输出保留原始文本及 0～N 个疑似发票号码引用，每项包含纯数字 Digits、长度、原始片段/必要上下文、Flags。

Parser 只产生候选，不证明它是有效发票，不与 A 表匹配；不使用金额、供应商名称或其他业务字段猜号码。禁止宽松提取全文所有数字串，采用保守结构解析，无法确定时输出 `PARSE_REVIEW`。

## 必须支持的真实样例

| 序号 | 原文 | 期望 Digits（保持顺序） |
|---|---|---|
| 1 | 采购专用发票，广东省AAA有限公司（NO.）26578450000012121212 | 26578450000012121212 |
| 2 | 采购专用发票，南通BBB有限公司（NO.）151515151 | 151515151 |
| 3 | 采购专用发票，湖南省CCC有限公司（NO.）6868686868 | 6868686868 |
| 4 | 银（0000）付DD申请办公室EEE有限公司NO.26458900000001536415 | 26458900000001536415 |
| 5 | 采购专用发票，广东省TTT有限公司265784500000989898--lo5 | 265784500000989898 |
| 6 | 采购专用发票，广东省FFF有限公司265784500000989898(P0101204648221) | 265784500000989898 |
| 7 | 银（0000）付DD申请办公室EEE有限公司BC26458900000001536415 | 26458900000001536415 |
| 8 | 银（0000）付DD申请办公室EEE有限公司BC26458900000001536415（折让） | 26458900000001536415 |
| 9 | 采购专用发票，广东省AAA有限公司（NO.）26578450000012121212/65456454 | 26578450000012121212 / 65456454 |
| 10 | 采购专用发票，广东省AAA有限公司（NO.）26578450000012121212/26578450000004685792 | 26578450000012121212 / 26578450000004685792 |
| 11 | 采购专用发票，广东省AAA有限公司（NO.）12121212/54685487/979422543/9847412 | 12121212 / 54685487 / 979422543 / 9847412 |
| 12 | 采购专用发票，南通BBB有限公司（NO.）151515151/6451225/88512456 | 151515151 / 6451225 / 88512456 |

支持分隔符 `/`、`，`、`、` 及混合使用；支持 `NO.`、`NO:`、`NO：`、`（NO.）`。NO 可以不存在。

## 排除与风险规则

1. 中英文普通括号内容不作为号码，例如 `（0000）`、`(P0101204648221)`、`（折让）`；`（NO.）` 仅为标记，不影响后面的号码。
2. 英文字母前缀不是 Digits：`BC26458900000001536415` 提取后面的数字。
3. `--` 后包含英文字母时，整段尾注忽略，不能产生其中的数字候选。
4. `--` 后为纯数字时，不把尾注当正常号码；正常解析前面主号码，并增加记录风险 `NUMERIC_TRAILER_REVIEW`。
5. 保持多号码顺序；相同 Digits 不产生重复候选，保留首次出现位置。Digits 全程为 String，不作数字转换，保留前导零。
6. 无法保守确认的结构输出 `PARSE_REVIEW`。候选不等于已经验证的发票号码。

## 用户后续确认：结构解析与无 NO 风险提示

用户撤回了先前长度门槛，重新确认“采用结构解析＋无 NO 风险提示”：

- 取消 7–20 / 18–20 位硬门槛；长度仅如实记录，不决定是否为真实发票。
- NO 后符合号码组结构的纯数字项作为候选；允许英文字母前缀，Digits 不含前缀。
- 无 NO 仅解析文本末尾号码组，候选统一附带 `PARSE_REVIEW`，短号也不直接丢弃；长数字也不自动视为可靠号码。
- 括号及尾注先排除；不在全文中搜索所有数字。无 NO 且末尾组前还残留未排除的数字时，不挑选其中某一组，输出 `PARSE_REVIEW`。
- 多个独立 NO 标记、未配对括号、空分隔项、不支持的数字拼接/小数/尾部文字等输出 `PARSE_REVIEW`。可以保留同一明确号码组中已成功解析的项，风险传播到其 Flags；多个 NO 标记不擅自选择某一标记。
- 无 NO 的纯日期状数字也只能作为带风险的候选；Parser 无法仅凭数字文本区分日期、业务编码与票号，任何候选均不是确认结果。

## 实现与回归边界

- 新模块 `src/modVATStage2Parser.bas`、独立测试 `tests/modVATStage2ParserTests.bas`；只允许 build/test 的最小接线。
- 覆盖全部 12 个真实样例和空串、无数字、括号、前缀、NO 标点、中英文括号、各分隔符及混合、两类尾注、P 尾注、折让、顺序、去重等边界。
- Parser 测试全部 PASS；原 Stage 1 **89/89** 必须重新通过，旧测试不删除、不修改、不放宽。
- C1 不进入正式 release 功能，不改 Stage 1 业务规则。完成后写 `STAGE2_CHECKPOINT_C1.md` 并停止。

## 本轮禁止

不实现 A/B 号码匹配、完整号/尾号匹配、AMBIGUOUS 判断、1→1/多→1 分组、金额比较、完成颜色逻辑、全表 fallback、baseline/fingerprint、业务本地状态文件、Stage 2 UI/最终报告、手动匹配、OCR、Python、数据库或 Web。不得开始 C2。

## C2.1 — 单个 B 引用到 A 完整号码候选集合（2026-09-11）

本节为用户单独授权的 C2.1，前文禁止匹配属于 C1 当时的范围；C1/C1.1 的行为与风险接口继续冻结。本节不授权 C2.2 或后续阶段。

- 独立模块 `src/modVATStage2Matcher.bas`，不依赖 Parser、Excel 对象或其他业务字段。
- 输入为一个纯 ASCII 数字 String 引用及 A 完整号码 String 数组、有效元素数。号码不 Trim、不转数字、不规范化，保留前导零；不设引用最低长度。A 号码的数据质量由上游保证，本模块不扫描分类 A 数据异常。
- 首先逐项以 `StrComp(..., vbBinaryCompare) = 0` 收集 exact。存在任意 exact 时只返回所有 exact，绝不混入 suffix。
- 无 exact 时才比较 `Right$(A号码, Len(B引用))` 与 B 引用；A 短于 B 时不匹配。
- MatchKind 分为 NOT_FOUND、EXACT_UNIQUE、EXACT_MULTIPLE、SUFFIX_UNIQUE、SUFFIX_MULTIPLE。多个候选不选第一个、不消歧；重复的 A 输入保留各自索引。
- 结果保留 ReferenceDigits、ReferenceLength、MatchKind、CandidateCount、Candidates()、Flags。候选含 AIndex 和 InvoiceDigits，保持输入 A 顺序。
- 独立 Matcher 标志 SHORT_SUFFIX_REVIEW：非空 A 且没有 exact、实际执行 suffix 搜索时，引用长度 <=6 设置，>6 不设置。执行 suffix 后未找到候选也保留路径标志；空 A 未执行 suffix，不设置。所有 exact 均不设置。本标志不决定人工审核或风险豁免。
- 接口 `VATStage2MatchReference(ByVal referenceDigits As String, ByRef aInvoiceDigits() As String, ByVal aCount As Long) As VATS2ReferenceMatchResult`。
- A 数组为一维；从实际 LBound 起读取 aCount 个元素。AIndex 是原数组下标，支持 0、1 或其他下界。结果 Candidates 为 1-based；零候选未分配，按 CandidateCount 访问。
- 空/非 ASCII 数字引用返回 INVALID_REFERENCE。负 aCount、超过容量或未分配数组但 aCount>0 返回 INVALID_A_RANGE。aCount=0 允许未分配数组并返回 NOT_FOUND；B 引用合法性检查优先。这些状态不属于 A 数据质量分类。
- 独立测试接线 `tools/build.ps1 -MatcherTestOnly` 只导入 Matcher 及其测试；必须实际通过 Matcher 全测、冻结 Parser 70 项及 Stage 1 89 项，不删改旧测试语义。
- 必测 exact 唯一/多个及优先级、suffix 唯一/多个与长度边界、无匹配、B 更长、前导零、顺序、重复调用、空 A；补充输入及数组契约边界。
- 不实现多引用调度/分组、跨 B 冲突、DUPLICATE_REFERENCE、DUPLICATE_A_INVOICE 全表扫描、金额、颜色、fallback、baseline、fingerprint、冻结/解冻、UI 或 release 接入。完成实现、测试、C2.1 检查点后立即停止。

## C2.2a — A 完整号码完整性扫描与重复检测（2026-09-11）

本节为单独授权的纯内存 A 数据扫描；此前 Parser 与 Matcher 行为保持冻结，不进入 C2.2b。

- 独立模块 `src/modVATStage2AIntegrity.bas`，接口 `VATStage2ScanAIntegrity(ByRef aInvoiceDigits() As String, ByVal aCount As Long) As VATS2AIntegrityResult`。不依赖 Excel、Parser 或 Matcher，不读写文件。
- 从实际 LBound 开始扫描 aCount 个元素，支持非 1 下界；aCount=0 返回正常空结果，可使用未分配数组。负数、超容量、未分配但 aCount>0 返回明确 INVALID_A_RANGE（常量 `VATS2_A_INVALID_RANGE`）。
- 完整号码按二进制字符串完全相等判重，不 suffix、不转数值、不 Trim、不标准化。前导零保留，00123456 与 123456 不相同。
- 返回 RecordCount、DuplicateGroupCount、DuplicateGroups()、Status、Flags；每组包含 InvoiceDigits、OccurrenceCount、AIndexes()，保留全部原索引。
- 重复组按该号码第一次出现的顺序，组内按 A 原顺序；不按第二次出现时间排序，不修改输入。
- 发现重复设置 DUPLICATE_A_INVOICE。空字符串或非 ASCII 数字以 A_DATA_REVIEW 及 InvalidValueCount/InvalidAIndexes() 报告；非法值不作为发票号码参与重复分组，不修复或猜测。两种 Flags 可以同时存在。
- Status 表示范围是否有效、扫描是否正常完成；SCAN_OK 不代表数据无问题，调用方必须同时读取 Flags。合法范围 RecordCount=aCount（包含非法值），非法范围 RecordCount=0。所有输出数组为 1-based，零元素未分配。
- 独立测试接线 `-AIntegrityTestOnly`。覆盖无重复、两次/三次、多个组、不连续、前导零、相同尾号、任意下界、部分范围、空/非法范围、首次排序、组内顺序、确定性、输入不变和非法值报告；实际复验 Matcher 38、Parser 70、Stage 1 89 项。
- 禁止 B 调度、B 匹配调用、多引用/跨 B 检测、分组、金额、颜色、fallback、baseline/fingerprint、UI、release。完成 C2.2a 检查点后立即停止。

## C2.2b — 单个 B 行多引用解析与匹配汇总（2026-09-11）

本节单独授权组合冻结 C1/C1.1 和 C2.1，不修改 Parser、Matcher、AIntegrity 或 Stage 1。

- 接口：`VATStage2MatchBRow(ByVal rawText As String, ByRef aInvoiceDigits() As String, ByVal aCount As Long) As VATS2BRowMatchResult`。新模块 `src/modVATStage2BRowMatch.bas`，纯内存，无 Excel 对象、文件读写或 UI。
- 流程固定：原文调用 `VATStage2ParseBInvoices` 一次 → 按返回顺序逐引用调用 `VATStage2MatchReference` → 保留完整结果并汇总。不重新实现解析或匹配算法，不自动选择候选。
- A 数组及 aCount 原样交给 Matcher。正式调用前由调用方执行 AIntegrity 并处理 A_DATA_REVIEW；本模块不重扫、不清洗、不筛选 A。A 完整号码重复只保留 Matcher 的 EXACT_MULTIPLE，不在本层分类为 DUPLICATE_A_INVOICE。
- 每个引用保存 Digits、Length、RawFragment、StartIndex、Context、ParserFlags、MatchKind、MatcherFlags、CandidateCount、全部 Candidates（AIndex、InvoiceDigits）。保留 Parser 顺序、A 原索引及候选顺序，exact 优先于 suffix 和 MULTIPLE 语义不变。
- 行结果保存 OriginalText、ParserFlags、ReferenceCount、References()、UniqueMatchCount、NotFoundCount、MultipleMatchCount、InvalidMatchCount、MatcherFlags、MatchState。
- `NO_REFERENCES`：Parser 产生零引用，仍保留其 Flags；不调用 Matcher，因此此路径不判断 A 范围有效性。
- `ALL_UNIQUE`：非零引用且全部为 EXACT_UNIQUE / SUFFIX_UNIQUE，仅表示候选关系逐引用唯一，不代表金额正确、自动通过或免审核。
- `INCOMPLETE`：至少一个 NOT_FOUND / EXACT_MULTIPLE / SUFFIX_MULTIPLE，且无接口错误。保留每个成功及失败结果，不因部分不完整而丢弃已成功信息。
- `INVALID_INPUT`：存在 Matcher INVALID_REFERENCE / INVALID_A_RANGE 或未知接口状态，优先于其他行状态，不伪装成 NOT_FOUND。
- 四类计数分别归类 unique、not found、multiple、invalid；其和必须等于 ReferenceCount。
- ParserFlags 与 MatcherFlags 独立位空间；行 ParserFlags 原样保留，行 MatcherFlags 为全部引用 MatcherFlags 的 OR。缺 NO、结构异常、纯数字尾注、单格重复引用、短尾号均仅传播风险，不改变 unique 的分类，不决定人工审核。
- Parser 已去重的引用不重新添加；`--` 纯数字不作为引用。结果数组 1-based，零数量未分配，以 Count 遍历。
- 独立测试入口 `-BRowMatchTestOnly`，仅加载 Parser、Matcher、BRowMatch 及新测试。测试覆盖用户六个示例和全部要求，包括原文/上下文、计数守恒、风险并存、保序、前导零、非 1 下界、确定性、输入不变、无效 A 范围及空输入。旧 AIntegrity 29、Matcher 38、Parser 70、Stage 1 89 项必须实际 PASS。
- 不实现跨 B 行冲突、多 B 扫描、同 A 被多 B 引用、跨 B DUPLICATE_REFERENCE、金额/金额分组、颜色、fallback、baseline/fingerprint、冻结/解冻、UI、release、OCR、Python 或数据库。完成检查点和 STATUS 最小更新后停止，下一阶段另行授权。

## C2.2c — 已唯一匹配的 A 引用冲突扫描（2026-09-11）

本节单独授权扫描已生成的 BRowMatch 行结果；不修改前面四个 Stage 2 模块或 Stage 1，不调用 Parser/Matcher 重新匹配。

- 接口：`VATStage2ScanBConflicts(ByRef bRows() As VATS2BRowMatchResult, ByRef bRowIds() As Long, ByVal bCount As Long) As VATS2BConflictResult`。纯内存，无 Excel 对象或文件读写。
- 两个输入数组分别从其实际 LBound 开始读取 bCount 个元素，以相对位置对应。BIndex 保留 bRows 实际下标；BRowId 只作外部标识，不用于数组寻址或判断是否同一 B 记录。
- 只对 EXACT_UNIQUE / SUFFIX_UNIQUE 创建 association。NOT_FOUND、EXACT_MULTIPLE、SUFFIX_MULTIPLE、INVALID_REFERENCE、INVALID_A_RANGE 不参与，不从 MULTIPLE 选候选。行 MatchState 和 Parser/Matcher Flags 不过滤唯一关联。
- Association 保存 BIndex、BRowId、ReferenceIndex、ReferenceDigits、MatchKind、AIndex、AInvoiceDigits。
- 按 AIndex 分组，至少两个关联才输出冲突组；保留该组全部关联。组按 AIndex 首次遇到顺序，组内按 B 输入顺序及行内引用顺序。
- 组保存 AIndex、AInvoiceDigits、AssociationCount、DistinctBRowCount、Associations()、Flags。DistinctBRowCount 按不同 BIndex 计算，即使不同输入记录 RowId 相同也不能合并。
- 独立组 Flags：SAME_B_ROW_REUSE=1（同 BIndex 有至少两个引用）、CROSS_B_ROW_REUSE=2（至少两个 BIndex）；可同时设置为3。均只报告结构冲突，不自动消歧或决定最终审核。
- Parser 去重后只有一个引用的记录不会凭 DUPLICATE_REF_REVIEW 虚构第二个关联；A 完整号码自身重复仍由 AIntegrity 负责。
- 结果含 BRecordCount、UniqueAssociationCount（所有 UNIQUE 关联，包括无冲突的单次关联）、ConflictGroupCount、ConflictGroups()、Status。只输出冲突组，不另建无冲突关联结果表。
- bCount=0 正常空结果，允许输入数组未分配。负数、任一数组容量不足或非零但未分配返回 INVALID_RANGE，不伪装为无冲突。
- UNIQUE 必须 CandidateCount=1 且候选数组为1..1；引用数组数量/边界违反 BRowMatch 契约或同 AIndex 对应不同完整号码时，返回 INVALID_CONTRACT，并用 ErrorBIndex / ErrorReferenceIndex 指出首个错误（引用容器错误的 ReferenceIndex=0）。接口异常不输出部分关联计数或部分冲突组。范围有效时 BRecordCount=bCount，范围无效为0。
- 输出数组1-based、零组未分配。保留所有输入数据，不修改风险、计数、候选或 ID。
- 新测试入口 `-BConflictTestOnly` 加载 Parser、Matcher、BRowMatch、BConflict 和独立测试；新测试及 BRowMatch 34、AIntegrity 29、Matcher 38、Parser 70、Stage 1 89 项必须实际 PASS。
- 禁止金额/金额分组、颜色、fallback、A-only/B-only 最终分类、baseline/fingerprint、冻结/解冻、最终审核规则、UI、release。完成检查点和 STATUS 最小更新后停止，不进入 C3。

## C3.1 — 纯金额核对引擎（2026-09-11）

本节单独授权数学核对，不修改冻结 C1/C2 或 Stage 1，不判断号码关系、组完整性、冲突门控或业务审核。

- 新模块 `src/modVATStage2Amount.bas`，接口 `VATStage2CompareAmounts(ByRef aIndexes() As Long, ByRef aAmounts() As Variant, ByVal aCount As Long, ByVal bAmount As Variant) As VATS2AmountResult`。纯内存，无 Excel 对象或文件读写。
- 调用方提供已确定的 AIndex 及对应金额；A 金额唯一来源为 `有效抵扣税额*`，B 为 `税额`。本层只计算 ΣA 对 B，不按号码找金额或验证字段来源。
- 按两个数组各自 LBound 开始读取 aCount 个元素，按相对位置对应。保留 AIndex 原值、输入顺序及重复，0/负/正索引都允许，不去重或排序。
- 沿用 Stage 1 `VATTryAmount` 的类型白名单：Byte、Integer、Long、Single、Double、Currency、Decimal 输入经 CDec 转为 Variant/Decimal；拒绝所有文本（包括数字文本/空字符串）、Empty、Null、Boolean、Excel Error、Date 和其他非白名单类型。转换失败为 INVALID_AMOUNT，不猜值、不当0，也不跳过空金额。
- 所有金额输出、累计和差额均为 Variant/Decimal，累计使用 `CDec(total) + CDec(amount)`，差额固定 `CDec(AAmountSum) - CDec(BAmount)`。直接与 `CDec(0)` 比较；无 Double 最终比较、epsilon、分位容差或自动四舍五入。
- 成功返回 ACount、AIndexes()、AAmounts()（逐项 Decimal）、AAmountSum、BAmount、Difference、AmountState、Status。数组1-based，金额和索引顺序不变。
- AmountState：NOT_COMPARED=0、AMOUNT_EQUAL=1、AMOUNT_MISMATCH=2。EQUAL 仅表示本次数学相等，不代表号码正确、组完整或审核通过。
- Status：OK=0、EMPTY_GROUP=1、INVALID_RANGE=2、INVALID_AMOUNT=3、ARITHMETIC_ERROR=4。aCount=0 不比较；负数、容量不足或非零时未分配返回 INVALID_RANGE。Decimal 累加/差额失败明确返回 ARITHMETIC_ERROR。
- 首个非法 A/累加失败：ErrorSide=`A`，ErrorIndex 为 aAmounts 实际下标，ErrorAIndex 为相应原 AIndex。B 错误为 `B`、差额错误为 `DIFFERENCE`，两者索引字段为0；ErrorReason 保留原因。先按顺序处理 A，再处理 B，最后比较。
- 非 OK 不发布部分结果：AmountState=NOT_COMPARED、ACount=0、结果数组未分配、合计/B金额/差额为 Empty。调用方必须先检查 Status，不能把未比较当零差额或有效相等。
- Decimal 的有限范围/精度、CDec 对已有数值的转换与 Stage 1 一致；不恢复 Excel 或调用方此前已丢失的精度，不引入任意精度数据清洗。
- 独立入口 `-AmountTestOnly` 仅加载 Amount 和独立测试；新测试与 BConflict 41、BRowMatch 34、AIntegrity 29、Matcher 38、Parser 70、Stage 1 89 项全部实际回归。
- 不实现号码分组、组完整性、INCOMPLETE/冲突处理、颜色、fallback、A-only/B-only、baseline/fingerprint、冻结、人工审核、UI、release。完成检查点和 STATUS 最小更新后停止，不进入 C3.2。

## C3.2 — 单 B 行组完整性门控与金额协调器（2026-09-11）

本节单独授权关系门控及冻结 Amount 引擎的调用，不修改既有模块，不制定最终审核策略。

- 新模块 `src/modVATStage2GroupAmount.bas`；接口 `VATStage2EvaluateBGroupAmount(ByRef bRow As VATS2BRowMatchResult, ByVal bIndex As Long, ByRef conflicts As VATS2BConflictResult, ByRef aAmounts() As Variant, ByVal bAmount As Variant) As VATS2GroupAmountResult`。纯内存，无 Excel 对象、文件读写或 UI。
- 输入必须属于同一批关系及冲突扫描、同一 AIndex 映射空间。AIndex 是稳定标识：候选为15就直接读取 aAmounts(15)，支持0/负/正下标，禁止排序、重编号、按金额查找或丢失映射。
- NO_REFERENCES、INCOMPLETE 禁止金额比较；即使部分引用 UNIQUE 也不使用部分金额对完整 B 比较。INVALID_INPUT、关系计数/候选数组/状态契约损坏或冲突扫描状态异常，返回 GROUP_INVALID_INPUT。
- 当前行冲突必须查看组内 Associations：当前 BIndex 在同组出现至少两次为 SameBRowReuse；出现且该组 DistinctBRowCount>=2 为 CrossBRowReuse。不能把组级 SAME 标志用于该组每一行。另检查本行 UNIQUE 的 AIndex 是否重复，防止缺失冲突组时重复求和；不删除或合并引用。
- 当前行 SameBRowReuse 禁止比较；只有 CrossBRowReuse 可诊断比较并保留风险。同一组 B1 same+cross、B2仅cross 时，B1阻断而B2允许诊断。
- 只有全部引用 UNIQUE、ReferenceCount>0 且本行无 SAME 时，按 Parser 顺序组装 AIndexes/AAmounts。逐一验证候选 AIndex 位于 aAmounts 的真实边界内；未分配或越界返回 GROUP_INVALID_INPUT，不当0、不变成金额不等。
- 只调用 `VATStage2CompareAmounts`，不复制 Decimal 算术。Amount.Status=OK → GROUP_COMPARED；任何非OK → GROUP_AMOUNT_ERROR，完整嵌套保留 Amount 结果及错误。其 ErrorIndex 是组内临时金额数组位置，ErrorAIndex 仍是原 AIndex。
- 结果保留 BIndex、来源关系/冲突状态、引用及四类计数摘要、独立 ParserFlags/MatcherFlags、SameBRowReuse/CrossBRowReuse、GroupState、Reasons、映射错误位置、AmountEvaluated、Amount。原 bRow 不修改、不复制候选作为第二真源；详细未完成引用及候选由调用方保留的原 bRow 查询。
- GroupState：NOT_COMPARABLE=0、COMPARED=1、AMOUNT_ERROR=2、INVALID_INPUT=3。Reasons 独立位：NO_REFERENCES=1、INCOMPLETE_RELATION=2、INVALID_RELATION=4、SAME_ROW_A_REUSE=8、INVALID_CONFLICT=16、INVALID_A_MAPPING=32，可组合。
- 未调用引擎时 AmountEvaluated=False，AmountState=NOT_COMPARED、差额为Empty；嵌套 UDT 的默认 Status=0 不表示引擎执行成功，必须先检查 GroupState/AmountEvaluated。
- 缺NO、短尾号、结构异常、纯数字尾注、Parser单格去重提示、跨行复用均不单独阻断数学诊断，且不因金额相等被清除。各模块风险位空间不合并。
- 独立入口 `-GroupAmountTestOnly` 加载 Parser、Matcher、BRowMatch、BConflict、Amount、GroupAmount 及独立测试；新测试及 Amount48、BConflict41、BRowMatch34、AIntegrity29、Matcher38、Parser70、Stage1 89项必须全部实际 PASS。
- 不实现颜色/子集搜索、全表fallback、总体金额快速门、短尾号自动免审核、baseline/fingerprint、冻结/解冻、A-only/B-only最终报告、最终审核、UI或release。检查点及STATUS更新后停止。

## C4.1：Excel 全表只读快照（2026-09-11）

- 独立模块 `modVATStage2ExcelSnapshot`，入口 `VATStage2ReadASnapshot(sheet, headerRow, invoiceColumn, amountColumn, completedColor)` / `VATStage2ReadBSnapshot(sheet, headerRow, supplierColumn, amountColumn, completedColor)`。调用方提供已定位的 Worksheet、表头行、两个业务列及共用完成色；本层不重新查找表头、不解析或核对。
- 结果含 SheetName、HeaderRow、RecordCount、1-based Records()、Status、ErrorRow、ErrorReason。A 记录含 AIndex、ExcelRow、InvoiceDigitsRaw、AmountRaw、IsCompleted、InvoiceCellAddress、AmountCellAddress、InvoiceHasFormula、AmountHasFormula；B 对应 BIndex、SupplierTextRaw、SupplierCellAddress、SupplierHasFormula。原值用 Variant 保存 Value2，其余为标量，不保存任何 Excel 对象。
- 从 HeaderRow+1 扫描至 UsedRange 末行。只读取两业务字段及金额单元格的显示填色，隐藏/筛选行照常读取。两字段均为 Empty、均无公式且金额非完成色时才跳过；仅格式残留不产生记录。身份存在金额空、金额存在身份空、公式空串、完成色空行都保留。
- 每个保留行按 Excel 行序赋值 AIndex/BIndex=1..RecordCount，并保存原 ExcelRow。索引基于完整快照，不按完成色筛选或重新编号；已有记录仅变色不改变其索引。索引不是跨插行、删除、重排或记录集合改变的永久标识。
- 完成色语义等价于 Stage 1 VATScan 的计数条件：金额单元格 DisplayFormat.Interior 的 Pattern 非 xlPatternNone、ColorIndex 非 xlColorIndexNone、Color 等于 completedColor。无填充与白色不同，条件格式使用实际显示色。图案填充达到此条件也标记完成，与 Stage 1 计数一致；快照不作金额合法性或审核判断。
- 原始值不 Trim、不转字符串或数字、不 CDec、不修复、不调用 Parser/Matcher/Amount。原样保留 Empty、错误、布尔、数字文本和公式空串；读取当前未保存 Value2，不触发计算。
- Status：OK=0、INVALID_SHEET=1、INVALID_HEADER=2、INVALID_COLUMN=3、READ_ERROR=4，均以 VATS2_SNAPSHOT_ 为前缀。拒绝 Nothing、已关闭/失效工作表、越界表头及列号。读取失败清空部分记录并保留错误说明/行号；必须先检查 Status，再使用 Records。有效空结果的数组未分配。
- 模块不修改数据/公式/格式，不保存或关闭输入工作簿，不筛选排序。测试只创建自己的临时工作簿；`-SnapshotTestOnly` 仅为测试对照加载 Stage 1，生产快照模块不依赖 Stage 1 扫描器。
- 验证：Snapshot81、GroupAmount44、Amount48、BConflict41、BRowMatch34、AIntegrity29、Matcher38、Parser70、Stage1 89 全部真实 Excel PASS；冻结源码/旧测试/release 哈希不变。
- 本轮不建立完成子集、不做作用域匹配/fallback/颜色缺失异常、金额比较/总额快速门、baseline/fingerprint、冻结、报告、UI或release。C4.1 检查点及 STATUS 更新后停止，不进入 C4.2。

## C4.2：完成色作用域匹配（2026-09-12）

- 独立纯内存协调层 `modVATStage2CompletedScope`，入口 `VATStage2MatchCompletedScope(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult) As VATS2CompletedScopeResult`。不访问工作簿、不复制完整快照、不修改冻结模块。
- 先验证两 Snapshot.Status=OK，否则 SCOPE_INVALID_INPUT。记录数组必须对应冻结快照的1..RecordCount，记录内原Index等于数组位置，ExcelRow在HeaderRow之后严格递增；零记录时数组未分配。损坏返回 SCOPE_INVALID_CONTRACT，不筛选后掩盖接口错误。
- 从完整A按原序选出IsCompleted=True的记录，建立1..CompletedACount位置到OriginalAIndex及ExcelRow映射。String号码原样传递，不Trim、不CStr；非String保留scope位置，用空String占位，并记录AIdentityIssue（OriginalAIndex、ExcelRow、RawVarType），属于数据风险而非接口失败。
- BMatches、BInScope、BRowIds均按完整BIndex空间1..FullBRecordCount保存。仅完成B调用冻结VATStage2MatchBRow；未完成B直接保留合法零引用NO_REFERENCES占位，不调用Parser/Matcher，也不产生源类型Issue。BRowIds(BIndex)=ExcelRow。
- 完成B的SupplierTextRaw是String则原样传入；否则用空String取得NO_REFERENCES，同时记录BSourceIssue（OriginalBIndex、ExcelRow、RawVarType）。带源Issue的NO_REFERENCES不代表真实缺票。
- BRowMatch返回后只修改副本Candidate.AIndex：scope位置→OriginalAIndex。所有候选均验证位置范围，候选顺序、MatchKind、CandidateCount、InvoiceDigits、上下文、ParserFlags、MatcherFlags均保持原值。exact优先、multiple和短suffix风险完全由实际completed scope内冻结Matcher决定。
- 独立回映入口 `VATStage2RemapScopeRow(source, originalAIndexes(), scopeCount, remapped) As VATS2ScopeStatus` 校验映射和引用/候选数组，全部验证后才生成结果副本；失败输出空行并返回INVALID_CONTRACT。输入与输出应使用不同变量，此接口不调用Matcher、不消歧。
- 使用完整BMatches与BRowIds调用冻结VATStage2ScanBConflicts，因此Association的BIndex和AIndex均为原完整快照索引，BRowId为ExcelRow。未完成B零引用不生成association。Conflicts.BRecordCount是完整B数量，不是完成B数量。
- 结果保留Status、FullARecordCount、FullBRecordCount、CompletedACount、CompletedBCount、CompletedAIndexes、CompletedAExcelRows、AIdentityIssueCount/Issues、BInScope、BRowIds、BSourceIssueCount/Issues、BMatches、Conflicts，以及ErrorSide/Index/Reason。零数量数组未分配；必须先检查Status，再结合BInScope和Issues解释行结果。
- Status均带VATS2_前缀：SCOPE_OK=0、SCOPE_INVALID_INPUT=1、SCOPE_INVALID_CONTRACT=2、SCOPE_CONFLICT_ERROR=3。回映失败不发布部分结果；冲突扫描非OK时保留其嵌套错误及已经完成的行结果，但整次scope不能作为成功使用。运行时错误按所在阶段返回明确非OK。
- 结果只说明双方完成色范围内的号码关系。NOT_FOUND不意味着全表不存在、漏色或B_ONLY；未被引用的完成A也不判A_ONLY。非String之外的A字符串质量仍留给AIntegrity及调用层，不新增数据清洗。
- `-CompletedScopeTestOnly`新测试73/73实际PASS；Snapshot81、GroupAmount44、Amount48、BConflict41、BRowMatch34、AIntegrity29、Matcher38、Parser70、Stage1 89全部实际PASS。旧源码/测试/release哈希不变，不重建release。
- 不实现full A/B fallback、A_COLOR_MISSING/B_COLOR_MISSING、A_ONLY/B_ONLY、金额/总额快速门、短尾号免审核、baseline/fingerprint、冻结/解冻、最终审核或UI。完成检查点与STATUS后停止，C4.3待授权。

## C4.3：completed-scope 失败后的 full-table fallback（2026-09-12）

- 独立模块 `modVATStage2FullFallback`，入口 `VATStage2RunFullFallback(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, ByRef completed As VATS2CompletedScopeResult) As VATS2FullFallbackResult`。纯内存，读取已有结果，不修改C1～C4.2。
- 第一层始终以completed为准。B→A仅处理完成B中的NOT_FOUND引用，调用冻结Matcher搜索完整A。完整String号码数组下标直接等于原AIndex；非String用空String占位，不CStr。已unique或multiple的第一层引用不重匹配、不替换。
- BToAFallback保留BIndex、ExcelRow、ReferenceIndex、整个OriginalBMatch、完整FullAMatch、含原AIndex/ExcelRow/InvoiceDigits/IsCompleted的全部候选、Evidence及MissingEvidence。未完成A唯一命中为A_COLOR_MISSING证据；MULTIPLE不挑选；full unique completed为INVALID_CONTRACT并保留矛盾现场。
- 对每个完成A先检查所有completed B候选：存在UNIQUE为COVERAGE_UNIQUE；仅MULTIPLE为COVERAGE_MULTIPLE_ONLY；完全没有候选才为COVERAGE_NONE并触发A→B。前两者不fallback，不以全表结果改写。
- 只有存在A→B目标才逐行调用冻结BRowMatch搜索完整B，使用完整A数组。FullBMatches(OriginalBIndex)保存整行全部refs及成功/失败/歧义，不截取单个命中；FullBScanned明确是否执行此搜索。新full诊断不能整体覆盖completed第一层结果，即使其exact优先导致结果不同。
- AToBFallback保留原AIndex/ExcelRow、Evidence、全部B行Hits、MissingEvidence及SourceBlocked。每个Hit保留BIndex/ExcelRow/IsCompleted、UniqueReferenceCount、MultipleReferenceCount；整行关系通过FullBMatches(Hit.BIndex)取得。未完成B唯一命中为B_COLOR_MISSING证据，多个B全部保留；仅multiple为FULL_AMBIGUOUS，不能判缺色；异常的已完成B唯一命中返回INVALID_CONTRACT。
- AColorMissingCount/BColorMissingCount分别按原AIndex/BIndex去重统计缺色记录数，所有逐引用/逐A命中证据不删除。正反向均不处理金额、组金额完整性或冲突消歧。
- A blockers至少包括非String，另对空/非ASCII数字String标记遮蔽但仍原样传给Matcher。B blockers包括非String、Parser STRUCTURE_REVIEW、NUMERIC_TRAILER_REVIEW或零引用；普通NO_MARKER_MISSING和单格重复提示不单独构成搜索遮蔽。每项保存OriginalIndex、ExcelRow、RawVarType、Reasons、ParserFlags。
- 无候选且可能被相应方向blocker遮蔽时为FULL_NOT_FOUND_WITH_BLOCKERS，MissingEvidence=False。A→B源A自身不可搜索也视为遮蔽。无blocker的FULL_NOT_FOUND才设MissingEvidence=True，仍只是当前快照下可供后续使用的证据，不输出最终A_ONLY/B_ONLY或审核PASS/FAIL。
- Evidence枚举带VATS2_前缀：FULL_UNIQUE_COMPLETED=1、FULL_UNIQUE_UNCOMPLETED=2、FULL_MULTIPLE=3、FULL_NOT_FOUND=4、FULL_NOT_FOUND_WITH_BLOCKERS=5、FULL_AMBIGUOUS=6。反向UNIQUE指存在唯一匹配reference，不保证只有一个B行；全部Hits必须保留。
- 主Status带VATS2_前缀：FALLBACK_OK=0、FALLBACK_INVALID_INPUT=1、FALLBACK_INVALID_CONTRACT=2、FALLBACK_ENGINE_ERROR=3。检查非OK输入、快照数组/索引/行序及completed来源/颜色/引用容器与候选契约。先看Status再解释数据；矛盾路径保存部分现场用于诊断，不能作为完成结果使用。未执行full B时BBlockerCount=0不表示已确认B无blocker。
- 新增`-FullFallbackTestOnly`，新测试132/132 PASS。CompletedScope73、Snapshot81、GroupAmount44、Amount48、BConflict41、BRowMatch34、AIntegrity29、Matcher38、Parser70、Stage1 89全部真实Excel PASS。冻结源码、旧测试、release哈希不变。
- 不实现总体金额快速门、金额比较、短suffix免审核、baseline/fingerprint、冻结、最终审核、UI或release接入。C4.3检查点和STATUS更新后停止，不进入下一阶段。

## C5.1：fallback 后关系整合与冲突重扫（2026-09-12）

- 独立纯内存模块 `modVATStage2EffectiveRelations`，入口 `VATStage2BuildEffectiveRelations(a, b, completed, fallback) As VATS2EffectiveRelationsResult`，四个输入均ByRef只读。结果不是Snapshot第二真源，不调用Parser/Matcher重新匹配、不读金额、不调用GroupAmount。
- EffectiveBInScope、EffectiveBMatches、BRowIds和RowIssueFlags均使用完整1..FullBRecordCount空间。所有完成B进入；未完成B仅在C4.3存在真正UNIQUE命中、形成B_COLOR_MISSING证据时按原BIndex加入一次。其它未完成B为合法零引用NO_REFERENCES占位，BRowIds始终等于原ExcelRow。
- 完成B以C4.2行结果副本为基础。原EXACT/SUFFIX的UNIQUE/MULTIPLE逐引用保持不变；仅NOT_FOUND按BIndex+ReferenceIndex寻找恰好一条BToAFallback。复制FullAMatch的MatchKind、Flags→MatcherFlags、CandidateCount和全部Candidates，不重映射原AIndex；Parser产生的Digits/Length/RawFragment/StartIndex/Context/ParserFlags均保留。
- 完成B合并后重新计算四类引用计数、MatchState和逐引用MatcherFlags的OR。不能保留旧NOT_FOUND数量或旧短suffix标志。fallback的缺色、blocker或MissingEvidence标签不改变Matcher身份分类。
- 缺色未完成B直接采用FullBMatches(BIndex)的完整行，含所有成功、歧义和未找到引用，不截取触发命中的reference。逐A命中重叠时仍只纳入该B一次；核验Hit索引、ExcelRow、完成状态、unique/multiple计数与完整行一致。
- 重新建立完整A String数组，String原样，非String为空String，不CStr/Trim。只调用冻结VATStage2ScanAIntegrity，保存完整AIntegrity结果；建立AInvalidValue(AIndex)和ADuplicateGroupIndex(AIndex)，不自行判重或实现ASCII校验。
- 逐Effective UNIQUE关联查A质量摘要，设置独立RowIssueFlags：VATS2_RELATION_A_INVALID=1、VATS2_RELATION_A_DUPLICATE=2，可组合。RelationIssues保存每个有问题的唯一关联的BIndex、ExcelRow、ReferenceIndex、AIndex、AExcelRow、Flags。保留原关系用于诊断，不自动选择其它A；无关A问题不作全局关系失败。
- 完整EffectiveBMatches/BRowIds调用冻结VATStage2ScanBConflicts，保存EffectiveConflicts。可发现两个fallback B、同B不同fallback refs，以及completed B与缺色B之间新增的SAME/CROSS复用；不实现另一套冲突算法，也不因A质量风险跳过关联诊断。
- 主结果含Status、FullARecordCount、FullBRecordCount、EffectiveBCount、上述完整数组、AIntegrity、质量摘要、RelationIssueCount/Issues、EffectiveConflicts、ErrorSide/Index/ReferenceIndex/Reason。Status带VATS2_前缀：EFFECTIVE_OK=0、EFFECTIVE_INVALID_INPUT=1、EFFECTIVE_INVALID_CONTRACT=2、EFFECTIVE_ENGINE_ERROR=3。失败清空部分输出并保留错误定位。
- 必须拒绝非OK输入、record count/索引/行序不一致、NOT_FOUND补充缺失/重复/越界、错误原B关系、FullAMatch引用身份不一致或候选越界/号码来源不一致、缺色Hit指向完成B、FullBMatches数组缺失/原文错位、Hit与候选不对应、整行计数/状态/数组损坏。第一层优先不因全表exact改变。
- EFFECTIVE_OK只表示整合和重扫完成，不代表UNIQUE已可信。后续C5.2必须检查RowIssueFlags/RelationIssues及EffectiveConflicts；RELATION_A_INVALID或RELATION_A_DUPLICATE不得视为可信金额关系。本轮不调用金额层、不决定最终审核。
- `-EffectiveRelationsTestOnly`新测试135/135 PASS；FullFallback132、CompletedScope73、Snapshot81、GroupAmount44、Amount48、BConflict41、BRowMatch34、AIntegrity29、Matcher38、Parser70、Stage1 89全部真实Excel PASS。27个冻结源码/旧测试/release文件哈希不变。
- 本轮不实现金额比较、总体金额快速门、最终A_ONLY/B_ONLY、短suffix豁免、baseline/fingerprint、冻结、最终审核、UI或release。检查点和STATUS更新后停止，C5.2未开始。

## C5.2：Effective Relations 统一组金额诊断（2026-09-12）

- 独立模块 `modVATStage2EffectiveAmount`，入口 `VATStage2EvaluateEffectiveAmounts(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, ByRef effective As VATS2EffectiveRelationsResult) As VATS2EffectiveAmountResult`。纯内存只读，不访问 Workbook，不修改冻结模块。
- 建立完整 `aAmounts(1..a.RecordCount)`，逐原 AIndex 承接 Snapshot.AmountRaw 的 Variant，不重排、不压缩、不转换、不按完成色过滤。仅 EffectiveBInScope=True 的 B 使用既有整行 EffectiveBMatches、EffectiveConflicts 和该 B 原 AmountRaw。
- 逐行质量门控先于 GroupAmount：RowIssueFlags 含 RELATION_A_INVALID 或 RELATION_A_DUPLICATE，RowStates 返回独立 `VATS2_RELATION_QUALITY_BLOCKED`，不调用 GroupAmount，不比较金额。保留完整 RelationQualityFlags 与 RelationIssues（原 BIndex/ExcelRow/ReferenceIndex/AIndex/AExcelRow/Flags）。
- 无质量阻断才调用冻结 `VATStage2EvaluateBGroupAmount`；完整保存其 GroupState、Reasons、引用计数、ParserFlags、MatcherFlags、Same/CrossBRowReuse、AmountEvaluated、错误定位和嵌套 Amount。ΣA−B、Decimal、逐项 AIndexes/AAmounts 与 AMOUNT_EQUAL/MISMATCH 完全由冻结层产生，不复制金额算法。
- A_COLOR_MISSING 后补入的 A、B_COLOR_MISSING 后纳入的整行 B 正常参与数学诊断。INCOMPLETE、NO_REFERENCES、SAME 由 GroupAmount 阻断；仅 CROSS 不阻断。不因金额相等清除任何原关系风险或决定免人工。
- 结果使用完整 BIndex 空间：BInScope、BRowIds、RowStates、GroupAmounts、RelationQualityFlags 均1-based；零记录数组未分配。范围外 RowStates=EA_OUT_OF_SCOPE；质量阻断行的默认 GroupAmounts 未执行，不得把其嵌套默认值解释为成功。
- RowStates 前四项与冻结 GroupState 对齐：EA_NOT_COMPARABLE=0、EA_COMPARED=1、EA_AMOUNT_ERROR=2、EA_INVALID_INPUT=3；RELATION_QUALITY_BLOCKED=4、EA_OUT_OF_SCOPE=5（均带 VATS2_ 前缀）。GroupAmounts 保留冻结接口原枚举名称。
- 主 Status（VATS2_EFFECTIVE_AMOUNT_ 前缀）：OK=0、INVALID_INPUT=1（上游非OK）、INVALID_CONTRACT=2（输入结构/来源错误）、GROUP_ERROR=3（冻结 GroupAmount 返回 INVALID_INPUT）。后者保留完整行结果并增加 InvalidGroupCount，不能伪装金额不等；必须先检查主状态。
- 汇总 EffectiveBCount、ComparedCount、EqualCount、MismatchCount、NotComparableCount、AmountErrorCount、RelationQualityBlockedCount、InvalidGroupCount；前者等于各组状态数量之和，ComparedCount=EqualCount+MismatchCount。这些只是诊断统计，不是最终审核 PASS/FAIL。
- 拒绝 Snapshot/Effective 非OK、record count不一致、记录数组边界或原Index/ExcelRow错位、B范围/原文/引用计数/候选映射损坏。完整 AIntegrity 与质量摘要交叉检查，逐 UNIQUE 关联与有序 RelationIssues 一一核对，再核对 RowIssueFlags 的 OR；未知标志、缺失、多余、重复、越界问题记录均拒绝。
- EffectiveConflicts 必须OK并对应完整B数量；调用冻结 BConflict 扫描器作契约对照，逐字段核验原有冲突结果（含原 BRowId、AIndex、ReferenceIndex、计数和标志）。不另写冲突算法，不替换或修复传入结果。
- 契约失败清空部分输出，只保留明确 Status 和 ErrorSide/Index/ReferenceIndex/Reason。输入必须属于同一批次；本层不重新生成号码匹配或 AIntegrity 结果，不引入 fingerprint。
- 新测试入口 `./tools/build.ps1 -EffectiveAmountTestOnly`，新测试200/200断言PASS；十二组冻结回归全部PASS，共814项；合计1014项。没有修改旧测试语义、冻结源码或release。
- C5.2完成即停止。不实现总体A/B金额快速门、总体审核、short suffix豁免、最终A_ONLY/B_ONLY结论、baseline、fingerprint、冻结/解冻、UI、最终报告或release。C5.3未开始。

## C5.3：SHORT_SUFFIX_REVIEW 本次运行豁免策略（2026-09-12）

- 新模块 `modVATStage2ShortSuffixPolicy`；入口 `VATStage2EvaluateShortSuffixPolicy(a, b, effective, amounts, aggregateReliableEqual) As VATS2ShortSuffixPolicyResult`。前四项为只读ByRef的A/B Snapshot、EffectiveRelations、EffectiveAmount；最后为ByVal Boolean，由调用方确认Stage 1完成金额扫描可靠、没有完成色空金额warning/致命异常且Decimal总体精确相等。本模块不计算或推断该资格。
- 仅为Effective范围内原引用MatcherFlags实际含SHORT_SUFFIX_REVIEW的引用生成扁平Decision；不是按长度自行补加风险。冻结Matcher对short NOT_FOUND也可能保留此标志，其决策仍为REVIEW_REQUIRED；无标志的exact、长suffix和空A情况下NOT_FOUND不生成决策。
- WAIVED_THIS_RUN必须同时满足：aggregateReliableEqual=True；B整行完整且全部UNIQUE；C5.2该完整组实际比较成功且AMOUNT_EQUAL；RelationQualityFlags=0；整行及各引用ParserFlags=0；无SAME/CROSS冲突；B已完成且该组全部唯一匹配A均已完成；无其它身份结构风险。任一不满足则REVIEW_REQUIRED，理由位可组合。
- 仅豁免短suffix本次单独人工复核，不代表最终审核通过，不写回EffectiveRelations/EffectiveAmount，不清除MatcherFlags，不持久化。调用方不得用策略结果替代原始风险来源。
- 每个Decision保存原BIndex、ExcelRow、ReferenceIndex、AIndex、ReferenceDigits、Decision和ReasonFlags；按BIndex及引用原序排列。非唯一引用AIndex=0表示没有唯一对应，不从MULTIPLE中选第一项；完整候选仍由调用方保留的原关系提供。
- Decision常量：VATS2_REVIEW_REQUIRED=0、VATS2_WAIVED_THIS_RUN=1。原因位：VATS2_AGGREGATE_NOT_ELIGIBLE=1、VATS2_ROW_NOT_COMPLETE_UNIQUE=2、VATS2_GROUP_NOT_EQUAL=4、VATS2_PARSER_RISK=8、VATS2_RELATION_QUALITY_RISK=16、VATS2_CONFLICT_RISK=32、VATS2_COLOR_MISSING=64、VATS2_OTHER_STRUCTURE_RISK=128。未知的非short Matcher风险位视为其它结构风险，不自动豁免。
- 结果另有Status、AggregateReliableEqual、DecisionCount、Decisions()、WaivedCount、ReviewRequiredCount、ErrorBIndex/ReferenceIndex/Reason。Status为VATS2_SHORT_POLICY_OK=0、INVALID_INPUT=1、INVALID_CONTRACT=2（后三者同前缀）。失败清空部分决策，不能保留局部豁免；零决策数组未分配。
- 只校验本层所用状态、计数/索引空间、范围/行号、唯一候选身份、关系质量标志、组来源风险、实际比较状态与组内AIndex的一致性。对short标志与exact/长引用的矛盾拒绝，而不改写原flag。不复制C5.1/C5.2全部防御性校验，不重做AIntegrity。
- 仅读取C5.2的SAME/CROSS及已有EffectiveConflicts关联是否涉及本B，作为CONFLICT_RISK；不重新分组或扫描冲突。无金额重算、Parser/Matcher调度、最终审核、A_ONLY/B_ONLY、baseline/fingerprint、冻结、UI或release。
- 新入口 `./tools/build.ps1 -ShortSuffixPolicyTestOnly`：139/139断言PASS；十三组旧回归共1014项全部PASS，合计1153项。冻结源码、旧测试和release哈希不变。C5.3完成后停止，未开始下一阶段。
