Attribute VB_Name = "modVATStage2AuditTests"
Option Explicit

Private passed As Long, failed As Long, log As String
Private a As VATS2ASnapshotResult, b As VATS2BSnapshotResult
Private f As VATS2FullFallbackResult, e As VATS2EffectiveRelationsResult
Private m As VATS2EffectiveAmountResult, p As VATS2ShortSuffixPolicyResult

Public Function VATStage2AuditFindings_SelfTest() As String
    On Error GoTo Unexpected
    passed = 0: failed = 0: log = ""
    TestColorAndMissing
    TestRisks
    TestAdditionalEvidence
    TestContracts
    GoTo Done
Unexpected:
    failed = failed + 1: log = log & "UNEXPECTED " & Err.Number & " " & Err.Description & vbCrLf
Done:
    If failed = 0 Then
        VATStage2AuditFindings_SelfTest = "PASS: " & passed & " assertions" & vbCrLf & log
    Else
        VATStage2AuditFindings_SelfTest = "FAIL: " & failed & "; PASS: " & passed & vbCrLf & log
    End If
End Function

Private Sub Fixture(ByVal ac As Long, ByVal bc As Long)
    Dim aa As VATS2ASnapshotResult, bb As VATS2BSnapshotResult, i As Long
    a = aa: b = bb: a.HeaderRow = 3: b.HeaderRow = 5
    a.SheetName = "合成A": b.SheetName = "合成B": a.RecordCount = ac: b.RecordCount = bc
    If ac > 0 Then ReDim a.Records(1 To ac)
    If bc > 0 Then ReDim b.Records(1 To bc)
    For i = 1 To ac
        With a.Records(i)
            .AIndex = i: .ExcelRow = 10 + i * 2: .InvoiceDigitsRaw = "1000000" & i
            .AmountRaw = i: .IsCompleted = True
        End With
    Next i
    For i = 1 To bc
        With b.Records(i)
            .BIndex = i: .ExcelRow = 20 + i * 3: .SupplierTextRaw = "NO.1000000" & i
            .AmountRaw = i: .IsCompleted = True
        End With
    Next i
End Sub

Private Sub Prepare(Optional ByVal aggregateEqual As Boolean = True)
    Dim c As VATS2CompletedScopeResult
    '只在测试夹具运行冻结上游，被测Findings模块不调用任何业务引擎。
    c = VATStage2MatchCompletedScope(a, b)
    f = VATStage2RunFullFallback(a, b, c)
    e = VATStage2BuildEffectiveRelations(a, b, c, f)
    m = VATStage2EvaluateEffectiveAmounts(a, b, e)
    p = VATStage2EvaluateShortSuffixPolicy(a, b, e, m, aggregateEqual)
    If f.Status <> 0 Or e.Status <> 0 Or m.Status <> 0 Or p.Status <> 0 Then Err.Raise 5, , "夹具失败：" & f.ErrorReason & e.ErrorReason & m.ErrorReason & p.ErrorReason
End Sub

Private Sub TestColorAndMissing()
    Dim r As VATS2AuditFindingsResult, pos As Long
    Fixture 1, 1: Prepare: r = RunAudit()
    Check "正常exact没有事项", r.FindingCount = 0
    a.Records(1).IsCompleted = False: b.Records(1).AmountRaw = 9
    Call Prepare: r = RunAudit()
    Check "A缺色和金额差异共存", r.CodeCounts(VATS2_FIND_A_COLOR_MISSING) = 1 And r.CodeCounts(VATS2_FIND_AMOUNT_MISMATCH) = 1
    pos = FindCode(r, VATS2_FIND_A_COLOR_MISSING)
    Check "A缺色双方实际位置", r.Findings(pos).AIndex = 1 And r.Findings(pos).AExcelRow = 12 And r.Findings(pos).BIndex = 1 And r.Findings(pos).BExcelRow = 23 And r.Findings(pos).ReferenceIndex = 1
    pos = FindCode(r, VATS2_FIND_AMOUNT_MISMATCH)
    Check "差额原样为A减B且完整成员", r.Findings(pos).Difference = CDec(-8) And VarType(r.Findings(pos).Difference) = vbDecimal And r.Findings(pos).Members(1).AIndex = 1
    Fixture 2, 1: b.Records(1).IsCompleted = False: b.Records(1).SupplierTextRaw = "NO.10000001/10000002": b.Records(1).AmountRaw = 9
    Call Prepare: r = RunAudit()
    Check "B缺色按两条A-B关系定位而非总数", r.CodeCounts(VATS2_FIND_B_COLOR_MISSING) = 2 And r.CodeCounts(VATS2_FIND_AMOUNT_MISMATCH) = 1
    pos = FindCode(r, VATS2_FIND_AMOUNT_MISMATCH)
    Check "缺色合并B保留整组A与差额", r.Findings(pos).MemberCount = 2 And r.Findings(pos).Members(2).AIndex = 2 And r.Findings(pos).Difference = CDec(-6)
    Fixture 1, 1: b.Records(1).SupplierTextRaw = "NO.99999999": Prepare: r = RunAudit()
    Check "全表确定A_ONLY和B_ONLY同时存在", r.CodeCounts(VATS2_FIND_A_ONLY) = 1 And r.CodeCounts(VATS2_FIND_B_ONLY) = 1 And r.CodeCounts(VATS2_FIND_INCOMPLETE_GROUP) = 1
    a.Records(1).InvoiceDigitsRaw = "非法": Prepare: r = RunAudit()
    Check "A blocker和SourceBlocked禁止两侧错误missing", r.CodeCounts(VATS2_FIND_A_ONLY) = 0 And r.CodeCounts(VATS2_FIND_B_ONLY) = 0 And r.CodeCounts(VATS2_FIND_SEARCH_BLOCKER) = 1 And r.CodeCounts(VATS2_FIND_A_DATA_REVIEW) = 1
    Fixture 1, 1: b.Records(1).SupplierTextRaw = "NO.99999999--123": Prepare: r = RunAudit()
    Check "B blocker只阻止A_ONLY不抹去可靠B_ONLY", r.CodeCounts(VATS2_FIND_A_ONLY) = 0 And r.CodeCounts(VATS2_FIND_B_ONLY) = 1 And r.CodeCounts(VATS2_FIND_SEARCH_BLOCKER) = 1
    pos = FindCode(r, VATS2_FIND_SEARCH_BLOCKER)
    Check "blocker原始side位置原因ParserFlags保留", r.Findings(pos).Side = "B" And r.Findings(pos).BIndex = 1 And r.Findings(pos).BExcelRow = 23 And r.Findings(pos).ReasonFlags = f.BBlockers(1).Reasons And r.Findings(pos).ParserFlags = f.BBlockers(1).ParserFlags
    Fixture 0, 0: Prepare: r = RunAudit(): Check "空数据正常零事项", r.FindingCount = 0 And r.ShortSuffixWaivedCount = 0
End Sub

Private Sub TestRisks()
    Dim r As VATS2AuditFindingsResult, pos As Long, mode As Long, before As String
    Fixture 2, 1: a.Records(2).InvoiceDigitsRaw = "10000001": Prepare: r = RunAudit()
    Check "exact multiple和重复A及不完整并存", r.CodeCounts(VATS2_FIND_AMBIGUOUS_MATCH) = 1 And r.CodeCounts(VATS2_FIND_DUPLICATE_A_INVOICE) = 1 And r.CodeCounts(VATS2_FIND_INCOMPLETE_GROUP) = 1
    pos = FindCode(r, VATS2_FIND_DUPLICATE_A_INVOICE)
    Check "重复组全部原A索引行号", r.Findings(pos).MemberCount = 2 And r.Findings(pos).Members(1).AIndex = 1 And r.Findings(pos).Members(2).AExcelRow = 14
    pos = FindCode(r, VATS2_FIND_AMBIGUOUS_MATCH)
    Check "歧义不选择第一项", r.Findings(pos).AIndex = 0 And r.Findings(pos).MemberCount = 2 And r.Findings(pos).MatchKind = VATS2_EXACT_MULTIPLE
    Fixture 2, 1: a.Records(2).InvoiceDigitsRaw = "20000001": b.Records(1).SupplierTextRaw = "NO.0000001"
    Call Prepare: r = RunAudit(): pos = FindCode(r, VATS2_FIND_AMBIGUOUS_MATCH)
    Check "suffix multiple保留全部候选", r.Findings(pos).MatchKind = VATS2_SUFFIX_MULTIPLE And r.Findings(pos).MemberCount = 2
    a.Records(1).IsCompleted = False: a.Records(2).IsCompleted = False: Prepare: r = RunAudit()
    pos = FindCode(r, VATS2_FIND_AMBIGUOUS_MATCH)
    Check "FULL_MULTIPLE与effective同证据去重合并来源", r.CodeCounts(VATS2_FIND_AMBIGUOUS_MATCH) = 1 And r.Findings(pos).SourceFlags = (VATS2_FIND_FROM_EFFECTIVE Or VATS2_FIND_FROM_B_TO_A)
    Fixture 2, 1: a.Records(2).InvoiceDigitsRaw = "20000001": b.Records(1).SupplierTextRaw = "NO.0000001": b.Records(1).IsCompleted = False
    Call Prepare: r = RunAudit()
    Check "反向FULL_AMBIGUOUS两目标共享证据仅一项", f.AToBFallbacks(1).Evidence = VATS2_FULL_AMBIGUOUS And r.CodeCounts(VATS2_FIND_AMBIGUOUS_MATCH) = 1 And r.CodeCounts(VATS2_FIND_B_COLOR_MISSING) = 0
    Fixture 1, 1: b.Records(1).SupplierTextRaw = "无号码": Prepare: r = RunAudit()
    Check "NO_REFERENCES是INCOMPLETE且blocker可见", r.CodeCounts(VATS2_FIND_INCOMPLETE_GROUP) = 1 And r.CodeCounts(VATS2_FIND_SEARCH_BLOCKER) = 1 And r.CodeCounts(VATS2_FIND_A_ONLY) = 0
    For mode = 1 To 6
        Fixture 1, 1
        Select Case mode
            Case 1: b.Records(1).SupplierTextRaw = "NO.10000001/"
            Case 2: b.Records(1).SupplierTextRaw = "NO.10000001--123"
            Case 3: b.Records(1).SupplierTextRaw = "公司10000001"
            Case 4: b.Records(1).SupplierTextRaw = "NO.10000001/10000001"
            Case 5: b.Records(1).AmountRaw = "1"
            Case 6: a.Records(1).AmountRaw = CVErr(2042)
        End Select
        Call Prepare: r = RunAudit()
        Select Case mode
            Case 1: Check "结构风险", r.CodeCounts(VATS2_FIND_STRUCTURE_REVIEW) = 1
            Case 2: Check "纯数字尾注风险", r.CodeCounts(VATS2_FIND_NUMERIC_TRAILER_REVIEW) = 1
            Case 3: Check "无NO提示", r.CodeCounts(VATS2_FIND_NO_MARKER_MISSING) = 1
            Case 4: Check "Parser重复引用只一项且保留来源", r.CodeCounts(VATS2_FIND_DUPLICATE_REFERENCE) = 1 And r.Findings(FindCode(r, VATS2_FIND_DUPLICATE_REFERENCE)).SourceFlags = VATS2_FIND_FROM_PARSER
            Case 5, 6: Check "金额异常不伪装差异", r.CodeCounts(VATS2_FIND_AMOUNT_ERROR) = 1 And r.CodeCounts(VATS2_FIND_AMOUNT_MISMATCH) = 0
        End Select
    Next mode
    Fixture 1, 1: b.Records(1).SupplierTextRaw = "NO.10000001/000001": Prepare: r = RunAudit()
    pos = FindCode(r, VATS2_FIND_DUPLICATE_REFERENCE)
    Check "SAME来源保留全部重复关联", r.CodeCounts(VATS2_FIND_DUPLICATE_REFERENCE) = 1 And r.Findings(pos).SourceFlags = VATS2_FIND_FROM_CONFLICT And r.Findings(pos).MemberCount = 2
    Fixture 1, 2: b.Records(2).SupplierTextRaw = "NO.10000001": Prepare: r = RunAudit()
    pos = FindCode(r, VATS2_FIND_CROSS_B_ROW_REUSE)
    Check "CROSS冲突整组可追溯", r.CodeCounts(VATS2_FIND_CROSS_B_ROW_REUSE) = 1 And r.Findings(pos).MemberCount = 2 And r.Findings(pos).Members(2).BIndex = 2 And r.Findings(pos).Members(2).BExcelRow = 26
    Check "CROSS不抹去金额差异", r.CodeCounts(VATS2_FIND_AMOUNT_MISMATCH) = 1
    Fixture 2, 1: a.Records(2).InvoiceDigitsRaw = "10000001": a.Records(2).IsCompleted = False
    Call Prepare: r = RunAudit()
    Check "完整关系质量阻断不误标不完整", m.RowStates(1) = VATS2_RELATION_QUALITY_BLOCKED And r.CodeCounts(VATS2_FIND_INCOMPLETE_GROUP) = 0 And r.CodeCounts(VATS2_FIND_RELATION_QUALITY_REVIEW) = 1
    Fixture 2, 1: a.Records(2).InvoiceDigitsRaw = "非法": Prepare: r = RunAudit()
    Check "未参与A仍报告质量事实", r.CodeCounts(VATS2_FIND_A_DATA_REVIEW) = 1 And r.CodeCounts(VATS2_FIND_RELATION_QUALITY_REVIEW) = 0
    Fixture 1, 1: b.Records(1).SupplierTextRaw = "NO.000001": Prepare False: r = RunAudit()
    Check "short REVIEW直接保留理由", r.CodeCounts(VATS2_FIND_SHORT_SUFFIX_REVIEW) = 1 And r.Findings(FindCode(r, VATS2_FIND_SHORT_SUFFIX_REVIEW)).ReasonFlags = p.Decisions(1).ReasonFlags
    Prepare True: r = RunAudit()
    Check "short WAIVED只计数不生成异常", r.FindingCount = 0 And r.ShortSuffixWaivedCount = 1 And e.EffectiveBMatches(1).References(1).MatcherFlags = VATS2_SHORT_SUFFIX_REVIEW
    Fixture 1, 1: b.Records(1).SupplierTextRaw = "公司000001/000001--123": b.Records(1).AmountRaw = 9
    Call Prepare: r = RunAudit()
    Check "同一B四种独立事项全部保留", r.CodeCounts(VATS2_FIND_NO_MARKER_MISSING) = 1 And r.CodeCounts(VATS2_FIND_NUMERIC_TRAILER_REVIEW) = 1 And r.CodeCounts(VATS2_FIND_DUPLICATE_REFERENCE) = 1 And r.CodeCounts(VATS2_FIND_AMOUNT_MISMATCH) = 1 And r.CodeCounts(VATS2_FIND_SHORT_SUFFIX_REVIEW) = 1
    Fixture 3, 3
    For mode = 1 To 3: b.Records(mode).AmountRaw = 9: Next mode
    Call Prepare: r = RunAudit()
    Check "Finding顺序为原B索引", r.Findings(1).BIndex = 1 And r.Findings(2).BIndex = 2 And r.Findings(3).BIndex = 3
End Sub

Private Sub TestContracts()
    Dim mode As Long, r As VATS2AuditFindingsResult
    For mode = 1 To 22
        Fixture 1, 1: b.Records(1).SupplierTextRaw = "NO.999999": Prepare
        Select Case mode
            Case 1: a.Status = VATS2_SNAPSHOT_READ_ERROR
            Case 2: b.Status = VATS2_SNAPSHOT_READ_ERROR
            Case 3: f.Status = VATS2_FALLBACK_INVALID_INPUT
            Case 4: e.Status = VATS2_EFFECTIVE_INVALID_INPUT
            Case 5: m.Status = VATS2_EFFECTIVE_AMOUNT_INVALID_INPUT
            Case 6: p.Status = VATS2_SHORT_POLICY_INVALID_INPUT
            Case 7: e.AIntegrity.Status = VATS2_A_INVALID_RANGE
            Case 8: e.EffectiveConflicts.Status = VATS2_BCONFLICT_INVALID_RANGE
            Case 9: e.FullARecordCount = 2
            Case 10: m.FullBRecordCount = 2
            Case 11: a.Records(1).AIndex = 9
            Case 12: e.BRowIds(1) = 99
            Case 13: m.BInScope(1) = False
            Case 14: f.BToAFallbacks(1).MissingEvidence = False
            Case 15: f.AToBFallbacks(1).MissingEvidence = False
            Case 16: p.Decisions(1).ReferenceDigits = "0"
            Case 17: p.Decisions(1).BIndex = 2
            Case 18: p.Decisions(1).ReferenceIndex = 9
            Case 19: p.Decisions(1).AIndex = 1
            Case 20: e.EffectiveBMatches(1).References(1).MatcherFlags = 0
            Case 21: p.WaivedCount = 5
            Case 22: f.BToAFallbacks(1).ExcelRow = 999
        End Select
        r = VATStage2BuildAuditFindings(a, b, f, e, m, p)
        Check "拒绝损坏契约且无部分事项" & mode, r.Status <> VATS2_AUDIT_OK And r.FindingCount = 0 And r.ShortSuffixWaivedCount = 0 And Len(r.ErrorReason) > 0
    Next mode
    Fixture 2, 1: a.Records(2).InvoiceDigitsRaw = "10000001": Prepare
    e.AIntegrity.DuplicateGroups(1).AIndexes(2) = 99
    r = VATStage2BuildAuditFindings(a, b, f, e, m, p)
    Check "重复组A索引越界拒绝", r.Status = VATS2_AUDIT_INVALID_CONTRACT And r.FindingCount = 0
    Fixture 1, 1: a.Records(1).InvoiceDigitsRaw = "非法": b.Records(1).SupplierTextRaw = "NO.999999": Prepare
    f.ABlockers(1).ExcelRow = 99
    r = VATStage2BuildAuditFindings(a, b, f, e, m, p)
    Check "blocker位置错位拒绝", r.Status = VATS2_AUDIT_INVALID_CONTRACT
End Sub

Private Sub TestAdditionalEvidence()
    Dim r As VATS2AuditFindingsResult, pos As Long, i As Long
    Fixture 2, 1: b.Records(1).SupplierTextRaw = "NO.10000001/10000002": b.Records(1).AmountRaw = 3
    a.Records(2).AmountRaw = CVErr(2042): Prepare: r = RunAudit()
    pos = FindCode(r, VATS2_FIND_AMOUNT_ERROR)
    Check "非法A金额定位原A与原错误", r.Findings(pos).AIndex = 2 And r.Findings(pos).AExcelRow = 14 And r.Findings(pos).AmountErrorSide = "A" And r.Findings(pos).AmountErrorIndex = 2 And r.Findings(pos).Detail = m.GroupAmounts(1).Amount.ErrorReason
    Fixture 2, 1: a.Records(2).InvoiceDigitsRaw = "20000001": b.Records(1).SupplierTextRaw = "公司0000001": b.Records(1).IsCompleted = False
    Call Prepare: r = RunAudit()
    Check "full扫描范围外B解析风险仍保留", Not e.EffectiveBInScope(1) And r.CodeCounts(VATS2_FIND_NO_MARKER_MISSING) = 1 And r.CodeCounts(VATS2_FIND_AMBIGUOUS_MATCH) = 1
    Fixture 2, 2: a.Records(2).InvoiceDigitsRaw = "10000001": a.Records(2).IsCompleted = False
    b.Records(1).SupplierTextRaw = "NO.10000001/000001/000001": b.Records(2).SupplierTextRaw = "NO.10000001"
    Call Prepare: r = RunAudit()
    Check "Parser重复与SAME分别保留且CROSS不覆盖", r.CodeCounts(VATS2_FIND_DUPLICATE_REFERENCE) = 2 And r.CodeCounts(VATS2_FIND_CROSS_B_ROW_REUSE) = 1 And r.CodeCounts(VATS2_FIND_DUPLICATE_A_INVOICE) = 1
    Fixture 2, 1: a.Records(1).IsCompleted = False: b.Records(1).SupplierTextRaw = "NO.10000001/10000002": b.Records(1).AmountRaw = 3
    Call Prepare: r = RunAudit()
    Check "fallback完成整组后不误报INCOMPLETE", r.CodeCounts(VATS2_FIND_A_COLOR_MISSING) = 1 And r.CodeCounts(VATS2_FIND_INCOMPLETE_GROUP) = 0 And r.CodeCounts(VATS2_FIND_AMOUNT_MISMATCH) = 0
    Fixture 2, 1: b.Records(1).SupplierTextRaw = "NO.000001/000002": b.Records(1).AmountRaw = 3
    Call Prepare: r = RunAudit()
    Check "多个short豁免按引用计数", r.ShortSuffixWaivedCount = 2 And r.CodeCounts(VATS2_FIND_SHORT_SUFFIX_REVIEW) = 0
    For i = 1 To 7
        Fixture 2, 1: a.Records(1).InvoiceDigitsRaw = "非法": b.Records(1).SupplierTextRaw = "NO.99999999": Prepare
        Select Case i
            Case 1: f.AToBFallbacks(1).SourceBlocked = False: f.AToBFallbacks(1).MissingEvidence = True: f.AToBFallbacks(1).Evidence = VATS2_FULL_NOT_FOUND
            Case 2: f.BToAFallbacks(1).MissingEvidence = True: f.BToAFallbacks(1).Evidence = VATS2_FULL_NOT_FOUND
            Case 3: e.AIntegrity.InvalidAIndexes(1) = 0
            Case 4: f.ABlockers(1).OriginalIndex = 0
            Case 5: f.AToBFallbacks(1).AIndex = 0
            Case 6: f.FullBScanned = False
            Case 7: f.BToAFallbacks(1).Evidence = VATS2_FULL_MULTIPLE
        End Select
        r = VATStage2BuildAuditFindings(a, b, f, e, m, p)
        Check "矛盾missing或原索引拒绝" & i, r.Status = VATS2_AUDIT_INVALID_CONTRACT And r.FindingCount = 0
    Next i
    Fixture 2, 1: b.Records(1).AmountRaw = 9: Prepare: m.GroupAmounts(1).Amount.AIndexes(1) = 2
    r = VATStage2BuildAuditFindings(a, b, f, e, m, p)
    Check "金额成员不是原关系则拒绝", r.Status = VATS2_AUDIT_INVALID_CONTRACT
End Sub

Private Function RunAudit() As VATS2AuditFindingsResult
    Dim r As VATS2AuditFindingsResult, again As VATS2AuditFindingsResult, before As String, i As Long, total As Long
    before = AllInputsKey()
    r = VATStage2BuildAuditFindings(a, b, f, e, m, p)
    Check "合法上游正常生成", r.Status = VATS2_AUDIT_OK
    If r.Status <> VATS2_AUDIT_OK Then Err.Raise 5, , r.ErrorSide & r.ErrorIndex & "/" & r.ErrorReferenceIndex & ": " & r.ErrorReason
    again = VATStage2BuildAuditFindings(a, b, f, e, m, p)
    Check "六份输入全部字段保持不变", before = AllInputsKey()
    Check "重复调用全部输出确定", AuditKey(r) = AuditKey(again)
    For i = 1 To 18: total = total + r.CodeCounts(i): Next i
    Check "各Code计数与真实数量一致", total = r.FindingCount
    RunAudit = r
End Function

Private Function AllInputsKey() As String
    AllInputsKey = InputKey(a, b) & FallbackKey(f, a.RecordCount, b.RecordCount) & EffectiveKey(e) & ResultKey(m) & PolicyKey(p)
End Function

Private Function FindCode(ByRef r As VATS2AuditFindingsResult, ByVal code As VATS2FindingCode) As Long
    Dim i As Long
    For i = 1 To r.FindingCount
        If r.Findings(i).Code = code Then FindCode = i: Exit Function
    Next i
    Err.Raise 5, , "未找到预期FindingCode=" & code
End Function

Private Function AuditKey(ByRef r As VATS2AuditFindingsResult) As String
    Dim s As String, i As Long, j As Long
    s = r.Status & ":" & r.FindingCount & ":" & r.ShortSuffixWaivedCount & r.ErrorSide & r.ErrorIndex & r.ErrorReferenceIndex & r.ErrorReason
    For i = 1 To 18: s = s & ":" & r.CodeCounts(i): Next i
    For i = 1 To r.FindingCount
        With r.Findings(i)
            s = s & "|" & .Code & .Side & ":" & .AIndex & ":" & .AExcelRow & ":" & .BIndex & ":" & .BExcelRow & ":" & .ReferenceIndex & ":" & .ReferenceDigits
            s = s & ":" & .SourceFlags & ":" & .ReasonFlags & ":" & .ParserFlags & ":" & .MatcherFlags & ":" & .MatchKind & ":" & .GroupIndex & ":" & .MemberCount & ":" & RawKey(.Difference) & ":" & .GroupState & ":" & .AmountStatus & ":" & .Detail
            s = s & ":" & .AmountErrorSide & ":" & .AmountErrorIndex & ":" & .RawVarType
            For j = 1 To .MemberCount
                With .Members(j): s = s & "#" & .AIndex & ":" & .AExcelRow & ":" & .BIndex & ":" & .BExcelRow & ":" & .ReferenceIndex & ":" & .ReferenceDigits: End With
            Next j
        End With
    Next i
    AuditKey = s
End Function

Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then passed = passed + 1 Else failed = failed + 1: log = log & "FAIL " & name & vbCrLf
End Sub

Private Function PolicyKey(ByRef r As VATS2ShortSuffixPolicyResult) As String
    Dim s As String, i As Long
    s = r.Status & ":" & r.AggregateReliableEqual & ":" & r.DecisionCount & ":" & r.WaivedCount & ":" & r.ReviewRequiredCount & ":" & r.ErrorBIndex & ":" & r.ErrorReferenceIndex & ":" & r.ErrorReason
    For i = 1 To r.DecisionCount
        With r.Decisions(i)
            s = s & "|" & .BIndex & ":" & .ExcelRow & ":" & .ReferenceIndex & ":" & .AIndex & ":" & .ReferenceDigits & ":" & .Decision & ":" & .ReasonFlags
        End With
    Next i
    PolicyKey = s
End Function

Private Function EffectiveKey(ByRef r As VATS2EffectiveRelationsResult) As String
    Dim s As String, i As Long, j As Long
    s = r.Status & ":" & r.FullARecordCount & ":" & r.FullBRecordCount & ":" & r.EffectiveBCount & r.ErrorSide & r.ErrorIndex & r.ErrorReferenceIndex & r.ErrorReason
    For i = 1 To r.FullBRecordCount
        s = s & "|" & r.EffectiveBInScope(i) & ":" & r.BRowIds(i) & ":" & r.RowIssueFlags(i) & RowKey(r.EffectiveBMatches(i))
    Next i
    For i = 1 To r.FullARecordCount: s = s & "|" & r.AInvalidValue(i) & ":" & r.ADuplicateGroupIndex(i): Next i
    With r.AIntegrity
        s = s & .Status & ":" & .RecordCount & ":" & .Flags & ":" & .InvalidValueCount & ":" & .DuplicateGroupCount
        For i = 1 To .InvalidValueCount: s = s & "I" & .InvalidAIndexes(i): Next i
        For i = 1 To .DuplicateGroupCount
            s = s & .DuplicateGroups(i).InvoiceDigits & ":" & .DuplicateGroups(i).OccurrenceCount
            For j = 1 To .DuplicateGroups(i).OccurrenceCount: s = s & "D" & .DuplicateGroups(i).AIndexes(j): Next j
        Next i
    End With
    For i = 1 To r.RelationIssueCount
        With r.RelationIssues(i): s = s & "|" & .BIndex & ":" & .ExcelRow & ":" & .ReferenceIndex & ":" & .AIndex & ":" & .AExcelRow & ":" & .Flags: End With
    Next i
    With r.EffectiveConflicts
        s = s & .Status & ":" & .BRecordCount & ":" & .UniqueAssociationCount & ":" & .ConflictGroupCount & ":" & .ErrorBIndex & ":" & .ErrorReferenceIndex
        For i = 1 To .ConflictGroupCount
            With .ConflictGroups(i)
                s = s & .AIndex & ":" & .AInvoiceDigits & ":" & .AssociationCount & ":" & .DistinctBRowCount & ":" & .Flags
                For j = 1 To .AssociationCount
                    With .Associations(j): s = s & "|" & .BIndex & ":" & .BRowId & ":" & .ReferenceIndex & .ReferenceDigits & ":" & .AIndex & .AInvoiceDigits & ":" & .MatchKind: End With
                Next j
            End With
        Next i
    End With
    EffectiveKey = s
End Function

Private Function RawKey(ByVal value As Variant) As String
    If IsNull(value) Then
        RawKey = "Null"
    Else
        RawKey = VarType(value) & ":" & CStr(value)
    End If
End Function

Private Function InputKey(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult) As String
    Dim s As String, i As Long
    s = a.Status & a.SheetName & a.HeaderRow & a.RecordCount & a.ErrorRow & a.ErrorReason & b.Status & b.SheetName & b.HeaderRow & b.RecordCount & b.ErrorRow & b.ErrorReason
    For i = 1 To a.RecordCount
        With a.Records(i)
            s = s & "|" & .AIndex & ":" & .ExcelRow & RawKey(.InvoiceDigitsRaw) & RawKey(.AmountRaw) & .IsCompleted & .InvoiceCellAddress & .AmountCellAddress & .InvoiceHasFormula & .AmountHasFormula
        End With
    Next i
    For i = 1 To b.RecordCount
        With b.Records(i)
            s = s & "|" & .BIndex & ":" & .ExcelRow & RawKey(.SupplierTextRaw) & RawKey(.AmountRaw) & .IsCompleted & .SupplierCellAddress & .AmountCellAddress & .SupplierHasFormula & .AmountHasFormula
        End With
    Next i
    InputKey = s
End Function

Private Function RowKey(ByRef row As VATS2BRowMatchResult) As String
    Dim s As String, i As Long, j As Long
    s = row.OriginalText & "|" & row.ParserFlags & ":" & row.MatcherFlags & ":" & row.MatchState & ":" & row.ReferenceCount & ":" & row.UniqueMatchCount & ":" & row.NotFoundCount & ":" & row.MultipleMatchCount & ":" & row.InvalidMatchCount
    For i = 1 To row.ReferenceCount
        With row.References(i)
            s = s & "|" & .Digits & ":" & .Length & .RawFragment & ":" & .StartIndex & .Context & ":" & .ParserFlags & ":" & .MatcherFlags & ":" & .MatchKind & ":" & .CandidateCount
            For j = 1 To .CandidateCount
                s = s & "|" & .Candidates(j).AIndex & ":" & .Candidates(j).InvoiceDigits
            Next j
        End With
    Next i
    RowKey = s
End Function

Private Function GroupKey(ByRef g As VATS2GroupAmountResult) As String
    Dim s As String, i As Long
    s = g.BIndex & ":" & g.SourceMatchState & ":" & g.SourceConflictStatus & ":" & g.ReferenceCount & ":" & g.UniqueMatchCount & ":" & g.NotFoundCount & ":" & g.MultipleMatchCount & ":" & g.InvalidMatchCount
    s = s & "|" & g.ParserFlags & ":" & g.MatcherFlags & ":" & g.SameBRowReuse & ":" & g.CrossBRowReuse & ":" & g.GroupState & ":" & g.Reasons & ":" & g.ErrorReferenceIndex & ":" & g.ErrorAIndex & ":" & g.AmountEvaluated
    With g.Amount
        s = s & "|" & .ACount & ":" & .Status & ":" & .AmountState & ":" & .ErrorSide & ":" & .ErrorIndex & ":" & .ErrorAIndex & ":" & .ErrorReason
        s = s & "|" & RawKey(.AAmountSum) & "|" & RawKey(.BAmount) & "|" & RawKey(.Difference)
        For i = 1 To .ACount: s = s & "|" & .AIndexes(i) & ":" & RawKey(.AAmounts(i)): Next i
    End With
    GroupKey = s
End Function

Private Function ResultKey(ByRef r As VATS2EffectiveAmountResult) As String
    Dim s As String, i As Long
    s = r.Status & ":" & r.FullARecordCount & ":" & r.FullBRecordCount & ":" & r.EffectiveBCount & ":" & r.ComparedCount & ":" & r.EqualCount & ":" & r.MismatchCount & ":" & r.NotComparableCount & ":" & r.AmountErrorCount & ":" & r.RelationQualityBlockedCount & ":" & r.InvalidGroupCount
    s = s & "|" & r.ErrorSide & ":" & r.ErrorIndex & ":" & r.ErrorReferenceIndex & ":" & r.ErrorReason & ":" & r.RelationIssueCount
    For i = 1 To r.FullBRecordCount
        s = s & "|" & r.BInScope(i) & ":" & r.BRowIds(i) & ":" & r.RowStates(i) & ":" & r.RelationQualityFlags(i) & GroupKey(r.GroupAmounts(i))
    Next i
    For i = 1 To r.RelationIssueCount
        With r.RelationIssues(i): s = s & "|" & .BIndex & ":" & .ExcelRow & ":" & .ReferenceIndex & ":" & .AIndex & ":" & .AExcelRow & ":" & .Flags: End With
    Next i
    ResultKey = s
End Function

Private Function FallbackKey(ByRef r As VATS2FullFallbackResult, ByVal ac As Long, ByVal bc As Long) As String
    Dim s As String, i As Long, j As Long
    s = r.Status & ":" & r.BToAFallbackCount & ":" & r.AToBFallbackCount & ":" & r.AColorMissingCount & ":" & r.BColorMissingCount & r.FullBScanned & r.ErrorSide & r.ErrorIndex & r.ErrorReferenceIndex & r.ErrorReason
    For i = 1 To r.BToAFallbackCount
        With r.BToAFallbacks(i)
            s = s & "|" & .BIndex & ":" & .ExcelRow & ":" & .ReferenceIndex & RowKey(.OriginalBMatch) & .Evidence & .MissingEvidence
            s = s & .FullAMatch.ReferenceDigits & ":" & .FullAMatch.ReferenceLength & ":" & .FullAMatch.MatchKind & ":" & .FullAMatch.Flags & ":" & .FullAMatch.CandidateCount
            For j = 1 To .FullAMatch.CandidateCount
                s = s & "|" & .FullAMatch.Candidates(j).AIndex & .FullAMatch.Candidates(j).InvoiceDigits & .Candidates(j).AIndex & ":" & .Candidates(j).ExcelRow & .Candidates(j).InvoiceDigits & .Candidates(j).IsCompleted
            Next j
        End With
    Next i
    For i = 1 To r.AToBFallbackCount
        With r.AToBFallbacks(i)
            s = s & "|" & .AIndex & ":" & .ExcelRow & ":" & .Evidence & .SourceBlocked & .MissingEvidence & .HitCount
            For j = 1 To .HitCount
                s = s & "|" & .Hits(j).BIndex & ":" & .Hits(j).ExcelRow & .Hits(j).IsCompleted & .Hits(j).UniqueReferenceCount & ":" & .Hits(j).MultipleReferenceCount
            Next j
        End With
    Next i
    s = s & r.ABlockerCount & ":" & r.BBlockerCount
    For i = 1 To r.ABlockerCount
        With r.ABlockers(i): s = s & "|" & .OriginalIndex & ":" & .ExcelRow & ":" & .RawVarType & ":" & .Reasons & ":" & .ParserFlags: End With
    Next i
    For i = 1 To r.BBlockerCount
        With r.BBlockers(i): s = s & "|" & .OriginalIndex & ":" & .ExcelRow & ":" & .RawVarType & ":" & .Reasons & ":" & .ParserFlags: End With
    Next i
    For i = 1 To ac: s = s & "C" & r.CompletedACoverage(i): Next i
    If r.FullBScanned Then
        For i = 1 To bc: s = s & RowKey(r.FullBMatches(i)): Next i
    End If
    FallbackKey = s
End Function
