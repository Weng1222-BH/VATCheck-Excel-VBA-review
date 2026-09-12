Attribute VB_Name = "modVATStage2ShortPolicyTests"
Option Explicit

Private passed As Long, failed As Long, log As String

Public Function VATStage2ShortSuffixPolicy_SelfTest() As String
    On Error GoTo Unexpected
    passed = 0: failed = 0: log = ""
    TestPolicy
    TestReferences
    TestContracts
    GoTo Done
Unexpected:
    failed = failed + 1: log = log & "UNEXPECTED " & Err.Number & " " & Err.Description & vbCrLf
Done:
    If failed = 0 Then
        VATStage2ShortSuffixPolicy_SelfTest = "PASS: " & passed & " assertions" & vbCrLf & log
    Else
        VATStage2ShortSuffixPolicy_SelfTest = "FAIL: " & failed & "; PASS: " & passed & vbCrLf & log
    End If
End Function

Private Sub TestPolicy()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, e As VATS2EffectiveRelationsResult
    Dim m As VATS2EffectiveAmountResult, r As VATS2ShortSuffixPolicyResult, mode As Long
    Init a, b, 2, 1
    PreparePolicy a, b, e, m: r = RunPolicy(a, b, e, m, True)
    Check "干净短尾号本次豁免", r.DecisionCount = 1 And r.WaivedCount = 1 And r.Decisions(1).Decision = VATS2_WAIVED_THIS_RUN And r.Decisions(1).ReasonFlags = 0
    Check "保留原索引行号与引用Digits", r.Decisions(1).AIndex = 1 And r.Decisions(1).BIndex = 1 And r.Decisions(1).ExcelRow = 23 And r.Decisions(1).ReferenceIndex = 1 And r.Decisions(1).ReferenceDigits = "000001"
    r = RunPolicy(a, b, e, m, False)
    Check "总体不具备资格只有统一原因", r.Decisions(1).Decision = VATS2_REVIEW_REQUIRED And r.Decisions(1).ReasonFlags = VATS2_AGGREGATE_NOT_ELIGIBLE
    '同一来源随本次资格变化可以再次豁免；不持久化或清除旧Matcher标志。
    r = RunPolicy(a, b, e, m, True)
    Check "资格只属于本次运行", r.WaivedCount = 1 And e.EffectiveBMatches(1).References(1).MatcherFlags = VATS2_SHORT_SUFFIX_REVIEW
    For mode = 1 To 11
        Init a, b, 2, 1
        Select Case mode
            Case 1: b.Records(1).AmountRaw = 9
            Case 2: b.Records(1).SupplierTextRaw = "NO.000001/999999"
            Case 3: b.Records(1).SupplierTextRaw = "公司000001"
            Case 4: a.Records(1).InvoiceDigitsRaw = "XX000001"
            Case 5: b.Records(1).SupplierTextRaw = "NO.000001/10000001"
            Case 6: a.Records(1).IsCompleted = False
            Case 7: b.Records(1).IsCompleted = False
            Case 8: b.Records(1).AmountRaw = CVErr(2042)
            Case 9: a.Records(1).AmountRaw = "1"
            Case 10: a.Records(2).InvoiceDigitsRaw = "10000001": a.Records(2).IsCompleted = False
            Case 11: b.Records(1).SupplierTextRaw = "NO.000001--12345"
        End Select
        PreparePolicy a, b, e, m: r = RunPolicy(a, b, e, m, True)
        Check "非干净组不得豁免" & mode, r.WaivedCount = 0 And r.ReviewRequiredCount = r.DecisionCount And r.DecisionCount >= 1
        Select Case mode
            Case 1: Check "组金额不等", r.Decisions(1).ReasonFlags = VATS2_GROUP_NOT_EQUAL
            Case 2: Check "不完整组不能部分豁免", (r.Decisions(1).ReasonFlags And VATS2_ROW_NOT_COMPLETE_UNIQUE) <> 0 And r.Decisions(2).AIndex = 0 And r.Decisions(2).Decision = VATS2_REVIEW_REQUIRED
            Case 3, 11: Check "Parser风险保留", (r.Decisions(1).ReasonFlags And VATS2_PARSER_RISK) <> 0
            Case 4, 10: Check "关系质量风险阻止豁免", (r.Decisions(1).ReasonFlags And VATS2_RELATION_QUALITY_RISK) <> 0 And (r.Decisions(1).ReasonFlags And VATS2_GROUP_NOT_EQUAL) <> 0
            Case 5: Check "SAME复用风险", (r.Decisions(1).ReasonFlags And VATS2_CONFLICT_RISK) <> 0
            Case 6, 7: Check "缺色但金额相等仍需复核", m.EqualCount = 1 And r.Decisions(1).ReasonFlags = VATS2_COLOR_MISSING
            Case 8, 9: Check "金额异常不豁免", m.AmountErrorCount = 1 And (r.Decisions(1).ReasonFlags And VATS2_GROUP_NOT_EQUAL) <> 0
        End Select
    Next mode
    Init a, b, 1, 2
    b.Records(2).SupplierTextRaw = "NO.000001": b.Records(2).AmountRaw = 1
    PreparePolicy a, b, e, m: r = RunPolicy(a, b, e, m, True)
    Check "CROSS虽两组相等仍全部复核", m.EqualCount = 2 And r.ReviewRequiredCount = 2 And r.Decisions(1).ReasonFlags = VATS2_CONFLICT_RISK And r.Decisions(2).ReasonFlags = VATS2_CONFLICT_RISK
    Init a, b, 2, 1
    b.Records(1).SupplierTextRaw = "NO.000001/10000002": b.Records(1).AmountRaw = 3
    a.Records(2).IsCompleted = False
    PreparePolicy a, b, e, m: r = RunPolicy(a, b, e, m, True)
    Check "非short成员缺色也阻止本组short豁免", r.DecisionCount = 1 And m.EqualCount = 1 And r.Decisions(1).ReasonFlags = VATS2_COLOR_MISSING
    '模拟未来Matcher独立风险位：不定义新业务算法，未知身份风险只能复核。
    Init a, b, 1, 1: PreparePolicy a, b, e, m
    e.EffectiveBMatches(1).MatcherFlags = VATS2_SHORT_SUFFIX_REVIEW Or 2
    e.EffectiveBMatches(1).References(1).MatcherFlags = VATS2_SHORT_SUFFIX_REVIEW Or 2
    m.GroupAmounts(1).MatcherFlags = e.EffectiveBMatches(1).MatcherFlags
    r = RunPolicy(a, b, e, m, True)
    Check "未知其它身份结构风险不豁免", r.Decisions(1).ReasonFlags = VATS2_OTHER_STRUCTURE_RISK
End Sub

Private Sub TestReferences()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, e As VATS2EffectiveRelationsResult
    Dim m As VATS2EffectiveAmountResult, r As VATS2ShortSuffixPolicyResult
    Dim i As Long, combined As String
    Init a, b, 1, 1
    a.Records(1).InvoiceDigitsRaw = "000001"
    PreparePolicy a, b, e, m: r = RunPolicy(a, b, e, m, True)
    Check "短exact不生成决策", r.DecisionCount = 0
    b.Records(1).SupplierTextRaw = "NO.0000001": a.Records(1).InvoiceDigitsRaw = "10000001"
    PreparePolicy a, b, e, m: r = RunPolicy(a, b, e, m, True)
    Check "长suffix不生成决策", e.EffectiveBMatches(1).References(1).MatchKind = VATS2_SUFFIX_UNIQUE And r.DecisionCount = 0
    b.Records(1).SupplierTextRaw = "NO.10000001"
    PreparePolicy a, b, e, m: r = RunPolicy(a, b, e, m, False)
    Check "无short即使总体不具备资格仍0决策", r.DecisionCount = 0
    Init a, b, 3, 3
    b.Records(1).SupplierTextRaw = "NO.000002/000001": b.Records(1).AmountRaw = 3
    b.Records(2).SupplierTextRaw = "": b.Records(2).IsCompleted = False
    PreparePolicy a, b, e, m: r = RunPolicy(a, b, e, m, True)
    Check "多个short保留B和引用顺序", r.DecisionCount = 3 And r.WaivedCount = 3 And r.Decisions(1).AIndex = 2 And r.Decisions(2).AIndex = 1 And r.Decisions(2).ReferenceIndex = 2 And r.Decisions(3).BIndex = 3 And r.Decisions(3).ExcelRow = 29
    Init a, b, 2, 1
    a.Records(2).InvoiceDigitsRaw = "20000001"
    PreparePolicy a, b, e, m: r = RunPolicy(a, b, e, m, True)
    Check "short multiple不选择首候选", r.DecisionCount = 1 And r.Decisions(1).AIndex = 0 And r.ReviewRequiredCount = 1 And (r.Decisions(1).ReasonFlags And VATS2_ROW_NOT_COMPLETE_UNIQUE) <> 0
    b.Records(1).SupplierTextRaw = "NO.999999"
    PreparePolicy a, b, e, m: r = RunPolicy(a, b, e, m, False)
    Check "short NOT_FOUND仍有复核决策且理由组合", r.Decisions(1).AIndex = 0 And r.Decisions(1).ReasonFlags = (VATS2_AGGREGATE_NOT_ELIGIBLE Or VATS2_ROW_NOT_COMPLETE_UNIQUE Or VATS2_GROUP_NOT_EQUAL)
    Init a, b, 0, 0
    PreparePolicy a, b, e, m: r = RunPolicy(a, b, e, m, True)
    Check "双空输入正常0决策", r.Status = VATS2_SHORT_POLICY_OK And r.DecisionCount = 0
    Init a, b, 0, 1
    PreparePolicy a, b, e, m: r = RunPolicy(a, b, e, m, True)
    Check "短Digits但实际未带short标志不生成决策", e.EffectiveBMatches(1).References(1).MatcherFlags = 0 And r.DecisionCount = 0
    Init a, b, 69, 1
    combined = "NO."
    For i = 1 To 69
        If i > 1 Then combined = combined & "/"
        combined = combined & Right$(a.Records(i).InvoiceDigitsRaw, 6)
    Next i
    b.Records(1).SupplierTextRaw = combined: b.Records(1).AmountRaw = 2415
    PreparePolicy a, b, e, m: r = RunPolicy(a, b, e, m, True)
    Check "跨容量边界仍完整保序并裁剪输出", r.DecisionCount = 69 And r.WaivedCount = 69 And UBound(r.Decisions) = 69 And r.Decisions(65).ReferenceIndex = 65 And r.Decisions(69).AIndex = 69
End Sub

Private Sub TestContracts()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, e As VATS2EffectiveRelationsResult
    Dim m As VATS2EffectiveAmountResult, r As VATS2ShortSuffixPolicyResult, mode As Long
    For mode = 1 To 23
        Init a, b, 1, 1: PreparePolicy a, b, e, m
        Select Case mode
            Case 1: a.Status = VATS2_SNAPSHOT_READ_ERROR
            Case 2: b.Status = VATS2_SNAPSHOT_READ_ERROR
            Case 3: e.Status = VATS2_EFFECTIVE_INVALID_CONTRACT
            Case 4: m.Status = VATS2_EFFECTIVE_AMOUNT_GROUP_ERROR
            Case 5: m.FullBRecordCount = 2
            Case 6: e.FullARecordCount = 2
            Case 7: m.BInScope(1) = False
            Case 8: m.BRowIds(1) = 999
            Case 9: e.EffectiveBMatches(1).References(1).Candidates(1).AIndex = 2
            Case 10: m.RelationQualityFlags(1) = VATS2_RELATION_A_INVALID
            Case 11: m.GroupAmounts(1).BIndex = 2
            Case 12: m.GroupAmounts(1).AmountEvaluated = False
            Case 13: m.GroupAmounts(1).Amount.Status = VATS2_AMOUNT_INVALID_AMOUNT
            Case 14: m.GroupAmounts(1).Amount.AIndexes(1) = 2
            Case 15: e.EffectiveBMatches(1).References(1).Length = 7
            Case 16: e.EffectiveBMatches(1).References(1).MatchKind = VATS2_EXACT_UNIQUE
            Case 17: m.RowStates(1) = VATS2_EA_NOT_COMPARABLE
            Case 18: m.GroupAmounts(1).ParserFlags = 9
            Case 19: a.Records(1).AIndex = 2
            Case 20: e.EffectiveConflicts.Status = VATS2_BCONFLICT_INVALID_CONTRACT
            Case 21: Erase e.EffectiveBMatches
            Case 22: m.GroupAmounts(1).Amount.AmountState = VATS2_AMOUNT_NOT_COMPARED
            Case 23: m.GroupAmounts(1).Amount.ACount = 2
        End Select
        r = VATStage2EvaluateShortSuffixPolicy(a, b, e, m, True)
        Check "依赖契约损坏拒绝且无部分豁免" & mode, r.Status <> VATS2_SHORT_POLICY_OK And r.DecisionCount = 0 And r.WaivedCount = 0 And Len(r.ErrorReason) > 0
    Next mode
End Sub

Private Sub PreparePolicy(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, _
    ByRef e As VATS2EffectiveRelationsResult, ByRef m As VATS2EffectiveAmountResult)
    Prepare a, b, e
    m = VATStage2EvaluateEffectiveAmounts(a, b, e)
    If m.Status <> VATS2_EFFECTIVE_AMOUNT_OK Then Err.Raise 5, , "金额夹具失败：" & m.ErrorReason
End Sub

Private Function RunPolicy(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, _
    ByRef e As VATS2EffectiveRelationsResult, ByRef m As VATS2EffectiveAmountResult, ByVal aggregateEqual As Boolean) As VATS2ShortSuffixPolicyResult
    Dim before As String, r As VATS2ShortSuffixPolicyResult, again As VATS2ShortSuffixPolicyResult
    before = InputKey(a, b) & EffectiveKey(e) & ResultKey(m) & CStr(aggregateEqual)
    r = VATStage2EvaluateShortSuffixPolicy(a, b, e, m, aggregateEqual)
    If r.Status <> VATS2_SHORT_POLICY_OK Then Err.Raise 5, , "策略失败：" & r.ErrorBIndex & "/" & r.ErrorReferenceIndex & " " & r.ErrorReason
    Check "四份输入及所有原MatcherFlags不变", before = InputKey(a, b) & EffectiveKey(e) & ResultKey(m) & CStr(aggregateEqual)
    again = VATStage2EvaluateShortSuffixPolicy(a, b, e, m, aggregateEqual)
    Check "重复调用完整策略结果确定", PolicyKey(r) = PolicyKey(again)
    Check "策略数量与本次资格一致", r.WaivedCount + r.ReviewRequiredCount = r.DecisionCount And r.AggregateReliableEqual = aggregateEqual
    RunPolicy = r
End Function

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

Private Sub Init(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, ByVal ac As Long, ByVal bc As Long)
    Dim aa As VATS2ASnapshotResult, bb As VATS2BSnapshotResult, i As Long
    a = aa: b = bb
    a.Status = VATS2_SNAPSHOT_OK: b.Status = VATS2_SNAPSHOT_OK
    a.SheetName = "合成A": b.SheetName = "合成B": a.HeaderRow = 3: b.HeaderRow = 5
    a.RecordCount = ac: b.RecordCount = bc
    If ac > 0 Then ReDim a.Records(1 To ac)
    If bc > 0 Then ReDim b.Records(1 To bc)
    For i = 1 To ac
        With a.Records(i)
            .AIndex = i: .ExcelRow = 10 + i * 2: .IsCompleted = True
            .InvoiceDigitsRaw = "1000000" & i: .AmountRaw = i
            .InvoiceCellAddress = "D" & .ExcelRow: .AmountCellAddress = "I" & .ExcelRow
            .AmountHasFormula = (i Mod 2 = 0)
        End With
    Next i
    For i = 1 To bc
        With b.Records(i)
            .BIndex = i: .ExcelRow = 20 + i * 3: .IsCompleted = True
            .SupplierTextRaw = "NO.00000" & i: .AmountRaw = i
            .SupplierCellAddress = "B" & .ExcelRow: .AmountCellAddress = "F" & .ExcelRow
        End With
    Next i
End Sub

Private Sub Prepare(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, ByRef e As VATS2EffectiveRelationsResult)
    Dim c As VATS2CompletedScopeResult, f As VATS2FullFallbackResult
    c = VATStage2MatchCompletedScope(a, b)
    f = VATStage2RunFullFallback(a, b, c)
    e = VATStage2BuildEffectiveRelations(a, b, c, f)
    If e.Status <> VATS2_EFFECTIVE_OK Then Err.Raise 5, , "上游夹具失败：" & e.ErrorReason
End Sub

Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        passed = passed + 1
    Else
        failed = failed + 1: log = log & "FAIL " & name & vbCrLf
    End If
End Sub

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
