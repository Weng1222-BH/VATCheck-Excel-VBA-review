Attribute VB_Name = "modVATStage2FinalTests"
Option Explicit

Private passed As Long, failed As Long, log As String
Private ag As VATS2AggregateGateResult, am As VATS2EffectiveAmountResult, af As VATS2AuditFindingsResult

Public Function VATStage2FinalDecision_SelfTest() As String
    On Error GoTo Unexpected
    passed = 0: failed = 0: log = ""
    TestVerdicts
    TestUnavailable
    TestSummaries
    TestContracts
    GoTo Done
Unexpected:
    failed = failed + 1: log = log & "UNEXPECTED " & Err.Number & " " & Err.Description & vbCrLf
Done:
    If failed = 0 Then
        VATStage2FinalDecision_SelfTest = "PASS: " & passed & " assertions" & vbCrLf & log
    Else
        VATStage2FinalDecision_SelfTest = "FAIL: " & failed & "; PASS: " & passed & vbCrLf & log
    End If
End Function

Private Sub Fixture()
    Dim x As VATS2AggregateGateResult, y As VATS2EffectiveAmountResult, z As VATS2AuditFindingsResult
    ag = x: am = y: af = z
    ag.TotalA = CDec(3): ag.TotalB = CDec(3): ag.Difference = CDec(0): ag.TotalsEqual = True
    ag.AScanSucceeded = True: ag.BScanSucceeded = True: ag.CalculationState = xlDone
    ag.ACompletedCount = 2: ag.BCompletedCount = 1: ag.AIncludedCount = 2: ag.BIncludedCount = 1
    ag.ReliableEqualForShortSuffix = True
    am.FullARecordCount = 2: am.FullBRecordCount = 1: am.EffectiveBCount = 1: am.ComparedCount = 1: am.EqualCount = 1
    ReDim am.BInScope(1 To 1): ReDim am.BRowIds(1 To 1): ReDim am.RowStates(1 To 1)
    ReDim am.GroupAmounts(1 To 1): ReDim am.RelationQualityFlags(1 To 1)
    am.BInScope(1) = True: am.BRowIds(1) = 23: am.RowStates(1) = VATS2_EA_COMPARED
    With am.GroupAmounts(1)
        .BIndex = 1: .GroupState = VATS2_GROUP_COMPARED: .AmountEvaluated = True
        .Amount.ACount = 2: ReDim .Amount.AIndexes(1 To 2): ReDim .Amount.AAmounts(1 To 2)
        .Amount.AIndexes(1) = 1: .Amount.AIndexes(2) = 2: .Amount.AAmounts(1) = CDec(1): .Amount.AAmounts(2) = CDec(2)
        .Amount.AAmountSum = CDec(3): .Amount.BAmount = CDec(3): .Amount.Difference = CDec(0): .Amount.AmountState = VATS2_AMOUNT_EQUAL
    End With
End Sub

Private Sub AddFinding(ByVal code As Long)
    af.FindingCount = af.FindingCount + 1: ReDim Preserve af.Findings(1 To af.FindingCount)
    af.CodeCounts(code) = af.CodeCounts(code) + 1
    With af.Findings(af.FindingCount)
        .Code = code: .Side = "B": .BIndex = 1: .BExcelRow = 23: .ReferenceIndex = 2: .ReferenceDigits = "000002"
        .SourceFlags = VATS2_FIND_FROM_EFFECTIVE: .ParserFlags = 4: .MatcherFlags = 1: .Detail = "原始证据"
        .MemberCount = 1: ReDim .Members(1 To 1): .Members(1).AIndex = 2: .Members(1).AExcelRow = 14
    End With
End Sub

Private Sub Warning(ByVal side As String)
    If side = "A" Then
        ag.ACompletedCount = 3: ag.AEmptyWarningCount = 1: ag.AWarningCount = 1
        ReDim ag.AWarnings(1 To 1): ag.AWarnings(1) = "合成A!I9 空金额，已跳过未计入"
    Else
        ag.BCompletedCount = 2: ag.BEmptyWarningCount = 1: ag.BWarningCount = 1
        ReDim ag.BWarnings(1 To 1): ag.BWarnings(1) = "合成B!F9 空金额，已跳过未计入"
    End If
    ag.ReliableEqualForShortSuffix = False
End Sub

Private Sub TestVerdicts()
    Dim mode As Long, r As VATS2FinalDecisionResult
    For mode = 0 To 7
        Call Fixture
        If (mode And 1) <> 0 Then ag.TotalB = CDec(4): ag.Difference = CDec(-1): ag.TotalsEqual = False
        If (mode And 2) <> 0 Then AddFinding VATS2_FIND_AMOUNT_MISMATCH
        If (mode And 4) <> 0 Then Warning "A"
        r = RunFinal()
        If mode = 0 Then
            Check "可靠相等零事项零警告VERIFIED", r.Verdict = VATS2_FINAL_VERIFIED And Not r.RequiresManualReview
        Else
            Check "差异事项警告的所有组合均REVIEW" & mode, r.Verdict = VATS2_FINAL_REVIEW_REQUIRED And r.RequiresManualReview
        End If
        Check "三种独立提示完整保留" & mode, r.HasAggregateMismatch = ((mode And 1) <> 0) And r.HasFindings = ((mode And 2) <> 0) And r.HasAggregateWarnings = ((mode And 4) <> 0)
        Check "总体warning及差异不伪造Finding" & mode, r.FindingCount = af.FindingCount And r.CodeCounts(VATS2_FIND_AMOUNT_MISMATCH) = af.CodeCounts(VATS2_FIND_AMOUNT_MISMATCH)
    Next mode
    Call Fixture: Warning "B": r = RunFinal()
    Check "B警告独立阻止VERIFIED", r.Verdict = VATS2_FINAL_REVIEW_REQUIRED And r.BWarningCount = 1
    Call Fixture
    ag.ACompletedCount = 0: ag.BCompletedCount = 0: ag.AIncludedCount = 0: ag.BIncludedCount = 0
    ag.TotalA = CDec(0): ag.TotalB = CDec(0)
    am.FullARecordCount = 0: am.FullBRecordCount = 0: am.EffectiveBCount = 0: am.ComparedCount = 0: am.EqualCount = 0
    Erase am.BInScope: Erase am.BRowIds: Erase am.RowStates: Erase am.GroupAmounts: Erase am.RelationQualityFlags
    r = RunFinal(): Check "零完成可靠相等仍VERIFIED并注明", r.Verdict = VATS2_FINAL_VERIFIED And r.NoCompletedRecords
    Call Fixture: af.ShortSuffixWaivedCount = 2: r = RunFinal()
    Check "short豁免仅统计不改变VERIFIED", r.Verdict = VATS2_FINAL_VERIFIED And r.ShortSuffixWaivedCount = 2 And r.FindingCount = 0
    AddFinding VATS2_FIND_STRUCTURE_REVIEW: r = RunFinal()
    Check "short豁免不抵销现有事项", r.Verdict = VATS2_FINAL_REVIEW_REQUIRED And r.ShortSuffixWaivedCount = 2
    Call Fixture: ag.TotalA = CDec(9): ag.Difference = CDec(6): ag.TotalsEqual = False: r = RunFinal()
    Check "逐组全equal不覆盖总体不等", r.EqualCount = 1 And r.MismatchCount = 0 And r.Verdict = VATS2_FINAL_REVIEW_REQUIRED And r.Difference = CDec(6)
    Call Fixture: AddFinding VATS2_FIND_AMOUNT_MISMATCH: am.EqualCount = 0: am.MismatchCount = 1
    r = RunFinal(): Check "总体相等但逐组finding仍REVIEW", r.TotalsEqual And r.Verdict = VATS2_FINAL_REVIEW_REQUIRED
End Sub

Private Sub TestUnavailable()
    Dim mode As Long, r As VATS2FinalDecisionResult
    For mode = 1 To 3
        Call Fixture: ag.Status = mode
        '故意带残留金额和Boolean，验证非OK门绝不会把这些数据当总体结论。
        ag.TotalA = CDec(99): ag.TotalsEqual = False
        Warning "A": AddFinding VATS2_FIND_STRUCTURE_REVIEW
        r = RunFinal()
        Check "三种总体门失败均UNAVAILABLE" & mode, r.Verdict = VATS2_FINAL_UNAVAILABLE And r.Status = VATS2_FINAL_DECISION_OK
        Check "UNAVAILABLE金额和比较保持Empty" & mode, IsEmpty(r.TotalA) And IsEmpty(r.TotalB) And IsEmpty(r.Difference) And IsEmpty(r.TotalsEqual)
        Check "UNAVAILABLE不是人工复核或伪造零记录" & mode, Not r.RequiresManualReview And Not r.HasAggregateMismatch And Not r.NoCompletedRecords
        Check "UNAVAILABLE可保留原始诊断计数" & mode, r.HasFindings And r.HasAggregateWarnings And r.AggregateStatus = mode
        ag.ACompletedCount = 0: ag.BCompletedCount = 0: r = RunFinal()
        Check "未可靠扫描的零计数不判NoCompletedRecords" & mode, Not r.NoCompletedRecords
    Next mode
End Sub

Private Sub TestSummaries()
    Dim r As VATS2FinalDecisionResult, i As Long
    Call Fixture
    For i = 1 To 18: AddFinding i: Next i
    Warning "A": Warning "B"
    am.FullBRecordCount = 9: am.EffectiveBCount = 9: am.ComparedCount = 5: am.EqualCount = 3: am.MismatchCount = 2
    am.NotComparableCount = 1: am.AmountErrorCount = 2: am.RelationQualityBlockedCount = 1
    '本层只读摘要，无须逐行数组；建立占位仅让输入不变检查序列化全部空间。
    ReDim Preserve am.BInScope(1 To 9): ReDim Preserve am.BRowIds(1 To 9): ReDim Preserve am.RowStates(1 To 9)
    ReDim Preserve am.GroupAmounts(1 To 9): ReDim Preserve am.RelationQualityFlags(1 To 9)
    ag.TotalA = CDec("12345678901234567890.12345678"): ag.TotalB = ag.TotalA
    r = RunFinal()
    For i = 1 To 18: Check "CodeCounts原样传播" & i, r.CodeCounts(i) = af.CodeCounts(i): Next i
    Check "七个Stage2计数原样保留", r.EffectiveBCount = 9 And r.ComparedCount = 5 And r.EqualCount = 3 And r.MismatchCount = 2 And r.NotComparableCount = 1 And r.AmountErrorCount = 2 And r.RelationQualityBlockedCount = 1
    Check "完成计入warning统计原样保留", r.ACompletedCount = 3 And r.BCompletedCount = 2 And r.AIncludedCount = 2 And r.BIncludedCount = 1 And r.AWarningCount = 1 And r.BWarningCount = 1
    Check "金额直接复制且Decimal精度不变", VarType(r.TotalA) = vbDecimal And r.TotalA = ag.TotalA And r.TotalB = ag.TotalB And r.Difference = ag.Difference
    Call Fixture: ag.ACompletedCount = 0: ag.AIncludedCount = 0: r = RunFinal()
    Check "仅一侧零完成不判双方零记录", Not r.NoCompletedRecords
End Sub

Private Sub TestContracts()
    Dim r As VATS2FinalDecisionResult, i As Long
    For i = 1 To 33
        Call Fixture
        Select Case i
            Case 1: ag.Status = 99
            Case 2: am.Status = 99
            Case 3: af.Status = 99
            Case 4: ag.TotalsEqual = Empty
            Case 5: ag.TotalsEqual = "True"
            Case 6: ag.TotalsEqual = 1
            Case 7: ag.TotalsEqual = Null
            Case 8: ag.ACompletedCount = -1
            Case 9: ag.BCompletedCount = -1
            Case 10: ag.AIncludedCount = -1
            Case 11: ag.BIncludedCount = -1
            Case 12: ag.AWarningCount = -1
            Case 13: ag.BWarningCount = -1
            Case 14: ag.AEmptyWarningCount = -1
            Case 15: ag.AIssueCount = -1
            Case 16: am.FullARecordCount = -1
            Case 17: am.FullBRecordCount = -1
            Case 18: am.EffectiveBCount = -1
            Case 19: am.ComparedCount = -1
            Case 20: am.EqualCount = -1
            Case 21: am.MismatchCount = -1
            Case 22: am.NotComparableCount = -1
            Case 23: am.AmountErrorCount = -1
            Case 24: am.RelationQualityBlockedCount = -1
            Case 25: am.InvalidGroupCount = -1
            Case 26: af.FindingCount = -1
            Case 27: af.ShortSuffixWaivedCount = -1
            Case 28: af.CodeCounts(18) = -1
            Case 29: am.ComparedCount = 2
            Case 30: am.NotComparableCount = 1
            Case 31: am.FullBRecordCount = 0
            Case 32: af.FindingCount = 1
            Case 33: af.CodeCounts(18) = 2147483647: af.CodeCounts(1) = 2147483647: af.FindingCount = 2147483647
        End Select
        r = VATStage2BuildFinalDecision(ag, am, af)
        Check "必要契约损坏明确拒绝" & i, r.Status = VATS2_FINAL_INVALID_CONTRACT And r.Verdict = VATS2_FINAL_UNAVAILABLE And Not r.RequiresManualReview And IsEmpty(r.TotalsEqual) And r.FindingCount = 0 And Len(r.ErrorReason) > 0
    Next i
    For i = 1 To 5
        Call Fixture
        If i <= 3 Then am.Status = i Else af.Status = i - 3
        r = VATStage2BuildFinalDecision(ag, am, af)
        Check "失败上游不视为零异常" & i, r.Status = VATS2_FINAL_INVALID_INPUT And r.Verdict = VATS2_FINAL_UNAVAILABLE And Not r.RequiresManualReview
    Next i
    Call Fixture: am.FullBRecordCount = 2: am.EffectiveBCount = 2: am.InvalidGroupCount = 1
    r = VATStage2BuildFinalDecision(ag, am, af)
    Check "OK金额层不能带InvalidGroupCount", r.Status = VATS2_FINAL_INVALID_CONTRACT
End Sub

Private Function RunFinal() As VATS2FinalDecisionResult
    Dim r As VATS2FinalDecisionResult, again As VATS2FinalDecisionResult, before As String
    before = AggregateKey(ag) & ResultKey(am) & AuditKey(af)
    r = VATStage2BuildFinalDecision(ag, am, af)
    Check "合法摘要正常组合", r.Status = VATS2_FINAL_DECISION_OK
    again = VATStage2BuildFinalDecision(ag, am, af)
    Check "三个上游所有字段及嵌套数组不变", before = AggregateKey(ag) & ResultKey(am) & AuditKey(af)
    Check "重复调用全字段确定", FinalKey(r) = FinalKey(again)
    RunFinal = r
End Function

Private Function FinalKey(ByRef r As VATS2FinalDecisionResult) As String
    Dim s As String, i As Long
    s = r.Status & ":" & r.Verdict & ":" & r.AggregateStatus & ":" & r.EffectiveAmountStatus & ":" & r.AuditStatus
    s = s & "|" & RawKey(r.TotalA) & "|" & RawKey(r.TotalB) & "|" & RawKey(r.Difference) & "|" & RawKey(r.TotalsEqual)
    s = s & ":" & r.ACompletedCount & ":" & r.BCompletedCount & ":" & r.AIncludedCount & ":" & r.BIncludedCount & ":" & r.AWarningCount & ":" & r.BWarningCount
    s = s & ":" & r.EffectiveBCount & ":" & r.ComparedCount & ":" & r.EqualCount & ":" & r.MismatchCount & ":" & r.NotComparableCount & ":" & r.AmountErrorCount & ":" & r.RelationQualityBlockedCount
    s = s & ":" & r.FindingCount & ":" & r.ShortSuffixWaivedCount & ":" & r.HasAggregateMismatch & ":" & r.HasAggregateWarnings & ":" & r.HasFindings & ":" & r.NoCompletedRecords & ":" & r.RequiresManualReview & r.ErrorSource & r.ErrorReason
    For i = 1 To 18: s = s & ":" & r.CodeCounts(i): Next i
    FinalKey = s
End Function

Private Function AggregateKey(ByRef r As VATS2AggregateGateResult) As String
    Dim s As String, i As Long
    s = r.Status & ":" & r.CalculationState & ":" & r.AScanSucceeded & ":" & r.BScanSucceeded & ":" & r.ReliableEqualForShortSuffix & r.ErrorReason
    s = s & "|" & RawKey(r.TotalA) & "|" & RawKey(r.TotalB) & "|" & RawKey(r.Difference) & "|" & RawKey(r.TotalsEqual)
    s = s & ":" & r.ACompletedCount & ":" & r.BCompletedCount & ":" & r.AIncludedCount & ":" & r.BIncludedCount & ":" & r.AEmptyWarningCount & ":" & r.BEmptyWarningCount
    s = s & ":" & r.AIssueCount & ":" & r.BIssueCount & ":" & r.AWarningCount & ":" & r.BWarningCount
    For i = 1 To r.AIssueCount: s = s & "|" & r.AIssues(i): Next i
    For i = 1 To r.BIssueCount: s = s & "|" & r.BIssues(i): Next i
    For i = 1 To r.AWarningCount: s = s & "|" & r.AWarnings(i): Next i
    For i = 1 To r.BWarningCount: s = s & "|" & r.BWarnings(i): Next i
    AggregateKey = s
End Function

Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then passed = passed + 1 Else failed = failed + 1: log = log & "FAIL " & name & vbCrLf
End Sub

Private Function RawKey(ByVal value As Variant) As String
    If IsNull(value) Then
        RawKey = "Null"
    Else
        RawKey = VarType(value) & ":" & CStr(value)
    End If
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
