Attribute VB_Name = "modVATStage2EffAmountTests"
Option Explicit

Private passed As Long, failed As Long, log As String

Public Function VATStage2EffectiveAmount_SelfTest() As String
    On Error GoTo Unexpected
    passed = 0: failed = 0: log = ""
    TestAmounts
    TestScope
    TestQuality
    TestContracts
    GoTo Done
Unexpected:
    failed = failed + 1: log = log & "UNEXPECTED " & Err.Number & " " & Err.Description & vbCrLf
Done:
    If failed = 0 Then
        VATStage2EffectiveAmount_SelfTest = "PASS: " & passed & " assertions" & vbCrLf & log
    Else
        VATStage2EffectiveAmount_SelfTest = "FAIL: " & failed & "; PASS: " & passed & vbCrLf & log
    End If
End Function

Private Sub TestAmounts()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, e As VATS2EffectiveRelationsResult
    Dim r As VATS2EffectiveAmountResult, i As Long
    Init a, b, 3, 3
    b.Records(1).IsCompleted = False: b.Records(3).IsCompleted = False
    b.Records(1).SupplierTextRaw = "": b.Records(3).SupplierTextRaw = ""
    b.Records(2).SupplierTextRaw = "NO.10000002"
    Prepare a, b, e: r = Run(a, b, e)
    Check "1A1B相等", r.EqualCount = 1 And r.ComparedCount = 1 And r.RowStates(2) = VATS2_EA_COMPARED
    Check "完整索引不压缩且范围外不比较", r.FullBRecordCount = 3 And r.EffectiveBCount = 1 And r.GroupAmounts(2).Amount.AIndexes(1) = 2 And r.GroupAmounts(2).BIndex = 2 And r.BRowIds(2) = 26 And r.RowStates(1) = VATS2_EA_OUT_OF_SCOPE And r.RowStates(3) = VATS2_EA_OUT_OF_SCOPE
    b.Records(2).AmountRaw = 3
    r = Run(a, b, e)
    Check "1A1B不等与负差额", r.MismatchCount = 1 And r.GroupAmounts(2).Amount.Difference = CDec(-1)
    b.Records(2).AmountRaw = 1: r = Run(a, b, e)
    Check "正差额保持A减B", r.GroupAmounts(2).Amount.Difference = CDec(1)
    b.Records(2).SupplierTextRaw = "NO.10000003/10000001": b.Records(2).AmountRaw = 4
    Prepare a, b, e: r = Run(a, b, e)
    Check "多A1B相等按引用顺序保留原AIndex", r.EqualCount = 1 And r.GroupAmounts(2).Amount.AIndexes(1) = 3 And r.GroupAmounts(2).Amount.AIndexes(2) = 1 And r.GroupAmounts(2).Amount.AAmounts(1) = CDec(3)
    b.Records(2).AmountRaw = 5: r = Run(a, b, e)
    Check "多A1B不等", r.MismatchCount = 1 And r.GroupAmounts(2).Amount.Difference = CDec(-1)
    a.Records(3).AmountRaw = 0.1: a.Records(1).AmountRaw = 0.2: b.Records(2).AmountRaw = 0.3
    r = Run(a, b, e)
    Check "0.1加0.2等于0.3且保留Decimal", r.EqualCount = 1 And VarType(r.GroupAmounts(2).Amount.AAmountSum) = vbDecimal And VarType(r.GroupAmounts(2).Amount.Difference) = vbDecimal And r.GroupAmounts(2).Amount.Difference = CDec(0)
    For i = 1 To 5
        Select Case i
            Case 1: a.Records(3).AmountRaw = "0.1"
            Case 2: a.Records(3).AmountRaw = Empty
            Case 3: a.Records(3).AmountRaw = CVErr(2042)
            Case 4: a.Records(3).AmountRaw = True
            Case 5: a.Records(3).AmountRaw = ""
        End Select
        r = Run(a, b, e)
        Check "非法A金额只返回错误" & i, r.AmountErrorCount = 1 And r.EqualCount = 0 And r.GroupAmounts(2).Amount.ErrorSide = "A" And r.GroupAmounts(2).Amount.ErrorAIndex = 3 And IsEmpty(r.GroupAmounts(2).Amount.Difference)
    Next i
    a.Records(3).AmountRaw = 0.1: b.Records(2).AmountRaw = "0.3"
    r = Run(a, b, e)
    Check "非法B金额完整传播", r.AmountErrorCount = 1 And r.GroupAmounts(2).Amount.ErrorSide = "B" And Not r.GroupAmounts(2).Amount.AmountState = VATS2_AMOUNT_MISMATCH
    Init a, b, 3, 3
    b.Records(1).AmountRaw = 1: b.Records(2).AmountRaw = 8: b.Records(3).AmountRaw = CVErr(2042)
    Prepare a, b, e: r = Run(a, b, e)
    Check "统计只表示数学诊断", r.EffectiveBCount = 3 And r.ComparedCount = 2 And r.EqualCount = 1 And r.MismatchCount = 1 And r.AmountErrorCount = 1
    Init a, b, 0, 0
    Prepare a, b, e: r = Run(a, b, e)
    Check "空输入正常零统计", r.Status = VATS2_EFFECTIVE_AMOUNT_OK And r.EffectiveBCount = 0 And r.ComparedCount = 0
    Init a, b, 0, 1
    Prepare a, b, e: r = Run(a, b, e)
    Check "空A非空B不比较", r.NotComparableCount = 1
    Init a, b, 1, 0
    Prepare a, b, e: r = Run(a, b, e)
    Check "非空A空B正常", r.FullARecordCount = 1 And r.FullBRecordCount = 0
End Sub

Private Sub TestScope()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, e As VATS2EffectiveRelationsResult
    Dim c As VATS2CompletedScopeResult, f As VATS2FullFallbackResult, r As VATS2EffectiveAmountResult
    Init a, b, 3, 2
    a.Records(2).IsCompleted = False: b.Records(2).IsCompleted = False
    b.Records(1).SupplierTextRaw = "NO.10000001/10000002/10000003": b.Records(1).AmountRaw = 6
    Prepare a, b, e
    c = VATStage2MatchCompletedScope(a, b): f = VATStage2RunFullFallback(a, b, c)
    r = Run(a, b, e)
    Check "fallback补全merged且缺色A参与完整组", f.AColorMissingCount = 1 And r.EqualCount = 1 And r.GroupAmounts(1).Amount.ACount = 3 And r.GroupAmounts(1).Amount.AIndexes(2) = 2 And r.GroupAmounts(1).Amount.AAmountSum = CDec(6)
    Init a, b, 2, 3
    b.Records(1).IsCompleted = False: b.Records(2).IsCompleted = False: b.Records(3).IsCompleted = False
    b.Records(1).SupplierTextRaw = "": b.Records(2).SupplierTextRaw = ""
    b.Records(3).SupplierTextRaw = "NO.10000001/10000002": b.Records(3).AmountRaw = 3
    Prepare a, b, e: r = Run(a, b, e)
    Check "B缺色整组正常参与", r.EffectiveBCount = 1 And r.BInScope(3) And r.EqualCount = 1 And r.GroupAmounts(3).Amount.ACount = 2
    b.Records(3).SupplierTextRaw = "NO.10000001/99999999"
    Prepare a, b, e: r = Run(a, b, e)
    Check "缺色B的INCOMPLETE不使用部分金额", r.NotComparableCount = 1 And Not r.GroupAmounts(3).AmountEvaluated And IsEmpty(r.GroupAmounts(3).Amount.Difference)
    Init a, b, 1, 2
    b.Records(1).SupplierTextRaw = "NO.10000001/000001": b.Records(1).AmountRaw = 2
    b.Records(2).SupplierTextRaw = "NO.10000001": b.Records(2).AmountRaw = 1
    Prepare a, b, e: r = Run(a, b, e)
    Check "SAME阻断且不去重凑金额", r.NotComparableCount = 1 And r.GroupAmounts(1).SameBRowReuse And r.GroupAmounts(1).CrossBRowReuse And Not r.GroupAmounts(1).AmountEvaluated
    Check "仅CROSS仍允许诊断比较", r.EqualCount = 1 And r.GroupAmounts(2).CrossBRowReuse And Not r.GroupAmounts(2).SameBRowReuse
    Init a, b, 1, 1
    b.Records(1).SupplierTextRaw = "公司000001--12345"
    Prepare a, b, e: r = Run(a, b, e)
    Check "风险元数据不因相等而清除或豁免", r.EqualCount = 1 And r.GroupAmounts(1).ParserFlags = (VATS2_NO_MARKER_MISSING Or VATS2_NUMERIC_TRAILER_REVIEW) And r.GroupAmounts(1).MatcherFlags = VATS2_SHORT_SUFFIX_REVIEW
    b.Records(1).SupplierTextRaw = "无号码": Prepare a, b, e: r = Run(a, b, e)
    Check "NO_REFERENCES不比較", r.NotComparableCount = 1 And r.GroupAmounts(1).Reasons = VATS2_GROUP_NO_REFERENCES
End Sub

Private Sub TestQuality()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, e As VATS2EffectiveRelationsResult
    Dim r As VATS2EffectiveAmountResult
    Init a, b, 4, 1
    a.Records(1).InvoiceDigitsRaw = "XX000001": b.Records(1).SupplierTextRaw = "NO.000001"
    Prepare a, b, e: r = Run(a, b, e)
    Check "非法A关系即使金额相等仍阻断", r.RelationQualityBlockedCount = 1 And r.RowStates(1) = VATS2_RELATION_QUALITY_BLOCKED And r.ComparedCount = 0 And Not r.GroupAmounts(1).AmountEvaluated
    Check "保留具体非法关联原位置", r.RelationQualityFlags(1) = VATS2_RELATION_A_INVALID And r.RelationIssueCount = 1 And r.RelationIssues(1).AIndex = 1 And r.RelationIssues(1).AExcelRow = 12 And r.RelationIssues(1).ExcelRow = 23
    a.Records(1).AmountRaw = CVErr(2042): b.Records(1).AmountRaw = "非法"
    r = Run(a, b, e)
    Check "关系门控先于金额错误", r.RelationQualityBlockedCount = 1 And r.AmountErrorCount = 0
    Init a, b, 4, 1
    a.Records(2).InvoiceDigitsRaw = "10000001": a.Records(2).IsCompleted = False
    Prepare a, b, e: r = Run(a, b, e)
    Check "第一层unique但完整A重复必须阻断", r.RelationQualityFlags(1) = VATS2_RELATION_A_DUPLICATE And r.RelationQualityBlockedCount = 1 And r.EqualCount = 0
    a.Records(2).InvoiceDigitsRaw = "00999999": a.Records(3).InvoiceDigitsRaw = "00999999"
    a.Records(4).InvoiceDigitsRaw = Empty
    Prepare a, b, e: r = Run(a, b, e)
    Check "无关A质量异常不阻断正常B", e.AIntegrity.InvalidValueCount = 1 And e.AIntegrity.DuplicateGroupCount = 1 And r.EqualCount = 1 And r.RelationQualityBlockedCount = 0
    a.Records(4).InvoiceDigitsRaw = "XX000004"
    b.Records(1).SupplierTextRaw = "NO.10000001/000004"
    Prepare a, b, e: r = Run(a, b, e)
    Check "组中任一坏关联阻断完整组", r.RelationQualityBlockedCount = 1 And r.RelationIssues(1).ReferenceIndex = 2 And r.ComparedCount = 0
    a.Records(2).InvoiceDigitsRaw = "10000001": a.Records(2).IsCompleted = False
    Prepare a, b, e: r = Run(a, b, e)
    Check "同B两个质量风险组合且保留全部问题", r.RelationIssueCount = 2 And r.RelationQualityFlags(1) = (VATS2_RELATION_A_INVALID Or VATS2_RELATION_A_DUPLICATE) And r.RelationIssues(1).AIndex = 1 And r.RelationIssues(2).AIndex = 4 And r.RelationQualityBlockedCount = 1
End Sub

Private Sub TestContracts()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, e As VATS2EffectiveRelationsResult
    Dim r As VATS2EffectiveAmountResult, mode As Long, label As String
    '每次从真实冻结上游创建合法结果后，只破坏一项契约。
    For mode = 1 To 47
        Init a, b, 2, 2
        b.Records(2).SupplierTextRaw = "NO.10000001"
        Prepare a, b, e
        label = "拒绝契约损坏" & mode
        Select Case mode
            Case 1: a.Status = VATS2_SNAPSHOT_READ_ERROR
            Case 2: b.Status = VATS2_SNAPSHOT_READ_ERROR
            Case 3: e.Status = VATS2_EFFECTIVE_INVALID_INPUT
            Case 4: e.FullARecordCount = 1
            Case 5: e.FullBRecordCount = 1
            Case 6: a.Records(1).AIndex = 2
            Case 7: b.Records(1).BIndex = 2
            Case 8: a.Records(2).ExcelRow = a.Records(1).ExcelRow
            Case 9: b.Records(2).ExcelRow = b.Records(1).ExcelRow
            Case 10: e.BRowIds(1) = 999
            Case 11: Erase e.EffectiveBInScope
            Case 12: ReDim Preserve e.EffectiveBMatches(1 To 3)
            Case 13: ReDim e.RowIssueFlags(0 To 1)
            Case 14: e.EffectiveBCount = 1
            Case 15: e.EffectiveBInScope(1) = False
            Case 16: e.EffectiveBMatches(1).References(1).Candidates(1).AIndex = 99
            Case 17: e.EffectiveBMatches(1).References(1).Candidates(1).InvoiceDigits = "00000000"
            Case 18: e.EffectiveBMatches(1).OriginalText = "错误原文"
            Case 19: e.EffectiveBMatches(1).UniqueMatchCount = 0
            Case 20: e.EffectiveBMatches(1).References(1).CandidateCount = 2
            Case 21: e.EffectiveBMatches(1).References(1).MatchKind = VATS2_INVALID_REFERENCE
            Case 22: e.EffectiveBMatches(1).MatcherFlags = 123
            Case 23: e.RowIssueFlags(1) = VATS2_RELATION_A_INVALID
            Case 24: e.RowIssueFlags(1) = 8
            Case 25: e.RelationIssueCount = 1
            Case 26: ReDim e.RelationIssues(1 To 1)
            Case 27: e.AIntegrity.Status = VATS2_A_INVALID_RANGE
            Case 28: e.AIntegrity.RecordCount = 1
            Case 29: e.AInvalidValue(1) = True
            Case 30: e.ADuplicateGroupIndex(1) = 1
            Case 31: Erase e.AInvalidValue
            Case 32: e.EffectiveConflicts.Status = VATS2_BCONFLICT_INVALID_CONTRACT
            Case 33: e.EffectiveConflicts.BRecordCount = 1
            Case 34: e.EffectiveConflicts.UniqueAssociationCount = 1
            Case 35: e.EffectiveConflicts.ConflictGroups(1).Associations(1).BRowId = 999
            Case 36: e.EffectiveConflicts.ConflictGroups(1).Associations(1).BIndex = 99
            Case 37: e.EffectiveConflicts.ConflictGroups(1).Associations(1).AIndex = 99
            Case 38: e.EffectiveConflicts.ConflictGroups(1).Associations(1).ReferenceIndex = 99
            Case 39: e.EffectiveConflicts.ConflictGroupCount = 0: Erase e.EffectiveConflicts.ConflictGroups
            Case 40: e.EffectiveConflicts.ConflictGroups(1).Flags = VATS2_SAME_B_ROW_REUSE
            Case 41: e.EffectiveConflicts.ConflictGroups(1).DistinctBRowCount = 1
            Case 42: ReDim Preserve a.Records(1 To 3)
            Case 43: Erase b.Records
            Case 44: e.RelationIssueCount = -1
            Case 45: e.EffectiveBMatches(1).References(1).Length = 99
            Case 46: e.AIntegrity.Flags = VATS2_A_DATA_REVIEW
            Case 47: e.EffectiveConflicts.ConflictGroups(1).Associations(1).ReferenceDigits = "999"
        End Select
        r = VATStage2EvaluateEffectiveAmounts(a, b, e)
        Check label, r.Status <> VATS2_EFFECTIVE_AMOUNT_OK And r.ComparedCount = 0 And r.FullBRecordCount = 0 And Len(r.ErrorReason) > 0
    Next mode
    For mode = 1 To 12
        Init a, b, 2, 1
        a.Records(1).InvoiceDigitsRaw = "XX000001": b.Records(1).SupplierTextRaw = "NO.000001"
        Prepare a, b, e
        Select Case mode
            Case 1: e.RelationIssues(1).BIndex = 2
            Case 2: e.RelationIssues(1).AIndex = 2
            Case 3: e.RelationIssues(1).ReferenceIndex = 2
            Case 4: e.RelationIssues(1).ExcelRow = 999
            Case 5: e.RelationIssues(1).AExcelRow = 999
            Case 6: e.RelationIssues(1).Flags = VATS2_RELATION_A_DUPLICATE
            Case 7: e.RowIssueFlags(1) = 0
            Case 8: e.RowIssueFlags(1) = 0: e.RelationIssueCount = 0: Erase e.RelationIssues
            Case 9: e.RelationIssueCount = 2: ReDim Preserve e.RelationIssues(1 To 2): e.RelationIssues(2) = e.RelationIssues(1)
            Case 10: e.AIntegrity.InvalidAIndexes(1) = 99
            Case 11: Erase e.RowIssueFlags
            Case 12: ReDim e.RelationIssues(0 To 0)
        End Select
        r = VATStage2EvaluateEffectiveAmounts(a, b, e)
        Check "质量门控元数据损坏不容绕过" & mode, r.Status = VATS2_EFFECTIVE_AMOUNT_INVALID_CONTRACT And r.ComparedCount = 0
    Next mode
    Init a, b, 0, 0: Prepare a, b, e
    ReDim a.Records(1 To 1)
    r = VATStage2EvaluateEffectiveAmounts(a, b, e)
    Check "零计数仍分配A数组拒绝", r.Status = VATS2_EFFECTIVE_AMOUNT_INVALID_CONTRACT
    Init a, b, 0, 0: Prepare a, b, e
    ReDim e.EffectiveBMatches(1 To 1)
    r = VATStage2EvaluateEffectiveAmounts(a, b, e)
    Check "零计数仍分配关系数组拒绝", r.Status = VATS2_EFFECTIVE_AMOUNT_INVALID_CONTRACT
End Sub

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
            .SupplierTextRaw = "NO.1000000" & i: .AmountRaw = i
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

Private Function Run(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, ByRef e As VATS2EffectiveRelationsResult) As VATS2EffectiveAmountResult
    Dim r As VATS2EffectiveAmountResult, again As VATS2EffectiveAmountResult
    Dim before As String, i As Long, amounts() As Variant, expected As VATS2GroupAmountResult
    before = InputKey(a, b) & EffectiveKey(e)
    r = VATStage2EvaluateEffectiveAmounts(a, b, e)
    If r.Status <> VATS2_EFFECTIVE_AMOUNT_OK Then Err.Raise 5, , "EffectiveAmount失败：" & r.ErrorSide & r.ErrorIndex & "/" & r.ErrorReferenceIndex & " " & r.ErrorReason
    Check "全部输入字段不变", before = InputKey(a, b) & EffectiveKey(e)
    again = VATStage2EvaluateEffectiveAmounts(a, b, e)
    Check "完整输出重复调用确定", ResultKey(r) = ResultKey(again)
    Check "各诊断统计守恒", r.EffectiveBCount = r.ComparedCount + r.NotComparableCount + r.AmountErrorCount + r.RelationQualityBlockedCount + r.InvalidGroupCount And r.ComparedCount = r.EqualCount + r.MismatchCount
    If a.RecordCount > 0 Then ReDim amounts(1 To a.RecordCount)
    For i = 1 To a.RecordCount: amounts(i) = a.Records(i).AmountRaw: Next i
    For i = 1 To b.RecordCount
        If r.BInScope(i) And r.RelationQualityFlags(i) = 0 Then
            expected = VATStage2EvaluateBGroupAmount(e.EffectiveBMatches(i), i, e.EffectiveConflicts, amounts, b.Records(i).AmountRaw)
            Check "冻结GroupAmount逐字段传播B" & i, GroupKey(expected) = GroupKey(r.GroupAmounts(i))
        End If
    Next i
    Run = r
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
