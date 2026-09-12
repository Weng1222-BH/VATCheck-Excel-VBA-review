Attribute VB_Name = "modVATStage2EffectiveTests"
Option Explicit

Private passed As Long, failed As Long, log As String

Public Function VATStage2EffectiveRelations_SelfTest() As String
    On Error GoTo Unexpected
    passed = 0: failed = 0: log = vbNullString
    TestPatch
    TestScopeAndConflicts
    TestIntegrity
    TestContracts
    GoTo Done
Unexpected:
    failed = failed + 1: log = log & "UNEXPECTED " & Err.Number & " " & Err.Description & vbCrLf
Done:
    If failed = 0 Then
        VATStage2EffectiveRelations_SelfTest = "PASS: " & passed & " assertions" & vbCrLf & log
    Else
        VATStage2EffectiveRelations_SelfTest = "FAIL: " & failed & "; PASS: " & passed & vbCrLf & log
    End If
End Function

Private Sub TestPatch()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, c As VATS2CompletedScopeResult
    Dim f As VATS2FullFallbackResult, r As VATS2EffectiveRelationsResult, i As Long
    Init a, b, 3, 3
    a.Records(1).InvoiceDigitsRaw = "00111111": a.Records(1).IsCompleted = True
    a.Records(2).InvoiceDigitsRaw = "00222222"
    a.Records(3).InvoiceDigitsRaw = "00333333": a.Records(3).IsCompleted = True
    b.Records(2).SupplierTextRaw = "NO.00111111/00333333": b.Records(2).IsCompleted = True
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "第一层全部unique逐字段保持", r.Status = VATS2_EFFECTIVE_OK And RowKey(r.EffectiveBMatches(2)) = RowKey(c.BMatches(2))
    Check "完整BIndex空间不压缩且未涉及B占位", r.EffectiveBCount = 1 And Not r.EffectiveBInScope(1) And r.EffectiveBInScope(2) And r.EffectiveBMatches(3).ReferenceCount = 0 And UBound(r.EffectiveBMatches) = 3
    b.Records(2).SupplierTextRaw = "NO.00111111/00222222/00333333"
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "merged B只补全中间NOT_FOUND", r.EffectiveBMatches(2).ReferenceCount = 3 And r.EffectiveBMatches(2).References(2).Candidates(1).AIndex = 2 And r.EffectiveBMatches(2).References(1).Candidates(1).AIndex = 1
    Check "补全后整行计数和MatchState刷新", r.EffectiveBMatches(2).UniqueMatchCount = 3 And r.EffectiveBMatches(2).NotFoundCount = 0 And r.EffectiveBMatches(2).MultipleMatchCount = 0 And r.EffectiveBMatches(2).InvalidMatchCount = 0 And r.EffectiveBMatches(2).MatchState = VATS2_BROW_ALL_UNIQUE
    For i = 1 To 3
        Check "Parser所有字段保持ref" & i, ParserKey(r.EffectiveBMatches(2).References(i)) = ParserKey(c.BMatches(2).References(i))
    Next i
    Check "前导零原AIndex与B行号不变", r.EffectiveBMatches(2).References(2).Candidates(1).InvoiceDigits = "00222222" And r.BRowIds(2) = b.Records(2).ExcelRow
    a.Records(3).IsCompleted = False
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "两个NOT_FOUND分别补全且计数守恒", f.BToAFallbackCount = 2 And r.EffectiveBMatches(2).UniqueMatchCount = 3 And r.EffectiveBMatches(2).References(3).Candidates(1).AIndex = 3
    a.Records(3).InvoiceDigitsRaw = "00222222"
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "fallback multiple及notfound整合后仍INCOMPLETE", r.EffectiveBMatches(2).UniqueMatchCount = 1 And r.EffectiveBMatches(2).MultipleMatchCount = 1 And r.EffectiveBMatches(2).NotFoundCount = 1 And r.EffectiveBMatches(2).MatchState = VATS2_BROW_INCOMPLETE
    Check "multiple全部候选保序不选择", r.EffectiveBMatches(2).References(2).Candidates(1).AIndex = 2 And r.EffectiveBMatches(2).References(2).Candidates(2).AIndex = 3
    Init a, b, 1, 1
    a.Records(1).InvoiceDigitsRaw = "123"
    b.Records(1).SupplierTextRaw = "NO.123": b.Records(1).IsCompleted = True
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "空scope转full短exact后flags仍来自冻结Matcher", r.EffectiveBMatches(1).MatcherFlags = 0 And r.EffectiveBMatches(1).References(1).MatchKind = VATS2_EXACT_UNIQUE
    a.Records(1).InvoiceDigitsRaw = "00123"
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "补全后MatcherFlags重新OR为SHORT_SUFFIX", c.BMatches(1).MatcherFlags = 0 And r.EffectiveBMatches(1).MatcherFlags = VATS2_SHORT_SUFFIX_REVIEW And r.EffectiveBMatches(1).References(1).MatcherFlags = VATS2_SHORT_SUFFIX_REVIEW
    Init a, b, 2, 1
    a.Records(1).InvoiceDigitsRaw = "99999999": a.Records(1).IsCompleted = True
    a.Records(2).InvoiceDigitsRaw = "123"
    b.Records(1).SupplierTextRaw = "NO.123": b.Records(1).IsCompleted = True
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "替换NOT_FOUND的旧短suffix flags被清除", c.BMatches(1).MatcherFlags = VATS2_SHORT_SUFFIX_REVIEW And r.EffectiveBMatches(1).MatcherFlags = 0
    Init a, b, 3, 1
    a.Records(1).InvoiceDigitsRaw = "11123456": a.Records(1).IsCompleted = True
    a.Records(2).InvoiceDigitsRaw = "123456"
    a.Records(3).IsCompleted = True
    b.Records(1).SupplierTextRaw = "NO.123456": b.Records(1).IsCompleted = True
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "full exact不覆盖第一层suffix unique", f.FullBMatches(1).References(1).Candidates(1).AIndex = 2 And r.EffectiveBMatches(1).References(1).Candidates(1).AIndex = 1 And r.EffectiveBMatches(1).References(1).MatchKind = VATS2_SUFFIX_UNIQUE
    a.Records(3).InvoiceDigitsRaw = "22123456"
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "第一层multiple不被全表exact替换", c.BMatches(1).MultipleMatchCount = 1 And RowKey(c.BMatches(1)) = RowKey(r.EffectiveBMatches(1)) And r.EffectiveBMatches(1).References(1).Candidates(2).AIndex = 3
    Init a, b, 1, 1
    a.Records(1).InvoiceDigitsRaw = "00123456"
    b.Records(1).SupplierTextRaw = "公司00123456--12345": b.Records(1).IsCompleted = True
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "fallback补全不清除Parser来源风险及字段", r.EffectiveBMatches(1).ParserFlags = (VATS2_NO_MARKER_MISSING Or VATS2_NUMERIC_TRAILER_REVIEW) And ParserKey(r.EffectiveBMatches(1).References(1)) = ParserKey(c.BMatches(1).References(1))
End Sub

Private Sub TestScopeAndConflicts()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, c As VATS2CompletedScopeResult
    Dim f As VATS2FullFallbackResult, r As VATS2EffectiveRelationsResult
    Init a, b, 4, 5
    a.Records(2).InvoiceDigitsRaw = "00111111": a.Records(2).IsCompleted = True
    a.Records(3).InvoiceDigitsRaw = "00222222": a.Records(3).IsCompleted = True
    a.Records(4).InvoiceDigitsRaw = "33333333"
    b.Records(4).SupplierTextRaw = "NO.00111111/00222222/33333333/99999999"
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "同一缺色merged B被两个完成A命中只纳入一次", f.AToBFallbackCount = 2 And r.EffectiveBCount = 1 And r.EffectiveBInScope(4)
    Check "缺色B整行成功失败引用与候选完全保留", RowKey(r.EffectiveBMatches(4)) = RowKey(f.FullBMatches(4)) And r.EffectiveBMatches(4).ReferenceCount = 4 And r.EffectiveBMatches(4).NotFoundCount = 1
    Check "未涉及未完成B不进入effective或association", Not r.EffectiveBInScope(1) And r.EffectiveBMatches(1).ReferenceCount = 0 And r.EffectiveConflicts.UniqueAssociationCount = 3
    Check "纳入缺色B后原BIndex与ExcelRow仍保留", r.BRowIds(4) = b.Records(4).ExcelRow And r.EffectiveBMatches(4).References(1).Candidates(1).AIndex = 2
    a.Records(1).InvoiceDigitsRaw = "33333333"
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "缺色merged B中的multiple也保留", r.EffectiveBMatches(4).MultipleMatchCount = 1 And r.EffectiveBMatches(4).References(3).CandidateCount = 2
    Init a, b, 3, 5
    a.Records(3).InvoiceDigitsRaw = "00123456"
    b.Records(2).SupplierTextRaw = "NO.00123456": b.Records(2).IsCompleted = True
    b.Records(5).SupplierTextRaw = "NO.00123456": b.Records(5).IsCompleted = True
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "两个fallback B复用同A新产生CROSS", c.Conflicts.ConflictGroupCount = 0 And r.EffectiveConflicts.ConflictGroupCount = 1 And r.EffectiveConflicts.ConflictGroups(1).Flags = VATS2_CROSS_B_ROW_REUSE
    Check "post-fallback冲突保留原B2/B5和A3", r.EffectiveConflicts.ConflictGroups(1).AIndex = 3 And r.EffectiveConflicts.ConflictGroups(1).Associations(1).BIndex = 2 And r.EffectiveConflicts.ConflictGroups(1).Associations(2).BIndex = 5 And r.EffectiveConflicts.ConflictGroups(1).Associations(2).BRowId = b.Records(5).ExcelRow
    b.Records(5).IsCompleted = False
    b.Records(2).SupplierTextRaw = "NO.00123456/123456"
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "同B两个fallback ref指同A重扫产生SAME", r.EffectiveConflicts.ConflictGroups(1).Flags = VATS2_SAME_B_ROW_REUSE And r.EffectiveConflicts.ConflictGroups(1).Associations(2).BIndex = 2
    Init a, b, 2, 3
    a.Records(1).InvoiceDigitsRaw = "00111111": a.Records(1).IsCompleted = True
    a.Records(2).InvoiceDigitsRaw = "00222222": a.Records(2).IsCompleted = True
    b.Records(1).SupplierTextRaw = "NO.00111111": b.Records(1).IsCompleted = True
    b.Records(3).SupplierTextRaw = "NO.00111111/00222222"
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "completed B与缺色B之间新产生CROSS", r.EffectiveBCount = 2 And r.EffectiveConflicts.ConflictGroups(1).Flags = VATS2_CROSS_B_ROW_REUSE And r.EffectiveConflicts.ConflictGroups(1).Associations(2).BIndex = 3
End Sub

Private Sub TestIntegrity()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, c As VATS2CompletedScopeResult
    Dim f As VATS2FullFallbackResult, r As VATS2EffectiveRelationsResult
    Init a, b, 4, 2
    a.Records(1).InvoiceDigitsRaw = "00123456": a.Records(1).IsCompleted = True
    a.Records(2).InvoiceDigitsRaw = Empty
    a.Records(3).InvoiceDigitsRaw = "00999999": a.Records(4).InvoiceDigitsRaw = "00999999"
    b.Records(2).SupplierTextRaw = "NO.00123456": b.Records(2).IsCompleted = True
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "冻结AIntegrity完整保存非法与重复结果", r.AIntegrity.InvalidValueCount = 1 And r.AIntegrity.InvalidAIndexes(1) = 2 And r.AIntegrity.DuplicateGroupCount = 1 And r.AIntegrity.DuplicateGroups(1).AIndexes(2) = 4
    Check "原AIndex质量摘要准确", r.AInvalidValue(2) And Not r.AInvalidValue(1) And r.ADuplicateGroupIndex(3) = 1 And r.ADuplicateGroupIndex(4) = 1
    Check "无关问题A不删除正常关系", r.Status = VATS2_EFFECTIVE_OK And r.EffectiveBMatches(2).UniqueMatchCount = 1 And r.RowIssueFlags(2) = 0
    a.Records(1).InvoiceDigitsRaw = "XX123456"
    b.Records(2).SupplierTextRaw = "NO.123456"
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "真实suffix unique指向非法A标RELATION_A_INVALID", r.EffectiveBMatches(2).UniqueMatchCount = 1 And r.RowIssueFlags(2) = VATS2_RELATION_A_INVALID And r.AInvalidValue(1)
    Check "非法关联记录原索引行号与引用位置", r.RelationIssueCount = 1 And r.RelationIssues(1).BIndex = 2 And r.RelationIssues(1).AIndex = 1 And r.RelationIssues(1).ReferenceIndex = 1 And r.RelationIssues(1).AExcelRow = a.Records(1).ExcelRow
    a.Records(1).InvoiceDigitsRaw = "00999999"
    b.Records(2).SupplierTextRaw = "NO.00999999"
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "scope unique但全A重复成员标RELATION_A_DUPLICATE", c.BMatches(2).UniqueMatchCount = 1 And r.EffectiveBMatches(2).UniqueMatchCount = 1 And r.RowIssueFlags(2) = VATS2_RELATION_A_DUPLICATE And r.ADuplicateGroupIndex(1) = 1
    Init a, b, 0, 0
    Prepare a, b, c, f: r = Run(a, b, c, f)
    Check "双空输入正常零关系并执行完整性与冲突扫描", r.Status = VATS2_EFFECTIVE_OK And r.EffectiveBCount = 0 And r.AIntegrity.Status = VATS2_A_SCAN_OK And r.EffectiveConflicts.Status = VATS2_BCONFLICT_OK
End Sub

Private Sub TestContracts()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, c As VATS2CompletedScopeResult
    Dim f As VATS2FullFallbackResult, r As VATS2EffectiveRelationsResult, mode As Long
    For mode = 1 To 17
        Init a, b, 2, 1
        a.Records(1).InvoiceDigitsRaw = "00111111": a.Records(2).InvoiceDigitsRaw = "00222222"
        b.Records(1).SupplierTextRaw = "NO.00111111/00222222": b.Records(1).IsCompleted = True
        Prepare a, b, c, f
        Select Case mode
            Case 1: a.Status = VATS2_SNAPSHOT_READ_ERROR
            Case 2: c.Status = VATS2_SCOPE_INVALID_CONTRACT
            Case 3: f.Status = VATS2_FALLBACK_INVALID_INPUT
            Case 4: c.FullARecordCount = 3
            Case 5: a.Records(2).AIndex = 1
            Case 6: b.Records(1).BIndex = 5
            Case 7: f.BToAFallbackCount = 1
            Case 8: f.BToAFallbacks(2) = f.BToAFallbacks(1)
            Case 9: f.BToAFallbacks(1).BIndex = 2
            Case 10: f.BToAFallbacks(1).ReferenceIndex = 3
            Case 11: f.BToAFallbacks(1).FullAMatch.Candidates(1).AIndex = 3
            Case 12: f.BToAFallbacks(1).FullAMatch.ReferenceDigits = "错误引用"
            Case 13: c.BMatches(1).NotFoundCount = 0
            Case 14: Erase f.BToAFallbacks(1).FullAMatch.Candidates
            Case 15: c.BRowIds(1) = 999
            Case 16: f.BToAFallbacks(1).OriginalBMatch.ParserFlags = 999
            Case 17
                f.BToAFallbacks(1).FullAMatch.MatchKind = VATS2_INVALID_REFERENCE
                f.BToAFallbacks(1).FullAMatch.CandidateCount = 0
                Erase f.BToAFallbacks(1).FullAMatch.Candidates
        End Select
        r = VATStage2BuildEffectiveRelations(a, b, c, f)
        If mode <= 3 Then
            Check "非OK输入拒绝" & mode, r.Status = VATS2_EFFECTIVE_INVALID_INPUT
        Else
            Check "损坏patch映射契约拒绝且无部分关系" & mode, r.Status = VATS2_EFFECTIVE_INVALID_CONTRACT And r.EffectiveBCount = 0 And r.ErrorReason <> ""
        End If
    Next mode
    For mode = 1 To 9
        Init a, b, 1, 2
        a.Records(1).InvoiceDigitsRaw = "00111111": a.Records(1).IsCompleted = True
        b.Records(2).SupplierTextRaw = "NO.00111111"
        Prepare a, b, c, f
        Select Case mode
            Case 1: Erase f.FullBMatches
            Case 2: f.FullBMatches(2) = f.FullBMatches(1)
            Case 3: f.FullBMatches(2).UniqueMatchCount = 0
            Case 4: f.AToBFallbacks(1).Hits(1).IsCompleted = True
            Case 5: f.AToBFallbacks(1).Hits(1).BIndex = 1
            Case 6: f.AToBFallbacks(1).Hits(1).UniqueReferenceCount = 2
            Case 7: f.AToBFallbacks(1).Hits(1).ExcelRow = 99
            Case 8: f.BColorMissingCount = 0
            Case 9: f.FullBScanned = False
        End Select
        r = VATStage2BuildEffectiveRelations(a, b, c, f)
        Check "缺色B整行与Hit契约拒绝" & mode, r.Status = VATS2_EFFECTIVE_INVALID_CONTRACT And r.EffectiveBCount = 0
    Next mode
End Sub

Private Sub Prepare(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, ByRef c As VATS2CompletedScopeResult, ByRef f As VATS2FullFallbackResult)
    c = VATStage2MatchCompletedScope(a, b)
    f = VATStage2RunFullFallback(a, b, c)
    If c.Status <> VATS2_SCOPE_OK Or f.Status <> VATS2_FALLBACK_OK Then Err.Raise 5, , "上游测试夹具不是OK"
End Sub

Private Function Run(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, ByRef c As VATS2CompletedScopeResult, ByRef f As VATS2FullFallbackResult) As VATS2EffectiveRelationsResult
    Dim r As VATS2EffectiveRelationsResult, again As VATS2EffectiveRelationsResult
    Dim ab As String, cc As String, ff As String
    ab = InputKey(a, b): cc = CompletedKey(c): ff = FallbackKey(f, a.RecordCount, b.RecordCount)
    r = VATStage2BuildEffectiveRelations(a, b, c, f)
    If r.Status <> VATS2_EFFECTIVE_OK Then Err.Raise 5, , "Effective失败: " & r.ErrorSide & r.ErrorIndex & "/" & r.ErrorReferenceIndex & " " & r.ErrorReason
    Check "Snapshot所有字段不变", ab = InputKey(a, b)
    Check "completed完整输入不变", cc = CompletedKey(c)
    Check "fallback完整输入不变", ff = FallbackKey(f, a.RecordCount, b.RecordCount)
    again = VATStage2BuildEffectiveRelations(a, b, c, f)
    Check "Effective完整结果重复调用确定", EffectiveKey(r) = EffectiveKey(again)
    Run = r
End Function

Private Function ParserKey(ByRef item As VATS2BRowReference) As String
    ParserKey = item.Digits & ":" & item.Length & ":" & item.RawFragment & ":" & item.StartIndex & ":" & item.Context & ":" & item.ParserFlags
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

Private Sub Init(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, ByVal ac As Long, ByVal bc As Long)
    Dim emptyA As VATS2ASnapshotResult, emptyB As VATS2BSnapshotResult, i As Long
    a = emptyA: b = emptyB
    a.HeaderRow = 3: b.HeaderRow = 5: a.SheetName = "A": b.SheetName = "B"
    a.RecordCount = ac: b.RecordCount = bc
    If ac > 0 Then ReDim a.Records(1 To ac)
    If bc > 0 Then ReDim b.Records(1 To bc)
    For i = 1 To ac
        With a.Records(i)
            .AIndex = i: .ExcelRow = 10 + 2 * i
            .InvoiceDigitsRaw = "12345678": .AmountRaw = CVErr(2042)
            .InvoiceCellAddress = "D" & .ExcelRow: .AmountCellAddress = "I" & .ExcelRow
        End With
    Next i
    For i = 1 To bc
        With b.Records(i)
            .BIndex = i: .ExcelRow = 20 + 3 * i
            .SupplierTextRaw = "NO.99999999": .AmountRaw = "不允许处理金额"
            .SupplierCellAddress = "B" & .ExcelRow: .AmountCellAddress = "F" & .ExcelRow
        End With
    Next i
End Sub

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

Private Function CompletedKey(ByRef r As VATS2CompletedScopeResult) As String
    Dim s As String, i As Long, j As Long
    s = r.Status & ":" & r.FullARecordCount & ":" & r.FullBRecordCount & ":" & r.CompletedACount & ":" & r.CompletedBCount & r.ErrorSide & r.ErrorIndex & r.ErrorReason
    For i = 1 To r.CompletedACount: s = s & "|" & r.CompletedAIndexes(i) & ":" & r.CompletedAExcelRows(i): Next i
    For i = 1 To r.FullBRecordCount: s = s & "|" & r.BInScope(i) & ":" & r.BRowIds(i) & RowKey(r.BMatches(i)): Next i
    s = s & r.AIdentityIssueCount & ":" & r.BSourceIssueCount
    For i = 1 To r.AIdentityIssueCount
        With r.AIdentityIssues(i): s = s & "|" & .OriginalAIndex & ":" & .ExcelRow & ":" & .RawVarType: End With
    Next i
    For i = 1 To r.BSourceIssueCount
        With r.BSourceIssues(i): s = s & "|" & .OriginalBIndex & ":" & .ExcelRow & ":" & .RawVarType: End With
    Next i
    With r.Conflicts
        s = s & .Status & ":" & .BRecordCount & ":" & .UniqueAssociationCount & ":" & .ConflictGroupCount & ":" & .ErrorBIndex & ":" & .ErrorReferenceIndex
        For i = 1 To .ConflictGroupCount
            s = s & "|" & .ConflictGroups(i).AIndex & .ConflictGroups(i).AInvoiceDigits & ":" & .ConflictGroups(i).AssociationCount & ":" & .ConflictGroups(i).DistinctBRowCount & ":" & .ConflictGroups(i).Flags
            For j = 1 To .ConflictGroups(i).AssociationCount
                With .ConflictGroups(i).Associations(j)
                    s = s & "|" & .BIndex & ":" & .BRowId & ":" & .ReferenceIndex & .ReferenceDigits & ":" & .MatchKind & ":" & .AIndex & .AInvoiceDigits
                End With
            Next j
        Next i
    End With
    CompletedKey = s
End Function

Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        passed = passed + 1: log = log & "PASS " & name & vbCrLf
    Else
        failed = failed + 1: log = log & "FAIL " & name & vbCrLf
    End If
End Sub
