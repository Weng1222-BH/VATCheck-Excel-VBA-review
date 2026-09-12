Attribute VB_Name = "modVATStage2CompletedScopeTests"
Option Explicit

Private passed As Long, failed As Long, log As String

Public Function VATStage2CompletedScope_SelfTest() As String
    On Error GoTo Unexpected
    passed = 0: failed = 0: log = vbNullString
    TestMapping
    TestBoundaries
    TestIssues
    TestConflicts
    TestContracts
    TestRemap
    GoTo Done
Unexpected:
    failed = failed + 1: log = log & "UNEXPECTED " & Err.Number & " " & Err.Description & vbCrLf
Done:
    If failed = 0 Then
        VATStage2CompletedScope_SelfTest = "PASS: " & passed & " assertions" & vbCrLf & log
    Else
        VATStage2CompletedScope_SelfTest = "FAIL: " & failed & "; PASS: " & passed & vbCrLf & log
    End If
End Function

Private Sub TestMapping()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, r As VATS2CompletedScopeResult
    Dim again As VATS2CompletedScopeResult, raw As VATS2BRowMatchResult, expected As VATS2BRowMatchResult
    Dim digits(1 To 3) As String, map(1 To 3) As Long, before As String, i As Long, status As VATS2ScopeStatus
    Init a, b, 5, 8
    a.Records(1).InvoiceDigitsRaw = "00123456"
    a.Records(2).InvoiceDigitsRaw = "123456": a.Records(2).IsCompleted = False
    a.Records(3).InvoiceDigitsRaw = "111123456"
    a.Records(4).InvoiceDigitsRaw = "999999": a.Records(4).IsCompleted = False
    a.Records(5).InvoiceDigitsRaw = "222123456"
    For i = 1 To 8: b.Records(i).IsCompleted = False: Next i
    b.Records(1).SupplierTextRaw = "NO.00123456"
    b.Records(2).SupplierTextRaw = "NO.00123456": b.Records(2).IsCompleted = True
    b.Records(3).SupplierTextRaw = "NO.NO.((--123"
    b.Records(4).SupplierTextRaw = "NO.123456": b.Records(4).IsCompleted = True
    b.Records(5).SupplierTextRaw = "NO.00123456": b.Records(5).IsCompleted = True
    b.Records(6).SupplierTextRaw = CVErr(2042)
    b.Records(8).SupplierTextRaw = "NO.999999": b.Records(8).IsCompleted = True
    before = InputKey(a, b)
    r = VATStage2MatchCompletedScope(a, b)
    Check "完整数量与完成数量独立", r.Status = VATS2_SCOPE_OK And r.FullARecordCount = 5 And r.FullBRecordCount = 8 And r.CompletedACount = 3 And r.CompletedBCount = 4
    Check "完成A映射保留1,3,5", r.CompletedAIndexes(1) = 1 And r.CompletedAIndexes(2) = 3 And r.CompletedAIndexes(3) = 5
    Check "完成A映射保留ExcelRow", r.CompletedAExcelRows(2) = a.Records(3).ExcelRow And r.CompletedAExcelRows(3) = a.Records(5).ExcelRow
    Check "B完整索引空间没有压缩", LBound(r.BMatches) = 1 And UBound(r.BMatches) = 8 And r.BInScope(2) And r.BInScope(5) And r.BInScope(8) And Not r.BInScope(1)
    Check "完成色exact唯一且前导零不丢", r.BMatches(2).References(1).MatchKind = VATS2_EXACT_UNIQUE And r.BMatches(2).References(1).Candidates(1).InvoiceDigits = "00123456"
    Check "未完成A的exact不压制scope内suffix", r.BMatches(4).References(1).MatchKind = VATS2_SUFFIX_MULTIPLE And r.BMatches(4).References(1).CandidateCount = 3
    Check "multiple候选全部回映并保持顺序", r.BMatches(4).References(1).Candidates(1).AIndex = 1 And r.BMatches(4).References(1).Candidates(2).AIndex = 3 And r.BMatches(4).References(1).Candidates(3).AIndex = 5
    Check "SHORT_SUFFIX来自真实scope搜索", r.BMatches(4).MatcherFlags = VATS2_SHORT_SUFFIX_REVIEW
    Check "只存在未完成A时当前scope为NOT_FOUND", r.BMatches(8).References(1).MatchKind = VATS2_NOT_FOUND And r.BMatches(8).NotFoundCount = 1
    Check "未完成B合法匹配原文也没有被解析", r.BMatches(1).ReferenceCount = 0 And r.BMatches(1).OriginalText = "" And r.BMatches(1).MatchState = VATS2_BROW_NO_REFERENCES
    Check "未完成B异常文本和类型不产生风险或Issue", r.BMatches(3).ParserFlags = 0 And r.BSourceIssueCount = 0 And r.BMatches(6).ReferenceCount = 0
    Check "CROSS关联保留原B2和B5", r.Conflicts.ConflictGroupCount = 1 And r.Conflicts.ConflictGroups(1).Associations(1).BIndex = 2 And r.Conflicts.ConflictGroups(1).Associations(2).BIndex = 5
    Check "占位B不产生关联且BRowId为ExcelRow", r.Conflicts.UniqueAssociationCount = 2 And r.Conflicts.ConflictGroups(1).Associations(2).BRowId = b.Records(5).ExcelRow
    digits(1) = "00123456": digits(2) = "111123456": digits(3) = "222123456"
    map(1) = 1: map(2) = 3: map(3) = 5
    For i = 1 To b.RecordCount
        If r.BInScope(i) Then
            raw = VATStage2MatchBRow(b.Records(i).SupplierTextRaw, digits, 3)
            expected = raw
            RemapExpected expected, map
            Check "B" & i & "除AIndex外逐字段等于冻结BRowMatch", RowKey(r.BMatches(i)) = RowKey(expected)
        End If
    Next i
    Check "输入Snapshot所有字段保持不变", before = InputKey(a, b)
    again = VATStage2MatchCompletedScope(a, b)
    Check "重复调用完整结果确定", ResultKey(r) = ResultKey(again)
End Sub

Private Sub TestBoundaries()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, r As VATS2CompletedScopeResult
    Init a, b, 3, 1
    a.Records(1).InvoiceDigitsRaw = "12345678"
    a.Records(2).InvoiceDigitsRaw = "11112345678"
    a.Records(3).InvoiceDigitsRaw = "22212345678"
    b.Records(1).SupplierTextRaw = "NO.12345678"
    r = VATStage2MatchCompletedScope(a, b)
    Check "exact优先不混入suffix候选", r.BMatches(1).References(1).MatchKind = VATS2_EXACT_UNIQUE And r.BMatches(1).References(1).CandidateCount = 1 And r.BMatches(1).References(1).Candidates(1).AIndex = 1
    a.Records(1).IsCompleted = False: a.Records(2).IsCompleted = False
    r = VATStage2MatchCompletedScope(a, b)
    Check "suffix唯一回映A3且长度大于6无短号flag", r.BMatches(1).References(1).MatchKind = VATS2_SUFFIX_UNIQUE And r.BMatches(1).References(1).Candidates(1).AIndex = 3 And r.BMatches(1).MatcherFlags = 0
    a.Records(3).IsCompleted = False
    r = VATStage2MatchCompletedScope(a, b)
    Check "无完成A仍完成B解析并返回未找到", r.Status = VATS2_SCOPE_OK And r.CompletedACount = 0 And r.BMatches(1).NotFoundCount = 1
    b.Records(1).SupplierTextRaw = "NO.123"
    r = VATStage2MatchCompletedScope(a, b)
    Check "空scope未执行suffix故无短号flag", r.BMatches(1).MatcherFlags = 0
    b.Records(1).IsCompleted = False
    r = VATStage2MatchCompletedScope(a, b)
    Check "无完成B保留占位和零关联", r.CompletedBCount = 0 And r.Conflicts.UniqueAssociationCount = 0 And Not r.BInScope(1)
    Init a, b, 0, 0
    r = VATStage2MatchCompletedScope(a, b)
    Check "两个空Snapshot正常空scope", r.Status = VATS2_SCOPE_OK And r.CompletedACount = 0 And r.FullBRecordCount = 0 And r.Conflicts.Status = VATS2_BCONFLICT_OK
    Init a, b, 1, 0
    r = VATStage2MatchCompletedScope(a, b)
    Check "只有完成A不产生A_ONLY判断", r.Status = VATS2_SCOPE_OK And r.CompletedACount = 1 And r.Conflicts.UniqueAssociationCount = 0
    Init a, b, 3, 1
    a.Records(1).InvoiceDigitsRaw = "00123": a.Records(2).IsCompleted = False
    a.Records(3).InvoiceDigitsRaw = "00123": b.Records(1).SupplierTextRaw = "NO.00123"
    r = VATStage2MatchCompletedScope(a, b)
    Check "完成A重复exact全部保留不消歧", r.BMatches(1).References(1).MatchKind = VATS2_EXACT_MULTIPLE And r.BMatches(1).References(1).Candidates(2).AIndex = 3 And r.Conflicts.UniqueAssociationCount = 0
    Check "短exact multiple无SHORT_SUFFIX", r.BMatches(1).MatcherFlags = 0
    a.Records(3).IsCompleted = False
    r = VATStage2MatchCompletedScope(a, b)
    Check "短exact unique无SHORT_SUFFIX", r.BMatches(1).References(1).MatchKind = VATS2_EXACT_UNIQUE And r.BMatches(1).MatcherFlags = 0
    b.Records(1).SupplierTextRaw = "公司123--55"
    r = VATStage2MatchCompletedScope(a, b)
    Check "无NO与数字尾注ParserFlags保留", r.BMatches(1).ParserFlags = (VATS2_NO_MARKER_MISSING Or VATS2_NUMERIC_TRAILER_REVIEW)
    Check "前导零suffix及短尾号风险保留", r.BMatches(1).References(1).Candidates(1).InvoiceDigits = "00123" And r.BMatches(1).MatcherFlags = VATS2_SHORT_SUFFIX_REVIEW
    b.Records(1).SupplierTextRaw = "NO.00123/00123"
    r = VATStage2MatchCompletedScope(a, b)
    Check "Parser重复风险不制造第二个关联", r.BMatches(1).ParserFlags = VATS2_DUPLICATE_REF_REVIEW And r.Conflicts.UniqueAssociationCount = 1 And r.Conflicts.ConflictGroupCount = 0
    a.Records(1).InvoiceDigitsRaw = "00123 "
    b.Records(1).SupplierTextRaw = "NO.00123"
    r = VATStage2MatchCompletedScope(a, b)
    Check "String号码原样传递不Trim", r.BMatches(1).NotFoundCount = 1 And r.AIdentityIssueCount = 0
End Sub

Private Sub TestIssues()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, r As VATS2CompletedScopeResult
    Dim types As Variant, value As Variant, i As Long, before As String
    types = Array(Empty, Null, True, 12345678#, CVErr(2042), DateSerial(2026, 9, 12))
    Init a, b, 7, 7
    For i = 1 To 6
        a.Records(i).InvoiceDigitsRaw = types(i - 1)
        b.Records(i).SupplierTextRaw = types(i - 1)
    Next i
    a.Records(7).InvoiceDigitsRaw = "55555555"
    b.Records(7).SupplierTextRaw = "NO.12345678/55555555"
    before = InputKey(a, b)
    r = VATStage2MatchCompletedScope(a, b)
    Check "非String为数据Issue而非接口失败", r.Status = VATS2_SCOPE_OK And r.AIdentityIssueCount = 6 And r.BSourceIssueCount = 6
    For i = 1 To 6
        Check "源类型" & i & "同时保留A/B索引行号VarType", r.AIdentityIssues(i).OriginalAIndex = i And r.AIdentityIssues(i).ExcelRow = a.Records(i).ExcelRow And r.AIdentityIssues(i).RawVarType = VarType(types(i - 1)) And r.BSourceIssues(i).OriginalBIndex = i And r.BSourceIssues(i).ExcelRow = b.Records(i).ExcelRow And r.BSourceIssues(i).RawVarType = VarType(types(i - 1))
        If r.BMatches(i).MatchState <> VATS2_BROW_NO_REFERENCES Or Not r.BInScope(i) Then Err.Raise 5, , "非String B占位错误"
    Next i
    Check "非String B仍在scope且是带Issue的NO_REFERENCES", True
    Check "非String A仍在scope不转数值制造匹配", r.CompletedACount = 7 And r.BMatches(7).References(1).MatchKind = VATS2_NOT_FOUND And r.BMatches(7).References(2).Candidates(1).AIndex = 7
    Check "Issue路径输入原始类型和所有值不变", before = InputKey(a, b)
    a.Records(4).IsCompleted = False: b.Records(4).IsCompleted = False
    r = VATStage2MatchCompletedScope(a, b)
    Check "未完成非String不进入Issue列表", r.AIdentityIssueCount = 5 And r.BSourceIssueCount = 5 And r.CompletedAIndexes(4) = 5
    Init a, b, 1, 1
    a.Records(1).InvoiceDigitsRaw = "": b.Records(1).SupplierTextRaw = ""
    r = VATStage2MatchCompletedScope(a, b)
    Check "String空串不误标非StringIssue", r.AIdentityIssueCount = 0 And r.BSourceIssueCount = 0 And r.BMatches(1).MatchState = VATS2_BROW_NO_REFERENCES
End Sub

Private Sub TestConflicts()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, r As VATS2CompletedScopeResult, i As Long
    Init a, b, 5, 8
    For i = 1 To 5: a.Records(i).IsCompleted = (i = 5): Next i
    a.Records(5).InvoiceDigitsRaw = "00123456"
    For i = 1 To 8: b.Records(i).IsCompleted = (i = 5 Or i = 8): Next i
    b.Records(2).SupplierTextRaw = "NO.00123456"
    b.Records(5).SupplierTextRaw = "NO.00123456/123456"
    b.Records(8).SupplierTextRaw = "NO.00123456"
    r = VATStage2MatchCompletedScope(a, b)
    Check "SAME与CROSS同组共存", r.Conflicts.ConflictGroupCount = 1 And r.Conflicts.ConflictGroups(1).Flags = (VATS2_SAME_B_ROW_REUSE Or VATS2_CROSS_B_ROW_REUSE)
    Check "冲突组AIndex是原A5", r.Conflicts.ConflictGroups(1).AIndex = 5 And r.Conflicts.ConflictGroups(1).AssociationCount = 3
    Check "SAME关联保持原B5两引用", r.Conflicts.ConflictGroups(1).Associations(1).BIndex = 5 And r.Conflicts.ConflictGroups(1).Associations(2).BIndex = 5 And r.Conflicts.ConflictGroups(1).Associations(2).ReferenceIndex = 2
    Check "CROSS关联保持原B8与ExcelRow", r.Conflicts.ConflictGroups(1).Associations(3).BIndex = 8 And r.Conflicts.ConflictGroups(1).Associations(3).BRowId = b.Records(8).ExcelRow And r.Conflicts.ConflictGroups(1).Associations(3).AIndex = 5
    Check "未完成B2不参与冲突计数", r.Conflicts.UniqueAssociationCount = 3 And r.Conflicts.ConflictGroups(1).DistinctBRowCount = 2
End Sub

Private Sub TestContracts()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, r As VATS2CompletedScopeResult, mode As Long
    For mode = 1 To 14
        Init a, b, 2, 2
        Select Case mode
            Case 1: a.Status = VATS2_SNAPSHOT_READ_ERROR
            Case 2: b.Status = VATS2_SNAPSHOT_INVALID_SHEET
            Case 3: a.RecordCount = -1
            Case 4: Erase a.Records
            Case 5: b.RecordCount = 3
            Case 6: ReDim a.Records(0 To 1)
            Case 7: b.Records(2).BIndex = 1
            Case 8: a.Records(2).AIndex = 5
            Case 9: a.Records(2).ExcelRow = a.Records(1).ExcelRow
            Case 10: b.Records(2).ExcelRow = b.HeaderRow
            Case 11: b.HeaderRow = 0
            Case 12: ReDim b.Records(1 To 3)
            Case 13: a.RecordCount = 0
            Case 14: b.RecordCount = 0
        End Select
        r = VATStage2MatchCompletedScope(a, b)
        If mode <= 2 Then
            Check "Snapshot非法状态拒绝" & mode, r.Status = VATS2_SCOPE_INVALID_INPUT
        Else
            Check "Snapshot数组索引行序契约拒绝" & mode, r.Status = VATS2_SCOPE_INVALID_CONTRACT And r.CompletedBCount = 0
        End If
    Next mode
End Sub

Private Sub TestRemap()
    Dim digits(1 To 2) As String, map() As Long, source As VATS2BRowMatchResult
    Dim output As VATS2BRowMatchResult, original As VATS2BRowMatchResult, status As VATS2ScopeStatus, mode As Long, before As String
    digits(1) = "11112345678": digits(2) = "22212345678"
    original = VATStage2MatchBRow("NO.12345678", digits, 2)
    For mode = 1 To 8
        ReDim map(1 To 2): map(1) = 1: map(2) = 5
        source = original
        Select Case mode
            Case 1: source.References(1).Candidates(2).AIndex = 3
            Case 2: source.References(1).Candidates(1).AIndex = 0
            Case 3: source.References(1).Candidates(1).AIndex = -1
            Case 4: Erase source.References(1).Candidates
            Case 5: Erase source.References
            Case 6: Erase map
            Case 7: map(2) = 1
            Case 8: ReDim map(0 To 1)
        End Select
        output = original
        status = VATStage2RemapScopeRow(source, map, 2, output)
        Check "回映损坏或越界拒绝且无部分结果" & mode, status = VATS2_SCOPE_INVALID_CONTRACT And output.ReferenceCount = 0
    Next mode
    ReDim map(1 To 2): map(1) = 1: map(2) = 5
    source = original: before = RowKey(source)
    status = VATStage2RemapScopeRow(source, map, 2, output)
    Check "回映仅改副本AIndex且源结果不变", status = VATS2_SCOPE_OK And output.References(1).Candidates(2).AIndex = 5 And RowKey(source) = before
End Sub

Private Sub Init(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, ByVal ac As Long, ByVal bc As Long)
    Dim emptyA As VATS2ASnapshotResult, emptyB As VATS2BSnapshotResult, i As Long
    a = emptyA: b = emptyB
    a.HeaderRow = 3: b.HeaderRow = 5: a.SheetName = "A": b.SheetName = "B"
    a.RecordCount = ac: b.RecordCount = bc
    If ac > 0 Then ReDim a.Records(1 To ac)
    If bc > 0 Then ReDim b.Records(1 To bc)
    For i = 1 To ac
        With a.Records(i)
            .AIndex = i: .ExcelRow = 10 + 2 * i: .IsCompleted = True
            .InvoiceDigitsRaw = "88888888": .AmountRaw = CVErr(2042)
            .InvoiceCellAddress = "D" & .ExcelRow: .AmountCellAddress = "I" & .ExcelRow
        End With
    Next i
    For i = 1 To bc
        With b.Records(i)
            .BIndex = i: .ExcelRow = 20 + 3 * i: .IsCompleted = True
            .SupplierTextRaw = "NO.88888888": .AmountRaw = "不许处理金额"
            .SupplierCellAddress = "B" & .ExcelRow: .AmountCellAddress = "F" & .ExcelRow
        End With
    Next i
End Sub

Private Sub RemapExpected(ByRef row As VATS2BRowMatchResult, ByRef map() As Long)
    Dim i As Long, j As Long
    For i = 1 To row.ReferenceCount
        For j = 1 To row.References(i).CandidateCount
            row.References(i).Candidates(j).AIndex = map(row.References(i).Candidates(j).AIndex)
        Next j
    Next i
End Sub

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

Private Function ResultKey(ByRef r As VATS2CompletedScopeResult) As String
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
    ResultKey = s
End Function

Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        passed = passed + 1: log = log & "PASS " & name & vbCrLf
    Else
        failed = failed + 1: log = log & "FAIL " & name & vbCrLf
    End If
End Sub
