Attribute VB_Name = "modVATStage2FullFallbackTests"
Option Explicit

Private passed As Long, failed As Long, log As String

Public Function VATStage2FullFallback_SelfTest() As String
    On Error GoTo Unexpected
    passed = 0: failed = 0: log = vbNullString
    TestForward
    TestReverse
    TestBlockers
    TestContracts
    GoTo Done
Unexpected:
    failed = failed + 1: log = log & "UNEXPECTED " & Err.Number & " " & Err.Description & vbCrLf
Done:
    If failed = 0 Then
        VATStage2FullFallback_SelfTest = "PASS: " & passed & " assertions" & vbCrLf & log
    Else
        VATStage2FullFallback_SelfTest = "FAIL: " & failed & "; PASS: " & passed & vbCrLf & log
    End If
End Function

Private Sub TestForward()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, c As VATS2CompletedScopeResult
    Dim r As VATS2FullFallbackResult, before As String
    Init a, b, 3, 2
    a.Records(1).InvoiceDigitsRaw = "00111111": a.Records(1).IsCompleted = True
    a.Records(2).InvoiceDigitsRaw = "00222222"
    a.Records(3).InvoiceDigitsRaw = "00333333": a.Records(3).IsCompleted = True
    b.Records(2).SupplierTextRaw = "NO.00111111/00222222/00333333": b.Records(2).IsCompleted = True
    c = VATStage2MatchCompletedScope(a, b)
    r = Run(a, b, c)
    Check "合并B仅中间NOT_FOUND触发fallback", r.Status = VATS2_FALLBACK_OK And r.BToAFallbackCount = 1 And r.BToAFallbacks(1).ReferenceIndex = 2
    Check "未完成A exact形成缺色证据", r.BToAFallbacks(1).Evidence = VATS2_FULL_UNIQUE_UNCOMPLETED And r.AColorMissingCount = 1 And r.BToAFallbacks(1).FullAMatch.MatchKind = VATS2_EXACT_UNIQUE
    Check "正向原B索引行号及原A索引行号保持", r.BToAFallbacks(1).BIndex = 2 And r.BToAFallbacks(1).ExcelRow = b.Records(2).ExcelRow And r.BToAFallbacks(1).Candidates(1).AIndex = 2 And r.BToAFallbacks(1).Candidates(1).ExcelRow = a.Records(2).ExcelRow
    Check "前导零号码和未完成属性保持", r.BToAFallbacks(1).Candidates(1).InvoiceDigits = "00222222" And Not r.BToAFallbacks(1).Candidates(1).IsCompleted
    Check "保留原合并B整行关系而非局部金额组", RowKey(r.BToAFallbacks(1).OriginalBMatch) = RowKey(c.BMatches(2)) And r.BToAFallbacks(1).OriginalBMatch.ReferenceCount = 3
    Check "已覆盖完成A不反向搜索", r.AToBFallbackCount = 0 And Not r.FullBScanned And r.CompletedACoverage(1) = VATS2_COVERAGE_UNIQUE And r.CompletedACoverage(3) = VATS2_COVERAGE_UNIQUE
    b.Records(2).SupplierTextRaw = "NO.00111111/222222/00333333"
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "未完成A suffix unique及冻结短号风险", r.BToAFallbacks(1).FullAMatch.MatchKind = VATS2_SUFFIX_UNIQUE And r.BToAFallbacks(1).FullAMatch.Flags = VATS2_SHORT_SUFFIX_REVIEW And r.AColorMissingCount = 1
    a.Records(2).InvoiceDigitsRaw = "00999999"
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "完整A无命中且无blocker可作为missing evidence", r.BToAFallbacks(1).Evidence = VATS2_FULL_NOT_FOUND And r.BToAFallbacks(1).MissingEvidence And r.AColorMissingCount = 0
    Init a, b, 4, 1
    a.Records(2).InvoiceDigitsRaw = "00123456": a.Records(4).InvoiceDigitsRaw = "00123456"
    b.Records(1).IsCompleted = True: b.Records(1).SupplierTextRaw = "NO.00123456"
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "full A exact MULTIPLE不选择候选", r.BToAFallbacks(1).Evidence = VATS2_FULL_MULTIPLE And r.BToAFallbacks(1).FullAMatch.MatchKind = VATS2_EXACT_MULTIPLE And r.AColorMissingCount = 0
    Check "full多个候选保持原A2和A4顺序", r.BToAFallbacks(1).Candidates(1).AIndex = 2 And r.BToAFallbacks(1).Candidates(2).AIndex = 4
    a.Records(2).InvoiceDigitsRaw = "11123456": a.Records(4).InvoiceDigitsRaw = "22123456"
    b.Records(1).SupplierTextRaw = "NO.123456"
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "full A suffix MULTIPLE保持歧义", r.BToAFallbacks(1).FullAMatch.MatchKind = VATS2_SUFFIX_MULTIPLE And r.BToAFallbacks(1).Evidence = VATS2_FULL_MULTIPLE
    '第一层suffix unique即使全表另有exact，也绝不能被覆盖。
    Init a, b, 2, 1
    a.Records(1).InvoiceDigitsRaw = "11123456": a.Records(1).IsCompleted = True
    a.Records(2).InvoiceDigitsRaw = "123456"
    b.Records(1).SupplierTextRaw = "NO.123456": b.Records(1).IsCompleted = True
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "全表exact不覆盖已成立的completed suffix", r.BToAFallbackCount = 0 And r.AToBFallbackCount = 0 And c.BMatches(1).References(1).Candidates(1).AIndex = 1 And c.BMatches(1).References(1).MatchKind = VATS2_SUFFIX_UNIQUE
    Init a, b, 3, 1
    a.Records(1).InvoiceDigitsRaw = "11123456": a.Records(1).IsCompleted = True
    a.Records(2).InvoiceDigitsRaw = "123456"
    a.Records(3).InvoiceDigitsRaw = "77777777": a.Records(3).IsCompleted = True
    b.Records(1).SupplierTextRaw = "NO.123456": b.Records(1).IsCompleted = True
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "为另一A执行full B也不反改第一层", r.FullBScanned And r.AToBFallbackCount = 1 And r.AToBFallbacks(1).AIndex = 3 And r.CompletedACoverage(1) = VATS2_COVERAGE_UNIQUE And c.BMatches(1).References(1).Candidates(1).AIndex = 1 And r.FullBMatches(1).References(1).Candidates(1).AIndex = 2
    Init a, b, 1, 2
    b.Records(1).IsCompleted = True: b.Records(1).SupplierTextRaw = "NO.12345678"
    b.Records(2).IsCompleted = True: b.Records(2).SupplierTextRaw = "NO.12345678"
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "多B引用同一未完成A保留全部证据且缺色A按行计数", r.BToAFallbackCount = 2 And r.AColorMissingCount = 1 And r.BToAFallbacks(1).BIndex = 1 And r.BToAFallbacks(2).BIndex = 2
End Sub

Private Sub TestReverse()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, c As VATS2CompletedScopeResult, r As VATS2FullFallbackResult
    Init a, b, 3, 4
    a.Records(2).IsCompleted = True: a.Records(2).InvoiceDigitsRaw = "00123456"
    a.Records(3).InvoiceDigitsRaw = "00987654"
    b.Records(2).SupplierTextRaw = "NO.00123456/00987654/77777777"
    b.Records(4).SupplierTextRaw = "NO.123456"
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "未出现的完成A才执行反向full B", r.FullBScanned And r.AToBFallbackCount = 1 And r.AToBFallbacks(1).AIndex = 2 And r.AToBFallbacks(1).ExcelRow = a.Records(2).ExcelRow
    Check "多个未完成B唯一指向A全保留", r.AToBFallbacks(1).HitCount = 2 And r.BColorMissingCount = 2 And r.AToBFallbacks(1).Evidence = VATS2_FULL_UNIQUE_UNCOMPLETED
    Check "反向保留B2和B4原索引行号", r.AToBFallbacks(1).Hits(1).BIndex = 2 And r.AToBFallbacks(1).Hits(2).BIndex = 4 And r.AToBFallbacks(1).Hits(2).ExcelRow = b.Records(4).ExcelRow
    Check "未完成merged B返回全部三引用及失败引用", r.FullBMatches(2).ReferenceCount = 3 And r.FullBMatches(2).UniqueMatchCount = 2 And r.FullBMatches(2).NotFoundCount = 1
    Check "完整B关系仍使用原AIndex", r.FullBMatches(2).References(1).Candidates(1).AIndex = 2 And r.FullBMatches(2).References(2).Candidates(1).AIndex = 3
    a.Records(3).IsCompleted = True
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "同一缺色合并B被多A命中不重复计算缺色行", r.AToBFallbackCount = 2 And r.BColorMissingCount = 2 And r.AToBFallbacks(2).Hits(1).BIndex = 2
    Init a, b, 2, 1
    a.Records(1).InvoiceDigitsRaw = "11123456": a.Records(1).IsCompleted = True
    a.Records(2).InvoiceDigitsRaw = "22123456": a.Records(2).IsCompleted = True
    b.Records(1).SupplierTextRaw = "NO.123456": b.Records(1).IsCompleted = True
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "仅completed MULTIPLE覆盖不fallback或判缺色", r.BToAFallbackCount = 0 And r.AToBFallbackCount = 0 And r.CompletedACoverage(1) = VATS2_COVERAGE_MULTIPLE_ONLY And r.BColorMissingCount = 0
    b.Records(1).IsCompleted = False
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "full B只出现multiple候选为ambiguous", r.AToBFallbackCount = 2 And r.AToBFallbacks(1).Evidence = VATS2_FULL_AMBIGUOUS And r.BColorMissingCount = 0
    Check "ambiguous保留整行与全部候选", r.AToBFallbacks(1).Hits(1).MultipleReferenceCount = 1 And r.FullBMatches(1).References(1).CandidateCount = 2
    b.Records(1).SupplierTextRaw = "NO.11123456/123456"
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "unique与multiple并存时保留二者不丢歧义", r.AToBFallbacks(1).Evidence = VATS2_FULL_UNIQUE_UNCOMPLETED And r.AToBFallbacks(1).Hits(1).UniqueReferenceCount = 1 And r.AToBFallbacks(1).Hits(1).MultipleReferenceCount = 1 And r.AToBFallbacks(2).Evidence = VATS2_FULL_AMBIGUOUS
    b.Records(1).SupplierTextRaw = "NO.99999999"
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "full B无候选且无blocker形成可靠证据", r.AToBFallbacks(1).Evidence = VATS2_FULL_NOT_FOUND And r.AToBFallbacks(1).MissingEvidence And r.BBlockerCount = 0
    Init a, b, 0, 0
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "双空快照正常零fallback", r.Status = VATS2_FALLBACK_OK And r.BToAFallbackCount = 0 And r.AToBFallbackCount = 0
    Init a, b, 1, 0: a.Records(1).IsCompleted = True
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "空B也完成反向搜索并形成missing evidence", r.FullBScanned And r.AToBFallbacks(1).MissingEvidence
End Sub

Private Sub TestBlockers()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, c As VATS2CompletedScopeResult, r As VATS2FullFallbackResult
    Dim v As Variant, i As Long
    v = Array(12345678#, Empty, Null, True, CVErr(2042))
    For i = 0 To UBound(v)
        Init a, b, 2, 1
        a.Records(1).InvoiceDigitsRaw = "99999999"
        a.Records(2).InvoiceDigitsRaw = v(i)
        b.Records(1).SupplierTextRaw = "NO.12345678": b.Records(1).IsCompleted = True
        c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
        Check "非String A阻断确定缺失类型" & i, r.ABlockerCount = 1 And r.ABlockers(1).OriginalIndex = 2 And r.ABlockers(1).ExcelRow = a.Records(2).ExcelRow And r.ABlockers(1).RawVarType = VarType(v(i)) And r.BToAFallbacks(1).Evidence = VATS2_FULL_NOT_FOUND_WITH_BLOCKERS And Not r.BToAFallbacks(1).MissingEvidence
    Next i
    Init a, b, 1, 2
    a.Records(1).IsCompleted = True
    b.Records(2).SupplierTextRaw = 12345678#
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "非String B记录原索引类型和遮蔽证据", r.BBlockerCount = 1 And r.BBlockers(1).OriginalIndex = 2 And r.BBlockers(1).RawVarType = vbDouble And r.AToBFallbacks(1).Evidence = VATS2_FULL_NOT_FOUND_WITH_BLOCKERS
    b.Records(2).SupplierTextRaw = "NO.99999999//"
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "Parser结构异常产生blocker", (r.BBlockers(1).ParserFlags And VATS2_STRUCTURE_REVIEW) <> 0 And (r.BBlockers(1).Reasons And VATS2_BLOCK_STRUCTURE) <> 0 And Not r.AToBFallbacks(1).MissingEvidence
    b.Records(2).SupplierTextRaw = "NO.99999999--12345678"
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "纯数字尾注遮蔽也不判确定缺失", (r.BBlockers(1).Reasons And VATS2_BLOCK_NUMERIC_TRAILER) <> 0 And r.AToBFallbacks(1).Evidence = VATS2_FULL_NOT_FOUND_WITH_BLOCKERS
    b.Records(2).SupplierTextRaw = ""
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "零引用String保守记录遮蔽", (r.BBlockers(1).Reasons And VATS2_BLOCK_NO_REFERENCES) <> 0 And Not r.AToBFallbacks(1).MissingEvidence
    b.Records(2).SupplierTextRaw = "公司99999999"
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "正常无NO候选不误标结构blocker", r.BBlockerCount = 0 And r.AToBFallbacks(1).MissingEvidence And r.FullBMatches(2).ParserFlags = VATS2_NO_MARKER_MISSING
    a.Records(1).InvoiceDigitsRaw = Empty
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "反向源A自身不可搜索也不判可靠缺失", r.AToBFallbacks(1).SourceBlocked And r.AToBFallbacks(1).Evidence = VATS2_FULL_NOT_FOUND_WITH_BLOCKERS
    a.Records(1).InvoiceDigitsRaw = " 12345678 "
    c = VATStage2MatchCompletedScope(a, b): r = Run(a, b, c)
    Check "非法A String只标遮蔽不Trim修复", r.ABlockers(1).Reasons = VATS2_BLOCK_INVALID_A_STRING And r.AToBFallbacks(1).SourceBlocked
End Sub

Private Sub TestContracts()
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, c As VATS2CompletedScopeResult, r As VATS2FullFallbackResult
    Dim mode As Long
    Init a, b, 1, 1
    a.Records(1).IsCompleted = True: b.Records(1).IsCompleted = True
    b.Records(1).SupplierTextRaw = "NO.99999999"
    c = VATStage2MatchCompletedScope(a, b)
    '同批原引用字段不变，构造过期第一层：当前A变成此前NOT_FOUND号码。
    a.Records(1).InvoiceDigitsRaw = "99999999"
    r = VATStage2RunFullFallback(a, b, c)
    Check "full unique completed明确契约异常", r.Status = VATS2_FALLBACK_INVALID_CONTRACT And r.BToAFallbacks(1).Evidence = VATS2_FULL_UNIQUE_COMPLETED And r.ErrorSide = "B" And r.ErrorReferenceIndex = 1
    For mode = 1 To 8
        Init a, b, 1, 1
        a.Records(1).IsCompleted = True: b.Records(1).IsCompleted = True
        b.Records(1).SupplierTextRaw = "NO.12345678"
        c = VATStage2MatchCompletedScope(a, b)
        Select Case mode
            Case 1: a.Status = VATS2_SNAPSHOT_READ_ERROR
            Case 2: c.Status = VATS2_SCOPE_INVALID_CONTRACT
            Case 3: c.FullARecordCount = 2
            Case 4: c.BInScope(1) = False
            Case 5: c.BMatches(1).References(1).Candidates(1).AIndex = 2
            Case 6: Erase b.Records
            Case 7: c.BRowIds(1) = 999
            Case 8: c.BMatches(1).OriginalText = "另一批数据"
        End Select
        r = VATStage2RunFullFallback(a, b, c)
        If mode <= 2 Then
            Check "非OK输入拒绝" & mode, r.Status = VATS2_FALLBACK_INVALID_INPUT
        Else
            Check "损坏关系映射拒绝" & mode, r.Status = VATS2_FALLBACK_INVALID_CONTRACT
        End If
    Next mode
End Sub

Private Function Run(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, _
    ByRef c As VATS2CompletedScopeResult) As VATS2FullFallbackResult
    Dim r As VATS2FullFallbackResult, again As VATS2FullFallbackResult, inputBefore As String, cBefore As String
    inputBefore = InputKey(a, b): cBefore = CompletedKey(c)
    r = VATStage2RunFullFallback(a, b, c)
    Check "Snapshot输入所有字段不变", InputKey(a, b) = inputBefore
    Check "completed输入完整结果不变", CompletedKey(c) = cBefore
    again = VATStage2RunFullFallback(a, b, c)
    Check "fallback重复调用结果确定", FallbackKey(r, a.RecordCount, b.RecordCount) = FallbackKey(again, a.RecordCount, b.RecordCount)
    Run = r
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
