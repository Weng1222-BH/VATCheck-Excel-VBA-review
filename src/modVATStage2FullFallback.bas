Attribute VB_Name = "modVATStage2FullFallback"
Option Explicit

Public Enum VATS2FallbackStatus
    VATS2_FALLBACK_OK = 0
    VATS2_FALLBACK_INVALID_INPUT = 1
    VATS2_FALLBACK_INVALID_CONTRACT = 2
    VATS2_FALLBACK_ENGINE_ERROR = 3
End Enum

Public Enum VATS2FullEvidence
    VATS2_FULL_UNIQUE_COMPLETED = 1
    VATS2_FULL_UNIQUE_UNCOMPLETED = 2
    VATS2_FULL_MULTIPLE = 3
    VATS2_FULL_NOT_FOUND = 4
    VATS2_FULL_NOT_FOUND_WITH_BLOCKERS = 5
    VATS2_FULL_AMBIGUOUS = 6
End Enum

Public Enum VATS2CompletedACoverage
    VATS2_COVERAGE_NONE = 0
    VATS2_COVERAGE_UNIQUE = 1
    VATS2_COVERAGE_MULTIPLE_ONLY = 2
End Enum

Public Enum VATS2FallbackBlockerReason
    VATS2_BLOCK_NON_STRING = 1
    VATS2_BLOCK_STRUCTURE = 2
    VATS2_BLOCK_NUMERIC_TRAILER = 4
    VATS2_BLOCK_NO_REFERENCES = 8
    VATS2_BLOCK_INVALID_A_STRING = 16
End Enum

Public Type VATS2FallbackBlocker
    OriginalIndex As Long
    ExcelRow As Long
    RawVarType As Long
    Reasons As Long
    ParserFlags As Long
End Type

Public Type VATS2FallbackACandidate
    AIndex As Long
    ExcelRow As Long
    InvoiceDigits As String
    IsCompleted As Boolean
End Type

Public Type VATS2BToAFallback
    BIndex As Long
    ExcelRow As Long
    ReferenceIndex As Long
    OriginalBMatch As VATS2BRowMatchResult
    Evidence As VATS2FullEvidence
    FullAMatch As VATS2ReferenceMatchResult
    Candidates() As VATS2FallbackACandidate
    MissingEvidence As Boolean
End Type

Public Type VATS2FallbackBHit
    BIndex As Long
    ExcelRow As Long
    IsCompleted As Boolean
    UniqueReferenceCount As Long
    MultipleReferenceCount As Long
End Type

Public Type VATS2AToBFallback
    AIndex As Long
    ExcelRow As Long
    Evidence As VATS2FullEvidence
    HitCount As Long
    Hits() As VATS2FallbackBHit
    MissingEvidence As Boolean
    SourceBlocked As Boolean
End Type

Public Type VATS2FullFallbackResult
    Status As VATS2FallbackStatus
    BToAFallbackCount As Long
    BToAFallbacks() As VATS2BToAFallback
    AToBFallbackCount As Long
    AToBFallbacks() As VATS2AToBFallback
    AColorMissingCount As Long
    BColorMissingCount As Long
    CompletedACoverage() As VATS2CompletedACoverage
    FullBScanned As Boolean
    FullBMatches() As VATS2BRowMatchResult
    ABlockerCount As Long
    ABlockers() As VATS2FallbackBlocker
    BBlockerCount As Long
    BBlockers() As VATS2FallbackBlocker
    ErrorSide As String
    ErrorIndex As Long
    ErrorReferenceIndex As Long
    ErrorReason As String
End Type

Public Function VATStage2RunFullFallback(ByRef a As VATS2ASnapshotResult, _
    ByRef b As VATS2BSnapshotResult, ByRef completed As VATS2CompletedScopeResult) As VATS2FullFallbackResult
    Dim r As VATS2FullFallbackResult, blank As VATS2FullFallbackResult
    Dim digits() As String, blockedA() As Boolean, aMissing() As Boolean, bMissing() As Boolean
    Dim i As Long, j As Long, k As Long, ai As Long, n As Long, p As Long
    Dim reason As Long, raw As String, uniqueCount As Long, multipleCount As Long, anyUnique As Boolean
    Dim matched As VATS2ReferenceMatchResult
    r.Status = VATS2_FALLBACK_INVALID_INPUT
    If a.Status <> VATS2_SNAPSHOT_OK Or b.Status <> VATS2_SNAPSHOT_OK Or completed.Status <> VATS2_SCOPE_OK Then
        r.ErrorReason = "Snapshot 或 completed scope 非 OK。": GoTo Done
    End If
    r.Status = VATS2_FALLBACK_INVALID_CONTRACT
    If Not ValidFallbackInputs(a, b, completed) Then
        r.ErrorReason = "快照、完成范围数组或原索引契约不一致。": GoTo Done
    End If
    On Error GoTo Failed
    r.Status = VATS2_FALLBACK_OK
    If a.RecordCount > 0 Then
        ReDim digits(1 To a.RecordCount)
        ReDim blockedA(1 To a.RecordCount)
        ReDim aMissing(1 To a.RecordCount)
        ReDim r.CompletedACoverage(1 To a.RecordCount)
        ReDim r.ABlockers(1 To a.RecordCount)
    End If
    For i = 1 To a.RecordCount
        reason = 0
        If VarType(a.Records(i).InvoiceDigitsRaw) = vbString Then
            digits(i) = a.Records(i).InvoiceDigitsRaw
            '只标记搜索遮蔽，不清洗或修复原字符串，不改变 Matcher 输入。
            If Not ASCIIDigits(digits(i)) Then reason = VATS2_BLOCK_INVALID_A_STRING
        Else
            reason = VATS2_BLOCK_NON_STRING
        End If
        blockedA(i) = (reason <> 0)
        If blockedA(i) Then
            r.ABlockerCount = r.ABlockerCount + 1
            With r.ABlockers(r.ABlockerCount)
                .OriginalIndex = i: .ExcelRow = a.Records(i).ExcelRow
                .RawVarType = VarType(a.Records(i).InvoiceDigitsRaw): .Reasons = reason
            End With
        End If
    Next i
    If r.ABlockerCount > 0 Then
        ReDim Preserve r.ABlockers(1 To r.ABlockerCount)
    Else
        Erase r.ABlockers
    End If
    '先计算第一层覆盖关系，UNIQUE 优先于仅 MULTIPLE；不以 full 结果改写它。
    For i = 1 To b.RecordCount
        If completed.BInScope(i) Then
            For j = 1 To completed.BMatches(i).ReferenceCount
                With completed.BMatches(i).References(j)
                    If .MatchKind = VATS2_NOT_FOUND Then r.BToAFallbackCount = r.BToAFallbackCount + 1
                    For k = 1 To .CandidateCount
                        ai = .Candidates(k).AIndex
                        If .MatchKind = VATS2_EXACT_UNIQUE Or .MatchKind = VATS2_SUFFIX_UNIQUE Then
                            r.CompletedACoverage(ai) = VATS2_COVERAGE_UNIQUE
                        ElseIf r.CompletedACoverage(ai) = VATS2_COVERAGE_NONE Then
                            r.CompletedACoverage(ai) = VATS2_COVERAGE_MULTIPLE_ONLY
                        End If
                    Next k
                End With
            Next j
        End If
    Next i
    If r.BToAFallbackCount > 0 Then ReDim r.BToAFallbacks(1 To r.BToAFallbackCount)
    For i = 1 To b.RecordCount
        If completed.BInScope(i) Then
            For j = 1 To completed.BMatches(i).ReferenceCount
                If completed.BMatches(i).References(j).MatchKind = VATS2_NOT_FOUND Then
                    n = n + 1
                    matched = VATStage2MatchReference(completed.BMatches(i).References(j).Digits, digits, a.RecordCount)
                    With r.BToAFallbacks(n)
                        .BIndex = i: .ExcelRow = b.Records(i).ExcelRow: .ReferenceIndex = j
                        .OriginalBMatch = completed.BMatches(i)
                        .FullAMatch = matched
                        If matched.CandidateCount > 0 Then ReDim .Candidates(1 To matched.CandidateCount)
                        For k = 1 To matched.CandidateCount
                            ai = matched.Candidates(k).AIndex
                            .Candidates(k).AIndex = ai: .Candidates(k).ExcelRow = a.Records(ai).ExcelRow
                            .Candidates(k).InvoiceDigits = matched.Candidates(k).InvoiceDigits
                            .Candidates(k).IsCompleted = a.Records(ai).IsCompleted
                        Next k
                        Select Case matched.MatchKind
                            Case VATS2_EXACT_UNIQUE, VATS2_SUFFIX_UNIQUE
                                ai = matched.Candidates(1).AIndex
                                If a.Records(ai).IsCompleted Then
                                    .Evidence = VATS2_FULL_UNIQUE_COMPLETED
                                    r.ErrorSide = "B": r.ErrorIndex = i: r.ErrorReferenceIndex = j
                                    r.ErrorReason = "第一层 NOT_FOUND，但 full A 唯一候选已完成；输入关系契约矛盾。"
                                    GoTo Contradiction
                                End If
                                .Evidence = VATS2_FULL_UNIQUE_UNCOMPLETED
                                If Not aMissing(ai) Then r.AColorMissingCount = r.AColorMissingCount + 1
                                aMissing(ai) = True
                            Case VATS2_EXACT_MULTIPLE, VATS2_SUFFIX_MULTIPLE
                                .Evidence = VATS2_FULL_MULTIPLE
                            Case VATS2_NOT_FOUND
                                .MissingEvidence = (r.ABlockerCount = 0)
                                If .MissingEvidence Then .Evidence = VATS2_FULL_NOT_FOUND Else .Evidence = VATS2_FULL_NOT_FOUND_WITH_BLOCKERS
                            Case Else
                                Err.Raise 5, , "冻结 Matcher 返回无效状态。"
                        End Select
                    End With
                End If
            Next j
        End If
    Next i
    For i = 1 To a.RecordCount
        If a.Records(i).IsCompleted And r.CompletedACoverage(i) = VATS2_COVERAGE_NONE Then r.AToBFallbackCount = r.AToBFallbackCount + 1
    Next i
    '没有反向目标时不重跑完整 B；FullBScanned 明确区分未扫描与无 blocker。
    If r.AToBFallbackCount = 0 Then GoTo Done
    ReDim r.AToBFallbacks(1 To r.AToBFallbackCount)
    If b.RecordCount > 0 Then
        ReDim r.FullBMatches(1 To b.RecordCount)
        ReDim r.BBlockers(1 To b.RecordCount)
        ReDim bMissing(1 To b.RecordCount)
    End If
    For i = 1 To b.RecordCount
        reason = 0: raw = vbNullString
        If VarType(b.Records(i).SupplierTextRaw) = vbString Then
            raw = b.Records(i).SupplierTextRaw
        Else
            reason = VATS2_BLOCK_NON_STRING
        End If
        r.FullBMatches(i) = VATStage2MatchBRow(raw, digits, a.RecordCount)
        With r.FullBMatches(i)
            If .MatchState = VATS2_BROW_INVALID_INPUT Then Err.Raise 5, , "冻结 full BRowMatch 返回无效状态。"
            If (.ParserFlags And VATS2_STRUCTURE_REVIEW) <> 0 Then reason = reason Or VATS2_BLOCK_STRUCTURE
            If (.ParserFlags And VATS2_NUMERIC_TRAILER_REVIEW) <> 0 Then reason = reason Or VATS2_BLOCK_NUMERIC_TRAILER
            If .ReferenceCount = 0 Then reason = reason Or VATS2_BLOCK_NO_REFERENCES
        End With
        If reason <> 0 Then
            r.BBlockerCount = r.BBlockerCount + 1
            With r.BBlockers(r.BBlockerCount)
                .OriginalIndex = i: .ExcelRow = b.Records(i).ExcelRow
                .RawVarType = VarType(b.Records(i).SupplierTextRaw): .Reasons = reason
                .ParserFlags = r.FullBMatches(i).ParserFlags
            End With
        End If
    Next i
    r.FullBScanned = True
    If r.BBlockerCount > 0 Then ReDim Preserve r.BBlockers(1 To r.BBlockerCount) Else Erase r.BBlockers
    n = 0
    For ai = 1 To a.RecordCount
        If a.Records(ai).IsCompleted And r.CompletedACoverage(ai) = VATS2_COVERAGE_NONE Then
            n = n + 1: anyUnique = False
            With r.AToBFallbacks(n)
                .AIndex = ai: .ExcelRow = a.Records(ai).ExcelRow: .SourceBlocked = blockedA(ai)
                If b.RecordCount > 0 Then ReDim .Hits(1 To b.RecordCount)
                For i = 1 To b.RecordCount
                    uniqueCount = 0: multipleCount = 0
                    For j = 1 To r.FullBMatches(i).ReferenceCount
                        With r.FullBMatches(i).References(j)
                            For k = 1 To .CandidateCount
                                If .Candidates(k).AIndex = ai Then
                                    If .MatchKind = VATS2_EXACT_UNIQUE Or .MatchKind = VATS2_SUFFIX_UNIQUE Then
                                        uniqueCount = uniqueCount + 1
                                    Else
                                        multipleCount = multipleCount + 1
                                    End If
                                End If
                            Next k
                        End With
                    Next j
                    If uniqueCount + multipleCount > 0 Then
                        .HitCount = .HitCount + 1: p = .HitCount
                        .Hits(p).BIndex = i: .Hits(p).ExcelRow = b.Records(i).ExcelRow
                        .Hits(p).IsCompleted = b.Records(i).IsCompleted
                        .Hits(p).UniqueReferenceCount = uniqueCount: .Hits(p).MultipleReferenceCount = multipleCount
                        If uniqueCount > 0 Then
                            If b.Records(i).IsCompleted Then
                                .Evidence = VATS2_FULL_UNIQUE_COMPLETED
                                r.ErrorSide = "A": r.ErrorIndex = ai
                                r.ErrorReason = "第一层未覆盖 A，但 full B 中已完成行唯一引用它；关系契约矛盾。"
                                GoTo Contradiction
                            End If
                            anyUnique = True
                            '按缺色 B 行去重计数；所有命中仍逐 A 保留，不挑选某一行。
                            If Not bMissing(i) Then r.BColorMissingCount = r.BColorMissingCount + 1
                            bMissing(i) = True
                        End If
                    End If
                Next i
                If .HitCount > 0 Then ReDim Preserve .Hits(1 To .HitCount) Else Erase .Hits
                If anyUnique Then
                    .Evidence = VATS2_FULL_UNIQUE_UNCOMPLETED
                ElseIf .HitCount > 0 Then
                    .Evidence = VATS2_FULL_AMBIGUOUS
                Else
                    .MissingEvidence = (r.BBlockerCount = 0 And Not .SourceBlocked)
                    If .MissingEvidence Then .Evidence = VATS2_FULL_NOT_FOUND Else .Evidence = VATS2_FULL_NOT_FOUND_WITH_BLOCKERS
                End If
            End With
        End If
    Next ai
Done:
    VATStage2RunFullFallback = r
    Exit Function
Contradiction:
    '保留矛盾现场用于诊断，主状态禁止将任何结果作为成功使用。
    r.Status = VATS2_FALLBACK_INVALID_CONTRACT
    GoTo Done
Failed:
    blank.Status = VATS2_FALLBACK_ENGINE_ERROR
    blank.ErrorReason = Err.Description
    r = blank
    Resume Done
End Function

Private Function ASCIIDigits(ByVal text As String) As Boolean
    Dim i As Long, ch As String
    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    ASCIIDigits = True
End Function

Private Function ValidFallbackInputs(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, _
    ByRef c As VATS2CompletedScopeResult) As Boolean
    Dim i As Long, j As Long, k As Long, previous As Long, ca As Long, cb As Long, ai As Long
    On Error GoTo Invalid
    If a.RecordCount < 0 Or b.RecordCount < 0 Or a.HeaderRow < 1 Or b.HeaderRow < 1 Then Exit Function
    If c.FullARecordCount <> a.RecordCount Or c.FullBRecordCount <> b.RecordCount Then Exit Function
    If a.RecordCount > 0 Then
        If LBound(a.Records) <> 1 Or UBound(a.Records) <> a.RecordCount Then Exit Function
    End If
    If b.RecordCount > 0 Then
        If LBound(b.Records) <> 1 Or UBound(b.Records) <> b.RecordCount Then Exit Function
        If LBound(c.BMatches) <> 1 Or UBound(c.BMatches) <> b.RecordCount Then Exit Function
        If LBound(c.BInScope) <> 1 Or UBound(c.BInScope) <> b.RecordCount Then Exit Function
        If LBound(c.BRowIds) <> 1 Or UBound(c.BRowIds) <> b.RecordCount Then Exit Function
    End If
    previous = a.HeaderRow
    For i = 1 To a.RecordCount
        If a.Records(i).AIndex <> i Or a.Records(i).ExcelRow <= previous Then Exit Function
        previous = a.Records(i).ExcelRow
        If a.Records(i).IsCompleted Then
            ca = ca + 1
            If c.CompletedAIndexes(ca) <> i Or c.CompletedAExcelRows(ca) <> previous Then Exit Function
        End If
    Next i
    If ca <> c.CompletedACount Then Exit Function
    If ca > 0 Then
        If LBound(c.CompletedAIndexes) <> 1 Or UBound(c.CompletedAIndexes) <> ca Then Exit Function
        If LBound(c.CompletedAExcelRows) <> 1 Or UBound(c.CompletedAExcelRows) <> ca Then Exit Function
    End If
    previous = b.HeaderRow
    For i = 1 To b.RecordCount
        If b.Records(i).BIndex <> i Or b.Records(i).ExcelRow <= previous Then Exit Function
        previous = b.Records(i).ExcelRow
        If c.BInScope(i) <> b.Records(i).IsCompleted Or c.BRowIds(i) <> previous Then Exit Function
        If Not ValidCompletedRow(c.BMatches(i)) Then Exit Function
        If c.BInScope(i) Then
            cb = cb + 1
            If VarType(b.Records(i).SupplierTextRaw) = vbString Then
                If StrComp(c.BMatches(i).OriginalText, b.Records(i).SupplierTextRaw, vbBinaryCompare) <> 0 Then Exit Function
            ElseIf c.BMatches(i).OriginalText <> "" Then
                Exit Function
            End If
        Else
            If c.BMatches(i).ReferenceCount <> 0 Or c.BMatches(i).OriginalText <> "" Then Exit Function
        End If
        For j = 1 To c.BMatches(i).ReferenceCount
            With c.BMatches(i).References(j)
                For k = 1 To .CandidateCount
                    ai = .Candidates(k).AIndex
                    If ai < 1 Or ai > a.RecordCount Then Exit Function
                    If Not a.Records(ai).IsCompleted Or VarType(a.Records(ai).InvoiceDigitsRaw) <> vbString Then Exit Function
                    If StrComp(.Candidates(k).InvoiceDigits, a.Records(ai).InvoiceDigitsRaw, vbBinaryCompare) <> 0 Then Exit Function
                Next k
            End With
        Next j
    Next i
    If cb <> c.CompletedBCount Then Exit Function
    ValidFallbackInputs = True
Invalid:
End Function

Private Function ValidCompletedRow(ByRef row As VATS2BRowMatchResult) As Boolean
    Dim i As Long, u As Long, m As Long, n As Long
    On Error GoTo Invalid
    If row.ReferenceCount < 0 Then Exit Function
    If row.ReferenceCount > 0 Then
        If LBound(row.References) <> 1 Or UBound(row.References) <> row.ReferenceCount Then Exit Function
    End If
    For i = 1 To row.ReferenceCount
        With row.References(i)
            If Not ASCIIDigits(.Digits) Then Exit Function
            Select Case .MatchKind
                Case VATS2_NOT_FOUND
                    n = n + 1: If .CandidateCount <> 0 Then Exit Function
                Case VATS2_EXACT_UNIQUE, VATS2_SUFFIX_UNIQUE
                    u = u + 1: If .CandidateCount <> 1 Then Exit Function
                Case VATS2_EXACT_MULTIPLE, VATS2_SUFFIX_MULTIPLE
                    m = m + 1: If .CandidateCount < 2 Then Exit Function
                Case Else
                    Exit Function
            End Select
            If .CandidateCount > 0 Then
                If LBound(.Candidates) <> 1 Or UBound(.Candidates) <> .CandidateCount Then Exit Function
            End If
        End With
    Next i
    If row.UniqueMatchCount <> u Or row.NotFoundCount <> n Or row.MultipleMatchCount <> m Or row.InvalidMatchCount <> 0 Then Exit Function
    If row.ReferenceCount = 0 Then
        If row.MatchState <> VATS2_BROW_NO_REFERENCES Then Exit Function
    ElseIf u = row.ReferenceCount Then
        If row.MatchState <> VATS2_BROW_ALL_UNIQUE Then Exit Function
    Else
        If row.MatchState <> VATS2_BROW_INCOMPLETE Then Exit Function
    End If
    ValidCompletedRow = True
Invalid:
End Function
