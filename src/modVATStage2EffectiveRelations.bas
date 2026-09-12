Attribute VB_Name = "modVATStage2EffectiveRelations"
Option Explicit
Option Compare Binary

Public Enum VATS2EffectiveStatus
    VATS2_EFFECTIVE_OK = 0
    VATS2_EFFECTIVE_INVALID_INPUT = 1
    VATS2_EFFECTIVE_INVALID_CONTRACT = 2
    VATS2_EFFECTIVE_ENGINE_ERROR = 3
End Enum

Public Enum VATS2RelationIssue
    VATS2_RELATION_A_INVALID = 1
    VATS2_RELATION_A_DUPLICATE = 2
End Enum

Public Type VATS2RelationAIssue
    BIndex As Long
    ExcelRow As Long
    ReferenceIndex As Long
    AIndex As Long
    AExcelRow As Long
    Flags As Long
End Type

Public Type VATS2EffectiveRelationsResult
    Status As VATS2EffectiveStatus
    FullARecordCount As Long
    FullBRecordCount As Long
    EffectiveBCount As Long
    EffectiveBInScope() As Boolean
    EffectiveBMatches() As VATS2BRowMatchResult
    BRowIds() As Long
    AIntegrity As VATS2AIntegrityResult
    AInvalidValue() As Boolean
    ADuplicateGroupIndex() As Long
    RowIssueFlags() As Long
    RelationIssueCount As Long
    RelationIssues() As VATS2RelationAIssue
    EffectiveConflicts As VATS2BConflictResult
    ErrorSide As String
    ErrorIndex As Long
    ErrorReferenceIndex As Long
    ErrorReason As String
End Type

Public Function VATStage2BuildEffectiveRelations(ByRef a As VATS2ASnapshotResult, _
    ByRef b As VATS2BSnapshotResult, ByRef completed As VATS2CompletedScopeResult, _
    ByRef fallback As VATS2FullFallbackResult) As VATS2EffectiveRelationsResult
    Dim r As VATS2EffectiveRelationsResult, blank As VATS2EffectiveRelationsResult
    Dim digits() As String, patch As Object, key As String, seenA() As Boolean, missingB() As Boolean
    Dim coverage() As VATS2CompletedACoverage, i As Long, j As Long, k As Long, p As Long
    Dim ai As Long, bi As Long, ri As Long, ca As Long, cb As Long, u As Long, m As Long
    Dim expectedPatches As Long, expectedReverse As Long, targetHits As Long, missingCount As Long
    Dim previous As Long, flags As Long, uniqueCapacity As Long, phase As VATS2EffectiveStatus
    r.Status = VATS2_EFFECTIVE_INVALID_INPUT
    If a.Status <> VATS2_SNAPSHOT_OK Or b.Status <> VATS2_SNAPSHOT_OK Or _
       completed.Status <> VATS2_SCOPE_OK Or fallback.Status <> VATS2_FALLBACK_OK Then
        r.ErrorReason = "Snapshot、CompletedScope 或 FullFallback 非 OK。": GoTo Done
    End If
    phase = VATS2_EFFECTIVE_INVALID_CONTRACT
    On Error GoTo Invalid
    Need a.RecordCount >= 0 And b.RecordCount >= 0 And a.HeaderRow >= 1 And b.HeaderRow >= 1, "快照计数或表头非法。"
    Need completed.FullARecordCount = a.RecordCount And completed.FullBRecordCount = b.RecordCount, "完整记录数量不一致。"
    r.FullARecordCount = a.RecordCount: r.FullBRecordCount = b.RecordCount
    If a.RecordCount > 0 Then
        Need LBound(a.Records) = 1 And UBound(a.Records) = a.RecordCount, "A 记录数组契约无效。"
        Need LBound(fallback.CompletedACoverage) = 1 And UBound(fallback.CompletedACoverage) = a.RecordCount, "fallback A覆盖数组缺失。"
        ReDim digits(1 To a.RecordCount): ReDim coverage(1 To a.RecordCount): ReDim seenA(1 To a.RecordCount)
    Else
        Need EmptyARecords(a), "零计数 A 仍有记录。"
    End If
    If b.RecordCount > 0 Then
        Need LBound(b.Records) = 1 And UBound(b.Records) = b.RecordCount, "B 记录数组契约无效。"
        Need LBound(completed.BMatches) = 1 And UBound(completed.BMatches) = b.RecordCount, "completed BMatches数组错误。"
        Need LBound(completed.BInScope) = 1 And UBound(completed.BInScope) = b.RecordCount, "completed BInScope数组错误。"
        Need LBound(completed.BRowIds) = 1 And UBound(completed.BRowIds) = b.RecordCount, "completed BRowIds数组错误。"
        ReDim r.EffectiveBInScope(1 To b.RecordCount): ReDim r.EffectiveBMatches(1 To b.RecordCount)
        ReDim r.BRowIds(1 To b.RecordCount): ReDim r.RowIssueFlags(1 To b.RecordCount)
        ReDim missingB(1 To b.RecordCount)
    Else
        Need EmptyBRecords(b), "零计数 B 仍有记录。"
    End If
    previous = a.HeaderRow
    For i = 1 To a.RecordCount
        r.ErrorSide = "A": r.ErrorIndex = i
        Need a.Records(i).AIndex = i And a.Records(i).ExcelRow > previous, "原 AIndex 或 ExcelRow不一致。"
        previous = a.Records(i).ExcelRow
        If VarType(a.Records(i).InvoiceDigitsRaw) = vbString Then digits(i) = a.Records(i).InvoiceDigitsRaw
        If a.Records(i).IsCompleted Then
            ca = ca + 1
            Need completed.CompletedAIndexes(ca) = i And completed.CompletedAExcelRows(ca) = previous, "completed A映射错误。"
        End If
    Next i
    Need ca = completed.CompletedACount, "完成 A计数不一致。"
    If ca > 0 Then
        Need LBound(completed.CompletedAIndexes) = 1 And UBound(completed.CompletedAIndexes) = ca, "完成 A索引数组错误。"
        Need LBound(completed.CompletedAExcelRows) = 1 And UBound(completed.CompletedAExcelRows) = ca, "完成 A行号数组错误。"
    End If
    previous = b.HeaderRow
    For bi = 1 To b.RecordCount
        r.ErrorSide = "B": r.ErrorIndex = bi: r.ErrorReferenceIndex = 0
        Need b.Records(bi).BIndex = bi And b.Records(bi).ExcelRow > previous, "原 BIndex 或 ExcelRow不一致。"
        previous = b.Records(bi).ExcelRow
        Need completed.BInScope(bi) = b.Records(bi).IsCompleted And completed.BRowIds(bi) = previous, "completed B范围或行号错误。"
        Need ValidRow(completed.BMatches(bi), a, True), "completed行计数、候选或数组契约损坏。"
        r.BRowIds(bi) = previous
        If completed.BInScope(bi) Then
            cb = cb + 1
            Need SourceAgrees(completed.BMatches(bi), b.Records(bi).SupplierTextRaw), "completed B源原文不一致。"
            r.EffectiveBInScope(bi) = True
            r.EffectiveBMatches(bi) = completed.BMatches(bi)
            For ri = 1 To completed.BMatches(bi).ReferenceCount
                With completed.BMatches(bi).References(ri)
                    If .MatchKind = VATS2_NOT_FOUND Then expectedPatches = expectedPatches + 1
                    For k = 1 To .CandidateCount
                        ai = .Candidates(k).AIndex
                        If IsUnique(.MatchKind) Then
                            coverage(ai) = VATS2_COVERAGE_UNIQUE
                        ElseIf coverage(ai) = VATS2_COVERAGE_NONE Then
                            coverage(ai) = VATS2_COVERAGE_MULTIPLE_ONLY
                        End If
                    Next k
                End With
            Next ri
        Else
            Need completed.BMatches(bi).ReferenceCount = 0 And completed.BMatches(bi).OriginalText = "", "未完成 B占位不是零引用。"
        End If
    Next bi
    Need cb = completed.CompletedBCount, "完成 B计数不一致。"
    For ai = 1 To a.RecordCount
        Need coverage(ai) = fallback.CompletedACoverage(ai), "fallback覆盖与第一层关系不一致。"
        If a.Records(ai).IsCompleted And coverage(ai) = VATS2_COVERAGE_NONE Then expectedReverse = expectedReverse + 1
    Next ai
    '用 BIndex+ReferenceIndex 一一索引补充项，禁止重复、缺失或覆盖第一层成功/歧义。
    Need fallback.BToAFallbackCount = expectedPatches, "NOT_FOUND补充项数量缺失或多余。"
    Set patch = CreateObject("Scripting.Dictionary")
    If expectedPatches > 0 Then Need LBound(fallback.BToAFallbacks) = 1 And UBound(fallback.BToAFallbacks) = expectedPatches, "BToA数组错误。"
    For p = 1 To fallback.BToAFallbackCount
        With fallback.BToAFallbacks(p)
            bi = .BIndex: ri = .ReferenceIndex
            r.ErrorSide = "B": r.ErrorIndex = bi: r.ErrorReferenceIndex = ri
            Need bi >= 1 And bi <= b.RecordCount, "fallback BIndex越界。"
            Need completed.BInScope(bi), "fallback指向未完成B。"
            Need ri >= 1 And ri <= completed.BMatches(bi).ReferenceCount, "fallback ReferenceIndex越界。"
            Need completed.BMatches(bi).References(ri).MatchKind = VATS2_NOT_FOUND, "fallback未对应第一层NOT_FOUND。"
            Need .ExcelRow = b.Records(bi).ExcelRow, "fallback B行号不一致。"
            Need SameRow(.OriginalBMatch, completed.BMatches(bi)), "fallback保留的原B关系不一致。"
            key = bi & ":" & ri
            Need Not patch.Exists(key), "同一NOT_FOUND出现重复fallback。"
            patch.Add key, p
            With .FullAMatch
                Need .ReferenceDigits = completed.BMatches(bi).References(ri).Digits And .ReferenceLength = completed.BMatches(bi).References(ri).Length, "FullAMatch引用身份不一致。"
                Need ValidMatcher(.MatchKind, .CandidateCount, .Candidates, a, False), "FullAMatch候选契约损坏或AIndex越界。"
            End With
            '只覆盖 Matcher 部分；Parser字段始终来自第一层副本。
            With r.EffectiveBMatches(bi).References(ri)
                .MatchKind = fallback.BToAFallbacks(p).FullAMatch.MatchKind
                .MatcherFlags = fallback.BToAFallbacks(p).FullAMatch.Flags
                .CandidateCount = fallback.BToAFallbacks(p).FullAMatch.CandidateCount
                Erase .Candidates
                If .CandidateCount > 0 Then ReDim .Candidates(1 To .CandidateCount)
                For k = 1 To .CandidateCount
                    .Candidates(k) = fallback.BToAFallbacks(p).FullAMatch.Candidates(k)
                Next k
            End With
        End With
    Next p
    For bi = 1 To b.RecordCount
        If completed.BInScope(bi) Then
            For ri = 1 To completed.BMatches(bi).ReferenceCount
                If completed.BMatches(bi).References(ri).MatchKind = VATS2_NOT_FOUND Then Need patch.Exists(bi & ":" & ri), "NOT_FOUND缺少唯一fallback。"
            Next ri
            Summarize r.EffectiveBMatches(bi)
        End If
    Next bi
    '反向证据必须和 full B 整行候选一致；按原BIndex加入一次。
    Need fallback.AToBFallbackCount = expectedReverse, "AToB目标数量不一致。"
    Need fallback.FullBScanned = (expectedReverse > 0), "FullBScanned与反向目标不一致。"
    If expectedReverse > 0 Then
        Need LBound(fallback.AToBFallbacks) = 1 And UBound(fallback.AToBFallbacks) = expectedReverse, "AToB数组错误。"
        If b.RecordCount > 0 Then Need LBound(fallback.FullBMatches) = 1 And UBound(fallback.FullBMatches) = b.RecordCount, "FullBMatches缺失或数组错误。"
        For bi = 1 To b.RecordCount
            r.ErrorSide = "B": r.ErrorIndex = bi: r.ErrorReferenceIndex = 0
            Need ValidRow(fallback.FullBMatches(bi), a, False), "FullBMatches内部契约错误。"
            Need SourceAgrees(fallback.FullBMatches(bi), b.Records(bi).SupplierTextRaw), "FullBMatches源行错位。"
        Next bi
    End If
    For p = 1 To fallback.AToBFallbackCount
        With fallback.AToBFallbacks(p)
            ai = .AIndex: r.ErrorSide = "A": r.ErrorIndex = ai: r.ErrorReferenceIndex = 0
            Need ai >= 1 And ai <= a.RecordCount, "AToB AIndex越界。"
            Need Not seenA(ai), "AToB重复A目标。": seenA(ai) = True
            Need a.Records(ai).IsCompleted And coverage(ai) = VATS2_COVERAGE_NONE And .ExcelRow = a.Records(ai).ExcelRow, "AToB未对应未覆盖完成A。"
            Need .HitCount >= 0, "HitCount为负。"
            If .HitCount > 0 Then Need LBound(.Hits) = 1 And UBound(.Hits) = .HitCount, "Hit数组错误。"
            targetHits = 0
            '逐行对照命中计数，不重跑号码匹配，不以证据标签篡改MatchKind。
            For bi = 1 To b.RecordCount
                u = 0: m = 0
                For ri = 1 To fallback.FullBMatches(bi).ReferenceCount
                    With fallback.FullBMatches(bi).References(ri)
                        For k = 1 To .CandidateCount
                            If .Candidates(k).AIndex = ai Then
                                If IsUnique(.MatchKind) Then u = u + 1 Else m = m + 1
                            End If
                        Next k
                    End With
                Next ri
                If u + m > 0 Then
                    targetHits = targetHits + 1
                    Need targetHits <= .HitCount, "反向Hit缺失。"
                    With .Hits(targetHits)
                        Need .BIndex = bi And .ExcelRow = b.Records(bi).ExcelRow And .IsCompleted = b.Records(bi).IsCompleted, "Hit BIndex、行号或完成属性错位。"
                        Need .UniqueReferenceCount = u And .MultipleReferenceCount = m, "Hit计数与FullBMatch不一致。"
                    End With
                    If u > 0 Then
                        Need Not b.Records(bi).IsCompleted, "B_COLOR_MISSING Hit指向已完成B。"
                        Need .Evidence = VATS2_FULL_UNIQUE_UNCOMPLETED, "唯一未完成B命中缺少B_COLOR_MISSING证据。"
                        missingB(bi) = True
                        r.EffectiveBInScope(bi) = True
                        r.EffectiveBMatches(bi) = fallback.FullBMatches(bi)
                    End If
                End If
            Next bi
            Need targetHits = .HitCount, "反向Hit多余或不对应候选。"
        End With
    Next p
    For bi = 1 To b.RecordCount
        If missingB(bi) Then missingCount = missingCount + 1
    Next bi
    Need missingCount = fallback.BColorMissingCount, "缺色B去重计数不一致。"
    '完整 A质量由冻结引擎负责；本层只建立原索引到质量结果的查找摘要。
    phase = VATS2_EFFECTIVE_ENGINE_ERROR
    r.AIntegrity = VATStage2ScanAIntegrity(digits, a.RecordCount)
    Need r.AIntegrity.Status = VATS2_A_SCAN_OK, "冻结AIntegrity扫描失败。"
    If a.RecordCount > 0 Then
        ReDim r.AInvalidValue(1 To a.RecordCount): ReDim r.ADuplicateGroupIndex(1 To a.RecordCount)
    End If
    For i = 1 To r.AIntegrity.InvalidValueCount
        r.AInvalidValue(r.AIntegrity.InvalidAIndexes(i)) = True
    Next i
    For i = 1 To r.AIntegrity.DuplicateGroupCount
        For j = 1 To r.AIntegrity.DuplicateGroups(i).OccurrenceCount
            r.ADuplicateGroupIndex(r.AIntegrity.DuplicateGroups(i).AIndexes(j)) = i
        Next j
    Next i
    phase = VATS2_EFFECTIVE_INVALID_CONTRACT
    For bi = 1 To b.RecordCount
        r.ErrorSide = "B": r.ErrorIndex = bi: r.ErrorReferenceIndex = 0
        Need ValidRow(r.EffectiveBMatches(bi), a, False), "Effective内部计数或候选契约错误。"
        If r.EffectiveBInScope(bi) Then
            r.EffectiveBCount = r.EffectiveBCount + 1
            uniqueCapacity = uniqueCapacity + r.EffectiveBMatches(bi).UniqueMatchCount
        End If
    Next bi
    If uniqueCapacity > 0 Then ReDim r.RelationIssues(1 To uniqueCapacity)
    For bi = 1 To b.RecordCount
        If r.EffectiveBInScope(bi) Then
            For ri = 1 To r.EffectiveBMatches(bi).ReferenceCount
                With r.EffectiveBMatches(bi).References(ri)
                    If IsUnique(.MatchKind) Then
                        ai = .Candidates(1).AIndex: flags = 0
                        If r.AInvalidValue(ai) Then flags = flags Or VATS2_RELATION_A_INVALID
                        If r.ADuplicateGroupIndex(ai) > 0 Then flags = flags Or VATS2_RELATION_A_DUPLICATE
                        If flags <> 0 Then
                            r.RowIssueFlags(bi) = r.RowIssueFlags(bi) Or flags
                            r.RelationIssueCount = r.RelationIssueCount + 1
                            With r.RelationIssues(r.RelationIssueCount)
                                .BIndex = bi: .ExcelRow = b.Records(bi).ExcelRow: .ReferenceIndex = ri
                                .AIndex = ai: .AExcelRow = a.Records(ai).ExcelRow: .Flags = flags
                            End With
                        End If
                    End If
                End With
            Next ri
        End If
    Next bi
    If r.RelationIssueCount > 0 Then ReDim Preserve r.RelationIssues(1 To r.RelationIssueCount) Else Erase r.RelationIssues
    phase = VATS2_EFFECTIVE_ENGINE_ERROR
    r.EffectiveConflicts = VATStage2ScanBConflicts(r.EffectiveBMatches, r.BRowIds, b.RecordCount)
    Need r.EffectiveConflicts.Status = VATS2_BCONFLICT_OK, "冻结BConflict重扫失败。"
    r.Status = VATS2_EFFECTIVE_OK: r.ErrorSide = "": r.ErrorIndex = 0: r.ErrorReferenceIndex = 0
Done:
    VATStage2BuildEffectiveRelations = r
    Exit Function
Invalid:
    '接口失败不发布部分有效关系；错误位置保留，不能伪装成零匹配成功。
    blank.Status = phase: blank.ErrorSide = r.ErrorSide: blank.ErrorIndex = r.ErrorIndex
    blank.ErrorReferenceIndex = r.ErrorReferenceIndex: blank.ErrorReason = Err.Description
    r = blank
    Resume Done
End Function

Private Sub Need(ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then Err.Raise vbObjectError + 5101, "EffectiveRelations", reason
End Sub

Private Function IsUnique(ByVal kind As VATS2MatchKind) As Boolean
    IsUnique = (kind = VATS2_EXACT_UNIQUE Or kind = VATS2_SUFFIX_UNIQUE)
End Function

Private Function ValidMatcher(ByVal kind As VATS2MatchKind, ByVal count As Long, _
    ByRef candidates() As VATS2MatchCandidate, ByRef a As VATS2ASnapshotResult, ByVal completedOnly As Boolean) As Boolean
    Dim i As Long, ai As Long, previous As Long
    On Error GoTo Invalid
    Select Case kind
        Case VATS2_NOT_FOUND
            If count <> 0 Then Exit Function
        Case VATS2_EXACT_UNIQUE, VATS2_SUFFIX_UNIQUE
            If count <> 1 Then Exit Function
        Case VATS2_EXACT_MULTIPLE, VATS2_SUFFIX_MULTIPLE
            If count < 2 Then Exit Function
        Case Else
            Exit Function
    End Select
    If count > 0 Then
        If LBound(candidates) <> 1 Or UBound(candidates) <> count Then Exit Function
    End If
    For i = 1 To count
        ai = candidates(i).AIndex
        If ai <= previous Or ai > a.RecordCount Then Exit Function
        previous = ai
        If completedOnly And Not a.Records(ai).IsCompleted Then Exit Function
        If VarType(a.Records(ai).InvoiceDigitsRaw) <> vbString Then Exit Function
        If StrComp(candidates(i).InvoiceDigits, a.Records(ai).InvoiceDigitsRaw, vbBinaryCompare) <> 0 Then Exit Function
    Next i
    ValidMatcher = True
Invalid:
End Function

Private Function ValidRow(ByRef row As VATS2BRowMatchResult, ByRef a As VATS2ASnapshotResult, ByVal completedOnly As Boolean) As Boolean
    Dim i As Long, summary As VATS2BRowMatchResult
    On Error GoTo Invalid
    If row.ReferenceCount < 0 Then Exit Function
    If row.ReferenceCount > 0 Then
        If LBound(row.References) <> 1 Or UBound(row.References) <> row.ReferenceCount Then Exit Function
    End If
    For i = 1 To row.ReferenceCount
        With row.References(i)
            If .Length <> Len(.Digits) Or .Length = 0 Then Exit Function
            If Not ValidMatcher(.MatchKind, .CandidateCount, .Candidates, a, completedOnly) Then Exit Function
        End With
    Next i
    summary = row: Summarize summary
    ValidRow = (summary.UniqueMatchCount = row.UniqueMatchCount And summary.NotFoundCount = row.NotFoundCount And _
        summary.MultipleMatchCount = row.MultipleMatchCount And summary.InvalidMatchCount = row.InvalidMatchCount And _
        summary.MatcherFlags = row.MatcherFlags And summary.MatchState = row.MatchState)
Invalid:
End Function

Private Sub Summarize(ByRef row As VATS2BRowMatchResult)
    Dim i As Long
    row.UniqueMatchCount = 0: row.NotFoundCount = 0: row.MultipleMatchCount = 0: row.InvalidMatchCount = 0: row.MatcherFlags = 0
    For i = 1 To row.ReferenceCount
        With row.References(i)
            row.MatcherFlags = row.MatcherFlags Or .MatcherFlags
            Select Case .MatchKind
                Case VATS2_EXACT_UNIQUE, VATS2_SUFFIX_UNIQUE: row.UniqueMatchCount = row.UniqueMatchCount + 1
                Case VATS2_EXACT_MULTIPLE, VATS2_SUFFIX_MULTIPLE: row.MultipleMatchCount = row.MultipleMatchCount + 1
                Case VATS2_NOT_FOUND: row.NotFoundCount = row.NotFoundCount + 1
                Case Else: row.InvalidMatchCount = row.InvalidMatchCount + 1
            End Select
        End With
    Next i
    If row.ReferenceCount = 0 Then
        row.MatchState = VATS2_BROW_NO_REFERENCES
    ElseIf row.InvalidMatchCount > 0 Then
        row.MatchState = VATS2_BROW_INVALID_INPUT
    ElseIf row.UniqueMatchCount = row.ReferenceCount Then
        row.MatchState = VATS2_BROW_ALL_UNIQUE
    Else
        row.MatchState = VATS2_BROW_INCOMPLETE
    End If
End Sub

Private Function SourceAgrees(ByRef row As VATS2BRowMatchResult, ByVal raw As Variant) As Boolean
    If VarType(raw) = vbString Then
        SourceAgrees = (StrComp(row.OriginalText, raw, vbBinaryCompare) = 0)
    Else
        SourceAgrees = (row.OriginalText = "" And row.ReferenceCount = 0)
    End If
End Function

Private Function SameRow(ByRef x As VATS2BRowMatchResult, ByRef y As VATS2BRowMatchResult) As Boolean
    Dim i As Long, j As Long
    On Error GoTo Invalid
    If x.OriginalText <> y.OriginalText Or x.ReferenceCount <> y.ReferenceCount Or x.ParserFlags <> y.ParserFlags Then Exit Function
    If x.MatchState <> y.MatchState Or x.MatcherFlags <> y.MatcherFlags Or x.UniqueMatchCount <> y.UniqueMatchCount Or x.NotFoundCount <> y.NotFoundCount Or x.MultipleMatchCount <> y.MultipleMatchCount Or x.InvalidMatchCount <> y.InvalidMatchCount Then Exit Function
    For i = 1 To x.ReferenceCount
        With x.References(i)
            If .Digits <> y.References(i).Digits Or .Length <> y.References(i).Length Or .RawFragment <> y.References(i).RawFragment Or .StartIndex <> y.References(i).StartIndex Or .Context <> y.References(i).Context Or .ParserFlags <> y.References(i).ParserFlags Then Exit Function
            If .MatchKind <> y.References(i).MatchKind Or .MatcherFlags <> y.References(i).MatcherFlags Or .CandidateCount <> y.References(i).CandidateCount Then Exit Function
            For j = 1 To .CandidateCount
                If .Candidates(j).AIndex <> y.References(i).Candidates(j).AIndex Or .Candidates(j).InvoiceDigits <> y.References(i).Candidates(j).InvoiceDigits Then Exit Function
            Next j
        End With
    Next i
    SameRow = True
Invalid:
End Function

Private Function EmptyARecords(ByRef a As VATS2ASnapshotResult) As Boolean
    Dim n As Long
    On Error GoTo Unallocated
    n = LBound(a.Records)
    Exit Function
Unallocated:
    EmptyARecords = (Err.Number = 9)
End Function

Private Function EmptyBRecords(ByRef b As VATS2BSnapshotResult) As Boolean
    Dim n As Long
    On Error GoTo Unallocated
    n = LBound(b.Records)
    Exit Function
Unallocated:
    EmptyBRecords = (Err.Number = 9)
End Function
