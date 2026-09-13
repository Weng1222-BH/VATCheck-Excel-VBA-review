Attribute VB_Name = "modVATStage2AuditFindings"
Option Explicit
Option Compare Binary

Public Enum VATS2AuditStatus
    VATS2_AUDIT_OK = 0
    VATS2_AUDIT_INVALID_INPUT = 1
    VATS2_AUDIT_INVALID_CONTRACT = 2
End Enum

Public Enum VATS2FindingCode
    VATS2_FIND_A_COLOR_MISSING = 1
    VATS2_FIND_B_COLOR_MISSING = 2
    VATS2_FIND_B_ONLY = 3
    VATS2_FIND_A_ONLY = 4
    VATS2_FIND_AMBIGUOUS_MATCH = 5
    VATS2_FIND_AMOUNT_MISMATCH = 6
    VATS2_FIND_AMOUNT_ERROR = 7
    VATS2_FIND_INCOMPLETE_GROUP = 8
    VATS2_FIND_DUPLICATE_A_INVOICE = 9
    VATS2_FIND_A_DATA_REVIEW = 10
    VATS2_FIND_DUPLICATE_REFERENCE = 11
    VATS2_FIND_CROSS_B_ROW_REUSE = 12
    VATS2_FIND_STRUCTURE_REVIEW = 13
    VATS2_FIND_NUMERIC_TRAILER_REVIEW = 14
    VATS2_FIND_NO_MARKER_MISSING = 15
    VATS2_FIND_SHORT_SUFFIX_REVIEW = 16
    VATS2_FIND_SEARCH_BLOCKER = 17
    VATS2_FIND_RELATION_QUALITY_REVIEW = 18
End Enum

Public Enum VATS2FindingSource
    VATS2_FIND_FROM_EFFECTIVE = 1
    VATS2_FIND_FROM_B_TO_A = 2
    VATS2_FIND_FROM_A_TO_B = 4
    VATS2_FIND_FROM_INTEGRITY = 8
    VATS2_FIND_FROM_PARSER = 16
    VATS2_FIND_FROM_CONFLICT = 32
    VATS2_FIND_FROM_AMOUNT = 64
    VATS2_FIND_FROM_SHORT_POLICY = 128
    VATS2_FIND_FROM_BLOCKER = 256
    VATS2_FIND_FROM_RELATION_QUALITY = 512
End Enum

Public Type VATS2FindingLink
    AIndex As Long
    AExcelRow As Long
    BIndex As Long
    BExcelRow As Long
    ReferenceIndex As Long
    ReferenceDigits As String
End Type

Public Type VATS2AuditFinding
    Code As VATS2FindingCode
    Side As String
    AIndex As Long
    AExcelRow As Long
    BIndex As Long
    BExcelRow As Long
    ReferenceIndex As Long
    ReferenceDigits As String
    SourceFlags As Long
    ReasonFlags As Long
    ParserFlags As Long
    MatcherFlags As Long
    MatchKind As VATS2MatchKind
    GroupIndex As Long
    MemberCount As Long
    Members() As VATS2FindingLink
    Difference As Variant
    GroupState As VATS2GroupState
    AmountStatus As VATS2AmountStatus
    AmountErrorSide As String
    AmountErrorIndex As Long
    RawVarType As Long
    Detail As String
End Type

Public Type VATS2AuditFindingsResult
    Status As VATS2AuditStatus
    FindingCount As Long
    Findings() As VATS2AuditFinding
    CodeCounts(1 To 18) As Long
    ShortSuffixWaivedCount As Long
    ErrorSide As String
    ErrorIndex As Long
    ErrorReferenceIndex As Long
    ErrorReason As String
End Type

Public Function VATStage2BuildAuditFindings(ByRef a As VATS2ASnapshotResult, _
    ByRef b As VATS2BSnapshotResult, ByRef fallback As VATS2FullFallbackResult, _
    ByRef effective As VATS2EffectiveRelationsResult, ByRef amounts As VATS2EffectiveAmountResult, _
    ByRef shortPolicy As VATS2ShortSuffixPolicyResult) As VATS2AuditFindingsResult
    Dim r As VATS2AuditFindingsResult, blank As VATS2AuditFindingsResult, f As VATS2AuditFinding
    Dim i As Long, j As Long, k As Long, ai As Long, bi As Long, ri As Long, n As Long
    Dim shortCount As Long, uniqueHits As Long, multipleHits As Long, sourceBlocked As Boolean
    Dim seen As Object, decisions As Object, key As String
    r.Status = VATS2_AUDIT_INVALID_INPUT
    If a.Status <> VATS2_SNAPSHOT_OK Or b.Status <> VATS2_SNAPSHOT_OK Or _
        fallback.Status <> VATS2_FALLBACK_OK Or effective.Status <> VATS2_EFFECTIVE_OK Or _
        amounts.Status <> VATS2_EFFECTIVE_AMOUNT_OK Or shortPolicy.Status <> VATS2_SHORT_POLICY_OK Or _
        effective.AIntegrity.Status <> VATS2_A_SCAN_OK Or effective.EffectiveConflicts.Status <> VATS2_BCONFLICT_OK Then
        r.ErrorReason = "上游结果非 OK。": GoTo Done
    End If
    On Error GoTo Invalid
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbBinaryCompare
    Set decisions = CreateObject("Scripting.Dictionary"): decisions.CompareMode = vbBinaryCompare
    Need a.RecordCount >= 0 And b.RecordCount >= 0, "快照数量无效。"
    Need effective.FullARecordCount = a.RecordCount And amounts.FullARecordCount = a.RecordCount And effective.AIntegrity.RecordCount = a.RecordCount, "A数量不一致。"
    Need effective.FullBRecordCount = b.RecordCount And amounts.FullBRecordCount = b.RecordCount And effective.EffectiveConflicts.BRecordCount = b.RecordCount, "B数量不一致。"
    For i = 1 To a.RecordCount
        Context r, "A", i
        Need a.Records(i).AIndex = i And a.Records(i).ExcelRow > a.HeaderRow, "A原索引/行号错误。"
        If i > 1 Then Need a.Records(i).ExcelRow > a.Records(i - 1).ExcelRow, "A行序错误。"
    Next i
    For i = 1 To b.RecordCount
        Context r, "B", i
        Need b.Records(i).BIndex = i And b.Records(i).ExcelRow > b.HeaderRow, "B原索引/行号错误。"
        If i > 1 Then Need b.Records(i).ExcelRow > b.Records(i - 1).ExcelRow, "B行序错误。"
        Need effective.BRowIds(i) = b.Records(i).ExcelRow And amounts.BRowIds(i) = b.Records(i).ExcelRow, "B来源行号错位。"
        Need effective.EffectiveBInScope(i) = amounts.BInScope(i), "Effective范围不一致。"
        If b.Records(i).IsCompleted Then Need effective.EffectiveBInScope(i), "完成B不在有效范围。"
    Next i
    '全表A质量事实不因当前是否被引用而丢弃。
    Need effective.AIntegrity.InvalidValueCount >= 0 And effective.AIntegrity.DuplicateGroupCount >= 0, "A质量计数无效。"
    For i = 1 To effective.AIntegrity.InvalidValueCount
        ai = effective.AIntegrity.InvalidAIndexes(i): Context r, "A", ai
        Need ai >= 1, "非法A质量记录索引必须大于零。"
        f = NewFinding(a, b, VATS2_FIND_A_DATA_REVIEW, ai, 0, 0, "", VATS2_FIND_FROM_INTEGRITY)
        Emit r, f, seen
    Next i
    For i = 1 To effective.AIntegrity.DuplicateGroupCount
        With effective.AIntegrity.DuplicateGroups(i)
            Need .OccurrenceCount >= 2, "重复A组成员不足。"
            ai = .AIndexes(1): Context r, "A", ai
            f = NewFinding(a, b, VATS2_FIND_DUPLICATE_A_INVOICE, ai, 0, 0, .InvoiceDigits, VATS2_FIND_FROM_INTEGRITY)
            f.GroupIndex = i
            For j = 1 To .OccurrenceCount
                ai = .AIndexes(j)
                Need ai >= 1, "重复组A索引必须大于零。"
                AddMember f, a, b, ai, 0, 0, .InvoiceDigits
                Need VarType(a.Records(ai).InvoiceDigitsRaw) = vbString, "重复A身份类型错位。"
                Need StrComp(a.Records(ai).InvoiceDigitsRaw, .InvoiceDigits, vbBinaryCompare) = 0, "重复A身份不一致。"
            Next j
        End With
        Emit r, f, seen
    Next i
    Need fallback.ABlockerCount >= 0 And fallback.BBlockerCount >= 0, "blocker计数无效。"
    For i = 1 To fallback.ABlockerCount
        BlockerFinding a, b, fallback.ABlockers(i), True, r, seen
    Next i
    For i = 1 To fallback.BBlockerCount
        Need fallback.FullBScanned, "未搜索全B却存在B blocker。"
        BlockerFinding a, b, fallback.BBlockers(i), False, r, seen
    Next i
    '先整理有效B整行证据，金额与身份事项互不覆盖。
    For bi = 1 To b.RecordCount
        Context r, "B", bi
        If effective.EffectiveBInScope(bi) Then
            n = n + 1
            With effective.EffectiveBMatches(bi)
                ValidateRow a, .References, .ReferenceCount
                Need .MatchState <> VATS2_BROW_INVALID_INPUT, "有效关系含接口错误。"
                ParserFindings a, b, bi, .ParserFlags, r, seen
                For ri = 1 To .ReferenceCount
                    Context r, "B", bi, ri
                    If IsMultiple(.References(ri).MatchKind) Then Ambiguous a, b, bi, ri, .References(ri), VATS2_FIND_FROM_EFFECTIVE, r, seen
                    If (.References(ri).MatcherFlags And VATS2_SHORT_SUFFIX_REVIEW) <> 0 Then shortCount = shortCount + 1
                Next ri
                Context r, "B", bi
                If .MatchState <> VATS2_BROW_ALL_UNIQUE Then
                    f = NewFinding(a, b, VATS2_FIND_INCOMPLETE_GROUP, 0, bi, 0, "", VATS2_FIND_FROM_EFFECTIVE)
                    f.ReasonFlags = .MatchState: Emit r, f, seen
                End If
            End With
            Need amounts.RelationQualityFlags(bi) = effective.RowIssueFlags(bi), "关系质量来源不一致。"
            Select Case amounts.RowStates(bi)
                Case VATS2_RELATION_QUALITY_BLOCKED
                    Need amounts.RelationQualityFlags(bi) <> 0, "质量阻断无来源。"
                Case VATS2_EA_COMPARED, VATS2_EA_AMOUNT_ERROR, VATS2_EA_NOT_COMPARABLE
                    With amounts.GroupAmounts(bi)
                        Need .BIndex = bi And .GroupState = amounts.RowStates(bi), "组金额来源状态或索引错位。"
                        If .GroupState = VATS2_GROUP_COMPARED Then
                            Need .AmountEvaluated And .Amount.Status = VATS2_AMOUNT_OK, "未实际比较却发布金额结论。"
                            Need .Amount.AmountState = VATS2_AMOUNT_EQUAL Or .Amount.AmountState = VATS2_AMOUNT_MISMATCH, "比较状态错误。"
                            Need .Amount.ACount = effective.EffectiveBMatches(bi).ReferenceCount And .Amount.ACount > 0, "已比较组成员数错位。"
                            For j = 1 To .Amount.ACount
                                Need IsUnique(effective.EffectiveBMatches(bi).References(j).MatchKind), "比较组引用不唯一。"
                                Need .Amount.AIndexes(j) = effective.EffectiveBMatches(bi).References(j).Candidates(1).AIndex, "金额成员A索引与关系错位。"
                            Next j
                            If .Amount.AmountState = VATS2_AMOUNT_MISMATCH Then
                                f = NewFinding(a, b, VATS2_FIND_AMOUNT_MISMATCH, 0, bi, 0, "", VATS2_FIND_FROM_AMOUNT)
                                f.Difference = .Amount.Difference
                                For j = 1 To .Amount.ACount: AddMember f, a, b, .Amount.AIndexes(j), bi, 0, "": Next j
                                f.GroupState = .GroupState: f.AmountStatus = .Amount.Status: Emit r, f, seen
                            End If
                        ElseIf .GroupState = VATS2_GROUP_AMOUNT_ERROR Then
                            Need .AmountEvaluated And .Amount.Status <> VATS2_AMOUNT_OK, "金额异常缺少原始异常状态。"
                            f = NewFinding(a, b, VATS2_FIND_AMOUNT_ERROR, .Amount.ErrorAIndex, bi, .ErrorReferenceIndex, "", VATS2_FIND_FROM_AMOUNT)
                            f.GroupState = .GroupState: f.AmountStatus = .Amount.Status
                            f.AmountErrorSide = .Amount.ErrorSide: f.AmountErrorIndex = .Amount.ErrorIndex
                            f.Detail = .Amount.ErrorReason: f.ReasonFlags = .Reasons: Emit r, f, seen
                        End If
                    End With
                Case Else: Need False, "有效范围包含无效金额行状态。"
            End Select
        Else
            Need amounts.RowStates(bi) = VATS2_EA_OUT_OF_SCOPE, "范围外金额状态错位。"
            If fallback.FullBScanned Then ParserFindings a, b, bi, fallback.FullBMatches(bi).ParserFlags, r, seen
        End If
    Next bi
    Need n = effective.EffectiveBCount And n = amounts.EffectiveBCount, "有效B计数不一致。"
    'B到A的缺失仅采信明确全表无候选且没有遮蔽的证据。
    Need fallback.BToAFallbackCount >= 0 And fallback.AToBFallbackCount >= 0, "fallback计数无效。"
    For i = 1 To fallback.BToAFallbackCount
        With fallback.BToAFallbacks(i)
            bi = .BIndex: ri = .ReferenceIndex: Context r, "B", bi, ri
            CheckReference a, b, effective, bi, ri
            Need b.Records(bi).IsCompleted And .ExcelRow = b.Records(bi).ExcelRow, "B fallback来源错位。"
            Need .FullAMatch.ReferenceDigits = effective.EffectiveBMatches(bi).References(ri).Digits, "fallback引用身份错位。"
            Need .FullAMatch.CandidateCount = effective.EffectiveBMatches(bi).References(ri).CandidateCount And .FullAMatch.MatchKind = effective.EffectiveBMatches(bi).References(ri).MatchKind, "fallback与有效引用不一致。"
            For j = 1 To .FullAMatch.CandidateCount
                Need .FullAMatch.Candidates(j).AIndex = effective.EffectiveBMatches(bi).References(ri).Candidates(j).AIndex, "fallback候选A索引错位。"
            Next j
            Need .MissingEvidence = (.Evidence = VATS2_FULL_NOT_FOUND), "MissingEvidence与Evidence矛盾。"
            Select Case .Evidence
                Case VATS2_FULL_UNIQUE_UNCOMPLETED
                    Need IsUnique(.FullAMatch.MatchKind) And .FullAMatch.CandidateCount = 1, "缺色证据不是唯一。"
                    ai = .FullAMatch.Candidates(1).AIndex
                    f = NewFinding(a, b, VATS2_FIND_A_COLOR_MISSING, ai, bi, ri, .FullAMatch.ReferenceDigits, VATS2_FIND_FROM_B_TO_A)
                    Need Not a.Records(ai).IsCompleted And .Candidates(1).AIndex = ai And .Candidates(1).ExcelRow = a.Records(ai).ExcelRow, "缺色A来源错位。"
                    Emit r, f, seen
                Case VATS2_FULL_MULTIPLE
                    Need IsMultiple(.FullAMatch.MatchKind), "FULL_MULTIPLE与匹配类型矛盾。"
                    Ambiguous a, b, bi, ri, effective.EffectiveBMatches(bi).References(ri), VATS2_FIND_FROM_B_TO_A, r, seen
                Case VATS2_FULL_NOT_FOUND, VATS2_FULL_NOT_FOUND_WITH_BLOCKERS
                    Need .FullAMatch.MatchKind = VATS2_NOT_FOUND And .FullAMatch.CandidateCount = 0, "无候选证据矛盾。"
                    Need .MissingEvidence = (fallback.ABlockerCount = 0), "B缺失证据与A blocker矛盾。"
                    If .MissingEvidence Then
                        f = NewFinding(a, b, VATS2_FIND_B_ONLY, 0, bi, ri, .FullAMatch.ReferenceDigits, VATS2_FIND_FROM_B_TO_A): Emit r, f, seen
                    End If
                Case Else: Need False, "B fallback证据类型无效。"
            End Select
        End With
    Next i
    For i = 1 To fallback.AToBFallbackCount
        With fallback.AToBFallbacks(i)
            ai = .AIndex: Context r, "A", ai
            Need ai >= 1, "A反向索引必须大于零。"
            f = NewFinding(a, b, VATS2_FIND_A_ONLY, ai, 0, 0, "", VATS2_FIND_FROM_A_TO_B)
            Need fallback.FullBScanned And .ExcelRow = a.Records(ai).ExcelRow And a.Records(ai).IsCompleted, "A反向搜索来源无效。"
            Need .HitCount >= 0 And .MissingEvidence = (.Evidence = VATS2_FULL_NOT_FOUND), "A MissingEvidence矛盾。"
            sourceBlocked = False
            For j = 1 To fallback.ABlockerCount
                If fallback.ABlockers(j).OriginalIndex = ai Then sourceBlocked = True
            Next j
            Need .SourceBlocked = sourceBlocked, "SourceBlocked与原A blocker矛盾。"
            If .HitCount = 0 Then
                Need .Evidence = VATS2_FULL_NOT_FOUND Or .Evidence = VATS2_FULL_NOT_FOUND_WITH_BLOCKERS, "A零命中证据错误。"
                Need .MissingEvidence = (fallback.BBlockerCount = 0 And Not .SourceBlocked), "A缺失证据与blocker矛盾。"
                If .MissingEvidence Then Emit r, f, seen
            Else
                Need Not .MissingEvidence, "有候选不能声明缺失。"
                uniqueHits = 0
                For j = 1 To .HitCount
                    bi = .Hits(j).BIndex: Context r, "B", bi
                    f = NewFinding(a, b, VATS2_FIND_B_COLOR_MISSING, ai, bi, 0, "", VATS2_FIND_FROM_A_TO_B)
                    Need .Hits(j).ExcelRow = b.Records(bi).ExcelRow And .Hits(j).IsCompleted = b.Records(bi).IsCompleted, "反向Hit位置错误。"
                    ValidateRow a, fallback.FullBMatches(bi).References, fallback.FullBMatches(bi).ReferenceCount
                    n = 0: multipleHits = 0
                    For ri = 1 To fallback.FullBMatches(bi).ReferenceCount
                        For k = 1 To fallback.FullBMatches(bi).References(ri).CandidateCount
                            If fallback.FullBMatches(bi).References(ri).Candidates(k).AIndex = ai Then
                                If IsUnique(fallback.FullBMatches(bi).References(ri).MatchKind) Then
                                    n = n + 1: uniqueHits = uniqueHits + 1
                                    Need Not b.Records(bi).IsCompleted, "反向唯一Hit不是缺色B。"
                                    CheckReference a, b, effective, bi, ri
                                    Need effective.EffectiveBMatches(bi).References(ri).Candidates(1).AIndex = ai, "缺色B有效关系错位。"
                                    f = NewFinding(a, b, VATS2_FIND_B_COLOR_MISSING, ai, bi, ri, fallback.FullBMatches(bi).References(ri).Digits, VATS2_FIND_FROM_A_TO_B)
                                    Emit r, f, seen
                                ElseIf IsMultiple(fallback.FullBMatches(bi).References(ri).MatchKind) Then
                                    multipleHits = multipleHits + 1
                                    Ambiguous a, b, bi, ri, fallback.FullBMatches(bi).References(ri), VATS2_FIND_FROM_A_TO_B, r, seen
                                End If
                            End If
                        Next k
                    Next ri
                    Need n = .Hits(j).UniqueReferenceCount And multipleHits = .Hits(j).MultipleReferenceCount And n + multipleHits > 0, "反向Hit与引用候选不一致。"
                Next j
                If uniqueHits > 0 Then Need .Evidence = VATS2_FULL_UNIQUE_UNCOMPLETED, "反向unique证据错误。" Else Need .Evidence = VATS2_FULL_AMBIGUOUS, "反向歧义证据错误。"
            End If
        End With
    Next i
    For i = 1 To effective.RelationIssueCount
        With effective.RelationIssues(i)
            Context r, "B", .BIndex, .ReferenceIndex
            CheckReference a, b, effective, .BIndex, .ReferenceIndex
            f = NewFinding(a, b, VATS2_FIND_RELATION_QUALITY_REVIEW, .AIndex, .BIndex, .ReferenceIndex, effective.EffectiveBMatches(.BIndex).References(.ReferenceIndex).Digits, VATS2_FIND_FROM_RELATION_QUALITY)
            Need .ExcelRow = f.BExcelRow And .AExcelRow = f.AExcelRow And .Flags <> 0, "关联质量位置错位。"
            Need IsUnique(effective.EffectiveBMatches(.BIndex).References(.ReferenceIndex).MatchKind), "质量事项关联不唯一。"
            Need effective.EffectiveBMatches(.BIndex).References(.ReferenceIndex).Candidates(1).AIndex = .AIndex, "质量事项A关联错位。"
            f.ReasonFlags = .Flags: Emit r, f, seen
        End With
    Next i
    ConflictFindings a, b, effective, r, seen
    Need shortPolicy.DecisionCount = shortCount, "short决策数量与实际flag不一致。"
    For i = 1 To shortPolicy.DecisionCount
        With shortPolicy.Decisions(i)
            bi = .BIndex: ri = .ReferenceIndex: Context r, "B", bi, ri
            CheckReference a, b, effective, bi, ri
            key = bi & ":" & ri: Need Not decisions.Exists(key), "short决策重复。": decisions.Add key, True
            Need .ExcelRow = b.Records(bi).ExcelRow And .ReferenceDigits = effective.EffectiveBMatches(bi).References(ri).Digits, "short来源位置或Digits错误。"
            Need (effective.EffectiveBMatches(bi).References(ri).MatcherFlags And VATS2_SHORT_SUFFIX_REVIEW) <> 0, "决策没有原始short flag。"
            ai = 0
            If IsUnique(effective.EffectiveBMatches(bi).References(ri).MatchKind) Then ai = effective.EffectiveBMatches(bi).References(ri).Candidates(1).AIndex
            Need .AIndex = ai, "short AIndex错位。"
            Select Case .Decision
                Case VATS2_WAIVED_THIS_RUN: r.ShortSuffixWaivedCount = r.ShortSuffixWaivedCount + 1
                Case VATS2_REVIEW_REQUIRED
                    f = NewFinding(a, b, VATS2_FIND_SHORT_SUFFIX_REVIEW, ai, bi, ri, .ReferenceDigits, VATS2_FIND_FROM_SHORT_POLICY)
                    f.ReasonFlags = .ReasonFlags: f.MatcherFlags = effective.EffectiveBMatches(bi).References(ri).MatcherFlags: Emit r, f, seen
                Case Else: Need False, "未知short决策。"
            End Select
        End With
    Next i
    Need r.ShortSuffixWaivedCount = shortPolicy.WaivedCount And shortPolicy.ReviewRequiredCount = shortCount - r.ShortSuffixWaivedCount, "short汇总计数不一致。"
    SortFindings r
    r.Status = VATS2_AUDIT_OK: r.ErrorSide = "": r.ErrorIndex = 0: r.ErrorReferenceIndex = 0
Done:
    VATStage2BuildAuditFindings = r: Exit Function
Invalid:
    blank.Status = VATS2_AUDIT_INVALID_CONTRACT
    blank.ErrorSide = r.ErrorSide: blank.ErrorIndex = r.ErrorIndex: blank.ErrorReferenceIndex = r.ErrorReferenceIndex
    blank.ErrorReason = Err.Description: r = blank
    Resume Done
End Function

Private Sub Need(ByVal ok As Boolean, ByVal reason As String)
    If Not ok Then Err.Raise 5, , reason
End Sub

Private Sub Context(ByRef r As VATS2AuditFindingsResult, ByVal side As String, ByVal i As Long, Optional ByVal ri As Long = 0)
    r.ErrorSide = side: r.ErrorIndex = i: r.ErrorReferenceIndex = ri
End Sub

Private Function IsUnique(ByVal kind As VATS2MatchKind) As Boolean
    IsUnique = (kind = VATS2_EXACT_UNIQUE Or kind = VATS2_SUFFIX_UNIQUE)
End Function

Private Function IsMultiple(ByVal kind As VATS2MatchKind) As Boolean
    IsMultiple = (kind = VATS2_EXACT_MULTIPLE Or kind = VATS2_SUFFIX_MULTIPLE)
End Function

Private Function NewFinding(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, _
    ByVal code As VATS2FindingCode, ByVal ai As Long, ByVal bi As Long, ByVal ri As Long, ByVal digits As String, ByVal source As Long) As VATS2AuditFinding
    Dim f As VATS2AuditFinding
    Need ai >= 0 And ai <= a.RecordCount And bi >= 0 And bi <= b.RecordCount And ri >= 0, "Finding引用索引越界。"
    f.Code = code: f.AIndex = ai: f.BIndex = bi: f.ReferenceIndex = ri: f.ReferenceDigits = digits: f.SourceFlags = source
    If ai > 0 Then f.AExcelRow = a.Records(ai).ExcelRow: f.Side = "A"
    If bi > 0 Then f.BExcelRow = b.Records(bi).ExcelRow: f.Side = "B"
    NewFinding = f
End Function

Private Sub AddMember(ByRef f As VATS2AuditFinding, ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, _
    ByVal ai As Long, ByVal bi As Long, ByVal ri As Long, ByVal digits As String)
    Dim loc As VATS2AuditFinding
    loc = NewFinding(a, b, f.Code, ai, bi, ri, digits, f.SourceFlags)
    f.MemberCount = f.MemberCount + 1: ReDim Preserve f.Members(1 To f.MemberCount)
    With f.Members(f.MemberCount)
        .AIndex = ai: .AExcelRow = loc.AExcelRow: .BIndex = bi: .BExcelRow = loc.BExcelRow
        .ReferenceIndex = ri: .ReferenceDigits = digits
    End With
End Sub

Private Sub ValidateRow(ByRef a As VATS2ASnapshotResult, ByRef refs() As VATS2BRowReference, ByVal count As Long)
    Dim i As Long, j As Long, ai As Long
    Need count >= 0, "引用计数无效。"
    For i = 1 To count
        With refs(i)
            Need .CandidateCount >= 0, "候选计数无效。"
            If IsUnique(.MatchKind) Then
                Need .CandidateCount = 1, "唯一引用候选数量错误。"
            ElseIf IsMultiple(.MatchKind) Then
                Need .CandidateCount > 1, "歧义引用候选不足。"
            Else
                Need .MatchKind = VATS2_NOT_FOUND And .CandidateCount = 0, "引用含无效匹配状态。"
            End If
            For j = 1 To .CandidateCount
                ai = .Candidates(j).AIndex: Need ai >= 1 And ai <= a.RecordCount, "候选A越界。"
                Need VarType(a.Records(ai).InvoiceDigitsRaw) = vbString, "候选A来源不是String。"
                Need .Candidates(j).InvoiceDigits = a.Records(ai).InvoiceDigitsRaw, "候选A身份错位。"
            Next j
        End With
    Next i
End Sub

Private Sub CheckReference(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, ByRef e As VATS2EffectiveRelationsResult, ByVal bi As Long, ByVal ri As Long)
    Need bi >= 1 And bi <= b.RecordCount, "B引用越界。"
    Need e.EffectiveBInScope(bi), "引用不在有效范围。"
    Need ri >= 1 And ri <= e.EffectiveBMatches(bi).ReferenceCount, "行内引用越界。"
End Sub

Private Sub Ambiguous(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, ByVal bi As Long, ByVal ri As Long, _
    ByRef ref As VATS2BRowReference, ByVal source As Long, ByRef r As VATS2AuditFindingsResult, ByVal seen As Object)
    Dim f As VATS2AuditFinding, j As Long
    f = NewFinding(a, b, VATS2_FIND_AMBIGUOUS_MATCH, 0, bi, ri, ref.Digits, source)
    f.MatchKind = ref.MatchKind: f.ParserFlags = ref.ParserFlags: f.MatcherFlags = ref.MatcherFlags
    For j = 1 To ref.CandidateCount: AddMember f, a, b, ref.Candidates(j).AIndex, bi, ri, ref.Digits: Next j
    Emit r, f, seen
End Sub

Private Sub ParserFindings(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, ByVal bi As Long, ByVal flags As Long, ByRef r As VATS2AuditFindingsResult, ByVal seen As Object)
    Dim bit As Long, code As VATS2FindingCode, f As VATS2AuditFinding
    For bit = 0 To 3
        If (flags And CLng(2 ^ bit)) <> 0 Then
            Select Case bit
                Case 0: code = VATS2_FIND_STRUCTURE_REVIEW
                Case 1: code = VATS2_FIND_NUMERIC_TRAILER_REVIEW
                Case 2: code = VATS2_FIND_NO_MARKER_MISSING
                Case 3: code = VATS2_FIND_DUPLICATE_REFERENCE
            End Select
            f = NewFinding(a, b, code, 0, bi, 0, "", VATS2_FIND_FROM_PARSER)
            f.ParserFlags = flags: f.ReasonFlags = CLng(2 ^ bit): Emit r, f, seen
        End If
    Next bit
End Sub

Private Sub BlockerFinding(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, ByRef blocker As VATS2FallbackBlocker, _
    ByVal isA As Boolean, ByRef r As VATS2AuditFindingsResult, ByVal seen As Object)
    Dim f As VATS2AuditFinding
    If isA Then
        Context r, "A", blocker.OriginalIndex
        f = NewFinding(a, b, VATS2_FIND_SEARCH_BLOCKER, blocker.OriginalIndex, 0, 0, "", VATS2_FIND_FROM_BLOCKER)
        Need f.AIndex > 0 And f.AExcelRow = blocker.ExcelRow, "A blocker来源错误。"
    Else
        Context r, "B", blocker.OriginalIndex
        f = NewFinding(a, b, VATS2_FIND_SEARCH_BLOCKER, 0, blocker.OriginalIndex, 0, "", VATS2_FIND_FROM_BLOCKER)
        Need f.BIndex > 0 And f.BExcelRow = blocker.ExcelRow, "B blocker来源错误。"
    End If
    Need blocker.Reasons <> 0, "blocker缺少原因。"
    f.ReasonFlags = blocker.Reasons: f.ParserFlags = blocker.ParserFlags: f.RawVarType = blocker.RawVarType: Emit r, f, seen
End Sub

Private Sub ConflictFindings(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, ByRef e As VATS2EffectiveRelationsResult, ByRef r As VATS2AuditFindingsResult, ByVal seen As Object)
    Dim i As Long, j As Long, k As Long, bi As Long, f As VATS2AuditFinding
    Need e.EffectiveConflicts.ConflictGroupCount >= 0, "冲突组计数无效。"
    For i = 1 To e.EffectiveConflicts.ConflictGroupCount
        With e.EffectiveConflicts.ConflictGroups(i)
            Need .AssociationCount >= 2, "冲突组关联不足。"
            For j = 1 To .AssociationCount
                Context r, "B", .Associations(j).BIndex, .Associations(j).ReferenceIndex
                CheckReference a, b, e, .Associations(j).BIndex, .Associations(j).ReferenceIndex
                Need .Associations(j).BRowId = b.Records(.Associations(j).BIndex).ExcelRow And .Associations(j).AIndex = .AIndex, "冲突来源位置错误。"
                With .Associations(j)
                    Need IsUnique(e.EffectiveBMatches(.BIndex).References(.ReferenceIndex).MatchKind), "冲突关联非唯一。"
                    Need e.EffectiveBMatches(.BIndex).References(.ReferenceIndex).Candidates(1).AIndex = .AIndex And e.EffectiveBMatches(.BIndex).References(.ReferenceIndex).Digits = .ReferenceDigits, "冲突引用错位。"
                End With
            Next j
            If (.Flags And VATS2_CROSS_B_ROW_REUSE) <> 0 Then
                f = NewFinding(a, b, VATS2_FIND_CROSS_B_ROW_REUSE, .AIndex, .Associations(1).BIndex, 0, .AInvoiceDigits, VATS2_FIND_FROM_CONFLICT)
                f.GroupIndex = i: f.ReasonFlags = .Flags
                For j = 1 To .AssociationCount
                    AddMember f, a, b, .AIndex, .Associations(j).BIndex, .Associations(j).ReferenceIndex, .Associations(j).ReferenceDigits
                Next j
                Emit r, f, seen
            End If
            If (.Flags And VATS2_SAME_B_ROW_REUSE) <> 0 Then
                For j = 1 To .AssociationCount
                    bi = .Associations(j).BIndex
                    If j > 1 Then
                        If .Associations(j - 1).BIndex = bi Then GoTo NextAssociation
                    End If
                    f = NewFinding(a, b, VATS2_FIND_DUPLICATE_REFERENCE, .AIndex, bi, 0, .AInvoiceDigits, VATS2_FIND_FROM_CONFLICT)
                    f.GroupIndex = i: f.ReasonFlags = VATS2_SAME_B_ROW_REUSE
                    For k = 1 To .AssociationCount
                        If .Associations(k).BIndex = bi Then AddMember f, a, b, .AIndex, bi, .Associations(k).ReferenceIndex, .Associations(k).ReferenceDigits
                    Next k
                    If f.MemberCount > 1 Then Emit r, f, seen
NextAssociation:
                Next j
            End If
        End With
    Next i
End Sub

Private Sub Emit(ByRef r As VATS2AuditFindingsResult, ByRef f As VATS2AuditFinding, ByVal seen As Object)
    Dim key As String, i As Long, at As Long
    '键只作查重，不枚举Dictionary。相同歧义候选集合来自两层时合并来源位。
    key = f.Code & ":" & f.Side & ":" & f.AIndex & ":" & f.BIndex & ":" & f.ReferenceIndex & ":" & f.MatchKind & ":" & f.GroupIndex & ":" & f.ReferenceDigits
    For i = 1 To f.MemberCount
        key = key & "|" & f.Members(i).AIndex & ":" & f.Members(i).BIndex & ":" & f.Members(i).ReferenceIndex
    Next i
    If seen.Exists(key) Then
        at = seen(key): r.Findings(at).SourceFlags = r.Findings(at).SourceFlags Or f.SourceFlags
    Else
        r.FindingCount = r.FindingCount + 1: ReDim Preserve r.Findings(1 To r.FindingCount)
        r.Findings(r.FindingCount) = f: seen.Add key, r.FindingCount
        r.CodeCounts(f.Code) = r.CodeCounts(f.Code) + 1
    End If
End Sub

Private Sub SortFindings(ByRef r As VATS2AuditFindingsResult)
    Dim i As Long, j As Long, item As VATS2AuditFinding
    'A独立事项先按AIndex；B事项按BIndex、ReferenceIndex、Code。并列保持证据原序。
    For i = 2 To r.FindingCount
        item = r.Findings(i): j = i - 1
        Do While j >= 1
            If Not ComesAfter(r.Findings(j), item) Then Exit Do
            r.Findings(j + 1) = r.Findings(j): j = j - 1
        Loop
        r.Findings(j + 1) = item
    Next i
End Sub

Private Function ComesAfter(ByRef x As VATS2AuditFinding, ByRef y As VATS2AuditFinding) As Boolean
    If x.BIndex <> y.BIndex Then
        ComesAfter = x.BIndex > y.BIndex
    ElseIf x.BIndex = 0 And x.AIndex <> y.AIndex Then
        ComesAfter = x.AIndex > y.AIndex
    ElseIf x.ReferenceIndex <> y.ReferenceIndex Then
        ComesAfter = x.ReferenceIndex > y.ReferenceIndex
    Else
        ComesAfter = x.Code > y.Code
    End If
End Function
