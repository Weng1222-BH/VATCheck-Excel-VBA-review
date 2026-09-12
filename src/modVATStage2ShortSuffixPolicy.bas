Attribute VB_Name = "modVATStage2ShortSuffixPolicy"
Option Explicit
Option Compare Binary

Public Enum VATS2ShortPolicyStatus
    VATS2_SHORT_POLICY_OK = 0
    VATS2_SHORT_POLICY_INVALID_INPUT = 1
    VATS2_SHORT_POLICY_INVALID_CONTRACT = 2
End Enum

Public Enum VATS2ShortDecisionKind
    VATS2_REVIEW_REQUIRED = 0
    VATS2_WAIVED_THIS_RUN = 1
End Enum

Public Enum VATS2ShortReasonFlags
    VATS2_SHORT_REASON_NONE = 0
    VATS2_AGGREGATE_NOT_ELIGIBLE = 1
    VATS2_ROW_NOT_COMPLETE_UNIQUE = 2
    VATS2_GROUP_NOT_EQUAL = 4
    VATS2_PARSER_RISK = 8
    VATS2_RELATION_QUALITY_RISK = 16
    VATS2_CONFLICT_RISK = 32
    VATS2_COLOR_MISSING = 64
    VATS2_OTHER_STRUCTURE_RISK = 128
End Enum

Public Type VATS2ShortSuffixDecision
    BIndex As Long
    ExcelRow As Long
    ReferenceIndex As Long
    AIndex As Long
    ReferenceDigits As String
    Decision As VATS2ShortDecisionKind
    ReasonFlags As Long
End Type

Public Type VATS2ShortSuffixPolicyResult
    Status As VATS2ShortPolicyStatus
    AggregateReliableEqual As Boolean
    DecisionCount As Long
    Decisions() As VATS2ShortSuffixDecision
    WaivedCount As Long
    ReviewRequiredCount As Long
    ErrorBIndex As Long
    ErrorReferenceIndex As Long
    ErrorReason As String
End Type

Public Function VATStage2EvaluateShortSuffixPolicy(ByRef a As VATS2ASnapshotResult, _
    ByRef b As VATS2BSnapshotResult, ByRef effective As VATS2EffectiveRelationsResult, _
    ByRef amounts As VATS2EffectiveAmountResult, ByVal aggregateReliableEqual As Boolean) As VATS2ShortSuffixPolicyResult
    Dim r As VATS2ShortSuffixPolicyResult, blank As VATS2ShortSuffixPolicyResult
    Dim bi As Long, ri As Long, n As Long, capacity As Long, reasons As Long, hasShort As Boolean
    r.AggregateReliableEqual = aggregateReliableEqual: r.Status = VATS2_SHORT_POLICY_INVALID_INPUT
    If a.Status <> VATS2_SNAPSHOT_OK Or b.Status <> VATS2_SNAPSHOT_OK Or _
        effective.Status <> VATS2_EFFECTIVE_OK Or amounts.Status <> VATS2_EFFECTIVE_AMOUNT_OK Or _
        effective.EffectiveConflicts.Status <> VATS2_BCONFLICT_OK Then
        r.ErrorReason = "Snapshot、EffectiveRelations、EffectiveAmount 或冲突结果非 OK。": GoTo Done
    End If
    On Error GoTo Invalid
    '仅核对策略实际读取的索引空间，不重做上游的完整契约扫描。
    Need a.RecordCount >= 0 And b.RecordCount >= 0, "快照数量无效。"
    Need effective.FullARecordCount = a.RecordCount And amounts.FullARecordCount = a.RecordCount, "A数量不一致。"
    Need effective.FullBRecordCount = b.RecordCount And amounts.FullBRecordCount = b.RecordCount, "B数量不一致。"
    Need effective.EffectiveBCount = amounts.EffectiveBCount, "Effective范围计数不一致。"
    If a.RecordCount > 0 Then Need LBound(a.Records) = 1 And UBound(a.Records) = a.RecordCount, "A记录边界错误。"
    If b.RecordCount > 0 Then
        Need LBound(b.Records) = 1 And UBound(b.Records) = b.RecordCount, "B记录边界错误。"
        Need LBound(effective.EffectiveBInScope) = 1 And UBound(effective.EffectiveBInScope) = b.RecordCount, "Effective范围数组错误。"
        Need LBound(effective.EffectiveBMatches) = 1 And UBound(effective.EffectiveBMatches) = b.RecordCount, "Effective关系数组错误。"
        Need LBound(amounts.BInScope) = 1 And UBound(amounts.BInScope) = b.RecordCount, "金额范围数组错误。"
    End If
    For bi = 1 To b.RecordCount
        r.ErrorBIndex = bi: r.ErrorReferenceIndex = 0
        Need effective.EffectiveBInScope(bi) = amounts.BInScope(bi), "关系和金额范围不一致。"
        With effective.EffectiveBMatches(bi)
            Need .ReferenceCount >= 0, "引用计数错误。"
            If Not effective.EffectiveBInScope(bi) Then
                Need .ReferenceCount = 0, "范围外B不应包含引用。"
            Else
                Need b.Records(bi).BIndex = bi, "BIndex错位。"
                Need effective.BRowIds(bi) = b.Records(bi).ExcelRow And amounts.BRowIds(bi) = b.Records(bi).ExcelRow, "B行号错位。"
                Need b.Records(bi).ExcelRow > b.HeaderRow, "B行号不在表头之后。"
                If .ReferenceCount > 0 Then Need LBound(.References) = 1 And UBound(.References) = .ReferenceCount, "引用数组边界错误。"
                hasShort = False
                For ri = 1 To .ReferenceCount
                    If (.References(ri).MatcherFlags And VATS2_SHORT_SUFFIX_REVIEW) <> 0 Then hasShort = True
                Next ri
                If hasShort Then
                    reasons = RowReasons(a, b, effective, amounts, bi, aggregateReliableEqual)
                    For ri = 1 To .ReferenceCount
                        r.ErrorReferenceIndex = ri
                        With .References(ri)
                            If (.MatcherFlags And VATS2_SHORT_SUFFIX_REVIEW) <> 0 Then
                                Need .Length = Len(.Digits) And .Length >= 1 And .Length <= 6, "short标志与引用长度不一致。"
                                Need .MatchKind = VATS2_SUFFIX_UNIQUE Or .MatchKind = VATS2_SUFFIX_MULTIPLE Or .MatchKind = VATS2_NOT_FOUND, "short标志与匹配类型不一致。"
                                n = n + 1
                                If n > capacity Then
                                    capacity = capacity + 64: ReDim Preserve r.Decisions(1 To capacity)
                                End If
                                r.Decisions(n).BIndex = bi: r.Decisions(n).ExcelRow = b.Records(bi).ExcelRow
                                r.Decisions(n).ReferenceIndex = ri: r.Decisions(n).ReferenceDigits = .Digits
                                '没有唯一候选时AIndex=0；不能选择歧义列表中的第一项。
                                If .MatchKind = VATS2_SUFFIX_UNIQUE Then r.Decisions(n).AIndex = .Candidates(1).AIndex
                                r.Decisions(n).ReasonFlags = reasons
                                If reasons = 0 Then
                                    r.Decisions(n).Decision = VATS2_WAIVED_THIS_RUN: r.WaivedCount = r.WaivedCount + 1
                                Else
                                    r.Decisions(n).Decision = VATS2_REVIEW_REQUIRED: r.ReviewRequiredCount = r.ReviewRequiredCount + 1
                                End If
                            End If
                        End With
                    Next ri
                End If
            End If
        End With
    Next bi
    r.DecisionCount = n
    If n > 0 Then ReDim Preserve r.Decisions(1 To n)
    r.Status = VATS2_SHORT_POLICY_OK: r.ErrorBIndex = 0: r.ErrorReferenceIndex = 0
Done:
    VATStage2EvaluateShortSuffixPolicy = r
    Exit Function
Invalid:
    '失败不留下部分豁免；原关系和原金额结果始终只读。
    blank.Status = VATS2_SHORT_POLICY_INVALID_CONTRACT: blank.AggregateReliableEqual = aggregateReliableEqual
    blank.ErrorBIndex = r.ErrorBIndex: blank.ErrorReferenceIndex = r.ErrorReferenceIndex: blank.ErrorReason = Err.Description
    r = blank
    Resume Done
End Function

Private Function RowReasons(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, _
    ByRef e As VATS2EffectiveRelationsResult, ByRef m As VATS2EffectiveAmountResult, _
    ByVal bi As Long, ByVal aggregateEqual As Boolean) As Long
    Dim reasons As Long, ri As Long, ai As Long, complete As Boolean, parserRisk As Boolean
    Dim colorMissing As Boolean, otherRisk As Boolean, groupEqual As Boolean
    If Not aggregateEqual Then reasons = VATS2_AGGREGATE_NOT_ELIGIBLE
    Need m.RelationQualityFlags(bi) = e.RowIssueFlags(bi), "关系质量标志在两层之间不一致。"
    With e.EffectiveBMatches(bi)
        complete = (.ReferenceCount > 0 And .MatchState = VATS2_BROW_ALL_UNIQUE And .UniqueMatchCount = .ReferenceCount _
            And .NotFoundCount = 0 And .MultipleMatchCount = 0 And .InvalidMatchCount = 0)
        parserRisk = (.ParserFlags <> 0)
        otherRisk = ((.MatcherFlags And Not VATS2_SHORT_SUFFIX_REVIEW) <> 0)
        colorMissing = Not b.Records(bi).IsCompleted
        For ri = 1 To .ReferenceCount
            With .References(ri)
                parserRisk = parserRisk Or (.ParserFlags <> 0)
                otherRisk = otherRisk Or ((.MatcherFlags And Not VATS2_SHORT_SUFFIX_REVIEW) <> 0)
                If .MatchKind = VATS2_EXACT_UNIQUE Or .MatchKind = VATS2_SUFFIX_UNIQUE Then
                    Need .CandidateCount = 1, "唯一引用的候选数量不符。"
                    Need LBound(.Candidates) = 1 And UBound(.Candidates) = 1, "唯一候选数组错误。"
                    ai = .Candidates(1).AIndex
                    Need ai >= 1 And ai <= a.RecordCount, "候选AIndex越界。"
                    Need a.Records(ai).AIndex = ai, "候选AIndex与快照不一致。"
                    Need VarType(a.Records(ai).InvoiceDigitsRaw) = vbString, "候选A身份不是String。"
                    Need StrComp(.Candidates(1).InvoiceDigits, a.Records(ai).InvoiceDigitsRaw, vbBinaryCompare) = 0, "候选号码与A快照不一致。"
                    colorMissing = colorMissing Or Not a.Records(ai).IsCompleted
                Else
                    complete = False
                End If
            End With
        Next ri
    End With
    If Not complete Then reasons = reasons Or VATS2_ROW_NOT_COMPLETE_UNIQUE
    If parserRisk Then reasons = reasons Or VATS2_PARSER_RISK
    If colorMissing Then reasons = reasons Or VATS2_COLOR_MISSING
    If otherRisk Then reasons = reasons Or VATS2_OTHER_STRUCTURE_RISK
    If m.RelationQualityFlags(bi) <> 0 Then
        reasons = reasons Or VATS2_RELATION_QUALITY_RISK
        Need m.RowStates(bi) = VATS2_RELATION_QUALITY_BLOCKED And Not m.GroupAmounts(bi).AmountEvaluated, "质量阻断行不应已比较金额。"
    Else
        With m.GroupAmounts(bi)
            Need .BIndex = bi And .SourceMatchState = e.EffectiveBMatches(bi).MatchState, "金额组与关系行不一致。"
            Need .ReferenceCount = e.EffectiveBMatches(bi).ReferenceCount And .ParserFlags = e.EffectiveBMatches(bi).ParserFlags And .MatcherFlags = e.EffectiveBMatches(bi).MatcherFlags, "金额组的引用或风险来源错位。"
            Need m.RowStates(bi) = .GroupState, "金额行状态与组状态不一致。"
            If .GroupState = VATS2_GROUP_COMPARED Then
                Need .AmountEvaluated And .Amount.Status = VATS2_AMOUNT_OK, "COMPARED缺少实际成功金额结果。"
                Need .Amount.AmountState = VATS2_AMOUNT_EQUAL Or .Amount.AmountState = VATS2_AMOUNT_MISMATCH, "金额比较状态错误。"
                Need .Amount.ACount = e.EffectiveBMatches(bi).ReferenceCount, "组金额成员计数错位。"
                For ri = 1 To .Amount.ACount
                    Need .Amount.AIndexes(ri) = e.EffectiveBMatches(bi).References(ri).Candidates(1).AIndex, "金额成员AIndex错位。"
                Next ri
                groupEqual = (.Amount.AmountState = VATS2_AMOUNT_EQUAL)
            End If
            If .SameBRowReuse Or .CrossBRowReuse Then reasons = reasons Or VATS2_CONFLICT_RISK
        End With
    End If
    If Not groupEqual Then reasons = reasons Or VATS2_GROUP_NOT_EQUAL
    '仅读取已存在的冲突关联；不重新分组或计算SAME/CROSS。
    If HasConflict(e.EffectiveConflicts, bi) Then reasons = reasons Or VATS2_CONFLICT_RISK
    RowReasons = reasons
End Function

Private Function HasConflict(ByRef conflicts As VATS2BConflictResult, ByVal bi As Long) As Boolean
    Dim i As Long, j As Long
    Need conflicts.ConflictGroupCount >= 0, "冲突计数错误。"
    For i = 1 To conflicts.ConflictGroupCount
        For j = 1 To conflicts.ConflictGroups(i).AssociationCount
            If conflicts.ConflictGroups(i).Associations(j).BIndex = bi Then HasConflict = True: Exit Function
        Next j
    Next i
End Function

Private Sub Need(ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then Err.Raise vbObjectError + 5301, "ShortSuffixPolicy", reason
End Sub
