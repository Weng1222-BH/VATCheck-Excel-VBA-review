Attribute VB_Name = "modVATStage2EffectiveAmount"
Option Explicit
Option Compare Binary

Public Enum VATS2EffectiveAmountStatus
    VATS2_EFFECTIVE_AMOUNT_OK = 0
    VATS2_EFFECTIVE_AMOUNT_INVALID_INPUT = 1
    VATS2_EFFECTIVE_AMOUNT_INVALID_CONTRACT = 2
    VATS2_EFFECTIVE_AMOUNT_GROUP_ERROR = 3
End Enum

'前四个值与冻结 GroupState 对齐；新增状态不冒充金额结论。
Public Enum VATS2EffectiveAmountRowState
    VATS2_EA_NOT_COMPARABLE = 0
    VATS2_EA_COMPARED = 1
    VATS2_EA_AMOUNT_ERROR = 2
    VATS2_EA_INVALID_INPUT = 3
    VATS2_RELATION_QUALITY_BLOCKED = 4
    VATS2_EA_OUT_OF_SCOPE = 5
End Enum

Public Type VATS2EffectiveAmountResult
    Status As VATS2EffectiveAmountStatus
    FullARecordCount As Long
    FullBRecordCount As Long
    BInScope() As Boolean
    BRowIds() As Long
    RowStates() As VATS2EffectiveAmountRowState
    GroupAmounts() As VATS2GroupAmountResult
    RelationQualityFlags() As Long
    RelationIssueCount As Long
    RelationIssues() As VATS2RelationAIssue
    EffectiveBCount As Long
    ComparedCount As Long
    EqualCount As Long
    MismatchCount As Long
    NotComparableCount As Long
    AmountErrorCount As Long
    RelationQualityBlockedCount As Long
    InvalidGroupCount As Long
    ErrorSide As String
    ErrorIndex As Long
    ErrorReferenceIndex As Long
    ErrorReason As String
End Type

Public Function VATStage2EvaluateEffectiveAmounts(ByRef a As VATS2ASnapshotResult, _
    ByRef b As VATS2BSnapshotResult, ByRef effective As VATS2EffectiveRelationsResult) As VATS2EffectiveAmountResult
    Dim r As VATS2EffectiveAmountResult, blank As VATS2EffectiveAmountResult
    Dim aAmounts() As Variant, i As Long
    r.Status = VATS2_EFFECTIVE_AMOUNT_INVALID_INPUT
    If a.Status <> VATS2_SNAPSHOT_OK Or b.Status <> VATS2_SNAPSHOT_OK Or effective.Status <> VATS2_EFFECTIVE_OK Then
        r.ErrorReason = "Snapshot 或 EffectiveRelations 非 OK。": GoTo Done
    End If
    On Error GoTo Invalid
    r.Status = VATS2_EFFECTIVE_AMOUNT_INVALID_CONTRACT
    ValidateInputs a, b, effective, r
    r.FullARecordCount = a.RecordCount: r.FullBRecordCount = b.RecordCount
    r.EffectiveBCount = effective.EffectiveBCount
    r.RelationIssueCount = effective.RelationIssueCount
    If r.RelationIssueCount > 0 Then
        ReDim r.RelationIssues(1 To r.RelationIssueCount)
        For i = 1 To r.RelationIssueCount: r.RelationIssues(i) = effective.RelationIssues(i): Next i
    End If
    '完整 AIndex 空间直接承接原始 Variant，绝不压缩、转换或按颜色过滤。
    If a.RecordCount > 0 Then ReDim aAmounts(1 To a.RecordCount)
    For i = 1 To a.RecordCount
        If IsObject(a.Records(i).AmountRaw) Then
            Set aAmounts(i) = a.Records(i).AmountRaw
        Else
            aAmounts(i) = a.Records(i).AmountRaw
        End If
    Next i
    If b.RecordCount > 0 Then
        ReDim r.BInScope(1 To b.RecordCount): ReDim r.BRowIds(1 To b.RecordCount)
        ReDim r.RowStates(1 To b.RecordCount): ReDim r.GroupAmounts(1 To b.RecordCount)
        ReDim r.RelationQualityFlags(1 To b.RecordCount)
    End If
    r.Status = VATS2_EFFECTIVE_AMOUNT_OK
    For i = 1 To b.RecordCount
        r.BInScope(i) = effective.EffectiveBInScope(i): r.BRowIds(i) = b.Records(i).ExcelRow
        r.RowStates(i) = VATS2_EA_OUT_OF_SCOPE
        r.RelationQualityFlags(i) = effective.RowIssueFlags(i)
        If r.BInScope(i) Then
            If r.RelationQualityFlags(i) <> 0 Then
                '质量门控先于任何金额读取/比较；即使金额无效或碰巧相等也不绕过。
                r.RowStates(i) = VATS2_RELATION_QUALITY_BLOCKED
                r.RelationQualityBlockedCount = r.RelationQualityBlockedCount + 1
            Else
                r.GroupAmounts(i) = VATStage2EvaluateBGroupAmount(effective.EffectiveBMatches(i), i, _
                    effective.EffectiveConflicts, aAmounts, b.Records(i).AmountRaw)
                With r.GroupAmounts(i)
                    r.RowStates(i) = .GroupState
                    Select Case .GroupState
                        Case VATS2_GROUP_COMPARED
                            r.ComparedCount = r.ComparedCount + 1
                            If .Amount.AmountState = VATS2_AMOUNT_EQUAL Then
                                r.EqualCount = r.EqualCount + 1
                            Else
                                r.MismatchCount = r.MismatchCount + 1
                            End If
                        Case VATS2_GROUP_NOT_COMPARABLE
                            r.NotComparableCount = r.NotComparableCount + 1
                        Case VATS2_GROUP_AMOUNT_ERROR
                            r.AmountErrorCount = r.AmountErrorCount + 1
                        Case VATS2_GROUP_INVALID_INPUT
                            r.InvalidGroupCount = r.InvalidGroupCount + 1
                            r.Status = VATS2_EFFECTIVE_AMOUNT_GROUP_ERROR
                            If r.ErrorIndex = 0 Then
                                r.ErrorSide = "B": r.ErrorIndex = i
                                r.ErrorReferenceIndex = .ErrorReferenceIndex
                                r.ErrorReason = "冻结 GroupAmount 返回 GROUP_INVALID_INPUT；详见该行完整结果。"
                            End If
                    End Select
                End With
            End If
        End If
    Next i
Done:
    VATStage2EvaluateEffectiveAmounts = r
    Exit Function
Invalid:
    '契约失败不发布部分比较结果；不能解释为金额不等或默认相等。
    blank.Status = VATS2_EFFECTIVE_AMOUNT_INVALID_CONTRACT
    blank.ErrorSide = r.ErrorSide: blank.ErrorIndex = r.ErrorIndex
    blank.ErrorReferenceIndex = r.ErrorReferenceIndex: blank.ErrorReason = Err.Description
    r = blank
    Resume Done
End Function

Private Sub Need(ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then Err.Raise vbObjectError + 5201, "EffectiveAmount", reason
End Sub

Private Sub ValidateInputs(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, _
    ByRef e As VATS2EffectiveRelationsResult, ByRef r As VATS2EffectiveAmountResult)
    Dim i As Long, j As Long, k As Long, ai As Long, bi As Long, ri As Long, prior As Long
    Dim scopeCount As Long, flags As Long, issuePos As Long, summaryFlags As Long
    Dim invalidA() As Boolean, duplicateA() As Long, checkedConflicts As VATS2BConflictResult
    Need a.RecordCount >= 0 And b.RecordCount >= 0 And a.HeaderRow >= 1 And b.HeaderRow >= 1, "快照计数或表头无效。"
    Need e.FullARecordCount = a.RecordCount And e.FullBRecordCount = b.RecordCount, "完整记录数量不一致。"
    Need ShapeA(a.Records, a.RecordCount) And ShapeB(b.Records, b.RecordCount), "快照数组契约损坏。"
    Need ShapeBool(e.EffectiveBInScope, b.RecordCount) And ShapeRows(e.EffectiveBMatches, b.RecordCount), "B范围或关系数组契约损坏。"
    Need ShapeLong(e.BRowIds, b.RecordCount) And ShapeLong(e.RowIssueFlags, b.RecordCount), "B行号或质量标志数组契约损坏。"
    Need ShapeBool(e.AInvalidValue, a.RecordCount) And ShapeLong(e.ADuplicateGroupIndex, a.RecordCount), "A质量摘要数组契约损坏。"
    Need ShapeIssues(e.RelationIssues, e.RelationIssueCount), "RelationIssues数量或数组错误。"
    prior = a.HeaderRow
    For i = 1 To a.RecordCount
        r.ErrorSide = "A": r.ErrorIndex = i
        Need a.Records(i).AIndex = i And a.Records(i).ExcelRow > prior, "原 AIndex/ExcelRow 映射损坏。"
        prior = a.Records(i).ExcelRow
    Next i
    '只验证冻结 AIntegrity 与摘要的一致性，不重新实现号码清洗或判重。
    With e.AIntegrity
        Need .Status = VATS2_A_SCAN_OK And .RecordCount = a.RecordCount, "AIntegrity 状态或计数错误。"
        Need ShapeLong(.InvalidAIndexes, .InvalidValueCount) And ShapeDuplicates(.DuplicateGroups, .DuplicateGroupCount), "AIntegrity 数组错误。"
        If a.RecordCount > 0 Then
            ReDim invalidA(1 To a.RecordCount): ReDim duplicateA(1 To a.RecordCount)
        End If
        prior = 0
        For i = 1 To .InvalidValueCount
            ai = .InvalidAIndexes(i)
            Need ai > prior And ai <= a.RecordCount, "非法 A 索引越界或重复。"
            invalidA(ai) = True: prior = ai
        Next i
        prior = 0
        For i = 1 To .DuplicateGroupCount
            Need .DuplicateGroups(i).OccurrenceCount >= 2, "重复组少于两个成员。"
            Need ShapeLong(.DuplicateGroups(i).AIndexes, .DuplicateGroups(i).OccurrenceCount), "重复组索引数组错误。"
            Need .DuplicateGroups(i).AIndexes(1) > prior, "重复组首次位置顺序错误。"
            prior = .DuplicateGroups(i).AIndexes(1): k = 0
            For j = 1 To .DuplicateGroups(i).OccurrenceCount
                ai = .DuplicateGroups(i).AIndexes(j)
                Need ai > k And ai <= a.RecordCount, "重复组 AIndex 越界或乱序。"
                Need duplicateA(ai) = 0 And Not invalidA(ai), "重复组成员重复或与非法成员冲突。"
                Need VarType(a.Records(ai).InvoiceDigitsRaw) = vbString, "重复成员不是原 String。"
                Need StrComp(a.Records(ai).InvoiceDigitsRaw, .DuplicateGroups(i).InvoiceDigits, vbBinaryCompare) = 0, "重复组号码与原 A 不符。"
                duplicateA(ai) = i: k = ai
            Next j
        Next i
        If .InvalidValueCount > 0 Then summaryFlags = VATS2_A_DATA_REVIEW
        If .DuplicateGroupCount > 0 Then summaryFlags = summaryFlags Or VATS2_DUPLICATE_A_INVOICE
        Need .Flags = summaryFlags, "AIntegrity Flags 不一致。"
    End With
    For i = 1 To a.RecordCount
        Need e.AInvalidValue(i) = invalidA(i) And e.ADuplicateGroupIndex(i) = duplicateA(i), "A质量摘要与 AIntegrity 不一致。"
    Next i
    prior = b.HeaderRow
    For bi = 1 To b.RecordCount
        r.ErrorSide = "B": r.ErrorIndex = bi: r.ErrorReferenceIndex = 0
        Need b.Records(bi).BIndex = bi And b.Records(bi).ExcelRow > prior, "原 BIndex/ExcelRow 映射损坏。"
        prior = b.Records(bi).ExcelRow
        Need e.BRowIds(bi) = prior, "Effective B行号错位。"
        ValidateRow e.EffectiveBMatches(bi), a
        If e.EffectiveBInScope(bi) Then
            scopeCount = scopeCount + 1
            If VarType(b.Records(bi).SupplierTextRaw) = vbString Then
                Need StrComp(e.EffectiveBMatches(bi).OriginalText, b.Records(bi).SupplierTextRaw, vbBinaryCompare) = 0, "B原文错位。"
            Else
                Need e.EffectiveBMatches(bi).OriginalText = "" And e.EffectiveBMatches(bi).ReferenceCount = 0, "非 String B原文的占位错误。"
            End If
        Else
            Need Not b.Records(bi).IsCompleted, "完成 B 不可从 Effective 范围消失。"
            Need e.EffectiveBMatches(bi).OriginalText = "" And e.EffectiveBMatches(bi).ReferenceCount = 0, "范围外 B 必须为零引用占位。"
        End If
        flags = 0
        For ri = 1 To e.EffectiveBMatches(bi).ReferenceCount
            r.ErrorReferenceIndex = ri
            With e.EffectiveBMatches(bi).References(ri)
                If .MatchKind = VATS2_EXACT_UNIQUE Or .MatchKind = VATS2_SUFFIX_UNIQUE Then
                    ai = .Candidates(1).AIndex: summaryFlags = 0
                    If e.AInvalidValue(ai) Then summaryFlags = VATS2_RELATION_A_INVALID
                    If e.ADuplicateGroupIndex(ai) > 0 Then summaryFlags = summaryFlags Or VATS2_RELATION_A_DUPLICATE
                    If summaryFlags <> 0 Then
                        issuePos = issuePos + 1
                        Need issuePos <= e.RelationIssueCount, "问题关联缺少 RelationIssue。"
                        With e.RelationIssues(issuePos)
                            Need .BIndex = bi And .ReferenceIndex = ri And .AIndex = ai, "RelationIssue 索引错位、重复或乱序。"
                            Need .ExcelRow = b.Records(bi).ExcelRow And .AExcelRow = a.Records(ai).ExcelRow, "RelationIssue ExcelRow 错位。"
                            Need .Flags = summaryFlags, "RelationIssue Flags 与 A质量不符。"
                        End With
                        flags = flags Or summaryFlags
                    End If
                End If
            End With
        Next ri
        Need e.RowIssueFlags(bi) = flags, "RowIssueFlags 与 RelationIssues 不一致。"
    Next bi
    Need scopeCount = e.EffectiveBCount, "EffectiveBCount 与范围不符。"
    Need issuePos = e.RelationIssueCount, "RelationIssues 多余或越界。"
    r.ErrorSide = "CONFLICT": r.ErrorIndex = 0: r.ErrorReferenceIndex = 0
    Need e.EffectiveConflicts.Status = VATS2_BCONFLICT_OK, "EffectiveConflicts 非 OK。"
    '使用冻结扫描器验证完整冲突来源；不另写冲突算法，也不修复/替换输入。
    checkedConflicts = VATStage2ScanBConflicts(e.EffectiveBMatches, e.BRowIds, b.RecordCount)
    Need SameConflicts(checkedConflicts, e.EffectiveConflicts), "EffectiveConflicts 计数、关联或映射损坏。"
    r.ErrorSide = "": r.ErrorIndex = 0: r.ErrorReferenceIndex = 0
End Sub

Private Sub ValidateRow(ByRef row As VATS2BRowMatchResult, ByRef a As VATS2ASnapshotResult)
    Dim i As Long, j As Long, ai As Long, prior As Long, u As Long, nf As Long, m As Long, mf As Long
    Dim state As VATS2BRowMatchState
    Need ShapeReferences(row.References, row.ReferenceCount), "引用数组损坏。"
    For i = 1 To row.ReferenceCount
        With row.References(i)
            Need .Length > 0 And .Length = Len(.Digits), "引用长度损坏。"
            Need ShapeCandidates(.Candidates, .CandidateCount), "候选数组损坏。"
            Select Case .MatchKind
                Case VATS2_NOT_FOUND: Need .CandidateCount = 0, "NOT_FOUND 候选数量错误。": nf = nf + 1
                Case VATS2_EXACT_UNIQUE, VATS2_SUFFIX_UNIQUE: Need .CandidateCount = 1, "UNIQUE 候选数量错误。": u = u + 1
                Case VATS2_EXACT_MULTIPLE, VATS2_SUFFIX_MULTIPLE: Need .CandidateCount >= 2, "MULTIPLE 候选数量错误。": m = m + 1
                Case Else: Need False, "无效 Matcher 状态。"
            End Select
            mf = mf Or .MatcherFlags: prior = 0
            For j = 1 To .CandidateCount
                ai = .Candidates(j).AIndex
                Need ai > prior And ai <= a.RecordCount, "候选 AIndex 越界或错序。"
                Need VarType(a.Records(ai).InvoiceDigitsRaw) = vbString, "候选 A 身份不是 String。"
                Need StrComp(.Candidates(j).InvoiceDigits, a.Records(ai).InvoiceDigitsRaw, vbBinaryCompare) = 0, "候选完整号与原 A 不一致。"
                prior = ai
            Next j
        End With
    Next i
    If row.ReferenceCount = 0 Then
        state = VATS2_BROW_NO_REFERENCES
    ElseIf u = row.ReferenceCount Then
        state = VATS2_BROW_ALL_UNIQUE
    Else
        state = VATS2_BROW_INCOMPLETE
    End If
    Need row.UniqueMatchCount = u And row.NotFoundCount = nf And row.MultipleMatchCount = m And row.InvalidMatchCount = 0, "行引用计数不一致。"
    Need row.MatchState = state And row.MatcherFlags = mf, "行状态或 MatcherFlags 不一致。"
End Sub

Private Function SameConflicts(ByRef x As VATS2BConflictResult, ByRef y As VATS2BConflictResult) As Boolean
    Dim i As Long, j As Long
    On Error GoTo Invalid
    If x.Status <> VATS2_BCONFLICT_OK Or y.Status <> x.Status Then Exit Function
    If x.BRecordCount <> y.BRecordCount Or x.UniqueAssociationCount <> y.UniqueAssociationCount Or x.ConflictGroupCount <> y.ConflictGroupCount Then Exit Function
    If x.ErrorBIndex <> y.ErrorBIndex Or x.ErrorReferenceIndex <> y.ErrorReferenceIndex Then Exit Function
    If Not ShapeConflicts(y.ConflictGroups, y.ConflictGroupCount) Then Exit Function
    For i = 1 To x.ConflictGroupCount
        With y.ConflictGroups(i)
            If .AIndex <> x.ConflictGroups(i).AIndex Or .AInvoiceDigits <> x.ConflictGroups(i).AInvoiceDigits Then Exit Function
            If .AssociationCount <> x.ConflictGroups(i).AssociationCount Or .DistinctBRowCount <> x.ConflictGroups(i).DistinctBRowCount Or .Flags <> x.ConflictGroups(i).Flags Then Exit Function
            If Not ShapeAssociations(.Associations, .AssociationCount) Then Exit Function
            For j = 1 To .AssociationCount
                With .Associations(j)
                    If .BIndex <> x.ConflictGroups(i).Associations(j).BIndex Or .BRowId <> x.ConflictGroups(i).Associations(j).BRowId Then Exit Function
                    If .ReferenceIndex <> x.ConflictGroups(i).Associations(j).ReferenceIndex Or .AIndex <> x.ConflictGroups(i).Associations(j).AIndex Then Exit Function
                    If .MatchKind <> x.ConflictGroups(i).Associations(j).MatchKind Or .ReferenceDigits <> x.ConflictGroups(i).Associations(j).ReferenceDigits Or .AInvoiceDigits <> x.ConflictGroups(i).Associations(j).AInvoiceDigits Then Exit Function
                End With
            Next j
        End With
    Next i
    SameConflicts = True
Invalid:
End Function

Private Function ShapeA(ByRef items() As VATS2ASnapshotRecord, ByVal count As Long) As Boolean
    Dim first As Long, last As Long
    If count < 0 Then Exit Function
    On Error GoTo Unallocated
    first = LBound(items): last = UBound(items)
    ShapeA = (count > 0 And first = 1 And last = count)
    Exit Function
Unallocated:
    ShapeA = (Err.Number = 9 And count = 0)
End Function

Private Function ShapeB(ByRef items() As VATS2BSnapshotRecord, ByVal count As Long) As Boolean
    Dim first As Long, last As Long
    If count < 0 Then Exit Function
    On Error GoTo Unallocated
    first = LBound(items): last = UBound(items)
    ShapeB = (count > 0 And first = 1 And last = count)
    Exit Function
Unallocated:
    ShapeB = (Err.Number = 9 And count = 0)
End Function

Private Function ShapeBool(ByRef items() As Boolean, ByVal count As Long) As Boolean
    Dim first As Long, last As Long
    If count < 0 Then Exit Function
    On Error GoTo Unallocated
    first = LBound(items): last = UBound(items)
    ShapeBool = (count > 0 And first = 1 And last = count)
    Exit Function
Unallocated:
    ShapeBool = (Err.Number = 9 And count = 0)
End Function

Private Function ShapeLong(ByRef items() As Long, ByVal count As Long) As Boolean
    Dim first As Long, last As Long
    If count < 0 Then Exit Function
    On Error GoTo Unallocated
    first = LBound(items): last = UBound(items)
    ShapeLong = (count > 0 And first = 1 And last = count)
    Exit Function
Unallocated:
    ShapeLong = (Err.Number = 9 And count = 0)
End Function

Private Function ShapeRows(ByRef items() As VATS2BRowMatchResult, ByVal count As Long) As Boolean
    Dim first As Long, last As Long
    If count < 0 Then Exit Function
    On Error GoTo Unallocated
    first = LBound(items): last = UBound(items)
    ShapeRows = (count > 0 And first = 1 And last = count)
    Exit Function
Unallocated:
    ShapeRows = (Err.Number = 9 And count = 0)
End Function

Private Function ShapeIssues(ByRef items() As VATS2RelationAIssue, ByVal count As Long) As Boolean
    Dim first As Long, last As Long
    If count < 0 Then Exit Function
    On Error GoTo Unallocated
    first = LBound(items): last = UBound(items)
    ShapeIssues = (count > 0 And first = 1 And last = count)
    Exit Function
Unallocated:
    ShapeIssues = (Err.Number = 9 And count = 0)
End Function

Private Function ShapeDuplicates(ByRef items() As VATS2ADuplicateGroup, ByVal count As Long) As Boolean
    Dim first As Long, last As Long
    If count < 0 Then Exit Function
    On Error GoTo Unallocated
    first = LBound(items): last = UBound(items)
    ShapeDuplicates = (count > 0 And first = 1 And last = count)
    Exit Function
Unallocated:
    ShapeDuplicates = (Err.Number = 9 And count = 0)
End Function

Private Function ShapeReferences(ByRef items() As VATS2BRowReference, ByVal count As Long) As Boolean
    Dim first As Long, last As Long
    If count < 0 Then Exit Function
    On Error GoTo Unallocated
    first = LBound(items): last = UBound(items)
    ShapeReferences = (count > 0 And first = 1 And last = count)
    Exit Function
Unallocated:
    ShapeReferences = (Err.Number = 9 And count = 0)
End Function

Private Function ShapeCandidates(ByRef items() As VATS2MatchCandidate, ByVal count As Long) As Boolean
    Dim first As Long, last As Long
    If count < 0 Then Exit Function
    On Error GoTo Unallocated
    first = LBound(items): last = UBound(items)
    ShapeCandidates = (count > 0 And first = 1 And last = count)
    Exit Function
Unallocated:
    ShapeCandidates = (Err.Number = 9 And count = 0)
End Function

Private Function ShapeConflicts(ByRef items() As VATS2BConflictGroup, ByVal count As Long) As Boolean
    Dim first As Long, last As Long
    If count < 0 Then Exit Function
    On Error GoTo Unallocated
    first = LBound(items): last = UBound(items)
    ShapeConflicts = (count > 0 And first = 1 And last = count)
    Exit Function
Unallocated:
    ShapeConflicts = (Err.Number = 9 And count = 0)
End Function

Private Function ShapeAssociations(ByRef items() As VATS2BAssociation, ByVal count As Long) As Boolean
    Dim first As Long, last As Long
    If count < 0 Then Exit Function
    On Error GoTo Unallocated
    first = LBound(items): last = UBound(items)
    ShapeAssociations = (count > 0 And first = 1 And last = count)
    Exit Function
Unallocated:
    ShapeAssociations = (Err.Number = 9 And count = 0)
End Function
