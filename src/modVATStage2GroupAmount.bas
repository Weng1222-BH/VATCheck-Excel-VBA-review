Attribute VB_Name = "modVATStage2GroupAmount"
Option Explicit
Option Compare Binary

Public Enum VATS2GroupState
    VATS2_GROUP_NOT_COMPARABLE = 0
    VATS2_GROUP_COMPARED = 1
    VATS2_GROUP_AMOUNT_ERROR = 2
    VATS2_GROUP_INVALID_INPUT = 3
End Enum

Public Enum VATS2GroupReasons
    VATS2_GROUP_REASON_NONE = 0
    VATS2_GROUP_NO_REFERENCES = 1
    VATS2_GROUP_INCOMPLETE_RELATION = 2
    VATS2_GROUP_INVALID_RELATION = 4
    VATS2_GROUP_SAME_ROW_A_REUSE = 8
    VATS2_GROUP_INVALID_CONFLICT = 16
    VATS2_GROUP_INVALID_A_MAPPING = 32
End Enum

Public Type VATS2GroupAmountResult
    BIndex As Long
    SourceMatchState As VATS2BRowMatchState
    SourceConflictStatus As VATS2BConflictStatus
    ReferenceCount As Long
    UniqueMatchCount As Long
    NotFoundCount As Long
    MultipleMatchCount As Long
    InvalidMatchCount As Long
    ParserFlags As Long
    MatcherFlags As Long
    SameBRowReuse As Boolean
    CrossBRowReuse As Boolean
    GroupState As VATS2GroupState
    Reasons As Long
    ErrorReferenceIndex As Long
    ErrorAIndex As Long
    AmountEvaluated As Boolean
    Amount As VATS2AmountResult
End Type

Public Function VATStage2EvaluateBGroupAmount(ByRef bRow As VATS2BRowMatchResult, ByVal bIndex As Long, _
    ByRef conflicts As VATS2BConflictResult, ByRef aAmounts() As Variant, ByVal bAmount As Variant) As VATS2GroupAmountResult
    Dim result As VATS2GroupAmountResult, indexes() As Long, amounts() As Variant
    Dim i As Long, ai As Long, first As Long, last As Long, rowValid As Boolean, conflictValid As Boolean
    result.BIndex = bIndex
    result.SourceMatchState = bRow.MatchState: result.SourceConflictStatus = conflicts.Status
    result.ReferenceCount = bRow.ReferenceCount: result.UniqueMatchCount = bRow.UniqueMatchCount
    result.NotFoundCount = bRow.NotFoundCount: result.MultipleMatchCount = bRow.MultipleMatchCount
    result.InvalidMatchCount = bRow.InvalidMatchCount
    result.ParserFlags = bRow.ParserFlags: result.MatcherFlags = bRow.MatcherFlags
    Select Case bRow.MatchState
        Case VATS2_BROW_NO_REFERENCES: result.Reasons = VATS2_GROUP_NO_REFERENCES
        Case VATS2_BROW_INCOMPLETE: result.Reasons = VATS2_GROUP_INCOMPLETE_RELATION
        Case VATS2_BROW_INVALID_INPUT: result.Reasons = VATS2_GROUP_INVALID_RELATION
    End Select
    rowValid = GroupRowContract(bRow, result.SameBRowReuse)
    If Not rowValid Then result.Reasons = result.Reasons Or VATS2_GROUP_INVALID_RELATION
    If rowValid Then
        conflictValid = GroupCurrentConflicts(bRow, bIndex, conflicts, result.SameBRowReuse, result.CrossBRowReuse)
    End If
    If conflicts.Status <> VATS2_BCONFLICT_OK Or (rowValid And Not conflictValid) Then
        result.Reasons = result.Reasons Or VATS2_GROUP_INVALID_CONFLICT
    End If
    If result.SameBRowReuse Then result.Reasons = result.Reasons Or VATS2_GROUP_SAME_ROW_A_REUSE
    If (result.Reasons And (VATS2_GROUP_INVALID_RELATION Or VATS2_GROUP_INVALID_CONFLICT)) <> 0 Then
        result.GroupState = VATS2_GROUP_INVALID_INPUT
        GoTo Done
    End If
    If result.Reasons <> 0 Then GoTo Done
    '只有全 UNIQUE 且本行无重复计数歧义时才读取金额。
    If Not GroupAmountBounds(aAmounts, first, last) Then
        result.Reasons = VATS2_GROUP_INVALID_A_MAPPING
        result.GroupState = VATS2_GROUP_INVALID_INPUT
        GoTo Done
    End If
    ReDim indexes(1 To bRow.ReferenceCount): ReDim amounts(1 To bRow.ReferenceCount)
    For i = 1 To bRow.ReferenceCount
        ai = bRow.References(i).Candidates(1).AIndex
        If ai < first Or ai > last Then
            result.Reasons = VATS2_GROUP_INVALID_A_MAPPING
            result.GroupState = VATS2_GROUP_INVALID_INPUT
            result.ErrorReferenceIndex = i: result.ErrorAIndex = ai
            GoTo Done
        End If
        indexes(i) = ai
        '直接使用原 AIndex 读取，绝不把组内位置当成 AIndex。
        If IsObject(aAmounts(ai)) Then
            Set amounts(i) = aAmounts(ai)
        Else
            amounts(i) = aAmounts(ai)
        End If
    Next i
    result.Amount = VATStage2CompareAmounts(indexes, amounts, bRow.ReferenceCount, bAmount)
    result.AmountEvaluated = True
    If result.Amount.Status = VATS2_AMOUNT_OK Then
        result.GroupState = VATS2_GROUP_COMPARED
    Else
        result.GroupState = VATS2_GROUP_AMOUNT_ERROR
    End If
Done:
    VATStage2EvaluateBGroupAmount = result
End Function

Private Function GroupRowContract(ByRef row As VATS2BRowMatchResult, ByRef sameRow As Boolean) As Boolean
    Dim i As Long, u As Long, nf As Long, multi As Long, invalid As Long, expected As VATS2BRowMatchState
    Dim seen As Object, ai As Long
    On Error GoTo InvalidContract
    If row.ReferenceCount < 0 Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    If row.ReferenceCount > 0 Then
        If LBound(row.References) <> 1 Or UBound(row.References) <> row.ReferenceCount Then Exit Function
    End If
    For i = 1 To row.ReferenceCount
        With row.References(i)
            If .CandidateCount < 0 Then Exit Function
            If .CandidateCount > 0 Then
                If LBound(.Candidates) <> 1 Or UBound(.Candidates) <> .CandidateCount Then Exit Function
            End If
            Select Case .MatchKind
                Case VATS2_EXACT_UNIQUE, VATS2_SUFFIX_UNIQUE
                    If .CandidateCount <> 1 Then Exit Function
                    u = u + 1: ai = .Candidates(1).AIndex
                    '本行去重歧义的局部不变量检查，防止缺失冲突组时重复求和；不删除引用。
                    If seen.Exists(ai) Then sameRow = True Else seen.Add ai, True
                Case VATS2_NOT_FOUND
                    If .CandidateCount <> 0 Then Exit Function
                    nf = nf + 1
                Case VATS2_EXACT_MULTIPLE, VATS2_SUFFIX_MULTIPLE
                    If .CandidateCount < 2 Then Exit Function
                    multi = multi + 1
                Case Else
                    invalid = invalid + 1
            End Select
        End With
    Next i
    If u <> row.UniqueMatchCount Or nf <> row.NotFoundCount Or multi <> row.MultipleMatchCount Or invalid <> row.InvalidMatchCount Then Exit Function
    If row.ReferenceCount = 0 Then
        expected = VATS2_BROW_NO_REFERENCES
    ElseIf invalid > 0 Then
        expected = VATS2_BROW_INVALID_INPUT
    ElseIf u = row.ReferenceCount Then
        expected = VATS2_BROW_ALL_UNIQUE
    Else
        expected = VATS2_BROW_INCOMPLETE
    End If
    GroupRowContract = (row.MatchState = expected)
InvalidContract:
End Function

Private Function GroupCurrentConflicts(ByRef row As VATS2BRowMatchResult, ByVal bi As Long, _
    ByRef conflicts As VATS2BConflictResult, ByRef sameRow As Boolean, ByRef crossRow As Boolean) As Boolean
    Dim g As Long, j As Long, hits As Long, ri As Long, seenRefs As Object
    On Error GoTo InvalidContract
    If conflicts.Status <> VATS2_BCONFLICT_OK Then Exit Function
    If conflicts.BRecordCount < 1 Or conflicts.ConflictGroupCount < 0 Then Exit Function
    If conflicts.ConflictGroupCount > 0 Then
        If LBound(conflicts.ConflictGroups) <> 1 Or UBound(conflicts.ConflictGroups) <> conflicts.ConflictGroupCount Then Exit Function
    End If
    Set seenRefs = CreateObject("Scripting.Dictionary")
    For g = 1 To conflicts.ConflictGroupCount
        With conflicts.ConflictGroups(g)
            If .AssociationCount < 2 Or .DistinctBRowCount < 1 Or .DistinctBRowCount > .AssociationCount Then Exit Function
            If LBound(.Associations) <> 1 Or UBound(.Associations) <> .AssociationCount Then Exit Function
            hits = 0
            For j = 1 To .AssociationCount
                If .Associations(j).BIndex = bi Then
                    ri = .Associations(j).ReferenceIndex
                    If ri < 1 Or ri > row.ReferenceCount Then Exit Function
                    If seenRefs.Exists(ri) Then Exit Function
                    seenRefs.Add ri, True
                    If row.References(ri).CandidateCount <> 1 Then Exit Function
                    If row.References(ri).MatchKind <> VATS2_EXACT_UNIQUE And row.References(ri).MatchKind <> VATS2_SUFFIX_UNIQUE Then Exit Function
                    If .Associations(j).MatchKind <> row.References(ri).MatchKind Then Exit Function
                    If .Associations(j).ReferenceDigits <> row.References(ri).Digits Then Exit Function
                    If .Associations(j).AIndex <> .AIndex Or row.References(ri).Candidates(1).AIndex <> .AIndex Then Exit Function
                    If .Associations(j).AInvoiceDigits <> .AInvoiceDigits Or row.References(ri).Candidates(1).InvoiceDigits <> .AInvoiceDigits Then Exit Function
                    hits = hits + 1
                End If
            Next j
            '不把组级 SAME 标志套用给本组所有 B 行。
            If hits >= 2 Then sameRow = True
            If hits > 0 And .DistinctBRowCount >= 2 Then crossRow = True
        End With
    Next g
    GroupCurrentConflicts = True
InvalidContract:
End Function

Private Function GroupAmountBounds(ByRef amounts() As Variant, ByRef first As Long, ByRef last As Long) As Boolean
    On Error GoTo Unallocated
    first = LBound(amounts): last = UBound(amounts)
    GroupAmountBounds = True
Unallocated:
End Function
