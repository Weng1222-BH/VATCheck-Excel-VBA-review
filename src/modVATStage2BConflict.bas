Attribute VB_Name = "modVATStage2BConflict"
Option Explicit
Option Compare Binary

Public Enum VATS2BConflictStatus
    VATS2_BCONFLICT_OK = 0
    VATS2_BCONFLICT_INVALID_RANGE = 1
    VATS2_BCONFLICT_INVALID_CONTRACT = 2
End Enum

Public Enum VATS2BConflictFlags
    VATS2_BCONFLICT_NONE = 0
    VATS2_SAME_B_ROW_REUSE = 1
    VATS2_CROSS_B_ROW_REUSE = 2
End Enum

Public Type VATS2BAssociation
    BIndex As Long
    BRowId As Long
    ReferenceIndex As Long
    ReferenceDigits As String
    MatchKind As VATS2MatchKind
    AIndex As Long
    AInvoiceDigits As String
End Type

Public Type VATS2BConflictGroup
    AIndex As Long
    AInvoiceDigits As String
    AssociationCount As Long
    DistinctBRowCount As Long
    Associations() As VATS2BAssociation
    Flags As Long
End Type

Public Type VATS2BConflictResult
    BRecordCount As Long
    UniqueAssociationCount As Long
    ConflictGroupCount As Long
    ConflictGroups() As VATS2BConflictGroup
    Status As VATS2BConflictStatus
    ErrorBIndex As Long
    ErrorReferenceIndex As Long
End Type

Public Function VATStage2ScanBConflicts(ByRef bRows() As VATS2BRowMatchResult, _
    ByRef bRowIds() As Long, ByVal bCount As Long) As VATS2BConflictResult
    Dim result As VATS2BConflictResult, allLinks() As VATS2BAssociation
    Dim rowFirst As Long, idFirst As Long, offset As Long, bi As Long, ri As Long
    Dim totalRefs As Long, ai As Long, n As Long, g As Long, k As Long, previousB As Long
    Dim byA As Object, order As Collection, links As Collection
    result.Status = VATS2_BCONFLICT_INVALID_RANGE
    If bCount < 0 Then GoTo Done
    If bCount = 0 Then result.Status = VATS2_BCONFLICT_OK: GoTo Done
    If Not ConflictInputRange(bRows, bRowIds, bCount, rowFirst, idFirst) Then GoTo Done
    result.BRecordCount = bCount
    '先核对引用数组契约；不按行状态或风险过滤任何 UNIQUE。
    For offset = 0 To bCount - 1
        bi = rowFirst + offset: ri = 0
        If Not ConflictReferenceRange(bRows(bi)) Then GoTo BadContract
        totalRefs = totalRefs + bRows(bi).ReferenceCount
    Next offset
    result.Status = VATS2_BCONFLICT_OK
    If totalRefs = 0 Then GoTo Done
    ReDim allLinks(1 To totalRefs)
    Set byA = CreateObject("Scripting.Dictionary")
    Set order = New Collection
    For offset = 0 To bCount - 1
        bi = rowFirst + offset
        For ri = 1 To bRows(bi).ReferenceCount
            With bRows(bi).References(ri)
                If .MatchKind = VATS2_EXACT_UNIQUE Or .MatchKind = VATS2_SUFFIX_UNIQUE Then
                    If Not ConflictUniqueCandidate(bRows(bi).References(ri)) Then GoTo BadContract
                    ai = .Candidates(1).AIndex
                    If byA.Exists(ai) Then
                        Set links = byA.Item(ai)
                        '同一 AIndex 必须来自同一个 A 集合，号码不一致时不伪造冲突组。
                        If StrComp(allLinks(links(1)).AInvoiceDigits, .Candidates(1).InvoiceDigits, vbBinaryCompare) <> 0 Then GoTo BadContract
                    Else
                        Set links = New Collection
                        byA.Add ai, links
                        order.Add ai
                    End If
                    n = n + 1
                    allLinks(n).BIndex = bi
                    allLinks(n).BRowId = bRowIds(idFirst + offset)
                    allLinks(n).ReferenceIndex = ri
                    allLinks(n).ReferenceDigits = .Digits
                    allLinks(n).MatchKind = .MatchKind
                    allLinks(n).AIndex = ai
                    allLinks(n).AInvoiceDigits = .Candidates(1).InvoiceDigits
                    links.Add n
                    If links.Count = 2 Then result.ConflictGroupCount = result.ConflictGroupCount + 1
                End If
            End With
        Next ri
    Next offset
    result.UniqueAssociationCount = n
    If result.ConflictGroupCount = 0 Then GoTo Done
    ReDim result.ConflictGroups(1 To result.ConflictGroupCount)
    For k = 1 To order.Count
        Set links = byA.Item(order(k))
        If links.Count > 1 Then
            g = g + 1
            With result.ConflictGroups(g)
                .AIndex = order(k)
                .AInvoiceDigits = allLinks(links(1)).AInvoiceDigits
                .AssociationCount = links.Count
                ReDim .Associations(1 To links.Count)
                For n = 1 To links.Count
                    .Associations(n) = allLinks(links(n))
                    bi = .Associations(n).BIndex
                    '关联已按 B 输入顺序排列；同一 BIndex 的关联必然相邻。
                    If n = 1 Then
                        .DistinctBRowCount = 1
                    ElseIf bi = previousB Then
                        .Flags = .Flags Or VATS2_SAME_B_ROW_REUSE
                    Else
                        .DistinctBRowCount = .DistinctBRowCount + 1
                    End If
                    previousB = bi
                Next n
                If .DistinctBRowCount > 1 Then .Flags = .Flags Or VATS2_CROSS_B_ROW_REUSE
            End With
        End If
    Next k
    GoTo Done
BadContract:
    '接口异常不返回可误用的部分成功计数或冲突结果。
    result.Status = VATS2_BCONFLICT_INVALID_CONTRACT
    result.ErrorBIndex = bi: result.ErrorReferenceIndex = ri
    result.UniqueAssociationCount = 0: result.ConflictGroupCount = 0
    Erase result.ConflictGroups
Done:
    VATStage2ScanBConflicts = result
End Function

Private Function ConflictInputRange(ByRef rows() As VATS2BRowMatchResult, ByRef ids() As Long, _
    ByVal count As Long, ByRef rowFirst As Long, ByRef idFirst As Long) As Boolean
    On Error GoTo InvalidRange
    rowFirst = LBound(rows): idFirst = LBound(ids)
    If CDbl(count) > CDbl(UBound(rows)) - CDbl(rowFirst) + 1 Then Exit Function
    If CDbl(count) > CDbl(UBound(ids)) - CDbl(idFirst) + 1 Then Exit Function
    ConflictInputRange = True
InvalidRange:
End Function

Private Function ConflictReferenceRange(ByRef row As VATS2BRowMatchResult) As Boolean
    On Error GoTo InvalidRange
    If row.ReferenceCount < 0 Then Exit Function
    If row.ReferenceCount > 0 Then
        If LBound(row.References) <> 1 Then Exit Function
        If UBound(row.References) <> row.ReferenceCount Then Exit Function
    End If
    ConflictReferenceRange = True
InvalidRange:
End Function

Private Function ConflictUniqueCandidate(ByRef reference As VATS2BRowReference) As Boolean
    On Error GoTo InvalidCandidate
    If reference.CandidateCount <> 1 Then Exit Function
    If LBound(reference.Candidates) <> 1 Then Exit Function
    If UBound(reference.Candidates) <> 1 Then Exit Function
    ConflictUniqueCandidate = True
InvalidCandidate:
End Function
