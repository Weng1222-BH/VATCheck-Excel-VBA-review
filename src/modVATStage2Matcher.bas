Attribute VB_Name = "modVATStage2Matcher"
Option Explicit
Option Compare Binary

Public Enum VATS2MatchKind
    VATS2_NOT_FOUND = 0
    VATS2_EXACT_UNIQUE = 1
    VATS2_EXACT_MULTIPLE = 2
    VATS2_SUFFIX_UNIQUE = 3
    VATS2_SUFFIX_MULTIPLE = 4
    VATS2_INVALID_REFERENCE = 5
    VATS2_INVALID_A_RANGE = 6
End Enum

Public Enum VATS2MatcherFlags
    VATS2_MATCH_FLAG_NONE = 0
    VATS2_SHORT_SUFFIX_REVIEW = 1
End Enum

Public Type VATS2MatchCandidate
    AIndex As Long
    InvoiceDigits As String
End Type

Public Type VATS2ReferenceMatchResult
    ReferenceDigits As String
    ReferenceLength As Long
    MatchKind As VATS2MatchKind
    CandidateCount As Long
    Candidates() As VATS2MatchCandidate
    Flags As Long
End Type

Public Function VATStage2MatchReference(ByVal referenceDigits As String, _
    ByRef aInvoiceDigits() As String, ByVal aCount As Long) As VATS2ReferenceMatchResult
    Dim result As VATS2ReferenceMatchResult
    Dim first As Long, last As Long, i As Long, ch As String
    result.ReferenceDigits = referenceDigits
    result.ReferenceLength = Len(referenceDigits)
    '仅检查单个 B 引用的接口前提，不清洗、不转换号码。
    result.MatchKind = VATS2_INVALID_REFERENCE
    If result.ReferenceLength = 0 Then GoTo Done
    For i = 1 To result.ReferenceLength
        ch = Mid$(referenceDigits, i, 1)
        If ch < "0" Or ch > "9" Then GoTo Done
    Next i
    result.MatchKind = VATS2_INVALID_A_RANGE
    If aCount < 0 Then GoTo Done
    If aCount = 0 Then
        result.MatchKind = VATS2_NOT_FOUND
        GoTo Done
    End If
    If Not MatchArrayBounds(aInvoiceDigits, first, last) Then GoTo Done
    'aCount 表示从数组实际下界开始的有效元素数；候选保存原下标。
    If CDbl(aCount) > CDbl(last) - CDbl(first) + 1 Then GoTo Done
    last = CLng(CDbl(first) + CDbl(aCount) - 1)
    result.MatchKind = VATS2_NOT_FOUND
    ReDim result.Candidates(1 To aCount)
    For i = first To last
        If StrComp(aInvoiceDigits(i), referenceDigits, vbBinaryCompare) = 0 Then
            AddMatchCandidate result, i, aInvoiceDigits(i)
        End If
    Next i
    If result.CandidateCount > 0 Then
        If result.CandidateCount = 1 Then
            result.MatchKind = VATS2_EXACT_UNIQUE
        Else
            result.MatchKind = VATS2_EXACT_MULTIPLE
        End If
        GoTo FinishCandidates
    End If
    '只有不存在 exact 才进入 suffix；短引用即使未找到候选也保留此路径提示。
    If result.ReferenceLength <= 6 Then result.Flags = VATS2_SHORT_SUFFIX_REVIEW
    For i = first To last
        If Len(aInvoiceDigits(i)) >= result.ReferenceLength Then
            If StrComp(Right$(aInvoiceDigits(i), result.ReferenceLength), referenceDigits, vbBinaryCompare) = 0 Then
                AddMatchCandidate result, i, aInvoiceDigits(i)
            End If
        End If
    Next i
    If result.CandidateCount = 1 Then result.MatchKind = VATS2_SUFFIX_UNIQUE
    If result.CandidateCount > 1 Then result.MatchKind = VATS2_SUFFIX_MULTIPLE
FinishCandidates:
    If result.CandidateCount = 0 Then
        Erase result.Candidates
    Else
        ReDim Preserve result.Candidates(1 To result.CandidateCount)
    End If
Done:
    VATStage2MatchReference = result
End Function

Private Sub AddMatchCandidate(ByRef result As VATS2ReferenceMatchResult, _
    ByVal aIndex As Long, ByVal invoiceDigits As String)
    result.CandidateCount = result.CandidateCount + 1
    With result.Candidates(result.CandidateCount)
        .AIndex = aIndex
        .InvoiceDigits = invoiceDigits
    End With
End Sub

Private Function MatchArrayBounds(ByRef values() As String, ByRef first As Long, ByRef last As Long) As Boolean
    '未分配数组只转换成明确的接口错误，不吞掉匹配过程中的其他错误。
    On Error GoTo Unallocated
    first = LBound(values): last = UBound(values)
    MatchArrayBounds = True
Unallocated:
End Function
