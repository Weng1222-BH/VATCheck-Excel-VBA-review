Attribute VB_Name = "modVATStage2BRowMatch"
Option Explicit
Option Compare Binary

Public Enum VATS2BRowMatchState
    VATS2_BROW_NO_REFERENCES = 0
    VATS2_BROW_ALL_UNIQUE = 1
    VATS2_BROW_INCOMPLETE = 2
    VATS2_BROW_INVALID_INPUT = 3
End Enum

Public Type VATS2BRowReference
    Digits As String
    Length As Long
    RawFragment As String
    StartIndex As Long
    Context As String
    ParserFlags As Long
    MatchKind As VATS2MatchKind
    MatcherFlags As Long
    CandidateCount As Long
    Candidates() As VATS2MatchCandidate
End Type

Public Type VATS2BRowMatchResult
    OriginalText As String
    ParserFlags As Long
    ReferenceCount As Long
    References() As VATS2BRowReference
    UniqueMatchCount As Long
    NotFoundCount As Long
    MultipleMatchCount As Long
    InvalidMatchCount As Long
    MatcherFlags As Long
    MatchState As VATS2BRowMatchState
End Type

Public Function VATStage2MatchBRow(ByVal rawText As String, _
    ByRef aInvoiceDigits() As String, ByVal aCount As Long) As VATS2BRowMatchResult
    Dim result As VATS2BRowMatchResult, parsed As VATS2ParseResult
    Dim matched As VATS2ReferenceMatchResult, i As Long, j As Long
    'A 数据质量由调用方先通过 AIntegrity 处理；此处不清洗或重扫 A。
    parsed = VATStage2ParseBInvoices(rawText)
    result.OriginalText = parsed.OriginalText
    result.ParserFlags = parsed.Flags
    result.ReferenceCount = parsed.ReferenceCount
    If result.ReferenceCount = 0 Then
        result.MatchState = VATS2_BROW_NO_REFERENCES
        GoTo Done
    End If
    ReDim result.References(1 To result.ReferenceCount)
    For i = 1 To result.ReferenceCount
        matched = VATStage2MatchReference(parsed.References(i).Digits, aInvoiceDigits, aCount)
        With result.References(i)
            .Digits = parsed.References(i).Digits
            .Length = parsed.References(i).Length
            .RawFragment = parsed.References(i).RawFragment
            .StartIndex = parsed.References(i).StartIndex
            .Context = parsed.References(i).Context
            .ParserFlags = parsed.References(i).Flags
            .MatchKind = matched.MatchKind
            .MatcherFlags = matched.Flags
            .CandidateCount = matched.CandidateCount
            If .CandidateCount > 0 Then
                ReDim .Candidates(1 To .CandidateCount)
                For j = 1 To .CandidateCount
                    .Candidates(j) = matched.Candidates(j)
                Next j
            End If
        End With
        '两个 Flags 位空间始终分开，风险不改变唯一候选计数。
        result.MatcherFlags = result.MatcherFlags Or matched.Flags
        Select Case matched.MatchKind
            Case VATS2_EXACT_UNIQUE, VATS2_SUFFIX_UNIQUE
                result.UniqueMatchCount = result.UniqueMatchCount + 1
            Case VATS2_NOT_FOUND
                result.NotFoundCount = result.NotFoundCount + 1
            Case VATS2_EXACT_MULTIPLE, VATS2_SUFFIX_MULTIPLE
                result.MultipleMatchCount = result.MultipleMatchCount + 1
            Case Else
                '包括 INVALID_REFERENCE、INVALID_A_RANGE；未知接口状态也不能伪装为未找到。
                result.InvalidMatchCount = result.InvalidMatchCount + 1
        End Select
    Next i
    If result.InvalidMatchCount > 0 Then
        result.MatchState = VATS2_BROW_INVALID_INPUT
    ElseIf result.UniqueMatchCount = result.ReferenceCount Then
        '仅表示每个引用都有唯一候选，不代表审核通过或金额正确。
        result.MatchState = VATS2_BROW_ALL_UNIQUE
    Else
        result.MatchState = VATS2_BROW_INCOMPLETE
    End If
Done:
    VATStage2MatchBRow = result
End Function
