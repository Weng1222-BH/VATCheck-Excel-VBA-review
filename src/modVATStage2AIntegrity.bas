Attribute VB_Name = "modVATStage2AIntegrity"
Option Explicit
Option Compare Binary

Public Enum VATS2AIntegrityStatus
    VATS2_A_SCAN_OK = 0
    VATS2_A_INVALID_RANGE = 1
End Enum

Public Enum VATS2AIntegrityFlags
    VATS2_A_FLAG_NONE = 0
    VATS2_DUPLICATE_A_INVOICE = 1
    VATS2_A_DATA_REVIEW = 2
End Enum

Public Type VATS2ADuplicateGroup
    InvoiceDigits As String
    OccurrenceCount As Long
    AIndexes() As Long
End Type

Public Type VATS2AIntegrityResult
    RecordCount As Long
    DuplicateGroupCount As Long
    DuplicateGroups() As VATS2ADuplicateGroup
    InvalidValueCount As Long
    InvalidAIndexes() As Long
    Status As VATS2AIntegrityStatus
    Flags As Long
End Type

Public Function VATStage2ScanAIntegrity(ByRef aInvoiceDigits() As String, _
    ByVal aCount As Long) As VATS2AIntegrityResult
    Dim result As VATS2AIntegrityResult
    Dim first As Long, last As Long, i As Long, j As Long, g As Long
    Dim byDigits As Object, orderedDigits As Collection, positions As Collection
    Dim digits As String
    result.Status = VATS2_A_INVALID_RANGE
    If aCount < 0 Then GoTo Done
    If aCount = 0 Then
        result.Status = VATS2_A_SCAN_OK
        GoTo Done
    End If
    If Not IntegrityArrayBounds(aInvoiceDigits, first, last) Then GoTo Done
    If CDbl(aCount) > CDbl(last) - CDbl(first) + 1 Then GoTo Done
    last = CLng(CDbl(first) + CDbl(aCount) - 1)
    result.Status = VATS2_A_SCAN_OK
    result.RecordCount = aCount
    Set byDigits = CreateObject("Scripting.Dictionary")
    '二进制完整字符串键，等价于 StrComp(..., ..., vbBinaryCompare)=0；不做尾号匹配。
    byDigits.CompareMode = vbBinaryCompare
    Set orderedDigits = New Collection
    ReDim result.InvalidAIndexes(1 To aCount)
    For i = first To last
        digits = aInvoiceDigits(i)
        If Not IntegrityAsciiDigits(digits) Then
            '非法值只报告原索引，不清洗，不作为发票号码进入重复组。
            result.InvalidValueCount = result.InvalidValueCount + 1
            result.InvalidAIndexes(result.InvalidValueCount) = i
        Else
            If byDigits.Exists(digits) Then
                Set positions = byDigits.Item(digits)
            Else
                Set positions = New Collection
                byDigits.Add digits, positions
                orderedDigits.Add digits
            End If
            positions.Add i
            If positions.Count = 2 Then result.DuplicateGroupCount = result.DuplicateGroupCount + 1
        End If
    Next i
    If result.InvalidValueCount > 0 Then
        result.Flags = result.Flags Or VATS2_A_DATA_REVIEW
        ReDim Preserve result.InvalidAIndexes(1 To result.InvalidValueCount)
    Else
        Erase result.InvalidAIndexes
    End If
    If result.DuplicateGroupCount > 0 Then
        result.Flags = result.Flags Or VATS2_DUPLICATE_A_INVOICE
        ReDim result.DuplicateGroups(1 To result.DuplicateGroupCount)
        '按首次出现顺序生成结果，不依赖字典的枚举顺序或第二次出现顺序。
        For i = 1 To orderedDigits.Count
            digits = orderedDigits(i)
            Set positions = byDigits.Item(digits)
            If positions.Count > 1 Then
                g = g + 1
                With result.DuplicateGroups(g)
                    .InvoiceDigits = digits
                    .OccurrenceCount = positions.Count
                    ReDim .AIndexes(1 To positions.Count)
                    For j = 1 To positions.Count
                        .AIndexes(j) = positions(j)
                    Next j
                End With
            End If
        Next i
    End If
Done:
    VATStage2ScanAIntegrity = result
End Function

Private Function IntegrityAsciiDigits(ByVal digits As String) As Boolean
    Dim i As Long, ch As String
    If Len(digits) = 0 Then Exit Function
    For i = 1 To Len(digits)
        ch = Mid$(digits, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IntegrityAsciiDigits = True
End Function

Private Function IntegrityArrayBounds(ByRef values() As String, ByRef first As Long, ByRef last As Long) As Boolean
    '只处理未分配数组的边界错误，不吞掉扫描过程中的异常。
    On Error GoTo Unallocated
    first = LBound(values): last = UBound(values)
    IntegrityArrayBounds = True
Unallocated:
End Function
