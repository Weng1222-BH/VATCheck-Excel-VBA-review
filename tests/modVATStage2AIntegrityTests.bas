Attribute VB_Name = "modVATStage2AIntegrityTests"
Option Explicit
Option Compare Binary

Private caseCount As Long, failureCount As Long
Private details As String

Public Function VATStage2AIntegrity_SelfTest() As String
    On Error GoTo Unexpected
    caseCount = 0: failureCount = 0: details = vbNullString
    CheckCase "无重复", "11111111|22222222|33333333", "", ""
    CheckCase "两个相同号码", "26578450000012345678|26578450000012345678", "26578450000012345678:1,2", ""
    CheckCase "同号三次", "12345678|12345678|12345678", "12345678:1,2,3", ""
    CheckCase "两个重复组", "11111111|22222222|11111111|22222222", "11111111:1,3;22222222:2,4", ""
    CheckCase "不连续重复", "26578450000012345678|11111111111111111111|26578450000012345678", "26578450000012345678:1,3", ""
    CheckCase "前导零不同", "00123456|123456", "", ""
    CheckCase "前导零重复保持", "00123456|123456|00123456", "00123456:1,3", ""
    CheckCase "相同尾号非重复", "12345678|9912345678|8812345678", "", ""
    CheckCase "零下界", "11111111|22222222|11111111", "11111111:0,2", "", 0
    CheckCase "正任意下界", "11111111|22222222|11111111", "11111111:8,10", "", 8
    CheckCase "负下界", "11111111|22222222|11111111", "11111111:-3,-1", "", -3
    CheckCase "aCount 忽略后段重复", "11111111|22222222|11111111", "", "", 1, 2
    CheckCase "aCount 保留范围内重复", "11111111|11111111|22222222|22222222", "11111111:1,2", "", 1, 3
    CheckCase "未分配空数组", "", "", ""
    CheckCase "已分配但 aCount 零", "11111111|11111111", "", "", 1, 0
    CheckCase "负 aCount", "11111111", "", "", 1, -1, VATS2_A_INVALID_RANGE
    CheckCase "aCount 超容量", "11111111", "", "", 1, 2, VATS2_A_INVALID_RANGE
    CheckCase "未分配却声称非空", "", "", "", 1, 1, VATS2_A_INVALID_RANGE
    CheckCase "首次顺序不同于第二次出现顺序", "11111111|22222222|22222222|11111111|11111111", "11111111:1,4,5;22222222:2,3", ""
    CheckCase "单个号码", "11111111", "", ""
    CheckCase "多个空串仅报告非法值", "|", "", "1,2"
    CheckCase "非 ASCII 字符与空白不清洗", "12345678| 12345678|12345678 |１２３４５６７８|BC12345678|123.45", "", "2,3,4,5,6"
    CheckCase "非法值和重复风险共存", "11111111||11111111|x", "11111111:1,3", "2,4"
    CheckCase "重复非法值不当发票重复", "BAD|BAD", "", "1,2"
    CheckCase "非法索引保留负下界", "x|11111111|", "", "-2,0", -2
    CheckCase "有效范围外非法值不检查", "11111111|BAD|", "", "", 1, 1
    CheckCase "长号码字符串不丢精度", "0000123456789012345678901234567890|0000123456789012345678901234567891|0000123456789012345678901234567890", "0000123456789012345678901234567890:1,3", ""
    CheckCase "零与多位零保持不同", "0|00|0|00", "0:1,3;00:2,4", ""
    CheckCase "控制字符不清洗", "12345678" & vbTab & "|12345678" & vbLf, "", "1,2"
    If failureCount = 0 Then
        VATStage2AIntegrity_SelfTest = "PASS: " & caseCount & " cases" & vbCrLf & details
    Else
        VATStage2AIntegrity_SelfTest = "FAIL: " & failureCount & "/" & caseCount & " cases" & vbCrLf & details
    End If
    Exit Function
Unexpected:
    VATStage2AIntegrity_SelfTest = "FAIL: unexpected " & Err.Number & " " & Err.Description & vbCrLf & details
End Function

Private Sub CheckCase(ByVal title As String, ByVal source As String, ByVal groups As String, ByVal invalids As String, _
    Optional ByVal lower As Long = 1, Optional ByVal validCount As Long = -999, _
    Optional ByVal expectedStatus As VATS2AIntegrityStatus = VATS2_A_SCAN_OK)
    Dim values() As String, parts As Variant, result As VATS2AIntegrityResult, again As VATS2AIntegrityResult
    Dim n As Long, i As Long, expectedFlags As Long, expectedRecords As Long
    Dim actualGroups As String, actualInvalids As String, reason As String
    On Error GoTo Failed
    caseCount = caseCount + 1
    If Len(source) > 0 Then
        parts = Split(source, "|"): n = UBound(parts) + 1
        ReDim values(lower To lower + n - 1)
        For i = 0 To n - 1
            values(lower + i) = parts(i)
        Next i
    End If
    If validCount = -999 Then validCount = n
    If expectedStatus = VATS2_A_SCAN_OK Then expectedRecords = validCount
    If Len(groups) > 0 Then expectedFlags = VATS2_DUPLICATE_A_INVOICE
    If Len(invalids) > 0 Then expectedFlags = expectedFlags Or VATS2_A_DATA_REVIEW
    result = VATStage2ScanAIntegrity(values, validCount)
    again = VATStage2ScanAIntegrity(values, validCount)
    Require result.Status = expectedStatus, "Status"
    Require result.RecordCount = expectedRecords, "RecordCount"
    Require result.Flags = expectedFlags, "Flags"
    Require result.DuplicateGroupCount = ItemCount(groups, ";"), "DuplicateGroupCount"
    Require result.InvalidValueCount = ItemCount(invalids, ","), "InvalidValueCount"
    actualGroups = GroupText(result): actualInvalids = InvalidText(result)
    Require actualGroups = groups, "重复组号码、次数、顺序及索引"
    Require actualInvalids = invalids, "非法值索引及顺序"
    '所有用例重复调用，并逐项检查输入不变，不仅比较摘要状态。
    Require again.Status = result.Status, "重复调用 Status"
    Require again.RecordCount = result.RecordCount, "重复调用 RecordCount"
    Require again.Flags = result.Flags, "重复调用 Flags"
    Require again.DuplicateGroupCount = result.DuplicateGroupCount, "重复调用组数"
    Require again.InvalidValueCount = result.InvalidValueCount, "重复调用非法值数"
    Require GroupText(again) = actualGroups, "重复调用所有重复组字段"
    Require InvalidText(again) = actualInvalids, "重复调用非法索引"
    For i = 0 To n - 1
        Require values(lower + i) = parts(i), "输入数组未修改"
    Next i
    details = details & "PASS " & title & vbCrLf
    Exit Sub
Failed:
    reason = Err.Description
    failureCount = failureCount + 1
    details = details & "FAIL " & title & ": " & reason & vbCrLf
End Sub

Private Function GroupText(ByRef result As VATS2AIntegrityResult) As String
    Dim i As Long, j As Long, text As String
    If result.DuplicateGroupCount = 0 Then
        Require Not HasGroupArray(result), "零组数组未分配"
        Exit Function
    End If
    Require LBound(result.DuplicateGroups) = 1, "组数组下界"
    Require UBound(result.DuplicateGroups) = result.DuplicateGroupCount, "组数组容量"
    For i = 1 To result.DuplicateGroupCount
        If i > 1 Then text = text & ";"
        With result.DuplicateGroups(i)
            Require .OccurrenceCount >= 2, "重复组至少两次"
            Require LBound(.AIndexes) = 1, "索引数组下界"
            Require UBound(.AIndexes) = .OccurrenceCount, "OccurrenceCount 与索引数一致"
            text = text & .InvoiceDigits & ":"
            For j = 1 To .OccurrenceCount
                If j > 1 Then text = text & ","
                text = text & CStr(.AIndexes(j))
            Next j
        End With
    Next i
    GroupText = text
End Function

Private Function InvalidText(ByRef result As VATS2AIntegrityResult) As String
    Dim i As Long, text As String
    If result.InvalidValueCount = 0 Then
        Require Not HasInvalidArray(result), "零非法值数组未分配"
        Exit Function
    End If
    Require LBound(result.InvalidAIndexes) = 1, "非法索引下界"
    Require UBound(result.InvalidAIndexes) = result.InvalidValueCount, "非法索引容量"
    For i = 1 To result.InvalidValueCount
        If i > 1 Then text = text & ","
        text = text & CStr(result.InvalidAIndexes(i))
    Next i
    InvalidText = text
End Function

Private Function ItemCount(ByVal text As String, ByVal separator As String) As Long
    If Len(text) > 0 Then ItemCount = UBound(Split(text, separator)) + 1
End Function

Private Function HasGroupArray(ByRef result As VATS2AIntegrityResult) As Boolean
    Dim upper As Long
    On Error GoTo EmptyArray
    upper = UBound(result.DuplicateGroups): HasGroupArray = True
EmptyArray:
End Function

Private Function HasInvalidArray(ByRef result As VATS2AIntegrityResult) As Boolean
    Dim upper As Long
    On Error GoTo EmptyArray
    upper = UBound(result.InvalidAIndexes): HasInvalidArray = True
EmptyArray:
End Function

Private Sub Require(ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then Err.Raise vbObjectError + 2201, "AIntegrityTests", reason
End Sub
