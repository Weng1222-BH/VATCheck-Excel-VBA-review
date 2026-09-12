Attribute VB_Name = "modVATStage2AmountTests"
Option Explicit

Private caseCount As Long, failureCount As Long, details As String

Public Function VATStage2Amount_SelfTest() As String
    Dim maximum As Variant, tiny As Variant
    On Error GoTo Unexpected
    caseCount = 0: failureCount = 0: details = vbNullString
    maximum = CDec("79228162514264337593543950335")
    tiny = CDec(1) / CDec("10000000000000000000000000000")
    CheckCase "整数相等", Array(100), 100, CDec(100), CDec(0)
    CheckCase "小数相等", Array(12.34), 12.34, CDec(12.34), CDec(0)
    CheckCase "A 大于 B 正差额", Array(100), 98, CDec(100), CDec(2)
    CheckCase "A 小于 B 负差额", Array(98), 100, CDec(98), CDec(-2)
    CheckCase "多 A 相等", Array(10, 20, 30), 60, CDec(60), CDec(0)
    CheckCase "多 A 不等", Array(10, 20, 30), 59, CDec(60), CDec(1)
    CheckCase "0.1加0.2精确等于0.3", Array(0.1, 0.2), 0.3, CDec(0.3), CDec(0)
    CheckCase "大 Decimal 金额", Array(maximum), maximum, maximum, CDec(0)
    CheckCase "大数差一不丢失", Array(maximum), maximum - CDec(1), maximum, CDec(1)
    CheckCase "最小 Decimal 小数", Array(tiny), tiny, tiny, CDec(0)
    CheckCase "最小非零差额不容差", Array(tiny), CDec(0), tiny, tiny
    CheckCase "分以下差额不四舍五入", Array(CDec(1) / CDec(1000)), 0, CDec(1) / CDec(1000), CDec(1) / CDec(1000)
    CheckCase "不同小数位精确相等", Array(CDec("1.2300")), CDec("1.23"), CDec("1.23"), CDec(0)
    CheckCase "负数与零", Array(-10, 0, 5), -5, CDec(-5), CDec(0)
    CheckCase "合法零金额", Array(0), 0, CDec(0), CDec(0)
    CheckCase "保留顺序和重复 AIndex", Array(8, 3, 2, 7), 20, CDec(20), CDec(0)
    CheckCase "两个数组不同下界", Array(1, 2, 3), 6, CDec(6), CDec(0), VATS2_AMOUNT_OK, "", 0, -3, 8
    CheckCase "金额负下界索引零", Array(1, 2), 3, CDec(3), CDec(0), VATS2_AMOUNT_OK, "", 0, 0, -4
    CheckCase "只覆盖前部", Array(1, 2, "忽略范围外"), 3, CDec(3), CDec(0), VATS2_AMOUNT_OK, "", 0, 1, 1, 2
    CheckCase "空组不比较", Array(), 0, Empty, Empty, VATS2_AMOUNT_EMPTY_GROUP
    CheckCase "已分配但零组", Array(1), 1, Empty, Empty, VATS2_AMOUNT_EMPTY_GROUP, "", 0, 1, 1, 0
    CheckCase "负 aCount", Array(1), 1, Empty, Empty, VATS2_AMOUNT_INVALID_RANGE, "", 0, 1, 1, -1
    CheckCase "索引容量不足", Array(1, 2), 3, Empty, Empty, VATS2_AMOUNT_INVALID_RANGE, "", 0, 1, 1, 2, 1
    CheckCase "金额容量不足", Array(1, 2), 3, Empty, Empty, VATS2_AMOUNT_INVALID_RANGE, "", 0, 1, 1, 2, 2, 1
    CheckCase "索引数组未分配", Array(1), 1, Empty, Empty, VATS2_AMOUNT_INVALID_RANGE, "", 0, 1, 1, 1, 0
    CheckCase "金额数组未分配", Array(1), 1, Empty, Empty, VATS2_AMOUNT_INVALID_RANGE, "", 0, 1, 1, 1, 1, 0
    CheckCase "A 非数字文本", Array("abc"), 0, Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "A", 1
    CheckCase "B 非数字文本", Array(1), "abc", Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "B"
    CheckCase "A 数字文本拒绝", Array("1.23"), 1.23, Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "A", 1
    CheckCase "B 数字文本拒绝", Array(1), "1", Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "B"
    CheckCase "A Boolean", Array(True), 1, Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "A", 1
    CheckCase "B Boolean", Array(0), False, Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "B"
    CheckCase "A Excel Error", Array(CVErr(2042)), 1, Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "A", 1
    CheckCase "B Excel Error", Array(1), CVErr(2007), Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "B"
    CheckCase "A Empty", Array(Empty), 0, Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "A", 1
    CheckCase "B Empty", Array(0), Empty, Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "B"
    CheckCase "A Null", Array(Null), 0, Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "A", 1
    CheckCase "B Null", Array(0), Null, Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "B"
    CheckCase "空字符串不是零", Array(""), 0, Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "A", 1
    CheckCase "中间非法准确定位", Array(1, "非法", 2), 3, Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "A", 9, -2, 8
    CheckCase "中间非法负下标定位", Array(1, Null, 2), 3, Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "A", -3, 6, -4
    CheckCase "A 日期类型拒绝", Array(DateSerial(2026, 9, 11)), 0, Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "A", 1
    CheckCase "B 日期类型拒绝", Array(0), DateSerial(2026, 9, 11), Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "B"
    CheckCase "A CDec 转换溢出", Array(1E+29), 0, Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "A", 1
    CheckCase "B CDec 转换溢出", Array(0), 1E+29, Empty, Empty, VATS2_AMOUNT_INVALID_AMOUNT, "B"
    CheckCase "Decimal 累加溢出", Array(maximum, CDec(1)), 0, Empty, Empty, VATS2_AMOUNT_ARITHMETIC_ERROR, "A", 2
    CheckCase "Decimal 差额溢出", Array(maximum), CDec(-1), Empty, Empty, VATS2_AMOUNT_ARITHMETIC_ERROR, "DIFFERENCE"
    CheckCase "所有允许数值子类型转 Decimal", Array(CByte(1), CInt(2), CLng(3), CSng(4), CDbl(5), CCur(6), CDec(7)), 28, CDec(28), CDec(0)
    If failureCount = 0 Then
        VATStage2Amount_SelfTest = "PASS: " & caseCount & " cases" & vbCrLf & details
    Else
        VATStage2Amount_SelfTest = "FAIL: " & failureCount & "/" & caseCount & " cases" & vbCrLf & details
    End If
    Exit Function
Unexpected:
    VATStage2Amount_SelfTest = "FAIL: unexpected " & Err.Number & " " & Err.Description & vbCrLf & details
End Function

Private Sub CheckCase(ByVal title As String, ByVal fixture As Variant, ByVal b As Variant, _
    ByVal expectedSum As Variant, ByVal expectedDiff As Variant, Optional ByVal status As VATS2AmountStatus = VATS2_AMOUNT_OK, _
    Optional ByVal side As String = "", Optional ByVal errorIndex As Long = 0, _
    Optional ByVal indexLower As Long = 1, Optional ByVal amountLower As Long = 1, _
    Optional ByVal count As Long = -999, Optional ByVal indexCapacity As Long = -1, Optional ByVal amountCapacity As Long = -1)
    Dim indexes() As Long, amounts() As Variant, result As VATS2AmountResult, again As VATS2AmountResult
    Dim n As Long, i As Long, before As String, reason As String
    On Error GoTo Failed
    caseCount = caseCount + 1
    n = UBound(fixture) - LBound(fixture) + 1
    If count = -999 Then count = n
    If indexCapacity = -1 Then indexCapacity = n
    If amountCapacity = -1 Then amountCapacity = n
    If indexCapacity > 0 Then
        ReDim indexes(indexLower To indexLower + indexCapacity - 1)
        For i = 0 To indexCapacity - 1
            Select Case i Mod 3
                Case 0: indexes(indexLower + i) = 0
                Case 1: indexes(indexLower + i) = -7
                Case 2: indexes(indexLower + i) = 42
            End Select
        Next i
    End If
    If amountCapacity > 0 Then
        ReDim amounts(amountLower To amountLower + amountCapacity - 1)
        For i = 0 To amountCapacity - 1
            amounts(amountLower + i) = fixture(LBound(fixture) + i)
        Next i
    End If
    before = InputText(indexes, amounts, indexCapacity, amountCapacity, indexLower, amountLower, b)
    result = VATStage2CompareAmounts(indexes, amounts, count, b)
    again = VATStage2CompareAmounts(indexes, amounts, count, b)
    VerifyResult result, indexes, amounts, count, b, expectedSum, expectedDiff, status, side, errorIndex, indexLower, amountLower
    VerifyResult again, indexes, amounts, count, b, expectedSum, expectedDiff, status, side, errorIndex, indexLower, amountLower
    Require result.ErrorReason = again.ErrorReason, "重复调用错误原因一致"
    Require InputText(indexes, amounts, indexCapacity, amountCapacity, indexLower, amountLower, b) = before, "输入数组及 B 值类型和值未修改"
    details = details & "PASS " & title & vbCrLf
    Exit Sub
Failed:
    reason = Err.Description: failureCount = failureCount + 1
    details = details & "FAIL " & title & ": " & reason & vbCrLf
End Sub

Private Sub VerifyResult(ByRef result As VATS2AmountResult, ByRef indexes() As Long, ByRef amounts() As Variant, _
    ByVal count As Long, ByVal b As Variant, ByVal expectedSum As Variant, ByVal expectedDiff As Variant, _
    ByVal status As VATS2AmountStatus, ByVal side As String, ByVal errorIndex As Long, ByVal il As Long, ByVal al As Long)
    Dim i As Long, expectedState As VATS2AmountState
    Require result.Status = status, "Status"
    Require result.ErrorSide = side, "错误来源"
    Require result.ErrorIndex = errorIndex, "错误金额数组下标"
    If side = "A" Then
        Require result.ErrorAIndex = indexes(il + errorIndex - al), "错误原 AIndex"
    Else
        Require result.ErrorAIndex = 0, "非 A 错误无 AIndex"
    End If
    If status = VATS2_AMOUNT_OK Then
        Require result.ACount = count, "ACount"
        Require LBound(result.AIndexes) = 1 And UBound(result.AIndexes) = count, "输出索引数组容量"
        Require LBound(result.AAmounts) = 1 And UBound(result.AAmounts) = count, "输出金额数组容量"
        For i = 1 To count
            Require result.AIndexes(i) = indexes(il + i - 1), "AIndex 与顺序保持"
            Require VarType(result.AAmounts(i)) = vbDecimal, "逐项 Decimal 子类型"
            Require result.AAmounts(i) = CDec(amounts(al + i - 1)), "对应金额保持"
        Next i
        Require VarType(result.AAmountSum) = vbDecimal And VarType(result.BAmount) = vbDecimal And VarType(result.Difference) = vbDecimal, "比较结果全部 Decimal"
        Require result.AAmountSum = expectedSum, "精确合计"
        Require result.BAmount = CDec(b), "精确 B 金额"
        Require result.Difference = expectedDiff, "精确差额与方向"
        If expectedDiff = CDec(0) Then expectedState = VATS2_AMOUNT_EQUAL Else expectedState = VATS2_AMOUNT_MISMATCH
        Require result.AmountState = expectedState, "AmountState"
        Require Len(result.ErrorReason) = 0, "成功无残留错误"
    Else
        Require result.ACount = 0, "错误不发布部分数量"
        Require result.AmountState = VATS2_AMOUNT_NOT_COMPARED, "错误不得比较"
        Require IsEmpty(result.AAmountSum) And IsEmpty(result.BAmount) And IsEmpty(result.Difference), "错误无默认零或部分合计"
        Require NoArrays(result), "错误不发布部分数组"
        If status = VATS2_AMOUNT_INVALID_AMOUNT Or status = VATS2_AMOUNT_ARITHMETIC_ERROR Then Require Len(result.ErrorReason) > 0, "错误原因保留"
    End If
End Sub

Private Function InputText(ByRef indexes() As Long, ByRef amounts() As Variant, ByVal ic As Long, ByVal ac As Long, _
    ByVal il As Long, ByVal al As Long, ByVal b As Variant) As String
    Dim i As Long, text As String
    For i = 0 To ic - 1
        text = text & indexes(il + i) & ";"
    Next i
    For i = 0 To ac - 1
        text = text & ValueText(amounts(al + i)) & ";"
    Next i
    InputText = text & ValueText(b)
End Function

Private Function ValueText(ByVal value As Variant) As String
    Dim text As String
    If IsNull(value) Then text = "Null" Else text = CStr(value)
    ValueText = VarType(value) & ":" & Len(text) & ":" & text
End Function

Private Function NoArrays(ByRef result As VATS2AmountResult) As Boolean
    Dim upper As Long, indexError As Long, amountError As Long
    On Error Resume Next
    upper = UBound(result.AIndexes): indexError = Err.Number: Err.Clear
    upper = UBound(result.AAmounts): amountError = Err.Number: Err.Clear
    On Error GoTo 0
    NoArrays = (indexError <> 0 And amountError <> 0)
End Function

Private Sub Require(ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then Err.Raise vbObjectError + 2501, "AmountTests", reason
End Sub
