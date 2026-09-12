Attribute VB_Name = "modVATStage2Amount"
Option Explicit

Public Enum VATS2AmountState
    VATS2_AMOUNT_NOT_COMPARED = 0
    VATS2_AMOUNT_EQUAL = 1
    VATS2_AMOUNT_MISMATCH = 2
End Enum

Public Enum VATS2AmountStatus
    VATS2_AMOUNT_OK = 0
    VATS2_AMOUNT_EMPTY_GROUP = 1
    VATS2_AMOUNT_INVALID_RANGE = 2
    VATS2_AMOUNT_INVALID_AMOUNT = 3
    VATS2_AMOUNT_ARITHMETIC_ERROR = 4
End Enum

Public Type VATS2AmountResult
    ACount As Long
    AIndexes() As Long
    AAmounts() As Variant
    AAmountSum As Variant
    BAmount As Variant
    Difference As Variant
    AmountState As VATS2AmountState
    Status As VATS2AmountStatus
    ErrorSide As String
    ErrorIndex As Long
    ErrorAIndex As Long
    ErrorReason As String
End Type

Public Function VATStage2CompareAmounts(ByRef aIndexes() As Long, ByRef aAmounts() As Variant, _
    ByVal aCount As Long, ByVal bAmount As Variant) As VATS2AmountResult
    Dim result As VATS2AmountResult, indexFirst As Long, amountFirst As Long, offset As Long
    Dim total As Variant, converted As Variant, convertedB As Variant, difference As Variant
    result.Status = VATS2_AMOUNT_INVALID_RANGE
    If aCount < 0 Then GoTo Done
    If aCount = 0 Then result.Status = VATS2_AMOUNT_EMPTY_GROUP: GoTo Done
    If Not AmountInputRange(aIndexes, aAmounts, aCount, indexFirst, amountFirst) Then GoTo Done
    ReDim result.AIndexes(1 To aCount)
    ReDim result.AAmounts(1 To aCount)
    total = CDec(0)
    For offset = 0 To aCount - 1
        result.ErrorSide = "A"
        result.ErrorIndex = amountFirst + offset
        result.ErrorAIndex = aIndexes(indexFirst + offset)
        If Not AmountTryDecimal(aAmounts(amountFirst + offset), converted, result.ErrorReason) Then GoTo InvalidAmount
        result.AIndexes(offset + 1) = aIndexes(indexFirst + offset)
        result.AAmounts(offset + 1) = converted
        '与 Stage 1 一致：每项转 Decimal，Decimal 累加，不经 Double/Currency 合计。
        On Error GoTo ArithmeticError
        total = CDec(total) + CDec(converted)
        On Error GoTo 0
    Next offset
    result.ErrorSide = "B": result.ErrorIndex = 0: result.ErrorAIndex = 0
    If Not AmountTryDecimal(bAmount, convertedB, result.ErrorReason) Then GoTo InvalidAmount
    result.ErrorSide = "DIFFERENCE"
    On Error GoTo ArithmeticError
    difference = CDec(total) - CDec(convertedB)
    If difference = CDec(0) Then
        result.AmountState = VATS2_AMOUNT_EQUAL
    Else
        result.AmountState = VATS2_AMOUNT_MISMATCH
    End If
    On Error GoTo 0
    result.ACount = aCount
    result.AAmountSum = total: result.BAmount = convertedB: result.Difference = difference
    result.Status = VATS2_AMOUNT_OK
    result.ErrorSide = vbNullString: result.ErrorIndex = 0: result.ErrorAIndex = 0
    GoTo Done
InvalidAmount:
    result.Status = VATS2_AMOUNT_INVALID_AMOUNT
    GoTo ClearAmounts
ArithmeticError:
    result.Status = VATS2_AMOUNT_ARITHMETIC_ERROR
    result.ErrorReason = "Decimal 运算失败：" & Err.Number & " " & Err.Description
ClearAmounts:
    '失败不发布部分金额或默认零差额，避免被误用为有效比较结果。
    result.AmountState = VATS2_AMOUNT_NOT_COMPARED
    Erase result.AIndexes: Erase result.AAmounts
Done:
    VATStage2CompareAmounts = result
End Function

Private Function AmountTryDecimal(ByVal value As Variant, ByRef amount As Variant, ByRef reason As String) As Boolean
    On Error GoTo Failed
    '沿用 Stage 1 VATTryAmount 的类型白名单；数字文本也拒绝。
    Select Case VarType(value)
        Case vbEmpty
            reason = "金额为空"
        Case vbError
            reason = "金额为 Excel 错误值"
        Case vbString
            If Len(value) = 0 Then
                reason = "金额为空（包括公式返回空字符串）"
            Else
                reason = "金额为文本（数字文本也不自动转换）"
            End If
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            amount = CDec(value)
            AmountTryDecimal = True
        Case Else
            reason = "金额不是可用数值"
    End Select
    Exit Function
Failed:
    reason = "金额超出 Decimal 可处理范围或无法转换"
End Function

Private Function AmountInputRange(ByRef indexes() As Long, ByRef amounts() As Variant, ByVal count As Long, _
    ByRef indexFirst As Long, ByRef amountFirst As Long) As Boolean
    On Error GoTo InvalidRange
    indexFirst = LBound(indexes): amountFirst = LBound(amounts)
    '这里只计算数组容量；金额不经过 Double。
    If CDbl(count) > CDbl(UBound(indexes)) - CDbl(indexFirst) + 1 Then Exit Function
    If CDbl(count) > CDbl(UBound(amounts)) - CDbl(amountFirst) + 1 Then Exit Function
    AmountInputRange = True
InvalidRange:
End Function
