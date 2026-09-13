Attribute VB_Name = "modVATStage2Fingerprint"
Option Explicit
Option Compare Binary

Public Enum VATS2FingerprintStatus
    VATS2_FINGERPRINT_OK = 0
    VATS2_FINGERPRINT_INVALID_INPUT = 1
    VATS2_FINGERPRINT_INVALID_CONTRACT = 2
    VATS2_FINGERPRINT_UNSUPPORTED_VALUE = 3
    VATS2_FINGERPRINT_ENCODING_ERROR = 4
End Enum

Public Type VATS2FingerprintResult
    Status As VATS2FingerprintStatus
    ARecordCount As Long
    BRecordCount As Long
    AFingerprints() As String
    BFingerprints() As String
    ABatchFingerprint As String
    BBatchFingerprint As String
    CombinedFingerprint As String
    ErrorSide As String
    ErrorIndex As Long
    ErrorReason As String
End Type

'LSet仅在固定大小的纯数值UDT间复制，不使用指针、WinAPI或Variant内存布局。
Private Type SingleValue
    value As Single
End Type
Private Type DoubleValue
    value As Double
End Type
Private Type CurrencyValue
    value As Currency
End Type
Private Type Bytes4
    data(0 To 3) As Byte
End Type
Private Type Bytes8
    data(0 To 7) As Byte
End Type

Public Function VATStage2BuildFingerprints(ByRef a As VATS2ASnapshotResult, _
    ByRef b As VATS2BSnapshotResult) As VATS2FingerprintResult
    Dim r As VATS2FingerprintResult, blank As VATS2FingerprintResult
    Dim i As Long, phase As VATS2FingerprintStatus
    r.Status = VATS2_FINGERPRINT_INVALID_INPUT
    If a.Status <> VATS2_SNAPSHOT_OK Then r.ErrorSide = "A": r.ErrorReason = "A Snapshot非OK。": GoTo Done
    If b.Status <> VATS2_SNAPSHOT_OK Then r.ErrorSide = "B": r.ErrorReason = "B Snapshot非OK。": GoTo Done
    On Error GoTo Failed
    phase = VATS2_FINGERPRINT_INVALID_CONTRACT
    r.ErrorSide = "A"
    Need a.RecordCount >= 0, "RecordCount不得为负数。"
    If a.RecordCount = 0 Then
        Need AUnallocated(a), "零记录的A数组必须未分配。"
    Else
        Need LBound(a.Records) = 1 And UBound(a.Records) = a.RecordCount, "A数组必须为1..RecordCount。"
    End If
    For i = 1 To a.RecordCount
        r.ErrorIndex = i
        Need a.Records(i).AIndex = i And a.Records(i).ExcelRow > 0, "A当前索引或行号无效。"
        If i > 1 Then Need a.Records(i).ExcelRow > a.Records(i - 1).ExcelRow, "A当前ExcelRow必须递增。"
    Next i
    r.ErrorSide = "B": r.ErrorIndex = 0
    Need b.RecordCount >= 0, "RecordCount不得为负数。"
    If b.RecordCount = 0 Then
        Need BUnallocated(b), "零记录的B数组必须未分配。"
    Else
        Need LBound(b.Records) = 1 And UBound(b.Records) = b.RecordCount, "B数组必须为1..RecordCount。"
    End If
    For i = 1 To b.RecordCount
        r.ErrorIndex = i
        Need b.Records(i).BIndex = i And b.Records(i).ExcelRow > 0, "B当前索引或行号无效。"
        If i > 1 Then Need b.Records(i).ExcelRow > b.Records(i - 1).ExcelRow, "B当前ExcelRow必须递增。"
    Next i
    phase = VATS2_FINGERPRINT_ENCODING_ERROR
    r.ARecordCount = a.RecordCount: r.BRecordCount = b.RecordCount
    If a.RecordCount > 0 Then ReDim r.AFingerprints(1 To a.RecordCount)
    If b.RecordCount > 0 Then ReDim r.BFingerprints(1 To b.RecordCount)
    r.ErrorSide = "A"
    For i = 1 To a.RecordCount
        r.ErrorIndex = i
        With a.Records(i)
            r.AFingerprints(i) = RecordSignature("A", .InvoiceDigitsRaw, .AmountRaw, .IsCompleted, .InvoiceHasFormula, .AmountHasFormula)
        End With
    Next i
    r.ErrorSide = "B"
    For i = 1 To b.RecordCount
        r.ErrorIndex = i
        With b.Records(i)
            r.BFingerprints(i) = RecordSignature("B", .SupplierTextRaw, .AmountRaw, .IsCompleted, .SupplierHasFormula, .AmountHasFormula)
        End With
    Next i
    r.ErrorSide = "A": r.ErrorIndex = 0
    r.ABatchFingerprint = BatchSignature("A", r.AFingerprints, r.ARecordCount)
    r.ErrorSide = "B"
    r.BBatchFingerprint = BatchSignature("B", r.BFingerprints, r.BRecordCount)
    r.CombinedFingerprint = Pack("VATFP1-COMBINED") & Pack(r.ABatchFingerprint) & Pack(r.BBatchFingerprint)
    r.Status = VATS2_FINGERPRINT_OK: r.ErrorSide = "": r.ErrorIndex = 0
Done:
    VATStage2BuildFingerprints = r: Exit Function
Failed:
    blank.Status = phase
    If Err.Number = vbObjectError + 710 Then blank.Status = VATS2_FINGERPRINT_UNSUPPORTED_VALUE
    blank.ErrorSide = r.ErrorSide: blank.ErrorIndex = r.ErrorIndex: blank.ErrorReason = Err.Description
    r = blank
    Resume Done
End Function

Private Function Pack(ByVal text As String) As String
    '长度单位为VBA Len的UTF-16代码单元；长度本身是十进制整数，不依赖内容分隔符。
    Pack = CStr(Len(text)) & ":" & text
End Function

Private Function RecordSignature(ByVal side As String, ByVal identity As Variant, ByVal amount As Variant, _
    ByVal completed As Boolean, ByVal identityFormula As Boolean, ByVal amountFormula As Boolean) As String
    RecordSignature = Pack("VATFP1-RECORD") & Pack(side) & Pack(EncodeValue(identity)) & Pack(EncodeValue(amount)) & _
        Pack(EncodeValue(completed)) & Pack(EncodeValue(identityFormula)) & Pack(EncodeValue(amountFormula))
End Function

Private Function EncodeValue(ByVal value As Variant) As String
    Dim kind As Long, payload As String, one As SingleValue, two As DoubleValue, money As CurrencyValue
    Dim four As Bytes4, eight As Bytes8, i As Long
    If IsObject(value) Or IsArray(value) Then Err.Raise vbObjectError + 710, , "不支持对象或数组Variant。"
    kind = VarType(value)
    Select Case kind
        Case vbEmpty, vbNull: payload = ""
        Case vbString: payload = value
        Case vbByte, vbInteger, vbLong: payload = CStr(value)
        Case vbBoolean
            If value Then payload = "1" Else payload = "0"
        Case vbSingle
            one.value = value: LSet four = one
            For i = 0 To 3: payload = payload & Right$("0" & Hex$(four.data(i)), 2): Next i
        Case vbDouble, vbDate
            two.value = CDbl(value): LSet eight = two
            For i = 0 To 7: payload = payload & Right$("0" & Hex$(eight.data(i)), 2): Next i
        Case vbCurrency
            money.value = value: LSet eight = money
            For i = 0 To 7: payload = payload & Right$("0" & Hex$(eight.data(i)), 2): Next i
        Case vbDecimal
            'Decimal由VBA直接转完整十进制文本，不经Double或显示格式；环境区域设置须一致。
            payload = CStr(value)
        Case vbError
            'CStr(Error)保留错误编号；不将错误当金额或普通字符串，外层另带VarType。
            payload = CStr(value)
        Case Else: Err.Raise vbObjectError + 710, , "不支持的Variant类型：" & kind
    End Select
    EncodeValue = Pack(CStr(kind)) & Pack(payload)
End Function

Private Function BatchSignature(ByVal side As String, ByRef records() As String, ByVal count As Long) As String
    Dim ordered() As String, work() As String, parts() As String, i As Long
    ReDim parts(0 To count)
    parts(0) = Pack("VATFP1-BATCH") & Pack(side) & Pack(CStr(count))
    If count > 0 Then
        ReDim ordered(1 To count): ReDim work(1 To count)
        For i = 1 To count: ordered(i) = records(i): Next i
        MergeSort ordered, work, 1, count
        For i = 1 To count: parts(i) = Pack(ordered(i)): Next i
    End If
    BatchSignature = Join(parts, "")
End Function

Private Sub MergeSort(ByRef values() As String, ByRef work() As String, ByVal first As Long, ByVal last As Long)
    Dim middle As Long, left As Long, right As Long, at As Long
    If first >= last Then Exit Sub
    middle = first + (last - first) \ 2
    MergeSort values, work, first, middle
    MergeSort values, work, middle + 1, last
    left = first: right = middle + 1
    For at = first To last
        If left > middle Then
            work(at) = values(right): right = right + 1
        ElseIf right > last Then
            work(at) = values(left): left = left + 1
        ElseIf StrComp(values(left), values(right), vbBinaryCompare) <= 0 Then
            work(at) = values(left): left = left + 1
        Else
            work(at) = values(right): right = right + 1
        End If
    Next at
    For at = first To last: values(at) = work(at): Next at
End Sub

Private Sub Need(ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then Err.Raise 5, , reason
End Sub

Private Function AUnallocated(ByRef a As VATS2ASnapshotResult) As Boolean
    Dim bound As Long
    On Error GoTo Missing
    bound = LBound(a.Records): Exit Function
Missing:
    AUnallocated = True
End Function

Private Function BUnallocated(ByRef b As VATS2BSnapshotResult) As Boolean
    Dim bound As Long
    On Error GoTo Missing
    bound = LBound(b.Records): Exit Function
Missing:
    BUnallocated = True
End Function
