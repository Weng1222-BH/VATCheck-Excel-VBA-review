Attribute VB_Name = "modVATStage2FingerprintTests"
Option Explicit
Option Compare Binary

Private passed As Long, failed As Long, log As String
Private a As VATS2ASnapshotResult, b As VATS2BSnapshotResult

Public Function VATStage2Fingerprint_SelfTest() As String
    On Error GoTo Unexpected
    passed = 0: failed = 0: log = ""
    TestPositions
    TestContent
    TestVariants
    TestCanonicalVersion
    TestBatches
    TestContracts
    GoTo Done
Unexpected:
    failed = failed + 1: log = log & "UNEXPECTED " & Err.Number & " " & Err.Description & vbCrLf
Done:
    If failed = 0 Then
        VATStage2Fingerprint_SelfTest = "PASS: " & passed & " assertions" & vbCrLf & log
    Else
        VATStage2Fingerprint_SelfTest = "FAIL: " & failed & "; PASS: " & passed & vbCrLf & log
    End If
End Function

Private Sub Fixture(Optional ByVal count As Long = 1)
    Dim aa As VATS2ASnapshotResult, bb As VATS2BSnapshotResult, i As Long
    a = aa: b = bb: a.RecordCount = count: b.RecordCount = count
    a.HeaderRow = 3: b.HeaderRow = 5: a.SheetName = "合成A": b.SheetName = "合成B"
    If count > 0 Then ReDim a.Records(1 To count): ReDim b.Records(1 To count)
    For i = 1 To count
        With a.Records(i)
            .AIndex = i: .ExcelRow = 10 + i: .InvoiceDigitsRaw = "0000" & i: .AmountRaw = CDec(i)
            .InvoiceCellAddress = "D" & .ExcelRow: .AmountCellAddress = "I" & .ExcelRow: .IsCompleted = True
        End With
        With b.Records(i)
            .BIndex = i: .ExcelRow = 20 + i: .SupplierTextRaw = "NO.0000" & i: .AmountRaw = CDec(i)
            .SupplierCellAddress = "B" & .ExcelRow: .AmountCellAddress = "F" & .ExcelRow: .IsCompleted = True
        End With
    Next i
End Sub

Private Sub TestPositions()
    Dim r As VATS2FingerprintResult, s As VATS2FingerprintResult
    Dim ar As VATS2ASnapshotRecord, br As VATS2BSnapshotRecord
    Fixture 2: r = RunFingerprint()
    a.Records(1).ExcelRow = 101: a.Records(2).ExcelRow = 150: a.Records(1).InvoiceCellAddress = "K101": a.Records(1).AmountCellAddress = "Z101"
    b.Records(1).ExcelRow = 201: b.Records(2).ExcelRow = 230: b.Records(1).SupplierCellAddress = "R201": b.Records(1).AmountCellAddress = "T201"
    a.HeaderRow = 80: b.HeaderRow = 90: a.SheetName = "移动A": b.SheetName = "移动B"
    s = RunFingerprint()
    Check "行号地址表名变化不影响任何指纹", SameResult(r, s)
    ar = a.Records(1): a.Records(1) = a.Records(2): a.Records(2) = ar
    br = b.Records(1): b.Records(1) = b.Records(2): b.Records(2) = br
    a.Records(1).AIndex = 1: a.Records(2).AIndex = 2: a.Records(1).ExcelRow = 101: a.Records(2).ExcelRow = 150
    b.Records(1).BIndex = 1: b.Records(2).BIndex = 2: b.Records(1).ExcelRow = 201: b.Records(2).ExcelRow = 230
    s = RunFingerprint()
    Check "换序后A原顺序输出重新映射", s.AFingerprints(1) = r.AFingerprints(2) And s.AFingerprints(2) = r.AFingerprints(1)
    Check "换序后B原顺序输出重新映射", s.BFingerprints(1) = r.BFingerprints(2) And s.BFingerprints(2) = r.BFingerprints(1)
    Check "排序只变当前Index不变两侧Batch及Combined", s.ABatchFingerprint = r.ABatchFingerprint And s.BBatchFingerprint = r.BBatchFingerprint And s.CombinedFingerprint = r.CombinedFingerprint
End Sub

Private Sub TestContent()
    Dim r As VATS2FingerprintResult, s As VATS2FingerprintResult, i As Long
    For i = 1 To 14
        Call Fixture: r = RunFingerprint()
        Select Case i
            Case 1: a.Records(1).InvoiceDigitsRaw = "00002"
            Case 2: b.Records(1).SupplierTextRaw = "NO.00002"
            Case 3: a.Records(1).AmountRaw = CDec(2)
            Case 4: b.Records(1).AmountRaw = CDec(2)
            Case 5: a.Records(1).IsCompleted = False
            Case 6: b.Records(1).IsCompleted = False
            Case 7: a.Records(1).InvoiceHasFormula = True
            Case 8: b.Records(1).SupplierHasFormula = True
            Case 9: a.Records(1).AmountHasFormula = True
            Case 10: b.Records(1).AmountHasFormula = True
            Case 11: a.Records(1).InvoiceDigitsRaw = "1"
            Case 12: b.Records(1).SupplierTextRaw = " NO.00001"
            Case 13: b.Records(1).SupplierTextRaw = "no.00001"
            Case 14: a.Records(1).InvoiceDigitsRaw = "00001 "
        End Select
        s = RunFingerprint()
        Check "业务值类型或标志变化必须改变指纹" & i, s.CombinedFingerprint <> r.CombinedFingerprint
    Next i
    Call Fixture
    a.Records(1).InvoiceDigitsRaw = "A|BC": a.Records(1).AmountRaw = "D": r = RunFingerprint()
    a.Records(1).InvoiceDigitsRaw = "A": a.Records(1).AmountRaw = "BC|D": s = RunFingerprint()
    Check "长度前缀拒绝简单分隔碰撞", r.AFingerprints(1) <> s.AFingerprints(1)
    a.Records(1).InvoiceDigitsRaw = "中文:" & ChrW(0) & "3:AB" & vbCrLf & ChrW(&HD83D) & ChrW(&HDE00)
    r = RunFingerprint(): a.Records(1).InvoiceDigitsRaw = Replace(a.Records(1).InvoiceDigitsRaw, ChrW(0), "")
    s = RunFingerprint(): Check "保留Unicode空字符换行及长度语义", r.AFingerprints(1) <> s.AFingerprints(1)
End Sub

Private Sub TestVariants()
    Dim values(0 To 14) As Variant, signatures(0 To 14) As String, r As VATS2FingerprintResult
    Dim i As Long, j As Long, prior As String
    values(0) = Empty: values(1) = "": values(2) = CByte(1): values(3) = CInt(1): values(4) = CLng(1)
    values(5) = CSng(1): values(6) = CDbl(1): values(7) = CCur(1): values(8) = CDec(1)
    values(9) = True: values(10) = CVErr(2042): values(11) = CDate(1): values(12) = Null
    values(13) = "1": values(14) = CVErr(2007)
    Call Fixture
    For i = 0 To 14
        a.Records(1).AmountRaw = values(i): b.Records(1).SupplierTextRaw = values(i)
        r = RunFingerprint(): signatures(i) = r.AFingerprints(1)
        For j = 0 To i - 1: Check "Variant类型/值不混淆" & i & "/" & j, signatures(i) <> signatures(j): Next j
    Next i
    a.Records(1).AmountRaw = CDbl(1): r = RunFingerprint(): prior = r.AFingerprints(1)
    a.Records(1).AmountRaw = CDbl(1) + 2 ^ (-52): r = RunFingerprint()
    Check "Double最低有效位变化不被15位文本吞掉", prior <> r.AFingerprints(1)
    a.Records(1).AmountRaw = CSng(1): r = RunFingerprint(): prior = r.AFingerprints(1)
    a.Records(1).AmountRaw = CSng(1 + 2 ^ (-23)): r = RunFingerprint()
    Check "Single最低有效位变化可识别", prior <> r.AFingerprints(1)
    a.Records(1).AmountRaw = CDec("79228162514264337593543950335"): r = RunFingerprint(): prior = r.AFingerprints(1)
    a.Records(1).AmountRaw = CDec("79228162514264337593543950334"): r = RunFingerprint()
    Check "Decimal最大整数相差1仍区分", prior <> r.AFingerprints(1)
    a.Records(1).AmountRaw = CDec("0.0000000000000000000000000001"): r = RunFingerprint(): prior = r.AFingerprints(1)
    a.Records(1).AmountRaw = CDec("0.0000000000000000000000000002"): r = RunFingerprint()
    Check "Decimal第28位变化可识别", prior <> r.AFingerprints(1)
    a.Records(1).AmountRaw = CCur("922337203685477.5807"): r = RunFingerprint(): prior = r.AFingerprints(1)
    a.Records(1).AmountRaw = CCur("922337203685477.5806"): r = RunFingerprint()
    Check "Currency极值末位变化可识别", prior <> r.AFingerprints(1)
    a.Records(1).AmountRaw = CDate(0.5): r = RunFingerprint(): prior = r.AFingerprints(1)
    a.Records(1).AmountRaw = CDate(0.5 + 2 ^ (-53)): r = RunFingerprint()
    Check "Date不以显示秒数截断精度", prior <> r.AFingerprints(1)
    a.Records(1).AmountRaw = CDbl(-1E+200): r = RunFingerprint(): prior = r.AFingerprints(1)
    a.Records(1).AmountRaw = CDbl(-1E+200) * 2: r = RunFingerprint()
    Check "超财务范围Double仍可编码不修复", prior <> r.AFingerprints(1)
End Sub

Private Sub TestCanonicalVersion()
    Dim r As VATS2FingerprintResult, expected As String
    Call Fixture
    a.Records(1).InvoiceDigitsRaw = Empty: a.Records(1).AmountRaw = Empty: a.Records(1).IsCompleted = False
    r = RunFingerprint()
    expected = "13:VATFP1-RECORD1:A5:1:00:5:1:00:7:2:111:07:2:111:07:2:111:0"
    Check "VATFP1长度前缀协议固定样本", r.AFingerprints(1) = expected
    Check "Batch显式带版本侧别数量及完整记录长度", r.ABatchFingerprint = "12:VATFP1-BATCH1:A1:1" & CStr(Len(expected)) & ":" & expected
    a.Records(1).AmountRaw = CDbl(1): r = RunFingerprint()
    Check "Double值按Windows小端字节编码", InStr(1, r.AFingerprints(1), "16:000000000000F03F", vbBinaryCompare) > 0
    a.Records(1).AmountRaw = CCur(1): r = RunFingerprint()
    Check "Currency四位定点整数原样编码", InStr(1, r.AFingerprints(1), "16:1027000000000000", vbBinaryCompare) > 0
End Sub

Private Sub TestBatches()
    Dim r As VATS2FingerprintResult, s As VATS2FingerprintResult, i As Long, ar As VATS2ASnapshotRecord, br As VATS2BSnapshotRecord
    Fixture 2
    a.Records(2) = a.Records(1): a.Records(2).AIndex = 2: a.Records(2).ExcelRow = 12
    b.Records(2) = b.Records(1): b.Records(2).BIndex = 2: b.Records(2).ExcelRow = 22
    r = RunFingerprint()
    Check "重复记录保持相同record指纹", r.AFingerprints(1) = r.AFingerprints(2) And r.BFingerprints(1) = r.BFingerprints(2)
    a.RecordCount = 1: ReDim Preserve a.Records(1 To 1): b.RecordCount = 1: ReDim Preserve b.Records(1 To 1)
    s = RunFingerprint()
    Check "multiplicity不同两侧Batch均不同", r.ABatchFingerprint <> s.ABatchFingerprint And r.BBatchFingerprint <> s.BBatchFingerprint
    Call Fixture: r = RunFingerprint()
    a.Records(1).InvoiceDigitsRaw = "NO.00001": b.Records(1).SupplierTextRaw = "00001": s = RunFingerprint()
    Check "交换A/B内容Combined不相同", r.CombinedFingerprint <> s.CombinedFingerprint
    b.Records(1).SupplierTextRaw = a.Records(1).InvoiceDigitsRaw: s = RunFingerprint()
    Check "同内容也区分A/B record和Batch", s.AFingerprints(1) <> s.BFingerprints(1) And s.ABatchFingerprint <> s.BBatchFingerprint
    Fixture 0: r = RunFingerprint()
    Check "零记录有稳定且区分两侧的签名", r.ARecordCount = 0 And r.BRecordCount = 0 And Len(r.CombinedFingerprint) > 0 And r.ABatchFingerprint <> r.BBatchFingerprint
    Fixture 257: r = RunFingerprint()
    For i = 1 To 128
        ar = a.Records(i): a.Records(i) = a.Records(258 - i): a.Records(258 - i) = ar
        br = b.Records(i): b.Records(i) = b.Records(258 - i): b.Records(258 - i) = br
    Next i
    For i = 1 To 257
        a.Records(i).AIndex = i: a.Records(i).ExcelRow = i + 10
        b.Records(i).BIndex = i: b.Records(i).ExcelRow = i + 20
    Next i
    s = RunFingerprint(): Check "257条逆序批次相同且原数组顺序保留", r.CombinedFingerprint = s.CombinedFingerprint And s.AFingerprints(1) = r.AFingerprints(257)
End Sub

Private Sub TestContracts()
    Dim i As Long, r As VATS2FingerprintResult
    For i = 1 To 14
        Fixture 2
        Select Case i
            Case 1: a.Status = VATS2_SNAPSHOT_READ_ERROR
            Case 2: b.Status = VATS2_SNAPSHOT_READ_ERROR
            Case 3: a.RecordCount = -1
            Case 4: b.RecordCount = -1
            Case 5: Erase a.Records
            Case 6: Erase b.Records
            Case 7: ReDim a.Records(0 To 1)
            Case 8: ReDim b.Records(0 To 1)
            Case 9: a.Records(2).AIndex = 1
            Case 10: b.Records(2).BIndex = 1
            Case 11: a.Records(2).ExcelRow = a.Records(1).ExcelRow
            Case 12: b.Records(2).ExcelRow = 1
            Case 13: a.RecordCount = 0
            Case 14: b.RecordCount = 0
        End Select
        r = VATStage2BuildFingerprints(a, b)
        Check "契约拒绝且不发布部分Batch" & i, r.Status <> VATS2_FINGERPRINT_OK And Len(r.CombinedFingerprint) = 0 And Len(r.ABatchFingerprint) = 0 And Len(r.BBatchFingerprint) = 0 And r.ARecordCount = 0 And r.BRecordCount = 0 And Len(r.ErrorReason) > 0
    Next i
    Fixture 2: b.Records(2).AmountRaw = Array(1, 2): r = VATStage2BuildFingerprints(a, b)
    Check "不支持数组Variant明确定位且清空已编码A", r.Status = VATS2_FINGERPRINT_UNSUPPORTED_VALUE And r.ErrorSide = "B" And r.ErrorIndex = 2 And r.ARecordCount = 0 And Len(r.ABatchFingerprint) = 0
    Set b.Records(2).AmountRaw = New Collection: r = VATStage2BuildFingerprints(a, b)
    Check "不读取对象默认属性而明确拒绝", r.Status = VATS2_FINGERPRINT_UNSUPPORTED_VALUE And r.ErrorSide = "B" And r.ErrorIndex = 2
End Sub

Private Function RunFingerprint() As VATS2FingerprintResult
    Dim r As VATS2FingerprintResult, again As VATS2FingerprintResult, aa As VATS2ASnapshotResult, bb As VATS2BSnapshotResult
    aa = a: bb = b
    r = VATStage2BuildFingerprints(a, b)
    Check "合法输入编码成功", r.Status = VATS2_FINGERPRINT_OK
    If r.Status <> VATS2_FINGERPRINT_OK Then Err.Raise 5, , r.ErrorSide & r.ErrorIndex & r.ErrorReason
    again = VATStage2BuildFingerprints(a, b)
    Check "重复调用全部输出相同", SameResult(r, again)
    Check "所有输入字段含位置原样不变", SameInputs(aa, bb)
    RunFingerprint = r
End Function

Private Function SameValue(ByVal x As Variant, ByVal y As Variant) As Boolean
    If VarType(x) <> VarType(y) Then Exit Function
    Select Case VarType(x)
        Case vbNull, vbEmpty: SameValue = True
        Case vbError: SameValue = (CStr(x) = CStr(y))
        Case vbString: SameValue = (StrComp(x, y, vbBinaryCompare) = 0)
        Case Else: SameValue = (x = y)
    End Select
End Function

Private Function SameInputs(ByRef aa As VATS2ASnapshotResult, ByRef bb As VATS2BSnapshotResult) As Boolean
    Dim i As Long
    If aa.Status <> a.Status Or aa.RecordCount <> a.RecordCount Or aa.SheetName <> a.SheetName Or aa.HeaderRow <> a.HeaderRow Or aa.ErrorRow <> a.ErrorRow Or aa.ErrorReason <> a.ErrorReason Then Exit Function
    If bb.Status <> b.Status Or bb.RecordCount <> b.RecordCount Or bb.SheetName <> b.SheetName Or bb.HeaderRow <> b.HeaderRow Or bb.ErrorRow <> b.ErrorRow Or bb.ErrorReason <> b.ErrorReason Then Exit Function
    For i = 1 To aa.RecordCount
        With aa.Records(i)
            If .AIndex <> a.Records(i).AIndex Or .ExcelRow <> a.Records(i).ExcelRow Or .InvoiceCellAddress <> a.Records(i).InvoiceCellAddress Or .AmountCellAddress <> a.Records(i).AmountCellAddress Then Exit Function
            If Not SameValue(.InvoiceDigitsRaw, a.Records(i).InvoiceDigitsRaw) Or Not SameValue(.AmountRaw, a.Records(i).AmountRaw) Then Exit Function
            If .IsCompleted <> a.Records(i).IsCompleted Or .InvoiceHasFormula <> a.Records(i).InvoiceHasFormula Or .AmountHasFormula <> a.Records(i).AmountHasFormula Then Exit Function
        End With
    Next i
    For i = 1 To bb.RecordCount
        With bb.Records(i)
            If .BIndex <> b.Records(i).BIndex Or .ExcelRow <> b.Records(i).ExcelRow Or .SupplierCellAddress <> b.Records(i).SupplierCellAddress Or .AmountCellAddress <> b.Records(i).AmountCellAddress Then Exit Function
            If Not SameValue(.SupplierTextRaw, b.Records(i).SupplierTextRaw) Or Not SameValue(.AmountRaw, b.Records(i).AmountRaw) Then Exit Function
            If .IsCompleted <> b.Records(i).IsCompleted Or .SupplierHasFormula <> b.Records(i).SupplierHasFormula Or .AmountHasFormula <> b.Records(i).AmountHasFormula Then Exit Function
        End With
    Next i
    SameInputs = True
End Function

Private Function SameResult(ByRef x As VATS2FingerprintResult, ByRef y As VATS2FingerprintResult) As Boolean
    Dim i As Long
    If x.Status <> y.Status Or x.ARecordCount <> y.ARecordCount Or x.BRecordCount <> y.BRecordCount Or x.ErrorSide <> y.ErrorSide Or x.ErrorIndex <> y.ErrorIndex Or x.ErrorReason <> y.ErrorReason Then Exit Function
    If x.ABatchFingerprint <> y.ABatchFingerprint Or x.BBatchFingerprint <> y.BBatchFingerprint Or x.CombinedFingerprint <> y.CombinedFingerprint Then Exit Function
    For i = 1 To x.ARecordCount: If x.AFingerprints(i) <> y.AFingerprints(i) Then Exit Function
    Next i
    For i = 1 To x.BRecordCount: If x.BFingerprints(i) <> y.BFingerprints(i) Then Exit Function
    Next i
    SameResult = True
End Function

Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then passed = passed + 1 Else failed = failed + 1: log = log & "FAIL " & name & vbCrLf
End Sub
