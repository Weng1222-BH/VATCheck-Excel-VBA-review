Attribute VB_Name = "modVATStage2SnapshotTests"
Option Explicit

Private passed As Long, failed As Long, details As String

Public Function VATStage2Snapshot_SelfTest() As String
    Dim book As Workbook, a As Worksheet, b As Worksheet, parity As Worksheet
    Dim snapshot As VATS2ASnapshotResult, bSnapshot As VATS2BSnapshotResult
    Dim none As Worksheet, i As Long, failure As String
    On Error GoTo Unexpected
    passed = 0: failed = 0: details = vbNullString
    '只创建、写入和关闭本测试自己的临时工作簿。
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set a = book.Worksheets(1): a.Name = "快照A"
    Set b = book.Worksheets.Add: b.Name = "快照B"
    Set parity = book.Worksheets.Add: parity.Name = "颜色对照"
    TestA a
    TestB b
    TestParity parity
    TestBounds a, b
    TestCapacity parity
    snapshot = VATStage2ReadASnapshot(a, 9, 4, 9, RGB(0, 176, 80))
    bSnapshot = VATStage2ReadBSnapshot(b, 4, 2, 6, RGB(0, 176, 80))
    book.Close False
    Set book = Nothing
    Check "A快照在原工作簿关闭后仍可读取原值与位置", _
          snapshot.Records(1).InvoiceDigitsRaw = "00123456789012345678" And snapshot.Records(1).ExcelRow = 10
    Check "B快照在原工作簿关闭后仍可读取原文", bSnapshot.Records(1).SupplierTextRaw = " 当前未保存原文 "
    snapshot = VATStage2ReadASnapshot(a, 9, 4, 9, vbWhite)
    Check "已关闭A Worksheet明确失效且无部分记录", snapshot.Status = VATS2_SNAPSHOT_INVALID_SHEET And snapshot.RecordCount = 0
    bSnapshot = VATStage2ReadBSnapshot(b, 4, 2, 6, vbWhite)
    Check "已关闭B Worksheet明确失效", bSnapshot.Status = VATS2_SNAPSHOT_INVALID_SHEET
    GoTo Finished
Unexpected:
    failure = CStr(Err.Number) & " " & Err.Description
    failed = failed + 1: details = details & "UNEXPECTED: " & failure & vbCrLf
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
Finished:
    If failed = 0 Then
        VATStage2Snapshot_SelfTest = "PASS: " & passed & " assertions" & vbCrLf & details
    Else
        VATStage2Snapshot_SelfTest = "FAIL: " & failed & "; PASS: " & passed & vbCrLf & details
    End If
End Function

Private Sub TestA(ByVal sheet As Worksheet)
    Dim r As VATS2ASnapshotResult, again As VATS2ASnapshotResult
    Dim before As String, savedBefore As Boolean, i As Long, rows As Variant, objects As Boolean
    Dim green As Long: green = RGB(0, 176, 80)
    sheet.Range("D9").Value2 = "数电发票号码"
    sheet.Range("I9").Value2 = VAT_FIELD_A
    sheet.Range("I1").Formula = "=SUM(I10:I23)"
    sheet.Range("D10").NumberFormat = "@": sheet.Range("D10").Value2 = "00123456789012345678"
    sheet.Range("I10").Value2 = 12.34: sheet.Range("I10").Interior.Color = green
    sheet.Range("D11").Value2 = " 身份原文 ": sheet.Range("D11").Interior.Color = green
    sheet.Range("I11").NumberFormat = "@": sheet.Range("I11").Value2 = "0012.340"
    sheet.Range("D15").Value2 = "333": sheet.Range("I15").Interior.Color = green
    sheet.Range("I16").Value2 = 0
    sheet.Range("I17").Interior.Color = green
    sheet.Range("D18").Formula = "="""""
    sheet.Range("I19").Formula = "="""""
    sheet.Range("D20").Value2 = CVErr(xlErrNA): sheet.Range("I20").Value2 = CVErr(xlErrDiv0)
    sheet.Range("D21").Value2 = True: sheet.Range("I21").Value2 = False
    sheet.Range("D22").Value2 = "未填色": sheet.Range("I22").Value2 = 0
    sheet.Range("D23").Value2 = 12345: sheet.Range("I23").Value2 = 45678
    sheet.Range("I23").NumberFormat = "yyyy-mm-dd"
    sheet.Range("A12:H12").Interior.Color = green
    sheet.Range("I40").Interior.Color = vbYellow
    sheet.Range("Z45").Font.Bold = True
    sheet.Rows(11).Hidden = True
    before = SheetEvidence(sheet, "A1:Z45"): savedBefore = sheet.Parent.Saved
    r = VATStage2ReadASnapshot(sheet, 9, 4, 9, green)
    Check "A多行完整快照和表头元数据", r.Status = VATS2_SNAPSHOT_OK And r.RecordCount = 11 And r.SheetName = "快照A" And r.HeaderRow = 9
    rows = Array(10, 11, 15, 16, 17, 18, 19, 20, 21, 22, 23)
    For i = 1 To r.RecordCount
        If r.Records(i).AIndex <> i Or r.Records(i).ExcelRow <> rows(i - 1) Then Err.Raise 5, , "AIndex或ExcelRow映射错误"
        objects = objects Or IsObject(r.Records(i).InvoiceDigitsRaw) Or IsObject(r.Records(i).AmountRaw)
    Next i
    Check "AIndex连续且非连续Excel行映射准确", True
    Check "A原值字段不含Excel对象", Not objects
    Check "记录数组边界与RecordCount一致", LBound(r.Records) = 1 And UBound(r.Records) = r.RecordCount
    Check "完成状态混合且完成索引仍为1和3", r.Records(1).IsCompleted And Not r.Records(2).IsCompleted And r.Records(3).IsCompleted And r.Records(3).AIndex = 3
    Check "A只检查金额列，身份字段填色不算完成", Not r.Records(2).IsCompleted
    Check "A前导零和String类型保持", VarType(r.Records(1).InvoiceDigitsRaw) = vbString And r.Records(1).InvoiceDigitsRaw = "00123456789012345678"
    Check "A身份原文不Trim", r.Records(2).InvoiceDigitsRaw = " 身份原文 "
    Check "数字金额保留Value2 Double", VarType(r.Records(1).AmountRaw) = vbDouble And r.Records(1).AmountRaw = 12.34
    Check "文本金额不转换", VarType(r.Records(2).AmountRaw) = vbString And r.Records(2).AmountRaw = "0012.340"
    Check "身份存在但完成色空金额保留Empty", r.Records(3).ExcelRow = 15 And IsEmpty(r.Records(3).AmountRaw)
    Check "身份空但零金额仍保留", IsEmpty(r.Records(4).InvoiceDigitsRaw) And r.Records(4).AmountRaw = 0
    Check "双字段空但金额完成色仍保留", IsEmpty(r.Records(5).InvoiceDigitsRaw) And IsEmpty(r.Records(5).AmountRaw) And r.Records(5).IsCompleted
    Check "身份公式空串仍保留", r.Records(6).InvoiceHasFormula And VarType(r.Records(6).InvoiceDigitsRaw) = vbString And r.Records(6).InvoiceDigitsRaw = ""
    Check "金额公式空串仍保留", r.Records(7).AmountHasFormula And VarType(r.Records(7).AmountRaw) = vbString And r.Records(7).AmountRaw = ""
    Check "错误原值不拦截不转零", IsError(r.Records(8).InvoiceDigitsRaw) And IsError(r.Records(8).AmountRaw)
    Check "布尔原值保持", VarType(r.Records(9).InvoiceDigitsRaw) = vbBoolean And VarType(r.Records(9).AmountRaw) = vbBoolean
    Check "未完成零金额保留", r.Records(10).AmountRaw = 0 And Not r.Records(10).IsCompleted
    Check "数值身份和日期格式金额保持Value2类型", VarType(r.Records(11).InvoiceDigitsRaw) = vbDouble And VarType(r.Records(11).AmountRaw) = vbDouble And r.Records(11).AmountRaw = 45678
    Check "A单元格定位不写死列", r.Records(3).InvoiceCellAddress = "D15" And r.Records(3).AmountCellAddress = "I15"
    Check "未修改输入值公式格式与隐藏状态", before = SheetEvidence(sheet, "A1:Z45")
    Check "没有保存工作簿或修改Saved状态", sheet.Parent.Saved = savedBefore And Not sheet.Parent.Saved
    again = VATStage2ReadASnapshot(sheet, 9, 4, 9, green)
    Check "A重复调用完整结果一致", SameA(r, again)
    sheet.Range("I11").Interior.Color = green
    again = VATStage2ReadASnapshot(sheet, 9, 4, 9, green)
    Check "改变现有记录颜色不改变后续Index", again.RecordCount = r.RecordCount And again.Records(2).IsCompleted And again.Records(3).AIndex = 3 And again.Records(3).ExcelRow = 15
    sheet.Range("I10").Value2 = 99.25
    again = VATStage2ReadASnapshot(sheet, 9, 4, 9, green)
    Check "A当前未保存修改即时读取", again.Records(1).AmountRaw = 99.25
    Check "旧快照不随源数据变化", r.Records(1).AmountRaw = 12.34
End Sub

Private Sub TestB(ByVal sheet As Worksheet)
    Dim r As VATS2BSnapshotResult, again As VATS2BSnapshotResult, before As String
    Dim green As Long, original As String, i As Long, objects As Boolean
    green = RGB(0, 176, 80)
    original = " 采购（0000）公司NO.00123/456--lo5 " & vbLf & "（折让）"
    sheet.Range("B4").Value2 = "供应商名称": sheet.Range("F4").Value2 = VAT_FIELD_B
    sheet.Range("F1").Formula = "=SUM(F5:F12)"
    sheet.Range("B5").Value2 = original: sheet.Range("F5").Value2 = 10
    sheet.Range("F5").Interior.Color = green
    sheet.Range("B6").Value2 = "隐藏行": sheet.Range("B6").Interior.Color = green
    sheet.Range("F6").NumberFormat = "@": sheet.Range("F6").Value2 = "10.00"
    sheet.Range("B9").Value2 = "仅身份"
    sheet.Range("F10").Value2 = 20
    sheet.Range("F11").Interior.Color = green
    sheet.Range("B12").Formula = "=""""": sheet.Range("F13").Formula = "="""""
    sheet.Range("B14").Value2 = CVErr(xlErrNA): sheet.Range("F14").Value2 = True
    sheet.Range("A7:E7").Interior.Color = green
    sheet.Range("F40").Interior.Color = vbYellow
    sheet.Range("B4:F14").AutoFilter Field:=1, Criteria1:=original
    before = SheetEvidence(sheet, "A1:F40")
    r = VATStage2ReadBSnapshot(sheet, 4, 2, 6, green)
    Check "B全表快照不受筛选隐藏限制", r.Status = VATS2_SNAPSHOT_OK And r.RecordCount = 8 And r.SheetName = "快照B" And r.HeaderRow = 4
    Check "BIndex与ExcelRow映射及中间空行", r.Records(2).BIndex = 2 And r.Records(2).ExcelRow = 6 And r.Records(3).BIndex = 3 And r.Records(3).ExcelRow = 9
    Check "B供应商原文逐字保留且不调用Parser", r.Records(1).SupplierTextRaw = original
    Check "B完成状态只来自税额列", r.Records(1).IsCompleted And Not r.Records(2).IsCompleted
    Check "B金额String与Double原值保持", VarType(r.Records(1).AmountRaw) = vbDouble And VarType(r.Records(2).AmountRaw) = vbString
    Check "B身份存在金额空且无完成色仍保留", r.Records(3).SupplierTextRaw = "仅身份" And IsEmpty(r.Records(3).AmountRaw) And Not r.Records(3).IsCompleted
    Check "B仅金额仍保留", IsEmpty(r.Records(4).SupplierTextRaw) And r.Records(4).AmountRaw = 20
    Check "B双空完成行仍保留", r.Records(5).ExcelRow = 11 And r.Records(5).IsCompleted And IsEmpty(r.Records(5).AmountRaw)
    Check "B两个字段的公式空串分别保留", r.Records(6).SupplierHasFormula And r.Records(6).SupplierTextRaw = "" And r.Records(7).AmountHasFormula And r.Records(7).AmountRaw = ""
    Check "B错误与Boolean原始类型保持", IsError(r.Records(8).SupplierTextRaw) And VarType(r.Records(8).AmountRaw) = vbBoolean
    Check "B单元格地址来自实际列", r.Records(3).SupplierCellAddress = "B9" And r.Records(3).AmountCellAddress = "F9"
    For i = 1 To r.RecordCount
        objects = objects Or IsObject(r.Records(i).SupplierTextRaw) Or IsObject(r.Records(i).AmountRaw)
    Next i
    Check "B返回值不含Excel对象", Not objects
    again = VATStage2ReadBSnapshot(sheet, 4, 2, 6, green)
    Check "B重复调用完整结果一致", SameB(r, again)
    Check "B源值公式格式筛选及隐藏状态不变", before = SheetEvidence(sheet, "A1:F40") And sheet.FilterMode
    sheet.Range("F6").Interior.Color = green
    again = VATStage2ReadBSnapshot(sheet, 4, 2, 6, green)
    Check "B改变完成色后完整Index不变", again.RecordCount = r.RecordCount And again.Records(2).IsCompleted And again.Records(3).BIndex = 3 And again.Records(3).ExcelRow = 9
    sheet.Range("B5").Value2 = " 当前未保存原文 "
    again = VATStage2ReadBSnapshot(sheet, 4, 2, 6, green)
    Check "B未保存原文即时读取且原快照独立", again.Records(1).SupplierTextRaw = " 当前未保存原文 " And r.Records(1).SupplierTextRaw = original
End Sub

Private Sub TestParity(ByVal sheet As Worksheet)
    Dim mode As Long, color As Long, expected As Boolean, fc As FormatCondition
    Dim r As VATS2ASnapshotResult, b As VATS2BSnapshotResult, count As Long, total As Variant
    Dim issues As Collection, ok As Boolean, before As String
    For mode = 0 To 5
        sheet.Cells.Clear
        sheet.Range("A1").Value2 = "身份": sheet.Range("B1").Value2 = VAT_FIELD_B
        sheet.Range("A2").Value2 = "00123": sheet.Range("B2").Value2 = 10
        color = vbWhite: expected = True
        Select Case mode
            Case 0
                expected = False
            Case 1
                sheet.Range("B2").Interior.Color = vbWhite
            Case 2
                color = RGB(0, 176, 80): sheet.Range("B2").Interior.Color = color
            Case 3
                sheet.Range("B2").Interior.Color = vbYellow: expected = False
            Case 4
                color = RGB(0, 176, 80)
                Set fc = sheet.Range("B2").FormatConditions.Add(xlCellValue, xlGreater, "0")
                fc.Interior.Color = color
                sheet.Calculate
            Case 5
                color = RGB(0, 176, 80)
                With sheet.Range("B2").Interior
                    .Color = color: .Pattern = xlPatternGray25: .PatternColor = vbBlack
                End With
        End Select
        before = SheetEvidence(sheet, "A1:B2")
        Set issues = New Collection
        ok = VATScan(sheet.Range("B1"), VAT_FIELD_B, "B", color, total, count, issues)
        r = VATStage2ReadASnapshot(sheet, 1, 1, 2, color)
        b = VATStage2ReadBSnapshot(sheet, 1, 1, 2, color)
        Check "颜色模式" & mode & "原生预期与Stage1计数一致", (count = 1) = expected
        Check "颜色模式" & mode & " A/B DisplayFormat parity", r.Status = VATS2_SNAPSHOT_OK And b.Status = VATS2_SNAPSHOT_OK And r.Records(1).IsCompleted = expected And b.Records(1).IsCompleted = expected
        Check "颜色模式" & mode & "快照读取不改变格式", before = SheetEvidence(sheet, "A1:B2")
        If mode = 4 Then Check "条件格式使用实际显示而非基础Interior", sheet.Range("B2").Interior.Pattern = xlPatternNone And r.Records(1).IsCompleted
        If mode = 5 Then Check "图案仍保留Stage1完成计数但不做金额合法性判断", count = 1 And Not ok And issues.Count = 1 And r.Records(1).IsCompleted
    Next mode
End Sub

Private Sub TestBounds(ByVal a As Worksheet, ByVal b As Worksheet)
    Dim r As VATS2ASnapshotResult, s As VATS2BSnapshotResult, none As Worksheet
    Dim n As Variant, col As Variant
    r = VATStage2ReadASnapshot(none, 1, 1, 2, vbWhite)
    s = VATStage2ReadBSnapshot(none, 1, 1, 2, vbWhite)
    Check "Nothing Worksheet两侧明确失败", r.Status = VATS2_SNAPSHOT_INVALID_SHEET And s.Status = VATS2_SNAPSHOT_INVALID_SHEET And r.RecordCount = 0 And s.RecordCount = 0
    For Each n In Array(-1, 0, a.Rows.Count + 1)
        r = VATStage2ReadASnapshot(a, n, 1, 2, vbWhite)
        s = VATStage2ReadBSnapshot(b, n, 1, 2, vbWhite)
        Check "两侧非法HeaderRow=" & n, r.Status = VATS2_SNAPSHOT_INVALID_HEADER And s.Status = VATS2_SNAPSHOT_INVALID_HEADER And r.ErrorReason <> "" And s.RecordCount = 0
    Next n
    For Each col In Array(-1, 0, a.Columns.Count + 1)
        r = VATStage2ReadASnapshot(a, 1, col, 2, vbWhite)
        s = VATStage2ReadBSnapshot(b, 1, col, 2, vbWhite)
        Check "两侧身份列越界=" & col, r.Status = VATS2_SNAPSHOT_INVALID_COLUMN And s.Status = VATS2_SNAPSHOT_INVALID_COLUMN
        r = VATStage2ReadASnapshot(a, 1, 1, col, vbWhite)
        s = VATStage2ReadBSnapshot(b, 1, 1, col, vbWhite)
        Check "两侧金额列越界=" & col, r.Status = VATS2_SNAPSHOT_INVALID_COLUMN And s.Status = VATS2_SNAPSHOT_INVALID_COLUMN
    Next col
    r = VATStage2ReadASnapshot(a, a.Rows.Count, 1, a.Columns.Count, vbWhite)
    s = VATStage2ReadBSnapshot(b, b.Rows.Count, 1, b.Columns.Count, vbWhite)
    Check "合法最后行表头和最后列返回正常空结果", r.Status = VATS2_SNAPSHOT_OK And s.Status = VATS2_SNAPSHOT_OK And r.RecordCount = 0 And s.RecordCount = 0
End Sub

Private Sub TestCapacity(ByVal sheet As Worksheet)
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult
    sheet.Cells.Clear
    a = VATStage2ReadASnapshot(sheet, 1, 1, 2, vbWhite)
    Check "全空工作表正常零记录", a.Status = VATS2_SNAPSHOT_OK And a.RecordCount = 0
    sheet.Range("B600").NumberFormat = "0.000"
    a = VATStage2ReadASnapshot(sheet, 1, 1, 2, vbWhite)
    Check "仅格式UsedRange尾部不制造记录", a.Status = VATS2_SNAPSHOT_OK And a.RecordCount = 0
    sheet.Range("A2:A301").Value2 = "记录"
    a = VATStage2ReadASnapshot(sheet, 1, 1, 2, vbWhite)
    b = VATStage2ReadBSnapshot(sheet, 1, 1, 2, vbWhite)
    Check "A数组扩容跨256条不丢失映射", a.RecordCount = 300 And a.Records(300).AIndex = 300 And a.Records(300).ExcelRow = 301 And UBound(a.Records) = 300
    Check "B数组扩容跨256条不丢失映射", b.RecordCount = 300 And b.Records(300).BIndex = 300 And b.Records(300).ExcelRow = 301 And UBound(b.Records) = 300
End Sub

'测试证据包括业务值/公式/格式/隐藏/筛选，读前读后必须完全一致。
Private Function SheetEvidence(ByVal sheet As Worksheet, ByVal address As String) As String
    Dim cell As Range, value As String, fc As FormatCondition
    For Each cell In sheet.Range(address).Cells
        value = value & cell.Address & ":" & RawKey(cell.Value2) & ":" & RawKey(cell.Formula) & ":" & _
                cell.NumberFormat & ":" & cell.Interior.Pattern & ":" & cell.Interior.Color & ":" & _
                cell.Interior.ColorIndex & ":" & cell.Interior.PatternColor & ":" & cell.Font.Bold & ":" & _
                cell.EntireRow.Hidden & ":" & cell.FormatConditions.Count & ";"
        For Each fc In cell.FormatConditions
            value = value & fc.Type & ":" & fc.Formula1 & ":" & fc.Interior.Color & ";"
        Next fc
    Next cell
    SheetEvidence = value & sheet.AutoFilterMode & ":" & sheet.FilterMode & ":" & sheet.UsedRange.Address
End Function

Private Function RawKey(ByVal value As Variant) As String
    RawKey = CStr(VarType(value)) & ":" & CStr(value)
End Function

Private Function SameA(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2ASnapshotResult) As Boolean
    Dim i As Long
    If a.Status <> b.Status Or a.RecordCount <> b.RecordCount Or a.SheetName <> b.SheetName Or a.HeaderRow <> b.HeaderRow Then Exit Function
    For i = 1 To a.RecordCount
        If a.Records(i).AIndex <> b.Records(i).AIndex Or a.Records(i).ExcelRow <> b.Records(i).ExcelRow Then Exit Function
        If RawKey(a.Records(i).InvoiceDigitsRaw) <> RawKey(b.Records(i).InvoiceDigitsRaw) Or RawKey(a.Records(i).AmountRaw) <> RawKey(b.Records(i).AmountRaw) Then Exit Function
        If a.Records(i).IsCompleted <> b.Records(i).IsCompleted Or a.Records(i).InvoiceHasFormula <> b.Records(i).InvoiceHasFormula Or a.Records(i).AmountHasFormula <> b.Records(i).AmountHasFormula Then Exit Function
        If a.Records(i).InvoiceCellAddress <> b.Records(i).InvoiceCellAddress Or a.Records(i).AmountCellAddress <> b.Records(i).AmountCellAddress Then Exit Function
    Next i
    SameA = True
End Function

Private Function SameB(ByRef a As VATS2BSnapshotResult, ByRef b As VATS2BSnapshotResult) As Boolean
    Dim i As Long
    If a.Status <> b.Status Or a.RecordCount <> b.RecordCount Or a.SheetName <> b.SheetName Or a.HeaderRow <> b.HeaderRow Then Exit Function
    For i = 1 To a.RecordCount
        If a.Records(i).BIndex <> b.Records(i).BIndex Or a.Records(i).ExcelRow <> b.Records(i).ExcelRow Then Exit Function
        If RawKey(a.Records(i).SupplierTextRaw) <> RawKey(b.Records(i).SupplierTextRaw) Or RawKey(a.Records(i).AmountRaw) <> RawKey(b.Records(i).AmountRaw) Then Exit Function
        If a.Records(i).IsCompleted <> b.Records(i).IsCompleted Or a.Records(i).SupplierHasFormula <> b.Records(i).SupplierHasFormula Or a.Records(i).AmountHasFormula <> b.Records(i).AmountHasFormula Then Exit Function
        If a.Records(i).SupplierCellAddress <> b.Records(i).SupplierCellAddress Or a.Records(i).AmountCellAddress <> b.Records(i).AmountCellAddress Then Exit Function
    Next i
    SameB = True
End Function

Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        passed = passed + 1: details = details & "PASS " & name & vbCrLf
    Else
        failed = failed + 1: details = details & "FAIL " & name & vbCrLf
    End If
End Sub
