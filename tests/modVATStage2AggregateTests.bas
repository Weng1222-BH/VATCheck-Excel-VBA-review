Attribute VB_Name = "modVATStage2AggregateTests"
Option Explicit

Private passed As Long, failed As Long, log As String
Private testBook As Workbook, a As Worksheet, b As Worksheet, color As Long

Public Function VATStage2AggregateGate_SelfTest() As String
    Dim errorText As String
    On Error GoTo Unexpected
    passed = 0: failed = 0: log = "": color = RGB(255, 255, 0)
    Set testBook = Application.Workbooks.Add(xlWBATWorksheet)
    Set a = testBook.Worksheets(1): a.Name = "合成A"
    Set b = testBook.Worksheets.Add(After:=a): b.Name = "合成B"
    TestTotals
    TestWarnings
    TestBlocked
    TestHeadersAndOverflow
    log = log & "INFO CalculationState：xlDone路径原生验证；非xlDone为代码级guard，未改变全局计算策略以制造状态。" & vbCrLf
    GoTo Done
Unexpected:
    errorText = Err.Number & " " & Err.Description
    failed = failed + 1: log = log & "UNEXPECTED " & errorText & vbCrLf
Done:
    '只关闭本测试创建且从未保存的临时工作簿。
    On Error Resume Next
    If Not testBook Is Nothing Then testBook.Close SaveChanges:=False
    Set a = Nothing: Set b = Nothing: Set testBook = Nothing
    On Error GoTo 0
    If failed = 0 Then
        VATStage2AggregateGate_SelfTest = "PASS: " & passed & " assertions" & vbCrLf & log
    Else
        VATStage2AggregateGate_SelfTest = "FAIL: " & failed & "; PASS: " & passed & vbCrLf & log
    End If
End Function

Private Sub ResetFixture()
    a.Range("A1:M20").UnMerge: b.Range("A1:M20").UnMerge
    a.Range("A1:M20").Clear: b.Range("A1:M20").Clear
    a.Range("I3").Value2 = VAT_FIELD_A: b.Range("F5").Value2 = VAT_FIELD_B
End Sub

Private Sub PutAmount(ByVal cell As Range, ByVal value As Variant)
    cell.Interior.Pattern = xlSolid: cell.Interior.Color = color
    cell.Value2 = value
End Sub

Private Sub TestTotals()
    Dim r As VATS2AggregateGateResult
    Call ResetFixture
    PutAmount a.Range("I4"), 10: PutAmount b.Range("F6"), 10
    r = Run(a.Range("I3"), b.Range("F5"))
    Check "总体相等且提供本次短尾号资格", r.Status = VATS2_AGGREGATE_OK And r.TotalsEqual And r.ReliableEqualForShortSuffix
    Check "xlDone guard正常通路", r.CalculationState = xlDone And r.AScanSucceeded And r.BScanSucceeded
    b.Range("F6").Value2 = 8: r = Run(a.Range("I3"), b.Range("F5"))
    Check "正差为A减B", Not r.TotalsEqual And r.Difference = CDec(2) And Not r.ReliableEqualForShortSuffix
    b.Range("F6").Value2 = 12: r = Run(a.Range("I3"), b.Range("F5"))
    Check "负差为A减B", Not r.TotalsEqual And r.Difference = CDec(-2)
    PutAmount a.Range("I5"), 2: r = Run(a.Range("I3"), b.Range("F5"))
    Check "多A少B且行数不同仍可靠相等", r.TotalsEqual And r.ReliableEqualForShortSuffix And r.ACompletedCount = 2 And r.BCompletedCount = 1 And r.AIncludedCount = 2 And r.BIncludedCount = 1
    a.Range("I4").Value2 = 0.1: a.Range("I5").Value2 = 0.2: b.Range("F6").Value2 = 0.3
    r = Run(a.Range("I3"), b.Range("F5"))
    Check "0.1加0.2与0.3精确相等", r.TotalsEqual And r.Difference = CDec(0)
    Check "三项金额均保留Decimal子类型", VarType(r.TotalA) = vbDecimal And VarType(r.TotalB) = vbDecimal And VarType(r.Difference) = vbDecimal And VarType(r.TotalsEqual) = vbBoolean
    Call ResetFixture
    a.Range("I4").Value2 = 100: b.Range("F6").Value2 = 200
    r = Run(a.Range("I3"), b.Range("F5"))
    Check "零完成记录为可靠零总额且不按数量额外限制", r.ACompletedCount = 0 And r.BCompletedCount = 0 And r.TotalA = CDec(0) And r.TotalB = CDec(0) And r.TotalsEqual And r.ReliableEqualForShortSuffix
    Call ResetFixture
    PutAmount a.Range("I4"), 2: PutAmount b.Range("F6"), 2
    a.Range("I1").Formula = "=SUM(I4:I10)": a.Range("I1").Interior.Color = color
    b.Range("F1").Formula = "=SUM(F6:F10)": b.Range("F1").Interior.Color = color
    a.Range("I1").Calculate: b.Range("F1").Calculate
    a.Range("J4").Value2 = 999: a.Range("J4").Interior.Color = color
    a.Range("I5").Value2 = 888: a.Range("I5").Interior.Color = RGB(0, 255, 0)
    r = Run(a.Range("I3"), b.Range("F5"))
    Check "复用冻结目标列和动态数据起点", r.TotalA = CDec(2) And r.TotalB = CDec(2) And r.ACompletedCount = 1
End Sub

Private Sub TestWarnings()
    Dim r As VATS2AggregateGateResult
    Call ResetFixture
    PutAmount a.Range("I4"), 1: PutAmount b.Range("F6"), 1: PutAmount a.Range("I5"), Empty
    r = Run(a.Range("I3"), b.Range("F5"))
    Check "A真正空金额仍可比较但无豁免资格", r.Status = VATS2_AGGREGATE_OK And r.TotalsEqual And Not r.ReliableEqualForShortSuffix
    Check "A计数区分完成计入空金额与警告", r.ACompletedCount = 2 And r.AIncludedCount = 1 And r.AEmptyWarningCount = 1 And r.AWarningCount = 1 And r.AIssueCount = 0 And r.BWarningCount = 0
    Check "A原始警告位置保留", InStr(r.AWarnings(1), "!I5") > 0 And InStr(r.AWarnings(1), "空金额，已跳过未计入") > 0
    a.Range("I5").Interior.Pattern = xlPatternNone: PutAmount b.Range("F7"), Empty
    r = Run(a.Range("I3"), b.Range("F5"))
    Check "B真正空金额仍可比较但无豁免资格", r.Status = VATS2_AGGREGATE_OK And r.TotalsEqual And Not r.ReliableEqualForShortSuffix
    Check "B计数和位置保留", r.BCompletedCount = 2 And r.BIncludedCount = 1 And r.BEmptyWarningCount = 1 And r.BWarningCount = 1 And InStr(r.BWarnings(1), "!F7") > 0
    PutAmount a.Range("I5"), Empty: PutAmount a.Range("I9"), Empty: PutAmount b.Range("F10"), Empty
    r = Run(a.Range("I3"), b.Range("F5"))
    Check "两侧所有警告及原序完整保留", r.AWarningCount = 2 And r.BWarningCount = 2 And r.ACompletedCount = 3 And r.BCompletedCount = 3 And InStr(r.AWarnings(2), "!I9") > 0 And InStr(r.BWarnings(2), "!F10") > 0
    b.Range("F6").Value2 = 9: r = Run(a.Range("I3"), b.Range("F5"))
    Check "警告且有效总额不等仍报告数学差额", r.Status = VATS2_AGGREGATE_OK And Not r.TotalsEqual And r.Difference = CDec(-8) And Not r.ReliableEqualForShortSuffix
    Call ResetFixture: PutAmount a.Range("I4"), Empty: PutAmount b.Range("F6"), Empty
    r = Run(a.Range("I3"), b.Range("F5"))
    Check "只有完成空金额有效零相等但资格为False", r.TotalsEqual And r.AIncludedCount = 0 And r.BIncludedCount = 0 And Not r.ReliableEqualForShortSuffix
End Sub

Private Sub TestBlocked()
    Dim r As VATS2AggregateGateResult, mode As Long, target As Range
    For mode = 1 To 8
        Call ResetFixture: PutAmount a.Range("I4"), 1: PutAmount b.Range("F6"), 1
        If mode Mod 2 = 1 Then Set target = a.Range("I4") Else Set target = b.Range("F6")
        Select Case mode
            Case 1, 2
                target.Formula = "=" & Chr$(34) & Chr$(34)
                target.Calculate
            Case 3, 4
                target.NumberFormat = "@": target.Value2 = "1"
            Case 5, 6
                target.Value2 = CVErr(2042)
            Case 7, 8
                target.Value2 = True
        End Select
        r = Run(a.Range("I3"), b.Range("F5"))
        Check "致命金额异常阻断总体结论" & mode, r.Status = VATS2_SCAN_BLOCKED And IsEmpty(r.TotalsEqual) And IsEmpty(r.Difference) And Not r.ReliableEqualForShortSuffix
        Check "一侧失败另一侧仍返回扫描证据" & mode, (r.AScanSucceeded Xor r.BScanSucceeded) And r.AIssueCount + r.BIssueCount = 1
    Next mode
    Call ResetFixture
    PutAmount a.Range("I4"), CVErr(2042): PutAmount a.Range("I5"), True: PutAmount a.Range("I6"), Empty
    PutAmount b.Range("F6"), CVErr(2007): PutAmount b.Range("F7"), Empty
    r = Run(a.Range("I3"), b.Range("F5"))
    Check "两侧问题与警告同时原样保留", r.AIssueCount = 2 And r.BIssueCount = 1 And r.AWarningCount = 1 And r.BWarningCount = 1 And InStr(r.AIssues(1), "!I4") > 0 And InStr(r.AIssues(2), "!I5") > 0 And InStr(r.BIssues(1), "!F6") > 0
    Call ResetFixture: PutAmount a.Range("I4"), 2: PutAmount b.Range("F6"), 2
    a.Range("I4:I5").Merge
    r = Run(a.Range("I3"), b.Range("F5"))
    Check "合并金额仍由VATScan阻断", r.Status = VATS2_SCAN_BLOCKED And r.AIssueCount > 0
    Call ResetFixture: PutAmount a.Range("I4"), 1E+30: PutAmount b.Range("F6"), 2
    r = Run(a.Range("I3"), b.Range("F5"))
    Check "Decimal金额转换失败仍由VATScan处理", r.Status = VATS2_SCAN_BLOCKED And Not r.AScanSucceeded And r.BScanSucceeded
End Sub

Private Sub TestHeadersAndOverflow()
    Dim r As VATS2AggregateGateResult, missing As Range
    Call ResetFixture: PutAmount a.Range("I4"), 1: PutAmount b.Range("F6"), 1
    a.Range("I3").Value2 = "已修改"
    r = Run(a.Range("I3"), b.Range("F5"))
    Check "失效表头保留VATScan原始问题", r.Status = VATS2_SCAN_BLOCKED And InStr(r.AIssues(1), "表头已改变") > 0
    r = Run(missing, b.Range("F5"))
    Check "Nothing表头仍尝试另一侧并阻断", r.Status = VATS2_SCAN_BLOCKED And Not r.AScanSucceeded And r.BScanSucceeded
    Call ResetFixture: PutAmount a.Range("I4"), 7E+28: PutAmount b.Range("F6"), -7E+28
    r = Run(a.Range("I3"), b.Range("F5"))
    Check "双侧成功但差额溢出独立标识", r.AScanSucceeded And r.BScanSucceeded And r.Status = VATS2_AGGREGATE_ARITHMETIC_ERROR And IsEmpty(r.TotalsEqual) And Not r.ReliableEqualForShortSuffix And Len(r.ErrorReason) > 0
    Check "差额失败不伪造VATScan问题", r.AIssueCount = 0 And r.BIssueCount = 0
End Sub

Private Function Run(ByVal ha As Range, ByVal hb As Range) As VATS2AggregateGateResult
    Dim r As VATS2AggregateGateResult, again As VATS2AggregateGateResult
    Dim before As String, calculation As XlCalculation, saved As Boolean, books As Long
    Dim ia As New Collection, ib As New Collection, wa As New Collection, wb As New Collection
    Dim ta As Variant, tb As Variant, ca As Long, cb As Long, na As Long, nb As Long, ea As Long, eb As Long, oka As Boolean, okb As Boolean
    before = CellsKey(a) & CellsKey(b): calculation = Application.Calculation
    saved = testBook.Saved: books = Application.Workbooks.Count
    r = VATStage2EvaluateAggregateGate(ha, hb, color)
    Check "单元格值公式颜色及工作簿状态只读", before = CellsKey(a) & CellsKey(b) And saved = testBook.Saved And books = Application.Workbooks.Count And calculation = Application.Calculation
    again = VATStage2EvaluateAggregateGate(ha, hb, color)
    Check "重复调用完整输出确定", ResultKey(r) = ResultKey(again)
    '独立直接调用冻结扫描，逐字段对照原始证据；不复制业务扫描逻辑。
    oka = VATScan(ha, VAT_FIELD_A, "A", color, ta, ca, ia, na, ea, wa)
    okb = VATScan(hb, VAT_FIELD_B, "B", color, tb, cb, ib, nb, eb, wb)
    Check "冻结扫描状态和六项计数完整传播", r.AScanSucceeded = oka And r.BScanSucceeded = okb And r.ACompletedCount = ca And r.BCompletedCount = cb And r.AIncludedCount = na And r.BIncludedCount = nb And r.AEmptyWarningCount = ea And r.BEmptyWarningCount = eb
    Check "两侧Issue与Warning逐字逐项一致", SameMessages(ia, r.AIssues, r.AIssueCount) And SameMessages(ib, r.BIssues, r.BIssueCount) And SameMessages(wa, r.AWarnings, r.AWarningCount) And SameMessages(wb, r.BWarnings, r.BWarningCount)
    If r.Status = VATS2_AGGREGATE_OK Then
        Check "总额直接来自VATScan且资格符合契约", r.TotalA = ta And r.TotalB = tb And r.ReliableEqualForShortSuffix = (oka And okb And wa.Count = 0 And wb.Count = 0 And r.TotalsEqual)
    Else
        Check "非OK不泄露部分总额或一致性结论", IsEmpty(r.TotalA) And IsEmpty(r.TotalB) And IsEmpty(r.Difference) And IsEmpty(r.TotalsEqual) And Not r.ReliableEqualForShortSuffix
    End If
    Run = r
End Function

Private Function SameMessages(ByVal source As Collection, ByRef values() As String, ByVal count As Long) As Boolean
    Dim i As Long
    If source.Count <> count Then Exit Function
    For i = 1 To count
        If StrComp(source(i), values(i), vbBinaryCompare) <> 0 Then Exit Function
    Next i
    SameMessages = True
End Function

Private Function RawKey(ByVal value As Variant) As String
    If IsNull(value) Then RawKey = "Null" Else RawKey = VarType(value) & ":" & CStr(value)
End Function

Private Function CellsKey(ByVal sheet As Worksheet) As String
    Dim cell As Range, s As String
    For Each cell In sheet.Range("A1:M12").Cells
        s = s & "|" & cell.Address & RawKey(cell.Value2) & RawKey(cell.Formula) & ":" & cell.HasFormula & ":" & cell.MergeCells
        s = s & ":" & cell.Interior.Pattern & ":" & cell.Interior.Color & ":" & cell.Interior.ColorIndex & ":" & cell.NumberFormat
    Next cell
    CellsKey = s
End Function

Private Function ResultKey(ByRef r As VATS2AggregateGateResult) As String
    Dim s As String, i As Long
    s = r.Status & ":" & r.CalculationState & ":" & r.AScanSucceeded & ":" & r.BScanSucceeded & ":" & RawKey(r.TotalA) & ":" & RawKey(r.TotalB) & ":" & RawKey(r.Difference) & ":" & RawKey(r.TotalsEqual) & ":" & r.ReliableEqualForShortSuffix
    s = s & ":" & r.ACompletedCount & ":" & r.BCompletedCount & ":" & r.AIncludedCount & ":" & r.BIncludedCount & ":" & r.AEmptyWarningCount & ":" & r.BEmptyWarningCount
    s = s & ":" & r.AIssueCount & ":" & r.BIssueCount & ":" & r.AWarningCount & ":" & r.BWarningCount & ":" & r.ErrorReason
    For i = 1 To r.AIssueCount: s = s & "|" & r.AIssues(i): Next i
    For i = 1 To r.BIssueCount: s = s & "|" & r.BIssues(i): Next i
    For i = 1 To r.AWarningCount: s = s & "|" & r.AWarnings(i): Next i
    For i = 1 To r.BWarningCount: s = s & "|" & r.BWarnings(i): Next i
    ResultKey = s
End Function

Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        passed = passed + 1
    Else
        failed = failed + 1: log = log & "FAIL " & name & vbCrLf
    End If
End Sub
