Attribute VB_Name = "modVATCheck"
Option Explicit

Public Const VAT_FIELD_A As String = "有效抵扣税额*"
Public Const VAT_FIELD_B As String = "税额"

Public Sub VATCheck_Start()
    '非模态窗体允许用户保留界面并继续在 Excel 中填色。
    VATActivateHost
    frmVATCheck.Show vbModeless
End Sub

Public Sub VATActivateHost()
    Dim window As Window
    '将窗体依附于工具宿主窗口，避免关闭业务文件时连带丢失窗体。
    For Each window In ThisWorkbook.Windows
        If window.Visible Then
            window.Activate
            Exit For
        End If
    Next window
End Sub

Public Function VATOpenWorkbook(ByVal filePath As String) As Workbook
    Dim book As Workbook, oldSecurity As MsoAutomationSecurity
    Dim errorNumber As Long, errorText As String
    For Each book In Application.Workbooks
        If StrComp(book.FullName, filePath, vbTextCompare) = 0 Then
            Set VATOpenWorkbook = book
            Exit Function
        End If
    Next book
    '只在用户选择后打开文件；禁用该次打开的宏，不更新外部链接。
    oldSecurity = Application.AutomationSecurity
    On Error GoTo Failed
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Set VATOpenWorkbook = Application.Workbooks.Open(Filename:=filePath, UpdateLinks:=0, _
                                                    Notify:=False, AddToMru:=False)
    Application.AutomationSecurity = oldSecurity
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Application.AutomationSecurity = oldSecurity
    Err.Raise errorNumber, "VATOpenWorkbook", errorText
End Function

Public Function VATWorkbookOpen(ByVal book As Workbook) As Boolean
    Dim candidate As Workbook
    If book Is Nothing Then Exit Function
    For Each candidate In Application.Workbooks
        If candidate Is book Then
            VATWorkbookOpen = True
            Exit Function
        End If
    Next candidate
End Function

Public Function VATFindHeaders(ByVal sheet As Worksheet, ByVal field As String) As Collection
    Dim found As New Collection, cell As Range, firstAddress As String
    Dim area As Range, literal As String
    Set area = sheet.UsedRange
    '星号是表头的真实字符，必须转义，不能当作通配符。
    literal = Replace(Replace(Replace(field, "~", "~~"), "*", "~*"), "?", "~?")
    '自定义格式可能给显示文本添加空白；先找候选，再按 Value2 严格确认。
    Set cell = area.Find(What:=literal, After:=area.Cells(area.Rows.Count, area.Columns.Count), _
                        LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, MatchCase:=True, MatchByte:=True, SearchFormat:=False)
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        Do
            If VarType(cell.Value2) = vbString Then
                If StrComp(CStr(cell.Value2), field, vbBinaryCompare) = 0 Then found.Add cell
            End If
            Set cell = area.Find(What:=literal, After:=cell, LookIn:=xlValues, LookAt:=xlPart, _
                                SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                MatchCase:=True, MatchByte:=True, SearchFormat:=False)
            If cell Is Nothing Then Exit Do
        Loop While cell.Address <> firstAddress
    End If
    Set VATFindHeaders = found
End Function

Public Function VATColumnLetter(ByVal column As Long) As String
    Dim result As String
    Do While column > 0
        column = column - 1
        result = Chr$(65 + column Mod 26) & result
        column = column \ 26
    Loop
    VATColumnLetter = result
End Function

Public Function VATLocation(ByVal header As Range) As String
    VATLocation = "第" & header.Column & "列（" & VATColumnLetter(header.Column) & "列），表头 " & header.Address(False, False)
End Function

Public Function VATReadColor(ByVal sample As Range, ByRef color As Long, ByRef reason As String) As Boolean
    On Error GoTo Failed
    If sample.CountLarge <> 1 Then
        reason = "请只选择一个已正确填色的单元格。"
        Exit Function
    End If
    With sample.DisplayFormat.Interior
        If .Pattern = xlPatternNone Or .ColorIndex = xlColorIndexNone Then
            reason = "该单元格没有填充颜色，请重新选择。"
            Exit Function
        End If
        If .Pattern <> xlSolid Then
            reason = "请选择使用纯色填充的单元格。"
            Exit Function
        End If
        color = .Color
    End With
    VATReadColor = True
    Exit Function
Failed:
    reason = "无法读取所选单元格的填充颜色：" & Err.Description
End Function

Public Function VATTryAmount(ByVal value As Variant, ByRef amount As Variant, ByRef reason As String) As Boolean
    On Error GoTo Failed
    'Decimal 只能作为 Variant 的子类型保存；绝不使用 Double 或 Currency 累加。
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
            VATTryAmount = True
        Case Else
            reason = "金额不是可用数值"
    End Select
    Exit Function
Failed:
    reason = "金额超出 Decimal 可处理范围或无法转换"
End Function

Public Function VATMoney(ByVal amount As Variant) As String
    Dim decimalText As String, whole As String, fraction As String, sign As String
    Dim separator As String, position As Long, grouped As String
    '直接格式化 Decimal 字符串，避免 Format 的中间转换，也不隐藏分以下的差额。
    decimalText = CStr(CDec(amount))
    separator = Mid$(CStr(CDec(1) / CDec(2)), 2, 1)
    If Left$(decimalText, 1) = "-" Then
        sign = "-"
        decimalText = Mid$(decimalText, 2)
    End If
    position = InStr(decimalText, separator)
    If position > 0 Then
        whole = Left$(decimalText, position - 1)
        fraction = Mid$(decimalText, position + 1)
    Else
        whole = decimalText
    End If
    Do While Len(fraction) > 2 And Right$(fraction, 1) = "0"
        fraction = Left$(fraction, Len(fraction) - 1)
    Loop
    Do While Len(fraction) < 2
        fraction = fraction & "0"
    Loop
    Do While Len(whole) > 3
        grouped = "," & Right$(whole, 3) & grouped
        whole = Left$(whole, Len(whole) - 3)
    Loop
    VATMoney = "¥" & sign & whole & grouped & "." & fraction
End Function

Private Sub AddIssue(ByVal issues As Collection, ByVal side As String, ByVal cell As Range, ByVal reason As String)
    issues.Add "表" & side & " [" & cell.Parent.Parent.Name & "]" & cell.Parent.Name & _
               "!" & cell.Address(False, False) & "：" & reason
End Sub

Public Function VATScan(ByVal header As Range, ByVal field As String, ByVal side As String, _
                        ByVal completedColor As Long, ByRef total As Variant, _
                        ByRef completedCount As Long, ByVal issues As Collection, _
                        Optional ByRef includedCount As Long = 0, Optional ByRef emptyCount As Long = 0, _
                        Optional ByVal warnings As Collection = Nothing) As Boolean
    Dim sheet As Worksheet, book As Workbook, area As Range, cell As Range
    Dim lastRow As Long, row As Long, amount As Variant, reason As String, startIssues As Long
    Dim value As Variant, fill As Object
    On Error GoTo Failed
    total = CDec(0)
    completedCount = 0
    includedCount = 0
    emptyCount = 0
    startIssues = issues.Count
    If header Is Nothing Then
        issues.Add "表" & side & "：尚未明确选择目标表头。"
        Exit Function
    End If
    Set sheet = header.Parent
    Set book = sheet.Parent
    If Not VATWorkbookOpen(book) Then
        issues.Add "表" & side & "：原工作簿已关闭，请重新选择 Excel。"
        Exit Function
    End If
    value = header.Value2
    If VarType(value) <> vbString Then
        AddIssue issues, side, header, "表头已改变，请重新选择 Sheet 和目标字段。"
        Exit Function
    End If
    If StrComp(CStr(value), field, vbBinaryCompare) <> 0 Then
        AddIssue issues, side, header, "表头已改变，请重新选择 Sheet 和目标字段。"
        Exit Function
    End If
    If header.MergeCells Then
        AddIssue issues, side, header, "表头为合并单元格，无法确定唯一金额列。"
        Exit Function
    End If
    '使用已用区域的末行，包含只有填色而没有金额的尾行；不能用 End(xlUp) 漏掉异常。
    Set area = sheet.UsedRange
    lastRow = area.Row + area.Rows.Count - 1
    For row = header.Row + 1 To lastRow
        Set cell = sheet.Cells(row, header.Column)
        '只检查目标列，隐藏/筛选行也参与；显示颜色可兼容条件格式。
        Set fill = cell.DisplayFormat.Interior
        If fill.Pattern <> xlPatternNone And fill.ColorIndex <> xlColorIndexNone Then
            If fill.Color = completedColor Then
                completedCount = completedCount + 1
                If cell.MergeCells Then
                    AddIssue issues, side, cell, "已完成金额单元格被合并，无法可靠按行汇总。"
                ElseIf fill.Pattern <> xlSolid Then
                    AddIssue issues, side, cell, "已完成颜色使用了非纯色图案，无法可靠识别。"
                Else
                    reason = vbNullString
                    value = cell.Value2
                    '只放宽真正空单元格；公式空字符串及所有原致命异常仍阻断。
                    If IsEmpty(value) And Not cell.HasFormula Then
                        emptyCount = emptyCount + 1
                        If Not warnings Is Nothing Then AddIssue warnings, side, cell, "空金额，已跳过未计入"
                    ElseIf VATTryAmount(value, amount, reason) Then
                        total = CDec(total) + CDec(amount)
                        includedCount = includedCount + 1
                    Else
                        AddIssue issues, side, cell, reason
                    End If
                End If
            End If
        End If
    Next row
    VATScan = (issues.Count = startIssues)
    Exit Function
Failed:
    If cell Is Nothing Then
        issues.Add "表" & side & "：读取失败，请重新选择 Excel / Sheet。" & Err.Description
    Else
        reason = Err.Description
        AddIssue issues, side, cell, "读取或 Decimal 汇总失败：" & reason
    End If
End Function

Public Function VATResult(ByVal headerA As Range, ByVal headerB As Range, ByVal completedColor As Long) As String
    Dim totalA As Variant, totalB As Variant, difference As Variant
    Dim countA As Long, countB As Long, issues As New Collection
    Dim includedA As Long, includedB As Long, emptyA As Long, emptyB As Long, warnings As New Collection
    Dim okA As Boolean, okB As Boolean, item As Variant, result As String
    On Error GoTo Failed
    If Application.CalculationState <> xlDone Then
        VATResult = "本次无法完成可靠核对：Excel 尚有待计算内容。请先完成计算（F9），再点击一键核对。"
        Exit Function
    End If
    okA = VATScan(headerA, VAT_FIELD_A, "A", completedColor, totalA, countA, issues, includedA, emptyA, warnings)
    okB = VATScan(headerB, VAT_FIELD_B, "B", completedColor, totalB, countB, issues, includedB, emptyB, warnings)
    If Not okA Or Not okB Then
        result = "本次无法完成可靠核对" & vbCrLf & "发现 " & issues.Count & " 个问题；不输出一致性结论。" & vbCrLf & _
                 "表A 已完成颜色：" & countA & "行；表B 已完成颜色：" & countB & "行" & vbCrLf & vbCrLf
        For Each item In issues
            result = result & item & vbCrLf
        Next item
        result = result & VATWarningText(warnings)
        VATResult = result
        Exit Function
    End If
    difference = CDec(totalA) - CDec(totalB)
    If difference = CDec(0) Then
        result = "核对一致"
    Else
        result = "核对不一致"
    End If
    If warnings.Count > 0 Then result = "有效金额部分" & result
    result = result & vbCrLf & vbCrLf & "表A 有效抵扣税额：" & VATMoney(totalA) & vbCrLf & _
             VATCountText(countA, includedA, emptyA) & vbCrLf & vbCrLf & _
             "表B 税额：" & VATMoney(totalB) & vbCrLf & VATCountText(countB, includedB, emptyB) & vbCrLf & vbCrLf & _
             "差额（A - B）：" & VATMoney(difference)
    If difference > CDec(0) Then result = result & vbCrLf & "A 比 B 多 " & VATMoney(difference)
    If difference < CDec(0) Then result = result & vbCrLf & "B 比 A 多 " & VATMoney(-difference)
    If countA = 0 And countB = 0 Then result = result & vbCrLf & "提示：两边均无已完成记录，合计均为零。"
    result = result & VATWarningText(warnings)
    VATResult = result & vbCrLf & vbCrLf & "读取时间：" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
                vbCrLf & "这是点击时的结果；修改金额或颜色后，请再次核对。"
    Exit Function
Failed:
    VATResult = "本次无法完成可靠核对：" & Err.Description
End Function

Private Function VATCountText(ByVal completed As Long, ByVal included As Long, ByVal skipped As Long) As String
    VATCountText = "完成色行数：" & completed & "行；实际计入核算行数：" & included & _
                   "行；空金额未计入行数：" & skipped & "行"
End Function

Private Function VATWarningText(ByVal warnings As Collection) As String
    Dim item As Variant, result As String
    If warnings.Count = 0 Then Exit Function
    result = vbCrLf & vbCrLf & "警告：以下 " & warnings.Count & " 个完成色空金额单元格已跳过；结论仅针对有效金额。" & vbCrLf
    For Each item In warnings
        result = result & item & vbCrLf
    Next item
    VATWarningText = result
End Function
