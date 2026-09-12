Attribute VB_Name = "modVATTests"
Option Explicit

Private assertions As Long
Private logText As String

Private Sub Expect(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 600, , "测试失败：" & message
    assertions = assertions + 1
    logText = logText & "PASS " & message & vbCrLf
End Sub

Private Sub Paint(ByVal cell As Range, ByVal color As Long)
    '仅用于本模块自己创建的临时测试工作簿。
    cell.Interior.Pattern = xlSolid
    cell.Interior.Color = color
End Sub

Private Sub TestSummaryAboveHeaders(ByVal sa As Worksheet, ByVal sb As Worksheet, ByVal green As Long)
    Dim foundA As Collection, foundB As Collection, issues As Collection, warnings As Collection
    Dim total As Variant, count As Long, included As Long, skipped As Long, ok As Boolean
    Dim ui As frmVATCheck, errorNumber As Long, errorText As String
    On Error GoTo Failed
    '只操作调用者自建的测试工作表；连上方合计也填完成色，证明扫描起点正确。
    sa.Range("I1").Formula = "=SUM(I4:I20)"
    sa.Range("I3").Value2 = VAT_FIELD_A
    sa.Range("I4").Value2 = 3580.5: sa.Range("I5").Value2 = 1200
    sa.Range("I6").Value2 = 999
    Paint sa.Range("I1,I4:I5"), green
    sb.Range("F1").Formula = "=SUM(F6:F20)"
    sb.Range("F5").Value2 = VAT_FIELD_B
    sb.Range("F6").Value2 = 4780.5: sb.Range("F7").Value2 = 999
    Paint sb.Range("F1,F6"), green
    sa.Calculate: sb.Calculate
    Set foundA = VATFindHeaders(sa, VAT_FIELD_A)
    Expect foundA.Count = 1, "首行 SUM：A 唯一找到非首行表头"
    Expect VATLocation(foundA(1)) = "第9列（I列），表头 I3", "首行 SUM：A 准确定位 I3 和第9列 I"
    Set issues = New Collection
    ok = VATScan(foundA(1), VAT_FIELD_A, "A", green, total, count, issues)
    Expect ok And total = CDec(4780.5) And count = 2, "首行 SUM：A 从 I4 扫描，完成色 I1 合计不参与"
    Set foundB = VATFindHeaders(sb, VAT_FIELD_B)
    Expect foundB.Count = 1, "首行 SUM：B 唯一找到非首行表头"
    Expect VATLocation(foundB(1)) = "第6列（F列），表头 F5", "首行 SUM：B 准确定位 F5 和第6列 F"
    Set issues = New Collection
    ok = VATScan(foundB(1), VAT_FIELD_B, "B", green, total, count, issues)
    Expect ok And total = CDec(4780.5) And count = 1, "首行 SUM：B 从 F6 扫描，完成色 F1 合计不参与"
    Expect InStr(VATResult(foundA(1), foundB(1), green), "核对一致") = 1, "首行 SUM：两边有效明细正常比较"
    Set ui = New frmVATCheck
    Load ui
    ui.BindWorkbook "A", sa.Parent: ui.BindWorkbook "B", sb.Parent
    ui.Controls("panelA").Controls("sheetA").ListIndex = 0
    ui.Controls("panelB").Controls("sheetB").ListIndex = 0
    Expect ui.Controls("panelA").Controls("locationA").Caption = "第9列（I列），表头 I3" And _
           ui.Controls("panelB").Controls("locationB").Caption = "第6列（F列），表头 F5", "首行 SUM：真实窗体选择 Sheet 后自动定位双表"
    Unload ui: Set ui = Nothing

    '第三场景：上方数字、公式、文本和空行混合，均不能成为核算记录。
    sa.Cells.Clear: sb.Cells.Clear
    sa.Range("I1").Formula = "=SUM(I8:I20)": sb.Range("F1").Formula = "=SUM(F8:F20)"
    sa.Range("I2").Value2 = 123: sb.Range("F2").Value2 = 456
    sa.Range("I3").Value2 = "说明文字": sb.Range("F3").Value2 = "说明文字"
    sa.Range("I5").Formula = "=""""": sb.Range("F5").Formula = "="""""
    sa.Range("I7").Value2 = VAT_FIELD_A: sb.Range("F7").Value2 = VAT_FIELD_B
    sa.Range("I8").Value2 = 0.1: sa.Range("I9").Value2 = 0.2: sb.Range("F8").Value2 = 0.3
    Paint sa.Range("I1:I5,I8:I9"), green: Paint sb.Range("F1:F5,F8"), green
    sa.Calculate: sb.Calculate
    Set foundA = VATFindHeaders(sa, VAT_FIELD_A): Set foundB = VATFindHeaders(sb, VAT_FIELD_B)
    Expect foundA.Count = 1 And foundA(1).Address(False, False) = "I7", "混合前置内容：A 仍准确找到 I7"
    Expect foundB.Count = 1 And foundB(1).Address(False, False) = "F7", "混合前置内容：B 仍准确找到 F7"
    Set issues = New Collection: Set warnings = New Collection
    ok = VATScan(foundA(1), VAT_FIELD_A, "A", green, total, count, issues, included, skipped, warnings)
    Expect ok And total = CDec(0.3) And count = 2 And included = 2 And skipped = 0 And warnings.Count = 0, "混合前置内容：A 上方数字文本空行公式全部不参与"
    Set issues = New Collection: Set warnings = New Collection
    ok = VATScan(foundB(1), VAT_FIELD_B, "B", green, total, count, issues, included, skipped, warnings)
    Expect ok And total = CDec(0.3) And count = 1 And included = 1 And skipped = 0 And warnings.Count = 0, "混合前置内容：B 上方数字文本空行公式全部不参与"
    Expect InStr(VATResult(foundA(1), foundB(1), green), "核对一致") = 1, "混合前置内容：Decimal 有效明细比较正常"
    sa.Cells.Clear: sb.Cells.Clear
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If Not ui Is Nothing Then Unload ui
    On Error GoTo 0
    Err.Raise errorNumber, "TestSummaryAboveHeaders", errorText
End Sub

Private Sub TestFormattedHeaders(ByVal sa As Worksheet, ByVal sb As Worksheet)
    Dim found As Collection, legacy As Range, area As Range, item As Variant, row As Long
    Dim numberFormat As String, ui As frmVATCheck, errorNumber As Long, errorText As String
    On Error GoTo Failed
    '复用真实 F2 的格式：显示文本带空白，但 Value2 不带空白。
    numberFormat = "_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * ""-""??_ ;_ @_ "
    sb.Range("F2").Value2 = VAT_FIELD_B
    sb.Range("F2").NumberFormat = numberFormat
    Expect sb.Range("F2").Value2 = VAT_FIELD_B And sb.Range("F2").Text <> VAT_FIELD_B, "格式回归：真实格式改变显示文本而不改变 Value2"
    Set area = sb.UsedRange
    Set legacy = area.Find(What:=VAT_FIELD_B, After:=area.Cells(area.Rows.Count, area.Columns.Count), _
                          LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, _
                          SearchDirection:=xlNext, MatchCase:=True, MatchByte:=True, SearchFormat:=False)
    Expect legacy Is Nothing, "格式回归：旧 xlValues + xlWhole 确实漏掉真实表头"
    Set found = VATFindHeaders(sb, VAT_FIELD_B)
    Expect found.Count = 1, "格式回归：新版找到格式化的精确 B 表头"
    Expect VATLocation(found(1)) = "第6列（F列），表头 F2", "格式回归：B 定位 F2 和第6列 F"
    row = 1
    For Each item In Array("税额合计", "应交税额", " 税额", "税额 ", ChrW(160) & "税额", "税额" & vbLf, "税" & ChrW(&H3000) & "额")
        sb.Cells(row, 8).Value2 = item: row = row + 1
    Next item
    Set found = VATFindHeaders(sb, VAT_FIELD_B)
    Expect found.Count = 1 And found(1).Address(False, False) = "F2", "格式回归：B 拒绝前后缀、真实空白、换行及不间断空格"
    sa.Range("I3").Value2 = VAT_FIELD_A: sa.Range("I3").NumberFormat = numberFormat
    Set found = VATFindHeaders(sa, VAT_FIELD_A)
    Expect found.Count = 1, "格式回归：格式化 A 表头的字面星号仍能匹配"
    Expect VATLocation(found(1)) = "第9列（I列），表头 I3", "格式回归：A 仍动态定位 I3"
    row = 1
    For Each item In Array(VAT_FIELD_A & "合计", "应交" & VAT_FIELD_A, " " & VAT_FIELD_A, VAT_FIELD_A & " ", _
                           "有效抵扣税额", "有效抵扣税额123", "有效抵扣税额" & ChrW(&HFF0A))
        sa.Cells(row, 12).Value2 = item: row = row + 1
    Next item
    Set found = VATFindHeaders(sa, VAT_FIELD_A)
    Expect found.Count = 1 And found(1).Address(False, False) = "I3", "格式回归：A 拒绝近似字段、真实空白及全角星号"
    Set ui = New frmVATCheck
    Load ui
    ui.BindWorkbook "A", sa.Parent: ui.BindWorkbook "B", sb.Parent
    ui.Controls("panelA").Controls("sheetA").ListIndex = 0
    ui.Controls("panelB").Controls("sheetB").ListIndex = 0
    Expect ui.Controls("panelA").Controls("locationA").Caption = "第9列（I列），表头 I3" And _
           ui.Controls("panelB").Controls("locationB").Caption = "第6列（F列），表头 F2", "格式回归：双栏自动定位仍沿用严格 Value2 匹配"
    Unload ui: Set ui = Nothing
    sa.Cells.Clear: sb.Cells.Clear
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If Not ui Is Nothing Then Unload ui
    On Error GoTo 0
    Err.Raise errorNumber, "TestFormattedHeaders", errorText
End Sub

Public Function VATCheck_SelfTest() As String
    Dim a As Workbook, b As Workbook, c As Workbook, reused As Workbook, sa As Worksheet, sb As Worksheet
    Dim green As Long, red As Long, white As Long, total As Variant, amount As Variant
    Dim count As Long, issues As Collection, found As Collection, reason As String, text As String
    Dim ok As Boolean, oldSaved As Boolean, formulaBefore As Variant, colorBefore As Long
    Dim errorMessage As String, i As Long, startTime As Single, ui As frmVATCheck
    Dim warnings As Collection, included As Long, skipped As Long
    Dim oldEvents As Boolean, oldSecurity As Long, fixtureDir As String, pathA As String, pathB As String, pathC As String
    Dim createdFiles As New Collection, file As Variant, bookCount As Long, openError As Long
    On Error GoTo Failed
    oldEvents = Application.EnableEvents
    assertions = 0: logText = vbNullString
    green = RGB(146, 208, 80): red = RGB(255, 0, 0): white = RGB(255, 255, 255)
    Set a = Application.Workbooks.Add(xlWBATWorksheet)
    Set b = Application.Workbooks.Add(xlWBATWorksheet)
    Set sa = a.Worksheets(1): Set sb = b.Worksheets(1)
    sa.Name = "测试A": sb.Name = "测试B"
    TestSummaryAboveHeaders sa, sb, green
    TestFormattedHeaders sa, sb
    sa.Range("K3").Value2 = VAT_FIELD_A
    sa.Range("L3").Value2 = "有效抵扣税额其他"
    sb.Range("AA5").Value2 = VAT_FIELD_B
    Set found = VATFindHeaders(sa, VAT_FIELD_A)
    Expect found.Count = 1, "表头星号精确匹配，不匹配后缀文字"
    Expect found(1).Column = 11 And found(1).Row = 3, "动态定位非首行 K3 表头"
    Expect VATColumnLetter(27) = "AA" And VATColumnLetter(16384) = "XFD", "多字母列定位"
    Expect VATLocation(found(1)) = "第11列（K列），表头 K3", "列号、字母和表头地址展示"
    sa.Range("M3").Value2 = VAT_FIELD_A: sa.Range("K10").Value2 = VAT_FIELD_A
    Set found = VATFindHeaders(sa, VAT_FIELD_A)
    Expect found.Count = 3, "不同列及同列不同行的重复表头都列出"
    sa.Range("M3,K10").ClearContents
    Set found = VATFindHeaders(sb, VAT_FIELD_A)
    Expect found.Count = 0, "缺少表头不猜测列"

    sa.Range("K4").Value2 = 0.1: Paint sa.Range("K4"), green
    sa.Range("K5").Value2 = 0.2: Paint sa.Range("K5"), green
    sb.Range("AA6").Value2 = 0.3: Paint sb.Range("AA6"), green
    sa.Range("K6").Value2 = 999
    sa.Range("K7").Value2 = 888: Paint sa.Range("K7"), red
    Paint sa.Range("J6"), green
    a.Saved = True: b.Saved = True
    formulaBefore = sa.Range("K4").Formula: colorBefore = sa.Range("K4").Interior.Color
    text = VATResult(sa.Range("K3"), sb.Range("AA5"), green)
    Expect InStr(text, "核对一致") = 1 And InStr(text, "差额（A - B）：¥0.00") > 0, "Decimal：0.1 + 0.2 与 0.3 精确相等"
    Expect InStr(text, "完成色行数：2行") > 0 And InStr(text, "完成色行数：1行") > 0, "记录数量不同仍按金额判断一致"
    Expect a.Saved And b.Saved, "核对不将工作簿标记为已修改"
    Expect sa.Range("K4").Formula = formulaBefore And sa.Range("K4").Interior.Color = colorBefore, "核对不改金额、公式或颜色"
    Expect InStr(text, "999") = 0 And InStr(text, "888") = 0, "无填充、其他颜色及相邻列颜色不参与"

    sb.Range("AA6").Value2 = 0.2
    Expect Not b.Saved, "测试修改保持未保存"
    text = VATResult(sa.Range("K3"), sb.Range("AA5"), green)
    Expect InStr(text, "A 比 B 多 ¥0.10") > 0, "读取未保存金额并显示 A 较多方向"
    Expect Not b.Saved, "核对不会自动保存"
    sb.Range("AA6").Value2 = 1
    text = VATResult(sa.Range("K3"), sb.Range("AA5"), green)
    Expect InStr(text, "B 比 A 多 ¥0.70") > 0, "显示 B 较多方向"
    Paint sa.Range("K5"), red
    Set issues = New Collection
    ok = VATScan(sa.Range("K3"), VAT_FIELD_A, "A", green, total, count, issues)
    Expect ok And total = CDec(0.1) And count = 1, "读取未保存颜色修改"
    Paint sa.Range("K5"), green
    sa.Rows(5).Hidden = True
    Set issues = New Collection
    ok = VATScan(sa.Range("K3"), VAT_FIELD_A, "A", green, total, count, issues)
    Expect ok And total = CDec(0.3) And count = 2, "隐藏行照常参与"
    sa.Rows(5).Hidden = False

    Paint sa.Range("K20"), green
    text = VATResult(sa.Range("K3"), sb.Range("AA5"), green)
    Expect InStr(text, "有效金额部分核对不一致") = 1 And InStr(text, "!K20：空金额，已跳过未计入") > 0 And _
           InStr(text, "B 比 A 多 ¥0.70") > 0, "填色空白尾行警告并跳过，有效金额继续比较"
    Set warnings = New Collection: Set issues = New Collection
    ok = VATScan(sa.Range("K3"), VAT_FIELD_A, "A", green, total, count, issues, included, skipped, warnings)
    Expect ok And total = CDec(0.3) And count = 3 And included = 2 And skipped = 1 And issues.Count = 0 And warnings.Count = 1, "空金额分别统计完成色、实际计入、未计入行数"
    Expect InStr(text, "完成色行数：3行；实际计入核算行数：2行；空金额未计入行数：1行") > 0, "结果完整展示三类行数"
    Paint sa.Range("K21"), green: Paint sb.Range("AA20"), green
    sb.Range("AA6").Value2 = 0.3
    text = VATResult(sa.Range("K3"), sb.Range("AA5"), green)
    Expect InStr(text, "有效金额部分核对一致") = 1 And InStr(text, "差额（A - B）：¥0.00") > 0, "空金额警告不影响有效金额精确相等"
    Expect InStr(text, "!K20：") > 0 And InStr(text, "!K21：") > 0 And InStr(text, "!AA20：") > 0 And _
           InStr(text, "以下 3 个") > 0, "列出两边全部被跳过的空金额位置"
    sa.Range("K21").Interior.Pattern = xlPatternNone: sb.Range("AA20").Interior.Pattern = xlPatternNone
    sa.Range("K4").Value2 = True
    text = VATResult(sa.Range("K3"), sb.Range("AA5"), green)
    Expect InStr(text, "本次无法完成可靠核对") = 1 And InStr(text, "!K4：") > 0 And InStr(text, "!K20：") > 0, "空金额警告与布尔值并存仍阻断，并保留警告位置"
    sa.Range("K4").Value2 = 1E+30
    text = VATResult(sa.Range("K3"), sb.Range("AA5"), green)
    Expect InStr(text, "本次无法完成可靠核对") = 1 And InStr(text, "Decimal") > 0, "空金额警告不会放宽 Decimal 转换失败"
    sa.Range("K4").Value2 = 0.1
    sa.Range("K20").Interior.Pattern = xlPatternNone
    sa.Range("K4").NumberFormat = "@": sa.Range("K4").Value2 = "0.10"
    sb.Range("AA6").Value2 = CVErr(xlErrDiv0)
    text = VATResult(sa.Range("K3"), sb.Range("AA5"), green)
    Expect InStr(text, "表A [") > 0 And InStr(text, "!K4：金额为文本") > 0, "数字文本也必须阻止核对"
    Expect InStr(text, "表B [") > 0 And InStr(text, "!AA6：金额为 Excel 错误值") > 0, "两边异常一起显示准确单元格"
    sa.Range("K4").NumberFormat = "General"
    sa.Range("K4").Formula = "=" & Chr$(34) & Chr$(34)
    sa.Calculate
    text = VATResult(sa.Range("K3"), sb.Range("AA5"), green)
    Expect InStr(text, "公式返回空字符串") > 0, "公式空字符串作为空金额报错"
    sa.Range("K4").Formula = "=0.1+0.2": sa.Calculate
    Set issues = New Collection
    ok = VATScan(sa.Range("K3"), VAT_FIELD_A, "A", green, total, count, issues)
    Expect ok And total = CDec(0.5), "数值公式结果转 Decimal 后累加"
    sa.Range("K4").Value2 = -5.25: sa.Range("K5").Value2 = 0
    Set issues = New Collection
    ok = VATScan(sa.Range("K3"), VAT_FIELD_A, "A", green, total, count, issues)
    Expect ok And total = CDec(-5.25) And count = 2, "负数及零金额参与"

    reason = vbNullString
    Expect Not VATTryAmount(True, amount, reason), "布尔值不是金额"
    reason = vbNullString
    Expect Not VATTryAmount(1E+30, amount, reason), "超出 Decimal 范围明确失败"
    Expect VATMoney(CDec(128563.47)) = "¥128,563.47", "财务金额分组显示"
    Expect VATMoney(CDec(-0.001)) = "¥-0.001", "保留分以下差额，避免显示零差额却不一致"
    Expect VATMoney(CDec("79228162514264337593543950335")) = "¥79,228,162,514,264,337,593,543,950,335.00", "大 Decimal 金额显示不经过浮点格式化"

    reason = vbNullString
    Expect Not VATReadColor(sa.Range("K6"), i, reason), "无填充不能作为完成颜色样本"
    Paint sa.Range("K4"), white
    reason = vbNullString
    Expect VATReadColor(sa.Range("K4"), i, reason) And i = white, "白色填充与无填充分开"
    Set issues = New Collection
    ok = VATScan(sa.Range("K3"), VAT_FIELD_A, "A", white, total, count, issues)
    Expect ok And count = 1 And total = CDec(-5.25), "选择白色时不计入无填充记录"
    reason = vbNullString
    Expect Not VATReadColor(sa.Range("K4:K5"), i, reason), "多单元格颜色样本明确拒绝"
    With sa.Range("K6").FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
        .Interior.Color = green
    End With
    Set issues = New Collection
    ok = VATScan(sa.Range("K3"), VAT_FIELD_A, "A", green, total, count, issues)
    Expect ok And total = CDec(999), "条件格式按当前显示颜色匹配"
    sa.Range("K6").FormatConditions.Delete

    sa.Range("K3").Value2 = "新表头"
    text = VATResult(sa.Range("K3"), sb.Range("AA5"), green)
    Expect InStr(text, "表头已改变") > 0, "表头被修改后不使用旧列输出一致"
    sa.Range("K3").Value2 = VAT_FIELD_A
    sa.Range("K4:K5").Merge
    Paint sa.Range("K4"), green
    text = VATResult(sa.Range("K3"), sb.Range("AA5"), green)
    Expect InStr(text, "被合并") > 0, "已完成金额合并单元格阻止核对"
    sa.Range("K4:K5").UnMerge
    sa.Range("K4:K7").Interior.Pattern = xlPatternNone
    sb.Range("AA6").Interior.Pattern = xlPatternNone
    text = VATResult(sa.Range("K3"), sb.Range("AA5"), green)
    Expect InStr(text, "两边均无已完成记录") > 0, "双方零记录明确提示"
    Paint sa.Range("K20"), green: Paint sb.Range("AA20"), green
    text = VATResult(sa.Range("K3"), sb.Range("AA5"), green)
    Expect InStr(text, "有效金额部分核对一致") = 1 And InStr(text, "实际计入核算行数：0行") > 0, "只有空金额时明确部分结论且实际计入零行"
    sa.Range("K20").Interior.Pattern = xlPatternNone: sb.Range("AA20").Interior.Pattern = xlPatternNone

    '直接操作真实 MSForms 控件，验证下拉联动和重复表头选择状态。
    sa.Range("M3").Value2 = VAT_FIELD_A
    Set ui = New frmVATCheck
    Load ui
    ui.BindWorkbook "A", a
    ui.BindWorkbook "B", b
    Expect ui.Controls("panelA").Left < ui.Controls("panelB").Left And ui.Controls("panelA").Top = ui.Controls("panelB").Top, "真实窗体 A/B 同屏左右排列"
    Expect ui.Controls("panelA").Controls("sheetA").ListCount = 1, "选择 Excel 后加载真实 Sheet 列表"
    ui.Controls("panelA").Controls("sheetA").ListIndex = 0
    ui.Controls("panelB").Controls("sheetB").ListIndex = 0
    Expect ui.Controls("panelA").Controls("fieldA").ListCount = 2 And ui.Controls("panelA").Controls("fieldA").ListIndex = -1, "重复表头必须手动选择，不默认猜测"
    Expect ui.Controls("panelB").Controls("fieldB").ListIndex = 0, "唯一表头自动定位"
    ui.Controls("panelA").Controls("fieldA").ListIndex = 1
    Expect InStr(ui.Controls("panelA").Controls("locationA").Caption, "第13列（M列）") = 1, "手选重复表头更新实际列号"
    Unload ui
    Set ui = Nothing

    sa.Range("M3").ClearContents
    sa.Range("K4:K10003").Value2 = 0.01
    Paint sa.Range("K4:K10003"), green
    Set issues = New Collection
    startTime = Timer
    ok = VATScan(sa.Range("K3"), VAT_FIELD_A, "A", green, total, count, issues)
    Expect ok And total = CDec(100) And count = 10000, "一万行 0.01 精确累加为 100.00"
    logText = logText & "INFO 一万行扫描用时 " & Format$(Timer - startTime, "0.00") & " 秒" & vbCrLf

    '在自建临时目录里验证真实文件打开/复用/关闭事件；不触碰用户业务文件。
    fixtureDir = Environ$("TEMP") & "\VATCheck-" & Format$(Now, "yyyymmddhhnnss") & "-" & CStr(CLng(Timer * 100))
    MkDir fixtureDir
    pathA = fixtureDir & "\旧表A.xlsx": pathB = fixtureDir & "\表B.xlsx": pathC = fixtureDir & "\新表A.xlsx"
    sb.Range("AA6").Value2 = 100: Paint sb.Range("AA6"), green
    a.SaveAs Filename:=pathA, FileFormat:=xlOpenXMLWorkbook: createdFiles.Add pathA
    b.SaveAs Filename:=pathB, FileFormat:=xlOpenXMLWorkbook: createdFiles.Add pathB
    Set c = Application.Workbooks.Add(xlWBATWorksheet)
    c.Worksheets(1).Name = "重新选择A"
    c.Worksheets(1).Range("R7").Value2 = VAT_FIELD_A
    c.Worksheets(1).Range("R8").Value2 = 100.01: Paint c.Worksheets(1).Range("R8"), green
    c.SaveAs Filename:=pathC, FileFormat:=xlOpenXMLWorkbook: createdFiles.Add pathC
    c.Close SaveChanges:=False: Set c = Nothing
    Set ui = New frmVATCheck
    Load ui
    Application.EnableEvents = True
    ui.BindFile "A", pathA: ui.BindFile "B", pathB
    ui.Controls("panelA").Controls("sheetA").ListIndex = 0
    ui.Controls("panelB").Controls("sheetB").ListIndex = 0
    VATActivateHost
    ui.Show vbModeless
    Expect ui.Visible, "文件切换前工具窗体非模态保持可见"
    a.Activate: sa.Range("K4").Select
    ui.Controls("panelA").Controls("pickA").Value = True
    Expect ui.Controls("panelA").Controls("colorA").BackColor = green And _
           ui.Controls("panelB").Controls("colorB").BackColor = green, "取色按钮事件保持两表共享完成颜色"
    sa.Range("K4").Value2 = 0.02
    sb.Range("AA6").Value2 = 100.01
    bookCount = Application.Workbooks.Count
    Set reused = VATOpenWorkbook(UCase$(pathA))
    Expect reused Is a, "同一路径忽略大小写复用当前 Workbook"
    Expect Application.Workbooks.Count = bookCount And Not a.Saved And sa.Range("K4").Value2 = 0.02, "文件复用不重复打开且保留未保存修改"
    ui.BindFile "A", pathA
    ui.Controls("panelA").Controls("sheetA").ListIndex = 0
    Expect ui.Controls("panelB").Controls("fieldB").ListIndex = 0, "重新绑定 A 不清空 B 的 Sheet 或字段"
    ui.Controls("checkButton").Value = True
    Expect InStr(ui.Controls("resultBox").Text, "核对一致") = 1 And _
           InStr(ui.Controls("resultBox").Text, "¥100.01") > 0, "文件绑定后按钮核对读取两边未保存金额"
    a.Close SaveChanges:=False
    Set a = Nothing: Set reused = Nothing
    Expect ui.Visible, "关闭原业务 Workbook 后工具窗体继续存在"
    Expect ui.Controls("panelA").Controls("excelA").Value = "" And _
           ui.Controls("panelA").Controls("sheetA").ListCount = 0 And _
           ui.Controls("panelA").Controls("fieldA").ListCount = 0 And _
           InStr(ui.Controls("panelA").Controls("locationA").Caption, "失效") > 0, "关闭 A 后只清空 A 的文件、Sheet、字段和定位"
    Expect ui.Controls("panelB").Controls("excelB").Value = b.Name And _
           ui.Controls("panelB").Controls("sheetB").ListIndex = 0 And _
           ui.Controls("panelB").Controls("fieldB").ListIndex = 0 And Not b.Saved, "关闭 A 保留 B 的绑定、定位和未保存状态"
    ui.Controls("checkButton").Value = False: ui.Controls("checkButton").Value = True
    Expect InStr(ui.Controls("resultBox").Text, "请先") = 1, "关闭一侧后不能沿用旧表头输出结果"
    bookCount = Application.Workbooks.Count
    oldSecurity = Application.AutomationSecurity
    ui.BindFile "A", pathC
    Set c = VATOpenWorkbook(pathC)
    Expect Application.Workbooks.Count = bookCount + 1 And Application.AutomationSecurity = oldSecurity, "新文件只打开一次并恢复原宏打开设置"
    Expect ui.Controls("panelA").Controls("sheetA").List(0) = "重新选择A", "重新选择新 Excel 文件后加载新 Sheet"
    ui.Controls("panelA").Controls("sheetA").ListIndex = 0
    Expect InStr(ui.Controls("panelA").Controls("locationA").Caption, "第18列（R列），表头 R7") = 1, "重新绑定后按新表头动态定位 R7"
    ui.Controls("checkButton").Value = False: ui.Controls("checkButton").Value = True
    Expect InStr(ui.Controls("resultBox").Text, "核对一致") = 1 And InStr(ui.Controls("resultBox").Text, "¥100.01") > 0, "关闭旧文件到重新选择新文件后可以继续核对"
    On Error Resume Next
    ui.BindFile "A", fixtureDir & "\不存在.xlsx"
    openError = Err.Number
    On Error GoTo Failed
    Expect openError <> 0 And ui.Controls("panelA").Controls("fieldA").ListIndex = 0 And _
           ui.Controls("panelB").Controls("fieldB").ListIndex = 0 And Application.AutomationSecurity = oldSecurity, "打开失败保留原绑定和另一侧状态并恢复设置"
    ui.BindFile "A", pathB
    Expect VATWorkbookOpen(c) And VATWorkbookOpen(b) And Not b.Saved, "切换文件不自动关闭或保存原业务 Workbook"
    ui.BindFile "A", pathC
    ui.Controls("panelA").Controls("sheetA").ListIndex = 0
    b.Close SaveChanges:=False
    Expect Not VATWorkbookOpen(b), "识别原工作簿已关闭"
    Set b = Nothing
    Expect ui.Controls("panelB").Controls("sheetB").ListCount = 0 And _
           ui.Controls("panelA").Controls("fieldA").ListIndex = 0 And ui.Visible, "关闭 B 同样只使 B 失效且 A 和窗体继续可用"
    Unload ui: Set ui = Nothing
    c.Close SaveChanges:=False: Set c = Nothing
    Application.EnableEvents = oldEvents
    For Each file In createdFiles
        Kill CStr(file)
    Next file
    RmDir fixtureDir
    VATCheck_SelfTest = "PASS: " & assertions & " assertions" & vbCrLf & logText
    Exit Function
Failed:
    errorMessage = Err.Description
    On Error Resume Next
    If Not ui Is Nothing Then Unload ui
    If Not a Is Nothing Then a.Close SaveChanges:=False
    If Not b Is Nothing Then b.Close SaveChanges:=False
    If Not c Is Nothing Then c.Close SaveChanges:=False
    Application.EnableEvents = oldEvents
    For Each file In createdFiles
        Kill CStr(file)
    Next file
    If Len(fixtureDir) > 0 Then RmDir fixtureDir
    On Error GoTo 0
    VATCheck_SelfTest = "FAIL: " & errorMessage & vbCrLf & logText
End Function
