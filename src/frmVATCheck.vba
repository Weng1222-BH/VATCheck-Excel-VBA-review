Option Explicit

Private excelA As MSForms.TextBox, excelB As MSForms.TextBox
Private WithEvents chooseA As MSForms.CommandButton
Private WithEvents chooseB As MSForms.CommandButton
Private WithEvents excelEvents As Excel.Application
Private WithEvents sheetA As MSForms.ComboBox
Private WithEvents sheetB As MSForms.ComboBox
Private WithEvents fieldA As MSForms.ComboBox
Private WithEvents fieldB As MSForms.ComboBox
Private WithEvents pickA As MSForms.CommandButton
Private WithEvents pickB As MSForms.CommandButton
Private WithEvents refreshButton As MSForms.CommandButton
Private WithEvents checkButton As MSForms.CommandButton
Private resultBox As MSForms.TextBox
Private locationA As MSForms.Label, locationB As MSForms.Label
Private colorA As MSForms.Label, colorB As MSForms.Label
Private bookA As Workbook, bookB As Workbook
Private sheetsA As Collection, sheetsB As Collection
Private headersA As Collection, headersB As Collection
Private headerA As Range, headerB As Range
Private completedColor As Long, hasColor As Boolean, updating As Boolean

Private Sub UserForm_Initialize()
    On Error GoTo Failed
    Me.Caption = "增值税专票双表一键核对"
    Me.Width = 712
    Me.Height = 638
    Me.BackColor = RGB(245, 247, 250)
    Me.Font.Name = "Microsoft YaHei UI"
    Me.Font.Size = 10
    MakeLabel Me, "heading", "增值税专票 · 双表核对", 20, 12, 470, 25, 16, True
    MakeLabel Me, "hint", "选择两份 Excel 文件；已打开文件直接复用（含未保存修改），两表共用完成颜色。", 20, 43, 660, 22, 10
    BuildPanel "A", 20, VAT_FIELD_A
    BuildPanel "B", 364, VAT_FIELD_B
    Set refreshButton = MakeButton(Me, "refreshButton", "刷新 Sheet 列表", 20, 350, 183, 28)
    Set checkButton = MakeButton(Me, "checkButton", "一键核对", 274, 346, 150, 36)
    checkButton.BackColor = RGB(32, 102, 180)
    checkButton.ForeColor = vbWhite
    checkButton.Font.Bold = True
    MakeLabel Me, "resultTitle", "核对结果", 20, 392, 110, 20, 12, True
    Set resultBox = Me.Controls.Add("Forms.TextBox.1", "resultBox", True)
    With resultBox
        .Left = 20: .Top = 418: .Width = 668: .Height = 160
        .MultiLine = True: .WordWrap = True: .ScrollBars = fmScrollBarsVertical
        .Locked = True: .TabStop = True: .MaxLength = 0
        .Font.Name = "Microsoft YaHei UI": .Font.Size = 11
        .BackColor = vbWhite: .BorderStyle = fmBorderStyleSingle
    End With
    MakeLabel Me, "footer", "只读核对 · 不保存或关闭工作簿 · 修改后需再次点击核对", 20, 586, 668, 18, 9
    Set sheetsA = New Collection: Set sheetsB = New Collection
    Set headersA = New Collection: Set headersB = New Collection
    Set excelEvents = Application
    InvalidateResult
    Exit Sub
Failed:
    MsgBox "界面初始化失败：" & Err.Description, vbExclamation
End Sub

Private Sub BuildPanel(ByVal side As String, ByVal left As Single, ByVal field As String)
    Dim panel As MSForms.Frame, bookBox As MSForms.TextBox, sheetBox As MSForms.ComboBox
    Dim fieldBox As MSForms.ComboBox, location As MSForms.Label, swatch As MSForms.Label
    Dim button As MSForms.CommandButton, chooseButton As MSForms.CommandButton
    Set panel = Me.Controls.Add("Forms.Frame.1", "panel" & side, True)
    With panel
        .Caption = "表 " & side: .Left = left: .Top = 74: .Width = 324: .Height = 261
        .Font.Name = "Microsoft YaHei UI": .Font.Size = 12: .Font.Bold = True
        .BackColor = vbWhite
    End With
    MakeLabel panel, "excelLabel" & side, "Excel", 12, 20, 70, 18, 10
    Set bookBox = panel.Controls.Add("Forms.TextBox.1", "excel" & side, True)
    With bookBox
        .Left = 12: .Top = 40: .Width = 188: .Height = 24
        .Locked = True: .Font.Name = "Microsoft YaHei UI": .Font.Size = 10
    End With
    Set chooseButton = MakeButton(panel, "choose" & side, "选择 Excel…", 206, 39, 104, 26)
    MakeLabel panel, "sheetLabel" & side, "Sheet", 12, 71, 70, 18, 10
    Set sheetBox = MakeCombo(panel, "sheet" & side, 12, 91, 298)
    MakeLabel panel, "fieldLabel" & side, "目标字段：" & field, 12, 122, 298, 18, 10
    Set fieldBox = MakeCombo(panel, "field" & side, 12, 142, 298)
    Set location = MakeLabel(panel, "location" & side, "定位结果：请先选择 Sheet", 12, 173, 298, 22, 10)
    MakeLabel panel, "colorLabel" & side, "已完成颜色（A / B 共用）", 12, 200, 240, 18, 10
    Set swatch = MakeLabel(panel, "color" & side, "未选择", 12, 222, 118, 24, 10)
    swatch.BackStyle = fmBackStyleOpaque
    swatch.BackColor = RGB(237, 240, 244)
    swatch.TextAlign = fmTextAlignCenter
    Set button = MakeButton(panel, "pick" & side, "从选中单元格读取", 143, 219, 167, 27)
    If side = "A" Then
        Set excelA = bookBox: Set sheetA = sheetBox: Set fieldA = fieldBox
        Set locationA = location: Set colorA = swatch: Set pickA = button
        Set chooseA = chooseButton
    Else
        Set excelB = bookBox: Set sheetB = sheetBox: Set fieldB = fieldBox
        Set locationB = location: Set colorB = swatch: Set pickB = button
        Set chooseB = chooseButton
    End If
End Sub

Private Function MakeLabel(ByVal parent As Object, ByVal name As String, ByVal caption As String, _
                           ByVal left As Single, ByVal top As Single, ByVal width As Single, _
                           ByVal height As Single, ByVal size As Single, Optional ByVal bold As Boolean = False) As MSForms.Label
    Dim control As MSForms.Label
    Set control = parent.Controls.Add("Forms.Label.1", name, True)
    With control
        .Caption = caption: .Left = left: .Top = top: .Width = width: .Height = height
        .Font.Name = "Microsoft YaHei UI": .Font.Size = size: .Font.Bold = bold
        .BackStyle = fmBackStyleTransparent
    End With
    Set MakeLabel = control
End Function

Private Function MakeCombo(ByVal parent As Object, ByVal name As String, ByVal left As Single, _
                           ByVal top As Single, ByVal width As Single) As MSForms.ComboBox
    Dim control As MSForms.ComboBox
    Set control = parent.Controls.Add("Forms.ComboBox.1", name, True)
    With control
        .Left = left: .Top = top: .Width = width: .Height = 24
        .Font.Name = "Microsoft YaHei UI": .Font.Size = 10
        .Style = fmStyleDropDownList: .ListRows = 12
    End With
    Set MakeCombo = control
End Function

Private Function MakeButton(ByVal parent As Object, ByVal name As String, ByVal caption As String, _
                            ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single) As MSForms.CommandButton
    Dim control As MSForms.CommandButton
    Set control = parent.Controls.Add("Forms.CommandButton.1", name, True)
    With control
        .Caption = caption: .Left = left: .Top = top: .Width = width: .Height = height
        .Font.Name = "Microsoft YaHei UI": .Font.Size = 10
    End With
    Set MakeButton = control
End Function

Private Sub InvalidateResult()
    If Not resultBox Is Nothing Then resultBox.Text = "尚未核对。请选择 Excel、Sheet、目标字段和完成颜色，然后点击【一键核对】。"
End Sub

Private Sub ClearBinding(ByVal side As String)
    updating = True
    If side = "A" Then
        Set bookA = Nothing: Set headerA = Nothing
        Set sheetsA = New Collection: Set headersA = New Collection
        excelA.Value = vbNullString: excelA.ControlTipText = vbNullString
        sheetA.Clear: fieldA.Clear
        locationA.Caption = "文件已失效，请重新选择 Excel"
    Else
        Set bookB = Nothing: Set headerB = Nothing
        Set sheetsB = New Collection: Set headersB = New Collection
        excelB.Value = vbNullString: excelB.ControlTipText = vbNullString
        sheetB.Clear: fieldB.Clear
        locationB.Caption = "文件已失效，请重新选择 Excel"
    End If
    updating = False
    InvalidateResult
End Sub

Private Sub ValidateBindings()
    '事件被其他宏禁用时，用户再次操作窗体也会识别已关闭文件。
    If Not bookA Is Nothing Then
        If Not VATWorkbookOpen(bookA) Then ClearBinding "A"
    End If
    If Not bookB Is Nothing Then
        If Not VATWorkbookOpen(bookB) Then ClearBinding "B"
    End If
End Sub

Public Sub BindWorkbook(ByVal side As String, ByVal book As Workbook)
    If side <> "A" And side <> "B" Then Err.Raise vbObjectError + 203, , "未知表侧。"
    If Not VATWorkbookOpen(book) Then Err.Raise vbObjectError + 204, , "工作簿已关闭，请重新选择 Excel。"
    If book Is ThisWorkbook Then Err.Raise vbObjectError + 205, , "请选择业务文件，不要选择工具自身。"
    If book.IsAddin Or UCase$(book.Name) = "PERSONAL.XLSB" Then Err.Raise vbObjectError + 206, , "请选择业务文件。"
    If side = "A" Then
        Set bookA = book
        excelA.Value = book.Name: excelA.ControlTipText = book.FullName
    Else
        Set bookB = book
        excelB.Value = book.Name: excelB.ControlTipText = book.FullName
    End If
    LoadSheets side
End Sub

Public Sub BindFile(ByVal side As String, ByVal filePath As String)
    Dim book As Workbook
    ValidateBindings
    Set book = VATOpenWorkbook(filePath)
    BindWorkbook side, book
End Sub

Private Sub ChooseFile(ByVal side As String)
    Dim picker As FileDialog
    On Error GoTo Failed
    ValidateBindings
    Set picker = Application.FileDialog(msoFileDialogFilePicker)
    With picker
        .Title = "选择表 " & side & " 的 Excel 文件"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel 工作簿", "*.xlsx;*.xlsm;*.xlsb;*.xls"
        If .Show <> -1 Then Exit Sub
        BindFile side, .SelectedItems(1)
    End With
    Exit Sub
Failed:
    MsgBox "表" & side & "：无法选择或打开文件。" & Err.Description, vbExclamation
End Sub

Private Sub excelEvents_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    Dim affected As Boolean
    If Cancel Then Exit Sub
    '仅使将关闭文件对应的侧失效，不保存、不关闭，也不改变 Cancel。
    If Not bookA Is Nothing Then
        If Wb Is bookA Then
            ClearBinding "A": affected = True
        End If
    End If
    If Not bookB Is Nothing Then
        If Wb Is bookB Then
            ClearBinding "B": affected = True
        End If
    End If
    If affected Then VATActivateHost
End Sub

Private Sub UserForm_Activate()
    ValidateBindings
End Sub

Private Sub UserForm_Terminate()
    Set excelEvents = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '先断开应用事件，避免等到 Terminate 才处理造成循环引用。
    Set excelEvents = Nothing
    Set bookA = Nothing: Set bookB = Nothing
    Set headerA = Nothing: Set headerB = Nothing
End Sub

Private Sub LoadSheets(ByVal side As String)
    Dim sheetBox As MSForms.ComboBox, fieldBox As MSForms.ComboBox
    Dim book As Workbook, sheet As Object, entries As New Collection
    On Error GoTo Failed
    If updating Then Exit Sub
    If side = "A" Then
        Set book = bookA: Set sheetBox = sheetA: Set fieldBox = fieldA
        Set headerA = Nothing: Set headersA = New Collection
        locationA.Caption = "定位结果：请先选择 Sheet"
    Else
        Set book = bookB: Set sheetBox = sheetB: Set fieldBox = fieldB
        Set headerB = Nothing: Set headersB = New Collection
        locationB.Caption = "定位结果：请先选择 Sheet"
    End If
    updating = True
    sheetBox.Clear: fieldBox.Clear
    If Not book Is Nothing Then
        If Not VATWorkbookOpen(book) Then Err.Raise vbObjectError + 200, , "工作簿已关闭，请刷新列表。"
        For Each sheet In book.Sheets
            entries.Add sheet
            If TypeOf sheet Is Worksheet Then
                sheetBox.AddItem sheet.Name
            Else
                sheetBox.AddItem sheet.Name & "（非工作表，不能核对）"
            End If
        Next sheet
    End If
    If side = "A" Then
        Set sheetsA = entries
    Else
        Set sheetsB = entries
    End If
    updating = False
    InvalidateResult
    Exit Sub
Failed:
    updating = False
    InvalidateResult
    MsgBox "表" & side & "：" & Err.Description, vbExclamation
End Sub

Private Sub LoadFields(ByVal side As String)
    Dim sheetBox As MSForms.ComboBox, fieldBox As MSForms.ComboBox, location As MSForms.Label
    Dim entries As Collection, found As Collection, sheet As Object, cell As Range, field As String
    On Error GoTo Failed
    If updating Then Exit Sub
    If side = "A" Then
        Set sheetBox = sheetA: Set fieldBox = fieldA: Set entries = sheetsA: Set location = locationA
        Set headerA = Nothing: Set headersA = New Collection: field = VAT_FIELD_A
    Else
        Set sheetBox = sheetB: Set fieldBox = fieldB: Set entries = sheetsB: Set location = locationB
        Set headerB = Nothing: Set headersB = New Collection: field = VAT_FIELD_B
    End If
    InvalidateResult
    updating = True
    fieldBox.Clear
    location.Caption = "定位结果：未定位"
    If sheetBox.ListIndex < 0 Then GoTo Done
    Set sheet = entries(sheetBox.ListIndex + 1)
    If Not VATWorkbookOpen(sheet.Parent) Then Err.Raise vbObjectError + 201, , "工作簿已关闭，请刷新列表。"
    If Not TypeOf sheet Is Worksheet Then Err.Raise vbObjectError + 202, , "此 Sheet 不是工作表，请选择数据工作表。"
    Set found = VATFindHeaders(sheet, field)
    If side = "A" Then
        Set headersA = found
    Else
        Set headersB = found
    End If
    For Each cell In found
        fieldBox.AddItem field & " | " & cell.Address(False, False) & " | 第" & cell.Column & "列（" & VATColumnLetter(cell.Column) & "）"
    Next cell
    If found.Count = 0 Then
        location.Caption = "未找到完全相同的表头"
        MsgBox "表" & side & " / " & sheet.Name & "：未找到完全相同的表头【" & field & "】。请检查 Sheet、空格及星号。", vbExclamation
    ElseIf found.Count = 1 Then
        fieldBox.ListIndex = 0
    Else
        location.Caption = "发现" & found.Count & "个同名表头，请明确选择"
    End If
Done:
    updating = False
    SelectHeader side
    Exit Sub
Failed:
    updating = False
    MsgBox "表" & side & "：" & Err.Description, vbExclamation
End Sub

Private Sub SelectHeader(ByVal side As String)
    If updating Then Exit Sub
    On Error GoTo Failed
    InvalidateResult
    If side = "A" Then
        Set headerA = Nothing
        If fieldA.ListIndex >= 0 Then
            Set headerA = headersA(fieldA.ListIndex + 1)
            locationA.Caption = VATLocation(headerA)
        End If
    Else
        Set headerB = Nothing
        If fieldB.ListIndex >= 0 Then
            Set headerB = headersB(fieldB.ListIndex + 1)
            locationB.Caption = VATLocation(headerB)
        End If
    End If
    Exit Sub
Failed:
    MsgBox "表头位置已失效，请重新选择 Sheet。", vbExclamation
End Sub

Private Sub ReadSelectedColor()
    Dim sample As Range, reason As String, newColor As Long, foreground As Long
    On Error GoTo Failed
    If TypeName(Application.Selection) <> "Range" Then
        MsgBox "请先在 Excel 中选中一个正确填色的单元格，再点击读取按钮。窗体可以保持打开。", vbInformation
        Exit Sub
    End If
    Set sample = Application.Selection
    If Not VATReadColor(sample, newColor, reason) Then
        MsgBox reason, vbExclamation
        Exit Sub
    End If
    completedColor = newColor: hasColor = True
    foreground = vbBlack
    If ((newColor Mod 256) * 299# + ((newColor \ 256) Mod 256) * 587# + ((newColor \ 65536) Mod 256) * 114#) < 128000# Then foreground = vbWhite
    colorA.BackColor = newColor: colorB.BackColor = newColor
    colorA.ForeColor = foreground: colorB.ForeColor = foreground
    colorA.Caption = "已选择（两表共用）": colorB.Caption = colorA.Caption
    colorA.ControlTipText = sample.Address(External:=True)
    colorB.ControlTipText = colorA.ControlTipText
    InvalidateResult
    Exit Sub
Failed:
    MsgBox "无法读取颜色：" & Err.Description, vbExclamation
End Sub

Private Sub checkButton_Click()
    On Error GoTo Failed
    ValidateBindings
    InvalidateResult
    If headerA Is Nothing Or headerB Is Nothing Then
        resultBox.Text = "请先为表 A、表 B 分别选择 Excel、Sheet 和唯一目标表头。"
        Exit Sub
    End If
    If Not hasColor Then
        resultBox.Text = "请先选中 Excel 中已正确填色的一个单元格，再点击【从选中单元格读取】。"
        Exit Sub
    End If
    checkButton.Enabled = False
    resultBox.Text = "正在核对，请稍候……"
    Me.Repaint
    '核对仅在点击时执行；关闭事件只负责文件绑定失效。
    resultBox.Text = VATResult(headerA, headerB, completedColor)
    checkButton.Enabled = True
    On Error Resume Next
    locationA.Caption = VATLocation(headerA)
    locationB.Caption = VATLocation(headerB)
    On Error GoTo 0
    Exit Sub
Failed:
    checkButton.Enabled = True
    resultBox.Text = "本次无法完成可靠核对：" & Err.Description
End Sub

Private Sub chooseA_Click()
    ChooseFile "A"
End Sub
Private Sub chooseB_Click()
    ChooseFile "B"
End Sub
Private Sub sheetA_Change()
    LoadFields "A"
End Sub
Private Sub sheetB_Change()
    LoadFields "B"
End Sub
Private Sub fieldA_Change()
    SelectHeader "A"
End Sub
Private Sub fieldB_Change()
    SelectHeader "B"
End Sub
Private Sub pickA_Click()
    ReadSelectedColor
End Sub
Private Sub pickB_Click()
    ReadSelectedColor
End Sub
Private Sub refreshButton_Click()
    ValidateBindings
    If Not bookA Is Nothing Then LoadSheets "A"
    If Not bookB Is Nothing Then LoadSheets "B"
End Sub
