Attribute VB_Name = "modVATStage2ExcelSnapshot"
Option Explicit

Public Enum VATS2SnapshotStatus
    VATS2_SNAPSHOT_OK = 0
    VATS2_SNAPSHOT_INVALID_SHEET = 1
    VATS2_SNAPSHOT_INVALID_HEADER = 2
    VATS2_SNAPSHOT_INVALID_COLUMN = 3
    VATS2_SNAPSHOT_READ_ERROR = 4
End Enum

Public Type VATS2ASnapshotRecord
    AIndex As Long
    ExcelRow As Long
    InvoiceDigitsRaw As Variant
    AmountRaw As Variant
    IsCompleted As Boolean
    InvoiceCellAddress As String
    AmountCellAddress As String
    InvoiceHasFormula As Boolean
    AmountHasFormula As Boolean
End Type

Public Type VATS2BSnapshotRecord
    BIndex As Long
    ExcelRow As Long
    SupplierTextRaw As Variant
    AmountRaw As Variant
    IsCompleted As Boolean
    SupplierCellAddress As String
    AmountCellAddress As String
    SupplierHasFormula As Boolean
    AmountHasFormula As Boolean
End Type

Public Type VATS2ASnapshotResult
    SheetName As String
    HeaderRow As Long
    RecordCount As Long
    Records() As VATS2ASnapshotRecord
    Status As VATS2SnapshotStatus
    ErrorRow As Long
    ErrorReason As String
End Type

Public Type VATS2BSnapshotResult
    SheetName As String
    HeaderRow As Long
    RecordCount As Long
    Records() As VATS2BSnapshotRecord
    Status As VATS2SnapshotStatus
    ErrorRow As Long
    ErrorReason As String
End Type

'内部临时记录也只保存值；不将 Excel 对象带入返回结果。
Private Type SnapshotRow
    IdentityRaw As Variant
    AmountRaw As Variant
    IsCompleted As Boolean
    IdentityAddress As String
    AmountAddress As String
    IdentityHasFormula As Boolean
    AmountHasFormula As Boolean
End Type

Public Function VATStage2ReadASnapshot(ByVal sheet As Worksheet, ByVal headerRow As Long, _
    ByVal invoiceColumn As Long, ByVal amountColumn As Long, _
    ByVal completedColor As Long) As VATS2ASnapshotResult
    Dim result As VATS2ASnapshotResult, item As SnapshotRow
    Dim lastRow As Long, row As Long, capacity As Long
    On Error GoTo Failed
    result.HeaderRow = headerRow
    result.Status = ValidateSnapshot(sheet, headerRow, invoiceColumn, amountColumn, _
                                    result.SheetName, lastRow, result.ErrorReason)
    If result.Status <> VATS2_SNAPSHOT_OK Then GoTo Done
    For row = headerRow + 1 To lastRow
        If ReadSnapshotRow(sheet, row, invoiceColumn, amountColumn, completedColor, item) Then
            result.RecordCount = result.RecordCount + 1
            If result.RecordCount > capacity Then
                capacity = capacity + 256
                ReDim Preserve result.Records(1 To capacity)
            End If
            With result.Records(result.RecordCount)
                .AIndex = result.RecordCount
                .ExcelRow = row
                .InvoiceDigitsRaw = item.IdentityRaw
                .AmountRaw = item.AmountRaw
                .IsCompleted = item.IsCompleted
                .InvoiceCellAddress = item.IdentityAddress
                .AmountCellAddress = item.AmountAddress
                .InvoiceHasFormula = item.IdentityHasFormula
                .AmountHasFormula = item.AmountHasFormula
            End With
        End If
    Next row
    If result.RecordCount > 0 Then ReDim Preserve result.Records(1 To result.RecordCount)
Done:
    VATStage2ReadASnapshot = result
    Exit Function
Failed:
    result.Status = VATS2_SNAPSHOT_READ_ERROR
    result.ErrorRow = row
    result.ErrorReason = Err.Description
    '失败不返回半份快照；调用方必须先检查 Status。
    result.RecordCount = 0
    Erase result.Records
    Resume Done
End Function

Public Function VATStage2ReadBSnapshot(ByVal sheet As Worksheet, ByVal headerRow As Long, _
    ByVal supplierColumn As Long, ByVal amountColumn As Long, _
    ByVal completedColor As Long) As VATS2BSnapshotResult
    Dim result As VATS2BSnapshotResult, item As SnapshotRow
    Dim lastRow As Long, row As Long, capacity As Long
    On Error GoTo Failed
    result.HeaderRow = headerRow
    result.Status = ValidateSnapshot(sheet, headerRow, supplierColumn, amountColumn, _
                                    result.SheetName, lastRow, result.ErrorReason)
    If result.Status <> VATS2_SNAPSHOT_OK Then GoTo Done
    For row = headerRow + 1 To lastRow
        If ReadSnapshotRow(sheet, row, supplierColumn, amountColumn, completedColor, item) Then
            result.RecordCount = result.RecordCount + 1
            If result.RecordCount > capacity Then
                capacity = capacity + 256
                ReDim Preserve result.Records(1 To capacity)
            End If
            With result.Records(result.RecordCount)
                .BIndex = result.RecordCount
                .ExcelRow = row
                .SupplierTextRaw = item.IdentityRaw
                .AmountRaw = item.AmountRaw
                .IsCompleted = item.IsCompleted
                .SupplierCellAddress = item.IdentityAddress
                .AmountCellAddress = item.AmountAddress
                .SupplierHasFormula = item.IdentityHasFormula
                .AmountHasFormula = item.AmountHasFormula
            End With
        End If
    Next row
    If result.RecordCount > 0 Then ReDim Preserve result.Records(1 To result.RecordCount)
Done:
    VATStage2ReadBSnapshot = result
    Exit Function
Failed:
    result.Status = VATS2_SNAPSHOT_READ_ERROR
    result.ErrorRow = row
    result.ErrorReason = Err.Description
    result.RecordCount = 0
    Erase result.Records
    Resume Done
End Function

Private Function ValidateSnapshot(ByVal sheet As Worksheet, ByVal headerRow As Long, _
    ByVal identityColumn As Long, ByVal amountColumn As Long, ByRef sheetName As String, _
    ByRef lastRow As Long, ByRef reason As String) As VATS2SnapshotStatus
    Dim book As Workbook, candidate As Worksheet, found As Boolean, area As Range
    On Error GoTo InvalidSheet
    If sheet Is Nothing Then GoTo InvalidSheet
    '关闭或删除的 Worksheet 引用也必须失效，不能伪装成成功空表。
    For Each book In sheet.Application.Workbooks
        For Each candidate In book.Worksheets
            If candidate Is sheet Then found = True: Exit For
        Next candidate
        If found Then Exit For
    Next book
    If Not found Then GoTo InvalidSheet
    sheetName = sheet.Name
    If headerRow < 1 Or headerRow > sheet.Rows.Count Then
        ValidateSnapshot = VATS2_SNAPSHOT_INVALID_HEADER
        reason = "HeaderRow 超出工作表范围。"
        Exit Function
    End If
    If identityColumn < 1 Or amountColumn < 1 Or _
       identityColumn > sheet.Columns.Count Or amountColumn > sheet.Columns.Count Then
        ValidateSnapshot = VATS2_SNAPSHOT_INVALID_COLUMN
        reason = "业务字段列号超出工作表范围。"
        Exit Function
    End If
    On Error GoTo ReadFailed
    '与 Stage 1 相同：包括仅有格式的尾行，随后按业务内容决定是否保留。
    Set area = sheet.UsedRange
    lastRow = area.Row + area.Rows.Count - 1
    ValidateSnapshot = VATS2_SNAPSHOT_OK
    Exit Function
InvalidSheet:
    ValidateSnapshot = VATS2_SNAPSHOT_INVALID_SHEET
    reason = "Worksheet 未提供、已关闭或已失效。"
    Exit Function
ReadFailed:
    ValidateSnapshot = VATS2_SNAPSHOT_READ_ERROR
    reason = Err.Description
End Function

Private Function ReadSnapshotRow(ByVal sheet As Worksheet, ByVal row As Long, _
    ByVal identityColumn As Long, ByVal amountColumn As Long, ByVal completedColor As Long, _
    ByRef item As SnapshotRow) As Boolean
    Dim identityCell As Range, amountCell As Range, fill As Object
    Set identityCell = sheet.Cells(row, identityColumn)
    Set amountCell = sheet.Cells(row, amountColumn)
    item.IdentityRaw = identityCell.Value2
    item.AmountRaw = amountCell.Value2
    item.IdentityHasFormula = identityCell.HasFormula
    item.AmountHasFormula = amountCell.HasFormula
    Set fill = amountCell.DisplayFormat.Interior
    '等价于 VATScan 的完成色计数条件；图案/合并/金额是否合法不由快照层判断。
    item.IsCompleted = (fill.Pattern <> xlPatternNone And fill.ColorIndex <> xlColorIndexNone _
                        And fill.Color = completedColor)
    If IsEmpty(item.IdentityRaw) And IsEmpty(item.AmountRaw) And _
       Not item.IdentityHasFormula And Not item.AmountHasFormula And Not item.IsCompleted Then Exit Function
    item.IdentityAddress = identityCell.Address(False, False)
    item.AmountAddress = amountCell.Address(False, False)
    ReadSnapshotRow = True
End Function
