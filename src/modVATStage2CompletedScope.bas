Attribute VB_Name = "modVATStage2CompletedScope"
Option Explicit

Public Enum VATS2ScopeStatus
    VATS2_SCOPE_OK = 0
    VATS2_SCOPE_INVALID_INPUT = 1
    VATS2_SCOPE_INVALID_CONTRACT = 2
    VATS2_SCOPE_CONFLICT_ERROR = 3
End Enum

Public Type VATS2AIdentityIssue
    OriginalAIndex As Long
    ExcelRow As Long
    RawVarType As Long
End Type

Public Type VATS2BSourceIssue
    OriginalBIndex As Long
    ExcelRow As Long
    RawVarType As Long
End Type

Public Type VATS2CompletedScopeResult
    Status As VATS2ScopeStatus
    FullARecordCount As Long
    FullBRecordCount As Long
    CompletedACount As Long
    CompletedBCount As Long
    CompletedAIndexes() As Long
    CompletedAExcelRows() As Long
    AIdentityIssueCount As Long
    AIdentityIssues() As VATS2AIdentityIssue
    BInScope() As Boolean
    BRowIds() As Long
    BSourceIssueCount As Long
    BSourceIssues() As VATS2BSourceIssue
    BMatches() As VATS2BRowMatchResult
    Conflicts As VATS2BConflictResult
    ErrorSide As String
    ErrorIndex As Long
    ErrorReason As String
End Type

Public Function VATStage2MatchCompletedScope(ByRef a As VATS2ASnapshotResult, _
    ByRef b As VATS2BSnapshotResult) As VATS2CompletedScopeResult
    Dim r As VATS2CompletedScopeResult, emptyResult As VATS2CompletedScopeResult
    Dim digits() As String, source As VATS2BRowMatchResult, text As String
    Dim i As Long, p As Long, phase As VATS2ScopeStatus
    r.Status = VATS2_SCOPE_INVALID_INPUT
    If a.Status <> VATS2_SNAPSHOT_OK Or b.Status <> VATS2_SNAPSHOT_OK Then
        r.ErrorReason = "输入 Snapshot.Status 非 OK。": GoTo Done
    End If
    r.Status = VATS2_SCOPE_INVALID_CONTRACT
    r.ErrorSide = "A"
    If Not ValidASnapshot(a) Then r.ErrorReason = "A 快照记录数组、索引或行序契约无效。": GoTo Done
    r.ErrorSide = "B"
    If Not ValidBSnapshot(b) Then r.ErrorReason = "B 快照记录数组、索引或行序契约无效。": GoTo Done
    r.ErrorSide = vbNullString
    '从此处开始只使用已经验证的完整快照空间，不复制金额或其他业务值。
    phase = VATS2_SCOPE_INVALID_CONTRACT
    On Error GoTo Failed
    r.FullARecordCount = a.RecordCount: r.FullBRecordCount = b.RecordCount
    For i = 1 To a.RecordCount
        If a.Records(i).IsCompleted Then r.CompletedACount = r.CompletedACount + 1
    Next i
    If r.CompletedACount > 0 Then
        ReDim digits(1 To r.CompletedACount)
        ReDim r.CompletedAIndexes(1 To r.CompletedACount)
        ReDim r.CompletedAExcelRows(1 To r.CompletedACount)
        ReDim r.AIdentityIssues(1 To r.CompletedACount)
    End If
    For i = 1 To a.RecordCount
        If a.Records(i).IsCompleted Then
            p = p + 1
            r.CompletedAIndexes(p) = a.Records(i).AIndex
            r.CompletedAExcelRows(p) = a.Records(i).ExcelRow
            If VarType(a.Records(i).InvoiceDigitsRaw) = vbString Then
                digits(p) = a.Records(i).InvoiceDigitsRaw
            Else
                '保留该 scope 位置的默认空字符串；绝不 CStr 生成猜测号码。
                r.AIdentityIssueCount = r.AIdentityIssueCount + 1
                With r.AIdentityIssues(r.AIdentityIssueCount)
                    .OriginalAIndex = a.Records(i).AIndex
                    .ExcelRow = a.Records(i).ExcelRow
                    .RawVarType = VarType(a.Records(i).InvoiceDigitsRaw)
                End With
            End If
        End If
    Next i
    If r.AIdentityIssueCount > 0 Then
        ReDim Preserve r.AIdentityIssues(1 To r.AIdentityIssueCount)
    Else
        Erase r.AIdentityIssues
    End If
    If b.RecordCount > 0 Then
        ReDim r.BInScope(1 To b.RecordCount)
        ReDim r.BRowIds(1 To b.RecordCount)
        ReDim r.BMatches(1 To b.RecordCount)
        ReDim r.BSourceIssues(1 To b.RecordCount)
    End If
    For i = 1 To b.RecordCount
        r.BRowIds(i) = b.Records(i).ExcelRow
        r.BInScope(i) = b.Records(i).IsCompleted
        '未完成 B 保留默认合法 NO_REFERENCES 占位，不调用 Parser/Matcher。
        If r.BInScope(i) Then
            r.CompletedBCount = r.CompletedBCount + 1
            text = vbNullString
            If VarType(b.Records(i).SupplierTextRaw) = vbString Then
                text = b.Records(i).SupplierTextRaw
            Else
                r.BSourceIssueCount = r.BSourceIssueCount + 1
                With r.BSourceIssues(r.BSourceIssueCount)
                    .OriginalBIndex = b.Records(i).BIndex
                    .ExcelRow = b.Records(i).ExcelRow
                    .RawVarType = VarType(b.Records(i).SupplierTextRaw)
                End With
            End If
            source = VATStage2MatchBRow(text, digits, r.CompletedACount)
            If source.MatchState = VATS2_BROW_INVALID_INPUT Then
                r.ErrorReason = "冻结 BRowMatch 返回无效输入状态。": GoTo ContractFailed
            End If
            If VATStage2RemapScopeRow(source, r.CompletedAIndexes, r.CompletedACount, r.BMatches(i)) <> VATS2_SCOPE_OK Then
                r.ErrorReason = "候选位置超出 completed scope 映射或结果数组契约无效。": GoTo ContractFailed
            End If
        End If
    Next i
    If r.BSourceIssueCount > 0 Then
        ReDim Preserve r.BSourceIssues(1 To r.BSourceIssueCount)
    Else
        Erase r.BSourceIssues
    End If
    phase = VATS2_SCOPE_CONFLICT_ERROR
    r.Conflicts = VATStage2ScanBConflicts(r.BMatches, r.BRowIds, b.RecordCount)
    If r.Conflicts.Status <> VATS2_BCONFLICT_OK Then
        r.Status = VATS2_SCOPE_CONFLICT_ERROR
        r.ErrorSide = "B": r.ErrorIndex = r.Conflicts.ErrorBIndex
        r.ErrorReason = "冻结 BConflict 扫描失败；请检查嵌套 Conflicts 状态。"
        GoTo Done
    End If
    r.Status = VATS2_SCOPE_OK
Done:
    VATStage2MatchCompletedScope = r
    Exit Function
ContractFailed:
    r.Status = VATS2_SCOPE_INVALID_CONTRACT
    r.ErrorSide = "B": r.ErrorIndex = i
    '返回失败状态和定位；不发布已经回映一部分的行结果。
    emptyResult.Status = r.Status: emptyResult.ErrorSide = r.ErrorSide
    emptyResult.ErrorIndex = r.ErrorIndex: emptyResult.ErrorReason = r.ErrorReason
    r = emptyResult
    GoTo Done
Failed:
    emptyResult.Status = phase
    emptyResult.ErrorReason = Err.Description
    r = emptyResult
    Resume Done
End Function

'独立回映接口便于验证损坏候选；仅改结果副本 AIndex，失败时输出为空。
Public Function VATStage2RemapScopeRow(ByRef source As VATS2BRowMatchResult, _
    ByRef originalAIndexes() As Long, ByVal scopeCount As Long, _
    ByRef remapped As VATS2BRowMatchResult) As VATS2ScopeStatus
    Dim copy As VATS2BRowMatchResult, blank As VATS2BRowMatchResult
    Dim i As Long, j As Long, p As Long, previous As Long
    remapped = blank
    VATStage2RemapScopeRow = VATS2_SCOPE_INVALID_CONTRACT
    On Error GoTo Invalid
    If scopeCount < 0 Or source.ReferenceCount < 0 Then Exit Function
    If scopeCount > 0 Then
        If LBound(originalAIndexes) <> 1 Or UBound(originalAIndexes) <> scopeCount Then Exit Function
        For i = 1 To scopeCount
            If originalAIndexes(i) <= previous Then Exit Function
            previous = originalAIndexes(i)
        Next i
    End If
    If source.ReferenceCount > 0 Then
        If LBound(source.References) <> 1 Or UBound(source.References) <> source.ReferenceCount Then Exit Function
    End If
    '先验证全部位置，避免失败后泄漏混合索引空间。
    For i = 1 To source.ReferenceCount
        With source.References(i)
            If .CandidateCount < 0 Then Exit Function
            If .CandidateCount > 0 Then
                If LBound(.Candidates) <> 1 Or UBound(.Candidates) <> .CandidateCount Then Exit Function
            End If
            For j = 1 To .CandidateCount
                p = .Candidates(j).AIndex
                If p < 1 Or p > scopeCount Then Exit Function
            Next j
        End With
    Next i
    copy = source
    For i = 1 To copy.ReferenceCount
        For j = 1 To copy.References(i).CandidateCount
            p = copy.References(i).Candidates(j).AIndex
            copy.References(i).Candidates(j).AIndex = originalAIndexes(p)
        Next j
    Next i
    remapped = copy
    VATStage2RemapScopeRow = VATS2_SCOPE_OK
Invalid:
End Function

Private Function ValidASnapshot(ByRef a As VATS2ASnapshotResult) As Boolean
    Dim i As Long, previous As Long
    On Error GoTo Invalid
    If a.RecordCount < 0 Or a.HeaderRow < 1 Then Exit Function
    If a.RecordCount = 0 Then
        '冻结快照的零记录数组未分配，不能把已有记录藏在零计数后。
        On Error GoTo EmptyArray
        i = LBound(a.Records)
        Exit Function
    End If
    If a.RecordCount > 0 Then
        If LBound(a.Records) <> 1 Or UBound(a.Records) <> a.RecordCount Then Exit Function
    End If
    previous = a.HeaderRow
    For i = 1 To a.RecordCount
        If a.Records(i).AIndex <> i Or a.Records(i).ExcelRow <= previous Then Exit Function
        previous = a.Records(i).ExcelRow
    Next i
    ValidASnapshot = True
    Exit Function
EmptyArray:
    ValidASnapshot = (Err.Number = 9)
Invalid:
End Function

Private Function ValidBSnapshot(ByRef b As VATS2BSnapshotResult) As Boolean
    Dim i As Long, previous As Long
    On Error GoTo Invalid
    If b.RecordCount < 0 Or b.HeaderRow < 1 Then Exit Function
    If b.RecordCount = 0 Then
        On Error GoTo EmptyArray
        i = LBound(b.Records)
        Exit Function
    End If
    If b.RecordCount > 0 Then
        If LBound(b.Records) <> 1 Or UBound(b.Records) <> b.RecordCount Then Exit Function
    End If
    previous = b.HeaderRow
    For i = 1 To b.RecordCount
        If b.Records(i).BIndex <> i Or b.Records(i).ExcelRow <= previous Then Exit Function
        previous = b.Records(i).ExcelRow
    Next i
    ValidBSnapshot = True
    Exit Function
EmptyArray:
    ValidBSnapshot = (Err.Number = 9)
Invalid:
End Function
