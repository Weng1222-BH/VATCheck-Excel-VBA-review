Attribute VB_Name = "modVATStage2BConflictTests"
Option Explicit
Option Compare Binary

Private caseCount As Long, failureCount As Long, details As String

Public Function VATStage2BConflict_SelfTest() As String
    On Error GoTo Unexpected
    caseCount = 0: failureCount = 0: details = vbNullString
    '期望组格式：AIndex:Flags:DistinctBRowCount:BIndex/ReferenceIndex,...；组间分号。
    CheckCase "不同 A 无冲突", "NO.11111111|NO.22222222", "11111111|22222222", 2, ""
    CheckCase "两个 exact 跨 B", "NO.12345678|NO.12345678", "12345678", 2, "1:2:2:1/1,2/1"
    CheckCase "exact 与 suffix 跨 B", "NO.12345678|NO.345678", "12345678", 2, "1:2:2:1/1,2/1"
    CheckCase "三个 B 共同引用", "NO.12345678|NO.345678|NO.5678", "12345678", 3, "1:2:3:1/1,2/1,3/1"
    CheckCase "同 B 不同引用复用", "NO.12345678/345678", "12345678", 2, "1:1:1:1/1,1/2"
    CheckCase "同组同时两种复用", "NO.12345678/345678|NO.12345678", "12345678", 3, "1:3:2:1/1,1/2,2/1"
    CheckCase "NOT_FOUND 不参与", "NO.12345678/99999999|NO.99999999", "12345678", 1, ""
    CheckCase "suffix multiple 不取其中一个", "NO.345678|NO.12345678", "12345678|9912345678", 1, ""
    CheckCase "exact multiple 不参与", "NO.12345678|NO.12345678", "12345678|12345678", 0, ""
    CheckCase "INCOMPLETE 中 unique 参与", "NO.12345678/99999999|NO.345678", "12345678", 2, "1:2:2:1/1,2/1"
    CheckCase "Parser 无 NO 风险不阻断", "供应商ABC12345678|NO.12345678", "12345678", 2, "1:2:2:1/1,2/1"
    CheckCase "结构与尾注风险不阻断", "NO.12345678//--123|NO.12345678", "12345678", 2, "1:2:2:1/1,2/1"
    CheckCase "短尾号风险不阻断", "NO.345678|NO.5678", "12345678", 2, "1:2:2:1/1,2/1"
    CheckCase "Parser 去重不虚构同 B 关联", "NO.12345678/12345678", "12345678", 1, ""
    CheckCase "去重后仍可跨 B", "NO.12345678/12345678|NO.12345678", "12345678", 2, "1:2:2:1/1,2/1"
    CheckCase "多组按首次而非成组时间排序", "NO.22222222/11111111|NO.11111111|NO.22222222", "11111111|22222222", 4, "2:2:2:1/1,3/1;1:2:2:1/2,2/1"
    CheckCase "组内跳过非 unique 保留引用索引", "NO.99999999/12345678/345678|NO.5678", "12345678", 3, "1:3:2:1/2,1/3,2/1"
    CheckCase "独立的数组下界", "NO.12345678|NO.345678", "12345678", 2, "1:2:2:-2/1,-1/1", -2, 8
    CheckCase "正下界 B 与负下界 ID", "NO.12345678|NO.345678", "12345678", 2, "1:2:2:7/1,8/1", 7, -4
    CheckCase "只检查 bCount 前部", "NO.12345678|NO.12345678", "12345678", 1, "", 1, 1, 1
    CheckCase "空未分配数组", "", "12345678", 0, ""
    CheckCase "已分配 bCount 零", "NO.12345678", "12345678", 0, "", 1, 1, 0
    CheckCase "负 bCount", "NO.12345678", "12345678", 0, "", 1, 1, -1, VATS2_BCONFLICT_INVALID_RANGE
    CheckCase "bRows 容量不足", "NO.12345678", "12345678", 0, "", 1, 1, 2, VATS2_BCONFLICT_INVALID_RANGE, 2
    CheckCase "bRowIds 容量不足", "NO.12345678|NO.345678", "12345678", 0, "", 1, 1, 2, VATS2_BCONFLICT_INVALID_RANGE, 1
    CheckCase "bRows 未分配", "", "12345678", 0, "", 1, 1, 1, VATS2_BCONFLICT_INVALID_RANGE, 1
    CheckCase "bRowIds 未分配", "NO.12345678", "12345678", 0, "", 1, 1, 1, VATS2_BCONFLICT_INVALID_RANGE, 0
    CheckCase "RowId 相同也算两个记录", "NO.12345678|NO.345678", "12345678", 2, "1:2:2:1/1,2/1", 1, 1, -999, VATS2_BCONFLICT_OK, -1, True
    CheckCase "零引用 B 行", "无号码|无数字", "12345678", 0, ""
    CheckCase "UNIQUE 声称零候选", "NO.12345678", "12345678", 0, "", 1, 1, -999, VATS2_BCONFLICT_INVALID_CONTRACT, -1, False, 1
    CheckCase "UNIQUE 声称两个候选", "NO.12345678", "12345678", 0, "", 1, 1, -999, VATS2_BCONFLICT_INVALID_CONTRACT, -1, False, 2
    CheckCase "UNIQUE 候选数组未分配", "NO.12345678", "12345678", 0, "", 1, 1, -999, VATS2_BCONFLICT_INVALID_CONTRACT, -1, False, 3
    CheckCase "引用数量非法", "NO.12345678", "12345678", 0, "", 1, 1, -999, VATS2_BCONFLICT_INVALID_CONTRACT, -1, False, 4
    CheckCase "引用数组未分配", "NO.12345678", "12345678", 0, "", 1, 1, -999, VATS2_BCONFLICT_INVALID_CONTRACT, -1, False, 5
    CheckCase "INVALID_REFERENCE 不参与", "NO.12345678|NO.345678", "12345678", 1, "", 1, 1, -999, VATS2_BCONFLICT_OK, -1, False, 6
    CheckCase "INVALID_A_RANGE 不参与", "NO.12345678|NO.345678", "12345678", 1, "", 1, 1, -999, VATS2_BCONFLICT_OK, -1, False, 7
    CheckCase "同 AIndex 不同完整号码拒绝", "NO.12345678|NO.345678", "12345678", 0, "", 1, 1, -999, VATS2_BCONFLICT_INVALID_CONTRACT, -1, False, 8
    CheckCase "仅按 AIndex 而非完整号码分组", "NO.12345678|NO.345678", "12345678", 2, "", 1, 1, -999, VATS2_BCONFLICT_OK, -1, False, 9
    CheckCase "行 INVALID_INPUT 不过滤其 unique", "NO.12345678|NO.345678", "12345678", 2, "1:2:2:1/1,2/1", 1, 1, -999, VATS2_BCONFLICT_OK, -1, False, 10
    CheckCase "AIndex 可以为负", "NO.12345678|NO.345678", "12345678", 2, "-5:2:2:1/1,2/1", 1, 1, -999, VATS2_BCONFLICT_OK, -1, False, 11
    CheckCase "晚发契约异常不保留部分计数", "NO.12345678|NO.345678", "12345678", 0, "", 1, 1, -999, VATS2_BCONFLICT_INVALID_CONTRACT, -1, False, 12
    If failureCount = 0 Then
        VATStage2BConflict_SelfTest = "PASS: " & caseCount & " cases" & vbCrLf & details
    Else
        VATStage2BConflict_SelfTest = "FAIL: " & failureCount & "/" & caseCount & " cases" & vbCrLf & details
    End If
    Exit Function
Unexpected:
    VATStage2BConflict_SelfTest = "FAIL: unexpected " & Err.Number & " " & Err.Description & vbCrLf & details
End Function

Private Sub CheckCase(ByVal title As String, ByVal texts As String, ByVal aText As String, _
    ByVal expectedLinks As Long, ByVal expectedGroups As String, Optional ByVal rowLower As Long = 1, _
    Optional ByVal idLower As Long = 1, Optional ByVal scanCount As Long = -999, _
    Optional ByVal expectedStatus As VATS2BConflictStatus = VATS2_BCONFLICT_OK, _
    Optional ByVal idCount As Long = -1, Optional ByVal sameId As Boolean = False, Optional ByVal mutation As Long = 0)
    Dim rows() As VATS2BRowMatchResult, ids() As Long, a() As String, parts As Variant
    Dim n As Long, i As Long, groups As Long, expectedRecords As Long, before As String, reason As String
    Dim result As VATS2BConflictResult, again As VATS2BConflictResult
    On Error GoTo Failed
    caseCount = caseCount + 1
    a = Split(aText, "|")
    If Len(texts) > 0 Then
        parts = Split(texts, "|"): n = UBound(parts) + 1
        ReDim rows(rowLower To rowLower + n - 1)
        For i = 0 To n - 1
            rows(rowLower + i) = VATStage2MatchBRow(CStr(parts(i)), a, UBound(a) + 1)
            'A 由 Split 产生零下界，测试显式映射为从1开始，便于核对期望索引。
            ShiftAIndexes rows(rowLower + i), 1
        Next i
    End If
    If scanCount = -999 Then scanCount = n
    If idCount = -1 Then idCount = n
    If idCount > 0 Then
        ReDim ids(idLower To idLower + idCount - 1)
        For i = 0 To idCount - 1
            If sameId Then ids(idLower + i) = 777 Else ids(idLower + i) = 100 + i * 10
        Next i
    End If
    If mutation > 0 Then MutateFixture rows, rowLower, mutation
    before = InputText(rows, ids, n, idCount, rowLower, idLower)
    result = VATStage2ScanBConflicts(rows, ids, scanCount)
    again = VATStage2ScanBConflicts(rows, ids, scanCount)
    If Len(expectedGroups) > 0 Then groups = UBound(Split(expectedGroups, ";")) + 1
    If expectedStatus <> VATS2_BCONFLICT_INVALID_RANGE Then expectedRecords = scanCount
    CheckResult result, rows, ids, rowLower, idLower, expectedRecords, expectedLinks, groups, expectedGroups, expectedStatus
    CheckResult again, rows, ids, rowLower, idLower, expectedRecords, expectedLinks, groups, expectedGroups, expectedStatus
    Require result.ErrorBIndex = again.ErrorBIndex And result.ErrorReferenceIndex = again.ErrorReferenceIndex, "错误定位重复调用一致"
    If expectedStatus = VATS2_BCONFLICT_INVALID_CONTRACT Then
        i = rowLower
        If mutation = 8 Or mutation = 12 Then i = i + 1
        Require result.ErrorBIndex = i, "契约错误 BIndex"
        If mutation = 4 Or mutation = 5 Then i = 0 Else i = 1
        Require result.ErrorReferenceIndex = i, "契约错误 ReferenceIndex"
    End If
    Require InputText(rows, ids, n, idCount, rowLower, idLower) = before, "所有 B 输入字段及 RowIds 未修改"
    details = details & "PASS " & title & vbCrLf
    Exit Sub
Failed:
    reason = Err.Description: failureCount = failureCount + 1
    details = details & "FAIL " & title & ": " & reason & vbCrLf
End Sub

Private Sub CheckResult(ByRef result As VATS2BConflictResult, ByRef rows() As VATS2BRowMatchResult, _
    ByRef ids() As Long, ByVal rowLower As Long, ByVal idLower As Long, ByVal records As Long, _
    ByVal links As Long, ByVal groups As Long, ByVal expected As String, ByVal status As VATS2BConflictStatus)
    Dim i As Long, j As Long, text As String, upper As Long
    Require result.Status = status, "Status"
    Require result.BRecordCount = records, "BRecordCount"
    Require result.UniqueAssociationCount = links, "UniqueAssociationCount"
    Require result.ConflictGroupCount = groups, "ConflictGroupCount"
    If groups > 0 Then
        Require LBound(result.ConflictGroups) = 1 And UBound(result.ConflictGroups) = groups, "组数组容量"
    Else
        On Error Resume Next
        upper = UBound(result.ConflictGroups)
        j = Err.Number: Err.Clear
        On Error GoTo 0
        Require j <> 0, "零组数组未分配"
    End If
    For i = 1 To groups
        With result.ConflictGroups(i)
            If i > 1 Then text = text & ";"
            text = text & .AIndex & ":" & .Flags & ":" & .DistinctBRowCount & ":"
            Require .AssociationCount >= 2, "至少两个关联"
            Require LBound(.Associations) = 1 And UBound(.Associations) = .AssociationCount, "关联数组数量"
            For j = 1 To .AssociationCount
                CheckAssociation .Associations(j), rows, ids, rowLower, idLower, .AIndex, .AInvoiceDigits
                If j > 1 Then text = text & ","
                text = text & .Associations(j).BIndex & "/" & .Associations(j).ReferenceIndex
            Next j
        End With
    Next i
    Require text = expected, "全部组及关联顺序、索引、Flags、DistinctBRowCount"
End Sub

Private Sub CheckAssociation(ByRef link As VATS2BAssociation, ByRef rows() As VATS2BRowMatchResult, _
    ByRef ids() As Long, ByVal rowLower As Long, ByVal idLower As Long, ByVal ai As Long, ByVal digits As String)
    With rows(link.BIndex).References(link.ReferenceIndex)
        Require link.BRowId = ids(idLower + link.BIndex - rowLower), "BRowId 按相对位置映射"
        Require link.ReferenceDigits = .Digits, "ReferenceDigits"
        Require link.MatchKind = .MatchKind, "MatchKind"
        Require link.AIndex = .Candidates(1).AIndex And link.AIndex = ai, "AIndex"
        Require link.AInvoiceDigits = .Candidates(1).InvoiceDigits And link.AInvoiceDigits = digits, "A 完整号码"
    End With
End Sub

Private Sub ShiftAIndexes(ByRef row As VATS2BRowMatchResult, ByVal delta As Long)
    Dim i As Long, j As Long
    For i = 1 To row.ReferenceCount
        For j = 1 To row.References(i).CandidateCount
            row.References(i).Candidates(j).AIndex = row.References(i).Candidates(j).AIndex + delta
        Next j
    Next i
End Sub

Private Sub MutateFixture(ByRef rows() As VATS2BRowMatchResult, ByVal first As Long, ByVal mode As Long)
    '仅构造测试输入契约异常，不改冻结模块。
    Select Case mode
        Case 1: rows(first).References(1).CandidateCount = 0
        Case 2: rows(first).References(1).CandidateCount = 2
        Case 3: Erase rows(first).References(1).Candidates
        Case 4: rows(first).ReferenceCount = -1
        Case 5: Erase rows(first).References
        Case 6: rows(first).References(1).MatchKind = VATS2_INVALID_REFERENCE
        Case 7: rows(first).References(1).MatchKind = VATS2_INVALID_A_RANGE
        Case 8: rows(first + 1).References(1).Candidates(1).InvoiceDigits = "99999999"
        Case 9: rows(first + 1).References(1).Candidates(1).AIndex = 2
        Case 10: rows(first).MatchState = VATS2_BROW_INVALID_INPUT
        Case 11
            rows(first).References(1).Candidates(1).AIndex = -5
            rows(first + 1).References(1).Candidates(1).AIndex = -5
        Case 12: rows(first + 1).References(1).CandidateCount = 0
    End Select
End Sub

Private Function InputText(ByRef rows() As VATS2BRowMatchResult, ByRef ids() As Long, _
    ByVal count As Long, ByVal idCount As Long, ByVal first As Long, ByVal idFirst As Long) As String
    Dim i As Long, j As Long, k As Long, lo As Long, hi As Long, text As String
    For i = first To first + count - 1
        With rows(i)
            text = text & Pack(.OriginalText) & "|" & .ParserFlags & ":" & .ReferenceCount & ":" & .MatcherFlags & ":" & .MatchState
            text = text & ":" & .UniqueMatchCount & ":" & .NotFoundCount & ":" & .MultipleMatchCount & ":" & .InvalidMatchCount
        End With
        lo = 1: hi = 0
        On Error Resume Next
        lo = LBound(rows(i).References): hi = UBound(rows(i).References)
        Err.Clear: On Error GoTo 0
        text = text & "[" & lo & "," & hi & "]"
        For j = lo To hi
            With rows(i).References(j)
                text = text & Pack(.Digits) & .Length & Pack(.RawFragment) & .StartIndex & Pack(.Context)
                text = text & ":" & .ParserFlags & ":" & .MatchKind & ":" & .MatcherFlags & ":" & .CandidateCount
            End With
            text = text & CandidateText(rows(i).References(j))
        Next j
    Next i
    For k = idFirst To idFirst + idCount - 1
        text = text & "ID:" & ids(k) & ";"
    Next k
    InputText = text
End Function

Private Function CandidateText(ByRef reference As VATS2BRowReference) As String
    Dim lo As Long, hi As Long, k As Long, text As String
    lo = 1: hi = 0
    On Error Resume Next
    lo = LBound(reference.Candidates): hi = UBound(reference.Candidates)
    Err.Clear: On Error GoTo 0
    text = "[" & lo & "," & hi & "]"
    For k = lo To hi
        text = text & reference.Candidates(k).AIndex & Pack(reference.Candidates(k).InvoiceDigits)
    Next k
    CandidateText = text
End Function

Private Function Pack(ByVal text As String) As String
    Pack = Len(text) & ":" & text
End Function

Private Sub Require(ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then Err.Raise vbObjectError + 2401, "BConflictTests", reason
End Sub
