Attribute VB_Name = "modVATStage2GroupAmountTests"
Option Explicit

Private caseCount As Long, failureCount As Long, details As String

Public Function VATStage2GroupAmount_SelfTest() As String
    Dim maximum As Variant
    On Error GoTo Unexpected
    caseCount = 0: failureCount = 0: details = vbNullString
    maximum = CDec("79228162514264337593543950335")
    CheckCase "单 A 相等", "NO.11111111", "11111111", Array(10), 10, VATS2_GROUP_COMPARED, 0, 0, 0, 0
    CheckCase "单 A 不等正差额", "NO.11111111", "11111111", Array(10), 8, VATS2_GROUP_COMPARED, 2, 0, 0, 0
    CheckCase "单 A 不等负差额", "NO.11111111", "11111111", Array(8), 10, VATS2_GROUP_COMPARED, -2, 0, 0, 0
    CheckCase "多 A 相等", "NO.11111111/22222222", "11111111|22222222", Array(10, 20), 30, VATS2_GROUP_COMPARED, 0, 0, 0, 0
    CheckCase "多 A 不等", "NO.11111111/22222222", "11111111|22222222", Array(10, 20), 29, VATS2_GROUP_COMPARED, 1, 0, 0, 0
    CheckCase "B 顺序决定金额组顺序", "NO.22222222/11111111", "11111111|22222222", Array(10, 20), 30, VATS2_GROUP_COMPARED, 0, 0, 0, 0
    CheckCase "零引用不比较非法金额", "无发票", "11111111", Array("非法"), "非法", VATS2_GROUP_NOT_COMPARABLE, Empty, 0, 0, VATS2_GROUP_NO_REFERENCES
    CheckCase "unique加未找到不比部分金额", "NO.11111111/33333333", "11111111|22222222", Array(10, 20), 10, VATS2_GROUP_NOT_COMPARABLE, Empty, 0, 0, VATS2_GROUP_INCOMPLETE_RELATION
    CheckCase "unique加suffix multiple不比较", "NO.11111111/345678", "11111111|12345678|99345678", Array(10, 20, 30), 10, VATS2_GROUP_NOT_COMPARABLE, Empty, 0, 1, VATS2_GROUP_INCOMPLETE_RELATION
    CheckCase "重复 A 导致 exact multiple", "NO.11111111/22222222", "11111111|22222222|22222222", Array(10, 20, 20), 50, VATS2_GROUP_NOT_COMPARABLE, Empty, 0, 0, VATS2_GROUP_INCOMPLETE_RELATION
    CheckCase "关系 INVALID_INPUT", "NO.11111111", "11111111", Array(10), 10, VATS2_GROUP_INVALID_INPUT, Empty, 0, 0, VATS2_GROUP_INVALID_RELATION, False, False, 0, 1, -999, 1
    CheckCase "本行重复 A 禁止求和", "NO.12345678/345678", "12345678", Array(10), 20, VATS2_GROUP_NOT_COMPARABLE, Empty, 0, 1, VATS2_GROUP_SAME_ROW_A_REUSE, True
    CheckCase "跨 B 允许诊断比较", "NO.12345678|NO.12345678", "12345678", Array(10), 10, VATS2_GROUP_COMPARED, 0, 0, 0, 0, False, True
    CheckCase "同组 B1 same加cross 被阻断", "NO.12345678/345678|NO.12345678", "12345678", Array(10), 20, VATS2_GROUP_NOT_COMPARABLE, Empty, 0, 1, VATS2_GROUP_SAME_ROW_A_REUSE, True, True
    CheckCase "同组 B2 仅cross仍可比较", "NO.12345678/345678|NO.12345678", "12345678", Array(10), 10, VATS2_GROUP_COMPARED, 0, 0, 0, 0, False, True, 1
    CheckCase "缺 NO 风险保留", "供应商ABC11111111", "11111111", Array(10), 10, VATS2_GROUP_COMPARED, 0, VATS2_NO_MARKER_MISSING, 0, 0
    CheckCase "短尾号风险保留", "NO.345678", "12345678", Array(10), 10, VATS2_GROUP_COMPARED, 0, 0, 1, 0
    CheckCase "结构风险仍可诊断比较", "NO.11111111//22222222", "11111111|22222222", Array(10, 20), 30, VATS2_GROUP_COMPARED, 0, VATS2_STRUCTURE_REVIEW, 0, 0
    CheckCase "纯数字尾注风险保留", "NO.11111111--123", "11111111", Array(10), 10, VATS2_GROUP_COMPARED, 0, VATS2_NUMERIC_TRAILER_REVIEW, 0, 0
    CheckCase "Parser去重风险仍可比较", "NO.11111111/11111111", "11111111", Array(10), 10, VATS2_GROUP_COMPARED, 0, VATS2_DUPLICATE_REF_REVIEW, 0, 0
    CheckCase "所有风险不因相等消失", "供应商BC345678//345678--123|NO.12345678", "12345678", Array(10), 10, VATS2_GROUP_COMPARED, 0, 15, 1, 0, False, True
    CheckCase "跨行风险和不一致并存", "NO.12345678|NO.345678", "12345678", Array(10), 9, VATS2_GROUP_COMPARED, 1, 0, 0, 0, False, True
    CheckCase "非法 A 金额错误", "NO.11111111", "11111111", Array("10"), 10, VATS2_GROUP_AMOUNT_ERROR, Empty, 0, 0, 0
    CheckCase "非法 B 金额错误", "NO.11111111", "11111111", Array(10), "10", VATS2_GROUP_AMOUNT_ERROR, Empty, 0, 0, 0
    CheckCase "Decimal 累加错误传播", "NO.11111111/22222222", "11111111|22222222", Array(maximum, CDec(1)), 0, VATS2_GROUP_AMOUNT_ERROR, Empty, 0, 0, 0
    CheckCase "Decimal 差额错误传播", "NO.11111111", "11111111", Array(maximum), CDec(-1), VATS2_GROUP_AMOUNT_ERROR, Empty, 0, 0, 0
    CheckCase "AIndex 零直接读取", "NO.11111111", "11111111", Array(10), 10, VATS2_GROUP_COMPARED, 0, 0, 0, 0, False, False, 0, 0
    CheckCase "负 AIndex 和金额负下界", "NO.22222222/11111111", "11111111|22222222", Array(10, 20), 30, VATS2_GROUP_COMPARED, 0, 0, 0, 0, False, False, 0, -5
    CheckCase "AIndex15不重新编号", "NO.22222222/11111111", "11111111|22222222", Array(10, 20), 30, VATS2_GROUP_COMPARED, 0, 0, 0, 0, False, False, 0, 15
    CheckCase "AIndex超过金额上界", "NO.11111111/22222222", "11111111|22222222", Array(10), 10, VATS2_GROUP_INVALID_INPUT, Empty, 0, 0, VATS2_GROUP_INVALID_A_MAPPING
    CheckCase "AIndex低于金额下界", "NO.11111111", "11111111", Array(10), 10, VATS2_GROUP_INVALID_INPUT, Empty, 0, 0, VATS2_GROUP_INVALID_A_MAPPING, False, False, 0, 1, 2
    CheckCase "金额数组未分配", "NO.11111111", "11111111", Array(), 10, VATS2_GROUP_INVALID_INPUT, Empty, 0, 0, VATS2_GROUP_INVALID_A_MAPPING
    CheckCase "UNIQUE数量损坏", "NO.11111111", "11111111", Array(10), 10, VATS2_GROUP_INVALID_INPUT, Empty, 0, 0, VATS2_GROUP_INVALID_RELATION, False, False, 0, 1, -999, 2
    CheckCase "ALL_UNIQUE含非唯一引用", "NO.11111111", "11111111", Array(10), 10, VATS2_GROUP_INVALID_INPUT, Empty, 0, 0, VATS2_GROUP_INVALID_RELATION, False, False, 0, 1, -999, 3
    CheckCase "冲突 INVALID_RANGE", "NO.11111111", "11111111", Array(10), 10, VATS2_GROUP_INVALID_INPUT, Empty, 0, 0, VATS2_GROUP_INVALID_CONFLICT, False, False, 0, 1, -999, 4
    CheckCase "冲突 INVALID_CONTRACT", "NO.11111111", "11111111", Array(10), 10, VATS2_GROUP_INVALID_INPUT, Empty, 0, 0, VATS2_GROUP_INVALID_CONFLICT, False, False, 0, 1, -999, 5
    CheckCase "0.1加0.2集成精确相等", "NO.11111111/22222222", "11111111|22222222", Array(0.1, 0.2), 0.3, VATS2_GROUP_COMPARED, 0, 0, 0, 0
    CheckCase "引用数组损坏", "NO.11111111", "11111111", Array(10), 10, VATS2_GROUP_INVALID_INPUT, Empty, 0, 0, VATS2_GROUP_INVALID_RELATION, False, False, 0, 1, -999, 6
    CheckCase "不完整加本行复用保留两原因", "NO.12345678/345678/99999999", "12345678", Array(10), 20, VATS2_GROUP_NOT_COMPARABLE, Empty, 0, 1, 10, True
    CheckCase "冲突关联不属于当前引用", "NO.11111111|NO.11111111", "11111111", Array(10), 10, VATS2_GROUP_INVALID_INPUT, Empty, 0, 0, VATS2_GROUP_INVALID_CONFLICT, False, False, 0, 1, -999, 7
    CheckCase "缺失冲突组也不能本行重复求和", "NO.12345678/345678", "12345678", Array(10), 20, VATS2_GROUP_NOT_COMPARABLE, Empty, 0, 1, 8, True, False, 0, 1, -999, 8
    CheckCase "多A错误保留原索引和组位置", "NO.22222222/11111111", "11111111|22222222", Array(Null, 20), 20, VATS2_GROUP_AMOUNT_ERROR, Empty, 0, 0, 0, False, False, 0, 15
    CheckCase "真实空金额不跳过", "NO.11111111", "11111111", Array(Empty), 0, VATS2_GROUP_AMOUNT_ERROR, Empty, 0, 0, 0
    CheckCase "无关 B 的 same不影响当前行", "NO.12345678/345678|NO.11111111", "12345678|11111111", Array(10, 20), 20, VATS2_GROUP_COMPARED, 0, 0, 0, 0, False, False, 1
    If failureCount = 0 Then
        VATStage2GroupAmount_SelfTest = "PASS: " & caseCount & " cases" & vbCrLf & details
    Else
        VATStage2GroupAmount_SelfTest = "FAIL: " & failureCount & "/" & caseCount & " cases" & vbCrLf & details
    End If
    Exit Function
Unexpected:
    VATStage2GroupAmount_SelfTest = "FAIL: unexpected " & Err.Number & " " & Err.Description & vbCrLf & details
End Function

Private Sub CheckCase(ByVal title As String, ByVal texts As String, ByVal aText As String, ByVal fixture As Variant, _
    ByVal b As Variant, ByVal state As VATS2GroupState, ByVal diff As Variant, ByVal pf As Long, ByVal mf As Long, _
    ByVal reasons As Long, Optional ByVal sameRow As Boolean = False, Optional ByVal crossRow As Boolean = False, _
    Optional ByVal target As Long = 0, Optional ByVal aLower As Long = 1, Optional ByVal mLower As Long = -999, Optional ByVal mutation As Long = 0)
    Dim rows() As VATS2BRowMatchResult, ids() As Long, digits() As String, amounts() As Variant
    Dim conflicts As VATS2BConflictResult, result As VATS2GroupAmountResult, again As VATS2GroupAmountResult
    Dim parts As Variant, bParts As Variant, n As Long, mc As Long, i As Long, bi As Long, before As String, reason As String
    On Error GoTo Failed
    caseCount = caseCount + 1
    parts = Split(aText, "|"): n = UBound(parts) + 1
    ReDim digits(aLower To aLower + n - 1)
    For i = 0 To n - 1: digits(aLower + i) = parts(i): Next i
    bParts = Split(texts, "|"): n = UBound(bParts) + 1
    ReDim rows(-3 To n - 4): ReDim ids(9 To n + 8)
    For i = 0 To n - 1
        rows(i - 3) = VATStage2MatchBRow(CStr(bParts(i)), digits, UBound(parts) + 1)
        ids(i + 9) = 100 + i
    Next i
    conflicts = VATStage2ScanBConflicts(rows, ids, n)
    bi = target - 3
    If mLower = -999 Then mLower = aLower
    mc = UBound(fixture) - LBound(fixture) + 1
    If mc > 0 Then
        ReDim amounts(mLower To mLower + mc - 1)
        For i = 0 To mc - 1: amounts(mLower + i) = fixture(LBound(fixture) + i): Next i
    End If
    Select Case mutation
        Case 1
            rows(bi).MatchState = VATS2_BROW_INVALID_INPUT
            rows(bi).References(1).MatchKind = VATS2_INVALID_A_RANGE
            rows(bi).References(1).CandidateCount = 0: Erase rows(bi).References(1).Candidates
            rows(bi).UniqueMatchCount = 0: rows(bi).InvalidMatchCount = 1
        Case 2: rows(bi).References(1).CandidateCount = 0
        Case 3: rows(bi).References(1).MatchKind = VATS2_NOT_FOUND
        Case 4: conflicts.Status = VATS2_BCONFLICT_INVALID_RANGE
        Case 5: conflicts.Status = VATS2_BCONFLICT_INVALID_CONTRACT
        Case 6: Erase rows(bi).References
        Case 7: conflicts.ConflictGroups(1).Associations(1).ReferenceIndex = 99
        Case 8: conflicts.ConflictGroupCount = 0: Erase conflicts.ConflictGroups
    End Select
    before = InputsText(rows, n, conflicts, amounts, mc, mLower, b)
    result = VATStage2EvaluateBGroupAmount(rows(bi), bi, conflicts, amounts, b)
    again = VATStage2EvaluateBGroupAmount(rows(bi), bi, conflicts, amounts, b)
    VerifyResult result, rows(bi), bi, conflicts, amounts, b, state, diff, pf, mf, reasons, sameRow, crossRow
    VerifyResult again, rows(bi), bi, conflicts, amounts, b, state, diff, pf, mf, reasons, sameRow, crossRow
    Require result.ErrorReferenceIndex = again.ErrorReferenceIndex And result.ErrorAIndex = again.ErrorAIndex, "重复调用映射错误一致"
    Require InputsText(rows, n, conflicts, amounts, mc, mLower, b) = before, "全部输入值及嵌套关系未修改"
    details = details & "PASS " & title & vbCrLf
    Exit Sub
Failed:
    reason = Err.Description: failureCount = failureCount + 1
    details = details & "FAIL " & title & ": " & reason & vbCrLf
End Sub

Private Sub VerifyResult(ByRef result As VATS2GroupAmountResult, ByRef row As VATS2BRowMatchResult, ByVal bi As Long, _
    ByRef conflicts As VATS2BConflictResult, ByRef amounts() As Variant, ByVal b As Variant, ByVal state As VATS2GroupState, _
    ByVal diff As Variant, ByVal pf As Long, ByVal mf As Long, ByVal reasons As Long, ByVal sameRow As Boolean, ByVal crossRow As Boolean)
    Dim indexes() As Long, values() As Variant, expected As VATS2AmountResult, i As Long, ai As Long
    Require result.BIndex = bi, "来源 BIndex"
    Require result.SourceMatchState = row.MatchState And result.SourceConflictStatus = conflicts.Status, "来源状态"
    Require result.ReferenceCount = row.ReferenceCount And result.UniqueMatchCount = row.UniqueMatchCount, "引用摘要"
    Require result.NotFoundCount = row.NotFoundCount And result.MultipleMatchCount = row.MultipleMatchCount And result.InvalidMatchCount = row.InvalidMatchCount, "不完整来源摘要"
    Require result.GroupState = state, "GroupState"
    Require result.Reasons = reasons, "不可比较原因"
    Require result.ParserFlags = pf And result.ParserFlags = row.ParserFlags, "ParserFlags"
    Require result.MatcherFlags = mf And result.MatcherFlags = row.MatcherFlags, "MatcherFlags"
    Require result.SameBRowReuse = sameRow And result.CrossBRowReuse = crossRow, "当前行冲突而非组级套用"
    If state = VATS2_GROUP_COMPARED Or state = VATS2_GROUP_AMOUNT_ERROR Then
        Require result.AmountEvaluated, "必须实际调用 Amount"
        ReDim indexes(1 To row.ReferenceCount): ReDim values(1 To row.ReferenceCount)
        For i = 1 To row.ReferenceCount
            ai = row.References(i).Candidates(1).AIndex
            indexes(i) = ai: values(i) = amounts(ai)
        Next i
        expected = VATStage2CompareAmounts(indexes, values, row.ReferenceCount, b)
        Require AmountText(result.Amount) = AmountText(expected), "Amount 全部结果及错误字段完整传播"
        If state = VATS2_GROUP_COMPARED Then
            Require result.Amount.Status = VATS2_AMOUNT_OK, "金额成功状态"
            Require result.Amount.Difference = CDec(diff), "差额方向及独立期望"
            Require VarType(result.Amount.Difference) = vbDecimal And VarType(result.Amount.AAmountSum) = vbDecimal, "Decimal精度"
            For i = 1 To row.ReferenceCount
                Require result.Amount.AIndexes(i) = indexes(i), "AIndex与Parser引用顺序不变"
                Require result.Amount.AAmounts(i) = CDec(amounts(indexes(i))), "AIndex直接金额映射"
            Next i
        Else
            Require result.Amount.Status <> VATS2_AMOUNT_OK And result.Amount.AmountState = VATS2_AMOUNT_NOT_COMPARED, "金额错误不是不一致"
        End If
    Else
        Require Not result.AmountEvaluated, "门控禁止调用 Amount"
        Require result.Amount.AmountState = VATS2_AMOUNT_NOT_COMPARED And IsEmpty(result.Amount.Difference), "门控无金额结论"
    End If
End Sub

Private Function AmountText(ByRef amount As VATS2AmountResult) As String
    Dim text As String, i As Long
    text = amount.Status & ":" & amount.AmountState & ":" & amount.ACount & ":" & amount.ErrorSide & ":" & amount.ErrorIndex & ":" & amount.ErrorAIndex & Pack(amount.ErrorReason)
    text = text & ValText(amount.AAmountSum) & ValText(amount.BAmount) & ValText(amount.Difference)
    For i = 1 To amount.ACount
        text = text & amount.AIndexes(i) & ValText(amount.AAmounts(i))
    Next i
    AmountText = text
End Function

Private Function InputsText(ByRef rows() As VATS2BRowMatchResult, ByVal n As Long, ByRef conflicts As VATS2BConflictResult, _
    ByRef amounts() As Variant, ByVal mc As Long, ByVal ml As Long, ByVal b As Variant) As String
    Dim text As String, i As Long, j As Long, k As Long, hi As Long
    For i = -3 To n - 4
        With rows(i)
            text = text & Pack(.OriginalText) & ":" & .MatchState & ":" & .ParserFlags & ":" & .MatcherFlags & ":" & .ReferenceCount
            text = text & ":" & .UniqueMatchCount & ":" & .NotFoundCount & ":" & .MultipleMatchCount & ":" & .InvalidMatchCount
        End With
        hi = 0
        On Error Resume Next
        hi = UBound(rows(i).References)
        Err.Clear: On Error GoTo 0
        text = text & "R" & hi
        For j = 1 To hi
            text = text & ReferenceText(rows(i).References(j))
        Next j
    Next i
    With conflicts
        text = text & "C" & .Status & ":" & .BRecordCount & ":" & .UniqueAssociationCount & ":" & .ConflictGroupCount & ":" & .ErrorBIndex & ":" & .ErrorReferenceIndex
        For i = 1 To .ConflictGroupCount
            With .ConflictGroups(i)
                text = text & .AIndex & Pack(.AInvoiceDigits) & ":" & .AssociationCount & ":" & .DistinctBRowCount & ":" & .Flags
                For j = 1 To .AssociationCount
                    With .Associations(j)
                        text = text & .BIndex & ":" & .BRowId & ":" & .ReferenceIndex & Pack(.ReferenceDigits) & .MatchKind & ":" & .AIndex & Pack(.AInvoiceDigits)
                    End With
                Next j
            End With
        Next i
    End With
    For k = 0 To mc - 1: text = text & ValText(amounts(ml + k)): Next k
    InputsText = text & ValText(b)
End Function

Private Function ReferenceText(ByRef reference As VATS2BRowReference) As String
    Dim text As String, hi As Long, j As Long
    With reference
        text = Pack(.Digits) & .Length & Pack(.RawFragment) & .StartIndex & Pack(.Context) & .ParserFlags & ":" & .MatchKind & ":" & .MatcherFlags & ":" & .CandidateCount
        On Error Resume Next
        hi = UBound(.Candidates)
        Err.Clear: On Error GoTo 0
        text = text & "K" & hi
        For j = 1 To hi: text = text & .Candidates(j).AIndex & Pack(.Candidates(j).InvoiceDigits): Next j
    End With
    ReferenceText = text
End Function

Private Function Pack(ByVal text As String) As String
    Pack = Len(text) & ":" & text
End Function

Private Function ValText(ByVal value As Variant) As String
    If IsNull(value) Then ValText = "Null" Else ValText = VarType(value) & Pack(CStr(value))
End Function

Private Sub Require(ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then Err.Raise vbObjectError + 2601, "GroupAmountTests", reason
End Sub
