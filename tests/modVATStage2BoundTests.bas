Attribute VB_Name = "modVATStage2BoundTests"
Option Explicit
Option Compare Binary

Private passed As Long, failed As Long, log As String
Private book As Workbook, a As Worksheet, b As Worksheet, color As Long
Private Const NUMBER1 As String = "11111111111112345678"
Private Const NUMBER2 As String = "22222222222287654321"

Public Function VATStage2BoundAuditRun_SelfTest() As String
    On Error GoTo Unexpected
    passed = 0: failed = 0: log = "": color = RGB(255, 255, 0)
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set a = book.Worksheets(1): a.Name = "合成A"
    Set b = book.Worksheets.Add(After:=a): b.Name = "合成B"
    TestChain
    TestHeaders
    TestStability
    TestReadOnly
    TestEntry
    log = log & "INFO SourceChanged并发窗口为代码级guard；纯内存比较逻辑已确定性验证，未引入DoEvents/timer或修改事件、计算策略。" & vbCrLf
    GoTo Done
Unexpected:
    failed = failed + 1: log = log & "UNEXPECTED " & Err.Number & " " & Err.Description & vbCrLf
Done:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    Set book = Nothing: Set a = Nothing: Set b = Nothing
    On Error GoTo 0
    If failed = 0 Then
        VATStage2BoundAuditRun_SelfTest = "PASS: " & passed & " assertions" & vbCrLf & log
    Else
        VATStage2BoundAuditRun_SelfTest = "FAIL: " & failed & "; PASS: " & passed & vbCrLf & log
    End If
End Function
Private Sub Fixture(Optional ByVal records As Boolean = True)
    a.Range("A1:P30").UnMerge: b.Range("A1:P30").UnMerge
    a.Range("A1:P30").Clear: b.Range("A1:P30").Clear
    a.Range("D3").Value2 = "数电发票号码": a.Range("I3").Value2 = VAT_FIELD_A
    b.Range("B5").Value2 = "供应商名称": b.Range("F5").Value2 = VAT_FIELD_B
    If records Then
        a.Range("D4").NumberFormat = "@": a.Range("D4").Value2 = NUMBER1
        b.Range("B6").Value2 = "采购合成样本NO." & NUMBER1
        a.Range("I4").Value2 = 10: a.Range("I4").Interior.Color = color
        b.Range("F6").Value2 = 10: b.Range("F6").Interior.Color = color
    End If
End Sub
Private Function Run() As VATS2BoundAuditRunResult
    Run = VATStage2RunBoundFullAudit(a.Range("D3"), a.Range("I3"), b.Range("B5"), b.Range("F5"), color)
End Function
Private Sub CheckBinding(ByRef r As VATS2BoundAuditRunResult)
    Dim sa As VATS2ASnapshotResult, sb As VATS2BSnapshotResult, fp As VATS2FingerprintResult, baseline As VATS2BaselineState
    sa = VATStage2ReadASnapshot(a, 3, 4, 9, color): sb = VATStage2ReadBSnapshot(b, 5, 2, 6, color)
    fp = VATStage2BuildFingerprints(sa, sb): baseline = VATStage2BuildBaseline(fp)
    Check "Bound技术成功", r.Status = VATS2_BOUND_RUN_OK
    Check "绑定版本", r.BindingProtocol = "VATRUN1" And r.FingerprintProtocol = "VATFP1" And r.DigestProtocol = "SHA256-UTF8"
    Check "当前A绑定", r.ARecordCount = baseline.ARecordCount And r.ABatchDigest = baseline.ABatchDigest
    Check "当前B绑定", r.BRecordCount = baseline.BRecordCount And r.BBatchDigest = baseline.BBatchDigest
    Check "当前Combined绑定", r.CombinedDigest = baseline.CombinedDigest And Len(r.CombinedDigest) = 64
    Check "内部Final和Finding契约", r.FinalDecision.Status = VATS2_FINAL_DECISION_OK And r.AuditFindings.Status = VATS2_AUDIT_OK And r.FinalDecision.FindingCount = r.AuditFindings.FindingCount
End Sub
Private Sub TestChain()
    Dim r As VATS2BoundAuditRunResult, other As VATS2BoundAuditRunResult, i As Long
    Call Fixture: r = Run(): CheckBinding r
    Check "完整一致VERIFIED", r.FinalDecision.Verdict = VATS2_FINAL_VERIFIED And r.AuditFindings.FindingCount = 0
    Check "本次Aggregate可靠相等", r.Aggregate.Status = VATS2_AGGREGATE_OK And r.Aggregate.ReliableEqualForShortSuffix
    a.Range("D4").Value2 = NUMBER2: b.Range("B6").Value2 = "NO." & NUMBER2
    other = Run(): CheckBinding other
    Check "两批均VERIFIED仍绑定各自身份", other.FinalDecision.Verdict = VATS2_FINAL_VERIFIED And other.CombinedDigest <> r.CombinedDigest
    Call Fixture
    b.Range("B6").Value2 = "NO.99999999999999999999": other = Run(): CheckBinding other
    Check "同金额不同批不拼接旧VERIFIED", other.CombinedDigest <> r.CombinedDigest And other.FinalDecision.Verdict = VATS2_FINAL_REVIEW_REQUIRED
    Check "本次Finding含B_ONLY", other.AuditFindings.CodeCounts(VATS2_FIND_B_ONLY) > 0
    Check "本次Finding含A_ONLY", other.AuditFindings.CodeCounts(VATS2_FIND_A_ONLY) > 0
    Call Fixture: b.Range("F6").Value2 = 11: r = Run(): CheckBinding r
    Check "总体金额不等REVIEW", r.FinalDecision.Verdict = VATS2_FINAL_REVIEW_REQUIRED And r.Aggregate.Difference = CDec(-1)
    Check "组金额Finding及Final来源", r.AuditFindings.CodeCounts(VATS2_FIND_AMOUNT_MISMATCH) = 1 And r.FinalDecision.HasFindings And r.FinalDecision.HasAggregateMismatch
    other = Run()
    Check "包含组成员的Finding完整重复确定", FindingsSignature(r.AuditFindings) = FindingsSignature(other.AuditFindings) And r.AuditFindings.Findings(1).MemberCount > 0
    For i = 1 To 3
        Call Fixture
        Select Case i
            Case 1: a.Range("I4").NumberFormat = "@": a.Range("I4").Value2 = "10"
            Case 2: a.Range("I4").Formula = "=""""": a.Range("I4").Calculate
            Case 3: a.Range("I4").Value2 = CVErr(xlErrNA)
        End Select
        r = Run(): CheckBinding r
        Check "扫描阻断仍合法UNAVAILABLE" & i, r.Aggregate.Status = VATS2_SCAN_BLOCKED And r.FinalDecision.Verdict = VATS2_FINAL_UNAVAILABLE And Not r.FinalDecision.RequiresManualReview
        Check "阻断不发布总体结论" & i, IsEmpty(r.FinalDecision.TotalsEqual) And IsEmpty(r.FinalDecision.TotalA) And r.Aggregate.AIssueCount > 0
    Next i
    Call Fixture
    a.Range("D4").Value2 = "11111111111111123456": b.Range("B6").Value2 = "NO.123456"
    r = Run(): CheckBinding r
    Check "本次可靠Aggregate驱动短尾号豁免", r.Aggregate.ReliableEqualForShortSuffix And r.AuditFindings.ShortSuffixWaivedCount = 1 And r.AuditFindings.CodeCounts(VATS2_FIND_SHORT_SUFFIX_REVIEW) = 0 And r.FinalDecision.Verdict = VATS2_FINAL_VERIFIED
    '额外完成色真空金额产生warning；不能把前一次可靠资格沿用到本次。
    a.Range("I5").Interior.Color = color: r = Run(): CheckBinding r
    Check "warning使本次资格失效", r.Aggregate.TotalsEqual And r.Aggregate.AWarningCount = 1 And Not r.Aggregate.ReliableEqualForShortSuffix
    Check "warning不沿用旧豁免", r.AuditFindings.ShortSuffixWaivedCount = 0 And r.AuditFindings.CodeCounts(VATS2_FIND_SHORT_SUFFIX_REVIEW) = 1 And r.FinalDecision.Verdict = VATS2_FINAL_REVIEW_REQUIRED
    Call Fixture
    a.Range("D5").NumberFormat = "@": a.Range("D5").Value2 = NUMBER2: a.Range("I5").Value2 = 20
    b.Range("B6").Value2 = "NO." & NUMBER1 & "/" & NUMBER2: b.Range("F6").Value2 = 30
    r = Run(): CheckBinding r
    Check "fallback补全成员进入内部金额链", r.FinalDecision.EqualCount = 1 And r.AuditFindings.CodeCounts(VATS2_FIND_A_COLOR_MISSING) = 1 And r.AuditFindings.CodeCounts(VATS2_FIND_AMOUNT_MISMATCH) = 0
    Fixture False: r = Run(): CheckBinding r
    Check "零记录合法VERIFIED绑定", r.ARecordCount = 0 And r.BRecordCount = 0 And r.FinalDecision.Verdict = VATS2_FINAL_VERIFIED And r.FinalDecision.NoCompletedRecords
    Call Fixture: r = Run()
    a.Range("D4:I4").Copy Destination:=a.Range("D8"): a.Range("D4:I4").Clear
    b.Range("B6:F6").Copy Destination:=b.Range("B9"): b.Range("B6:F6").Clear
    other = Run(): CheckBinding other
    Check "独立完整运行移动行不改长期身份", r.CombinedDigest = other.CombinedDigest And r.ABatchDigest = other.ABatchDigest And r.BBatchDigest = other.BBatchDigest And other.FinalDecision.Verdict = VATS2_FINAL_VERIFIED
    Call Fixture
    a.Range("D5").NumberFormat = "@": a.Range("D5").Value2 = NUMBER2: a.Range("I5").Value2 = 20: a.Range("I5").Interior.Color = color
    b.Range("B7").Value2 = "NO." & NUMBER2: b.Range("F7").Value2 = 20: b.Range("F7").Interior.Color = color
    r = Run()
    a.Range("D4:I4").Copy a.Range("D20"): a.Range("D5:I5").Copy a.Range("D4"): a.Range("D20:I20").Copy a.Range("D5"): a.Range("D20:I20").Clear
    b.Range("B6:F6").Copy b.Range("B20"): b.Range("B7:F7").Copy b.Range("B6"): b.Range("B20:F20").Copy b.Range("B7"): b.Range("B20:F20").Clear
    other = Run(): CheckBinding other
    Check "独立完整运行两侧排序保持长期batch", r.FinalDecision.Verdict = VATS2_FINAL_VERIFIED And other.FinalDecision.Verdict = VATS2_FINAL_VERIFIED And r.CombinedDigest = other.CombinedDigest
    '非法UTF-16可被Snapshot保留，但冻结SHA明确拒绝；自然触发最后阶段失败。
    Call Fixture: b.Range("B6").Value2 = "NO." & NUMBER1 & ChrW$(&HD800)
    r = Run()
    Check "摘要阶段错误可定位", r.Status = VATS2_BOUND_RUN_STAGE_FAILED And r.ErrorStage = "BASELINE_DIGEST"
    CheckUnavailable r
End Sub
Private Sub TestHeaders()
    Dim r As VATS2BoundAuditRunResult, ah As Range, am As Range, bh As Range, bm As Range, none As Range, i As Long
    Dim closedBook As Workbook
    For i = 1 To 13
        Call Fixture
        Set ah = a.Range("D3"): Set am = a.Range("I3"): Set bh = b.Range("B5"): Set bm = b.Range("F5")
        Select Case i
            Case 1: Set ah = b.Range("D3")
            Case 2: Set am = a.Range("I2")
            Case 3: Set bh = a.Range("B5")
            Case 4: Set bm = b.Range("F4")
            Case 5: Set ah = a.Range("D3:E3")
            Case 6: Set bm = b.Range("F5:F6")
            Case 7: Set ah = none
            Case 8: Set bm = none
            Case 9: Set ah = am
            Case 10: a.Range("D3:E3").Merge
            Case 11: a.Range("I3").Value2 = "税额"
            Case 12: b.Range("F5").Value2 = "税额 "
            Case 13: a.Range("D3").ClearContents
        End Select
        r = VATStage2RunBoundFullAudit(ah, am, bh, bm, color)
        Check "非法header拒绝" & i, r.Status = VATS2_BOUND_RUN_INVALID_INPUT And r.ErrorStage = "INPUT"
        CheckUnavailable r
    Next i
    Call Fixture
    Set closedBook = Application.Workbooks.Add(xlWBATWorksheet)
    Set ah = closedBook.Worksheets(1).Range("D3"): Set am = closedBook.Worksheets(1).Range("I3")
    ah.Value2 = "数电发票号码": am.Value2 = VAT_FIELD_A: closedBook.Close False
    r = VATStage2RunBoundFullAudit(ah, am, b.Range("B5"), b.Range("F5"), color)
    Check "已关闭Workbook表头拒绝", r.Status = VATS2_BOUND_RUN_INVALID_INPUT And r.ErrorStage = "INPUT"
    CheckUnavailable r
End Sub
Private Sub CheckUnavailable(ByRef r As VATS2BoundAuditRunResult)
    Check "失败不发布部分digest", r.BindingProtocol = "" And r.ABatchDigest = "" And r.BBatchDigest = "" And r.CombinedDigest = "" And r.ARecordCount = 0 And r.BRecordCount = 0
    Check "失败不发布可信Final或Finding", r.FinalDecision.Status <> VATS2_FINAL_DECISION_OK And r.FinalDecision.Verdict = VATS2_FINAL_UNAVAILABLE And r.AuditFindings.Status <> VATS2_AUDIT_OK And r.AuditFindings.FindingCount = 0 And Len(r.ErrorReason) > 0
End Sub
Private Sub TestStability()
    Dim sa As VATS2ASnapshotResult, sb As VATS2BSnapshotResult, ca As VATS2ASnapshotResult, cb As VATS2BSnapshotResult
    Dim fp As VATS2FingerprintResult, cp As VATS2FingerprintResult, ar As VATS2ASnapshotRecord, i As Long, before As String
    Call Fixture
    a.Range("D5").NumberFormat = "@": a.Range("D5").Value2 = NUMBER2: a.Range("I5").Value2 = 2
    sa = VATStage2ReadASnapshot(a, 3, 4, 9, color): sb = VATStage2ReadBSnapshot(b, 5, 2, 6, color)
    fp = VATStage2BuildFingerprints(sa, sb): ca = sa: cb = sb: cp = fp
    Check "同次确认一致", VATStage2BoundSourcesStable(sa, sb, fp, ca, cb, cp)
    For i = 1 To 12
        ca = sa: cb = sb: cp = fp
        Select Case i
            Case 1: ca.Records(1).AmountRaw = 12: cp = VATStage2BuildFingerprints(ca, cb)
            Case 2: cb.Records(1).SupplierTextRaw = "变更": cp = VATStage2BuildFingerprints(ca, cb)
            Case 3: ca.Records(1).ExcelRow = 7
            Case 4: cb.Records(1).ExcelRow = 8
            Case 5: ca.RecordCount = 1
            Case 6: cb.RecordCount = 0
            Case 7: ca.Records(1).IsCompleted = False: cp = VATStage2BuildFingerprints(ca, cb)
            Case 8: cb.Records(1).AmountHasFormula = True: cp = VATStage2BuildFingerprints(ca, cb)
            Case 9
                ar = ca.Records(1): ca.Records(1) = ca.Records(2): ca.Records(2) = ar
                ca.Records(1).AIndex = 1: ca.Records(1).ExcelRow = 4: ca.Records(2).AIndex = 2: ca.Records(2).ExcelRow = 5
                cp = VATStage2BuildFingerprints(ca, cb)
            Case 10: cp.Status = VATS2_FINGERPRINT_INVALID_INPUT
            Case 11: ca.Status = VATS2_SNAPSHOT_READ_ERROR
            Case 12: Erase cp.BFingerprints
        End Select
        Check "同次变化或契约损坏拒绝" & i, Not VATStage2BoundSourcesStable(sa, sb, fp, ca, cb, cp)
        If i = 9 Then Check "长期batch相同仍拒绝当次重排", cp.CombinedFingerprint = fp.CombinedFingerprint
    Next i
    ca = sa: cb = sb: cp = fp: before = fp.CombinedFingerprint
    Check "确认重复确定", VATStage2BoundSourcesStable(sa, sb, fp, ca, cb, cp) = VATStage2BoundSourcesStable(sa, sb, fp, ca, cb, cp)
    Check "确认不修改快照及签名", sa.Records(1).AmountRaw = 10 And sa.Records(1).ExcelRow = 4 And sa.Records(1).InvoiceDigitsRaw = NUMBER1 And fp.CombinedFingerprint = before And cp.CombinedFingerprint = before
End Sub
Private Sub TestReadOnly()
    Dim r As VATS2BoundAuditRunResult, s As VATS2BoundAuditRunResult, before As String, saved As Boolean
    Dim count As Long, calc As XlCalculation, events As Boolean, i As Long
    Call Fixture
    a.Range("I4").Formula = "=5+5": a.Range("I4").Calculate
    b.Range("F6").Formula = "=20/2": b.Range("F6").Calculate
    b.Range("B6").Value2 = "NO.99999999999999999999"
    before = Evidence(a) & Evidence(b): saved = book.saved: count = Application.Workbooks.count
    calc = Application.Calculation: events = Application.EnableEvents
    r = Run(): s = Run()
    Check "值公式颜色未改", before = Evidence(a) & Evidence(b)
    Check "未保存状态保持", book.saved = saved
    Check "工作簿未保存未关闭", VATWorkbookOpen(book) And Application.Workbooks.count = count And book.Path = ""
    Check "Excel策略保持", Application.Calculation = calc And Application.EnableEvents = events
    Check "重复digest和Final全部字段确定", r.CombinedDigest = s.CombinedDigest And FinalSignature(r.FinalDecision) = FinalSignature(s.FinalDecision)
    Check "重复Findings完整结构确定", FindingsSignature(r.AuditFindings) = FindingsSignature(s.AuditFindings)
    '只更改本测试工作簿的Saved标志，验证入口不会变脏或保存；不调用Save。
    book.saved = True: r = Run(): Check "干净Workbook状态保持", book.saved
    book.saved = False
End Sub
Private Sub TestEntry()
    Dim cm As Object, source As String, signature As String, p As Long
    Set cm = ThisWorkbook.VBProject.VBComponents("modVATStage2BoundAuditRun").CodeModule
    source = cm.Lines(1, cm.CountOfLines)
    p = InStr(source, "Public Function VATStage2RunBoundFullAudit")
    signature = Mid$(source, p, InStr(p, source, ") As VATS2BoundAuditRunResult") - p)
    Check "主入口只接受四个Range", (Len(signature) - Len(Replace(signature, "As Range", ""))) / Len("As Range") = 4
    Check "主入口不接受外部审核结果", InStr(signature, "FinalDecision") = 0 And InStr(signature, "Fingerprint") = 0 And InStr(signature, "Findings") = 0 And InStr(signature, "Effective") = 0
    Check "新模块无写入或执行让步", InStr(source, "DoEvents") = 0 And InStr(source, "SaveBaseline") = 0 And InStr(source, "LoadBaseline") = 0 And InStr(source, ".Save") = 0 And InStr(source, ".Close") = 0 And InStr(source, "EnableEvents =") = 0 And InStr(source, "Calculation =") = 0
End Sub
Private Function Evidence(ByVal sheet As Worksheet) As String
    Dim cell As Range, s As String
    For Each cell In sheet.Range("A1:P15")
        s = s & Pack(CStr(VarType(cell.Value2))) & Pack(CStr(cell.Value2)) & Pack(CStr(cell.Formula)) & Pack(CStr(cell.HasFormula)) & Pack(CStr(cell.Interior.Pattern)) & Pack(CStr(cell.Interior.Color))
    Next cell
    Evidence = s
End Function
Private Function Pack(ByVal value As String) As String
    Pack = CStr(Len(value)) & ":" & value
End Function
Private Function FinalSignature(ByRef r As VATS2FinalDecisionResult) As String
    Dim value As Variant, s As String, i As Long
    For Each value In Array(r.Status, r.Verdict, r.AggregateStatus, r.EffectiveAmountStatus, r.AuditStatus, r.TotalA, r.TotalB, r.Difference, r.TotalsEqual, _
        r.ACompletedCount, r.BCompletedCount, r.AIncludedCount, r.BIncludedCount, r.AWarningCount, r.BWarningCount, r.EffectiveBCount, r.ComparedCount, r.EqualCount, _
        r.MismatchCount, r.NotComparableCount, r.AmountErrorCount, r.RelationQualityBlockedCount, r.FindingCount, r.ShortSuffixWaivedCount, _
        r.HasAggregateMismatch, r.HasAggregateWarnings, r.HasFindings, r.NoCompletedRecords, r.RequiresManualReview, r.ErrorSource, r.ErrorReason)
        s = s & Pack(CStr(VarType(value))) & Pack(CStr(value))
    Next value
    For i = 1 To 18: s = s & Pack(CStr(r.CodeCounts(i))): Next i
    FinalSignature = s
End Function
Private Function FindingsSignature(ByRef r As VATS2AuditFindingsResult) As String
    Dim value As Variant, s As String, i As Long, j As Long
    s = Pack(CStr(r.Status)) & Pack(CStr(r.FindingCount)) & Pack(CStr(r.ShortSuffixWaivedCount)) & Pack(r.ErrorSide) & Pack(CStr(r.ErrorIndex)) & Pack(CStr(r.ErrorReferenceIndex)) & Pack(r.ErrorReason)
    For i = 1 To 18: s = s & Pack(CStr(r.CodeCounts(i))): Next i
    For i = 1 To r.FindingCount
        With r.Findings(i)
            For Each value In Array(.Code, .Side, .AIndex, .AExcelRow, .BIndex, .BExcelRow, .ReferenceIndex, .ReferenceDigits, .SourceFlags, .ReasonFlags, .ParserFlags, .MatcherFlags, .MatchKind, .GroupIndex, .MemberCount, .Difference, .GroupState, .AmountStatus, .AmountErrorSide, .AmountErrorIndex, .RawVarType, .Detail)
                s = s & Pack(CStr(VarType(value))) & Pack(CStr(value))
            Next value
            For j = 1 To .MemberCount
                With .Members(j)
                    s = s & Pack(CStr(.AIndex)) & Pack(CStr(.AExcelRow)) & Pack(CStr(.BIndex)) & Pack(CStr(.BExcelRow)) & Pack(CStr(.ReferenceIndex)) & Pack(.ReferenceDigits)
                End With
            Next j
        End With
    Next i
    FindingsSignature = s
End Function
Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        passed = passed + 1
    Else
        failed = failed + 1: log = log & "FAIL " & name & vbCrLf
    End If
End Sub
