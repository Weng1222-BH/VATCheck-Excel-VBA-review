Attribute VB_Name = "modVATStage2SealTests"
Option Explicit
Option Compare Binary

Private passed As Long, failed As Long, log As String
Private book As Workbook, a As Worksheet, b As Worksheet, color As Long
Private fs As Object, root As String
Private Const INV As String = "11111111111112345678"
Private Const SUPPLIER As String = "合成隐私样本NO.11111111111112345678"
Private Const AMOUNT As Double = 987654.321

Public Function VATStage2VerificationSeal_SelfTest() As String
    On Error GoTo Unexpected
    passed = 0: failed = 0: log = "": color = RGB(255, 255, 0)
    Set fs = CreateObject("Scripting.FileSystemObject")
    root = fs.BuildPath(fs.GetSpecialFolder(2).Path, "VATCheck-C75-" & fs.GetTempName)
    fs.CreateFolder root
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set a = book.Worksheets(1): Set b = book.Worksheets.Add(After:=a)
    TestBuild
    TestStore
    TestReuse
    TestPurityAndOrder
    GoTo Done
Unexpected:
    failed = failed + 1: log = log & "UNEXPECTED " & Err.Number & " " & Err.Description & vbCrLf
Done:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    '只删除本测试创建且父目录为系统临时目录的唯一目录。
    If Len(root) > 0 Then
        If fs.GetParentFolderName(fs.GetAbsolutePathName(root)) = fs.GetSpecialFolder(2).Path Then
            If Left$(fs.GetFileName(root), 12) = "VATCheck-C75-" Then fs.DeleteFolder root, True
        End If
    End If
    On Error GoTo 0
    If failed = 0 Then
        VATStage2VerificationSeal_SelfTest = "PASS: " & passed & " assertions" & vbCrLf & log
    Else
        VATStage2VerificationSeal_SelfTest = "FAIL: " & failed & "; PASS: " & passed & vbCrLf & log
    End If
End Function

Private Sub Fixture(Optional ByVal records As Boolean = True)
    a.Range("A1:P30").Clear: b.Range("A1:P30").Clear
    a.Range("D3").Value2 = "数电发票号码": a.Range("I3").Value2 = VAT_FIELD_A
    b.Range("B5").Value2 = "供应商名称": b.Range("F5").Value2 = VAT_FIELD_B
    If records Then
        a.Range("D4").NumberFormat = "@": a.Range("D4").Value2 = INV
        b.Range("B6").Value2 = SUPPLIER
        a.Range("I4").Value2 = AMOUNT: a.Range("I4").Interior.Color = color
        b.Range("F6").Value2 = AMOUNT: b.Range("F6").Interior.Color = color
    End If
End Sub
Private Function Run() As VATS2BoundAuditRunResult
    Run = VATStage2RunBoundFullAudit(a.Range("D3"), a.Range("I3"), b.Range("B5"), b.Range("F5"), color)
End Function
Private Function CurrentFP() As VATS2FingerprintResult
    Dim sa As VATS2ASnapshotResult, sb As VATS2BSnapshotResult
    sa = VATStage2ReadASnapshot(a, 3, 4, 9, color): sb = VATStage2ReadBSnapshot(b, 5, 2, 6, color)
    CurrentFP = VATStage2BuildFingerprints(sa, sb)
End Function
Private Function Stamp(ByRef r As VATS2VerificationSealState) As String
    Stamp = r.Status & ":" & r.SchemaVersion & ":" & r.BindingProtocol & ":" & r.PolicyProtocol & ":" & _
        r.FingerprintProtocol & ":" & r.DigestProtocol & ":" & r.ARecordCount & ":" & r.BRecordCount & ":" & _
        r.VerifiedUtc & ":" & r.ABatchDigest & ":" & r.BBatchDigest & ":" & r.CombinedDigest
End Function

Private Sub TestBuild()
    Dim original As VATS2BoundAuditRunResult, r As VATS2BoundAuditRunResult, seal As VATS2VerificationSealState
    Dim fp As VATS2FingerprintResult, baseline As VATS2BaselineState, i As Long
    Call Fixture: original = Run(): seal = VATStage2BuildVerificationSeal(original)
    Check "受控真实VERIFIED生成Seal", original.FinalDecision.Verdict = VATS2_FINAL_VERIFIED And seal.Status = VATS2_VERIFY_OK
    fp = CurrentFP(): baseline = VATStage2BuildBaseline(fp)
    Check "绑定内部三项digest", seal.ABatchDigest = baseline.ABatchDigest And seal.BBatchDigest = baseline.BBatchDigest And seal.CombinedDigest = baseline.CombinedDigest
    Check "绑定数量", seal.ARecordCount = 1 And seal.BRecordCount = 1
    Check "固定五个协议", seal.SchemaVersion = "VATVERIFY1" And seal.BindingProtocol = "VATRUN1" And seal.PolicyProtocol = "VATSTAGE2-POLICY1" And seal.FingerprintProtocol = "VATFP1" And seal.DigestProtocol = "SHA256-UTF8"
    Check "UTC存在", Len(seal.VerifiedUtc) = 20 And Right$(seal.VerifiedUtc, 1) = "Z"
    For i = 1 To 38
        r = original
        Select Case i
            Case 1: r.FinalDecision.Verdict = VATS2_FINAL_REVIEW_REQUIRED
            Case 2: r.FinalDecision.Verdict = VATS2_FINAL_UNAVAILABLE
            Case 3: r.Status = VATS2_BOUND_RUN_SOURCE_CHANGED
            Case 4: r.Aggregate.Status = VATS2_SCAN_BLOCKED
            Case 5: r.AuditFindings.FindingCount = 1
            Case 6: r.Aggregate.AWarningCount = 1
            Case 7: r.Aggregate.BWarningCount = 1
            Case 8: r.ABatchDigest = String$(64, "G")
            Case 9: r.BBatchDigest = ""
            Case 10: r.CombinedDigest = UCase$(r.CombinedDigest)
            Case 11: r.BindingProtocol = "VATRUN2"
            Case 12: r.FingerprintProtocol = "VATFP2"
            Case 13: r.DigestProtocol = "OTHER"
            Case 14: r.FinalDecision.ACompletedCount = 0
            Case 15: r.FinalDecision.BIncludedCount = 0
            Case 16: r.FinalDecision.RequiresManualReview = True
            Case 17: r.FinalDecision.HasAggregateMismatch = True
            Case 18: r.FinalDecision.HasAggregateWarnings = True
            Case 19: r.FinalDecision.HasFindings = True
            Case 20: r.FinalDecision.FindingCount = 1
            Case 21: r.Aggregate.TotalsEqual = 1
            Case 22: r.Aggregate.TotalsEqual = False
            Case 23: r.FinalDecision.Status = VATS2_FINAL_INVALID_CONTRACT
            Case 24: r.AuditFindings.Status = VATS2_AUDIT_INVALID_INPUT
            Case 25: r.ARecordCount = -1
            Case 26: r.BRecordCount = -1
            Case 27: r.FinalDecision.TotalA = CDec(0)
            Case 28: r.FinalDecision.Difference = CDec(1)
            Case 29: r.FinalDecision.CodeCounts(18) = 1
            Case 30: r.AuditFindings.CodeCounts(18) = 1
            Case 31: r.FinalDecision.ShortSuffixWaivedCount = -1
            Case 32: r.FinalDecision.AggregateStatus = VATS2_CALCULATION_PENDING
            Case 33: r.FinalDecision.NoCompletedRecords = True
            Case 34: r.Aggregate.AScanSucceeded = False
            Case 35: r.Aggregate.BIssueCount = 1
            Case 36: r.FinalDecision.AWarningCount = 1
            Case 37: r.FinalDecision.TotalsEqual = Null
            Case 38: r.FinalDecision.TotalB = CDbl(r.FinalDecision.TotalB)
        End Select
        seal = VATStage2BuildVerificationSeal(r)
        Check "Build拒绝必要契约" & i, seal.Status = VATS2_VERIFY_INVALID_INPUT
        Check "Build失败不发布部分凭证" & i, seal.CombinedDigest = "" And seal.SchemaVersion = "" And seal.ARecordCount = 0
    Next i
    seal = VATStage2BuildVerificationSeal(original)
    Check "Build不改变原运行", original.Status = VATS2_BOUND_RUN_OK And original.FinalDecision.Verdict = VATS2_FINAL_VERIFIED And original.CombinedDigest = baseline.CombinedDigest And original.Aggregate.TotalA = CDec(AMOUNT)
    Check "重复Build摘要确定", seal.CombinedDigest = baseline.CombinedDigest
    b.Range("F6").Value2 = AMOUNT + 1: r = Run(): seal = VATStage2BuildVerificationSeal(r)
    Check "真实REVIEW拒绝", r.FinalDecision.Verdict = VATS2_FINAL_REVIEW_REQUIRED And seal.Status = VATS2_VERIFY_INVALID_INPUT
    a.Range("I4").Formula = "=""""": a.Range("I4").Calculate
    r = Run(): seal = VATStage2BuildVerificationSeal(r)
    Check "真实UNAVAILABLE拒绝", r.FinalDecision.Verdict = VATS2_FINAL_UNAVAILABLE And seal.Status = VATS2_VERIFY_INVALID_INPUT
    Fixture False: r = Run(): seal = VATStage2BuildVerificationSeal(r)
    Check "零记录VERIFIED允许", r.FinalDecision.NoCompletedRecords And seal.Status = VATS2_VERIFY_OK And seal.ARecordCount = 0 And seal.BRecordCount = 0
End Sub

Private Sub TestStore()
    Dim r As VATS2BoundAuditRunResult, seal As VATS2VerificationSealState, invalid As VATS2VerificationSealState
    Dim loaded As VATS2VerificationSealLoadResult, io As VATS2VerificationSealIOResult
    Dim fp As VATS2FingerprintResult, text As String, changed As String, parts As Variant, i As Long, expectedField As String
    Dim before As String, f As Integer, pending As String, target As String, canary As Variant
    Call Fixture: r = Run()
    r.AuditFindings.ErrorReason = "合成Finding隐私哨兵"
    seal = VATStage2BuildVerificationSeal(r): fp = CurrentFP()
    target = fs.BuildPath(root, "verified-v1.dat"): pending = fs.BuildPath(root, "verified-v1.pending.tmp")
    loaded = VATStage2LoadVerificationSeal(root)
    Check "NOT_FOUND安全默认", loaded.Status = VATS2_VERIFY_NOT_FOUND And loaded.Seal.CombinedDigest = ""
    WriteText fs.BuildPath(root, "baseline-v1.dat"), "BASELINE_SENTINEL"
    before = Stamp(seal): io = VATStage2SaveVerificationSeal(seal, root)
    Check "保存成功", io.Status = VATS2_VERIFY_OK
    loaded = VATStage2LoadVerificationSeal(root)
    Check "Roundtrip字段完整", loaded.Status = VATS2_VERIFY_OK And Stamp(loaded.Seal) = before
    Check "Save不修改输入", Stamp(seal) = before
    text = ReadText(target): parts = Split(text, vbLf)
    Check "固定ASCII LF和EOF", UBound(parts) = 12 And parts(12) = "" And InStr(text, vbCr) = 0
    Check "校验和复用冻结SHA", parts(11) = VATStage2SHA256(Left$(text, Len(text) - 65))
    For Each canary In Array(INV, SUPPLIER, CStr(AMOUNT), fp.AFingerprints(1), fp.BFingerprints(1), fp.CombinedFingerprint, "合成Finding隐私哨兵", book.FullName)
        Check "直接文件隐私排除", InStr(1, text, CStr(canary), vbBinaryCompare) = 0
    Next canary
    For i = 1 To Len(text)
        If AscW(Mid$(text, i, 1)) > 127 Then Err.Raise 5, , "非ASCII文件"
    Next i
    Check "字节内容为ASCII", True
    io = VATStage2SaveVerificationSeal(seal, root)
    Check "重复保存字节确定", io.Status = VATS2_VERIFY_OK And ReadText(target) = text
    loaded = VATStage2LoadVerificationSeal(root)
    Check "重复读取确定", Stamp(loaded.Seal) = before
    '每次重算checksum，确保版本/数据解析本身拒绝，不仅是checksum错误。
    For i = 1 To 23
        parts = Split(text, vbLf): expectedField = ""
        Select Case i
            Case 1: parts(0) = "WRONG"
            Case 2: parts(0) = "VATVERIFY2": expectedField = "SCHEMA"
            Case 3: parts(1) = "VATRUN2": expectedField = "BINDING"
            Case 4: parts(2) = "VATSTAGE2-POLICY2": expectedField = "POLICY"
            Case 5: parts(3) = "VATFP2": expectedField = "FINGERPRINT"
            Case 6: parts(4) = "SHA1": expectedField = "DIGEST"
            Case 7: parts(8) = String$(64, "G")
            Case 8: parts(9) = String$(63, "a")
            Case 9: parts(10) = UCase$(parts(10))
            Case 10: parts(5) = "-1"
            Case 11: parts(6) = "01"
            Case 12: parts(5) = "2147483648"
            Case 13: parts(5) = "1.0"
            Case 14: parts(7) = "2026-02-29T01:02:03Z"
            Case 15: parts(7) = "2026-09-15T24:00:00Z"
            Case 16: parts(7) = "2026-09-15T00:00:00z"
            Case 17: parts(7) = "0000-01-01T00:00:00Z"
            Case 18: parts(6) = ""
            Case 19: parts(7) = "2026-13-01T00:00:00Z"
            Case 20: parts(8) = ""
        End Select
        changed = Rechecksum(parts)
        Select Case i
            Case 21: changed = Left$(changed, Len(changed) - 66) & "x" & Right$(changed, 65)
            Case 22: changed = Left$(changed, Len(changed) - 10)
            Case 23: changed = changed & "junk"
        End Select
        WriteText target, changed: loaded = VATStage2LoadVerificationSeal(root)
        If Len(expectedField) > 0 Then
            Check "不支持版本" & i, loaded.Status = VATS2_VERIFY_UNSUPPORTED_VERSION And loaded.UnsupportedField = expectedField
        Else
            Check "损坏拒绝" & i, loaded.Status = VATS2_VERIFY_CORRUPT
        End If
        Check "Load失败无部分Seal" & i, loaded.Seal.Status = loaded.Status And loaded.Seal.CombinedDigest = "" And loaded.Seal.PolicyProtocol = ""
    Next i
    WriteText target, text
    WriteText pending, "OTHER_WRITER"
    io = VATStage2SaveVerificationSeal(seal, root)
    Check "已有pending拒绝", io.Status = VATS2_VERIFY_IO_ERROR
    Check "不抢占pending或删除旧seal", ReadText(pending) = "OTHER_WRITER" And ReadText(target) = text
    fs.DeleteFile pending
    f = FreeFile: Open target For Binary Access Read Lock Read Write As #f
    io = VATStage2SaveVerificationSeal(seal, root)
    loaded = VATStage2LoadVerificationSeal(root)
    Close #f
    Check "替换失败保留旧文件", io.Status = VATS2_VERIFY_IO_ERROR And ReadText(target) = text
    Check "只清理自己pending", Not fs.FileExists(pending)
    Check "锁定读取IO_ERROR", loaded.Status = VATS2_VERIFY_IO_ERROR And loaded.Seal.CombinedDigest = ""
    invalid = seal: invalid.VerifiedUtc = "BAD"
    io = VATStage2SaveVerificationSeal(invalid, root)
    Check "非法state不写入", io.Status = VATS2_VERIFY_INVALID_INPUT And ReadText(target) = text
    invalid = seal: invalid.PolicyProtocol = "VATSTAGE2-POLICY2"
    io = VATStage2SaveVerificationSeal(invalid, root)
    Check "不保存未知policy", io.Status = VATS2_VERIFY_INVALID_INPUT And ReadText(target) = text
    invalid = seal: invalid.VerifiedUtc = "2024-02-29T23:59:59Z"
    io = VATStage2SaveVerificationSeal(invalid, root): loaded = VATStage2LoadVerificationSeal(root)
    Check "UTC闰年合法", io.Status = VATS2_VERIFY_OK And loaded.Seal.VerifiedUtc = invalid.VerifiedUtc
    Check "baseline完全不触碰", ReadText(fs.BuildPath(root, "baseline-v1.dat")) = "BASELINE_SENTINEL"
End Sub

Private Sub TestReuse()
    Dim runResult As VATS2BoundAuditRunResult, fp As VATS2FingerprintResult, changedFP As VATS2FingerprintResult
    Dim base As VATS2BaselineLoadResult, baseBad As VATS2BaselineLoadResult
    Dim seal As VATS2VerificationSealLoadResult, sealBad As VATS2VerificationSealLoadResult
    Dim delta As VATS2IncrementalDeltaResult, d As VATS2IncrementalDeltaResult
    Dim r As VATS2ReuseEligibilityResult, again As VATS2ReuseEligibilityResult, i As Long, expected As Long, before As String
    Call Fixture: runResult = Run(): fp = CurrentFP(): base.Baseline = VATStage2BuildBaseline(fp)
    seal.Seal = VATStage2BuildVerificationSeal(runResult): delta = VATStage2CompareBaselineDelta(fp, base)
    before = Stamp(seal.Seal)
    r = VATStage2EvaluateReuseEligibility(delta, base, seal)
    Check "三方完全一致可复用", r.Status = VATS2_REUSE_OK And r.Decision = VATS2_REUSE_ELIGIBLE And r.ReasonFlags = 0
    again = VATStage2EvaluateReuseEligibility(delta, base, seal)
    Check "重复调用确定", r.Status = again.Status And r.Decision = again.Decision And r.ReasonFlags = again.ReasonFlags And r.ErrorReason = again.ErrorReason
    For i = 1 To 12
        sealBad = seal: expected = VATS2_REUSE_SEAL_BASELINE_MISMATCH
        Select Case i
            Case 1: sealBad.Status = VATS2_VERIFY_NOT_FOUND: expected = VATS2_REUSE_SEAL_NOT_FOUND
            Case 2: sealBad.Seal.ARecordCount = 2
            Case 3: sealBad.Seal.BRecordCount = 2
            Case 4: sealBad.Seal.ABatchDigest = String$(64, "a")
            Case 5: sealBad.Seal.BBatchDigest = String$(64, "b")
            Case 6: sealBad.Seal.CombinedDigest = String$(64, "c")
            Case 7: sealBad.Seal.PolicyProtocol = "VATSTAGE2-POLICY2": expected = VATS2_REUSE_POLICY_VERSION_MISMATCH
            Case 8: sealBad.Status = VATS2_VERIFY_CORRUPT: expected = VATS2_REUSE_SEAL_UNAVAILABLE
            Case 9: sealBad.Status = VATS2_VERIFY_IO_ERROR: expected = VATS2_REUSE_SEAL_UNAVAILABLE
            Case 10: sealBad.Status = VATS2_VERIFY_UNSUPPORTED_VERSION: sealBad.UnsupportedField = "POLICY": sealBad.ErrorReason = "无关文本": expected = VATS2_REUSE_POLICY_VERSION_MISMATCH
            Case 11: sealBad.Status = VATS2_VERIFY_UNSUPPORTED_VERSION: sealBad.UnsupportedField = "BINDING": sealBad.ErrorReason = "POLICY": expected = VATS2_REUSE_SEAL_UNAVAILABLE
            Case 12: sealBad.Seal.SchemaVersion = "VATVERIFY2": expected = VATS2_REUSE_SEAL_UNAVAILABLE
        End Select
        r = VATStage2EvaluateReuseEligibility(delta, base, sealBad)
        Check "Seal阻止复用" & i, r.Status = VATS2_REUSE_OK And r.Decision = VATS2_FULL_VALIDATION_REQUIRED And (r.ReasonFlags And expected) <> 0
        If i = 11 Then Check "不解析ErrorReason判断policy", (r.ReasonFlags And VATS2_REUSE_POLICY_VERSION_MISMATCH) = 0
    Next i
    For i = 1 To 4
        baseBad = base
        Select Case i
            Case 1: baseBad.Status = VATS2_BASELINE_NOT_FOUND
            Case 2: baseBad.Status = VATS2_BASELINE_CORRUPT
            Case 3: baseBad.Status = VATS2_BASELINE_IO_ERROR
            Case 4: baseBad.Status = VATS2_BASELINE_UNSUPPORTED_VERSION
        End Select
        d = VATStage2CompareBaselineDelta(fp, baseBad)
        r = VATStage2EvaluateReuseEligibility(d, baseBad, seal)
        expected = VATS2_REUSE_BASELINE_UNAVAILABLE
        If i = 1 Then expected = VATS2_REUSE_FIRST_RUN
        Check "首次及旧baseline不可用" & i, r.Status = VATS2_REUSE_OK And r.Decision = VATS2_FULL_VALIDATION_REQUIRED And (r.ReasonFlags And expected) <> 0
    Next i
    '实际调用冻结指纹及delta层构造A/B变化，不手工伪造正常分类。
    For i = 1 To 6
        Call Fixture
        Select Case i
            Case 1: a.Range("D5").NumberFormat = "@": a.Range("D5").Value2 = INV
            Case 2: b.Range("B7").Value2 = SUPPLIER
            Case 3: a.Range("D4:I4").Clear
            Case 4: b.Range("B6:F6").Clear
            Case 5: a.Range("I4").Value2 = 123
            Case 6: b.Range("F6").Value2 = 123
        End Select
        changedFP = CurrentFP(): d = VATStage2CompareBaselineDelta(changedFP, base)
        r = VATStage2EvaluateReuseEligibility(d, base, seal)
        Check "任意业务变化必须全量" & i, d.Status = VATS2_DELTA_OK And r.Status = VATS2_REUSE_OK And r.Decision = VATS2_FULL_VALIDATION_REQUIRED And (r.ReasonFlags And VATS2_REUSE_BATCH_CHANGED) <> 0
        If i <= 2 Or i >= 5 Then Check "新增变化原因" & i, (r.ReasonFlags And VATS2_REUSE_DATA_NEW_OR_CHANGED) <> 0
        If i >= 3 Then Check "剩余变化原因" & i, (r.ReasonFlags And VATS2_REUSE_DATA_REMOVED_OR_CHANGED) <> 0
    Next i
    Call Fixture
    a.Range("D5:I5").Value2 = a.Range("D4:I4").Value2: a.Range("I5").Interior.Color = color
    b.Range("B7:F7").Value2 = b.Range("B6:F6").Value2: b.Range("F7").Interior.Color = color
    changedFP = CurrentFP(): d = VATStage2CompareBaselineDelta(changedFP, base)
    r = VATStage2EvaluateReuseEligibility(d, base, seal)
    Check "完全重复multiplicity增加不复用", r.Decision = VATS2_FULL_VALIDATION_REQUIRED And d.ANewOrChangedCount = 1 And d.BNewOrChangedCount = 1
    baseBad.Baseline = VATStage2BuildBaseline(changedFP): baseBad.Status = VATS2_BASELINE_OK
    d = VATStage2CompareBaselineDelta(fp, baseBad): r = VATStage2EvaluateReuseEligibility(d, baseBad, seal)
    Check "multiplicity减少不复用", r.Decision = VATS2_FULL_VALIDATION_REQUIRED And (r.ReasonFlags And VATS2_REUSE_DATA_REMOVED_OR_CHANGED) <> 0
    '行移动后长期身份不变。
    Call Fixture
    a.Range("D10").NumberFormat = "@": a.Range("D10").Value2 = INV
    a.Range("I10").Value2 = AMOUNT: a.Range("I10").Interior.Color = color: a.Range("D4:I4").Clear
    b.Range("B12").Value2 = SUPPLIER: b.Range("F12").Value2 = AMOUNT: b.Range("F12").Interior.Color = color: b.Range("B6:F6").Clear
    changedFP = CurrentFP(): d = VATStage2CompareBaselineDelta(changedFP, base)
    r = VATStage2EvaluateReuseEligibility(d, base, seal)
    Check "Excel纯移动可复用", changedFP.CombinedFingerprint = fp.CombinedFingerprint And r.Decision = VATS2_REUSE_ELIGIBLE
    For i = 1 To 10
        d = delta: baseBad = base: sealBad = seal
        Select Case i
            Case 1: d.ACurrent(1).Digest = String$(64, "a")
            Case 2: d.BCurrent(1).CurrentIndex = 2
            Case 3: d.ACurrent(1).State = VATS2_DELTA_NEW_OR_CHANGED
            Case 4: d.AUnchangedCount = -1
            Case 5: d.CombinedUnchanged = False
            Case 6: d.BaselineSourceStatus = VATS2_BASELINE_NOT_FOUND
            Case 7: baseBad.Baseline.ARecordDigests(1) = String$(64, "b")
            Case 8: sealBad.Seal.Status = VATS2_VERIFY_CORRUPT
            Case 9: d.Status = VATS2_DELTA_INVALID_INPUT
            Case 10: baseBad.Status = VATS2_BASELINE_INVALID_INPUT
        End Select
        r = VATStage2EvaluateReuseEligibility(d, baseBad, sealBad)
        Check "损坏必要契约不授予资格" & i, r.Status <> VATS2_REUSE_OK And r.Decision = VATS2_FULL_VALIDATION_REQUIRED
    Next i
    r = VATStage2EvaluateReuseEligibility(delta, base, seal)
    Check "全部正常输入保持", Stamp(seal.Seal) = before And delta.ACurrent(1).Digest = base.Baseline.ARecordDigests(1) And delta.BCurrent(1).CurrentIndex = 1 And r.Decision = VATS2_REUSE_ELIGIBLE
    Fixture False: runResult = Run(): fp = CurrentFP()
    base.Baseline = VATStage2BuildBaseline(fp): seal.Seal = VATStage2BuildVerificationSeal(runResult)
    delta = VATStage2CompareBaselineDelta(fp, base): r = VATStage2EvaluateReuseEligibility(delta, base, seal)
    Check "零记录三方一致可复用", r.Status = VATS2_REUSE_OK And r.Decision = VATS2_REUSE_ELIGIBLE
End Sub

Private Function Rechecksum(ByVal parts As Variant) As String
    Dim body As String, i As Long
    For i = 0 To 10: body = body & parts(i) & vbLf: Next i
    Rechecksum = body & VATStage2SHA256(body) & vbLf
End Function

Private Sub TestPurityAndOrder()
    Dim runResult As VATS2BoundAuditRunResult, seal As VATS2VerificationSealState, sl As VATS2VerificationSealLoadResult
    Dim baseline As VATS2BaselineLoadResult, fp As VATS2FingerprintResult, delta As VATS2IncrementalDeltaResult
    Dim r As VATS2ReuseEligibilityResult, beforeRun As String, beforeBase As String, beforeDelta As String, beforeSeal As String
    Dim temporary As String, originalFP As String, cm As Object, source As String, signature As String, p As Long
    Call Fixture
    a.Range("D5").NumberFormat = "@": a.Range("D5").Value2 = "22222222222287654321"
    a.Range("I5").Value2 = 123: a.Range("I5").Interior.Color = color
    b.Range("B7").Value2 = "NO.22222222222287654321": b.Range("F7").Value2 = 123: b.Range("F7").Interior.Color = color
    runResult = Run(): beforeRun = RunSignature(runResult)
    seal = VATStage2BuildVerificationSeal(runResult)
    Check "Build保持BoundRun全部字段", beforeRun = RunSignature(runResult)
    Check "两记录完整审核可Seal", seal.Status = VATS2_VERIFY_OK
    fp = CurrentFP(): originalFP = fp.CombinedFingerprint
    baseline.Baseline = VATStage2BuildBaseline(fp): sl.Seal = seal
    '交换业务记录内容并保持当前运行行号递增，模拟Excel排序。
    a.Range("D4").Value2 = "22222222222287654321": a.Range("I4").Value2 = 123
    a.Range("D5").Value2 = INV: a.Range("I5").Value2 = AMOUNT
    b.Range("B6").Value2 = "NO.22222222222287654321": b.Range("F6").Value2 = 123
    b.Range("B7").Value2 = SUPPLIER: b.Range("F7").Value2 = AMOUNT
    fp = CurrentFP(): delta = VATStage2CompareBaselineDelta(fp, baseline)
    beforeBase = OldSignature(baseline): beforeDelta = ResultSignature(delta)
    beforeSeal = Pack(CStr(sl.Status)) & Pack(Stamp(sl.Seal)) & Pack(sl.UnsupportedField) & Pack(sl.ErrorReason)
    r = VATStage2EvaluateReuseEligibility(delta, baseline, sl)
    Check "两侧排序后整批资格保持", originalFP = fp.CombinedFingerprint And r.Status = VATS2_REUSE_OK And r.Decision = VATS2_REUSE_ELIGIBLE
    Check "Reuse保持baseline全部字段", beforeBase = OldSignature(baseline)
    Check "Reuse保持delta全部字段", beforeDelta = ResultSignature(delta)
    Check "Reuse保持sealLoad全部字段", beforeSeal = Pack(CStr(sl.Status)) & Pack(Stamp(sl.Seal)) & Pack(sl.UnsupportedField) & Pack(sl.ErrorReason)
    b.Range("F7").Value2 = 456: runResult = Run(): beforeRun = RunSignature(runResult)
    seal = VATStage2BuildVerificationSeal(runResult)
    Check "拒绝时完整Finding成员及运行输入保持", beforeRun = RunSignature(runResult) And seal.Status = VATS2_VERIFY_INVALID_INPUT
    fp = CurrentFP(): delta = VATStage2CompareBaselineDelta(fp, baseline)
    beforeDelta = ResultSignature(delta): r = VATStage2EvaluateReuseEligibility(delta, baseline, sl)
    Check "变化及residual全部输入保持", beforeDelta = ResultSignature(delta) And r.Decision = VATS2_FULL_VALIDATION_REQUIRED
    sl.Status = VATS2_VERIFY_NOT_FOUND
    r = VATStage2EvaluateReuseEligibility(delta, baseline, sl)
    Check "多个ReasonFlags同时保留", (r.ReasonFlags And VATS2_REUSE_SEAL_NOT_FOUND) <> 0 And (r.ReasonFlags And VATS2_REUSE_DATA_NEW_OR_CHANGED) <> 0 And (r.ReasonFlags And VATS2_REUSE_DATA_REMOVED_OR_CHANGED) <> 0
    Set cm = ThisWorkbook.VBProject.VBComponents("modVATStage2VerificationSeal").CodeModule
    source = cm.Lines(1, cm.CountOfLines): p = InStr(source, "Public Function VATStage2BuildVerificationSeal")
    signature = Mid$(source, p, InStr(p, source, ") As VATS2VerificationSealState") - p)
    Check "唯一构建签名只接受BoundRun", InStr(1, signature, "ByRef run As VATS2BoundAuditRunResult", vbTextCompare) > 0 And InStr(signature, "FinalDecision") = 0 And InStr(signature, "Fingerprint") = 0
    Check "不自动联动baseline或审核", InStr(source, "VATStage2SaveBaseline") = 0 And InStr(source, "VATStage2RunBoundFullAudit(") = 0 And InStr(source, "DoEvents") = 0
End Sub

Private Function RunSignature(ByRef r As VATS2BoundAuditRunResult) As String
    Dim s As String, v As Variant, i As Long
    For Each v In Array(r.Status, r.BindingProtocol, r.FingerprintProtocol, r.DigestProtocol, r.ARecordCount, r.BRecordCount, r.ABatchDigest, r.BBatchDigest, r.CombinedDigest, r.ErrorStage, r.ErrorReason)
        s = s & Pack(CStr(v))
    Next v
    s = s & Pack(FinalSignature(r.FinalDecision)) & Pack(FindingsSignature(r.AuditFindings))
    With r.Aggregate
        For Each v In Array(.Status, .CalculationState, .AScanSucceeded, .BScanSucceeded, .TotalA, .TotalB, .Difference, .TotalsEqual, .ReliableEqualForShortSuffix, _
            .ACompletedCount, .BCompletedCount, .AIncludedCount, .BIncludedCount, .AEmptyWarningCount, .BEmptyWarningCount, .AIssueCount, .BIssueCount, .AWarningCount, .BWarningCount, .ErrorReason)
            s = s & Pack(CStr(VarType(v))) & Pack(CStr(v))
        Next v
        For i = 1 To .AIssueCount: s = s & Pack(.AIssues(i)): Next i
        For i = 1 To .BIssueCount: s = s & Pack(.BIssues(i)): Next i
        For i = 1 To .AWarningCount: s = s & Pack(.AWarnings(i)): Next i
        For i = 1 To .BWarningCount: s = s & Pack(.BWarnings(i)): Next i
    End With
    RunSignature = s
End Function
Private Sub WriteText(ByVal path As String, ByVal value As String)
    Dim stream As Object
    Set stream = fs.CreateTextFile(path, True, False): stream.Write value: stream.Close
End Sub
Private Function ReadText(ByVal path As String) As String
    Dim stream As Object
    Set stream = fs.OpenTextFile(path, 1, False, 0): ReadText = stream.ReadAll: stream.Close
End Function
Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        passed = passed + 1: log = log & "PASS " & name & vbCrLf
    Else
        failed = failed + 1: log = log & "FAIL " & name & vbCrLf
    End If
End Sub

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
Private Function OldSignature(ByRef value As VATS2BaselineLoadResult) As String
    Dim s As String, i As Long
    s = Pack(CStr(value.Status)) & Pack(value.ErrorReason)
    With value.Baseline
        s = s & Pack(CStr(.Status)) & Pack(.SchemaVersion) & Pack(.FingerprintProtocol) & Pack(.DigestProtocol) & Pack(CStr(.ARecordCount)) & Pack(CStr(.BRecordCount))
        s = s & Pack(.ABatchDigest) & Pack(.BBatchDigest) & Pack(.CombinedDigest) & Pack(.CreatedUtc) & Pack(.UpdatedUtc)
        For i = 1 To .ARecordCount: s = s & Pack(.ARecordDigests(i)): Next i
        For i = 1 To .BRecordCount: s = s & Pack(.BRecordDigests(i)): Next i
    End With
    OldSignature = s
End Function
Private Function ResultSignature(ByRef r As VATS2IncrementalDeltaResult) As String
    Dim s As String, item As Variant, i As Long
    For Each item In Array(r.Status, r.BaselineSourceStatus, r.ARecordCount, r.BRecordCount, r.AUnchangedCount, r.BUnchangedCount, _
        r.ANewOrChangedCount, r.BNewOrChangedCount, r.AResidualCount, r.BResidualCount, r.AResidualGroupCount, r.BResidualGroupCount, _
        r.ABatchUnchanged, r.BBatchUnchanged, r.CombinedUnchanged, r.ErrorSide, r.ErrorIndex, r.ErrorReason)
        s = s & Pack(CStr(item))
    Next item
    For i = 1 To r.ARecordCount: s = s & Pack(CStr(r.ACurrent(i).CurrentIndex)) & Pack(r.ACurrent(i).Digest) & Pack(CStr(r.ACurrent(i).State)): Next i
    For i = 1 To r.BRecordCount: s = s & Pack(CStr(r.BCurrent(i).CurrentIndex)) & Pack(r.BCurrent(i).Digest) & Pack(CStr(r.BCurrent(i).State)): Next i
    For i = 1 To r.AResidualGroupCount: s = s & Pack(r.AResiduals(i).Digest) & Pack(CStr(r.AResiduals(i).OccurrenceCount)): Next i
    For i = 1 To r.BResidualGroupCount: s = s & Pack(r.BResiduals(i).Digest) & Pack(CStr(r.BResiduals(i).OccurrenceCount)): Next i
    ResultSignature = s
End Function
