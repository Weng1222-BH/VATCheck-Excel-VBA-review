Attribute VB_Name = "modVATStage2CommitTests"
Option Explicit
Option Compare Binary

Private passed As Long, failed As Long, log As String
Private book As Workbook, a As Worksheet, b As Worksheet, color As Long
Private fs As Object, testRoot As String
Private Const INV_X As String = "11111111111112345678"
Private Const INV_Y As String = "22222222222287654321"
Private Const AMOUNT As Double = 987654.321

Public Function VATStage2VerifiedStateCommit_SelfTest() As String
    On Error GoTo Unexpected
    passed = 0: failed = 0: log = "": color = RGB(255, 255, 0)
    Set fs = CreateObject("Scripting.FileSystemObject")
    testRoot = fs.BuildPath(fs.GetSpecialFolder(2).Path, "VATCheck-C76-" & fs.GetTempName)
    fs.CreateFolder testRoot
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set a = book.Worksheets(1): Set b = book.Worksheets.Add(After:=a)
    TestNormal
    TestReject
    TestFailures
    TestPurity
    GoTo Done
Unexpected:
    failed = failed + 1: log = log & "UNEXPECTED " & Err.Number & " " & Err.Description & vbCrLf
Done:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    '测试仅清理自己创建且位于系统临时目录的唯一子目录。
    If Len(testRoot) > 0 Then
        If fs.GetParentFolderName(fs.GetAbsolutePathName(testRoot)) = fs.GetSpecialFolder(2).Path Then
            If Left$(fs.GetFileName(testRoot), 12) = "VATCheck-C76-" Then fs.DeleteFolder testRoot, True
        End If
    End If
    On Error GoTo 0
    If failed = 0 Then
        VATStage2VerifiedStateCommit_SelfTest = "PASS: " & passed & " assertions" & vbCrLf & log
    Else
        VATStage2VerifiedStateCommit_SelfTest = "FAIL: " & failed & "; PASS: " & passed & vbCrLf & log
    End If
End Function

Private Sub Fixture(Optional ByVal which As String = "X")
    a.Range("A1:P30").Clear: b.Range("A1:P30").Clear
    a.Range("D3").Value2 = "数电发票号码": a.Range("I3").Value2 = VAT_FIELD_A
    b.Range("B5").Value2 = "供应商名称": b.Range("F5").Value2 = VAT_FIELD_B
    If which = "EMPTY" Then Exit Sub
    a.Range("D4").NumberFormat = "@"
    If which = "X" Then a.Range("D4").Value2 = INV_X Else a.Range("D4").Value2 = INV_Y
    b.Range("B6").Value2 = "合成提交隐私样本NO." & a.Range("D4").Value2
    a.Range("I4").Value2 = AMOUNT: a.Range("I4").Interior.Color = color
    b.Range("F6").Value2 = AMOUNT: b.Range("F6").Interior.Color = color
End Sub
Private Function Run() As VATS2BoundAuditRunResult
    Run = VATStage2RunBoundFullAudit(a.Range("D3"), a.Range("I3"), b.Range("B5"), b.Range("F5"), color)
End Function
Private Function CurrentFP() As VATS2FingerprintResult
    Dim sa As VATS2ASnapshotResult, sb As VATS2BSnapshotResult
    sa = VATStage2ReadASnapshot(a, 3, 4, 9, color): sb = VATStage2ReadBSnapshot(b, 5, 2, 6, color)
    CurrentFP = VATStage2BuildFingerprints(sa, sb)
End Function
Private Function DirFor(ByVal name As String) As String
    DirFor = fs.BuildPath(testRoot, name)
End Function
Private Function DiskText(ByVal root As String, ByVal name As String) As String
    Dim stream As Object
    Set stream = fs.OpenTextFile(fs.BuildPath(root, name), 1, False, 0)
    DiskText = stream.ReadAll: stream.Close
End Function
Private Sub PutFile(ByVal root As String, ByVal name As String, ByVal text As String)
    Dim stream As Object
    If Not fs.FolderExists(root) Then fs.CreateFolder root
    Set stream = fs.CreateTextFile(fs.BuildPath(root, name), False, False)
    stream.Write text: stream.Close
End Sub
Private Function Eligibility(ByVal root As String, ByRef fp As VATS2FingerprintResult) As VATS2ReuseEligibilityResult
    Dim base As VATS2BaselineLoadResult, seal As VATS2VerificationSealLoadResult, delta As VATS2IncrementalDeltaResult
    base = VATStage2LoadBaseline(root): seal = VATStage2LoadVerificationSeal(root)
    delta = VATStage2CompareBaselineDelta(fp, base)
    Eligibility = VATStage2EvaluateReuseEligibility(delta, base, seal)
End Function
Private Sub CheckDisk(ByVal root As String, ByRef runResult As VATS2BoundAuditRunResult)
    Dim base As VATS2BaselineLoadResult, seal As VATS2VerificationSealLoadResult
    base = VATStage2LoadBaseline(root): seal = VATStage2LoadVerificationSeal(root)
    Check "两个文件可可靠读取", base.Status = VATS2_BASELINE_OK And seal.Status = VATS2_VERIFY_OK
    Check "baseline数量与BoundRun一致", base.Baseline.ARecordCount = runResult.ARecordCount And base.Baseline.BRecordCount = runResult.BRecordCount
    Check "seal数量与BoundRun一致", seal.Seal.ARecordCount = runResult.ARecordCount And seal.Seal.BRecordCount = runResult.BRecordCount
    Check "baseline三项绑定digest一致", base.Baseline.ABatchDigest = runResult.ABatchDigest And base.Baseline.BBatchDigest = runResult.BBatchDigest And base.Baseline.CombinedDigest = runResult.CombinedDigest
    Check "seal三项绑定digest一致", seal.Seal.ABatchDigest = runResult.ABatchDigest And seal.Seal.BBatchDigest = runResult.BBatchDigest And seal.Seal.CombinedDigest = runResult.CombinedDigest
End Sub
Private Sub CheckSuccess(ByRef r As VATS2VerifiedStateCommitResult)
    Check "提交成功", r.Status = VATS2_COMMIT_OK And r.BaselineWritten And r.SealWritten
    Check "保存状态完整传播", r.BaselineStatus = VATS2_BASELINE_OK And r.SealStatus = VATS2_VERIFY_OK
    Check "成功无错误定位", r.ErrorStage = "" And r.ErrorReason = ""
End Sub

Private Sub TestNormal()
    Dim runResult As VATS2BoundAuditRunResult, fp As VATS2FingerprintResult, bound As VATS2BoundBaselineResult
    Dim r As VATS2VerifiedStateCommitResult, reuse As VATS2ReuseEligibilityResult
    Dim root As String, oldDigest As String, text As String, name As Variant, canary As Variant
    Call Fixture: runResult = Run(): fp = CurrentFP(): root = DirFor("normal")
    bound = VATStage2BuildBaselineForBoundRun(runResult, fp)
    Check "绑定辅助构建成功", bound.Status = VATS2_COMMIT_OK And bound.Baseline.CombinedDigest = runResult.CombinedDigest
    Check "绑定辅助不创建目录", Not fs.FolderExists(root)
    r = VATStage2CommitVerifiedState(runResult, fp, root): CheckSuccess r: CheckDisk root, runResult
    reuse = Eligibility(root, fp)
    Check "完整提交后具备C7.5资格", reuse.Status = VATS2_REUSE_OK And reuse.Decision = VATS2_REUSE_ELIGIBLE
    Check "只产生两个正式状态文件", fs.GetFolder(root).Files.Count = 2
    For Each name In Array("baseline-v1.dat", "verified-v1.dat")
        text = DiskText(root, CStr(name))
        For Each canary In Array(INV_X, "合成提交隐私样本", CStr(AMOUNT), fp.AFingerprints(1), fp.BFingerprints(1), fp.CombinedFingerprint, book.FullName)
            Check "状态文件不含业务原文", InStr(1, text, CStr(canary), vbBinaryCompare) = 0
        Next canary
    Next name
    oldDigest = runResult.CombinedDigest
    Fixture "Y": runResult = Run(): fp = CurrentFP()
    r = VATStage2CommitVerifiedState(runResult, fp, root): CheckSuccess r: CheckDisk root, runResult
    Check "第二次覆盖为新批次", runResult.CombinedDigest <> oldDigest
    reuse = Eligibility(root, fp): Check "覆盖后仍三方一致", reuse.Decision = VATS2_REUSE_ELIGIBLE
    r = VATStage2CommitVerifiedState(runResult, fp, root): CheckSuccess r: CheckDisk root, runResult
    Check "重复提交不产生第三文件", fs.GetFolder(root).Files.Count = 2
    Fixture "EMPTY": runResult = Run(): fp = CurrentFP()
    r = VATStage2CommitVerifiedState(runResult, fp, DirFor("empty")): CheckSuccess r: CheckDisk DirFor("empty"), runResult
    Check "零记录VERIFIED", runResult.FinalDecision.NoCompletedRecords And runResult.ARecordCount = 0 And runResult.BRecordCount = 0
End Sub

Private Sub TestReject()
    Dim original As VATS2BoundAuditRunResult, candidate As VATS2BoundAuditRunResult
    Dim fp As VATS2FingerprintResult, badFP As VATS2FingerprintResult, bound As VATS2BoundBaselineResult
    Dim r As VATS2VerifiedStateCommitResult, i As Long, expected As VATS2VerifiedStateCommitStatus
    Dim root As String, baselineBefore As String, sealBefore As String
    Call Fixture: original = Run(): fp = CurrentFP(): root = DirFor("reject")
    r = VATStage2CommitVerifiedState(original, fp, root): CheckSuccess r
    baselineBefore = DiskText(root, "baseline-v1.dat"): sealBefore = DiskText(root, "verified-v1.dat")
    For i = 1 To 17
        candidate = original: badFP = fp: expected = VATS2_COMMIT_INVALID_CONTRACT
        Select Case i
            Case 1: candidate.FinalDecision.Verdict = VATS2_FINAL_REVIEW_REQUIRED: expected = VATS2_COMMIT_INVALID_INPUT
            Case 2: candidate.FinalDecision.Verdict = VATS2_FINAL_UNAVAILABLE: expected = VATS2_COMMIT_INVALID_INPUT
            Case 3: candidate.Status = VATS2_BOUND_RUN_SOURCE_CHANGED: expected = VATS2_COMMIT_INVALID_INPUT
            Case 4: candidate.FinalDecision.Status = VATS2_FINAL_INVALID_INPUT: expected = VATS2_COMMIT_INVALID_INPUT
            Case 5: badFP.Status = VATS2_FINGERPRINT_INVALID_INPUT: expected = VATS2_COMMIT_INVALID_INPUT
            Case 6: candidate.ARecordCount = 2
            Case 7: candidate.BRecordCount = 2
            Case 8: candidate.ABatchDigest = String$(64, "a")
            Case 9: candidate.BBatchDigest = String$(64, "b")
            Case 10: candidate.CombinedDigest = String$(64, "c")
            Case 11: badFP.ABatchFingerprint = "unrelated batch"
            Case 12: candidate.BindingProtocol = "VATRUN2"
            Case 13: candidate.Aggregate.AWarningCount = 1
            Case 14: candidate.FinalDecision.HasAggregateMismatch = True
            Case 15: candidate.AuditFindings.FindingCount = 1
            Case 16: Erase badFP.AFingerprints
            Case 17: badFP.AFingerprints(1) = ChrW(&HD800)
        End Select
        r = VATStage2CommitVerifiedState(candidate, badFP, root)
        Check "拒绝非法输入或拼接" & i, r.Status = expected
        Check "拒绝时两个Save均未执行" & i, Not r.BaselineWritten And Not r.SealWritten And r.BaselineStatus = VATS2_COMMIT_NOT_ATTEMPTED And r.SealStatus = VATS2_COMMIT_NOT_ATTEMPTED
        Check "原磁盘内容不改" & i, DiskText(root, "baseline-v1.dat") = baselineBefore And DiskText(root, "verified-v1.dat") = sealBefore And fs.GetFolder(root).Files.Count = 2
        Check "错误定位不为空" & i, Len(r.ErrorStage) > 0 And Len(r.ErrorReason) > 0
    Next i
    Fixture "Y": badFP = CurrentFP()
    r = VATStage2CommitVerifiedState(original, badFP, DirFor("cross-batch"))
    Check "真实X运行加Y指纹被拒绝", r.Status = VATS2_COMMIT_INVALID_CONTRACT And r.ErrorStage = "BASELINE_BINDING"
    Check "跨批拒绝不创建任何目录文件", Not fs.FolderExists(DirFor("cross-batch"))
    bound = VATStage2BuildBaselineForBoundRun(original, badFP)
    Check "辅助失败无部分baseline", bound.Status = VATS2_COMMIT_INVALID_CONTRACT And bound.Baseline.Status <> VATS2_BASELINE_OK And bound.Baseline.CombinedDigest = "" And bound.Baseline.ARecordCount = 0
    Call Fixture: b.Range("F6").Value2 = 123: candidate = Run(): badFP = CurrentFP()
    r = VATStage2CommitVerifiedState(candidate, badFP, DirFor("real-review"))
    Check "真实REVIEW不落盘", candidate.FinalDecision.Verdict = VATS2_FINAL_REVIEW_REQUIRED And r.Status = VATS2_COMMIT_INVALID_INPUT And Not fs.FolderExists(DirFor("real-review"))
    a.Range("I4").Formula = "=""""": a.Range("I4").Calculate: candidate = Run(): badFP = CurrentFP()
    r = VATStage2CommitVerifiedState(candidate, badFP, DirFor("real-unavailable"))
    Check "真实UNAVAILABLE不落盘", candidate.FinalDecision.Verdict = VATS2_FINAL_UNAVAILABLE And r.Status = VATS2_COMMIT_INVALID_INPUT And Not fs.FolderExists(DirFor("real-unavailable"))
End Sub

Private Sub TestFailures()
    Dim runX As VATS2BoundAuditRunResult, runY As VATS2BoundAuditRunResult, fpX As VATS2FingerprintResult, fpY As VATS2FingerprintResult
    Dim r As VATS2VerifiedStateCommitResult, reuse As VATS2ReuseEligibilityResult, loaded As VATS2BaselineLoadResult
    Dim root As String, baseBefore As String, sealBefore As String, f As Integer
    Call Fixture: runX = Run(): fpX = CurrentFP()
    Fixture "Y": runY = Run(): fpY = CurrentFP()
    root = DirFor("baseline-pending")
    r = VATStage2CommitVerifiedState(runX, fpX, root): CheckSuccess r
    baseBefore = DiskText(root, "baseline-v1.dat"): sealBefore = DiskText(root, "verified-v1.dat")
    PutFile root, "baseline-v1.pending.tmp", "BASELINE_OTHER_WRITER"
    r = VATStage2CommitVerifiedState(runY, fpY, root)
    Check "baseline失败阻止后续Save", r.Status = VATS2_COMMIT_BASELINE_FAILED And r.BaselineStatus = VATS2_BASELINE_IO_ERROR And r.SealStatus = VATS2_COMMIT_NOT_ATTEMPTED
    Check "baseline失败Written均False", Not r.BaselineWritten And Not r.SealWritten And r.ErrorStage = "BASELINE_SAVE"
    Check "两份旧状态及他人pending保持", DiskText(root, "baseline-v1.dat") = baseBefore And DiskText(root, "verified-v1.dat") = sealBefore And DiskText(root, "baseline-v1.pending.tmp") = "BASELINE_OTHER_WRITER"
    root = DirFor("different-batch")
    r = VATStage2CommitVerifiedState(runX, fpX, root): CheckSuccess r
    sealBefore = DiskText(root, "verified-v1.dat")
    PutFile root, "verified-v1.pending.tmp", "SEAL_OTHER_WRITER"
    r = VATStage2CommitVerifiedState(runY, fpY, root): CheckPartial r
    loaded = VATStage2LoadBaseline(root)
    Check "Partial保留新Y baseline不回滚", loaded.Status = VATS2_BASELINE_OK And loaded.Baseline.CombinedDigest = runY.CombinedDigest
    Check "不同批旧X seal不删除不篡改", DiskText(root, "verified-v1.dat") = sealBefore
    reuse = Eligibility(root, fpY)
    Check "不同批Partial强制全量", reuse.Status = VATS2_REUSE_OK And reuse.Decision = VATS2_FULL_VALIDATION_REQUIRED And (reuse.ReasonFlags And VATS2_REUSE_SEAL_BASELINE_MISMATCH) <> 0
    Check "seal他人pending保留", DiskText(root, "verified-v1.pending.tmp") = "SEAL_OTHER_WRITER" And fs.GetFolder(root).Files.Count = 3
    root = DirFor("same-batch")
    r = VATStage2CommitVerifiedState(runX, fpX, root): CheckSuccess r
    sealBefore = DiskText(root, "verified-v1.dat")
    PutFile root, "verified-v1.pending.tmp", "SAME_BATCH_PENDING"
    r = VATStage2CommitVerifiedState(runX, fpX, root): CheckPartial r
    reuse = Eligibility(root, fpX)
    Check "同批例外仍由C7.5授予资格", reuse.Status = VATS2_REUSE_OK And reuse.Decision = VATS2_REUSE_ELIGIBLE And reuse.ReasonFlags = 0
    Check "同批旧Seal和pending保持", DiskText(root, "verified-v1.dat") = sealBefore And DiskText(root, "verified-v1.pending.tmp") = "SAME_BATCH_PENDING"
    root = DirFor("missing-seal")
    PutFile root, "verified-v1.pending.tmp", "PENDING_WITHOUT_SEAL"
    r = VATStage2CommitVerifiedState(runY, fpY, root): CheckPartial r
    reuse = Eligibility(root, fpY)
    Check "缺seal的Partial必须全量", reuse.Status = VATS2_REUSE_OK And reuse.Decision = VATS2_FULL_VALIDATION_REQUIRED And (reuse.ReasonFlags And VATS2_REUSE_SEAL_NOT_FOUND) <> 0
    Check "不伪造或补写缺失seal", Not fs.FileExists(fs.BuildPath(root, "verified-v1.dat")) And fs.GetFolder(root).Files.Count = 2
    root = DirFor("locked-seal")
    r = VATStage2CommitVerifiedState(runX, fpX, root): CheckSuccess r
    sealBefore = DiskText(root, "verified-v1.dat")
    f = FreeFile: Open fs.BuildPath(root, "verified-v1.dat") For Binary Access Read Lock Read Write As #f
    r = VATStage2CommitVerifiedState(runY, fpY, root)
    Close #f
    CheckPartial r
    Check "替换失败清理自己的pending且旧seal保持", DiskText(root, "verified-v1.dat") = sealBefore And fs.GetFolder(root).Files.Count = 2
    reuse = Eligibility(root, fpY): Check "锁失败Partial安全降级", reuse.Decision = VATS2_FULL_VALIDATION_REQUIRED
    '正常重试显式调用协调器，无后台恢复、回滚或补写。
    r = VATStage2CommitVerifiedState(runY, fpY, root): CheckSuccess r: CheckDisk root, runY
    reuse = Eligibility(root, fpY): Check "解除锁后显式重试成功", reuse.Decision = VATS2_REUSE_ELIGIBLE
End Sub
Private Sub CheckPartial(ByRef r As VATS2VerifiedStateCommitResult)
    Check "返回精确Partial状态", r.Status = VATS2_COMMIT_SEAL_FAILED_PARTIAL
    Check "Partial仅baseline写成功", r.BaselineWritten And Not r.SealWritten And r.BaselineStatus = VATS2_BASELINE_OK And r.SealStatus = VATS2_VERIFY_IO_ERROR
    Check "Partial定位seal保存阶段", r.ErrorStage = "SEAL_SAVE" And Len(r.ErrorReason) > 0
End Sub

Private Sub TestPurity()
    Dim runResult As VATS2BoundAuditRunResult, fp As VATS2FingerprintResult, r As VATS2VerifiedStateCommitResult
    Dim beforeRun As String, beforeFP As String, beforeCells As String, saved As Boolean, count As Long
    Dim cm As Object, source As String, signature As String, p As Long, v As Variant
    Call Fixture
    a.Range("I4").Formula = "=987654.321": a.Range("I4").Calculate
    b.Range("F6").Formula = "=987654.321": b.Range("F6").Calculate
    runResult = Run(): fp = CurrentFP()
    beforeRun = RunSignature(runResult): beforeFP = FpSignature(fp): beforeCells = Evidence(a) & Evidence(b)
    saved = book.Saved: count = Application.Workbooks.Count
    r = VATStage2CommitVerifiedState(runResult, fp, DirFor("purity"))
    Check "Run所有字段完全不改", beforeRun = RunSignature(runResult)
    Check "FP所有字段完全不改", beforeFP = FpSignature(fp)
    Check "Workbook值公式颜色完全不改", beforeCells = Evidence(a) & Evidence(b)
    Check "Workbook未保存未关闭", book.Saved = saved And book.Path = "" And Application.Workbooks.Count = count
    book.Saved = True
    r = VATStage2CommitVerifiedState(runResult, fp, DirFor("clean-book"))
    Check "原干净Workbook保持干净", book.Saved
    book.Saved = False
    Set cm = ThisWorkbook.VBProject.VBComponents("modVATStage2VerifiedStateCommit").CodeModule
    source = cm.Lines(1, cm.CountOfLines)
    For Each v In Array("CreateObject", "Open ", "Kill ", "MoveFile", "SaveAs", "DoEvents", "VATStage2EvaluateReuseEligibility", "VATStage2LoadBaseline", "VATStage2LoadVerificationSeal")
        Check "Coordinator不直接I/O或判定reuse", InStr(1, source, CStr(v), vbTextCompare) = 0
    Next v
    p = InStr(source, "Public Function VATStage2CommitVerifiedState")
    signature = Mid$(source, p, InStr(p, source, ") As VATS2VerifiedStateCommitResult") - p)
    Check "入口不接受裸FinalDecision", InStr(signature, "VATS2BoundAuditRunResult") > 0 And InStr(signature, "VATS2FingerprintResult") > 0 And InStr(signature, "VATS2FinalDecisionResult") = 0
End Sub

Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        passed = passed + 1: log = log & "PASS " & name & vbCrLf
    Else
        failed = failed + 1: log = log & "FAIL " & name & vbCrLf
    End If
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
Private Function FpSignature(ByRef fp As VATS2FingerprintResult) As String
    Dim s As String, i As Long
    s = Pack(CStr(fp.Status)) & Pack(CStr(fp.ARecordCount)) & Pack(CStr(fp.BRecordCount)) & Pack(fp.ABatchFingerprint) & Pack(fp.BBatchFingerprint) & Pack(fp.CombinedFingerprint)
    s = s & Pack(fp.ErrorSide) & Pack(CStr(fp.ErrorIndex)) & Pack(fp.ErrorReason)
    For i = 1 To fp.ARecordCount: s = s & Pack(fp.AFingerprints(i)): Next i
    For i = 1 To fp.BRecordCount: s = s & Pack(fp.BFingerprints(i)): Next i
    FpSignature = s
End Function
Private Function Evidence(ByVal sheet As Worksheet) As String
    Dim cell As Range, s As String
    For Each cell In sheet.Range("A1:P15")
        s = s & Pack(CStr(VarType(cell.Value2))) & Pack(CStr(cell.Value2)) & Pack(CStr(cell.Formula)) & Pack(CStr(cell.HasFormula)) & Pack(CStr(cell.Interior.Pattern)) & Pack(CStr(cell.Interior.Color))
    Next cell
    Evidence = s
End Function
