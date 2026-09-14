Attribute VB_Name = "modVATStage2DeltaTests"
Option Explicit
Option Compare Binary

Private passed As Long, failed As Long, log As String
Private a As VATS2ASnapshotResult, b As VATS2BSnapshotResult
Private old As VATS2BaselineLoadResult

Public Function VATStage2IncrementalDelta_SelfTest() As String
    On Error GoTo Unexpected
    passed = 0: failed = 0: log = ""
    TestChanges
    TestDuplicates
    TestModes
    TestContracts
    TestPurity
    GoTo Done
Unexpected:
    failed = failed + 1: log = log & "UNEXPECTED " & Err.Number & " " & Err.Description & vbCrLf
Done:
    If failed = 0 Then
        VATStage2IncrementalDelta_SelfTest = "PASS: " & passed & " assertions" & vbCrLf & log
    Else
        VATStage2IncrementalDelta_SelfTest = "FAIL: " & failed & "; PASS: " & passed & vbCrLf & log
    End If
End Function
Private Sub Fixture(Optional ByVal ac As Long = 2, Optional ByVal bc As Long = 2)
    Dim aa As VATS2ASnapshotResult, bb As VATS2BSnapshotResult, i As Long
    a = aa: b = bb: a.RecordCount = ac: b.RecordCount = bc
    If ac > 0 Then ReDim a.Records(1 To ac)
    If bc > 0 Then ReDim b.Records(1 To bc)
    For i = 1 To ac
        With a.Records(i)
            .AIndex = i: .ExcelRow = 10 + i: .InvoiceDigitsRaw = "0000" & i
            .AmountRaw = CDec(i): .IsCompleted = True: .InvoiceCellAddress = "I" & .ExcelRow: .AmountCellAddress = "J" & .ExcelRow
        End With
    Next i
    For i = 1 To bc
        With b.Records(i)
            .BIndex = i: .ExcelRow = 20 + i: .SupplierTextRaw = "合成记录NO.0000" & i
            .AmountRaw = CDec(i): .IsCompleted = True: .SupplierCellAddress = "B" & .ExcelRow: .AmountCellAddress = "F" & .ExcelRow
        End With
    Next i
End Sub
Private Function Fingerprints() As VATS2FingerprintResult
    Fingerprints = VATStage2BuildFingerprints(a, b)
End Function
Private Sub Remember()
    Dim fp As VATS2FingerprintResult, blank As VATS2BaselineLoadResult
    old = blank: fp = Fingerprints(): old.Baseline = VATStage2BuildBaseline(fp)
End Sub
Private Function RunDelta() As VATS2IncrementalDeltaResult
    Dim fp As VATS2FingerprintResult
    fp = Fingerprints(): RunDelta = VATStage2CompareBaselineDelta(fp, old)
End Function
Private Sub CheckCounts(ByRef r As VATS2IncrementalDeltaResult, ByVal au As Long, ByVal an As Long, ByVal ar As Long, _
    ByVal bu As Long, ByVal bn As Long, ByVal br As Long)
    Check "正常比较Status", r.Status = VATS2_DELTA_OK
    Check "A分类计数", r.AUnchangedCount = au And r.ANewOrChangedCount = an And r.AResidualCount = ar And r.ARecordCount = au + an
    Check "B分类计数", r.BUnchangedCount = bu And r.BNewOrChangedCount = bn And r.BResidualCount = br And r.BRecordCount = bu + bn
    Check "A批次契约", r.ABatchUnchanged = (an = 0 And ar = 0)
    Check "B批次契约", r.BBatchUnchanged = (bn = 0 And br = 0)
    Check "Combined契约", r.CombinedUnchanged = (an = 0 And ar = 0 And bn = 0 And br = 0)
End Sub
Private Sub TestChanges()
    Dim r As VATS2IncrementalDeltaResult, fp As VATS2FingerprintResult, ar As VATS2ASnapshotRecord, br As VATS2BSnapshotRecord, i As Long
    Call Fixture: Call Remember: r = RunDelta(): CheckCounts r, 2, 0, 0, 2, 0, 0
    Check "无剩余时数组未分配", NoResiduals(r)
    ar = a.Records(1): a.Records(1) = a.Records(2): a.Records(2) = ar
    br = b.Records(1): b.Records(1) = b.Records(2): b.Records(2) = br
    For i = 1 To 2
        a.Records(i).AIndex = i: a.Records(i).ExcelRow = 100 + i: a.Records(i).InvoiceCellAddress = "Z" & (100 + i)
        b.Records(i).BIndex = i: b.Records(i).ExcelRow = 200 + i: b.Records(i).SupplierCellAddress = "T" & (200 + i)
    Next i
    r = RunDelta(): fp = Fingerprints(): CheckCounts r, 2, 0, 0, 2, 0, 0
    For i = 1 To 2
        Check "换序后A当前映射" & i, r.ACurrent(i).CurrentIndex = i And r.ACurrent(i).Digest = VATStage2SHA256(fp.AFingerprints(i)) And r.ACurrent(i).State = VATS2_DELTA_UNCHANGED
        Check "换序后B当前映射" & i, r.BCurrent(i).CurrentIndex = i And r.BCurrent(i).Digest = VATStage2SHA256(fp.BFingerprints(i)) And r.BCurrent(i).State = VATS2_DELTA_UNCHANGED
    Next i
    For i = 1 To 10
        Call Fixture: Call Remember
        Select Case i
            Case 1: a.Records(1).AmountRaw = CDec(7)
            Case 2: b.Records(1).SupplierTextRaw = "合成变更记录"
            Case 3: a.Records(1).IsCompleted = False
            Case 4: b.Records(1).IsCompleted = False
            Case 5: a.Records(1).InvoiceHasFormula = True
            Case 6: b.Records(1).SupplierHasFormula = True
            Case 7: a.Records(1).AmountHasFormula = True
            Case 8: b.Records(1).AmountHasFormula = True
            Case 9: a.Records(1).InvoiceDigitsRaw = "000099"
            Case 10: b.Records(1).AmountRaw = "2"
        End Select
        r = RunDelta()
        If i Mod 2 = 1 Then
            CheckCounts r, 1, 1, 1, 2, 0, 0
            Check "变更A分类" & i, r.ACurrent(1).State = VATS2_DELTA_NEW_OR_CHANGED And r.ACurrent(2).State = VATS2_DELTA_UNCHANGED
            Check "A剩余次数" & i, r.AResidualGroupCount = 1 And r.AResiduals(1).OccurrenceCount = 1
        Else
            CheckCounts r, 2, 0, 0, 1, 1, 1
            Check "变更B分类" & i, r.BCurrent(1).State = VATS2_DELTA_NEW_OR_CHANGED And r.BCurrent(2).State = VATS2_DELTA_UNCHANGED
            Check "B剩余次数" & i, r.BResidualGroupCount = 1 And r.BResiduals(1).OccurrenceCount = 1
        End If
    Next i
    Call Fixture: Call Remember: Fixture 3, 2: r = RunDelta(): CheckCounts r, 2, 1, 0, 2, 0, 0
    Check "A新增当前位置", r.ACurrent(3).CurrentIndex = 3 And r.ACurrent(3).State = VATS2_DELTA_NEW_OR_CHANGED
    Fixture 1, 2: r = RunDelta(): CheckCounts r, 1, 0, 1, 2, 0, 0
    Fixture 2, 3: r = RunDelta(): CheckCounts r, 2, 0, 0, 2, 1, 0
    Fixture 2, 1: r = RunDelta(): CheckCounts r, 2, 0, 0, 1, 0, 1
    Fixture 0, 0: Call Remember: r = RunDelta(): CheckCounts r, 0, 0, 0, 0, 0, 0
    Check "空双方未分配记录数组", NoCurrent(r)
    Fixture 0, 2: Call Remember: r = RunDelta(): CheckCounts r, 0, 0, 0, 2, 0, 0
    Fixture 2, 0: Call Remember: r = RunDelta(): CheckCounts r, 2, 0, 0, 0, 0, 0
    Fixture 0, 0: r = RunDelta(): CheckCounts r, 0, 0, 2, 0, 0, 0
    Check "剩余按digest排序", r.AResidualGroupCount = 2 And StrComp(r.AResiduals(1).Digest, r.AResiduals(2).Digest, vbBinaryCompare) < 0
    Fixture 0, 0: Call Remember: Call Fixture: r = RunDelta(): CheckCounts r, 0, 2, 0, 0, 2, 0
    Call Remember: Fixture 0, 0: r = RunDelta(): CheckCounts r, 0, 0, 2, 0, 0, 2
End Sub
Private Sub DuplicateFixture(ByVal count As Long)
    Dim i As Long
    Fixture count, count
    For i = 1 To count
        a.Records(i).InvoiceDigitsRaw = "00001": a.Records(i).AmountRaw = CDec(1)
        b.Records(i).SupplierTextRaw = "合成重复NO.00001": b.Records(i).AmountRaw = CDec(1)
    Next i
End Sub
Private Sub TestDuplicates()
    Dim r As VATS2IncrementalDeltaResult, fp As VATS2FingerprintResult, i As Long
    DuplicateFixture 2: Call Remember: DuplicateFixture 3: r = RunDelta(): CheckCounts r, 2, 1, 0, 2, 1, 0
    For i = 1 To 3
        Check "按当前顺序消费A出现次数" & i, (r.ACurrent(i).State = VATS2_DELTA_UNCHANGED) = (i <= 2)
        Check "按当前顺序消费B出现次数" & i, (r.BCurrent(i).State = VATS2_DELTA_UNCHANGED) = (i <= 2)
    Next i
    DuplicateFixture 3: Call Remember: DuplicateFixture 2: r = RunDelta(): CheckCounts r, 2, 0, 1, 2, 0, 1
    Check "重复减少只剩一次", r.AResidualGroupCount = 1 And r.AResiduals(1).OccurrenceCount = 1 And r.BResiduals(1).OccurrenceCount = 1
    DuplicateFixture 0: r = RunDelta(): CheckCounts r, 0, 0, 3, 0, 0, 3
    Check "剩余聚合保留三次", r.AResidualGroupCount = 1 And r.AResiduals(1).OccurrenceCount = 3 And r.BResidualGroupCount = 1 And r.BResiduals(1).OccurrenceCount = 3
    'X/Y/X换序为Y/X/X，同摘要的预算不依赖旧索引。
    DuplicateFixture 3: a.Records(2).InvoiceDigitsRaw = "other": Call Remember
    a.Records(1).InvoiceDigitsRaw = "other": a.Records(2).InvoiceDigitsRaw = "00001": r = RunDelta(): CheckCounts r, 3, 0, 0, 3, 0, 0
    '人为构造合法接口中的相同digest文本，证明没有跨侧消费；不解析VATFP1内部结构。
    Fixture 0, 1: fp = Fingerprints(): fp.BFingerprints(1) = "synthetic-shared-canonical"
    old.Baseline = VATStage2BuildBaseline(fp)
    Fixture 1, 0: fp = Fingerprints(): fp.AFingerprints(1) = "synthetic-shared-canonical"
    r = VATStage2CompareBaselineDelta(fp, old): CheckCounts r, 0, 1, 0, 0, 0, 1
    Check "相同digest不能跨侧匹配", r.ACurrent(1).Digest = r.BResiduals(1).Digest And r.ACurrent(1).State = VATS2_DELTA_NEW_OR_CHANGED
    DuplicateFixture 3: a.Records(3).InvoiceDigitsRaw = "Y": b.Records(3).SupplierTextRaw = "Y": Call Remember
    DuplicateFixture 4: a.Records(4).InvoiceDigitsRaw = "Y": b.Records(4).SupplierTextRaw = "Y": r = RunDelta()
    CheckCounts r, 3, 1, 0, 3, 1, 0
    Check "X预算耗尽不影响Y", r.ACurrent(3).State = VATS2_DELTA_NEW_OR_CHANGED And r.ACurrent(4).State = VATS2_DELTA_UNCHANGED And r.BCurrent(3).State = VATS2_DELTA_NEW_OR_CHANGED And r.BCurrent(4).State = VATS2_DELTA_UNCHANGED
End Sub
Private Sub TestModes()
    Dim fp As VATS2FingerprintResult, r As VATS2IncrementalDeltaResult, i As Long
    Call Fixture: Call Remember: fp = Fingerprints(): old.Status = VATS2_BASELINE_NOT_FOUND
    r = VATStage2CompareBaselineDelta(fp, old)
    Check "FIRST_RUN独立模式", r.Status = VATS2_DELTA_FIRST_RUN And r.BaselineSourceStatus = VATS2_BASELINE_NOT_FOUND
    Check "FIRST_RUN全部待完整验证", r.ANewOrChangedCount = 2 And r.BNewOrChangedCount = 2 And r.AUnchangedCount = 0 And r.BUnchangedCount = 0
    Check "FIRST_RUN无旧残留无批次证明", r.AResidualCount = 0 And r.BResidualCount = 0 And Not r.ABatchUnchanged And Not r.BBatchUnchanged And Not r.CombinedUnchanged
    Check "FIRST_RUN不输出异常原因", r.ErrorReason = "" And NoResiduals(r)
    For i = 1 To 2: Check "FIRST_RUN记录状态" & i, r.ACurrent(i).State = VATS2_DELTA_NEW_OR_CHANGED And r.BCurrent(i).State = VATS2_DELTA_NEW_OR_CHANGED: Next i
    For i = VATS2_BASELINE_CORRUPT To VATS2_BASELINE_UNSUPPORTED_VERSION
        old.Status = i: old.ErrorReason = "synthetic source error " & i
        r = VATStage2CompareBaselineDelta(fp, old)
        Check "不可用来源" & i, r.Status = VATS2_DELTA_BASELINE_UNAVAILABLE And r.BaselineSourceStatus = i And r.ErrorReason = old.ErrorReason
        Check "不可用不分类" & i, NoCurrent(r) And NoResiduals(r) And r.ARecordCount = 0 And r.BRecordCount = 0 And Not r.CombinedUnchanged
    Next i
    old.Status = VATS2_BASELINE_INVALID_INPUT: r = VATStage2CompareBaselineDelta(fp, old)
    Check "来源INVALID_INPUT拒绝", r.Status = VATS2_DELTA_INVALID_INPUT And NoCurrent(r)
    Fixture 0, 0: fp = Fingerprints(): old.Status = VATS2_BASELINE_NOT_FOUND
    r = VATStage2CompareBaselineDelta(fp, old)
    Check "零记录FIRST_RUN无伪造相同", r.Status = VATS2_DELTA_FIRST_RUN And NoCurrent(r) And Not r.CombinedUnchanged
End Sub
Private Sub TestContracts()
    Dim fp As VATS2FingerprintResult, original As VATS2FingerprintResult, source As VATS2BaselineLoadResult
    Dim r As VATS2IncrementalDeltaResult, i As Long, temp As String
    Call Fixture: Call Remember: original = Fingerprints(): source = old
    For i = 1 To 28
        fp = original: old = source
        Select Case i
            Case 1: fp.Status = VATS2_FINGERPRINT_INVALID_INPUT
            Case 2: fp.ARecordCount = -1
            Case 3: fp.BRecordCount = 3
            Case 4: Erase fp.AFingerprints
            Case 5: ReDim fp.AFingerprints(0 To 1)
            Case 6: fp.BFingerprints(1) = ""
            Case 7: fp.ABatchFingerprint = ""
            Case 8: fp.BBatchFingerprint = ""
            Case 9: fp.CombinedFingerprint = ""
            Case 10: old.Baseline.Status = VATS2_BASELINE_CORRUPT
            Case 11: old.Baseline.SchemaVersion = "VATBASE2"
            Case 12: old.Baseline.FingerprintProtocol = "VATFP2"
            Case 13: old.Baseline.DigestProtocol = "other"
            Case 14: old.Baseline.ARecordCount = -1
            Case 15: old.Baseline.BRecordCount = 3
            Case 16: Erase old.Baseline.ARecordDigests
            Case 17: ReDim old.Baseline.BRecordDigests(0 To 1)
            Case 18: old.Baseline.ARecordDigests(1) = String$(63, "a")
            Case 19: old.Baseline.BRecordDigests(1) = String$(64, "g")
            Case 20: old.Baseline.ARecordDigests(1) = String$(64, "A")
            Case 21
                temp = old.Baseline.ARecordDigests(1): old.Baseline.ARecordDigests(1) = old.Baseline.ARecordDigests(2): old.Baseline.ARecordDigests(2) = temp
            Case 22: old.Baseline.CombinedDigest = ""
            Case 23: old.Baseline.ABatchDigest = String$(63, "a")
            Case 24: old.Baseline.BBatchDigest = String$(64, "z")
            Case 25: fp.ARecordCount = 0
            Case 26: old.Baseline.BRecordCount = 0
            Case 27: fp.AFingerprints(1) = ChrW$(&HD800)
            Case 28: old.Status = 99
        End Select
        r = VATStage2CompareBaselineDelta(fp, old)
        If i = 1 Or i = 28 Then
            Check "主状态拒绝" & i, r.Status = VATS2_DELTA_INVALID_INPUT
        Else
            Check "必要契约拒绝" & i, r.Status = VATS2_DELTA_INVALID_CONTRACT
        End If
        Check "失败清空部分数组和分类" & i, NoCurrent(r) And NoResiduals(r) And r.ARecordCount = 0 And r.BRecordCount = 0 And r.AUnchangedCount = 0 And Not r.CombinedUnchanged And Len(r.ErrorReason) > 0
    Next i
    For i = 1 To 6
        fp = original: old = source
        Select Case i
            Case 1: old.Baseline.CombinedDigest = String$(64, "0")
            Case 2: old.Baseline.ABatchDigest = String$(64, "0")
            Case 3: old.Baseline.BBatchDigest = String$(64, "0")
            Case 4: fp.AFingerprints(1) = "changed without batch"
            Case 5: fp.BFingerprints(1) = "changed without batch"
            Case 6
                fp.AFingerprints(1) = "changed": fp.ABatchFingerprint = "changed batch"
        End Select
        r = VATStage2CompareBaselineDelta(fp, old)
        Check "批次与多重集双向矛盾拒绝" & i, r.Status = VATS2_DELTA_INVALID_CONTRACT And NoCurrent(r) And NoResiduals(r)
    Next i
End Sub
Private Sub TestPurity()
    Dim fp As VATS2FingerprintResult, r As VATS2IncrementalDeltaResult, s As VATS2IncrementalDeltaResult
    Dim beforeFp As String, beforeOld As String, beforeFiles As String, i As Long
    DuplicateFixture 3: Call Remember: DuplicateFixture 2: fp = Fingerprints()
    beforeFp = FpSignature(fp): beforeOld = OldSignature(old): beforeFiles = FileListing()
    r = VATStage2CompareBaselineDelta(fp, old)
    For i = 1 To 3
        s = VATStage2CompareBaselineDelta(fp, old)
        Check "重复调用完全确定" & i, ResultSignature(r) = ResultSignature(s)
    Next i
    Check "输入FP所有字段完全不改", beforeFp = FpSignature(fp)
    Check "输入baseline所有字段完全不改", beforeOld = OldSignature(old)
    Check "不创建修改baseline或项目文件", beforeFiles = FileListing()
End Sub
Private Function Pack(ByVal s As String) As String
    Pack = CStr(Len(s)) & ":" & s
End Function
Private Function FpSignature(ByRef fp As VATS2FingerprintResult) As String
    Dim s As String, i As Long
    s = Pack(CStr(fp.Status)) & Pack(CStr(fp.ARecordCount)) & Pack(CStr(fp.BRecordCount)) & Pack(fp.ABatchFingerprint) & Pack(fp.BBatchFingerprint) & Pack(fp.CombinedFingerprint)
    s = s & Pack(fp.ErrorSide) & Pack(CStr(fp.ErrorIndex)) & Pack(fp.ErrorReason)
    For i = 1 To fp.ARecordCount: s = s & Pack(fp.AFingerprints(i)): Next i
    For i = 1 To fp.BRecordCount: s = s & Pack(fp.BFingerprints(i)): Next i
    FpSignature = s
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
Private Function NoCurrent(ByRef r As VATS2IncrementalDeltaResult) As Boolean
    Dim n As Long, ae As Long, be As Long
    On Error Resume Next
    n = LBound(r.ACurrent): ae = Err.Number: Err.Clear
    n = LBound(r.BCurrent): be = Err.Number: Err.Clear
    On Error GoTo 0
    NoCurrent = ae <> 0 And be <> 0
End Function
Private Function NoResiduals(ByRef r As VATS2IncrementalDeltaResult) As Boolean
    Dim n As Long, ae As Long, be As Long
    On Error Resume Next
    n = LBound(r.AResiduals): ae = Err.Number: Err.Clear
    n = LBound(r.BResiduals): be = Err.Number: Err.Clear
    On Error GoTo 0
    NoResiduals = ae <> 0 And be <> 0
End Function
Private Function FileListing() As String
    '只读验证相关目录现状；测试不保存任何baseline，不生成临时文件。
    Dim fs As Object, folder As Variant, file As Object, s As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    For Each folder In Array(CurDir$, Environ$("LOCALAPPDATA") & "\VATCheck", Environ$("TEMP") & "\VATCheck")
        s = s & Pack(CStr(fs.FolderExists(folder)))
        If fs.FolderExists(folder) Then
            For Each file In fs.GetFolder(folder).Files
                s = s & Pack(file.Name) & Pack(CStr(file.Size)) & Pack(CStr(file.DateLastModified))
            Next file
        End If
    Next folder
    FileListing = s
End Function
Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        passed = passed + 1
    Else
        failed = failed + 1: log = log & "FAIL " & name & vbCrLf
    End If
End Sub
