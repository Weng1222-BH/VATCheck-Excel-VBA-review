Attribute VB_Name = "modVATStage2BaselineTests"
Option Explicit
Option Compare Binary

Private passed As Long, failed As Long, log As String
Private a As VATS2ASnapshotResult, b As VATS2BSnapshotResult
Private fs As Object, testRoot As String, root As String, target As String

Public Function VATStage2BaselineStore_SelfTest() As String
    On Error GoTo Unexpected
    passed = 0: failed = 0: log = ""
    Set fs = CreateObject("Scripting.FileSystemObject")
    testRoot = fs.BuildPath(fs.GetSpecialFolder(2).Path, "VATCheck-C72-" & fs.GetTempName)
    fs.CreateFolder testRoot
    root = fs.BuildPath(testRoot, "nested\store"): target = fs.BuildPath(root, "baseline-v1.dat")
    TestSHA
    TestBuild
    TestFiles
    TestInvalidState
    GoTo Done
Unexpected:
    failed = failed + 1: log = log & "UNEXPECTED " & Err.Number & " " & Err.Description & vbCrLf
Done:
    '仅删除本次创建、经绝对路径及前缀检查的测试目录；绝不访问正式baseline。
    On Error Resume Next
    If Len(testRoot) > 0 Then
        If fs.GetParentFolderName(fs.GetAbsolutePathName(testRoot)) = fs.GetSpecialFolder(2).Path Then
            If Left$(fs.GetFileName(testRoot), 13) = "VATCheck-C72-" Then fs.DeleteFolder testRoot, True
        End If
    End If
    On Error GoTo 0
    If failed = 0 Then
        VATStage2BaselineStore_SelfTest = "PASS: " & passed & " assertions" & vbCrLf & log
    Else
        VATStage2BaselineStore_SelfTest = "FAIL: " & failed & "; PASS: " & passed & vbCrLf & log
    End If
End Function

Private Sub TestSHA()
    Dim expected As Variant, lengths As Variant, i As Long, result As String, errorCode As Long
    'FIPS 180-4 / NIST NSRL公开向量，绝不以本实现的结果生成预期。
    'https://www.nist.gov/itl/ai/ai-standards-and-guidelines-group/nsrl-test-data
    Check "SHA空字符串", VATStage2SHA256("") = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
    Check "SHA abc", VATStage2SHA256("abc") = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    Check "SHA quick brown fox", VATStage2SHA256("The quick brown fox jumps over the lazy dog") = "d7a8fbb307d7809469ca9abcb0082e4f8d5651e46d3cdb762d02d0bf37c9e592"
    Check "SHA百万a公开向量", VATStage2SHA256(String$(1000000, "a")) = "cdc76e5c9914fb9281a1c7e284d73e67f1809a48a497200e046d39ccc7112cd0"
    '以下附加向量预期由独立.NET SHA256及严格UTF8无BOM生成，生产代码无此依赖。
    Check "SHA中文", VATStage2SHA256("中文") = "72726d8818f693066ceb69afa364218b692e62ea92b385782363780f47529c21"
    Check "SHA补充平面Unicode", VATStage2SHA256(ChrW$(&HD83D) & ChrW$(&HDE00)) = "f0443a342c5ef54783a111b51ba56c938e474c32324d90c3a60c9c8e3a37e2d9"
    Check "SHA嵌入NUL", VATStage2SHA256("a" & vbNullChar & "b") = "59b271ae1bbcb1d31d41929817f4b16fb439eb4f31520b5ad1d5ce98920a7138"
    lengths = Array(55, 56, 64, 65)
    expected = Array("9f4390f8d30c2dd92ec9f095b65e2b9ae9b0a925a5258e241c9f1e910f734318", _
        "b35439a4ac6f0948b6d6f9e3c6af0f5f590ce20f1bde7090ef7970686ec6738a", _
        "ffe054fe7ae0cb6dc65c3af9b61d5209f439851db43d0ba5997337df154668eb", _
        "635361c48bb9eab14198e76ea8ab7f1a41685d6ad62aa9146d301d4f17eb0ae0")
    For i = 0 To 3: Check "SHA填充边界" & lengths(i), VATStage2SHA256(String$(lengths(i), "a")) = expected(i): Next i
    Check "SHA重复调用", VATStage2SHA256("中文abc") = VATStage2SHA256("中文abc")
    On Error Resume Next
    result = VATStage2SHA256(ChrW$(&HD800)): errorCode = Err.Number: Err.Clear
    On Error GoTo 0
    Check "孤立高代理项拒绝", errorCode <> 0
    On Error Resume Next
    result = VATStage2SHA256(ChrW$(&HDC00)): errorCode = Err.Number: Err.Clear
    On Error GoTo 0
    Check "孤立低代理项拒绝", errorCode <> 0
End Sub

Private Sub Fixture(Optional ByVal count As Long = 1)
    Dim aa As VATS2ASnapshotResult, bb As VATS2BSnapshotResult, i As Long
    a = aa: b = bb: a.RecordCount = count: b.RecordCount = count
    If count > 0 Then ReDim a.Records(1 To count): ReDim b.Records(1 To count)
    For i = 1 To count
        With a.Records(i)
            .AIndex = i: .ExcelRow = i + 10: .InvoiceDigitsRaw = "C72_SYNTHETIC_INVOICE_000000" & i
            .AmountRaw = "C72_AMOUNT_CANARY_987654.321": .IsCompleted = True
            .InvoiceCellAddress = "D" & .ExcelRow: .AmountCellAddress = "I" & .ExcelRow
        End With
        With b.Records(i)
            .BIndex = i: .ExcelRow = i + 20: .SupplierTextRaw = "C72_SYNTHETIC_SUPPLIER_仅合成测试_" & i
            .AmountRaw = "C72_B_AMOUNT_CANARY_123456.789": .IsCompleted = True
            .SupplierCellAddress = "B" & .ExcelRow: .AmountCellAddress = "F" & .ExcelRow
        End With
    Next i
End Sub
Private Function Fingerprints() As VATS2FingerprintResult
    Fingerprints = VATStage2BuildFingerprints(a, b)
End Function
Private Function Built() As VATS2BaselineState
    Dim fp As VATS2FingerprintResult
    fp = Fingerprints(): Built = VATStage2BuildBaseline(fp)
End Function
Private Function Identity(ByRef r As VATS2BaselineState) As String
    Dim s As String, i As Long
    s = r.Status & ":" & r.SchemaVersion & r.FingerprintProtocol & r.DigestProtocol & ":" & r.ARecordCount & ":" & r.BRecordCount
    s = s & r.ABatchDigest & r.BBatchDigest & r.CombinedDigest
    For i = 1 To r.ARecordCount: s = s & r.ARecordDigests(i): Next i
    For i = 1 To r.BRecordCount: s = s & r.BRecordDigests(i): Next i
    Identity = s
End Function
Private Sub TestBuild()
    Dim fp As VATS2FingerprintResult, copy As VATS2FingerprintResult, r As VATS2BaselineState, s As VATS2BaselineState
    Dim ar As VATS2ASnapshotRecord, br As VATS2BSnapshotRecord, i As Long
    Fixture 2: fp = Fingerprints(): copy = fp: r = VATStage2BuildBaseline(fp)
    Check "正常build及版本", r.Status = VATS2_BASELINE_OK And r.SchemaVersion = "VATBASE1" And r.FingerprintProtocol = "VATFP1" And r.DigestProtocol = "SHA256-UTF8"
    Check "完整记录数量", r.ARecordCount = 2 And r.BRecordCount = 2
    Check "记录摘要来源A", r.ARecordDigests(1) = VATStage2SHA256(fp.AFingerprints(1)) Or r.ARecordDigests(2) = VATStage2SHA256(fp.AFingerprints(1))
    Check "记录摘要来源B", r.BRecordDigests(1) = VATStage2SHA256(fp.BFingerprints(1)) Or r.BRecordDigests(2) = VATStage2SHA256(fp.BFingerprints(1))
    Check "三个批次摘要来源", r.ABatchDigest = VATStage2SHA256(fp.ABatchFingerprint) And r.BBatchDigest = VATStage2SHA256(fp.BBatchFingerprint) And r.CombinedDigest = VATStage2SHA256(fp.CombinedFingerprint)
    Check "UTC元数据", Len(r.CreatedUtc) = 20 And Right$(r.CreatedUtc, 1) = "Z" And r.CreatedUtc = r.UpdatedUtc
    Check "digest排序", StrComp(r.ARecordDigests(1), r.ARecordDigests(2), vbBinaryCompare) <= 0 And StrComp(r.BRecordDigests(1), r.BRecordDigests(2), vbBinaryCompare) <= 0
    s = VATStage2BuildBaseline(fp): Check "build重复业务身份确定", Identity(r) = Identity(s)
    Check "输入fp元数据不变", fp.Status = copy.Status And fp.ARecordCount = copy.ARecordCount And fp.BRecordCount = copy.BRecordCount And fp.ErrorSide = copy.ErrorSide And fp.ErrorIndex = copy.ErrorIndex And fp.ErrorReason = copy.ErrorReason
    Check "输入fp批次不变", fp.ABatchFingerprint = copy.ABatchFingerprint And fp.BBatchFingerprint = copy.BBatchFingerprint And fp.CombinedFingerprint = copy.CombinedFingerprint
    For i = 1 To 2: Check "输入fp数组不变" & i, fp.AFingerprints(i) = copy.AFingerprints(i) And fp.BFingerprints(i) = copy.BFingerprints(i): Next i
    ar = a.Records(1): a.Records(1) = a.Records(2): a.Records(2) = ar
    br = b.Records(1): b.Records(1) = b.Records(2): b.Records(2) = br
    For i = 1 To 2
        a.Records(i).AIndex = i: a.Records(i).ExcelRow = 100 + i: a.Records(i).InvoiceCellAddress = "K" & (100 + i)
        b.Records(i).BIndex = i: b.Records(i).ExcelRow = 200 + i: b.Records(i).SupplierCellAddress = "R" & (200 + i)
    Next i
    s = Built(): Check "换序当前位置不影响baseline业务身份", Identity(r) = Identity(s)
    For i = 1 To 10
        Call Fixture: r = Built()
        Select Case i
            Case 1: a.Records(1).InvoiceDigitsRaw = "changed"
            Case 2: b.Records(1).SupplierTextRaw = "changed"
            Case 3: a.Records(1).AmountRaw = CDec(1)
            Case 4: b.Records(1).AmountRaw = CDec(1)
            Case 5: a.Records(1).IsCompleted = False
            Case 6: b.Records(1).IsCompleted = False
            Case 7: a.Records(1).InvoiceHasFormula = True
            Case 8: b.Records(1).SupplierHasFormula = True
            Case 9: a.Records(1).AmountHasFormula = True
            Case 10: b.Records(1).AmountHasFormula = True
        End Select
        s = Built(): Check "业务内容变化" & i, r.CombinedDigest <> s.CombinedDigest
    Next i
    Call Fixture: r = Built()
    Fixture 2: a.Records(2).InvoiceDigitsRaw = a.Records(1).InvoiceDigitsRaw: b.Records(2).SupplierTextRaw = b.Records(1).SupplierTextRaw
    s = Built()
    Check "重复摘要保留multiplicity", s.ARecordCount = 2 And s.ARecordDigests(1) = s.ARecordDigests(2) And s.BRecordDigests(1) = s.BRecordDigests(2)
    Check "重复次数改变批次", s.ABatchDigest <> r.ABatchDigest And s.BBatchDigest <> r.BBatchDigest And s.CombinedDigest <> r.CombinedDigest
    Fixture 0: r = Built(): Check "零记录build", r.Status = VATS2_BASELINE_OK And r.ARecordCount = 0 And r.BRecordCount = 0 And Len(r.CombinedDigest) = 64
    fp = Fingerprints(): fp.Status = VATS2_FINGERPRINT_INVALID_INPUT: r = VATStage2BuildBaseline(fp)
    Check "非OK指纹拒绝且无部分摘要", r.Status = VATS2_BASELINE_INVALID_INPUT And r.CombinedDigest = "" And r.ARecordCount = 0
    fp = copy: fp.ARecordCount = 3: r = VATStage2BuildBaseline(fp): Check "指纹数组契约拒绝", r.Status = VATS2_BASELINE_INVALID_INPUT
    fp = copy: fp.AFingerprints(1) = ChrW$(&HD800): r = VATStage2BuildBaseline(fp)
    Check "指纹非法UTF16拒绝且不发布部分结果", r.Status = VATS2_BASELINE_INVALID_INPUT And r.ARecordCount = 0 And r.CombinedDigest = ""
End Sub

Private Sub TestFiles()
    Dim r As VATS2BaselineState, copy As VATS2BaselineState, fp As VATS2FingerprintResult
    Dim io As VATS2BaselineIOResult, loaded As VATS2BaselineLoadResult, s As String, lines As Variant
    Dim i As Long, original As String, part As Variant, f As Integer, pending As String
    loaded = VATStage2LoadBaseline(root)
    Check "首次NOT_FOUND", loaded.Status = VATS2_BASELINE_NOT_FOUND And loaded.Baseline.Status = VATS2_BASELINE_NOT_FOUND
    Check "load不创建目录", Not fs.FolderExists(root)
    Call Fixture: fp = Fingerprints(): r = VATStage2BuildBaseline(fp): copy = r
    io = VATStage2SaveBaseline(r, root): Check "自动创建嵌套目录并保存", io.Status = VATS2_BASELINE_OK And fs.FileExists(target)
    loaded = VATStage2LoadBaseline(root): Check "save/load往返", loaded.Status = VATS2_BASELINE_OK And Identity(loaded.Baseline) = Identity(r)
    Check "UTC完整往返", loaded.Baseline.CreatedUtc = r.CreatedUtc And loaded.Baseline.UpdatedUtc = r.UpdatedUtc
    Check "save输入完全不改", Identity(r) = Identity(copy) And r.CreatedUtc = copy.CreatedUtc And r.UpdatedUtc = copy.UpdatedUtc
    s = ReadText(target): original = s
    For Each part In Array(CStr(a.Records(1).InvoiceDigitsRaw), CStr(b.Records(1).SupplierTextRaw), CStr(a.Records(1).AmountRaw), CStr(b.Records(1).AmountRaw), _
        fp.AFingerprints(1), fp.BFingerprints(1), fp.ABatchFingerprint, fp.BBatchFingerprint, fp.CombinedFingerprint, "VATFP1-RECORD", "VATFP1-BATCH", "VATFP1-COMBINED")
        Check "落盘隐私检查" & Len(CStr(part)), InStr(1, s, CStr(part), vbBinaryCompare) = 0
    Next part
    Check "文件只含digest及元数据", InStr(s, r.ARecordDigests(1)) > 0 And InStr(s, r.CombinedDigest) > 0
    Check "无BOM固定LF", Left$(s, 8) = "VATBASE1" And InStr(s, vbCr) = 0 And Right$(s, 1) = vbLf
    io = VATStage2SaveBaseline(r, root): Check "重复保存字节完全一致", io.Status = VATS2_BASELINE_OK And ReadText(target) = original
    loaded = VATStage2LoadBaseline(root): Check "重复load确定", Identity(loaded.Baseline) = Identity(r)
    '已有暂存文件使创建失败，旧baseline保持，且不删除非本次文件。
    pending = fs.BuildPath(root, "baseline-v1.pending.tmp"): WriteText pending, "test-owned-blocker"
    a.Records(1).IsCompleted = False: r = Built(): io = VATStage2SaveBaseline(r, root)
    Check "临时创建失败返回IO_ERROR", io.Status = VATS2_BASELINE_IO_ERROR
    Check "临时失败旧文件逐字节不变", ReadText(target) = original
    Check "不删除他人暂存文件", ReadText(pending) = "test-owned-blocker"
    fs.DeleteFile pending
    '锁住正式文件阻止最终替换，验证本次暂存文件清理及旧文件保留。
    f = FreeFile: Open target For Binary Access Read Lock Read Write As #f
    io = VATStage2SaveBaseline(r, root): Close #f
    Check "替换失败明确IO_ERROR", io.Status = VATS2_BASELINE_IO_ERROR
    Check "替换失败保留旧baseline", ReadText(target) = original
    Check "替换失败清理自有临时文件", Not fs.FileExists(pending)
    io = VATStage2SaveBaseline(r, root): loaded = VATStage2LoadBaseline(root)
    Check "解除锁后可正常替换", io.Status = VATS2_BASELINE_OK And Identity(loaded.Baseline) = Identity(r) And ReadText(target) <> original
    f = FreeFile: Open target For Binary Access Read Lock Read Write As #f
    loaded = VATStage2LoadBaseline(root): Close #f
    Check "读取被锁文件IO_ERROR", loaded.Status = VATS2_BASELINE_IO_ERROR
    '每个损坏用例均独立重置原始文件；结构问题即使重算校验和也必须拒绝。
    For i = 1 To 17
        lines = Split(original, vbLf)
        Select Case i
            Case 1: lines(0) = "WRONG"
            Case 2: lines(0) = "VATBASE2"
            Case 3: lines(1) = "VATFP2"
            Case 4: lines(2) = "SHA256-ANSI"
            Case 5: lines(7) = String$(63, "a")
            Case 6: lines(8) = String$(64, "g")
            Case 7: lines(3) = "2"
            Case 8: lines(3) = "-1"
            Case 9: lines(3) = "01"
            Case 10: lines(4) = "2147483648"
            Case 11: lines(5) = "2026-02-30T00:00:00Z"
            Case 12: lines(6) = "0001-01-01T00:00:00Z"
            Case 13: lines(10) = String$(64, "A")
            Case 14: lines(5) = "2026-09-14T25:00:00Z"
            Case 15: lines(10) = ""
            Case 16: lines(11) = String$(63, "a")
            Case 17: lines(3) = " 1"
        End Select
        s = Join(lines, vbLf): s = Recheck(s): WriteText target, s
        loaded = VATStage2LoadBaseline(root)
        If i >= 2 And i <= 4 Then
            Check "不支持版本" & i, loaded.Status = VATS2_BASELINE_UNSUPPORTED_VERSION
        Else
            Check "损坏结构拒绝" & i, loaded.Status = VATS2_BASELINE_CORRUPT
        End If
        Check "错误不返回部分baseline" & i, loaded.Baseline.CombinedDigest = "" And loaded.Baseline.ARecordCount = 0
    Next i
    WriteText target, Left$(original, Len(original) - 30): loaded = VATStage2LoadBaseline(root)
    Check "截断CORRUPT", loaded.Status = VATS2_BASELINE_CORRUPT
    WriteText target, original & "junk": loaded = VATStage2LoadBaseline(root)
    Check "尾部垃圾CORRUPT", loaded.Status = VATS2_BASELINE_CORRUPT
    WriteText target, original & vbLf: loaded = VATStage2LoadBaseline(root)
    Check "多余空行CORRUPT", loaded.Status = VATS2_BASELINE_CORRUPT
    WriteText target, "": loaded = VATStage2LoadBaseline(root)
    Check "空文件CORRUPT", loaded.Status = VATS2_BASELINE_CORRUPT
    lines = Split(original, vbLf): lines(10) = String$(64, "a")
    WriteText target, Join(lines, vbLf): loaded = VATStage2LoadBaseline(root)
    Check "合法hex变动校验和拒绝", loaded.Status = VATS2_BASELINE_CORRUPT
    Fixture 0: r = Built(): io = VATStage2SaveBaseline(r, root): loaded = VATStage2LoadBaseline(root)
    Check "零记录往返", io.Status = VATS2_BASELINE_OK And loaded.Status = VATS2_BASELINE_OK And Identity(r) = Identity(loaded.Baseline)
    Fixture 2: a.Records(2).InvoiceDigitsRaw = a.Records(1).InvoiceDigitsRaw: b.Records(2).SupplierTextRaw = b.Records(1).SupplierTextRaw
    r = Built(): io = VATStage2SaveBaseline(r, root): loaded = VATStage2LoadBaseline(root)
    Check "重复multiset往返", loaded.Status = VATS2_BASELINE_OK And Identity(r) = Identity(loaded.Baseline) And loaded.Baseline.ARecordCount = 2
    loaded = VATStage2LoadBaseline("relative"): Check "相对目录无隐式回退", loaded.Status = VATS2_BASELINE_IO_ERROR
End Sub

Private Sub TestInvalidState()
    Dim r As VATS2BaselineState, changed As VATS2BaselineState, io As VATS2BaselineIOResult, i As Long, oldFile As String
    Fixture 2: r = Built(): oldFile = ReadText(target)
    For i = 1 To 12
        changed = r
        Select Case i
            Case 1: changed.Status = VATS2_BASELINE_CORRUPT
            Case 2: changed.SchemaVersion = "VATBASE2"
            Case 3: changed.FingerprintProtocol = "VATFP2"
            Case 4: changed.DigestProtocol = "other"
            Case 5: changed.ARecordCount = -1
            Case 6: changed.BRecordCount = 3
            Case 7: changed.ARecordDigests(1) = "C72_SYNTHETIC_RAW_CONTENT"
            Case 8: changed.CombinedDigest = "raw canonical"
            Case 9: changed.CreatedUtc = "raw metadata"
            Case 10: changed.UpdatedUtc = "0001-01-01T00:00:00Z"
            Case 11: changed.ARecordDigests(1) = String$(64, "f"): changed.ARecordDigests(2) = String$(64, "0")
            Case 12: changed.ARecordCount = 0
        End Select
        io = VATStage2SaveBaseline(changed, root)
        Check "非法state拒绝" & i, io.Status = VATS2_BASELINE_INVALID_INPUT
        Check "非法state不写磁盘" & i, ReadText(target) = oldFile And Not fs.FileExists(fs.BuildPath(root, "baseline-v1.pending.tmp"))
    Next i
End Sub
Private Function Recheck(ByVal s As String) As String
    Dim body As String
    body = Left$(s, Len(s) - 65): Recheck = body & VATStage2SHA256(body) & vbLf
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
        passed = passed + 1
    Else
        failed = failed + 1: log = log & "FAIL " & name & vbCrLf
    End If
End Sub
