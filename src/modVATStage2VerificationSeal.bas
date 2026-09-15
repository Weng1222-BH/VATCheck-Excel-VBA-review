Attribute VB_Name = "modVATStage2VerificationSeal"
Option Explicit
Option Compare Binary

Public Enum VATS2VerificationStatus
    VATS2_VERIFY_OK = 0
    VATS2_VERIFY_NOT_FOUND = 1
    VATS2_VERIFY_INVALID_INPUT = 2
    VATS2_VERIFY_CORRUPT = 3
    VATS2_VERIFY_IO_ERROR = 4
    VATS2_VERIFY_UNSUPPORTED_VERSION = 5
End Enum
Public Type VATS2VerificationSealState
    Status As VATS2VerificationStatus
    SchemaVersion As String
    BindingProtocol As String
    PolicyProtocol As String
    FingerprintProtocol As String
    DigestProtocol As String
    ARecordCount As Long
    BRecordCount As Long
    VerifiedUtc As String
    ABatchDigest As String
    BBatchDigest As String
    CombinedDigest As String
End Type
Public Type VATS2VerificationSealIOResult
    Status As VATS2VerificationStatus
    ErrorReason As String
End Type
Public Type VATS2VerificationSealLoadResult
    Status As VATS2VerificationStatus
    Seal As VATS2VerificationSealState
    UnsupportedField As String
    ErrorReason As String
End Type
Public Enum VATS2ReuseStatus
    VATS2_REUSE_OK = 0
    VATS2_REUSE_INVALID_INPUT = 1
    VATS2_REUSE_INVALID_CONTRACT = 2
End Enum
Public Enum VATS2ReuseDecision
    VATS2_FULL_VALIDATION_REQUIRED = 0
    VATS2_REUSE_ELIGIBLE = 1
End Enum
Public Enum VATS2ReuseReason
    VATS2_REUSE_FIRST_RUN = 1
    VATS2_REUSE_BASELINE_UNAVAILABLE = 2
    VATS2_REUSE_SEAL_NOT_FOUND = 4
    VATS2_REUSE_SEAL_UNAVAILABLE = 8
    VATS2_REUSE_DATA_NEW_OR_CHANGED = 16
    VATS2_REUSE_DATA_REMOVED_OR_CHANGED = 32
    VATS2_REUSE_BATCH_CHANGED = 64
    VATS2_REUSE_SEAL_BASELINE_MISMATCH = 128
    VATS2_REUSE_POLICY_VERSION_MISMATCH = 256
End Enum
Public Type VATS2ReuseEligibilityResult
    Status As VATS2ReuseStatus
    Decision As VATS2ReuseDecision
    ReasonFlags As Long
    ErrorReason As String
End Type

Private Type SystemTime
    Year As Integer
    Month As Integer
    DayOfWeek As Integer
    Day As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    Milliseconds As Integer
End Type
#If VBA7 Then
    Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef value As SystemTime)
    Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal oldName As LongPtr, ByVal newName As LongPtr, ByVal flags As Long) As Long
#Else
    Private Declare Sub GetSystemTime Lib "kernel32" (ByRef value As SystemTime)
    Private Declare Function MoveFileExW Lib "kernel32" (ByVal oldName As Long, ByVal newName As Long, ByVal flags As Long) As Long
#End If
Private Const BASE_FILE As String = "verified-v1.dat"
Private Const TEMP_FILE As String = "verified-v1.pending.tmp"

'唯一构建入口只接受受控运行结果；不接受裸FinalDecision或任意fp的事后拼接。
Public Function VATStage2BuildVerificationSeal(ByRef run As VATS2BoundAuditRunResult) As VATS2VerificationSealState
    Dim r As VATS2VerificationSealState, blank As VATS2VerificationSealState, i As Long
    On Error GoTo Invalid
    Need run.Status = VATS2_BOUND_RUN_OK
    Need run.BindingProtocol = "VATRUN1" And run.FingerprintProtocol = "VATFP1" And run.DigestProtocol = "SHA256-UTF8"
    Need run.ARecordCount >= 0 And run.BRecordCount >= 0
    Need IsDigest(run.ABatchDigest) And IsDigest(run.BBatchDigest) And IsDigest(run.CombinedDigest)
    With run.Aggregate
        Need .Status = VATS2_AGGREGATE_OK And .AScanSucceeded And .BScanSucceeded
        Need VarType(.TotalsEqual) = vbBoolean
        Need .TotalsEqual
        Need .AWarningCount = 0 And .BWarningCount = 0 And .AIssueCount = 0 And .BIssueCount = 0
        Need .AEmptyWarningCount = 0 And .BEmptyWarningCount = 0
        Need .ACompletedCount >= 0 And .ACompletedCount <= run.ARecordCount
        Need .BCompletedCount >= 0 And .BCompletedCount <= run.BRecordCount
        Need .AIncludedCount = .ACompletedCount And .BIncludedCount = .BCompletedCount
    End With
    Need run.AuditFindings.Status = VATS2_AUDIT_OK And run.AuditFindings.FindingCount = 0
    With run.FinalDecision
        Need .Status = VATS2_FINAL_DECISION_OK And .Verdict = VATS2_FINAL_VERIFIED
        Need .AggregateStatus = run.Aggregate.Status And .AuditStatus = run.AuditFindings.Status
        Need .EffectiveAmountStatus = VATS2_EFFECTIVE_AMOUNT_OK
        Need Not .RequiresManualReview And Not .HasAggregateMismatch And Not .HasAggregateWarnings And Not .HasFindings
        Need .FindingCount = 0
        Need VarType(.TotalsEqual) = vbBoolean
        Need .TotalsEqual
        SameDecimal .TotalA, run.Aggregate.TotalA
        SameDecimal .TotalB, run.Aggregate.TotalB
        SameDecimal .Difference, run.Aggregate.Difference
        Need .ACompletedCount = run.Aggregate.ACompletedCount And .BCompletedCount = run.Aggregate.BCompletedCount
        Need .AIncludedCount = run.Aggregate.AIncludedCount And .BIncludedCount = run.Aggregate.BIncludedCount
        Need .AWarningCount = run.Aggregate.AWarningCount And .BWarningCount = run.Aggregate.BWarningCount
        Need .NoCompletedRecords = (.ACompletedCount = 0 And .BCompletedCount = 0)
        Need .ShortSuffixWaivedCount >= 0 And .ShortSuffixWaivedCount = run.AuditFindings.ShortSuffixWaivedCount
        For i = 1 To 18
            Need .CodeCounts(i) = 0 And run.AuditFindings.CodeCounts(i) = 0
        Next i
    End With
    r.SchemaVersion = "VATVERIFY1": r.BindingProtocol = run.BindingProtocol
    r.PolicyProtocol = "VATSTAGE2-POLICY1": r.FingerprintProtocol = run.FingerprintProtocol: r.DigestProtocol = run.DigestProtocol
    r.ARecordCount = run.ARecordCount: r.BRecordCount = run.BRecordCount
    r.ABatchDigest = run.ABatchDigest: r.BBatchDigest = run.BBatchDigest: r.CombinedDigest = run.CombinedDigest
    r.VerifiedUtc = UtcNow()
    VATStage2BuildVerificationSeal = r: Exit Function
Invalid:
    blank.Status = VATS2_VERIFY_INVALID_INPUT: VATStage2BuildVerificationSeal = blank
End Function

Private Sub SameDecimal(ByVal left As Variant, ByVal right As Variant)
    '只核实已发布摘要一致，不重算金额或重新判断业务结论。
    Need VarType(left) = vbDecimal And VarType(right) = vbDecimal
    Need left = right
End Sub

Private Function Serialize(ByRef r As VATS2VerificationSealState) As String
    Dim s As String
    s = r.SchemaVersion & vbLf & r.BindingProtocol & vbLf & r.PolicyProtocol & vbLf & _
        r.FingerprintProtocol & vbLf & r.DigestProtocol & vbLf & CStr(r.ARecordCount) & vbLf & _
        CStr(r.BRecordCount) & vbLf & r.VerifiedUtc & vbLf & r.ABatchDigest & vbLf & _
        r.BBatchDigest & vbLf & r.CombinedDigest & vbLf
    Serialize = s & VATStage2SHA256(s) & vbLf
End Function

Private Function Parse(ByVal s As String) As VATS2VerificationSealLoadResult
    Dim r As VATS2VerificationSealLoadResult, blank As VATS2VerificationSealLoadResult, lines As Variant
    Dim field As String, phase As VATS2VerificationStatus
    phase = VATS2_VERIFY_CORRUPT
    On Error GoTo Invalid
    lines = Split(s, vbLf): Need UBound(lines) = 12: Need lines(12) = ""
    If lines(0) <> "VATVERIFY1" Then
        If Left$(lines(0), 9) = "VATVERIFY" Then phase = VATS2_VERIFY_UNSUPPORTED_VERSION: field = "SCHEMA"
        GoTo Invalid
    End If
    With r.Seal
        .SchemaVersion = lines(0): .BindingProtocol = lines(1): .PolicyProtocol = lines(2)
        .FingerprintProtocol = lines(3): .DigestProtocol = lines(4)
        field = Unsupported(.BindingProtocol, .PolicyProtocol, .FingerprintProtocol, .DigestProtocol)
        If Len(field) > 0 Then phase = VATS2_VERIFY_UNSUPPORTED_VERSION: GoTo Invalid
        .ARecordCount = ParseCount(CStr(lines(5))): .BRecordCount = ParseCount(CStr(lines(6)))
        .VerifiedUtc = lines(7): .ABatchDigest = lines(8): .BBatchDigest = lines(9): .CombinedDigest = lines(10)
    End With
    ValidateState r.Seal
    Need IsDigest(CStr(lines(11)))
    Need VATStage2SHA256(Left$(s, Len(s) - 65)) = lines(11)
    Parse = r: Exit Function
Invalid:
    blank.Status = phase: blank.Seal.Status = phase: blank.UnsupportedField = field
    blank.ErrorReason = "Seal格式、版本或完整性校验失败。": Parse = blank
End Function

Private Function Unsupported(ByVal binding As String, ByVal policy As String, ByVal fingerprint As String, ByVal digest As String) As String
    If binding <> "VATRUN1" Then
        Unsupported = "BINDING"
    ElseIf policy <> "VATSTAGE2-POLICY1" Then
        Unsupported = "POLICY"
    ElseIf fingerprint <> "VATFP1" Then
        Unsupported = "FINGERPRINT"
    ElseIf digest <> "SHA256-UTF8" Then
        Unsupported = "DIGEST"
    End If
End Function

Private Sub ValidateState(ByRef r As VATS2VerificationSealState)
    Need r.Status = VATS2_VERIFY_OK And r.SchemaVersion = "VATVERIFY1"
    Need Len(Unsupported(r.BindingProtocol, r.PolicyProtocol, r.FingerprintProtocol, r.DigestProtocol)) = 0
    Need r.ARecordCount >= 0 And r.BRecordCount >= 0
    Need IsDigest(r.ABatchDigest) And IsDigest(r.BBatchDigest) And IsDigest(r.CombinedDigest)
    Need IsUtc(r.VerifiedUtc)
End Sub

Public Function VATStage2SaveVerificationSeal(ByRef baseline As VATS2VerificationSealState, _
    Optional ByVal rootOverride As String = "") As VATS2VerificationSealIOResult
    Dim r As VATS2VerificationSealIOResult, check As VATS2VerificationSealLoadResult
    Dim fs As Object, stream As Object, root As String, target As String, temporary As String
    Dim payload As String, ownTemp As Boolean, phase As VATS2VerificationStatus, errorCode As Long
    On Error GoTo Failed
    phase = VATS2_VERIFY_INVALID_INPUT
    ValidateState baseline
    payload = Serialize(baseline)
    Set fs = CreateObject("Scripting.FileSystemObject")
    root = StoreRoot(rootOverride, fs): target = fs.BuildPath(root, BASE_FILE): temporary = fs.BuildPath(root, TEMP_FILE)
    phase = VATS2_VERIFY_IO_ERROR
    EnsureFolder fs, root
    '固定暂存名通过创建且禁止覆盖实现单写者互斥；已有暂存文件绝不视为本次所有。
    Set stream = fs.CreateTextFile(temporary, False, False): ownTemp = True
    '此处payload仅含ASCII协议、digest和UTC元数据，ANSI输出与UTF-8无BOM字节相同。
    stream.Write payload: stream.Close: Set stream = Nothing
    check = ReadSealFile(temporary)
    If check.Status <> VATS2_VERIFY_OK Then Err.Raise vbObjectError + 722
    Need Serialize(check.Seal) = payload
    '同目录同卷rename替换，不先删除旧文件，不允许跨卷copy/delete降级。
    If MoveFileExW(StrPtr(temporary), StrPtr(target), 9) = 0 Then
        errorCode = Err.LastDllError: Err.Raise vbObjectError + 723
    End If
    ownTemp = False: r.Status = VATS2_VERIFY_OK
    VATStage2SaveVerificationSeal = r: Exit Function
Failed:
    If errorCode = 0 Then errorCode = Err.Number
    r.Status = phase: r.ErrorReason = "Seal保存失败，错误码=" & CStr(errorCode)
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If ownTemp Then fs.DeleteFile temporary, False
    On Error GoTo 0
    VATStage2SaveVerificationSeal = r
End Function

Public Function VATStage2LoadVerificationSeal(Optional ByVal rootOverride As String = "") As VATS2VerificationSealLoadResult
    Dim r As VATS2VerificationSealLoadResult, fs As Object, root As String
    On Error GoTo Failed
    Set fs = CreateObject("Scripting.FileSystemObject")
    root = StoreRoot(rootOverride, fs)
    r = ReadSealFile(fs.BuildPath(root, BASE_FILE))
    VATStage2LoadVerificationSeal = r: Exit Function
Failed:
    r.Status = VATS2_VERIFY_IO_ERROR: r.Seal.Status = r.Status
    r.ErrorReason = "Seal读取路径失败，错误码=" & CStr(Err.Number)
    VATStage2LoadVerificationSeal = r
End Function

Private Function ReadSealFile(ByVal path As String) As VATS2VerificationSealLoadResult
    Dim r As VATS2VerificationSealLoadResult, f As Integer, opened As Boolean, s As String, fs As Object
    On Error GoTo Failed
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FileExists(path) Then
        '同名目录不是正常首次运行。
        If fs.FolderExists(path) Then Err.Raise vbObjectError + 724
        r.Status = VATS2_VERIFY_NOT_FOUND: r.Seal.Status = r.Status
        ReadSealFile = r: Exit Function
    End If
    f = FreeFile: Open path For Binary Access Read Lock Write As #f: opened = True
    s = String$(LOF(f), vbNullChar)
    If Len(s) > 0 Then Get #f, , s
    Close #f: opened = False
    ReadSealFile = Parse(s): Exit Function
Failed:
    r.Status = VATS2_VERIFY_IO_ERROR: r.Seal.Status = r.Status
    r.ErrorReason = "Seal读取失败，错误码=" & CStr(Err.Number)
    On Error Resume Next
    If opened Then Close #f
    On Error GoTo 0
    ReadSealFile = r
End Function

Private Function IsDigest(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) <> 64 Then Exit Function
    For i = 1 To 64
        ch = Mid$(s, i, 1)
        If InStr(1, "0123456789abcdef", ch, vbBinaryCompare) = 0 Then Exit Function
    Next i
    IsDigest = True
End Function
Private Function ParseCount(ByVal s As String) As Long
    Dim i As Long, n As Double, ch As Long
    Need Len(s) > 0 And Len(s) <= 10
    If Len(s) > 1 Then Need Left$(s, 1) <> "0"
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1)): Need ch >= 48 And ch <= 57
        n = n * 10 + ch - 48: Need n <= 2147483647#
    Next i
    ParseCount = CLng(n)
End Function
Private Function IsUtc(ByVal s As String) As Boolean
    Dim i As Long, y As Long, m As Long, d As Long, maxDay As Long, ch As String
    If Len(s) <> 20 Then Exit Function
    If Mid$(s, 5, 1) <> "-" Or Mid$(s, 8, 1) <> "-" Or Mid$(s, 11, 1) <> "T" Or _
        Mid$(s, 14, 1) <> ":" Or Mid$(s, 17, 1) <> ":" Or Right$(s, 1) <> "Z" Then Exit Function
    For i = 1 To 19
        Select Case i
            Case 5, 8, 11, 14, 17
            Case Else
                ch = Mid$(s, i, 1): If ch < "0" Or ch > "9" Then Exit Function
        End Select
    Next i
    y = CLng(Left$(s, 4)): m = CLng(Mid$(s, 6, 2)): d = CLng(Mid$(s, 9, 2))
    If y < 1 Or m < 1 Or m > 12 Or d < 1 Then Exit Function
    maxDay = 31
    Select Case m
        Case 4, 6, 9, 11: maxDay = 30
        Case 2
            maxDay = 28
            If (y Mod 4 = 0 And y Mod 100 <> 0) Or y Mod 400 = 0 Then maxDay = 29
    End Select
    If d > maxDay Then Exit Function
    If CLng(Mid$(s, 12, 2)) > 23 Or CLng(Mid$(s, 15, 2)) > 59 Or CLng(Mid$(s, 18, 2)) > 59 Then Exit Function
    IsUtc = True
End Function
Private Function UtcNow() As String
    Dim t As SystemTime
    GetSystemTime t
    UtcNow = Right$("0000" & CStr(t.Year), 4) & "-" & Two(t.Month) & "-" & Two(t.Day) & "T" & _
        Two(t.Hour) & ":" & Two(t.Minute) & ":" & Two(t.Second) & "Z"
End Function
Private Function Two(ByVal n As Long) As String
    Two = Right$("0" & CStr(n), 2)
End Function
Private Function StoreRoot(ByVal override As String, ByVal fs As Object) As String
    Dim root As String
    root = override
    If Len(root) = 0 Then
        root = Environ$("LOCALAPPDATA"): Need Len(root) > 0
        root = fs.BuildPath(root, "VATCheck")
    End If
    '只接受本地绝对路径，绝不以当前项目目录作为隐式回退。
    Need Len(root) >= 3: Need Mid$(root, 2, 2) = ":\"
    Need InStr(root, vbNullChar) = 0
    StoreRoot = fs.GetAbsolutePathName(root)
End Function
Private Sub EnsureFolder(ByVal fs As Object, ByVal root As String)
    If fs.FolderExists(root) Then Exit Sub
    Need Not fs.FileExists(root)
    Need Len(fs.GetParentFolderName(root)) > 0
    EnsureFolder fs, fs.GetParentFolderName(root)
    fs.CreateFolder root
End Sub
Private Sub Need(ByVal condition As Boolean)
    If Not condition Then Err.Raise vbObjectError + 721, , "Seal必要契约不满足。"
End Sub

'纯内存资格判断。绝不调用审核链、执行skip或读取历史审核结果。
Public Function VATStage2EvaluateReuseEligibility(ByRef delta As VATS2IncrementalDeltaResult, _
    ByRef baselineLoad As VATS2BaselineLoadResult, ByRef sealLoad As VATS2VerificationSealLoadResult) As VATS2ReuseEligibilityResult
    Dim r As VATS2ReuseEligibilityResult, phase As VATS2ReuseStatus, field As String
    On Error GoTo Invalid
    phase = VATS2_REUSE_INVALID_INPUT
    Need delta.Status >= VATS2_DELTA_OK And delta.Status <= VATS2_DELTA_BASELINE_UNAVAILABLE
    Need baselineLoad.Status >= VATS2_BASELINE_OK And baselineLoad.Status <= VATS2_BASELINE_UNSUPPORTED_VERSION
    Need baselineLoad.Status <> VATS2_BASELINE_INVALID_INPUT
    Need sealLoad.Status >= VATS2_VERIFY_OK And sealLoad.Status <= VATS2_VERIFY_UNSUPPORTED_VERSION
    Need sealLoad.Status <> VATS2_VERIFY_INVALID_INPUT
    phase = VATS2_REUSE_INVALID_CONTRACT
    Need delta.BaselineSourceStatus = baselineLoad.Status
    Select Case delta.Status
        Case VATS2_DELTA_FIRST_RUN
            Need baselineLoad.Status = VATS2_BASELINE_NOT_FOUND
            r.ReasonFlags = r.ReasonFlags Or VATS2_REUSE_FIRST_RUN
        Case VATS2_DELTA_BASELINE_UNAVAILABLE
            Need baselineLoad.Status = VATS2_BASELINE_CORRUPT Or baselineLoad.Status = VATS2_BASELINE_IO_ERROR Or baselineLoad.Status = VATS2_BASELINE_UNSUPPORTED_VERSION
            r.ReasonFlags = r.ReasonFlags Or VATS2_REUSE_BASELINE_UNAVAILABLE
        Case VATS2_DELTA_OK
            Need baselineLoad.Status = VATS2_BASELINE_OK
            ValidateBaselineSummary baselineLoad.Baseline
            CheckDeltaSide delta.ARecordCount, delta.AUnchangedCount, delta.ANewOrChangedCount, delta.AResidualCount, baselineLoad.Baseline.ARecordCount
            CheckDeltaSide delta.BRecordCount, delta.BUnchangedCount, delta.BNewOrChangedCount, delta.BResidualCount, baselineLoad.Baseline.BRecordCount
            Need delta.ABatchUnchanged = (delta.ANewOrChangedCount = 0 And delta.AResidualCount = 0)
            Need delta.BBatchUnchanged = (delta.BNewOrChangedCount = 0 And delta.BResidualCount = 0)
            Need delta.CombinedUnchanged = (delta.ABatchUnchanged And delta.BBatchUnchanged)
            If delta.ANewOrChangedCount > 0 Or delta.BNewOrChangedCount > 0 Then r.ReasonFlags = r.ReasonFlags Or VATS2_REUSE_DATA_NEW_OR_CHANGED
            If delta.AResidualCount > 0 Or delta.BResidualCount > 0 Then r.ReasonFlags = r.ReasonFlags Or VATS2_REUSE_DATA_REMOVED_OR_CHANGED
            If Not delta.CombinedUnchanged Then r.ReasonFlags = r.ReasonFlags Or VATS2_REUSE_BATCH_CHANGED
    End Select
    Select Case sealLoad.Status
        Case VATS2_VERIFY_NOT_FOUND
            r.ReasonFlags = r.ReasonFlags Or VATS2_REUSE_SEAL_NOT_FOUND
        Case VATS2_VERIFY_UNSUPPORTED_VERSION
            r.ReasonFlags = r.ReasonFlags Or VATS2_REUSE_SEAL_UNAVAILABLE
            If sealLoad.UnsupportedField = "POLICY" Then r.ReasonFlags = r.ReasonFlags Or VATS2_REUSE_POLICY_VERSION_MISMATCH
        Case VATS2_VERIFY_CORRUPT, VATS2_VERIFY_IO_ERROR
            r.ReasonFlags = r.ReasonFlags Or VATS2_REUSE_SEAL_UNAVAILABLE
        Case VATS2_VERIFY_OK
            Need sealLoad.Seal.Status = VATS2_VERIFY_OK
            field = Unsupported(sealLoad.Seal.BindingProtocol, sealLoad.Seal.PolicyProtocol, sealLoad.Seal.FingerprintProtocol, sealLoad.Seal.DigestProtocol)
            If sealLoad.Seal.SchemaVersion <> "VATVERIFY1" Or Len(field) > 0 Then
                r.ReasonFlags = r.ReasonFlags Or VATS2_REUSE_SEAL_UNAVAILABLE
                If sealLoad.Seal.PolicyProtocol <> "VATSTAGE2-POLICY1" Then r.ReasonFlags = r.ReasonFlags Or VATS2_REUSE_POLICY_VERSION_MISMATCH
            Else
                ValidateState sealLoad.Seal
                If baselineLoad.Status = VATS2_BASELINE_OK Then
                    If Not SameBatch(sealLoad.Seal, baselineLoad.Baseline) Then r.ReasonFlags = r.ReasonFlags Or VATS2_REUSE_SEAL_BASELINE_MISMATCH
                End If
            End If
    End Select
    If r.ReasonFlags = 0 Then
        '只在即将授予整批资格时核实delta确实消费本次传入的baseline多重集。
        '防止把另一份全UNCHANGED的delta与当前baseline混用；不重跑C7.3差异算法。
        Need delta.Status = VATS2_DELTA_OK And baselineLoad.Status = VATS2_BASELINE_OK And sealLoad.Status = VATS2_VERIFY_OK
        Need delta.AResidualGroupCount = 0 And delta.BResidualGroupCount = 0
        CheckUnchanged delta.ACurrent, delta.ARecordCount, baselineLoad.Baseline.ARecordDigests
        CheckUnchanged delta.BCurrent, delta.BRecordCount, baselineLoad.Baseline.BRecordDigests
        r.Decision = VATS2_REUSE_ELIGIBLE
    End If
    VATStage2EvaluateReuseEligibility = r: Exit Function
Invalid:
    r.Status = phase: r.Decision = VATS2_FULL_VALIDATION_REQUIRED
    r.ErrorReason = "Reuse必要输入或一致性契约无效。": VATStage2EvaluateReuseEligibility = r
End Function

Private Function SameBatch(ByRef seal As VATS2VerificationSealState, ByRef baseline As VATS2BaselineState) As Boolean
    SameBatch = (seal.ARecordCount = baseline.ARecordCount And seal.BRecordCount = baseline.BRecordCount And _
        StrComp(seal.ABatchDigest, baseline.ABatchDigest, vbBinaryCompare) = 0 And _
        StrComp(seal.BBatchDigest, baseline.BBatchDigest, vbBinaryCompare) = 0 And _
        StrComp(seal.CombinedDigest, baseline.CombinedDigest, vbBinaryCompare) = 0)
End Function

Private Sub ValidateBaselineSummary(ByRef baseline As VATS2BaselineState)
    Need baseline.Status = VATS2_BASELINE_OK And baseline.SchemaVersion = "VATBASE1"
    Need baseline.FingerprintProtocol = "VATFP1" And baseline.DigestProtocol = "SHA256-UTF8"
    Need baseline.ARecordCount >= 0 And baseline.BRecordCount >= 0
    Need IsDigest(baseline.ABatchDigest) And IsDigest(baseline.BBatchDigest) And IsDigest(baseline.CombinedDigest)
End Sub

Private Sub CheckDeltaSide(ByVal count As Long, ByVal unchanged As Long, ByVal changed As Long, ByVal residual As Long, ByVal oldCount As Long)
    Need count >= 0 And unchanged >= 0 And changed >= 0 And residual >= 0
    Need CDbl(unchanged) + changed = count
    Need CDbl(unchanged) + residual = oldCount
End Sub

Private Sub CheckUnchanged(ByRef current() As VATS2DeltaCurrent, ByVal count As Long, ByRef oldDigests() As String)
    Dim budget As Object, i As Long, key As String, allocated As Boolean, bound As Long
    If count = 0 Then
        On Error Resume Next
        Err.Clear: bound = LBound(current): allocated = (Err.Number = 0)
        Err.Clear: bound = LBound(oldDigests): allocated = allocated Or (Err.Number = 0)
        On Error GoTo 0
        Need Not allocated: Exit Sub
    End If
    Need LBound(current) = 1 And UBound(current) = count
    Need LBound(oldDigests) = 1 And UBound(oldDigests) = count
    Set budget = CreateObject("Scripting.Dictionary"): budget.CompareMode = vbBinaryCompare
    For i = 1 To count
        key = oldDigests(i): Need IsDigest(key)
        If i > 1 Then Need StrComp(oldDigests(i - 1), key, vbBinaryCompare) <= 0
        If Not budget.Exists(key) Then budget.Add key, 0&
        budget(key) = CLng(budget(key)) + 1
    Next i
    For i = 1 To count
        Need current(i).CurrentIndex = i And current(i).State = VATS2_DELTA_UNCHANGED
        key = current(i).Digest: Need budget.Exists(key)
        Need CLng(budget(key)) > 0
        budget(key) = CLng(budget(key)) - 1
    Next i
End Sub
