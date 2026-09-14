Attribute VB_Name = "modVATStage2BaselineStore"
Option Explicit
Option Compare Binary

Public Enum VATS2BaselineStatus
    VATS2_BASELINE_OK = 0
    VATS2_BASELINE_NOT_FOUND = 1
    VATS2_BASELINE_INVALID_INPUT = 2
    VATS2_BASELINE_CORRUPT = 3
    VATS2_BASELINE_IO_ERROR = 4
    VATS2_BASELINE_UNSUPPORTED_VERSION = 5
End Enum
Public Type VATS2BaselineState
    Status As VATS2BaselineStatus
    SchemaVersion As String
    FingerprintProtocol As String
    DigestProtocol As String
    ARecordCount As Long
    BRecordCount As Long
    ARecordDigests() As String
    BRecordDigests() As String
    ABatchDigest As String
    BBatchDigest As String
    CombinedDigest As String
    CreatedUtc As String
    UpdatedUtc As String
End Type
Public Type VATS2BaselineIOResult
    Status As VATS2BaselineStatus
    ErrorReason As String
End Type
Public Type VATS2BaselineLoadResult
    Status As VATS2BaselineStatus
    Baseline As VATS2BaselineState
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
Private Const BASE_FILE As String = "baseline-v1.dat"
Private Const TEMP_FILE As String = "baseline-v1.pending.tmp"

Public Function VATStage2BuildBaseline(ByRef fp As VATS2FingerprintResult) As VATS2BaselineState
    Dim r As VATS2BaselineState, blank As VATS2BaselineState, i As Long
    On Error GoTo Invalid
    Need fp.Status = VATS2_FINGERPRINT_OK
    CheckArray fp.AFingerprints, fp.ARecordCount, False
    CheckArray fp.BFingerprints, fp.BRecordCount, False
    Need Len(fp.ABatchFingerprint) > 0 And Len(fp.BBatchFingerprint) > 0 And Len(fp.CombinedFingerprint) > 0
    r.SchemaVersion = "VATBASE1": r.FingerprintProtocol = "VATFP1": r.DigestProtocol = "SHA256-UTF8"
    r.ARecordCount = fp.ARecordCount: r.BRecordCount = fp.BRecordCount
    If r.ARecordCount > 0 Then ReDim r.ARecordDigests(1 To r.ARecordCount)
    If r.BRecordCount > 0 Then ReDim r.BRecordDigests(1 To r.BRecordCount)
    For i = 1 To r.ARecordCount: r.ARecordDigests(i) = VATStage2SHA256(fp.AFingerprints(i)): Next i
    For i = 1 To r.BRecordCount: r.BRecordDigests(i) = VATStage2SHA256(fp.BFingerprints(i)): Next i
    SortDigests r.ARecordDigests, r.ARecordCount: SortDigests r.BRecordDigests, r.BRecordCount
    r.ABatchDigest = VATStage2SHA256(fp.ABatchFingerprint)
    r.BBatchDigest = VATStage2SHA256(fp.BBatchFingerprint)
    r.CombinedDigest = VATStage2SHA256(fp.CombinedFingerprint)
    r.CreatedUtc = UtcNow(): r.UpdatedUtc = r.CreatedUtc
    VATStage2BuildBaseline = r: Exit Function
Invalid:
    blank.Status = VATS2_BASELINE_INVALID_INPUT: VATStage2BuildBaseline = blank
End Function

Public Function VATStage2SaveBaseline(ByRef baseline As VATS2BaselineState, _
    Optional ByVal rootOverride As String = "") As VATS2BaselineIOResult
    Dim r As VATS2BaselineIOResult, check As VATS2BaselineLoadResult
    Dim fs As Object, stream As Object, root As String, target As String, temporary As String
    Dim payload As String, ownTemp As Boolean, phase As VATS2BaselineStatus, errorCode As Long
    On Error GoTo Failed
    phase = VATS2_BASELINE_INVALID_INPUT
    ValidateState baseline
    payload = Serialize(baseline)
    Set fs = CreateObject("Scripting.FileSystemObject")
    root = StoreRoot(rootOverride, fs): target = fs.BuildPath(root, BASE_FILE): temporary = fs.BuildPath(root, TEMP_FILE)
    phase = VATS2_BASELINE_IO_ERROR
    EnsureFolder fs, root
    '固定暂存名通过创建且禁止覆盖实现单写者互斥；已有暂存文件绝不视为本次所有。
    Set stream = fs.CreateTextFile(temporary, False, False): ownTemp = True
    '此处payload仅含ASCII协议、digest和UTC元数据，ANSI输出与UTF-8无BOM字节相同。
    stream.Write payload: stream.Close: Set stream = Nothing
    check = ReadBaselineFile(temporary)
    If check.Status <> VATS2_BASELINE_OK Then Err.Raise vbObjectError + 722
    Need Serialize(check.Baseline) = payload
    '同目录同卷rename替换，不先删除旧文件，不允许跨卷copy/delete降级。
    If MoveFileExW(StrPtr(temporary), StrPtr(target), 9) = 0 Then
        errorCode = Err.LastDllError: Err.Raise vbObjectError + 723
    End If
    ownTemp = False: r.Status = VATS2_BASELINE_OK
    VATStage2SaveBaseline = r: Exit Function
Failed:
    If errorCode = 0 Then errorCode = Err.Number
    r.Status = phase: r.ErrorReason = "Baseline保存失败，错误码=" & CStr(errorCode)
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If ownTemp Then fs.DeleteFile temporary, False
    On Error GoTo 0
    VATStage2SaveBaseline = r
End Function

Public Function VATStage2LoadBaseline(Optional ByVal rootOverride As String = "") As VATS2BaselineLoadResult
    Dim r As VATS2BaselineLoadResult, fs As Object, root As String
    On Error GoTo Failed
    Set fs = CreateObject("Scripting.FileSystemObject")
    root = StoreRoot(rootOverride, fs)
    r = ReadBaselineFile(fs.BuildPath(root, BASE_FILE))
    VATStage2LoadBaseline = r: Exit Function
Failed:
    r.Status = VATS2_BASELINE_IO_ERROR: r.Baseline.Status = r.Status
    r.ErrorReason = "Baseline读取路径失败，错误码=" & CStr(Err.Number)
    VATStage2LoadBaseline = r
End Function

'VATBASE1固定顺序LF文本；末行是此前完整ASCII正文的SHA-256校验和，不是签名/身份认证。
Private Function Serialize(ByRef r As VATS2BaselineState) As String
    Dim s As String, i As Long
    s = r.SchemaVersion & vbLf & r.FingerprintProtocol & vbLf & r.DigestProtocol & vbLf & _
        CStr(r.ARecordCount) & vbLf & CStr(r.BRecordCount) & vbLf & r.CreatedUtc & vbLf & r.UpdatedUtc & vbLf & _
        r.ABatchDigest & vbLf & r.BBatchDigest & vbLf & r.CombinedDigest & vbLf
    For i = 1 To r.ARecordCount: s = s & r.ARecordDigests(i) & vbLf: Next i
    For i = 1 To r.BRecordCount: s = s & r.BRecordDigests(i) & vbLf: Next i
    Serialize = s & VATStage2SHA256(s) & vbLf
End Function

Private Function ReadBaselineFile(ByVal path As String) As VATS2BaselineLoadResult
    Dim r As VATS2BaselineLoadResult, f As Integer, opened As Boolean, s As String, fs As Object
    On Error GoTo Failed
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FileExists(path) Then
        '同名目录不是正常首次运行。
        If fs.FolderExists(path) Then Err.Raise vbObjectError + 724
        r.Status = VATS2_BASELINE_NOT_FOUND: r.Baseline.Status = r.Status
        ReadBaselineFile = r: Exit Function
    End If
    f = FreeFile: Open path For Binary Access Read Lock Write As #f: opened = True
    s = String$(LOF(f), vbNullChar)
    If Len(s) > 0 Then Get #f, , s
    Close #f: opened = False
    ReadBaselineFile = Parse(s): Exit Function
Failed:
    r.Status = VATS2_BASELINE_IO_ERROR: r.Baseline.Status = r.Status
    r.ErrorReason = "Baseline读取失败，错误码=" & CStr(Err.Number)
    On Error Resume Next
    If opened Then Close #f
    On Error GoTo 0
    ReadBaselineFile = r
End Function

Private Function Parse(ByVal s As String) As VATS2BaselineLoadResult
    Dim r As VATS2BaselineLoadResult, blank As VATS2BaselineLoadResult, lines As Variant
    Dim i As Long, p As Long, count As Double, phase As VATS2BaselineStatus
    phase = VATS2_BASELINE_CORRUPT
    On Error GoTo Invalid
    lines = Split(s, vbLf): Need UBound(lines) >= 2
    If lines(0) <> "VATBASE1" Then
        If Left$(lines(0), 7) = "VATBASE" Then phase = VATS2_BASELINE_UNSUPPORTED_VERSION
        GoTo Invalid
    End If
    If lines(1) <> "VATFP1" Or lines(2) <> "SHA256-UTF8" Then phase = VATS2_BASELINE_UNSUPPORTED_VERSION: GoTo Invalid
    Need UBound(lines) >= 11
    With r.Baseline
        .SchemaVersion = lines(0): .FingerprintProtocol = lines(1): .DigestProtocol = lines(2)
        .ARecordCount = ParseCount(CStr(lines(3))): .BRecordCount = ParseCount(CStr(lines(4)))
        count = CDbl(.ARecordCount) + .BRecordCount
        Need count + 11 = UBound(lines): Need lines(UBound(lines)) = ""
        .CreatedUtc = lines(5): .UpdatedUtc = lines(6)
        .ABatchDigest = lines(7): .BBatchDigest = lines(8): .CombinedDigest = lines(9)
        p = 10
        If .ARecordCount > 0 Then ReDim .ARecordDigests(1 To .ARecordCount)
        If .BRecordCount > 0 Then ReDim .BRecordDigests(1 To .BRecordCount)
        For i = 1 To .ARecordCount: .ARecordDigests(i) = lines(p): p = p + 1: Next i
        For i = 1 To .BRecordCount: .BRecordDigests(i) = lines(p): p = p + 1: Next i
    End With
    ValidateState r.Baseline
    Need IsDigest(CStr(lines(p)))
    Need VATStage2SHA256(Left$(s, Len(s) - 65)) = lines(p)
    r.Status = VATS2_BASELINE_OK: Parse = r: Exit Function
Invalid:
    blank.Status = phase: blank.Baseline.Status = phase
    blank.ErrorReason = "Baseline格式、版本或完整性校验失败。"
    Parse = blank
End Function

Private Sub ValidateState(ByRef r As VATS2BaselineState)
    Need r.Status = VATS2_BASELINE_OK
    Need r.SchemaVersion = "VATBASE1" And r.FingerprintProtocol = "VATFP1" And r.DigestProtocol = "SHA256-UTF8"
    CheckArray r.ARecordDigests, r.ARecordCount, True: CheckArray r.BRecordDigests, r.BRecordCount, True
    Need IsDigest(r.ABatchDigest) And IsDigest(r.BBatchDigest) And IsDigest(r.CombinedDigest)
    Need IsUtc(r.CreatedUtc) And IsUtc(r.UpdatedUtc)
    Need StrComp(r.CreatedUtc, r.UpdatedUtc, vbBinaryCompare) <= 0
End Sub
Private Sub CheckArray(ByRef values() As String, ByVal count As Long, ByVal digests As Boolean)
    Dim i As Long
    Need count >= 0
    If count = 0 Then
        Need Not Allocated(values): Exit Sub
    End If
    Need LBound(values) = 1 And UBound(values) = count
    For i = 1 To count
        If digests Then
            Need IsDigest(values(i))
            If i > 1 Then Need StrComp(values(i - 1), values(i), vbBinaryCompare) <= 0
        Else
            Need Len(values(i)) > 0
        End If
    Next i
End Sub
Private Function Allocated(ByRef values() As String) As Boolean
    Dim n As Long
    On Error GoTo EmptyArray
    n = LBound(values): Allocated = True
EmptyArray:
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
Private Sub SortDigests(ByRef values() As String, ByVal count As Long)
    Dim scratch() As String
    If count < 2 Then Exit Sub
    ReDim scratch(1 To count): MergeSort values, scratch, 1, count
End Sub
Private Sub MergeSort(ByRef values() As String, ByRef scratch() As String, ByVal lo As Long, ByVal hi As Long)
    Dim mid As Long, i As Long, j As Long, p As Long
    If lo >= hi Then Exit Sub
    mid = lo + (hi - lo) \ 2
    MergeSort values, scratch, lo, mid: MergeSort values, scratch, mid + 1, hi
    i = lo: j = mid + 1
    For p = lo To hi
        If i > mid Then
            scratch(p) = values(j): j = j + 1
        ElseIf j > hi Then
            scratch(p) = values(i): i = i + 1
        ElseIf StrComp(values(i), values(j), vbBinaryCompare) <= 0 Then
            scratch(p) = values(i): i = i + 1
        Else
            scratch(p) = values(j): j = j + 1
        End If
    Next p
    For p = lo To hi: values(p) = scratch(p): Next p
End Sub
Private Sub Need(ByVal condition As Boolean)
    If Not condition Then Err.Raise vbObjectError + 721, , "Baseline必要契约不满足。"
End Sub
