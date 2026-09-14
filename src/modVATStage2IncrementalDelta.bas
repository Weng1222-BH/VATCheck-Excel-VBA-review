Attribute VB_Name = "modVATStage2IncrementalDelta"
Option Explicit
Option Compare Binary

Public Enum VATS2IncrementalDeltaStatus
    VATS2_DELTA_OK = 0
    VATS2_DELTA_FIRST_RUN = 1
    VATS2_DELTA_BASELINE_UNAVAILABLE = 2
    VATS2_DELTA_INVALID_INPUT = 3
    VATS2_DELTA_INVALID_CONTRACT = 4
End Enum
Public Enum VATS2DeltaRecordState
    VATS2_DELTA_UNCHANGED = 0
    VATS2_DELTA_NEW_OR_CHANGED = 1
End Enum
Public Type VATS2DeltaCurrent
    CurrentIndex As Long
    Digest As String
    State As VATS2DeltaRecordState
End Type
Public Type VATS2DeltaResidual
    Digest As String
    OccurrenceCount As Long
End Type
Public Type VATS2IncrementalDeltaResult
    Status As VATS2IncrementalDeltaStatus
    BaselineSourceStatus As VATS2BaselineStatus
    ARecordCount As Long
    BRecordCount As Long
    ACurrent() As VATS2DeltaCurrent
    BCurrent() As VATS2DeltaCurrent
    AUnchangedCount As Long
    BUnchangedCount As Long
    ANewOrChangedCount As Long
    BNewOrChangedCount As Long
    AResidualCount As Long
    BResidualCount As Long
    AResidualGroupCount As Long
    BResidualGroupCount As Long
    AResiduals() As VATS2DeltaResidual
    BResiduals() As VATS2DeltaResidual
    ABatchUnchanged As Boolean
    BBatchUnchanged As Boolean
    CombinedUnchanged As Boolean
    ErrorSide As String
    ErrorIndex As Long
    ErrorReason As String
End Type

'只比较摘要多重集，不执行文件读取、保存、业务匹配、freeze或skip。
Public Function VATStage2CompareBaselineDelta(ByRef fp As VATS2FingerprintResult, _
    ByRef baselineLoad As VATS2BaselineLoadResult) As VATS2IncrementalDeltaResult
    Dim r As VATS2IncrementalDeltaResult, blank As VATS2IncrementalDeltaResult
    Dim current As VATS2BaselineState, firstRun As Boolean, sameA As Boolean, sameB As Boolean
    Dim phase As VATS2IncrementalDeltaStatus
    r.BaselineSourceStatus = baselineLoad.Status
    phase = VATS2_DELTA_INVALID_INPUT
    On Error GoTo Invalid
    Need fp.Status = VATS2_FINGERPRINT_OK, "当前Fingerprint非OK。"
    Select Case baselineLoad.Status
        Case VATS2_BASELINE_OK, VATS2_BASELINE_NOT_FOUND, VATS2_BASELINE_CORRUPT, VATS2_BASELINE_IO_ERROR, VATS2_BASELINE_UNSUPPORTED_VERSION
        Case Else: Need False, "Baseline来源状态无效。"
    End Select
    phase = VATS2_DELTA_INVALID_CONTRACT
    r.ErrorSide = "A": CheckArray fp.AFingerprints, fp.ARecordCount, False, r.ErrorIndex
    r.ErrorSide = "B": CheckArray fp.BFingerprints, fp.BRecordCount, False, r.ErrorIndex
    r.ErrorSide = "": r.ErrorIndex = 0
    Need Len(fp.ABatchFingerprint) > 0 And Len(fp.BBatchFingerprint) > 0 And Len(fp.CombinedFingerprint) > 0, "当前批次签名不能为空。"
    If baselineLoad.Status <> VATS2_BASELINE_OK And baselineLoad.Status <> VATS2_BASELINE_NOT_FOUND Then
        blank.Status = VATS2_DELTA_BASELINE_UNAVAILABLE: blank.BaselineSourceStatus = baselineLoad.Status
        blank.ErrorReason = baselineLoad.ErrorReason
        If Len(blank.ErrorReason) = 0 Then blank.ErrorReason = "Baseline不可用，调用方需要完整验证。"
        VATStage2CompareBaselineDelta = blank: Exit Function
    End If
    firstRun = (baselineLoad.Status = VATS2_BASELINE_NOT_FOUND)
    If Not firstRun Then
        With baselineLoad.Baseline
            Need .Status = VATS2_BASELINE_OK, "Baseline内外状态不一致。"
            Need .SchemaVersion = "VATBASE1" And .FingerprintProtocol = "VATFP1" And .DigestProtocol = "SHA256-UTF8", "Baseline版本契约无效。"
            r.ErrorSide = "A": CheckArray .ARecordDigests, .ARecordCount, True, r.ErrorIndex
            r.ErrorSide = "B": CheckArray .BRecordDigests, .BRecordCount, True, r.ErrorIndex
            r.ErrorSide = "": r.ErrorIndex = 0
            Need IsDigest(.ABatchDigest) And IsDigest(.BBatchDigest) And IsDigest(.CombinedDigest), "Baseline批次摘要无效。"
        End With
    End If
    r.ARecordCount = fp.ARecordCount: r.BRecordCount = fp.BRecordCount
    r.ErrorSide = "A"
    CompareSide fp.AFingerprints, fp.ARecordCount, baselineLoad.Baseline.ARecordDigests, _
        baselineLoad.Baseline.ARecordCount, firstRun, r.ACurrent, r.AUnchangedCount, r.ANewOrChangedCount, _
        r.AResiduals, r.AResidualCount, r.AResidualGroupCount, r.ErrorIndex
    r.ErrorSide = "B"
    CompareSide fp.BFingerprints, fp.BRecordCount, baselineLoad.Baseline.BRecordDigests, _
        baselineLoad.Baseline.BRecordCount, firstRun, r.BCurrent, r.BUnchangedCount, r.BNewOrChangedCount, _
        r.BResiduals, r.BResidualCount, r.BResidualGroupCount, r.ErrorIndex
    r.ErrorSide = "": r.ErrorIndex = 0
    '复用冻结构建入口获取批次摘要；结果只留在内存，不保存baseline。
    current = VATStage2BuildBaseline(fp)
    Need current.Status = VATS2_BASELINE_OK, "当前摘要构建失败。"
    If firstRun Then
        r.Status = VATS2_DELTA_FIRST_RUN
    Else
        r.ABatchUnchanged = (StrComp(current.ABatchDigest, baselineLoad.Baseline.ABatchDigest, vbBinaryCompare) = 0)
        r.BBatchUnchanged = (StrComp(current.BBatchDigest, baselineLoad.Baseline.BBatchDigest, vbBinaryCompare) = 0)
        r.CombinedUnchanged = (StrComp(current.CombinedDigest, baselineLoad.Baseline.CombinedDigest, vbBinaryCompare) = 0)
        sameA = (r.ANewOrChangedCount = 0 And r.AResidualCount = 0)
        sameB = (r.BNewOrChangedCount = 0 And r.BResidualCount = 0)
        r.ErrorSide = "A": Need r.ABatchUnchanged = sameA, "A批次摘要与多重集矛盾。"
        r.ErrorSide = "B": Need r.BBatchUnchanged = sameB, "B批次摘要与多重集矛盾。"
        r.ErrorSide = "": Need r.CombinedUnchanged = (sameA And sameB), "Combined摘要与多重集矛盾。"
        r.Status = VATS2_DELTA_OK
    End If
    r.ErrorSide = "": VATStage2CompareBaselineDelta = r: Exit Function
Invalid:
    blank.Status = phase: blank.BaselineSourceStatus = baselineLoad.Status
    blank.ErrorSide = r.ErrorSide: blank.ErrorIndex = r.ErrorIndex: blank.ErrorReason = Err.Description
    VATStage2CompareBaselineDelta = blank
End Function

Private Sub CompareSide(ByRef raw() As String, ByVal count As Long, ByRef oldDigests() As String, _
    ByVal oldCount As Long, ByVal firstRun As Boolean, ByRef records() As VATS2DeltaCurrent, _
    ByRef unchanged As Long, ByRef newOrChanged As Long, ByRef residuals() As VATS2DeltaResidual, _
    ByRef residualCount As Long, ByRef groupCount As Long, ByRef errorIndex As Long)
    Dim budget() As Long, i As Long, start As Long, pos As Long, digest As String
    '每个相同摘要区间的首元素保存剩余次数；只减副本，不改旧baseline。
    If Not firstRun And oldCount > 0 Then
        ReDim budget(1 To oldCount): start = 1
        For i = 1 To oldCount
            If i > 1 Then
                If StrComp(oldDigests(i), oldDigests(i - 1), vbBinaryCompare) <> 0 Then start = i
            End If
            budget(start) = budget(start) + 1
        Next i
    End If
    If count > 0 Then ReDim records(1 To count)
    For i = 1 To count
        errorIndex = i: digest = VATStage2SHA256(raw(i))
        records(i).CurrentIndex = i: records(i).Digest = digest: records(i).State = VATS2_DELTA_NEW_OR_CHANGED
        pos = 0
        If Not firstRun Then pos = FindFirst(oldDigests, oldCount, digest)
        If pos > 0 Then
            If budget(pos) > 0 Then
                budget(pos) = budget(pos) - 1: records(i).State = VATS2_DELTA_UNCHANGED
            End If
        End If
        If records(i).State = VATS2_DELTA_UNCHANGED Then unchanged = unchanged + 1 Else newOrChanged = newOrChanged + 1
    Next i
    errorIndex = 0
    If firstRun Or oldCount = 0 Then Exit Sub
    For i = 1 To oldCount
        If budget(i) > 0 Then groupCount = groupCount + 1: residualCount = residualCount + budget(i)
    Next i
    If groupCount = 0 Then Exit Sub
    ReDim residuals(1 To groupCount): pos = 0
    For i = 1 To oldCount
        If budget(i) > 0 Then
            pos = pos + 1: residuals(pos).Digest = oldDigests(i): residuals(pos).OccurrenceCount = budget(i)
        End If
    Next i
End Sub
Private Function FindFirst(ByRef values() As String, ByVal count As Long, ByVal digest As String) As Long
    Dim lo As Long, hi As Long, mid As Long
    lo = 1: hi = count
    Do While lo <= hi
        mid = lo + (hi - lo) \ 2
        If StrComp(values(mid), digest, vbBinaryCompare) < 0 Then
            lo = mid + 1
        Else
            If values(mid) = digest Then FindFirst = mid
            hi = mid - 1
        End If
    Loop
End Function
Private Sub CheckArray(ByRef values() As String, ByVal count As Long, ByVal digests As Boolean, ByRef errorIndex As Long)
    Dim i As Long
    errorIndex = 0: Need count >= 0, "记录数量不得为负。"
    If count = 0 Then Need Not Allocated(values), "零记录数组必须未分配。": Exit Sub
    Need LBound(values) = 1 And UBound(values) = count, "数组必须为1..RecordCount。"
    For i = 1 To count
        errorIndex = i
        If digests Then
            Need IsDigest(values(i)), "记录摘要必须为64位小写hex。"
            If i > 1 Then Need StrComp(values(i - 1), values(i), vbBinaryCompare) <= 0, "Baseline摘要数组必须已排序。"
        Else
            Need Len(values(i)) > 0, "当前record签名不能为空。"
        End If
    Next i
    errorIndex = 0
End Sub
Private Function Allocated(ByRef values() As String) As Boolean
    Dim n As Long
    On Error GoTo EmptyArray
    n = LBound(values): Allocated = True
EmptyArray:
End Function
Private Function IsDigest(ByVal value As String) As Boolean
    Dim i As Long
    If Len(value) <> 64 Then Exit Function
    For i = 1 To 64
        If InStr(1, "0123456789abcdef", Mid$(value, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next i
    IsDigest = True
End Function
Private Sub Need(ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then Err.Raise vbObjectError + 730, , reason
End Sub
