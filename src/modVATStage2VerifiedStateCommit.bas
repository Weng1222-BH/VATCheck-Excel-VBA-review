Attribute VB_Name = "modVATStage2VerifiedStateCommit"
Option Explicit
Option Compare Binary

Public Enum VATS2VerifiedStateCommitStatus
    VATS2_COMMIT_OK = 0
    VATS2_COMMIT_INVALID_INPUT = 1
    VATS2_COMMIT_BASELINE_FAILED = 2
    VATS2_COMMIT_SEAL_FAILED_PARTIAL = 3
    VATS2_COMMIT_INVALID_CONTRACT = 4
End Enum
Public Const VATS2_COMMIT_NOT_ATTEMPTED As Long = -1

Public Type VATS2BoundBaselineResult
    Status As VATS2VerifiedStateCommitStatus
    Baseline As VATS2BaselineState
    ErrorStage As String
    ErrorReason As String
End Type

Public Type VATS2VerifiedStateCommitResult
    Status As VATS2VerifiedStateCommitStatus
    '这两个字段只记录Save结果；-1表示尚未调用对应存储入口。
    BaselineStatus As Long
    SealStatus As Long
    BaselineWritten As Boolean
    SealWritten As Boolean
    ErrorStage As String
    ErrorReason As String
End Type

'纯内存绑定检查，不授予审核结论，不写文件，也不接受裸FinalDecision。
Public Function VATStage2BuildBaselineForBoundRun(ByRef run As VATS2BoundAuditRunResult, _
    ByRef fp As VATS2FingerprintResult) As VATS2BoundBaselineResult
    Dim r As VATS2BoundBaselineResult, blank As VATS2BoundBaselineResult
    Dim baseline As VATS2BaselineState, phase As VATS2VerifiedStateCommitStatus, stage As String
    phase = VATS2_COMMIT_INVALID_INPUT: stage = "INPUT"
    On Error GoTo Invalid
    Need run.Status = VATS2_BOUND_RUN_OK And fp.Status = VATS2_FINGERPRINT_OK
    phase = VATS2_COMMIT_INVALID_CONTRACT: stage = "BASELINE_BUILD"
    baseline = VATStage2BuildBaseline(fp)
    Need baseline.Status = VATS2_BASELINE_OK
    stage = "BASELINE_BINDING"
    Need run.BindingProtocol = "VATRUN1"
    Need run.FingerprintProtocol = baseline.FingerprintProtocol And run.DigestProtocol = baseline.DigestProtocol
    Need run.ARecordCount = baseline.ARecordCount And run.BRecordCount = baseline.BRecordCount
    Need StrComp(run.ABatchDigest, baseline.ABatchDigest, vbBinaryCompare) = 0
    Need StrComp(run.BBatchDigest, baseline.BBatchDigest, vbBinaryCompare) = 0
    Need StrComp(run.CombinedDigest, baseline.CombinedDigest, vbBinaryCompare) = 0
    r.Baseline = baseline: VATStage2BuildBaselineForBoundRun = r: Exit Function
Invalid:
    blank.Status = phase: blank.Baseline.Status = VATS2_BASELINE_INVALID_INPUT
    blank.ErrorStage = stage: blank.ErrorReason = "当前Fingerprint与BoundRun输入或绑定契约不满足。"
    VATStage2BuildBaselineForBoundRun = blank
End Function

'只协调两次冻结Save调用；不伪造双文件原子事务，不回滚或恢复旧文件。
Public Function VATStage2CommitVerifiedState(ByRef run As VATS2BoundAuditRunResult, _
    ByRef fp As VATS2FingerprintResult, Optional ByVal rootOverride As String = "") As VATS2VerifiedStateCommitResult
    Dim r As VATS2VerifiedStateCommitResult, bound As VATS2BoundBaselineResult
    Dim seal As VATS2VerificationSealState, baselineIO As VATS2BaselineIOResult, sealIO As VATS2VerificationSealIOResult
    r.Status = VATS2_COMMIT_INVALID_INPUT
    r.BaselineStatus = VATS2_COMMIT_NOT_ATTEMPTED: r.SealStatus = VATS2_COMMIT_NOT_ATTEMPTED
    r.ErrorStage = "INPUT"
    On Error GoTo Unexpected
    If run.Status <> VATS2_BOUND_RUN_OK Or run.FinalDecision.Status <> VATS2_FINAL_DECISION_OK Or _
        run.FinalDecision.Verdict <> VATS2_FINAL_VERIFIED Then
        r.ErrorReason = "只有受控完整审核VERIFIED结果可提交。": GoTo Done
    End If
    bound = VATStage2BuildBaselineForBoundRun(run, fp)
    If bound.Status <> VATS2_COMMIT_OK Then
        r.Status = bound.Status: r.ErrorStage = bound.ErrorStage: r.ErrorReason = bound.ErrorReason: GoTo Done
    End If
    r.ErrorStage = "SEAL_BUILD"
    seal = VATStage2BuildVerificationSeal(run)
    If seal.Status <> VATS2_VERIFY_OK Then
        r.Status = VATS2_COMMIT_INVALID_CONTRACT
        r.ErrorReason = "BoundRun未通过冻结Seal构建契约。": GoTo Done
    End If
    r.ErrorStage = "BASELINE_SAVE"
    baselineIO = VATStage2SaveBaseline(bound.Baseline, rootOverride)
    r.BaselineStatus = baselineIO.Status
    If baselineIO.Status <> VATS2_BASELINE_OK Then
        r.Status = VATS2_COMMIT_BASELINE_FAILED: r.ErrorReason = baselineIO.ErrorReason: GoTo Done
    End If
    r.BaselineWritten = True
    r.ErrorStage = "SEAL_SAVE"
    sealIO = VATStage2SaveVerificationSeal(seal, rootOverride)
    r.SealStatus = sealIO.Status
    If sealIO.Status <> VATS2_VERIFY_OK Then
        r.Status = VATS2_COMMIT_SEAL_FAILED_PARTIAL: r.ErrorReason = sealIO.ErrorReason: GoTo Done
    End If
    r.SealWritten = True: r.Status = VATS2_COMMIT_OK: r.ErrorStage = "": r.ErrorReason = ""
Done:
    VATStage2CommitVerifiedState = r: Exit Function
Unexpected:
    '冻结Save自行捕获I/O错误；意外错误仍保留已经完成的写入事实，不调用补救写入。
    If r.BaselineWritten Then
        r.Status = VATS2_COMMIT_SEAL_FAILED_PARTIAL
    ElseIf r.ErrorStage = "BASELINE_SAVE" Then
        r.Status = VATS2_COMMIT_BASELINE_FAILED
    Else
        r.Status = VATS2_COMMIT_INVALID_CONTRACT
    End If
    r.ErrorReason = "提交阶段出现意外错误，错误码=" & CStr(Err.Number)
    Resume Done
End Function

Private Sub Need(ByVal condition As Boolean)
    If Not condition Then Err.Raise vbObjectError + 761, , "Bound baseline契约不满足。"
End Sub
