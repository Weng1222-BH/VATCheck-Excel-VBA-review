Attribute VB_Name = "modVATStage2BoundAuditRun"
Option Explicit
Option Compare Binary

Public Enum VATS2BoundRunStatus
    VATS2_BOUND_RUN_OK = 0
    VATS2_BOUND_RUN_INVALID_INPUT = 1
    VATS2_BOUND_RUN_STAGE_FAILED = 2
    VATS2_BOUND_RUN_SOURCE_CHANGED = 3
    VATS2_BOUND_RUN_INVALID_CONTRACT = 4
End Enum
Public Type VATS2BoundAuditRunResult
    Status As VATS2BoundRunStatus
    BindingProtocol As String
    FingerprintProtocol As String
    DigestProtocol As String
    ARecordCount As Long
    BRecordCount As Long
    ABatchDigest As String
    BBatchDigest As String
    CombinedDigest As String
    FinalDecision As VATS2FinalDecisionResult
    AuditFindings As VATS2AuditFindingsResult
    Aggregate As VATS2AggregateGateResult
    ErrorStage As String
    ErrorReason As String
End Type

'唯一绑定入口：所有审核结果和指纹均在本次调用中产生，不接受外部审核对象。
Public Function VATStage2RunBoundFullAudit(ByVal aInvoiceHeader As Range, ByVal aAmountHeader As Range, _
    ByVal bSupplierHeader As Range, ByVal bAmountHeader As Range, ByVal completedColor As Long) As VATS2BoundAuditRunResult
    Dim r As VATS2BoundAuditRunResult, blank As VATS2BoundAuditRunResult, phase As VATS2BoundRunStatus
    Dim wa As Worksheet, wb As Worksheet, ar As Long, br As Long, ai As Long, aa As Long, bi As Long, ba As Long
    Dim a As VATS2ASnapshotResult, b As VATS2BSnapshotResult, confirmA As VATS2ASnapshotResult, confirmB As VATS2BSnapshotResult
    Dim fp As VATS2FingerprintResult, confirmFP As VATS2FingerprintResult, digest As VATS2BaselineState
    Dim aggregate As VATS2AggregateGateResult, scope As VATS2CompletedScopeResult, fallback As VATS2FullFallbackResult
    Dim effective As VATS2EffectiveRelationsResult, amounts As VATS2EffectiveAmountResult
    Dim shortPolicy As VATS2ShortSuffixPolicyResult, findings As VATS2AuditFindingsResult, decision As VATS2FinalDecisionResult
    On Error GoTo Failed
    phase = VATS2_BOUND_RUN_INVALID_INPUT: r.ErrorStage = "INPUT"
    CheckHeaders aInvoiceHeader, aAmountHeader, VAT_FIELD_A
    CheckHeaders bSupplierHeader, bAmountHeader, VAT_FIELD_B
    Set wa = aInvoiceHeader.Parent: Set wb = bSupplierHeader.Parent
    ar = aInvoiceHeader.Row: ai = aInvoiceHeader.Column: aa = aAmountHeader.Column
    br = bSupplierHeader.Row: bi = bSupplierHeader.Column: ba = bAmountHeader.Column
    phase = VATS2_BOUND_RUN_STAGE_FAILED
    r.ErrorStage = "SNAPSHOT_A": a = VATStage2ReadASnapshot(wa, ar, ai, aa, completedColor)
    Need a.Status = VATS2_SNAPSHOT_OK, a.ErrorReason
    r.ErrorStage = "SNAPSHOT_B": b = VATStage2ReadBSnapshot(wb, br, bi, ba, completedColor)
    Need b.Status = VATS2_SNAPSHOT_OK, b.ErrorReason
    r.ErrorStage = "FINGERPRINT": fp = VATStage2BuildFingerprints(a, b)
    Need fp.Status = VATS2_FINGERPRINT_OK, fp.ErrorReason
    r.ErrorStage = "AGGREGATE"
    aggregate = VATStage2EvaluateAggregateGate(aAmountHeader, bAmountHeader, completedColor)
    phase = VATS2_BOUND_RUN_INVALID_CONTRACT
    Need aggregate.Status >= VATS2_AGGREGATE_OK And aggregate.Status <= VATS2_AGGREGATE_ARITHMETIC_ERROR, "未知Aggregate状态。"
    'Aggregate非OK仍交给冻结FinalDecision决定UNAVAILABLE，不能冒充通过或业务复核。
    r.ErrorStage = "SOURCE_STABILITY": phase = VATS2_BOUND_RUN_SOURCE_CHANGED
    CheckHeaders aInvoiceHeader, aAmountHeader, VAT_FIELD_A
    CheckHeaders bSupplierHeader, bAmountHeader, VAT_FIELD_B
    Need aInvoiceHeader.Row = ar And aAmountHeader.Column = aa And aInvoiceHeader.Column = ai, "A表头位置在运行期间变化。"
    Need bSupplierHeader.Row = br And bAmountHeader.Column = ba And bSupplierHeader.Column = bi, "B表头位置在运行期间变化。"
    phase = VATS2_BOUND_RUN_STAGE_FAILED
    confirmA = VATStage2ReadASnapshot(wa, ar, ai, aa, completedColor)
    Need confirmA.Status = VATS2_SNAPSHOT_OK, confirmA.ErrorReason
    confirmB = VATStage2ReadBSnapshot(wb, br, bi, ba, completedColor)
    Need confirmB.Status = VATS2_SNAPSHOT_OK, confirmB.ErrorReason
    confirmFP = VATStage2BuildFingerprints(confirmA, confirmB)
    Need confirmFP.Status = VATS2_FINGERPRINT_OK, confirmFP.ErrorReason
    phase = VATS2_BOUND_RUN_SOURCE_CHANGED
    Need VATStage2BoundSourcesStable(a, b, fp, confirmA, confirmB, confirmFP), "前后Snapshot的记录、顺序或ExcelRow发生变化。"
    '后续链只使用第一份已确认稳定的内存快照，不再读取业务单元格。
    phase = VATS2_BOUND_RUN_STAGE_FAILED
    r.ErrorStage = "COMPLETED_SCOPE": scope = VATStage2MatchCompletedScope(a, b)
    Need scope.Status = VATS2_SCOPE_OK, scope.ErrorReason
    r.ErrorStage = "FULL_FALLBACK": fallback = VATStage2RunFullFallback(a, b, scope)
    Need fallback.Status = VATS2_FALLBACK_OK, fallback.ErrorReason
    r.ErrorStage = "EFFECTIVE_RELATIONS": effective = VATStage2BuildEffectiveRelations(a, b, scope, fallback)
    Need effective.Status = VATS2_EFFECTIVE_OK, effective.ErrorReason
    r.ErrorStage = "EFFECTIVE_AMOUNT": amounts = VATStage2EvaluateEffectiveAmounts(a, b, effective)
    Need amounts.Status = VATS2_EFFECTIVE_AMOUNT_OK, amounts.ErrorReason
    r.ErrorStage = "SHORT_SUFFIX"
    shortPolicy = VATStage2EvaluateShortSuffixPolicy(a, b, effective, amounts, aggregate.ReliableEqualForShortSuffix)
    Need shortPolicy.Status = VATS2_SHORT_POLICY_OK, shortPolicy.ErrorReason
    r.ErrorStage = "AUDIT_FINDINGS": findings = VATStage2BuildAuditFindings(a, b, fallback, effective, amounts, shortPolicy)
    Need findings.Status = VATS2_AUDIT_OK, findings.ErrorReason
    r.ErrorStage = "FINAL_DECISION": decision = VATStage2BuildFinalDecision(aggregate, amounts, findings)
    Need decision.Status = VATS2_FINAL_DECISION_OK, decision.ErrorReason
    r.ErrorStage = "BASELINE_DIGEST": digest = VATStage2BuildBaseline(fp)
    Need digest.Status = VATS2_BASELINE_OK, "当前Fingerprint的摘要构建失败。"
    '只有全部成功才一次性发布绑定；不会保存baseline或调用历史结果。
    r.BindingProtocol = "VATRUN1": r.FingerprintProtocol = "VATFP1": r.DigestProtocol = "SHA256-UTF8"
    r.ARecordCount = digest.ARecordCount: r.BRecordCount = digest.BRecordCount
    r.ABatchDigest = digest.ABatchDigest: r.BBatchDigest = digest.BBatchDigest: r.CombinedDigest = digest.CombinedDigest
    r.FinalDecision = decision: r.AuditFindings = findings: r.Aggregate = aggregate
    r.Status = VATS2_BOUND_RUN_OK: r.ErrorStage = ""
    VATStage2RunBoundFullAudit = r: Exit Function
Failed:
    blank.Status = phase: blank.ErrorStage = r.ErrorStage: blank.ErrorReason = Err.Description
    blank.FinalDecision.Status = VATS2_FINAL_INVALID_INPUT: blank.FinalDecision.Verdict = VATS2_FINAL_UNAVAILABLE
    blank.AuditFindings.Status = VATS2_AUDIT_INVALID_INPUT: blank.Aggregate.Status = VATS2_SCAN_BLOCKED
    VATStage2RunBoundFullAudit = blank
End Function

Private Sub CheckHeaders(ByVal identityHeader As Range, ByVal amountHeader As Range, ByVal amountName As String)
    Dim sheet As Worksheet, item As Worksheet, live As Boolean
    Need Not identityHeader Is Nothing, "身份表头未提供。"
    Need Not amountHeader Is Nothing, "金额表头未提供。"
    Need identityHeader.CountLarge = 1 And amountHeader.CountLarge = 1, "表头必须为单个单元格。"
    Need identityHeader.Application Is Application, "表头必须属于当前Excel实例。"
    Need identityHeader.Parent Is amountHeader.Parent, "同一侧两个表头必须属于同一Worksheet。"
    Need identityHeader.Row = amountHeader.Row, "同一侧表头必须处于同一行。"
    Need identityHeader.Column <> amountHeader.Column, "身份与金额必须是不同列。"
    Need Not identityHeader.MergeCells And Not amountHeader.MergeCells, "表头不能为合并单元格。"
    Set sheet = identityHeader.Parent
    Need VATWorkbookOpen(sheet.Parent), "表头工作簿已关闭。"
    For Each item In sheet.Parent.Worksheets
        If item Is sheet Then live = True: Exit For
    Next item
    Need live, "表头Worksheet已失效。"
    Need VarType(identityHeader.Value2) = vbString, "身份表头必须是非空文本。"
    Need Len(CStr(identityHeader.Value2)) > 0, "身份表头不能为空。"
    Need VarType(amountHeader.Value2) = vbString, "金额表头文本无效。"
    Need StrComp(CStr(amountHeader.Value2), amountName, vbBinaryCompare) = 0, "金额表头与冻结字段名称不一致。"
End Sub

'纯内存稳定性辅助入口仅返回Boolean，不能创建BoundRun或注入外部审核结果。
Public Function VATStage2BoundSourcesStable(ByRef a As VATS2ASnapshotResult, ByRef b As VATS2BSnapshotResult, _
    ByRef fp As VATS2FingerprintResult, ByRef confirmA As VATS2ASnapshotResult, ByRef confirmB As VATS2BSnapshotResult, _
    ByRef confirmFP As VATS2FingerprintResult) As Boolean
    Dim i As Long
    On Error GoTo Unstable
    If a.Status <> VATS2_SNAPSHOT_OK Or b.Status <> VATS2_SNAPSHOT_OK Or confirmA.Status <> VATS2_SNAPSHOT_OK Or confirmB.Status <> VATS2_SNAPSHOT_OK Then Exit Function
    If fp.Status <> VATS2_FINGERPRINT_OK Or confirmFP.Status <> VATS2_FINGERPRINT_OK Then Exit Function
    If a.RecordCount < 0 Or b.RecordCount < 0 Then Exit Function
    If a.RecordCount <> confirmA.RecordCount Or b.RecordCount <> confirmB.RecordCount Then Exit Function
    If fp.ARecordCount <> a.RecordCount Or fp.BRecordCount <> b.RecordCount Or confirmFP.ARecordCount <> a.RecordCount Or confirmFP.BRecordCount <> b.RecordCount Then Exit Function
    If a.RecordCount > 0 Then
        If LBound(a.Records) <> 1 Or UBound(a.Records) <> a.RecordCount Or LBound(confirmA.Records) <> 1 Or UBound(confirmA.Records) <> a.RecordCount Then Exit Function
        If LBound(fp.AFingerprints) <> 1 Or UBound(fp.AFingerprints) <> a.RecordCount Or LBound(confirmFP.AFingerprints) <> 1 Or UBound(confirmFP.AFingerprints) <> a.RecordCount Then Exit Function
    End If
    If b.RecordCount > 0 Then
        If LBound(b.Records) <> 1 Or UBound(b.Records) <> b.RecordCount Or LBound(confirmB.Records) <> 1 Or UBound(confirmB.Records) <> b.RecordCount Then Exit Function
        If LBound(fp.BFingerprints) <> 1 Or UBound(fp.BFingerprints) <> b.RecordCount Or LBound(confirmFP.BFingerprints) <> 1 Or UBound(confirmFP.BFingerprints) <> b.RecordCount Then Exit Function
    End If
    For i = 1 To a.RecordCount
        If a.Records(i).ExcelRow <> confirmA.Records(i).ExcelRow Then Exit Function
        If Len(fp.AFingerprints(i)) = 0 Or StrComp(fp.AFingerprints(i), confirmFP.AFingerprints(i), vbBinaryCompare) <> 0 Then Exit Function
    Next i
    For i = 1 To b.RecordCount
        If b.Records(i).ExcelRow <> confirmB.Records(i).ExcelRow Then Exit Function
        If Len(fp.BFingerprints(i)) = 0 Or StrComp(fp.BFingerprints(i), confirmFP.BFingerprints(i), vbBinaryCompare) <> 0 Then Exit Function
    Next i
    VATStage2BoundSourcesStable = True
Unstable:
End Function
Private Sub Need(ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then
        If Len(reason) = 0 Then reason = "受控审核阶段未成功。"
        Err.Raise vbObjectError + 740, , reason
    End If
End Sub
