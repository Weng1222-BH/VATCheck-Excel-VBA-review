Attribute VB_Name = "modVATStage2FinalDecision"
Option Explicit

Public Enum VATS2FinalDecisionStatus
    VATS2_FINAL_DECISION_OK = 0
    VATS2_FINAL_INVALID_INPUT = 1
    VATS2_FINAL_INVALID_CONTRACT = 2
End Enum

Public Enum VATS2FinalVerdict
    VATS2_FINAL_UNAVAILABLE = 0
    VATS2_FINAL_VERIFIED = 1
    VATS2_FINAL_REVIEW_REQUIRED = 2
End Enum

Public Type VATS2FinalDecisionResult
    Status As VATS2FinalDecisionStatus
    Verdict As VATS2FinalVerdict
    AggregateStatus As VATS2AggregateGateStatus
    EffectiveAmountStatus As VATS2EffectiveAmountStatus
    AuditStatus As VATS2AuditStatus
    TotalA As Variant
    TotalB As Variant
    Difference As Variant
    TotalsEqual As Variant
    ACompletedCount As Long
    BCompletedCount As Long
    AIncludedCount As Long
    BIncludedCount As Long
    AWarningCount As Long
    BWarningCount As Long
    EffectiveBCount As Long
    ComparedCount As Long
    EqualCount As Long
    MismatchCount As Long
    NotComparableCount As Long
    AmountErrorCount As Long
    RelationQualityBlockedCount As Long
    FindingCount As Long
    CodeCounts(1 To 18) As Long
    ShortSuffixWaivedCount As Long
    HasAggregateMismatch As Boolean
    HasAggregateWarnings As Boolean
    HasFindings As Boolean
    NoCompletedRecords As Boolean
    RequiresManualReview As Boolean
    ErrorSource As String
    ErrorReason As String
End Type

Public Function VATStage2BuildFinalDecision(ByRef aggregate As VATS2AggregateGateResult, _
    ByRef amounts As VATS2EffectiveAmountResult, ByRef findings As VATS2AuditFindingsResult) As VATS2FinalDecisionResult
    Dim r As VATS2FinalDecisionResult, blank As VATS2FinalDecisionResult
    Dim i As Long, countSum As Double
    r.Verdict = VATS2_FINAL_UNAVAILABLE
    r.AggregateStatus = aggregate.Status: r.EffectiveAmountStatus = amounts.Status: r.AuditStatus = findings.Status
    r.Status = VATS2_FINAL_INVALID_CONTRACT
    On Error GoTo Invalid
    r.ErrorSource = "AGGREGATE"
    Need aggregate.Status >= VATS2_AGGREGATE_OK And aggregate.Status <= VATS2_AGGREGATE_ARITHMETIC_ERROR, "未知Aggregate状态。"
    r.ErrorSource = "EFFECTIVE_AMOUNT"
    Need amounts.Status >= VATS2_EFFECTIVE_AMOUNT_OK And amounts.Status <= VATS2_EFFECTIVE_AMOUNT_GROUP_ERROR, "未知EffectiveAmount状态。"
    r.ErrorSource = "FINDINGS"
    Need findings.Status >= VATS2_AUDIT_OK And findings.Status <= VATS2_AUDIT_INVALID_CONTRACT, "未知AuditFindings状态。"
    '失败的上游不能被解释为零异常；与合法的总体门UNAVAILABLE分开标识。
    r.Status = VATS2_FINAL_INVALID_INPUT
    If amounts.Status <> VATS2_EFFECTIVE_AMOUNT_OK Then
        r.ErrorSource = "EFFECTIVE_AMOUNT": r.ErrorReason = "EffectiveAmount非OK。": GoTo Done
    End If
    If findings.Status <> VATS2_AUDIT_OK Then
        r.ErrorSource = "FINDINGS": r.ErrorReason = "AuditFindings非OK。": GoTo Done
    End If
    r.Status = VATS2_FINAL_INVALID_CONTRACT
    r.ErrorSource = "AGGREGATE"
    NonNegative aggregate.ACompletedCount, aggregate.BCompletedCount, aggregate.AIncludedCount, aggregate.BIncludedCount, _
        aggregate.AWarningCount, aggregate.BWarningCount, aggregate.AEmptyWarningCount, aggregate.BEmptyWarningCount, aggregate.AIssueCount, aggregate.BIssueCount
    If aggregate.Status = VATS2_AGGREGATE_OK Then Need VarType(aggregate.TotalsEqual) = vbBoolean, "AGGREGATE_OK的TotalsEqual必须是Boolean。"
    r.ErrorSource = "EFFECTIVE_AMOUNT"
    NonNegative amounts.FullARecordCount, amounts.FullBRecordCount, amounts.EffectiveBCount, amounts.ComparedCount, _
        amounts.EqualCount, amounts.MismatchCount, amounts.NotComparableCount, amounts.AmountErrorCount, _
        amounts.RelationQualityBlockedCount, amounts.InvalidGroupCount, amounts.RelationIssueCount
    '这里只对Long计数求和，Double可精确容纳这些计数之和；不读取或重算金额。
    Need CDbl(amounts.ComparedCount) = CDbl(amounts.EqualCount) + amounts.MismatchCount, "ComparedCount不等于EqualCount加MismatchCount。"
    countSum = CDbl(amounts.ComparedCount) + amounts.NotComparableCount + amounts.AmountErrorCount + amounts.RelationQualityBlockedCount + amounts.InvalidGroupCount
    Need countSum = amounts.EffectiveBCount And amounts.EffectiveBCount <= amounts.FullBRecordCount, "EffectiveBCount分类计数或完整B范围错误。"
    Need amounts.InvalidGroupCount = 0, "OK的EffectiveAmount不能含接口错误组。"
    r.ErrorSource = "FINDINGS"
    NonNegative findings.FindingCount, findings.ShortSuffixWaivedCount
    countSum = 0
    For i = 1 To 18
        NonNegative findings.CodeCounts(i)
        countSum = countSum + findings.CodeCounts(i)
    Next i
    Need countSum = findings.FindingCount, "FindingCount与CodeCounts总和不一致。"
    '不访问逐行金额或Findings数组，不按事项类型重新分类。
    r.ACompletedCount = aggregate.ACompletedCount: r.BCompletedCount = aggregate.BCompletedCount
    r.AIncludedCount = aggregate.AIncludedCount: r.BIncludedCount = aggregate.BIncludedCount
    r.AWarningCount = aggregate.AWarningCount: r.BWarningCount = aggregate.BWarningCount
    r.EffectiveBCount = amounts.EffectiveBCount: r.ComparedCount = amounts.ComparedCount
    r.EqualCount = amounts.EqualCount: r.MismatchCount = amounts.MismatchCount
    r.NotComparableCount = amounts.NotComparableCount: r.AmountErrorCount = amounts.AmountErrorCount
    r.RelationQualityBlockedCount = amounts.RelationQualityBlockedCount
    r.FindingCount = findings.FindingCount: r.ShortSuffixWaivedCount = findings.ShortSuffixWaivedCount
    For i = 1 To 18: r.CodeCounts(i) = findings.CodeCounts(i): Next i
    r.HasFindings = (findings.FindingCount > 0)
    r.HasAggregateWarnings = (aggregate.AWarningCount > 0 Or aggregate.BWarningCount > 0)
    r.Status = VATS2_FINAL_DECISION_OK: r.ErrorSource = ""
    '门未成功时即使输入带残留部分金额也不复制，不发布一致/不一致及零记录结论。
    If aggregate.Status <> VATS2_AGGREGATE_OK Then GoTo Done
    r.TotalA = aggregate.TotalA: r.TotalB = aggregate.TotalB: r.Difference = aggregate.Difference
    r.TotalsEqual = aggregate.TotalsEqual
    r.HasAggregateMismatch = Not aggregate.TotalsEqual
    r.NoCompletedRecords = (aggregate.ACompletedCount = 0 And aggregate.BCompletedCount = 0)
    If r.HasAggregateMismatch Or r.HasAggregateWarnings Or r.HasFindings Then
        r.Verdict = VATS2_FINAL_REVIEW_REQUIRED
        r.RequiresManualReview = True
    Else
        r.Verdict = VATS2_FINAL_VERIFIED
    End If
Done:
    VATStage2BuildFinalDecision = r: Exit Function
Invalid:
    blank.Status = VATS2_FINAL_INVALID_CONTRACT: blank.Verdict = VATS2_FINAL_UNAVAILABLE
    blank.AggregateStatus = aggregate.Status: blank.EffectiveAmountStatus = amounts.Status: blank.AuditStatus = findings.Status
    blank.ErrorSource = r.ErrorSource: blank.ErrorReason = Err.Description: r = blank
    Resume Done
End Function

Private Sub Need(ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then Err.Raise 5, , reason
End Sub

Private Sub NonNegative(ParamArray counts() As Variant)
    Dim i As Long
    For i = LBound(counts) To UBound(counts)
        Need counts(i) >= 0, "计数不得为负数。"
    Next i
End Sub
