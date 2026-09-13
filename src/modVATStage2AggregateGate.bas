Attribute VB_Name = "modVATStage2AggregateGate"
Option Explicit

Public Enum VATS2AggregateGateStatus
    VATS2_AGGREGATE_OK = 0
    VATS2_CALCULATION_PENDING = 1
    VATS2_SCAN_BLOCKED = 2
    VATS2_AGGREGATE_ARITHMETIC_ERROR = 3
End Enum

Public Type VATS2AggregateGateResult
    Status As VATS2AggregateGateStatus
    CalculationState As XlCalculationState
    AScanSucceeded As Boolean
    BScanSucceeded As Boolean
    TotalA As Variant
    TotalB As Variant
    Difference As Variant
    TotalsEqual As Variant
    ReliableEqualForShortSuffix As Boolean
    ACompletedCount As Long
    BCompletedCount As Long
    AIncludedCount As Long
    BIncludedCount As Long
    AEmptyWarningCount As Long
    BEmptyWarningCount As Long
    AIssueCount As Long
    AIssues() As String
    BIssueCount As Long
    BIssues() As String
    AWarningCount As Long
    AWarnings() As String
    BWarningCount As Long
    BWarnings() As String
    ErrorReason As String
End Type

Public Function VATStage2EvaluateAggregateGate(ByVal headerA As Range, ByVal headerB As Range, _
    ByVal completedColor As Long) As VATS2AggregateGateResult
    Dim r As VATS2AggregateGateResult, totalA As Variant, totalB As Variant, difference As Variant
    Dim issuesA As New Collection, issuesB As New Collection
    Dim warningsA As New Collection, warningsB As New Collection, subtracting As Boolean
    r.Status = VATS2_SCAN_BLOCKED
    On Error GoTo Failed
    r.CalculationState = Application.CalculationState
    If r.CalculationState <> xlDone Then
        r.Status = VATS2_CALCULATION_PENDING
        GoTo Done
    End If
    '两边分别保留原始证据；一侧失败也读取另一侧，不据此输出总体结论。
    r.AScanSucceeded = VATScan(headerA, VAT_FIELD_A, "A", completedColor, totalA, _
        r.ACompletedCount, issuesA, r.AIncludedCount, r.AEmptyWarningCount, warningsA)
    r.BScanSucceeded = VATScan(headerB, VAT_FIELD_B, "B", completedColor, totalB, _
        r.BCompletedCount, issuesB, r.BIncludedCount, r.BEmptyWarningCount, warningsB)
    CopyMessages issuesA, r.AIssueCount, r.AIssues
    CopyMessages issuesB, r.BIssueCount, r.BIssues
    CopyMessages warningsA, r.AWarningCount, r.AWarnings
    CopyMessages warningsB, r.BWarningCount, r.BWarnings
    If Not r.AScanSucceeded Or Not r.BScanSucceeded Or r.AIssueCount > 0 Or r.BIssueCount > 0 Then GoTo Done
    '扫描结束仍须处于已完成计算状态；不触发计算、不改变计算策略。
    r.CalculationState = Application.CalculationState
    If r.CalculationState <> xlDone Then
        r.Status = VATS2_CALCULATION_PENDING
        GoTo Done
    End If
    subtracting = True
    difference = CDec(totalA) - CDec(totalB)
    '只有双侧可靠且差额成功时，才发布总体金额和Boolean一致性结果。
    r.TotalA = totalA: r.TotalB = totalB: r.Difference = difference
    r.TotalsEqual = (difference = CDec(0))
    r.Status = VATS2_AGGREGATE_OK
    r.ReliableEqualForShortSuffix = (r.AWarningCount = 0 And r.BWarningCount = 0 And r.TotalsEqual)
Done:
    VATStage2EvaluateAggregateGate = r
    Exit Function
Failed:
    If subtracting Then r.Status = VATS2_AGGREGATE_ARITHMETIC_ERROR Else r.Status = VATS2_SCAN_BLOCKED
    r.ErrorReason = Err.Number & " " & Err.Description
    r.TotalA = Empty: r.TotalB = Empty: r.Difference = Empty: r.TotalsEqual = Empty
    r.ReliableEqualForShortSuffix = False
    Resume Done
End Function

Private Sub CopyMessages(ByVal source As Collection, ByRef count As Long, ByRef target() As String)
    Dim i As Long
    count = source.Count
    If count > 0 Then ReDim target(1 To count)
    For i = 1 To count
        '只复制VATScan生成的文本，不解析、重写或合并两侧的问题。
        target(i) = source(i)
    Next i
End Sub
