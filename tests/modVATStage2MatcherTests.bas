Attribute VB_Name = "modVATStage2MatcherTests"
Option Explicit
Option Compare Binary

Private caseCount As Long
Private failureCount As Long
Private details As String

Public Function VATStage2Matcher_SelfTest() As String
    On Error GoTo Unexpected
    caseCount = 0: failureCount = 0: details = vbNullString
    CheckCase "exact 唯一", "12345678", "88888888,12345678", VATS2_EXACT_UNIQUE, "2", 0
    CheckCase "exact 优先于两个 suffix", "12345678", "12345678,11112345678,22212345678", VATS2_EXACT_UNIQUE, "1", 0
    CheckCase "exact 多个", "12345678", "12345678,12345678", VATS2_EXACT_MULTIPLE, "1,2", 0
    CheckCase "九位 suffix", "151515151", "26578450000151515151", VATS2_SUFFIX_UNIQUE, "1", 0
    CheckCase "六位 suffix", "123456", "111111123456", VATS2_SUFFIX_UNIQUE, "1", VATS2_SHORT_SUFFIX_REVIEW
    CheckCase "五位 suffix", "12345", "11111112345", VATS2_SUFFIX_UNIQUE, "1", VATS2_SHORT_SUFFIX_REVIEW
    CheckCase "suffix 多个", "123456", "111111123456,222222123456", VATS2_SUFFIX_MULTIPLE, "1,2", VATS2_SHORT_SUFFIX_REVIEW
    CheckCase "无匹配", "12345678", "88888888,99999999", VATS2_NOT_FOUND, "", 0
    CheckCase "B 比 A 长", "00123456", "123456", VATS2_NOT_FOUND, "", 0
    CheckCase "前导零 exact", "00123456", "00123456,123456", VATS2_EXACT_UNIQUE, "1", 0
    CheckCase "前导零 suffix", "00123456", "9900123456,99123456", VATS2_SUFFIX_UNIQUE, "1", 0
    CheckCase "候选保留 A 顺序", "1234567", "991234567,88888888,111234567,771234567", VATS2_SUFFIX_MULTIPLE, "1,3,4", 0
    CheckCase "六位 exact 无短号风险", "123456", "99123456,123456", VATS2_EXACT_UNIQUE, "2", 0
    CheckCase "一位 exact 无短号风险", "0", "10,0,100", VATS2_EXACT_UNIQUE, "2", 0
    CheckCase "一位 suffix 不设门槛", "0", "10,22,30", VATS2_SUFFIX_MULTIPLE, "1,3", VATS2_SHORT_SUFFIX_REVIEW
    CheckCase "多个 exact 排除 suffix", "12345678", "9912345678,12345678,8812345678,12345678", VATS2_EXACT_MULTIPLE, "2,4", 0
    CheckCase "未分配空 A", "12345678", "", VATS2_NOT_FOUND, "", 0
    CheckCase "已分配但有效数为零", "123456", "123456", VATS2_NOT_FOUND, "", 0, 1, 0
    CheckCase "只检查 aCount 指定部分", "12345678", "9912345678,12345678", VATS2_SUFFIX_UNIQUE, "1", 0, 1, 1
    CheckCase "零下界保留索引", "12345678", "9912345678,8812345678", VATS2_SUFFIX_MULTIPLE, "0,1", 0, 0
    CheckCase "任意下界保留索引", "12345678", "99999999,12345678,12345678", VATS2_EXACT_MULTIPLE, "6,7", 0, 5
    CheckCase "负下界保留索引", "12345678", "9912345678,8812345678", VATS2_SUFFIX_MULTIPLE, "-2,-1", 0, -2
    CheckCase "重复 A suffix 不去重", "12345678", "9912345678,9912345678", VATS2_SUFFIX_MULTIPLE, "1,2", 0
    CheckCase "七位 suffix 无短号风险", "1234567", "991234567", VATS2_SUFFIX_UNIQUE, "1", 0
    CheckCase "短号未找到仍记录 suffix 路径", "12", "333,444", VATS2_NOT_FOUND, "", VATS2_SHORT_SUFFIX_REVIEW
    CheckCase "空 A 不执行 suffix", "12", "", VATS2_NOT_FOUND, "", 0
    CheckCase "空引用无效", "", "12345678", VATS2_INVALID_REFERENCE, "", 0
    CheckCase "不 Trim", " 12345678", "12345678", VATS2_INVALID_REFERENCE, "", 0
    CheckCase "不解析字母前缀", "BC12345678", "12345678", VATS2_INVALID_REFERENCE, "", 0
    CheckCase "全角数字无效", "１２３４５６７８", "12345678", VATS2_INVALID_REFERENCE, "", 0
    CheckCase "小数无效", "123.45", "12345", VATS2_INVALID_REFERENCE, "", 0
    CheckCase "分隔符无效", "123/456", "123456", VATS2_INVALID_REFERENCE, "", 0
    CheckCase "尾换行无效", "12345678" & vbLf, "12345678", VATS2_INVALID_REFERENCE, "", 0
    CheckCase "负 aCount", "12345678", "12345678", VATS2_INVALID_A_RANGE, "", 0, 1, -1
    CheckCase "aCount 超过容量", "12345678", "12345678", VATS2_INVALID_A_RANGE, "", 0, 1, 2
    CheckCase "未分配 A 却声称有元素", "12345678", "", VATS2_INVALID_A_RANGE, "", 0, 1, 1
    CheckCase "长字符串不转数值", "0000123456789012345678901234567890", "0000123456789012345678901234567890,99123456789012345678901234567890", VATS2_EXACT_UNIQUE, "1", 0
    CheckCase "长字符串 suffix", "0000123456789012345678901234567890", "990000123456789012345678901234567890", VATS2_SUFFIX_UNIQUE, "1", 0
    If failureCount = 0 Then
        VATStage2Matcher_SelfTest = "PASS: " & caseCount & " cases" & vbCrLf & details
    Else
        VATStage2Matcher_SelfTest = "FAIL: " & failureCount & "/" & caseCount & " cases" & vbCrLf & details
    End If
    Exit Function
Unexpected:
    VATStage2Matcher_SelfTest = "FAIL: unexpected " & Err.Number & " " & Err.Description & vbCrLf & details
End Function

Private Sub CheckCase(ByVal title As String, ByVal referenceDigits As String, ByVal aSource As String, _
    ByVal expectedKind As VATS2MatchKind, ByVal expectedIndexes As String, ByVal expectedFlags As Long, _
    Optional ByVal lower As Long = 1, Optional ByVal validCount As Long = -999)
    Dim values() As String, parts As Variant, indexes As Variant
    Dim result As VATS2ReferenceMatchResult, again As VATS2ReferenceMatchResult
    Dim n As Long, expectedCount As Long, i As Long, aIndex As Long, reason As String
    On Error GoTo Failed
    caseCount = caseCount + 1
    If Len(aSource) > 0 Then
        parts = Split(aSource, ","): n = UBound(parts) + 1
        ReDim values(lower To lower + n - 1)
        For i = 0 To n - 1
            values(lower + i) = parts(i)
        Next i
    End If
    If validCount = -999 Then validCount = n
    If Len(expectedIndexes) > 0 Then
        indexes = Split(expectedIndexes, ","): expectedCount = UBound(indexes) + 1
    End If
    result = VATStage2MatchReference(referenceDigits, values, validCount)
    '每个用例都重复调用，核对确定性、输入不变和全部返回字段。
    again = VATStage2MatchReference(referenceDigits, values, validCount)
    Require result.ReferenceDigits = referenceDigits, "原文"
    Require result.ReferenceLength = Len(referenceDigits), "引用长度"
    Require result.MatchKind = expectedKind, "MatchKind"
    Require result.Flags = expectedFlags, "Flags"
    Require result.CandidateCount = expectedCount, "候选数量"
    Require again.ReferenceDigits = result.ReferenceDigits, "重复调用原文"
    Require again.ReferenceLength = result.ReferenceLength, "重复调用长度"
    Require again.MatchKind = result.MatchKind, "重复调用类型"
    Require again.Flags = result.Flags, "重复调用 Flags"
    Require again.CandidateCount = result.CandidateCount, "重复调用数量"
    If expectedCount > 0 Then
        Require LBound(result.Candidates) = 1, "候选数组下界"
        Require UBound(result.Candidates) = expectedCount, "候选数组上界"
        Require UBound(again.Candidates) = expectedCount, "重复调用数组上界"
    Else
        Require Not HasCandidates(result), "零候选数组未分配"
        Require Not HasCandidates(again), "重复调用零候选数组未分配"
    End If
    For i = 1 To expectedCount
        aIndex = CLng(indexes(i - 1))
        Require result.Candidates(i).AIndex = aIndex, "原 A 索引及顺序"
        Require result.Candidates(i).InvoiceDigits = values(aIndex), "候选号码"
        Require again.Candidates(i).AIndex = aIndex, "重复调用索引"
        Require again.Candidates(i).InvoiceDigits = values(aIndex), "重复调用号码"
    Next i
    For i = 0 To n - 1
        Require values(lower + i) = parts(i), "A 输入未修改"
    Next i
    details = details & "PASS " & title & vbCrLf
    Exit Sub
Failed:
    reason = Err.Description
    failureCount = failureCount + 1
    details = details & "FAIL " & title & ": " & reason & vbCrLf
End Sub

Private Sub Require(ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then Err.Raise vbObjectError + 2101, "MatcherTests", reason
End Sub

Private Function HasCandidates(ByRef result As VATS2ReferenceMatchResult) As Boolean
    Dim upper As Long
    On Error GoTo EmptyArray
    upper = UBound(result.Candidates)
    HasCandidates = True
EmptyArray:
End Function
