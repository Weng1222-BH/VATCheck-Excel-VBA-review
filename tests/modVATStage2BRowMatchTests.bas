Attribute VB_Name = "modVATStage2BRowMatchTests"
Option Explicit
Option Compare Binary

Private caseCount As Long, failureCount As Long
Private details As String

Public Function VATStage2BRowMatch_SelfTest() As String
    On Error GoTo Unexpected
    caseCount = 0: failureCount = 0: details = vbNullString
    '期望逐引用：Digits:MatchKind:MatcherFlags:A索引列表；各引用以分号分隔。
    'MatchKind 冻结值：0未找到，1/2 exact唯一/多个，3/4 suffix唯一/多个，6无效A范围。
    '期望计数依次为 Unique、NotFound、Multiple、Invalid。
    CheckCase "单引用 exact", "NO.11111111", "11111111", "11111111:1:0:1", "1,0,0,0", VATS2_BROW_ALL_UNIQUE, 0, 0
    CheckCase "单引用 suffix", "NO.22222222", "99999922222222", "22222222:3:0:1", "1,0,0,0", VATS2_BROW_ALL_UNIQUE, 0, 0
    CheckCase "两个 exact 唯一", "NO.11111111/22222222", "11111111|22222222", "11111111:1:0:1;22222222:1:0:2", "2,0,0,0", VATS2_BROW_ALL_UNIQUE, 0, 0
    CheckCase "exact 与 suffix 混合", "NO.11111111/22222222", "11111111|99999922222222", "11111111:1:0:1;22222222:3:0:2", "2,0,0,0", VATS2_BROW_ALL_UNIQUE, 0, 0
    CheckCase "部分未找到保留成功结果", "NO.11111111/33333333", "11111111|22222222", "11111111:1:0:1;33333333:0:0:", "1,1,0,0", VATS2_BROW_INCOMPLETE, 0, 0
    CheckCase "部分 multiple 保留成功结果", "NO.11111111/123456", "11111111|111111123456|222222123456", "11111111:1:0:1;123456:4:1:2,3", "1,0,1,0", VATS2_BROW_INCOMPLETE, 0, 1
    CheckCase "多个未找到", "NO.11111111/22222222", "33333333", "11111111:0:0:;22222222:0:0:", "0,2,0,0", VATS2_BROW_INCOMPLETE, 0, 0
    CheckCase "multiple 所有候选顺序", "NO.123456", "99123456|55555555|88123456|77123456", "123456:4:1:1,3,4", "0,0,1,0", VATS2_BROW_INCOMPLETE, 0, 1
    CheckCase "空原文零引用", "", "11111111", "", "0,0,0,0", VATS2_BROW_NO_REFERENCES, 0, 0
    CheckCase "普通文字零引用", "供应商名称（折让）", "11111111", "", "0,0,0,0", VATS2_BROW_NO_REFERENCES, 0, 0
    CheckCase "零引用保留结构风险", "NO.11111111 NO.22222222", "11111111", "", "0,0,0,0", VATS2_BROW_NO_REFERENCES, VATS2_STRUCTURE_REVIEW, 0
    CheckCase "无 NO 仍可唯一", "供应商ABC12345678", "12345678", "12345678:1:0:1", "1,0,0,0", VATS2_BROW_ALL_UNIQUE, VATS2_NO_MARKER_MISSING, 0
    CheckCase "短尾号仍可唯一", "NO.123456", "999999123456", "123456:3:1:1", "1,0,0,0", VATS2_BROW_ALL_UNIQUE, 0, 1
    CheckCase "缺 NO 与短尾号风险分开", "供应商ABC123456", "999999123456", "123456:3:1:1", "1,0,0,0", VATS2_BROW_ALL_UNIQUE, VATS2_NO_MARKER_MISSING, 1
    CheckCase "空分隔风险不否定唯一", "NO.11111111//22222222", "11111111|22222222", "11111111:1:0:1;22222222:1:0:2", "2,0,0,0", VATS2_BROW_ALL_UNIQUE, VATS2_STRUCTURE_REVIEW, 0
    CheckCase "纯数字尾注不新增引用", "NO.11111111--12345", "11111111|12345", "11111111:1:0:1", "1,0,0,0", VATS2_BROW_ALL_UNIQUE, VATS2_NUMERIC_TRAILER_REVIEW, 0
    CheckCase "重复引用只传播风险", "NO.11111111/11111111", "11111111", "11111111:1:0:1", "1,0,0,0", VATS2_BROW_ALL_UNIQUE, VATS2_DUPLICATE_REF_REVIEW, 0
    CheckCase "exact 优先于两个 suffix", "NO.12345678", "12345678|11112345678|22212345678", "12345678:1:0:1", "1,0,0,0", VATS2_BROW_ALL_UNIQUE, 0, 0
    CheckCase "多个 exact 不混入 suffix", "NO.12345678", "9912345678|12345678|12345678", "12345678:2:0:2,3", "0,0,1,0", VATS2_BROW_INCOMPLETE, 0, 0
    CheckCase "混合分隔引用顺序与 A 顺序独立", "采购（NO.）33333333，11111111、22222222", "11111111|22222222|33333333", "33333333:1:0:3;11111111:1:0:1;22222222:1:0:2", "3,0,0,0", VATS2_BROW_ALL_UNIQUE, 0, 0
    CheckCase "零下界 A", "NO.12345678", "9912345678|8812345678", "12345678:4:0:0,1", "0,0,1,0", VATS2_BROW_INCOMPLETE, 0, 0, 0
    CheckCase "负下界 A", "NO.11111111/22222222", "11111111|22222222", "11111111:1:0:-2;22222222:1:0:-1", "2,0,0,0", VATS2_BROW_ALL_UNIQUE, 0, 0, -2
    CheckCase "前导零保留", "NO.00123456/123456", "00123456|123456", "00123456:1:0:1;123456:1:0:2", "2,0,0,0", VATS2_BROW_ALL_UNIQUE, 0, 0
    CheckCase "空 A 所有引用未找到", "NO.11111111/123456", "", "11111111:0:0:;123456:0:0:", "0,2,0,0", VATS2_BROW_INCOMPLETE, 0, 0
    CheckCase "无效 A 范围不伪装未找到", "NO.11111111/22222222", "11111111", "11111111:6:0:;22222222:6:0:", "0,0,0,2", VATS2_BROW_INVALID_INPUT, 0, 0, 1, -1
    CheckCase "未分配 A 非零计数", "NO.11111111", "", "11111111:6:0:", "0,0,0,1", VATS2_BROW_INVALID_INPUT, 0, 0, 1, 1
    CheckCase "A 数量越界", "NO.11111111", "11111111", "11111111:6:0:", "0,0,0,1", VATS2_BROW_INVALID_INPUT, 0, 0, 1, 2
    CheckCase "零引用不调用 Matcher 检查 A", "无号码", "", "", "0,0,0,0", VATS2_BROW_NO_REFERENCES, 0, 0, 1, -1
    CheckCase "只读有效 A 前缀范围", "NO.11111111/22222222", "11111111|22222222", "11111111:1:0:1;22222222:0:0:", "1,1,0,0", VATS2_BROW_INCOMPLETE, 0, 0, 1, 1
    CheckCase "四种 Parser 风险加 Matcher 风险", "供应商BC123456//123456--123", "99123456", "123456:3:1:1", "1,0,0,0", VATS2_BROW_ALL_UNIQUE, 15, 1
    CheckCase "短号未找到仍传播 Matcher 风险", "NO.12/11111111", "11111111", "12:0:1:;11111111:1:0:1", "1,1,0,0", VATS2_BROW_INCOMPLETE, 0, 1
    CheckCase "前缀括号原文定位完整保留", "前置（0000）采购（NO.）BC00123456/ZZ22222222（折让）", "00123456|22222222", "00123456:1:0:1;22222222:1:0:2", "2,0,0,0", VATS2_BROW_ALL_UNIQUE, 0, 0
    CheckCase "英文数字尾注完全排除", "NO.11111111--lo5", "11111111|5", "11111111:1:0:1", "1,0,0,0", VATS2_BROW_ALL_UNIQUE, 0, 0
    CheckCase "唯一多个未找到同时保留", "NO.11111111/22222222/33333333", "11111111|22222222|22222222", "11111111:1:0:1;22222222:2:0:2,3;33333333:0:0:", "1,1,1,0", VATS2_BROW_INCOMPLETE, 0, 0
    If failureCount = 0 Then
        VATStage2BRowMatch_SelfTest = "PASS: " & caseCount & " cases" & vbCrLf & details
    Else
        VATStage2BRowMatch_SelfTest = "FAIL: " & failureCount & "/" & caseCount & " cases" & vbCrLf & details
    End If
    Exit Function
Unexpected:
    VATStage2BRowMatch_SelfTest = "FAIL: unexpected " & Err.Number & " " & Err.Description & vbCrLf & details
End Function

Private Sub CheckCase(ByVal title As String, ByVal rawText As String, ByVal aSource As String, _
    ByVal expectedRefs As String, ByVal expectedCounts As String, ByVal expectedState As VATS2BRowMatchState, _
    ByVal expectedParserFlags As Long, ByVal expectedMatcherFlags As Long, _
    Optional ByVal lower As Long = 1, Optional ByVal validCount As Long = -999)
    Dim values() As String, parts As Variant, result As VATS2BRowMatchResult, again As VATS2BRowMatchResult
    Dim n As Long, i As Long, expectedCount As Long, reason As String
    On Error GoTo Failed
    caseCount = caseCount + 1
    If Len(aSource) > 0 Then
        parts = Split(aSource, "|"): n = UBound(parts) + 1
        ReDim values(lower To lower + n - 1)
        For i = 0 To n - 1
            values(lower + i) = parts(i)
        Next i
    End If
    If validCount = -999 Then validCount = n
    If Len(expectedRefs) > 0 Then expectedCount = UBound(Split(expectedRefs, ";")) + 1
    result = VATStage2MatchBRow(rawText, values, validCount)
    again = VATStage2MatchBRow(rawText, values, validCount)
    VerifyResult result, rawText, values, validCount, expectedRefs, expectedCounts, expectedState, expectedParserFlags, expectedMatcherFlags, expectedCount
    VerifyResult again, rawText, values, validCount, expectedRefs, expectedCounts, expectedState, expectedParserFlags, expectedMatcherFlags, expectedCount
    For i = 0 To n - 1
        Require values(lower + i) = parts(i), "A 输入未修改"
    Next i
    details = details & "PASS " & title & vbCrLf
    Exit Sub
Failed:
    reason = Err.Description: failureCount = failureCount + 1
    details = details & "FAIL " & title & ": " & reason & vbCrLf
End Sub

Private Sub VerifyResult(ByRef result As VATS2BRowMatchResult, ByVal rawText As String, _
    ByRef values() As String, ByVal validCount As Long, ByVal expectedRefs As String, _
    ByVal expectedCounts As String, ByVal expectedState As VATS2BRowMatchState, _
    ByVal expectedParserFlags As Long, ByVal expectedMatcherFlags As Long, ByVal expectedCount As Long)
    Dim parsed As VATS2ParseResult, matched As VATS2ReferenceMatchResult
    Dim i As Long, j As Long, text As String, counts As String
    Require result.OriginalText = rawText, "B 原文完整保留"
    Require result.ReferenceCount = expectedCount, "引用数量"
    Require result.MatchState = expectedState, "行状态"
    Require result.ParserFlags = expectedParserFlags, "ParserFlags 独立空间"
    Require result.MatcherFlags = expectedMatcherFlags, "MatcherFlags OR 汇总"
    Require result.ReferenceCount = result.UniqueMatchCount + result.NotFoundCount + result.MultipleMatchCount + result.InvalidMatchCount, "四类计数守恒"
    counts = result.UniqueMatchCount & "," & result.NotFoundCount & "," & result.MultipleMatchCount & "," & result.InvalidMatchCount
    Require counts = expectedCounts, "四类计数精确值"
    parsed = VATStage2ParseBInvoices(rawText)
    Require result.ReferenceCount = parsed.ReferenceCount, "不新增或丢弃 Parser 引用"
    If result.ReferenceCount = 0 Then
        Require Not HasReferences(result), "零引用数组未分配"
    Else
        Require LBound(result.References) = 1, "引用数组下界"
        Require UBound(result.References) = result.ReferenceCount, "引用数组容量"
    End If
    For i = 1 To result.ReferenceCount
        matched = VATStage2MatchReference(parsed.References(i).Digits, values, validCount)
        With result.References(i)
            '对照冻结模块的完整输出，确保集成层没有截断上下文、位置或风险。
            Require .Digits = parsed.References(i).Digits, "Digits 及 Parser 原顺序"
            Require .Length = parsed.References(i).Length, "Length"
            Require .RawFragment = parsed.References(i).RawFragment, "RawFragment"
            Require .StartIndex = parsed.References(i).StartIndex, "StartIndex"
            Require .Context = parsed.References(i).Context, "Context"
            Require .ParserFlags = parsed.References(i).Flags, "逐引用 ParserFlags"
            Require Mid$(rawText, .StartIndex, Len(.RawFragment)) = .RawFragment, "原文位置"
            Require .MatchKind = matched.MatchKind, "逐引用 MatchKind"
            Require .MatcherFlags = matched.Flags, "逐引用 MatcherFlags"
            Require .CandidateCount = matched.CandidateCount, "CandidateCount"
            If .CandidateCount > 0 Then
                Require LBound(.Candidates) = 1, "候选下界"
                Require UBound(.Candidates) = .CandidateCount, "候选容量"
            Else
                Require Not HasCandidates(result.References(i)), "零候选数组未分配"
            End If
            If i > 1 Then text = text & ";"
            text = text & .Digits & ":" & .MatchKind & ":" & .MatcherFlags & ":"
            For j = 1 To .CandidateCount
                Require .Candidates(j).AIndex = matched.Candidates(j).AIndex, "全部候选原索引和顺序"
                Require .Candidates(j).InvoiceDigits = matched.Candidates(j).InvoiceDigits, "全部候选完整号码"
                Require .Candidates(j).InvoiceDigits = values(.Candidates(j).AIndex), "候选与输入对应"
                If j > 1 Then text = text & ","
                text = text & .Candidates(j).AIndex
            Next j
        End With
    Next i
    Require text = expectedRefs, "独立期望：逐引用类型、顺序、候选索引及风险"
End Sub

Private Function HasReferences(ByRef result As VATS2BRowMatchResult) As Boolean
    Dim upper As Long
    On Error GoTo EmptyArray
    upper = UBound(result.References): HasReferences = True
EmptyArray:
End Function

Private Function HasCandidates(ByRef reference As VATS2BRowReference) As Boolean
    Dim upper As Long
    On Error GoTo EmptyArray
    upper = UBound(reference.Candidates): HasCandidates = True
EmptyArray:
End Function

Private Sub Require(ByVal condition As Boolean, ByVal reason As String)
    If Not condition Then Err.Raise vbObjectError + 2301, "BRowMatchTests", reason
End Sub
