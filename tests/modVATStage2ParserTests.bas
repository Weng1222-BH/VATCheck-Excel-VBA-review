Attribute VB_Name = "modVATStage2ParserTests"
Option Explicit
Option Compare Binary

Private caseCount As Long
Private logText As String

Private Sub Require(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 710, "C1 Parser", message
End Sub

Private Sub CheckCase(ByVal name As String, ByVal raw As String, ByVal expected As String, ByVal flags As Long)
    Dim parsed As VATS2ParseResult, expectedParts As Variant, expectedCount As Long
    Dim i As Long, j As Long, actual As String, ch As String
    parsed = VATStage2ParseBInvoices(raw)
    expectedCount = 0
    If Len(expected) > 0 Then
        expectedParts = Split(expected, "|")
        expectedCount = UBound(expectedParts) + 1
    End If
    Require StrComp(parsed.OriginalText, raw, vbBinaryCompare) = 0, name & "：原文发生变化"
    Require parsed.ReferenceCount = expectedCount, name & "：候选数量实际=" & parsed.ReferenceCount & "，期望=" & expectedCount
    Require parsed.Flags = flags, name & "：Flags实际=" & parsed.Flags & "，期望=" & flags
    For i = 1 To parsed.ReferenceCount
        With parsed.References(i)
            Require .Digits = CStr(expectedParts(i - 1)), name & "：Digits 或顺序不符"
            Require .Length = Len(.Digits), name & "：长度不符"
            Require .Flags = flags, name & "：引用遗漏记录风险"
            Require .StartIndex >= 1 And Len(.RawFragment) > 0, name & "：原文位置缺失"
            Require Mid$(raw, .StartIndex, Len(.RawFragment)) = .RawFragment, name & "：原片段与位置不符"
            Require InStr(1, .RawFragment, .Digits, vbBinaryCompare) > 0, name & "：数字被错误拼接"
            Require InStr(1, .Context, .RawFragment, vbBinaryCompare) > 0, name & "：上下文缺少原片段"
            For j = 1 To Len(.Digits)
                ch = Mid$(.Digits, j, 1)
                Require ch >= "0" And ch <= "9", name & "：Digits含非ASCII数字"
            Next j
            For j = 1 To i - 1
                Require .Digits <> parsed.References(j).Digits, name & "：重复候选"
                Require .StartIndex > parsed.References(j).StartIndex, name & "：原文顺序被改变"
            Next j
        End With
    Next i
    caseCount = caseCount + 1
    logText = logText & "PASS " & name & vbCrLf
End Sub

Public Function VATStage2Parser_SelfTest() As String
    Dim p As VATS2ParseResult, again As VATS2ParseResult, raw As String
    Dim review As Long, numeric As Long, missing As Long, duplicate As Long
    On Error GoTo Failed
    caseCount = 0: logText = vbNullString
    review = VATS2_STRUCTURE_REVIEW: numeric = VATS2_NUMERIC_TRAILER_REVIEW
    missing = VATS2_NO_MARKER_MISSING: duplicate = VATS2_DUPLICATE_REF_REVIEW
    '每个用例同时验证原文、数量、数字、长度、顺序、去重、定位、上下文和风险。
    CheckCase "真实01：括号NO＋20位", "采购专用发票，广东省AAA有限公司（NO.）26578450000012121212", "26578450000012121212", 0
    CheckCase "真实02：9位", "采购专用发票，南通BBB有限公司（NO.）151515151", "151515151", 0
    CheckCase "真实03：10位", "采购专用发票，湖南省CCC有限公司（NO.）6868686868", "6868686868", 0
    CheckCase "真实04：排除前置0000", "银（0000）付DD申请办公室EEE有限公司NO.26458900000001536415", "26458900000001536415", 0
    CheckCase "真实05：无NO与lo5尾注", "采购专用发票，广东省TTT有限公司265784500000989898--lo5", "265784500000989898", missing
    CheckCase "真实06：排除P括号尾注", "采购专用发票，广东省FFF有限公司265784500000989898(P0101204648221)", "265784500000989898", missing
    CheckCase "真实07：BC前缀", "银（0000）付DD申请办公室EEE有限公司BC26458900000001536415", "26458900000001536415", missing
    CheckCase "真实08：折让括号", "银（0000）付DD申请办公室EEE有限公司BC26458900000001536415（折让）", "26458900000001536415", missing
    CheckCase "真实09：20位与8位", "采购专用发票，广东省AAA有限公司（NO.）26578450000012121212/65456454", "26578450000012121212|65456454", 0
    CheckCase "真实10：两个20位", "采购专用发票，广东省AAA有限公司（NO.）26578450000012121212/26578450000004685792", "26578450000012121212|26578450000004685792", 0
    CheckCase "真实11：四号码保序", "采购专用发票，广东省AAA有限公司（NO.）12121212/54685487/979422543/9847412", "12121212|54685487|979422543|9847412", 0
    CheckCase "真实12：9位7位8位", "采购专用发票，南通BBB有限公司（NO.）151515151/6451225/88512456", "151515151|6451225|88512456", 0

    CheckCase "边界：空字符串", "", "", 0
    CheckCase "边界：全空白", " " & vbTab & vbCrLf & ChrW(160) & ChrW(&H3000), "", 0
    CheckCase "边界：无数字中文", "采购专用发票，某公司", "", 0
    CheckCase "边界：无数字英文", "SupplierABC", "", 0
    CheckCase "排除：只有普通括号数字", "（0000）(12345678901234567890)", "", 0
    CheckCase "排除：只有P业务编码括号", "(P0101204648221)（P99999999999999999999）", "", 0
    CheckCase "排除：括号内NO不提升为标记", "（付款说明 NO.12345678）", "", 0
    CheckCase "排除：嵌套中英文括号", "NO.12345678（备注(12345678901234567890)）", "12345678", 0
    CheckCase "标记：NO英文冒号", "供应商NO:12345678", "12345678", 0
    CheckCase "标记：NO中文冒号", "供应商NO：12345678", "12345678", 0
    CheckCase "标记：英文括号", "供应商(NO.)12345678", "12345678", 0
    CheckCase "标记：小写no", "供应商no.12345678", "12345678", 0
    CheckCase "标记：括号内外空白", "供应商（ NO： ）" & ChrW(160) & "12345678", "12345678", 0
    CheckCase "前缀：有NO的英文字母", "NO.BC12345678", "12345678", 0
    CheckCase "前缀：无NO短号仅待复核", "供应商BC12345678", "12345678", missing
    CheckCase "分隔：斜杠", "NO.12345678/98765432", "12345678|98765432", 0
    CheckCase "分隔：中文逗号", "NO.12345678，98765432", "12345678|98765432", 0
    CheckCase "分隔：顿号", "NO.12345678、98765432", "12345678|98765432", 0
    CheckCase "分隔：混合及空白", "NO.12345678 / 98765432， BC00123456、7654321", "12345678|98765432|00123456|7654321", 0
    CheckCase "尾注：英文数字整体忽略", "NO.12345678--lo5/98765432", "12345678", 0
    CheckCase "尾注：纯数字不作为候选", "NO.12345678--12345", "12345678", numeric
    CheckCase "尾注：数字前后空白", "NO.12345678-- 00012345 ", "12345678", numeric
    CheckCase "尾注：无NO风险与纯数字风险并存", "供应商12345678--12345", "12345678", missing Or numeric
    CheckCase "尾注：字母在括号内也整体忽略", "NO.12345678--123(P9)", "12345678", 0
    CheckCase "尾注：未知中文结构待复核", "NO.12345678--备注5", "12345678", review
    CheckCase "尾注：空尾注待复核", "NO.12345678--", "12345678", review
    CheckCase "排除：数字不得跨括号拼接", "NO.1234(P0000)5678", "", review
    CheckCase "保守：数字内部空格不拼接", "NO.1234 5678", "", review
    CheckCase "保守：小数不拆号码", "NO.12345678.90", "", review
    CheckCase "保守：单横线不拆号码", "NO.12345678-98765432", "", review
    CheckCase "保守：全角数字不自动转换", "NO.１２３４５６７８", "", review
    CheckCase "保守：无NO不收集前置业务数字", "业务20260911供应商12345678", "", review
    CheckCase "保守：非末尾数字不提取", "供应商12345678待处理", "", review
    CheckCase "保守：无NO纯日期状数字仅为待复核候选", "20260911", "20260911", missing
    CheckCase "边界：长度不作为硬门槛", "NO.5/123456789012345678901234567890", "5|123456789012345678901234567890", 0
    CheckCase "去重：保留首次出现及不同前导零", "NO.00123456/12345678/BC00123456/12345678/012345678", "00123456|12345678|012345678", duplicate
    CheckCase "保守：多个NO不擅自选一组", "NO.12345678/NO.98765432", "", review
    CheckCase "保守：空分隔项保留有效项并标风险", "NO.12345678//98765432", "12345678|98765432", review
    CheckCase "保守：首尾分隔符标风险", "NO./12345678/", "12345678", review
    CheckCase "保守：未闭合括号", "NO.12345678(P999", "12345678", review
    CheckCase "保守：单独左括号", "(", "", review
    CheckCase "保守：单独右括号", "）", "", review
    CheckCase "保守：不配对的中英文括号", "（NO.)12345678", "12345678", review Or missing
    CheckCase "排除：普通括号内NO与外部标记区分", "（内部NO.0000）NO.12345678", "12345678", 0
    CheckCase "保守：不支持的项不拆数字", "NO.12345678、付款P999业务", "12345678", review
    CheckCase "标记：不把单词内NO当显式标记", "UNO.12345678", "12345678", missing
    CheckCase "边界：空括号无候选", "（）（）()", "", 0

    raw = "供应商BC12345678/00001234--123"
    p = VATStage2ParseBInvoices(raw): again = VATStage2ParseBInvoices(raw)
    Require p.ReferenceCount = again.ReferenceCount And p.Flags = again.Flags And p.OriginalText = again.OriginalText, "重复调用结果不稳定"
    Require p.References(1).Digits = again.References(1).Digits And p.References(2).StartIndex = again.References(2).StartIndex, "重复调用引用不稳定"
    Require p.References(1).RawFragment = "BC12345678" And p.References(1).StartIndex = 4, "前缀及原文一基索引不符"
    Require VATStage2ParserFlagNames(0) = "" And VATStage2ParserFlagNames(review) = "STRUCTURE_REVIEW" And _
            VATStage2ParserFlagNames(missing) = "NO_MARKER_MISSING" And _
            VATStage2ParserFlagNames(duplicate) = "DUPLICATE_REF_REVIEW" And _
            VATStage2ParserFlagNames(numeric) = "NUMERIC_TRAILER_REVIEW" And _
            VATStage2ParserFlagNames(review Or numeric) = "STRUCTURE_REVIEW|NUMERIC_TRAILER_REVIEW" And _
            VATStage2ParserFlagNames(missing Or review Or numeric Or duplicate) = _
            "NO_MARKER_MISSING|STRUCTURE_REVIEW|NUMERIC_TRAILER_REVIEW|DUPLICATE_REF_REVIEW", "风险名称不符"
    caseCount = caseCount + 1
    logText = logText & "PASS 接口：重复调用确定、前缀定位正确、风险名称稳定" & vbCrLf
    'C1.1 仅补强元数据：原 60 项的候选、排除、位置及顺序断言保留。
    CheckCase "C1.1：同一单元格重复两次", "NO.12345678/12345678", "12345678", duplicate
    CheckCase "C1.1：多次重复仍只有首次候选", "NO.12345678/12345678/12345678/12345678", "12345678", duplicate
    CheckCase "C1.1：不同前缀的同Digits也记录重复", "NO.BC12345678/12345678/BC12345678", "12345678", duplicate
    CheckCase "C1.1：无NO与重复同时存在", "供应商12345678/12345678", "12345678", missing Or duplicate
    CheckCase "C1.1：重复与空分隔独立标记", "NO.12345678//12345678", "12345678", review Or duplicate
    CheckCase "C1.1：四种风险可同时存在", "供应商BC12345678//12345678--123", "12345678", missing Or review Or numeric Or duplicate
    CheckCase "C1.1：已排除括号内容不算重复引用", "NO.12345678(12345678)", "12345678", 0
    CheckCase "C1.1：纯数字尾注不算重复引用", "NO.12345678--12345678", "12345678", numeric
    CheckCase "C1.1：无NO解析失败只有结构风险", "供应商1234 5678", "", review
    p = VATStage2ParseBInvoices("NO.BC12345678/12345678/BC12345678")
    again = VATStage2ParseBInvoices("NO.BC12345678")
    Require p.ReferenceCount = 1 And p.Flags = duplicate And p.References(1).Flags = duplicate, "重复元数据缺失"
    Require p.References(1).StartIndex = 4 And p.References(1).StartIndex = again.References(1).StartIndex And _
            p.References(1).RawFragment = "BC12345678" And p.References(1).RawFragment = again.References(1).RawFragment, "重复覆盖了首次位置或片段"
    Require VATS2_PARSE_REVIEW = (review Or missing) And (review And missing) = 0 And _
            (VATS2_PARSE_REVIEW And numeric) = 0 And (VATS2_PARSE_REVIEW And duplicate) = 0, "旧兼容掩码或独立风险位不符"
    caseCount = caseCount + 1
    logText = logText & "PASS C1.1：首次引用定位不变、独立风险位与兼容掩码正确" & vbCrLf
    VATStage2Parser_SelfTest = "PASS: " & caseCount & " cases" & vbCrLf & logText
    Exit Function
Failed:
    VATStage2Parser_SelfTest = "FAIL: " & Err.Description & vbCrLf & logText
End Function
