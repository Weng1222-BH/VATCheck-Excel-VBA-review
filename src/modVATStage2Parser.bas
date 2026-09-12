Attribute VB_Name = "modVATStage2Parser"
Option Explicit
Option Compare Binary

Public Enum VATS2ParserFlags
    VATS2_FLAG_NONE = 0
    VATS2_STRUCTURE_REVIEW = 1
    VATS2_NUMERIC_TRAILER_REVIEW = 2
    VATS2_NO_MARKER_MISSING = 4
    VATS2_DUPLICATE_REF_REVIEW = 8
    '兼容旧的宽泛风险检查：这是结构异常与无标记提示的组合掩码，不是独立风险位。
    VATS2_PARSE_REVIEW = 5
End Enum

Public Type VATS2InvoiceReference
    Digits As String
    Length As Long
    RawFragment As String
    StartIndex As Long
    Context As String
    Flags As Long
End Type

Public Type VATS2ParseResult
    OriginalText As String
    ReferenceCount As Long
    References() As VATS2InvoiceReference
    Flags As Long
End Type

Public Function VATStage2ParseBInvoices(ByVal rawText As String) As VATS2ParseResult
    Dim parsed As VATS2ParseResult, work As String, trailer As String
    Dim i As Long, trailerAt As Long, markerCount As Long, groupStart As Long, groupEnd As Long
    Dim tokenStart As Long, firstToken As Boolean, noMarker As Boolean
    parsed.OriginalText = rawText
    '用等长空白屏蔽括号，保留所有候选在原文中的位置；不能将两段数字拼接。
    work = MaskParentheses(rawText, parsed.Flags)
    trailerAt = InStr(1, work, "--", vbBinaryCompare)
    If trailerAt > 0 Then
        trailer = TrimSpaces(Mid$(rawText, trailerAt + 2))
        If Not HasLetter(trailer) Then
            If AllDigits(trailer) Then
                parsed.Flags = parsed.Flags Or VATS2_NUMERIC_TRAILER_REVIEW
            Else
                parsed.Flags = parsed.Flags Or VATS2_STRUCTURE_REVIEW
            End If
        End If
        work = Left$(work, trailerAt - 1)
    End If
    groupEnd = Len(work)
    Do While groupEnd > 0
        If Not IsSpace(Mid$(work, groupEnd, 1)) Then Exit Do
        groupEnd = groupEnd - 1
    Loop
    For i = 1 To groupEnd - 2
        If IsMarkerAt(work, i) Then
            markerCount = markerCount + 1
            groupStart = i + 3
        End If
    Next i
    If markerCount > 1 Then
        parsed.Flags = parsed.Flags Or VATS2_STRUCTURE_REVIEW
        GoTo Done
    End If
    noMarker = (markerCount = 0)
    If noMarker Then
        If Not HasDigitLike(work) Then GoTo Done
        '无 NO 时只读取末尾号码组并保留风险，不用长度代替真实性判断。
        groupStart = groupEnd + 1
        Do While groupStart > 1
            If Not IsGroupChar(Mid$(work, groupStart - 1, 1)) Then Exit Do
            groupStart = groupStart - 1
        Loop
        If HasDigitLike(Left$(work, groupStart - 1)) Then
            parsed.Flags = parsed.Flags Or VATS2_STRUCTURE_REVIEW
            GoTo Done
        End If
    End If
    tokenStart = groupStart: firstToken = True
    For i = groupStart To groupEnd + 1
        If i = groupEnd + 1 Or IsSeparator(Mid$(work, i, 1)) Then
            If Not ReadToken(work, tokenStart, i - 1, parsed) Then
                '无标记号码组首项不成立时，不从后面另找一段数字作为起点。
                If noMarker And firstToken Then GoTo Done
            End If
            firstToken = False: tokenStart = i + 1
        End If
    Next i
    If markerCount = 1 And parsed.ReferenceCount = 0 Then parsed.Flags = parsed.Flags Or VATS2_STRUCTURE_REVIEW
Done:
    '仅在无标记且已得到候选时提示缺少 NO；结构问题独立保留，不在这里决定审核方式。
    If noMarker And parsed.ReferenceCount > 0 Then parsed.Flags = parsed.Flags Or VATS2_NO_MARKER_MISSING
    '记录级风险传播给引用，调用方读取候选时也不会遗漏尾注或结构风险。
    For i = 1 To parsed.ReferenceCount
        parsed.References(i).Flags = parsed.Flags
    Next i
    VATStage2ParseBInvoices = parsed
End Function

Public Function VATStage2ParserFlagNames(ByVal flags As Long) As String
    Dim names As String
    If (flags And VATS2_NO_MARKER_MISSING) <> 0 Then names = "NO_MARKER_MISSING"
    If (flags And VATS2_STRUCTURE_REVIEW) <> 0 Then
        If Len(names) > 0 Then names = names & "|"
        names = names & "STRUCTURE_REVIEW"
    End If
    If (flags And VATS2_NUMERIC_TRAILER_REVIEW) <> 0 Then
        If Len(names) > 0 Then names = names & "|"
        names = names & "NUMERIC_TRAILER_REVIEW"
    End If
    If (flags And VATS2_DUPLICATE_REF_REVIEW) <> 0 Then
        If Len(names) > 0 Then names = names & "|"
        names = names & "DUPLICATE_REF_REVIEW"
    End If
    VATStage2ParserFlagNames = names
End Function

Private Function ReadToken(ByVal work As String, ByVal first As Long, ByVal last As Long, _
                           ByRef parsed As VATS2ParseResult) As Boolean
    Dim position As Long, digitsStart As Long, digits As String
    Dim i As Long, contextStart As Long
    Do While first <= last
        If Not IsSpace(Mid$(work, first, 1)) Then Exit Do
        first = first + 1
    Loop
    Do While last >= first
        If Not IsSpace(Mid$(work, last, 1)) Then Exit Do
        last = last - 1
    Loop
    If first > last Then
        '完全空白或已全部排除的原文不是异常；号码组内空项则需要复核。
        If Len(TrimSpaces(work)) > 0 Then parsed.Flags = parsed.Flags Or VATS2_STRUCTURE_REVIEW
        Exit Function
    End If
    position = first
    Do While position <= last
        If Not IsLetter(Mid$(work, position, 1)) Then Exit Do
        position = position + 1
    Loop
    Do While position <= last
        If Not IsSpace(Mid$(work, position, 1)) Then Exit Do
        position = position + 1
    Loop
    digitsStart = position
    Do While position <= last
        If Not IsDigit(Mid$(work, position, 1)) Then Exit Do
        position = position + 1
    Loop
    digits = Mid$(work, digitsStart, position - digitsStart)
    If position <= last Or Len(digits) = 0 Then
        parsed.Flags = parsed.Flags Or VATS2_STRUCTURE_REVIEW
        Exit Function
    End If
    ReadToken = True
    For i = 1 To parsed.ReferenceCount
        If parsed.References(i).Digits = digits Then
            '只表示本次输入单元格内重复引用；候选仍保留首次位置，不涉及跨 B 行关系。
            parsed.Flags = parsed.Flags Or VATS2_DUPLICATE_REF_REVIEW
            Exit Function
        End If
    Next i
    parsed.ReferenceCount = parsed.ReferenceCount + 1
    ReDim Preserve parsed.References(1 To parsed.ReferenceCount)
    contextStart = first - 16
    If contextStart < 1 Then contextStart = 1
    With parsed.References(parsed.ReferenceCount)
        .Digits = digits
        .Length = Len(digits)
        .RawFragment = Mid$(parsed.OriginalText, first, last - first + 1)
        .StartIndex = first
        .Context = Mid$(parsed.OriginalText, contextStart, last - contextStart + 17)
    End With
End Function

Private Function MaskParentheses(ByVal text As String, ByRef flags As Long) As String
    Dim work As String, i As Long, j As Long, stack As String, ch As String, inner As String
    Dim malformed As Boolean, closed As Boolean
    work = text: i = 1
    Do While i <= Len(text)
        ch = Mid$(text, i, 1)
        If ch = "(" Or ch = "（" Then
            stack = ClosingFor(ch): malformed = False: closed = False
            For j = i + 1 To Len(text)
                ch = Mid$(text, j, 1)
                If ch = "(" Or ch = "（" Then
                    stack = stack & ClosingFor(ch)
                ElseIf ch = ")" Or ch = "）" Then
                    If ch <> Right$(stack, 1) Then malformed = True
                    stack = Left$(stack, Len(stack) - 1)
                    If Len(stack) = 0 Then
                        closed = True
                        Exit For
                    End If
                End If
            Next j
            If Not closed Then j = Len(text): malformed = True
            If malformed Then flags = flags Or VATS2_STRUCTURE_REVIEW
            inner = vbNullString
            If j > i Then inner = TrimSpaces(Mid$(text, i + 1, j - i - 1))
            If Not malformed And Len(inner) = 3 And IsMarkerAt(inner, 1) Then
                Mid$(work, i, 1) = " ": Mid$(work, j, 1) = " "
            Else
                Mid$(work, i, j - i + 1) = Space$(j - i + 1)
            End If
            i = j + 1
        Else
            If ch = ")" Or ch = "）" Then
                flags = flags Or VATS2_STRUCTURE_REVIEW
                Mid$(work, i, 1) = " "
            End If
            i = i + 1
        End If
    Loop
    MaskParentheses = work
End Function

Private Function ClosingFor(ByVal opening As String) As String
    If opening = "(" Then ClosingFor = ")" Else ClosingFor = "）"
End Function

Private Function IsMarkerAt(ByVal text As String, ByVal position As Long) As Boolean
    Dim punctuation As String
    If position < 1 Or position + 2 > Len(text) Then Exit Function
    If UCase$(Mid$(text, position, 2)) <> "NO" Then Exit Function
    If position > 1 Then
        If IsLetter(Mid$(text, position - 1, 1)) Or IsDigit(Mid$(text, position - 1, 1)) Then Exit Function
    End If
    punctuation = Mid$(text, position + 2, 1)
    IsMarkerAt = (punctuation = "." Or punctuation = ":" Or punctuation = "：")
End Function

Private Function IsDigit(ByVal ch As String) As Boolean
    IsDigit = (Len(ch) = 1 And ch >= "0" And ch <= "9")
End Function

Private Function IsLetter(ByVal ch As String) As Boolean
    IsLetter = (Len(ch) = 1 And ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z")))
End Function

Private Function HasLetter(ByVal text As String) As Boolean
    Dim i As Long
    For i = 1 To Len(text)
        If IsLetter(Mid$(text, i, 1)) Then HasLetter = True: Exit Function
    Next i
End Function

Private Function HasDigitLike(ByVal text As String) As Boolean
    Dim i As Long, ch As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If IsDigit(ch) Or (ch >= ChrW(&HFF10) And ch <= ChrW(&HFF19)) Then HasDigitLike = True: Exit Function
    Next i
End Function

Private Function AllDigits(ByVal text As String) As Boolean
    Dim i As Long
    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        If Not IsDigit(Mid$(text, i, 1)) Then Exit Function
    Next i
    AllDigits = True
End Function

Private Function IsSpace(ByVal ch As String) As Boolean
    IsSpace = (ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or ch = ChrW(160) Or ch = ChrW(&H3000))
End Function

Private Function TrimSpaces(ByVal text As String) As String
    Dim first As Long, last As Long
    first = 1: last = Len(text)
    Do While first <= last
        If Not IsSpace(Mid$(text, first, 1)) Then Exit Do
        first = first + 1
    Loop
    Do While last >= first
        If Not IsSpace(Mid$(text, last, 1)) Then Exit Do
        last = last - 1
    Loop
    TrimSpaces = Mid$(text, first, last - first + 1)
End Function

Private Function IsSeparator(ByVal ch As String) As Boolean
    IsSeparator = (ch = "/" Or ch = "，" Or ch = "、")
End Function

Private Function IsGroupChar(ByVal ch As String) As Boolean
    IsGroupChar = IsDigit(ch) Or IsLetter(ch) Or IsSpace(ch) Or IsSeparator(ch)
End Function
