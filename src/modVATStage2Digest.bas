Attribute VB_Name = "modVATStage2Digest"
Option Explicit
Option Compare Binary

'SHA256-UTF8：FIPS 180-4 SHA-256；UTF-8无BOM。无文件、COM、API或外部程序。
'无配对UTF-16代理项明确拒绝，不能用替换字符悄悄改变输入。
Public Function VATStage2SHA256(ByVal text As String) As String
    Dim data() As Byte, n As Long, size As Long, bits As Double
    Dim h(0 To 7) As Long, w(0 To 63) As Long, k(0 To 63) As Long
    Dim constants As Variant, i As Long, j As Long, offset As Long
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, hh As Long
    Dim s0 As Long, s1 As Long, t1 As Long, t2 As Long
    EncodeUTF8 text, data, n
    size = CLng((Fix((CDbl(n) + 8) / 64) + 1) * 64)
    ReDim Preserve data(0 To size - 1)
    data(n) = 128: bits = CDbl(n) * 8
    For i = size - 1 To size - 8 Step -1
        data(i) = CByte(bits - Fix(bits / 256) * 256): bits = Fix(bits / 256)
    Next i
    constants = Split("428a2f98 71374491 b5c0fbcf e9b5dba5 3956c25b 59f111f1 923f82a4 ab1c5ed5 " & _
        "d807aa98 12835b01 243185be 550c7dc3 72be5d74 80deb1fe 9bdc06a7 c19bf174 " & _
        "e49b69c1 efbe4786 0fc19dc6 240ca1cc 2de92c6f 4a7484aa 5cb0a9dc 76f988da " & _
        "983e5152 a831c66d b00327c8 bf597fc7 c6e00bf3 d5a79147 06ca6351 14292967 " & _
        "27b70a85 2e1b2138 4d2c6dfc 53380d13 650a7354 766a0abb 81c2c92e 92722c85 " & _
        "a2bfe8a1 a81a664b c24b8b70 c76c51a3 d192e819 d6990624 f40e3585 106aa070 " & _
        "19a4c116 1e376c08 2748774c 34b0bcb5 391c0cb3 4ed8aa4a 5b9cca4f 682e6ff3 " & _
        "748f82ee 78a5636f 84c87814 8cc70208 90befffa a4506ceb bef9a3f7 c67178f2", " ")
    For i = 0 To 63: k(i) = CLng("&H" & constants(i)): Next i
    constants = Split("6a09e667 bb67ae85 3c6ef372 a54ff53a 510e527f 9b05688c 1f83d9ab 5be0cd19", " ")
    For i = 0 To 7: h(i) = CLng("&H" & constants(i)): Next i
    For offset = 0 To size - 1 Step 64
        For i = 0 To 15
            j = offset + i * 4
            w(i) = Word32(CDbl(data(j)) * 16777216# + CDbl(data(j + 1)) * 65536# + CLng(data(j + 2)) * 256& + data(j + 3))
        Next i
        For i = 16 To 63
            s0 = Rotate(w(i - 15), 7) Xor Rotate(w(i - 15), 18) Xor Shift(w(i - 15), 3)
            s1 = Rotate(w(i - 2), 17) Xor Rotate(w(i - 2), 19) Xor Shift(w(i - 2), 10)
            w(i) = Word32(CDbl(w(i - 16)) + s0 + w(i - 7) + s1)
        Next i
        a = h(0): b = h(1): c = h(2): d = h(3): e = h(4): f = h(5): g = h(6): hh = h(7)
        For i = 0 To 63
            s1 = Rotate(e, 6) Xor Rotate(e, 11) Xor Rotate(e, 25)
            t1 = Word32(CDbl(hh) + s1 + ((e And f) Xor ((Not e) And g)) + k(i) + w(i))
            s0 = Rotate(a, 2) Xor Rotate(a, 13) Xor Rotate(a, 22)
            t2 = Word32(CDbl(s0) + ((a And b) Xor (a And c) Xor (b And c)))
            hh = g: g = f: f = e: e = Word32(CDbl(d) + t1)
            d = c: c = b: b = a: a = Word32(CDbl(t1) + t2)
        Next i
        h(0) = Word32(CDbl(h(0)) + a): h(1) = Word32(CDbl(h(1)) + b)
        h(2) = Word32(CDbl(h(2)) + c): h(3) = Word32(CDbl(h(3)) + d)
        h(4) = Word32(CDbl(h(4)) + e): h(5) = Word32(CDbl(h(5)) + f)
        h(6) = Word32(CDbl(h(6)) + g): h(7) = Word32(CDbl(h(7)) + hh)
    Next offset
    For i = 0 To 7: VATStage2SHA256 = VATStage2SHA256 & LCase$(Right$("00000000" & Hex$(h(i)), 8)): Next i
End Function

'所有中间整数均小于2^53，Double仅用于无损32位模加，不参与业务金额处理。
Private Function Word32(ByVal value As Double) As Long
    value = value - Int(value / 4294967296#) * 4294967296#
    If value >= 2147483648# Then value = value - 4294967296#
    Word32 = CLng(value)
End Function
Private Function Unsigned(ByVal value As Long) As Double
    Unsigned = value
    If value < 0 Then Unsigned = Unsigned + 4294967296#
End Function
Private Function Shift(ByVal value As Long, ByVal n As Long) As Long
    Shift = CLng(Fix(Unsigned(value) / (2# ^ n)))
End Function
Private Function Rotate(ByVal value As Long, ByVal n As Long) As Long
    Dim u As Double, q As Double
    u = Unsigned(value): q = Fix(u / (2# ^ n))
    Rotate = Word32(q + (u - q * (2# ^ n)) * (2# ^ (32 - n)))
End Function
Private Sub EncodeUTF8(ByVal text As String, ByRef data() As Byte, ByRef n As Long)
    Dim i As Long, u As Long, low As Long
    ReDim data(0 To CLng(CDbl(Len(text)) * 3))
    i = 1
    Do While i <= Len(text)
        u = AscW(Mid$(text, i, 1)) And &HFFFF&
        If u >= &HD800& And u <= &HDBFF& Then
            If i = Len(text) Then Err.Raise vbObjectError + 720, , "UTF-16代理项未配对。"
            low = AscW(Mid$(text, i + 1, 1)) And &HFFFF&
            If low < &HDC00& Or low > &HDFFF& Then Err.Raise vbObjectError + 720, , "UTF-16代理项未配对。"
            u = &H10000 + (u - &HD800&) * 1024 + low - &HDC00&: i = i + 1
        ElseIf u >= &HDC00& And u <= &HDFFF& Then
            Err.Raise vbObjectError + 720, , "UTF-16代理项未配对。"
        End If
        Select Case u
            Case 0 To 127: data(n) = u: n = n + 1
            Case 128 To 2047
                data(n) = 192 Or (u \ 64): data(n + 1) = 128 Or (u And 63): n = n + 2
            Case 2048 To 65535
                data(n) = 224 Or (u \ 4096): data(n + 1) = 128 Or ((u \ 64) And 63)
                data(n + 2) = 128 Or (u And 63): n = n + 3
            Case Else
                data(n) = 240 Or (u \ 262144): data(n + 1) = 128 Or ((u \ 4096) And 63)
                data(n + 2) = 128 Or ((u \ 64) And 63): data(n + 3) = 128 Or (u And 63): n = n + 4
        End Select
        i = i + 1
    Loop
End Sub
