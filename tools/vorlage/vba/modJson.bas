Attribute VB_Name = "modJson"
Option Explicit
' ---------------------------------------------------------------------------
'  Minimaler JSON-Parser + UTF-8-Leser (ohne Verweise, ohne Scripting.Dictionary)
'  Objekte  -> Collection mit Schluessel "k_<name>" und Markierung "__obj"
'  Arrays   -> Collection ohne Schluessel
'  null     -> Null
' ---------------------------------------------------------------------------

Private js As String
Private jp As Long
Private jn As Long
Private jbs As Long

Public Function JsonParse(ByVal s As String) As Variant
    Dim v As Variant
    js = s
    jp = 1
    jn = Len(s)
    jbs = 0
    If jn > 0 Then
        If Mid$(js, 1, 1) = ChrW(&HFEFF&) Then jp = 2
    End If
    ParseValue v
    If IsObject(v) Then
        Set JsonParse = v
    Else
        JsonParse = v
    End If
    js = ""
End Function

Public Sub JsonParseInto(ByVal s As String, ByRef v As Variant)
    js = s
    jp = 1
    jn = Len(s)
    jbs = 0
    If jn > 0 Then
        If Mid$(js, 1, 1) = ChrW(&HFEFF&) Then jp = 2
    End If
    ParseValue v
    js = ""
End Sub

Private Sub SkipWs()
    Dim ch As String
    Do While jp <= jn
        ch = Mid$(js, jp, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            jp = jp + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub ParseValue(ByRef v As Variant)
    Dim ch As String
    SkipWs
    If jp > jn Then
        v = Empty
        Exit Sub
    End If
    ch = Mid$(js, jp, 1)
    Select Case ch
        Case "{"
            Set v = ParseObject()
        Case "["
            Set v = ParseArray()
        Case """"
            v = ParseString()
        Case "t"
            jp = jp + 4
            v = True
        Case "f"
            jp = jp + 5
            v = False
        Case "n"
            jp = jp + 4
            v = Null
        Case Else
            v = ParseNumber()
    End Select
End Sub

Private Function ParseObject() As Collection
    Dim c As Collection
    Dim k As String
    Dim v As Variant
    Dim ch As String
    Set c = New Collection
    c.Add True, "__obj"
    jp = jp + 1
    SkipWs
    If Mid$(js, jp, 1) = "}" Then
        jp = jp + 1
        Set ParseObject = c
        Exit Function
    End If
    Do
        SkipWs
        k = ParseString()
        SkipWs
        jp = jp + 1
        v = Empty
        ParseValue v
        On Error Resume Next
        c.Add v, "k_" & k
        Err.Clear
        On Error GoTo 0
        SkipWs
        ch = Mid$(js, jp, 1)
        jp = jp + 1
        If ch <> "," Then Exit Do
    Loop
    Set ParseObject = c
End Function

Private Function ParseArray() As Collection
    Dim c As Collection
    Dim v As Variant
    Dim ch As String
    Set c = New Collection
    jp = jp + 1
    SkipWs
    If Mid$(js, jp, 1) = "]" Then
        jp = jp + 1
        Set ParseArray = c
        Exit Function
    End If
    Do
        v = Empty
        ParseValue v
        c.Add v
        SkipWs
        ch = Mid$(js, jp, 1)
        jp = jp + 1
        If ch <> "," Then Exit Do
    Loop
    Set ParseArray = c
End Function

Private Function HexVal(ByVal h As String) As Long
    Dim i As Long, d As Long, r As Long
    h = LCase$(h)
    For i = 1 To Len(h)
        d = InStr(1, "0123456789abcdef", Mid$(h, i, 1)) - 1
        If d < 0 Then d = 0
        r = r * 16 + d
    Next i
    HexVal = r
End Function

Private Function ParseString() As String
    Dim res As String
    Dim q As Long
    Dim ch As String
    jp = jp + 1
    Do
        q = InStr(jp, js, """")
        If q = 0 Then
            res = res & Mid$(js, jp)
            jp = jn + 1
            Exit Do
        End If
        If jbs < jp And jbs <= jn Then
            jbs = InStr(jp, js, "\")
            If jbs = 0 Then jbs = jn + 1
        End If
        If q < jbs Then
            res = res & Mid$(js, jp, q - jp)
            jp = q + 1
            Exit Do
        End If
        res = res & Mid$(js, jp, jbs - jp)
        ch = Mid$(js, jbs + 1, 1)
        Select Case ch
            Case "n"
                res = res & vbLf
                jp = jbs + 2
            Case "r"
                res = res & vbCr
                jp = jbs + 2
            Case "t"
                res = res & vbTab
                jp = jbs + 2
            Case "b"
                res = res & Chr(8)
                jp = jbs + 2
            Case "f"
                res = res & Chr(12)
                jp = jbs + 2
            Case "u"
                res = res & ChrW(HexVal(Mid$(js, jbs + 2, 4)))
                jp = jbs + 6
            Case Else
                res = res & ch
                jp = jbs + 2
        End Select
    Loop
    ParseString = res
End Function

Private Function ParseNumber() As Variant
    Dim st As Long
    Dim ch As String
    st = jp
    Do While jp <= jn
        ch = Mid$(js, jp, 1)
        If InStr(1, "+-0123456789.eE", ch) > 0 Then
            jp = jp + 1
        Else
            Exit Do
        End If
    Loop
    ParseNumber = Val(Mid$(js, st, jp - st))
End Function

' ------------------------------------------------------------------ Zugriff
Public Function JIsObj(ByVal c As Variant) As Boolean
    Dim t As Variant
    If Not IsObject(c) Then Exit Function
    If c Is Nothing Then Exit Function
    On Error Resume Next
    t = c("__obj")
    JIsObj = (Err.Number = 0)
    Err.Clear
End Function

Public Function JStr(ByVal o As Object, ByVal k As String) As String
    Dim v As Variant
    If o Is Nothing Then Exit Function
    On Error Resume Next
    v = o("k_" & k)
    If Err.Number <> 0 Then
        Err.Clear
        JStr = ""
        Exit Function
    End If
    On Error GoTo 0
    If IsNull(v) Or IsEmpty(v) Then
        JStr = ""
    ElseIf VarType(v) = vbBoolean Then
        JStr = IIf(v, "true", "false")
    Else
        JStr = CStr(v)
    End If
End Function

Public Function JNum(ByVal o As Object, ByVal k As String) As Double
    Dim v As Variant
    If o Is Nothing Then Exit Function
    On Error Resume Next
    v = o("k_" & k)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    If IsNumeric(v) Then JNum = CDbl(v)
End Function

Public Function JObj(ByVal o As Object, ByVal k As String) As Object
    Dim r As Object
    If o Is Nothing Then
        Set JObj = Nothing
        Exit Function
    End If
    On Error Resume Next
    Set r = o("k_" & k)
    If Err.Number <> 0 Then
        Err.Clear
        Set r = Nothing
    End If
    On Error GoTo 0
    Set JObj = r
End Function

' ------------------------------------------------------------------ Datei lesen (UTF-8)
Public Function LeseUtf8(ByVal pfad As String) As String
    Dim s As String
    If LeseUtf8Ado(pfad, s) Then
        LeseUtf8 = s
    Else
        LeseUtf8 = LeseUtf8Manuell(pfad)
    End If
End Function

Private Function LeseUtf8Ado(ByVal pfad As String, ByRef s As String) As Boolean
    Dim st As Object
    On Error GoTo Nein
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile pfad
    s = st.ReadText(-1)
    st.Close
    If Len(s) > 0 Then
        If Left$(s, 1) = ChrW(&HFEFF&) Then s = Mid$(s, 2)
    End If
    LeseUtf8Ado = True
    Exit Function
Nein:
    LeseUtf8Ado = False
End Function

Private Function LeseUtf8Manuell(ByVal pfad As String) As String
    Dim f As Integer
    Dim b() As Byte
    Dim n As Long, i As Long, c As Long, cnt As Long, np As Long
    Dim chunk As String
    Dim parts() As String
    f = FreeFile
    Open pfad For Binary Access Read As #f
    n = LOF(f)
    If n = 0 Then
        Close #f
        Exit Function
    End If
    ReDim b(0 To n - 1)
    Get #f, , b
    Close #f
    ReDim parts(0 To n \ 256 + 2)
    i = 0
    If n >= 3 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then i = 3
    End If
    Do While i < n
        c = b(i)
        If c < &H80& Then
            i = i + 1
        ElseIf c < &HE0& Then
            c = ((c And &H1F&) * 64&) Or (b(i + 1) And &H3F&)
            i = i + 2
        ElseIf c < &HF0& Then
            c = ((c And &HF&) * 4096&) Or ((b(i + 1) And &H3F&) * 64&) Or (b(i + 2) And &H3F&)
            i = i + 3
        Else
            c = ((c And &H7&) * 262144) Or ((b(i + 1) And &H3F&) * 4096&) Or ((b(i + 2) And &H3F&) * 64&) Or (b(i + 3) And &H3F&)
            i = i + 4
            c = c - &H10000
            chunk = chunk & ChrW(&HD800& + (c \ 1024&))
            c = &HDC00& + (c And &H3FF&)
        End If
        chunk = chunk & ChrW(c)
        cnt = cnt + 1
        If cnt >= 256 Then
            parts(np) = chunk
            np = np + 1
            chunk = ""
            cnt = 0
        End If
    Loop
    parts(np) = chunk
    ReDim Preserve parts(0 To np)
    LeseUtf8Manuell = Join(parts, "")
End Function
