Attribute VB_Name = "modImport"
Option Explicit
' ===========================================================================
'  George-Import (JSON-Umsatzexporte) - alles kommt aus der Exportdatei:
'  Konten, Buchungen, Zahlungsart. Danach läuft die Abo-Erkennung (modAbo).
'  Einstieg:  ImportDateienAuswaehlen   - Dateidialog (Mehrfachauswahl)
'             ImportOrdnerEinlesen      - alle *.json im Import-Ordner
'             ImportOrdnerOeffnen       - Ordner im Explorer öffnen
' ===========================================================================

Private Const E_ORDNER As String = "B5"     ' Einstellungen: Import-Ordner
Private Const E_KWBAR As String = "B12"     ' Einstellungen: Stichwörter Bargeld

' ---- Konten (Arbeitsdaten)
Private accKey() As String, accNum() As String, accRaw() As String, accTitle() As String
Private accName() As String, accTyp() As String, accFile() As String
Private accCnt() As Long, accNew() As Long, accPN() As Long, nAcc As Long
Private kwBar As String

' ===========================================================================
'  Einstiegspunkte
' ===========================================================================
Public Sub ImportDateienAuswaehlen()
    Dim f As Variant
    On Error Resume Next
    If Environ$("USERPROFILE") <> "" Then ChDrive Left$(Environ$("USERPROFILE"), 1): ChDir Environ$("USERPROFILE") & "\Downloads"
    On Error GoTo 0
    f = Application.GetOpenFilename("George-Export (*.json),*.json", , "George-Export(e) auswählen - mehrere mit Strg markieren", , True)
    If Not IsArray(f) Then Exit Sub
    ImportiereDateien f, False, False
End Sub

Public Sub ImportOrdnerEinlesen()
    Dim ordner As String, fn As String
    Dim files() As String
    Dim n As Long
    ordner = ImportOrdner(True)
    If ordner = "" Then Exit Sub
    fn = Dir(ordner & PS() & "*.json")
    Do While fn <> ""
        n = n + 1
        ReDim Preserve files(1 To n)
        files(n) = ordner & PS() & fn
        fn = Dir()
    Loop
    If n = 0 Then
        MsgBox "Im Import-Ordner liegen keine JSON-Dateien." & vbLf & vbLf & ordner & vbLf & vbLf & _
               "Exportiere die Umsätze in George als JSON und speichere/ziehe sie in diesen Ordner.", vbInformation, "George-Import"
        Exit Sub
    End If
    ImportiereDateien files, True, False
End Sub

Public Sub ImportOrdnerOeffnen()
    Dim ordner As String
    ordner = ImportOrdner(True)
    If ordner = "" Then Exit Sub
    On Error Resume Next
    ThisWorkbook.FollowHyperlink ordner
    If Err.Number <> 0 Then MsgBox "Import-Ordner: " & ordner, vbInformation, "George-Import"
End Sub

Public Sub PruefeImportOrdnerBeimOeffnen()
    Dim ordner As String, fn As String
    Dim n As Long
    On Error GoTo Ende
    ordner = ImportOrdner(False)
    If ordner = "" Then Exit Sub
    If Dir(ordner, vbDirectory) = "" Then Exit Sub
    fn = Dir(ordner & PS() & "*.json")
    Do While fn <> ""
        n = n + 1
        fn = Dir()
    Loop
    If n > 0 Then
        If MsgBox(n & " Datei(en) im Import-Ordner gefunden." & vbLf & "Jetzt importieren?", vbYesNo + vbQuestion, "George-Import") = vbYes Then
            ImportOrdnerEinlesen
        End If
    End If
Ende:
End Sub

' Menüband „Finanzen“
Public Sub Ribbon_Import(control As Object)
    ImportDateienAuswaehlen
End Sub

Public Sub Ribbon_Ordner(control As Object)
    ImportOrdnerEinlesen
End Sub

Public Sub Ribbon_OrdnerOeffnen(control As Object)
    ImportOrdnerOeffnen
End Sub

' Für automatisierte Tests: Pfade mit | getrennt, ohne Dialoge
Public Function ImportTest(ByVal pfade As String) As String
    ImportTest = ImportiereDateien(Split(pfade, "|"), False, True)
End Function

' ===========================================================================
'  Kern
' ===========================================================================
Public Function ImportiereDateien(ByVal files As Variant, ByVal verschieben As Boolean, ByVal still As Boolean) As String
    Dim wsB As Worksheet, wsO As Worksheet
    Dim txs As Collection, txF As Collection, keys As Collection, occ As Collection
    Dim root As Variant, lst As Variant, t As Variant
    Dim i As Long, j As Long, n As Long, r0 As Long, lastRow As Long, nNeu As Long, nDup As Long
    Dim out() As Variant
    Dim fpath As String, txt As String, key As String, msg As String, aktDatei As String, aboMsg As String
    Dim tx As Object, pa As Object
    Dim val As Double, dat As Variant
    Dim partner As String, ref As String, iban As String, gl As String, mand As String, ft As String
    Dim ai As Long
    Dim oldCalc As Long

    On Error GoTo Fehler
    Set wsB = Blatt("shBuchungen", "Buchungen")
    Set wsO = Blatt("shUebersicht", "Übersicht")
    kwBar = ""
    On Error Resume Next
    kwBar = SStr(Blatt("shEinst", "Einstellungen").Range(E_KWBAR).Value)
    Application.ScreenUpdating = False
    oldCalc = Application.Calculation
    Application.Calculation = -4135
    Application.StatusBar = "George-Import läuft ..."
    On Error GoTo Fehler
    If kwBar = "" Then kwBar = "bargeld;bankomat;behebung;atm;automat;auszahlung"

    ' ---- Dateien lesen
    Set txs = New Collection
    Set txF = New Collection
    For i = LBound(files) To UBound(files)
        fpath = CStr(files(i))
        If Len(fpath) > 0 Then
            aktDatei = fpath
            txt = LeseUtf8(fpath)
            root = Empty
            JsonParseInto txt, root
            If IsObject(root) Then
                Set lst = root
                If JIsObj(root) Then
                    Set lst = JObj(root, "transactions")
                    If lst Is Nothing Then Set lst = JObj(root, "collection")
                    If lst Is Nothing Then Set lst = New Collection
                End If
                For Each t In lst
                    If JIsObj(t) Then
                        txs.Add t
                        txF.Add DateiName(fpath)
                    End If
                Next t
            End If
        End If
    Next i
    aktDatei = ""

    ' ---- Konten aus dem Export erkennen und in der Übersicht eintragen
    ErkenneKonten txs, txF, wsO

    ' ---- vorhandene Buchungen (Schlüssel)
    Set keys = New Collection
    Set occ = New Collection
    lastRow = LetzteZeile(wsB)
    LadeSchluessel wsB, lastRow, keys

    ' ---- neue Buchungen aufbauen
    n = txs.Count
    If n > 0 Then ReDim out(1 To n, 1 To N_COLS)
    nNeu = 0
    For i = 1 To n
        Set tx = txs(i)
        ai = AccIndex(GruppenKey(tx, CStr(txF(i))))
        key = TxKey(tx, accName(ai))
        ' gleiche Buchung mehrfach in derselben Datei (z. B. zwei gleiche Einkäufe am selben Tag)
        If InSet(occ, key) Then
            j = CLng(occ("x" & key)) + 1
            occ.Remove "x" & key
            occ.Add j, "x" & key
            key = key & "#" & j
        Else
            occ.Add 1, "x" & key
        End If
        If InSet(keys, key) Then
            nDup = nDup + 1
        Else
            keys.Add 0, "x" & key
            val = TxBetrag(tx)
            dat = TxDatum(tx)
            partner = Trim$(JStr(tx, "partnerName"))
            If partner = "" Then partner = Trim$(JStr(tx, "merchantName"))
            If partner = "" Then partner = Trim$(IIf(val < 0, JStr(tx, "receiverName"), JStr(tx, "senderName")))
            ref = TxReferenz(tx)
            If partner = "" Then partner = SauberName(ref)
            Set pa = JObj(tx, "partnerAccount")
            iban = ""
            If Not pa Is Nothing Then
                iban = Trim$(JStr(pa, "iban"))
                If iban = "" Then iban = Trim$(JStr(pa, "number"))
            End If
            gl = Trim$(JStr(tx, "sepaCreditorId"))
            mand = Trim$(JStr(tx, "sepaMandateId"))
            ft = " " & Fold(partner & " " & ref & " " & JStr(tx, "receiverReference")) & " "
            nNeu = nNeu + 1
            out(nNeu, C_DATUM) = dat
            out(nNeu, C_KONTO) = TxtZelle(accName(ai))
            out(nNeu, C_PARTNER) = TxtZelle(partner)
            out(nNeu, C_REF) = TxtZelle(ref)
            out(nNeu, C_BETRAG) = val
            out(nNeu, C_ZART) = Zahlungsart(ft, iban, gl, mand, accTyp(ai), val)
            out(nNeu, C_ABO) = Empty
            out(nNeu, C_ART) = Empty
            out(nNeu, C_NOTIZ) = Empty
            out(nNeu, C_MONAT) = Empty
            out(nNeu, C_IBAN) = TxtZelle(iban)
            out(nNeu, C_GLAEUB) = TxtZelle(gl)
            out(nNeu, C_MANDAT) = TxtZelle(mand)
            out(nNeu, C_KEY) = key
            accNew(ai) = accNew(ai) + 1
        End If
    Next i

    ' ---- schreiben
    If nNeu > 0 Then
        r0 = lastRow + 1
        If r0 < B_FIRST Then r0 = B_FIRST
        Dim blk() As Variant
        ReDim blk(1 To nNeu, 1 To N_COLS)
        For i = 1 To nNeu
            For j = 1 To N_COLS
                blk(i, j) = out(i, j)
            Next j
        Next i
        wsB.Range(wsB.Cells(r0, 1), wsB.Cells(r0 + nNeu - 1, N_COLS)).Value = blk
        FormatiereZeilen wsB, r0, r0 + nNeu - 1
        SortiereBuchungen wsB, r0 + nNeu - 1
        FormatiereZeilen wsB, B_FIRST, r0 + nNeu - 1
    End If
    On Error Resume Next
    wsB.PageSetup.PrintArea = "$A$1:$I$" & Application.WorksheetFunction.Max(B_FIRST, LetzteZeile(wsB))
    On Error GoTo Fehler

    ' ---- Abo-Erkennung über alle Buchungen
    aboMsg = ErkenneAbos(True)

    ' ---- Zusammenfassung
    msg = nNeu & " neue Buchung(en) importiert"
    If nDup > 0 Then msg = msg & ", " & nDup & " schon vorhanden (übersprungen)"
    msg = msg & "." & vbLf & vbLf & "Konten:" & vbLf
    For j = 1 To nAcc
        msg = msg & "   " & accName(j) & " (" & accTyp(j) & "): " & accNew(j) & " neu von " & accCnt(j) & "  [" & accFile(j) & "]" & vbLf
    Next j
    msg = msg & vbLf & "Abo-Erkennung:" & vbLf & Replace(aboMsg, " | ", vbLf)

    On Error Resume Next
    wsO.Range("B5").Value = Format(Now, "dd.mm.yyyy hh:mm") & "  -  " & nNeu & " neu, " & nDup & " Duplikate"
    If verschieben Then VerschiebeErledigt files
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True
    Application.StatusBar = False
    On Error GoTo 0
    If Not still Then MsgBox msg, vbInformation, "George-Import"
    ImportiereDateien = Replace(msg, vbLf, " | ")
    Exit Function

Fehler:
    Dim em As String
    em = "Fehler beim Import" & IIf(aktDatei <> "", " (Datei " & DateiName(aktDatei) & ")", "") & ":" & vbLf & Err.Description
    On Error GoTo -1
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If Not still Then MsgBox em, vbExclamation, "George-Import"
    ImportiereDateien = "FEHLER: " & Replace(em, vbLf, " ")
End Function

' Zahlungsart aus den Feldern des Exports
Private Function Zahlungsart(ByVal ft As String, ByVal iban As String, ByVal gl As String, ByVal mand As String, _
                             ByVal typ As String, ByVal val As Double) As String
    If gl <> "" Or mand <> "" Then
        Zahlungsart = "Lastschrift"
    ElseIf InStr(ft, "dauerauftrag") > 0 Then
        Zahlungsart = "Dauerauftrag"
    ElseIf HatStichwort(ft, kwBar) Then
        Zahlungsart = "Bargeld"
    ElseIf typ = "Kreditkarte" Then
        Zahlungsart = IIf(val > 0, "Gutschrift", "Karte")
    ElseIf InStr(ft, " pos ") > 0 Or InStr(ft, "kartenzahlung") > 0 Or InStr(ft, " dankt ") > 0 Or InStr(ft, "bezahlung") > 0 Or InStr(ft, "e-com") > 0 Then
        Zahlungsart = "Karte"
    ElseIf Len(NormKonto(iban)) >= 8 Then
        Zahlungsart = "Überweisung"
    ElseIf val > 0 Then
        Zahlungsart = "Gutschrift"
    Else
        Zahlungsart = "Sonstiges"
    End If
End Function

Private Sub LadeSchluessel(ByVal ws As Worksheet, ByVal lastRow As Long, ByRef keys As Collection)
    Dim r As Long
    Dim data As Variant
    Dim k As String
    If lastRow < B_FIRST Then Exit Sub
    data = ws.Range(ws.Cells(B_FIRST, C_KEY), ws.Cells(lastRow, C_KEY)).Value
    For r = 1 To UBound(data, 1)
        k = SStr(data(r, 1))
        If k <> "" Then
            If Not InSet(keys, k) Then keys.Add B_FIRST + r - 1, "x" & k
        End If
    Next r
End Sub

' ===========================================================================
'  Konten erkennen (nur aus dem Export: Kontonummer + Kontobezeichnung)
' ===========================================================================
Private Function GruppenKey(ByVal tx As Object, ByVal fname As String) As String
    Dim num As String
    num = NormKonto(JStr(tx, "ownerAccountNumber"))
    If Len(num) >= 4 Then
        GruppenKey = "N:" & num
    Else
        GruppenKey = "F:" & LCase$(fname) & "|" & LCase$(JStr(tx, "ownerAccountTitle"))
    End If
End Function

Private Sub ErkenneKonten(ByVal txs As Collection, ByVal txF As Collection, ByVal wsO As Worksheet)
    Dim i As Long, r As Long, a As Long, frei As Long
    Dim num As String, ib As String, typ As String, gk As String, title As String, fname As String, nm As String
    Dim tx As Object
    Dim gefunden As Boolean
    nAcc = 0
    ReDim accNum(1 To 1): ReDim accRaw(1 To 1): ReDim accTyp(1 To 1): ReDim accName(1 To 1): ReDim accTitle(1 To 1)
    ReDim accFile(1 To 1): ReDim accCnt(1 To 1): ReDim accNew(1 To 1): ReDim accKey(1 To 1): ReDim accPN(1 To 1)
    ' 1. Durchlauf: Gruppen bilden (je Kontonummer, sonst je Datei)
    For i = 1 To txs.Count
        Set tx = txs(i)
        fname = CStr(txF(i))
        gk = GruppenKey(tx, fname)
        a = AccIndex(gk)
        If a = 0 Then
            nAcc = nAcc + 1
            ReDim Preserve accNum(1 To nAcc): ReDim Preserve accRaw(1 To nAcc): ReDim Preserve accTyp(1 To nAcc)
            ReDim Preserve accName(1 To nAcc): ReDim Preserve accTitle(1 To nAcc): ReDim Preserve accFile(1 To nAcc)
            ReDim Preserve accCnt(1 To nAcc): ReDim Preserve accNew(1 To nAcc): ReDim Preserve accKey(1 To nAcc)
            ReDim Preserve accPN(1 To nAcc)
            a = nAcc
            accKey(a) = gk
            accNum(a) = NormKonto(JStr(tx, "ownerAccountNumber"))
            accRaw(a) = Trim$(JStr(tx, "ownerAccountNumber"))
            accTitle(a) = Trim$(JStr(tx, "ownerAccountTitle"))
            accFile(a) = fname
        End If
        accCnt(a) = accCnt(a) + 1
        If Trim$(JStr(tx, "partnerName")) <> "" Then accPN(a) = accPN(a) + 1
    Next i
    ' 2. Durchlauf: Konto in der Übersicht finden oder anlegen
    For a = 1 To nAcc
        num = accNum(a)
        title = accTitle(a)
        gefunden = False
        ' a) Kontonummer steht schon in der Übersicht
        If Len(num) >= 4 Then
            For r = K_FIRST To K_LAST
                ib = NormKonto(SStr(wsO.Cells(r, 3).Value))
                If SStr(wsO.Cells(r, 1).Value) <> "" And GleichesKonto(num, ib) Then
                    accName(a) = SStr(wsO.Cells(r, 1).Value)
                    accTyp(a) = SStr(wsO.Cells(r, 2).Value)
                    gefunden = True
                    Exit For
                End If
            Next r
        End If
        ' b) Kontobezeichnung = Kontoname in der Übersicht (ohne hinterlegte Nummer)
        If Not gefunden And title <> "" Then
            For r = K_FIRST To K_LAST
                If LCase$(SStr(wsO.Cells(r, 1).Value)) = LCase$(title) And Trim$(SStr(wsO.Cells(r, 3).Value)) = "" Then
                    accName(a) = SStr(wsO.Cells(r, 1).Value)
                    accTyp(a) = SStr(wsO.Cells(r, 2).Value)
                    If Len(num) >= 4 Then wsO.Cells(r, 3).Value = "'" & accRaw(a)
                    gefunden = True
                    Exit For
                End If
            Next r
        End If
        ' c) neu anlegen: Name = Kontobezeichnung aus George, Typ geraten
        If Not gefunden Then
            typ = RateTyp(title)
            If typ = "" Then typ = RateTyp(accFile(a))
            If typ = "" Then
                If accPN(a) = 0 Then typ = "Kreditkarte" Else typ = "Girokonto"
            End If
            nm = title
            If nm = "" Then nm = typ
            ' Name schon vergeben (anderes Konto)? -> Nummer anhängen
            For r = K_FIRST To K_LAST
                If LCase$(SStr(wsO.Cells(r, 1).Value)) = LCase$(nm) Then
                    If Len(num) >= 4 Then nm = nm & " " & Right$(num, 4) Else nm = nm & " " & a
                    Exit For
                End If
            Next r
            frei = 0
            For r = K_FIRST To K_LAST
                If Trim$(SStr(wsO.Cells(r, 1).Value)) = "" Then frei = r: Exit For
            Next r
            If frei > 0 Then
                wsO.Cells(frei, 1).Value = TxtZelle(nm)
                wsO.Cells(frei, 2).Value = typ
                If Len(num) >= 4 Then wsO.Cells(frei, 3).Value = "'" & accRaw(a)
            End If
            accName(a) = nm
            accTyp(a) = typ
        End If
        If accTyp(a) = "" Then accTyp(a) = "Girokonto"
    Next a
End Sub

Private Function AccIndex(ByVal gk As String) As Long
    Dim i As Long
    For i = 1 To nAcc
        If accKey(i) = gk Then AccIndex = i: Exit Function
    Next i
End Function

Public Function RateTyp(ByVal s As String) As String
    Dim t As String
    t = " " & Fold(s) & " "
    t = Replace(t, "sparkasse", " ")
    If Trim$(t) = "" Then Exit Function
    If InStr(t, "kredit") > 0 Or InStr(t, "card") > 0 Or InStr(t, "visa") > 0 Or InStr(t, "master") > 0 Or InStr(t, "karte") > 0 Then
        RateTyp = "Kreditkarte"
    ElseIf InStr(t, "bauspar") > 0 Then
        RateTyp = "Bausparkonto"
    ElseIf InStr(t, "spar") > 0 Or InStr(t, "kapital") > 0 Then
        RateTyp = "Sparkonto"
    ElseIf InStr(t, "giro") > 0 Or InStr(t, "konto") > 0 Or InStr(t, "gehalt") > 0 Then
        RateTyp = "Girokonto"
    End If
End Function

' ===========================================================================
'  Schreiben / Sortieren
' ===========================================================================
Private Function GenutzteBis(ByVal ws As Worksheet) As Long
    With ws.UsedRange
        GenutzteBis = .Row + .Rows.Count - 1
    End With
End Function

Public Function LetzteZeile(ByVal ws As Worksheet) As Long
    Dim bis As Long, r As Long
    Dim a As Variant, k As Variant
    LetzteZeile = B_FIRST - 1
    bis = GenutzteBis(ws)
    If bis < B_FIRST Then Exit Function
    a = ws.Range(ws.Cells(B_FIRST, C_DATUM), ws.Cells(bis + 1, C_DATUM)).Value
    k = ws.Range(ws.Cells(B_FIRST, C_KEY), ws.Cells(bis + 1, C_KEY)).Value
    For r = UBound(a, 1) To 1 Step -1
        If SStr(a(r, 1)) <> "" Or SStr(k(r, 1)) <> "" Then
            LetzteZeile = B_FIRST + r - 1
            Exit Function
        End If
    Next r
End Function

' Zahlenformate und Monatsformel für die Zeilen vonZeile..bisZeile
Private Sub FormatiereZeilen(ByVal ws As Worksheet, ByVal vonZeile As Long, ByVal bisZeile As Long)
    On Error Resume Next
    ws.Range(ws.Cells(vonZeile, C_DATUM), ws.Cells(bisZeile, C_DATUM)).NumberFormat = "DD.MM.YYYY"
    ws.Range(ws.Cells(vonZeile, C_BETRAG), ws.Cells(bisZeile, C_BETRAG)).NumberFormat = "#,##0.00 €;[Red]-#,##0.00 €"
    ws.Range(ws.Cells(vonZeile, C_MONAT), ws.Cells(bisZeile, C_MONAT)).NumberFormat = "MMM YYYY"
    ws.Range(ws.Cells(vonZeile, C_MONAT), ws.Cells(bisZeile, C_MONAT)).FormulaR1C1 = _
        "=IF(RC[-9]="""","""",DATE(YEAR(RC[-9]),MONTH(RC[-9]),1))"
End Sub

Private Sub SortiereBuchungen(ByVal ws As Worksheet, ByVal bisZeile As Long)
    Dim rng As Range
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    Set rng = ws.Range(ws.Cells(B_FIRST, 1), ws.Cells(bisZeile, N_COLS))
    rng.Sort Key1:=ws.Cells(B_FIRST, C_DATUM), Order1:=1, Key2:=ws.Cells(B_FIRST, C_KONTO), Order2:=1, Header:=2
    On Error GoTo 0
End Sub

Private Sub VerschiebeErledigt(ByVal files As Variant)
    Dim ordner As String, ziel As String, fn As String
    Dim i As Long
    On Error Resume Next
    ordner = ImportOrdner(True)
    If ordner = "" Then Exit Sub
    ordner = ordner & PS() & "erledigt"
    If Dir(ordner, vbDirectory) = "" Then MkDir ordner
    For i = LBound(files) To UBound(files)
        fn = DateiName(CStr(files(i)))
        ziel = ordner & PS() & fn
        If Dir(ziel) <> "" Then ziel = ordner & PS() & Format(Now, "yyyymmdd-hhmmss") & "_" & fn
        Name CStr(files(i)) As ziel
    Next i
End Sub

' ===========================================================================
'  Hilfsfunktionen (auch von modAbo genutzt)
' ===========================================================================
Public Function Blatt(ByVal code As String, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.CodeName = code Then
            Set Blatt = ws
            Exit Function
        End If
    Next ws
    Set Blatt = ThisWorkbook.Worksheets(nm)
End Function

Private Function PS() As String
    PS = Application.PathSeparator
End Function

Private Function ImportOrdner(ByVal anlegen As Boolean) As String
    Dim p As String, basis As String
    On Error Resume Next
    p = Trim$(CStr(Blatt("shEinst", "Einstellungen").Range(E_ORDNER).Value))
    If p = "" Then
        basis = MappenOrdner()
        If basis = "" And anlegen Then basis = OrdnerWaehlen()
        If basis = "" Then
            ImportOrdner = ""
            Exit Function
        End If
        p = basis & PS() & "Import"
    End If
    If Right$(p, 1) = PS() Then p = Left$(p, Len(p) - 1)
    If anlegen Then
        If Dir(p, vbDirectory) = "" Then MkDir p
        If Dir(p & PS() & "erledigt", vbDirectory) = "" Then MkDir p & PS() & "erledigt"
    End If
    ImportOrdner = p
End Function

' Lokaler Ordner der Mappe - auch wenn sie in OneDrive/SharePoint liegt (ThisWorkbook.Path ist dann eine URL)
Private Function MappenOrdner() As String
    Dim p As String
    p = ThisWorkbook.Path
    If p = "" Then Exit Function
    If LCase$(Left$(p, 4)) <> "http" Then
        MappenOrdner = p
        Exit Function
    End If
    MappenOrdner = OneDriveLokal(p)
End Function

Public Function OneDriveLokal(ByVal url As String) As String
    Dim teile() As String, roots As Variant, root As Variant
    Dim i As Long, j As Long
    Dim rest As String, kandidat As String
    url = UrlDecode(url)
    url = Replace(url, "\", "/")
    teile = Split(Mid$(url, InStr(url, "//") + 2), "/")
    roots = Array(Environ$("OneDriveCommercial"), Environ$("OneDriveConsumer"), Environ$("OneDrive"))
    On Error Resume Next
    For Each root In roots
        If CStr(root) <> "" Then
            For i = 1 To UBound(teile)
                rest = ""
                For j = i To UBound(teile)
                    rest = rest & PS() & teile(j)
                Next j
                kandidat = CStr(root) & rest
                If Dir(kandidat & PS() & ThisWorkbook.Name) <> "" Then
                    OneDriveLokal = kandidat
                    Exit Function
                End If
            Next i
            If Dir(CStr(root) & PS() & ThisWorkbook.Name) <> "" Then
                OneDriveLokal = CStr(root)
                Exit Function
            End If
        End If
    Next root
End Function

Private Function UrlDecode(ByVal s As String) As String
    Dim i As Long
    Dim out As String, h As String
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 1) = "%" And i + 2 <= Len(s) Then
            h = Mid$(s, i + 1, 2)
            out = out & Chr$(CLng("&H" & h))
            i = i + 3
        Else
            out = out & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    out = Replace(out, Chr$(195) & Chr$(164), ChrW(228))
    out = Replace(out, Chr$(195) & Chr$(182), ChrW(246))
    out = Replace(out, Chr$(195) & Chr$(188), ChrW(252))
    out = Replace(out, Chr$(195) & Chr$(132), ChrW(196))
    out = Replace(out, Chr$(195) & Chr$(150), ChrW(214))
    out = Replace(out, Chr$(195) & Chr$(156), ChrW(220))
    out = Replace(out, Chr$(195) & Chr$(159), ChrW(223))
    UrlDecode = out
End Function

' Fallback: Ordner einmal auswählen lassen und in Einstellungen merken
Private Function OrdnerWaehlen() As String
    Dim fd As Object
    Dim p As String
    On Error Resume Next
    MsgBox "Der Speicherort der Mappe konnte nicht ermittelt werden." & vbLf & _
           "Bitte wähle den Ordner, in dem die Mappe liegt - dort wird der Ordner „Import“ angelegt.", vbInformation, "George-Import"
    Set fd = Application.FileDialog(4)
    fd.Title = "Ordner der Mappe auswählen"
    If fd.Show = -1 Then p = fd.SelectedItems(1)
    If p <> "" Then
        Blatt("shEinst", "Einstellungen").Range(E_ORDNER).Value = p & PS() & "Import"
        OrdnerWaehlen = p
    End If
End Function

Private Function DateiName(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If InStrRev(p, "/") > i Then i = InStrRev(p, "/")
    If InStrRev(p, ":") > i And Application.PathSeparator = ":" Then i = InStrRev(p, ":")
    DateiName = Mid$(p, i + 1)
End Function

Public Function SStr(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        SStr = ""
    Else
        SStr = CStr(v)
    End If
End Function

' Text so schreiben, dass Excel ihn nicht als Formel, Zahl oder Datum deutet
Public Function TxtZelle(ByVal s As String) As Variant
    If s = "" Then
        TxtZelle = Empty
    ElseIf InStr("=+-@'", Left$(s, 1)) > 0 Or IsNumeric(s) Or IsDate(s) Then
        TxtZelle = "'" & s
    Else
        TxtZelle = s
    End If
End Function

Private Function InSet(ByVal c As Collection, ByVal k As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = c("x" & k)
    InSet = (Err.Number = 0)
    Err.Clear
End Function

' Kleinbuchstaben, Umlaute aufgelöst, alles außer Buchstaben/Ziffern -> Leerzeichen
Public Function Fold(ByVal s As String) As String
    Dim i As Long, code As Long
    Dim out As String, ch As String
    s = LCase$(s)
    s = Replace(s, ChrW(228), "ae")
    s = Replace(s, ChrW(246), "oe")
    s = Replace(s, ChrW(252), "ue")
    s = Replace(s, ChrW(223), "ss")
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        If (code >= 97 And code <= 122) Or (code >= 48 And code <= 57) Then
            out = out & ch
        ElseIf code = 39 Or code = 8217 Then
            ' Apostroph weglassen
        ElseIf code >= 224 And code <= 255 Then
            out = out & Mid$("aaaaaaaceeeeiiiidnooooo/ouuuuyty", code - 223, 1)
        Else
            out = out & " "
        End If
    Next i
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    Fold = Trim$(out)
End Function

Public Function NormKonto(ByVal s As String) As String
    Dim i As Long
    Dim ch As String, out As String
    s = UCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then out = out & ch
    Next i
    NormKonto = out
End Function

Public Function GleichesKonto(ByVal a As String, ByVal b As String) As Boolean
    If Len(a) < 4 Or Len(b) < 4 Then Exit Function
    GleichesKonto = (Right$(a, Len(b)) = b) Or (Right$(b, Len(a)) = a)
End Function

Private Function TxBetrag(ByVal tx As Object) As Double
    Dim a As Object
    Dim p As Long
    Set a = JObj(tx, "amount")
    If a Is Nothing Then Exit Function
    p = CLng(JNum(a, "precision"))
    TxBetrag = Round(JNum(a, "value") / (10 ^ p), 2)
End Function

Private Function TxDatum(ByVal tx As Object) As Variant
    Dim s As String
    s = JStr(tx, "booking")
    If s = "" Then s = JStr(tx, "valuation")
    If s Like "####-##-##*" Then
        TxDatum = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
    ElseIf s Like "##.##.####*" Then
        TxDatum = DateSerial(CInt(Mid$(s, 7, 4)), CInt(Mid$(s, 4, 2)), CInt(Left$(s, 2)))
    Else
        TxDatum = Empty
    End If
End Function

Private Function TxReferenz(ByVal tx As Object) As String
    Dim s As String
    Dim z As String
    s = Trim$(JStr(tx, "reference"))
    z = Trim$(JStr(tx, "receiverReference"))
    If z <> "" And InStr(1, s, z, vbTextCompare) = 0 Then
        If s = "" Then s = z Else s = s & "  ·  " & z
    End If
    TxReferenz = s
End Function

Private Function TxKey(ByVal tx As Object, ByVal konto As String) As String
    Dim a As Object
    Dim s As String
    Set a = JObj(tx, "amount")
    s = LCase$(konto) & "|" & Left$(JStr(tx, "booking"), 10) & "|"
    If Not a Is Nothing Then s = s & JStr(a, "value")
    s = s & "|" & Left$(Fold(JStr(tx, "partnerName") & " " & JStr(tx, "reference")), 80)
    If JStr(tx, "referenceNumber") <> "" Then s = s & "|" & JStr(tx, "referenceNumber")
    TxKey = s
End Function

' Macht aus 'POS 12,40 0001 K1 02.09. 10:15 O TRAFIK SCHOENHERR' etwas Lesbares
Public Function SauberName(ByVal roh As String) As String
    Dim w As Variant
    Dim s As String, u As String, out As String
    Dim i As Long, digits As Long
    roh = Replace(Replace(roh, vbLf, " "), vbCr, " ")
    roh = Replace(roh, "Apple Pay", " ", , , vbTextCompare)
    roh = Replace(roh, "Google Pay", " ", , , vbTextCompare)
    For Each w In Split(roh, " ")
        s = Trim$(CStr(w))
        u = UCase$(s)
        Do While Len(s) > 1 And InStr("*#+", Left$(s, 1)) > 0
            s = Mid$(s, 2)
        Loop
        u = UCase$(s)
        If s = "" Then GoTo Weiter
        If InStr(" POS ATM O NFC E-COM E-COMM ECOM BEZAHLUNG KARTE KARTENZAHLUNG MC DANKT AT IE DE HR IT NL LU SE ", " " & u & " ") > 0 Then GoTo Weiter
        If u Like "K#" Then GoTo Weiter
        digits = 0
        For i = 1 To Len(s)
            If Mid$(s, i, 1) Like "#" Then digits = digits + 1
        Next i
        If digits > 0 And (digits >= 4 Or InStr(s, ".") > 0 Or InStr(s, ",") > 0 Or InStr(s, ":") > 0 Or InStr(s, "/") > 0 Or digits = Len(s)) Then GoTo Weiter
        out = out & " " & s
Weiter:
    Next w
    out = Trim$(out)
    Do While Len(out) > 1 And InStr(",;.:-/*", Right$(out, 1)) > 0
        out = Trim$(Left$(out, Len(out) - 1))
    Loop
    If out = "" Then out = Trim$(roh)
    If out = UCase$(out) And out <> LCase$(out) Then
        s = ""
        For Each w In Split(out, " ")
            If Len(w) > 3 Then
                s = s & " " & UCase$(Left$(w, 1)) & LCase$(Mid$(w, 2))
            Else
                s = s & " " & w
            End If
        Next w
        out = Trim$(s)
    End If
    SauberName = Left$(out, 40)
End Function
