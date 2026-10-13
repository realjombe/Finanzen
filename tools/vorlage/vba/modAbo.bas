Attribute VB_Name = "modAbo"
Option Explicit
' ===========================================================================
'  Abo-Erkennung - ausschließlich aus den importierten George-Buchungen
'
'  1) Gleicher Abbucher: Buchungen mit derselben Mandatsreferenz, Gläubiger-ID,
'     IBAN oder demselben Namen, die in einem regelmäßigen Rhythmus kommen
'     (wöchentlich ... jährlich). Fester Betrag -> „Abo“ / „Teilzahlung“ /
'     „Dauerauftrag“, schwankender Betrag -> „Abo (variabel)“.
'  2) Gleicher Betrag: derselbe Betrag jeden Monat (bzw. in festem Rhythmus),
'     auch wenn der Empfängertext wechselt -> Teilzahlung / Dauerauftrag / Abo.
'
'  Einstieg: AbosErkennen (Schaltfläche/Menüband) - wird auch nach jedem
'  Import automatisch ausgeführt.
' ===========================================================================

' ---- Blatt Buchungen (Spalten)
Public Const B_FIRST As Long = 6
Public Const C_DATUM As Long = 1
Public Const C_KONTO As Long = 2
Public Const C_PARTNER As Long = 3
Public Const C_REF As Long = 4
Public Const C_BETRAG As Long = 5
Public Const C_ZART As Long = 6
Public Const C_ABO As Long = 7
Public Const C_ART As Long = 8
Public Const C_NOTIZ As Long = 9
Public Const C_MONAT As Long = 10
Public Const C_IBAN As Long = 11
Public Const C_GLAEUB As Long = 12
Public Const C_MANDAT As Long = 13
Public Const C_KEY As Long = 14
Public Const N_COLS As Long = 14

' ---- Blatt Abos (Spalten)
Private Const A_FIRST As Long = 6
Private Const A_MAX As Long = 205
Private Const A_NAME As Long = 1
Private Const A_ART As Long = 2
Private Const A_STATUS As Long = 3
Private Const A_RHY As Long = 4
Private Const A_BETRAG As Long = 5
Private Const A_MITTEL As Long = 6
Private Const A_MONAT As Long = 7
Private Const A_JAHR As Long = 8
Private Const A_ANZ As Long = 9
Private Const A_ERSTE As Long = 10
Private Const A_LETZTE As Long = 11
Private Const A_NAECHSTE As Long = 12
Private Const A_KONTO As Long = 13
Private Const A_ZART As Long = 14
Private Const A_GRUND As Long = 15
Private Const A_RATEN As Long = 16
Private Const A_REF As Long = 17
Private Const A_KAT As Long = 18
Private Const A_IGN As Long = 19
Private Const A_KEY As Long = 20
Private Const A_NCOLS As Long = 20

' ---- Blatt Übersicht (Konten)
Public Const K_FIRST As Long = 17
Public Const K_LAST As Long = 24

' ---- Blatt Monatsübersicht
Private Const M_FIRST As Long = 6
Private Const M_LAST As Long = 41

' ---- Blatt Einstellungen
Private Const E_MIN As String = "B6"
Private Const E_TOLPCT As String = "B7"
Private Const E_TOLTAGE As String = "B8"
Private Const E_MINBETRAG As String = "B9"
Private Const E_KWRATE As String = "B10"
Private Const E_KWUMB As String = "B11"
Private Const E_KWBAR As String = "B12"

' ---- Buchungen (Arbeitsdaten)
Private nB As Long
Private bDat() As Double, bAmt() As Double, bCent() As Long
Private bKonto() As String, bName() As String, bRef() As String, bZart() As String
Private bIban() As String, bGl() As String, bMand() As String
Private bKey() As String, bText() As String
Private bCand() As Boolean, bAbo() As Long, bArt() As String

' ---- Abos (Ergebnis)
Private nA As Long
Private aName() As String, aArt() As String, aStatus() As String, aRhy() As String
Private aBetrag() As Double, aMittel() As Double, aMonat() As Double, aJahr() As Double
Private aAnz() As Long, aErste() As Double, aLetzte() As Double, aNaechste() As Double
Private aKonto() As String, aZart() As String, aGrund() As String, aRaten() As String
Private aRef() As String, aKey() As String, aSort() As Double
Private aP() As Double, aCent() As Long, aFest() As Boolean, aDel() As Boolean

' ---- Einstellungen / Hilfsdaten
Private minAnz As Long, tolPct As Double, tolTage As Double, minBetrag As Double
Private kwRate As String, kwUmb As String, kwBar As String
Private ownNum() As String, nOwn As Long
Private kTyp As Collection
Private kStand As Collection
Private monCent As Collection
Private monDbg As String
Private sMode As Long

' ===========================================================================
'  Einstiegspunkte
' ===========================================================================
Public Sub AbosErkennen()
    ErkenneAbos False
End Sub

Public Sub Ribbon_Abos(control As Object)
    AbosErkennen
End Sub

' Für automatisierte Tests (ohne Dialoge)
Public Function AbosTest() As String
    AbosTest = ErkenneAbos(True)
End Function

' Diagnose nach einem Lauf: Beträge mit monatlicher Serie, Test des Collection-Zugriffs
Public Function AboDiagnose(ByVal cent As Long) As String
    AboDiagnose = "monCent:" & monDbg & " | HatKey(" & cent & ")=" & HatKey(monCent, "x" & cent) & " | HatKey(fehlt)=" & HatKey(monCent, "xgibtesnicht")
End Function

Public Function ErkenneAbos(ByVal still As Boolean) As String
    Dim wsB As Worksheet, wsA As Worksheet
    Dim msg As String
    Dim oldCalc As Long
    Dim i As Long, nAbo As Long, nTeil As Long, nDauer As Long, nVar As Long
    Dim summe As Double
    On Error GoTo Fehler
    Set wsB = Blatt("shBuchungen", "Buchungen")
    Set wsA = Blatt("shAbos", "Abos")
    On Error Resume Next
    Application.ScreenUpdating = False
    oldCalc = Application.Calculation
    Application.Calculation = -4135
    Application.StatusBar = "Abos werden erkannt ..."
    On Error GoTo Fehler

    LadeEinstellungen
    LadeKonten
    LadeBuchungen wsB
    nA = 0
    If nB > 0 Then
        ReDim aName(1 To nB + 1): ReDim aArt(1 To nB + 1): ReDim aStatus(1 To nB + 1): ReDim aRhy(1 To nB + 1)
        ReDim aBetrag(1 To nB + 1): ReDim aMittel(1 To nB + 1): ReDim aMonat(1 To nB + 1): ReDim aJahr(1 To nB + 1)
        ReDim aAnz(1 To nB + 1): ReDim aErste(1 To nB + 1): ReDim aLetzte(1 To nB + 1): ReDim aNaechste(1 To nB + 1)
        ReDim aKonto(1 To nB + 1): ReDim aZart(1 To nB + 1): ReDim aGrund(1 To nB + 1): ReDim aRaten(1 To nB + 1)
        ReDim aRef(1 To nB + 1): ReDim aKey(1 To nB + 1): ReDim aSort(1 To nB + 1)
        ReDim aP(1 To nB + 1): ReDim aCent(1 To nB + 1): ReDim aFest(1 To nB + 1): ReDim aDel(1 To nB + 1)
        KlassifiziereBuchungen
        FindeMonatsBetraege
        FindeAbbucherSerien
        FindeBetragSerien
        MergeFolgeserien
    End If
    SchreibeErgebnis wsB, wsA

    For i = 1 To nA
        If aDel(i) Then
            ' zusammengeführt
        ElseIf aStatus(i) = "aktiv" Or aStatus(i) = "überfällig" Then
            If Left$(aArt(i), 3) = "Abo" Then
                nAbo = nAbo + 1
                If aArt(i) <> "Abo" Then nVar = nVar + 1
            ElseIf aArt(i) = "Teilzahlung" Then
                nTeil = nTeil + 1
            ElseIf aArt(i) = "Dauerauftrag" Then
                nDauer = nDauer + 1
            End If
            summe = summe + aMonat(i)
        End If
    Next i
    msg = AnzahlAbos() & " regelmäßige Zahlungen erkannt (" & nB & " Buchungen)." & vbLf & vbLf & _
          "   aktive Abos: " & nAbo & IIf(nVar > 0, " (davon " & nVar & " mit variablem Betrag)", "") & vbLf & _
          "   laufende Teilzahlungen: " & nTeil & vbLf & _
          "   Daueraufträge: " & nDauer & vbLf & vbLf & _
          "   fixe Belastung: " & Format(summe, "#,##0.00") & " € pro Monat" & vbLf & vbLf & _
          "Details im Blatt „Abos“. Dort kannst du Kategorie eintragen oder Zeilen mit „x“ ignorieren."
    wsA.Range("A3").Value = "Zuletzt erkannt: " & Format(Now, "dd.mm.yyyy hh:mm") & "  -  " & nAbo & " Abos, " & nTeil & " Teilzahlungen, " & nDauer & " Daueraufträge aktiv"

    On Error Resume Next
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True
    Application.StatusBar = False
    On Error GoTo 0
    If Not still Then MsgBox msg, vbInformation, "Abo-Erkennung"
    ErkenneAbos = Replace(msg, vbLf, " | ")
    Exit Function

Fehler:
    Dim em As String
    em = "Fehler bei der Abo-Erkennung: " & Err.Description & " (" & Err.Number & ")"
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If Not still Then MsgBox em, vbExclamation, "Abo-Erkennung"
    ErkenneAbos = "FEHLER: " & em
End Function

' ===========================================================================
'  Daten laden
' ===========================================================================
Private Sub LadeEinstellungen()
    Dim ws As Worksheet
    Set ws = Blatt("shEinst", "Einstellungen")
    minAnz = 3: tolPct = 0.01: tolTage = 7: minBetrag = 1
    On Error Resume Next
    minAnz = CLng(ws.Range(E_MIN).Value)
    tolPct = CDbl(ws.Range(E_TOLPCT).Value)
    tolTage = CDbl(ws.Range(E_TOLTAGE).Value)
    minBetrag = CDbl(ws.Range(E_MINBETRAG).Value)
    kwRate = SStr(ws.Range(E_KWRATE).Value)
    kwUmb = SStr(ws.Range(E_KWUMB).Value)
    kwBar = SStr(ws.Range(E_KWBAR).Value)
    On Error GoTo 0
    If minAnz < 2 Then minAnz = 2
    If tolPct <= 0 Then tolPct = 0.01
    If tolPct > 1 Then tolPct = tolPct / 100
    If tolTage <= 0 Then tolTage = 7
    If minBetrag < 0 Then minBetrag = 0
    If kwRate = "" Then kwRate = "rate;raten;ratenzahlung;ratenkauf;teilzahlung;teilbetrag;finanzierung;kreditrate;kreditrückzahlung;tilgung;leasing;abzahlung"
    If kwUmb = "" Then kwUmb = "kreditkartenabrechnung;kartenabrechnung;kreditkartenrechnung;kartenausgleich;umbuchung;eigenübertrag;eigenüberweisung"
    If kwBar = "" Then kwBar = "bargeld;bankomat;behebung;atm;automat;auszahlung"
End Sub

' Konten aus der Übersicht: Typ je Konto, eigene Kontonummern
Private Sub LadeKonten()
    Dim wsO As Worksheet
    Dim r As Long
    Dim nm As String, ib As String
    Set wsO = Blatt("shUebersicht", "Übersicht")
    Set kTyp = New Collection
    nOwn = 0
    ReDim ownNum(1 To K_LAST - K_FIRST + 2)
    For r = K_FIRST To K_LAST
        nm = Trim$(SStr(wsO.Cells(r, 1).Value))
        ib = NormKonto(SStr(wsO.Cells(r, 3).Value))
        If nm <> "" Then
            If Not HatKey(kTyp, "x" & LCase$(nm)) Then kTyp.Add Trim$(SStr(wsO.Cells(r, 2).Value)), "x" & LCase$(nm)
        End If
        If Len(ib) >= 4 Then nOwn = nOwn + 1: ownNum(nOwn) = ib
    Next r
End Sub

Private Function KontoTyp(ByVal konto As String) As String
    KontoTyp = CStr(HoleWert(kTyp, "x" & LCase$(Trim$(konto)), ""))
End Function

Private Function IstEigenesKonto(ByVal iban As String) As Boolean
    Dim i As Long
    Dim n As String
    n = NormKonto(iban)
    If Len(n) < 4 Then Exit Function
    For i = 1 To nOwn
        If GleichesKonto(n, ownNum(i)) Then IstEigenesKonto = True: Exit Function
    Next i
End Function

Private Sub LadeBuchungen(ByVal wsB As Worksheet)
    Dim lastRow As Long, r As Long, i As Long
    Dim data As Variant
    Dim v As Variant
    nB = 0
    lastRow = LetzteZeile(wsB)
    If lastRow < B_FIRST Then Exit Sub
    data = wsB.Range(wsB.Cells(B_FIRST, 1), wsB.Cells(lastRow, N_COLS)).Value
    ReDim bDat(1 To lastRow): ReDim bAmt(1 To lastRow): ReDim bCent(1 To lastRow)
    ReDim bKonto(1 To lastRow): ReDim bName(1 To lastRow): ReDim bRef(1 To lastRow): ReDim bZart(1 To lastRow)
    ReDim bIban(1 To lastRow): ReDim bGl(1 To lastRow): ReDim bMand(1 To lastRow)
    ReDim bKey(1 To lastRow): ReDim bText(1 To lastRow)
    ReDim bCand(1 To lastRow): ReDim bAbo(1 To lastRow): ReDim bArt(1 To lastRow)
    Set kStand = New Collection
    For r = 1 To UBound(data, 1)
        v = data(r, C_DATUM)
        If IsDate(v) And IsNumeric(data(r, C_BETRAG)) Then
            nB = nB + 1
            bDat(nB) = CDbl(CDate(v))
            bAmt(nB) = CDbl(data(r, C_BETRAG))
            bCent(nB) = CLng(Fix(Abs(bAmt(nB)) * 100 + 0.5))
            bKonto(nB) = Trim$(SStr(data(r, C_KONTO)))
            bName(nB) = Trim$(SStr(data(r, C_PARTNER)))
            bRef(nB) = Trim$(SStr(data(r, C_REF)))
            bZart(nB) = Trim$(SStr(data(r, C_ZART)))
            bIban(nB) = Trim$(SStr(data(r, C_IBAN)))
            bGl(nB) = Trim$(SStr(data(r, C_GLAEUB)))
            bMand(nB) = Trim$(SStr(data(r, C_MANDAT)))
            bText(nB) = " " & Fold(bName(nB) & " " & bRef(nB)) & " "
            bAbo(nB) = 0
            bCand(nB) = False
            MerkeStand bKonto(nB), bDat(nB)
        End If
    Next r
End Sub

' Datenstand je Konto = jüngste Buchung des Kontos
Private Sub MerkeStand(ByVal konto As String, ByVal d As Double)
    Dim k As String
    k = "x" & LCase$(konto)
    If Not HatKey(kStand, k) Then
        kStand.Add d, k
    ElseIf d > CDbl(HoleWert(kStand, k, 0)) Then
        kStand.Remove k
        kStand.Add d, k
    End If
End Sub

Private Function Datenstand(ByVal konto As String) As Double
    Datenstand = CDbl(HoleWert(kStand, "x" & LCase$(konto), CDbl(Date)))
End Function

' ===========================================================================
'  Grundeinordnung je Buchung: Einnahme / Umbuchung / Ausgabe / Ignorieren
' ===========================================================================
Private Sub KlassifiziereBuchungen()
    Dim i As Long
    Dim typ As String
    For i = 1 To nB
        typ = KontoTyp(bKonto(i))
        bCand(i) = False
        If bCent(i) = 0 Then
            bArt(i) = "Ignorieren"
        ElseIf IstEigenesKonto(bIban(i)) Or HatStichwort(bText(i), kwUmb) Then
            bArt(i) = "Umbuchung"
        ElseIf typ = "Sparkonto" Or typ = "Bausparkonto" Then
            If InStr(bText(i), "zins") > 0 Then
                bArt(i) = IIf(bAmt(i) > 0, "Einnahme", "Ausgabe")
            ElseIf InStr(bText(i), " kest ") > 0 Or InStr(bText(i), "kapitalertrag") > 0 Then
                bArt(i) = "Ausgabe"
            Else
                bArt(i) = "Umbuchung"
            End If
        ElseIf bAmt(i) > 0 Then
            bArt(i) = IIf(typ = "Kreditkarte", "Umbuchung", "Einnahme")
        ElseIf bZart(i) = "Bargeld" Or HatStichwort(bText(i), kwBar) Then
            bArt(i) = "Ausgabe"
        Else
            bArt(i) = "Ausgabe"
            bCand(i) = (Abs(bAmt(i)) >= minBetrag)
        End If
        bKey(i) = AbbucherKey(i)
    Next i
End Sub

' Schlüssel „gleicher Abbucher“: Mandat > Gläubiger-ID > IBAN > Name
Private Function AbbucherKey(ByVal i As Long) As String
    Dim k As String
    If bMand(i) <> "" Then
        AbbucherKey = "M:" & UCase$(bMand(i))
    ElseIf bGl(i) <> "" Then
        AbbucherKey = "C:" & UCase$(bGl(i))
    ElseIf Len(NormKonto(bIban(i))) >= 8 Then
        AbbucherKey = "I:" & NormKonto(bIban(i))
    Else
        k = NameKey(bName(i))
        If k = "" Then k = NameKey(bRef(i))
        AbbucherKey = "N:" & k
    End If
End Function

' Erstes aussagekräftiges Wort des Namens (ohne Zahlen, Rechtsformen, Orte ...)
Public Function NameKey(ByVal s As String) As String
    Dim w As Variant
    Dim t As String
    Dim i As Long, hasDigit As Boolean
    Const STOPW As String = " dankt gmbh ag kg ltd com www wien graz linz salzburg austria at de nl ie lu se ch sarl bv co inc the pos bezahlung karte kartenzahlung mc visa und der die das von zahlung mein meine apple pay google nfc "
    For Each w In Split(Fold(s), " ")
        t = CStr(w)
        If Len(t) >= 3 And InStr(STOPW, " " & t & " ") = 0 Then
            hasDigit = False
            For i = 1 To Len(t)
                If Mid$(t, i, 1) Like "#" Then hasDigit = True: Exit For
            Next i
            If Not hasDigit Then
                NameKey = t
                Exit Function
            End If
        End If
    Next w
    NameKey = Left$(Fold(s), 12)
End Function

' Stichwortliste (;-getrennt): kurze Wörter nur als ganzes Wort, lange auch als Teilstring
Public Function HatStichwort(ByVal ft As String, ByVal liste As String) As Boolean
    Dim w As Variant
    Dim k As String
    For Each w In Split(liste, ";")
        k = Fold(CStr(w))
        If k <> "" Then
            If InStr(k, " ") > 0 Then
                If InStr(ft, " " & k & " ") > 0 Or InStr(ft, k) > 0 Then HatStichwort = True: Exit Function
            ElseIf Len(k) <= 5 Then
                If InStr(ft, " " & k & " ") > 0 Then HatStichwort = True: Exit Function
            Else
                If InStr(ft, k) > 0 Then HatStichwort = True: Exit Function
            End If
        End If
    Next w
End Function

' ===========================================================================
'  1) Serien je Abbucher
' ===========================================================================
Private Sub FindeAbbucherSerien()
    Dim idx() As Long, grp() As Long, cl() As Long
    Dim n As Long, i As Long, g0 As Long, g1 As Long, m As Long, j As Long
    Dim P As Double, rhy As String, fak As Double, frac As Double
    Dim c0 As Long, c1 As Long, cnt As Long
    n = 0
    ReDim idx(1 To nB)
    For i = 1 To nB
        If bCand(i) Then n = n + 1: idx(n) = i
    Next i
    If n = 0 Then Exit Sub
    ReDim Preserve idx(1 To n)
    sMode = 1
    SortIdx idx, 1, n
    g0 = 1
    Do While g0 <= n
        g1 = g0
        Do While g1 < n
            If bKey(idx(g1 + 1)) <> bKey(idx(g0)) Then Exit Do
            g1 = g1 + 1
        Loop
        m = g1 - g0 + 1
        If m >= 2 Then
            ReDim grp(1 To m)
            For j = 1 To m: grp(j) = idx(g0 + j - 1): Next j
            ' a) ganze Gruppe regelmäßig? (eine Buchung je Periode)
            If BewerteSerie(grp, m, True, False, P, rhy, fak, frac) Then
                ' Teilserie einer monatlichen Zahlung mit gleichem Betrag (Empfängertext wechselt)? -> überlassen wir Schritt 2
                If Not (P >= 45 And NurMonatsBetrag(grp, m)) Then NeuesAbo grp, m, P, rhy, fak, "gleicher Abbucher"
            Else
                ' b) sonst: je Betrag getrennt (z. B. zwei Abos beim selben Anbieter, Jahresabo neben Einzelkäufen)
                sMode = 2
                SortIdx grp, 1, m
                c0 = 1
                Do While c0 <= m
                    c1 = c0
                    Do While c1 < m
                        If bCent(grp(c1 + 1)) <> bCent(grp(c0)) Then Exit Do
                        c1 = c1 + 1
                    Loop
                    cnt = c1 - c0 + 1
                    If cnt >= 2 Then
                        ReDim cl(1 To cnt)
                        For j = 1 To cnt: cl(j) = grp(c0 + j - 1): Next j
                        If BewerteSerie(cl, cnt, True, False, P, rhy, fak, frac) Then
                            If Not (P >= 45 And NurMonatsBetrag(cl, cnt)) Then NeuesAbo cl, cnt, P, rhy, fak, "gleicher Abbucher und Betrag"
                        End If
                    End If
                    c0 = c1 + 1
                Loop
                sMode = 1
            End If
        End If
        g0 = g1 + 1
    Loop
End Sub

' ===========================================================================
'  2) Serien mit gleichem Betrag (Empfängertext darf wechseln)
' ===========================================================================
Private Sub FindeBetragSerien()
    Dim idx() As Long, grp() As Long
    Dim n As Long, i As Long, g0 As Long, g1 As Long, m As Long, j As Long
    Dim P As Double, rhy As String, fak As Double, frac As Double
    n = 0
    ReDim idx(1 To nB)
    For i = 1 To nB
        If bCand(i) And bAbo(i) = 0 Then n = n + 1: idx(n) = i
    Next i
    If n = 0 Then Exit Sub
    ReDim Preserve idx(1 To n)
    sMode = 2
    SortIdx idx, 1, n
    g0 = 1
    Do While g0 <= n
        g1 = g0
        Do While g1 < n
            If bCent(idx(g1 + 1)) <> bCent(idx(g0)) Then Exit Do
            g1 = g1 + 1
        Loop
        m = g1 - g0 + 1
        If m >= minAnz Then
            ReDim grp(1 To m)
            For j = 1 To m: grp(j) = idx(g0 + j - 1): Next j
            If BewerteSerie(grp, m, False, True, P, rhy, fak, frac) Then
                NeuesAbo grp, m, P, rhy, fak, "gleicher Betrag " & rhy
            End If
        End If
        g0 = g1 + 1
    Loop
    sMode = 1
End Sub

' Vorab: Beträge, die über alle Kandidaten hinweg eine strenge monatliche (oder kürzere) Serie bilden
Private Sub FindeMonatsBetraege()
    Dim idx() As Long, grp() As Long
    Dim n As Long, i As Long, g0 As Long, g1 As Long, m As Long, j As Long
    Dim P As Double, rhy As String, fak As Double, frac As Double
    Set monCent = New Collection
    monDbg = ""
    n = 0
    ReDim idx(1 To nB)
    For i = 1 To nB
        If bCand(i) Then n = n + 1: idx(n) = i
    Next i
    If n = 0 Then Exit Sub
    ReDim Preserve idx(1 To n)
    sMode = 2
    SortIdx idx, 1, n
    g0 = 1
    Do While g0 <= n
        g1 = g0
        Do While g1 < n
            If bCent(idx(g1 + 1)) <> bCent(idx(g0)) Then Exit Do
            g1 = g1 + 1
        Loop
        m = g1 - g0 + 1
        If m >= minAnz And m >= 3 Then
            ReDim grp(1 To m)
            For j = 1 To m: grp(j) = idx(g0 + j - 1): Next j
            If BewerteSerie(grp, m, False, True, P, rhy, fak, frac) Then
                If P <= 35 Then
                    If Not HatKey(monCent, "x" & bCent(idx(g0))) Then monCent.Add True, "x" & bCent(idx(g0))
                    monDbg = monDbg & " " & bCent(idx(g0))
                End If
            End If
        End If
        g0 = g1 + 1
    Loop
    sMode = 1
End Sub

' Alle Buchungen der Serie haben denselben Betrag, und dieser Betrag bildet insgesamt eine monatliche Serie
Private Function NurMonatsBetrag(ByRef idx() As Long, ByVal n As Long) As Boolean
    Dim j As Long
    Dim v As Variant
    For j = 2 To n
        If bCent(idx(j)) <> bCent(idx(1)) Then Exit Function
    Next j
    NurMonatsBetrag = HatKey(monCent, "x" & bCent(idx(1)))
End Function

' Zwei feste Serien mit gleichem Betrag und Rhythmus, die nahtlos aufeinander folgen
' (Empfänger umbenannt, neue IBAN, Wechsel Lastschrift -> Karte) -> eine Zahlung
Private Sub MergeFolgeserien()
    Dim i As Long, j As Long, k As Long, n As Long
    Dim gap As Double, tol As Double
    Dim u() As Long
    Dim erst As Long, zweit As Long
    i = 1
    Do While i <= nA
        If Not aDel(i) And aFest(i) Then
            j = i + 1
            Do While j <= nA
                If Not aDel(j) And aFest(j) And aRhy(j) = aRhy(i) And GleicherBetrag(aCent(i), aCent(j)) Then
                    If aErste(j) > aLetzte(i) Then
                        gap = aErste(j) - aLetzte(i): erst = i: zweit = j
                    ElseIf aErste(i) > aLetzte(j) Then
                        gap = aErste(i) - aLetzte(j): erst = j: zweit = i
                    Else
                        gap = -1
                    End If
                    If aP(i) < 20 Then tol = 2 Else tol = tolTage + aP(i) * 0.08
                    If gap >= aP(i) - tol And gap <= 2 * aP(i) + tol Then
                        n = 0
                        ReDim u(1 To aAnz(i) + aAnz(j))
                        For k = 1 To nB
                            If bAbo(k) = i Or bAbo(k) = j Then n = n + 1: u(n) = k
                        Next k
                        NeuesAbo u, n, aP(i), aRhy(i), MonatsFaktor(aRhy(i)), "gleicher Betrag, Empfänger/Zahlungsweg gewechselt (" & aName(erst) & " -> " & aName(zweit) & ")"
                        aDel(i) = True: aDel(j) = True
                        Exit Do
                    End If
                End If
                j = j + 1
            Loop
        End If
        i = i + 1
    Loop
End Sub

Private Function MonatsFaktor(ByVal rhy As String) As Double
    Select Case rhy
        Case "wöchentlich": MonatsFaktor = 52 / 12
        Case "14-tägig": MonatsFaktor = 26 / 12
        Case "2-monatlich": MonatsFaktor = 0.5
        Case "vierteljährlich": MonatsFaktor = 1 / 3
        Case "halbjährlich": MonatsFaktor = 1 / 6
        Case "jährlich": MonatsFaktor = 1 / 12
        Case Else: MonatsFaktor = 1
    End Select
End Function

Private Function AnzahlAbos() As Long
    Dim i As Long
    For i = 1 To nA
        If Not aDel(i) Then AnzahlAbos = AnzahlAbos + 1
    Next i
End Function

' ===========================================================================
'  Regelmäßigkeit einer Serie (idx nach Datum aufsteigend)
' ===========================================================================
Private Function BewerteSerie(ByRef idx() As Long, ByVal n As Long, ByVal mitWoche As Boolean, ByVal streng As Boolean, _
                              ByRef P As Double, ByRef rhy As String, ByRef fak As Double, ByRef frac As Double) As Boolean
    Dim d() As Double
    Dim i As Long, k As Long, best As Long, reg As Long, minN As Long
    Dim med As Double, bestD As Double, tol As Double, dd As Double
    Dim per As Variant, nam As Variant, fk As Variant
    BewerteSerie = False
    frac = 0
    If n < 2 Then Exit Function
    SortNachDatum idx, n
    ReDim d(1 To n - 1)
    For i = 1 To n - 1
        d(i) = bDat(idx(i + 1)) - bDat(idx(i))
    Next i
    med = Median(d, n - 1)
    per = Array(7#, 14#, 30.4375, 60.875, 91.3125, 182.625, 365.25)
    nam = Array("wöchentlich", "14-tägig", "monatlich", "2-monatlich", "vierteljährlich", "halbjährlich", "jährlich")
    fk = Array(52 / 12, 26 / 12, 1, 0.5, 1 / 3, 1 / 6, 1 / 12)
    bestD = 1E+99
    best = 2
    For k = IIf(mitWoche, 0, 2) To 6
        dd = Abs(med - per(k)) / per(k)
        If dd < bestD Then bestD = dd: best = k
    Next k
    P = per(best): rhy = CStr(nam(best)): fak = CDbl(fk(best))
    If P < 20 Then tol = 2 Else tol = tolTage + P * 0.08
    If n = 2 Then tol = tolTage + 8                   ' nur zwei Buchungen (Jahresabo): enger prüfen
    If Abs(med - P) > tol Then Exit Function
    reg = 0
    For i = 1 To n - 1
        If Abs(d(i) - P) <= tol Then
            reg = reg + 1
        ElseIf streng Then
            If d(i) < P * 0.5 Then Exit Function      ' zwei Buchungen im selben Zeitraum
        ElseIf Abs(d(i) - 2 * P) <= tol * 1.5 Then
            reg = reg + 1                             ' ein Zeitraum ausgelassen
        End If
    Next i
    frac = reg / (n - 1)
    If P <= 100 Then
        minN = minAnz
    ElseIf P <= 200 Then
        minN = IIf(minAnz > 3, minAnz, 3)
    Else
        minN = 2
    End If
    If n < minN Then Exit Function
    If frac < IIf(streng, 0.8, 0.7) Then Exit Function
    BewerteSerie = True
End Function

' ===========================================================================
'  Abo anlegen
' ===========================================================================
Private Sub NeuesAbo(ByRef idx() As Long, ByVal n As Long, ByVal P As Double, ByVal rhy As String, ByVal fak As Double, ByVal grund As String)
    Dim i As Long, last As Long, j As Long
    Dim fest As Boolean, lastschrift As Boolean, ueberweisung As Boolean, karte As Boolean, rate As Boolean
    Dim summe As Double, vorher As String
    Dim stand As Double, naechste As Double, tol As Double
    Dim rn As Long, rm As Long, rinfo As String
    Dim konto As String, mehrere As Boolean
    SortNachDatum idx, n
    last = idx(n)
    nA = nA + 1
    For j = 1 To n
        i = idx(j)
        bAbo(i) = nA
        summe = summe + Abs(bAmt(i))
        If bMand(i) <> "" Or bGl(i) <> "" Or bZart(i) = "Lastschrift" Then lastschrift = True
        If bZart(i) = "Überweisung" Or bZart(i) = "Dauerauftrag" Then ueberweisung = True
        If bZart(i) = "Karte" Then karte = True
        If HatStichwort(bText(i), kwRate) Then rate = True
        If RatenZaehler(bRef(i) & " " & bName(i), rn, rm) Then rate = True
        If konto = "" Then
            konto = bKonto(i)
        ElseIf konto <> bKonto(i) Then
            mehrere = True
        End If
    Next j
    BetragsAnalyse idx, n, fest, vorher

    aName(nA) = SauberName(bName(last))
    If aName(nA) = "" Then aName(nA) = SauberName(bRef(last))
    If aName(nA) = "" Then aName(nA) = "(ohne Namen)"
    If fest Then
        If rate Then
            aArt(nA) = "Teilzahlung"
        ElseIf ueberweisung And Not lastschrift And Not karte Then
            aArt(nA) = "Dauerauftrag"
        Else
            aArt(nA) = "Abo"
        End If
        aBetrag(nA) = Abs(bAmt(last))
    Else
        aArt(nA) = "Abo (variabel)"
        aBetrag(nA) = Abs(bAmt(last))
    End If
    aMittel(nA) = summe / n
    aP(nA) = P
    aCent(nA) = bCent(last)
    aFest(nA) = fest
    aDel(nA) = False
    aRhy(nA) = rhy
    aMonat(nA) = Round(IIf(fest, aBetrag(nA), aMittel(nA)) * fak, 2)
    aJahr(nA) = Round(aMonat(nA) * 12, 2)
    aAnz(nA) = n
    aErste(nA) = bDat(idx(1))
    aLetzte(nA) = bDat(last)
    naechste = bDat(last) + Round(P)
    aNaechste(nA) = naechste
    aKonto(nA) = IIf(mehrere, "mehrere", konto)
    aZart(nA) = bZart(last)
    aGrund(nA) = grund & IIf(vorher <> "", "; " & vorher, "")
    aRef(nA) = Left$(bRef(last), 80)
    If Left$(grund, 8) = "gleicher" And InStr(grund, "Betrag") > 0 And InStr(grund, "Abbucher") = 0 Then
        aKey(nA) = "B|" & bCent(last) & "|" & rhy
    Else
        aKey(nA) = "A|" & bKey(last) & "|" & IIf(grund = "gleicher Abbucher und Betrag", CStr(bCent(last)), "*")
    End If

    ' Raten „n von m“ aus der letzten Buchung (sonst aus einer früheren)
    rinfo = ""
    For j = n To 1 Step -1
        If RatenZaehler(bRef(idx(j)) & " " & bName(idx(j)), rn, rm) Then
            If j = n Then
                rinfo = rn & " von " & rm
            Else
                rinfo = (rn + (n - j)) & " von " & rm
            End If
            Exit For
        End If
    Next j
    If rinfo = "" And aArt(nA) = "Teilzahlung" Then rinfo = n & " bisher"
    aRaten(nA) = rinfo

    ' Status anhand des Datenstands des Kontos
    stand = Datenstand(bKonto(last))
    If P < 20 Then tol = 2 Else tol = tolTage + P * 0.08
    If rinfo <> "" And rn > 0 And rm > 0 And (rn + (n - j)) >= rm Then
        aStatus(nA) = "abgeschlossen"
    ElseIf stand <= naechste + tol Then
        aStatus(nA) = "aktiv"
    ElseIf stand <= naechste + P Then
        aStatus(nA) = "überfällig"
    Else
        aStatus(nA) = "beendet"
    End If
    aSort(nA) = IIf(aStatus(nA) = "aktiv" Or aStatus(nA) = "überfällig", 0, 1) * 1000000 - aMonat(nA)
End Sub

' Fester Betrag? (gleich innerhalb Toleranz, oder in Blöcken geändert - z. B. Preiserhöhung)
Private Sub BetragsAnalyse(ByRef idx() As Long, ByVal n As Long, ByRef fest As Boolean, ByRef vorher As String)
    Dim j As Long, i As Long, k As Long
    Dim blk() As Long, nBlk As Long
    Dim cur As Long, c As Long
    Dim dominant As Long, domCnt As Long, cnt As Long
    fest = False: vorher = ""
    ' Blöcke gleicher Beträge in Datumsreihenfolge
    ReDim blk(1 To n)
    nBlk = 0
    cur = -1
    For j = 1 To n
        c = bCent(idx(j))
        If nBlk = 0 Then
            nBlk = 1: blk(1) = c: cur = c
        ElseIf Not GleicherBetrag(c, cur) Then
            nBlk = nBlk + 1: blk(nBlk) = c: cur = c
        End If
    Next j
    If nBlk = 1 Then
        fest = True
        Exit Sub
    End If
    ' dominanter Betrag (>= 80 %) - einzelne Ausreißer ignorieren
    For j = 1 To n
        cnt = 0
        For i = 1 To n
            If GleicherBetrag(bCent(idx(i)), bCent(idx(j))) Then cnt = cnt + 1
        Next i
        If cnt > domCnt Then domCnt = cnt: dominant = bCent(idx(j))
    Next j
    If domCnt >= n * 0.8 Then
        fest = True
        Exit Sub
    End If
    ' Blöcke ohne Wiederkehr eines früheren Betrags = Betragsänderung (fest, geändert)
    For j = 2 To nBlk
        For k = 1 To j - 1
            If GleicherBetrag(blk(j), blk(k)) Then Exit Sub        ' verschränkt -> variabel
        Next k
    Next j
    If nBlk <= 3 Then
        fest = True
        vorher = "Betrag geändert (vorher " & Format(blk(nBlk - 1) / 100, "#,##0.00") & " €)"
    End If
End Sub

Private Function GleicherBetrag(ByVal c1 As Long, ByVal c2 As Long) As Boolean
    Dim t As Double
    t = tolPct * ((c1 + c2) / 2)
    If t < 2 Then t = 2
    GleicherBetrag = (Abs(c1 - c2) <= t)
End Function

' „3/12“, „Rate 3 von 12“, „3. von 12 Raten“ -> n, m
Public Function RatenZaehler(ByVal s As String, ByRef rn As Long, ByRef rm As Long) As Boolean
    Dim w() As String
    Dim i As Long, p As Long
    Dim a As String, b As String
    rn = 0: rm = 0
    s = Replace(Replace(Replace(s, "(", " "), ")", " "), ",", " ")
    w = Split(Trim$(s), " ")
    For i = 0 To UBound(w)
        p = InStr(w(i), "/")
        If p > 1 And p < Len(w(i)) Then
            a = Left$(w(i), p - 1): b = Mid$(w(i), p + 1)
            If NurZiffern(a) And NurZiffern(b) Then
                If Len(a) <= 3 And Len(b) <= 3 Then
                    rn = CLng(a): rm = CLng(b)
                    If rm > 1 And rn <= rm And rn >= 1 Then RatenZaehler = True: Exit Function
                End If
            End If
        End If
    Next i
    For i = 1 To UBound(w) - 1
        If LCase$(w(i)) = "von" Then
            a = Replace(w(i - 1), ".", ""): b = w(i + 1)
            If NurZiffern(a) And NurZiffern(b) Then
                If Len(a) <= 3 And Len(b) <= 3 Then
                    rn = CLng(a): rm = CLng(b)
                    If rm > 1 And rn <= rm And rn >= 1 Then RatenZaehler = True: Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function NurZiffern(ByVal s As String) As Boolean
    Dim i As Long
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        If Not (Mid$(s, i, 1) Like "#") Then Exit Function
    Next i
    NurZiffern = True
End Function

' ===========================================================================
'  Ergebnis schreiben
' ===========================================================================
Private Sub SchreibeErgebnis(ByVal wsB As Worksheet, ByVal wsA As Worksheet)
    Dim out() As Variant, bo() As Variant
    Dim ord() As Long
    Dim i As Long, j As Long, k As Long, r As Long, t As Long
    Dim altKat As Collection, altIgn As Collection
    Dim key As String, lastA As Long
    Dim wsM As Worksheet
    Dim monate As Collection
    Dim mon As Double

    ' --- Kategorie / Ignorieren aus dem alten Blatt merken
    Set altKat = New Collection
    Set altIgn = New Collection
    lastA = A_FIRST - 1
    For r = A_MAX To A_FIRST Step -1
        If SStr(wsA.Cells(r, A_NAME).Value) <> "" Or SStr(wsA.Cells(r, A_KEY).Value) <> "" Then lastA = r: Exit For
    Next r
    For r = A_FIRST To lastA
        key = SStr(wsA.Cells(r, A_KEY).Value)
        If key <> "" Then
            If SStr(wsA.Cells(r, A_KAT).Value) <> "" And Not HatKey(altKat, "x" & key) Then altKat.Add SStr(wsA.Cells(r, A_KAT).Value), "x" & key
            If SStr(wsA.Cells(r, A_IGN).Value) <> "" And Not HatKey(altIgn, "x" & key) Then altIgn.Add SStr(wsA.Cells(r, A_IGN).Value), "x" & key
        End If
    Next r
    If lastA >= A_FIRST Then wsA.Range(wsA.Cells(A_FIRST, 1), wsA.Cells(lastA, A_NCOLS)).ClearContents

    ' --- Reihenfolge: aktive zuerst, dann nach Monatsbetrag absteigend (zusammengeführte Zeilen weglassen)
    Dim nOut As Long
    nOut = 0
    If nA > 0 Then
        ReDim ord(1 To nA)
        For i = 1 To nA
            If Not aDel(i) Then nOut = nOut + 1: ord(nOut) = i
        Next i
    End If
    If nOut > 0 Then
        For i = 2 To nOut
            t = ord(i): j = i - 1
            Do While j >= 1
                If aSort(ord(j)) <= aSort(t) Then Exit Do
                ord(j + 1) = ord(j): j = j - 1
            Loop
            ord(j + 1) = t
        Next i
        If nOut > A_MAX - A_FIRST + 1 Then nOut = A_MAX - A_FIRST + 1
        ReDim out(1 To nOut, 1 To A_NCOLS)
        For k = 1 To nOut
            i = ord(k)
            out(k, A_NAME) = TxtZelle(aName(i))
            out(k, A_ART) = aArt(i)
            out(k, A_STATUS) = aStatus(i)
            out(k, A_RHY) = aRhy(i)
            out(k, A_BETRAG) = aBetrag(i)
            out(k, A_MITTEL) = Round(aMittel(i), 2)
            out(k, A_MONAT) = aMonat(i)
            out(k, A_JAHR) = aJahr(i)
            out(k, A_ANZ) = aAnz(i)
            out(k, A_ERSTE) = CDate(aErste(i))
            out(k, A_LETZTE) = CDate(aLetzte(i))
            out(k, A_NAECHSTE) = CDate(aNaechste(i))
            out(k, A_KONTO) = TxtZelle(aKonto(i))
            out(k, A_ZART) = aZart(i)
            out(k, A_GRUND) = aGrund(i)
            out(k, A_RATEN) = TxtZelle(aRaten(i))
            out(k, A_REF) = TxtZelle(aRef(i))
            out(k, A_KAT) = HoleWert(altKat, "x" & aKey(i), Empty)
            out(k, A_IGN) = HoleWert(altIgn, "x" & aKey(i), Empty)
            out(k, A_KEY) = aKey(i)
        Next k
        wsA.Range(wsA.Cells(A_FIRST, 1), wsA.Cells(A_FIRST + nOut - 1, A_NCOLS)).Value = out
        wsA.Range(wsA.Cells(A_FIRST, A_ERSTE), wsA.Cells(A_FIRST + nOut - 1, A_NAECHSTE)).NumberFormat = "DD.MM.YYYY"
        wsA.Range(wsA.Cells(A_FIRST, A_BETRAG), wsA.Cells(A_FIRST + nOut - 1, A_JAHR)).NumberFormat = "#,##0.00 €"
    End If

    ' --- Buchungen: Abo-Name und Art
    If nB > 0 Then
        ReDim bo(1 To nB, 1 To 2)
        For i = 1 To nB
            If bAbo(i) > 0 Then
                bo(i, 1) = TxtZelle(aName(bAbo(i)))
                bo(i, 2) = aArt(bAbo(i))
            Else
                bo(i, 1) = Empty
                bo(i, 2) = bArt(i)
            End If
        Next i
        wsB.Range(wsB.Cells(B_FIRST, C_ABO), wsB.Cells(B_FIRST + nB - 1, C_ART)).Value = bo
    End If

    ' --- Monatsübersicht: Monate mit Buchungen
    On Error Resume Next
    Set wsM = Blatt("shMonat", "Monatsübersicht")
    On Error GoTo 0
    If Not wsM Is Nothing Then
        Set monate = New Collection
        For i = 1 To nB
            mon = CDbl(DateSerial(Year(CDate(bDat(i))), Month(CDate(bDat(i))), 1))
            If Not HatKey(monate, "x" & CStr(mon)) Then monate.Add mon, "x" & CStr(mon)
        Next i
        Dim ml() As Double, nM As Long, tmp As Double
        nM = monate.Count
        wsM.Range(wsM.Cells(M_FIRST, 1), wsM.Cells(M_LAST, 1)).ClearContents
        If nM > 0 Then
            ReDim ml(1 To nM)
            For i = 1 To nM: ml(i) = monate(i): Next i
            For i = 2 To nM
                tmp = ml(i): j = i - 1
                Do While j >= 1
                    If ml(j) <= tmp Then Exit Do
                    ml(j + 1) = ml(j): j = j - 1
                Loop
                ml(j + 1) = tmp
            Next i
            ' aufsteigend, höchstens so viele Monate wie Zeilen vorhanden (die jüngsten)
            k = 0
            If nM > M_LAST - M_FIRST + 1 Then k = nM - (M_LAST - M_FIRST + 1)
            For i = k + 1 To nM
                wsM.Cells(M_FIRST + i - k - 1, 1).Value = CDate(ml(i))
            Next i
            nM = nM - k
            wsM.Range(wsM.Cells(M_FIRST, 1), wsM.Cells(M_FIRST + nM - 1, 1)).NumberFormat = "MMM YYYY"
        End If
    End If
End Sub

' ===========================================================================
'  Sortieren / Hilfsfunktionen
' ===========================================================================
' Collection-Zugriffe nur über diese Helfer: Parameter (lokal) + .Item + IsEmpty-Prüfung,
' damit Excel und LibreOffice fehlende Schlüssel gleich behandeln.
Private Function HatKey(ByVal c As Collection, ByVal k As String) As Boolean
    Dim v As Variant
    If c Is Nothing Then Exit Function
    v = Empty
    On Error Resume Next
    v = c.Item(k)
    If Err.Number <> 0 Then
        Err.Clear
        HatKey = False
    Else
        HatKey = Not IsEmpty(v)
    End If
    On Error GoTo 0
End Function

Private Function HoleWert(ByVal c As Collection, ByVal k As String, ByVal std As Variant) As Variant
    Dim v As Variant
    HoleWert = std
    If c Is Nothing Then Exit Function
    v = Empty
    On Error Resume Next
    v = c.Item(k)
    If Err.Number <> 0 Then
        Err.Clear
    ElseIf Not IsEmpty(v) Then
        HoleWert = v
    End If
    On Error GoTo 0
End Function

' sMode 1: Schlüssel, dann Datum   -   sMode 2: Betrag (Cent), dann Datum
Private Function Cmp(ByVal a As Long, ByVal b As Long) As Long
    If sMode = 1 Then
        If bKey(a) < bKey(b) Then Cmp = -1: Exit Function
        If bKey(a) > bKey(b) Then Cmp = 1: Exit Function
    Else
        If bCent(a) < bCent(b) Then Cmp = -1: Exit Function
        If bCent(a) > bCent(b) Then Cmp = 1: Exit Function
    End If
    If bDat(a) < bDat(b) Then
        Cmp = -1
    ElseIf bDat(a) > bDat(b) Then
        Cmp = 1
    Else
        Cmp = 0
    End If
End Function

Private Sub SortIdx(ByRef idx() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Long, t As Long
    If lo >= hi Then Exit Sub
    i = lo: j = hi
    p = idx((lo + hi) \ 2)
    Do While i <= j
        Do While Cmp(idx(i), p) < 0: i = i + 1: Loop
        Do While Cmp(idx(j), p) > 0: j = j - 1: Loop
        If i <= j Then
            t = idx(i): idx(i) = idx(j): idx(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then SortIdx idx, lo, j
    If i < hi Then SortIdx idx, i, hi
End Sub

Private Sub SortNachDatum(ByRef idx() As Long, ByVal n As Long)
    Dim i As Long, j As Long, t As Long
    For i = 2 To n
        t = idx(i): j = i - 1
        Do While j >= 1
            If bDat(idx(j)) <= bDat(t) Then Exit Do
            idx(j + 1) = idx(j): j = j - 1
        Loop
        idx(j + 1) = t
    Next i
End Sub

Private Function Median(ByRef d() As Double, ByVal n As Long) As Double
    Dim s() As Double
    Dim i As Long, j As Long, t As Double
    ReDim s(1 To n)
    For i = 1 To n: s(i) = d(i): Next i
    For i = 2 To n
        t = s(i): j = i - 1
        Do While j >= 1
            If s(j) <= t Then Exit Do
            s(j + 1) = s(j): j = j - 1
        Loop
        s(j + 1) = t
    Next i
    If n Mod 2 = 1 Then
        Median = s((n + 1) \ 2)
    Else
        Median = (s(n \ 2) + s(n \ 2 + 1)) / 2
    End If
End Function
