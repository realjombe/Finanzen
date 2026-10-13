Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    PruefeImportOrdnerBeimOeffnen
End Sub

' Schaltflächen sind Hyperlinks auf die eigene Zelle - hier werden sie ausgewertet
Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
    Dim t As String
    On Error Resume Next
    t = CStr(Target.Range.Cells(1, 1).Value)
    On Error GoTo 0
    If InStr(t, "Dateien importieren") > 0 Then
        ImportDateienAuswaehlen
    ElseIf InStr(t, "Import-Ordner einlesen") > 0 Then
        ImportOrdnerEinlesen
    ElseIf InStr(t, "Import-Ordner") > 0 And InStr(t, "ffnen") > 0 Then
        ImportOrdnerOeffnen
    ElseIf InStr(t, "Abos") > 0 And InStr(t, "erkennen") > 0 Then
        AbosErkennen
    End If
End Sub
