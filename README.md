# Finanzen-Excel

Excel-Mappe (`Finanzen_George.xlsm`) für die Übersicht aller Einnahmen und Ausgaben aus den
George-Umsatzexporten (JSON) von Girokonto, Sparkonto, Bausparkonto und Kreditkarte.

## Blätter

| Blatt | Inhalt |
|---|---|
| **Übersicht** | Import-Schaltflächen, Kennzahlen für den Auswertungsmonat (Einnahmen, Ausgaben, Saldo, Verfügbar, Nettovermögen, Buchungen ohne Kategorie), Konten mit Kontostand, Ausgaben nach Kategorie |
| **Buchungen** | Alle importierten Buchungen: Datum, Konto, Empfänger, Verwendungszweck, Betrag, Bezeichnung, Kategorie, Art, Notiz. Editierbar sind Bezeichnung, Kategorie, Art und Notiz |
| **Monatsübersicht** | Einnahmen, Fixkosten, variable Ausgaben und Saldo je Monat des gewählten Jahres, Diagramm, Kategorien × Monate |
| **Fixkosten** | Geplante Einnahmen und Fixkosten mit Suchbegriff für den Import, Soll (Ø pro Monat), Aufteilung geteilter Abos (Anzahl Personen oder Anteil in %) mit „Mein Anteil“, Ist im Auswertungsmonat, letzte Buchung |
| **Importregeln** | Suchbegriff → Bezeichnung, Kategorie, Art (erste passende Regel gewinnt) |
| **Einstellungen** | Eine Kategorienliste für die ganze Mappe, Kontotypen, Arten, Import-Ordner, Toleranz |
| **Hinweise** | Bedienung und Logik des Imports |

**Art** einer Buchung: `Ausgabe` (variabel) · `Fixkosten` · `Einnahme` · `Umbuchung` (zwischen eigenen
Konten, zählt nirgends) · `Ignorieren`.

## Import

1. In George je Konto den Umsatz-Export als JSON herunterladen.
2. Dateien in den Ordner `Import` neben der Mappe legen und in der Mappe auf „Import-Ordner einlesen“
   klicken (oder „Dateien importieren …“ und die Dateien auswählen). Eingelesene Dateien wandern nach
   `Import\erledigt`.
3. Buchungen ohne Kategorie sind gelb markiert – Kategorie eintragen; der Import merkt sie sich für
   denselben Händler.

## VBA

Der Makro-Code ist in der Mappe enthalten. `VBA-Module.zip` enthält die Quellen (`modImport.bas`,
`modJson.bas`, Code für `DieseArbeitsmappe`, Menüband `customUI14.xml`) zum Nachlesen oder erneuten
Importieren im VBA-Editor (Alt+F11).
