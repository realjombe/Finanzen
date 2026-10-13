# Finanzen-Excel

Zwei Excel-Mappen für die George-Umsatzexporte (JSON) von Girokonto, Sparkonto, Bausparkonto und Kreditkarte:

| Mappe | Zweck |
|---|---|
| **`Finanzen_Vorlage.xlsm`** | Neue Vorlage: braucht nur die JSON-Exportdateien. Erkennt Abos, Teilzahlungen und Daueraufträge automatisch – über gleiche Abbucher und über gleiche Beträge jeden Monat. Keine Pflege von Fixkosten-Listen oder Importregeln. |
| `Finanzen_George.xlsm` | Bisherige Mappe mit Kategorien, Fixkosten-Planung, Importregeln und Kontoständen (Beschreibung weiter unten). |

Beide Mappen lesen Dateien aus dem Ordner `Import` neben der Mappe und verschieben sie danach nach `Import\erledigt`.

## Vorlage mit Abo-Erkennung (`Finanzen_Vorlage.xlsm`)

1. In George je Konto den Umsatz-Export als JSON herunterladen.
2. In der Mappe „Dateien importieren …“ klicken (oder Dateien nach `Import` legen und „Import-Ordner einlesen“).
3. Konten werden aus dem Export angelegt (Kontobezeichnung, IBAN/Kartennummer), Buchungen importiert, danach
   läuft die Abo-Erkennung automatisch. Ergebnis im Blatt **Abos**; „Abos neu erkennen“ startet sie jederzeit neu.

#### Blätter

| Blatt | Inhalt |
|---|---|
| **Übersicht** | Schaltflächen, Kennzahlen (aktive Abos, Abo-Kosten/Monat, laufende Teilzahlungen, Daueraufträge, fixe Belastung pro Monat und Jahr, Einnahmen/Ausgaben/Saldo im Auswertungsmonat), Konten aus dem Export, Fälligkeiten der nächsten 7/30 Tage |
| **Buchungen** | Alle importierten Buchungen: Datum, Konto, Empfänger, Verwendungszweck, Betrag, Zahlungsart (Lastschrift/Überweisung/Karte/Bargeld …), erkanntes Abo, Art, Notiz, IBAN Partner, Gläubiger-ID, Mandatsreferenz |
| **Abos** | Erkannte regelmäßige Zahlungen: Art (Abo · Abo (variabel) · Teilzahlung · Dauerauftrag), Status (aktiv · überfällig · beendet · abgeschlossen), Rhythmus, Betrag, pro Monat/Jahr, erste/letzte/nächste Buchung, Konto, „Erkannt durch“, Raten („11 von 12“), Kategorie und Ignorieren (Eingaben bleiben erhalten) |
| **Monatsübersicht** | Je Monat: Einnahmen, Abos, Teilzahlungen, Daueraufträge, sonstige Ausgaben, Saldo + Diagramm |
| **Einstellungen** | Mindestanzahl Buchungen, Toleranz Betrag/Termin, Mindestbetrag, Stichwörter (Teilzahlung, Umbuchung, Bargeld), Kategorien, Kontotypen |
| **Hinweise** | Bedienung und Logik |

### So erkennt die Vorlage Abos

* **Gleicher Abbucher:** Buchungen mit derselben Mandatsreferenz, Gläubiger-ID, IBAN oder demselben Namen
  bilden eine Serie. Kommen sie in festem Abstand (wöchentlich bis jährlich, Toleranz einstellbar) und mindestens
  so oft wie „Mindestanzahl“ (Jahresabos: 2), gilt die Serie als regelmäßige Zahlung.
  * fester Betrag → **Abo** (Lastschrift/Karte), **Dauerauftrag** (Überweisung) oder **Teilzahlung**
    (Stichwort wie Rate/Ratenzahlung/Teilbetrag oder Zähler wie „3/12“ im Text)
  * schwankender Betrag (Handy, Strom) → **Abo (variabel)**; Preisänderungen (10,99 → 11,99) bleiben ein Abo
  * mehrere Abos beim selben Anbieter (zwei Apple-Abos am selben Tag, Jahresabo neben Einzelkäufen) werden über
    den Betrag getrennt
* **Gleicher Betrag:** derselbe Betrag jeden Monat (bzw. in festem Rhythmus, eine Zahlung je Zeitraum) wird auch
  dann als Teilzahlung/Dauerauftrag/Abo erkannt, wenn der Empfängertext wechselt. Wechselt eine feste Zahlung den
  Empfänger oder Zahlungsweg (neue IBAN, Überweisung → Karte), werden die Serien zusammengeführt.
* **Nie Abos:** Umbuchungen zwischen eigenen Konten, Kreditkartenabrechnungen, Bargeldbehebungen, Eingänge.
* **Status:** aktiv = nächste Zahlung steht aus · überfällig = laut Datenstand des Kontos erwartet, aber nicht
  gekommen · beendet = seit mehr als einem Zeitraum nichts mehr · abgeschlossen = letzte Rate erreicht.

### Technik / Entwicklung

Die Vorlage wird ohne Excel gebaut: `tools/vorlage/build_vorlage.py` erzeugt die Blätter (openpyxl) und bettet den
VBA-Code aus `tools/vorlage/vba/` als `vbaProject.bin` ein (`ovba.py`, eigener MS-OVBA/CFB-Writer).
`tools/vorlage/test_vorlage.py` baut die Mappe, erzeugt Beispiel-Exporte (`make_samples.py`) und führt Import und
Abo-Erkennung headless in LibreOffice aus. `VBA-Module-Vorlage.zip` enthält die Module zum Nachlesen oder zum
Import im VBA-Editor (Alt+F11).

## Bisherige Mappe (`Finanzen_George.xlsm`)

Übersicht aller Einnahmen und Ausgaben mit Kategorien, Fixkosten-Planung und Importregeln.

### Blätter

| Blatt | Inhalt |
|---|---|
| **Übersicht** | Import-Schaltflächen, Kennzahlen für den Auswertungsmonat (Einnahmen, Ausgaben, Saldo, Verfügbar, Nettovermögen, Buchungen ohne Kategorie), Konten mit Kontostand und Monatswechsel je Konto (Kreditkarte: ab dem 28. zählt zum Folgemonat), Ausgaben nach Kategorie |
| **Buchungen** | Alle importierten Buchungen: Datum, Konto, Empfänger, Verwendungszweck, Betrag, Bezeichnung, Kategorie, Art, Notiz, berechneter Auswertungsmonat. Editierbar sind Bezeichnung, Kategorie, Art und Notiz |
| **Monatsübersicht** | Einnahmen, Fixkosten, variable Ausgaben und Saldo je Monat des gewählten Jahres, Diagramm, Kategorien × Monate |
| **Fixkosten** | Geplante Einnahmen und Fixkosten mit Suchbegriff für den Import, Soll (Ø pro Monat), Aufteilung geteilter Abos (Anzahl Personen oder Anteil in %) mit „Mein Anteil“, Ist im Auswertungsmonat, letzte Buchung |
| **Importregeln** | Suchbegriff → Bezeichnung, Kategorie, Art (erste passende Regel gewinnt) |
| **Einstellungen** | Eine Kategorienliste für die ganze Mappe, Kontotypen, Arten, Import-Ordner, Toleranz |
| **Hinweise** | Bedienung und Logik des Imports |

**Art** einer Buchung: `Ausgabe` (variabel) · `Fixkosten` · `Einnahme` · `Umbuchung` (zwischen eigenen
Konten, zählt nirgends) · `Ignorieren`.

### Import

1. In George je Konto den Umsatz-Export als JSON herunterladen.
2. Dateien in den Ordner `Import` neben der Mappe legen und in der Mappe auf „Import-Ordner einlesen“
   klicken (oder „Dateien importieren …“ und die Dateien auswählen). Eingelesene Dateien wandern nach
   `Import\erledigt`.
3. Buchungen ohne Kategorie sind gelb markiert – Kategorie eintragen; der Import merkt sie sich für
   denselben Händler.

### VBA

Der Makro-Code ist in der Mappe enthalten. `VBA-Module.zip` enthält die Quellen (`modImport.bas`,
`modJson.bas`, Code für `DieseArbeitsmappe`, Menüband `customUI14.xml`) zum Nachlesen oder erneuten
Importieren im VBA-Editor (Alt+F11).
