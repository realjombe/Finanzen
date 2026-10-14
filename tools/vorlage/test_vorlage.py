# -*- coding: utf-8 -*-
"""
Testlauf der Vorlage in LibreOffice (headless): baut die Mappe, erzeugt Beispiel-Exporte,
führt den VBA-Import + die Abo-Erkennung aus und druckt das Blatt „Abos“.

    python3 tools/vorlage/test_vorlage.py [arbeitsordner]

Voraussetzung: soffice + python3-uno. Makros laufen über ein eigenes LibreOffice-Profil
(Makrosicherheit niedrig, VBA ausführbar).
"""
import os
import subprocess
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_vorlage  # noqa: E402
import make_samples  # noqa: E402

PORT = 2083
XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Filter/Microsoft/Import"><prop oor:name="LoadExcelBasicCode" oor:op="fuse"><value>true</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Filter/Microsoft/Import"><prop oor:name="LoadExcelBasicExecutable" oor:op="fuse"><value>true</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Filter/Microsoft/Import"><prop oor:name="LoadExcelBasicStorage" oor:op="fuse"><value>true</value></prop></item>
</oor:items>
"""


def pv(n, v):
    from com.sun.star.beans import PropertyValue
    p = PropertyValue()
    p.Name = n
    p.Value = v
    return p


def soffice(profile):
    import uno
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    url = "uno:socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext" % PORT
    try:
        return resolver.resolve(url)
    except Exception:
        pass
    os.makedirs(os.path.join(profile, "user"), exist_ok=True)
    with open(os.path.join(profile, "user", "registrymodifications.xcu"), "w", encoding="utf-8") as f:
        f.write(XCU)
    subprocess.Popen(["soffice", "--headless", "--norestore", "--nologo", "--nodefault", "--nofirststartwizard",
                      "-env:UserInstallation=file://" + profile, "--accept=socket,host=127.0.0.1,port=%d;urp;" % PORT],
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(120):
        try:
            return resolver.resolve(url)
        except Exception:
            time.sleep(1)
    raise RuntimeError("soffice nicht erreichbar")


def load(ctx, path):
    import uno
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(path)), "_blank", 0,
                                        (pv("Hidden", True), pv("MacroExecutionMode", 4)))


def macro(doc, mod, fn, args=()):
    s = doc.getScriptProvider().getScript("vnd.sun.star.script:VBAProject.%s.%s?language=Basic&location=document" % (mod, fn))
    return s.invoke(tuple(args), (), ())[0]


def rows(doc, sheet, rng):
    return doc.Sheets.getByName(sheet).getCellRangeByName(rng).DataArray


def fmt(v):
    if isinstance(v, float):
        if 40000 < v < 60000 and abs(v - round(v)) < 1e-9:   # Datum
            import datetime
            return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))).strftime("%d.%m.%Y")
        return ("%.2f" % v).rstrip("0").rstrip(".")
    return str(v)


def main(work):
    os.makedirs(work, exist_ok=True)
    samples = os.path.join(work, "samples")
    os.makedirs(samples, exist_ok=True)
    make_samples.main(samples)
    mappe = os.path.join(work, "Finanzen_Vorlage_test.xlsm")
    build_vorlage.build(mappe)
    ctx = soffice(os.path.join(work, "lo_profile"))
    doc = load(ctx, mappe)
    try:
        libs = {n: list(doc.BasicLibraries.getByName(n).ElementNames) for n in doc.BasicLibraries.ElementNames}
        print("Module:", libs)
        files = "|".join(os.path.join(samples, "George-Export-%s.json" % n) for n in ("Girokonto", "Kreditkarte", "Sparkonto"))
        t = time.time()
        print("ImportTest:", macro(doc, "modImport", "ImportTest", (files,)))
        print("Dauer %.1f s" % (time.time() - t))
        print("\n== Übersicht Konten ==")
        for r in rows(doc, "Übersicht", "A18:C25"):
            if r[0]:
                print("  ", [fmt(v) for v in r])
        print("\n== Abos ==")
        kopf = rows(doc, "Abos", "A5:T5")[0]
        print("  ", " | ".join(kopf[:17]))
        for r in rows(doc, "Abos", "A6:T60"):
            if r[0]:
                print("  ", " | ".join(fmt(v) for v in r[:17]))
        print("\n== Übersicht Kennzahlen ==")
        for rng in ("A9:L9", "A10:L10", "A13:L13", "A14:L14"):
            print("  ", [fmt(v) for v in rows(doc, "Übersicht", rng)[0] if v != ""])
        print("== Übersicht Konten (Formeln) ==")
        for r in rows(doc, "Übersicht", "A18:I25"):
            if r[0]:
                print("  ", [fmt(v) for v in r])
        print("== Fälligkeiten ==")
        for r in rows(doc, "Übersicht", "A30:C32"):
            print("  ", [fmt(v) for v in r])
        print("  Datenstand:", fmt(rows(doc, "Übersicht", "I5:I5")[0][0]), "| letzter Import:", rows(doc, "Übersicht", "B5:B5")[0][0])
        print("\n== Buchungen (Auszug) ==")
        for r in rows(doc, "Buchungen", "A6:J14"):
            print("  ", [fmt(v) for v in r])
        print("\n== Monatsübersicht ==")
        for r in rows(doc, "Monatsübersicht", "A6:H12"):
            if r[0]:
                print("  ", [fmt(v) for v in r])
        # zweiter Import derselben Dateien: nur Duplikate
        print("\nZweiter Import:", macro(doc, "modImport", "ImportTest", (files,))[:120])
        # Kategorie eintragen, erneut erkennen -> bleibt erhalten
        sh = doc.Sheets.getByName("Abos")
        sh.getCellRangeByName("R6").String = "Testkategorie"
        sh.getCellRangeByName("S7").String = "x"
        key6 = sh.getCellRangeByName("T6").String
        print("AbosTest:", macro(doc, "modAbo", "AbosTest")[:100])
        for r in rows(doc, "Abos", "A6:T60"):
            if r[19] == key6:
                print("Kategorie erhalten:", r[17] == "Testkategorie", "| Zeile:", r[0])
        print("Diagnose:", macro(doc, "modAbo", "AboDiagnose", (12030,)))
        print("Art-Verteilung Buchungen:")
        arten = {}
        for r in rows(doc, "Buchungen", "H6:H2000"):
            if r[0]:
                arten[r[0]] = arten.get(r[0], 0) + 1
        print("  ", arten)
    finally:
        doc.close(True)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else tempfile.mkdtemp(prefix="vorlage_test_"))
