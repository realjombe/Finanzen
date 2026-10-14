# -*- coding: utf-8 -*-
"""Erzeugt Beispiel-George-Exporte (JSON) zum Testen der Abo-Erkennung: Girokonto, Kreditkarte, Sparkonto."""
import json
import random
import sys
from datetime import date, timedelta

GIRO_IBAN = "AT021200000012345678"
SPAR_IBAN = "AT481200000098765432"
CARD_NO = "5401 **** **** 1234"
MONTHS = [(2025, m) for m in range(9, 13)] + [(2026, m) for m in range(1, 11)]   # 2025-09 .. 2026-10
TODAY = date(2026, 10, 13)


def iso(d):
    return d.strftime("%Y-%m-%dT00:00:00.000+0100")


def tx(d, amount, owner_no, owner_title, partner="", iban="", reference="", mandate="", creditor="", ref_no=None, extra=None):
    t = {
        "booking": iso(d),
        "valuation": iso(d),
        "partnerName": partner,
        "partnerAccount": {"iban": iban, "bic": "", "number": "", "bankCode": "", "countryCode": iban[:2] if iban else ""},
        "amount": {"value": int(round(amount * 100)), "precision": 2, "currency": "EUR"},
        "reference": reference,
        "referenceNumber": ref_no or "",
        "receiverReference": "",
        "sepaMandateId": mandate,
        "sepaCreditorId": creditor,
        "ownerAccountNumber": owner_no,
        "ownerAccountTitle": owner_title,
        "categories": [],
        "note": "",
    }
    if extra:
        t.update(extra)
    return t


def monthly(day, start=(2025, 9), end=(2026, 10)):
    for y, m in MONTHS:
        if (y, m) < start or (y, m) > end:
            continue
        yield date(y, m, min(day, 28))


def main(outdir):
    rnd = random.Random(42)
    giro, card, spar = [], [], []
    G = dict(owner_no=GIRO_IBAN, owner_title="Girokonto")
    C = dict(owner_no=CARD_NO, owner_title="Mastercard Gold")
    S = dict(owner_no=SPAR_IBAN, owner_title="Sparkonto")

    # 1 Spotify: Lastschrift, Preiserhoehung ab 04/2026
    for d in monthly(5):
        amt = -10.99 if (d.year, d.month) < (2026, 4) else -11.99
        giro.append(tx(d, amt, partner="Spotify AB", iban="SE1550000000054910000003", reference="Spotify Premium",
                       mandate="SPOT-AT-7781", creditor="SE87ZZZ0000012345", **G))
    # 2 A1: Lastschrift, variabler Betrag
    a1 = [29.90, 31.45, 29.90, 29.90, 34.10, 29.90, 29.90, 30.20, 29.90, 29.90, 29.90, 33.70, 29.90, 29.90]
    for d, amt in zip(monthly(15), a1):
        giro.append(tx(d, -amt, partner="A1 Telekom Austria AG", iban="AT611904300234573201",
                       reference="Rechnung %02d/%d Kundennr 123456" % (d.month, d.year), mandate="A1-M-556677", creditor="AT12ZZZ00000000123", **G))
    # 3 Wien Energie: Teilbetrag fest
    for d in monthly(2):
        giro.append(tx(d, -86.00, partner="Wien Energie GmbH", iban="AT483200000012345864", reference="Teilbetrag Strom/Gas Vertrag 998877",
                       mandate="WE-998877", creditor="AT98ZZZ00000000777", **G))
    # 4 Miete: Dauerauftrag (Ueberweisung, fester Betrag)
    for d in monthly(1):
        giro.append(tx(d, -850.00, partner="Hausverwaltung Müller GmbH", iban="AT611904300234573999", reference="Miete Wohnung Top 12", **G))
    # 5 Ratenzahlung Moebel: 12 Raten ab 11/2025
    n = 0
    for d in monthly(20, start=(2025, 11)):
        n += 1
        giro.append(tx(d, -125.00, partner="XXXLutz KG", iban="AT021100000045678901", reference="Rate %d/12 KV 4711" % n, **G))
    # 6 Kreditrate
    for d in monthly(10):
        giro.append(tx(d, -310.50, partner="Erste Bank der oesterreichischen Sparkassen AG", iban="AT022011100000000011",
                       reference="Kreditrate Konto 12345", mandate="ERSTE-KR-1", creditor="AT61ZZZ00000000001", **G))
    # 7 Versicherung vierteljaehrlich
    for (y, m) in [(2025, 9), (2025, 12), (2026, 3), (2026, 6), (2026, 9)]:
        giro.append(tx(date(y, m, 8), -120.30, partner="Wiener Städtische Versicherung AG", iban="AT731200000055512345",
                       reference="Polizze 55-123 Haushaltsversicherung", mandate="WST-55-123", creditor="AT55ZZZ00000000055", **G))
    # 9 Fitness: gekuendigt nach 05/2026
    for d in monthly(3, end=(2026, 5)):
        giro.append(tx(d, -29.90, partner="FitInn GmbH", iban="AT401200000066612345", reference="Mitgliedsbeitrag", mandate="FIT-1001", creditor="AT77ZZZ00000000099", **G))
    # 10 Gehalt
    for d in monthly(28):
        giro.append(tx(d, 2450.00, partner="Muster GmbH", iban="AT951400000012340000", reference="Gehalt %02d/%d" % (d.month, d.year), **G))
    # 12 Taschengeld (Ueberweisung fester Betrag)
    for d in monthly(1):
        giro.append(tx(d, -50.00, partner="Max Mustermann", iban="AT111200000011111111", reference="Taschengeld", **G))
    # 13 Umbuchung auf Sparkonto (eigenes Konto) – darf kein Abo werden
    for d in monthly(2):
        giro.append(tx(d, -200.00, partner="Max Mustermann", iban=SPAR_IBAN, reference="Sparen", **G))
        spar.append(tx(d, 200.00, partner="Max Mustermann", iban=GIRO_IBAN, reference="Sparen", **S))
    # 16 Kindergartenbeitrag: gleicher Betrag, wechselnder Empfaengertext, keine IBAN (eps-Zahlung)
    names = ["Stadt Wien MA10", "MA 10 Kindergarten", "Magistrat Wien", "Wien MA10 Kindergaerten", "Kindergarten Stadt Wien", "MA10 Wien"]
    for i, d in enumerate(monthly(6)):
        giro.append(tx(d, -95.00, partner=names[i % len(names)], reference="eps Zahlung Kindergartenbeitrag %02d/%d" % (d.month, d.year), **G))
    # 17 Nachhilfe: erst Ueberweisung an IBAN A, dann Kartenzahlung (gleicher Betrag) -> eine Zahlung
    for d in monthly(9, end=(2026, 2)):
        giro.append(tx(d, -75.00, partner="Lisa Lehrer", iban="AT371200000077777777", reference="Nachhilfe", **G))
    for d in monthly(9, start=(2026, 3)):
        giro.append(tx(d, -75.00, partner="SUMUP *LISA LEHRER WIEN", reference="POS 75,00 AT K1 %s 17:00" % d.strftime("%d.%m."), **G))
    # 11 Lebensmittel (Karte am Girokonto): unregelmaessig
    shops = ["BILLA DANKT", "SPAR DANKT", "HOFER DANKT", "DM DROGERIE"]
    for y, m in MONTHS:
        for _ in range(rnd.randint(6, 10)):
            d = date(y, m, rnd.randint(1, 28))
            if d > TODAY:
                continue
            shop = rnd.choice(shops)
            amt = round(rnd.uniform(5, 80), 2)
            giro.append(tx(d, -amt, partner="%s %04d WIEN" % (shop, rnd.randint(1000, 9999)),
                           reference="POS %s AT K1 %s 10:15" % (("%.2f" % amt).replace(".", ","), d.strftime("%d.%m.")), **G))
    # 15 Bargeld: ca. monatlich 100 (gleicher Betrag, aber kein Abo)
    for y, m in MONTHS:
        d = date(y, m, rnd.randint(1, 28))
        if d <= TODAY:
            giro.append(tx(d, -100.00, partner="", reference="Bargeldbehebung Bankomat 12345 WIEN %s" % d.strftime("%d.%m."), **G))
    # 14 Kreditkartenabrechnung (variabel, monatlich) – darf kein Abo werden
    for d in monthly(27):
        amt = round(rnd.uniform(80, 400), 2)
        giro.append(tx(d, -amt, partner="Erste Bank Kreditkarte", iban="AT022011100000000099",
                       reference="Kreditkartenabrechnung %02d/%d Karte %s" % (d.month, d.year, CARD_NO), **G))
        card.append(tx(d, amt, partner="", reference="Zahlung Kreditkartenabrechnung", **C))

    # ---- Kreditkarte
    netflix = ["NETFLIX.COM AMSTERDAM NL", "Netflix International B.V.", "NETFLIX.COM 866-579-7172", "NETFLIX AMSTERDAM"]
    for i, d in enumerate(monthly(4)):
        card.append(tx(d, -13.99, reference=netflix[i % len(netflix)], **C))
    for d in monthly(11):                                  # Apple: zwei Abos, gleicher Haendler, gleicher Tag
        card.append(tx(d, -2.99, reference="APPLE.COM/BILL ITUNES.COM IE", **C))
        card.append(tx(d, -10.99, reference="APPLE.COM/BILL ITUNES.COM IE", **C))
    for y in (2025, 2026):                                 # Prime jaehrlich
        card.append(tx(date(y, 10, 12), -89.90, reference="AMAZON PRIME*2K3LX LUXEMBOURG", **C))
    for y, m in MONTHS:                                    # Amazon/Restaurants unregelmaessig
        for _ in range(rnd.randint(1, 4)):
            d = date(y, m, rnd.randint(1, 28))
            if d > TODAY:
                continue
            card.append(tx(d, -round(rnd.uniform(8, 120), 2), reference=rnd.choice(["AMAZON.DE LUXEMBOURG", "AMZN MKTP DE", "RESTAURANT ZUM GOLDENEN WIEN", "OEBB TICKETS WIEN"]), **C))

    # ---- Sparkonto: Zinsen / KESt
    spar.append(tx(date(2025, 12, 31), 1.23, partner="", reference="Zinsen", **S))
    spar.append(tx(date(2025, 12, 31), -0.31, partner="", reference="KESt", **S))

    for name, lst in (("Girokonto", giro), ("Kreditkarte", card), ("Sparkonto", spar)):
        lst = [t for t in lst if t["booking"][:10] <= TODAY.isoformat()]
        lst.sort(key=lambda t: t["booking"], reverse=True)
        with open("%s/George-Export-%s.json" % (outdir, name), "w", encoding="utf-8") as f:
            json.dump(lst, f, ensure_ascii=False, indent=2)
        print(name, len(lst), "Buchungen")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else ".")
