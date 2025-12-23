#!/usr/bin/env python3
"""
Skript zur Korrektur der Theken-Abrechnungsliste 2026
Führt folgende Schritte durch:
1. Korrektur der Datumsangaben für März-Dezember 2026
2. Anpassung der Zeilennummern (20001-20360, 30 pro Monat)
3. Einheitliche Formatierung für DIN-A4 Druck
4. Überprüfung aller Freitage, Samstage, Sonntage
"""

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.page import PageMargins
from datetime import datetime, timedelta
import calendar
import sys

# Datei-Pfad
DATEI_PFAD = '/Users/dr.rolandspiehl/skills-introduction-to-github2/Theken-Abrechnungsliste2026.xlsx'

def lade_arbeitsbuch():
    """Lädt die Excel-Datei"""
    print(f"Lade Datei: {DATEI_PFAD}")
    try:
        wb = openpyxl.load_workbook(DATEI_PFAD)
        print(f"✓ Datei geladen. Arbeitsblätter: {wb.sheetnames}")
        return wb
    except FileNotFoundError:
        print(f"✗ Fehler: Datei nicht gefunden: {DATEI_PFAD}")
        sys.exit(1)
    except Exception as e:
        print(f"✗ Fehler beim Laden: {e}")
        sys.exit(1)

def analysiere_struktur(wb):
    """Analysiert die Struktur der Excel-Datei"""
    print("\n=== ANALYSE DER DATEI-STRUKTUR ===")

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        print(f"\nArbeitsblatt: {sheet_name}")
        print(f"  Dimensionen: {ws.dimensions}")

        # Erste Zeilen anzeigen
        print(f"  Erste 5 Zeilen:")
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), 1):
            print(f"    Zeile {row_idx}: {row}")

    return True

def hole_wochenendtage_2026(monat):
    """Gibt alle Freitage, Samstage und Sonntage eines Monats in 2026 zurück"""
    jahr = 2026
    wochenendtage = []

    # Anzahl der Tage im Monat
    tage_im_monat = calendar.monthrange(jahr, monat)[1]

    # Durchlaufe alle Tage des Monats
    for tag in range(1, tage_im_monat + 1):
        datum = datetime(jahr, monat, tag)
        # 4 = Freitag, 5 = Samstag, 6 = Sonntag
        if datum.weekday() in [4, 5, 6]:
            wochenendtage.append(datum)

    return wochenendtage

def schritt1_korrigiere_datumswerte(wb):
    """Schritt 1: Korrigiert Datumsangaben für März-Dezember 2026"""
    print("\n=== SCHRITT 1: Korrektur der Datumsangaben ===")

    # TODO: Implementierung nach Analyse der Dateistruktur
    print("Bitte warten - analysiere erst die Struktur...")

    return wb

def schritt2_korrigiere_zeilennummern(wb):
    """Schritt 2: Korrigiert Zeilennummern (20001-20360)"""
    print("\n=== SCHRITT 2: Korrektur der Zeilennummern ===")

    # TODO: Implementierung
    print("Wird nach Strukturanalyse implementiert...")

    return wb

def schritt3_formatierung(wb):
    """Schritt 3: Einheitliche Formatierung für DIN-A4"""
    print("\n=== SCHRITT 3: DIN-A4 Formatierung ===")

    # TODO: Implementierung
    print("Wird nach Strukturanalyse implementiert...")

    return wb

def schritt4_pruefe_wochenenden(wb):
    """Schritt 4: Überprüft alle Wochenendtage"""
    print("\n=== SCHRITT 4: Überprüfung der Wochenendtage ===")

    for monat in range(1, 13):
        monat_name = calendar.month_name[monat]
        wochenendtage = hole_wochenendtage_2026(monat)
        print(f"\n{monat_name} 2026: {len(wochenendtage)} Wochenendtage")
        for datum in wochenendtage:
            wochentag = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'][datum.weekday()]
            print(f"  {wochentag} {datum.strftime('%d.%m.%Y')}")

    return wb

def main():
    """Hauptfunktion"""
    print("=" * 60)
    print("THEKEN-ABRECHNUNGSLISTE 2026 - KORREKTUR-SKRIPT")
    print("=" * 60)

    # Lade Arbeitsbuch
    wb = lade_arbeitsbuch()

    # Analysiere Struktur
    analysiere_struktur(wb)

    # Frage ob fortfahren
    print("\n" + "=" * 60)
    antwort = input("Struktur analysiert. Mit Korrekturen fortfahren? (j/n): ")
    if antwort.lower() != 'j':
        print("Abgebrochen.")
        return

    # Führe Schritte durch
    wb = schritt1_korrigiere_datumswerte(wb)
    wb = schritt2_korrigiere_zeilennummern(wb)
    wb = schritt3_formatierung(wb)
    wb = schritt4_pruefe_wochenenden(wb)

    # Speichern
    backup_pfad = DATEI_PFAD.replace('.xlsx', '_backup.xlsx')
    print(f"\n=== SPEICHERN ===")
    print(f"Backup: {backup_pfad}")
    wb.save(backup_pfad)
    print(f"✓ Backup gespeichert")

    print(f"Korrigierte Datei: {DATEI_PFAD}")
    wb.save(DATEI_PFAD)
    print(f"✓ Datei gespeichert")

    print("\n" + "=" * 60)
    print("FERTIG!")
    print("=" * 60)

if __name__ == "__main__":
    main()
