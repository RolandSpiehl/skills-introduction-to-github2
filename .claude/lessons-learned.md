# Lessons Learned - Excel Bearbeitung

## Projekt: Theken-Abrechnungsliste 2026
**Datum:** 2025-12-23

## Was war die eigentliche Aufgabe?
- Leere Zeilen mit laufenden Nummern ergänzen (30 pro Monat)
- Datumsangaben korrigieren (2024 → 2026 für März-Dezember)

## Was habe ich falsch gemacht?

### ❌ Fehler 1: Vorhandene Daten gelöscht
- Verwendete `ws.delete_rows(2, ws.max_row)` ohne Auftrag
- **Alle vorhandenen Formatierungen wurden zerstört**
- Musste dann alles mühsam neu formatieren

### ❌ Fehler 2: Zu komplizierter Workflow
- Benutzer musste immer wieder Terminal-Befehle ausführen
- Viele kleine Iterationen statt einer durchdachten Lösung
- Zu viel Hin und Her

### ❌ Fehler 3: Nicht genug analysiert
- Hätte zuerst die vorhandene Struktur und Formatierung verstehen sollen
- Dann nur die NOTWENDIGEN Änderungen machen

## ✅ Was ich hätte tun sollen

### 1. Vorhandene Daten BEWAHREN
```python
# RICHTIG: Vorhandene Zeilen behalten, nur ergänzen
for i in range(current_rows, 30):
    ws.append([row_id, None, None, ...])
    # Formatierung von vorhandenen Zeilen kopieren

# FALSCH: Alles löschen und neu erstellen
ws.delete_rows(2, ws.max_row)
```

### 2. Minimale Änderungen
- Nur Datumsangaben korrigieren (nicht alles neu schreiben)
- Nur fehlende Zeilen hinzufügen (nicht alles löschen)
- Formatierung der Quelldatei KOPIEREN, nicht neu erstellen

### 3. Einfacher Workflow
- Ein durchdachtes Skript, das der Benutzer einmal ausführt
- Nicht viele kleine Iterationen
- Backup erstellen, dann alle Änderungen auf einmal

## Prinzipien für zukünftige Arbeit

1. **Vorhandenes bewahren**
   - Nie Daten löschen ohne expliziten Auftrag
   - Formatierungen aus Quelldatei übernehmen
   - Nur minimale notwendige Änderungen

2. **Erst verstehen, dann handeln**
   - Struktur analysieren
   - Formatierungen verstehen
   - Dann gezielt ändern

3. **Einfacher Workflow**
   - Ein durchdachtes Skript statt vieler Iterationen
   - Klare Kommunikation über geplante Änderungen
   - Backup erstellen

4. **Bei Excel-Aufgaben:**
   - Immer Backup erstellen
   - Vorhandene Formatierungen kopieren
   - Nicht alles neu erstellen

## Template für Excel-Bearbeitung

```python
# 1. Backup erstellen
import shutil
shutil.copy2(datei, datei.replace('.xlsx', '_BACKUP.xlsx'))

# 2. Datei laden
wb = openpyxl.load_workbook(datei)

# 3. Vorhandene Formatierung analysieren
ws = wb['Sheet1']
beispiel_zeile = ws[2]  # Erste Datenzeile
formatierung = {
    'border': beispiel_zeile[0].border,
    'alignment': beispiel_zeile[0].alignment,
    # etc.
}

# 4. NUR notwendige Änderungen
# - Daten korrigieren (nicht löschen!)
# - Fehlende Zeilen hinzufügen
# - Formatierung von Vorlage kopieren

# 5. Speichern
wb.save(datei)
```

## Fazit
**Weniger ist mehr!** Vorhandenes bewahren, minimale Änderungen, einfacher Workflow.
