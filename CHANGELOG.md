# Changelog

## [Unreleased]

### Dokumentation
- 2026-10-24: Doku: Übernahme aus Claude-Memory (claude_setup CS-019) – CLAUDE.md: veraltete ahlib-API (`screen_and_log`) durch Verweis auf `ClaudeWork\_templates\python_standard.md` ersetzt

## [2026-10-11]

### Dokumentation
- README: Zeitkorrektur Apple MOV korrigiert (war vertauscht). QuickTime speichert UTC,
  daher Sommerzeit +2 h, Winterzeit +1 h.
- TODO: automatische Umrechnung per ExifTool (`-api QuickTimeUTC`) aufgenommen.

## [2026-04-05]

### Erkenntnisse
- Bei EH (iPhone) besteht zwischen MOV und JPG/HEIC eine Zeitdifferenz von +1 h (Sommerzeit).
  Für MOV-Dateien muss `Time_Shift = +1` eingetragen werden.
- Unklar, ob sich der Offset mit der Winterzeit auf +2 h ändert (bereits früher beobachtet).

## [2026-08-01]

### Erkenntnisse
- Aktuell gilt für EH (iPhone): Zeitdifferenz zwischen MOV und JPG/HEIC beträgt **+2 h**.
  Für MOV-Dateien `Time_Shift = +2` eintragen.

## [2025-09-01]

### Bugfix
- Einige MOV-Dateien von EHs iPhone hatten nach der Umbenennung einen um 2 h falschen
  Zeitstempel und befanden sich noch im Foto-Ordner des Sammelordners.

## [2025-05-14] — Version 20250514

### mfmo_1_analyze.py (20250514a)
- Lauffähige Version mit Tkinter-Dateiauswahl-Dialog

### mfmo_2_rename.py (20250514b)
- Auswahl der `_analyse.xlsx` per Dialog
- `_rename.xlsx` wird in separate Datei geschrieben (nicht dieselbe wie `_analyse.xlsx`)

### mfmo_3_execute.py (20250514a)
- Erste lauffähige Version der Umbenennung

## [2025-05-13] — Version 20250513

### mfmo_1_analyze.py (20250513a)
- Log-Warnung eingebaut, wenn weder EXIF-Datum erkannt wird noch der Ordnername ein Datum enthält

### mfmo_1_analyze.py (20250513b)
- (weitere Korrekturen, Details im Code-Verlauf)

## [2025-05-12] — Konzeptentscheidungen

- 3-stufiger Prozess festgelegt:
  1. Analyse-Tool → Excel
  2. Manuelle Nacharbeit in Excel
  3. Umbennungs-Tool auf Basis der Excel
- Samsung-Dateiname als zuverlässige Datumsquelle bestätigt (JPG und MP4)
- Samsung EXIF für MP4 als unzuverlässig eingestuft (kein Kameramodell-Eintrag)
- Dubletten-Bereinigung: erster gleicher Name erhält `a`, nächster `b` usw.

## [2025-05-10] — Projektstart

- Ziel definiert: Mediadateien unterschiedlicher Autoren und Geräte chronologisch umbenennen
- Format festgelegt: `YYYYMMDD_HHMMSS[buchstabe].[ext]`
- Ansatz: Python + ExifTool + Excel als manuell editierbare Zwischenstufe
- Technologie-Stack gewählt: Python 3.12, pandas, openpyxl, tqdm, configparser, tkinter
