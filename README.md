# Media From Mobile Organize (MFMO)

Mediadateien (Fotos, Videos, PDFs) von Samsung- und Apple-Geräten chronologisch
umbenennen, so dass eine Sortierung nach Dateiname der Aufnahmereihenfolge entspricht.

## Zieldateiformat

```
YYYYMMDD_HHMMSSa.jpg
YYYYMMDD_HHMMSSb.jpg   ← bei Zeitgleichheit
```

## Voraussetzungen

- Python 3.12
- [ExifTool](https://exiftool.org/) (Windows-Binary unter `C:\Tools\ExifTool\exiftool.exe`)

```bash
pip install pandas openpyxl tqdm
```

## Dateien

| Datei                | Beschreibung                                          |
|----------------------|-------------------------------------------------------|
| `mfmo.py`            | Kombinierter Workflow (empfohlen)                     |
| `mfmo_1_analyze.py`  | Nur Schritt 1: Analyse                                |
| `mfmo_2_rename.py`   | Nur Schritt 2: Namenberechnung                        |
| `mfmo_3_execute.py`  | Nur Schritt 3: Umbenennung                            |
| `mfmo.ini`           | Konfiguration für kombinierten Workflow               |
| `mfmo_1_analyze.ini` | Konfiguration für `mfmo_1_analyze.py`                 |
| `mfmo_2_rename.ini`  | Konfiguration für `mfmo_2_rename.py`                  |

## Konfiguration (INI)

```ini
[ALLGEMEIN]
quelle_pfad   = C:\Users\ah\Documents\Sammelordner\Fotos\20230825
log_datei     = mfmo.log
exiftool_pfad = C:\Tools\ExifTool\exiftool.exe

[Artist]
SM-S901B      = JH
SM-G973F      = LH
SM-F711B      = AH
iPhone 16 Pro = EH
iphone 11 pro = EH
```

`[Artist]` ordnet Kameramodell-Bezeichnungen Autorkürzeln zu.

## Verwendung

### Kombinierter Workflow (empfohlen)

```bash
python mfmo.py
```

Der Workflow pausiert zweimal und wartet auf manuelle Bearbeitung der Excel-Dateien.

### Einzelne Schritte

```bash
python mfmo_1_analyze.py   # → <ordner>_analyse.xlsx erzeugen
# Excel manuell prüfen/bearbeiten (Time_Shift eintragen, Ziel_Name korrigieren)
python mfmo_2_rename.py    # → <ordner>_rename.xlsx erzeugen
# Excel prüfen
python mfmo_3_execute.py   # → Dateien umbenennen
```

## Workflow im Detail

### Schritt 1 — Analyse (`mfmo_1_analyze.py`)

- Liest alle Dateien im Quellordner ein
- Ruft ExifTool auf und extrahiert: Aufnahmedatum, Kamera-Fabrikat/-Modell, GPS, Artist
- Leitet `Ziel_Name` aus EXIF-Datum oder Dateiname ab (Muster `YYYYMMDD_HHMMSS`)
- Erstellt `<ordnername>_analyse.xlsx` mit fixierter Kopfzeile und Tabellenformat

**Manuelle Bearbeitung danach:**
- `Time_Shift` eintragen (Stunden-Offset, z. B. `+1` für iPhone-MOV im Sommer)
- `Ziel_Name` korrigieren, falls kein Datum erkennbar war
- `Artist` ergänzen, falls leer

### Schritt 2 — Rename (`mfmo_2_rename.py`)

- Liest `_analyse.xlsx` ein
- Wendet `Time_Shift` auf `Ziel_Name` an → `Ziel_Name_Time_Shift`
- Bereinigt Dubletten durch Anhängen von Buchstaben (`a`, `b`, `c`, …) → `Ziel_Name_bereinigt`
- Erstellt `<ordnername>_rename.xlsx`

**Manuelle Prüfung danach:** `Ziel_Name_bereinigt` kontrollieren

### Schritt 3 — Execute (`mfmo_3_execute.py`)

- Liest `_rename.xlsx` ein
- Benennt jede Datei von `Dateiname` nach `Ziel_Name_bereinigt` um (`os.rename`)
- Überspringt Zeilen mit fehlenden Werten oder bereits existierendem Zielnamen
- Gibt Zusammenfassung (umbenannt / Fehler) aus

## Gerätespezifische Hinweise

| Gerät / Format  | Besonderheit                                               |
|-----------------|------------------------------------------------------------|
| Samsung JPG     | Dateiname und EXIF-Datum zuverlässig                       |
| Samsung MP4     | EXIF-Datum abweichend, kein Kameramodell-Eintrag           |
| Apple HEIC      | EXIF vollständig                                           |
| Apple MOV       | Zeitstempel weicht ab: +1 h (Sommerzeit), +2 h (Winter)   |

## Log-Dateien

Jeder Lauf schreibt eine `.log`-Datei (Pfad aus INI). Format:
```
2026-08-01 10:00:00,000 - INFO - Nachricht
```
