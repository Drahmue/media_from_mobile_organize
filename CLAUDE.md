# CLAUDE.md — Media From Mobile Organize (MFMO)

## Projektüberblick

Python-Tool zum chronologischen Umbenennen von Mediadateien (Fotos, Videos, PDFs) von
Samsung- und Apple-Geräten. Ziel: Nach Dateiname sortiert ergibt sich die korrekte
zeitliche Reihenfolge (`YYYYMMDD_HHMMSS[a|b|c].ext`).

## Dateistruktur

| Datei               | Zweck                                                         |
|---------------------|---------------------------------------------------------------|
| `mfmo.py`           | Kombinierter Workflow (alle 3 Schritte, mit Pausen)           |
| `mfmo_1_analyze.py` | Schritt 1: EXIF auslesen → `_analyse.xlsx`                    |
| `mfmo_2_rename.py`  | Schritt 2: Time_Shift + Dublettenbereinigung → `_rename.xlsx` |
| `mfmo_3_execute.py` | Schritt 3: Dateien physisch umbenennen                        |
| `mfmo.ini`          | Konfiguration für `mfmo.py`                                   |
| `mfmo_1_analyze.ini`| Konfiguration für `mfmo_1_analyze.py`                         |
| `mfmo_2_rename.ini` | Konfiguration für `mfmo_2_rename.py`                          |

## Coding-Konventionen (immer einhalten)

- **Coding-Standard:** `C:\Users\ah\Dropbox\ClaudeWork\_templates\python_standard.md` (ahlib-API: `create_extended_logger`, `StructuredConfigParser`)
- **Konfiguration** immer aus INI-Datei lesen (`configparser`)
- **Log-Datei** immer anlegen (`logging.basicConfig`, UTF-8)
- **Hauptprogramm** nur aus Funktionsaufrufen (`main()` ruft Funktionen auf)
- **Arbeitsverzeichnis** zu Skript-Beginn auf Skriptpfad setzen (`set_arbeitsverzeichnis()`)
- **UTF-8 überall**: Dateien, Logging, INI (`encoding='utf-8'`)
- **Fehlerbehandlung**: alle wahrscheinlichen Fehler abfangen, niemals stumm scheitern
- **Python 3.12**, Pandas 2.2.3, openpyxl, tqdm, tkinter, configparser

## Excel-Ausgabe-Konventionen

- Fixierte Kopfzeile (`freeze_panes = "A2"`)
- Tabellenformat `TableStyleMedium9` mit `displayName` (eindeutig pro Datei!)
- Auto-Spaltenbreite (max Zeichenbreite + 2)
- Spaltenreihenfolge `_analyse.xlsx`: Dateiname, Endung, EXIF-Aufnahmedatum, Kamera-Fabrikat, Kamera-Modell, GPS-Breite, GPS-Länge, Artist, **Time_Shift**, Ziel_Name
- Spalten `_rename.xlsx`: alle obigen + Ziel_Name_Time_Shift, Ziel_Name_bereinigt

## Zieldateiname-Format

```
YYYYMMDD_HHMMSS[buchstabe].[ext]
```

- Buchstabe (`a`, `b`, `c`, …) bei Zeitgleichheit (immer angehängt, auch beim ersten)
- `ext` in Kleinbuchstaben

## INI-Struktur

```ini
[ALLGEMEIN]
quelle_pfad   = C:\Pfad\zum\Quellordner
log_datei     = mfmo.log
exiftool_pfad = C:\Tools\ExifTool\exiftool.exe

[Artist]
sm-s901b      = JH
sm-g973f      = LH
sm-f711b      = AH
iphone 16 pro = EH
iphone 11 pro = EH
```

Artist-Keys werden **lowercase** verglichen (`.strip().lower()`).

## Externe Abhängigkeit: ExifTool

- Pfad aus INI (`exiftool_pfad`)
- Aufruf: `exiftool -j <dateipfad>` → JSON-Ausgabe
- Felder: `DateTimeOriginal`, `CreateDate`, `Make`, `Model`, `GPSLatitude`, `GPSLongitude`, `Artist`

## Bekannte Geräteeigenarten

| Gerät/Format     | Problem                                                              |
|------------------|----------------------------------------------------------------------|
| Samsung MP4      | EXIF-Datum abweichend / unzuverlässig; kein Kameramodell-Eintrag    |
| iPhone MOV       | Zeitstempel in UTC (QuickTime): +2h (Sommerzeit), +1h (Winter)      |
| Samsung JPG/MP4  | Dateiname enthält verlässlich Datum (`YYYYMMDD_HHMMSS`)             |

## Workflow (mfmo.py)

1. Verzeichnis auswählen (Tkinter-Dialog)
2. **Schritt 1**: EXIF-Analyse → `<ordner>_analyse.xlsx` → Pause zur manuellen Bearbeitung
3. **Schritt 2**: Time_Shift anwenden + Dubletten bereinigen → `<ordner>_rename.xlsx` → Pause zur Prüfung
4. **Schritt 3**: Dateien umbenennen (`os.rename`)

## Time_Shift-Spalte

- Ganzzahl (Stunden), positiv oder negativ
- Wird per `timedelta(hours=Time_Shift)` auf `Ziel_Name` addiert
- Muss für iPhone-MOV-Dateien manuell in der Excel-Datei eingetragen werden
