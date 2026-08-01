# TODO

## Offen

### Analyse (`mfmo_1_analyze.py`)
- [ ] **Default-Ordner als Fallback**: Falls der `quelle_pfad` aus der INI nicht existiert,
  einen sinnvollen Standardpfad (z. B. Desktop oder aktuelles Verzeichnis) verwenden
  statt mit Fehler abzubrechen.

### Samsung MP4 — Autor-Erkennung
- [ ] Strategie für Samsung-MP4-Dateien ohne Kameramodell-Eintrag im EXIF klären.
  Optionen:
  - Dateien vorab nach Autor in Unterordner sortieren und Kamera/Autor per ExifTool
    in die EXIF-Daten schreiben
  - Dropbox-Metadaten auslesen, um Autor zu identifizieren

### iPhone MOV — Zeitoffset
- [ ] Beobachten, ob der Zeitoffset (+2 h, Stand 2026-08-01) sich mit der Sommerzeit
  auf +1 h reduziert. Aktuell gilt: `Time_Shift = +2` für EH-MOV-Dateien eintragen.

### Metadaten schreiben
- [ ] Entscheiden, ob EXIF-Metadaten (Artist, Kameramodell) nach der Umbenennung
  nachträglich per ExifTool eingeschrieben werden sollen.

## Erledigte Punkte (Referenz)

- [x] Log-Warnung, wenn weder EXIF-Datum noch Datum im Ordnernamen (20250513a)
- [x] Tkinter-Dateiauswahl statt hartcodiertem Pfad (20250514a)
- [x] `_rename.xlsx` als separate Datei (nicht `_analyse.xlsx` überschreiben) (20250514b)
- [x] Dubletten-Bereinigung mit aufsteigenden Buchstaben (a, b, c …)
- [x] Time_Shift-Spalte in `_analyse.xlsx` für Zeitkorrektur pro Datei
- [x] Kombinierter Workflow `mfmo.py` mit Pause-Punkten zur manuellen Excel-Bearbeitung
