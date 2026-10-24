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

### iPhone MOV — Zeitoffset automatisch umrechnen (2026-10-11)
- [ ] ExifTool die Zeit von MOV-Dateien selbst umrechnen lassen (`-api QuickTimeUTC`),
  statt `Time_Shift` von Hand einzutragen. Hintergrund: QuickTime speichert `CreateDate`
  in UTC. Daraus folgt Sommerzeit +2 h, Winterzeit +1 h (passt zu den Beobachtungen
  2026-04-05 mit +1 h und 2026-08-01 mit +2 h, sofern die April-Videos noch aus der
  Winterzeit stammten). Umsetzung in `mfmo_1_analyze.py` (und `mfmo.py`): Option nur für
  `.mov`/`.mp4` setzen, vorher an einem EH-Video aus Sommer- und Winterzeit gegen ein
  gleichzeitiges HEIC/JPG prüfen. Danach `Time_Shift` nur noch für Reisen/Zeitzonen.
  Bis dahin gilt: `Time_Shift = +2` (Sommerzeit) für EH-MOV-Dateien.

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
