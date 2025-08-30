import os
import sys
import logging
import pandas as pd
from tkinter import Tk, filedialog

def set_arbeitsverzeichnis():
    skriptpfad = os.path.abspath(sys.argv[0])
    os.chdir(os.path.dirname(skriptpfad))

def screen_and_log(message):
    print(message)
    logging.info(message)

def initialisiere_logging(log_datei):
    logging.basicConfig(
        filename=log_datei,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        encoding='utf-8'
    )

def waehle_excel_datei(startpfad):
    """Zeigt eine Dialogbox zur Auswahl einer Excel-Datei."""
    root = Tk()
    root.withdraw()
    datei = filedialog.askopenfilename(
        initialdir=startpfad,
        title="Bitte Rename-Excel-Datei auswählen",
        filetypes=[("Excel-Dateien", "*.xlsx")]
    )
    root.destroy()
    return os.path.normpath(datei) if datei else None

def fuehre_umbenennung_aus(excel_pfad, ordnerpfad):
    try:
        df = pd.read_excel(excel_pfad)
    except Exception as e:
        screen_and_log(f"Fehler beim Einlesen der Excel-Datei: {e}")
        return

    if 'Dateiname' not in df.columns or 'Ziel_Name_bereinigt' not in df.columns:
        screen_and_log("Spalten 'Dateiname' oder 'Ziel_Name_bereinigt' fehlen in der Excel-Datei.")
        return

    umbenannt = 0
    fehlerhaft = 0

    for _, row in df.iterrows():
        alt = row['Dateiname']
        neu = row['Ziel_Name_bereinigt']

        if pd.isna(alt) or pd.isna(neu):
            continue

        quellpfad = os.path.join(ordnerpfad, alt)
        zielpfad = os.path.join(ordnerpfad, neu)

        if not os.path.exists(quellpfad):
            screen_and_log(f"❌ FEHLT: Quelldatei nicht gefunden: {quellpfad}")
            fehlerhaft += 1
            continue

        if os.path.exists(zielpfad):
            screen_and_log(f"⚠️  ZIEL EXISTIERT BEREITS: {zielpfad} – übersprungen")
            fehlerhaft += 1
            continue

        try:
            os.rename(quellpfad, zielpfad)
            screen_and_log(f"✅ UMBENANNT: {alt} -> {neu}")
            umbenannt += 1
        except Exception as e:
            screen_and_log(f"❌ FEHLER bei {alt}: {e}")
            fehlerhaft += 1

    screen_and_log(f"\nZusammenfassung: {umbenannt} Dateien umbenannt, {fehlerhaft} Fehler oder Konflikte.")

def main():
    set_arbeitsverzeichnis()
    log_datei = "Fotos_Videos_Sortieren_Umbenennen_Log.txt"
    initialisiere_logging(log_datei)

    screen_and_log("Toolstart: Dateien umbenennen")

    excel_pfad = waehle_excel_datei(os.getcwd())
    if not excel_pfad:
        screen_and_log("Keine Excel-Datei ausgewählt. Vorgang abgebrochen.")
        return

    ordnerpfad = os.path.dirname(excel_pfad)
    fuehre_umbenennung_aus(excel_pfad, ordnerpfad)

    screen_and_log("Toolende.")

if __name__ == "__main__":
    main()
