#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Thadden27 - dateien.json Generator & Datei-Sammler
====================================================
Sammelt MP3s und PDFs aus dem Quellordner (OneDrive), KOPIERT sie in die
Ordner audio/ und noten/ im Ziel-Ordner (lokales Git-Repo) und erstellt
dort automatisch die dateien.json.

VERWENDUNG:
  1. Sowohl der Quellordner als auch der Zielordner (lokales Git-Repo)
     liegen in OneDrive. Das Script sucht deshalb bei jedem Start
     zuerst automatisch den OneDrive-Ordner im Benutzerordner (egal wie
     der genau heißt, z.B. "OneDrive - thaddenschule.de") und hängt dann
     die festen Unterpfade SOURCE_SUBPATH ("09 Musik\\2025 Elisabeth")
     bzw. DEST_SUBPATH ("12 GitHub\\thadden27") an. Das Script selbst
     kann dadurch an einer beliebigen Stelle liegen (es muss NICHT mehr
     im Git-Ordner liegen) - auf einem neuen Rechner muss i.d.R. gar
     nichts von Hand angepasst werden, solange die Ordnerstruktur
     innerhalb von OneDrive gleich bleibt.
  2. Script ausführen: Doppelklick  ODER  in der Kommandozeile:
       python3 make_dateien.py
  3. Die Ordner "audio/", "noten/" und die Datei "dateien.json" liegen
     danach im Git-Ordner bereit zum Commit/Push auf GitHub.

QUELLORDNER (OneDrive) - VORHER:
  09 Musik/2025 Elisabeth/
  |---- Partien aus 1.1 Marsch/       <- MP3s, Szene-ID = 1.1
  |---- Partien aus 1.3 Nun will.../  <- MP3s, Szene-ID = 1.3
  |---- Partien aus 14 Finale/        <- MP3s, Szene-ID = 14
  |---- pdf/                          <- alle PDFs
  |   |---- 1.3 Nun will... - Flöte 1.pdf   <- Szene-ID = 1.3
  |   `---- 14 Finale - Partitur.pdf        <- Szene-ID = 14
  `---- TXT/                          <- alle Text-Dateien
      |---- 1.3 Nun will... - Text.txt      <- Szene-ID = 1.3
      `---- 14 Finale - Text.txt            <- Szene-ID = 14

ZIELORDNER (Git-Repo) - NACHHER:
  thadden27/
  |---- audio/                        <- alle MP3s aus allen "Partien aus ..."
  |                                    Unterordnern, hierher kopiert (flach,
  |                                    ohne Unterordner)
  |---- noten/                        <- alle PDFs aus pdf/, hierher kopiert
  |---- TXT/                          <- alle TXT-Dateien aus TXT/, hierher kopiert
  `---- dateien.json                  <- verweist auf audio/..., noten/... und TXT/...

Die Original-Dateien im Quellordner (OneDrive) bleiben dabei unangetastet
- es wird kopiert, nicht verschoben/gelöscht. Kopiert wird nur, was neu
ist oder sich geändert hat (Größe/Änderungsdatum). Unveränderte Dateien
werden nicht erneut geschrieben, damit OneDrive nicht bei jedem Lauf
alles neu hochlädt.

Haben MP3s aus unterschiedlichen "Partien aus ..."-Ordnern denselben
Dateinamen, würden sie sich in audio/ gegenseitig überschreiben. Das
Script erkennt das und stellt bei diesen Dateien die Szenen-Nummer vor
den Namen (z.B. "1.3 - Sopran.mp3" und "14 - Sopran.mp3"), mit einer
Warnung in der Konsole.

Die Datei "besetzung.xlsx" liegt ausschließlich im Git-Repo (DEST_DIR,
"12 GitHub/thadden27"). Aus ihr wird dort bei jedem Lauf die
"besetzung.csv" frisch erzeugt. Der Musik-Quellordner wird für die
Besetzungsdateien NICHT durchsucht. Die Besetzung kann also ganz normal
in Excel gepflegt werden. Dafür wird das Python-Paket "openpyxl"
gebraucht; fehlt es, bietet das Script an, es automatisch zu installieren
(fragt vorher in der Konsole nach Bestätigung). Fehlt die Excel-Datei
oder kann sie nicht umgewandelt werden, bricht das Script ab, OHNE etwas
zu kopieren oder hochzuladen - eine alte besetzung.csv wird bewusst
nicht ersatzweise verwendet.
Tipp: Die Abschnitts-Spalte in Excel als "Text" formatieren. Sonst macht
Excel z.B. aus "1.10" die Zahl 1,1 - das Script warnt in diesem Fall.

Außerdem prüft das Script anhand von "besetzung.csv" (Liste der
gültigen Abschnitte), ob es im Quellordner MP3s, PDFs oder TXTs gibt, die
KEINEM gültigen Abschnitt zugeordnet werden können (weil kein erkennbares
Zahlen-Präfix wie "1.1 ..." im Ordner-/Dateinamen steht, ODER weil die
erkannte Nummer gar nicht in besetzung.csv vorkommt). Diese Dateien
werden TROTZDEM ganz normal nach audio/, noten/ bzw. TXT/ kopiert (nur
eben ohne Eintrag in dateien.json) und am Ende als "nicht zugeordnet"
aufgelistet - das Script wartet dabei auf [Enter], um die Liste zu
bestätigen.

Danach prüft das Script, ob in audio/, noten/ bzw. TXT/ Dateien liegen,
die nicht (mehr) aus dem Quellordner stammen - z.B. weil eine Stimme
weggefallen ist, eine Datei umbenannt/gelöscht oder ein ganzer
"Partien aus ..."-Ordner entfernt wurde. Das gilt für ALLE Dateien in
diesen Ordnern, egal welcher Dateityp. Diese "verwaisten" Dateien werden
aufgelistet, und das Script fragt, ob sie gelöscht werden sollen.
Gelöscht wird dabei IMMER nur im Zielordner (Git-Repo), der Quellordner
(OneDrive) bleibt davon unberührt. In dateien.json tauchen verwaiste
Dateien in keinem Fall mehr auf - auch dann nicht, wenn man das Löschen
ablehnt (dann liegen sie nur noch ungenutzt im Repo).

Am Ende bietet das Script an, die Änderungen in DEST_DIR gleich zu
committen und zu pushen (fragt dafür in der Konsole nach Bestätigung
und nach der Commit-Nachricht). Gepusht wird der aktuell ausgecheckte
Branch (siehe GIT_BRANCH). DEST_DIR muss dafür ein bestehendes Git-Repo
mit eingerichtetem Remote sein (git clone / git remote add vorher
einmalig von Hand erledigen).

Läuft das Script ohne Tastatur-Eingabe (z.B. automatisch gestartet),
gilt bei allen Rückfragen "Nein": es wird dann nichts installiert,
nichts gelöscht und nichts gepusht.

ANPASSEN:
  - SOURCE_SUBPATH: Unterordner-Pfad INNERHALB von OneDrive, in dem der
    Quellordner liegt (ohne den Namen des OneDrive-Ordners selbst - der
    wird automatisch gefunden, siehe find_onedrive_dir()).
  - DEST_SUBPATH: dasselbe für den Zielordner (lokales Git-Repo).
  - Beide werden automatisch mit dem gefundenen OneDrive-Ordner zu
    SOURCE_DIR bzw. DEST_DIR zusammengesetzt - auf jedem Rechner passt
    sich das automatisch an, egal wie der OneDrive-Ordner genau heißt
    oder wer angemeldet ist.
  - IGNORE_FILENAMES: Dateinamen, die ganz normal kopiert werden (auch
    erneut, wenn sich was geändert hat), aber NICHT der Abschnitts-
    Zuordnung unterworfen werden: kein Eintrag in dateien.json, nicht in
    der "nicht zugeordnet"-Liste - z.B. für eine allgemeine Infodatei
    ohne Bezug zu einem einzelnen Abschnitt.
  - PDF_FOLDER: Name des PDF-Unterordners in SOURCE_DIR
  - TXT_FOLDER: Name des TXT-Unterordners in SOURCE_DIR
  - PARTIEN_PREFIX: Anfang der Audio-Unterordnernamen in SOURCE_DIR
  - GITHUB_AUDIO_PATH / GITHUB_NOTEN_PATH / GITHUB_TXT_PATH: Namen der
    Zielordner in DEST_DIR (gleichzeitig die Pfade, wie sie in
    dateien.json stehen)
  - GIT_AUTO_COMMIT_PUSH: auf False setzen, um den Commit/Push-Teil
    komplett zu deaktivieren (dann wird nur kopiert + dateien.json
    geschrieben)
  - GIT_REMOTE / GIT_BRANCH: wohin gepusht wird. GIT_BRANCH = None
    bedeutet: der aktuell ausgecheckte Branch.
  - BESETZUNG_XLSX / BESETZUNG_CSV: Die Dateien "besetzung.xlsx" und
    "besetzung.csv" liegen ausschließlich in DEST_DIR (12 GitHub/thadden27),
    NICHT im Musik-Quellordner.
"""

import os
import io
import re
import csv
import sys
import json
import shutil
import tempfile
import subprocess

# Windows-Konsolen benutzen oft NICHT UTF-8 als Codepage. Ohne diesen
# Block konnte das Script beim allerersten "print" mit einem Umlaut
# (ü/ö/ä/ß) sofort mit einem UnicodeEncodeError abstuerzen - und zwar
# so schnell, dass das Konsolenfenster nur kurz aufblitzt und sich ohne
# jede sichtbare Meldung wieder schliesst. Deshalb hier robust auf
# UTF-8 umschalten und nicht darstellbare Zeichen im Zweifel ersetzen
# statt abzustuerzen.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# -- KONFIGURATION ------------------------------------------------------

def find_onedrive_dir():
    """
    Sucht direkt im Benutzerordner (~, z.B. "C:\\Users\\flori") nach dem
    OneDrive-Ordner. Der heißt je nach Rechner/Konto unterschiedlich
    (z.B. "OneDrive", "OneDrive - thaddenschule.de", ...) - deshalb wird
    hier NICHT ein fest eingetragener Name erwartet, sondern nach allen
    Ordnern gesucht, deren Name mit "onedrive" beginnt (Groß-/
    Kleinschreibung egal).

    Gibt den vollständigen Pfad zum gefundenen OneDrive-Ordner zurück,
    oder None, wenn keiner gefunden wurde.

    - Genau ein Treffer -> der wird verwendet.
    - Mehrere Treffer -> der mit "thaddenschule" im Namen wird bevorzugt
      (Schul-Konto), sonst der alphabetisch erste - dabei wird eine
      Info-Meldung ausgegeben, damit das nicht unbemerkt bleibt.
    - Kein Treffer -> None (der Aufrufer gibt dann eine klare
      Fehlermeldung aus).
    """
    home = os.path.expanduser("~")
    try:
        entries = [
            name for name in os.listdir(home)
            if name.lower().startswith("onedrive") and os.path.isdir(os.path.join(home, name))
        ]
    except OSError:
        return None

    if not entries:
        return None

    if len(entries) > 1:
        preferred = [e for e in entries if "thaddenschule" in e.lower()]
        chosen = sorted(preferred)[0] if preferred else sorted(entries)[0]
        print(f"  [Info] Mehrere OneDrive-Ordner gefunden ({', '.join(sorted(entries))}) "
              f"- verwende '{chosen}'.")
    else:
        chosen = entries[0]

    return os.path.join(home, chosen)


_ONEDRIVE_DIR = find_onedrive_dir()

# Unterordner-Pfade INNERHALB des OneDrive-Ordners (ohne den Namen des
# OneDrive-Ordners selbst - der wird ja automatisch gefunden, siehe
# find_onedrive_dir() oben). Läuft das Script auf einem neuen Rechner,
# muss i.d.R. nichts angepasst werden, solange diese Unterordner-Struktur
# innerhalb von OneDrive gleich bleibt.
SOURCE_SUBPATH = ["09 Musik", "2025 Elisabeth"]
DEST_SUBPATH = ["12 GitHub", "thadden27"]

if _ONEDRIVE_DIR:
    # Quellordner: hier liegen die "Partien aus ..."-Ordner und "pdf/".
    SOURCE_DIR = os.path.join(_ONEDRIVE_DIR, *SOURCE_SUBPATH)
    # Zielordner: lokales Git-Repo, in dem audio/, noten/ und dateien.json
    # landen sollen - liegt ebenfalls in OneDrive.
    DEST_DIR = os.path.join(_ONEDRIVE_DIR, *DEST_SUBPATH)
else:
    # Kein OneDrive-Ordner gefunden - Pfade trotzdem zusammenbauen (unter
    # einem klar erkennbaren Platzhalter-Namen), damit main() weiter unten
    # mit einer verständlichen Fehlermeldung abbricht, statt mit einem
    # rohen Python-Fehler abzustürzen.
    SOURCE_DIR = os.path.join(os.path.expanduser("~"), "OneDrive-NICHT-GEFUNDEN", *SOURCE_SUBPATH)
    DEST_DIR = os.path.join(os.path.expanduser("~"), "OneDrive-NICHT-GEFUNDEN", *DEST_SUBPATH)

# Lokale Ordnernamen (innerhalb SOURCE_DIR)
PDF_FOLDER     = "pdf"           # Ordner mit allen PDF-Dateien
TXT_FOLDER     = "TXT"           # Ordner mit allen TXT-Dateien
PARTIEN_PREFIX = "Partien aus "  # Unterordner-Präfix für MP3-Ordner

# Dateinamen (nur der Dateiname, ohne Pfad - Groß-/Kleinschreibung egal),
# die das Script GANZ NORMAL kopiert (auch erneut, wenn sich was geändert
# hat), aber NICHT der Abschnitts-Zuordnung unterwirft: kein Eintrag in
# dateien.json, taucht nicht in der "nicht zugeordnet"-Liste auf. Für
# allgemeine Dateien ohne Bezug zu einem einzelnen Abschnitt, z.B. eine
# Datei mit generellen Infos zum Stück statt zu einer Szene.
IGNORE_FILENAMES = {
    "Generelle Informationen zum Stück.txt",
}

# Die Datei mit den generellen Infos zum Stück wird zusätzlich als "_info"
# in dateien.json eingetragen - die Website zeigt sie dann oben in der
# Szenenliste als eigene Karte an.
INFO_FILENAME = "Generelle Informationen zum Stück.txt"

# Liedtexte: TXT-Dateien, deren Name auf "_Lyrics.txt" endet (Groß-/Klein-
# schreibung egal), z.B. "2.1_Lyrics.txt" oder "2.1 Geh aus mein Herz_Lyrics.txt".
# Sie landen in dateien.json NICHT unter "txt" (Beschreibung), sondern unter
# "lyrics" - die Website zeigt sie auf der Szenen-Seite als eigenen Abschnitt.
LYRICS_PATTERN = re.compile(r'[_\s-]lyrics\.txt$', re.IGNORECASE)


def is_lyrics_file(filename):
    """True, wenn der Dateiname auf '_Lyrics.txt' endet (siehe LYRICS_PATTERN)."""
    return bool(LYRICS_PATTERN.search(os.path.basename(filename)))


_IGNORE_FILENAMES_LOWER = {n.lower() for n in IGNORE_FILENAMES}

# Namen der Zielordner in DEST_DIR - gleichzeitig die Pfade, wie sie in
# der dateien.json stehen sollen (z.B. "audio/1.1 Marsch.mp3")
GITHUB_AUDIO_PATH = "audio"
GITHUB_NOTEN_PATH = "noten"
GITHUB_TXT_PATH   = "TXT"

# Vollständige lokale Zielpfade, in die tatsächlich kopiert wird
AUDIO_OUT_DIR = os.path.join(DEST_DIR, GITHUB_AUDIO_PATH)
NOTEN_OUT_DIR = os.path.join(DEST_DIR, GITHUB_NOTEN_PATH)
TXT_OUT_DIR   = os.path.join(DEST_DIR, GITHUB_TXT_PATH)

# Ausgabedatei (liegt im Ziel-Ordner, direkt neben audio/ und noten/)
OUTPUT_FILE = os.path.join(DEST_DIR, "dateien.json")

# Nach dem Kopieren in der Konsole fragen, ob committet + gepusht werden
# soll (inkl. Abfrage der Commit-Nachricht). False = nur kopieren, kein Git.
GIT_AUTO_COMMIT_PUSH = True
# Wohin gepusht wird. GIT_BRANCH = None -> aktuell ausgecheckter Branch.
GIT_REMOTE = "origin"
GIT_BRANCH = None

# Die Besetzungsdateien liegen AUSSCHLIESSLICH in DEST_DIR, also im lokalen
# Git-Repo "12 GitHub/thadden27". Die Excel-Datei ist die Quelle; die CSV
# wird daneben bei jedem Lauf frisch erzeugt.
BESETZUNG_XLSX = os.path.join(DEST_DIR, "besetzung.xlsx")
BESETZUNG_CSV = os.path.join(DEST_DIR, "besetzung.csv")

# Präfix für Zwischendateien beim Kopieren. Bleiben solche Dateien nach
# einem Absturz liegen, werden sie beim nächsten Lauf ohne Nachfrage
# entfernt (siehe find_orphans()).
TMP_PREFIX = ".tmp_kopie_"

# Systemdateien, die Windows/macOS/OneDrive selbst in Ordnern anlegen -
# werden beim Aufräumen nie als "verwaist" gemeldet.
SYSTEM_FILENAMES = {"desktop.ini", "thumbs.db", ".ds_store"}
# ----------------------------------------------------------------------

# Zählt Probleme und Kopiervorgänge während des Laufs, damit am Ende eine
# ehrliche Zusammenfassung möglich ist.
STATS = {"copy_errors": 0, "delete_errors": 0,
         "copied_new": 0, "copied_updated": 0, "unchanged": 0}

# Spaltennamen (klein geschrieben), die als "Abschnitt-Spalte" in
# besetzung.xlsx/.csv erkannt werden.
ABSCHNITT_SPALTENNAMEN = (
    "abschnitt", "abschnitt-id", "abschnittid", "abschnittsnummer",
    "szene", "szene-id", "szenenid", "id", "nr", "nummer",
)


def ask_yes_no(question, default_yes):
    """
    Stellt eine Ja/Nein-Frage in der Konsole. Leere Eingabe = default_yes.
    Gibt es keine Tastatur-Eingabe (EOF, z.B. automatisch gestartet), gilt
    IMMER "Nein" - damit nie unbeaufsichtigt installiert, gelöscht oder
    gepusht wird.
    """
    hint = "[J/n]" if default_yes else "[j/N]"
    try:
        answer = input(f"{question} {hint}: ").strip().lower()
    except EOFError:
        print()
        return False
    if answer == "":
        return default_yes
    return answer in ("j", "ja", "y", "yes")


def write_text_atomic(path, text, encoding="utf-8"):
    """
    Schreibt 'text' nach 'path', und zwar erst in eine Zwischendatei im
    selben Ordner, die dann in einem Schritt umbenannt wird. Bricht das
    Schreiben ab (Absturz, voller Datenträger, ...), bleibt die alte Datei
    unversehrt statt halb geschrieben zurück.
    """
    folder = os.path.dirname(path) or "."
    tmp_fd, tmp_path = tempfile.mkstemp(dir=folder, prefix=".tmp_", suffix=".part")
    try:
        with os.fdopen(tmp_fd, "w", encoding=encoding, newline="") as f:
            f.write(text)
        os.replace(tmp_path, path)
    except Exception:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise


def _cell_to_str(value):
    """
    Wandelt einen Excel-Zellwert in einen sauberen String für die CSV um.
    - None -> "" (leere Zelle)
    - Ganzzahlige Floats -> ohne ".0" (Excel liefert Zahlen fast immer als
      float, "14" würde sonst als "14.0" rausfallen)
    - datetime/date -> "Tag.Monat", falls Excel eine eingetragene Zahl wie
      "1.1" versehentlich als Datum interpretiert hat (bei europäischen
      Excel-Einstellungen wird "1.1" oft zu "1. Januar"). Falls das bei
      euch falsche Ergebnisse liefert: die Abschnitts-Spalte in Excel
      vorher als "Text" formatieren, dann bleibt sie unangetastet.
    """
    if value is None:
        return ""
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value)
    if hasattr(value, "day") and hasattr(value, "month"):
        return f"{value.day}.{value.month}"
    return str(value)


def find_id_col(header_row):
    """
    Sucht in der Kopfzeile die Spalte mit der Abschnitts-ID (per
    Spaltenname, siehe ABSCHNITT_SPALTENNAMEN, sonst jede Spalte mit
    "abschnitt"/"szene" im Namen). Gibt den Spaltenindex zurück oder None.
    """
    header = [("" if h is None else str(h)).strip().lower() for h in header_row]
    idx = next((i for i, name in enumerate(header) if name in ABSCHNITT_SPALTENNAMEN), None)
    if idx is None:
        idx = next((i for i, name in enumerate(header)
                    if "abschnitt" in name or "szene" in name), None)
    return idx


def warn_numeric_scene_ids(raw_rows):
    """
    Warnt, wenn Abschnitts-Nummern in der Excel-Datei als Zahl oder Datum
    statt als Text gespeichert sind. Excel macht dann z.B. aus "1.10" die
    Zahl 1,1 - und "1.10" wäre von "1.1" nicht mehr zu unterscheiden.
    Ganze Zahlen (z.B. 14) sind unproblematisch und werden nicht gemeldet.
    """
    if not raw_rows:
        return
    id_col = find_id_col(raw_rows[0])
    if id_col is None:
        id_col = 0

    suspicious = []
    for rownum, row in enumerate(raw_rows[1:], start=2):
        if id_col >= len(row):
            continue
        v = row[id_col]
        if isinstance(v, bool):
            continue
        if isinstance(v, float) and not v.is_integer():
            suspicious.append((rownum, _cell_to_str(v), "Zahl"))
        elif hasattr(v, "day") and hasattr(v, "month"):
            suspicious.append((rownum, _cell_to_str(v), "Datum"))

    if not suspicious:
        return

    print(f"  [WARNUNG] {len(suspicious)} Abschnitts-Nummer(n) in 'besetzung.xlsx' sind als")
    print("     Zahl/Datum statt als Text gespeichert. Excel kann dabei z.B. aus")
    print("     \"1.10\" die Zahl 1,1 machen - dann wäre 1.10 von 1.1 nicht mehr")
    print("     zu unterscheiden. Bitte die Spalte in Excel als \"Text\" formatieren")
    print("     und die Nummern neu eintippen. Betroffen:")
    for rownum, shown, kind in suspicious[:10]:
        print(f"       - Zeile {rownum}: {shown}  ({kind})")
    if len(suspicious) > 10:
        print(f"       ... und {len(suspicious) - 10} weitere")


def ensure_openpyxl():
    """
    Versucht "import openpyxl" - falls das Paket fehlt, wird in der
    Konsole gefragt, ob es automatisch installiert werden soll (per
    "pip install openpyxl" mit demselben Python, das gerade läuft).
    Gibt das openpyxl-Modul zurück, oder None, wenn es weder vorhanden
    ist noch installiert werden konnte/sollte.
    """
    try:
        import openpyxl
        return openpyxl
    except ImportError:
        pass

    print("  [Info]  Paket 'openpyxl' wird für die xlsx->csv-Umwandlung gebraucht,")
    print("     ist aber nicht installiert.")
    if not ask_yes_no("  Jetzt automatisch installieren (pip install openpyxl)?", default_yes=True):
        print("  Übersprungen. Von Hand installierbar mit:")
        print(f"       {sys.executable} -m pip install openpyxl")
        return None

    print("  -> installiere openpyxl ...")
    try:
        proc = subprocess.run(
            [sys.executable, "-m", "pip", "install", "openpyxl"],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=120,
        )
    except (OSError, subprocess.TimeoutExpired) as e:
        print(f"  [FEHLER] Installation fehlgeschlagen: {e}")
        return None

    if proc.returncode != 0:
        print("  [FEHLER] 'pip install openpyxl' ist fehlgeschlagen:")
        print(f"     {(proc.stdout or '') + (proc.stderr or '')}".strip())
        print("     Von Hand versuchen mit:")
        print(f"       {sys.executable} -m pip install openpyxl")
        return None

    try:
        import openpyxl
        print("  [OK] openpyxl installiert.")
        return openpyxl
    except ImportError as e:
        print(f"  [FEHLER] openpyxl wurde installiert, kann aber nicht geladen werden: {e}")
        return None


def convert_xlsx_to_csv(xlsx_path, csv_path):
    """
    Wandelt das erste/aktive Tabellenblatt von xlsx_path in csv_path um
    (UTF-8, Semikolon als Trennzeichen). Gibt True bei Erfolg zurück,
    sonst False (mit erklärender Meldung in der Konsole).
    """
    openpyxl = ensure_openpyxl()
    if openpyxl is None:
        print(f"     '{os.path.basename(xlsx_path)}' kann so nicht in eine CSV")
        print("     umgewandelt werden.")
        return False

    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
        try:
            ws = wb.active
            raw_rows = [list(row) for row in ws.iter_rows(values_only=True)]
        finally:
            wb.close()
    except Exception as e:
        print(f"  [FEHLER] Konnte '{xlsx_path}' nicht lesen: {e}")
        return False

    warn_numeric_scene_ids(raw_rows)
    rows = [[_cell_to_str(v) for v in row] for row in raw_rows]

    buf = io.StringIO()
    csv.writer(buf, delimiter=";", lineterminator="\r\n").writerows(rows)
    try:
        write_text_atomic(csv_path, buf.getvalue(), encoding="utf-8-sig")
    except OSError as e:
        print(f"  [FEHLER] Konnte '{csv_path}' nicht schreiben: {e}")
        return False

    print(f"  [OK] '{os.path.basename(xlsx_path)}' -> '{os.path.basename(csv_path)}'")
    print(f"     umgewandelt ({len(rows)} Zeile(n)).")
    return True


def sync_besetzung_csv_from_xlsx():
    """Erzeugt IMMER besetzung.csv aus der Excel-Datei im Git-Repo.

    Dabei wird nicht auf Dateizeitstempel vertraut (OneDrive und Excel
    koennen Zeitstempel unterschiedlich setzen). Fehlt die XLSX oder
    schlaegt die Konvertierung fehl, ist dies ein Fehler und das gesamte
    Skript darf keine alte CSV stillschweigend weiterverwenden.
    """
    if not os.path.isfile(BESETZUNG_XLSX):
        print(f"  [FEHLER] Excel-Quelldatei fehlt: {BESETZUNG_XLSX}")
        print("     Bitte besetzung.xlsx in den GitHub-Ordner thadden27 legen.")
        print("     Die alte besetzung.csv wird NICHT ersatzweise verwendet.")
        return False

    print("  [Info] 'besetzung.xlsx' gefunden - CSV wird immer frisch erzeugt.")
    return convert_xlsx_to_csv(BESETZUNG_XLSX, BESETZUNG_CSV)


def load_valid_scene_ids():
    """
    Erzeugt besetzung.csv frisch aus besetzung.xlsx und liest daraus die
    gültigen Abschnitts-IDs (Spalte per Name erkannt, sonst erste Spalte).

    Gibt (valid_ids, csv_path) zurück:
      valid_ids = set() der gültigen IDs. Ist es LEER (Besetzung enthält
                  keine erkennbaren Nummern), wird NICHT gegen die
                  Besetzung geprüft - jede per Zahlen-Präfix erkannte ID
                  gilt dann als gültig.
                  None, wenn die Excel-Datei fehlt oder nicht umgewandelt
                  werden konnte -> main() bricht dann ab.
      csv_path  = Pfad der benutzten CSV, oder None.
    """
    if not sync_besetzung_csv_from_xlsx():
        return None, None

    # Die CSV hat das Script gerade selbst geschrieben (UTF-8, ";") - die
    # Kodierung oder das Trennzeichen muss hier also nicht geraten werden.
    try:
        with open(BESETZUNG_CSV, "r", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f, delimiter=";"))
    except OSError as e:
        print(f"  [FEHLER] Konnte '{BESETZUNG_CSV}' nicht lesen: {e}")
        return None, None

    name = os.path.basename(BESETZUNG_CSV)
    if not rows or not rows[0]:
        print(f"  [WARNUNG]  '{name}' ist leer - Abschnitts-Zuordnung wird NICHT")
        print("     gegen eine Besetzung geprüft.")
        return set(), BESETZUNG_CSV

    id_col = find_id_col(rows[0])
    if id_col is None:
        id_col = 0
        print(f"  [Info]  Keine Spalte 'Abschnitt'/'Szene' in '{name}'")
        print(f"     erkannt - benutze stattdessen die erste Spalte ('{rows[0][0]}').")
    else:
        print(f"  [Info]  Abschnitts-Spalte in '{name}': '{rows[0][id_col]}'")

    valid_ids = set()
    for row in rows[1:]:
        if id_col >= len(row):
            continue
        m = re.match(r'^(\d+(?:\.\d+)?)', row[id_col].strip())
        if m:
            valid_ids.add(m.group(1))

    if not valid_ids:
        print(f"  [WARNUNG]  Keine Abschnitts-Nummern in '{name}' gefunden -")
        print("     Abschnitts-Zuordnung wird NICHT gegen die Besetzung geprüft.")
    else:
        print(f"  [OK] {len(valid_ids)} gültige Abschnitts-ID(s) aus '{name}' geladen.")
    return valid_ids, BESETZUNG_CSV


def paths_overlap(a, b):
    """
    True, wenn die Ordner a und b identisch sind ODER einer der beiden im
    anderen liegt (z.B. DEST_DIR irgendwo unterhalb von SOURCE_DIR). Wird
    benutzt, um VOR jedem Kopieren sicherzustellen, dass Quell- und
    Zielordner garantiert getrennt sind - eine Überschneidung wäre
    gefährlich, weil dann versehentlich Dateien im Quellordner als
    "Ziel" landen könnten (siehe copy_file() für den Schutz auf
    Einzeldatei-Ebene).
    """
    a = os.path.normcase(os.path.normpath(os.path.abspath(a)))
    b = os.path.normcase(os.path.normpath(os.path.abspath(b)))
    return a == b or a.startswith(b + os.sep) or b.startswith(a + os.sep)


def extract_scene_id(filename):
    """
    Extrahiert die Szenen-ID aus einem Datei- oder Ordnernamen.
    Beispiele:
      "1.1 Marsch.mp3"              -> "1.1"
      "1.3 Nun will... - Flöte.pdf" -> "1.3"
      "14 Finale - Partitur.pdf"    -> "14"
      "3.2 Eine Heldin A.mp3"       -> "3.2"
    """
    # Muster: optional "Partien aus " am Anfang, dann Ziffern[.Ziffern]
    name = filename
    if name.startswith(PARTIEN_PREFIX):
        name = name[len(PARTIEN_PREFIX):]

    m = re.match(r'^(\d+(?:\.\d+)?)', name.strip())
    if m:
        return m.group(1)
    return None


def is_unchanged(src_path, dest_path):
    """
    True, wenn dest_path schon dieselbe Größe und dasselbe Änderungsdatum
    wie src_path hat (beim Kopieren wird das Änderungsdatum mitkopiert).
    2 Sekunden Toleranz, weil manche Dateisysteme Zeiten nur grob speichern.
    """
    try:
        s = os.stat(src_path)
        d = os.stat(dest_path)
    except OSError:
        return False
    return s.st_size == d.st_size and abs(s.st_mtime - d.st_mtime) <= 2


def copy_file(src_path, dest_dir, filename):
    """
    Kopiert eine Datei als 'filename' nach dest_dir (legt den Ordner bei
    Bedarf an) - aber nur, wenn sie dort noch fehlt oder sich geändert hat.
    Bricht bei einem Kopierfehler NICHT das ganze Script ab, sondern
    meldet den Fehler und zählt ihn für die Abschluss-Zusammenfassung.

    Gibt zurück: "neu", "aktualisiert", "unverändert" oder None (Fehler).

    WICHTIG - Schutz vor Datenverlust im Quellordner:
    Es wird NIE direkt in die Zieldatei geschrieben, sondern immer erst in
    eine NEUE Zwischendatei im Zielordner (kann also unmöglich mit der
    Quelle identisch sein), die danach in einem Schritt (os.replace) auf
    den echten Zielnamen umbenannt wird. Selbst falls Quelle und Ziel durch
    einen Pfad-Sonderfall (Groß-/Kleinschreibung, OneDrive-Verknüpfung, ...)
    doch dieselbe Datei sein sollten, wird die Quelle dadurch nie geleert.
    """
    dest_path = os.path.join(dest_dir, filename)
    try:
        os.makedirs(dest_dir, exist_ok=True)
        already_existed = os.path.exists(dest_path)
        if already_existed and is_unchanged(src_path, dest_path):
            STATS["unchanged"] += 1
            return "unverändert"

        tmp_fd, tmp_path = tempfile.mkstemp(
            dir=dest_dir, prefix=TMP_PREFIX, suffix=os.path.splitext(filename)[1]
        )
        os.close(tmp_fd)
        try:
            shutil.copyfile(src_path, tmp_path)
            try:
                # Änderungsdatum mitkopieren - wird für is_unchanged() gebraucht.
                shutil.copystat(src_path, tmp_path)
            except OSError:
                pass  # dann wird die Datei beim nächsten Lauf eben erneut kopiert
            os.replace(tmp_path, dest_path)
        except Exception:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
            raise
    except OSError as e:
        STATS["copy_errors"] += 1
        print(f"     [FEHLER] Fehler beim Kopieren von '{filename}': {e}")
        return None

    if already_existed:
        STATS["copied_updated"] += 1
        return "aktualisiert"
    STATS["copied_new"] += 1
    return "neu"


def _status_suffix(status):
    """Kurzer Zusatz für die Konsolenausgabe je nach Ergebnis von copy_file()."""
    if status is None:
        return "  (FEHLER beim Kopieren)"
    if status == "unverändert":
        return ""
    return f"  ({status})"


def collect_audio(valid_ids):
    """
    Sucht MP3s in allen "Partien aus X.X NAME"-Unterordnern von
    SOURCE_DIR und kopiert ALLE (flach, ohne Unterordner) nach
    AUDIO_OUT_DIR - auch die, die keinem gültigen Abschnitt zugeordnet
    werden können. Ein Ordner gilt als gültig zugeordnet, wenn sein Name
    eine Szenen-ID enthält UND (valid_ids leer ist ODER diese ID in
    valid_ids vorkommt); nur dann landen seine MP3s auch im result-Dict
    (also in dateien.json).

    Kommt derselbe Dateiname in mehreren Ordnern vor, wird bei ALLEN
    betroffenen Dateien die Szenen-Nummer vorangestellt
    ("1.3 - Sopran.mp3"), damit sich nichts gegenseitig überschreibt.

    Gibt zurück:
      (result, expected, unassigned)
      result     = {szene_id: ["audio/dateiname.mp3", ...], ...}
      expected   = set der Dateinamen (klein geschrieben), die nach diesem
                   Lauf in audio/ liegen sollen - alles andere dort ist
                   "verwaist" (siehe find_orphans()). None, wenn es gar
                   keinen "Partien aus ..."-Ordner gibt.
      unassigned = Liste von (ordnername, dateiname) für MP3s ohne
                   gültigen Abschnitt (werden trotzdem kopiert, aber nicht
                   in dateien.json aufgenommen).
    """
    result = {}
    expected = set()
    unassigned = []

    # 1. Erst alles einsammeln (noch nichts kopieren), damit doppelte
    #    Dateinamen über mehrere Ordner hinweg erkannt werden können.
    folders = []
    for entry in sorted(os.scandir(SOURCE_DIR), key=lambda x: x.name.lower()):
        if not entry.is_dir() or not entry.name.startswith(PARTIEN_PREFIX):
            continue

        sid = extract_scene_id(entry.name)
        sid_ok = sid is not None and (not valid_ids or sid in valid_ids)
        if sid is None:
            print(f"  [WARNUNG]  Ordner ohne erkennbare ID (wird trotzdem kopiert): {entry.name}")
        elif not sid_ok:
            print(f"  [WARNUNG]  Abschnitt '{sid}' nicht in besetzung.csv - wird trotzdem")
            print(f"     kopiert, aber nicht in dateien.json aufgenommen: {entry.name}")

        mp3s = sorted(
            (f for f in os.scandir(entry.path) if f.is_file() and f.name.lower().endswith('.mp3')),
            key=lambda x: x.name.lower(),
        )
        folders.append((entry, sid, sid_ok, mp3s))

    if not folders:
        print(f"  [WARNUNG]  Kein einziger '{PARTIEN_PREFIX}...'-Ordner gefunden - keine MP3s.")
        return result, None, unassigned

    name_count = {}
    for _, _, _, mp3s in folders:
        for f in mp3s:
            key = f.name.lower()
            name_count[key] = name_count.get(key, 0) + 1
    duplicates = {k for k, n in name_count.items() if n > 1}

    # 2. Kopieren und verzeichnen
    for entry, sid, sid_ok, mp3s in folders:
        files = []
        for f in mp3s:
            target = f.name
            if f.name.lower() in duplicates:
                tag = sid if sid else entry.name[len(PARTIEN_PREFIX):].strip()
                target = f"{tag} - {f.name}"
                print(f"  [WARNUNG]  '{f.name}' gibt es in mehreren Ordnern - wird als")
                print(f"     '{target}' gespeichert (Ordner: {entry.name}).")
            if target.lower() in expected:
                # Sehr unwahrscheinlich, aber sicher ist sicher: Name ist trotzdem schon vergeben
                base, ext = os.path.splitext(target)
                n = 2
                while f"{base} ({n}){ext}".lower() in expected:
                    n += 1
                target = f"{base} ({n}){ext}"
                print(f"     -> Name schon vergeben, stattdessen: '{target}'")
            expected.add(target.lower())

            status = copy_file(f.path, AUDIO_OUT_DIR, target)
            suffix = _status_suffix(status)

            if f.name.lower() in _IGNORE_FILENAMES_LOWER:
                print(f"  [Audio] (allgemein, ohne Abschnitts-Zuordnung): {target}{suffix}")
                continue

            if sid_ok:
                files.append(f"{GITHUB_AUDIO_PATH}/{target}")
                print(f"  [Audio] {sid}: {target}{suffix}")
            else:
                unassigned.append((entry.name, target))
                print(f"  [Audio] (nicht zugeordnet): {target}{suffix}")

        if files:
            result.setdefault(sid, []).extend(files)

    return result, expected, unassigned


def collect_flat_folder(source_subfolder, extension, out_dir, github_path_prefix,
                         valid_ids, label):
    """
    Generische Variante für einen FLACHEN Quell-Unterordner (keine weiteren
    Unterordner darin) mit einem bestimmten Dateityp - wird sowohl für
    Noten (PDF) als auch für Text-Dateien (TXT) benutzt. Kopiert ALLE
    passenden Dateien nach out_dir - auch die, die keinem gültigen
    Abschnitt zugeordnet werden können (siehe collect_audio für die Logik
    von valid_ids).

    source_subfolder   = Name des Unterordners in SOURCE_DIR (z.B. "pdf", "TXT")
    extension           = Dateiendung, klein geschrieben, mit Punkt (z.B. ".pdf")
    out_dir             = lokaler Zielordner, in den kopiert wird
    github_path_prefix  = Prefix für den Pfad in dateien.json (z.B. "noten")
    valid_ids           = siehe collect_audio()
    label               = Text für Konsolenausgaben (z.B. "Noten", "TXT")

    Gibt zurück:
      (result, expected, unassigned)
      result     = {szene_id: ["<prefix>/dateiname.ext", ...], ...}
      expected   = set der Dateinamen (klein geschrieben), die nach diesem
                   Lauf in out_dir liegen sollen, oder None, falls der
                   Quell-Unterordner gar nicht existiert.
      unassigned = Liste von Dateinamen ohne gültigen Abschnitt (werden
                   trotzdem kopiert, aber nicht in dateien.json aufgenommen).
    """
    result = {}
    src_dir = os.path.join(SOURCE_DIR, source_subfolder)

    if not os.path.isdir(src_dir):
        print(f"  [WARNUNG]  Kein '{source_subfolder}'-Ordner gefunden - keine {label}-Dateien.")
        return result, None, []

    expected = set()
    unassigned = []

    for f in sorted(os.scandir(src_dir), key=lambda x: x.name.lower()):
        if not f.is_file():
            continue
        if not f.name.lower().endswith(extension):
            continue

        expected.add(f.name.lower())
        status = copy_file(f.path, out_dir, f.name)
        suffix = _status_suffix(status)

        if f.name.lower() in _IGNORE_FILENAMES_LOWER:
            print(f"  [{label}] (allgemein, ohne Abschnitts-Zuordnung): {f.name}{suffix}")
            continue

        sid = extract_scene_id(f.name)
        sid_ok = sid is not None and (not valid_ids or sid in valid_ids)

        if not sid_ok:
            if sid is None:
                print(f"  [WARNUNG]  {label}-Datei ohne erkennbare ID (wird trotzdem kopiert): {f.name}{suffix}")
            else:
                print(f"  [WARNUNG]  Abschnitt '{sid}' nicht in besetzung.csv - '{f.name}' wird")
                print(f"     trotzdem kopiert, aber nicht in dateien.json aufgenommen.{suffix}")
            unassigned.append(f.name)
            continue

        result.setdefault(sid, []).append(f"{github_path_prefix}/{f.name}")
        print(f"  [{label}] {sid}: {f.name}{suffix}")

    return result, expected, unassigned


def collect_noten(valid_ids):
    """Sucht PDFs in PDF_FOLDER und kopiert sie nach NOTEN_OUT_DIR (siehe collect_flat_folder)."""
    return collect_flat_folder(PDF_FOLDER, '.pdf', NOTEN_OUT_DIR, GITHUB_NOTEN_PATH,
                                valid_ids, "Noten")


def collect_txt(valid_ids):
    """Sucht TXT-Dateien in TXT_FOLDER und kopiert sie nach TXT_OUT_DIR (siehe collect_flat_folder)."""
    return collect_flat_folder(TXT_FOLDER, '.txt', TXT_OUT_DIR, GITHUB_TXT_PATH,
                                valid_ids, "TXT")


def report_unassigned(unassigned_audio, unassigned_noten, unassigned_txt):
    """
    Zeigt MP3s/PDFs/TXTs im Quellordner an, die keinem gültigen Abschnitt
    zugeordnet werden konnten. Sie wurden trotzdem kopiert, stehen aber
    NICHT in dateien.json. Wartet danach auf [Enter], um die Liste zu
    bestätigen (reine Anzeige, es wird nichts verändert).
    """
    print(f"\n-- Nicht zugeordnete Dateien ---------------------------")

    if not unassigned_audio and not unassigned_noten and not unassigned_txt:
        print("  [OK] Alle MP3s, PDFs und TXT-Dateien im Quellordner konnten")
        print("    einem Abschnitt zugeordnet werden.")
        return

    if unassigned_audio:
        print(f"  [Audio] {len(unassigned_audio)} MP3-Datei(en) ohne gültigen Abschnitt:")
        for ordner, name in unassigned_audio:
            print(f"    - {name}   (Ordner: {ordner})")

    if unassigned_noten:
        print(f"  [Noten] {len(unassigned_noten)} PDF-Datei(en) ohne gültigen Abschnitt:")
        for name in unassigned_noten:
            print(f"    - {name}")

    if unassigned_txt:
        print(f"  [TXT] {len(unassigned_txt)} TXT-Datei(en) ohne gültigen Abschnitt:")
        for name in unassigned_txt:
            print(f"    - {name}")

    print("\n  Diese Dateien wurden trotzdem ganz normal nach audio/, noten/")
    print("  bzw. TXT/ kopiert - sie tauchen nur NICHT in dateien.json auf.")
    print("  Um sie zuzuordnen: Datei-/Ordnernamen im Quellordner mit der")
    print(f"  passenden Abschnitts-Nummer beginnen lassen (z.B. \"1.1 ...\",")
    print(f"  bei Audio-Ordnern mit Präfix \"{PARTIEN_PREFIX}\") UND sicherstellen,")
    print("  dass diese Nummer auch in besetzung.xlsx vorkommt.")

    try:
        input("\n  [Enter] um diese Liste zu bestätigen und fortzufahren...")
    except EOFError:
        print()


def find_orphans(out_dir, expected):
    """
    Findet ALLE Dateien in out_dir (egal welcher Dateityp), deren Name
    nicht in 'expected' vorkommt - also Dateien, die bei diesem Lauf nicht
    aus dem Quellordner kopiert wurden (Stimme weggefallen, umbenannt,
    gelöscht, Ordner entfernt, ...). Groß-/Kleinschreibung wird ignoriert.

    Ausnahmen: versteckte Dateien (Name beginnt mit ".") und Systemdateien
    (siehe SYSTEM_FILENAMES) werden nie gemeldet. Liegengebliebene
    Zwischendateien eines abgebrochenen Kopiervorgangs (TMP_PREFIX) werden
    ohne Nachfrage entfernt - sie stammen vom Script selbst.
    """
    if not os.path.isdir(out_dir):
        return []
    orphans = []
    for f in sorted(os.scandir(out_dir), key=lambda x: x.name.lower()):
        if not f.is_file():
            continue
        if f.name.startswith(TMP_PREFIX):
            try:
                os.remove(f.path)
                print(f"  [Info] Liegengebliebene Zwischendatei entfernt: {f.name}")
            except OSError as e:
                print(f"  [WARNUNG] Zwischendatei konnte nicht entfernt werden: {f.name} ({e})")
            continue
        if f.name.startswith(".") or f.name.lower() in SYSTEM_FILENAMES:
            continue
        if f.name.lower() not in expected:
            orphans.append(f.name)
    return orphans


def handle_orphans(label, out_dir, orphans, source_missing=False):
    """
    Zeigt verwaiste Dateien (im Zielordner, aber nicht mehr im Quellordner
    vorhanden) an und fragt in der Konsole, ob sie gelöscht werden sollen.
    Löscht NUR im Zielordner (DEST_DIR) - der Quellordner wird nie
    angefasst. source_missing=True: der Quell-Unterordner fehlt komplett,
    dann wird besonders deutlich gewarnt.
    """
    if not orphans:
        return

    print(f"\n  {label}: {len(orphans)} Datei(en) in '{out_dir}',")
    print("  die im Quellordner nicht mehr gefunden wurden:")
    for name in orphans:
        print(f"    - {name}")
    if source_missing:
        print("  [ACHTUNG] Der passende Quell-Ordner fehlt KOMPLETT. Falls OneDrive")
        print("     noch nicht fertig synchronisiert ist, bitte NICHT löschen.")

    if not ask_yes_no(f"  Diese {len(orphans)} Datei(en) jetzt löschen?", default_yes=False):
        print("  Übersprungen - nichts gelöscht (in dateien.json stehen sie trotzdem nicht).")
        return

    deleted = 0
    for name in orphans:
        path = os.path.join(out_dir, name)
        try:
            os.remove(path)
            deleted += 1
            print(f"    [gelöscht]  {name}")
        except OSError as e:
            STATS["delete_errors"] += 1
            print(f"    [FEHLER] konnte nicht gelöscht werden: {name} ({e})")
    print(f"  [OK] {deleted} von {len(orphans)} Datei(en) gelöscht.")


def run_git(args, cwd, timeout=None):
    """
    Führt einen git-Befehl aus. Gibt (erfolgreich: bool, ausgabe: str) zurück.
    'ausgabe' enthält bei Misserfolg immer eine lesbare Fehlermeldung -
    auch wenn git selbst gar nicht erreichbar war/nicht reagiert hat.
    """
    try:
        proc = subprocess.run(
            ["git"] + args,
            cwd=cwd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
    except FileNotFoundError:
        return False, "git wurde nicht gefunden. Ist Git installiert und im PATH?"
    except subprocess.TimeoutExpired:
        return False, (
            f"Zeitüberschreitung nach {timeout}s - es kam keine Antwort. "
            "Vermutlich besteht keine Verbindung zum Git-Server (z.B. "
            "GitHub). Bitte Internetverbindung / VPN / Proxy prüfen."
        )
    output = ((proc.stdout or "") + (proc.stderr or "")).strip()
    return proc.returncode == 0, output


def ignore_excel_lockfiles_in_git():
    """Excel-Sperrdateien (~$...) ignorieren und bereits getrackte entfernen.

    'git rm --cached' entfernt die Dateien NUR aus Git, nicht vom Datenträger.
    Dadurch werden temporäre OneDrive-/Excel-Dateien nicht veröffentlicht.
    """
    ignore_path = os.path.join(DEST_DIR, ".gitignore")
    try:
        if os.path.isfile(ignore_path):
            with open(ignore_path, "r", encoding="utf-8-sig") as f:
                existing = f.read()
        else:
            existing = ""
        rule = "~$*.xls*"
        if rule not in existing.splitlines():
            with open(ignore_path, "a", encoding="utf-8", newline="\n") as f:
                if existing and not existing.endswith(("\n", "\r")):
                    f.write("\n")
                f.write("\n# Temporäre Excel-Sperrdateien nicht hochladen\n" + rule + "\n")
            print("  [Git] Excel-Sperrdateien in .gitignore eingetragen.")
    except OSError as e:
        print(f"  [FEHLER] Konnte .gitignore nicht aktualisieren: {e}")
        return False

    ok, out = run_git(["ls-files", "-z"], DEST_DIR)
    if not ok:
        print(f"  [FEHLER] Konnte Git-Dateiliste nicht lesen: {out}")
        return False
    for relative_path in out.split("\0"):
        name = os.path.basename(relative_path)
        if name.startswith("~$") and name.lower().endswith((".xlsx", ".xlsm", ".xls")):
            ok, result = run_git(["rm", "-f", "--cached", "--", relative_path], DEST_DIR)
            if not ok:
                print(f"  [FEHLER] Konnte Excel-Sperrdatei nicht aus Git entfernen: {result}")
                return False
            print(f"  [Git] Entferne temporäre Datei aus Git (lokal bleibt sie): {relative_path}")
    return True


def git_commit_and_push():
    """
    Fragt in der Konsole nach dem Upload und führt Git aus:
    temporäre Excel-Sperrdateien ignorieren -> git add -> ggf. commit ->
    git pull --no-rebase <remote> <branch> -> git push <remote> <branch>.
    So klappt der Push auch, wenn seit dem letzten Lauf auf GitHub etwas
    geändert wurde oder ein vorheriger Push fehlgeschlagen ist.

    Gibt einen kurzen Status-String zurück, den main() für die
    Abschluss-Zusammenfassung auswertet:
      "no_git"            - git nicht installiert/im PATH
      "not_a_repo"        - DEST_DIR ist kein Git-Repository
      "no_branch"         - aktueller Branch nicht ermittelbar
      "skipped_by_user"   - Nutzer hat die Frage mit Nein beantwortet
      "add_failed"        - 'git add' fehlgeschlagen
      "commit_failed"     - 'git commit' fehlgeschlagen
      "pull_failed"       - Remote-Änderungen konnten nicht integriert werden
      "push_failed"       - 'git push' fehlgeschlagen
      "csv_verify_failed" - besetzung.csv fehlt im gepushten Commit
      "pushed"            - alles erfolgreich (ggf. ohne neuen Commit) gepusht
    """
    print("\n-- Git -------------------------------------------------")

    if shutil.which("git") is None:
        print("  [WARNUNG]  git wurde nicht gefunden - Commit/Push übersprungen.")
        return "no_git"

    ok, _ = run_git(["rev-parse", "--is-inside-work-tree"], DEST_DIR)
    if not ok:
        print(f"  [WARNUNG]  '{DEST_DIR}' ist kein Git-Repository - Commit/Push übersprungen.")
        print("     (Einmalig 'git clone ...' bzw. 'git init' + 'git remote add' nötig.)")
        return "not_a_repo"

    branch = GIT_BRANCH
    if not branch:
        ok, out = run_git(["rev-parse", "--abbrev-ref", "HEAD"], DEST_DIR)
        branch = out.strip() if ok else ""
        if not branch or branch == "HEAD":
            print("  [WARNUNG]  Konnte den aktuellen Git-Branch nicht ermitteln")
            print("     (evtl. 'detached HEAD' oder noch kein Commit) - Commit/Push übersprungen.")
            print("     Alternativ GIT_BRANCH oben im Script fest eintragen.")
            return "no_branch"

    if not ask_yes_no(f"  Änderungen committen und nach {GIT_REMOTE}/{branch} pushen?",
                      default_yes=True):
        print("  Übersprungen - nichts committet/gepusht.")
        return "skipped_by_user"

    # Vor dem Staging temporäre Excel-Dateien aus Git ausschließen.
    # Sie sind technisch nur Sperrdateien von Excel und gehören nicht ins Repo.
    if not ignore_excel_lockfiles_in_git():
        return "add_failed"

    # CSV explizit aufnehmen, auch falls sie durch globale Git-Regeln
    # oder eine lokale .git/info/exclude ignoriert werden sollte.
    if not os.path.isfile(BESETZUNG_CSV):
        print("  [FEHLER] besetzung.csv fehlt; Git-Upload wird abgebrochen.")
        return "add_failed"
    ok, out = run_git(["add", "-f", "--", "besetzung.csv"], DEST_DIR)
    if not ok:
        print(f"  [FEHLER] Konnte besetzung.csv nicht zum Commit vormerken: {out}")
        return "add_failed"

    # Alle übrigen Änderungen erfassen (inkl. gelöschter Dateien).
    ok, out = run_git(["add", "-A", "."], DEST_DIR)
    if not ok:
        print(f"  [FEHLER] 'git add' fehlgeschlagen:\n{out}")
        return "add_failed"

    # Anzeige, ob die neu erzeugte CSV andere Inhalte als der letzte
    # Git-Commit hat. Nur tatsächlich geänderte Inhalte erzeugen eine
    # neue Git-Dateiversion (erneutes Speichern alleine reicht nicht).
    ok, out = run_git(["diff", "--cached", "--name-only", "--", "besetzung.csv"], DEST_DIR)
    if not ok:
        print(f"  [FEHLER] CSV-Aenderungspruefung fehlgeschlagen: {out}")
        return "add_failed"
    if out.strip():
        print("  [Git/CSV] besetzung.csv wurde geaendert und wird mitcommittet.")
    else:
        print("  [Git/CSV] Inhalt von besetzung.csv ist bereits identisch mit dem Git-Stand.")
        print("            Sie wird deshalb nicht erneut als Dateiaenderung angezeigt.")

    ok, out = run_git(["diff", "--cached", "--quiet"], DEST_DIR)
    if ok:
        print("  [Info] Keine neuen Änderungen zu committen.")
    else:
        try:
            message = input("  Commit-Nachricht (leer = Standardtext): ").strip()
        except EOFError:
            print()
            message = ""
        if not message:
            message = "Update dateien.json, audio und noten"

        ok, out = run_git(["commit", "-m", message], DEST_DIR)
        if out:
            print(f"  {out}")
        if not ok:
            print("  [FEHLER] 'git commit' fehlgeschlagen - Push übersprungen.")
            return "commit_failed"

    # Auch wenn es nichts Neues zu committen gibt: Ein vorheriger Push
    # könnte fehlgeschlagen sein, daher IMMER Remote-Abgleich und Push.
    print(f"  -> GitHub-Änderungen abgleichen (git pull --no-rebase {GIT_REMOTE} {branch}) ...")
    ok, out = run_git(["pull", "--no-rebase", GIT_REMOTE, branch], DEST_DIR, timeout=120)
    if out:
        print(f"  {out}")
    if not ok:
        print("  [FEHLER] GitHub-Änderungen konnten nicht automatisch integriert werden.")
        print("     Kein Push durchgeführt. Die lokalen Commits bleiben erhalten.")
        print("     Falls ein Merge-Konflikt angezeigt wird, bitte erst klären;")
        print("     zur Wiederherstellung des Zustands davor: git merge --abort")
        return "pull_failed"

    print(f"  -> Push (git push {GIT_REMOTE} {branch}) ...")
    ok, out = run_git(["push", GIT_REMOTE, branch], DEST_DIR, timeout=120)
    if out:
        print(f"  {out}")
    if ok:
        print("  [OK] Erfolgreich gepusht.")
        # Der Push beinhaltet den aktuellen HEAD. Daher pruefen wir,
        # ob genau diese CSV im gepushten Commit gespeichert ist.
        exists, details = run_git(["cat-file", "-e", "HEAD:besetzung.csv"], DEST_DIR)
        if not exists:
            print("  [FEHLER] Im gepushten Commit fehlt besetzung.csv!", details)
            return "csv_verify_failed"
        identical, details = run_git(["diff", "--quiet", "HEAD", "--", "besetzung.csv"], DEST_DIR)
        if not identical:
            print("  [FEHLER] Lokale besetzung.csv unterscheidet sich vom gepushten Commit!")
            print("     Bitte den Git-Status pruefen:", details)
            return "csv_verify_failed"
        print("  [Git/CSV] Geprueft: besetzung.csv ist im gepushten Commit enthalten")
        print("            und stimmt mit der lokalen CSV ueberein.")
        return "pushed"

    print("  [FEHLER] 'git push' fehlgeschlagen - Änderungen sind lokal")
    print("     committet, aber NICHT auf GitHub angekommen!")
    print("     Bitte Fehlermeldung oben prüfen (z.B. Verbindung oder Rechte).")
    return "push_failed"


def _wait_and_exit():
    try:
        input("\n[Enter] zum Beenden...")
    except EOFError:
        pass


def main():
    print("=" * 55)
    print("  Thadden27 - dateien.json Generator & Datei-Sammler")
    print("=" * 55)
    print(f"Quellordner: {SOURCE_DIR}")
    print(f"Zielordner:  {DEST_DIR}\n")

    if _ONEDRIVE_DIR is None:
        print(f"[FEHLER] Kein OneDrive-Ordner in '{os.path.expanduser('~')}' gefunden.")
        print("  Erwartet wird ein Ordner, dessen Name mit 'OneDrive' beginnt,")
        print("  direkt im Benutzerordner (z.B. 'OneDrive - thaddenschule.de').")
        print("  Mögliche Ursachen: OneDrive ist auf diesem Rechner noch nicht")
        print("  eingerichtet/fertig synchronisiert, oder der Ordner liegt an")
        print("  einer anderen Stelle als üblich.")
        _wait_and_exit()
        return

    if not os.path.isdir(SOURCE_DIR):
        print(f"[FEHLER] Quellordner nicht gefunden: {SOURCE_DIR}")
        print(f"  (gefundener OneDrive-Ordner: {_ONEDRIVE_DIR})")
        print("  Mögliche Ursachen: OneDrive ist auf diesem Rechner noch nicht")
        print("  fertig synchronisiert, oder der Unterordner heißt hier anders")
        print("  als '09 Musik\\2025 Elisabeth'. Dann bitte SOURCE_SUBPATH oben")
        print("  im Script anpassen.")
        _wait_and_exit()
        return

    if not os.path.isdir(DEST_DIR):
        print(f"[FEHLER] Zielordner nicht gefunden: {DEST_DIR}")
        print(f"  (gefundener OneDrive-Ordner: {_ONEDRIVE_DIR})")
        print("  Mögliche Ursachen: der Unterordner heißt hier anders als")
        print("  '12 GitHub\\thadden27' (dann bitte DEST_SUBPATH oben im Script")
        print("  anpassen), oder das Git-Repo wurde dort noch nicht angelegt/")
        print("  geklont.")
        _wait_and_exit()
        return

    if paths_overlap(SOURCE_DIR, DEST_DIR):
        print("[FEHLER] Quellordner und Zielordner überschneiden sich:")
        print(f"  Quellordner: {SOURCE_DIR}")
        print(f"  Zielordner:  {DEST_DIR}")
        print("  Einer der beiden Ordner liegt (ganz oder teilweise) im")
        print("  jeweils anderen - dadurch könnten beim Kopieren aus Versehen")
        print("  Original-Dateien im Quellordner überschrieben/geleert werden.")
        print("  Script wurde deshalb zur Sicherheit abgebrochen, OHNE etwas")
        print("  zu verändern. Bitte SOURCE_SUBPATH und DEST_SUBPATH oben im")
        print("  Script prüfen - die zusammengesetzten Pfade dürfen sich in")
        print("  keinem Fall überschneiden.")
        _wait_and_exit()
        return

    print("-- Besetzung (gültige Abschnitte) ----------------------")
    valid_ids, _ = load_valid_scene_ids()
    if valid_ids is None:
        print("\n[FEHLER] Abbruch: besetzung.xlsx konnte nicht erfolgreich in CSV umgewandelt werden.")
        print("         Es werden keine Dateien kopiert oder zu GitHub gepusht.")
        _wait_and_exit()
        return

    print("\n-- Audio (MP3) -> wird nach audio/ kopiert -------------")
    audio, audio_expected, unassigned_audio = collect_audio(valid_ids)

    print("\n-- Noten (PDF) -> wird nach noten/ kopiert -------------")
    noten, noten_expected, unassigned_noten = collect_noten(valid_ids)

    print(f"\n-- TXT -> wird nach {GITHUB_TXT_PATH}/ kopiert -------------")
    txt, txt_expected, unassigned_txt = collect_txt(valid_ids)

    print(f"\n  Kopiert: {STATS['copied_new']} neu, {STATS['copied_updated']} aktualisiert, "
          f"{STATS['unchanged']} unverändert (übersprungen).")

    # Liedtexte ("..._Lyrics.txt") von den normalen Beschreibungen trennen
    lyrics = {}
    for sid in list(txt.keys()):
        ly = [p for p in txt[sid] if is_lyrics_file(p)]
        if ly:
            lyrics[sid] = ly
            txt[sid] = [p for p in txt[sid] if not is_lyrics_file(p)]
            if not txt[sid]:
                del txt[sid]
            for p in ly:
                print(f"  [Lyrics] {sid}: {os.path.basename(p)}")

    # Zusammenführen
    all_ids = sorted(set(list(audio.keys()) + list(noten.keys()) + list(txt.keys())
                         + list(lyrics.keys())),
                     key=lambda x: [int(n) for n in x.split('.')])

    result = {}
    for sid in all_ids:
        entry = {}
        if sid in noten:
            entry["noten"] = noten[sid]
        if sid in audio:
            entry["audio"] = audio[sid]
        if sid in txt:
            entry["txt"] = txt[sid]
        if sid in lyrics:
            entry["lyrics"] = lyrics[sid]
        result[sid] = entry

    # Generelle Infos zum Stück (falls vorhanden) als "_info" eintragen
    if txt_expected and INFO_FILENAME.lower() in txt_expected:
        result["_info"] = [f"{GITHUB_TXT_PATH}/{INFO_FILENAME}"]

    # JSON schreiben (über Zwischendatei - bei einem Abbruch bleibt die
    # alte dateien.json heil)
    json_str = json.dumps(result, ensure_ascii=False, indent=2)
    write_text_atomic(OUTPUT_FILE, json_str, encoding="utf-8")

    print(f"\n-- Ergebnis ------------------------------------------")
    print(f"[OK] {len(all_ids)} Szenen gefunden")
    print(f"[OK] Dateien kopiert nach: {AUDIO_OUT_DIR}")
    print(f"                    und: {NOTEN_OUT_DIR}")
    print(f"                    und: {TXT_OUT_DIR}")
    print(f"[OK] Datei gespeichert: {OUTPUT_FILE}\n")
    print("Vorschau:")
    print(json_str[:800] + ("..." if len(json_str) > 800 else ""))
    print(f"\n[OK] Fertig! '{GITHUB_AUDIO_PATH}/', '{GITHUB_NOTEN_PATH}/', '{GITHUB_TXT_PATH}/' und")
    print(f"  'dateien.json' liegen jetzt in {DEST_DIR} bereit zum Commit/Push.")

    # -- Nicht zugeordnete Dateien: MP3s/PDFs/TXTs im Quellordner, die
    #    keinem Abschnitt zugeordnet werden konnten (siehe extract_scene_id) --
    report_unassigned(unassigned_audio, unassigned_noten, unassigned_txt)

    # -- Aufräumen: ALLE Dateien im Zielordner, die bei diesem Lauf nicht
    #    aus dem Quellordner kamen (Stimme weggefallen, umbenannt, ...) --
    print(f"\n-- Aufräumen -------------------------------------------")
    checks = [
        ("Audio (MP3)", AUDIO_OUT_DIR, audio_expected),
        ("Noten (PDF)", NOTEN_OUT_DIR, noten_expected),
        ("TXT", TXT_OUT_DIR, txt_expected),
    ]
    orphan_lists = [(label, out_dir, find_orphans(out_dir, expected or set()), expected is None)
                    for label, out_dir, expected in checks]

    if not any(orphans for _, _, orphans, _ in orphan_lists):
        print("  [OK] Keine verwaisten Dateien gefunden - alles im Zielordner")
        print("    hat noch eine Entsprechung im Quellordner.")
    else:
        for label, out_dir, orphans, source_missing in orphan_lists:
            handle_orphans(label, out_dir, orphans, source_missing)

    git_status = "disabled"
    if GIT_AUTO_COMMIT_PUSH:
        git_status = git_commit_and_push()

    # -- Abschluss-Zusammenfassung ---------------------------------------
    problems = []
    if STATS["copy_errors"]:
        problems.append(f"{STATS['copy_errors']} Datei(en) konnten nicht kopiert werden")
    if STATS["delete_errors"]:
        problems.append(f"{STATS['delete_errors']} Datei(en) konnten nicht gelöscht werden")
    if git_status == "no_git":
        problems.append("git ist nicht installiert / nicht im PATH")
    elif git_status == "not_a_repo":
        problems.append(f"'{DEST_DIR}' ist kein Git-Repository")
    elif git_status == "no_branch":
        problems.append("aktueller Git-Branch konnte nicht ermittelt werden (kein Push)")
    elif git_status == "add_failed":
        problems.append("'git add' ist fehlgeschlagen")
    elif git_status == "commit_failed":
        problems.append("'git commit' ist fehlgeschlagen")
    elif git_status == "pull_failed":
        problems.append("'git pull --no-rebase' ist fehlgeschlagen (kein Push durchgeführt)")
    elif git_status == "push_failed":
        problems.append("'git push' ist fehlgeschlagen (lokal aber committet)")
    elif git_status == "csv_verify_failed":
        problems.append("besetzung.csv fehlt im gepushten Commit oder stimmt nicht ueberein")
    # "disabled", "skipped_by_user" und "pushed" sind in Ordnung - kein
    # Problem, sondern erwartetes/gewolltes Verhalten.

    print("\n" + "=" * 55)
    if problems:
        print("  [WARNUNG] FERTIG MIT PROBLEMEN")
        print("=" * 55)
        for p in problems:
            print(f"  - {p}")
        print("\n  Bitte die Meldungen weiter oben in der Konsole prüfen.")
    else:
        print("  [OK] ALLES ERFOLGREICH")
    print("=" * 55)

    unassigned_count = len(unassigned_audio) + len(unassigned_noten) + len(unassigned_txt)
    if unassigned_count:
        print(f"  [Info] Hinweis: {unassigned_count} Datei(en) ohne gültigen")
        print("    Abschnitt (siehe 'Nicht zugeordnete Dateien' weiter oben).")

    # Konsole bleibt offen, bis eine Taste gedrückt wird (auch bei Doppelklick)
    _wait_and_exit()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        # Fängt WIRKLICH ALLES ab, was main() nicht schon selbst behandelt
        # (z.B. ein Tippfehler-Bug, eine unerwartete Datei-/Rechteproblem).
        # Ohne das wuerde die Konsole bei einem unerwarteten Fehler sonst
        # sofort und ohne jede sichtbare Meldung wieder zugehen.
        import traceback
        print("\n" + "=" * 55)
        print("  UNERWARTETER FEHLER - Script wurde abgebrochen")
        print("=" * 55)
        traceback.print_exc()
        print("\n  Bitte diese Meldung an Flo weitergeben, dann kann der")
        print("  Fehler behoben werden.")
        _wait_and_exit()
