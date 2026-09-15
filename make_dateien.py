#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Thadden27 - dateien.json Generator & Datei-Sammler
====================================================
Sammelt MP3s und PDFs aus dem Quellordner (OneDrive), KOPIERT sie in die
Ordner audio/ und noten/ im Ziel-Ordner (lokales Git-Repo) und erstellt
dort automatisch die dateien.json.

VERWENDUNG:
  1. Dieses Script DIREKT in den lokalen Git-Ordner legen (z.B. neben die
     Ordner audio/ und noten/, egal ob der Git-Ordner "C:\\Git\\thadden27"
     oder "C:\\Thadden27\\thadden27" heißt oder sonst irgendwo liegt - das
     Ziel wird automatisch als "der Ordner, in dem dieses Script liegt"
     erkannt). SOURCE_DIR (OneDrive) wird ebenfalls automatisch aus dem
     Windows-Benutzernamen zusammengebaut, siehe ONEDRIVE_SUBPATH weiter
     unten. Auf einem neuen Rechner muss also i.d.R. gar nichts mehr von
     Hand angepasst werden - einfach das Script in den dortigen
     Git-Ordner legen und ausführen.
  2. Script ausführen: Doppelklick  ODER  in der Kommandozeile:
       python3 make_dateien.py
  3. Die Ordner "audio/", "noten/" und die Datei "dateien.json" liegen
     danach im Git-Ordner bereit zum Commit/Push auf GitHub.

QUELLORDNER (OneDrive) - VORHER:
  09 Musik/2025 Elisabeth/
  |---- Partien aus 1.1 Marsch/       <- MP3s, Szene-ID = 1.1
  |---- Partien aus 1.3 Nun will.../  <- MP3s, Szene-ID = 1.3
  |---- Partien aus 14 Finale/        <- MP3s, Szene-ID = 14
  `---- pdf/                          <- alle PDFs
      |---- 1.3 Nun will... - Flöte 1.pdf   <- Szene-ID = 1.3
      `---- 14 Finale - Partitur.pdf        <- Szene-ID = 14

ZIELORDNER (Git-Repo) - NACHHER:
  thadden27/
  |---- audio/                        <- alle MP3s aus allen "Partien aus ..."
  |                                    Unterordnern, hierher kopiert (flach,
  |                                    ohne Unterordner)
  |---- noten/                        <- alle PDFs aus pdf/, hierher kopiert
  `---- dateien.json                  <- verweist auf audio/... und noten/...

Die Original-Dateien im Quellordner (OneDrive) bleiben dabei unangetastet
- es wird kopiert, nicht verschoben/gelöscht. Wenn zwei MP3s aus
unterschiedlichen "Partien aus ..."-Ordnern zufällig denselben Dateinamen
haben, überschreibt die zuletzt kopierte Datei die vorherige in audio/
(mit einer kurzen Meldung in der Konsole).

Existiert eine "besetzung.xlsx" (gleiche Orte wie besetzung.csv, siehe
BESETZUNG_XLSX_CANDIDATES), wird daraus automatisch die gleichnamige
besetzung.csv erzeugt/aktualisiert (nur wenn nötig) - die Besetzung kann
also ganz normal in Excel gepflegt werden. Dafür wird das Python-Paket
"openpyxl" gebraucht; fehlt es, bietet das Script an, es automatisch zu
installieren (fragt vorher in der Konsole nach Bestätigung). Bei Nein
oder wenn die Installation fehlschlägt, macht das Script mit einer
eventuell vorhandenen besetzung.csv normal weiter.

Außerdem prüft das Script anhand von "besetzung.csv" (Liste der
gültigen Abschnitte - wird in DEST_DIR und SOURCE_DIR gesucht), ob es im
Quellordner MP3s oder PDFs gibt, die KEINEM gültigen Abschnitt zugeordnet
werden können (weil kein erkennbares Zahlen-Präfix wie "1.1 ..." im
Ordner-/Dateinamen steht, ODER weil die erkannte Nummer gar nicht in
besetzung.csv vorkommt). Diese Dateien werden TROTZDEM ganz normal nach
audio/ bzw. noten/ kopiert (nur eben ohne Eintrag in dateien.json) und
am Ende als "nicht zugeordnet" aufgelistet - das Script wartet dabei auf
[Enter], um die Liste zu bestätigen. Fehlt besetzung.csv ganz, wird nur
anhand des erkennbaren Zahlen-Präfixes entschieden (wie zuvor).

Danach prüft das Script, ob in audio/ bzw. noten/ Dateien liegen, die im
Quellordner nicht mehr existieren (z.B. weil dort etwas umbenannt oder
gelöscht wurde) - diese "verwaisten" Dateien werden aufgelistet, und das
Script fragt, ob sie gelöscht werden sollen. Gelöscht wird dabei IMMER
nur im Zielordner (Git-Repo), der Quellordner (OneDrive) bleibt davon
unberührt.

Am Ende bietet das Script an, die Änderungen in DEST_DIR gleich zu
committen und zu pushen (fragt dafür in der Konsole nach Bestätigung
und nach der Commit-Nachricht). DEST_DIR muss dafür ein bestehendes
Git-Repo mit eingerichtetem Remote sein (git clone / git remote add
vorher einmalig von Hand erledigen).

ANPASSEN:
  - ONEDRIVE_SUBPATH: der Teil des Quellordner-Pfads NACH "C:\\Users\\<Name>\\"
    (also ohne Benutzername) - wird zusammen mit dem automatisch erkannten
    Windows-Benutzernamen zu SOURCE_DIR zusammengesetzt. Auf jedem Rechner
    passt sich SOURCE_DIR damit automatisch an, egal wer angemeldet ist.
  - DEST_DIR:   wird automatisch als der Ordner erkannt, in dem dieses
    Script liegt (siehe VERWENDUNG oben). Nur nötig anzupassen, falls das
    Script ausnahmsweise NICHT im Git-Ordner liegen soll - dann hier
    einen festen Pfad eintragen.
  - PDF_FOLDER: Name des PDF-Unterordners in SOURCE_DIR
  - PARTIEN_PREFIX: Anfang der Audio-Unterordnernamen in SOURCE_DIR
  - GITHUB_AUDIO_PATH / GITHUB_NOTEN_PATH: Namen der Zielordner in
    DEST_DIR (gleichzeitig die Pfade, wie sie in dateien.json stehen)
  - GIT_AUTO_COMMIT_PUSH: auf False setzen, um den Commit/Push-Teil
    komplett zu deaktivieren (dann wird nur kopiert + dateien.json
    geschrieben, wie bisher)
  - BESETZUNG_CSV_CANDIDATES: Pfade, an denen nach "besetzung.csv"
    gesucht wird (der erste gefundene Treffer gewinnt). Eigenen Pfad bei
    Bedarf vorne in die Liste einfügen.
  - BESETZUNG_XLSX_CANDIDATES: dasselbe für "besetzung.xlsx" - wird
    gefunden, automatisch in die gleichnamige .csv umgewandelt (braucht
    "openpyxl" - wird bei Bedarf mit Rückfrage automatisch installiert)
"""

import os
import io
import re
import csv
import sys
import json
import shutil
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
# Quellordner: hier liegen die "Partien aus ..."-Ordner und "pdf/".
# Nur der Teil NACH "C:\Users\<Benutzername>\" wird hier eingetragen -
# der Benutzername selbst wird unten automatisch anhand des aktuell an-
# gemeldeten Windows-Kontos ermittelt (os.path.expanduser("~")). Dadurch
# läuft das Script unverändert auf jedem Rechner (z.B. "lepold" oder
# "flori"), solange die OneDrive-Ordnerstruktur darunter gleich heißt.
ONEDRIVE_SUBPATH = ["OneDrive - thaddenschule.de", "09 Musik", "2025 Elisabeth"]
SOURCE_DIR = os.path.join(os.path.expanduser("~"), *ONEDRIVE_SUBPATH)

# Zielordner: lokales Git-Repo, in dem audio/, noten/ und dateien.json
# landen sollen. Standardmäßig automatisch der Ordner, in dem DIESES
# Script liegt - das Script also einfach direkt in den Git-Ordner legen
# (egal ob der "C:\\Git\\thadden27", "C:\\Thadden27\\thadden27" oder anders
# heißt/liegt). Nur bei Bedarf hier stattdessen einen festen Pfad
# eintragen, z.B. DEST_DIR = r"C:\Git\thadden27".
DEST_DIR = os.path.dirname(os.path.abspath(__file__))

# Lokale Ordnernamen (innerhalb SOURCE_DIR)
PDF_FOLDER     = "pdf"           # Ordner mit allen PDF-Dateien
PARTIEN_PREFIX = "Partien aus "  # Unterordner-Präfix für MP3-Ordner

# Namen der Zielordner in DEST_DIR - gleichzeitig die Pfade, wie sie in
# der dateien.json stehen sollen (z.B. "audio/1.1 Marsch.mp3")
GITHUB_AUDIO_PATH = "audio"
GITHUB_NOTEN_PATH = "noten"

# Vollständige lokale Zielpfade, in die tatsächlich kopiert wird
AUDIO_OUT_DIR = os.path.join(DEST_DIR, GITHUB_AUDIO_PATH)
NOTEN_OUT_DIR = os.path.join(DEST_DIR, GITHUB_NOTEN_PATH)

# Ausgabedatei (liegt im Ziel-Ordner, direkt neben audio/ und noten/)
OUTPUT_FILE = os.path.join(DEST_DIR, "dateien.json")

# Nach dem Kopieren in der Konsole fragen, ob committet + gepusht werden
# soll (inkl. Abfrage der Commit-Nachricht). False = nur kopieren, kein Git.
GIT_AUTO_COMMIT_PUSH = True

# Kandidaten-Pfade für "besetzung.csv" (Liste der gültigen Abschnitte).
# Das Script probiert diese der Reihe nach durch und nimmt die erste
# Datei, die existiert. Wird keine gefunden, gilt jede per Zahlen-Präfix
# erkannte ID automatisch als gültig (altes Verhalten ohne CSV-Abgleich).
BESETZUNG_CSV_CANDIDATES = [
    os.path.join(DEST_DIR, "besetzung.csv"),
    os.path.join(SOURCE_DIR, "besetzung.csv"),
]

# Kandidaten-Pfade für "besetzung.xlsx". Existiert eine davon, wird daraus
# automatisch die gleichnamige .csv im selben Ordner erzeugt/aktualisiert
# (nur wenn die .csv fehlt oder älter ist als die .xlsx) - die Besetzung
# kann also ganz normal in Excel gepflegt werden, ohne von Hand als CSV
# exportieren zu müssen. Braucht das Paket "openpyxl" (pip install
# openpyxl); fehlt es, wird das klar gemeldet und stattdessen direkt mit
# einer vorhandenen besetzung.csv weitergemacht (falls es eine gibt).
BESETZUNG_XLSX_CANDIDATES = [
    os.path.join(DEST_DIR, "besetzung.xlsx"),
    os.path.join(SOURCE_DIR, "besetzung.xlsx"),
]
# ----------------------------------------------------------------------

# Zählt Probleme während des Laufs (Kopier-/Löschfehler), damit am Ende
# eine ehrliche "Alles erfolgreich"- oder Fehler-Zusammenfassung möglich ist.
STATS = {"copy_errors": 0, "delete_errors": 0}

# Spaltennamen (klein geschrieben), die als "Abschnitt-Spalte" in
# besetzung.csv erkannt werden.
ABSCHNITT_SPALTENNAMEN = (
    "abschnitt", "abschnitt-id", "abschnittid", "abschnittsnummer",
    "szene", "szene-id", "szenenid", "id", "nr", "nummer",
)


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
    try:
        answer = input("  Jetzt automatisch installieren (pip install openpyxl)? [J/n]: ").strip().lower()
    except EOFError:
        answer = "j"

    if answer not in ("", "j", "ja", "y", "yes"):
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
        print("     umgewandelt werden (oder die Datei einmalig von Hand in Excel")
        print("     als CSV speichern).")
        return False

    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
        try:
            ws = wb.active
            rows = [[_cell_to_str(v) for v in row] for row in ws.iter_rows(values_only=True)]
        finally:
            wb.close()
    except Exception as e:
        print(f"  [FEHLER] Konnte '{xlsx_path}' nicht lesen: {e}")
        return False

    try:
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f, delimiter=";").writerows(rows)
    except OSError as e:
        print(f"  [FEHLER] Konnte '{csv_path}' nicht schreiben: {e}")
        return False

    print(f"  [OK] '{os.path.basename(xlsx_path)}' -> '{os.path.basename(csv_path)}'")
    print(f"     umgewandelt ({len(rows)} Zeile(n)).")
    return True


def sync_besetzung_csv_from_xlsx():
    """
    Prüft BESETZUNG_XLSX_CANDIDATES der Reihe nach; existiert eine
    besetzung.xlsx, wird daraus die gleichnamige besetzung.csv im selben
    Ordner erzeugt/aktualisiert - aber nur, wenn diese fehlt oder älter
    als die xlsx ist (spart unnötige Konvertierungen bei jedem Lauf).
    """
    for xlsx_path in BESETZUNG_XLSX_CANDIDATES:
        if not os.path.isfile(xlsx_path):
            continue
        csv_path = os.path.splitext(xlsx_path)[0] + ".csv"
        stale = not os.path.isfile(csv_path) or \
            os.path.getmtime(xlsx_path) > os.path.getmtime(csv_path)
        if stale:
            print(f"  [Info]  '{os.path.basename(xlsx_path)}' gefunden - erzeuge/aktualisiere CSV ...")
            convert_xlsx_to_csv(xlsx_path, csv_path)
        break  # nur die erste gefundene xlsx verwenden (gleiche Prioritaet wie bei den CSV-Kandidaten)


def load_valid_scene_ids():
    """
    Lädt die Liste der gültigen Abschnitts-IDs aus "besetzung.csv"
    (erster Treffer aus BESETZUNG_CSV_CANDIDATES). Erkennt automatisch
    das Trennzeichen (";" oder ",") und die Spalte mit der Abschnitts-ID
    (per Spaltenname, sonst wird die erste Spalte genommen).

    Existiert eine besetzung.xlsx (siehe BESETZUNG_XLSX_CANDIDATES), wird
    die passende CSV zuerst automatisch daraus erzeugt/aktualisiert.

    Gibt (valid_ids, csv_path) zurück:
      valid_ids = set() der gültigen IDs, oder ein LEERES set(), wenn
                  keine besetzung.csv gefunden/gelesen werden konnte -
                  dann wird NICHT gegen die Besetzung geprüft (jede per
                  Zahlen-Präfix erkannte ID gilt automatisch als gültig).
      csv_path  = Pfad der benutzten CSV, oder None.
    """
    sync_besetzung_csv_from_xlsx()

    csv_path = next((c for c in BESETZUNG_CSV_CANDIDATES if os.path.isfile(c)), None)

    if not csv_path:
        searched = ", ".join(BESETZUNG_CSV_CANDIDATES)
        print(f"  [WARNUNG]  Keine 'besetzung.csv' gefunden (gesucht: {searched}).")
        print("     Abschnitts-Zuordnung wird NICHT gegen eine Besetzung geprüft -")
        print("     es zählt nur, ob eine Nummer im Namen erkannt wurde.")
        return set(), None

    try:
        with open(csv_path, 'rb') as f:
            raw = f.read()
    except OSError as e:
        print(f"  [FEHLER] Konnte '{csv_path}' nicht lesen: {e}")
        print("     Abschnitts-Zuordnung wird NICHT gegen eine Besetzung geprüft.")
        return set(), None

    # besetzung.csv kommt oft aus Excel/Windows und ist dann NICHT UTF-8,
    # sondern Windows-1252 ("ANSI") - einfach mehrere Kodierungen der
    # Reihe nach probieren, statt beim ersten Umlaut abzustürzen.
    text = None
    used_encoding = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            used_encoding = enc
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode("utf-8", errors="replace")
        used_encoding = "utf-8 (mit Ersatzzeichen fuer unbekannte Bytes)"
    if used_encoding != "utf-8-sig":
        print(f"  [Info]  '{os.path.basename(csv_path)}' ist nicht UTF-8 - wurde")
        print(f"     als '{used_encoding}' gelesen.")

    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=";,\t")
    except csv.Error:
        dialect = csv.excel
        dialect.delimiter = ';' if sample.count(';') >= sample.count(',') else ','
    rows = list(csv.reader(io.StringIO(text), dialect))

    if not rows or not rows[0]:
        print(f"  [WARNUNG]  '{csv_path}' ist leer.")
        return set(), csv_path

    header = [h.strip().lower() for h in rows[0]]
    id_col = next((i for i, name in enumerate(header) if name in ABSCHNITT_SPALTENNAMEN), None)
    if id_col is None:
        id_col = next((i for i, name in enumerate(header)
                        if "abschnitt" in name or "szene" in name), None)

    if id_col is None:
        id_col = 0
        print(f"  [Info]  Keine Spalte 'Abschnitt'/'Szene' in '{os.path.basename(csv_path)}'")
        print(f"     erkannt - benutze stattdessen die erste Spalte ('{rows[0][0]}').")
    else:
        print(f"  [Info]  Abschnitts-Spalte in '{os.path.basename(csv_path)}': '{rows[0][id_col]}'")

    valid_ids = set()
    for row in rows[1:]:
        if id_col >= len(row):
            continue
        m = re.match(r'^(\d+(?:\.\d+)?)', row[id_col].strip())
        if m:
            valid_ids.add(m.group(1))

    print(f"  [OK] {len(valid_ids)} gültige Abschnitts-ID(s) aus '{os.path.basename(csv_path)}' geladen.")
    return valid_ids, csv_path


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


def copy_file(src_path, dest_dir, filename):
    """
    Kopiert eine Datei nach dest_dir (legt den Ordner bei Bedarf an).
    Überschreibt stillschweigend, falls dort schon eine gleichnamige
    Datei liegt (z.B. weil zwei "Partien aus ..."-Ordner eine Datei mit
    demselben Namen enthalten) - gibt dabei aber eine kurze Meldung aus.
    Bricht bei einem Kopierfehler NICHT das ganze Script ab, sondern
    meldet den Fehler und zählt ihn für die Abschluss-Zusammenfassung.
    """
    try:
        os.makedirs(dest_dir, exist_ok=True)
        dest_path = os.path.join(dest_dir, filename)

        if os.path.abspath(src_path) == os.path.abspath(dest_path):
            return

        if os.path.exists(dest_path):
            print(f"     * überschreibe vorhandene Datei im Zielordner: {filename}")

        shutil.copy2(src_path, dest_path)
    except OSError as e:
        STATS["copy_errors"] += 1
        print(f"     [FEHLER] Fehler beim Kopieren von '{filename}': {e}")


def collect_audio(valid_ids):
    """
    Sucht MP3s in allen "Partien aus X.X NAME"-Unterordnern von
    SOURCE_DIR und kopiert ALLE (flach, ohne Unterordner) nach
    AUDIO_OUT_DIR - auch die, die keinem gültigen Abschnitt zugeordnet
    werden können. Ein Ordner gilt als gültig zugeordnet, wenn sein Name
    eine Szenen-ID enthält UND (valid_ids leer ist ODER diese ID in
    valid_ids vorkommt); nur dann landen seine MP3s auch im result-Dict
    (also in dateien.json).

    Gibt zurück:
      (result, source_filenames, unassigned)
      result           = {szene_id: ["audio/dateiname.mp3", ...], ...}
      source_filenames = set aller MP3-Dateinamen, die aktuell im
                         Quellordner existieren - wird benutzt, um später
                         verwaiste Dateien im Zielordner zu finden.
      unassigned       = Liste von (ordnername, dateiname) für MP3s ohne
                         gültigen Abschnitt (werden trotzdem kopiert,
                         aber nicht in dateien.json aufgenommen).
    """
    result = {}
    source_filenames = set()
    unassigned = []

    for entry in os.scandir(SOURCE_DIR):
        if not entry.is_dir():
            continue
        if not entry.name.startswith(PARTIEN_PREFIX):
            continue

        # Szenen-ID aus Ordnername extrahieren und gegen besetzung.csv prüfen
        sid = extract_scene_id(entry.name)
        sid_ok = sid is not None and (not valid_ids or sid in valid_ids)
        if sid is None:
            print(f"  [WARNUNG]  Ordner ohne erkennbare ID (wird trotzdem kopiert): {entry.name}")
        elif not sid_ok:
            print(f"  [WARNUNG]  Abschnitt '{sid}' nicht in besetzung.csv - wird trotzdem")
            print(f"     kopiert, aber nicht in dateien.json aufgenommen: {entry.name}")

        # MP3s in diesem Ordner finden, IMMER kopieren, und verzeichnen
        mp3s = []
        for f in sorted(os.scandir(entry.path), key=lambda x: x.name.lower()):
            if f.is_file() and f.name.lower().endswith('.mp3'):
                source_filenames.add(f.name)
                copy_file(f.path, AUDIO_OUT_DIR, f.name)
                if sid_ok:
                    github_path = f"{GITHUB_AUDIO_PATH}/{f.name}"
                    mp3s.append(github_path)
                    print(f"  [Audio] {sid}: {f.name}")
                else:
                    unassigned.append((entry.name, f.name))
                    print(f"  [Audio] (nicht zugeordnet): {f.name}")

        if mp3s:
            if sid not in result:
                result[sid] = []
            result[sid].extend(mp3s)

    return result, source_filenames, unassigned


def collect_noten(valid_ids):
    """
    Sucht PDFs im "pdf"-Unterordner von SOURCE_DIR und kopiert ALLE nach
    NOTEN_OUT_DIR - auch die, die keinem gültigen Abschnitt zugeordnet
    werden können (siehe collect_audio für die Logik von valid_ids).

    Gibt zurück:
      (result, source_filenames, unassigned)
      result           = {szene_id: ["noten/dateiname.pdf", ...], ...}
      source_filenames = set aller PDF-Dateinamen im "pdf"-Ordner, oder
                         None, falls der "pdf"-Ordner gar nicht existiert
                         (dann lässt sich nicht verlässlich sagen, was
                         "verwaist" ist - siehe find_orphans()).
      unassigned       = Liste von PDF-Dateinamen ohne gültigen Abschnitt
                         (werden trotzdem kopiert, aber nicht in
                         dateien.json aufgenommen).
    """
    result = {}
    pdf_dir = os.path.join(SOURCE_DIR, PDF_FOLDER)

    if not os.path.isdir(pdf_dir):
        print(f"  [WARNUNG]  Kein '{PDF_FOLDER}'-Ordner gefunden - keine Noten.")
        return result, None, []

    source_filenames = set()
    unassigned = []

    for f in sorted(os.scandir(pdf_dir), key=lambda x: x.name.lower()):
        if not f.is_file():
            continue
        if not f.name.lower().endswith('.pdf'):
            continue

        source_filenames.add(f.name)
        copy_file(f.path, NOTEN_OUT_DIR, f.name)

        sid = extract_scene_id(f.name)
        sid_ok = sid is not None and (not valid_ids or sid in valid_ids)

        if not sid_ok:
            if sid is None:
                print(f"  [WARNUNG]  PDF ohne erkennbare ID (wird trotzdem kopiert): {f.name}")
            else:
                print(f"  [WARNUNG]  Abschnitt '{sid}' nicht in besetzung.csv - '{f.name}' wird")
                print("     trotzdem kopiert, aber nicht in dateien.json aufgenommen.")
            unassigned.append(f.name)
            continue

        github_path = f"{GITHUB_NOTEN_PATH}/{f.name}"
        if sid not in result:
            result[sid] = []
        result[sid].append(github_path)
        print(f"  [Noten] {sid}: {f.name}")

    return result, source_filenames, unassigned


def report_unassigned(unassigned_audio, unassigned_noten):
    """
    Zeigt MP3s/PDFs im Quellordner an, die keinem Abschnitt (keiner
    erkennbaren Szenen-ID) zugeordnet werden konnten und deshalb NICHT
    mit kopiert/in dateien.json aufgenommen wurden. Wartet danach auf
    [Enter], um die Liste zu bestätigen (reine Anzeige, es wird nichts
    automatisch verändert).
    """
    print(f"\n-- Nicht zugeordnete Dateien ---------------------------")

    if not unassigned_audio and not unassigned_noten:
        print("  [OK] Alle MP3s und PDFs im Quellordner konnten einem")
        print("    Abschnitt zugeordnet werden.")
        return

    if unassigned_audio:
        print(f"  [Audio] {len(unassigned_audio)} MP3-Datei(en) ohne erkennbaren Abschnitt:")
        for ordner, name in unassigned_audio:
            print(f"    - {name}   (Ordner: {ordner})")

    if unassigned_noten:
        print(f"  [Noten] {len(unassigned_noten)} PDF-Datei(en) ohne erkennbaren Abschnitt:")
        for name in unassigned_noten:
            print(f"    - {name}")

    print("\n  Diese Dateien wurden trotzdem ganz normal nach audio/ bzw.")
    print("  noten/ kopiert - sie tauchen nur NICHT in dateien.json auf.")
    print("  Um sie zuzuordnen: Datei-/Ordnernamen im Quellordner mit der")
    print(f"  passenden Abschnitts-Nummer beginnen lassen (z.B. \"1.1 ...\",")
    print(f"  bei Audio-Ordnern mit Präfix \"{PARTIEN_PREFIX}\") UND sicherstellen,")
    print("  dass diese Nummer auch in besetzung.csv vorkommt.")

    try:
        input("\n  [Enter] um diese Liste zu bestätigen und fortzufahren...")
    except EOFError:
        pass


def find_orphans(out_dir, source_filenames, extension):
    """
    Findet Dateien mit 'extension' in out_dir, deren Name NICHT (mehr) in
    source_filenames vorkommt - also im Quellordner nicht mehr existiert.
    Gibt [] zurück, wenn source_filenames None ist (siehe collect_noten)
    oder out_dir noch gar nicht existiert - dann lässt sich nichts
    Verlässliches über "verwaist" aussagen, also lieber nichts vorschlagen.
    """
    if source_filenames is None or not os.path.isdir(out_dir):
        return []
    orphans = []
    for f in sorted(os.scandir(out_dir), key=lambda x: x.name.lower()):
        if f.is_file() and f.name.lower().endswith(extension) \
                and f.name not in source_filenames:
            orphans.append(f.name)
    return orphans


def handle_orphans(label, out_dir, orphans):
    """
    Zeigt verwaiste Dateien (im Zielordner, aber nicht mehr im Quellordner
    vorhanden) an und fragt in der Konsole, ob sie gelöscht werden sollen.
    Löscht NUR im Zielordner (DEST_DIR) - der Quellordner wird nie
    angefasst.
    """
    if not orphans:
        return

    print(f"\n  {label}: {len(orphans)} Datei(en) in '{out_dir}',")
    print("  die im Quellordner nicht mehr gefunden wurden:")
    for name in orphans:
        print(f"    - {name}")

    try:
        answer = input(f"  Diese {len(orphans)} Datei(en) jetzt löschen? [j/N]: ").strip().lower()
    except EOFError:
        answer = "n"

    if answer not in ("j", "ja", "y", "yes"):
        print("  Übersprungen - nichts gelöscht.")
        return

    deleted = 0
    for name in orphans:
        path = os.path.join(out_dir, name)
        try:
            os.remove(path)
            deleted += 1
            print(f"    [geloescht]  gelöscht: {name}")
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


def git_commit_and_push():
    """
    Fragt in der Konsole nach, ob die kopierten Dateien in DEST_DIR
    committet und gepusht werden sollen, und macht das dann
    (git add -> git commit -m "..." -> git push).

    Gibt einen kurzen Status-String zurück, den main() für die
    Abschluss-Zusammenfassung auswertet:
      "no_git"          - git nicht installiert/im PATH
      "not_a_repo"      - DEST_DIR ist kein Git-Repository
      "skipped_by_user" - Nutzer hat die Frage mit Nein beantwortet
      "add_failed"      - 'git add' fehlgeschlagen
      "no_changes"      - nichts zu committen
      "commit_failed"   - 'git commit' fehlgeschlagen
      "push_failed"     - Commit ok, aber 'git push' fehlgeschlagen
      "pushed"          - alles erfolgreich committet UND gepusht
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

    try:
        answer = input("  Änderungen committen und pushen? [J/n]: ").strip().lower()
    except EOFError:
        answer = "j"
    if answer not in ("", "j", "ja", "y", "yes"):
        print("  Übersprungen - nichts committet/gepusht.")
        return "skipped_by_user"

    # "-A ." statt einzelner Pfade: dadurch werden WIRKLICH ALLE Änderungen
    # im Git-Ordner erfasst (audio/, noten/, dateien.json, besetzung.csv,
    # besetzung.xlsx, neue/gelöschte Dateien, etc.) - nicht nur eine feste
    # Auswahl von Pfaden, bei der leicht mal eine Datei vergessen wird.
    ok, out = run_git(["add", "-A", "."], DEST_DIR)
    if not ok:
        print(f"  [FEHLER] 'git add' fehlgeschlagen:\n{out}")
        return "add_failed"

    ok, out = run_git(["status", "--porcelain"], DEST_DIR)
    if ok and not out.strip():
        print("  [Info] Keine Änderungen gegenüber dem letzten Commit - nichts zu tun.")
        return "no_changes"

    try:
        message = input("  Commit-Nachricht (leer = Standardtext): ").strip()
    except EOFError:
        message = ""
    if not message:
        message = "Update dateien.json, audio und noten"

    ok, out = run_git(["commit", "-m", message], DEST_DIR)
    if out:
        print(f"  {out}")
    if not ok:
        print("  [FEHLER] 'git commit' fehlgeschlagen - Push übersprungen.")
        return "commit_failed"

    print("  -> Push...")
    ok, out = run_git(["push"], DEST_DIR, timeout=30)
    if out:
        print(f"  {out}")
    if ok:
        print("  [OK] Erfolgreich gepusht.")
        return "pushed"
    else:
        print("  [FEHLER] 'git push' fehlgeschlagen - die Änderungen sind lokal")
        print("     committet, aber NICHT auf GitHub angekommen!")
        print("     Fehlermeldung oben prüfen (z.B. kein Internet, kein")
        print("     Remote/Upstream eingerichtet, Merge-Konflikt, fehlende")
        print("     Berechtigung). Danach reicht später ein einfaches")
        print("     'git push' im Git-Ordner, ohne das Script nochmal")
        print("     laufen zu lassen.")
        return "push_failed"


def main():
    print("=" * 55)
    print("  Thadden27 - dateien.json Generator & Datei-Sammler")
    print("=" * 55)
    print(f"Quellordner: {SOURCE_DIR}")
    print(f"Zielordner:  {DEST_DIR}\n")

    if not os.path.isdir(SOURCE_DIR):
        print(f"[FEHLER] Quellordner nicht gefunden: {SOURCE_DIR}")
        print(f"  (automatisch erkannter Benutzername: {os.path.expanduser('~')})")
        print("  Mögliche Ursachen: OneDrive ist auf diesem Rechner noch nicht")
        print("  fertig synchronisiert, oder der OneDrive-Ordner heißt hier anders")
        print("  als 'OneDrive - thaddenschule.de\\09 Musik\\2025 Elisabeth'.")
        print("  Dann bitte ONEDRIVE_SUBPATH oben im Script anpassen.")
        try:
            input("\n[Enter] zum Beenden...")
        except EOFError:
            pass
        return

    if not os.path.isdir(DEST_DIR):
        # Kann eigentlich nicht passieren, da DEST_DIR = Ordner des Scripts
        # selbst ist - nur relevant, falls DEST_DIR oben von Hand auf
        # einen festen Pfad umgestellt wurde.
        print(f"[FEHLER] Zielordner nicht gefunden: {DEST_DIR}")
        print("  Bitte DEST_DIR oben im Script prüfen (Git-Repo lokal")
        print("  geklont/vorhanden?).")
        try:
            input("\n[Enter] zum Beenden...")
        except EOFError:
            pass
        return

    print("-- Besetzung (gültige Abschnitte) ----------------------")
    valid_ids, besetzung_csv_path = load_valid_scene_ids()

    print("\n-- Audio (MP3) -> wird nach audio/ kopiert -------------")
    audio, audio_source_files, unassigned_audio = collect_audio(valid_ids)

    print("\n-- Noten (PDF) -> wird nach noten/ kopiert -------------")
    noten, noten_source_files, unassigned_noten = collect_noten(valid_ids)

    # Zusammenführen
    all_ids = sorted(set(list(audio.keys()) + list(noten.keys())),
                     key=lambda x: [int(n) for n in x.split('.')])

    result = {}
    for sid in all_ids:
        entry = {}
        if sid in noten:
            entry["noten"] = noten[sid]
        if sid in audio:
            entry["audio"] = audio[sid]
        result[sid] = entry

    # JSON schreiben
    json_str = json.dumps(result, ensure_ascii=False, indent=2)
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(json_str)

    print(f"\n-- Ergebnis ------------------------------------------")
    print(f"[OK] {len(all_ids)} Szenen gefunden")
    print(f"[OK] Dateien kopiert nach: {AUDIO_OUT_DIR}")
    print(f"                    und: {NOTEN_OUT_DIR}")
    print(f"[OK] Datei gespeichert: {OUTPUT_FILE}\n")
    print("Vorschau:")
    print(json_str[:800] + ("..." if len(json_str) > 800 else ""))
    print(f"\n[OK] Fertig! '{GITHUB_AUDIO_PATH}/', '{GITHUB_NOTEN_PATH}/' und")
    print(f"  'dateien.json' liegen jetzt in {DEST_DIR} bereit zum Commit/Push.")

    # -- Nicht zugeordnete Dateien: MP3s/PDFs im Quellordner, die keinem
    #    Abschnitt zugeordnet werden konnten (siehe extract_scene_id) --
    report_unassigned(unassigned_audio, unassigned_noten)

    # -- Aufräumen: Dateien im Zielordner, die im Quellordner nicht mehr
    #    existieren (z.B. weil im Quellordner umbenannt/gelöscht wurde) --
    print(f"\n-- Aufräumen -------------------------------------------")
    audio_orphans = find_orphans(AUDIO_OUT_DIR, audio_source_files, '.mp3')
    noten_orphans = find_orphans(NOTEN_OUT_DIR, noten_source_files, '.pdf')

    if not audio_orphans and not noten_orphans:
        print("  [OK] Keine verwaisten Dateien gefunden - alles im Zielordner")
        print("    hat noch eine Entsprechung im Quellordner.")
    else:
        handle_orphans("Audio (MP3)", AUDIO_OUT_DIR, audio_orphans)
        handle_orphans("Noten (PDF)", NOTEN_OUT_DIR, noten_orphans)

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
    elif git_status == "add_failed":
        problems.append("'git add' ist fehlgeschlagen")
    elif git_status == "commit_failed":
        problems.append("'git commit' ist fehlgeschlagen")
    elif git_status == "push_failed":
        problems.append("'git push' ist fehlgeschlagen (lokal aber committet)")
    # "disabled", "skipped_by_user", "no_changes" und "pushed" sind alle
    # in Ordnung - kein Problem, sondern erwartetes/gewolltes Verhalten.

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

    unassigned_count = len(unassigned_audio) + len(unassigned_noten)
    if unassigned_count:
        print(f"  [Info] Hinweis: {unassigned_count} Datei(en) ohne erkennbaren")
        print("    Abschnitt (siehe 'Nicht zugeordnete Dateien' weiter oben).")

    # Konsole bleibt offen, bis eine Taste gedrückt wird (auch bei Doppelklick)
    try:
        input("\n[Enter] zum Beenden...")
    except EOFError:
        pass


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
        try:
            input("\n[Enter] zum Beenden...")
        except EOFError:
            pass
