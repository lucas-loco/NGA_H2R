#!/usr/bin/env python3
"""
kobo_anonymize.py – Anonymisierung von KoboToolbox-Exporten
============================================================

Liest einen Kobo-Export (XLSX mit allen Sheets inkl. Repeat-Groups, oder CSV),
entfernt bzw. vergröbert personenbezogene Daten und schreibt:

  <name>_anonymized.xlsx / .csv      -> anonymisierte Daten (zum Teilen/Analysieren)
  <name>_anonymization_report.csv    -> was mit jeder Spalte passiert ist (zur Kontrolle)
  <name>_small_groups.csv            -> Kombinationen von Quasi-Identifikatoren mit n < k

Was passiert standardmäßig?
  * DROP       Direkte Identifikatoren: Namen, Telefon, E-Mail, Adresse, Geräte-IDs
               (deviceid, imei, simserial, ...), Kobo-Metadaten (_submitted_by,
               _attachments, _notes, ...), Fotos/Audio/Unterschriften.
  * PSEUDO     IDs (_id, _uuid, Enumerator-ID, Haushalts-ID ...) -> gesalzener Hash
               (HMAC-SHA256). Gleicher Wert -> gleiches Pseudonym, Verknüpfungen
               zwischen Haupt- und Repeat-Sheets bleiben also erhalten.
  * GPS        Koordinaten werden entfernt (oder auf N Nachkommastellen gerundet).
  * DATUM      start/end/today/_submission_time -> nur Monat (YYYY-MM).
               Vorher wird die Interviewdauer in Minuten berechnet.
  * ALTER      Numerisches Alter -> Altersgruppen (z. B. "18-29", "60+").
  * FREITEXT   Telefonnummern, E-Mails und lange Ziffernfolgen werden geschwärzt.
  * K-CHECK    Prüft, wie viele Befragte in seltenen Kombinationen von
               Quasi-Identifikatoren (Geschlecht, Altersgruppe, Ort ...) liegen.

Empfohlener Ablauf
------------------
  1. Export in Kobo: "XLS", Werte und Kopfzeilen = "XML values and headers",
     "Include groups in headers" egal (beides wird erkannt).
  2. Probelauf, nur Plan anzeigen:
        python kobo_anonymize.py daten.xlsx --form formular.xlsx --dry-run
  3. Report prüfen, ggf. Spalten mit --keep / --drop / --pseudo anpassen.
  4. Echter Lauf:
        python kobo_anonymize.py daten.xlsx --form formular.xlsx
  5. Die Datei *_anonymized.* weitergeben. Die Salt-Datei (.kobo_salt) und
     ggf. die Mapping-Datei NIEMALS weitergeben!

--form (das XLSForm aus Kobo: Projekt > Formular > "Download XLS") ist optional,
macht die Erkennung aber deutlich zuverlässiger (Freitext-, GPS-, Foto-, Datumsfelder).

Abhängigkeiten: pandas, openpyxl   (pip install -r requirements.txt)
"""
from __future__ import annotations

import argparse
import hashlib
import hmac
import re
import secrets
import sys
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

# =============================================================================
# KONFIGURATION – hier bei Bedarf anpassen
# =============================================================================

# Regex-Regeln werden auf den kleingeschriebenen Spaltennamen angewendet
# (und zusätzlich auf den letzten Teil nach "/" bei Gruppen-Headern).
# Reihenfolge der Prüfung: manuelle Overrides > XLSForm-Typ > diese Regeln.

# Kobo-Strukturspalten, die nie verändert werden (Verknüpfung Haupt-/Repeat-Sheets)
KEEP_EXACT = {"_index", "_parent_index", "_parent_table_name", "_submission__submission_time",
              "_submission__version__", "__version__", "_version_"}

DROP_EXACT = {
    # Kobo/ODK-Metadaten
    "deviceid", "subscriberid", "imei", "simserial", "simid", "phonenumber",
    "username", "email", "audit", "instanceid", "instancename", "deprecatedid",
    "_submitted_by", "_attachments", "_notes", "_tags", "_status",
    "_validation_status", "_geolocation", "_xform_id_string", "formhub/uuid",
    "meta/instanceid", "meta/instancename", "meta/deprecatedid", "meta/audit",
    "meta/rootuuid", "rootuuid",
}

# Tokens (durch "_" / "/" / "-" / "." getrennt), die auf direkte Identifikatoren hindeuten
DROP_TOKEN_PATTERN = re.compile(
    r"(^|[_/.\-])("
    r"name|names|nom|noms|surname|firstname|first_name|lastname|last_name|fullname|full_name|"
    r"phone|phones|telephone|tel|mobile|cell|whatsapp|phone_number|contact_number|"
    r"e_?mail|address|adresse|street|house_number|"
    r"signature|photo|photos|picture|image|img|selfie|audio|video|recording|"
    r"nin|passport|id_card|idcard|national_id|bvn|ssn"
    r")($|[_/.\-])"
)

GPS_PATTERN = re.compile(
    r"(geopoint|geotrace|geoshape|gps|_?latitude|_?longitude|_?altitude|_?precision|"
    r"coordinates|coordinate|(^|_)lat($|_)|(^|_)lon($|_)|(^|_)lng($|_))"
)

PSEUDO_EXACT = {"_id", "_uuid", "_submission__id", "_submission__uuid", "uuid"}
PSEUDO_TOKEN_PATTERN = re.compile(
    r"(^|[_/])("
    r"enumerator|enumerator_id|enum_id|enum|interviewer|surveyor|enumerator_name|"
    r"hh_id|household_id|respondent_id|ki_id|participant_id|beneficiary_id|case_id|"
    r"registration_id|reg_id|ration_card"
    r")($|[_/])"
)

DATETIME_EXACT = {"start", "end", "today", "_submission_time", "starttime", "endtime",
                  "meta/start", "meta/end", "meta/today"}
BIRTHDATE_PATTERN = re.compile(r"(dob|date_of_birth|birth_?date|birthday|naissance)")

AGE_PATTERN = re.compile(r"(^|[_/])(age|age_years|ki_age|resp_age|respondent_age|hoh_age)($|[_/])")
AGE_BINS = [0, 5, 12, 18, 30, 45, 60, 200]           # Grenzen (links inklusive)
AGE_LABELS = ["0-4", "5-11", "12-17", "18-29", "30-44", "45-59", "60+"]

# Spalten, die als Freitext behandelt werden (zusätzlich zu XLSForm-Typ "text")
FREETEXT_PATTERN = re.compile(r"(_other$|/other$|_specify$|_comment|comments?$|_remarks?$|_note_text$|_explain)")

# Quasi-Identifikatoren für den k-Anonymitäts-Check (Regex auf Spaltennamen).
# Leer lassen = automatische Auswahl über QI_AUTO_PATTERN.
QUASI_IDENTIFIERS: list[str] = []
# Automatisch: Spalten, deren Name auf eines dieser Wörter ENDET (z. B. "B_ki_gender", "C_info_lga")
QI_AUTO_PATTERN = re.compile(
    r"(^|[_/])(gender|sex|age_group|age|disp_status|displacement_status|"
    r"settlement|village|community|ward|lga|district|camp|site)$"
)
QI_AUTO_EXCLUDE = re.compile(r"(source|origin|_other$|_final$)")
K_THRESHOLD = 5

# XLSForm-Typen -> Aktion
FORM_TYPE_ACTIONS = {
    "text": "freetext",
    "geopoint": "gps", "geotrace": "gps", "geoshape": "gps",
    "image": "drop", "audio": "drop", "video": "drop", "file": "drop",
    "background-audio": "drop", "barcode": "drop",
    "deviceid": "drop", "subscriberid": "drop", "simserial": "drop",
    "phonenumber": "drop", "username": "drop", "email": "drop", "audit": "drop",
    "start": "datetime", "end": "datetime", "today": "datetime",
    "date": "datetime", "datetime": "datetime",
}

# =============================================================================
# Ende der Konfiguration
# =============================================================================

PHONE_RE = re.compile(r"(?<!\w)\+?\d[\d\s\-/().]{6,}\d(?!\w)")
EMAIL_RE = re.compile(r"[\w.+\-]+@[\w\-]+\.[\w.\-]+")
LONGNUM_RE = re.compile(r"\d{7,}")


@dataclass
class Settings:
    gps: str = "drop"            # drop | round | keep
    gps_decimals: int = 2        # 2 Nachkommastellen ~ 1,1 km
    dates: str = "month"         # month | date | keep | drop
    freetext: str = "redact"     # redact | drop | keep
    age: str = "bin"             # bin | keep | drop
    k: int = K_THRESHOLD
    keep: set[str] = field(default_factory=set)
    drop: set[str] = field(default_factory=set)
    pseudo: set[str] = field(default_factory=set)
    qi: list[str] = field(default_factory=list)
    write_mapping: bool = False


# -----------------------------------------------------------------------------
# Hilfsfunktionen
# -----------------------------------------------------------------------------

def load_salt(path: Path) -> bytes:
    """Lädt den geheimen Salt oder erzeugt einen neuen (für konsistente Pseudonyme)."""
    if path.exists():
        return path.read_text().strip().encode()
    salt = secrets.token_hex(32)
    path.write_text(salt + "\n")
    try:
        path.chmod(0o600)
    except OSError:
        pass
    print(f"[i] Neuer Salt erzeugt: {path}  (geheim halten, nicht weitergeben!)")
    return salt.encode()


def normalize_value(v) -> str:
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def make_pseudonymizer(salt: bytes, mapping: dict):
    def pseudo(v):
        if pd.isna(v) or normalize_value(v) == "":
            return v
        raw = normalize_value(v)
        digest = hmac.new(salt, raw.encode(), hashlib.sha256).hexdigest()[:12]
        out = f"P_{digest}"
        mapping.setdefault(out, raw)
        return out
    return pseudo


def name_variants(col: str) -> tuple[str, str]:
    low = str(col).strip().lower()
    last = low.split("/")[-1]
    return low, last


def read_form(path: Path) -> dict[str, str]:
    """Liest das XLSForm (Sheet 'survey') und gibt {feldname: typ} zurück."""
    survey = pd.read_excel(path, sheet_name="survey", dtype=str)
    survey.columns = [c.strip().lower() for c in survey.columns]
    types = {}
    for _, row in survey.iterrows():
        name, typ = row.get("name"), row.get("type")
        if pd.isna(name) or pd.isna(typ):
            continue
        typ = str(typ).strip().lower().split(" ")[0]
        appearance = str(row.get("appearance", "")).lower()
        if typ == "image" and "signature" in appearance:
            typ = "image"
        types[str(name).strip().lower()] = typ
    return types


def form_type_for(col: str, form_types: dict[str, str]) -> str | None:
    """Ordnet eine Spalte einem Formularfeld zu (auch mit Gruppenpräfix 'grp/' oder 'grp_')."""
    if not form_types:
        return None
    low, last = name_variants(col)
    if last in form_types:
        return form_types[last]
    best = None
    for name in form_types:
        if low.endswith("_" + name) or low.endswith("/" + name):
            if best is None or len(name) > len(best):
                best = name
    return form_types[best] if best else None


def is_text(s: pd.Series) -> bool:
    """True für Textspalten (pandas 2: object, pandas 3: str)."""
    return s.dtype == object or pd.api.types.is_string_dtype(s)


def looks_like_freetext(s: pd.Series) -> bool:
    """Heuristik ohne XLSForm: viele unterschiedliche, längere Texte."""
    if not is_text(s):
        return False
    vals = s.dropna().astype(str)
    vals = vals[vals.str.strip() != ""]
    if len(vals) < 5:
        return False
    if pd.to_numeric(vals, errors="coerce").notna().mean() > 0.9:
        return False
    if pd.to_datetime(vals.head(50), errors="coerce", format="mixed").notna().mean() > 0.9:
        return False
    uniq_ratio = vals.nunique() / len(vals)
    has_spaces = vals.str.contains(r"\s").mean()
    return uniq_ratio > 0.3 and has_spaces > 0.3 and vals.str.len().mean() > 12


def redact_text(v):
    if pd.isna(v) or not isinstance(v, str):
        return v
    v = EMAIL_RE.sub("[EMAIL]", v)
    v = PHONE_RE.sub("[NUMMER]", v)
    v = LONGNUM_RE.sub("[NUMMER]", v)
    return v


def to_datetime(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce", utc=True, format="mixed")


# -----------------------------------------------------------------------------
# Klassifikation der Spalten
# -----------------------------------------------------------------------------

def classify(col: str, series: pd.Series, form_types: dict, st: Settings) -> tuple[str, str]:
    """Gibt (aktion, begründung) für eine Spalte zurück."""
    low, last = name_variants(col)

    # 1) Manuelle Overrides
    if low in st.keep or col in st.keep:
        return "keep", "manuell: --keep"
    if low in st.drop or col in st.drop:
        return "drop", "manuell: --drop"
    if low in st.pseudo or col in st.pseudo:
        return "pseudo", "manuell: --pseudo"

    # 2) Kobo-Metadaten & IDs
    if low in KEEP_EXACT:
        return "keep", "Kobo-Struktur"
    if low in DROP_EXACT or last in DROP_EXACT:
        return "drop", "Kobo-Metadaten / Geräte-ID"
    if low in PSEUDO_EXACT or last in PSEUDO_EXACT:
        return "pseudo", "Datensatz-ID"
    if low in DATETIME_EXACT or last in DATETIME_EXACT:
        return "datetime", "Zeitstempel"
    if BIRTHDATE_PATTERN.search(low):
        return "birthdate", "Geburtsdatum"

    # 3) XLSForm-Typ
    ftype = form_type_for(col, form_types)
    if ftype in FORM_TYPE_ACTIONS:
        action = FORM_TYPE_ACTIONS[ftype]
        # GPS-Teilspalten (_latitude etc.) und Freitext-Felder mit Namen/Telefon -> genauer prüfen
        if action == "freetext" and DROP_TOKEN_PATTERN.search(low):
            return "drop", f"XLSForm text + Name/Kontakt im Spaltennamen"
        if action == "freetext" and PSEUDO_TOKEN_PATTERN.search(low):
            return "pseudo", "XLSForm text + ID/Enumerator im Spaltennamen"
        return action, f"XLSForm-Typ '{ftype}'"

    # 4) Namensregeln
    if GPS_PATTERN.search(low):
        return "gps", "GPS/Koordinaten"
    if PSEUDO_TOKEN_PATTERN.search(low):
        return "pseudo", "ID/Enumerator"
    if DROP_TOKEN_PATTERN.search(low):
        return "drop", "Name/Kontakt/Medien"
    if AGE_PATTERN.search(low) and pd.to_numeric(series, errors="coerce").notna().sum() > 0:
        return "age", "Alter (numerisch)"
    if FREETEXT_PATTERN.search(low) and is_text(series):
        return "freetext", "'other'/Kommentarfeld"
    if not form_types and looks_like_freetext(series):
        return "freetext", "Heuristik: sieht nach Freitext aus – bitte prüfen"
    if ftype is not None and AGE_PATTERN.search(low) and ftype in ("integer", "decimal"):
        return "age", "Alter (XLSForm)"

    return "keep", ""


# -----------------------------------------------------------------------------
# Anwendung
# -----------------------------------------------------------------------------

def round_geopoint_string(v, decimals: int):
    """'lat lon alt prec' -> 'lat lon' gerundet."""
    if pd.isna(v):
        return v
    parts = str(v).split()
    if len(parts) < 2:
        return pd.NA
    try:
        return f"{round(float(parts[0]), decimals)} {round(float(parts[1]), decimals)}"
    except ValueError:
        return pd.NA


def anonymize_sheet(df: pd.DataFrame, sheet: str, form_types: dict, st: Settings,
                    pseudo, report: list) -> pd.DataFrame:
    df = df.copy()
    plan = {c: classify(c, df[c], form_types, st) for c in df.columns}

    # Interviewdauer berechnen, bevor start/end vergröbert werden
    starts = [c for c, (a, _) in plan.items() if a == "datetime" and name_variants(c)[1] in ("start", "starttime")]
    ends = [c for c, (a, _) in plan.items() if a == "datetime" and name_variants(c)[1] in ("end", "endtime")]
    duration_col = None
    if starts and ends:
        dur = (to_datetime(df[ends[0]]) - to_datetime(df[starts[0]])).dt.total_seconds() / 60
        duration_col = "interview_duration_min"
        df[duration_col] = dur.round(1)

    out_cols = {}
    for col, (action, reason) in plan.items():
        s = df[col]
        detail = reason
        low, last = name_variants(col)

        if action == "keep":
            out_cols[col] = s

        elif action == "drop":
            pass  # Spalte wird nicht übernommen

        elif action == "pseudo":
            out_cols[col] = s.map(pseudo)

        elif action == "gps":
            if st.gps == "keep":
                out_cols[col] = s
            elif st.gps == "round":
                if re.search(r"(altitude|precision|accuracy)", low):
                    detail += " -> entfernt (Höhe/Genauigkeit)"
                elif re.search(r"(latitude|longitude|(^|_)lat($|_)|(^|_)lon($|_)|lng)", low):
                    out_cols[col] = pd.to_numeric(s, errors="coerce").round(st.gps_decimals)
                    detail += f" -> gerundet auf {st.gps_decimals} Stellen"
                else:
                    out_cols[col] = s.map(lambda v: round_geopoint_string(v, st.gps_decimals))
                    detail += f" -> gerundet auf {st.gps_decimals} Stellen"
            action = f"gps:{st.gps}"

        elif action == "datetime":
            if st.dates == "keep":
                out_cols[col] = s
            elif st.dates != "drop":
                dt = to_datetime(s)
                fmt = "%Y-%m" if st.dates == "month" else "%Y-%m-%d"
                out_cols[col] = dt.dt.strftime(fmt)
            action = f"datetime:{st.dates}"

        elif action == "birthdate":
            out_cols[col] = to_datetime(s).dt.year.astype("Int64")
            detail += " -> nur Geburtsjahr"

        elif action == "age":
            if st.age == "keep":
                out_cols[col] = s
            elif st.age == "bin":
                num = pd.to_numeric(s, errors="coerce")
                out_cols[col + "_group"] = pd.cut(num, bins=AGE_BINS, labels=AGE_LABELS,
                                                  right=False).astype(object)
                detail += f" -> ersetzt durch '{col}_group'"
            action = f"age:{st.age}"

        elif action == "freetext":
            if st.freetext == "keep":
                out_cols[col] = s
            elif st.freetext == "redact":
                red = s.map(redact_text)
                mask = s.notna()
                n_changed = int((red[mask].astype(str) != s[mask].astype(str)).sum())
                out_cols[col] = red
                detail += f" -> {n_changed} Werte geschwärzt"
            action = f"freetext:{st.freetext}"

        report.append({"sheet": sheet, "column": col, "action": action, "detail": detail,
                       "non_empty_values": int(s.notna().sum())})

    if duration_col:
        out_cols[duration_col] = df[duration_col]
        report.append({"sheet": sheet, "column": duration_col, "action": "added",
                       "detail": f"aus {starts[0]} / {ends[0]} berechnet",
                       "non_empty_values": int(df[duration_col].notna().sum())})

    return pd.DataFrame(out_cols, index=df.index)


def k_anonymity_check(df: pd.DataFrame, st: Settings) -> tuple[list[str], pd.DataFrame]:
    if st.qi:
        pats = [re.compile(p, re.I) for p in st.qi]
        qi = [c for c in df.columns if any(p.search(str(c)) for p in pats)]
    else:
        qi = [c for c in df.columns if QI_AUTO_PATTERN.search(str(c).lower())
              and not QI_AUTO_EXCLUDE.search(str(c).lower())
              and df[c].nunique(dropna=True) > 1]
        # Rohes Alter ignorieren, wenn Altersgruppe vorhanden
        qi = [c for c in qi if not (c + "_group") in qi]
    if not qi:
        return [], pd.DataFrame()
    groups = df[qi].astype(str).value_counts(dropna=False).reset_index(name="n")
    small = groups[groups["n"] < st.k].sort_values("n")
    return qi, small


# -----------------------------------------------------------------------------
# Ein-/Ausgabe
# -----------------------------------------------------------------------------

def read_input(path: Path) -> dict[str, pd.DataFrame]:
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path, sheet_name=None)
    if path.suffix.lower() in (".csv", ".txt"):
        return {"data": pd.read_csv(path, sep=None, engine="python", encoding="utf-8-sig")}
    raise SystemExit(f"Unbekanntes Dateiformat: {path.suffix} (erwartet .xlsx oder .csv)")


def write_output(sheets: dict[str, pd.DataFrame], path: Path):
    if path.suffix.lower() == ".csv":
        if len(sheets) == 1:
            next(iter(sheets.values())).to_csv(path, index=False, encoding="utf-8-sig")
        else:
            for name, df in sheets.items():
                df.to_csv(path.with_name(f"{path.stem}_{name}.csv"), index=False, encoding="utf-8-sig")
    else:
        with pd.ExcelWriter(path, engine="openpyxl") as xw:
            for name, df in sheets.items():
                df.to_excel(xw, sheet_name=name[:31], index=False)


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Anonymisiert KoboToolbox-Exporte (XLSX/CSV).")
    p.add_argument("input", type=Path, help="Kobo-Export (.xlsx oder .csv)")
    p.add_argument("--form", type=Path, help="XLSForm (.xlsx) des Projekts – empfohlen")
    p.add_argument("-o", "--output", type=Path, help="Ausgabedatei (Standard: <input>_anonymized.<ext>)")
    p.add_argument("--salt-file", type=Path, default=Path(".kobo_salt"),
                   help="Datei mit geheimem Salt (Standard: ./.kobo_salt)")
    p.add_argument("--gps", choices=["drop", "round", "keep"], default="drop")
    p.add_argument("--gps-decimals", type=int, default=2, help="Nachkommastellen bei --gps round")
    p.add_argument("--dates", choices=["month", "date", "keep", "drop"], default="month")
    p.add_argument("--free-text", choices=["redact", "drop", "keep"], default="redact")
    p.add_argument("--age", choices=["bin", "keep", "drop"], default="bin")
    p.add_argument("--keep", nargs="*", default=[], help="Spalten unverändert behalten")
    p.add_argument("--drop", nargs="*", default=[], help="Spalten zusätzlich entfernen")
    p.add_argument("--pseudo", nargs="*", default=[], help="Spalten zusätzlich pseudonymisieren")
    p.add_argument("--qi", nargs="*", default=[], help="Quasi-Identifikatoren (Regex) für k-Check")
    p.add_argument("-k", type=int, default=K_THRESHOLD, help="Mindestgruppengröße für k-Check")
    p.add_argument("--write-mapping", action="store_true",
                   help="Pseudonym->Original-Tabelle schreiben (NUR lokal aufbewahren!)")
    p.add_argument("--dry-run", action="store_true", help="Nur Plan/Report, keine Daten schreiben")
    return p.parse_args(argv)


def main(argv=None):
    a = parse_args(argv)
    lower = lambda xs: {x.strip().lower() for x in xs}
    st = Settings(gps=a.gps, gps_decimals=a.gps_decimals, dates=a.dates, freetext=a.free_text,
                  age=a.age, k=a.k, keep=lower(a.keep), drop=lower(a.drop),
                  pseudo=lower(a.pseudo), qi=a.qi, write_mapping=a.write_mapping)

    if not a.input.exists():
        raise SystemExit(f"Datei nicht gefunden: {a.input}")
    out = a.output or a.input.with_name(f"{a.input.stem}_anonymized{a.input.suffix if a.input.suffix.lower() == '.csv' else '.xlsx'}")
    report_path = out.with_name(f"{a.input.stem}_anonymization_report.csv")
    small_path = out.with_name(f"{a.input.stem}_small_groups.csv")

    form_types = read_form(a.form) if a.form else {}
    if a.form:
        print(f"[i] XLSForm geladen: {len(form_types)} Felder")

    sheets = read_input(a.input)
    print(f"[i] Eingelesen: {a.input.name} – " +
          ", ".join(f"{n} ({len(d)} Zeilen, {d.shape[1]} Spalten)" for n, d in sheets.items()))

    mapping: dict[str, str] = {}
    pseudo = make_pseudonymizer(load_salt(a.salt_file), mapping)
    report: list[dict] = []
    result = {name: anonymize_sheet(df, name, form_types, st, pseudo, report)
              for name, df in sheets.items()}

    rep = pd.DataFrame(report)
    print("\n=== Plan / Ergebnis (nur veränderte Spalten) ===")
    changed = rep[rep["action"] != "keep"]
    with pd.option_context("display.max_rows", 500, "display.width", 200, "display.max_colwidth", 70):
        print(changed[["sheet", "column", "action", "detail"]].to_string(index=False) if len(changed) else "(keine)")
    print("\nZusammenfassung:", rep["action"].str.split(":").str[0].value_counts().to_dict())

    # k-Anonymitäts-Check auf dem Hauptsheet
    main_name = next(iter(result))
    qi, small = k_anonymity_check(result[main_name], st)
    if qi:
        n_rows = len(result[main_name])
        n_small = int(small["n"].sum()) if len(small) else 0
        print(f"\n=== k-Anonymitäts-Check (k={st.k}) auf '{main_name}' ===")
        print(f"Quasi-Identifikatoren: {qi}")
        print(f"{len(small)} Kombinationen mit n<{st.k}; betrifft {n_small} von {n_rows} Zeilen "
              f"({n_small / max(n_rows, 1):.1%}).")
        if n_small:
            print("-> Erwägen: Orte gröber machen (--drop für Dorf/Settlement), Kategorien zusammenfassen,"
                  " oder mit --qi gezielt prüfen. Details: " + small_path.name)
    else:
        print("\n[i] Keine Quasi-Identifikatoren gefunden – k-Check übersprungen (ggf. --qi angeben).")

    unmatched = [c for c in rep["column"] if form_types and form_type_for(c, form_types) is None
                 and not str(c).startswith("_")]
    if form_types and len(unmatched) > 0.5 * len(rep):
        print("\n[!] Viele Spalten passen nicht zum XLSForm – wurde mit Labels statt "
              "'XML values and headers' exportiert?")

    if a.dry_run:
        rep.to_csv(report_path, index=False, encoding="utf-8-sig")
        print(f"\n[dry-run] Keine Daten geschrieben. Report: {report_path}")
        return

    write_output(result, out)
    rep.to_csv(report_path, index=False, encoding="utf-8-sig")
    if qi:
        small.to_csv(small_path, index=False, encoding="utf-8-sig")
    if st.write_mapping:
        mp = out.with_name(f"{a.input.stem}_PSEUDONYM_MAPPING_PRIVATE.csv")
        pd.DataFrame(mapping.items(), columns=["pseudonym", "original"]).to_csv(mp, index=False)
        print(f"[!] Mapping geschrieben (NICHT weitergeben): {mp}")

    print(f"\n[OK] Anonymisierte Daten: {out}")
    print(f"[OK] Report:             {report_path}")
    print("     Bitte den Report kurz prüfen, bevor die Datei geteilt wird.")


if __name__ == "__main__":
    sys.exit(main())
