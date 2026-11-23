# Kobo-Daten anonymisieren

## Installation

```bash
pip install -r requirements.txt
```

## Export aus KoboToolbox

1. **Daten**: Projekt > *Data* > *Downloads* > Typ **XLS**, Werte und Kopfzeilen = **XML values and headers**.
2. **Formular** (empfohlen): Projekt > *Form* > *Download XLS*.

## Anwendung

Probelauf (schreibt nur den Report, keine Daten):

```bash
python kobo_anonymize.py export.xlsx --form formular.xlsx --dry-run
```

Echter Lauf:

```bash
python kobo_anonymize.py export.xlsx --form formular.xlsx
```

Weitergeben darfst du nur `export_anonymized.xlsx` (und gern auch den Report).
**Nie weitergeben:** `.kobo_salt` und `*_PSEUDONYM_MAPPING_PRIVATE.csv`.

## Wichtige Optionen

| Option | Bedeutung |
|---|---|
| `--gps drop\|round\|keep` | GPS entfernen (Standard), runden (`--gps-decimals 2` ≈ 1 km) oder behalten |
| `--dates month\|date\|keep\|drop` | Zeitstempel auf Monat (Standard) oder Tag vergröbern |
| `--free-text redact\|drop\|keep` | Telefonnummern/E-Mails in Freitext schwärzen (Standard) oder Freitext entfernen |
| `--age bin\|keep\|drop` | Alter in Altersgruppen umwandeln (Standard) |
| `--keep SPALTE ...` | Spalte unverändert lassen (falls fälschlich entfernt) |
| `--drop SPALTE ...` | Spalte zusätzlich entfernen |
| `--pseudo SPALTE ...` | Spalte zusätzlich pseudonymisieren |
| `--qi REGEX ...`, `-k 5` | Quasi-Identifikatoren und Mindestgruppengröße für den k-Anonymitäts-Check |
| `--write-mapping` | Zuordnung Pseudonym → Original lokal speichern |

Die Regeln (welche Spaltennamen als Name, Telefon, ID … gelten) stehen oben im Skript
im Abschnitt `KONFIGURATION` und lassen sich dort erweitern.
