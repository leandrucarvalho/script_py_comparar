# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Script Does

`comparar_concluintes.py` reconciles two data sources for a BFD training program:

1. An Excel spreadsheet listing students (concluintes) per state and class (turma)
2. A local folder tree downloaded from Google Drive containing evidence folders named after students

It reports which Drive folders have no matching spreadsheet entry (candidates for removal) and which spreadsheet students have no Drive folder (missing evidence).

## Running the Script

```powershell
python comparar_concluintes.py
```

**Dependency:** `openpyxl` — install with `pip install openpyxl`

## Configuration

Edit the three constants at the top of the file before running:

| Constant | Purpose |
|---|---|
| `EXCEL_PATH` | Absolute path to the `.xlsx` file |
| `SHEET_NAME` | Name of the worksheet tab to read |
| `ESTADOS` | List of state names to filter (matches column 0 of the sheet) |
| `DRIVE_ROOT` | Root folder containing `turma/` subdirectories, each with one subfolder per student |

## Expected Folder Structure

```
DRIVE_ROOT/
  T01PEC1/          ← turma folder (or "Turma01PEC1" — prefix normalized)
    joao_silva/     ← student folder
    maria_souza/
  T02PEC1/
    ...
```

The Excel sheet expects columns in order: `[estado, turma, nome]`.

## Matching Logic

Names are matched with `nomes_em_comum()`: two names match if they share at least 2 significant words after normalizing (lowercasing, stripping accents, removing particles like *de/da/do/dos/das/e*). This tolerates abbreviations and minor name variations between the spreadsheet and folder names.
