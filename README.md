# PV Elite Extraction Platform

Extracts data from PV Elite pressure-vessel calculation reports (PDF) into
Excel and a SQLite database, with an optional natural-language search layer
(vector DB + local LLM) on top.

## What it does

1. **Import** - reads PV Elite PDF reports, extracts element-level and
   equipment-level data (geometry, thickness, weights, loads, platforms,
   weight breakdown).
2. **Store** - writes each tag to a SQLite database. Re-importing the same
   Project + Tag No updates it instead of duplicating.
3. **Export** - produces a multi-sheet Excel workbook, either straight from
   PDFs or regenerated later from the database (no PDFs needed).
4. **Manage** - browse, add, edit, and delete database records from the GUI.
5. **Ask** (optional) - question the data in plain English (exact figures via
   SQL, explanations via vector search over report text).

## Files

| File | Purpose |
|---|---|
| `pvelite_gui.py` | Desktop app - start here |
| `pvelite_dynamic_extractor.py` | Core PDF extraction engine |
| `bulk_extract.py` | Batch runner; builds the Excel workbook |
| `db_schema_sqlite.sql` | Database schema (frozen, versioned) |
| `db_writer.py` | Writes extraction results to the database |
| `db_export.py` | Regenerates the Excel workbook from the database |
| `db_manager_gui.py` | Add / edit / delete records window |
| `vector_db.py` | Embeds report sections for semantic search |
| `copilot.py` | Natural-language question routing (SQL vs vector) |

## Setup

Requires Python 3.10+.

```
pip install pandas openpyxl pymupdf
```

Optional, only for the natural-language search layer:

```
pip install chromadb sentence-transformers
ollama pull qwen3:14b
```

## Input folder layout

```
Input/
  ProjectA/
    Columns/            *.pdf
    HeatExchangers/     *.pdf
    Vessels_Hor/        *.pdf
    Vessels_Vert/       *.pdf
  ProjectB/
    ...
```

## Using the app

```
python pvelite_gui.py
```

1. **Input folder** - browse to your `Input` folder. Each project appears as
   a column with its categories listed underneath; tick what to import.
2. **Output file** - where the Excel workbook is saved.
3. **Also save to database** - tick to write to SQLite as well; pick the
   database file. To share one database across people, put the file on a
   shared drive and have everyone point at it.
4. **Run Extraction**.
5. **Manage Database...** - view, add, edit, delete records.
6. **Export DB to Excel...** - regenerate a workbook from stored data.

Command line alternatives:

```
python bulk_extract.py <reports_folder> <output.xlsx>
python db_export.py <database.db> <output.xlsx> [--project P] [--tag T]
```

## Excel output sheets

`Master_Element_Data` (one row per element), `Loads_Foundation_Support`,
`Work_Volume`, `ML_Summary`, `ML_Encoding_Key`, `Platform_List`,
`weight_summ`, `Equipment_Summary`. Every sheet carries Project and Category.

## Database tables

`equipment_master` (one row per tag), `element_data`, `platform_data`,
`weight_summation`, `mechanical_details`, `documents`, `schema_version`.
Deleting a tag removes all its child rows automatically. Schema changes are
additive only, and `db_writer.ensure_schema()` upgrades older databases in
place safely.

## Validation status - read this

| Category | Status |
|---|---|
| Vessels_Vert, Columns | **Validated** against real reports |
| Vessels_Hor | **Partial** - see below |
| HeatExchangers | **Not supported** - see below |

**Horizontal vessels:** element-level tables (thickness, weight, surface area)
share PV Elite's general layout and may extract correctly, but this has not
been checked against a real horizontal report. Saddle-supported vessels are
now detected and flagged with a warning; saddle stress and geometry data are
**not** extracted. Skirt/basering fields will be empty.

**Heat exchangers:** a different report structure entirely (shell/tube,
tubesheet, channel, floating head). Nothing here parses it. Expect zero
elements and a warning. The database schema is already prepared for it
(`equipment_type`, `mechanical_details`), but the extractor must be written
against a real sample report.

Both are selectable in the GUI and labelled "(unvalidated)". Do not rely on
their numbers until they are checked.

## Known limitations

- Semantic search and the LLM layer (`vector_db.py`, `copilot.py`) were
  built and tested with stand-in components only - the real embedding model
  and Ollama server could not be reached in the build environment. Test both
  on your own machine before relying on them.
- SQLite suits people importing at different times. Several people writing
  to a shared file at the same instant is not reliable.
- Div-1 stress tables are converted from kgf/cm2 to MPa; skirt allowable
  stress uses the material's ambient value as a fallback.

## Packaging as a Windows .exe

See `PACKAGING_INSTRUCTIONS.md`. Build on Windows; include all `.py` files
and `db_schema_sqlite.sql` next to `pvelite_gui.py`.
