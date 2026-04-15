# ITGC SAP HANA App

Initial MVP for loading Excel and TXT files and performing integrity checks.

## First-phase scope

- Load `.txt`, `.csv`, `.xlsx`, and `.xlsm` files
- Validate required columns
- Detect missing values in mandatory fields
- Produce a summary of valid and invalid rows
- Export an Excel validation report to the output folder
- Work through a local PySide6 desktop screen in Hebrew with native RTL support

## Project structure

- `src/readers/` - file readers for text and Excel
- `src/validators/` - validation engine
- `src/models/` - shared result models
- `src/pipeline.py` - orchestrates reading and validation
- `data/input/` - sample incoming files
- `data/output/` - generated Excel validation reports
- `tests/` - automated smoke tests

## Quick start

1. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

2. Open the desktop application:

   ```bash
   python -m src.main
   ```

3. Validate a file from the command line if needed:

   ```bash
   python -m src.main data/input/sample.txt --required user_id name
   ```

4. Run tests:

   ```bash
   python -m unittest discover -s tests -v
   ```
