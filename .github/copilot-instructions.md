# Copilot Instructions for ConciliacionBancaria

## Project Overview
This application compares bank statement files (PDF/Excel) from Itaú and BROU against records in a PostgreSQL database. It is designed for financial reconciliation, focusing on matching transaction amounts and dates between bank files and database records.

## Architecture & Data Flow
- **UI Layer**: `src/main.py` provides a Tkinter GUI for selecting files, configuring date ranges, and running comparisons.
- **File Readers**: 
  - `src/lectorItau.py` and `src/lectorBrou.py` parse Excel files from Itaú and BROU, handling header normalization, flexible column mapping, and conversion from `.xls` to `.xlsx` using Excel COM (Windows only).
  - Both readers output a standardized DataFrame with columns like `Fecha`, `Débito`, `Crédito`, `Monto`, etc.
- **Database Access**: 
  - `src/db.py` connects to a PostgreSQL database using SQLAlchemy and psycopg2. The main query pulls from the `cpf_contaux` table.
- **Comparison Logic**: 
  - `main.py` matches transactions by absolute amount and date, reporting found/missing records and exporting results to Excel.

## Key Workflows
- **Run the App**: Launch via `python src/main.py` (requires Python, pandas, openpyxl, SQLAlchemy, psycopg2, and Excel installed for `.xls` conversion).
- **File Processing**: Select an Itaú/BROU file (Excel), process it, then query the database for the selected date range.
- **Comparison**: Matches are based on absolute value of `Monto` (Excel) vs `imp_neto` (DB) and date (`Fecha` vs `fec_doc`).
- **Export**: Results can be exported to Excel, including a summary sheet.

## Conventions & Patterns
- **Flexible Header Mapping**: Readers use regex and accent-stripping to map diverse column headers to a standard schema.
- **Amount Normalization**: Handles negative values in parentheses, thousands separators, and missing values.
- **Footer Detection**: Skips summary/footer rows using keyword hints.
- **Windows-Only XLS Conversion**: `.xls` files are converted to `.xlsx` using Excel COM automation; this requires Excel to be installed.
- **Error Handling**: GUI logs errors and shows message boxes for user feedback.

## Integration Points
- **External Dependencies**: pandas, openpyxl, SQLAlchemy, psycopg2, win32com (for Excel automation), pythoncom.
- **Database**: PostgreSQL at `10.10.1.162`, database `m_cpf_contaux`, table `cpf_contaux`.
- **File Inputs**: Excel files from Itaú and BROU; PDF support is not implemented in code (despite README mention).

## Examples
- To process an Itaú file: `procesar_itau('Archivos/Estado_De_Cuenta_2769087_-_2025-10-22T171649.713.xls')`
- To process a BROU file: `procesar_brou('Archivos/Detalle_Movimiento_Cuenta_-_2025-10-22T171522.294.xls')`
- To query the database: `obtener_df_bd()`

## File Reference
- `src/main.py`: GUI, workflow orchestration
- `src/lectorItau.py`: Itaú file reader
- `src/lectorBrou.py`: BROU file reader
- `src/db.py`: Database connection/query
- `Archivos/`: Example input files

## Special Notes
- All `.xls` file processing requires Windows and Excel installed.
- Matching logic is strict: only exact matches on absolute amount and date are considered.
- No automated tests or CI/CD scripts are present.

---
For questions or unclear patterns, please ask for clarification or provide feedback to improve these instructions.
