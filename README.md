# EndpointUDFs

This project discovers Domino model endpoints, generates Excel-DNA UDFs, builds the Excel add-in, and publishes the add-in artifacts.

## Run

Run the end-to-end workflow from the repo root:

```bash
python run_all.py
```

What `run_all.py` does:
- Discovers model endpoints in the configured Domino project.
- Generates the C# UDFs and Excel-DNA config.
- Builds the Excel add-in.
- Copies the add-in files to `/mnt/artifacts` as:
  - `DominoExcelUDFsAddIn.xll` (32-bit)
  - `DominoExcelUDFsAddIn64.xll` (64-bit)

## Load The Add-In In Excel

============================================================

To use the add-in:
  1. Open Excel
  2. Go to File > Options > Add-ins
  3. At the bottom, select 'Excel Add-ins' and click 'Go...'
  4. Click 'Browse...' and select the .xll file
  5. Click OK to enable the add-in
