# SPC GUI Replica

A standalone Tkinter desktop application that recreates the supplied SPC analysis screen and wires the **DefectSPC** / **SPC Excel** actions to a linked backend Excel workbook.

## Run

```bash
python3 spc_gui.py
```

## SPC Excel workflow

You have two options:

1. **Default folder option**
   - Put your workbook at `spc_backend/SPC_Backend.xlsx`.
   - Click **DefectSPC** or **SPC Excel**.
   - The workbook opens directly in the operating system's default spreadsheet application.

2. **Link-file option**
   - Click **Link SPC Excel** in the ribbon or **Link File** in the left panel.
   - Select the exact `.xlsx`, `.xls`, `.xlsm`, or `.csv` workbook one time.
   - After that, **DefectSPC** / **SPC Excel** opens that linked file directly without opening a folder or asking again.

CSV files are also previewed in the top SPC result grid before the external spreadsheet application opens.
