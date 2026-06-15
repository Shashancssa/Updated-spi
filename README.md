# SPC GUI Replica

A standalone Tkinter desktop application that recreates the supplied SPC analysis screen and wires the **DefectSPC** / **Excel** actions to an SPC backend Excel-file workflow.

## Run

```bash
python3 spc_gui.py
```

## Excel workflow

- Click **DefectSPC** in the ribbon or **Excel** in the main condition panel.
- Select an `.xlsx`, `.xls`, `.xlsm`, or `.csv` file.
- The app opens the selected file with the operating system's default spreadsheet application.
- CSV files are also previewed in the top SPC result grid before the external application opens.
