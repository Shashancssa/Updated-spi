# SPC Analysis GUI Replica

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

## Editing fields

- Double-click any visible cell in the main result grid to edit values such as **Model**, **StationID**, **TotalBoard**, **YieldRate**, and counts.
- Double-click any visible cell in the lower condition grid to edit values such as **Tester**, **Model**, **StationID**, **Lotno**, and **Modifie**.
- Press **Enter** or click away to save the edit; press **Escape** to cancel.

## DB-process-only options

Options that are not part of the Excel workflow, such as **Utilities**, **Other**, **ModelList**, **ListView/Report**, **Histogram**, and **Defect Chart**, show a **DB Process Error** message because they require the real database process.

## Application icon

The GUI creates `.spc_runtime/spc_backend.ico` at runtime and uses it as the Windows title-bar/taskbar icon, so the PR contains only text files and avoids binary-file upload errors. Tkinter also applies the same green icon colors directly at runtime.

## PCB View window

Click **PcbView** in the ribbon to open a separate PCB View window matching the supplied inspection screen. The PCB View includes editable filter fields, editable panel-information entries, an editable defect grid, an inspection-board drawing, right-side analysis plots, and an editable actual-test-data grid.
