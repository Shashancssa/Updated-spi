"""SPC replica GUI with Excel-launching backend action.

Run with: python3 spc_gui.py
"""
from __future__ import annotations

import csv
import json
import os
import platform
import subprocess
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "SPC Analysis Backend"
CONFIG_FILE = Path(".spc_gui_config.json")
DEFAULT_SPC_EXCEL = Path("spc_backend/SPC_Backend.xlsx")
SAMPLE_ROWS = [
    ("SPI_TRI_356_R0_(ME420079)_SS", "SPI-LINE01SPI", "4", "100%", "0%", "0", "0%", "0", "100%", "4", "0", "100%", "9"),
    ("SPI_TRI_357_R0_(ME420079)_CS", "SPI-LINE01SPI", "4", "100%", "0%", "0", "0%", "0", "100%", "4", "0", "100%", "0"),
]


def open_with_default_app(path: Path) -> None:
    """Open a file with the operating system's default application."""
    system = platform.system()
    if system == "Windows":
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif system == "Darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


class SPCReplica(tk.Tk):
    """Desktop GUI replica for the SPC analysis screen."""

    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1366x730")
        self.minsize(1100, 650)
        self.configure(bg="#dfeaf8")
        self.loaded_excel: Path | None = None
        self.spc_excel_path = self._load_configured_excel_path()
        self._build_styles()
        self._set_window_icon()
        self._build_ui()

    def _set_window_icon(self) -> None:
        """Create a small in-memory SPC icon for the application window."""
        icon = tk.PhotoImage(width=32, height=32)
        icon.put("#1d6fa5", to=(0, 0, 32, 32))
        icon.put("#d7e8fb", to=(3, 3, 29, 29))
        icon.put("#1d6fa5", to=(5, 22, 27, 25))
        icon.put("#ffb347", to=(7, 15, 12, 22))
        icon.put("#2f80ed", to=(14, 10, 19, 22))
        icon.put("#27ae60", to=(21, 6, 26, 22))
        self.iconphoto(True, icon)
        self._icon_image = icon

    def _load_configured_excel_path(self) -> Path:
        """Return the linked SPC workbook path saved by the operator, or the default path."""
        if CONFIG_FILE.exists():
            try:
                data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
                configured = data.get("spc_excel_path")
                if configured:
                    return Path(configured).expanduser()
            except json.JSONDecodeError:
                pass
        return DEFAULT_SPC_EXCEL

    def _save_configured_excel_path(self, path: Path) -> None:
        CONFIG_FILE.write_text(json.dumps({"spc_excel_path": str(path)}, indent=2), encoding="utf-8")

    def _build_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Ribbon.TFrame", background="#c6ddf5")
        style.configure("Panel.TLabelframe", background="#dfeaf8", bordercolor="#9bbbd8")
        style.configure("Panel.TLabelframe.Label", background="#dfeaf8", foreground="#1b3f5e")
        style.configure("Blue.TButton", padding=5, background="#b9dcf5", foreground="#153b59")
        style.map("Blue.TButton", background=[("active", "#8ec1ea")])
        style.configure("Treeview", rowheight=22, fieldbackground="#e6f1ff", background="#e6f1ff")
        style.configure("Treeview.Heading", background="#c9dff6", foreground="#173d5c")

    def _build_ui(self) -> None:
        self._build_title_bar()
        self._build_ribbon()
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True)
        self._build_left_panel(body)
        self._build_workspace(body)
        self._build_status_bar()

    def _build_title_bar(self) -> None:
        bar = tk.Frame(self, height=28, bg="#d7e8fb", highlightthickness=1, highlightbackground="#a7c0df")
        bar.pack(fill="x")
        tk.Label(bar, text="⚙", bg="#ffd65a", fg="#b87400", font=("Segoe UI", 14), width=3).pack(side="left", padx=(3, 12), pady=3)

    def _build_ribbon(self) -> None:
        tabs = tk.Frame(self, bg="#b8d5f0")
        tabs.pack(fill="x")
        for i, name in enumerate(("PCB", "Utilities", "Other")):
            bg = "#d7e8fb" if i == 0 else "#b8d5f0"
            tab = tk.Label(tabs, text=name, bg=bg, fg="#143856", padx=18, pady=8, relief="ridge" if i == 0 else "flat")
            tab.pack(side="left")
            if name != "PCB":
                tab.bind("<Button-1>", lambda _event, option=name: self._db_process_error(option))

        ribbon = tk.Frame(self, height=92, bg="#c6ddf5", highlightthickness=1, highlightbackground="#94b7d8")
        ribbon.pack(fill="x")
        tools = [
            ("🔍", "ModelList", self._noop),
            ("▤", "ListView/Report", self._noop),
            ("▦", "PcbView", self._noop),
            ("▥", "Histogram", self._noop),
            ("📊", "Defect\nChart", self._noop),
            ("📁", "Link SPC\nExcel", self.choose_spc_excel_link),
            ("〰", "DefectSPC", self.open_linked_spc_excel),
        ]
        for icon, label, command in tools:
            box = tk.Frame(ribbon, width=78, height=90, bg="#c6ddf5", highlightbackground="#aac4de", highlightthickness=1)
            box.pack(side="left", padx=1, pady=2)
            box.pack_propagate(False)
            tk.Button(box, text=icon, command=command, font=("Segoe UI", 20), bg="#c6ddf5", relief="flat", activebackground="#9fc5eb").pack(pady=(8, 0))
            tk.Label(box, text=label, bg="#c6ddf5", fg="#163d5c", font=("Segoe UI", 8)).pack()
        tk.Label(ribbon, text="Analysis by PCB", bg="#abcbe8", fg="#254e70").pack(side="left", fill="y", padx=(0, 8), pady=(70, 0))

    def _build_left_panel(self, parent: ttk.Frame) -> None:
        side = tk.Frame(parent, width=310, bg="#dce9f8", highlightthickness=1, highlightbackground="#9fbdda")
        side.pack(side="left", fill="y", padx=(2, 5), pady=5)
        side.pack_propagate(False)
        lf = ttk.LabelFrame(side, text="Main Condition", style="Panel.TLabelframe")
        lf.pack(fill="both", expand=True, padx=22, pady=8)
        tk.Label(lf, text="◷", bg="#dfeaf8", font=("Segoe UI", 26)).grid(row=0, column=0, rowspan=2, padx=8, pady=15)
        tk.Label(lf, text="Start", bg="#dfeaf8").grid(row=0, column=1, sticky="e")
        tk.Label(lf, text="End", bg="#dfeaf8").grid(row=1, column=1, sticky="e")
        for r, date in enumerate(("2026/05/01", "2026/06/04")):
            ttk.Entry(lf, width=14).grid(row=r, column=2, padx=4, pady=6)
            lf.grid_slaves(row=r, column=2)[0].insert(0, date)
            ttk.Entry(lf, width=10).grid(row=r, column=3, padx=4)
            lf.grid_slaves(row=r, column=3)[0].insert(0, "06:28:40" if r == 0 else "23:28:40")
        tk.Label(lf, text="Period", bg="#dfeaf8").grid(row=2, column=0, sticky="w", padx=8, pady=(10, 0))
        for i, p in enumerate(("1 Hour", "8 Hour", "1 Day", "2 Day", "Week", "Month")):
            ttk.Button(lf, text=p, style="Blue.TButton").grid(row=3, column=i % 4, padx=2, pady=4, sticky="ew")
        ttk.Button(lf, text="🔎 Query", style="Blue.TButton", state="disabled").grid(row=4, column=2, columnspan=2, sticky="ew", padx=5, pady=22)
        ttk.Button(lf, text="📗 SPC Excel", style="Blue.TButton", command=self.open_linked_spc_excel).grid(row=5, column=2, columnspan=2, sticky="ew", padx=5, pady=0)
        ttk.Button(lf, text="🔗 Link File", style="Blue.TButton", command=self.choose_spc_excel_link).grid(row=5, column=0, columnspan=2, sticky="ew", padx=5, pady=0)
        tk.Checkbutton(lf, text="Select Condition", bg="#dfeaf8", variable=tk.BooleanVar(value=True)).grid(row=6, column=0, columnspan=2, sticky="w", pady=18)
        for idx, (lab, val) in enumerate((("Type", "By Model"), ("Mode", "By Panel"), ("PCS", "NONE"), ("LotNo", "*")), start=7):
            tk.Label(lf, text=lab, bg="#dfeaf8").grid(row=idx, column=0, sticky="w", padx=8, pady=4)
            combo = ttk.Combobox(lf, values=[val], width=35)
            combo.set(val)
            combo.grid(row=idx, column=1, columnspan=3, sticky="ew", pady=4)
        self.condition_grid = ttk.Treeview(lf, columns=("Tester", "Model", "StationID", "Lotno", "Modifie"), show="headings", height=3)
        for col in self.condition_grid["columns"]:
            self.condition_grid.heading(col, text=col)
            self.condition_grid.column(col, width=60)
        self.condition_grid.grid(row=11, column=0, columnspan=4, sticky="ew", padx=8, pady=20)
        for row in SAMPLE_ROWS:
            self.condition_grid.insert("", "end", values=("SPI", row[0][:20], "LINE01SPI", "*", "*"))
        self._make_tree_editable(self.condition_grid)
        ttk.Button(lf, text="Refresh", style="Blue.TButton").grid(row=12, column=0, columnspan=2, sticky="ew", padx=8)
        ttk.Button(lf, text="Apply", style="Blue.TButton").grid(row=12, column=2, columnspan=2, sticky="ew", padx=8)

    def _build_workspace(self, parent: ttk.Frame) -> None:
        main = tk.Frame(parent, bg="#eef5ff")
        main.pack(side="left", fill="both", expand=True, pady=5, padx=(0, 3))
        cols = ("Model", "StationID", "TotalBoard", "YieldRate", "PassRate", "PassCount", "FailRate", "FailCount", "FalseAlarmR", "FalseAlarmC", "DPMO", "YieldRateSol", "SkipBoard")
        self.grid = ttk.Treeview(main, columns=cols, show="headings", height=8)
        for col in cols:
            self.grid.heading(col, text=col)
            self.grid.column(col, width=105 if col == "Model" else 80, stretch=True)
        self.grid.pack(fill="x")
        for i, row in enumerate(SAMPLE_ROWS):
            self.grid.insert("", "end", values=row, tags=("selected",) if i == 1 else ())
        self.grid.tag_configure("selected", background="#6d8dd8", foreground="white")
        self._make_tree_editable(self.grid)
        self.chart = tk.Canvas(main, bg="white", highlightthickness=1, highlightbackground="#707070")
        self.chart.pack(fill="both", expand=True, padx=60, pady=8)
        self._draw_chart()

    def _make_tree_editable(self, tree: ttk.Treeview) -> None:
        """Allow operators to edit Model, StationID, LotNo, and other visible grid fields."""
        tree.bind("<Double-1>", lambda event, target=tree: self._start_tree_cell_edit(event, target))

    def _start_tree_cell_edit(self, event: tk.Event, tree: ttk.Treeview) -> None:
        row_id = tree.identify_row(event.y)
        column_id = tree.identify_column(event.x)
        if not row_id or not column_id:
            return
        column_index = int(column_id.replace("#", "")) - 1
        bbox = tree.bbox(row_id, column_id)
        if not bbox:
            return
        x, y, width, height = bbox
        values = list(tree.item(row_id, "values"))
        current_value = values[column_index] if column_index < len(values) else ""
        editor = ttk.Entry(tree)
        editor.insert(0, current_value)
        editor.select_range(0, "end")
        editor.focus_set()
        editor.place(x=x, y=y, width=width, height=height)

        def save_edit(_event: tk.Event | None = None) -> None:
            if not editor.winfo_exists():
                return
            while column_index >= len(values):
                values.append("")
            values[column_index] = editor.get()
            tree.item(row_id, values=values)
            editor.destroy()

        def cancel_edit(_event: tk.Event | None = None) -> None:
            if editor.winfo_exists():
                editor.destroy()

        editor.bind("<Return>", save_edit)
        editor.bind("<FocusOut>", save_edit)
        editor.bind("<Escape>", cancel_edit)

    def _draw_chart(self) -> None:
        self.chart.bind("<Configure>", lambda _event: self._render_chart())
        self.after(100, self._render_chart)

    def _render_chart(self) -> None:
        c = self.chart
        c.delete("all")
        w, h = max(c.winfo_width(), 900), max(c.winfo_height(), 420)
        left, bottom, top, right = 60, h - 45, 20, w - 130
        c.create_line(left, top, left, bottom, fill="black", width=2)
        c.create_line(left, bottom, right, bottom, fill="black", width=2)
        for i in range(0, 106, 5):
            y = bottom - (bottom - top) * i / 105
            c.create_text(left - 18, y, text=str(i), font=("Segoe UI", 8))
        groups = [left + (right - left) * .28, left + (right - left) * .78]
        labels = ["LINE01SPI/SPI_TRI_356_R0_(ME420079)_SS", "LINE01SPI/SPI_TRI_357_R0_(ME420079)_CS"]
        colors = [("green", 0), ("red", 0), ("#ff7f00", 100), ("blue", 100)]
        for x, label in zip(groups, labels):
            c.create_text(x, bottom + 20, text=label, font=("Segoe UI", 8))
            for j, (color, val) in enumerate(colors):
                bar_h = (bottom - top) * val / 105
                x0 = x - 55 + j * 35
                c.create_rectangle(x0, bottom - bar_h, x0 + 35, bottom, fill=color, outline=color)
                c.create_text(x0 + 17, bottom - bar_h - 16, text=str(val), font=("Segoe UI", 8))
        for k, (name, color) in enumerate((("PassRate", "green"), ("FailRate", "red"), ("FalseAlarmRate", "#ff7f00"), ("YieldRate", "blue"))):
            y = 52 + k * 18
            c.create_rectangle(right + 20, y - 7, right + 30, y + 3, fill=color)
            c.create_text(right + 38, y, text=name, anchor="w", font=("Segoe UI", 8))

    def _build_status_bar(self) -> None:
        status = tk.Frame(self, height=22, bg="#d7e8fb", highlightthickness=1, highlightbackground="#a7c0df")
        status.pack(fill="x", side="bottom")
        self.status_left = tk.Label(status, text=f"Linked SPC Excel: {self.spc_excel_path}", bg="#d7e8fb", fg="#173d5c", anchor="w")
        self.status_left.pack(side="left", padx=8)
        tk.Label(status, text="SPC Backend", bg="#d7e8fb", fg="#173d5c", anchor="e").pack(side="right", padx=8)

    def choose_spc_excel_link(self) -> None:
        """Let the operator link the exact workbook that the SPC button should open."""
        filetypes = [("Excel workbooks", "*.xlsx *.xls *.xlsm *.csv"), ("All files", "*.*")]
        filename = filedialog.askopenfilename(title="Link SPC Excel workbook", filetypes=filetypes)
        if not filename:
            return
        self.spc_excel_path = Path(filename)
        self._save_configured_excel_path(self.spc_excel_path)
        if hasattr(self, "status_left"):
            self.status_left.configure(text=f"Linked SPC Excel: {self.spc_excel_path}")
        messagebox.showinfo("SPC Backend", f"SPC option linked to:\n{self.spc_excel_path}")

    def open_linked_spc_excel(self) -> None:
        """Open the linked workbook directly when the SPC option is clicked."""
        path = self.spc_excel_path
        if not path.exists():
            messagebox.showwarning(
                "SPC Backend",
                "No linked SPC Excel workbook was found.\n\n"
                f"Place your workbook here:\n{DEFAULT_SPC_EXCEL}\n\n"
                "Or click 'Link SPC Excel' / 'Link File' to choose the workbook once.",
            )
            return
        self.loaded_excel = path
        if path.suffix.lower() == ".csv":
            self._load_csv_preview(path)
        try:
            open_with_default_app(path)
            messagebox.showinfo("SPC Backend", f"Linked SPC Excel opened:\n{path}")
        except Exception as exc:  # noqa: BLE001 - display OS integration failures to the operator.
            messagebox.showwarning("SPC Backend", f"Linked file:\n{path}\n\nCould not open it automatically: {exc}")

    def _load_csv_preview(self, path: Path) -> None:
        with path.open(newline="", encoding="utf-8-sig") as handle:
            reader = csv.reader(handle)
            rows = list(reader)[:20]
        if not rows:
            return
        for item in self.grid.get_children():
            self.grid.delete(item)
        headers = list(self.grid["columns"])
        for row in rows[1:]:
            padded = (row + [""] * len(headers))[: len(headers)]
            self.grid.insert("", "end", values=padded)

    def _db_process_error(self, option: str = "Selected option") -> None:
        messagebox.showerror("DB Process Error", f"{option} requires the DB process and is not available in this replica.")

    def _noop(self) -> None:
        self._db_process_error()


if __name__ == "__main__":
    SPCReplica().mainloop()
