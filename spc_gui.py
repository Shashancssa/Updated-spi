"""SPC replica GUI with Excel-launching backend action.

Run with: python3 spc_gui.py
"""
from __future__ import annotations

import csv
import ctypes
import json
import os
import platform
import struct
import subprocess
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "SPC Analysis"
CONFIG_FILE = Path(".spc_gui_config.json")
DEFAULT_SPC_EXCEL = Path("spc_backend/SPC_Backend.xlsx")
WINDOW_ICON = Path(".spc_runtime/spc_backend.ico")
WINDOW_APP_ID = "SPC.Analysis"
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
        """Set the same SPC icon for the title bar and Windows taskbar/app switcher."""
        if platform.system() == "Windows":
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(WINDOW_APP_ID)
            icon_path = self._ensure_window_icon_file()
            self.iconbitmap(default=str(icon_path))

        icon = tk.PhotoImage(width=32, height=32)
        icon.put("#1d6fa5", to=(0, 0, 32, 32))
        icon.put("#d7e8fb", to=(3, 3, 29, 29))
        icon.put("#1d6fa5", to=(5, 22, 27, 25))
        icon.put("#27ae60", to=(7, 15, 12, 22))
        icon.put("#2f80ed", to=(14, 10, 19, 22))
        icon.put("#27ae60", to=(21, 6, 26, 22))
        self.iconphoto(True, icon)
        self._icon_image = icon

    def _ensure_window_icon_file(self) -> Path:
        """Create the Windows .ico file at runtime so the PR contains no binary files."""
        if WINDOW_ICON.exists():
            return WINDOW_ICON
        WINDOW_ICON.parent.mkdir(exist_ok=True)
        size = 32
        pixels = []
        for y in range(size):
            for x in range(size):
                color = (0x1D, 0x6F, 0xA5)
                if 3 <= x < 29 and 3 <= y < 29:
                    color = (0xD7, 0xE8, 0xFB)
                if 5 <= x < 27 and 22 <= y < 25:
                    color = (0x1D, 0x6F, 0xA5)
                if 7 <= x < 12 and 15 <= y < 22:
                    color = (0x27, 0xAE, 0x60)
                if 14 <= x < 19 and 10 <= y < 22:
                    color = (0x2F, 0x80, 0xED)
                if 21 <= x < 26 and 6 <= y < 22:
                    color = (0x27, 0xAE, 0x60)
                pixels.append(color)

        bitmap = struct.pack("<IIIHHIIIIII", 40, size, size * 2, 1, 32, 0, size * size * 4, 0, 0, 0, 0)
        for y in range(size - 1, -1, -1):
            for x in range(size):
                red, green, blue = pixels[y * size + x]
                bitmap += bytes([blue, green, red, 255])
        bitmap += bytes(size * 4)

        icon_header = struct.pack("<HHH", 0, 1, 1)
        icon_directory = struct.pack("<BBBBHHII", size, size, 0, 0, 1, 32, len(bitmap), 22)
        WINDOW_ICON.write_bytes(icon_header + icon_directory + bitmap)
        return WINDOW_ICON

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
        tk.Label(bar, text="⚙", bg="#9be7b0", fg="#0b6b2b", font=("Segoe UI", 14), width=3).pack(side="left", padx=(3, 12), pady=3)

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
            ("▦", "PcbView", self.open_pcb_view),
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
        colors = [("green", 0), ("red", 0), ("#00a651", 100), ("blue", 100)]
        for x, label in zip(groups, labels):
            c.create_text(x, bottom + 20, text=label, font=("Segoe UI", 8))
            for j, (color, val) in enumerate(colors):
                bar_h = (bottom - top) * val / 105
                x0 = x - 55 + j * 35
                c.create_rectangle(x0, bottom - bar_h, x0 + 35, bottom, fill=color, outline=color)
                c.create_text(x0 + 17, bottom - bar_h - 16, text=str(val), font=("Segoe UI", 8))
        for k, (name, color) in enumerate((("PassRate", "green"), ("FailRate", "red"), ("FalseAlarmRate", "#00a651"), ("YieldRate", "blue"))):
            y = 52 + k * 18
            c.create_rectangle(right + 20, y - 7, right + 30, y + 3, fill=color)
            c.create_text(right + 38, y, text=name, anchor="w", font=("Segoe UI", 8))

    def _build_status_bar(self) -> None:
        status = tk.Frame(self, height=22, bg="#d7e8fb", highlightthickness=1, highlightbackground="#a7c0df")
        status.pack(fill="x", side="bottom")
        self.status_left = tk.Label(status, text=f"Linked SPC Excel: {self.spc_excel_path}", bg="#d7e8fb", fg="#173d5c", anchor="w")
        self.status_left.pack(side="left", padx=8)
        tk.Label(status, text="SPC Analysis", bg="#d7e8fb", fg="#173d5c", anchor="e").pack(side="right", padx=8)

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
        messagebox.showinfo("SPC Analysis", f"SPC option linked to:\n{self.spc_excel_path}")

    def open_linked_spc_excel(self) -> None:
        """Open the linked workbook directly when the SPC option is clicked."""
        path = self.spc_excel_path
        if not path.exists():
            messagebox.showwarning(
                "SPC Analysis",
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
            messagebox.showinfo("SPC Analysis", f"Linked SPC Excel opened:\n{path}")
        except Exception as exc:  # noqa: BLE001 - display OS integration failures to the operator.
            messagebox.showwarning("SPC Analysis", f"Linked file:\n{path}\n\nCould not open it automatically: {exc}")

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

    def open_pcb_view(self) -> None:
        """Open the PCB View replica window with editable inspection fields."""
        win = tk.Toplevel(self)
        win.title("SPC Analysis - PCB View")
        win.geometry("1366x720")
        win.minsize(1100, 620)
        win.configure(bg="#dfeaf8")
        win.transient(self)

        ribbon = tk.Frame(win, height=70, bg="#c6ddf5", highlightthickness=1, highlightbackground="#94b7d8")
        ribbon.pack(fill="x")
        for icon, label in (("🔍", "ModelList"), ("▤", "ListView/Report"), ("▦", "PcbView"), ("▥", "Histogram"), ("📊", "Defect\nChart"), ("〰", "DefectSPC")):
            item = tk.Frame(ribbon, width=78, height=68, bg="#c6ddf5", highlightbackground="#aac4de", highlightthickness=1)
            item.pack(side="left", padx=1, pady=2)
            item.pack_propagate(False)
            tk.Label(item, text=icon, bg="#c6ddf5", font=("Segoe UI", 18)).pack(pady=(4, 0))
            tk.Label(item, text=label, bg="#c6ddf5", fg="#163d5c", font=("Segoe UI", 8)).pack()

        middle = tk.Frame(win, bg="#dfeaf8")
        middle.pack(fill="both", expand=True)
        left = tk.Frame(middle, width=260, bg="#dce9f8", highlightbackground="#9fbdda", highlightthickness=1)
        left.pack(side="left", fill="y", padx=(2, 3), pady=3)
        left.pack_propagate(False)
        center = tk.Frame(middle, bg="black")
        center.pack(side="left", fill="both", expand=True, pady=3)
        right = tk.Frame(middle, width=225, bg="#eef5ff", highlightbackground="#9fbdda", highlightthickness=1)
        right.pack(side="right", fill="y", padx=(3, 2), pady=3)
        right.pack_propagate(False)

        self._build_pcb_filter_panel(left)
        self._build_pcb_canvas(center)
        self._build_pcb_analysis_panel(right)
        self._build_pcb_data_grid(win)

    def _build_pcb_filter_panel(self, parent: tk.Frame) -> None:
        filter_box = ttk.LabelFrame(parent, text="Data Filter:", style="Panel.TLabelframe")
        filter_box.pack(fill="x", padx=6, pady=4)
        tk.Label(filter_box, text="Filter Type:", bg="#dfeaf8").grid(row=0, column=0, sticky="w", padx=5, pady=3)
        filter_type = ttk.Combobox(filter_box, values=("*", "Volume", "Area", "Height"), width=12)
        filter_type.set("*")
        filter_type.grid(row=0, column=1, padx=4, pady=3)
        tk.Checkbutton(filter_box, text="Analysis", bg="#dfeaf8").grid(row=0, column=2, sticky="w")
        metric = ttk.Combobox(filter_box, values=("Volume", "Area", "Height"), width=9)
        metric.set("Volume")
        metric.grid(row=1, column=0, padx=5, pady=3)
        for row, sign, value in ((1, ">", "0"), (2, "<", "100")):
            op = ttk.Combobox(filter_box, values=(">", "<", "="), width=3)
            op.set(sign)
            op.grid(row=row, column=1, sticky="w", padx=4)
            entry = ttk.Entry(filter_box, width=8)
            entry.insert(0, value)
            entry.grid(row=row, column=1, sticky="e", padx=4)
        ttk.Button(filter_box, text="🔎 QUERY", style="Blue.TButton").grid(row=1, column=2, rowspan=2, padx=8, pady=6, sticky="nsew")

        info = ttk.LabelFrame(parent, text="Panel Info:", style="Panel.TLabelframe")
        info.pack(fill="x", padx=6, pady=4)
        values = (
            ("Board Status:", "RPASS"),
            ("Model Name:", "SPI_TRI_356_R0_(ME420079)_SS"),
            ("Inspect Time:", "2026/06/03 20:48:54"),
            ("Barcode:", "4P540010436261G0006"),
            ("TotalPad:", "2292"),
            ("Station:", "LINE01SPI"),
        )
        for row, (label, value) in enumerate(values):
            tk.Label(info, text=label, bg="#dfeaf8").grid(row=row, column=0, sticky="w", padx=4, pady=3)
            entry = ttk.Entry(info, width=24)
            entry.insert(0, value)
            entry.grid(row=row, column=1, sticky="ew", padx=4, pady=3)

        defect = ttk.LabelFrame(parent, text="Defect", style="Panel.TLabelframe")
        defect.pack(fill="both", expand=True, padx=6, pady=4)
        defect_grid = ttk.Treeview(defect, columns=("Result", "Count"), show="headings", height=6)
        for col in defect_grid["columns"]:
            defect_grid.heading(col, text=col)
            defect_grid.column(col, width=110 if col == "Result" else 55)
        for row in (("Over Height", "4"), ("PASS", "2288"), ("DefectCount", "4")):
            defect_grid.insert("", "end", values=row)
        defect_grid.pack(fill="both", expand=True, padx=4, pady=4)
        self._make_tree_editable(defect_grid)
        buttons = tk.Frame(parent, bg="#dce9f8")
        buttons.pack(fill="x", padx=6, pady=5)
        for text in ("🔍 To Condition", "📗 Report", "▰ Warpage"):
            ttk.Button(buttons, text=text, style="Blue.TButton").pack(side="left", expand=True, fill="x", padx=2)

    def _build_pcb_canvas(self, parent: tk.Frame) -> None:
        canvas = tk.Canvas(parent, bg="black", highlightthickness=0)
        canvas.pack(fill="both", expand=True)

        def draw(_event: tk.Event | None = None) -> None:
            canvas.delete("all")
            width = max(canvas.winfo_width(), 700)
            height = max(canvas.winfo_height(), 430)
            board_w = min(520, width - 220)
            board_h = height - 30
            x0 = (width - board_w) // 2
            y0 = 0
            canvas.create_rectangle(x0, y0, x0 + board_w, y0 + board_h, fill="#008b13", outline="#008b13")
            cell_w = board_w / 3
            cell_h = board_h / 2
            for row in range(2):
                for col in range(3):
                    ox = x0 + col * cell_w + 20
                    oy = y0 + row * cell_h + 20
                    self._draw_board_cell(canvas, ox, oy, cell_w - 35, cell_h - 35)
            canvas.create_text(5, height - 18, anchor="w", text="Current Zoom1", fill="black", font=("Segoe UI", 9, "bold"))
            for i, symbol in enumerate(("↔", "✣", "↖", "↗", "↙", "↘", "🔍")):
                canvas.create_rectangle(95 + i * 28, height - 32, 120 + i * 28, height - 7, fill="#e9f3ff", outline="#7fa8c8")
                canvas.create_text(107 + i * 28, height - 20, text=symbol, fill="#1c3f5d", font=("Segoe UI", 10))

        canvas.bind("<Configure>", draw)
        canvas.after(100, draw)

    def _draw_board_cell(self, canvas: tk.Canvas, x: float, y: float, width: float, height: float) -> None:
        for i in range(10):
            px = x + 8 + (i % 5) * 12
            py = y + 8 + (i // 5) * 12
            canvas.create_rectangle(px, py, px + 4, py + 4, outline="white")
        for i in range(55):
            px = x + 70 + (i * 17 % int(width - 90))
            py = y + 22 + (i * 29 % int(height - 55))
            size = 3 + (i % 3)
            canvas.create_rectangle(px, py, px + size, py + size, outline="white")
        for i in range(5):
            canvas.create_rectangle(x + 16 + i * 24, y + height - 36, x + 28 + i * 24, y + height - 22, outline="white")
        canvas.create_rectangle(x + width / 2 - 18, y + height - 55, x + width / 2 + 18, y + height - 18, outline="white", width=2)
        for i in range(10):
            canvas.create_rectangle(x + width / 2 - 25, y + height - 52 + i * 3, x + width / 2 - 20, y + height - 50 + i * 3, outline="white")
            canvas.create_rectangle(x + width / 2 + 20, y + height - 52 + i * 3, x + width / 2 + 25, y + height - 50 + i * 3, outline="white")

    def _build_pcb_analysis_panel(self, parent: tk.Frame) -> None:
        kind = ttk.LabelFrame(parent, text="Kind", style="Panel.TLabelframe")
        kind.pack(fill="x", padx=6, pady=4)
        for text in ("Volume(%)", "Area(%)", "Height"):
            tk.Radiobutton(kind, text=text, bg="#eef5ff", value=text).pack(side="left", padx=2)
        hist = tk.Canvas(parent, height=135, bg="white", highlightbackground="#808080", highlightthickness=1)
        hist.pack(fill="x", padx=10, pady=6)
        for x in range(15, 210, 12):
            hist.create_rectangle(x, 10, x + 4, 125, fill="yellow", outline="yellow")
        points = [70, 74, 76, 79, 84, 87, 88, 92, 94, 97, 100, 104, 108, 110, 114]
        for i, x in enumerate(points):
            y = 120 - (i * 7 % 85)
            hist.create_line(x, 120, x + 3, y, fill="black", width=3)
        scatter = tk.Canvas(parent, height=110, bg="white", highlightbackground="#808080", highlightthickness=1)
        scatter.pack(fill="x", padx=25, pady=14)
        scatter.create_line(10, 55, 195, 55)
        scatter.create_line(102, 10, 102, 100)
        for i in range(180):
            x = 102 + ((i * 37) % 70) - 35
            y = 55 + ((i * 23) % 46) - 23
            scatter.create_oval(x, y, x + 1, y + 1, fill="black")
        ttk.Combobox(parent, values=("Height CP/CPK",), width=25).pack(fill="x", padx=10, pady=5)
        tk.Checkbutton(parent, text="ColorCadMap", bg="#eef5ff").pack(anchor="w", padx=10)
        ttk.Button(parent, text="📁 Show Image", style="Blue.TButton").pack(padx=45, pady=8, fill="x")
        nav = tk.Frame(parent, bg="#eef5ff")
        nav.pack(fill="x", padx=8, pady=6)
        ttk.Button(nav, text="◀ Prior", style="Blue.TButton", state="disabled").pack(side="left", expand=True, fill="x")
        tk.Label(nav, text="1/1", bg="#eef5ff").pack(side="left", padx=8)
        ttk.Button(nav, text="Next ▶", style="Blue.TButton", state="disabled").pack(side="left", expand=True, fill="x")

    def _build_pcb_data_grid(self, parent: tk.Toplevel) -> None:
        bottom = tk.Frame(parent, height=175, bg="#dce9f8", highlightbackground="#9fbdda", highlightthickness=1)
        bottom.pack(fill="x", side="bottom")
        columns = ("winname", "V%", "A%", "H", "PX", "PY", "CircumRatio", "H%", "Boardsn", "Vhigh(%)", "Vstd(%)", "Vlow(%)")
        data_grid = ttk.Treeview(bottom, columns=columns, show="headings", height=7)
        data_grid.pack(fill="x", padx=45, pady=(8, 0))
        for col in columns:
            data_grid.heading(col, text=col)
            data_grid.column(col, width=90 if col != "Boardsn" else 150)
        rows = (
            ("1", "152.19", "101.88", "5.9764", "-1.1654", "3.9606", "0", "149.38", "4P540010436261G0002", "180", "100", "60"),
            ("1", "124.53", "83.1", "5.9961", "0.2835", "2.0591", "0", "149.85", "4P540010436261G0003", "180", "100", "60"),
            ("2", "146.45", "97.55", "6.0039", "-0.4213", "0.8898", "0", "150.13", "4P540010436261G0005", "180", "100", "60"),
            ("2", "132.7", "86.34", "6.1457", "0.2087", "2.1969", "0", "153.69", "4P540010436261G0001", "180", "100", "60"),
        )
        for row in rows:
            data_grid.insert("", "end", values=row)
        self._make_tree_editable(data_grid)
        summary = tk.Frame(bottom, bg="#dce9f8")
        summary.pack(fill="x", padx=580, pady=5)
        for values in (("AVG=118.65", "AVG=99.50", "AVG=4.77", "AVG=0.02", "AVG=0.46"), ("MAX=152.19", "MAX=113.75", "MAX=6.15", "MAX=5.28", "MAX=6.85"), ("MIN=76.14", "MIN=69.97", "MIN=3.24", "MIN=-3.09", "MIN=-4.39")):
            row_frame = tk.Frame(summary, bg="#dce9f8")
            row_frame.pack(fill="x")
            for value in values:
                entry = ttk.Entry(row_frame, width=14)
                entry.insert(0, value)
                entry.pack(side="left", padx=2, pady=1)

    def _db_process_error(self, option: str = "Selected option") -> None:
        messagebox.showerror("DB Process Error", f"{option} requires the DB process and is not available in this replica.")

    def _noop(self) -> None:
        self._db_process_error()


if __name__ == "__main__":
    SPCReplica().mainloop()
