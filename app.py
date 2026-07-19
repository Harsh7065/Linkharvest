"""
app.py
Modern desktop UI for LinkHarvest (Windows).
Sidebar navigation + swappable pages, grid-based responsive layout.
Run with:  python app.py
"""
import os
import threading
import queue
import subprocess
import webbrowser

import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
from PIL import Image

from downloader import load_sheet, build_download_tasks, run_downloads
from donation import generate_qr_image
from updater import check_for_update
from auth import verify_credentials, start_periodic_recheck, AuthError
from version import __version__
import pdf_extractor as pe
import data_profiler as dp
import sheet_editor as se
import dashboard_builder as db
import ai_assistant as aa
from donut_chart import render_donut
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

DEFAULT_THREADS = 10
DEFAULT_TIMEOUT = 30
DEFAULT_PDF_THREADS = 10

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ICON_PATH = os.path.join(BASE_DIR, "assets", "icon.ico")

SIDEBAR_WIDTH = 230
ACCENT = ("#1f6feb", "#3c8cff")
SIDEBAR_BG = ("gray92", "gray14")
SIDEBAR_BTN_INACTIVE = "transparent"
SIDEBAR_BTN_HOVER = ("gray85", "gray22")

# Each entry here is one row in the sidebar. Add a tuple to this list to
# register a new feature/page without touching the layout code.
NAV_ITEMS = [
    ("downloader", "🔗", "Link Downloader"),
    ("pdf_extractor", "📄", "PDF Extractor"),
    ("data_profiler", "🧹", "Data Profiler"),
    ("sheet_editor", "✏️", "Excel Editor"),
    ("dashboard_builder", "📊", "Dashboard Builder"),
    ("ai_assistant", "🤖", "AI Assistant"),
    ("support", "❤", "Support"),
]


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"LinkHarvest v{__version__}")
        self.geometry("1080x760")
        self.minsize(880, 640)
        try:
            self.iconbitmap(ICON_PATH)
        except Exception:
            pass

        self.log_queue = queue.Queue()
        self.download_thread = None

        self.pdf_log_queue = queue.Queue()
        self.pdf_thread = None

        self.editor_queue = queue.Queue()
        self.editor_thread = None
        self.editor_session = None

        self.dashboard_queue = queue.Queue()
        self.dashboard_thread = None
        self.dashboard_df = None
        self.dashboard_column_info = None
        self.dashboard_plan = None
        self.dashboard_chart_canvas_widget = None
        self.dashboard_summary_text = ""

        self.assistant_queue = queue.Queue()
        self.assistant_thread = None
        self.assistant_image_entries = []

        self.nav_buttons = {}
        self.pages = {}
        self.current_page = None

        self._build_layout()
        self._show_page("downloader")

        self.after(150, self._poll_log_queue)
        self.after(150, self._poll_pdf_log_queue)
        self.after(150, self._poll_editor_queue)
        self.after(150, self._poll_dashboard_queue)
        self.after(150, self._poll_assistant_queue)
        threading.Thread(target=self._check_update_then_login, daemon=True).start()

    # ==================================================================
    # Top-level responsive layout: sidebar (fixed) + content (expands)
    # ==================================================================
    def _build_layout(self):
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()

        # Content host: every page frame lives here, stacked via grid,
        # and we raise the active one. This avoids rebuilding widgets
        # every time the user switches pages.
        self.content_host = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.content_host.grid(row=0, column=1, sticky="nsew")
        self.content_host.grid_columnconfigure(0, weight=1)
        self.content_host.grid_rowconfigure(0, weight=1)

        self.pages["downloader"] = self._build_downloader_page(self.content_host)
        self.pages["pdf_extractor"] = self._build_pdf_extractor_page(self.content_host)
        self.pages["data_profiler"] = self._build_data_profiler_page(self.content_host)
        self.pages["sheet_editor"] = self._build_sheet_editor_page(self.content_host)
        self.pages["dashboard_builder"] = self._build_dashboard_builder_page(self.content_host)
        self.pages["ai_assistant"] = self._build_ai_assistant_page(self.content_host)
        self.pages["support"] = self._build_support_page(self.content_host)

        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")

    def _build_sidebar(self):
        sidebar = ctk.CTkFrame(self, width=SIDEBAR_WIDTH, corner_radius=0, fg_color=SIDEBAR_BG)
        sidebar.grid(row=0, column=0, sticky="nsw")
        sidebar.grid_propagate(False)

        # Branding
        brand_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        brand_frame.pack(fill="x", pady=(26, 6), padx=18)
        ctk.CTkLabel(brand_frame, text="🔗 LinkHarvest",
                     font=ctk.CTkFont(size=19, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(brand_frame, text=f"v{__version__}",
                     font=ctk.CTkFont(size=11), text_color="gray").pack(anchor="w", pady=(2, 0))

        ctk.CTkFrame(sidebar, height=1, fg_color=("gray80", "gray25")).pack(fill="x", padx=18, pady=(14, 14))

        ctk.CTkLabel(sidebar, text="FEATURES", font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="gray", anchor="w").pack(fill="x", padx=20, pady=(0, 6))

        # Nav buttons — driven by NAV_ITEMS, so adding a future feature
        # is just adding one line to that list.
        for key, icon, label in NAV_ITEMS:
            btn = ctk.CTkButton(
                sidebar, text=f"{icon}   {label}", anchor="w", height=42,
                corner_radius=8, font=ctk.CTkFont(size=14),
                fg_color=SIDEBAR_BTN_INACTIVE, hover_color=SIDEBAR_BTN_HOVER,
                text_color=("gray10", "gray90"),
                command=lambda k=key: self._show_page(k)
            )
            btn.pack(fill="x", padx=14, pady=3)
            self.nav_buttons[key] = btn

        # Footer
        footer = ctk.CTkFrame(sidebar, fg_color="transparent")
        footer.pack(side="bottom", fill="x", padx=18, pady=18)
        ctk.CTkLabel(footer, text="More features coming soon", font=ctk.CTkFont(size=10),
                     text_color="gray", wraplength=SIDEBAR_WIDTH - 40, justify="left").pack(anchor="w")

    def _show_page(self, key):
        for k, btn in self.nav_buttons.items():
            if k == key:
                btn.configure(fg_color=ACCENT, text_color="white")
            else:
                btn.configure(fg_color=SIDEBAR_BTN_INACTIVE, text_color=("gray10", "gray90"))
        self.pages[key].tkraise()
        self.current_page = key

    # Small helper so every page gets a consistent scrollable, padded body
    # (keeps things usable if the window gets short/narrow).
    def _new_page(self, parent, title, subtitle=None):
        page = ctk.CTkFrame(parent, corner_radius=0, fg_color="transparent")
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(page, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=28, pady=(24, 4))
        ctk.CTkLabel(header, text=title, font=ctk.CTkFont(size=22, weight="bold")).pack(anchor="w")
        if subtitle:
            ctk.CTkLabel(header, text=subtitle, font=ctk.CTkFont(size=13),
                         text_color="gray").pack(anchor="w", pady=(2, 0))

        body = ctk.CTkScrollableFrame(page, corner_radius=0, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=20, pady=(4, 20))
        body.grid_columnconfigure(0, weight=1)
        return page, body

    # ==================================================================
    # PAGE: Link Downloader
    # ==================================================================
    def _build_downloader_page(self, parent):
        page, body = self._new_page(
            parent, "Link Downloader",
            "Scan an Excel sheet for links and download images/videos/audio in parallel")

        # File section
        file_frame = ctk.CTkFrame(body, corner_radius=12)
        file_frame.pack(fill="x", pady=10)

        self.excel_path = self._path_row(file_frame, "Excel file", self._browse_excel)
        self.excel_path.insert(0, r"C:\Users\Harsh\Documents\links.xlsx")

        self.sheet_name = self._entry_row(file_frame, "Sheet name", "Sheet4")

        self.folder_path = self._path_row(file_frame, "Save folder", self._browse_folder)
        self.folder_path.insert(0, r"C:\Users\Harsh\Downloads\HarvestedFiles")

        # Row / column range section
        range_frame = ctk.CTkFrame(body, corner_radius=12)
        range_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(range_frame, text="Rows & Columns to scan",
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(12, 6))

        row_line = ctk.CTkFrame(range_frame, fg_color="transparent")
        row_line.pack(fill="x", padx=16, pady=6)
        ctk.CTkLabel(row_line, text="Rows:  From").pack(side="left")
        self.row_from = ctk.CTkEntry(row_line, width=70, placeholder_text="2")
        self.row_from.pack(side="left", padx=6)
        ctk.CTkLabel(row_line, text="To").pack(side="left")
        self.row_to = ctk.CTkEntry(row_line, width=70, placeholder_text="last")
        self.row_to.pack(side="left", padx=6)
        ctk.CTkLabel(row_line, text="(leave 'To' empty for the last row)",
                     font=ctk.CTkFont(size=11), text_color="gray").pack(side="left", padx=6)

        col_line = ctk.CTkFrame(range_frame, fg_color="transparent")
        col_line.pack(fill="x", padx=16, pady=(6, 14))
        ctk.CTkLabel(col_line, text="Columns:  From").pack(side="left")
        self.col_from = ctk.CTkEntry(col_line, width=70, placeholder_text="1")
        self.col_from.pack(side="left", padx=6)
        ctk.CTkLabel(col_line, text="To").pack(side="left")
        self.col_to = ctk.CTkEntry(col_line, width=70, placeholder_text="last")
        self.col_to.pack(side="left", padx=6)
        ctk.CTkLabel(col_line, text="(e.g. From=5 To=5 scans only column E)",
                     font=ctk.CTkFont(size=11), text_color="gray").pack(side="left", padx=6)

        # Advanced section
        adv_frame = ctk.CTkFrame(body, corner_radius=12)
        adv_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(adv_frame, text="Advanced (optional)",
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(12, 6))

        adv_line = ctk.CTkFrame(adv_frame, fg_color="transparent")
        adv_line.pack(fill="x", padx=16, pady=(0, 6))
        ctk.CTkLabel(adv_line, text="Threads").pack(side="left")
        self.threads_entry = ctk.CTkEntry(adv_line, width=60)
        self.threads_entry.insert(0, str(DEFAULT_THREADS))
        self.threads_entry.pack(side="left", padx=6)
        ctk.CTkLabel(adv_line, text="Timeout (sec)").pack(side="left", padx=(16, 0))
        self.timeout_entry = ctk.CTkEntry(adv_line, width=60)
        self.timeout_entry.insert(0, str(DEFAULT_TIMEOUT))
        self.timeout_entry.pack(side="left", padx=6)

        self.adv_warning = ctk.CTkLabel(adv_frame, text="", text_color="#e0a020",
                                         font=ctk.CTkFont(size=11), wraplength=650, justify="left")
        self.adv_warning.pack(anchor="w", padx=16, pady=(2, 12))
        self.threads_entry.bind("<KeyRelease>", self._check_advanced_changed)
        self.timeout_entry.bind("<KeyRelease>", self._check_advanced_changed)

        # Progress
        self.progress_bar = ctk.CTkProgressBar(body)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", pady=(16, 4))
        self.progress_label = ctk.CTkLabel(body, text="0%  (0 / 0)", font=ctk.CTkFont(size=12))
        self.progress_label.pack()

        # Log box
        self.log_box = ctk.CTkTextbox(body, height=150, corner_radius=8)
        self.log_box.pack(fill="both", expand=True, pady=10)

        # Action Button
        self.start_btn = ctk.CTkButton(body, text="Start Download", height=44,
                                        font=ctk.CTkFont(size=15, weight="bold"),
                                        command=self._start_download)
        self.start_btn.pack(pady=(8, 12), fill="x")

        return page

    def _path_row(self, parent, label, browse_cmd):
        line = ctk.CTkFrame(parent, fg_color="transparent")
        line.pack(fill="x", padx=16, pady=(10, 6))
        ctk.CTkLabel(line, text=label, width=90, anchor="w").pack(side="left")
        entry = ctk.CTkEntry(line)
        entry.pack(side="left", fill="x", expand=True, padx=6)
        ctk.CTkButton(line, text="Browse", width=80, command=lambda: browse_cmd(entry)).pack(side="left")
        return entry

    def _entry_row(self, parent, label, default=""):
        line = ctk.CTkFrame(parent, fg_color="transparent")
        line.pack(fill="x", padx=16, pady=6)
        ctk.CTkLabel(line, text=label, width=90, anchor="w").pack(side="left")
        entry = ctk.CTkEntry(line)
        entry.insert(0, default)
        entry.pack(side="left", fill="x", expand=True, padx=6)
        return entry

    def _browse_excel(self, entry):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)

    def _browse_folder(self, entry):
        path = filedialog.askdirectory()
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)

    def _check_advanced_changed(self, _event=None):
        changed = (self.threads_entry.get() != str(DEFAULT_THREADS) or
                   self.timeout_entry.get() != str(DEFAULT_TIMEOUT))
        if changed:
            self.adv_warning.configure(
                text="⚠ Default settings changed — with more threads or a shorter timeout, "
                     "some links may fail or remain un-downloaded.")
        else:
            self.adv_warning.configure(text="")

    def _start_download(self):
        if self.download_thread and self.download_thread.is_alive():
            messagebox.showinfo("Busy", "A download is already running.")
            return

        excel_path = self.excel_path.get().strip()
        sheet_name = self.sheet_name.get().strip()
        folder_path = self.folder_path.get().strip()

        if not excel_path or not os.path.isfile(excel_path):
            messagebox.showerror("Missing info", "Please select a valid Excel file.")
            return
        if not sheet_name:
            messagebox.showerror("Missing info", "Please enter a sheet name.")
            return
        if not folder_path:
            messagebox.showerror("Missing info", "Please select a save folder.")
            return
        os.makedirs(folder_path, exist_ok=True)

        try:
            row_from = int(self.row_from.get() or 2)
            row_to = int(self.row_to.get()) if self.row_to.get().strip() else None
            col_from = int(self.col_from.get() or 1)
            col_to = int(self.col_to.get()) if self.col_to.get().strip() else None
            threads = int(self.threads_entry.get() or DEFAULT_THREADS)
            timeout = int(self.timeout_entry.get() or DEFAULT_TIMEOUT)
        except ValueError:
            messagebox.showerror("Invalid input", "Rows, columns, threads and timeout must be numbers.")
            return

        self.start_btn.configure(state="disabled", text="Downloading...")
        self.progress_bar.set(0)
        self.progress_label.configure(text="0%  (0 / 0)")
        self.log_box.delete("1.0", "end")

        self.download_thread = threading.Thread(
            target=self._run_download_job,
            args=(excel_path, sheet_name, folder_path, row_from, row_to,
                  col_from, col_to, threads, timeout),
            daemon=True)
        self.download_thread.start()

    def _run_download_job(self, excel_path, sheet_name, folder_path, row_from,
                           row_to, col_from, col_to, threads, timeout):
        try:
            ws = load_sheet(excel_path, sheet_name)
            tasks = build_download_tasks(ws, row_from, row_to or ws.max_row,
                                          col_from, col_to or ws.max_column, folder_path)
        except Exception as e:
            self.log_queue.put(("error", f"Failed to read workbook: {e}"))
            self.log_queue.put(("done", None))
            return

        if not tasks:
            self.log_queue.put(("error", "No links found in the selected rows/columns."))
            self.log_queue.put(("done", None))
            return

        self.log_queue.put(("info", f"Found {len(tasks)} link(s). Downloading with {threads} thread(s)..."))

        def progress_cb(done, total):
            self.log_queue.put(("progress", (done, total)))

        def log_cb(msg):
            self.log_queue.put(("warn", msg))

        ok, failed = run_downloads(tasks, threads, timeout, progress_cb, log_cb)
        summary = f"Done. {ok} succeeded, {failed} failed."
        if failed:
            summary += " Some links may remain un-downloaded — see the log above."
        self.log_queue.put(("info", summary))
        self.log_queue.put(("done", None))

    def _poll_log_queue(self):
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()
                if kind == "progress":
                    done, total = payload
                    pct = done / total if total else 0
                    self.progress_bar.set(pct)
                    self.progress_label.configure(text=f"{int(pct * 100)}%  ({done} / {total})")
                elif kind in ("info", "warn", "error"):
                    self.log_box.insert("end", f"{payload}\n")
                    self.log_box.see("end")
                elif kind == "done":
                    self.start_btn.configure(state="normal", text="Start Download")
        except queue.Empty:
            pass
        self.after(150, self._poll_log_queue)

    # ==================================================================
    # PAGE: PDF Intelligent Extractor
    # ==================================================================
    def _build_pdf_extractor_page(self, parent):
        page, body = self._new_page(
            parent, "PDF Extractor",
            "Describe what to pull from a folder of PDFs — accurate, multithreaded, your choice of AI engine")

        # --- Configuration: AI provider + API key ---
        cfg_frame = ctk.CTkFrame(body, corner_radius=12)
        cfg_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(cfg_frame, text="Configuration",
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(12, 6))

        provider_line = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        provider_line.pack(fill="x", padx=16, pady=(0, 4))
        ctk.CTkLabel(provider_line, text="AI Engine", width=110, anchor="w").pack(side="left")
        self.provider_selector = ctk.CTkSegmentedButton(
            provider_line, values=["Gemini (Flash) — Free tier", "OpenAI (gpt-5)"],
            command=self._on_provider_change)
        self.provider_selector.set("Gemini (Flash) — Free tier")
        self.provider_selector.pack(side="left", fill="x", expand=True, padx=6)
        self._pdf_provider = pe.PROVIDER_GEMINI  # internal state, kept in sync with the selector

        get_key_line = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        get_key_line.pack(fill="x", padx=16, pady=(0, 10))
        ctk.CTkLabel(get_key_line, text="", width=110).pack(side="left")
        self.get_key_btn = ctk.CTkButton(
            get_key_line, text="🔗 Get a free Gemini API key", width=220, height=26,
            font=ctk.CTkFont(size=11), fg_color="transparent", border_width=1,
            text_color=("gray20", "gray85"),
            command=self._open_key_signup_page)
        self.get_key_btn.pack(side="left")

        key_line = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        key_line.pack(fill="x", padx=16, pady=(0, 6))
        self.api_key_field_label = ctk.CTkLabel(key_line, text="Gemini API Key", width=110, anchor="w")
        self.api_key_field_label.pack(side="left")
        self.api_key_entry = ctk.CTkEntry(key_line, show="*")
        self.api_key_entry.pack(side="left", fill="x", expand=True, padx=6)
        try:
            existing_key = pe.load_api_key(self._pdf_provider)
            if existing_key:
                self.api_key_entry.insert(0, existing_key)
        except Exception:
            pass
        ctk.CTkButton(key_line, text="Save", width=70, command=self._save_api_key).pack(side="left")

        self.api_key_status = ctk.CTkLabel(cfg_frame, text="Key is stored locally in a .env file, never uploaded anywhere.",
                                            font=ctk.CTkFont(size=11), text_color="gray")
        self.api_key_status.pack(anchor="w", padx=16, pady=(0, 12))

        model_line = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        model_line.pack(fill="x", padx=16, pady=(0, 4))
        ctk.CTkLabel(model_line, text="Model", width=110, anchor="w").pack(side="left")
        self.model_combo = ctk.CTkComboBox(
            model_line, values=pe.SUGGESTED_MODELS[self._pdf_provider])
        self.model_combo.set(pe.DEFAULT_GEMINI_MODEL)
        self.model_combo.pack(side="left", fill="x", expand=True, padx=6)
        ctk.CTkLabel(model_line, text="", width=70).pack(side="left")  # align with Save button above

        ctk.CTkLabel(cfg_frame,
                     text="Pick from the list or type any model name (e.g. gemini-flash-lite-latest for lower cost/faster).",
                     font=ctk.CTkFont(size=11), text_color="gray").pack(anchor="w", padx=16, pady=(0, 12))

        self.provider_warning = ctk.CTkLabel(cfg_frame, text="", text_color="#e0a020",
                                              font=ctk.CTkFont(size=11), wraplength=650, justify="left")
        self.provider_warning.pack(anchor="w", padx=16, pady=(0, 10))
        self._refresh_provider_availability_warning()

        # --- Step 1: target + source ---
        step1_frame = ctk.CTkFrame(body, corner_radius=12)
        step1_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(step1_frame, text="Step 1: Define Target & Source",
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(12, 6))

        ctk.CTkLabel(step1_frame, text="What should be extracted? (natural language, one or more fields)",
                     font=ctk.CTkFont(size=11), text_color="gray").pack(anchor="w", padx=16)
        self.pdf_instructions = ctk.CTkTextbox(step1_frame, height=80, corner_radius=8)
        self.pdf_instructions.pack(fill="x", padx=16, pady=(4, 12))
        self.pdf_instructions.insert(
            "1.0", "Invoice number, invoice date, vendor name, total amount, payment due date")

        self.pdf_source_folder = self._path_row(step1_frame, "Source Folder", self._browse_pdf_source)

        # --- Step 2: output ---
        step2_frame = ctk.CTkFrame(body, corner_radius=12)
        step2_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(step2_frame, text="Step 2: Output Destination",
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(12, 6))

        self.pdf_output_excel = self._path_row(step2_frame, "Output Excel", self._browse_pdf_output)
        self.pdf_sheet_name = self._entry_row(step2_frame, "Sheet Name", "extracted_results")

        # --- Advanced: thread slider ---
        pdf_adv_frame = ctk.CTkFrame(body, corner_radius=12)
        pdf_adv_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(pdf_adv_frame, text="Advanced (optional)",
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(12, 6))

        slider_line = ctk.CTkFrame(pdf_adv_frame, fg_color="transparent")
        slider_line.pack(fill="x", padx=16, pady=(0, 12))
        self.pdf_threads_label = ctk.CTkLabel(slider_line, text=f"Concurrent Threads: {DEFAULT_PDF_THREADS}")
        self.pdf_threads_label.pack(anchor="w")
        self.pdf_threads_slider = ctk.CTkSlider(slider_line, from_=1, to=20, number_of_steps=19,
                                                 command=self._on_pdf_threads_change)
        self.pdf_threads_slider.set(DEFAULT_PDF_THREADS)
        self.pdf_threads_slider.pack(fill="x", pady=(6, 0))

        # --- Progress + start ---
        self.pdf_progress_bar = ctk.CTkProgressBar(body)
        self.pdf_progress_bar.set(0)
        self.pdf_progress_bar.pack(fill="x", pady=(16, 4))
        self.pdf_progress_label = ctk.CTkLabel(body, text="0%  (0 / 0)", font=ctk.CTkFont(size=12))
        self.pdf_progress_label.pack()

        self.pdf_log_box = ctk.CTkTextbox(body, height=150, corner_radius=8)
        self.pdf_log_box.pack(fill="both", expand=True, pady=10)

        self.pdf_start_btn = ctk.CTkButton(body, text="Start Extraction", height=44,
                                            font=ctk.CTkFont(size=15, weight="bold"),
                                            command=self._start_pdf_extraction)
        self.pdf_start_btn.pack(pady=(8, 12), fill="x")

        return page

    def _browse_pdf_source(self, entry):
        path = filedialog.askdirectory()
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)

    def _browse_pdf_output(self, entry):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            confirmoverwrite=False)  # we append, not overwrite, so don't scare the user
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)

    def _on_pdf_threads_change(self, value):
        self.pdf_threads_label.configure(text=f"Concurrent Threads: {int(value)}")

    def _on_provider_change(self, selection):
        self._pdf_provider = (pe.PROVIDER_OPENAI if selection.startswith("OpenAI")
                               else pe.PROVIDER_GEMINI)
        label = "OpenAI API Key" if self._pdf_provider == pe.PROVIDER_OPENAI else "Gemini API Key"
        self.api_key_field_label.configure(text=label)

        # Swap in whichever key is already saved for the newly selected provider.
        self.api_key_entry.delete(0, "end")
        try:
            existing_key = pe.load_api_key(self._pdf_provider)
            if existing_key:
                self.api_key_entry.insert(0, existing_key)
        except Exception:
            pass
        self.api_key_status.configure(
            text="Key is stored locally in a .env file, never uploaded anywhere.",
            text_color="gray")

        # Swap the model dropdown's suggestions + default for the new provider.
        default_model = (pe.DEFAULT_OPENAI_MODEL if self._pdf_provider == pe.PROVIDER_OPENAI
                          else pe.DEFAULT_GEMINI_MODEL)
        self.model_combo.configure(values=pe.SUGGESTED_MODELS[self._pdf_provider])
        self.model_combo.set(default_model)

        key_label = "OpenAI" if self._pdf_provider == pe.PROVIDER_OPENAI else "Gemini"
        free_note = " (free)" if self._pdf_provider == pe.PROVIDER_GEMINI else ""
        self.get_key_btn.configure(text=f"🔗 Get a{free_note} {key_label} API key")

        self._refresh_provider_availability_warning()

    def _open_key_signup_page(self):
        url = ("https://aistudio.google.com/apikey" if self._pdf_provider == pe.PROVIDER_GEMINI
               else "https://platform.openai.com/api-keys")
        webbrowser.open(url)

    def _refresh_provider_availability_warning(self):
        if pe.is_provider_available(self._pdf_provider):
            self.provider_warning.configure(text="")
            return
        pkg = "openai" if self._pdf_provider == pe.PROVIDER_OPENAI else "google-generativeai"
        self.provider_warning.configure(
            text=f"⚠ The '{pkg}' package isn't installed. Run: pip install {pkg}")

    def _save_api_key(self):
        key = self.api_key_entry.get().strip()
        if not key:
            messagebox.showerror("Missing key", "Please enter an API key first.")
            return
        try:
            pe.save_api_key(key, provider=self._pdf_provider)
            self.api_key_status.configure(text="✓ Key saved to .env", text_color="#3ca34d")
        except Exception as e:
            messagebox.showerror("Couldn't save key", str(e))

    def _start_pdf_extraction(self):
        if self.pdf_thread and self.pdf_thread.is_alive():
            messagebox.showinfo("Busy", "An extraction is already running.")
            return

        provider = self._pdf_provider
        if not pe.is_provider_available(provider):
            pkg = "openai" if provider == pe.PROVIDER_OPENAI else "google-generativeai"
            messagebox.showerror("Missing dependency",
                                  f"The '{pkg}' package isn't installed.\nRun: pip install {pkg}")
            return

        api_key = self.api_key_entry.get().strip()
        model_name = self.model_combo.get().strip()
        instructions = self.pdf_instructions.get("1.0", "end").strip()
        source_folder = self.pdf_source_folder.get().strip()
        output_excel = self.pdf_output_excel.get().strip()
        sheet_name = self.pdf_sheet_name.get().strip()
        threads = int(self.pdf_threads_slider.get())

        if not api_key:
            messagebox.showerror("Missing info", f"Please enter and save your {provider.title()} API key.")
            return
        if not instructions:
            messagebox.showerror("Missing info", "Please describe what to extract.")
            return
        if not source_folder or not os.path.isdir(source_folder):
            messagebox.showerror("Missing info", "Please select a valid source folder.")
            return
        if not output_excel:
            messagebox.showerror("Missing info", "Please choose an output Excel file path.")
            return
        if not sheet_name:
            messagebox.showerror("Missing info", "Please enter a sheet/workbook name.")
            return

        pdfs = pe.list_pdfs(source_folder)
        if not pdfs:
            messagebox.showerror("No PDFs found", "No .pdf files were found in that folder.")
            return

        self.pdf_start_btn.configure(state="disabled", text="Extracting...")
        self.pdf_progress_bar.set(0)
        self.pdf_progress_label.configure(text="0%  (0 / 0)")
        self.pdf_log_box.delete("1.0", "end")

        self.pdf_thread = threading.Thread(
            target=self._run_pdf_extraction_job,
            args=(pdfs, instructions, provider, api_key, model_name, threads, output_excel, sheet_name),
            daemon=True)
        self.pdf_thread.start()

    def _run_pdf_extraction_job(self, pdfs, instructions, provider, api_key, model_name, threads, output_excel, sheet_name):
        engine_name = "OpenAI" if provider == pe.PROVIDER_OPENAI else "Gemini"
        self.pdf_log_queue.put(("info", f"Found {len(pdfs)} PDF(s). Extracting with {engine_name} ({model_name}) using {threads} thread(s)..."))

        def progress_cb(done, total):
            self.pdf_log_queue.put(("progress", (done, total)))

        def log_cb(msg):
            self.pdf_log_queue.put(("warn", msg))

        try:
            results = pe.run_extraction(pdfs, instructions, provider, api_key, threads, progress_cb, log_cb,
                                         model_name=model_name)
        except pe.AuthError as e:
            self.pdf_log_queue.put(("error", f"Authentication failed: {e}"))
            self.pdf_log_queue.put(("done", None))
            return
        except Exception as e:
            self.pdf_log_queue.put(("error", f"Extraction failed: {e}"))
            self.pdf_log_queue.put(("done", None))
            return

        try:
            pe.compile_to_excel(results, output_excel, sheet_name)
        except Exception as e:
            self.pdf_log_queue.put(("error", f"Extraction finished but writing Excel failed: {e}"))
            self.pdf_log_queue.put(("done", None))
            return

        ok = sum(1 for r in results if "_error" not in r)
        failed = len(results) - ok
        summary = f"Done. {ok} succeeded, {failed} failed. Saved to {output_excel}"
        self.pdf_log_queue.put(("info", summary))
        self.pdf_log_queue.put(("success", output_excel))
        self.pdf_log_queue.put(("done", None))

    def _poll_pdf_log_queue(self):
        try:
            while True:
                kind, payload = self.pdf_log_queue.get_nowait()
                if kind == "progress":
                    done, total = payload
                    pct = done / total if total else 0
                    self.pdf_progress_bar.set(pct)
                    self.pdf_progress_label.configure(text=f"{int(pct * 100)}%  ({done} / {total})")
                elif kind in ("info", "warn", "error"):
                    self.pdf_log_box.insert("end", f"{payload}\n")
                    self.pdf_log_box.see("end")
                elif kind == "success":
                    self._on_pdf_extraction_success(payload)
                elif kind == "done":
                    self.pdf_start_btn.configure(state="normal", text="Start Extraction")
        except queue.Empty:
            pass
        self.after(150, self._poll_pdf_log_queue)

    def _on_pdf_extraction_success(self, output_excel_path):
        messagebox.showinfo("Extraction complete",
                             f"Extraction finished.\nResults saved to:\n{output_excel_path}")
        try:
            folder = os.path.dirname(os.path.abspath(output_excel_path))
            if os.name == "nt":
                os.startfile(folder)
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception:
            pass  # opening the folder is a convenience, never fatal

    # ==================================================================
    # PAGE: Data Profiler
    # ==================================================================
    def _build_data_profiler_page(self, parent):
        page, body = self._new_page(
            parent, "Data Profiler",
            "Load a CSV/Excel file, detect data-quality issues, and export a cleaned version")

        self.profiler_df = None
        self.profiler_results = None
        self.profiler_chart_canvas_widget = None

        # --- Load data ---
        load_frame = ctk.CTkFrame(body, corner_radius=12)
        load_frame.pack(fill="x", pady=10)
        self.profiler_input_path = self._path_row(load_frame, "Data file", self._browse_profiler_input)

        analyze_line = ctk.CTkFrame(load_frame, fg_color="transparent")
        analyze_line.pack(fill="x", padx=16, pady=(0, 14))
        self.profiler_analyze_btn = ctk.CTkButton(analyze_line, text="Analyze", height=36,
                                                    command=self._run_profiler_analysis)
        self.profiler_analyze_btn.pack(side="left")
        self.profiler_status_label = ctk.CTkLabel(analyze_line, text="Accepts .csv, .xlsx, .xls",
                                                    font=ctk.CTkFont(size=11), text_color="gray")
        self.profiler_status_label.pack(side="left", padx=12)

        # --- KPI cards ---
        kpi_frame = ctk.CTkFrame(body, corner_radius=12)
        kpi_frame.pack(fill="x", pady=10)
        kpi_row = ctk.CTkFrame(kpi_frame, fg_color="transparent")
        kpi_row.pack(fill="x", padx=16, pady=16)
        self.profiler_kpi_labels = {}
        for i, (key, label) in enumerate([
            ("health_pct", "Health Score"),
            ("total_records", "Total Records"),
            ("total_columns", "Total Columns"),
            ("total_anomalies", "Total Anomalies"),
        ]):
            kpi_row.grid_columnconfigure(i, weight=1)
            card = ctk.CTkFrame(kpi_row, corner_radius=10, fg_color=("gray88", "gray20"))
            card.grid(row=0, column=i, sticky="nsew", padx=6)
            value_lbl = ctk.CTkLabel(card, text="—", font=ctk.CTkFont(size=24, weight="bold"))
            value_lbl.pack(pady=(14, 2))
            ctk.CTkLabel(card, text=label, font=ctk.CTkFont(size=11), text_color="gray").pack(pady=(0, 14))
            self.profiler_kpi_labels[key] = value_lbl

        # --- Filters (left) + donut chart (right) ---
        split_frame = ctk.CTkFrame(body, fg_color="transparent")
        split_frame.pack(fill="both", expand=True, pady=10)
        split_frame.grid_columnconfigure(0, weight=1)
        split_frame.grid_columnconfigure(1, weight=1)
        split_frame.grid_rowconfigure(0, weight=1)

        filters_frame = ctk.CTkFrame(split_frame, corner_radius=12)
        filters_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        ctk.CTkLabel(filters_frame, text="Anomalies to fix", font=ctk.CTkFont(weight="bold", size=14)
                     ).pack(pady=(12, 8), padx=16, anchor="w")

        self.profiler_filter_vars = {}
        self.profiler_checkboxes = {}
        for key, label in dp.ANOMALY_LABELS.items():
            var = ctk.BooleanVar(value=False)
            cb = ctk.CTkCheckBox(filters_frame, text=f"{label} (— cases)", variable=var,
                                  command=self._on_profiler_filter_toggle, state="disabled")
            cb.pack(anchor="w", padx=16, pady=6)
            self.profiler_filter_vars[key] = var
            self.profiler_checkboxes[key] = cb

        ctk.CTkFrame(filters_frame, height=1, fg_color=("gray80", "gray30")).pack(fill="x", padx=16, pady=(4, 8))

        self.profiler_fill_mean_var = ctk.BooleanVar(value=False)
        self.profiler_fill_mean_cb = ctk.CTkCheckBox(
            filters_frame, text="Fill missing values with column mean (— cases, numeric only)",
            variable=self.profiler_fill_mean_var, command=self._on_profiler_filter_toggle,
            state="disabled")
        self.profiler_fill_mean_cb.pack(anchor="w", padx=16, pady=6)
        ctk.CTkLabel(filters_frame,
                     text="Only applies to numeric columns — a text column has no valid "
                          "\"average\" to fall back on, so those stay untouched by this option.",
                     font=ctk.CTkFont(size=10), text_color="gray", wraplength=280,
                     justify="left").pack(anchor="w", padx=16, pady=(0, 8))

        ctk.CTkLabel(filters_frame,
                     text="Note: Mixed Data Types is shown for review but is never "
                          "auto-fixed — it's ambiguous, so those cells are left as-is.",
                     font=ctk.CTkFont(size=10), text_color="gray", wraplength=280,
                     justify="left").pack(anchor="w", padx=16, pady=(8, 16))

        chart_frame = ctk.CTkFrame(split_frame, corner_radius=12)
        chart_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        ctk.CTkLabel(chart_frame, text="Anomaly Breakdown", font=ctk.CTkFont(weight="bold", size=14)
                     ).pack(pady=(12, 8))
        self.profiler_chart_host = ctk.CTkFrame(chart_frame, fg_color="transparent")
        self.profiler_chart_host.pack(fill="both", expand=True, padx=8, pady=(0, 12))
        ctk.CTkLabel(self.profiler_chart_host, text="Analyze a file to see the breakdown",
                     font=ctk.CTkFont(size=12), text_color="gray").pack(expand=True)

        # --- Export ---
        export_frame = ctk.CTkFrame(body, corner_radius=12)
        export_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(export_frame, text="Export Cleaned Data", font=ctk.CTkFont(weight="bold", size=14)
                     ).pack(pady=(12, 6))
        self.profiler_clean_btn = ctk.CTkButton(export_frame, text="Clean & Export", height=44,
                                                  font=ctk.CTkFont(size=15, weight="bold"),
                                                  command=self._run_profiler_clean_export, state="disabled")
        self.profiler_clean_btn.pack(pady=(0, 16), padx=16, fill="x")

        return page

    def _browse_profiler_input(self, entry):
        path = filedialog.askopenfilename(filetypes=[("Data files", "*.csv *.xlsx *.xls")])
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)

    def _run_profiler_analysis(self):
        filepath = self.profiler_input_path.get().strip()
        if not filepath:
            messagebox.showerror("Missing info", "Please select a CSV or Excel file first.")
            return
        try:
            df = dp.load_data(filepath)
        except ValueError as e:
            messagebox.showerror("Couldn't load file", str(e))
            return

        self.profiler_df = df
        self.profiler_results = dp.analyze_data(df)
        r = self.profiler_results

        self.profiler_kpi_labels["health_pct"].configure(text=f"{r['health_pct']}%")
        self.profiler_kpi_labels["total_records"].configure(text=str(r["total_records"]))
        self.profiler_kpi_labels["total_columns"].configure(text=str(r["total_columns"]))
        self.profiler_kpi_labels["total_anomalies"].configure(text=str(r["total_anomalies"]))

        for key, label in dp.ANOMALY_LABELS.items():
            count = r.get(key, 0)
            cb = self.profiler_checkboxes[key]
            cb.configure(text=f"{label} ({count} cases)")
            if count > 0:
                cb.configure(state="normal")
                self.profiler_filter_vars[key].set(True)
            else:
                cb.configure(state="disabled")
                self.profiler_filter_vars[key].set(False)

        numeric_missing = r.get("missing_values_numeric", 0)
        self.profiler_fill_mean_cb.configure(
            text=f"Fill missing values with column mean ({numeric_missing} cases, numeric only)")
        if numeric_missing > 0:
            self.profiler_fill_mean_cb.configure(state="normal")
            self.profiler_fill_mean_var.set(True)
        else:
            self.profiler_fill_mean_cb.configure(state="disabled")
            self.profiler_fill_mean_var.set(False)

        self.profiler_status_label.configure(
            text=f"Loaded {r['total_records']} rows, {r['total_columns']} columns.",
            text_color="gray")
        self.profiler_clean_btn.configure(state="normal")

        self._refresh_profiler_chart()

    def _on_profiler_filter_toggle(self):
        self._refresh_profiler_chart()

    def _refresh_profiler_chart(self):
        if not self.profiler_results:
            return
        for widget in self.profiler_chart_host.winfo_children():
            widget.destroy()
        self.profiler_chart_canvas_widget = None

        active_filters = {k: v.get() for k, v in self.profiler_filter_vars.items()}
        fig = render_donut(self.profiler_results, active_filters)
        canvas = FigureCanvasTkAgg(fig, master=self.profiler_chart_host)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self.profiler_chart_canvas_widget = canvas

    def _run_profiler_clean_export(self):
        if self.profiler_df is None:
            messagebox.showerror("Nothing to export", "Please analyze a file first.")
            return

        filters = {k: v.get() for k, v in self.profiler_filter_vars.items()}
        filters["fill_missing_numeric_mean"] = self.profiler_fill_mean_var.get()
        cleaned = dp.clean_data(self.profiler_df, filters)

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if not save_path:
            return
        try:
            dp.save_data(cleaned, save_path)
        except ValueError as e:
            messagebox.showerror("Couldn't save file", str(e))
            return

        messagebox.showinfo(
            "Export complete",
            f"Cleaned data saved to:\n{save_path}\n\n"
            f"{len(self.profiler_df)} rows → {len(cleaned)} rows after cleaning.")
        try:
            folder = os.path.dirname(os.path.abspath(save_path))
            if os.name == "nt":
                os.startfile(folder)
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception:
            pass

    # ==================================================================
    # PAGE: Excel Editor
    # ==================================================================
    def _build_sheet_editor_page(self, parent):
        page, body = self._new_page(
            parent, "Excel Editor",
            "Load a spreadsheet, describe an edit in plain English, and apply it — undo, redo, and save when you're happy")

        # --- Load file ---
        load_frame = ctk.CTkFrame(body, corner_radius=12)
        load_frame.pack(fill="x", pady=10)
        self.editor_input_path = self._path_row(load_frame, "Spreadsheet", self._browse_editor_input)

        load_line = ctk.CTkFrame(load_frame, fg_color="transparent")
        load_line.pack(fill="x", padx=16, pady=(0, 6))
        ctk.CTkLabel(load_line, text="", width=90).pack(side="left")
        self.editor_load_btn = ctk.CTkButton(load_line, text="Load", width=90, height=32,
                                              command=self._load_editor_file)
        self.editor_load_btn.pack(side="left")
        ctk.CTkLabel(load_line, text="Sheet", width=50, anchor="w").pack(side="left", padx=(16, 0))
        self.editor_sheet_selector = ctk.CTkOptionMenu(
            load_line, values=["—"], width=180, state="disabled",
            command=self._on_editor_sheet_change)
        self.editor_sheet_selector.pack(side="left", padx=6)
        self.editor_load_status = ctk.CTkLabel(load_frame, text="Accepts .csv, .xlsx, .xls",
                                                font=ctk.CTkFont(size=11), text_color="gray")
        self.editor_load_status.pack(anchor="w", padx=16, pady=(0, 14))

        # --- AI configuration (reuses the same provider/key setup as PDF Extractor) ---
        cfg_frame = ctk.CTkFrame(body, corner_radius=12)
        cfg_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(cfg_frame, text="AI Engine (used to turn your instruction into edits)",
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(12, 6))

        editor_provider_line = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        editor_provider_line.pack(fill="x", padx=16, pady=(0, 4))
        ctk.CTkLabel(editor_provider_line, text="AI Engine", width=110, anchor="w").pack(side="left")
        self.editor_provider_selector = ctk.CTkSegmentedButton(
            editor_provider_line, values=["Gemini (Flash) — Free tier", "OpenAI (gpt-5)"],
            command=self._on_editor_provider_change)
        self.editor_provider_selector.set("Gemini (Flash) — Free tier")
        self.editor_provider_selector.pack(side="left", fill="x", expand=True, padx=6)
        self._editor_provider = pe.PROVIDER_GEMINI

        editor_key_line = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        editor_key_line.pack(fill="x", padx=16, pady=(0, 6))
        self.editor_api_key_field_label = ctk.CTkLabel(editor_key_line, text="Gemini API Key", width=110, anchor="w")
        self.editor_api_key_field_label.pack(side="left")
        self.editor_api_key_entry = ctk.CTkEntry(editor_key_line, show="*")
        self.editor_api_key_entry.pack(side="left", fill="x", expand=True, padx=6)
        try:
            existing_key = pe.load_api_key(self._editor_provider)
            if existing_key:
                self.editor_api_key_entry.insert(0, existing_key)
        except Exception:
            pass
        ctk.CTkButton(editor_key_line, text="Save", width=70, command=self._save_editor_api_key).pack(side="left")

        self.editor_api_key_status = ctk.CTkLabel(
            cfg_frame, text="Same key/model already saved for PDF Extractor is reused here automatically.",
            font=ctk.CTkFont(size=11), text_color="gray")
        self.editor_api_key_status.pack(anchor="w", padx=16, pady=(0, 12))

        editor_model_line = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        editor_model_line.pack(fill="x", padx=16, pady=(0, 4))
        ctk.CTkLabel(editor_model_line, text="Model", width=110, anchor="w").pack(side="left")
        self.editor_model_combo = ctk.CTkComboBox(
            editor_model_line, values=pe.SUGGESTED_MODELS[self._editor_provider])
        self.editor_model_combo.set(pe.DEFAULT_GEMINI_MODEL)
        self.editor_model_combo.pack(side="left", fill="x", expand=True, padx=6)

        self.editor_provider_warning = ctk.CTkLabel(cfg_frame, text="", text_color="#e0a020",
                                                      font=ctk.CTkFont(size=11), wraplength=650, justify="left")
        self.editor_provider_warning.pack(anchor="w", padx=16, pady=(0, 10))
        self._refresh_editor_provider_warning()

        # --- Prompt ---
        prompt_frame = ctk.CTkFrame(body, corner_radius=12)
        prompt_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(prompt_frame, text="Describe the edit", font=ctk.CTkFont(weight="bold", size=14)
                     ).pack(pady=(12, 6))
        ctk.CTkLabel(prompt_frame,
                     text="e.g. \"Remove rows where Status is Cancelled, then sort by Date descending\"",
                     font=ctk.CTkFont(size=11), text_color="gray").pack(anchor="w", padx=16)
        self.editor_prompt = ctk.CTkTextbox(prompt_frame, height=70, corner_radius=8)
        self.editor_prompt.pack(fill="x", padx=16, pady=(4, 6))

        apply_line = ctk.CTkFrame(prompt_frame, fg_color="transparent")
        apply_line.pack(fill="x", padx=16, pady=(0, 14))
        self.editor_apply_btn = ctk.CTkButton(apply_line, text="Apply Edit", height=36,
                                               font=ctk.CTkFont(weight="bold"),
                                               command=self._start_editor_apply, state="disabled")
        self.editor_apply_btn.pack(side="left")
        self.editor_apply_status = ctk.CTkLabel(apply_line, text="Load a spreadsheet first",
                                                 font=ctk.CTkFont(size=11), text_color="gray")
        self.editor_apply_status.pack(side="left", padx=12)

        # --- Undo / redo / save toolbar ---
        toolbar_frame = ctk.CTkFrame(body, corner_radius=12)
        toolbar_frame.pack(fill="x", pady=10)
        toolbar_line = ctk.CTkFrame(toolbar_frame, fg_color="transparent")
        toolbar_line.pack(fill="x", padx=16, pady=12)
        self.editor_undo_btn = ctk.CTkButton(toolbar_line, text="↶ Undo", width=100,
                                              command=self._on_editor_undo, state="disabled")
        self.editor_undo_btn.pack(side="left", padx=(0, 8))
        self.editor_redo_btn = ctk.CTkButton(toolbar_line, text="↷ Redo", width=100,
                                              command=self._on_editor_redo, state="disabled")
        self.editor_redo_btn.pack(side="left", padx=(0, 8))
        self.editor_save_btn = ctk.CTkButton(toolbar_line, text="💾 Save", width=100,
                                              command=self._on_editor_save, state="disabled")
        self.editor_save_btn.pack(side="left", padx=(0, 8))
        self.editor_saveas_btn = ctk.CTkButton(toolbar_line, text="Save As...", width=100,
                                                fg_color="transparent", border_width=1,
                                                text_color=("gray20", "gray85"),
                                                command=self._on_editor_save_as, state="disabled")
        self.editor_saveas_btn.pack(side="left")

        # --- Table preview ---
        table_frame = ctk.CTkFrame(body, corner_radius=12)
        table_frame.pack(fill="both", expand=True, pady=10)
        ctk.CTkLabel(table_frame, text="Sheet Preview", font=ctk.CTkFont(weight="bold", size=14)
                     ).pack(pady=(12, 4), anchor="w", padx=16)
        self.editor_table_info = ctk.CTkLabel(table_frame, text="No file loaded yet.",
                                               font=ctk.CTkFont(size=11), text_color="gray")
        self.editor_table_info.pack(anchor="w", padx=16, pady=(0, 8))

        tree_host = ctk.CTkFrame(table_frame, fg_color="transparent")
        tree_host.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        self._style_editor_treeview()
        y_scroll = ttk.Scrollbar(tree_host, orient="vertical")
        x_scroll = ttk.Scrollbar(tree_host, orient="horizontal")
        self.editor_tree = ttk.Treeview(
            tree_host, style="Editor.Treeview", height=14,
            yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        y_scroll.configure(command=self.editor_tree.yview)
        x_scroll.configure(command=self.editor_tree.xview)
        self.editor_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        tree_host.grid_rowconfigure(0, weight=1)
        tree_host.grid_columnconfigure(0, weight=1)

        # --- Edit log ---
        log_frame = ctk.CTkFrame(body, corner_radius=12)
        log_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(log_frame, text="Edit Log", font=ctk.CTkFont(weight="bold", size=14)
                     ).pack(pady=(12, 6), anchor="w", padx=16)
        self.editor_log_box = ctk.CTkTextbox(log_frame, height=110, corner_radius=8)
        self.editor_log_box.pack(fill="x", padx=16, pady=(0, 14))

        return page

    def _style_editor_treeview(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Editor.Treeview",
                         background="#1f1f1f", fieldbackground="#1f1f1f",
                         foreground="#e6e6e6", rowheight=26, borderwidth=0)
        style.configure("Editor.Treeview.Heading",
                         background="#2b2b2b", foreground="#e6e6e6",
                         font=("Segoe UI", 10, "bold"), borderwidth=0)
        style.map("Editor.Treeview", background=[("selected", "#3c8cff")])

    def _browse_editor_input(self, entry):
        path = filedialog.askopenfilename(filetypes=[("Spreadsheet files", "*.csv *.xlsx *.xls")])
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)

    def _load_editor_file(self):
        filepath = self.editor_input_path.get().strip()
        if not filepath:
            messagebox.showerror("Missing info", "Please select a spreadsheet file first.")
            return
        try:
            sheets, sheet_order = se.load_workbook(filepath)
        except ValueError as e:
            messagebox.showerror("Couldn't load file", str(e))
            return

        self.editor_session = se.EditSession(filepath, sheets, sheet_order)
        self.editor_sheet_selector.configure(values=sheet_order, state="normal")
        self.editor_sheet_selector.set(sheet_order[0])
        self.editor_log_box.delete("1.0", "end")
        self.editor_load_status.configure(
            text=f"Loaded {len(sheet_order)} sheet(s) from {os.path.basename(filepath)}.",
            text_color="gray")
        self._refresh_editor_table()
        self._update_editor_toolbar_state()
        self.editor_apply_btn.configure(state="normal")
        self.editor_apply_status.configure(text="Ready", text_color="gray")

    def _on_editor_sheet_change(self, selection):
        if not self.editor_session:
            return
        self.editor_session.set_active_sheet(selection)
        self._refresh_editor_table()
        self._update_editor_toolbar_state()

    def _refresh_editor_table(self):
        tree = self.editor_tree
        tree.delete(*tree.get_children())
        if not self.editor_session:
            tree["columns"] = ()
            self.editor_table_info.configure(text="No file loaded yet.")
            return

        df = self.editor_session.df
        columns = [str(c) for c in df.columns]
        tree["columns"] = columns
        tree["show"] = "headings"
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="w", stretch=True)

        preview = df.head(se.MAX_PREVIEW_ROWS)
        for _, row in preview.iterrows():
            values = ["" if pd.isna(v) else str(v) for v in row.tolist()]
            tree.insert("", "end", values=values)

        total_rows = len(df)
        note = ""
        if total_rows > se.MAX_PREVIEW_ROWS:
            note = f" (showing first {se.MAX_PREVIEW_ROWS} — edits and saving still cover all rows)"
        self.editor_table_info.configure(
            text=f"'{self.editor_session.active_sheet}' — {total_rows} rows, {len(columns)} columns{note}")

    def _on_editor_provider_change(self, selection):
        self._editor_provider = (pe.PROVIDER_OPENAI if selection.startswith("OpenAI")
                                  else pe.PROVIDER_GEMINI)
        label = "OpenAI API Key" if self._editor_provider == pe.PROVIDER_OPENAI else "Gemini API Key"
        self.editor_api_key_field_label.configure(text=label)

        self.editor_api_key_entry.delete(0, "end")
        try:
            existing_key = pe.load_api_key(self._editor_provider)
            if existing_key:
                self.editor_api_key_entry.insert(0, existing_key)
        except Exception:
            pass

        default_model = (pe.DEFAULT_OPENAI_MODEL if self._editor_provider == pe.PROVIDER_OPENAI
                          else pe.DEFAULT_GEMINI_MODEL)
        self.editor_model_combo.configure(values=pe.SUGGESTED_MODELS[self._editor_provider])
        self.editor_model_combo.set(default_model)
        self._refresh_editor_provider_warning()

    def _refresh_editor_provider_warning(self):
        if se.is_provider_available(self._editor_provider):
            self.editor_provider_warning.configure(text="")
            return
        pkg = "openai" if self._editor_provider == pe.PROVIDER_OPENAI else "google-generativeai"
        self.editor_provider_warning.configure(
            text=f"⚠ The '{pkg}' package isn't installed. Run: pip install {pkg}")

    def _save_editor_api_key(self):
        key = self.editor_api_key_entry.get().strip()
        if not key:
            messagebox.showerror("Missing key", "Please enter an API key first.")
            return
        try:
            pe.save_api_key(key, provider=self._editor_provider)
            self.editor_api_key_status.configure(text="✓ Key saved to .env", text_color="#3ca34d")
        except Exception as e:
            messagebox.showerror("Couldn't save key", str(e))

    def _start_editor_apply(self):
        if not self.editor_session:
            messagebox.showerror("Nothing loaded", "Please load a spreadsheet first.")
            return
        if self.editor_thread and self.editor_thread.is_alive():
            messagebox.showinfo("Busy", "An edit is already being applied.")
            return

        provider = self._editor_provider
        if not se.is_provider_available(provider):
            pkg = "openai" if provider == pe.PROVIDER_OPENAI else "google-generativeai"
            messagebox.showerror("Missing dependency",
                                  f"The '{pkg}' package isn't installed.\nRun: pip install {pkg}")
            return

        api_key = self.editor_api_key_entry.get().strip()
        model_name = self.editor_model_combo.get().strip()
        instruction = self.editor_prompt.get("1.0", "end").strip()

        if not api_key:
            messagebox.showerror("Missing info", f"Please enter and save your {provider.title()} API key.")
            return
        if not instruction:
            messagebox.showerror("Missing info", "Please describe the edit you want.")
            return

        self.editor_apply_btn.configure(state="disabled", text="Applying...")
        self.editor_apply_status.configure(text="Thinking through your edit...", text_color="gray")

        self.editor_thread = threading.Thread(
            target=self._run_editor_apply_job,
            args=(instruction, provider, api_key, model_name),
            daemon=True)
        self.editor_thread.start()

    def _run_editor_apply_job(self, instruction, provider, api_key, model_name):
        session = self.editor_session
        try:
            ops, explanation = se.get_edit_plan(
                instruction, session.df, session.active_sheet, provider, api_key, model_name=model_name)
        except se.AuthError as e:
            self.editor_queue.put(("error", f"Authentication failed: {e}"))
            return
        except Exception as e:
            self.editor_queue.put(("error", f"Couldn't plan the edit: {e}"))
            return

        if not ops:
            self.editor_queue.put(("info", explanation or "The AI didn't suggest any changes for that instruction."))
            return

        try:
            session.apply(ops, instruction=instruction, explanation=explanation)
        except se.OperationError as e:
            self.editor_queue.put(("error", f"Couldn't apply the suggested edit: {e}"))
            return

        self.editor_queue.put(("applied", (instruction, explanation, ops)))

    def _poll_editor_queue(self):
        try:
            while True:
                kind, payload = self.editor_queue.get_nowait()
                if kind == "applied":
                    instruction, explanation, ops = payload
                    self._refresh_editor_table()
                    self._update_editor_toolbar_state()
                    self.editor_log_box.insert(
                        "end", f"✓ \"{instruction}\"\n   → {explanation}\n")
                    self.editor_log_box.see("end")
                    self.editor_apply_status.configure(text="Applied", text_color="#3ca34d")
                    self.editor_prompt.delete("1.0", "end")
                elif kind == "info":
                    self.editor_log_box.insert("end", f"ℹ {payload}\n")
                    self.editor_log_box.see("end")
                    self.editor_apply_status.configure(text="No changes made", text_color="#e0a020")
                elif kind == "error":
                    self.editor_log_box.insert("end", f"✗ {payload}\n")
                    self.editor_log_box.see("end")
                    self.editor_apply_status.configure(text="Failed — see log", text_color="#e05050")
                if kind in ("applied", "info", "error"):
                    self.editor_apply_btn.configure(state="normal", text="Apply Edit")
        except queue.Empty:
            pass
        self.after(150, self._poll_editor_queue)

    def _update_editor_toolbar_state(self):
        if not self.editor_session:
            for btn in (self.editor_undo_btn, self.editor_redo_btn,
                        self.editor_save_btn, self.editor_saveas_btn):
                btn.configure(state="disabled")
            return
        self.editor_undo_btn.configure(state="normal" if self.editor_session.can_undo() else "disabled")
        self.editor_redo_btn.configure(state="normal" if self.editor_session.can_redo() else "disabled")
        self.editor_save_btn.configure(state="normal")
        self.editor_saveas_btn.configure(state="normal")

    def _on_editor_undo(self):
        try:
            self.editor_session.undo()
        except se.OperationError as e:
            messagebox.showinfo("Nothing to undo", str(e))
            return
        self._refresh_editor_table()
        self._update_editor_toolbar_state()
        self.editor_log_box.insert("end", f"↶ Undid last edit on '{self.editor_session.active_sheet}'\n")
        self.editor_log_box.see("end")

    def _on_editor_redo(self):
        try:
            self.editor_session.redo()
        except se.OperationError as e:
            messagebox.showinfo("Nothing to redo", str(e))
            return
        self._refresh_editor_table()
        self._update_editor_toolbar_state()
        self.editor_log_box.insert("end", f"↷ Redid last edit on '{self.editor_session.active_sheet}'\n")
        self.editor_log_box.see("end")

    def _on_editor_save(self):
        if not self.editor_session:
            return
        try:
            saved_path = self.editor_session.save()
        except ValueError as e:
            messagebox.showerror("Couldn't save file", str(e))
            return
        messagebox.showinfo("Saved", f"Saved to:\n{saved_path}")

    def _on_editor_save_as(self):
        if not self.editor_session:
            return
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if not save_path:
            return
        try:
            saved_path = self.editor_session.save(save_path)
        except ValueError as e:
            messagebox.showerror("Couldn't save file", str(e))
            return
        messagebox.showinfo("Saved", f"Saved to:\n{saved_path}")
        try:
            folder = os.path.dirname(os.path.abspath(saved_path))
            if os.name == "nt":
                os.startfile(folder)
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception:
            pass

    # ==================================================================
    # PAGE: Dashboard Builder
    # ==================================================================
    def _build_dashboard_builder_page(self, parent):
        page, body = self._new_page(
            parent, "Dashboard Builder",
            "Upload a data file and get a reference dashboard: which chart fits which "
            "column, a live preview, an editable Excel dashboard, and a plain-English summary")

        # --- Load data ---
        load_frame = ctk.CTkFrame(body, corner_radius=12)
        load_frame.pack(fill="x", pady=10)
        self.dash_input_path = self._path_row(load_frame, "Data file", self._browse_dash_input)
        analyze_line = ctk.CTkFrame(load_frame, fg_color="transparent")
        analyze_line.pack(fill="x", padx=16, pady=(0, 14))
        self.dash_analyze_btn = ctk.CTkButton(analyze_line, text="Build Dashboard Plan", height=36,
                                               command=self._start_dashboard_analysis)
        self.dash_analyze_btn.pack(side="left")
        self.dash_status_label = ctk.CTkLabel(analyze_line, text="Accepts .csv, .xlsx, .xls",
                                               font=ctk.CTkFont(size=11), text_color="gray")
        self.dash_status_label.pack(side="left", padx=12)

        # --- AI configuration (reuses the same provider/key already saved for PDF Extractor) ---
        cfg_frame = ctk.CTkFrame(body, corner_radius=12)
        cfg_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(cfg_frame, text="AI Engine (writes the plain-English summary — optional)",
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(12, 6))

        dash_provider_line = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        dash_provider_line.pack(fill="x", padx=16, pady=(0, 4))
        ctk.CTkLabel(dash_provider_line, text="AI Engine", width=110, anchor="w").pack(side="left")
        self.dash_provider_selector = ctk.CTkSegmentedButton(
            dash_provider_line, values=["Gemini (Flash) — Free tier", "OpenAI (gpt-5)"],
            command=self._on_dash_provider_change)
        self.dash_provider_selector.set("Gemini (Flash) — Free tier")
        self.dash_provider_selector.pack(side="left", fill="x", expand=True, padx=6)
        self._dash_provider = pe.PROVIDER_GEMINI

        dash_key_line = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        dash_key_line.pack(fill="x", padx=16, pady=(0, 6))
        self.dash_api_key_field_label = ctk.CTkLabel(dash_key_line, text="Gemini API Key", width=110, anchor="w")
        self.dash_api_key_field_label.pack(side="left")
        self.dash_api_key_entry = ctk.CTkEntry(dash_key_line, show="*")
        self.dash_api_key_entry.pack(side="left", fill="x", expand=True, padx=6)
        try:
            existing_key = pe.load_api_key(self._dash_provider)
            if existing_key:
                self.dash_api_key_entry.insert(0, existing_key)
        except Exception:
            pass
        ctk.CTkButton(dash_key_line, text="Save", width=70, command=self._save_dash_api_key).pack(side="left")

        self.dash_api_key_status = ctk.CTkLabel(
            cfg_frame, text="Same key/model already saved for PDF Extractor is reused here automatically.",
            font=ctk.CTkFont(size=11), text_color="gray")
        self.dash_api_key_status.pack(anchor="w", padx=16, pady=(0, 12))

        dash_model_line = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        dash_model_line.pack(fill="x", padx=16, pady=(0, 4))
        ctk.CTkLabel(dash_model_line, text="Model", width=110, anchor="w").pack(side="left")
        self.dash_model_combo = ctk.CTkComboBox(dash_model_line, values=pe.SUGGESTED_MODELS[self._dash_provider])
        self.dash_model_combo.set(pe.DEFAULT_GEMINI_MODEL)
        self.dash_model_combo.pack(side="left", fill="x", expand=True, padx=6)

        self.dash_summary_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(cfg_frame, text="Also write an AI summary explaining the dashboard's logic",
                         variable=self.dash_summary_var).pack(anchor="w", padx=16, pady=(4, 4))

        self.dash_provider_warning = ctk.CTkLabel(cfg_frame, text="", text_color="#e0a020",
                                                    font=ctk.CTkFont(size=11), wraplength=650, justify="left")
        self.dash_provider_warning.pack(anchor="w", padx=16, pady=(0, 10))
        self._refresh_dash_provider_warning()

        # --- Column classification ---
        classify_frame = ctk.CTkFrame(body, corner_radius=12)
        classify_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(classify_frame, text="Column Classification", font=ctk.CTkFont(weight="bold", size=14)
                     ).pack(pady=(12, 6), anchor="w", padx=16)
        self.dash_columns_box = ctk.CTkTextbox(classify_frame, height=110, corner_radius=8)
        self.dash_columns_box.pack(fill="x", padx=16, pady=(0, 14))
        self.dash_columns_box.insert("1.0", "Build a dashboard plan to see how each column was classified.")
        self.dash_columns_box.configure(state="disabled")

        # --- Chart plan + preview ---
        plan_frame = ctk.CTkFrame(body, corner_radius=12)
        plan_frame.pack(fill="both", expand=True, pady=10)
        ctk.CTkLabel(plan_frame, text="Suggested Charts", font=ctk.CTkFont(weight="bold", size=14)
                     ).pack(pady=(12, 6), anchor="w", padx=16)
        self.dash_plan_box = ctk.CTkTextbox(plan_frame, height=110, corner_radius=8)
        self.dash_plan_box.pack(fill="x", padx=16, pady=(0, 10))
        self.dash_plan_box.insert("1.0", "Chart-by-chart reasoning will appear here.")
        self.dash_plan_box.configure(state="disabled")

        self.dash_chart_host = ctk.CTkFrame(plan_frame, fg_color="transparent", height=320)
        self.dash_chart_host.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        ctk.CTkLabel(self.dash_chart_host, text="Preview will appear here after analysis",
                     font=ctk.CTkFont(size=12), text_color="gray").pack(expand=True)

        # --- Summary ---
        summary_frame = ctk.CTkFrame(body, corner_radius=12)
        summary_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(summary_frame, text="Dashboard Logic — Summary", font=ctk.CTkFont(weight="bold", size=14)
                     ).pack(pady=(12, 6), anchor="w", padx=16)
        self.dash_summary_box = ctk.CTkTextbox(summary_frame, height=110, corner_radius=8)
        self.dash_summary_box.pack(fill="x", padx=16, pady=(0, 14))

        # --- Export ---
        export_frame = ctk.CTkFrame(body, corner_radius=12)
        export_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(export_frame, text="Export", font=ctk.CTkFont(weight="bold", size=14)
                     ).pack(pady=(12, 6))
        self.dash_export_btn = ctk.CTkButton(export_frame, text="Export Dashboard Excel", height=44,
                                              font=ctk.CTkFont(size=15, weight="bold"),
                                              command=self._export_dashboard_excel, state="disabled")
        self.dash_export_btn.pack(pady=(0, 16), padx=16, fill="x")

        return page

    def _browse_dash_input(self, entry):
        path = filedialog.askopenfilename(filetypes=[("Data files", "*.csv *.xlsx *.xls")])
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)

    def _on_dash_provider_change(self, selection):
        self._dash_provider = (pe.PROVIDER_OPENAI if selection.startswith("OpenAI")
                                else pe.PROVIDER_GEMINI)
        label = "OpenAI API Key" if self._dash_provider == pe.PROVIDER_OPENAI else "Gemini API Key"
        self.dash_api_key_field_label.configure(text=label)
        self.dash_api_key_entry.delete(0, "end")
        try:
            existing_key = pe.load_api_key(self._dash_provider)
            if existing_key:
                self.dash_api_key_entry.insert(0, existing_key)
        except Exception:
            pass
        default_model = (pe.DEFAULT_OPENAI_MODEL if self._dash_provider == pe.PROVIDER_OPENAI
                          else pe.DEFAULT_GEMINI_MODEL)
        self.dash_model_combo.configure(values=pe.SUGGESTED_MODELS[self._dash_provider])
        self.dash_model_combo.set(default_model)
        self._refresh_dash_provider_warning()

    def _refresh_dash_provider_warning(self):
        if db.is_provider_available(self._dash_provider):
            self.dash_provider_warning.configure(text="")
            return
        pkg = "openai" if self._dash_provider == pe.PROVIDER_OPENAI else "google-generativeai"
        self.dash_provider_warning.configure(
            text=f"⚠ The '{pkg}' package isn't installed. Run: pip install {pkg}")

    def _save_dash_api_key(self):
        key = self.dash_api_key_entry.get().strip()
        if not key:
            messagebox.showerror("Missing key", "Please enter an API key first.")
            return
        try:
            pe.save_api_key(key, provider=self._dash_provider)
            self.dash_api_key_status.configure(text="✓ Key saved to .env", text_color="#3ca34d")
        except Exception as e:
            messagebox.showerror("Couldn't save key", str(e))

    def _start_dashboard_analysis(self):
        filepath = self.dash_input_path.get().strip()
        if not filepath:
            messagebox.showerror("Missing info", "Please select a CSV or Excel file first.")
            return
        if self.dashboard_thread and self.dashboard_thread.is_alive():
            messagebox.showinfo("Busy", "A dashboard is already being built.")
            return

        want_summary = self.dash_summary_var.get()
        provider = self._dash_provider
        api_key = self.dash_api_key_entry.get().strip()
        model_name = self.dash_model_combo.get().strip()

        if want_summary and not api_key:
            proceed = messagebox.askyesno(
                "No API key",
                "No API key is set, so the AI summary will be skipped — the charts and "
                "Excel dashboard will still be built. Continue?")
            if not proceed:
                return
            want_summary = False

        self.dash_analyze_btn.configure(state="disabled", text="Building...")
        self.dash_export_btn.configure(state="disabled")
        self.dash_status_label.configure(text="Reading file and classifying columns...", text_color="gray")

        self.dashboard_thread = threading.Thread(
            target=self._run_dashboard_job,
            args=(filepath, want_summary, provider, api_key, model_name),
            daemon=True)
        self.dashboard_thread.start()

    def _run_dashboard_job(self, filepath, want_summary, provider, api_key, model_name):
        try:
            df = dp.load_data(filepath)
        except ValueError as e:
            self.dashboard_queue.put(("error", str(e)))
            return

        column_info = db.classify_columns(df)
        plan = db.suggest_chart_plan(df, column_info)

        summary = ""
        summary_error = None
        if want_summary:
            try:
                summary = db.generate_ai_summary(column_info, plan, provider, api_key, model_name=model_name)
            except db.AuthError as e:
                summary_error = f"Authentication failed, summary skipped: {e}"
            except Exception as e:
                summary_error = f"Summary generation failed (charts still built): {e}"

        self.dashboard_queue.put(("built", (df, column_info, plan, summary, summary_error)))

    def _poll_dashboard_queue(self):
        try:
            while True:
                kind, payload = self.dashboard_queue.get_nowait()
                if kind == "built":
                    self._on_dashboard_built(*payload)
                elif kind == "error":
                    messagebox.showerror("Couldn't build dashboard", payload)
                    self.dash_status_label.configure(text=payload, text_color="#e05050")
                self.dash_analyze_btn.configure(state="normal", text="Build Dashboard Plan")
        except queue.Empty:
            pass
        self.after(150, self._poll_dashboard_queue)

    def _on_dashboard_built(self, df, column_info, plan, summary, summary_error):
        self.dashboard_df = df
        self.dashboard_column_info = column_info
        self.dashboard_plan = plan
        self.dashboard_summary_text = summary

        kind_lines = [f"{c['column']}: {c['kind']}  ({c['unique_count']} unique, {c['missing']} missing)"
                      for c in column_info]
        self._set_textbox(self.dash_columns_box, "\n".join(kind_lines) or "No columns found.")

        if plan:
            plan_lines = [f"[{p['chart_type'].upper()}] {p['title']} — {p['reason']}" for p in plan]
        else:
            plan_lines = ["Not enough structured columns to suggest charts for this file."]
        self._set_textbox(self.dash_plan_box, "\n\n".join(plan_lines))

        for widget in self.dash_chart_host.winfo_children():
            widget.destroy()
        fig = db.render_dashboard_preview(df, plan)
        canvas = FigureCanvasTkAgg(fig, master=self.dash_chart_host)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self.dashboard_chart_canvas_widget = canvas

        self.dash_summary_box.delete("1.0", "end")
        if summary:
            self.dash_summary_box.insert("1.0", summary)
        elif summary_error:
            self.dash_summary_box.insert("1.0", summary_error)
        else:
            self.dash_summary_box.insert("1.0", "(No AI summary requested — charts and column classification are above.)")

        self.dash_status_label.configure(
            text=f"{len(df)} rows, {len(df.columns)} columns — {len(plan)} chart(s) suggested.",
            text_color="gray")
        self.dash_export_btn.configure(state="normal")

    def _set_textbox(self, box, text):
        box.configure(state="normal")
        box.delete("1.0", "end")
        box.insert("1.0", text)
        box.configure(state="disabled")

    def _export_dashboard_excel(self):
        if self.dashboard_df is None or self.dashboard_plan is None:
            messagebox.showerror("Nothing to export", "Please build a dashboard plan first.")
            return
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not save_path:
            return
        try:
            db.build_excel_dashboard(self.dashboard_df, self.dashboard_plan, save_path,
                                      summary_text=self.dashboard_summary_text)
        except Exception as e:
            messagebox.showerror("Couldn't save dashboard", str(e))
            return
        messagebox.showinfo("Dashboard exported", f"Saved to:\n{save_path}")
        try:
            folder = os.path.dirname(os.path.abspath(save_path))
            if os.name == "nt":
                os.startfile(folder)
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception:
            pass

    # ==================================================================
    # PAGE: AI Assistant
    # ==================================================================
    def _build_ai_assistant_page(self, parent):
        page, body = self._new_page(
            parent, "AI Assistant",
            "Attach up to 3 images (a table, schema, error, sketch — anything visual) and "
            "describe what to write — SQL, DAX, Python, an Excel formula, VBA, and more")

        # --- AI configuration (reuses the same provider/key already saved for PDF Extractor) ---
        cfg_frame = ctk.CTkFrame(body, corner_radius=12)
        cfg_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(cfg_frame, text="AI Engine", font=ctk.CTkFont(weight="bold", size=14)
                     ).pack(pady=(12, 6))

        asst_provider_line = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        asst_provider_line.pack(fill="x", padx=16, pady=(0, 4))
        ctk.CTkLabel(asst_provider_line, text="AI Engine", width=110, anchor="w").pack(side="left")
        self.asst_provider_selector = ctk.CTkSegmentedButton(
            asst_provider_line, values=["Gemini (Flash) — Free tier", "OpenAI (gpt-5)"],
            command=self._on_asst_provider_change)
        self.asst_provider_selector.set("Gemini (Flash) — Free tier")
        self.asst_provider_selector.pack(side="left", fill="x", expand=True, padx=6)
        self._asst_provider = pe.PROVIDER_GEMINI

        asst_key_line = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        asst_key_line.pack(fill="x", padx=16, pady=(0, 6))
        self.asst_api_key_field_label = ctk.CTkLabel(asst_key_line, text="Gemini API Key", width=110, anchor="w")
        self.asst_api_key_field_label.pack(side="left")
        self.asst_api_key_entry = ctk.CTkEntry(asst_key_line, show="*")
        self.asst_api_key_entry.pack(side="left", fill="x", expand=True, padx=6)
        try:
            existing_key = pe.load_api_key(self._asst_provider)
            if existing_key:
                self.asst_api_key_entry.insert(0, existing_key)
        except Exception:
            pass
        ctk.CTkButton(asst_key_line, text="Save", width=70, command=self._save_asst_api_key).pack(side="left")

        self.asst_api_key_status = ctk.CTkLabel(
            cfg_frame, text="Same key/model already saved for PDF Extractor is reused here automatically.",
            font=ctk.CTkFont(size=11), text_color="gray")
        self.asst_api_key_status.pack(anchor="w", padx=16, pady=(0, 12))

        asst_model_line = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        asst_model_line.pack(fill="x", padx=16, pady=(0, 4))
        ctk.CTkLabel(asst_model_line, text="Model", width=110, anchor="w").pack(side="left")
        self.asst_model_combo = ctk.CTkComboBox(asst_model_line, values=pe.SUGGESTED_MODELS[self._asst_provider])
        self.asst_model_combo.set(pe.DEFAULT_GEMINI_MODEL)
        self.asst_model_combo.pack(side="left", fill="x", expand=True, padx=6)

        ctk.CTkLabel(cfg_frame,
                     text="Use a vision-capable model — the suggested defaults both support images.",
                     font=ctk.CTkFont(size=11), text_color="gray").pack(anchor="w", padx=16, pady=(0, 10))

        self.asst_provider_warning = ctk.CTkLabel(cfg_frame, text="", text_color="#e0a020",
                                                    font=ctk.CTkFont(size=11), wraplength=650, justify="left")
        self.asst_provider_warning.pack(anchor="w", padx=16, pady=(0, 10))
        self._refresh_asst_provider_warning()

        # --- Image attachments ---
        img_frame = ctk.CTkFrame(body, corner_radius=12)
        img_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(img_frame, text="Attach Images (up to 3, optional)",
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(12, 6), anchor="w", padx=16)

        self.assistant_image_entries = []
        for i in range(aa.MAX_IMAGES):
            row = ctk.CTkFrame(img_frame, fg_color="transparent")
            row.pack(fill="x", padx=16, pady=4)
            ctk.CTkLabel(row, text=f"Image {i + 1}", width=70, anchor="w").pack(side="left")
            entry = ctk.CTkEntry(row)
            entry.pack(side="left", fill="x", expand=True, padx=6)
            ctk.CTkButton(row, text="Browse", width=80,
                          command=lambda e=entry: self._browse_asst_image(e)).pack(side="left", padx=(0, 4))
            ctk.CTkButton(row, text="Clear", width=60, fg_color="transparent", border_width=1,
                          text_color=("gray20", "gray85"),
                          command=lambda e=entry: e.delete(0, "end")).pack(side="left")
            self.assistant_image_entries.append(entry)
        ctk.CTkLabel(img_frame, text="", height=6).pack()

        # --- Prompt ---
        prompt_frame = ctk.CTkFrame(body, corner_radius=12)
        prompt_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(prompt_frame, text="What should it write?", font=ctk.CTkFont(weight="bold", size=14)
                     ).pack(pady=(12, 6))
        ctk.CTkLabel(prompt_frame,
                     text="e.g. \"Write a SQL query that joins these two tables on customer_id\" or "
                          "\"Turn this into a DAX measure\" or \"Write the VBA macro for this\"",
                     font=ctk.CTkFont(size=11), text_color="gray").pack(anchor="w", padx=16)
        self.asst_prompt = ctk.CTkTextbox(prompt_frame, height=80, corner_radius=8)
        self.asst_prompt.pack(fill="x", padx=16, pady=(4, 6))

        ask_line = ctk.CTkFrame(prompt_frame, fg_color="transparent")
        ask_line.pack(fill="x", padx=16, pady=(0, 14))
        self.asst_ask_btn = ctk.CTkButton(ask_line, text="Ask AI Assistant", height=36,
                                           font=ctk.CTkFont(weight="bold"), command=self._start_asst_ask)
        self.asst_ask_btn.pack(side="left")
        self.asst_status_label = ctk.CTkLabel(ask_line, text="Ready", font=ctk.CTkFont(size=11), text_color="gray")
        self.asst_status_label.pack(side="left", padx=12)

        # --- Output ---
        out_frame = ctk.CTkFrame(body, corner_radius=12)
        out_frame.pack(fill="both", expand=True, pady=10)
        out_header = ctk.CTkFrame(out_frame, fg_color="transparent")
        out_header.pack(fill="x", padx=16, pady=(12, 6))
        ctk.CTkLabel(out_header, text="Response", font=ctk.CTkFont(weight="bold", size=14)).pack(side="left")
        ctk.CTkButton(out_header, text="Copy", width=70, height=26,
                      command=self._copy_asst_output).pack(side="right")
        self.asst_output_box = ctk.CTkTextbox(out_frame, height=280, corner_radius=8, wrap="word")
        self.asst_output_box.pack(fill="both", expand=True, padx=16, pady=(0, 14))

        return page

    def _browse_asst_image(self, entry):
        path = filedialog.askopenfilename(
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.webp *.gif *.bmp")])
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)

    def _on_asst_provider_change(self, selection):
        self._asst_provider = (pe.PROVIDER_OPENAI if selection.startswith("OpenAI")
                                else pe.PROVIDER_GEMINI)
        label = "OpenAI API Key" if self._asst_provider == pe.PROVIDER_OPENAI else "Gemini API Key"
        self.asst_api_key_field_label.configure(text=label)
        self.asst_api_key_entry.delete(0, "end")
        try:
            existing_key = pe.load_api_key(self._asst_provider)
            if existing_key:
                self.asst_api_key_entry.insert(0, existing_key)
        except Exception:
            pass
        default_model = (pe.DEFAULT_OPENAI_MODEL if self._asst_provider == pe.PROVIDER_OPENAI
                          else pe.DEFAULT_GEMINI_MODEL)
        self.asst_model_combo.configure(values=pe.SUGGESTED_MODELS[self._asst_provider])
        self.asst_model_combo.set(default_model)
        self._refresh_asst_provider_warning()

    def _refresh_asst_provider_warning(self):
        if aa.is_provider_available(self._asst_provider):
            self.asst_provider_warning.configure(text="")
            return
        pkg = "openai" if self._asst_provider == pe.PROVIDER_OPENAI else "google-generativeai"
        self.asst_provider_warning.configure(
            text=f"⚠ The '{pkg}' package isn't installed. Run: pip install {pkg}")

    def _save_asst_api_key(self):
        key = self.asst_api_key_entry.get().strip()
        if not key:
            messagebox.showerror("Missing key", "Please enter an API key first.")
            return
        try:
            pe.save_api_key(key, provider=self._asst_provider)
            self.asst_api_key_status.configure(text="✓ Key saved to .env", text_color="#3ca34d")
        except Exception as e:
            messagebox.showerror("Couldn't save key", str(e))

    def _start_asst_ask(self):
        if self.assistant_thread and self.assistant_thread.is_alive():
            messagebox.showinfo("Busy", "Already working on a request.")
            return

        provider = self._asst_provider
        if not aa.is_provider_available(provider):
            pkg = "openai" if provider == pe.PROVIDER_OPENAI else "google-generativeai"
            messagebox.showerror("Missing dependency",
                                  f"The '{pkg}' package isn't installed.\nRun: pip install {pkg}")
            return

        api_key = self.asst_api_key_entry.get().strip()
        model_name = self.asst_model_combo.get().strip()
        prompt = self.asst_prompt.get("1.0", "end").strip()
        image_paths = [e.get().strip() for e in self.assistant_image_entries if e.get().strip()]

        if not api_key:
            messagebox.showerror("Missing info", f"Please enter and save your {provider.title()} API key.")
            return
        if not prompt:
            messagebox.showerror("Missing info", "Please describe what you want written.")
            return

        images = aa.validate_images(image_paths)
        if image_paths and not images:
            messagebox.showerror("Invalid images",
                                  "None of the attached paths are readable image files "
                                  "(png/jpg/jpeg/webp/gif/bmp).")
            return

        self.asst_ask_btn.configure(state="disabled", text="Thinking...")
        self.asst_status_label.configure(text="Reading images and writing a response...", text_color="gray")

        self.assistant_thread = threading.Thread(
            target=self._run_asst_job,
            args=(prompt, images, provider, api_key, model_name),
            daemon=True)
        self.assistant_thread.start()

    def _run_asst_job(self, prompt, images, provider, api_key, model_name):
        try:
            response = aa.ask_assistant(prompt, images, provider, api_key, model_name=model_name)
        except aa.AuthError as e:
            self.assistant_queue.put(("error", f"Authentication failed: {e}"))
            return
        except Exception as e:
            self.assistant_queue.put(("error", f"Request failed: {e}"))
            return
        self.assistant_queue.put(("done", response))

    def _poll_assistant_queue(self):
        try:
            while True:
                kind, payload = self.assistant_queue.get_nowait()
                if kind == "done":
                    self.asst_output_box.delete("1.0", "end")
                    self.asst_output_box.insert("1.0", payload)
                    self.asst_status_label.configure(text="Done", text_color="#3ca34d")
                elif kind == "error":
                    self.asst_output_box.delete("1.0", "end")
                    self.asst_output_box.insert("1.0", payload)
                    self.asst_status_label.configure(text="Failed — see response box", text_color="#e05050")
                self.asst_ask_btn.configure(state="normal", text="Ask AI Assistant")
        except queue.Empty:
            pass
        self.after(150, self._poll_assistant_queue)

    def _copy_asst_output(self):
        text = self.asst_output_box.get("1.0", "end").strip()
        if not text:
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self.asst_status_label.configure(text="Copied to clipboard", text_color="#3ca34d")

    # ==================================================================
    # PAGE: Support / donation
    # ==================================================================
    def _build_support_page(self, parent):
        page, body = self._new_page(
            parent, "Support Development",
            "LinkHarvest is free — this is just an optional way to say thanks")

        donate_frame = ctk.CTkFrame(body, corner_radius=12)
        donate_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(donate_frame, text="❤ Leave a UPI donation",
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(16, 4))
        ctk.CTkLabel(donate_frame, text="If this tool saved you time, you can scan to leave a UPI donation.",
                     font=ctk.CTkFont(size=11), text_color="gray").pack()
        self.qr_label = ctk.CTkLabel(donate_frame, text="")
        self.qr_label.pack(pady=16)
        self._load_donation_qr()

        return page

    def _load_donation_qr(self):
        qr_path = os.path.join(BASE_DIR, "assets", "donate_qr.png")
        try:
            if not os.path.exists(qr_path):
                generate_qr_image(qr_path)
            pil_img = Image.open(qr_path)
            ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(180, 180))
            self.qr_label.configure(image=ctk_img)
        except Exception:
            self.qr_label.configure(text="(Donation QR unavailable)")

    # ==================================================================
    # Update checker (applies to the whole app, shown once on startup)
    # ==================================================================
    def _check_update_then_login(self):
        info = check_for_update(__version__)
        if info:
            self.after(0, lambda: self._show_update_prompt(info))
            return
        # Only reached if the version check passed - now require sign-in.
        self.after(0, self._show_login_gate)

    def _show_login_gate(self):
        """Blocks interaction until valid credentials are entered.
        Runs after the update check passes. Closing this dialog quits
        the app, same behavior as the update-required dialog."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Sign In")
        dialog.geometry("360x260")
        dialog.resizable(False, False)
        try:
            dialog.iconbitmap(ICON_PATH)
        except Exception:
            pass
        dialog.transient(self)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", self.destroy)

        ctk.CTkLabel(dialog, text="🔒 Sign In Required",
                     font=ctk.CTkFont(size=17, weight="bold")).pack(pady=(24, 12))

        id_entry = ctk.CTkEntry(dialog, placeholder_text="ID")
        id_entry.pack(padx=32, fill="x", pady=(0, 8))
        pwd_entry = ctk.CTkEntry(dialog, placeholder_text="Password", show="*")
        pwd_entry.pack(padx=32, fill="x")

        status = ctk.CTkLabel(dialog, text="", font=ctk.CTkFont(size=11), text_color="#e05555")
        status.pack(pady=(6, 0))

        def attempt_login():
            uid, pwd = id_entry.get().strip(), pwd_entry.get()
            login_btn.configure(state="disabled", text="Checking...")

            def do_check():
                try:
                    ok = verify_credentials(uid, pwd)
                except AuthError:
                    self.after(0, lambda: status.configure(
                        text="Can't reach license server - check your connection."))
                    self.after(0, lambda: login_btn.configure(state="normal", text="Sign In"))
                    return
                if ok:
                    def on_success():
                        dialog.grab_release()
                        dialog.destroy()
                        start_periodic_recheck(
                            uid, pwd,
                            on_revoked=lambda: self.after(0, self._force_quit_revoked)
                        )
                    self.after(0, on_success)
                else:
                    self.after(0, lambda: status.configure(text="Invalid ID or password."))
                    self.after(0, lambda: login_btn.configure(state="normal", text="Sign In"))

            threading.Thread(target=do_check, daemon=True).start()

        login_btn = ctk.CTkButton(dialog, text="Sign In", height=40, command=attempt_login)
        login_btn.pack(padx=32, pady=16, fill="x")
        pwd_entry.bind("<Return>", lambda e: attempt_login())
        dialog.grab_set()

    def _force_quit_revoked(self):
        """Called from the background recheck thread (auth.py) when
        access is revoked mid-session - the login was valid at startup
        but the row was since removed/changed in the sheet."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Access Revoked")
        dialog.geometry("360x160")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", self.destroy)
        ctk.CTkLabel(dialog, text="Your access has been revoked.\nThe app will now close.",
                     font=ctk.CTkFont(size=13), justify="center").pack(pady=30)
        self.after(2500, self.destroy)

    def _show_update_prompt(self, info):
        target_url = info.get("download_url") or info["url"]

        dialog = ctk.CTkToplevel(self)
        dialog.title("Update Required")
        dialog.geometry("440x260")
        dialog.resizable(False, False)
        try:
            dialog.iconbitmap(ICON_PATH)
        except Exception:
            pass

        dialog.transient(self)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", self.destroy)

        ctk.CTkLabel(dialog, text="🔄 Update Required",
                     font=ctk.CTkFont(size=19, weight="bold")).pack(pady=(24, 6))
        ctk.CTkLabel(
            dialog,
            text=(f"A newer version ({info['version']}) is available.\n"
                  f"You're currently on v{__version__}.\n\n"
                  "Please download and install the update to keep using LinkHarvest."),
            font=ctk.CTkFont(size=13), justify="center", wraplength=380
        ).pack(pady=(0, 18))

        def do_download():
            webbrowser.open(target_url)
            self.destroy()

        ctk.CTkButton(dialog, text="⬇  Download Update", height=44,
                      font=ctk.CTkFont(size=14, weight="bold"),
                      command=do_download).pack(pady=(0, 8), padx=32, fill="x")
        ctk.CTkButton(dialog, text="Quit", height=32, fg_color="transparent",
                      border_width=1, text_color=("gray20", "gray80"),
                      command=self.destroy).pack(padx=32, fill="x")

        dialog.grab_set()


if __name__ == "__main__":
    app = App()
    app.mainloop()
