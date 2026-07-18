"""
app.py
Modern desktop UI for LinkHarvest (Windows) — sidebar navigation between
feature pages: Link Downloader and PDF Extractor.
Run with:  python app.py
"""
import os
import threading
import queue
import webbrowser

import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from downloader import load_sheet, build_download_tasks, run_downloads
from donation import generate_qr_image
from updater import check_for_update
from version import __version__
import pdf_extractor
import data_profiler
from donut_chart import render_donut

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

DEFAULT_THREADS = 10
DEFAULT_TIMEOUT = 30

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ICON_PATH = os.path.join(BASE_DIR, "assets", "icon.ico")

SIDEBAR_WIDTH = 200


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"LinkHarvest v{__version__}")
        self.geometry("980x840")
        self.minsize(860, 780)
        try:
            self.iconbitmap(ICON_PATH)
        except Exception:
            pass

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.pages = {}
        self._build_sidebar()
        self._build_content_area()

        self.show_page("downloader")

        self.after(150, self._poll_downloader_queue)
        self.after(150, self._poll_pdf_queue)
        threading.Thread(target=self._check_update_background, daemon=True).start()

    # ---------------- Sidebar ----------------
    def _build_sidebar(self):
        sidebar = ctk.CTkFrame(self, width=SIDEBAR_WIDTH, corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsw")
        sidebar.grid_propagate(False)

        ctk.CTkLabel(sidebar, text="🔗 LinkHarvest",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(24, 4), padx=16)
        ctk.CTkLabel(sidebar, text=f"v{__version__}", font=ctk.CTkFont(size=11),
                     text_color="gray").pack(pady=(0, 24))

        self.nav_buttons = {}
        nav_items = [
            ("downloader", "🔗  Link Downloader"),
            ("pdf_extractor", "📄  PDF Extractor"),
            ("data_profiler", "📊  Data Profiler"),
        ]
        for key, label in nav_items:
            btn = ctk.CTkButton(
                sidebar, text=label, anchor="w", height=42, corner_radius=8,
                fg_color="transparent", hover_color=("gray80", "gray25"),
                font=ctk.CTkFont(size=14),
                command=lambda k=key: self.show_page(k))
            btn.pack(fill="x", padx=12, pady=4)
            self.nav_buttons[key] = btn

        # Spacer pushes the footer to the bottom
        ctk.CTkFrame(sidebar, fg_color="transparent").pack(fill="both", expand=True)

        ctk.CTkLabel(sidebar, text="More features coming soon",
                     font=ctk.CTkFont(size=10), text_color="gray",
                     wraplength=SIDEBAR_WIDTH - 24, justify="center").pack(pady=(0, 20), padx=12)

    def show_page(self, key):
        for page_key, page in self.pages.items():
            page.grid_forget()
        self.pages[key].grid(row=0, column=0, sticky="nsew")
        for page_key, btn in self.nav_buttons.items():
            if page_key == key:
                btn.configure(fg_color=("gray75", "gray30"))
            else:
                btn.configure(fg_color="transparent")

    # ---------------- Content area ----------------
    def _build_content_area(self):
        content = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        content.grid(row=0, column=1, sticky="nsew")
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(0, weight=1)
        self.content_container = content

        downloader_page = ctk.CTkScrollableFrame(content, corner_radius=0, fg_color="transparent")
        self.pages["downloader"] = downloader_page
        self._build_downloader_page(downloader_page)

        pdf_page = ctk.CTkScrollableFrame(content, corner_radius=0, fg_color="transparent")
        self.pages["pdf_extractor"] = pdf_page
        self._build_pdf_extractor_page(pdf_page)

        profiler_page = ctk.CTkScrollableFrame(content, corner_radius=0, fg_color="transparent")
        self.pages["data_profiler"] = profiler_page
        self._build_data_profiler_page(profiler_page)

    # ================================================================
    # PAGE 1: Link Downloader  (unchanged behavior from the original app)
    # ================================================================
    def _build_downloader_page(self, parent):
        ctk.CTkLabel(parent, text="Link Downloader",
                     font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(25, 2), anchor="w", padx=24)
        ctk.CTkLabel(parent, text="Download images / videos / audio from links in your Excel sheet",
                     font=ctk.CTkFont(size=13), text_color="gray").pack(pady=(0, 15), anchor="w", padx=24)

        file_frame = ctk.CTkFrame(parent, corner_radius=12)
        file_frame.pack(fill="x", padx=24, pady=10)

        self.excel_path = self._path_row(file_frame, "Excel file", self._browse_excel)
        self.excel_path.insert(0, r"C:\Users\Harsh\Documents\links.xlsx")

        self.sheet_name = self._entry_row(file_frame, "Sheet name", "Sheet4")

        self.folder_path = self._path_row(file_frame, "Save folder", self._browse_folder)
        self.folder_path.insert(0, r"C:\Users\Harsh\Downloads\HarvestedFiles")

        range_frame = ctk.CTkFrame(parent, corner_radius=12)
        range_frame.pack(fill="x", padx=24, pady=10)
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

        adv_frame = ctk.CTkFrame(parent, corner_radius=12)
        adv_frame.pack(fill="x", padx=24, pady=10)
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

        donate_frame = ctk.CTkFrame(parent, corner_radius=12)
        donate_frame.pack(fill="x", padx=24, pady=10)
        ctk.CTkLabel(donate_frame, text="❤ Support Development (optional)",
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(12, 4))
        ctk.CTkLabel(donate_frame, text="If this tool saved you time, you can scan to leave a UPI donation.",
                     font=ctk.CTkFont(size=11), text_color="gray").pack()
        self.qr_label = ctk.CTkLabel(donate_frame, text="")
        self.qr_label.pack(pady=12)
        self._load_donation_qr()

        self.progress_bar = ctk.CTkProgressBar(parent)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=24, pady=(16, 4))
        self.progress_label = ctk.CTkLabel(parent, text="0%  (0 / 0)", font=ctk.CTkFont(size=12))
        self.progress_label.pack()

        self.log_box = ctk.CTkTextbox(parent, height=140, corner_radius=8)
        self.log_box.pack(fill="both", expand=True, padx=24, pady=10)

        self.start_btn = ctk.CTkButton(parent, text="Start Download", height=44,
                                        font=ctk.CTkFont(size=15, weight="bold"),
                                        command=self._start_download)
        self.start_btn.pack(pady=(8, 24), fill="x", padx=24)

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

    def _load_donation_qr(self):
        qr_path = os.path.join(BASE_DIR, "assets", "donate_qr.png")
        try:
            if not os.path.exists(qr_path):
                generate_qr_image(qr_path)
            pil_img = Image.open(qr_path)
            ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(150, 150))
            self.qr_label.configure(image=ctk_img)
        except Exception:
            self.qr_label.configure(text="(Donation QR unavailable)")

    def _check_update_background(self):
        info = check_for_update(__version__)
        if info:
            self.after(0, lambda: self._show_update_prompt(info))

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
        if getattr(self, "download_thread", None) and self.download_thread.is_alive():
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

        self.downloader_queue = queue.Queue()
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
            self.downloader_queue.put(("error", f"Failed to read workbook: {e}"))
            self.downloader_queue.put(("done", None))
            return

        if not tasks:
            self.downloader_queue.put(("error", "No links found in the selected rows/columns."))
            self.downloader_queue.put(("done", None))
            return

        self.downloader_queue.put(("info", f"Found {len(tasks)} link(s). Downloading with {threads} thread(s)..."))

        def progress_cb(done, total):
            self.downloader_queue.put(("progress", (done, total)))

        def log_cb(msg):
            self.downloader_queue.put(("warn", msg))

        ok, failed = run_downloads(tasks, threads, timeout, progress_cb, log_cb)
        summary = f"Done. {ok} succeeded, {failed} failed."
        if failed:
            summary += " Some links may remain un-downloaded — see the log above."
        self.downloader_queue.put(("info", summary))
        self.downloader_queue.put(("done", None))

    def _poll_downloader_queue(self):
        q = getattr(self, "downloader_queue", None)
        if q is not None:
            try:
                while True:
                    kind, payload = q.get_nowait()
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
        self.after(150, self._poll_downloader_queue)

    # ================================================================
    # PAGE 2: PDF Extractor
    # ================================================================
    def _build_pdf_extractor_page(self, parent):
        self.pdf_paths = []

        ctk.CTkLabel(parent, text="PDF Extractor",
                     font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(25, 2), anchor="w", padx=24)
        ctk.CTkLabel(parent, text="Extract structured data from PDFs using AI, saved straight to Excel",
                     font=ctk.CTkFont(size=13), text_color="gray").pack(pady=(0, 15), anchor="w", padx=24)

        # API key section
        key_frame = ctk.CTkFrame(parent, corner_radius=12)
        key_frame.pack(fill="x", padx=24, pady=10)
        ctk.CTkLabel(key_frame, text="OpenAI API Key",
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(12, 6), padx=16, anchor="w")
        key_line = ctk.CTkFrame(key_frame, fg_color="transparent")
        key_line.pack(fill="x", padx=16, pady=(0, 14))
        self.api_key_entry = ctk.CTkEntry(key_line, show="•", placeholder_text="sk-...")
        self.api_key_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        saved_key = pdf_extractor.load_api_key()
        if saved_key:
            self.api_key_entry.insert(0, saved_key)
        ctk.CTkButton(key_line, text="Save Key", width=90,
                      command=self._save_pdf_api_key).pack(side="left")

        # PDF file selection
        files_frame = ctk.CTkFrame(parent, corner_radius=12)
        files_frame.pack(fill="x", padx=24, pady=10)
        ctk.CTkLabel(files_frame, text="PDF files",
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(12, 6), padx=16, anchor="w")
        files_line = ctk.CTkFrame(files_frame, fg_color="transparent")
        files_line.pack(fill="x", padx=16, pady=(0, 6))
        ctk.CTkButton(files_line, text="Choose PDFs...", width=140,
                      command=self._browse_pdfs).pack(side="left")
        ctk.CTkButton(files_line, text="Clear", width=80, fg_color="transparent",
                      border_width=1, text_color=("gray20", "gray80"),
                      command=self._clear_pdfs).pack(side="left", padx=(8, 0))
        self.pdf_files_label = ctk.CTkLabel(files_frame, text="No files selected.",
                                             font=ctk.CTkFont(size=12), text_color="gray",
                                             wraplength=650, justify="left")
        self.pdf_files_label.pack(anchor="w", padx=16, pady=(4, 14))

        # Prompt
        prompt_frame = ctk.CTkFrame(parent, corner_radius=12)
        prompt_frame.pack(fill="x", padx=24, pady=10)
        ctk.CTkLabel(prompt_frame, text="Extraction prompt",
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(12, 6), padx=16, anchor="w")
        self.pdf_prompt_box = ctk.CTkTextbox(prompt_frame, height=80, corner_radius=8)
        self.pdf_prompt_box.pack(fill="x", padx=16, pady=(0, 14))
        self.pdf_prompt_box.insert("1.0", "Extract the invoice number, date, and total amount.")

        # Save location
        save_frame = ctk.CTkFrame(parent, corner_radius=12)
        save_frame.pack(fill="x", padx=24, pady=10)
        self.pdf_save_path = self._path_row_save(save_frame, "Save results to",
                                                  self._browse_pdf_excel_save)
        self.pdf_save_path.insert(0, r"C:\Users\Harsh\Documents\pdf_extraction_results.xlsx")

        # Progress
        self.pdf_progress_bar = ctk.CTkProgressBar(parent)
        self.pdf_progress_bar.set(0)
        self.pdf_progress_bar.pack(fill="x", padx=24, pady=(16, 4))
        self.pdf_progress_label = ctk.CTkLabel(parent, text="0%  (0 / 0)", font=ctk.CTkFont(size=12))
        self.pdf_progress_label.pack()

        # Log
        self.pdf_log_box = ctk.CTkTextbox(parent, height=140, corner_radius=8)
        self.pdf_log_box.pack(fill="both", expand=True, padx=24, pady=10)

        self.pdf_start_btn = ctk.CTkButton(parent, text="Start Extraction", height=44,
                                            font=ctk.CTkFont(size=15, weight="bold"),
                                            command=self._start_pdf_extraction)
        self.pdf_start_btn.pack(pady=(8, 24), fill="x", padx=24)

    def _path_row_save(self, parent, label, browse_cmd):
        ctk.CTkLabel(parent, text=label,
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(12, 6), padx=16, anchor="w")
        line = ctk.CTkFrame(parent, fg_color="transparent")
        line.pack(fill="x", padx=16, pady=(0, 14))
        entry = ctk.CTkEntry(line)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        ctk.CTkButton(line, text="Browse", width=80, command=lambda: browse_cmd(entry)).pack(side="left")
        return entry

    def _save_pdf_api_key(self):
        key = self.api_key_entry.get().strip()
        if not key:
            messagebox.showerror("Missing key", "Please enter an API key first.")
            return
        pdf_extractor.save_api_key(key)
        messagebox.showinfo("Saved", "API key saved.")

    def _browse_pdfs(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if paths:
            self.pdf_paths = list(paths)
            self.pdf_files_label.configure(
                text=f"{len(self.pdf_paths)} file(s) selected:\n" +
                     "\n".join(os.path.basename(p) for p in self.pdf_paths[:8]) +
                     ("\n..." if len(self.pdf_paths) > 8 else ""))

    def _clear_pdfs(self):
        self.pdf_paths = []
        self.pdf_files_label.configure(text="No files selected.")

    def _browse_pdf_excel_save(self, entry):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)

    def _start_pdf_extraction(self):
        if getattr(self, "pdf_thread", None) and self.pdf_thread.is_alive():
            messagebox.showinfo("Busy", "An extraction is already running.")
            return

        api_key = self.api_key_entry.get().strip()
        prompt = self.pdf_prompt_box.get("1.0", "end").strip()
        save_path = self.pdf_save_path.get().strip()

        if not api_key:
            messagebox.showerror("Missing info", "Please enter and save your OpenAI API key.")
            return
        if not self.pdf_paths:
            messagebox.showerror("Missing info", "Please choose at least one PDF file.")
            return
        if not prompt:
            messagebox.showerror("Missing info", "Please enter an extraction prompt.")
            return
        if not save_path:
            messagebox.showerror("Missing info", "Please choose where to save the results.")
            return

        self.pdf_start_btn.configure(state="disabled", text="Extracting...")
        self.pdf_progress_bar.set(0)
        self.pdf_progress_label.configure(text="0%  (0 / 0)")
        self.pdf_log_box.delete("1.0", "end")

        self.pdf_queue = queue.Queue()
        self.pdf_thread = threading.Thread(
            target=self._run_pdf_job,
            args=(list(self.pdf_paths), api_key, prompt, save_path),
            daemon=True)
        self.pdf_thread.start()

    def _run_pdf_job(self, pdf_paths, api_key, prompt, save_path):
        def progress_cb(done, total):
            self.pdf_queue.put(("progress", (done, total)))

        def log_cb(msg):
            self.pdf_queue.put(("warn", msg))

        self.pdf_queue.put(("info", f"Processing {len(pdf_paths)} PDF(s)..."))
        ok, failed = pdf_extractor.process_pdfs(
            pdf_paths, api_key, prompt, save_path,
            progress_cb=progress_cb, log_cb=log_cb)
        summary = f"Done. {ok} succeeded, {failed} failed."
        if ok:
            summary += f" Results saved to {save_path}"
        self.pdf_queue.put(("info", summary))
        self.pdf_queue.put(("done", None))

    def _poll_pdf_queue(self):
        q = getattr(self, "pdf_queue", None)
        if q is not None:
            try:
                while True:
                    kind, payload = q.get_nowait()
                    if kind == "progress":
                        done, total = payload
                        pct = done / total if total else 0
                        self.pdf_progress_bar.set(pct)
                        self.pdf_progress_label.configure(text=f"{int(pct * 100)}%  ({done} / {total})")
                    elif kind in ("info", "warn", "error"):
                        self.pdf_log_box.insert("end", f"{payload}\n")
                        self.pdf_log_box.see("end")
                    elif kind == "done":
                        self.pdf_start_btn.configure(state="normal", text="Start Extraction")
            except queue.Empty:
                pass
        self.after(150, self._poll_pdf_queue)


    # ================================================================
    # PAGE 3: Data Profiler  (3-step flow: Ingestion -> Audit -> Export)
    # ================================================================
    PROFILER_STAGES = [
        ("upload", "Ingestion Target"),
        ("audit", "Dashboard Audit"),
        ("export", "Export Package"),
    ]

    def _build_data_profiler_page(self, parent):
        self.profiler_df = None              # working DataFrame at the Audit stage
        self.profiler_cleaned_df = None       # result after Execute Clean Data Vector
        self.profiler_before_results = None   # analyze_data() snapshot right after upload
        self.profiler_after_results = None    # analyze_data() snapshot right after cleaning
        self.profiler_results = None          # whatever is currently shown in the Audit stage
        self.profiler_filepath = None
        self.profiler_filter_vars = {}
        self.profiler_chart_canvas = None
        self.profiler_stage = "upload"
        self.profiler_max_reached = 0  # index into PROFILER_STAGES the user has unlocked

        ctk.CTkLabel(parent, text="📊 Data Profiling & Processing Suite",
                     font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(25, 2), anchor="w", padx=24)
        ctk.CTkLabel(parent, text="Upload a file, audit and clean it, then export the result",
                     font=ctk.CTkFont(size=13), text_color="gray").pack(pady=(0, 15), anchor="w", padx=24)

        wrapper = ctk.CTkFrame(parent, fg_color="transparent")
        wrapper.pack(fill="both", expand=True, padx=24, pady=(0, 24))
        wrapper.grid_columnconfigure(1, weight=1)
        wrapper.grid_rowconfigure(0, weight=1)

        # --- Stage rail (left) ---
        rail = ctk.CTkFrame(wrapper, corner_radius=12, width=190)
        rail.grid(row=0, column=0, sticky="ns", padx=(0, 12))
        rail.grid_propagate(False)
        ctk.CTkLabel(rail, text="SYSTEM STAGE", font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="gray").pack(pady=(16, 10), padx=14, anchor="w")
        self.profiler_stage_buttons = {}
        for key, label in self.PROFILER_STAGES:
            btn = ctk.CTkButton(
                rail, text=f"⚪  {label}", anchor="w", height=38,
                fg_color="transparent", hover_color=("gray80", "gray25"),
                font=ctk.CTkFont(size=13),
                command=lambda k=key: self._go_to_profiler_stage(k))
            btn.pack(fill="x", padx=10, pady=3)
            self.profiler_stage_buttons[key] = btn

        # --- Content area (right): one frame per stage, swapped in/out ---
        content = ctk.CTkFrame(wrapper, fg_color="transparent")
        content.grid(row=0, column=1, sticky="nsew")
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(0, weight=1)

        self.profiler_stage_frames = {
            "upload": self._build_profiler_upload_stage(content),
            "audit": self._build_profiler_audit_stage(content),
            "export": self._build_profiler_export_stage(content),
        }

        self._go_to_profiler_stage("upload", force=True)

    # ---------------- Stage navigation ----------------

    def _go_to_profiler_stage(self, key, force=False):
        stage_index = [k for k, _ in self.PROFILER_STAGES].index(key)
        if not force and stage_index > self.profiler_max_reached:
            return  # rail clicks can't skip ahead of what's actually been done
        self.profiler_stage = key
        for frame in self.profiler_stage_frames.values():
            frame.grid_forget()
        self.profiler_stage_frames[key].grid(row=0, column=0, sticky="nsew")
        self._refresh_profiler_stage_rail()

    def _refresh_profiler_stage_rail(self):
        for i, (key, label) in enumerate(self.PROFILER_STAGES):
            if key == self.profiler_stage:
                icon = "🟢"
            elif i < self.profiler_max_reached:
                icon = "✅"
            else:
                icon = "⚪"
            self.profiler_stage_buttons[key].configure(text=f"{icon}  {label}")

    # ---------------- Stage 1: Ingestion (upload) ----------------

    def _build_profiler_upload_stage(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")

        card = ctk.CTkFrame(frame, corner_radius=12)
        card.pack(fill="x", pady=10)
        ctk.CTkLabel(card, text="Step 1 — Ingestion Target",
                     font=ctk.CTkFont(weight="bold", size=16)).pack(pady=(16, 4), padx=16, anchor="w")
        ctk.CTkLabel(card, text="Select the CSV or Excel file you want to profile and clean.",
                     font=ctk.CTkFont(size=12), text_color="gray").pack(padx=16, anchor="w", pady=(0, 10))
        self.profiler_upload_entry = self._path_row(card, "Data file", self._browse_profiler_file)

        self.profiler_upload_status = ctk.CTkLabel(
            frame, text="", font=ctk.CTkFont(size=12), text_color="gray")
        self.profiler_upload_status.pack(anchor="w", pady=(4, 10))

        self.profiler_continue_to_audit_btn = ctk.CTkButton(
            frame, text="Continue to Audit  →", height=42, state="disabled",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=lambda: self._go_to_profiler_stage("audit", force=True))
        self.profiler_continue_to_audit_btn.pack(fill="x", pady=(6, 0))

        return frame

    def _browse_profiler_file(self, entry):
        path = filedialog.askopenfilename(filetypes=[("Data files", "*.csv *.xlsx *.xls")])
        if not path:
            return
        entry.delete(0, "end")
        entry.insert(0, path)

        try:
            df = data_profiler.load_data(path)
        except ValueError as e:
            messagebox.showerror("Could not load file", str(e))
            return

        self.profiler_filepath = path
        self.profiler_df = df
        self.profiler_before_results = data_profiler.analyze_data(df)
        self.profiler_max_reached = max(self.profiler_max_reached, 1)

        self.profiler_upload_status.configure(
            text=f"Loaded {len(df):,} record(s), {len(df.columns)} column(s).",
            text_color="#2FD675")
        self.profiler_continue_to_audit_btn.configure(state="normal")

        self._refresh_profiler_audit_view()
        self._refresh_profiler_stage_rail()

    # ---------------- Stage 2: Dashboard Audit ----------------

    def _build_profiler_audit_stage(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")

        kpi_row = ctk.CTkFrame(frame, fg_color="transparent")
        kpi_row.pack(fill="x", pady=(10, 10))
        self.profiler_kpi_labels = {}
        for key, label in [("total_records", "RECORDS"), ("total_columns", "COLUMNS"),
                            ("total_anomalies", "ANOMALIES"), ("health_pct", "HEALTH")]:
            card = ctk.CTkFrame(kpi_row, corner_radius=12)
            card.pack(side="left", expand=True, fill="both", padx=6)
            value_lbl = ctk.CTkLabel(card, text="--", font=ctk.CTkFont(size=22, weight="bold"))
            value_lbl.pack(pady=(14, 0))
            ctk.CTkLabel(card, text=label, font=ctk.CTkFont(size=11),
                         text_color="gray").pack(pady=(0, 14))
            self.profiler_kpi_labels[key] = value_lbl

        body = ctk.CTkFrame(frame, fg_color="transparent")
        body.pack(fill="both", expand=True, pady=(0, 10))
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)

        filter_frame = ctk.CTkFrame(body, corner_radius=12)
        filter_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        ctk.CTkLabel(filter_frame, text="SELECT OPTIMIZATION FILTERS",
                     font=ctk.CTkFont(weight="bold", size=13)).pack(pady=(14, 8), padx=16, anchor="w")
        self.profiler_checkbox_container = ctk.CTkFrame(filter_frame, fg_color="transparent")
        self.profiler_checkbox_container.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        ctk.CTkLabel(self.profiler_checkbox_container, text="Upload a file in Step 1 first.",
                     text_color="gray").pack(anchor="w")

        chart_frame = ctk.CTkFrame(body, corner_radius=12)
        chart_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        ctk.CTkLabel(chart_frame, text="ANOMALIES DISTRIBUTION MAP",
                     font=ctk.CTkFont(weight="bold", size=13)).pack(pady=(14, 8), padx=16, anchor="w")
        self.profiler_chart_container = ctk.CTkFrame(chart_frame, fg_color="transparent")
        self.profiler_chart_container.pack(fill="both", expand=True, padx=8, pady=(0, 16))

        nav_row = ctk.CTkFrame(frame, fg_color="transparent")
        nav_row.pack(fill="x")
        ctk.CTkButton(nav_row, text="←  Back to Ingestion", width=160, fg_color="transparent",
                      border_width=1, text_color=("gray20", "gray80"),
                      command=lambda: self._go_to_profiler_stage("upload")).pack(side="left")
        self.profiler_execute_btn = ctk.CTkButton(
            nav_row, text="✨ Execute Clean Data Vector", height=44,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#2FD675", hover_color="#25B563", text_color="#0A0A0A",
            command=self._on_profiler_execute_clean)
        self.profiler_execute_btn.pack(side="left", fill="x", expand=True, padx=(10, 0))

        return frame

    def _refresh_profiler_audit_view(self):
        self.profiler_results = data_profiler.analyze_data(self.profiler_df)
        self._rebuild_profiler_filters()
        self._update_profiler_kpis(self.profiler_results)
        self._redraw_profiler_chart()

    def _rebuild_profiler_filters(self):
        for child in self.profiler_checkbox_container.winfo_children():
            child.destroy()
        self.profiler_filter_vars = {}

        for key, label in data_profiler.ANOMALY_LABELS.items():
            count = self.profiler_results.get(key, 0)
            var = ctk.BooleanVar(value=count > 0)
            self.profiler_filter_vars[key] = var
            cb = ctk.CTkCheckBox(
                self.profiler_checkbox_container, text=f"{label} ({count} cases)",
                variable=var, command=self._redraw_profiler_chart)
            cb.pack(anchor="w", pady=6)
            if count == 0:
                cb.configure(state="disabled")

    def _update_profiler_kpis(self, results):
        self.profiler_kpi_labels["total_records"].configure(text=f'{results["total_records"]:,}')
        self.profiler_kpi_labels["total_columns"].configure(text=str(results["total_columns"]))
        self.profiler_kpi_labels["total_anomalies"].configure(text=str(results["total_anomalies"]))
        health = results["health_pct"]
        color = "#2FD675" if health >= 85 else ("#e0a020" if health >= 60 else "#ff5d5d")
        self.profiler_kpi_labels["health_pct"].configure(text=f"{health}%", text_color=color)

    def _redraw_profiler_chart(self):
        if self.profiler_chart_canvas is not None:
            self.profiler_chart_canvas.get_tk_widget().destroy()
            self.profiler_chart_canvas = None

        active = {k: v.get() for k, v in self.profiler_filter_vars.items()}
        fig = render_donut(self.profiler_results, active)
        self.profiler_chart_canvas = FigureCanvasTkAgg(fig, master=self.profiler_chart_container)
        self.profiler_chart_canvas.draw()
        self.profiler_chart_canvas.get_tk_widget().pack(fill="both", expand=True)

    def _on_profiler_execute_clean(self):
        if self.profiler_df is None:
            messagebox.showwarning("No data loaded", "Select a CSV or Excel file first.")
            return

        active_filters = {k: v.get() for k, v in self.profiler_filter_vars.items()}
        if not any(active_filters.values()):
            messagebox.showinfo("Nothing selected", "Check at least one filter to clean.")
            return

        cleaned = data_profiler.clean_data(self.profiler_df, active_filters)
        self.profiler_cleaned_df = cleaned
        self.profiler_after_results = data_profiler.analyze_data(cleaned)
        self.profiler_max_reached = max(self.profiler_max_reached, 2)

        self._update_profiler_export_comparison()
        self.profiler_export_status.configure(text="", text_color="gray")
        self._go_to_profiler_stage("export", force=True)

    # ---------------- Stage 3: Export Package ----------------

    def _build_profiler_export_stage(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")

        ctk.CTkLabel(frame, text="Step 3 — Export Package",
                     font=ctk.CTkFont(weight="bold", size=16)).pack(pady=(10, 4), anchor="w")
        ctk.CTkLabel(frame, text="Compare before/after results and save your cleaned file.",
                     font=ctk.CTkFont(size=12), text_color="gray").pack(anchor="w", pady=(0, 14))

        compare_row = ctk.CTkFrame(frame, fg_color="transparent")
        compare_row.pack(fill="x", pady=(0, 16))
        self.profiler_compare_labels = {}
        for key, label in [("total_records", "RECORDS"), ("total_columns", "COLUMNS"),
                            ("total_anomalies", "ANOMALIES"), ("health_pct", "HEALTH")]:
            card = ctk.CTkFrame(compare_row, corner_radius=12)
            card.pack(side="left", expand=True, fill="both", padx=6)
            ctk.CTkLabel(card, text=label, font=ctk.CTkFont(size=11),
                         text_color="gray").pack(pady=(14, 2))
            value_lbl = ctk.CTkLabel(card, text="-- → --", font=ctk.CTkFont(size=15, weight="bold"))
            value_lbl.pack(pady=(0, 4))
            delta_lbl = ctk.CTkLabel(card, text="", font=ctk.CTkFont(size=11))
            delta_lbl.pack(pady=(0, 14))
            self.profiler_compare_labels[key] = (value_lbl, delta_lbl)

        self.profiler_export_status = ctk.CTkLabel(
            frame, text="", font=ctk.CTkFont(size=12), text_color="gray",
            wraplength=650, justify="left")
        self.profiler_export_status.pack(anchor="w", pady=(0, 10))

        nav_row = ctk.CTkFrame(frame, fg_color="transparent")
        nav_row.pack(fill="x", pady=(6, 0))
        ctk.CTkButton(nav_row, text="←  Back to Audit", width=140, fg_color="transparent",
                      border_width=1, text_color=("gray20", "gray80"),
                      command=lambda: self._go_to_profiler_stage("audit")).pack(side="left")
        ctk.CTkButton(nav_row, text="Start Over", width=110, fg_color="transparent",
                      border_width=1, text_color=("gray20", "gray80"),
                      command=self._reset_profiler_flow).pack(side="left", padx=(8, 0))
        ctk.CTkButton(
            nav_row, text="💾  Save Cleaned File", height=44,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2FD675", hover_color="#25B563", text_color="#0A0A0A",
            command=self._on_profiler_save_export).pack(side="left", fill="x", expand=True, padx=(10, 0))

        return frame

    def _update_profiler_export_comparison(self):
        before = self.profiler_before_results
        after = self.profiler_after_results
        for key in ("total_records", "total_columns", "total_anomalies", "health_pct"):
            value_lbl, delta_lbl = self.profiler_compare_labels[key]
            b, a = before[key], after[key]
            suffix = "%" if key == "health_pct" else ""
            value_lbl.configure(text=f"{b:,}{suffix}  →  {a:,}{suffix}")

            diff = a - b
            if key == "total_anomalies":
                good = diff <= 0
            elif key == "health_pct":
                good = diff >= 0
            else:
                good = True  # record/column count changes are neutral, not good/bad
            if diff == 0:
                delta_lbl.configure(text="no change", text_color="gray")
            else:
                sign = "+" if diff > 0 else ""
                delta_lbl.configure(text=f"{sign}{diff}{suffix}",
                                     text_color=("#2FD675" if good else "#ff5d5d"))

    def _on_profiler_save_export(self):
        if self.profiler_cleaned_df is None:
            messagebox.showwarning("Nothing to export", "Clean the data in the Audit step first.")
            return

        default_name = "cleaned_" + os.path.basename(self.profiler_filepath or "export.csv")
        out_path = filedialog.asksaveasfilename(
            title="Save cleaned data as", initialfile=default_name,
            defaultextension=".csv", filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx")])
        if not out_path:
            return

        try:
            data_profiler.save_data(self.profiler_cleaned_df, out_path)
        except ValueError as e:
            messagebox.showerror("Could not save file", str(e))
            return

        self.profiler_export_status.configure(
            text=f"Saved {len(self.profiler_cleaned_df):,} record(s) to:\n{out_path}",
            text_color="#2FD675")

    def _reset_profiler_flow(self):
        self.profiler_df = None
        self.profiler_cleaned_df = None
        self.profiler_before_results = None
        self.profiler_after_results = None
        self.profiler_results = None
        self.profiler_filepath = None
        self.profiler_max_reached = 0

        self.profiler_upload_entry.delete(0, "end")
        self.profiler_upload_status.configure(text="")
        self.profiler_continue_to_audit_btn.configure(state="disabled")
        self.profiler_export_status.configure(text="")

        self._go_to_profiler_stage("upload", force=True)


if __name__ == "__main__":
    app = App()
    app.mainloop()
