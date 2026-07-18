"""
app.py
Modern desktop UI for LinkHarvest (Windows).
Run with:  python app.py
"""
import os
import threading
import queue
import webbrowser

import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image

from downloader import load_sheet, build_download_tasks, run_downloads
from donation import generate_qr_image
from updater import check_for_update
from version import __version__

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

DEFAULT_THREADS = 10
DEFAULT_TIMEOUT = 30

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ICON_PATH = os.path.join(BASE_DIR, "assets", "icon.ico")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"LinkHarvest v{__version__}")
        self.geometry("740x840")  # Slightly taller to cleanly fit the padded layout
        self.minsize(700, 780)
        try:
            self.iconbitmap(ICON_PATH)
        except Exception:
            pass

        self.log_queue = queue.Queue()
        self.download_thread = None

        self._build_ui()
        self.after(150, self._poll_log_queue)
        threading.Thread(target=self._check_update_background, daemon=True).start()

    # ---------------- UI ----------------
    def _build_ui(self):
        ctk.CTkLabel(self, text="🔗 LinkHarvest",
                     font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(25, 2))
        ctk.CTkLabel(self, text="Download images / videos / audio from links in your Excel sheet",
                     font=ctk.CTkFont(size=13), text_color="gray").pack(pady=(0, 15))

        # File section
        file_frame = ctk.CTkFrame(self, corner_radius=12)
        file_frame.pack(fill="x", padx=24, pady=10)

        self.excel_path = self._path_row(file_frame, "Excel file", self._browse_excel)
        self.excel_path.insert(0, r"C:\Users\Harsh\Documents\links.xlsx")  # Clean matching path

        self.sheet_name = self._entry_row(file_frame, "Sheet name", "Sheet4")

        self.folder_path = self._path_row(file_frame, "Save folder", self._browse_folder)
        self.folder_path.insert(0, r"C:\Users\Harsh\Downloads\HarvestedFiles")  # Clean matching folder

        # Row / column range section
        range_frame = ctk.CTkFrame(self, corner_radius=12)
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

        # Advanced section
        adv_frame = ctk.CTkFrame(self, corner_radius=12)
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

        # Support / donation section
        donate_frame = ctk.CTkFrame(self, corner_radius=12)
        donate_frame.pack(fill="x", padx=24, pady=10)
        ctk.CTkLabel(donate_frame, text="❤ Support Development (optional)",
                     font=ctk.CTkFont(weight="bold", size=14)).pack(pady=(12, 4))
        ctk.CTkLabel(donate_frame, text="If this tool saved you time, you can scan to leave a UPI donation.",
                     font=ctk.CTkFont(size=11), text_color="gray").pack()
        self.qr_label = ctk.CTkLabel(donate_frame, text="")
        self.qr_label.pack(pady=12)
        self._load_donation_qr()

        # Progress
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=24, pady=(16, 4))
        self.progress_label = ctk.CTkLabel(self, text="0%  (0 / 0)", font=ctk.CTkFont(size=12))
        self.progress_label.pack()

        # Log box
        self.log_box = ctk.CTkTextbox(self, height=140, corner_radius=8)
        self.log_box.pack(fill="both", expand=True, padx=24, pady=10)

        # Action Button
        self.start_btn = ctk.CTkButton(self, text="Start Download", height=44,
                                        font=ctk.CTkFont(size=15, weight="bold"),
                                        command=self._start_download)
        self.start_btn.pack(pady=(8, 20), fill="x", padx=24)

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

        # Make it modal: blocks interaction with the main window until resolved.
        dialog.transient(self)
        dialog.grab_set()
        # Closing the dialog's X button quits the whole app instead of silently
        # letting the user bypass the update.
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
            # After sending them to the download page, close this running copy —
            # they need to install and launch the new .exe to continue.
            self.destroy()

        ctk.CTkButton(dialog, text="⬇  Download Update", height=44,
                      font=ctk.CTkFont(size=14, weight="bold"),
                      command=do_download).pack(pady=(0, 8), padx=32, fill="x")
        ctk.CTkButton(dialog, text="Quit", height=32, fg_color="transparent",
                      border_width=1, text_color=("gray20", "gray80"),
                      command=self.destroy).pack(padx=32, fill="x")

        dialog.grab_set()  # re-assert focus after widgets are laid out

    def _check_advanced_changed(self, _event=None):
        changed = (self.threads_entry.get() != str(DEFAULT_THREADS) or
                   self.timeout_entry.get() != str(DEFAULT_TIMEOUT))
        if changed:
            self.adv_warning.configure(
                text="⚠ Default settings changed — with more threads or a shorter timeout, "
                     "some links may fail or remain un-downloaded.")
        else:
            self.adv_warning.configure(text="")

    # ---------------- Logic ----------------
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


if __name__ == "__main__":
    app = App()
    app.mainloop()
