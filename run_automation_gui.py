import os
import json
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

import pdf_parser
import billsync_formatter
import tracker_manager


# ── Persist user settings next to the EXE ─────────────────────────────────────
def _config_path():
    """
    Store config.json beside the running EXE (when frozen by PyInstaller)
    or beside the .py script (when run directly).
    """
    import sys
    if getattr(sys, "frozen", False):
        # Running as a PyInstaller EXE - use the EXE actual folder
        base = os.path.dirname(sys.executable)
    else:
        # Running as a plain .py script
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "quovant_config.json")

def load_config() -> dict:
    try:
        with open(_config_path(), "r") as f:
            return json.load(f)
    except Exception:
        return {}

def save_config(cfg: dict):
    try:
        with open(_config_path(), "w") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass


# ── Colours & Fonts ───────────────────────────────────────────────────────────
BG         = "#0f1117"
PANEL      = "#1a1d27"
ACCENT     = "#4f8ef7"
ACCENT2    = "#2ecc71"
DANGER     = "#e74c3c"
TEXT       = "#e8eaf0"
SUBTEXT    = "#7a7f99"
BORDER     = "#2a2d3e"
FONT_TITLE = ("Georgia", 20, "bold")
FONT_SUB   = ("Georgia", 10, "italic")
FONT_BODY  = ("Courier New", 10)
FONT_BTN   = ("Courier New", 10, "bold")
FONT_LABEL = ("Courier New", 9)


class AutomationGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("QUOVANT · BillSync Automation")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.geometry("900x780")
        self.minsize(760, 620)

        self.selected_pdfs: list[str] = []
        self.is_running = False

        # Load saved paths (empty strings if first run)
        cfg = load_config()
        self.output_dir    = tk.StringVar(value=cfg.get("output_dir",    ""))
        self.processed_dir = tk.StringVar(value=cfg.get("processed_dir", ""))

        self._build_ui()

    # ── UI Construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=BG, pady=18)
        hdr.pack(fill="x", padx=30)

        tk.Label(hdr, text="QUOVANT", font=FONT_TITLE,
                 bg=BG, fg=ACCENT).pack(anchor="w")
        tk.Label(hdr, text="BillSync Automation  ·  Manual PDF Mode",
                 font=FONT_SUB, bg=BG, fg=SUBTEXT).pack(anchor="w")

        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=30)

        # ── Main Body ─────────────────────────────────────────────────────────
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=30, pady=16)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(2, weight=1)   # log panel expands

        # ── Folder Settings Panel ──────────────────────────────────────────────
        settings_panel = tk.Frame(body, bg=PANEL, bd=0,
                                  highlightbackground=BORDER, highlightthickness=1)
        settings_panel.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        settings_panel.columnconfigure(1, weight=1)

        tk.Label(settings_panel, text="  FOLDER SETTINGS",
                 font=FONT_LABEL, bg=PANEL, fg=SUBTEXT,
                 anchor="w", pady=8).grid(row=0, column=0, columnspan=3, sticky="ew")

        # Output folder row
        tk.Label(settings_panel, text="  Output Folder:",
                 font=FONT_LABEL, bg=PANEL, fg=TEXT,
                 width=18, anchor="w").grid(row=1, column=0, padx=(10, 0), pady=4, sticky="w")

        tk.Entry(settings_panel, textvariable=self.output_dir,
                 bg="#12151f", fg=TEXT, insertbackground=ACCENT,
                 font=FONT_BODY, bd=0, highlightthickness=1,
                 highlightbackground=BORDER, highlightcolor=ACCENT
                 ).grid(row=1, column=1, sticky="ew", padx=8, pady=4)

        self._btn(settings_panel, "Browse", ACCENT,
                  lambda: self._browse_folder(self.output_dir, "Select Output Folder"),
                  width=8).grid(row=1, column=2, padx=(0, 10), pady=4)

        # Processed folder row
        tk.Label(settings_panel, text="  Processed Folder:",
                 font=FONT_LABEL, bg=PANEL, fg=TEXT,
                 width=18, anchor="w").grid(row=2, column=0, padx=(10, 0), pady=(0, 10), sticky="w")

        tk.Entry(settings_panel, textvariable=self.processed_dir,
                 bg="#12151f", fg=TEXT, insertbackground=ACCENT,
                 font=FONT_BODY, bd=0, highlightthickness=1,
                 highlightbackground=BORDER, highlightcolor=ACCENT
                 ).grid(row=2, column=1, sticky="ew", padx=8, pady=(0, 10))

        self._btn(settings_panel, "Browse", ACCENT,
                  lambda: self._browse_folder(self.processed_dir, "Select Processed Folder"),
                  width=8).grid(row=2, column=2, padx=(0, 10), pady=(0, 10))

        tk.Label(settings_panel,
                 text="  i  Paths are saved automatically and remembered next time.",
                 font=FONT_LABEL, bg=PANEL, fg=SUBTEXT, anchor="w", pady=4
                 ).grid(row=3, column=0, columnspan=3, sticky="ew")

        # ── File Selection Panel ───────────────────────────────────────────────
        file_panel = tk.Frame(body, bg=PANEL, bd=0,
                              highlightbackground=BORDER, highlightthickness=1)
        file_panel.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        file_panel.columnconfigure(0, weight=1)

        tk.Label(file_panel, text="  SELECTED PDF FILES",
                 font=FONT_LABEL, bg=PANEL, fg=SUBTEXT,
                 anchor="w", pady=8).grid(row=0, column=0, columnspan=3, sticky="ew")

        list_frame = tk.Frame(file_panel, bg=PANEL)
        list_frame.grid(row=1, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 8))
        list_frame.columnconfigure(0, weight=1)

        self.file_listbox = tk.Listbox(
            list_frame, height=5,
            bg="#12151f", fg=TEXT, selectbackground=ACCENT,
            font=FONT_BODY, bd=0, highlightthickness=0,
            selectforeground="white", activestyle="none"
        )
        self.file_listbox.grid(row=0, column=0, sticky="ew")

        sb = tk.Scrollbar(list_frame, orient="vertical",
                          command=self.file_listbox.yview, bg=PANEL)
        sb.grid(row=0, column=1, sticky="ns")
        self.file_listbox.config(yscrollcommand=sb.set)

        btn_row = tk.Frame(file_panel, bg=PANEL)
        btn_row.grid(row=2, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 10))

        self._btn(btn_row, "+ Add PDFs",  ACCENT,  self._add_pdfs).pack(side="left", padx=(0, 8))
        self._btn(btn_row, "x Remove",    SUBTEXT, self._remove_selected).pack(side="left", padx=(0, 8))
        self._btn(btn_row, "x Clear All", DANGER,  self._clear_files).pack(side="left")

        self.file_count_label = tk.Label(btn_row, text="0 files selected",
                                         font=FONT_LABEL, bg=PANEL, fg=SUBTEXT)
        self.file_count_label.pack(side="right", padx=4)

        # ── Log Panel ─────────────────────────────────────────────────────────
        log_panel = tk.Frame(body, bg=PANEL, bd=0,
                             highlightbackground=BORDER, highlightthickness=1)
        log_panel.grid(row=2, column=0, sticky="nsew", pady=(0, 12))
        log_panel.columnconfigure(0, weight=1)
        log_panel.rowconfigure(1, weight=1)

        tk.Label(log_panel, text="  PROCESSING LOG",
                 font=FONT_LABEL, bg=PANEL, fg=SUBTEXT,
                 anchor="w", pady=8).grid(row=0, column=0, sticky="ew")

        log_frame = tk.Frame(log_panel, bg=PANEL)
        log_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_frame, bg="#0a0c12", fg="#a0f0a0",
            font=FONT_BODY, bd=0, highlightthickness=0,
            wrap="word", state="disabled", insertbackground=ACCENT
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

        log_sb = tk.Scrollbar(log_frame, orient="vertical",
                              command=self.log_text.yview, bg=PANEL)
        log_sb.grid(row=0, column=1, sticky="ns")
        self.log_text.config(yscrollcommand=log_sb.set)

        # ── Progress Bar ───────────────────────────────────────────────────────
        self.progress = ttk.Progressbar(body, mode="indeterminate",
                                        style="Accent.Horizontal.TProgressbar")
        self.progress.grid(row=3, column=0, sticky="ew", pady=(0, 12))

        # ── Action Row ────────────────────────────────────────────────────────
        action_row = tk.Frame(body, bg=BG)
        action_row.grid(row=4, column=0, sticky="ew")

        self.run_btn = self._btn(
            action_row, "RUN AUTOMATION", ACCENT2, self._run_automation,
            width=22, font=("Courier New", 11, "bold")
        )
        self.run_btn.pack(side="left", padx=(0, 12))

        self._btn(action_row, "Open Output Folder", SUBTEXT,
                  self._open_output_folder).pack(side="left")

        self.status_label = tk.Label(action_row, text="Ready.",
                                     font=FONT_LABEL, bg=BG, fg=SUBTEXT)
        self.status_label.pack(side="right")

        # Progressbar style
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Accent.Horizontal.TProgressbar",
                        troughcolor=PANEL, background=ACCENT, thickness=6)

    # ── Widget Helper ──────────────────────────────────────────────────────────

    def _btn(self, parent, text, color, cmd, width=14, font=FONT_BTN) -> tk.Button:
        return tk.Button(
            parent, text=text, command=cmd,
            bg=color, fg="white" if color != SUBTEXT else BG,
            font=font, bd=0, padx=14, pady=6,
            relief="flat", cursor="hand2",
            activebackground=color, activeforeground="white",
            width=width
        )

    # ── Folder Browsing ────────────────────────────────────────────────────────

    def _browse_folder(self, var: tk.StringVar, title: str):
        chosen = filedialog.askdirectory(title=title)
        if chosen:
            var.set(chosen)
            self._save_current_config()

    def _save_current_config(self):
        save_config({
            "output_dir":    self.output_dir.get(),
            "processed_dir": self.processed_dir.get(),
        })

    def _validate_folders(self) -> bool:
        if not self.output_dir.get().strip():
            messagebox.showwarning("Missing Folder",
                                   "Please select an Output Folder before running.")
            return False
        if not self.processed_dir.get().strip():
            messagebox.showwarning("Missing Folder",
                                   "Please select a Processed Folder before running.")
            return False
        return True

    # ── Helpers ────────────────────────────────────────────────────────────────

    def _log(self, msg: str):
        def _insert():
            self.log_text.config(state="normal")
            ts = datetime.now().strftime("%H:%M:%S")
            self.log_text.insert("end", f"[{ts}]  {msg}\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.after(0, _insert)

    def _set_status(self, msg: str):
        self.after(0, lambda: self.status_label.config(text=msg))

    def _update_file_count(self):
        n = len(self.selected_pdfs)
        self.file_count_label.config(text=f"{n} file{'s' if n != 1 else ''} selected")

    # ── File Management ────────────────────────────────────────────────────────

    def _add_pdfs(self):
        files = filedialog.askopenfilenames(
            title="Select PDF Invoice(s)",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        for f in files:
            if f not in self.selected_pdfs:
                self.selected_pdfs.append(f)
                self.file_listbox.insert(
                    "end", f"  {os.path.basename(f)}   [{os.path.dirname(f)}]"
                )
        self._update_file_count()

    def _remove_selected(self):
        for i in reversed(self.file_listbox.curselection()):
            self.file_listbox.delete(i)
            self.selected_pdfs.pop(i)
        self._update_file_count()

    def _clear_files(self):
        self.file_listbox.delete(0, "end")
        self.selected_pdfs.clear()
        self._update_file_count()

    def _open_output_folder(self):
        path = self.output_dir.get().strip()
        if path and os.path.exists(path):
            os.startfile(path)
        else:
            messagebox.showinfo("Not Found",
                                "Output folder is not set or does not exist yet.")

    # ── Core Automation ────────────────────────────────────────────────────────

    def _run_automation(self):
        if self.is_running:
            return
        if not self.selected_pdfs:
            messagebox.showwarning("No Files", "Please add at least one PDF before running.")
            return
        if not self._validate_folders():
            return

        os.makedirs(self.output_dir.get(),    exist_ok=True)
        os.makedirs(self.processed_dir.get(), exist_ok=True)
        self._save_current_config()

        self.is_running = True
        self.run_btn.config(state="disabled", text="Running...")
        self.progress.start(12)
        self._set_status("Processing...")

        threading.Thread(target=self._automation_worker, daemon=True).start()

    def _automation_worker(self):
        output_dir    = self.output_dir.get().strip()
        processed_dir = self.processed_dir.get().strip()
        timestamp     = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_excel  = os.path.join(output_dir, f"BillSync_Output_{timestamp}.xlsx")
        tracker_path  = os.path.join(output_dir, "Master_Tracker.xlsx")

        try:
            self._log("=" * 58)
            self._log(f"Starting automation for {len(self.selected_pdfs)} file(s).")
            self._log(f"Output folder   : {output_dir}")
            self._log(f"Processed folder: {processed_dir}")
            self._log("=" * 58)

            # ── Step 1: Parse PDFs ─────────────────────────────────────────────
            self._log("\n  STEP 1 - Parsing PDFs")
            all_items = []

            for pdf_path in self.selected_pdfs:
                fname = os.path.basename(pdf_path)
                self._log(f"   Parsing: {fname}")
                try:
                    items, metadata = pdf_parser.extract_invoice_data(pdf_path)
                    self._log(f"   OK  {len(items)} line items · Invoice: {metadata.get('invoice_number', 'N/A')}")
                    all_items.extend(items)

                    final_pdf_path = os.path.join(processed_dir, fname)
                    tracker_manager.update_tracker(tracker_path, metadata, final_pdf_path)
                    self._log("   OK  Tracker updated.")

                except Exception as e:
                    self._log(f"   ERROR  {fname}: {e}")

            self._log(f"\n   Total line items extracted: {len(all_items)}")

            # ── Step 2: Generate BillSync Report ───────────────────────────────
            self._log("\n  STEP 2 - Generating BillSync Report")
            if all_items:
                try:
                    billsync_formatter.format_to_billsync(all_items, None, output_excel)
                    self._log(f"   OK  Saved: {os.path.basename(output_excel)}")
                except Exception as e:
                    self._log(f"   ERROR  generating report: {e}")
            else:
                self._log("   WARNING  No data extracted. Skipping report.")

            # ── Step 3: Move PDFs to Processed ────────────────────────────────
            self._log("\n  STEP 3 - Moving PDFs to Processed Folder")
            for pdf_path in self.selected_pdfs:
                fname = os.path.basename(pdf_path)
                dest  = os.path.join(processed_dir, fname)
                try:
                    if os.path.exists(dest):
                        os.remove(dest)
                    shutil.move(pdf_path, dest)
                    self._log(f"   OK  Moved: {fname}")
                except Exception as e:
                    self._log(f"   ERROR  Could not move {fname}: {e}")

            # ── Done ───────────────────────────────────────────────────────────
            self._log("\n" + "=" * 58)
            self._log("DONE  Automation complete.")
            self._log("=" * 58)
            self._set_status("Done")

            self.after(0, lambda: messagebox.showinfo(
                "Complete",
                f"Automation finished successfully.\n\nOutput saved to:\n{output_excel}"
            ))
            self.after(0, self._clear_files)

        except Exception as e:
            self._log(f"\n  Fatal error: {e}")
            self._set_status("Error.")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))

        finally:
            self.after(0, self._reset_ui)

    def _reset_ui(self):
        self.progress.stop()
        self.run_btn.config(state="normal", text="RUN AUTOMATION")
        self.is_running = False


# ── Entry Point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = AutomationGUI()
    app.mainloop()
