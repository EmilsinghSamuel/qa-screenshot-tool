"""
QA Screenshot Capture Tool
Portable Windows application — no admin or installation required.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os
import sys
import json
import time
from datetime import datetime
from io import BytesIO

try:
    from PIL import ImageGrab
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    import keyboard
except ImportError as exc:
    import tkinter.messagebox as _mb
    _mb.showerror(
        "Missing dependencies",
        f"Required package missing: {exc}\n\n"
        "pip install Pillow python-docx keyboard",
    )
    sys.exit(1)


def app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


class Settings:
    _DEFAULTS = {
        "hotkey": "f1",
        "auto_timestamp": True,
        "auto_number": True,
        "prompt_description": False,
        "minimize_on_capture": True,
        "capture_delay": 0,
        "tester_name": "",
        "project_name": "",
        "save_path": "",
    }

    def __init__(self):
        self._path = os.path.join(app_dir(), "config.json")
        self._data: dict = dict(self._DEFAULTS)
        self._load()

    def _load(self):
        try:
            if os.path.exists(self._path):
                with open(self._path, "r") as fh:
                    self._data.update(json.load(fh))
        except Exception:
            pass

    def save(self):
        try:
            with open(self._path, "w") as fh:
                json.dump(self._data, fh, indent=2)
        except Exception:
            pass

    def __getitem__(self, key):
        return self._data.get(key, self._DEFAULTS.get(key))

    def __setitem__(self, key, value):
        self._data[key] = value


class QATool:
    APP_TITLE = "QA Screenshot Tool"
    WIN_SIZE  = "540x500"

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(self.APP_TITLE)
        self.root.geometry(self.WIN_SIZE)
        self.root.resizable(False, False)

        self.settings = Settings()

        # Session state
        self.active     = False
        self.document   = None
        self.doc_path   = ""
        self.count      = 0
        self.sess_start = None

        # FIX: lock prevents concurrent captures corrupting the document
        self._lock = threading.Lock()

        self._build_ui()
        self._restore_settings_to_ui()

    # ── UI ───────────────────────────────────────────────────────────────────

    def _build_ui(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        capture_tab  = ttk.Frame(nb, padding=12)
        settings_tab = ttk.Frame(nb, padding=12)
        about_tab    = ttk.Frame(nb, padding=12)

        nb.add(capture_tab,  text="  Capture  ")
        nb.add(settings_tab, text="  Settings  ")
        nb.add(about_tab,    text="  About  ")

        self._build_capture_tab(capture_tab)
        self._build_settings_tab(settings_tab)
        self._build_about_tab(about_tab)

    def _build_capture_tab(self, p):
        tk.Label(p, text="QA Screenshot Capture Tool",
                 font=("Segoe UI", 13, "bold")).pack(pady=(4, 2))
        tk.Label(p, text="Capture test screenshots directly into Word documents",
                 font=("Segoe UI", 9), fg="#555").pack(pady=(0, 10))

        doc = ttk.LabelFrame(p, text="Document", padding=8)
        doc.pack(fill="x", pady=4)

        row = tk.Frame(doc)
        row.pack(fill="x", pady=2)
        tk.Label(row, text="File name:", width=12, anchor="w").pack(side="left")
        self.fname_var = tk.StringVar(
            value=f"TestReport_{datetime.now().strftime('%Y%m%d')}")
        ttk.Entry(row, textvariable=self.fname_var, width=34).pack(side="left", padx=4)

        row2 = tk.Frame(doc)
        row2.pack(fill="x", pady=2)
        tk.Label(row2, text="Save to:", width=12, anchor="w").pack(side="left")
        self.save_path_var = tk.StringVar(value=os.path.expanduser("~/Desktop"))
        ttk.Entry(row2, textvariable=self.save_path_var, width=28).pack(side="left", padx=4)
        ttk.Button(row2, text="Browse", command=self._browse, width=7).pack(side="left")

        hk_frame = ttk.LabelFrame(p, text="Active hotkey", padding=6)
        hk_frame.pack(fill="x", pady=4)
        self.hotkey_hint_var = tk.StringVar()
        tk.Label(hk_frame, textvariable=self.hotkey_hint_var,
                 font=("Segoe UI", 10, "bold"), fg="#006400").pack()

        btn = tk.Frame(p)
        btn.pack(pady=10)
        self.start_btn = ttk.Button(btn, text="Start Session",
                                    command=self.start_session, width=20)
        self.start_btn.pack(side="left", padx=6)
        self.stop_btn  = ttk.Button(btn, text="Stop & Save",
                                    command=self.stop_session,
                                    state="disabled", width=18)
        self.stop_btn.pack(side="left", padx=6)

        sf = ttk.LabelFrame(p, text="Session status", padding=8)
        sf.pack(fill="x", pady=4)
        self.status_var = tk.StringVar(
            value="No active session. Enter a file name and press Start.")
        tk.Label(sf, textvariable=self.status_var, font=("Segoe UI", 9),
                 fg="#003399", wraplength=460).pack()
        self.count_var = tk.StringVar(value="")
        tk.Label(sf, textvariable=self.count_var,
                 font=("Segoe UI", 9, "bold")).pack()
        self.last_var = tk.StringVar(value="")
        tk.Label(sf, textvariable=self.last_var,
                 font=("Segoe UI", 8), fg="#666").pack()

    def _build_settings_tab(self, p):
        hk = ttk.LabelFrame(p, text="Hotkey", padding=8)
        hk.pack(fill="x", pady=6)

        r = tk.Frame(hk)
        r.pack(fill="x")
        tk.Label(r, text="Screenshot key:", width=18, anchor="w").pack(side="left")
        self.hotkey_var = tk.StringVar(value="f1")
        hotkeys = [
            "f1", "f2", "f3", "f4", "f5", "f6",
            "f7", "f8", "f9", "f10", "f11", "f12",
            "ctrl+shift+s", "ctrl+shift+c",
            "ctrl+alt+s",   "ctrl+alt+c",
            "print screen",
        ]
        ttk.Combobox(r, textvariable=self.hotkey_var,
                     values=hotkeys, width=20).pack(side="left", padx=4)
        ttk.Button(r, text="Apply", command=self._apply_hotkey,
                   width=8).pack(side="left", padx=4)
        tk.Label(hk, text="Takes effect immediately in an active session.",
                 font=("Segoe UI", 8), fg="#777").pack(anchor="w", pady=(4, 0))

        co = ttk.LabelFrame(p, text="Capture options", padding=8)
        co.pack(fill="x", pady=6)

        self.auto_ts_var  = tk.BooleanVar(value=True)
        self.auto_num_var = tk.BooleanVar(value=True)
        self.prompt_d_var = tk.BooleanVar(value=False)
        self.minimize_var = tk.BooleanVar(value=True)
        self.delay_var    = tk.StringVar(value="0")

        ttk.Checkbutton(co, text="Add timestamp below each screenshot",
                        variable=self.auto_ts_var).pack(anchor="w", pady=2)
        ttk.Checkbutton(co, text="Auto-number screenshots (Screenshot 1, 2, 3 ...)",
                        variable=self.auto_num_var).pack(anchor="w", pady=2)
        ttk.Checkbutton(co, text="Prompt for description / test-step before each capture",
                        variable=self.prompt_d_var).pack(anchor="w", pady=2)
        ttk.Checkbutton(co, text="Minimize window during capture (keeps screenshots clean)",
                        variable=self.minimize_var).pack(anchor="w", pady=2)

        dr = tk.Frame(co)
        dr.pack(anchor="w", pady=2)
        tk.Label(dr, text="Capture delay (0-5 sec):").pack(side="left")
        ttk.Spinbox(dr, from_=0, to=5, textvariable=self.delay_var,
                    width=5).pack(side="left", padx=6)

        dh = ttk.LabelFrame(p, text="Document header info", padding=8)
        dh.pack(fill="x", pady=6)

        r1 = tk.Frame(dh)
        r1.pack(fill="x", pady=2)
        tk.Label(r1, text="Tester name:", width=15, anchor="w").pack(side="left")
        self.tester_var = tk.StringVar()
        ttk.Entry(r1, textvariable=self.tester_var, width=28).pack(side="left", padx=4)

        r2 = tk.Frame(dh)
        r2.pack(fill="x", pady=2)
        tk.Label(r2, text="Project / App:", width=15, anchor="w").pack(side="left")
        self.project_var = tk.StringVar()
        ttk.Entry(r2, textvariable=self.project_var, width=28).pack(side="left", padx=4)

        ttk.Button(p, text="Save settings", command=self._save_settings,
                   width=16).pack(pady=8)

    def _build_about_tab(self, p):
        tk.Label(p, text="QA Screenshot Capture Tool",
                 font=("Segoe UI", 14, "bold")).pack(pady=(20, 4))
        tk.Label(p, text="Version 1.0  -  Portable, no installation required",
                 font=("Segoe UI", 9), fg="#555").pack()
        instructions = (
            "\n"
            "How to use\n"
            "----------\n"
            "1.  Enter a file name on the Capture tab.\n"
            "2.  Choose where to save the document.\n"
            "3.  Press  Start Session.\n"
            "4.  Switch to your application under test.\n"
            "5.  Press your hotkey (default F1) to capture a screenshot.\n"
            "    The tool minimizes briefly so it stays out of the screenshot.\n"
            "6.  Press  Stop & Save  when done.\n"
            "    The Word document opens automatically.\n\n"
            "Tips\n"
            "----\n"
            "* Change the hotkey on the Settings tab at any time.\n"
            "* Enable 'Prompt for description' to annotate each screenshot.\n"
            "* The config.json file next to the .exe stores your settings.\n"
            "* The exe is fully portable - copy it to a USB drive and run it anywhere."
        )
        tk.Label(p, text=instructions, font=("Segoe UI", 9), justify="left",
                 wraplength=440).pack(padx=10)

    # ── Settings ─────────────────────────────────────────────────────────────

    def _restore_settings_to_ui(self):
        s = self.settings
        self.tester_var.set(s["tester_name"])
        self.project_var.set(s["project_name"])
        self.hotkey_var.set(s["hotkey"])
        self.auto_ts_var.set(s["auto_timestamp"])
        self.auto_num_var.set(s["auto_number"])
        self.prompt_d_var.set(s["prompt_description"])
        self.minimize_var.set(s["minimize_on_capture"])
        self.delay_var.set(str(s["capture_delay"]))
        if s["save_path"]:
            self.save_path_var.set(s["save_path"])
        self._refresh_hotkey_hint()

    def _refresh_hotkey_hint(self):
        hk = self.settings["hotkey"].upper()
        self.hotkey_hint_var.set(f"Press  [ {hk} ]  to capture a screenshot")

    def _apply_hotkey(self):
        hk = self.hotkey_var.get().strip()
        if not hk:
            messagebox.showwarning("Invalid hotkey", "Hotkey cannot be empty.")
            return
        self.settings["hotkey"] = hk
        self.settings.save()
        if self.active:
            self._register_hotkey()
        self._refresh_hotkey_hint()
        messagebox.showinfo("Hotkey updated", f"Hotkey set to: {hk.upper()}")

    def _save_settings(self):
        s = self.settings
        s["hotkey"]              = self.hotkey_var.get().strip()
        s["auto_timestamp"]      = self.auto_ts_var.get()
        s["auto_number"]         = self.auto_num_var.get()
        s["prompt_description"]  = self.prompt_d_var.get()
        s["minimize_on_capture"] = self.minimize_var.get()
        s["capture_delay"]       = int(self.delay_var.get() or 0)
        s["tester_name"]         = self.tester_var.get().strip()
        s["project_name"]        = self.project_var.get().strip()
        s["save_path"]           = self.save_path_var.get().strip()
        s.save()
        self._refresh_hotkey_hint()
        messagebox.showinfo("Saved", "Settings saved successfully.")

    def _browse(self):
        folder = filedialog.askdirectory(initialdir=self.save_path_var.get())
        if folder:
            self.save_path_var.set(folder)

    # ── Session ───────────────────────────────────────────────────────────────

    def start_session(self):
        fname = self.fname_var.get().strip()
        if not fname:
            messagebox.showerror("Missing name", "Please enter a document file name.")
            return

        for ch in r'<>:"/\|?*':
            fname = fname.replace(ch, "_")
        if not fname.lower().endswith(".docx"):
            fname += ".docx"

        save_dir = self.save_path_var.get().strip() or os.path.expanduser("~/Desktop")
        os.makedirs(save_dir, exist_ok=True)

        self.doc_path = os.path.join(save_dir, fname)

        # FIX: only mark session active after document is successfully created
        try:
            self.document = self._create_document()
        except Exception as exc:
            messagebox.showerror("Document error",
                                 f"Could not create document:\n{exc}")
            return

        self.count      = 0
        self.sess_start = datetime.now()
        self.active     = True

        self._register_hotkey()

        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        hk = self.settings["hotkey"].upper()
        self.status_var.set(f"Session active - press [{hk}] to capture a screenshot!")
        self.count_var.set("Screenshots captured: 0")
        self.last_var.set("")

    def stop_session(self):
        if not self.active:
            return
        self.active = False

        try:
            keyboard.unhook_all()
        except Exception:
            pass

        if self.document:
            try:
                self._append_summary()
            except Exception:
                pass

        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.status_var.set(f"Session ended - {self.count} screenshot(s) saved.")
        self.count_var.set("")
        self.last_var.set("")

        if self.doc_path and os.path.exists(self.doc_path):
            if messagebox.askyesno(
                    "Session complete",
                    f"{self.count} screenshot(s) captured.\n\n"
                    f"Saved to:\n{self.doc_path}\n\nOpen document now?"):
                # FIX: os.startfile wrapped — fails gracefully if Word isn't installed
                try:
                    os.startfile(self.doc_path)
                except Exception:
                    messagebox.showinfo(
                        "File saved",
                        f"Document saved to:\n{self.doc_path}\n\n"
                        "Open it manually in Microsoft Word.")

    # ── Document ──────────────────────────────────────────────────────────────

    def _create_document(self) -> Document:
        doc = Document()

        title = doc.add_heading("Test Report", level=0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        tbl = doc.add_table(rows=4, cols=2)
        tbl.style = "Table Grid"

        def fill_row(i, label, value):
            cells = tbl.rows[i].cells
            cells[0].text = label
            cells[1].text = value
            # Bold the label cell
            for para in cells[0].paragraphs:
                for run in para.runs:
                    run.bold = True

        fill_row(0, "Tester",  self.tester_var.get()  or "-")
        fill_row(1, "Project", self.project_var.get() or "-")
        fill_row(2, "Date",    datetime.now().strftime("%Y-%m-%d"))
        fill_row(3, "File",    self.fname_var.get())

        doc.add_paragraph("")
        doc.add_heading("Test Screenshots", level=1)
        doc.save(self.doc_path)
        return doc

    def _append_summary(self):
        end_time = datetime.now()
        self.document.add_heading("Session Summary", level=1)
        p = self.document.add_paragraph()
        p.add_run("Total screenshots: ").bold = True
        p.add_run(str(self.count))
        p.add_run("    ")
        p.add_run("Start: ").bold = True
        start_str = (self.sess_start.strftime("%Y-%m-%d %H:%M:%S")
                     if self.sess_start else "-")
        p.add_run(start_str)
        p.add_run("    ")
        p.add_run("End: ").bold = True
        p.add_run(end_time.strftime("%Y-%m-%d %H:%M:%S"))
        self.document.save(self.doc_path)

    # ── Hotkey & capture ──────────────────────────────────────────────────────

    def _register_hotkey(self):
        try:
            keyboard.unhook_all()
        except Exception:
            pass
        try:
            keyboard.add_hotkey(
                self.settings["hotkey"],
                self._on_hotkey,
                suppress=False,
            )
        except Exception as exc:
            messagebox.showerror(
                "Hotkey error",
                f"Could not register hotkey '{self.settings['hotkey']}'.\n\n"
                f"Error: {exc}\n\n"
                "Try a different key in Settings.",
            )

    def _on_hotkey(self):
        if not self.active:
            return
        # FIX: snapshot all option values here (keyboard thread) before spawning
        # a worker thread — avoids reading tkinter vars from the wrong thread.
        opts = {
            "delay":       self.settings["capture_delay"],
            "minimize":    self.settings["minimize_on_capture"],
            "auto_ts":     self.settings["auto_timestamp"],
            "auto_num":    self.settings["auto_number"],
            "prompt_desc": self.settings["prompt_description"],
        }
        threading.Thread(target=self._capture_flow, args=(opts,), daemon=True).start()

    def _capture_flow(self, opts: dict):
        if opts["prompt_desc"]:
            self.root.after(0, lambda: self._desc_dialog(opts))
        else:
            self._do_capture("", opts)

    def _desc_dialog(self, opts: dict):
        dlg = tk.Toplevel(self.root)
        dlg.title("Screenshot description")
        dlg.geometry("400x140")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.lift()
        dlg.focus_force()
        dlg.attributes("-topmost", True)

        tk.Label(dlg, text="Description / test step (press Enter or leave blank):",
                 font=("Segoe UI", 9)).pack(pady=(12, 4))
        desc_var = tk.StringVar()
        entry = ttk.Entry(dlg, textvariable=desc_var, width=46)
        entry.pack(padx=14)
        entry.focus()

        btns = tk.Frame(dlg)
        btns.pack(pady=10)

        def capture_with_desc():
            desc = desc_var.get().strip()
            dlg.destroy()
            threading.Thread(target=self._do_capture, args=(desc, opts), daemon=True).start()

        def skip():
            dlg.destroy()
            threading.Thread(target=self._do_capture, args=("", opts), daemon=True).start()

        ttk.Button(btns, text="Capture", command=capture_with_desc, width=12).pack(side="left", padx=8)
        ttk.Button(btns, text="Skip",    command=skip,               width=8 ).pack(side="left", padx=4)
        entry.bind("<Return>", lambda _: capture_with_desc())

    def _do_capture(self, description: str, opts: dict):
        # FIX: lock prevents two concurrent captures writing to the document at once
        if not self._lock.acquire(blocking=False):
            return  # a capture is already in progress — skip this keypress

        try:
            if opts["minimize"]:
                self.root.after(0, self.root.iconify)

            # FIX: honour user delay (minimum 0.4 s so minimize animation finishes)
            time.sleep(max(0.4, opts["delay"]))

            # FIX: all_screens=True captures across multiple monitors
            screenshot = ImageGrab.grab(all_screens=True)
            buf = BytesIO()
            screenshot.save(buf, format="PNG")
            buf.seek(0)

            # FIX: increment count inside the lock — thread-safe
            self.count += 1
            n  = self.count
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            heading = f"Screenshot {n}" if opts["auto_num"] else "Screenshot"
            self.document.add_heading(heading, level=2)

            if opts["auto_ts"]:
                para = self.document.add_paragraph(f"Captured: {ts}")
                run  = para.runs[0]
                run.font.size      = Pt(8)
                run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

            if description:
                notes = self.document.add_paragraph(f"Notes: {description}")
                notes.runs[0].bold = True

            self.document.add_picture(buf, width=Inches(6.0))
            self.document.add_paragraph("")
            self.document.save(self.doc_path)

            self.root.after(0, lambda: self._update_status(n, ts))
            self.root.after(0, self.root.deiconify)

        except Exception as exc:
            self.root.after(0, self.root.deiconify)
            self.root.after(0, lambda: messagebox.showerror("Capture failed", str(exc)))
        finally:
            self._lock.release()

    def _update_status(self, n: int, ts: str):
        self.count_var.set(f"Screenshots captured: {n}")
        self.last_var.set(f"Last capture: {ts}")
        self.status_var.set(f"Screenshot {n} saved to document.")

    # ── Close ─────────────────────────────────────────────────────────────────

    def _on_close(self):
        if self.active:
            if not messagebox.askyesno(
                    "Quit", "A session is active. Stop and save before quitting?"):
                return
            self.stop_session()
        try:
            keyboard.unhook_all()
        except Exception:
            pass
        self.root.destroy()

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()


if __name__ == "__main__":
    QATool().run()
