#!/usr/bin/env python3
"""
DOCX Image Swap — GUI (ttkbootstrap)
"""

import sys
import json
import shutil
import subprocess
import threading
import os
from pathlib import Path
from tkinter import filedialog

try:
    import ttkbootstrap as ttk
    from ttkbootstrap.constants import *
    try:
        from ttkbootstrap.widgets.scrolled import ScrolledText
    except ImportError:
        from ttkbootstrap.scrolled import ScrolledText
except ImportError:
    print("ttkbootstrap is required: pip install ttkbootstrap")
    sys.exit(1)

ROOT = Path(__file__).resolve().parent
CONFIG_FILE = ROOT / "config.json"
WORK_DIR = ROOT / "temp"
IMAGES_DIR = ROOT / "images"
OUTPUT_DIR = ROOT / "output"

APP_TITLE = "DOCX Image Swap"
VERSION = "1.0"
FONT = "Segoe UI"
MONO = "Consolas"

WINDOW_W, WINDOW_H = 850, 620
STEP_ICONS = {"pending": "○", "running": "▶", "done": "✓", "error": "✗", "skip": "–"}
STEP_COLORS = {"pending": "#6c7a89", "running": "#3498db", "done": "#27ae60", "error": "#e74c3c", "skip": "#95a5a6"}


class App:

    def __init__(self):
        self.root = ttk.Window(title=APP_TITLE, themename="darkly", resizable=(True, True))
        self.root.minsize(700, 500)
        self._center()

        self.config = self._load_config()
        self.input_file: Path | None = None
        self.current_step = 0
        self.running = False
        self.step_labels: list[ttk.Label] = []

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    # ── Layout ─────────────────────────────────────────────────

    def _center(self):
        x = (self.root.winfo_screenwidth() - WINDOW_W) // 2
        y = (self.root.winfo_screenheight() - WINDOW_H) // 2
        self.root.geometry(f"{WINDOW_W}x{WINDOW_H}+{x}+{y}")

    def _load_config(self):
        with open(CONFIG_FILE, encoding="utf-8") as f:
            return json.load(f)

    def _build_ui(self):
        # ── Navbar ──
        navbar = ttk.Frame(self.root, bootstyle="dark")
        navbar.pack(fill=X)

        logo = ttk.Frame(navbar, bootstyle="dark")
        logo.pack(side=LEFT, padx=(16, 0), pady=6)
        ttk.Label(logo, text="DIS", font=(FONT, 11, "bold"), foreground="#5a9fd4",
                  bootstyle="inverse-dark").pack(side=LEFT, padx=(0, 8))
        ttk.Label(logo, text=f"v{VERSION}", font=(FONT, 8), foreground="#6c7a89",
                  bootstyle="inverse-dark").pack(side=LEFT, pady=(2, 0))

        # ── Content ──
        content = ttk.Frame(self.root)
        content.pack(fill=BOTH, expand=YES, padx=20, pady=(12, 0))

        ttk.Label(content, text="DOCX Image Swap", font=(FONT, 14, "bold"),
                  bootstyle="light").pack(anchor=W)
        ttk.Label(content, text="Extract → Replace → Reinject images without touching the text",
                  font=(FONT, 9), bootstyle="secondary").pack(anchor=W, pady=(0, 12))

        # ── File picker ──
        file_frame = ttk.Frame(content)
        file_frame.pack(fill=X, pady=(0, 10))

        self.file_var = ttk.StringVar(value="No file selected")
        ttk.Entry(file_frame, textvariable=self.file_var, state="readonly",
                  font=(FONT, 10)).pack(side=LEFT, fill=X, expand=YES, padx=(0, 8))
        ttk.Button(file_frame, text="Browse…", bootstyle="info-outline",
                   command=self._browse).pack(side=RIGHT)

        # ── Steps ──
        ttk.Label(content, text="Pipeline", font=(FONT, 10, "bold"),
                  foreground="#5a9fd4").pack(anchor=W, pady=(4, 2))
        steps_frame = ttk.Frame(content)
        steps_frame.pack(fill=X, pady=(0, 8))

        for i, step in enumerate(self.config.get("steps", [])):
            name = step.get("name", f"Step {i+1}")
            lbl = ttk.Label(steps_frame, text=f"  {STEP_ICONS['pending']}  {name}",
                            font=(FONT, 10), foreground=STEP_COLORS["pending"])
            lbl.pack(anchor=W, pady=1)
            self.step_labels.append(lbl)

        # ── Buttons ──
        btn_frame = ttk.Frame(content)
        btn_frame.pack(fill=X, pady=(4, 8))

        self.run_btn = ttk.Button(btn_frame, text="▶  Run All Steps", bootstyle="success",
                                  command=self._run, state=DISABLED)
        self.run_btn.pack(side=LEFT)

        self.continue_btn = ttk.Button(btn_frame, text="Continue after swap", bootstyle="warning",
                                       command=self._continue, state=DISABLED)
        self.continue_btn.pack(side=LEFT, padx=(8, 0))

        self.open_img_btn = ttk.Button(btn_frame, text="📂 images/", bootstyle="secondary-outline",
                                       command=lambda: self._open_dir(IMAGES_DIR), state=DISABLED)
        self.open_img_btn.pack(side=LEFT, padx=(8, 0))

        self.open_out_btn = ttk.Button(btn_frame, text="📂 output/", bootstyle="secondary-outline",
                                       command=lambda: self._open_dir(OUTPUT_DIR), state=DISABLED)
        self.open_out_btn.pack(side=RIGHT)

        # ── Log ──
        ttk.Label(content, text="Log", font=(FONT, 10, "bold"),
                  foreground="#5a9fd4").pack(anchor=W, pady=(4, 2))
        log_frame = ttk.Frame(content)
        log_frame.pack(fill=BOTH, expand=YES, pady=(0, 10))

        self.log = ScrolledText(log_frame, font=(MONO, 9), wrap="word", height=10, state=DISABLED,
                                autohide=True)
        self.log.pack(fill=BOTH, expand=YES)

    # ── Actions ────────────────────────────────────────────────

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select DOCX file",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if path:
            self.input_file = Path(path)
            self.file_var.set(self.input_file.name)
            self.run_btn.configure(state=NORMAL)
            self._log(f"Selected: {self.input_file}")

    def _run(self):
        if not self.input_file:
            return
        self.running = True
        self.run_btn.configure(state=DISABLED)
        self.current_step = 0

        for i in range(len(self.step_labels)):
            self._set_step("pending", i)

        self._log(f"\n{'═' * 60}")
        self._log(f"  Starting: {self.input_file.name}")
        self._log(f"{'═' * 60}\n")

        WORK_DIR.mkdir(exist_ok=True)
        IMAGES_DIR.mkdir(exist_ok=True)
        OUTPUT_DIR.mkdir(exist_ok=True)

        shutil.copy2(self.input_file, WORK_DIR / "working.docx")
        info = {"original_name": self.input_file.stem, "original_path": str(self.input_file)}
        with open(WORK_DIR / "source_info.json", "w", encoding="utf-8") as f:
            json.dump(info, f, ensure_ascii=False)

        self._next_step()

    def _next_step(self):
        steps = self.config.get("steps", [])
        if self.current_step >= len(steps):
            self._finish()
            return

        cfg = steps[self.current_step]
        idx = self.current_step

        if not cfg.get("enabled", True):
            self._set_step("skip", idx)
            self._log(f"[SKIP] {cfg.get('name', '')}")
            self.current_step += 1
            self.root.after(50, self._next_step)
            return

        # Manual step
        if cfg.get("type") == "manual":
            self._set_step("running", idx)
            self._log(f"\n[{idx+1}/4] {cfg.get('name', '')} — WAITING")
            self._log(f"  Open images/ folder, replace files, then click 'Continue'.\n")
            self.continue_btn.configure(state=NORMAL)
            self.open_img_btn.configure(state=NORMAL)
            return

        # Regular step
        self._set_step("running", idx)
        self._log(f"\n[{idx+1}/4] {cfg.get('name', '')}")

        def worker():
            try:
                fn = self._get_fn(idx)
                merged = {**cfg}
                if "config" in cfg:
                    merged.update(cfg["config"])
                ok = fn(ROOT, merged, lambda m: self.root.after(0, self._log, m))
                self.root.after(0, self._step_done, idx, ok)
            except Exception as e:
                self.root.after(0, self._log, f"  Error: {e}")
                self.root.after(0, self._step_done, idx, False)

        threading.Thread(target=worker, daemon=True).start()

    def _get_fn(self, idx):
        from modules.step_01_extract_images import run as s1
        from modules.step_02_manual_swap import run as s2
        from modules.step_03_insert_images import run as s3
        from modules.step_04_save_result import run as s4
        return [s1, s2, s3, s4][idx]

    def _step_done(self, idx, ok):
        if ok:
            self._set_step("done", idx)
            self.current_step += 1
            self.root.after(100, self._next_step)
        else:
            self._set_step("error", idx)
            self._log(f"\n[ERROR] Step {idx+1} failed.")
            self._reset()

    def _continue(self):
        self.continue_btn.configure(state=DISABLED)
        self.open_img_btn.configure(state=DISABLED)
        self._set_step("done", self.current_step)
        self._log("  Swap confirmed. Continuing…\n")
        self.current_step += 1
        self.root.after(100, self._next_step)

    def _finish(self):
        self._log(f"\n{'═' * 60}")
        self._log("  DONE! Check the output/ folder.")
        self._log(f"{'═' * 60}")
        self.open_out_btn.configure(state=NORMAL)
        self._reset()

    def _reset(self):
        self.running = False
        self.run_btn.configure(state=NORMAL if self.input_file else DISABLED)

    # ── Helpers ────────────────────────────────────────────────

    def _set_step(self, status, idx):
        step = self.config["steps"][idx]
        name = step.get("name", f"Step {idx+1}")
        icon = STEP_ICONS.get(status, "○")
        color = STEP_COLORS.get(status, "#6c7a89")
        self.step_labels[idx].configure(text=f"  {icon}  {name}", foreground=color)

    def _log(self, msg: str):
        self.log.text.configure(state=NORMAL)
        self.log.text.insert(END, msg + "\n")
        self.log.text.see(END)
        self.log.text.configure(state=DISABLED)

    def _open_dir(self, path: Path):
        path.mkdir(exist_ok=True)
        p = str(path.resolve())
        if sys.platform == "win32":
            os.startfile(p)
        elif sys.platform == "darwin":
            subprocess.run(["open", p])
        else:
            subprocess.run(["xdg-open", p])

    def _on_close(self):
        self.root.destroy()


if __name__ == "__main__":
    App()
