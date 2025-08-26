import os
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox
from script_core import run_automation  # your existing automation

ALL_ENTITIES = [
    "D341_HSO_HGM", "D342_HSO_HGMD", "CC41_HSO_HMIN", "C741_HSO_HMSH",
    "AN41_HSO_HMSP", "D941_HSO_HMSZ", "J34V_HSO_HOME", "EM41_HSO_HSEU",
    "A441_HSO_HSOT", "WB41_HSOU", "D841_HSO_HSOK", "CY41_HSO_INNOVIA",
    "WM41_HSO_MID LAB INC", "GG41_HSO_FRITZ RUCK", "GH41_HSO_EOS",
]

DEFAULT_SELECTED = [
    "D341_HSO_HGM", "D342_HSO_HGMD", "CC41_HSO_HMIN", "C741_HSO_HMSH",
    "AN41_HSO_HMSP", "D941_HSO_HMSZ", "J34V_HSO_HOME", "EM41_HSO_HSEU",
    "A441_HSO_HSOT", "WB41_HSOU", "D841_HSO_HSOK",
]

def downloads_dir() -> str:
    return os.path.join(os.path.expanduser("~"), "Downloads")

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("STRAVIS Automation Runner")
        self.geometry("760x560")
        self.resizable(False, False)

        ttk.Label(self, text="STRAVIS Automation Runner", font=("Segoe UI", 14, "bold")).pack(pady=(12, 2))
        ttk.Label(
            self,
            text="Fill the inputs, click Run, then immediately switch focus to STRAVIS.",
            foreground="#555"
        ).pack()

        # Login requirement banner
        login_banner = ttk.Label(
            self,
            text="You need to be logged into STRAVIS before running the automation.",
            foreground="#0a5"
        )
        login_banner.pack(pady=(2, 8))

        # Where files go
        ttk.Label(
            self,
            text=f"ℹ️ All files will be saved into your system's Downloads folder: {downloads_dir()}",
            foreground="#0a5"
        ).pack(pady=(0, 10))

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=16, pady=10)

        # Target period
        row = ttk.Frame(body)
        row.pack(fill="x", pady=6)
        ttk.Label(row, text="Target period (YYYY.MM):", width=26).pack(side="left")
        self.target_period_var = tk.StringVar(value="2025.03")
        ttk.Entry(row, textvariable=self.target_period_var, width=18).pack(side="left")

        # Selection helpers
        helper = ttk.Frame(body)
        helper.pack(fill="x", pady=6)
        ttk.Label(
            helper,
            text="Select the entities to INCLUDE (others will be deselected):",
            font=("Segoe UI", 10, "bold")
        ).pack(side="left")
        right = ttk.Frame(helper)
        right.pack(side="right")
        ttk.Button(right, text="Select defaults", command=self.select_defaults).pack(side="left", padx=4)
        ttk.Button(right, text="Select all", command=self.select_all).pack(side="left", padx=4)
        ttk.Button(right, text="Clear all", command=self.clear_all).pack(side="left")

        # Checkboxes in 3 columns
        self.vars = {}
        grid = ttk.Frame(body)
        grid.pack(fill="x", pady=6)
        cols = 3
        for i, ent in enumerate(ALL_ENTITIES):
            r, c = divmod(i, cols)
            var = tk.BooleanVar(value=(ent in DEFAULT_SELECTED))
            self.vars[ent] = var
            ttk.Checkbutton(grid, text=ent, variable=var).grid(row=r, column=c, sticky="w", padx=6, pady=4)

        # “I’m logged in” gate
        gate_row = ttk.Frame(body)
        gate_row.pack(fill="x", pady=(8, 2))
        self.logged_in_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            gate_row,
            text="I confirm I am logged into STRAVIS (required)",
            variable=self.logged_in_var,
            command=self._update_run_state
        ).pack(side="left")

        # Status + Run button
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=16, pady=12)

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(bottom, textvariable=self.status_var, foreground="#444").pack(side="left")

        self.run_btn = ttk.Button(bottom, text="Run automation", command=self.on_run, width=20)
        self.run_btn.pack(side="right")
        self._update_run_state()

        # Style
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("TButton", padding=6)
        style.configure("TCheckbutton", padding=2)

    # ----- selection helpers -----
    def select_defaults(self):
        for e, v in self.vars.items():
            v.set(e in DEFAULT_SELECTED)

    def select_all(self):
        for v in self.vars.values():
            v.set(True)

    def clear_all(self):
        for v in self.vars.values():
            v.set(False)

    def _update_run_state(self):
        self.run_btn.config(state=("normal" if self.logged_in_var.get() else "disabled"))

    # ----- run flow -----
    def on_run(self):
        if not self.logged_in_var.get():
            messagebox.showwarning("Login required",
                                   "You need to be logged into STRAVIS before running the automation.")
            return

        selected = [e for e, v in self.vars.items() if v.get()]
        if not selected:
            messagebox.showerror("Validation error", "Please select at least one entity.")
            return

        target_period = self.target_period_var.get().strip()
        if not target_period:
            messagebox.showerror("Validation error", "Please enter a target period (e.g. 2025.03).")
            return

        to_deselect = [e for e in ALL_ENTITIES if e not in selected]
        iterations = len(selected)

        # Final heads-up
        msg = (
            "After you press OK, you have ~3 seconds to bring STRAVIS to the foreground.\n"
            "Don't touch mouse/keyboard while it runs.\n"
            "Tip: Move mouse to top-left if you need to abort."
        )
        if not messagebox.askokcancel("Heads up", msg):
            return

        self.run_btn.config(state="disabled")
        self.status_var.set("Starting in 3 seconds… switch to STRAVIS now")

        threading.Thread(
            target=self._run_worker,
            args=(target_period, to_deselect, iterations),
            daemon=True
        ).start()

    def _run_worker(self, target_period, to_deselect, iterations):
        """Worker thread: COM init here so pywin32/uiautomation/pywinauto can operate."""
        pythoncom = None
        try:
            import pythoncom
            pythoncom.CoInitialize()

            time.sleep(3)  # let user focus STRAVIS
            run_automation(target_period, to_deselect, select_n=20, iterations=iterations)

        except ImportError as e:
            self.after(0, self._done, False,
                       "pywin32 is required for Windows automation.\n\nFix: pip install pywin32\n\n"
                       f"Details: {e}")
        except Exception as e:
            self.after(0, self._done, False, f"Automation failed: {e}")
        finally:
            try:
                if pythoncom:
                    pythoncom.CoUninitialize()
            except Exception:
                pass

    def _done(self, ok: bool, message: str):
        self.status_var.set(message)
        if ok:
            # Completed popup with option to open Downloads
            if messagebox.askyesno(
                "Downloads complete",
                f"{message}\n\nFiles should be in:\n{downloads_dir()}\n\nOpen the folder now?"
            ):
                try:
                    os.startfile(downloads_dir())
                except Exception:
                    messagebox.showinfo("Note", f"Could not open folder. Please navigate to:\n{downloads_dir()}")
            else:
                messagebox.showinfo("Done", message)
        else:
            messagebox.showerror("Error", message)
        self.run_btn.config(state="normal")

if __name__ == "__main__":
    App().mainloop()