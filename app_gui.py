import os
import time
import tkinter as tk
import multiprocessing as mp
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


def _worker_entry(target_period, to_deselect, iterations, result_q):
    """
    Child process entry point.
    Initializes COM, waits 3s for focus, runs automation, reports result back to parent via Queue.
    """
    try:
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            # If pywin32 not available or COM init fails, still try to run; parent will show error if needed
            pythoncom = None  # noqa: F841

        time.sleep(3)
        # Hardcode Shift+Down rows to 20 (same as before)
        run_automation(target_period, to_deselect, select_n=20, iterations=iterations)
        result_q.put(("ok", "Automation finished without raising errors."))
    except Exception as e:
        result_q.put(("err", f"Automation failed: {e}"))


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("STRAVIS Automation Runner")
        # ✅ make window resizable
        self.geometry("820x600")
        self.resizable(True, True)

        # State for child process & queue
        self.proc: mp.Process | None = None
        self.result_q: mp.Queue | None = None
        self.poll_job = None

        ttk.Label(self, text="STRAVIS Automation Runner", font=("Segoe UI", 14, "bold")).pack(pady=(12, 2))
        ttk.Label(
            self,
            text="Fill the inputs, click Run, then immediately switch focus to STRAVIS.",
            foreground="#555"
        ).pack()

        # Login requirement banner
        ttk.Label(
            self,
            text="You need to be logged into STRAVIS before running the automation.",
            foreground="#0a5"
        ).pack(pady=(2, 8))

        # Where files go
        ttk.Label(
            self,
            text=f"ℹ️ All files will be saved into your system's Downloads folder",
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

        # Checkboxes in a grid that expands
        self.vars = {}
        grid_wrap = ttk.Frame(body)
        grid_wrap.pack(fill="both", expand=True, pady=6)
        grid = ttk.Frame(grid_wrap)
        grid.pack(anchor="nw")
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

        # Status + Run/Stop buttons
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=16, pady=12)

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(bottom, textvariable=self.status_var, foreground="#444").pack(side="left")

        btns = ttk.Frame(bottom)
        btns.pack(side="right")
        self.stop_btn = ttk.Button(btns, text="Stop", command=self.on_stop, width=12, state="disabled")
        self.stop_btn.pack(side="right", padx=(8, 0))
        self.run_btn = ttk.Button(btns, text="Run automation", command=self.on_run, width=16)
        self.run_btn.pack(side="right")

        self._update_run_state()

        # Close handler to ensure child process is terminated
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Simple style improvements
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
        self.run_btn.config(state=("normal" if self.logged_in_var.get() and self.proc is None else "disabled"))

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

        # Spin up child process
        self.result_q = mp.Queue()
        self.proc = mp.Process(target=_worker_entry, args=(target_period, to_deselect, iterations, self.result_q))
        self.proc.daemon = True  # auto-kill with parent if needed
        self.proc.start()

        self.status_var.set("Starting in 3 seconds… switch to STRAVIS now")
        self.run_btn.config(state="disabled")
        self.stop_btn.config(state="normal")

        # start polling for completion messages
        self._poll_results()

    def _poll_results(self):
        # If process ended, try to read final message
        if self.proc is not None and not self.proc.is_alive():
            try:
                kind, msg = (self.result_q.get_nowait() if self.result_q else ("err", "Unknown error"))
            except Exception:
                kind, msg = ("err", "Automation process ended unexpectedly.")
            finally:
                self._finalize_run(kind, msg)
            return

        # Process still running — check queue non-blocking
        if self.result_q is not None:
            try:
                kind, msg = self.result_q.get_nowait()
                # Got a result even though proc might still be alive — finalize and ensure process is gone
                self._finalize_run(kind, msg)
                return
            except Exception:
                pass

        # keep polling ~200ms
        self.poll_job = self.after(200, self._poll_results)

    def on_stop(self):
        if self.proc is None:
            return
        if messagebox.askokcancel("Stop automation", "Are you sure you want to stop the automation now?"):
            try:
                if self.poll_job:
                    self.after_cancel(self.poll_job)
                    self.poll_job = None
                if self.proc.is_alive():
                    self.proc.terminate()
                    self.proc.join(timeout=2)
            finally:
                self._finalize_run("err", "Automation stopped by user.")

    def _finalize_run(self, kind: str, message: str):
        # Clean up process & polling
        if self.poll_job:
            self.after_cancel(self.poll_job)
            self.poll_job = None
        if self.proc:
            try:
                if self.proc.is_alive():
                    self.proc.terminate()
                self.proc.close()
            except Exception:
                pass
        self.proc = None
        self.result_q = None

        self.status_var.set(message)
        self.stop_btn.config(state="disabled")
        self._update_run_state()

        if kind == "ok":
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
            # Error / Stopped
            messagebox.showerror("Automation ended", message)

    def _on_close(self):
        # Ensure child process is killed on exit
        try:
            if self.proc and self.proc.is_alive():
                self.proc.terminate()
        except Exception:
            pass
        self.destroy()


if __name__ == "__main__":
    # Required for multiprocessing on Windows when frozen (PyInstaller)
    mp.freeze_support()
    App().mainloop()