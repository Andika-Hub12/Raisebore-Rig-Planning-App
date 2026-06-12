"""
rig_planner_gui.py
==================
A simple Windows desktop window (GUI) for the rig planner.

It reuses ALL the logic from rig_planner.py - this file only adds the
window, the input boxes, and the results table. Keep this file in the
SAME folder as rig_planner.py and rig_database.xlsx.

Run it with:   python rig_planner_gui.py
Or, once built into an .exe, just double-click the .exe.
"""

import os
import sys
from datetime import date, datetime
import tkinter as tk
from tkinter import ttk, messagebox

# Reuse the engine we already built.
from rig_planner import RigDatabase, Job, plan_jobs


# ---------------------------------------------------------------------
# Find the database file whether we run as a .py or as a bundled .exe.
# When PyInstaller builds an .exe it unpacks to a temp folder exposed
# as sys._MEIPASS; otherwise we use the folder this script lives in.
# ---------------------------------------------------------------------
def resource_path(filename: str) -> str:
    if hasattr(sys, "_MEIPASS"):
        base = sys._MEIPASS                      # running as bundled .exe
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


DB_FILE = resource_path("rig_database.xlsx")


class RigPlannerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Rig Planner")
        self.geometry("900x600")
        self.minsize(820, 540)

        self.jobs: list[Job] = []

        self._build_header()
        self._build_input_form()
        self._build_job_list()
        self._build_results()
        self._build_actions()

    # -----------------------------------------------------------------
    # UI sections
    # -----------------------------------------------------------------
    def _build_header(self):
        bar = tk.Frame(self, bg="#0b3d91", height=56)
        bar.pack(fill="x")
        tk.Label(bar, text="  Rig Planner", bg="#0b3d91", fg="white",
                 font=("Segoe UI", 16, "bold")).pack(side="left", pady=10)
        tk.Label(bar, text="Date  >  Reamer  >  Rods    ",
                 bg="#0b3d91", fg="#cfe0ff",
                 font=("Segoe UI", 10)).pack(side="right", pady=14)

    def _build_input_form(self):
        frm = tk.LabelFrame(self, text="Add a job", padx=10, pady=10,
                            font=("Segoe UI", 10, "bold"))
        frm.pack(fill="x", padx=12, pady=(12, 6))

        labels = ["Job ID", "Start date (YYYY-MM-DD)", "Duration (days)",
                  "Reamer size (in)", "Hole length (m)"]
        self.entries = {}
        for i, lab in enumerate(labels):
            tk.Label(frm, text=lab, font=("Segoe UI", 9)).grid(
                row=0, column=i, padx=4, sticky="w")
            e = tk.Entry(frm, width=18)
            e.grid(row=1, column=i, padx=4, pady=2)
            self.entries[lab] = e

        # Helpful defaults so the user can just hit "Add"
        self.entries["Job ID"].insert(0, "J-201")
        self.entries["Start date (YYYY-MM-DD)"].insert(0, "2026-06-01")
        self.entries["Duration (days)"].insert(0, "2")
        self.entries["Reamer size (in)"].insert(0, "12.25")
        self.entries["Hole length (m)"].insert(0, "1800")

        tk.Button(frm, text="Add job", command=self.add_job,
                  bg="#0b3d91", fg="white",
                  font=("Segoe UI", 9, "bold"), width=12).grid(
            row=1, column=len(labels), padx=8)

    def _build_job_list(self):
        frm = tk.LabelFrame(self, text="Jobs to plan", padx=6, pady=6,
                            font=("Segoe UI", 10, "bold"))
        frm.pack(fill="x", padx=12, pady=6)

        cols = ("Job", "Start", "Days", "Reamer", "Hole")
        self.job_tree = ttk.Treeview(frm, columns=cols, show="headings",
                                     height=4)
        for c in cols:
            self.job_tree.heading(c, text=c)
            self.job_tree.column(c, width=120, anchor="center")
        self.job_tree.pack(side="left", fill="x", expand=True)

        tk.Button(frm, text="Remove selected", command=self.remove_job,
                  width=16).pack(side="left", padx=8)

    def _build_results(self):
        frm = tk.LabelFrame(self, text="Plan results", padx=6, pady=6,
                            font=("Segoe UI", 10, "bold"))
        frm.pack(fill="both", expand=True, padx=12, pady=6)

        self.out = tk.Text(frm, wrap="word", font=("Consolas", 10),
                           bg="#f7f8fa")
        self.out.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(frm, command=self.out.yview)
        sb.pack(side="right", fill="y")
        self.out.config(yscrollcommand=sb.set)

        # Colour tags for pass / fail lines
        self.out.tag_config("pass", foreground="#15803d")
        self.out.tag_config("fail", foreground="#b91c1c")
        self.out.tag_config("head", foreground="#0b3d91",
                            font=("Consolas", 10, "bold"))

    def _build_actions(self):
        bar = tk.Frame(self)
        bar.pack(fill="x", padx=12, pady=(0, 12))
        tk.Button(bar, text="Run plan", command=self.run_plan,
                  bg="#15803d", fg="white",
                  font=("Segoe UI", 11, "bold"), width=16, height=1).pack(
            side="left")
        tk.Button(bar, text="Clear results", command=lambda:
                  self.out.delete("1.0", "end"), width=14).pack(
            side="left", padx=8)

    # -----------------------------------------------------------------
    # Actions
    # -----------------------------------------------------------------
    def add_job(self):
        try:
            job_id = self.entries["Job ID"].get().strip()
            start = datetime.strptime(
                self.entries["Start date (YYYY-MM-DD)"].get().strip(),
                "%Y-%m-%d").date()
            days = int(self.entries["Duration (days)"].get())
            reamer = float(self.entries["Reamer size (in)"].get())
            hole = float(self.entries["Hole length (m)"].get())
        except ValueError as e:
            messagebox.showerror("Invalid input",
                                 f"Please check your entries.\n\n{e}")
            return

        if not job_id:
            messagebox.showerror("Invalid input", "Job ID cannot be empty.")
            return

        self.jobs.append(Job(job_id, start, days, reamer, hole))
        self.job_tree.insert("", "end", values=(
            job_id, start, days, reamer, hole))

    def remove_job(self):
        sel = self.job_tree.selection()
        if not sel:
            return
        for item in sel:
            idx = self.job_tree.index(item)
            self.job_tree.delete(item)
            del self.jobs[idx]

    def run_plan(self):
        self.out.delete("1.0", "end")

        if not self.jobs:
            messagebox.showinfo("No jobs", "Add at least one job first.")
            return

        if not os.path.exists(DB_FILE):
            messagebox.showerror(
                "Database not found",
                f"Could not find rig_database.xlsx at:\n{DB_FILE}\n\n"
                "Make sure it sits in the same folder as this program.")
            return

        try:
            db = RigDatabase(DB_FILE)
            results = plan_jobs(db, self.jobs)
        except Exception as e:
            messagebox.showerror("Error while planning", str(e))
            return

        for job, winner, attempts in results:
            self.out.insert("end", f"JOB {job.job_id}\n", "head")
            self.out.insert(
                "end",
                f"  start={job.start_date}  days={job.duration_days}  "
                f"reamer={job.reamer_size_in}\"  hole={job.hole_length_m} m\n")
            for a in attempts:
                tag = "pass" if a.is_assignable else "fail"
                mark = "OK " if a.is_assignable else " - "
                self.out.insert("end",
                                f"  [{mark}] {a.rig_name}: {a.reason}\n", tag)
            if winner:
                self.out.insert(
                    "end",
                    f"  >>> ASSIGNED to {winner.rig_name} "
                    f"({winner.rods_needed} x {winner.rod_size_in}\" rods)\n\n",
                    "pass")
            else:
                self.out.insert("end", "  >>> NO RIG AVAILABLE\n\n", "fail")


if __name__ == "__main__":
    app = RigPlannerApp()
    app.mainloop()
