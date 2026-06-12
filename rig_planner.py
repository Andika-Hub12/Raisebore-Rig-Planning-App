"""
rig_planner.py
==============
A small planning app that decides which rig can take a job by running
THREE sequential checks against an Excel database:

    Step 1 - DATE         : is the rig free on every day the job needs?
    Step 2 - REAMER SIZE  : can the rig physically run that reamer?
    Step 3 - RODS         : do we own enough rods (right size) for the hole?

The checks are *short-circuiting*: as soon as one fails, the rig is
rejected and the next checks are skipped (this matches the user's
"if one got crossed, dont check the rest" rule).

Run:
    python rig_planner.py
"""

from __future__ import annotations
from dataclasses import dataclass
from datetime import date, timedelta
from typing import List, Dict, Set, Tuple
import math
import pandas as pd

DB_PATH = "rig_database.xlsx"


# ---------------------------------------------------------------------
# Data classes - lightweight "record" types so code reads like English.
# ---------------------------------------------------------------------
@dataclass
class Job:
    """A single job we want to plan."""
    job_id: str
    start_date: date
    duration_days: int
    reamer_size_in: float
    hole_length_m: float

    @property
    def required_dates(self) -> List[date]:
        """Every calendar day the rig is needed for this job."""
        return [self.start_date + timedelta(days=i)
                for i in range(self.duration_days)]


@dataclass
class PlanResult:
    """Outcome of trying to assign one job to one rig."""
    job_id: str
    rig_name: str
    passed_date: bool = False
    passed_reamer: bool = False
    passed_rods: bool = False
    rods_needed: int = 0
    rod_size_in: float = 0.0
    reason: str = ""

    @property
    def is_assignable(self) -> bool:
        return self.passed_date and self.passed_reamer and self.passed_rods


# ---------------------------------------------------------------------
# Database loader - reads every sheet of the workbook into pandas
# DataFrames, then converts them to fast lookup structures.
# ---------------------------------------------------------------------
class RigDatabase:
    def __init__(self, path: str = DB_PATH):
        self.path = path
        sheets = pd.read_excel(path, sheet_name=None)

        self.rigs_df         = sheets["Rigs"]
        self.schedule_df     = sheets["RigSchedule"]
        self.capability_df   = sheets["ReamerCapability"]
        self.rods_df         = sheets["Rods"]
        self.mapping_df      = sheets["ReamerRodMapping"]

        # Normalise the date column to plain `date` objects.
        self.schedule_df["BusyDate"] = pd.to_datetime(
            self.schedule_df["BusyDate"]).dt.date

        # ---- Pre-build lookups ----
        # 1) Busy days per rig: {"RIG-A": {date, date, ...}}
        self.busy: Dict[str, Set[date]] = {}
        for rig in self.rigs_df["RigName"]:
            self.busy[rig] = set(
                self.schedule_df.loc[
                    self.schedule_df["RigName"] == rig, "BusyDate"])

        # 2) Reamer capability per rig: {"RIG-A": {8.5, 12.25, 17.5}}
        self.capable: Dict[str, Set[float]] = {}
        for rig in self.rigs_df["RigName"]:
            self.capable[rig] = set(
                self.capability_df.loc[
                    self.capability_df["RigName"] == rig,
                    "ReamerSize_in"].astype(float))

        # 3) Reamer -> rod size map: {8.5: 3.5, 12.25: 4.5, 17.5: 5.5}
        self.reamer_to_rod: Dict[float, float] = dict(
            zip(self.mapping_df["ReamerSize_in"].astype(float),
                self.mapping_df["RequiredRodSize_in"].astype(float)))

        # 4) Rod inventory keyed by size:
        #    {3.5: {"length_per_rod": 9, "available": 400}}
        self.rod_inventory: Dict[float, Dict[str, float]] = {}
        for _, row in self.rods_df.iterrows():
            self.rod_inventory[float(row["RodSize_in"])] = {
                "length_per_rod": float(row["LengthPerRod_m"]),
                "available":      int(row["AvailableCount"]),
            }

    # ----- Convenience -----
    @property
    def rig_names(self) -> List[str]:
        return list(self.rigs_df["RigName"])

    def max_hole_length(self, rig: str) -> float:
        return float(self.rigs_df.loc[
            self.rigs_df["RigName"] == rig, "MaxHoleLength_m"].iloc[0])


# ---------------------------------------------------------------------
# The three sequential checks - each returns (passed, reason).
# ---------------------------------------------------------------------
def check_date(db: RigDatabase, rig: str, job: Job) -> Tuple[bool, str]:
    """Step 1: rig must be free on every day in job.required_dates."""
    clashes = [d for d in job.required_dates if d in db.busy[rig]]
    if clashes:
        return False, f"busy on {', '.join(str(d) for d in clashes)}"
    return True, "all required dates free"


def check_reamer(db: RigDatabase, rig: str, job: Job) -> Tuple[bool, str]:
    """Step 2: rig must be able to run the requested reamer size."""
    if job.reamer_size_in not in db.capable[rig]:
        return False, (f"cannot run {job.reamer_size_in}\" reamer "
                       f"(capable: {sorted(db.capable[rig])})")
    return True, f"{job.reamer_size_in}\" reamer supported"


def check_rods(db: RigDatabase, rig: str, job: Job
               ) -> Tuple[bool, str, int, float]:
    """Step 3: enough rods of the right size for the hole length?
    Also enforces the rig's max hole length."""
    # 3a) rig's own depth limit
    if job.hole_length_m > db.max_hole_length(rig):
        return (False,
                f"hole length {job.hole_length_m} m exceeds rig limit "
                f"{db.max_hole_length(rig)} m",
                0, 0.0)

    # 3b) which rod pairs with this reamer?
    rod_size = db.reamer_to_rod.get(job.reamer_size_in)
    if rod_size is None:
        return False, f"no rod mapping for reamer {job.reamer_size_in}\"", 0, 0.0

    inv = db.rod_inventory.get(rod_size)
    if inv is None:
        return False, f"rod size {rod_size}\" not in inventory", 0, rod_size

    # 3c) how many rods does this hole need? Always round up.
    rods_needed = math.ceil(job.hole_length_m / inv["length_per_rod"])

    if rods_needed > inv["available"]:
        return (False,
                f"need {rods_needed} x {rod_size}\" rods, "
                f"only {inv['available']} available",
                rods_needed, rod_size)

    return (True,
            f"{rods_needed} x {rod_size}\" rods reserved "
            f"(stock {inv['available']})",
            rods_needed, rod_size)


# ---------------------------------------------------------------------
# Engine - runs the three checks SEQUENTIALLY and short-circuits.
# ---------------------------------------------------------------------
def evaluate(db: RigDatabase, rig: str, job: Job) -> PlanResult:
    res = PlanResult(job_id=job.job_id, rig_name=rig)

    # Step 1 ---------------------------------------------------------
    ok, why = check_date(db, rig, job)
    res.passed_date = ok
    if not ok:
        res.reason = f"[DATE FAIL] {why}"
        return res                      # short-circuit, skip 2 & 3

    # Step 2 ---------------------------------------------------------
    ok, why = check_reamer(db, rig, job)
    res.passed_reamer = ok
    if not ok:
        res.reason = f"[REAMER FAIL] {why}"
        return res                      # short-circuit, skip 3

    # Step 3 ---------------------------------------------------------
    ok, why, n_rods, rod_size = check_rods(db, rig, job)
    res.passed_rods   = ok
    res.rods_needed   = n_rods
    res.rod_size_in   = rod_size
    res.reason        = ("PASS - " + why) if ok else f"[ROD FAIL] {why}"
    return res


# ---------------------------------------------------------------------
# Greedy assignment: for each job, pick the first rig that passes all
# three checks, then reserve its dates & rods so later jobs see the
# updated state.
# ---------------------------------------------------------------------
def plan_jobs(db: RigDatabase, jobs: List[Job]
              ) -> List[Tuple[Job, PlanResult | None, List[PlanResult]]]:
    """Returns a list of tuples: (job, winning_result_or_None, all_attempts)."""
    output = []
    for job in jobs:
        attempts: List[PlanResult] = []
        winner: PlanResult | None = None

        for rig in db.rig_names:
            r = evaluate(db, rig, job)
            attempts.append(r)
            if r.is_assignable:
                winner = r
                # ---- reserve resources so the next job sees them used ----
                for d in job.required_dates:
                    db.busy[rig].add(d)
                db.rod_inventory[r.rod_size_in]["available"] -= r.rods_needed
                break

        output.append((job, winner, attempts))
    return output


# ---------------------------------------------------------------------
# Pretty printing
# ---------------------------------------------------------------------
def print_plan(results) -> None:
    print("=" * 72)
    print("RIG ASSIGNMENT PLAN")
    print("=" * 72)
    for job, winner, attempts in results:
        print(f"\nJOB {job.job_id}")
        print(f"  start={job.start_date}  days={job.duration_days}  "
              f"reamer={job.reamer_size_in}\"  hole={job.hole_length_m} m")
        for a in attempts:
            mark = "OK " if a.is_assignable else " - "
            print(f"  [{mark}] {a.rig_name}: {a.reason}")
        if winner:
            print(f"  >>> ASSIGNED to {winner.rig_name} "
                  f"({winner.rods_needed} x {winner.rod_size_in}\" rods)")
        else:
            print("  >>> NO RIG AVAILABLE")
    print("\n" + "=" * 72)


# ---------------------------------------------------------------------
# Demo: three jobs (matches the example in the user's request)
# ---------------------------------------------------------------------
if __name__ == "__main__":
    db = RigDatabase(DB_PATH)

    jobs = [
        Job("J-201", date(2026, 6, 1), duration_days=2,
            reamer_size_in=12.25, hole_length_m=1800),
        Job("J-202", date(2026, 6, 5), duration_days=3,
            reamer_size_in=17.5,  hole_length_m=1080),
        Job("J-203", date(2026, 6, 8), duration_days=1,
            reamer_size_in=8.5,   hole_length_m=900),
    ]

    plan = plan_jobs(db, jobs)
    print_plan(plan)
