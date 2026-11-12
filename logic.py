"""Business rules: checklist scoring, status, and fair reassignment."""
from __future__ import annotations

import re
from collections import Counter

import pandas as pd

# Regions in travel order, so each analyst's share is one connected block of sites.
GEO_ORDER = ["South West", "Western Region", "National Office", "Eastern Region",
             "Karamoja", "Northern Region", "West Nile"]

# Name, Email, Team, Active, Role  -- confirm emails before going live
DEFAULT_ROSTER = [
    ["Moreen Banshanga", "moreen_banshanga@wvi.org", "Team 1", "Yes", "Analyst"],
    ["Kevin Mugarura", "kevin_mugarura@wvi.org", "Team 1", "Yes", "Analyst"],
    ["Alex Senyondwa", "alex_senyondwa@wvi.org", "Team 2", "Yes", "Analyst"],
    ["Gabriella Oboma", "gabriella_oboma@wvi.org", "Team 2", "Yes", "Analyst"],
    ["Abura Kamilo", "", "", "No", "Analyst"],
    ["Timothy Ateesa", "", "", "No", "Admin"],
    ["Lynate Kobusingye", "lynate_kobusinge@wvi.org", "Team 1", "No", "Exited"],
    ["Patricia Ainembabazi", "ainembabazi_patricia@wvi.org", "Team 2", "No", "Exited"],
]


TRACKING_COLS = ["Last Updated By", "Last Updated At"]
LOG_HEADERS = ["Timestamp", "Analyst", "Analyst Email", "Employee ID", "Client Name",
               "Location", "Action", "Items Complete", "Status", "Changes"]


def checklist_cols(columns) -> list[str]:
    """Columns that start with '1. ', '2. ' ... are the pre-flight checklist."""
    return [c for c in columns if re.match(r"^\d+\.\s", str(c))]


def is_yes(v) -> bool:
    return str(v).strip().lower() in ("yes", "y", "true", "1", "done")


def score(values: dict, checklist: list[str]) -> tuple[int, float, str]:
    done = sum(is_yes(values.get(c, "")) for c in checklist)
    pct = done / len(checklist) if checklist else 0
    status = "Complete" if done == len(checklist) else ("In Progress" if done else "Not Started")
    return done, round(pct, 4), status


def status_of(row: pd.Series, checklist: list[str]) -> str:
    return score(row.to_dict(), checklist)[2]


def rebalance(df: pd.DataFrame, analysts: list[str], checklist: list[str],
              include_in_progress: bool = False) -> pd.DataFrame:
    """Share outstanding endpoints equally across `analysts`.

    * Rows that are Not Started, or owned by someone not in `analysts`, go into the pool.
    * In-progress rows stay with their current active owner unless include_in_progress.
    * The pool is ordered geographically and handed out in contiguous blocks, so each
      analyst gets neighbouring sites, and totals end up as even as possible.
    Returns the rows whose analyst changes, with columns _row, Employee ID, From, To.
    """
    work = df.copy()
    work["_status"] = work.apply(lambda r: status_of(r, checklist), axis=1)
    out = work[work["_status"] != "Complete"]
    active = set(analysts)
    owner = out["Assigned Analyst"].fillna("").str.strip()
    keep_mask = owner.isin(active) & ((out["_status"] == "In Progress") & (not include_in_progress))
    kept, pool = out[keep_mask], out[~keep_mask]

    geo = {r: i for i, r in enumerate(GEO_ORDER)}
    pool = pool.assign(_g=pool["Region"].map(lambda r: geo.get(r, 99))).sort_values(["_g", "Location", "_row"])

    loads = Counter({a: 0 for a in analysts})
    loads.update(kept["Assigned Analyst"].str.strip())
    sizes = Counter()
    for _ in range(len(pool)):  # water-fill: next endpoint goes to the lightest load
        a = min(analysts, key=lambda x: (loads[x], analysts.index(x)))
        loads[a] += 1
        sizes[a] += 1

    to, i = [], 0
    for a in analysts:
        to += [a] * sizes[a]
    pool = pool.assign(To=to)
    pool["From"] = pool["Assigned Analyst"].fillna("").str.strip()
    return pool[["_row", "Employee ID", "Client Name", "Location", "Region", "_status", "From", "To"]]


def data_issues(df: pd.DataFrame, checklist: list[str]) -> pd.DataFrame:
    issues = []
    ids = df["Employee ID"].str.strip().value_counts()
    sn = df["Device Serial Number"].str.strip().str.upper()
    sn_counts = sn[sn != ""].value_counts()
    for _, r in df.iterrows():
        eid, name = r["Employee ID"].strip(), r["Client Name"].strip()
        status = status_of(r, checklist)
        add = lambda msg: issues.append({"Employee ID": eid or "(blank)", "Client Name": name, "Issue": msg, "_row": r["_row"]})
        if not eid:
            add("Missing Employee ID")
        elif ids.get(eid, 0) > 1:
            add("Duplicate Employee ID")
        s = r["Device Serial Number"].strip().upper()
        if s and sn_counts.get(s, 0) > 1:
            add(f"Serial {s} appears on more than one row")
        if status == "Complete" and not r["Assigned Analyst"].strip():
            add("Complete but no analyst recorded")
        if status == "Complete" and not r.get("Date Completed", "").strip():
            add("Complete but no completion date")
        if "|" in r.get("Client Email", ""):
            add(f"Bad email: {r['Client Email']}")
        if r.get("Status", "").strip() and r["Status"].strip() != status:
            add(f"Status column says '{r['Status']}' but checklist says '{status}'")
    return pd.DataFrame(issues)
