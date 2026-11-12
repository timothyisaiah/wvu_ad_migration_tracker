"""Azure AD Join Migration tracker - Streamlit front end for the shared Excel workbook.

Run on the host PC:   streamlit run app.py --server.address 0.0.0.0
The workbook is a file on the host PC inside a OneDrive/SharePoint-synced folder; OneDrive
keeps it in sync with everyone else. Anyone using the app can set its path on the Workbook page.
There is no sign-in: people pick their own name from the roster.
"""
from __future__ import annotations

import json
import os
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

from excel_file import (ExcelAppBackend, OpenpyxlFileBackend, excel_available,
                        file_status, find_workbooks)
from logic import (DEFAULT_ROSTER, GEO_ORDER, LOG_HEADERS, TRACKING_COLS,
                   checklist_cols, data_issues, is_yes, rebalance, score,
                   status_of)

st.set_page_config(page_title="Migration Tracker", page_icon="🖥️", layout="wide")
APP_DIR = Path(__file__).parent
SETTINGS_FILE = APP_DIR / "settings.json"


# ----------------------------------------------------------------------------- config
def secret(section: str, default=None):
    try:
        return dict(st.secrets[section])
    except Exception:
        return default


CFG = {"tracker_sheet": "Endpoint Tracker", "analysts_sheet": "Analysts",
       "log_sheet": "Activity Log", "log_table": "ActivityLog"}
CFG.update(secret("tracker", {}) or {})


HOME_VAR = "%USERPROFILE%"


def portable(path: str) -> str:
    """C:/Users/timot/Box/x.xlsx -> %USERPROFILE%/Box/x.xlsx, so the same setting works for any user."""
    home = os.path.expandvars(HOME_VAR)
    if home != HOME_VAR and path.lower().startswith(home.lower() + os.sep):
        return HOME_VAR + path[len(home):]
    return path


def local_path(path: str) -> str:
    """Expand %USERPROFILE% etc. to this PC's own folders."""
    return os.path.expandvars(path) if path else path


def read_settings() -> dict:
    try:
        s = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    s["file_path"] = local_path(s.get("file_path", ""))
    return s


def write_settings(s: dict):
    s = {**s, "file_path": portable(s.get("file_path", ""))}
    SETTINGS_FILE.write_text(json.dumps(s, indent=2), encoding="utf-8")
    get_backend.clear()
    refresh()


SETTINGS = read_settings()
ENGINE = SETTINGS.get("engine") or ("excel" if excel_available() else "openpyxl")


@st.cache_resource
def get_backend(path: str, engine: str):
    return ExcelAppBackend(path) if engine == "excel" else OpenpyxlFileBackend(path)


def backend():
    return get_backend(SETTINGS.get("file_path", ""), ENGINE)


@st.cache_data(ttl=30, show_spinner="Reading the tracker workbook…")
def load(sheet: str) -> pd.DataFrame:
    return backend().read_sheet(sheet)


def refresh():
    load.clear()


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


# ----------------------------------------------------------------------------- workbook settings
def workbook_settings():
    """Point the app at the workbook on the host PC and prepare it."""
    cur = SETTINGS.get("file_path", "")
    st.write("Enter the path of the shared tracker workbook **on the PC running this app**, inside a Box Drive, "
             "OneDrive or synced SharePoint folder, which keeps it in sync with everyone else.")
    if cur:
        st.caption(f"In use now: `{cur}` (set by {SETTINGS.get('set_by', '?')} on {SETTINGS.get('set_at', '?')})")

    if "wb_path" not in st.session_state:
        st.session_state["wb_path"] = cur
    with st.form("wb_path_form", border=True):
        typed = st.text_input("📁 Workbook file path", st.session_state["wb_path"],
                              placeholder=r"C:\Users\you\OneDrive - World Vision\Migration Tracker.xlsx",
                              help="In File Explorer: Shift + right-click the file → Copy as path, then paste it here. "
                                   "Paths in your own user folder are saved as %USERPROFILE%\\…, so the same "
                                   "setting finds the file on any colleague's PC with the same Box folder.")
        if st.form_submit_button("Open this workbook", type="primary"):
            st.session_state["wb_path"] = typed
    with st.expander("🔎 Don't know the path? Search the Box / OneDrive / SharePoint folders on this PC"):
        term = st.text_input("File name contains", "migration")
        if st.button("Search"):
            st.session_state["found"] = find_workbooks(term)
        found = st.session_state.get("found")
        if found is not None:
            if found:
                pick = st.radio("Found", found, index=None)
                if pick and pick != st.session_state["wb_path"]:
                    st.session_state["wb_path"] = pick
                    st.rerun()
            else:
                st.caption("Nothing found. Paste the full path above instead.")

    path = local_path(st.session_state["wb_path"].strip().strip('"'))
    engines = {"excel": "Microsoft Excel (recommended — keeps charts/PivotTables, formulas)",
               "openpyxl": "Direct file write (data-only workbooks, no Excel needed)"}
    options = list(engines) if excel_available() else ["openpyxl"]
    engine = st.radio("Engine", options, format_func=engines.get,
                      index=options.index(ENGINE) if ENGINE in options else 0)

    if not path:
        return
    p = Path(path)
    if not p.exists():
        st.error("No file at that path on this PC. Check the path, and that Box Drive / OneDrive is running and signed in.")
        return
    if p.suffix.lower() not in (".xlsx", ".xlsm"):
        st.error("Pick an .xlsx or .xlsm workbook.")
        return
    fs = file_status(path)
    c = st.columns(3)
    c[0].metric("Size", f"{fs['size_kb']} KB")
    c[1].metric("Last modified", fs["modified"])
    c[2].metric("Kept in sync by", fs["synced_by"] or "Nothing")
    if not fs["synced_by"]:
        st.warning("This file isn't inside a Box Drive, OneDrive or SharePoint folder on this PC, so colleagues won't see the changes.")
    if engine == "openpyxl" and fs["rich"]:
        st.error(f"This workbook has {', '.join(fs['rich'])}. A direct file write would delete them — use the Excel engine.")
        return

    be = get_backend(path, engine)
    try:
        with st.spinner("Opening the workbook…"):
            info = be.info()
    except Exception as e:  # noqa: BLE001
        st.error(f"Couldn't open it: {e}")
        return
    if engine == "excel":
        if info["cloud"]:
            st.success("Excel opened it as a **cloud document with AutoSave** — saves merge live with colleagues' edits.")
        else:
            st.info(f"Excel opens it as a plain file: each save is written to disk, closed straight away, and "
                    f"{fs['synced_by'] or 'the sync app'} uploads it within seconds. Ask colleagues to update through "
                    f"this app rather than editing the file in Excel at the same time, or their copy may clash with it.")
    sheets = info["sheets"]
    tracker_ok = CFG["tracker_sheet"] in sheets
    missing_sheets = [s for s in (CFG["analysts_sheet"], CFG["log_sheet"]) if s not in sheets]
    st.write(f"Sheets: {', '.join(sheets)}")
    if not tracker_ok:
        st.error(f"There's no '{CFG['tracker_sheet']}' sheet in this workbook.")
        return
    cols = [] if not tracker_ok else [c for c in be.read_sheet(CFG["tracker_sheet"]).columns if c != "_row"]
    missing_cols = [c for c in TRACKING_COLS if c not in cols]
    if missing_sheets or missing_cols:
        st.warning("The workbook needs: " + ", ".join([f"sheet '{s}'" for s in missing_sheets] +
                                                    [f"column '{c}'" for c in missing_cols]))
        if st.button("Set up the workbook", type="primary"):
            with st.spinner("Adding sheets and columns…"):
                done = be.setup(CFG["tracker_sheet"], TRACKING_COLS,
                                [(CFG["analysts_sheet"], "Analysts", ["Name", "Email", "Team", "Active", "Role"], DEFAULT_ROSTER),
                                 (CFG["log_sheet"], CFG["log_table"], LOG_HEADERS, [])])
            st.success("Done: " + "; ".join(done) + ". Check the emails on the Analysts sheet.")
            st.rerun()
        return
    if path != cur or engine != SETTINGS.get("engine"):
        if st.button("Use this workbook", type="primary"):
            write_settings({"file_path": path, "engine": engine, "set_by": st.session_state.get("me") or "setup", "set_at": now()})
            st.session_state["flash"] = f"Tracker now uses {p.name}."
            st.rerun()
    else:
        st.success("✅ This is the workbook in use.")


# ----------------------------------------------------------------------------- data
def setup_screen(reason: str):
    st.title("Choose the tracker workbook")
    st.info(reason)
    workbook_settings()
    st.stop()


if not SETTINGS.get("file_path"):
    setup_screen("No workbook chosen yet.")
try:
    tracker = load(CFG["tracker_sheet"])
    roster = load(CFG["analysts_sheet"])
    log = load(CFG["log_sheet"])
except Exception as e:  # noqa: BLE001
    setup_screen(f"Couldn't read the tracker workbook: {e}")

CHECK = checklist_cols(tracker.columns)
HEADERS = [c for c in tracker.columns if c != "_row"]
tracker["Live Status"] = tracker.apply(lambda r: status_of(r, CHECK), axis=1)
for c in ("Name", "Email", "Team", "Active", "Role"):
    if c not in roster:
        roster[c] = ""
roster["Active"] = roster["Active"].str.strip().str.lower().isin(["yes", "y", "true"])
ACTIVE = roster[roster["Active"]]["Name"].tolist()


# ----------------------------------------------------------------------------- identity
def whoami() -> dict:
    """No sign-in: each person picks their own name. Saves are stamped with it."""
    names = ACTIVE + [n for n in roster["Name"].tolist() if n not in ACTIVE]
    if not names:
        setup_screen("The Analysts sheet is empty — add the team there first.")
    saved = st.session_state.get("me") or st.query_params.get("me")  # survives a page reload / bookmark
    if saved not in names:
        st.title("Azure AD Join Migration Tracker")
        pick = st.selectbox("Who are you?", names, index=None, placeholder="Pick your name…")
        st.caption("Every save is recorded under the name you pick. Not listed? Ask an admin to add you "
                   "on the Analysts sheet.")
        if pick is None:
            st.stop()
        st.session_state["me"] = pick
        st.rerun()
    pick = st.sidebar.selectbox("I am", names, index=names.index(saved))
    st.session_state["me"] = pick
    st.query_params["me"] = pick
    r = roster[roster["Name"] == pick].iloc[0]
    return {"name": r["Name"], "email": r["Email"], "team": r["Team"],
            "admin": r["Role"].strip().lower() == "admin"}


ME = whoami()
with st.sidebar:
    st.markdown(f"{ME['team'] or ''} {'· Admin' if ME['admin'] else ''}")
    if st.button("↻ Refresh data"):
        refresh()
        st.rerun()
    st.caption("Workbook: " + Path(SETTINGS.get("file_path", "")).name)


def write_log(entries: list[list]):
    backend().append_rows(CFG["log_sheet"], CFG["log_table"], entries)


# ----------------------------------------------------------------------------- pages
def page_update():
    st.title("Update an endpoint")
    if msg := st.session_state.pop("flash", None):
        st.success(msg, icon="✅")
    mine = tracker[tracker["Assigned Analyst"].str.strip() == ME["name"]]
    c1, c2 = st.columns([1, 1])
    scope = c1.radio("Show", ["My endpoints", "All outstanding", "Everyone"], horizontal=True,
                     index=0 if len(mine) else 1)
    pool = {"My endpoints": mine,
            "All outstanding": tracker[tracker["Live Status"] != "Complete"],
            "Everyone": tracker}[scope]
    locs = ["All locations"] + sorted(pool["Location"].unique())
    loc = c2.selectbox("Location", locs)
    if loc != "All locations":
        pool = pool[pool["Location"] == loc]

    icon = {"Complete": "✅", "In Progress": "🟡", "Not Started": "⚪"}
    labels = {int(r["_row"]): f"{icon[r['Live Status']]} {r['Client Name'].strip()} · {r['Employee ID']} · {r['Location']}"
              for _, r in pool.sort_values(["Live Status", "Client Name"], ascending=[False, True]).iterrows()}
    row_no = st.selectbox(f"Staff member ({len(labels)})", list(labels), format_func=labels.get,
                          index=None, placeholder="Type a name or Employee ID…")
    if row_no is None:
        st.info("Pick a staff member above. Then tick what's done and press Save — "
                "or **Mark all done & save** if the device is finished.")
        return

    r = tracker[tracker["_row"] == row_no].iloc[0]
    assigned = r["Assigned Analyst"].strip()
    a, b, c, d = st.columns(4)
    a.metric("Location", r["Location"])
    b.metric("Assigned to", assigned or "Unassigned")
    c.metric("Checklist", f"{score(r.to_dict(), CHECK)[0]}/{len(CHECK)}")
    d.metric("Last updated", r.get("Last Updated By", "") or "—", r.get("Last Updated At", "") or None,
             delta_color="off")
    if assigned and assigned != ME["name"]:
        st.warning(f"This endpoint is assigned to **{assigned}**. Your save will be logged under your name.")

    ver = r.get("Last Updated At", "")  # new widget keys after each save so the form shows saved values
    with st.form(f"f{row_no}-{ver}", border=True):
        x, y, z = st.columns(3)
        email = x.text_input("Client email", r["Client Email"])
        serial = y.text_input("Device serial number", r["Device Serial Number"])
        devname = z.text_input("WVI device name", r["WVI Device Name"], placeholder="UG-<serial> (auto-filled)")
        st.markdown("**Pre-flight checklist**")
        ticks = {}
        cols = st.columns(2)
        for i, col in enumerate(CHECK):
            ticks[col] = cols[i % 2].checkbox(col, value=is_yes(r[col]), key=f"{row_no}-{i}-{ver}")
        notes = st.text_area("Notes / blockers", r["Notes / Blockers"], height=70)
        take = st.checkbox("Assign this endpoint to me", value=not assigned)
        s1, s2, _ = st.columns([1, 1.4, 3])
        save = s1.form_submit_button("Save", type="primary")
        save_all = s2.form_submit_button("✓ Mark all done & save")

    if save or save_all:
        if save_all:
            ticks = {k: True for k in ticks}
        save_endpoint(r, ticks, email.strip(), serial.strip(), devname.strip(), notes.strip(), take)


def save_endpoint(r, ticks, email, serial, devname, notes, take):
    new = {col: ("Yes" if t else ("No" if r[col].strip() else "")) for col, t in ticks.items()}
    done, pct, status = score(new, CHECK)
    if status == "Complete" and not serial:
        st.error("Enter the device serial number before marking the endpoint complete.")
        return
    if serial and not devname:
        devname = f"UG-{serial.upper()}"
    new.update({"Client Email": email, "Device Serial Number": serial, "WVI Device Name": devname,
                "Notes / Blockers": notes, "Items Complete": done, "% Complete": pct, "Status": status})
    if take:
        new["Assigned Analyst"] = ME["name"]
    if done and not r.get("Date Started", "").strip():
        new["Date Started"] = datetime.now().strftime("%Y-%m-%d")
    if status == "Complete" and not r.get("Date Completed", "").strip():
        new["Date Completed"] = datetime.now().strftime("%Y-%m-%d")
    if status != "Complete" and r.get("Date Completed", "").strip():
        new["Date Completed"] = ""

    changed = {k: v for k, v in new.items() if k in HEADERS and str(v) != str(r.get(k, ""))}
    if not changed:
        st.info("Nothing changed.")
        return

    be = backend()
    # Guard against the sheet having been sorted/edited since we loaded it.
    live = be.read_row(CFG["tracker_sheet"], int(r["_row"]), HEADERS)
    if live.get("Employee ID", "") != r["Employee ID"]:
        refresh()
        st.error("Rows in the workbook have moved since you loaded it (someone sorted or inserted rows). "
                 "Data reloaded — pick the staff member again.")
        return
    if live.get("Last Updated At", "") != r.get("Last Updated At", ""):
        refresh()
        st.error(f"{live.get('Last Updated By') or 'Someone'} updated this endpoint at "
                 f"{live.get('Last Updated At')} while you were editing. Data reloaded — check and save again.")
        return

    changed.update({"Last Updated By": ME["name"], "Last Updated At": now()})
    try:
        kwargs = {"headers": HEADERS}
        be.write_cells(CFG["tracker_sheet"], [(int(r["_row"]), changed)], **kwargs)
        diff = "; ".join(f"{k.split('(')[0].strip()}: {r.get(k, '') or '∅'} → {v or '∅'}"
                         for k, v in changed.items() if k not in ("Last Updated By", "Last Updated At",
                                                                   "Items Complete", "% Complete"))
        action = "Completed" if status == "Complete" and r["Live Status"] != "Complete" else "Updated"
        write_log([[now(), ME["name"], ME["email"], r["Employee ID"], r["Client Name"].strip(),
                    r["Location"], action, done, status, diff[:1000]]])
    except Exception as e:  # noqa: BLE001
        st.error(f"Save failed: {e}")
        return
    refresh()
    st.session_state["flash"] = f"Saved: **{r['Client Name'].strip()}** is now **{status}** ({done}/{len(CHECK)})."
    st.rerun()


def page_mine():
    st.title("My workload")
    mine = tracker[tracker["Assigned Analyst"].str.strip() == ME["name"]]
    week = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
    my_log = log[(log["Analyst"] == ME["name"])] if len(log) else log
    c = st.columns(5)
    c[0].metric("Assigned", len(mine))
    c[1].metric("Complete", int((mine["Live Status"] == "Complete").sum()))
    c[2].metric("In progress", int((mine["Live Status"] == "In Progress").sum()))
    c[3].metric("Not started", int((mine["Live Status"] == "Not Started").sum()))
    c[4].metric("Completed last 7 days",
                int(((my_log["Action"] == "Completed") & (my_log["Timestamp"] >= week)).sum()) if len(my_log) else 0)
    if len(mine):
        st.progress((mine["Live Status"] == "Complete").mean(), text="Your completion")
    todo = mine[mine["Live Status"] != "Complete"]
    st.subheader(f"Still to do ({len(todo)})")
    if len(todo):
        st.dataframe(todo.groupby(["Region", "Location"]).size().rename("Endpoints").reset_index()
                     .sort_values("Endpoints", ascending=False), hide_index=True, use_container_width=True)
        st.dataframe(todo[["Client Name", "Employee ID", "Location", "Live Status", "Notes / Blockers"]],
                     hide_index=True, use_container_width=True)
    else:
        st.success("Nothing outstanding.")


def page_team():
    st.title("Team dashboard")
    st.caption("Every figure is calculated from the Endpoint Tracker checklist — nothing on this page is typed in.")
    total, comp = len(tracker), int((tracker["Live Status"] == "Complete").sum())
    c = st.columns(4)
    c[0].metric("Endpoints", total)
    c[1].metric("Complete", comp, f"{comp / total:.0%}" if total else None, delta_color="off")
    c[2].metric("In progress", int((tracker["Live Status"] == "In Progress").sum()))
    c[3].metric("Not started", int((tracker["Live Status"] == "Not Started").sum()))
    st.progress(comp / total if total else 0)

    week = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
    t = tracker.assign(Analyst=tracker["Assigned Analyst"].str.strip().replace("", "(unassigned)"))
    by = t.groupby("Analyst")["Live Status"].value_counts().unstack(fill_value=0)
    for col in ["Complete", "In Progress", "Not Started"]:
        if col not in by:
            by[col] = 0
    by["Assigned"] = by[["Complete", "In Progress", "Not Started"]].sum(axis=1)
    by["% Complete"] = (by["Complete"] / by["Assigned"]).round(3)
    if len(log):
        recent = log[log["Timestamp"] >= week]
        by["Completed (7 days)"] = recent[recent["Action"] == "Completed"].groupby("Analyst").size()
        by["Updates (7 days)"] = recent.groupby("Analyst").size()
        by["Last activity"] = log.groupby("Analyst")["Timestamp"].max()
    by = by.fillna(0).reset_index()
    by["Active"] = by["Analyst"].map(lambda a: "Yes" if a in ACTIVE else "")
    st.subheader("By analyst")
    st.dataframe(by, hide_index=True, use_container_width=True,
                 column_config={"% Complete": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1)})
    st.bar_chart(by.set_index("Analyst")[["Complete", "In Progress", "Not Started"]], horizontal=True)

    st.subheader("By region")
    rg = tracker.groupby("Region")["Live Status"].value_counts().unstack(fill_value=0)
    rg["Target"] = rg.sum(axis=1)
    rg["% Complete"] = (rg.get("Complete", 0) / rg["Target"]).round(3)
    rg = rg.reindex([r for r in GEO_ORDER if r in rg.index] + [r for r in rg.index if r not in GEO_ORDER])
    st.dataframe(rg.reset_index(), hide_index=True, use_container_width=True,
                 column_config={"% Complete": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1)})

    issues = data_issues(tracker, CHECK)
    with st.expander(f"⚠️ Data issues to fix ({len(issues)})"):
        if len(issues):
            st.dataframe(issues["Issue"].str.replace(r"Serial .*", "Serial on more than one row", regex=True)
                         .str.replace(r"Bad email.*", "Bad email", regex=True)
                         .str.replace(r"Status column says.*", "Status column disagrees with checklist", regex=True)
                         .value_counts().rename("Rows"), use_container_width=True)
            st.dataframe(issues.drop(columns="_row"), hide_index=True, use_container_width=True)


def page_log():
    st.title("Activity log")
    st.caption("Every save made through this app, with who made it and what changed.")
    if not len(log):
        st.info("No activity yet.")
        return
    who = st.multiselect("Analyst", sorted(log["Analyst"].unique()))
    view = log[log["Analyst"].isin(who)] if who else log
    st.dataframe(view.drop(columns="_row").sort_values("Timestamp", ascending=False),
                 hide_index=True, use_container_width=True)


def page_admin():
    st.title("Admin")
    if msg := st.session_state.pop("flash", None):
        st.success(msg, icon="✅")
    tab1, tab2 = st.tabs(["Share out outstanding endpoints", "Team roster"])
    with tab1:
        st.write("Splits outstanding endpoints evenly. Each analyst gets a connected block of "
                 "neighbouring sites; anything assigned to inactive or exited staff is always reassigned.")
        who = st.multiselect("Share between", roster["Name"].tolist(), default=ACTIVE)
        inc = st.checkbox("Also move in-progress endpoints (normally they stay with whoever started them)")
        if not who:
            return
        moves = rebalance(tracker, who, CHECK, include_in_progress=inc)
        changes = moves[moves["From"] != moves["To"]]
        out = tracker[tracker["Live Status"] != "Complete"]
        before = out["Assigned Analyst"].str.strip().replace("", "(unassigned)").value_counts()
        after = before.copy()
        for _, m in changes.iterrows():
            after[m["From"] or "(unassigned)"] -= 1
            after[m["To"]] = after.get(m["To"], 0) + 1
        summary = pd.DataFrame({"Outstanding now": before, "After": after}).fillna(0).astype(int)
        summary = summary[(summary.sum(axis=1)) > 0]
        st.dataframe(summary, use_container_width=True)
        st.write(f"**{len(changes)}** endpoints will change hands.")
        st.dataframe(changes.drop(columns="_row").rename(columns={"_status": "Status"}),
                     hide_index=True, use_container_width=True, height=280)
        if len(changes) and st.button(f"Apply — reassign {len(changes)} endpoints", type="primary"):
            be = backend()
            kwargs = {"headers": HEADERS}
            with st.spinner("Writing to the workbook…"):
                be.write_cells(CFG["tracker_sheet"],
                               [(int(m["_row"]), {"Assigned Analyst": m["To"]}) for _, m in changes.iterrows()], **kwargs)
                write_log([[now(), ME["name"], ME["email"], m["Employee ID"], m["Client Name"].strip(),
                            m["Location"], "Reassigned", "", m["_status"], f"{m['From'] or 'unassigned'} → {m['To']}"]
                           for _, m in changes.iterrows()])
            refresh()
            st.session_state["flash"] = f"Reassigned {len(changes)} endpoints. Each move is in the Activity log."
            st.rerun()

    with tab2:
        st.write("Mark leavers **Active = No** (their history is kept). Add new analysts in the empty row at the bottom. "
                 "Set Role = Admin for people who may reassign work and edit the roster.")
        ed = roster.assign(Active=roster["Active"].map({True: "Yes", False: "No"}))
        edited = st.data_editor(ed, num_rows="dynamic", hide_index=True, use_container_width=True,
                                column_config={"_row": None,
                                               "Active": st.column_config.SelectboxColumn(options=["Yes", "No"]),
                                               "Role": st.column_config.SelectboxColumn(options=["Analyst", "Admin", "Exited"])})
        if st.button("Save roster", type="primary"):
            cols = ["Name", "Email", "Team", "Active", "Role"]
            updates, new_rows = [], []
            for _, e in edited.iterrows():
                if pd.isna(e.get("_row")):
                    if str(e.get("Name") or "").strip():
                        new_rows.append([str(e.get(c) or "").strip() for c in cols])
                    continue
                old = ed[ed["_row"] == e["_row"]].iloc[0]
                diff = {c: str(e[c] or "").strip() for c in cols if str(e[c] or "").strip() != str(old[c]).strip()}
                if diff:
                    updates.append((int(e["_row"]), diff))
            be = backend()
            if updates:
                kwargs = {"headers": [c for c in roster.columns if c != "_row"]}
                be.write_cells(CFG["analysts_sheet"], updates, **kwargs)
            if new_rows:
                be.append_rows(CFG["analysts_sheet"], "Analysts", new_rows)
            refresh()
            st.session_state["flash"] = f"Roster saved ({len(updates)} changed, {len(new_rows)} added)."
            st.rerun()


def page_workbook():
    st.title("Workbook")
    if msg := st.session_state.pop("flash", None):
        st.success(msg, icon="✅")
    workbook_settings()


pages = [st.Page(page_update, title="Update endpoint", icon="✏️", default=True),
         st.Page(page_mine, title="My workload", icon="📋"),
         st.Page(page_team, title="Team dashboard", icon="📊"),
         st.Page(page_log, title="Activity log", icon="🕑"),
         st.Page(page_workbook, title="Workbook", icon="📁")]
if ME["admin"]:
    pages.append(st.Page(page_admin, title="Admin", icon="⚙️"))
st.navigation(pages, position="top").run()
