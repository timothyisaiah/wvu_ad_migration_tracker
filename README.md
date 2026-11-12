# Azure AD Join Migration Tracker (Streamlit)

A web form analysts use instead of typing into the shared Excel. It writes straight into
the tracker workbook, stamps **who** did each update and **when**, and keeps an **Activity Log**
of every change. Dashboards are calculated from the checklist, never typed.

| Page | Who | What it does |
|---|---|---|
| Update endpoint | Everyone | Pick a staff member → tick items → **Save**, or **Mark all done & save** (3 clicks) |
| My workload | Everyone | My assigned endpoints, what's left, by location |
| Team dashboard | Everyone | Totals, per-analyst and per-region progress, data-quality issues |
| Activity log | Everyone | Every save: who, when, what changed |
| Workbook | Everyone | Enter the workbook's file path; sets it up on first use |
| Admin | Role = Admin | Share outstanding endpoints evenly · team roster |

## How it fits together

```
Analysts' browsers ──(open the URL, pick their name)──► Streamlit app on the HOST PC
                                                            │  edits the workbook through Excel
                                                            ▼
                         C:\Users\…\World Vision\…\Migration Tracker.xlsx   (OneDrive-synced)
                                                            │  OneDrive / SharePoint sync
                                                            ▼
                                           Everyone else viewing it in Excel Online
```

* The app runs on **one Windows PC or server** (the host) with Excel and Box Drive (or OneDrive) signed in
  with an account that can edit the shared file.
* Open the URL, go to **Workbook**, and enter the file's path **on the host PC**, e.g. a file under
  `C:\Users\<you>\Box\…`. Box Drive (or OneDrive) keeps it in sync.
  A website can't reach files on a *visitor's* PC, so the path is always a path on the host.
* There is **no sign-in**. Each person picks their name from the Analysts sheet, and every save is
  stamped with it. Anyone who can reach the URL can use it, so keep the URL inside WVI
  (office network, VPN or Tailscale) rather than on the open internet.

### Engines

| Engine | When | Notes |
|---|---|---|
| **Microsoft Excel** (default) | Host has Excel | A hidden Excel edits the file, so PivotTables, charts, formulas and formatting are untouched. If Excel opens it as a cloud document (AutoSave on), each save merges cell-by-cell with colleagues' edits in Excel Online. Otherwise (always the case for Box files) the file is opened only for each save and closed straight after, so Box/OneDrive can upload it. Colleagues should update through the app rather than editing the file in Excel at the same moment. |
| Direct file write | No Excel on host | Rewrites the whole file. **Refuses** any workbook with PivotTables, charts, slicers or pictures (it would delete them) — the real tracker has these, so use Excel. |

Formula cells (Items Complete, % Complete, Status in the current workbook) are never overwritten.

## 1. Install on the host PC

```
pip install -r requirements.txt
```

Use a PC that stays on during working hours. Sign Office (Excel) and OneDrive in with an account
that can edit the tracker, and sync the SharePoint library / OneDrive folder that holds it.
**Don't keep the tracker open in desktop Excel on the host** — view it in Excel Online instead.

## 2. Run it

```
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

Colleagues open `http://<host PC name or IP>:8501`. To keep it running, add a Task Scheduler task
(*At log on* → `streamlit run …`); Excel automation needs a signed-in desktop session.
For people outside the office, use **Tailscale** or the office VPN rather than a public link.

## 3. Point it at the workbook (first time)

1. Open the app → **Workbook**. Paste the path (File Explorer: Shift + right-click → *Copy as path*),
   or use **Find** to search the host's OneDrive/SharePoint folders.
2. **Set up the workbook** — adds `Last Updated By` / `Last Updated At` columns to *Endpoint Tracker*,
   and *Analysts* and *Activity Log* sheets. Nothing else in the workbook is touched.
3. On the new **Analysts** sheet check names, Active and Role (Role = Admin unlocks the Admin page).
   Delete the `SAMPLE-001` row, and fix items under Team dashboard → *Data issues*.
4. **Use this workbook.** The choice is saved in `settings.json` next to the app and applies to everyone.

## How it keeps the numbers straight

- Status, Items Complete and % Complete come from the 10 checklist ticks.
- Analyst name, Date Started, Date Completed, Last Updated By/At are filled in automatically.
- Completing an endpoint requires the device serial; the WVI device name defaults to `UG-<serial>`.
- If someone saved the same endpoint while you were editing, or rows were sorted/inserted, the
  save is stopped and the data reloads.
- All Excel work goes through one queue, so two analysts saving at once can't collide.

## Files

| File | Purpose |
|---|---|
| `app.py` | The Streamlit app |
| `excel_file.py` | Excel and direct-file engines, OneDrive file search |
| `backend.py` | Shared helpers |
| `logic.py` | Checklist scoring, even reassignment, data-quality checks, default roster |
| `settings.json` | Which workbook and engine (written by the Workbook page) |
| `seed_workbook.py` | Builds a test workbook from the CSV export |
