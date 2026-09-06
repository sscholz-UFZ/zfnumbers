"""
zf_ufz_statistics.py
======================

Load and analyse the monthly "Bestandserfassung" (stock/inventory) Excel
exports from:

    C:\\Users\\fish\\Documents\\Fischzucht\\Bestandserfassung

Each file covers one month and is named like "BestandAugust2026.xls" -
the month and year are only encoded in the *file name*, not inside the
sheet itself, so they're parsed out of it here and added as columns to
every row, before all months are stacked into a single table.

GUI workflow
------------
1. Load Data: browse to the folder (defaults to the path above) and load
   every .xls/.xlsx file in it.
2. Inspect Data: the combined table is shown in a scrollable grid.
3. Analysis: optionally restrict to one Year / Line / Strain and/or a DOB
   (date of birth) date range, then either
   - "Time series per batch code": one small line graph per batch code
     (there can be many - shown as a scrollable grid of small graphs), or
   - "Total Num. animals per year/month": a single line graph of the
     summed Num. animals per month, for whatever the current filters
     select - e.g. selecting one Line or Strain and leaving Year at "All"
     shows that line/strain's total over time; adding a Year filter
     restricts any analysis to a single year.

Notes on the source files
--------------------------
- They are old-format .xls files (not .xlsx), which pandas reads via the
  `xlrd` package - make sure it's installed (`pip install xlrd`).
- Each file has exactly one sheet, but the sheet's *name* includes the
  export date and therefore differs between files (e.g.
  "Animals02.09.2026") - so the sheet is always selected by position
  (the first one), never by name.
- xlrd sometimes prints a harmless "OLE2 inconsistency" warning for these
  particular files; the data still reads correctly despite it.

Requirements
------------
    pip install pandas xlrd openpyxl matplotlib tkcalendar

Run:
    python zf_ufz_statistics.py
"""

import io
import math
import re
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import quote, unquote, urlparse

import pandas as pd
import requests
from requests.adapters import HTTPAdapter

import tkinter as tk
from tkinter import ttk, messagebox

from tkcalendar import DateEntry
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

DATA_FOLDER = Path(r"C:\Users\fish\Documents\Fischzucht\Bestandserfassung")
NUM_ANIMALS_COL = "Num. animals"
BATCH_COL = "Batch code"

# Recognized month names -> month number. Covers German (with and without
# the "ä" in März, in case a filename avoids special characters) and
# English; months spelled identically in both languages (April, August,
# September, November) only need to appear once.
MONTH_NAME_TO_NUM = {
    "januar": 1, "january": 1,
    "februar": 2, "february": 2,
    "märz": 3, "maerz": 3, "march": 3,
    "april": 4,
    "mai": 5, "may": 5,
    "juni": 6, "june": 6,
    "juli": 7, "july": 7,
    "august": 8,
    "september": 9,
    "oktober": 10, "october": 10,
    "november": 11,
    "dezember": 12, "december": 12,
}

# Canonical display name for each month number (German, matching the data's
# own language) - used regardless of whether a given filename happened to
# spell the month in German or English, so the resulting column is
# consistent across all loaded files.
MONTH_NUM_TO_NAME = {
    1: "Januar", 2: "Februar", 3: "März", 4: "April", 5: "Mai", 6: "Juni",
    7: "Juli", 8: "August", 9: "September", 10: "Oktober", 11: "November",
    12: "Dezember",
}

# Longest names first, so e.g. a hypothetical shorter name can't shadow a
# longer one that contains it as a substring.
_MONTH_PATTERN = re.compile(
    "(" + "|".join(sorted(MONTH_NAME_TO_NUM, key=len, reverse=True)) + ")",
    re.IGNORECASE,
)
_YEAR_PATTERN = re.compile(r"(19|20)\d{2}")


def parse_month_year_from_filename(filename: str):
    """Extract (month_num, year) from a file name such as
    "BestandAugust2026.xls". Raises ValueError if either isn't found."""
    stem = Path(filename).stem
    month_match = _MONTH_PATTERN.search(stem)
    year_match = _YEAR_PATTERN.search(stem)
    if not month_match:
        raise ValueError(f"No recognizable month name found in: {filename}")
    if not year_match:
        raise ValueError(f"No 4-digit year found in: {filename}")
    month_num = MONTH_NAME_TO_NUM[month_match.group(1).lower()]
    year = int(year_match.group(0))
    return month_num, year


def load_bestand_folder(folder=DATA_FOLDER, verbose=True):
    """Read every .xls/.xlsx file in `folder`, tag each row with the month
    and year parsed from its file name, and return one combined DataFrame."""
    folder = Path(folder)
    paths = sorted(folder.glob("*.xls")) + sorted(folder.glob("*.xlsx"))
    if not paths:
        raise FileNotFoundError(f"No .xls/.xlsx files found in: {folder}")

    frames = []
    for path in paths:
        try:
            month_num, year = parse_month_year_from_filename(path.name)
        except ValueError as e:
            print(f"[warn] Skipping {path.name}: {e}")
            continue

        df = pd.read_excel(path, sheet_name=0)  # first sheet - its name varies per file
        df["source_file"] = path.name
        df["month"] = MONTH_NUM_TO_NAME[month_num]
        df["month_num"] = month_num
        df["year"] = year

        # Put the new columns first, for readability when inspecting the table.
        meta_cols = ["source_file", "month", "month_num", "year"]
        df = df[meta_cols + [c for c in df.columns if c not in meta_cols]]

        frames.append(df)
        if verbose:
            print(f"Loaded {path.name}: {len(df)} rows  (month={MONTH_NUM_TO_NAME[month_num]}, year={year})")

    if not frames:
        raise ValueError(f"No files with a recognizable month/year were loaded from: {folder}")

    combined = pd.concat(frames, ignore_index=True)
    if verbose:
        print(f"\nCombined: {len(combined)} rows from {len(frames)} file(s).")
    return combined


# WebDAV request body for listing a folder's contents (Depth: 1 header) -
# requesting just resourcetype is enough to tell files apart from
# sub-folders, which is all the listing step needs.
_WEBDAV_PROPFIND_BODY = (
    '<?xml version="1.0"?>'
    '<d:propfind xmlns:d="DAV:"><d:prop><d:resourcetype/></d:prop></d:propfind>'
)
_WEBDAV_NS = {"d": "DAV:"}


def _list_nextcloud_excel_files(session, folder_url):
    """PROPFIND `folder_url` and return the .xls/.xlsx file names directly
    inside it (not sub-folders, not recursing into them)."""
    resp = session.request(
        "PROPFIND", folder_url, headers={"Depth": "1"}, data=_WEBDAV_PROPFIND_BODY, timeout=30,
    )
    resp.raise_for_status()
    root = ET.fromstring(resp.content)

    names = []
    for response_el in root.findall("d:response", _WEBDAV_NS):
        href = response_el.findtext("d:href", default="", namespaces=_WEBDAV_NS)
        is_folder = response_el.find(".//d:resourcetype/d:collection", _WEBDAV_NS) is not None
        if is_folder:
            continue
        name = unquote(href.rstrip("/").rsplit("/", 1)[-1])
        if name.lower().endswith((".xls", ".xlsx")):
            names.append(name)
    return sorted(set(names))


# How many files to download at once. Downloads are I/O-bound (waiting on
# the network, not the CPU), so a small thread pool lets several requests
# be in flight together instead of paying each one's round-trip latency
# back to back - this is what actually matters as the number of monthly
# files grows over time, unlike e.g. compressing the request.
_MAX_CONCURRENT_DOWNLOADS = 6


def _download_one_excel(session, folder_url, name):
    """Download and parse a single Excel file - the part of the loop in
    _load_bestand_from_webdav that's slow enough (one HTTPS round trip
    each) to be worth running several of at once."""
    file_resp = session.get(folder_url + quote(name), timeout=60)
    file_resp.raise_for_status()
    engine = "xlrd" if name.lower().endswith(".xls") else "openpyxl"
    return pd.read_excel(io.BytesIO(file_resp.content), sheet_name=0, engine=engine)


def _load_bestand_from_webdav(session, folder_url, location_desc, verbose):
    """Shared by both Nextcloud loaders below: list `folder_url` over
    WebDAV, download each .xls/.xlsx found directly in it (several at a
    time), tag each row with the month/year parsed from its file name
    (same as load_bestand_folder), and return one combined DataFrame."""
    filenames = _list_nextcloud_excel_files(session, folder_url)
    if not filenames:
        raise FileNotFoundError(f"No .xls/.xlsx files found in {location_desc}")

    to_fetch = []
    for name in filenames:
        try:
            month_num, year = parse_month_year_from_filename(name)
        except ValueError as e:
            if verbose:
                print(f"[warn] Skipping {name}: {e}")
            continue
        to_fetch.append((name, month_num, year))

    frames = [None] * len(to_fetch)
    with ThreadPoolExecutor(max_workers=min(_MAX_CONCURRENT_DOWNLOADS, len(to_fetch) or 1)) as pool:
        future_to_index = {
            pool.submit(_download_one_excel, session, folder_url, name): i
            for i, (name, _, _) in enumerate(to_fetch)
        }
        for future in as_completed(future_to_index):
            i = future_to_index[future]
            name, month_num, year = to_fetch[i]
            df = future.result()  # re-raises here if that download failed
            df["source_file"] = name
            df["month"] = MONTH_NUM_TO_NAME[month_num]
            df["month_num"] = month_num
            df["year"] = year

            meta_cols = ["source_file", "month", "month_num", "year"]
            df = df[meta_cols + [c for c in df.columns if c not in meta_cols]]

            frames[i] = df
            if verbose:
                print(f"Loaded {name}: {len(df)} rows  (month={MONTH_NUM_TO_NAME[month_num]}, year={year})")

    frames = [f for f in frames if f is not None]
    if not frames:
        raise ValueError(f"No files with a recognizable month/year were found in {location_desc}")

    combined = pd.concat(frames, ignore_index=True)
    if verbose:
        print(f"\nCombined: {len(combined)} rows from {len(frames)} file(s).")
    return combined


def _new_webdav_session():
    """A requests.Session with a large enough connection pool for
    _load_bestand_from_webdav's concurrent downloads to each get their own
    connection instead of queuing behind requests' small default pool."""
    session = requests.Session()
    adapter = HTTPAdapter(pool_maxsize=_MAX_CONCURRENT_DOWNLOADS)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


def load_bestand_from_nextcloud(url, username, app_password, folder="", verbose=True):
    """Fetch every .xls/.xlsx file from a folder in a regular Nextcloud
    *account* over WebDAV. See load_bestand_from_public_share() below for
    the simpler alternative that needs no account at all - just a share
    link (and its password, if it has one).

    `app_password` should be a Nextcloud *app password* (Settings > Security
    > "Create new app password"), not the account's real login password -
    it can be revoked independently without changing the main password.
    """
    folder_url = f"{url.rstrip('/')}/remote.php/dav/files/{quote(username)}/{quote(folder.strip('/'), safe='/')}/"
    with _new_webdav_session() as session:
        session.auth = (username, app_password)
        return _load_bestand_from_webdav(
            session, folder_url, f"Nextcloud folder: {folder or '/'}", verbose,
        )


def load_bestand_from_public_share(share_url, share_password="", verbose=True):
    """Fetch every .xls/.xlsx file from a Nextcloud *public share link*
    (e.g. https://nc.ufz.de/s/zebrafishufz) over WebDAV.

    Public shares use a different WebDAV endpoint than a normal account -
    https://<host>/public.php/webdav/ - authenticated with the share's own
    token (the part after "/s/" in the link) as the "username" and the
    share's password (empty string if it has none) as the "password". No
    real Nextcloud account or app password is involved, which makes this
    the simpler option for giving someone outside your own Nextcloud
    access to just this one folder.
    """
    parsed = urlparse(share_url)
    token = parsed.path.rstrip("/").rsplit("/", 1)[-1]
    if not token:
        raise ValueError(f"Could not find a share token (the part after \"/s/\") in: {share_url}")
    folder_url = f"{parsed.scheme}://{parsed.netloc}/public.php/webdav/"

    with _new_webdav_session() as session:
        session.auth = (token, share_password)
        return _load_bestand_from_webdav(session, folder_url, "the Nextcloud share", verbose)


def period_label(year, month_num):
    """Short, sortable-by-construction label like "06-2026" for plot axes."""
    return f"{month_num:02d}-{year}"


def y_axis_top_with_margin(values, margin_fraction=0.1, minimum=1):
    """Top limit for a Y axis pinned at bottom=0: the highest data point
    plus a small headroom margin, so a marker or line sitting right at the
    top of the data is never visually clipped by the plot's own edge.
    Accepts a plain list (NaN placeholders for missing months are fine -
    they're dropped) or a pandas Series."""
    numeric = [v for v in values if v == v]  # drop NaN (NaN != NaN)
    max_val = max(numeric, default=0)
    return max(max_val * (1 + margin_fraction), minimum)


# ----------------------------------------------------------------------
# GUI
# ----------------------------------------------------------------------
class BestandGUI:
    MANY_BATCHES_WARN_THRESHOLD = 80

    def __init__(self, root):
        self.root = root
        self.root.title("Zebrafish Stock (Bestand) Statistics")
        self.root.geometry("1150x780")

        self.data = None  # combined DataFrame once loaded
        self.folder = DATA_FOLDER  # fixed - not user-selectable
        self.status_var = tk.StringVar(value="No data loaded yet.")

        self.year_var = tk.StringVar(value="All")
        self.month_from_var = tk.StringVar(value="")
        self.month_to_var = tk.StringVar(value="")
        self.line_var = tk.StringVar(value="All")
        self.strain_var = tk.StringVar(value="All")
        self.mode_var = tk.StringVar(value="batch")  # "batch" or "total"

        self._build_ui()

    # ---------------- UI construction ----------------
    def _build_ui(self):
        upload = ttk.LabelFrame(self.root, text="1. Load Data")
        upload.pack(side="top", fill="x", padx=8, pady=6)
        row = ttk.Frame(upload)
        row.pack(side="top", fill="x", padx=6, pady=4)
        ttk.Label(row, text="Data is loaded from a fixed internal location - not user-selectable.",
                  foreground="gray30").pack(side="left")
        row2 = ttk.Frame(upload)
        row2.pack(side="top", fill="x", padx=6, pady=(0, 6))
        ttk.Button(row2, text="Load Data", command=self._on_load).pack(side="left")
        ttk.Label(row2, textvariable=self.status_var, foreground="blue").pack(side="left", padx=10)

        filters = ttk.LabelFrame(self.root, text="2. Filters")
        filters.pack(side="top", fill="x", padx=8, pady=6)
        ttk.Label(
            filters,
            text="Year and Month range refer to which monthly snapshot file(s) the data came "
                 "from (e.g. \"BestandAugust2026.xls\") - not a date recorded for an individual "
                 "animal. These filters apply to both the data table below and the analysis.",
            foreground="gray30", wraplength=1080, justify="left",
        ).pack(side="top", fill="x", padx=6, pady=(4, 2))

        row1 = ttk.Frame(filters)
        row1.pack(side="top", fill="x", padx=6, pady=4)
        ttk.Label(row1, text="Year:").pack(side="left")
        self.year_combo = ttk.Combobox(row1, textvariable=self.year_var, state="readonly",
                                        width=8, values=["All"])
        self.year_combo.pack(side="left", padx=(2, 14))
        ttk.Label(row1, text="Month range:").pack(side="left")
        self.month_from_combo = ttk.Combobox(row1, textvariable=self.month_from_var,
                                              state="disabled", width=10, values=[])
        self.month_from_combo.pack(side="left", padx=(2, 2))
        ttk.Label(row1, text="to").pack(side="left")
        self.month_to_combo = ttk.Combobox(row1, textvariable=self.month_to_var,
                                            state="disabled", width=10, values=[])
        self.month_to_combo.pack(side="left", padx=(2, 14))
        ttk.Label(row1, text="DOB from:").pack(side="left")
        self.dob_from = DateEntry(row1, date_pattern="yyyy-mm-dd", width=10, state="disabled")
        self.dob_from.pack(side="left", padx=(2, 6))
        ttk.Label(row1, text="to:").pack(side="left")
        self.dob_to = DateEntry(row1, date_pattern="yyyy-mm-dd", width=10, state="disabled")
        self.dob_to.pack(side="left", padx=(2, 6))
        self.dob_reset_button = ttk.Button(row1, text="Reset DOB", command=self._reset_dob_range,
                                            state="disabled")
        self.dob_reset_button.pack(side="left", padx=(4, 0))

        row2 = ttk.Frame(filters)
        row2.pack(side="top", fill="x", padx=6, pady=4)
        ttk.Label(row2, text="Line:").pack(side="left")
        self.line_combo = ttk.Combobox(row2, textvariable=self.line_var, state="readonly",
                                        width=20, values=["All"])
        self.line_combo.pack(side="left", padx=(2, 14))
        ttk.Label(row2, text="Strain:").pack(side="left")
        self.strain_combo = ttk.Combobox(row2, textvariable=self.strain_var, state="readonly",
                                          width=14, values=["All"])
        self.strain_combo.pack(side="left", padx=(2, 14))
        ttk.Button(row2, text="Apply Filters", command=self._refresh_tree).pack(side="left")

        # Auto-refresh the Inspect Data table whenever a filter changes -
        # the Plot button already always reads live filter values regardless.
        for combo in (self.year_combo, self.month_from_combo, self.month_to_combo,
                      self.line_combo, self.strain_combo):
            combo.bind("<<ComboboxSelected>>", lambda e: self._refresh_tree())
        for picker in (self.dob_from, self.dob_to):
            picker.bind("<<DateEntrySelected>>", lambda e: self._refresh_tree())

        inspect = ttk.LabelFrame(self.root, text="3. Inspect Data")
        inspect.pack(side="top", fill="both", expand=True, padx=8, pady=6)
        self.inspect_caption_var = tk.StringVar(value="")
        ttk.Label(inspect, textvariable=self.inspect_caption_var, foreground="gray30").pack(
            side="top", fill="x", padx=6, pady=(4, 0))
        tree_frame = ttk.Frame(inspect)
        tree_frame.pack(fill="both", expand=True, padx=6, pady=6)
        self.tree = ttk.Treeview(tree_frame, show="headings")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        analysis = ttk.LabelFrame(self.root, text="4. Analysis")
        analysis.pack(side="top", fill="x", padx=8, pady=6)
        mode_row = ttk.Frame(analysis)
        mode_row.pack(side="top", fill="x", padx=6, pady=4)
        ttk.Radiobutton(mode_row, text="Time series per batch code (many small graphs)",
                        variable=self.mode_var, value="batch").pack(side="left")
        ttk.Radiobutton(mode_row, text="Total Num. animals per year/month",
                        variable=self.mode_var, value="total").pack(side="left", padx=(20, 0))

        ttk.Button(analysis, text="Plot", command=self._on_plot).pack(side="left", padx=6, pady=(0, 6))

    # ---------------- load ----------------
    def _on_load(self):
        folder = self.folder
        if not folder.exists():
            messagebox.showerror("Not found", f"Folder not found:\n{folder}")
            return
        self.status_var.set("Loading...")
        self.root.update_idletasks()
        try:
            self.data = load_bestand_folder(folder, verbose=False)
        except Exception as e:
            messagebox.showerror("Load failed", str(e))
            self.status_var.set("Load failed.")
            return
        n_files = self.data["source_file"].nunique()
        self.status_var.set(f"Loaded {len(self.data)} rows from {n_files} file(s).")
        self._populate_filters()
        self._refresh_tree()

    def _populate_tree(self, df):
        self.tree.delete(*self.tree.get_children())
        cols = list(df.columns)
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=110, stretch=False)
        for _, row in df.iterrows():
            values = ["" if pd.isna(v) else str(v) for v in row]
            self.tree.insert("", "end", values=values)

    def _refresh_tree(self):
        """Re-apply the current filters to the Inspect Data table."""
        if self.data is None:
            return
        filtered = self._filtered_data()
        self._populate_tree(filtered)
        self.inspect_caption_var.set(
            f"{len(filtered)} of {len(self.data)} row(s) shown, based on the filters above: "
            f"{self._filter_subtitle()}"
        )

    def _populate_filters(self):
        years = sorted(self.data["year"].dropna().unique())
        self._available_months = sorted(self.data["month_num"].dropna().unique())
        month_names = [MONTH_NUM_TO_NAME[m] for m in self._available_months]
        lines = sorted(self.data["Line"].dropna().unique()) if "Line" in self.data.columns else []
        strains = sorted(self.data["Strain"].dropna().unique()) if "Strain" in self.data.columns else []
        self.year_combo["values"] = ["All"] + [str(int(y)) for y in years]
        self.line_combo["values"] = ["All"] + [str(v) for v in lines]
        self.strain_combo["values"] = ["All"] + [str(v) for v in strains]
        self.year_var.set("All")
        self.line_var.set("All")
        self.strain_var.set("All")

        self.month_from_combo.config(state="readonly", values=month_names)
        self.month_to_combo.config(state="readonly", values=month_names)
        self.month_from_var.set(month_names[0])
        self.month_to_var.set(month_names[-1])

        # DOB range: bound the pickers to the data's actual span and default
        # to covering all of it (i.e. no effective restriction until the
        # user narrows it).
        self._dob_min = self.data["DOB"].min().date()
        self._dob_max = self.data["DOB"].max().date()
        # Deliberately no mindate/maxdate: the pickers default to the data's
        # actual span, but the user should be free to pick any range at all
        # (e.g. "the last 3 months") - a range with no matching rows is a
        # valid, informative result, not something to block at the widget.
        for picker in (self.dob_from, self.dob_to):
            picker.config(state="normal")
        self.dob_reset_button.config(state="normal")
        self._reset_dob_range()

    def _reset_dob_range(self):
        self.dob_from.set_date(self._dob_min)
        self.dob_to.set_date(self._dob_max)
        self._refresh_tree()

    # ---------------- filtering ----------------
    def _dob_range_narrowed(self):
        """True if the DOB pickers have been moved in from the full data range."""
        return self.dob_from.get_date() != self._dob_min or self.dob_to.get_date() != self._dob_max

    def _month_range_nums(self):
        """(from_month_num, to_month_num) from the Month range combos."""
        from_num = MONTH_NAME_TO_NUM[self.month_from_var.get().lower()]
        to_num = MONTH_NAME_TO_NUM[self.month_to_var.get().lower()]
        return from_num, to_num

    def _month_range_narrowed(self):
        from_num, to_num = self._month_range_nums()
        return (from_num, to_num) != (self._available_months[0], self._available_months[-1])

    def _filtered_data(self):
        df = self.data
        if self.year_var.get() != "All":
            df = df[df["year"] == int(self.year_var.get())]
        from_num, to_num = self._month_range_nums()
        df = df[df["month_num"].between(from_num, to_num)]
        if self.line_var.get() != "All":
            df = df[df["Line"] == self.line_var.get()]
        if self.strain_var.get() != "All":
            df = df[df["Strain"] == self.strain_var.get()]
        if self._dob_range_narrowed():
            dob_from, dob_to = self.dob_from.get_date(), self.dob_to.get_date()
            df = df[(df["DOB"].dt.date >= dob_from) & (df["DOB"].dt.date <= dob_to)]
        return df

    def _filter_subtitle(self):
        bits = []
        if self.year_var.get() != "All":
            bits.append(f"Year {self.year_var.get()}")
        if self._month_range_narrowed():
            bits.append(f"Months {self.month_from_var.get()} to {self.month_to_var.get()}")
        if self.line_var.get() != "All":
            bits.append(f"Line {self.line_var.get()}")
        if self.strain_var.get() != "All":
            bits.append(f"Strain {self.strain_var.get()}")
        if self._dob_range_narrowed():
            bits.append(f"DOB {self.dob_from.get_date()} to {self.dob_to.get_date()}")
        return " | ".join(bits) if bits else "All data"

    # ---------------- plotting ----------------
    def _on_plot(self):
        if self.data is None:
            messagebox.showwarning("No data", "Load data first.")
            return
        df = self._filtered_data()
        if df.empty:
            messagebox.showinfo("No data", "No rows match the current filters.")
            return
        if self.mode_var.get() == "batch":
            self._plot_per_batch(df)
        else:
            self._plot_total(df)

    def _plot_total(self, df):
        grouped = (
            df.groupby(["year", "month_num"])[NUM_ANIMALS_COL]
            .sum()
            .reset_index()
            .sort_values(["year", "month_num"])
        )
        labels = [period_label(y, m) for y, m in zip(grouped["year"], grouped["month_num"])]
        subtitle = self._filter_subtitle()

        win = tk.Toplevel(self.root)
        win.title(f"Total {NUM_ANIMALS_COL} - {subtitle}")

        fig = Figure(figsize=(7.5, 4.5), dpi=100)
        ax = fig.add_subplot(111)
        ax.plot(labels, grouped[NUM_ANIMALS_COL], marker="o")
        ax.set_xlabel("Month-Year")
        ax.set_ylabel("Number of fish")
        ax.set_title(subtitle)
        # Always start the Y axis at zero (max stays auto, plus a small
        # margin so a marker/line at the very top is never clipped by the
        # plot's own edge) so a sharp decline is easy to read at a glance,
        # regardless of tank size.
        ax.set_ylim(bottom=0, top=y_axis_top_with_margin(grouped[NUM_ANIMALS_COL]))
        ax.tick_params(axis="x", rotation=45)
        for label in ax.get_xticklabels():
            label.set_ha("right")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    @staticmethod
    def _line_strain_label(sub, max_len=22):
        """A short "Strain / Line" label for a batch's subplot title, built
        from whichever of the two are actually populated for it."""
        def first_or_none(series):
            non_null = series.dropna()
            return str(non_null.iloc[0]) if len(non_null) else None

        strain_val = first_or_none(sub["Strain"]) if "Strain" in sub.columns else None
        line_val = first_or_none(sub["Line"]) if "Line" in sub.columns else None
        if line_val and len(line_val) > max_len:
            line_val = line_val[: max_len - 1] + "…"
        return " / ".join(p for p in (strain_val, line_val) if p)

    def _plot_per_batch(self, df):
        periods = sorted(set(zip(df["year"], df["month_num"])))
        period_labels = [period_label(y, m) for y, m in periods]
        period_index = {p: i for i, p in enumerate(periods)}

        batch_codes = sorted(df[BATCH_COL].dropna().unique())
        n = len(batch_codes)
        if n == 0:
            messagebox.showinfo("No data", "No batch codes match the current filters.")
            return
        if n > self.MANY_BATCHES_WARN_THRESHOLD:
            if not messagebox.askyesno(
                "Many graphs",
                f"This will create {n} small graphs (one per batch code), which may "
                "take a moment to render. Narrowing the Line/Strain/Year filters first "
                "will produce a more manageable view. Continue anyway?",
            ):
                return

        ncols = min(4, n)
        nrows = math.ceil(n / ncols)

        win = tk.Toplevel(self.root)
        win.title(f"Time series per batch code ({n} batch codes) - {self._filter_subtitle()}")
        win.geometry("1000x700")

        # Shown only while the subplots are being built, then replaced by
        # the actual (scrollable) plot area below.
        progress_frame = ttk.Frame(win)
        progress_frame.pack(side="top", fill="x", padx=10, pady=10)
        ttk.Label(progress_frame, text="Rendering plots...").pack(side="left")
        progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate",
                                    maximum=n, length=300)
        progress.pack(side="left", padx=10, fill="x", expand=True)
        progress_pct = ttk.Label(progress_frame, text="0%")
        progress_pct.pack(side="left")
        win.update_idletasks()

        fig = Figure(figsize=(ncols * 2.6, nrows * 2.1), dpi=100)
        for i, code in enumerate(batch_codes):
            ax = fig.add_subplot(nrows, ncols, i + 1)
            sub = df[df[BATCH_COL] == code]
            y = [float("nan")] * len(periods)
            for _, r in sub.iterrows():
                y[period_index[(r["year"], r["month_num"])]] = r[NUM_ANIMALS_COL]
            ax.plot(range(len(periods)), y, marker="o", markersize=3)
            label = self._line_strain_label(sub)
            title = f"{code}\n{label}" if label else str(code)
            ax.set_title(title, fontsize=7)
            ax.set_xticks(range(len(periods)))
            ax.set_xticklabels(period_labels, fontsize=6, rotation=45)
            ax.set_ylim(bottom=0, top=y_axis_top_with_margin(y))
            ax.tick_params(axis="y", labelsize=6)
            # Only the leftmost column gets a y-label: with hundreds of
            # batches this grid can end up many thousands of pixels tall, so
            # a single label shared across the whole figure (e.g. via
            # fig.supylabel) would land in the vertical middle of the image,
            # off-screen from whichever row is actually being viewed.
            if i % ncols == 0:
                ax.set_ylabel("Number of fish", fontsize=7)
            # Same reasoning for the x-label: only the bottom-most subplot
            # in each column gets one, instead of repeating it on every
            # subplot.
            if i + ncols >= n:
                ax.set_xlabel("Month-Year", fontsize=7)

            progress["value"] = i + 1
            progress_pct.config(text=f"{(i + 1) * 100 // n}%")
            win.update_idletasks()
        fig.tight_layout()

        progress_frame.destroy()

        outer = ttk.Frame(win)
        outer.pack(fill="both", expand=True)
        canvas_scroll = tk.Canvas(outer)
        vscroll = ttk.Scrollbar(outer, orient="vertical", command=canvas_scroll.yview)
        canvas_scroll.configure(yscrollcommand=vscroll.set)
        canvas_scroll.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")
        inner = ttk.Frame(canvas_scroll)
        canvas_scroll.create_window((0, 0), window=inner, anchor="nw")

        fig_canvas = FigureCanvasTkAgg(fig, master=inner)
        fig_canvas.draw()
        fig_canvas.get_tk_widget().pack()

        inner.update_idletasks()
        canvas_scroll.configure(scrollregion=canvas_scroll.bbox("all"))

        def _on_mousewheel(event):
            canvas_scroll.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas_scroll.bind("<MouseWheel>", _on_mousewheel)


def main():
    root = tk.Tk()
    BestandGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
