"""
zf_ufz_statistics.py (cloud deployment copy)
===============================================

Shared data-loading and plotting-support logic for the monthly
"Bestandserfassung" (stock/inventory) Excel exports, used by
zf_ufz_statistics_cloud.py. This is a trimmed copy of the same-named file
that also backs the internal LAN/desktop deployments - it holds only the
pure functions (filename parsing, local-folder and Nextcloud/WebDAV
loading, plot-axis helpers), with the desktop Tkinter GUI class removed
entirely, since Streamlit Community Cloud's container has no Tk system
libraries at all (see the comment further down, by the imports, for the
concrete error this caused before being trimmed).

Each source Excel file covers one month and is named like
"BestandAugust2026.xls" - the month and year are only encoded in the
*file name*, not inside the sheet itself, so they're parsed out of it
here and added as columns to every row, before all months are stacked
into a single table.

Notes on the source files
--------------------------
- They are old-format .xls files (not .xlsx), which pandas reads via the
  `xlrd` package.
- Each file has exactly one sheet, but the sheet's *name* includes the
  export date and therefore differs between files (e.g.
  "Animals02.09.2026") - so the sheet is always selected by position
  (the first one), never by name.
- xlrd sometimes prints a harmless "OLE2 inconsistency" warning for these
  particular files; the data still reads correctly despite it.
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

# Note: this is a trimmed copy for the cloud deployment - the desktop
# Tkinter GUI (BestandGUI) that the original zf_ufz_statistics.py also
# contains has been removed here, along with its tkinter/tkcalendar/
# matplotlib-Tkinter-backend imports. Streamlit Community Cloud's
# container has no Tk system libraries at all (it's headless), so
# importing tkinter there fails outright with "ImportError: libtk8.6.so:
# cannot open shared object file" - and none of the Streamlit apps
# (web/nc/cloud) ever used the GUI class anyway, only the data-loading
# and plotting-support functions below.

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


