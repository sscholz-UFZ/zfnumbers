"""
zf_ufz_statistics_cloud.py
=============================

Streamlit Community Cloud deployment copy of zf_ufz_statistics_nc.py -
identical interface and analysis, fetching the monthly "Bestandserfassung"
Excel exports over WebDAV from a Nextcloud *public share link*. It is a
deliberate separate copy (not the same file the internal LAN/PC setup
runs), so that this repository - which will live on GitHub - never
depends on, or gets confused with, anything on the office PC.

How the pieces fit together
-----------------------------
- GitHub hosts this code (this file + zf_ufz_statistics.py + the logo).
- Streamlit Community Cloud (share.streamlit.io) runs it, for free, and
  gives it a public https://...streamlit.app URL.
- Nextcloud is only ever the *data* source - the running app fetches the
  Excel files from there over the internet (WebDAV), exactly as it does
  when run locally. The code itself never lives in Nextcloud.

One-time setup, start to finish
----------------------------------
1. Create a free GitHub account (if you don't have one): github.com
2. Create a new repository and push this whole folder's contents to it
   (this file, zf_ufz_statistics.py, ufz_logo.png, requirements.txt) -
   see README.md in this same folder for the exact git commands.
3. Create a free Streamlit Community Cloud account at share.streamlit.io,
   signing in with that same GitHub account (no separate password).
4. Click "New app", pick the repository/branch, and set
   "Main file path" to: zf_ufz_statistics_cloud.py
5. Before or after deploying, open the app's Settings > Secrets in the
   Streamlit Cloud dashboard and paste in (see secrets.toml.example):

       [nextcloud]
       share_password = "the-real-share-password"

   This is Streamlit Cloud's own equivalent of a local
   .streamlit/secrets.toml file - it is never stored in the GitHub repo.
   Skipping this step is also fine: the page will just ask each visitor
   to type the password in themselves instead (see zf_ufz_statistics_nc.py
   for the full explanation of that trade-off).
6. Deploy. Share the resulting streamlit.app URL with whoever needs
   access - the Nextcloud share password (typed in or pre-set above) is
   what actually gates who can load any data.

Updating it later
-------------------
Push new commits to the same GitHub repo/branch - Streamlit Community
Cloud automatically redeploys the app from the latest commit.

Security note
--------------
This makes the app reachable to anyone with the URL - Streamlit itself
has no login screen. The Nextcloud share password (see
zf_ufz_statistics_nc.py's docstring for the "Remember"/"Forget" details)
is the actual access boundary here, so keep it only with the people who
should have access, same as for the internal deployment.
"""

import math
import json
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components
from matplotlib.figure import Figure

from zf_ufz_statistics import (
    BATCH_COL,
    MONTH_NUM_TO_NAME,
    NUM_ANIMALS_COL,
    load_bestand_from_public_share,
    period_label,
    y_axis_top_with_margin,
)

MANY_BATCHES_WARN_THRESHOLD = 80
LOGO_PATH = Path(__file__).parent / "ufz_logo.png"
# Just a link, not a secret - safe to keep directly in the code. Only the
# share's password (entered below) is treated as sensitive.
NEXTCLOUD_SHARE_URL = "https://nc.ufz.de/s/zebrafishufz"

st.set_page_config(page_title="Zebrafish Stock (Bestand) Statistics", layout="wide")
if LOGO_PATH.exists():
    st.logo(str(LOGO_PATH), size="large")
    # st.logo's built-in sizes top out at "large" (2rem tall) - double that
    # via CSS since there's no larger built-in option. Target the .stLogo
    # class (not a data-testid - "stHeaderLogo"/"stSidebarLogo" vary by
    # placement, but the class name is stable either way). !important is
    # needed because Streamlit's own generated class has equal CSS
    # specificity and is injected after ours, so it would otherwise win the
    # tie. The header bar itself is also enlarged to match (its default
    # 3.75rem is shorter than the doubled logo), so the logo isn't clipped.
    st.markdown(
        "<style>"
        ".stLogo { height: 4rem !important; }"
        "[data-testid='stHeader'] { height: 4.5rem !important; min-height: 4.5rem !important; }"
        "</style>",
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner="Fetching Excel files from Nextcloud...")
def _load_data(share_url, share_password):
    return load_bestand_from_public_share(share_url, share_password, verbose=False)


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


def _dob_label(sub):
    """"DOB: YYYY-MM-DD" for a batch, or None if it has no DOB recorded."""
    if "DOB" not in sub.columns:
        return None
    non_null = sub["DOB"].dropna()
    return f"DOB: {non_null.iloc[0]:%Y-%m-%d}" if len(non_null) else None


title_col, logo_col = st.columns([5, 1], vertical_alignment="center")
with title_col:
    st.title("Zebrafish Stock (Bestand) Statistics")
    st.caption("Monitor of the numbers of zebrafish based on monthly snapshots of the AniBio database.")
with logo_col:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH))

# ---------------------------------------------------------------- 1. Load --
st.header("1. Load Data")
st.caption("Data is fetched from a Nextcloud share over WebDAV.")

try:
    # st.secrets.get(...) is not safe here: when secrets.toml doesn't exist
    # at all (not just missing this one key), it raises rather than
    # returning the default - so the lookup itself has to be guarded.
    preset_password = st.secrets["nextcloud"]["share_password"]
except Exception:
    preset_password = None

if preset_password is not None:
    # An administrator has pre-configured the password (e.g. for an
    # always-on deployment) - nothing to ask the visitor for.
    share_password = preset_password
else:
    # --- "Remember this password on this device" -----------------------
    # Streamlit never renders a real HTML <form>, so the browser's own
    # native "save password?" prompt won't reliably trigger here - there
    # is no genuine form-submit event for it to key off. Instead, this
    # stores the password directly in this browser's own localStorage
    # (opt-in, via the checkbox below) and restores it on later visits.
    #
    # A components.html() block runs inside a same-origin sandboxed
    # iframe: it can read/write this page's real localStorage and reach
    # into the parent page's DOM (confirmed - the sandbox blocks
    # navigating the parent page, but not this), but it CANNOT redirect
    # or reload the page. So restoring a saved password is done by
    # programmatically setting the password field's own value and
    # simulating the same commit (focus, set, input event, blur) a real
    # user typing it and clicking away would produce - which is what
    # actually gets Streamlit to pick up the new value on its next rerun.
    _NC_REMEMBER_KEY = "zf_nc_share_password"
    _RESTORE_ATTEMPTED = "nc_restore_attempted"
    _REMEMBER_PREV = "nc_remember_prev"
    _PENDING_FORGET = "nc_pending_forget"

    # Handles a "Forget" button click from the *previous* run. This has to
    # happen before the password field/checkbox below are created: once a
    # widget with a given key exists in a run, Streamlit refuses to let
    # that run's own code overwrite its session_state value directly (a
    # StreamlitAPIException) - so the button itself only records that a
    # forget was requested and triggers a rerun, and the actual clearing
    # happens here, right at the top of the fresh run that follows, before
    # either widget is instantiated.
    if st.session_state.pop(_PENDING_FORGET, False):
        st.session_state["nc_share_password_field"] = ""
        st.session_state["nc_remember_checkbox"] = False
        st.session_state[_REMEMBER_PREV] = False
        components.html(
            f"<script>window.parent.localStorage.removeItem('{_NC_REMEMBER_KEY}');</script>",
            height=0,
        )

    share_password = st.text_input(
        "Nextcloud share password", type="password", key="nc_share_password_field",
        help="The password for the shared Bestandserfassung folder. Leave "
             "blank if the share has no password, then press Enter.",
    )

    is_first_run = not st.session_state.get(_RESTORE_ATTEMPTED, False)
    if is_first_run:
        # First run of this browser session: try a one-time restore. The
        # save/clear decision below is deliberately skipped this run -
        # share_password here still reflects the page *before* any
        # restore has landed, so no save/clear decision would be
        # meaningful yet (it shows up on the rerun this triggers, if the
        # restore actually found something saved).
        st.session_state[_RESTORE_ATTEMPTED] = True
        st.session_state[_REMEMBER_PREV] = False
        components.html(
            f"""
            <script>
            const saved = window.parent.localStorage.getItem('{_NC_REMEMBER_KEY}');
            if (saved) {{
                const input = window.parent.document.querySelector('input[type="password"]');
                if (input) {{
                    const setter = Object.getOwnPropertyDescriptor(
                        window.parent.HTMLInputElement.prototype, 'value').set;
                    input.focus();
                    setter.call(input, saved);
                    input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    input.blur();
                }}
            }}
            </script>
            """,
            height=0,
        )

    # One single call site for this widget, with a stable, explicit key -
    # not two separate st.checkbox(...) calls (one for the first run, one
    # for every run after). Two call sites with even slightly different
    # arguments (e.g. one carrying a help= tooltip the other lacks) hash
    # to two *different* widgets as far as Streamlit is concerned, which
    # silently discarded this checkbox's remembered state on some runs -
    # exactly the "forgets after one visit" symptom this was built to fix.
    # No value= is passed here either: once the user has actually clicked
    # it, Streamlit tracks its live state on its own from then on, and
    # re-passing a value on every rerun (even just echoing our own
    # prev_remember back) would fight that - on the very rerun a click
    # triggers, Streamlit still honors an explicitly-passed value=, which
    # would silently snap the box back right after the user changed it.
    remember = st.checkbox(
        "Remember this password on this device", key="nc_remember_checkbox",
        help="Stores the password in this browser's local storage so this "
             "page stops asking for it here. Anyone else using this same "
             "browser profile would then also be able to load the data "
             "without knowing the password.",
    )

    if not is_first_run:
        prev_remember = st.session_state.get(_REMEMBER_PREV, False)
        if remember and share_password:
            password_js = json.dumps(share_password).replace("</", "<\\/")
            components.html(
                f"<script>window.parent.localStorage.setItem"
                f"('{_NC_REMEMBER_KEY}', {password_js});</script>",
                height=0,
            )
        elif prev_remember and not remember:
            components.html(
                f"<script>window.parent.localStorage.removeItem('{_NC_REMEMBER_KEY}');</script>",
                height=0,
            )
        st.session_state[_REMEMBER_PREV] = remember

    # A dedicated, always-available action for a shared/public PC: rather
    # than relying on remembering to uncheck the box above before walking
    # away, this immediately wipes the saved password (and the password
    # field itself) in one click, regardless of the checkbox's state.
    if st.button("Forget saved password on this device"):
        st.session_state[_PENDING_FORGET] = True
        st.rerun()

if st.button("Reload Data (pick up newly added files)"):
    _load_data.clear()

try:
    data = _load_data(NEXTCLOUD_SHARE_URL, share_password)
except Exception as e:
    # The detailed error (which may include the Nextcloud URL) goes to
    # this process's own console/log only - the web page never shows it.
    print(f"[zf_ufz_statistics_nc] Failed to load data: {e}")
    st.error("Could not load data. Please check the share password and try again.")
    st.stop()

n_files = data["source_file"].nunique()
st.success(f"Loaded {len(data)} rows from {n_files} file(s).")

# ------------------------------------------------------------- 2. Filters --
st.header("2. Filters")
st.caption(
    "Snapshot time range refers to which monthly AniBio snapshot file(s) the data came from "
    "(e.g. “BestandAugust2026.xls”) - not a date recorded for an individual animal. "
    "These filters apply to both the data table below and the analysis further down."
)

available_periods = sorted(set(zip(data["year"], data["month_num"])))
period_option_labels = [period_label(y, m) for y, m in available_periods]
lines = sorted(data["Line"].dropna().unique()) if "Line" in data.columns else []
strains = sorted(data["Strain"].dropna().unique()) if "Strain" in data.columns else []
dob_min = data["DOB"].min().date()
dob_max = data["DOB"].max().date()

row1_col1, row1_col2 = st.columns(2)
with row1_col1:
    if len(period_option_labels) > 1:
        time_range = st.select_slider(
            "Snapshot time range", options=period_option_labels,
            value=(period_option_labels[0], period_option_labels[-1]),
            help="The range of monthly AniBio snapshots to include (spans Year and Month "
                 "together, e.g. 05-2026 to 07-2027).",
        )
    else:
        time_range = (period_option_labels[0], period_option_labels[0])
        st.text_input("Snapshot time range", value=period_option_labels[0], disabled=True,
                       help="Only one snapshot is currently loaded.")
with row1_col2:
    # No min_value/max_value here on purpose: the picker defaults to the data's
    # actual span, but a user should be free to pick any range at all (e.g.
    # "past 3 months" from today, well past the last recorded DOB) - a range
    # with no matching rows is a perfectly valid, informative result, not
    # something to block at the widget level.
    dob_range = st.date_input(
        "Date of birth (DOB) range", value=(dob_min, dob_max),
        help="Each animal's own date of birth - independent of which monthly "
             "snapshot file it appears in.",
    )

row2_col1, row2_col2 = st.columns(2)
with row2_col1:
    line_choice = st.selectbox("Line", ["All"] + [str(v) for v in lines])
with row2_col2:
    strain_choice = st.selectbox("Strain", ["All"] + [str(v) for v in strains])

# While the user has only picked one end of the DOB range, Streamlit briefly
# returns a 1-tuple - fall back to the full range until both are set.
if isinstance(dob_range, tuple) and len(dob_range) == 2:
    dob_from, dob_to = dob_range
else:
    dob_from, dob_to = dob_min, dob_max

from_idx = period_option_labels.index(time_range[0])
to_idx = period_option_labels.index(time_range[1])
selected_periods = available_periods[from_idx: to_idx + 1]
selected_period_keys = {y * 12 + m for y, m in selected_periods}

period_key = data["year"] * 12 + data["month_num"]
filtered = data[period_key.isin(selected_period_keys)]
if line_choice != "All":
    filtered = filtered[filtered["Line"] == line_choice]
if strain_choice != "All":
    filtered = filtered[filtered["Strain"] == strain_choice]
dob_narrowed = (dob_from, dob_to) != (dob_min, dob_max)
if dob_narrowed:
    filtered = filtered[(filtered["DOB"].dt.date >= dob_from) & (filtered["DOB"].dt.date <= dob_to)]

subtitle_bits = []
time_narrowed = (time_range[0], time_range[1]) != (period_option_labels[0], period_option_labels[-1])
if time_narrowed:
    subtitle_bits.append(f"Snapshots {time_range[0]} to {time_range[1]}")
if line_choice != "All":
    subtitle_bits.append(f"Line {line_choice}")
if strain_choice != "All":
    subtitle_bits.append(f"Strain {strain_choice}")
if dob_narrowed:
    subtitle_bits.append(f"DOB {dob_from} to {dob_to}")
subtitle = " | ".join(subtitle_bits) if subtitle_bits else "All data"

# ------------------------------------------------------------- 3. Inspect --
st.header("3. Inspect Data")
st.caption(f"{len(filtered)} of {len(data)} row(s) shown, based on the filters above: {subtitle}")
st.dataframe(filtered, width="stretch", height=300)

# ------------------------------------------------------------- 4. Analyse --
st.header("4. Analysis")

mode = st.radio(
    "Analysis type",
    [
        "Total Num. animals per year/month",
        "Time series per batch code (many small graphs)",
        "Single batch code (with comments)",
    ],
)

if filtered.empty:
    st.info("No rows match the current filters.")
    st.stop()

resolution_choice = st.selectbox(
    "Snapshot resolution for the plot", ["Every month", "Every 2nd month", "Every year"],
    help="Thins out how many monthly snapshots are actually plotted - useful once the "
         "database spans many months. The data table above is never affected by this, "
         "and the most recent snapshot in the current filters is always kept.",
)
_resolution_step = {"Every month": 1, "Every 2nd month": 2, "Every year": 12}[resolution_choice]

_periods_in_filtered = sorted(set(zip(filtered["year"], filtered["month_num"])))
# Anchor the stride at the most recent snapshot and step backward (e.g. for
# "every 2nd month" with Aug as the latest: Aug, Jun, Apr, ...) rather than
# striding forward from the oldest snapshot, which could leave an uneven
# final gap right before the most recent one.
_periods_to_plot = list(reversed(_periods_in_filtered[::-1][::_resolution_step]))
_plot_period_keys = {y * 12 + m for y, m in _periods_to_plot}
plot_df = filtered[(filtered["year"] * 12 + filtered["month_num"]).isin(_plot_period_keys)]

if mode.startswith("Total"):
    grouped = (
        plot_df.groupby(["year", "month_num"])[NUM_ANIMALS_COL]
        .sum()
        .reset_index()
        .sort_values(["year", "month_num"])
    )
    labels = [period_label(y, m) for y, m in zip(grouped["year"], grouped["month_num"])]

    fig = Figure(figsize=(6, 3.5), dpi=100)
    ax = fig.add_subplot(111)
    ax.plot(labels, grouped[NUM_ANIMALS_COL], marker="o")
    ax.set_xlabel("Month-Year")
    ax.set_ylabel("Number of fish")
    ax.set_title(subtitle)
    # Always start the Y axis at zero (max stays auto, plus a small margin
    # so a marker/line at the very top is never clipped by the plot's own
    # edge) so a sharp decline is easy to read at a glance, regardless of
    # tank size.
    ax.set_ylim(bottom=0, top=y_axis_top_with_margin(grouped[NUM_ANIMALS_COL]))
    ax.tick_params(axis="x", rotation=45)
    for lbl in ax.get_xticklabels():
        lbl.set_ha("right")
    fig.tight_layout()
    # width="content" shows the figure at its own natural size (per figsize
    # above) instead of Streamlit's default of stretching it to fill this
    # page's full "wide" layout width, which is what made it look huge.
    st.pyplot(fig, width="content")

elif mode.startswith("Time series per batch"):
    periods = sorted(set(zip(plot_df["year"], plot_df["month_num"])))
    period_labels = [period_label(y, m) for y, m in periods]
    period_index = {p: i for i, p in enumerate(periods)}
    batch_codes = sorted(plot_df[BATCH_COL].dropna().unique())
    n = len(batch_codes)

    if n == 0:
        st.info("No batch codes match the current filters.")
        st.stop()

    if n > MANY_BATCHES_WARN_THRESHOLD:
        st.info(
            f"Rendering {n} small graphs (one per batch code), which may take a moment. "
            "Narrowing the Line/Strain/time-range filters first will produce a more "
            "manageable view."
        )

    ncols = min(4, n)
    nrows = math.ceil(n / ncols)

    progress_bar = st.progress(0, text="🐟 Rendering plots...")
    fig = Figure(figsize=(ncols * 2.6, nrows * 2.1), dpi=100)
    for i, code in enumerate(batch_codes):
        ax = fig.add_subplot(nrows, ncols, i + 1)
        sub = plot_df[plot_df[BATCH_COL] == code]
        y = [float("nan")] * len(periods)
        for _, r in sub.iterrows():
            y[period_index[(r["year"], r["month_num"])]] = r[NUM_ANIMALS_COL]
        ax.plot(range(len(periods)), y, marker="o", markersize=3)
        title_lines = [str(code)]
        label = _line_strain_label(sub)
        if label:
            title_lines.append(label)
        dob = _dob_label(sub)
        if dob:
            title_lines.append(dob)
        ax.set_title("\n".join(title_lines), fontsize=7)
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
        # Same reasoning for the x-label: only the bottom-most subplot in
        # each column gets one, instead of repeating it on every subplot.
        if i + ncols >= n:
            ax.set_xlabel("Month-Year", fontsize=7)
        progress_bar.progress((i + 1) / n, text=f"🐟 Rendering plots... {i + 1}/{n}")
    fig.tight_layout()
    progress_bar.empty()

    st.pyplot(fig, width="content")

else:
    batch_codes_all = sorted(filtered[BATCH_COL].dropna().unique())
    if not batch_codes_all:
        st.info("No batch codes match the current filters.")
        st.stop()

    # One row per batch code, but every column - same shape as the Inspect
    # Data table above, just deduplicated so multi-selecting rows here means
    # selecting distinct batch codes rather than distinct monthly snapshots.
    batch_summary = (
        filtered.drop_duplicates(subset=[BATCH_COL])
        .sort_values(BATCH_COL)
        .reset_index(drop=True)
    )

    st.caption("Click one or more rows below to select their batch codes.")
    picker_state = st.dataframe(
        batch_summary, hide_index=True, width="stretch", height=250,
        on_select="rerun", selection_mode="multi-row", key="single_batch_picker",
    )
    selected_rows = [i for i in picker_state.selection.rows if i < len(batch_summary)]
    if selected_rows:
        selected_batches = list(batch_summary.iloc[selected_rows][BATCH_COL])
    else:
        # Nothing clicked yet (or the filters changed under a stale
        # selection) - default to the first batch code in the table.
        selected_batches = [batch_summary.iloc[0][BATCH_COL]]

    for selected_batch in selected_batches:
        batch_df = filtered[filtered[BATCH_COL] == selected_batch]
        batch_plot_df = plot_df[plot_df[BATCH_COL] == selected_batch]

        line_strain = _line_strain_label(batch_df)
        st.subheader(f"{selected_batch} — {line_strain}" if line_strain else str(selected_batch))
        dob = _dob_label(batch_df)
        if dob:
            st.caption(dob)

        graph_col, comments_col = st.columns([2, 1])

        with graph_col:
            periods = sorted(set(zip(batch_plot_df["year"], batch_plot_df["month_num"])))
            labels = [period_label(y, m) for y, m in periods]
            period_index = {p: i for i, p in enumerate(periods)}
            y = [float("nan")] * len(periods)
            for _, r in batch_plot_df.iterrows():
                y[period_index[(r["year"], r["month_num"])]] = r[NUM_ANIMALS_COL]

            title_lines = [str(selected_batch)]
            if line_strain:
                title_lines.append(line_strain)
            if dob:
                title_lines.append(dob)

            fig = Figure(figsize=(6, 3.5), dpi=100)
            ax = fig.add_subplot(111)
            ax.plot(labels, y, marker="o")
            ax.set_xlabel("Month-Year")
            ax.set_ylabel("Number of fish")
            ax.set_title("\n".join(title_lines))
            ax.set_ylim(bottom=0, top=y_axis_top_with_margin(y))
            ax.tick_params(axis="x", rotation=45)
            for lbl in ax.get_xticklabels():
                lbl.set_ha("right")
            fig.tight_layout()
            st.pyplot(fig, width="content")

        with comments_col:
            st.markdown("**Comments**")
            # Comments are shown for every snapshot in the current filters
            # (Line/Strain/time-range/DOB), independent of the plot's own
            # resolution setting - thinning out the graph shouldn't hide
            # text notes that exist on the in-between snapshots.
            comment_rows = batch_df[["year", "month_num", "Comments"]].dropna(subset=["Comments"])
            comment_rows = comment_rows[comment_rows["Comments"].astype(str).str.strip() != ""]
            comment_rows = comment_rows.sort_values(["year", "month_num"])
            if comment_rows.empty:
                st.caption("No comments available for this batch.")
            else:
                # The same comment is often just carried over unchanged into
                # every later snapshot - only show it again once its text
                # actually changes from the previous snapshot's comment.
                prev_comment = None
                for _, r in comment_rows.iterrows():
                    text = str(r["Comments"])
                    if text == prev_comment:
                        continue
                    st.markdown(f"**{period_label(r['year'], r['month_num'])}**  \n{text}")
                    prev_comment = text

        if len(selected_batches) > 1:
            st.divider()
