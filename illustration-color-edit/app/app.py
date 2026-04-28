"""
Streamlit entry point.

Thin scaffolding only — five horizontal tabs (Library, Editor, Global Map,
Batch Export, Settings) that delegate to ``src/`` modules for all logic.

Run with::

    streamlit run app/app.py

or via the Windows wrapper::

    launch_app.bat
"""

from __future__ import annotations

import sys
from pathlib import Path

# Make ``src`` importable when streamlit launches us from elsewhere.
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

import streamlit as st  # noqa: E402

from src.color_mapper import (  # noqa: E402
    ColorMapper,
    MatchKind,
    gray_value,
    suggest_from_history,
)
from src.config import configure_logging, load_config  # noqa: E402
from src.library_manager import LibraryManager  # noqa: E402
from src.mapping_store import (  # noqa: E402
    IllustrationMapping,
    MappingStore,
    VALID_STATUSES,
    merge_mappings,
)
from src.print_safety import check_mapping  # noqa: E402
from src.svg_parser import parse_svg  # noqa: E402
from src.svg_writer import apply_mapping_with_report, write_converted_svg  # noqa: E402


st.set_page_config(layout="wide", page_title="Illustration Color Edit")


# --------------------------------------------------------------------------- #
# Bootstrap config + state (once per session)
# --------------------------------------------------------------------------- #
def _bootstrap() -> None:
    if "config" in st.session_state:
        return
    cfg = load_config()
    cfg.ensure_dirs()
    configure_logging(cfg.log_level)

    cfg_path = cfg.source_path or (_PROJECT_ROOT / "config.json")
    st.session_state.config = cfg
    st.session_state.store = MappingStore(cfg_path, cfg.paths.metadata_dir)
    st.session_state.library = LibraryManager(cfg.paths.input_dir, st.session_state.store)
    st.session_state.current_file = None
    st.session_state.editor_picks = {}        # source_hex -> manual target_hex (current illustration)
    st.session_state.batch_report = None      # last batch run's summary


_bootstrap()


# --------------------------------------------------------------------------- #
# Caching wrappers (only cache pickleable returns)
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner=False)
def cached_color_extract(path_str: str, mtime: float) -> dict[str, int]:
    """Return ``{hex: usage_count}`` for an SVG, keyed by path + mtime."""
    parsed = parse_svg(Path(path_str))
    return {h: u.count for h, u in parsed.colors.items()}


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def render_inline_svg(svg_bytes: bytes, *, height: int = 480) -> None:
    """Render raw SVG bytes inline. Strips the XML decl so HTML doesn't choke."""
    text = svg_bytes.decode("utf-8", errors="replace")
    if text.lstrip().startswith("<?xml"):
        text = text.split("?>", 1)[1].lstrip()
    wrapper = (
        f'<div style="background:#fff;border:1px solid #e0e0e0;border-radius:6px;'
        f'padding:8px;height:{height}px;overflow:auto;display:flex;'
        f'align-items:center;justify-content:center;">{text}</div>'
    )
    st.markdown(wrapper, unsafe_allow_html=True)


def status_badge(status: str) -> str:
    color = {
        "pending": "#9CA3AF",
        "in_progress": "#F59E0B",
        "reviewed": "#10B981",
        "exported": "#3B82F6",
    }.get(status, "#9CA3AF")
    return (
        f'<span style="background:{color};color:#fff;padding:2px 8px;'
        f'border-radius:10px;font-size:0.8em;">{status}</span>'
    )


def color_swatch(hex_color: str, size: int = 22) -> str:
    return (
        f'<span style="display:inline-block;width:{size}px;height:{size}px;'
        f'background:{hex_color};border:1px solid #aaa;border-radius:3px;'
        f'vertical-align:middle;"></span>'
    )


def fresh_mapper() -> ColorMapper:
    """Build a ColorMapper from current config + live global map."""
    cfg = st.session_state.config
    store: MappingStore = st.session_state.store
    return ColorMapper(global_map=store.load_global_map(), matching=cfg.matching)


# --------------------------------------------------------------------------- #
# Tab: Library
# --------------------------------------------------------------------------- #
def tab_library() -> None:
    st.subheader("Library")
    library: LibraryManager = st.session_state.library
    cfg = st.session_state.config

    cols = st.columns([3, 1, 1, 1])
    cols[0].markdown(f"**Input directory:** `{cfg.paths.input_dir}`")
    if cols[1].button("Rescan", key="lib_rescan", width="content"):
        st.cache_data.clear()
        st.rerun()
    if cols[2].button("Open next pending", key="lib_open_next", width="content"):
        nxt = library.next_pending()
        if nxt:
            st.session_state.current_file = nxt.filename
            st.session_state.editor_picks = {}
            st.success(f"Opened {nxt.filename}. Switch to the **Editor** tab.")
        else:
            st.info("No pending illustrations.")

    entries = library.scan()
    counts = library.status_counts()
    cols[3].markdown(
        f"<div style='line-height:1.6'>"
        f"{status_badge('pending')} {counts['pending']} &nbsp;"
        f"{status_badge('in_progress')} {counts['in_progress']} &nbsp;"
        f"{status_badge('reviewed')} {counts['reviewed']} &nbsp;"
        f"{status_badge('exported')} {counts['exported']}"
        f"</div>",
        unsafe_allow_html=True,
    )

    if not entries:
        st.warning(f"No SVG files in {cfg.paths.input_dir}.")
        return

    # Header row
    h = st.columns([3, 1, 1, 1, 1, 1])
    for i, label in enumerate(["File", "Status", "Overrides", "Size (KB)", "Modified", "Open"]):
        h[i].markdown(f"**{label}**")

    for e in entries:
        c = st.columns([3, 1, 1, 1, 1, 1])
        c[0].write(e.filename)
        c[1].markdown(status_badge(e.status), unsafe_allow_html=True)
        c[2].write(e.override_count)
        c[3].write(f"{e.size_kb:.1f}")
        c[4].write(e.modified_iso[:19].replace("T", " ") if e.modified_iso else "—")
        if c[5].button("Open", key=f"open_{e.filename}", width="content"):
            st.session_state.current_file = e.filename
            st.session_state.editor_picks = {}
            st.toast(f"Opened {e.filename}. Switch to the **Editor** tab.")


# --------------------------------------------------------------------------- #
# Tab: Editor
# --------------------------------------------------------------------------- #
def tab_editor() -> None:
    st.subheader("Editor")
    library: LibraryManager = st.session_state.library
    store: MappingStore = st.session_state.store
    cfg = st.session_state.config

    current = st.session_state.get("current_file")
    if not current:
        st.info("No illustration selected. Pick one in the **Library** tab.")
        return

    svg_path = cfg.paths.input_dir / current
    if not svg_path.is_file():
        st.error(f"{svg_path} no longer exists. Rescan in the Library tab.")
        return

    # Mark in_progress on first open if currently pending.
    illu = store.load_illustration(current)
    if illu.status == "pending":
        illu.with_status("in_progress")
        store.save_illustration(illu)

    st.markdown(
        f"**File:** `{current}` &nbsp; **Status:** {status_badge(illu.status)}",
        unsafe_allow_html=True,
    )

    # Extract colors
    mtime = svg_path.stat().st_mtime
    colors = cached_color_extract(str(svg_path), mtime)
    if not colors:
        st.warning("No concrete colors found in this SVG (may be all `none`/`url(...)` references).")
        return

    mapper = fresh_mapper().with_overrides(illu.overrides)
    history = store.history()

    # Read existing picks: priority is editor_picks (in-progress edits)
    # > saved overrides > suggestion.
    picks: dict[str, str] = dict(st.session_state.editor_picks)

    left, right = st.columns(2)
    with left:
        st.markdown("**Original**")
        render_inline_svg(svg_path.read_bytes(), height=480)

    # Build the live mapping from picks+overrides+global before rendering right pane
    suggestions = {h: mapper.suggest(h) for h in sorted(colors)}
    effective: dict[str, str] = {}
    for src, sug in suggestions.items():
        if src in picks:
            effective[src] = picks[src]
        elif src in illu.overrides:
            effective[src] = illu.overrides[src]
        elif sug.target is not None:
            effective[src] = sug.target

    full_mapping = merge_mappings(store.load_global_map(), effective)
    converted_bytes, report = apply_mapping_with_report(svg_path, full_mapping)

    with right:
        st.markdown("**Converted (live preview)**")
        render_inline_svg(converted_bytes, height=480)

    st.divider()

    # Color mapping panel
    st.markdown(f"### Color mapping — {len(colors)} unique source colors")
    safety_warnings = check_mapping(effective, cfg.print_safety)
    safety_targets = {w.target for w in safety_warnings}

    sorted_colors = sorted(colors.items(), key=lambda kv: -kv[1])  # most-used first
    for src_hex, count in sorted_colors:
        sug = suggestions[src_hex]
        history_picks = suggest_from_history(src_hex, history)
        with st.container(border=True):
            row = st.columns([1, 2, 2, 3, 2])
            row[0].markdown(
                f"{color_swatch(src_hex)} <code>{src_hex}</code><br>"
                f"<small>{count} uses</small>",
                unsafe_allow_html=True,
            )

            # Suggestion summary
            if sug.kind is MatchKind.EXACT:
                badge = "<span style='color:#10B981'>● exact</span>"
                detail = sug.label or ""
            elif sug.kind is MatchKind.NEAR:
                badge = "<span style='color:#F59E0B'>● near</span>"
                detail = (
                    f"via <code>{sug.via}</code> · "
                    f"Δ{cfg.matching.metric.upper()}={sug.distance:.2f}"
                )
            else:
                badge = "<span style='color:#EF4444'>● none</span>"
                detail = "no exact or near match"
            row[1].markdown(f"{badge}<br><small>{detail}</small>", unsafe_allow_html=True)

            # Current picked target (initialized from suggestion / override)
            initial = (
                picks.get(src_hex)
                or illu.overrides.get(src_hex)
                or (sug.target if sug.target else "#888888")
            )
            picked = row[2].color_picker(
                "target",
                value=initial,
                key=f"pick_{current}_{src_hex}",
                label_visibility="collapsed",
            ).upper()

            # History suggestions (other illustrations)
            if history_picks:
                opts = [f"{t} ({c}x)" for t, c in history_picks[:5]]
                chosen = row[3].selectbox(
                    "history",
                    options=["(keep current)"] + opts,
                    key=f"hist_{current}_{src_hex}",
                    label_visibility="collapsed",
                )
                if chosen != "(keep current)":
                    chosen_hex = chosen.split(" ", 1)[0].upper()
                    picked = chosen_hex
            else:
                row[3].markdown("<small>no history yet</small>", unsafe_allow_html=True)

            # Print-safety hint for this row
            if picked in safety_targets:
                row[4].warning("⚠ light for print")
            elif gray_value(picked) <= 16:
                row[4].caption("very dark — OK")
            else:
                row[4].caption(f"luminance {gray_value(picked)}")

            picks[src_hex] = picked

    st.session_state.editor_picks = picks

    # Action bar
    st.divider()
    a1, a2, a3, a4 = st.columns([1, 1, 1, 3])
    if a1.button("Save (keep status)", key="ed_save", width="content"):
        illu.overrides = {k.upper(): v.upper() for k, v in picks.items()}
        store.save_illustration(illu)
        st.success(f"Saved {len(illu.overrides)} overrides for {current}.")
    if a2.button("Save & mark reviewed", key="ed_review", width="content", type="primary"):
        illu.overrides = {k.upper(): v.upper() for k, v in picks.items()}
        illu.with_status("reviewed")
        store.save_illustration(illu)
        # Promote new exact picks into the global map (only those that weren't there).
        gm = store.load_global_map()
        new_global = 0
        for src, tgt in illu.overrides.items():
            if src not in gm:
                store.upsert_global_entry(src, tgt, label="auto-promoted from editor", notes="")
                new_global += 1
        st.success(
            f"Saved & marked reviewed. {new_global} new entries promoted to the global map."
        )
    if a3.button("Promote ALL picks to global", key="ed_promote", width="content"):
        for src, tgt in picks.items():
            store.upsert_global_entry(src, tgt, label="manual promote", notes="")
        st.success(f"Promoted {len(picks)} entries to the global map.")

    if report.unmapped:
        a4.warning(
            f"{len(report.unmapped)} source colors are still unmapped: "
            + ", ".join(sorted(report.unmapped)[:8])
            + ("…" if len(report.unmapped) > 8 else "")
        )


# --------------------------------------------------------------------------- #
# Tab: Global Map
# --------------------------------------------------------------------------- #
def tab_global_map() -> None:
    st.subheader("Global color map")
    store: MappingStore = st.session_state.store

    gm = store.load_global_map()
    usage = store.usage_counts()

    if not gm:
        st.info("Global map is empty. Map a few colors in the Editor first.")
    else:
        h = st.columns([1, 2, 1, 2, 3, 1])
        for i, label in enumerate(["Source", "Target", "Used in", "Label", "Notes", ""]):
            h[i].markdown(f"**{label}**")

        for src in sorted(gm):
            entry = gm[src]
            c = st.columns([1, 2, 1, 2, 3, 1])
            c[0].markdown(
                f"{color_swatch(src)} <code>{src}</code>",
                unsafe_allow_html=True,
            )
            new_target = c[1].color_picker(
                "target", value=entry["target"], key=f"gm_t_{src}",
                label_visibility="collapsed",
            ).upper()
            c[2].write(usage.get(src, 0))
            new_label = c[3].text_input(
                "label", value=entry.get("label", ""), key=f"gm_l_{src}",
                label_visibility="collapsed",
            )
            new_notes = c[4].text_input(
                "notes", value=entry.get("notes", ""), key=f"gm_n_{src}",
                label_visibility="collapsed",
            )
            if c[5].button("✕", key=f"gm_del_{src}", help="Remove entry"):
                store.remove_global_entry(src)
                st.rerun()

            # Persist on change.
            if (
                new_target != entry["target"]
                or new_label != entry.get("label", "")
                or new_notes != entry.get("notes", "")
            ):
                store.upsert_global_entry(src, new_target, label=new_label, notes=new_notes)

    st.divider()
    st.markdown("**Add a new entry**")
    with st.form("add_global", clear_on_submit=True):
        f = st.columns([1, 1, 2, 3, 1])
        nsrc = f[0].text_input("source hex", value="#")
        ntgt = f[1].color_picker("target", value="#888888")
        nlbl = f[2].text_input("label")
        nnts = f[3].text_input("notes")
        if f[4].form_submit_button("Add"):
            if not nsrc.startswith("#") or len(nsrc) != 7:
                st.error("Source must be #RRGGBB.")
            else:
                store.upsert_global_entry(nsrc.upper(), ntgt.upper(), label=nlbl, notes=nnts)
                st.success(f"Added {nsrc.upper()} → {ntgt.upper()}.")


# --------------------------------------------------------------------------- #
# Tab: Batch Export
# --------------------------------------------------------------------------- #
def tab_batch() -> None:
    st.subheader("Batch export")
    library: LibraryManager = st.session_state.library
    store: MappingStore = st.session_state.store
    cfg = st.session_state.config

    only_reviewed = st.checkbox("Only export reviewed illustrations", value=True, key="batch_reviewed")
    st.markdown(f"**Output directory:** `{cfg.paths.output_dir}`")

    entries = library.scan()
    if only_reviewed:
        entries = [e for e in entries if e.status == "reviewed"]

    st.write(f"{len(entries)} illustration(s) queued.")

    if st.button("Run batch export", type="primary", key="batch_run", width="content"):
        if not entries:
            st.warning("Nothing to export.")
        else:
            cfg.paths.output_dir.mkdir(parents=True, exist_ok=True)
            global_map = store.load_global_map()
            log_rows: list[dict] = []
            progress = st.progress(0.0)
            for i, e in enumerate(entries, start=1):
                illu = store.load_illustration(e.filename)
                merged = merge_mappings(global_map, illu.overrides)
                dst = cfg.paths.output_dir / e.filename
                report = write_converted_svg(e.path, merged, dst)
                # Mark exported
                illu.with_status("exported")
                store.save_illustration(illu)
                log_rows.append({
                    "file": e.filename,
                    "replacements": report.replacements,
                    "unmapped_colors": len(report.unmapped),
                    "unmapped_list": ", ".join(sorted(report.unmapped)[:8]),
                })
                progress.progress(i / len(entries))
            st.session_state.batch_report = log_rows
            st.success(f"Exported {len(entries)} files to {cfg.paths.output_dir}.")

    rows = st.session_state.get("batch_report")
    if rows:
        st.markdown("### Last run report")
        st.dataframe(rows, width="stretch")


# --------------------------------------------------------------------------- #
# Tab: Settings
# --------------------------------------------------------------------------- #
def tab_settings() -> None:
    st.subheader("Settings")
    cfg = st.session_state.config
    st.caption(f"Config file: `{cfg.source_path}`")

    st.markdown("### Paths")
    st.write({
        "input_dir": str(cfg.paths.input_dir),
        "output_dir": str(cfg.paths.output_dir),
        "metadata_dir": str(cfg.paths.metadata_dir),
    })
    st.caption("Edit `config.json` directly to change paths, then restart the app.")

    st.markdown("### Matching")
    st.write({
        "nearest_enabled": cfg.matching.nearest_enabled,
        "metric": cfg.matching.metric,
        "threshold": cfg.matching.threshold,
    })

    st.markdown("### Print safety")
    st.write({
        "min_gray_value": cfg.print_safety.min_gray_value,
        "warn_only": cfg.print_safety.warn_only,
    })

    st.markdown("### Logging")
    st.write({"level": cfg.log_level})


# --------------------------------------------------------------------------- #
# Top-level layout
# --------------------------------------------------------------------------- #
st.title("Illustration Color Edit")
st.caption("SVG → grayscale conversion pipeline for the book project.")

t_lib, t_edit, t_global, t_batch, t_settings = st.tabs(
    ["Library", "Editor", "Global Map", "Batch Export", "Settings"]
)
with t_lib:
    tab_library()
with t_edit:
    tab_editor()
with t_global:
    tab_global_map()
with t_batch:
    tab_batch()
with t_settings:
    tab_settings()
