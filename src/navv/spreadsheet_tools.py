#!/usr/bin/env python3

# Copyright 2023 Battelle Energy Alliance, LLC

import os
import itertools
from collections import Counter
import socket
from copy import copy
import json
from typing import Any, Optional

from importlib.resources import files
import pickle
import string

import openpyxl
import openpyxl.styles
from openpyxl.styles import Alignment, NamedStyle
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.datavalidation import DataValidation
import netaddr
from tqdm import tqdm

from navv import data_types
from navv.utilities import timeit
from navv.message_handler import warning_msg
from navv import validators

DATA_PKL_FILE = str(files("navv").joinpath("data", "data.pkl"))
COL_NAMES = [
    "Count",
    "Src_IP",
    "Src_Desc",
    "Dest_IP",
    "Dest_Desc",
    "Port",
    "Service",
    "Proto",
    "Conn_State",
    "Src_Pur_Lev",
    "Src_Pur_Desc",
    "Dest_Pur_Lev",
    "Dest_Pur_Desc",
    "Notes",
]
HEADER_STYLE = NamedStyle(
    name="header_style",
    font=openpyxl.styles.Font(name="Calibri", size=11, bold=True),
    fill=openpyxl.styles.PatternFill("solid", fgColor="4286F4"),
)
IPV6_CELL_COLOR = (
    openpyxl.styles.PatternFill("solid", fgColor="FFFFFF"),
    openpyxl.styles.Font(name="Calibri", size=11, color="ff0000"),
)
EXTERNAL_NETWORK_CELL_COLOR = (
    openpyxl.styles.PatternFill("solid", fgColor="030303"),
    openpyxl.styles.Font(name="Calibri", size=11, color="ffff00"),
)
INTERNAL_NETWORK_CELL_COLOR = (
    openpyxl.styles.PatternFill("solid", fgColor="ffff00"),
    openpyxl.styles.Font(name="Calibri", size=11, color="000000"),
)
ICMP_CELL_COLOR = (
    openpyxl.styles.PatternFill("solid", fgColor="ff33cc"),
    openpyxl.styles.Font(name="Calibri", size=11, color="000000"),
)
UNKNOWN_EXTERNAL_CELL_COLOR = (
    openpyxl.styles.PatternFill("solid", fgColor="ffffff"),
    openpyxl.styles.Font(name="Calibri", size=11, color="000000"),
)

ALREADY_UNRESOLVED = list()


PURDUE_SHEET_NAME = "Purdue_Definitions"

# Purdue risk constants and styles
PURDUE_RISK_COL_NAME = "Risk"
PURDUE_RISK_LOW = "Low"
PURDUE_RISK_MED = "Medium"
PURDUE_RISK_HIGH = "High"

PURDUE_RISK_STYLE_LOW = (
    openpyxl.styles.PatternFill("solid", fgColor="FF00B050"),  # green
    openpyxl.styles.Font(name="Calibri", size=11, color="000000"),
)
PURDUE_RISK_STYLE_MED = (
    openpyxl.styles.PatternFill("solid", fgColor="FFFFA500"),  # orange
    openpyxl.styles.Font(name="Calibri", size=11, color="000000"),
)
PURDUE_RISK_STYLE_HIGH = (
    openpyxl.styles.PatternFill("solid", fgColor="FFFF0000"),  # red
    openpyxl.styles.Font(name="Calibri", size=11, color="FFFFFF"),
)

# Fixed Purdue levels (canonical list). Asset owners may customize colors/definitions, but not the level names.
PURDUE_LEVELS = [
    ("L0", "Physical process"),
    ("L1", "Basic control"),
    ("L2", "Area supervisory control"),
    ("L3", "Site operations"),
    ("L3.5", "Control System DMZ"),
    ("L4", "Enterprise IT"),
    ("L5", "Internet DMZ"),
    ("L6", "External / Cloud"),
]

# Default Purdue colors (cell fill). These are defaults and can be changed by the asset owner in Excel.
# NOTE: Colors are stored as the fill of the `Purdue_Color` cell; users can recolor cells directly.


PURDUE_DEFAULT_COLORS = {
    "L0": "3D6B35",     # green (Field I/O / process)
    "L1": "D6B85A",     # tan/yellow (Controller LAN)
    "L2": "E67E22",     # orange (Local HMI / supervisory)
    "L3": "0B7285",     # teal/blue (Site operations)
    "L3.5": "D4B24C",   # gold (Control System DMZ)
    "L4": "808080",     # gray (Enterprise IT)
    "L5": "0B3B5B",     # dark blue (Internet DMZ)
    "L6": "7030A0",     # purple (External / Cloud)
}


# Canonical worksheet order (anything not listed stays at the end)
CANONICAL_SHEET_ORDER: list[str] = [
    "Analysis",
    "Purdue Analysis",
    "Inventory Input",
    "Segments",
    "Purdue_Definitions",
    "MAC",
    "SNMP",
    "Inventory Report",
    "Stats",
    "Conn States",
    "Externals",
    "Unknown Internals",
]


def _reorder_standard_sheets(wb: openpyxl.Workbook) -> None:
    """Enforce a stable worksheet order in the workbook.

    Sheets listed in CANONICAL_SHEET_ORDER are placed first (when present) in that
    exact order. Any other sheets remain at the end in their current relative order.

    NOTE: We set `wb._sheets` directly because repeated `move_sheet(offset=...)`
    becomes unreliable once sheets are created/removed across multiple passes.
    """
    try:
        sheets = list(wb.worksheets)
        by_name = {ws.title: ws for ws in sheets}

        ordered = []
        seen = set()

        for name in CANONICAL_SHEET_ORDER:
            ws = by_name.get(name)
            if ws is not None:
                ordered.append(ws)
                seen.add(name)

        # Preserve relative order of any non-canonical sheets
        for ws in sheets:
            if ws.title not in seen:
                ordered.append(ws)

        # NOTE: `_sheets` is an internal openpyxl attribute used intentionally here to
        # enforce a stable, canonical worksheet order. This is safe at runtime but not
        # declared in openpyxl type stubs, so we silence the static type checker.
        wb._sheets = ordered  # type: ignore[attr-defined]
    except Exception:
        # Fail silently (ordering is nice-to-have)
        return


# Helper to ensure the header style is registered only once per workbook
def ensure_header_style(wb: openpyxl.Workbook) -> None:
    """Ensure the workbook has the named style used for header cells.

    OpenPyXL raises if you try to add a NamedStyle with an existing name, so we only
    register it once and then reference it by name.
    """
    existing: set[str] = set()
    for s in wb.named_styles:
        if isinstance(s, str):
            existing.add(s)
        else:
            name = getattr(s, "name", None)
            if name:
                existing.add(name)

    if HEADER_STYLE.name in existing:
        return

    try:
        wb.add_named_style(HEADER_STYLE)
    except ValueError:
        # Defensive: if another style with the same name is already present, do not fail.
        return


# --- Purdue helpers ---

def read_purdue_definitions(wb: openpyxl.Workbook):
    """Read Purdue_Definitions into a dict: level -> {fill, font, desc}, plus ordered levels."""
    if PURDUE_SHEET_NAME not in wb.sheetnames:
        return {}, []

    ws = wb[PURDUE_SHEET_NAME]
    levels_in_order: list[str] = []
    purdue_map: dict[str, dict] = {}

    for row in itertools.islice(ws.iter_rows(values_only=False), 1, None):
        lvl_cell = row[0]
        if lvl_cell is None or lvl_cell.value is None:
            continue
        lvl = str(lvl_cell.value).strip()
        if not lvl:
            continue

        color_cell = row[1] if len(row) > 1 else None
        desc_cell = row[2] if len(row) > 2 else None

        fill = (
            copy(color_cell.fill)
            if color_cell is not None
            else openpyxl.styles.PatternFill("solid", fgColor="FFFFFF")
        )
        font = openpyxl.styles.Font(name="Calibri", size=11, color="000000")
        desc = (
            str(desc_cell.value).strip()
            if (desc_cell is not None and desc_cell.value is not None)
            else ""
        )

        levels_in_order.append(lvl)
        purdue_map[lvl] = {"fill": fill, "font": font, "desc": desc}

    cleaned = validators.warn_purdue_definition_issues(levels_in_order)
    purdue_map = {lvl: purdue_map[lvl] for lvl in cleaned if lvl in purdue_map}

    return purdue_map, cleaned


def apply_purdue_dropdown_to_segments(wb: openpyxl.Workbook, seg_ws) -> None:
    """Apply a Purdue level dropdown to the Segments sheet Purdue_Level column."""
    if PURDUE_SHEET_NAME not in wb.sheetnames:
        return

    pur_ws = wb[PURDUE_SHEET_NAME]

    last_row = 1
    for r in range(2, pur_ws.max_row + 1):
        v = pur_ws.cell(row=r, column=1).value
        if v is not None and str(v).strip():
            last_row = r

    if last_row < 2:
        return

    src_range = f"'{PURDUE_SHEET_NAME}'!$A$2:$A${last_row}"
    dv = DataValidation(type="list", formula1=src_range, allow_blank=True)
    dv.error = "Select a Purdue level from the list."
    dv.errorTitle = "Invalid Purdue level"
    dv.prompt = "Choose a Purdue level (source: Purdue_Definitions)."
    dv.promptTitle = "Purdue level"

    seg_ws.add_data_validation(dv)
    dv.add("D2:D1048576")


def get_workbook(file_name):
    """Create the blank Inventory and Segment sheets for data input into the tool"""
    if os.path.isfile(file_name):
        wb = openpyxl.load_workbook(file_name)
        ensure_header_style(wb)
        # Ensure Segments sheet has Purdue_Level column + dropdown (non-destructive).
        if "Segments" in wb.sheetnames:
            seg_ws = wb["Segments"]
            if seg_ws.cell(row=1, column=4).value != "Purdue_Level":
                seg_ws.cell(row=1, column=4, value="Purdue_Level").style = HEADER_STYLE.name
            apply_purdue_dropdown_to_segments(wb, seg_ws)
            _reorder_standard_sheets(wb)
    else:
        wb = openpyxl.Workbook()
        ensure_header_style(wb)
        inv_sheet = wb.active
        assert inv_sheet is not None
        inv_sheet.title = "Inventory Input"
        seg_sheet = wb.create_sheet("Segments")

        inv_sheet.cell(row=1, column=1, value="IP").style = HEADER_STYLE.name
        inv_sheet.cell(row=1, column=2, value="Name").style = HEADER_STYLE.name

        seg_sheet.cell(row=1, column=1, value="Name").style = HEADER_STYLE.name
        seg_sheet.cell(row=1, column=2, value="Description").style = HEADER_STYLE.name
        seg_sheet.cell(row=1, column=3, value="CIDR").style = HEADER_STYLE.name
        seg_sheet.cell(row=1, column=4, value="Purdue_Level").style = HEADER_STYLE.name

        # Create Purdue definitions with defaults on initial workbook creation (Pass 1).
        write_purdue_definitions_sheet(wb, overwrite_with_defaults=True)
        apply_purdue_dropdown_to_segments(wb, seg_sheet)
        _reorder_standard_sheets(wb)
    return wb


@timeit
def get_inventory_data(ws, **kwargs):
    inventory = dict()
    for row in itertools.islice(ws.iter_rows(), 1, None):
        if not row[0].value or not row[1].value:
            continue
        inventory[row[0].value] = data_types.InventoryItem(
            ip=row[0].value,
            name=row[1].value,
            color=(copy(row[0].fill), copy(row[0].font)),
            mac_address="",
            vendor=""
        )
    return inventory


@timeit
def get_segments_data(ws, **kwargs):
    segments = []
    valid_purdue_levels: set[str] = set(kwargs.get("valid_purdue_levels") or [])
    network_ip = ""
    for row in itertools.islice(ws.iter_rows(), 1, None):
        if not row[2].value:
            continue
        network_ip = str(row[2].value).strip()
        pur_level = ""
        if len(row) > 3 and row[3].value is not None:
            pur_level = str(row[3].value).strip()
        if valid_purdue_levels:
            validators.warn_invalid_segment_purdue(str(row[0].value or ""), pur_level, valid_purdue_levels)
        try:
            network_ips = [str(ip) for ip in netaddr.IPNetwork(network_ip)]
        except (netaddr.AddrFormatError, ValueError):
            warning_msg(
                f"Invalid segment CIDR '{network_ip}' in Segments sheet — expected format like 10.10.10.0/24. Skipping."
            )
            continue

        for ip in network_ips:
            segments.append(
                data_types.Segment(
                    name=row[0].value,
                    description=row[1].value,
                    cidr=network_ip,
                    network=ip,
                    color=[copy(row[0].fill), copy(row[0].font)],
                    purdue_level=pur_level,
                )
            )
    return segments


def write_purdue_definitions_sheet(wb, *, overwrite_with_defaults: bool) -> None:
    """Create/seed the Purdue_Definitions sheet.

    Pass 1: overwrite_with_defaults=True  -> always rewrite defaults.
    Pass 2: overwrite_with_defaults=False -> preserve user edits; create only if missing.

    Purdue color is stored as the fill color of the `Purdue_Color` cell.
    Defaults are written top-down (L6 at the top through L0 at the bottom) for readability.
    """
    ensure_header_style(wb)
    if PURDUE_SHEET_NAME in wb.sheetnames and overwrite_with_defaults:
        wb.remove(wb[PURDUE_SHEET_NAME])

    if PURDUE_SHEET_NAME in wb.sheetnames and not overwrite_with_defaults:
        # Preserve user customizations.
        _reorder_standard_sheets(wb)
        return

    # Create the sheet and position it immediately after Segments (so it appears near the front).
    ws = wb.create_sheet(PURDUE_SHEET_NAME)
    _reorder_standard_sheets(wb)

    # Headers
    ws.cell(row=1, column=1, value="Purdue_Level").style = HEADER_STYLE.name
    ws.cell(row=1, column=2, value="Purdue_Color").style = HEADER_STYLE.name
    ws.cell(row=1, column=3, value="Purdue_Description").style = HEADER_STYLE.name

    # Defaults (display highest Purdue level at the top of the sheet)
    for r, (level, desc) in enumerate(reversed(PURDUE_LEVELS), start=2):
        ws.cell(row=r, column=1, value=level)

        color_cell = ws.cell(row=r, column=2, value="")
        fill_rgb = PURDUE_DEFAULT_COLORS.get(level, "FFFFFF")
        color_cell.fill = openpyxl.styles.PatternFill("solid", fgColor=fill_rgb)

        desc_cell = ws.cell(row=r, column=3, value=desc)
        desc_cell.alignment = openpyxl.styles.Alignment(wrap_text=True)

    ws.freeze_panes = "A2"
    auto_adjust_width(ws, 60)


# ----------------------------
# Segment guessing helpers
# ----------------------------


def _segments_sheet_existing_cidrs(seg_ws) -> set[str]:
    """Return a set of existing CIDR strings (normalized) from the Segments sheet."""
    existing: set[str] = set()
    for row in itertools.islice(seg_ws.iter_rows(values_only=True), 1, None):
        cidr = row[2] if len(row) > 2 else None
        if cidr is None:
            continue
        s = str(cidr).strip()
        if not s:
            continue
        existing.add(s)
    return existing


def guess_segments_from_ips(
    ips: set[str],
    *,
    min_ips_24: int = 3,
    enable_16: bool = True,
    min_ips_16: int = 256,
    min_24s_in_16: int = 8,
    enable_8: bool = False,
) -> list[tuple[str, str]]:
    """Return a list of (cidr, description) guesses from observed private IPv4 addresses.

    Bias:
      - /24 is the default segmentation unit (99% of real deployments).
      - /16 is proposed only when evidence is strong.
      - /8 is disabled by default.

    Returned list contains /24 candidates plus optional /16 roll-up candidates.
    """
    # Keep only RFC1918 IPv4 addresses
    v4_priv: set[netaddr.IPAddress] = set()
    for ip in ips:
        try:
            a = netaddr.IPAddress(str(ip).strip())
        except Exception:
            continue
        if a.version != 4:
            continue
        if not a.is_ipv4_private_use():
            continue
        v4_priv.add(a)

    if not v4_priv:
        return []

    # /24 candidates
    by_24: dict[str, set[str]] = {}
    for a in v4_priv:
        n24 = str(netaddr.IPNetwork(f"{a}/24").cidr)
        by_24.setdefault(n24, set()).add(str(a))

    guesses: list[tuple[str, str]] = []

    for cidr, members in sorted(by_24.items(), key=lambda kv: (-len(kv[1]), kv[0])):
        if len(members) < int(min_ips_24):
            continue
        desc = f"Auto-guessed from PCAP (observed {len(members)} unique IPs)"
        guesses.append((cidr, desc))

    # Optional /16 roll-ups (non-destructive proposals)
    if enable_16:
        # Map /16 -> (unique IPs, populated /24s)
        by_16_ips: dict[str, set[str]] = {}
        by_16_24s: dict[str, set[str]] = {}
        for a in v4_priv:
            n16 = str(netaddr.IPNetwork(f"{a}/16").cidr)
            n24 = str(netaddr.IPNetwork(f"{a}/24").cidr)
            by_16_ips.setdefault(n16, set()).add(str(a))
            by_16_24s.setdefault(n16, set()).add(n24)

        for n16 in sorted(by_16_ips.keys()):
            ip_count = len(by_16_ips[n16])
            distinct_24 = len(by_16_24s.get(n16, set()))
            if ip_count >= int(min_ips_16) and distinct_24 >= int(min_24s_in_16):
                desc = (
                    f"Roll-up candidate (covers {distinct_24} populated /24s, {ip_count} unique IPs). "
                    "Review before using."
                )
                guesses.append((n16, desc))

    # Extremely rare: /8 proposals (off by default)
    if enable_8:
        by_8_ips: dict[str, set[str]] = {}
        by_8_16s: dict[str, set[str]] = {}
        for a in v4_priv:
            n8 = str(netaddr.IPNetwork(f"{a}/8").cidr)
            n16 = str(netaddr.IPNetwork(f"{a}/16").cidr)
            by_8_ips.setdefault(n8, set()).add(str(a))
            by_8_16s.setdefault(n8, set()).add(n16)

        for n8 in sorted(by_8_ips.keys()):
            # Only propose /8 if it clearly spans many /16s (otherwise it's noise)
            if len(by_8_16s.get(n8, set())) >= 8 and len(by_8_ips[n8]) >= 2048:
                desc = (
                    f"Rare /8 roll-up candidate (covers {len(by_8_16s[n8])} /16s, {len(by_8_ips[n8])} IPs). "
                    "Review very carefully."
                )
                guesses.append((n8, desc))

    return guesses


def write_guessed_segments(
    wb: openpyxl.Workbook,
    ips: set[str],
    *,
    merge: bool = False,
    min_ips_24: int = 3,
    enable_16: bool = True,
    min_ips_16: int = 256,
    min_24s_in_16: int = 8,
    enable_8: bool = False,
) -> int:
    """Write guessed segments into the Segments tab.

    - If merge=False, the Segments sheet is cleared (rows 2..end) and replaced.
    - If merge=True, guessed segments are appended only if their CIDR does not already exist.

    Returns number of segments written.
    """
    if "Segments" not in wb.sheetnames:
        return 0

    seg_ws = wb["Segments"]

    # Ensure header exists (non-destructive)
    ensure_header_style(wb)
    if seg_ws.cell(row=1, column=1).value != "Name":
        seg_ws.cell(row=1, column=1, value="Name").style = HEADER_STYLE.name
    if seg_ws.cell(row=1, column=2).value != "Description":
        seg_ws.cell(row=1, column=2, value="Description").style = HEADER_STYLE.name
    if seg_ws.cell(row=1, column=3).value != "CIDR":
        seg_ws.cell(row=1, column=3, value="CIDR").style = HEADER_STYLE.name
    if seg_ws.cell(row=1, column=4).value != "Purdue_Level":
        seg_ws.cell(row=1, column=4, value="Purdue_Level").style = HEADER_STYLE.name

    if not merge:
        # Clear rows 2..end
        if seg_ws.max_row and seg_ws.max_row > 1:
            seg_ws.delete_rows(2, seg_ws.max_row - 1)

    existing_cidrs = _segments_sheet_existing_cidrs(seg_ws) if merge else set()

    guesses = guess_segments_from_ips(
        ips,
        min_ips_24=min_ips_24,
        enable_16=enable_16,
        min_ips_16=min_ips_16,
        min_24s_in_16=min_24s_in_16,
        enable_8=enable_8,
    )

    if not guesses:
        return 0

    written = 0
    next_row = seg_ws.max_row + 1

    for cidr, desc in guesses:
        if merge and cidr in existing_cidrs:
            continue

        # Create a deterministic AUTO name
        safe = cidr.replace("/", "_").replace(".", ".")
        suffix = ""
        if cidr.endswith("/16"):
            suffix = "_ROLLUP"
        if cidr.endswith("/8"):
            suffix = "_ROLLUP"

        name = f"AUTO_{safe}{suffix}"

        name_cell = seg_ws.cell(row=next_row, column=1, value=name)
        name_cell.fill = INTERNAL_NETWORK_CELL_COLOR[0]
        name_cell.font = INTERNAL_NETWORK_CELL_COLOR[1]

        dcell = seg_ws.cell(row=next_row, column=2, value=desc)
        dcell.alignment = openpyxl.styles.Alignment(wrap_text=True)
        dcell.fill = INTERNAL_NETWORK_CELL_COLOR[0]
        dcell.font = INTERNAL_NETWORK_CELL_COLOR[1]

        cidr_cell = seg_ws.cell(row=next_row, column=3, value=cidr)
        cidr_cell.fill = INTERNAL_NETWORK_CELL_COLOR[0]
        cidr_cell.font = INTERNAL_NETWORK_CELL_COLOR[1]

        seg_ws.cell(row=next_row, column=4, value="")  # Purdue_Level left blank intentionally

        next_row += 1
        written += 1

    # Re-apply dropdown after writing
    apply_purdue_dropdown_to_segments(wb, seg_ws)

    return written


def get_package_data():
    """Load services and conn_states data into memory"""
    with open(DATA_PKL_FILE, "rb") as f:
        services, conn_states = pickle.load(f)
    return services, conn_states


@timeit
def create_analysis_array(sort_input, **kwargs):
    arr = []
    # sort by count and source IP
    counted = sorted(
        list(
            str(count) + "\t" + item
            for item, count in sorted(Counter(sort_input).items(), key=lambda x: x[0])
        ),
        key=lambda x: int(x.split("\t")[0]),
        reverse=True,
    )
    for row in counted:
        cells = row.split("\t")
        arr.append(
            data_types.AnalysisRowItem(
                count=cells[0],
                src_ip=cells[1],
                dest_ip=cells[2],
                port=cells[3],
                proto=cells[4],
                conn=cells[5],
            )
        )

    return arr


@timeit
def perform_analysis(
    wb,
    rows,
    services,
    conn_states,
    inventory,
    segments,
    dns_data,
    json_path,
    ext_IPs,
    unk_int_IPs,
    **kwargs,
):
    # Purdue Definitions sheet handling:
    # - Pass 1 should overwrite with defaults.
    # - Pass 2 should preserve user edits.
    overwrite_purdue_defs = bool(kwargs.pop("overwrite_purdue_defs", False))
    write_purdue_definitions_sheet(wb, overwrite_with_defaults=overwrite_purdue_defs)

    sheet = make_sheet(wb, "Analysis", idx=0)
    sheet.append(
        [
            "Count",
            "Src_IP",
            "Src_Desc",
            "Dest_IP",
            "Dest_Desc",
            "Port",
            "Service",
            "Proto",
            "Conn_State",
            "Src_Pur_Lev",
            "Src_Pur_Desc",
            "Dest_Pur_Lev",
            "Dest_Pur_Desc",
            "Notes",
        ]
    )
    warning_msg("this may take awhile...")
    purdue_map, _purdue_levels = read_purdue_definitions(wb)
    matched_segment_names = set()
    for row_index, row in enumerate(tqdm(rows), start=2):
        row.src_desc = handle_ip(
            row.src_ip,
            dns_data,
            inventory,
            segments,
            ext_IPs,
            unk_int_IPs,
            matched_segment_names,
            purdue_map,
        )
        row.dest_desc = handle_ip(
            row.dest_ip,
            dns_data,
            inventory,
            segments,
            ext_IPs,
            unk_int_IPs,
            matched_segment_names,
            purdue_map,
        )
        handle_service(row, services)
        row.conn = (row.conn, conn_states[row.conn])
        write_row_to_sheet(row, row_index, sheet)
    # Optional but smart: warn on segments with zero matches in Analysis.
    # This helps catch CIDR typos (e.g., 10.10.x vs 10.100.x) or stale definitions.
    defined_segment_names = {s.name for s in segments if getattr(s, "name", None)}
    zero_match = sorted(defined_segment_names - matched_segment_names)
    if zero_match:
        warning_msg(
            "Segments defined in the 'Segments' sheet but not observed in Analysis (0 matches): "
            + ", ".join(zero_match)
            + ". If you expected traffic, double-check CIDR/IP notation in 'Segments'."
        )
    tab = Table(displayName="AnalysisTable", ref=f"A1:N{len(rows) + 1}")
    sheet.add_table(tab)

    # Segment-focused Purdue flow view (CIDR preferred; singleton IPs only when not in any segment)
    write_purdue_analysis_sheet(wb, rows, inventory, segments, dns_data, purdue_map)

    # Objective, qualitative risk roll-up per Purdue level (High/Medium/Low + color)

    # Ensure standard sheet ordering after all writes
    _reorder_standard_sheets(wb)

    # write lookup data to json file for future use
    with open(json_path, "w+") as fp:
        json.dump(dns_data, fp)


def _purdue_level_rank() -> dict[str, int]:
    """Return a stable ordering for Purdue levels used for adjacency/skip logic."""
    order = ["L0", "L1", "L2", "L3", "L3.5", "L4", "L5", "L6"]
    return {lvl: i for i, lvl in enumerate(order)}


def _flow_objective_risk(src_level: str, dst_level: str) -> str | None:
    """Compute objective qualitative risk for a single Purdue flow.

    Rules (directional):
      - High: any communication between L6 and OT levels (<= L3.5), either direction.
      - Medium: any higher->lower, any distance; OR upward skips (>1 level).
      - Low: same level; OR adjacent upward (one step only).

    Returns one of: High/Medium/Low, or None if levels are missing/unknown.
    """
    src = (src_level or "").strip()
    dst = (dst_level or "").strip()

    # Ignore unknown/blank levels for risk scoring
    if not src or not dst:
        return None
    if src.lower() == "unknown" or dst.lower() == "unknown":
        return None

    rank = _purdue_level_rank()
    if src not in rank or dst not in rank:
        return None

    s = rank[src]
    d = rank[dst]

    # High: OT (<= L3.5) to/from L6
    ot_max = rank["L3.5"]
    if (src == "L6" and d <= ot_max) or (dst == "L6" and s <= ot_max):
        return PURDUE_RISK_HIGH

    diff = d - s

    # Same level
    if diff == 0:
        return PURDUE_RISK_LOW

    # Upward (to higher level)
    if diff > 0:
        # Adjacent up one level is Low; any skip is Medium
        return PURDUE_RISK_LOW if diff == 1 else PURDUE_RISK_MED

    # Downward (to lower level) is always Medium
    return PURDUE_RISK_MED


def _worse_risk(a: str | None, b: str | None) -> str | None:
    """Return the worse of two risk strings, using High > Medium > Low."""
    order = {PURDUE_RISK_LOW: 1, PURDUE_RISK_MED: 2, PURDUE_RISK_HIGH: 3}
    if a is None:
        return b
    if b is None:
        return a
    return a if order.get(a, 0) >= order.get(b, 0) else b


def compute_purdue_level_risks_from_sheet(wb: openpyxl.Workbook) -> dict[str, str]:
    """Compute worst-case risk per Purdue level from the Purdue Analysis sheet.

    This function intentionally normalizes Excel cell values to strings at the boundary.
    OpenPyXL cell values can be None, numbers, formulas, etc. We treat any non-empty
    value as text for Purdue level comparison.
    """
    if "Purdue Analysis" not in wb.sheetnames:
        return {}

    ws = wb["Purdue Analysis"]

    # Expected header columns:
    # Count, Src_CIDR_or_IP, Src_Desc, Src_Purdue_Level, Src_Purdue_Desc,
    # Dst_CIDR_or_IP, Dst_Desc, Dst_Purdue_Level, Dst_Purdue_Desc

    def _cell_text(v: Any) -> str:
        """Return a trimmed string for an Excel cell value (or "" for None/blank)."""
        if v is None:
            return ""
        # Normalize common non-string values
        s = str(v).strip()
        return s

    risks: dict[str, Optional[str]] = {}

    for row in itertools.islice(ws.iter_rows(values_only=True), 1, None):
        if not row or len(row) < 9:
            continue

        src_lvl = _cell_text(row[3])
        dst_lvl = _cell_text(row[7])

        r = _flow_objective_risk(src_lvl, dst_lvl)
        if r is None:
            continue

        # Apply worst-case to both endpoints (objective per-level risk)
        risks[src_lvl] = _worse_risk(risks.get(src_lvl), r)
        risks[dst_lvl] = _worse_risk(risks.get(dst_lvl), r)

    # Normalize: drop Nones
    return {k: v for k, v in risks.items() if v is not None}


def apply_purdue_risk_to_definitions(wb: openpyxl.Workbook) -> None:
    """Write qualitative risk (High/Medium/Low + color) into Purdue_Definitions."""
    if PURDUE_SHEET_NAME not in wb.sheetnames:
        return

    ws = wb[PURDUE_SHEET_NAME]

    # Ensure header exists
    if ws.cell(row=1, column=4).value != PURDUE_RISK_COL_NAME:
        ws.cell(row=1, column=4, value=PURDUE_RISK_COL_NAME).style = HEADER_STYLE.name

    per_level = compute_purdue_level_risks_from_sheet(wb)

    # Default to Low if a level exists but has no flows? No: keep blank unless observed.
    for r in range(2, ws.max_row + 1):
        lvl = ws.cell(row=r, column=1).value
        if lvl is None:
            continue
        lvl_s = str(lvl).strip()
        if not lvl_s:
            continue

        risk = per_level.get(lvl_s)
        cell = ws.cell(row=r, column=4, value=risk or "")

        if risk == PURDUE_RISK_HIGH:
            cell.fill, cell.font = PURDUE_RISK_STYLE_HIGH
        elif risk == PURDUE_RISK_MED:
            cell.fill, cell.font = PURDUE_RISK_STYLE_MED
        elif risk == PURDUE_RISK_LOW:
            cell.fill, cell.font = PURDUE_RISK_STYLE_LOW
        else:
            # Clear styling if no risk
            cell.fill = openpyxl.styles.PatternFill()
            cell.font = openpyxl.styles.Font(name="Calibri", size=11, color="000000")


# --- Purdue Analysis Sheet ---
def write_purdue_analysis_sheet(
    wb: openpyxl.Workbook,
    rows: list,
    inventory: dict,
    segments: list[data_types.Segment],
    dns_data: dict,
    purdue_map: dict,
) -> None:
    """Create a segment-focused Purdue Analysis tab.

    - Prefer segments (display CIDR) when an IP is a member of a segment.
    - Use singleton IPs only when not part of any segment.
    - External/public IPs remain black and default to Purdue L6 (in Purdue columns only).
    - Segment/IP identity columns use segment color (or yellow internal / black external for singletons).
    - Purdue columns use Purdue colors only when a Purdue level is defined (or forced L6 for external singletons).

    The sheet is written as an Excel Table so the user can sort/filter on any column.
    """

    # Fast lookup: IP -> Segment (segments list is expanded per-IP)
    seg_by_ip: dict[str, data_types.Segment] = {}
    for s in segments:
        if getattr(s, "network", None):
            seg_by_ip[s.network] = s

    default_pur_style = (
        openpyxl.styles.PatternFill("solid", fgColor="FFFFFF"),
        openpyxl.styles.Font(name="Calibri", size=11, color="000000"),
    )

    def _is_private_v4(ip: str) -> bool:
        try:
            a = netaddr.IPAddress(ip)
        except Exception:
            return False
        return bool(a.version == 4 and a.is_ipv4_private_use())

    def _endpoint_for_ip(ip: str):
        """Return (cidr_or_ip, desc, seg_style(fill,font), pur_level, pur_desc, pur_style(fill,font))."""
        seg = seg_by_ip.get(ip)
        if seg is not None:
            # Segment endpoint: show CIDR and segment description
            cidr_or_ip = getattr(seg, "cidr", "") or ip
            desc = getattr(seg, "description", "") or ""
            seg_style = (seg.color[0], seg.color[1])

            pur_level = (getattr(seg, "purdue_level", "") or "").strip()
            pur_desc = ""
            pur_style = default_pur_style
            if pur_level and pur_level in purdue_map:
                pur_desc = purdue_map[pur_level].get("desc", "")
                pur_style = (
                    purdue_map[pur_level].get("fill"),
                    purdue_map[pur_level].get("font"),
                )
            else:
                # Internal segment with no Purdue level defined: use yellow "Unknown" so filters/sorts work
                pur_level = "Unknown"
                pur_desc = "Unknown"
                pur_style = INTERNAL_NETWORK_CELL_COLOR

            return cidr_or_ip, desc, seg_style, pur_level, pur_desc, pur_style

        # Singleton endpoint
        cidr_or_ip = ip
        # Prefer inventory name, then DNS, else Unknown
        if ip in inventory:
            desc = inventory[ip].name
        elif ip in dns_data:
            desc = dns_data[ip]
        else:
            desc = "Unknown"

        # Determine internal vs external for styling
        if _is_private_v4(ip):
            # Internal singleton (not in any segment): keep identity cells yellow and
            # also populate Purdue columns with "Unknown" so filters/sorts remain useful
            # even if the user hides the segment/IP columns.
            seg_style = INTERNAL_NETWORK_CELL_COLOR
            pur_level = "Unknown"
            pur_desc = "Unknown"
            pur_style = INTERNAL_NETWORK_CELL_COLOR
        else:
            # External/public: always black; force Purdue L6 in Purdue columns
            seg_style = EXTERNAL_NETWORK_CELL_COLOR
            pur_level = "L6"
            pur_desc = ""
            pur_style = default_pur_style
            if pur_level in purdue_map:
                pur_desc = purdue_map[pur_level].get("desc", "")
                pur_style = (
                    purdue_map[pur_level].get("fill"),
                    purdue_map[pur_level].get("font"),
                )

        return cidr_or_ip, desc, seg_style, pur_level, pur_desc, pur_style

    # Build unique flow pairs (segment CIDR preferred; singleton IP only if not in any segment)
    flows: dict[tuple[str, str], tuple] = {}
    flow_counts: dict[tuple[str, str], int] = {}
    for r in rows:
        src_key = getattr(r, "src_ip", "")
        dst_key = getattr(r, "dest_ip", "")
        if not src_key or not dst_key:
            continue

        src = _endpoint_for_ip(src_key)
        dst = _endpoint_for_ip(dst_key)

        pair_key = (src[0], dst[0])  # (Src_CIDR_or_IP, Dst_CIDR_or_IP)
        # Accumulate observed counts (rows are already aggregated in Analysis)
        try:
            c = int(getattr(r, "count", 0))
        except Exception:
            c = 0
        flow_counts[pair_key] = flow_counts.get(pair_key, 0) + c

        # Store endpoint metadata once (directional)
        if pair_key not in flows:
            flows[pair_key] = (src, dst)

    # Write sheet
    ws = make_sheet(wb, "Purdue Analysis", idx=1)
    ws.append(
        [
            "Count",
            "Src_CIDR_or_IP",
            "Src_Desc",
            "Src_Purdue_Level",
            "Src_Purdue_Desc",
            "Dst_CIDR_or_IP",
            "Dst_Desc",
            "Dst_Purdue_Level",
            "Dst_Purdue_Desc",
            "Risk",
        ]
    )

    # Apply header style
    for c in range(1, 11):
        ws.cell(row=1, column=c).style = HEADER_STYLE.name

    # Readable ordering: highest count first, then src/dst identifiers
    ordered = sorted(
        [(flow_counts.get(k, 0), flows[k][0], flows[k][1]) for k in flows.keys()],
        key=lambda t: (-int(t[0] or 0), str(t[1][0]), str(t[2][0])),
    )

    row_idx = 2
    for flow_count, src, dst in ordered:
        # src and dst are tuples from _endpoint_for_ip
        src_id, src_desc, src_style, src_pl, src_pd, src_pur_style = src
        dst_id, dst_desc, dst_style, dst_pl, dst_pd, dst_pur_style = dst

        # Count
        ws.cell(row=row_idx, column=1, value=int(flow_count) if flow_count is not None else 0)

        # Identity columns (segment/ip) use segment color or default yellow/black
        cB = ws.cell(row=row_idx, column=2, value=src_id)
        cB.fill, cB.font = src_style
        cC = ws.cell(row=row_idx, column=3, value=src_desc)
        cC.fill, cC.font = src_style

        cF = ws.cell(row=row_idx, column=6, value=dst_id)
        cF.fill, cF.font = dst_style
        cG = ws.cell(row=row_idx, column=7, value=dst_desc)
        cG.fill, cG.font = dst_style

        # Purdue columns use Purdue colors only if level exists (or forced L6 for external)
        cD = ws.cell(row=row_idx, column=4, value=src_pl)
        cE = ws.cell(row=row_idx, column=5, value=src_pd)
        if src_pl:
            cD.fill, cD.font = src_pur_style
            cE.fill, cE.font = src_pur_style

        cH = ws.cell(row=row_idx, column=8, value=dst_pl)
        cI = ws.cell(row=row_idx, column=9, value=dst_pd)
        if dst_pl:
            cH.fill, cH.font = dst_pur_style
            cI.fill, cI.font = dst_pur_style

        # Risk column is computed from the Purdue level flow for this row
        risk = _flow_objective_risk(str(src_pl or ""), str(dst_pl or ""))
        if not risk:
            # No blanks per design: explicitly mark unknown so it remains sortable/filterable
            risk = "Unknown"
            risk_cell = ws.cell(row=row_idx, column=10, value=risk)
            risk_cell.fill, risk_cell.font = INTERNAL_NETWORK_CELL_COLOR
        else:
            risk_cell = ws.cell(row=row_idx, column=10, value=risk)
            if risk == PURDUE_RISK_HIGH:
                risk_cell.fill, risk_cell.font = PURDUE_RISK_STYLE_HIGH
            elif risk == PURDUE_RISK_MED:
                risk_cell.fill, risk_cell.font = PURDUE_RISK_STYLE_MED
            elif risk == PURDUE_RISK_LOW:
                risk_cell.fill, risk_cell.font = PURDUE_RISK_STYLE_LOW
            else:
                # Fallback
                risk_cell.fill, risk_cell.font = INTERNAL_NETWORK_CELL_COLOR

        row_idx += 1

    # Make it sortable/filterable like Analysis
    last_row = max(2, row_idx - 1)
    tab = Table(displayName="PurdueAnalysisTable", ref=f"A1:J{last_row}")
    ws.add_table(tab)

    ws.freeze_panes = "A2"
    auto_adjust_width(ws, 100)


def write_row_to_sheet(row, row_index, sheet):
    sheet.cell(row=row_index, column=1, value=int(row.count))

    src_IP = sheet.cell(row=row_index, column=2, value=row.src_ip)
    src_IP.fill = row.src_desc[1][0]
    src_IP.font = row.src_desc[1][1]

    src_Desc = sheet.cell(row=row_index, column=3, value=row.src_desc[0])
    src_Desc.fill = row.src_desc[1][0]
    src_Desc.font = row.src_desc[1][1]

    dest_IP = sheet.cell(row=row_index, column=4, value=row.dest_ip)
    dest_IP.fill = row.dest_desc[1][0]
    dest_IP.font = row.dest_desc[1][1]

    dest_Desc = sheet.cell(row=row_index, column=5, value=row.dest_desc[0])
    dest_Desc.fill = row.dest_desc[1][0]
    dest_Desc.font = row.dest_desc[1][1]

    sheet.cell(row=row_index, column=6, value=int(row.port))

    service = sheet.cell(row=row_index, column=7, value=row.service[0])
    service.fill = row.service[1][0]
    service.font = row.service[1][1]

    sheet.cell(row=row_index, column=8, value=row.proto)

    conn_State = sheet.cell(row=row_index, column=9, value=row.conn[0])
    conn_State.fill = row.conn[1][0]
    conn_State.font = row.conn[1][1]

    # Purdue columns (level + description) with Purdue color fill
    src_pur_level = row.src_desc[2]
    src_pur_desc = row.src_desc[3]
    src_pur_style = row.src_desc[4]

    dest_pur_level = row.dest_desc[2]
    dest_pur_desc = row.dest_desc[3]
    dest_pur_style = row.dest_desc[4]

    src_pl = sheet.cell(row=row_index, column=10, value=src_pur_level)
    src_pl.fill = src_pur_style[0]
    src_pl.font = src_pur_style[1]

    src_pd = sheet.cell(row=row_index, column=11, value=src_pur_desc)
    src_pd.fill = src_pur_style[0]
    src_pd.font = src_pur_style[1]

    dest_pl = sheet.cell(row=row_index, column=12, value=dest_pur_level)
    dest_pl.fill = dest_pur_style[0]
    dest_pl.font = dest_pur_style[1]

    dest_pd = sheet.cell(row=row_index, column=13, value=dest_pur_desc)
    dest_pd.fill = dest_pur_style[0]
    dest_pd.font = dest_pur_style[1]

    # placeholder for notes cell
    sheet.cell(row=row_index, column=14, value="")


def handle_service(row, services):
    # { port: { proto: (name, (fill, font)} }
    if row.port in services and row.proto in services[row.port]:
        row.service = services[row.port][row.proto]
    else:
        if row.proto == "icmp":
            if netaddr.valid_ipv4(row.src_ip):
                row.proto = "ICMPv4"
                service_dict = data_types.icmp4_types
            else:
                row.proto = "ICMPv6"
                service_dict = data_types.icmp6_types
            if row.port in service_dict:
                row.service = (service_dict[row.port], ICMP_CELL_COLOR)
            else:
                row.service = ("unknown icmp", ICMP_CELL_COLOR)
        else:
            row.service = ("unknown service", UNKNOWN_EXTERNAL_CELL_COLOR)


def handle_ip(
    ip_to_check,
    dns_data,
    inventory,
    segments,
    ext_IPs,
    unk_int_IPs,
    matched_segment_names=None,
    purdue_map=None,
):
    """Function take IP Address and uses collected dns_data, inventory, and segment information to give IP Addresses in analysis context.

    Priority flow:
        * DHCP Broadcasting
        * Multicast
        * Within Segments identified
            * Resolution by DNS, then Inventory, and then Unknown
            * Appends name if External IP
        * Private Network
            * Resolution by DNS, Inventory, then Unknown
        * External (Public IP space) or Internet
            * Resolution by DNS, Unknown

    This will capture the name description and the color coding identified within the worksheet.
    """
    segment_ips = [segment.network for segment in segments]
    default_pur_style = (
        openpyxl.styles.PatternFill("solid", fgColor="FFFFFF"),
        openpyxl.styles.Font(name="Calibri", size=11, color="000000"),
    )
    desc_to_change = ("Not Triggered IP", IPV6_CELL_COLOR, "", "", default_pur_style)
    if purdue_map is None:
        purdue_map = {}
    if ip_to_check == str("0.0.0.0"):
        desc_to_change = (
            "Unassigned IPv4",
            IPV6_CELL_COLOR,
            "",
            "",
            default_pur_style,
        )
    elif ip_to_check == str("255.255.255.255"):
        desc_to_change = (
            "IPv4 All Subnet Broadcast",
            IPV6_CELL_COLOR,
            "",
            "",
            default_pur_style,
        )
    elif (
        netaddr.valid_ipv6(ip_to_check) or netaddr.IPAddress(ip_to_check).is_multicast()
    ):
        desc_to_change = (
            f"{'IPV6' if netaddr.valid_ipv6(ip_to_check) else 'IPV4'}{'_Multicast' if netaddr.IPAddress(ip_to_check).is_multicast() else ''}",
            IPV6_CELL_COLOR,
            "",
            "",
            default_pur_style,
        )
    elif ip_to_check in segment_ips:
        # Find the matching segment for this IP (do NOT drop the last segment).
        seg = next((s for s in segments if s.network == ip_to_check), None)
        if seg is not None:
            if matched_segment_names is not None and getattr(seg, "name", None):
                matched_segment_names.add(seg.name)
            if ip_to_check in dns_data:
                resolution = dns_data[ip_to_check]
            elif ip_to_check in inventory:
                resolution = inventory[ip_to_check].name
            else:
                resolution = f"Unknown device in {seg.name} network"
                unk_int_IPs.add(ip_to_check)

            if not netaddr.IPAddress(ip_to_check).is_ipv4_private_use():
                resolution = resolution + " {Non-Priv IP}"

            pur_level = getattr(seg, "purdue_level", "") or ""
            pur_desc = ""
            pur_style = default_pur_style
            if pur_level and pur_level in purdue_map:
                pur_desc = purdue_map[pur_level].get("desc", "")
                pur_style = (purdue_map[pur_level].get("fill"), purdue_map[pur_level].get("font"))
            else:
                pur_level = "Unknown"
                pur_desc = "Unknown"
                pur_style = INTERNAL_NETWORK_CELL_COLOR

            desc_to_change = (
                resolution,
                seg.color,
                pur_level,
                pur_desc,
                pur_style,
            )
    elif netaddr.IPAddress(ip_to_check).is_ipv4_private_use():
        if ip_to_check in dns_data:
            # Internal singleton, resolved by DNS: use yellow "Unknown" Purdue fields so filters/sorts work
            desc_to_change = (
                dns_data[ip_to_check],
                INTERNAL_NETWORK_CELL_COLOR,
                "Unknown",
                "Unknown",
                INTERNAL_NETWORK_CELL_COLOR,
            )
        elif ip_to_check in inventory:
            # Internal singleton, resolved by Inventory: use yellow "Unknown" Purdue fields so filters/sorts work
            desc_to_change = (
                inventory[ip_to_check].name,
                INTERNAL_NETWORK_CELL_COLOR,
                "Unknown",
                "Unknown",
                INTERNAL_NETWORK_CELL_COLOR,
            )
        else:
            # Unknown internal singleton: populate Purdue columns with yellow 'Unknown' so filters/sorts work
            # even if the user hides the segment/IP identity columns in Excel.
            desc_to_change = (
                "Unknown Internal address",
                INTERNAL_NETWORK_CELL_COLOR,
                "Unknown",
                "Unknown",
                INTERNAL_NETWORK_CELL_COLOR,
            )
            unk_int_IPs.add(ip_to_check)
    else:
        ext_IPs.add(ip_to_check)
        if ip_to_check in dns_data:
            resolution = dns_data[ip_to_check]
        elif ip_to_check in inventory:
            resolution = inventory[ip_to_check].name + " {Non-Priv IP}"
        else:
            try:
                resolution = socket.gethostbyaddr(ip_to_check)[0]
            except socket.herror:
                ALREADY_UNRESOLVED.append(ip_to_check)
            finally:
                resolution = "Unresolved external address"
        # External/public traffic defaults to the highest Purdue level (L6 = External / Cloud)
        ext_pur_level = "L6"
        ext_pur_desc = ""
        ext_pur_style = default_pur_style
        if ext_pur_level in purdue_map:
            ext_pur_desc = purdue_map[ext_pur_level].get("desc", "")
            ext_pur_style = (
                purdue_map[ext_pur_level].get("fill"),
                purdue_map[ext_pur_level].get("font"),
            )

        desc_to_change = (
            resolution,
            EXTERNAL_NETWORK_CELL_COLOR,
            ext_pur_level,
            ext_pur_desc,
            ext_pur_style,
        )
    return desc_to_change


def write_conn_states_sheet(conn_states, wb):
    new_ws = make_sheet(wb, "Conn States", idx=8)
    new_ws.append(["State", "Description"])
    for index, conn_state in enumerate(conn_states, start=2):
        # State column
        state_cell = new_ws[f"A{index}"]
        state_cell.value = conn_state
        state_cell.fill = conn_states[conn_state][0]
        state_cell.font = conn_states[conn_state][1]

        # Description column
        desc_cell = new_ws[f"B{index}"]
        desc_cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
        desc_cell.value = conn_states[conn_state][2]
        desc_cell.fill = conn_states[conn_state][0]
        desc_cell.font = conn_states[conn_state][1]
    auto_adjust_width(new_ws, 100)


def write_inventory_report_sheet(inventory_df, wb):
    """Get Mac Addresses with their associated IP addresses and manufacturer."""
    ir_sheet = make_sheet(wb, "Inventory Report", idx=4)
    ir_sheet.append(["MAC", "Vendor", "Hostname", "IPv4", "IPv6", "Port and Proto"])

    inventory_data = inventory_df.to_dict(orient="records")
    index = 2
    for row in inventory_data:
        # Mac column
        mac = row.get("mac", "")
        if not mac or not str(mac).strip():
            continue
        ir_sheet[f"A{index}"].value = mac

        # Vendor column
        ir_sheet[f"B{index}"].value = row["vendor"]

        # Hostname column
        hostname_column = ir_sheet[f"C{index}"]
        hostname_column.alignment = openpyxl.styles.Alignment(wrap_text=True)

        hostname = ""
        if row["hostname"]:
            hostname = ", ".join(each for each in row["hostname"] if each)
        hostname_column.value = hostname

        # IPv4 Address column
        ipv4_column = ir_sheet[f"D{index}"]
        ipv4_column.alignment = openpyxl.styles.Alignment(wrap_text=True)

        ipv4 = ""
        if row["ipv4"]:
            ipv4 = ", ".join(each for each in row["ipv4"] if each)
        ipv4_column.value = ipv4

        # IPv6 Address column
        ipv6_column = ir_sheet[f"E{index}"]
        ipv6_column.alignment = openpyxl.styles.Alignment(wrap_text=True)

        ipv6 = ""
        if row["ipv6"]:
            ipv6 = ", ".join(each for each in row["ipv6"] if each)
        ipv6_column.value = ipv6

        # Port and Protocol column
        pnp_column = ir_sheet[f"F{index}"]
        pnp_column.alignment = openpyxl.styles.Alignment(wrap_text=True)

        port_and_proto = ""
        if row["port_and_proto"]:
            port_and_proto = ", ".join(
                list(set(each for each in row["port_and_proto"] if each))[:10]
            )

        pnp_column.value = port_and_proto

        # Add styling to every other row
        if index % 2 == 0:
            for cell in ir_sheet[f"{index}:{index}"]:
                cell.fill = openpyxl.styles.PatternFill("solid", fgColor="AAAAAA")
        index += 1
    auto_adjust_width(ir_sheet, 40)


def write_snmp_sheet(snmp_df, wb):
    """Write SNMP log data to excel sheet."""
    sheet = make_sheet(wb, "SNMP", idx=4)
    sheet.append(
        ["Src IPv4", "Src Port", "Dest IPv4", "Dest Port", "Version", "Community"]
    )

    for index, row in enumerate(snmp_df.to_dict(orient="records"), start=2):
        # Source IPv4 column
        sheet[f"A{index}"].value = row["src_ip"]

        # Source Port column
        sheet[f"B{index}"].value = row["src_port"]

        # Destination IPv4 column
        sheet[f"C{index}"].value = row["dst_ip"]

        # Destination Port column
        sheet[f"D{index}"].value = row["dst_port"]

        # Version column
        sheet[f"E{index}"].value = row["version"]

        # Community column
        sheet[f"F{index}"].value = row["community"]

        # Add styling to every other row
        if index % 2 == 0:
            for cell in sheet[f"{index}:{index}"]:
                cell.fill = openpyxl.styles.PatternFill("solid", fgColor="AAAAAA")

    auto_adjust_width(sheet, 40)


def write_externals_sheet(IPs, wb):
    ext_sheet = make_sheet(wb, "Externals", idx=5)
    ext_sheet.append(["External IP"])
    for row_index, IP in enumerate(sorted(IPs), start=2):
        cell = ext_sheet[f"A{row_index}"]
        cell.value = IP
        if row_index % 2 == 0:
            cell.fill = openpyxl.styles.PatternFill("solid", fgColor="AAAAAA")
    auto_adjust_width(ext_sheet)


def write_unknown_internals_sheet(IPs, wb):
    int_sheet = make_sheet(wb, "Unknown Internals", idx=6)
    int_sheet.append(["Unknown Internal IP"])
    for row_index, IP in enumerate(sorted(IPs), start=2):
        cell = int_sheet[f"A{row_index}"]
        cell.value = IP
        if row_index % 2 == 0:
            cell.fill = openpyxl.styles.PatternFill("solid", fgColor="AAAAAA")
    auto_adjust_width(int_sheet)


def write_stats_sheet(wb, stats):
    stats_sheet = make_sheet(wb, "Stats", idx=7)
    stats_sheet.append(
        ["Length of Capture time"]
        + [column for column in stats if column != "Length of Capture time"]
    )
    stats_sheet["A2"] = stats.pop("Length of Capture time")
    for col_index, stat in enumerate(stats, 1):
        stats_sheet[f"{string.ascii_uppercase[col_index]}2"].value = stats[stat]
    auto_adjust_width(stats_sheet)


def write_mac_sheet(mac_df, wb):
    """Fill spreadsheet with MAC address -> IP address translation with manufacturer information"""
    sheet = make_sheet(wb, "MAC", idx=4)
    sheet.append(["MAC", "Manufacturer", "IPs"])
    for index, row in enumerate(mac_df.to_dict(orient="records"), start=2):
        # Source MAC column
        sheet[f"A{index}"].value = row["mac"]

        # Source Manufacturer column
        sheet[f"B{index}"].value = row["vendor"]

        # Source IPs
        sheet[f"C{index}"].value = row["associated_ip"]
        if len(row["associated_ip"]) > 16:
            sel_cell = sheet[f"C{index}"]
            sel_cell.alignment = Alignment(wrap_text=True)
            est_row_hght = int(len(row["associated_ip"]) / 50)
            if est_row_hght < 1:
                est_row_hght = 1
            sheet.row_dimensions[index].height = est_row_hght * 15

    auto_adjust_width(sheet)
    sheet.column_dimensions["C"].width = 39 * 1.2


def make_sheet(wb, sheet_name, idx=None):
    """Create the sheet if it doesn't already exist otherwise remove it and recreate it"""
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    return wb.create_sheet(sheet_name, index=idx)


def auto_adjust_width(sheet, width=40):
    """Adjust the width of the columns to fit the data"""
    for col in sheet.columns:
        max_width = max(len(f"{c.value}") for c in col if c.value) + 2
        sheet.column_dimensions[col[0].column_letter].width = (
            width if width < max_width else max_width
        )
