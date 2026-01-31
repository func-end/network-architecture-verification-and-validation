"""CLI Commands."""
import os
import shutil
import sys
import webbrowser
from typing import Set


# Third-Party Libraries
import click

# cisagov Libraries
from navv.gui.app import app
from navv.bll import get_inventory_report_df, get_snmp_df, get_zeek_df, get_mac_df
from navv.message_handler import success_msg, warning_msg
from navv.spreadsheet_tools import (
    auto_adjust_width,
    create_analysis_array,
    get_inventory_data,
    get_package_data,
    get_segments_data,
    get_workbook,
    perform_analysis,
    read_purdue_definitions,
    write_conn_states_sheet,
    write_externals_sheet,
    write_inventory_report_sheet,
    write_snmp_sheet,
    write_stats_sheet,
    write_unknown_internals_sheet,
    write_mac_sheet,
    write_guessed_segments,
    _reorder_standard_sheets,
)
from navv.zeek import (
    get_conn_data,
    get_dns_data,
    get_snmp_data,
    run_zeek,
    perform_zeekcut,
)


@click.command(
    "generate",
    help=(
        "Create or update the NAVV Excel workbook.\n\n"
        "Workflow\n"
        "  Pass 1 (create workbook)\n"
        "    Use --new-workbook to create/overwrite the XLSX template.\n"
        "    Optionally use -p/--pcap to refresh Zeek logs.\n\n"
        "  Pass 2..N (iterate)\n"
        "    Re-run against the existing workbook to refine Segments and Purdue definitions.\n"
        "    Use -p/--pcap to refresh Zeek logs (workbook edits are preserved).\n"
        "    Omit -p to reuse existing Zeek logs.\n\n"
        "Key options\n"
        "  -o, --output-dir      Output directory containing the workbook and artifacts.\n"
        "  -z, --zeek-logs       Directory to write/read Zeek logs.\n"
        "  -p, --pcap            Refresh Zeek logs from a PCAP (workbook preserved).\n"
        "  --new-workbook        Create/overwrite a new workbook with default tabs.\n"
        "  --guess-segments      Seed Segments with best-guess /24 CIDRs from observed private IPs.\n"
        "  --guess-merge         Merge guessed segments into an existing workbook (non-destructive).\n\n"
        "Examples\n"
        "  # Pass 1: new workbook + refresh Zeek logs\n"
        "  navv generate GeekLoungeBaseline \\\n"
        "    --new-workbook \\\n"
        "    -p ~/Projects/pcaps/capture.pcap \\\n"
        "    -o ~/Projects/navv/run1 \\\n"
        "    -z ~/Projects/navv/run1/zeek_logs\n\n"
        "  # Pass 1: new workbook + seed guessed segments\n"
        "  navv generate GeekLoungeBaseline \\\n"
        "    --new-workbook --guess-segments \\\n"
        "    -p ~/Projects/pcaps/capture.pcap \\\n"
        "    -o ~/Projects/navv/run1 \\\n"
        "    -z ~/Projects/navv/run1/zeek_logs\n\n"
        "  # Pass 2: iterate (reuse workbook + existing Zeek logs)\n"
        "  navv generate GeekLoungeBaseline \\\n"
        "    -o ~/Projects/navv/run1 \\\n"
        "    -z ~/Projects/navv/run1/zeek_logs\n\n"
        "  # Pass 2: iterate but refresh Zeek logs\n"
        "  navv generate GeekLoungeBaseline \\\n"
        "    -p ~/Projects/pcaps/capture.pcap \\\n"
        "    -o ~/Projects/navv/run1 \\\n"
        "    -z ~/Projects/navv/run1/zeek_logs\n"
    ),
)
@click.option(
    "-o",
    "--output-dir",
    required=False,
    help="Directory to place resultant analysis files in. Defaults to current working directory.",
    type=str,
)
@click.option(
    "-p",
    "--pcap",
    required=False,
    help="Path to pcap file. NAVV requires zeek logs or pcap. If used, zeek will run on pcap to create new logs.",
    type=str,
)
@click.option(
    "-z",
    "--zeek-logs",
    required=False,
    help="Path to store or contain zeek log files. Defaults to current working directory.",
    type=str,
)
@click.option(
    "--new-workbook",
    is_flag=True,
    help=(
        "Create/overwrite a new Excel workbook with default template tabs (Segments, Purdue_Definitions, Inventory Input). "
        "Use this to start a new assessment."
    ),
)
@click.option(
    "--guess-segments",
    is_flag=True,
    default=False,
    help="Populate Segments with best-guess CIDRs from observed private IPs (defaults to /24; may propose /16 rollups).",
)
@click.option(
    "--guess-merge",
    is_flag=True,
    default=False,
    help="Merge guessed segments into an existing workbook (non-destructive).",
)
@click.option(
    "--guess-min-ips-24",
    type=int,
    default=3,
    show_default=True,
    help="Minimum unique IPs observed in a /24 before creating a guessed segment.",
)
@click.option(
    "--guess-enable-16/--no-guess-enable-16",
    default=True,
    show_default=True,
    help="Allow conservative /16 roll-up suggestions.",
)
@click.option(
    "--guess-min-ips-16",
    type=int,
    default=256,
    show_default=True,
    help="Minimum unique IPs observed in a /16 before proposing a /16 roll-up.",
)
@click.option(
    "--guess-min-24s-in-16",
    type=int,
    default=8,
    show_default=True,
    help="Minimum populated /24s inside a /16 before proposing roll-up.",
)
@click.option(
    "--guess-enable-8",
    is_flag=True,
    default=False,
    help="Allow /8 roll-up suggestions (OFF by default; not recommended).",
)
@click.argument("customer_name")
def generate(
    customer_name,
    output_dir,
    pcap,
    zeek_logs,
    new_workbook,
    guess_segments,
    guess_merge,
    guess_min_ips_24,
    guess_enable_16,
    guess_min_ips_16,
    guess_min_24s_in_16,
    guess_enable_8,
):
    """Generate excel sheet."""
    # Defensive: some environments may inject Python flags into argv (e.g., "-B").
    # Click will treat these as unknown options; strip them if present.
    if "-B" in sys.argv:
        sys.argv = [a for a in sys.argv if a != "-B"]
    if not output_dir:
        output_dir = os.getcwd()
    if not zeek_logs:
        zeek_logs = output_dir

    # Ensure output directories exist
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(zeek_logs, exist_ok=True)

    if guess_segments and (not new_workbook) and (not guess_merge):
        warning_msg(
            "--guess-segments was requested but no action taken: use --new-workbook to seed a new workbook, or --guess-merge to add guessed segments to an existing workbook."
        )

    file_name = os.path.join(output_dir, customer_name + "_network_analysis.xlsx")

    # Workbook creation is independent from Zeek runs.

    if new_workbook:
        if os.path.exists(file_name):
            click.confirm(f"This will overwrite the existing workbook: {file_name}. Continue?", abort=True)
            os.remove(file_name)
        overwrite_purdue_defs = True
    else:
        if not os.path.exists(file_name):
            raise click.ClickException(
                f"Workbook {file_name} does not exist. Run with --new-workbook to create a new workbook."
            )
        overwrite_purdue_defs = False

    wb = get_workbook(file_name)

    services, conn_states = get_package_data()
    timer_data = dict()

    # Load Purdue level list (used for Segments validation / dropdown expectations)
    _purdue_map, _purdue_levels = read_purdue_definitions(wb)
    valid_purdue_levels = set(_purdue_levels)

    inventory = get_inventory_data(wb["Inventory Input"])

    # Determine Zeek availability and log readiness.
    zeek_available = shutil.which("zeek") is not None
    zeekcut_available = shutil.which("zeek-cut") is not None

    # If a PCAP is provided, Zeek is required to generate logs.
    if pcap and not zeek_available:
        raise click.ClickException(
            "Zeek is not installed or not on PATH. Install Zeek (and ensure `zeek` is on PATH), "
            "or omit -p/--pcap and point --zeek-logs at an existing Zeek logs directory."
        )

    if pcap:
        try:
            run_zeek(os.path.abspath(pcap), zeek_logs, timer=timer_data)
        except Exception as exc:
            # run_zeek already prints a user-facing message; convert to a clean CLI error
            raise click.ClickException(str(exc)) from exc
    else:
        timer_data["run_zeek"] = "NOT RAN"

    # Validate that Zeek logs exist (either freshly generated from PCAP or pre-existing).
    conn_log_path = os.path.join(zeek_logs, "conn.log")
    logs_ready = os.path.exists(conn_log_path)

    if not logs_ready:
        raise click.ClickException(
            f"No Zeek logs found in: {zeek_logs}. Expected at least conn.log. "
            "Provide -p/--pcap to generate logs (requires Zeek installed), "
            "or point -z/--zeek-logs to an existing Zeek logs directory."
        )

    # Get zeek data from conn.log, dns.log and snmp.log
    zeek_data = get_conn_data(zeek_logs)
    snmp_data = get_snmp_data(zeek_logs)
    dns_filtered = get_dns_data(customer_name, output_dir, zeek_logs)

    # Get dns data for resolution
    json_path = os.path.join(output_dir, f"{customer_name}_dns_data.json")

    # Get zeek dataframes
    zeek_df = get_zeek_df(zeek_data, dns_filtered)
    snmp_df = get_snmp_df(snmp_data)

    # Get inventory report dataframe
    inventory_df = get_inventory_report_df(zeek_df)
    mac_df = get_mac_df(zeek_df)

    # Turn zeekcut data into rows for spreadsheet
    rows = create_analysis_array(zeek_data, timer=timer_data)

    observed_ips: Set[str] = set()
    for r in rows:
        src = getattr(r, "src_ip", None)
        if isinstance(src, str) and src:
            observed_ips.add(src)
        dst = getattr(r, "dest_ip", None)
        if isinstance(dst, str) and dst:
            observed_ips.add(dst)

    # Optionally populate Segments with best-guess CIDRs.
    # By default this only makes sense on a new workbook, unless --guess-merge is provided.
    if guess_segments and (new_workbook or guess_merge):
        written = write_guessed_segments(
            wb,
            observed_ips,
            merge=bool(guess_merge and not new_workbook),
            min_ips_24=int(guess_min_ips_24),
            enable_16=bool(guess_enable_16),
            min_ips_16=int(guess_min_ips_16),
            min_24s_in_16=int(guess_min_24s_in_16),
            enable_8=bool(guess_enable_8),
        )
        if written:
            warning_msg(
                f"Guessed and wrote {written} segment(s) to the 'Segments' tab. Review/edit CIDRs before relying on them."
            )

    segments = get_segments_data(wb["Segments"], valid_purdue_levels=valid_purdue_levels)

    ext_IPs: set[str] = set()
    unk_int_IPs: set[str] = set()
    perform_analysis(
        wb,
        rows,
        services,
        conn_states,
        inventory,
        segments,
        dns_filtered,
        json_path,
        ext_IPs,
        unk_int_IPs,
        overwrite_purdue_defs=overwrite_purdue_defs,
        timer=timer_data,
    )

    write_inventory_report_sheet(inventory_df, wb)

    write_externals_sheet(ext_IPs, wb)

    write_unknown_internals_sheet(unk_int_IPs, wb)

    write_snmp_sheet(snmp_df, wb)

    write_mac_sheet(mac_df, wb)

    auto_adjust_width(wb["Analysis"])

    # Compute capture duration from conn.log timestamps if zeek-cut is available.
    if zeekcut_available and os.path.exists(conn_log_path):
        raw = perform_zeekcut(fields=["ts"], log_file=conn_log_path)
        if raw:
            times = raw.decode("utf-8").splitlines()
            times = [t for t in times if t.strip()]
            if len(times) >= 2:
                forward = sorted(times)
                try:
                    start = float(forward[0])
                    end = float(forward[-1])
                    cap_time = end - start
                    timer_data["Length of Capture time"] = "{} day(s) {} hour(s) {} minutes {} seconds".format(
                        int(cap_time / 86400),
                        int(cap_time % 86400 / 3600),
                        int(cap_time % 3600 / 60),
                        int(cap_time % 60),
                    )
                except ValueError:
                    timer_data["Length of Capture time"] = "Unknown (invalid timestamp values)"
            else:
                timer_data["Length of Capture time"] = "Unknown (insufficient timestamp data)"
        else:
            timer_data["Length of Capture time"] = "Unknown (zeek-cut unavailable or produced no output)"
    else:
        timer_data["Length of Capture time"] = "Unknown (zeek-cut not installed)"

    write_stats_sheet(wb, timer_data)
    write_conn_states_sheet(conn_states, wb)

    # Enforce canonical worksheet order for the saved workbook (locks NAVV-managed sheets in place)
    _reorder_standard_sheets(wb)

    wb.save(file_name)

    if new_workbook:
        success_msg(f"Successfully created file: {file_name}")
    elif pcap:
        success_msg(f"Successfully updated file (Zeek refreshed): {file_name}")
    else:
        success_msg(f"Successfully updated file: {file_name}")


@click.command(
    "launch",
    help=(
        "Launch the NAVV GUI (if installed/enabled).\n\n"
        "Usage\n"
        "  navv launch\n"
    ),
)
def launch():
    """Launch the NAVV GUI."""
    port = 5000
    warning_msg("Launching GUI in browser...")
    webbrowser.open(f"http://127.0.0.1:{port}/")
    app.run(port=port)
