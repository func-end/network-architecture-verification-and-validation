import json
import os
from subprocess import Popen, PIPE, STDOUT, check_call, CalledProcessError
import shutil
from typing import List, Optional

from navv.message_handler import error_msg
from navv.utilities import pushd, timeit, trim_dns_data


@timeit
def get_conn_data(zeek_logs):
    """Return a list of Zeek conn.log data."""
    return (
        perform_zeekcut(
            fields=[
                "id.orig_h",
                "id.resp_h",
                "id.resp_p",
                "proto",
                "conn_state",
                "orig_l2_addr",
                "resp_l2_addr",
            ],
            log_file=os.path.join(zeek_logs, "conn.log"),
        )
        .decode("utf-8")
        .split("\n")[:-1]
    )


@timeit
def get_dns_data(customer_name, output_dir, zeek_logs):
    """Get DNS data from zeek logs or from a json file if it exists"""
    json_path = os.path.join(output_dir, f"{customer_name}_dns_data.json")
    if os.path.exists(json_path):
        with open(json_path, "rb") as json_file:
            return json.load(json_file)

    dns_data = perform_zeekcut(
        fields=["query", "answers", "qtype", "rcode_name"],
        log_file=os.path.join(zeek_logs, "dns.log"),
    )
    return trim_dns_data(dns_data)


@timeit
def get_snmp_data(zeek_logs):
    """Get SNMP data from zeek logs or from a json file if it exists"""
    return (
        perform_zeekcut(
            fields=[
                "id.orig_h",
                "id.orig_p",
                "id.resp_h",
                "id.resp_p",
                "version",
                "community",
            ],
            log_file=os.path.join(zeek_logs, "snmp.log"),
        )
        .decode("utf-8")
        .split("\n")[:-1]
    )


def _which(cmd: str) -> Optional[str]:
    """Return full path to command or None if not found."""
    return shutil.which(cmd)


def _require_cmd(cmd: str, install_hint: str) -> str:
    """Return command path or raise a RuntimeError with a helpful message."""
    path = _which(cmd)
    if path is None:
        raise RuntimeError(
            f"Required dependency '{cmd}' was not found on PATH. {install_hint}"
        )
    return path


def perform_zeekcut(fields: List[str], log_file: str) -> bytes:
    """Perform the call to zeek-cut with the identified fields on the specified log file"""
    try:
        zeekcut_path = _require_cmd(
            "zeek-cut",
            "Install Zeek (which provides zeek-cut) or run NAVV using pre-generated Zeek logs.",
        )
        with open(log_file, "rb") as f:
            zeekcut = Popen(
                [zeekcut_path] + fields, stdout=PIPE, stdin=PIPE, stderr=STDOUT
            )
            return zeekcut.communicate(input=f.read())[0]
    except FileNotFoundError:
        # log file does not exist
        return b""
    except RuntimeError as exc:
        # zeek-cut missing
        error_msg(exc)
        return b""
    except OSError:
        return b""


@timeit
def run_zeek(pcap_path, zeek_logs_path, **kwargs):
    """Run Zeek on a PCAP into the provided logs directory.

    This function must fail fast (and loudly) when Zeek is not installed or when
    Zeek execution fails, otherwise NAVV will continue with empty logs and crash
    later with confusing errors.
    """
    try:
        zeek_path = _require_cmd(
            "zeek",
            "Install Zeek (e.g., 'brew install zeek' on macOS) or rerun NAVV without -p using existing Zeek logs (-z).",
        )
    except RuntimeError as exc:
        error_msg(exc)
        # Re-raise so callers can abort cleanly.
        raise

    # Ensure output directory exists
    os.makedirs(zeek_logs_path, exist_ok=True)

    with pushd(zeek_logs_path):
        # can we add Site::local_nets to the zeek call here?
        try:
            check_call([zeek_path, "-C", "-r", pcap_path, "local.zeek"])
        except (CalledProcessError, OSError, Exception) as exc:
            error_msg(exc)
            raise
