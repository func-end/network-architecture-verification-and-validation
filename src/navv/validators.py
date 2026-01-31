from ipaddress import IPv4Address, IPv6Address
import re


def is_ipv4_address(ip_address: str) -> bool:
    """Return True if address is a valid IPv4 address."""
    try:
        IPv4Address(ip_address)
        return True
    except ValueError:
        return False


def is_ipv6_address(ip_address: str) -> bool:
    """Return True if address is a valid IPv6 address."""
    try:
        IPv6Address(ip_address)
        return True
    except ValueError:
        return False


def is_mac_address(mac_address: str) -> bool:
    """Return True if address is a valid MAC address."""
    if re.match(
        "[0-9a-f]{2}([-:])[0-9a-f]{2}(\\1[0-9a-f]{2}){4}$", mac_address.lower()
    ):
        return True
    return False


# ----------------------------
# Purdue helpers
# ----------------------------


def warn_purdue_definition_issues(levels: list[str]) -> list[str]:
    """Validate Purdue level list from the Purdue_Definitions sheet.

    Returns a cleaned list of unique levels preserving order.
    Emits warnings for duplicates.
    """
    # Local import to avoid circular import issues.
    from navv.message_handler import warning_msg

    cleaned: list[str] = []
    seen: set[str] = set()

    for lvl in levels:
        if lvl is None:
            continue
        s = str(lvl).strip()
        if not s:
            continue
        if s in seen:
            warning_msg(
                f"Duplicate Purdue level in Purdue_Definitions: '{s}' (duplicates will be ignored)."
            )
            continue
        seen.add(s)
        cleaned.append(s)

    if not cleaned:
        warning_msg(
            "No Purdue levels found in Purdue_Definitions. Purdue dropdowns/annotations will be empty."
        )

    return cleaned


def warn_invalid_segment_purdue(
    segment_name: str, purdue_level: str, valid_levels: set[str]
) -> None:
    """Warn if a Segments row assigns a Purdue level not present in Purdue_Definitions."""
    from navv.message_handler import warning_msg

    if not purdue_level:
        return
    if purdue_level not in valid_levels:
        warning_msg(
            f"Segment '{segment_name}' assigns Purdue level '{purdue_level}', but it is not present in Purdue_Definitions."
        )
