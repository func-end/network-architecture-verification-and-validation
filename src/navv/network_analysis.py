#!/usr/bin/env python3

# Copyright 2023 Battelle Energy Alliance, LLC

# python std library imports

from importlib.resources import files

# third party imports
import click

# package imports
from navv.commands import generate, launch
from navv.message_handler import info_msg
from navv._version import __version__


CONTEXT_SETTINGS = dict(help_option_names=["-h", "--help"])
HEADER = f"NAVV: Network Architecture Verification and Validation {__version__}"

DATA_PATH = str(files("navv").joinpath("data"))


@click.group(context_settings=CONTEXT_SETTINGS, invoke_without_command=True)
@click.version_option(__version__)
@click.pass_context
def cli(ctx):
    """Network Architecture Verification and Validation (NAVV).

    \b
    Primary commands:
      generate   Create or update the NAVV Excel workbook
      launch     Launch the optional GUI

    \b
    Notes:
      • On each run, NAVV enforces a canonical worksheet tab order in the output workbook.
        This keeps Analysis/Purdue/Segments/Purdue_Definitions in consistent positions.

    \b
    For detailed usage and examples:
      navv generate --help
      navv launch --help
    """
    if ctx.invoked_subcommand is None:
        info_msg(HEADER)
        print(ctx.command.get_help(ctx))
    pass


def main():
    """Main function for performing zeek-cut commands and sorting the output"""

    cli.add_command(generate)
    cli.add_command(launch)
    cli()


if __name__ == "__main__":
    main()
