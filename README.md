# NAVV #

The **NAVV** (**N**etwork **A**rchitecture **V**erification and **V**alidation) tool creates an Excel-based workflow for analyzing network traffic from PCAP data and Zeek logs. NAVV automates Zeek analysis, collates Zeek logs, and produces a structured XLSX workbook that supports iterative network architecture analysis.

NAVV 4.x adds **Purdue Model awareness**, allowing analysts to define Purdue levels, assign them to network segments, and propagate those levels into traffic analysis results for sorting, filtering, and visual inspection.

* [Requirements](#Requirements)
* [Installation](#Installation)
    * [Development](#Development)
    * [Production](#Production)
* [Usage](#Usage)
    * [CLI](#CLI)
    * [Browser](#Browser)
    * [Analysis](#Analysis)
* [Docker](#Docker)
* [Copyright](#Copyright)
* [Contact](#Contact)

[![PyPI Release](https://img.shields.io/pypi/v/navv)](https://pypi.python.org/pypi/navv/)
[![Docker Image](https://github.com/cisagov/network-architecture-verification-and-validation/workflows/navv-build-push-ghcr/badge.svg)](https://github.com/cisagov/network-architecture-verification-and-validation/actions)

## Requirements ##

- This project only works on Linux or MacOS environments
- Zeek must be installed: [Get Zeek](https://zeek.org/get-zeek/)
- Python version 3.10 or later
  - As installation of Python varies from platform to platform, please refer to the [Python.org Wiki](https://wiki.python.org/moin/BeginnersGuide/Download) to install and configure Python on your system.

## Installation ##

### Development ###

If you intend to develop the NAVV tool:
- Verify you have the Zeek tool installed
  - [Install Zeek](https://zeek.org/get-zeek/)
- Clone this repository
  - `git clone https://github.com/cisagov/network-architecture-verification-and-validation.git`
- Setup your local virtual environment
  - `python3 -m venv .venv`
- Activate your local environment
  - `source .venv/bin/activate`
- Install the project and its dependencies to your local virtual environment
  - `pip install -e .`

### Production ###

If you would like to use the NAVV tool, its recommended you install it from PYPI
- Verify you have the Zeek tool installed
  - [Install Zeek](https://zeek.org/get-zeek/)
- Install the project using `pip`
  - The recommended method for installing packages with `pip` is using [User Installs](https://pip.pypa.io/en/stable/user_guide/#user-installs) which installs to a user-specific location rather than system-wide.
  - install the latest version of NAVV
    - `pip install -U navv`
  - or install a specific version of NAVV
    - example:  `pip install -U navv==4.0.0`

Verify the NAVV tool has been installed by running `navv` in your console:

```shell
NAVV: Network Architecture Verification and Validation 4.0.0
Usage: navv [OPTIONS] COMMAND [ARGS]...

  Network Architecture Verification and Validation.

Options:
  --version   Show the version and exit.
  -h, --help  Show this message and exit.

Commands:
  generate  Generate or update a NAVV analysis workbook.
  launch    Launch the optional NAVV GUI.
```

## Usage ##

### CLI ###

To run the NAVV tool in the CLI (Command Line Interface), Run the command: `navv generate`

Below are the available options and commands for `navv generate`:
```shell
Usage: navv generate [OPTIONS] CUSTOMER_NAME

  Generate excel sheet.

Options:
  -o, --output-dir TEXT  Directory to place resultant analysis files in.
                         Defaults to current working directory.
  -p, --pcap TEXT        Path to pcap file. NAVV requires zeek logs or pcap.
                         If used, zeek will run on pcap to create new logs.
  -z, --zeek-logs TEXT   Path to store or contain zeek log files. Defaults to
                         current working directory.
  -h, --help             Show this message and exit.
```

### Browser ###

To launch the NAVV tool in the browser, simply run: `navv launch`

This will automatically launch the tool into your default browser.

![](./docs/images/navv-gui.png)

The user will have two options:

- Generate a New Analysis:
  - Simply upload your PCAP file or a zipped file of your Zeek logs
  - Click Run Analysis
  - An excel sheet will be generated and downloaded via your browser

- Upload an Existing Analysis
  - Modify your generated excel spreadsheet, See [Analysis](#Analysis)
  - Upload your spreadsheet and your zipped Zeek logs file

### Analysis ###

NAVV’s workflow is intentionally **Excel-first**: NAVV generates an `.xlsx` workbook, you annotate it (segments, asset names, Purdue levels), then you re-run NAVV to apply your annotations back onto the traffic analysis.

#### Files NAVV creates

A typical run produces:

- **NAVV workbook**: `<customer>_NAVV.xlsx` (or whatever you name it)
- **Zeek logs**: either generated from a PCAP you provided, or reused from an existing logs directory

The workbook is designed to be edited and re-used across runs.

#### Excel workbook tabs

- **Analysis** – The collated traffic view (Zeek + NAVV enrichment). This is where you sort/filter and visually inspect flows.
- **Segments** – Your network segment definitions (CIDRs / ranges / labels) plus optional color coding. NAVV uses this tab to apply consistent coloring and labeling in `Analysis`.
- **Inventory** – Optional host/IP annotations (names, roles, owner, notes). Inventory colors apply **after** segment colors.
- **Purdue_Definitions** (NAVV 4.x) – Source-of-truth list of Purdue levels, descriptions, and colors used by dropdowns.

#### Suggested workflow

1. **Run NAVV on a PCAP or Zeek logs** to generate the initial workbook.
2. Open the workbook and fill in:
   - **Segments** tab: define segment names and CIDRs/ranges; pick segment colors.
   - **Inventory** tab (optional): label important hosts (jump boxes, historians, PLCs, etc.).
   - **Purdue levels** (NAVV 4.x): assign Purdue levels to segments.
3. **Save the workbook**.
4. **Re-run NAVV** using the same workbook + the same Zeek logs directory.
   - NAVV updates the `Analysis` tab, applies your segment/inventory colors, and keeps your definitions.

#### Segment matching and “guessing”

NAVV will attempt to map each source/destination IP to a segment using the CIDRs/ranges you define on the `Segments` tab.

- If an IP falls inside a defined segment, NAVV labels it and applies that segment’s formatting.
- If an IP does **not** match any defined segment, NAVV leaves it unclassified (so it stands out).

**Practical tip:** Start with coarse segments (a few broad CIDRs), re-run NAVV, then refine segments as you learn more. This is faster than trying to perfectly model the network up-front.

#### Host enrichment (DNS and manual overrides)

When available, NAVV uses Zeek `dns.log` responses to populate `Src_Desc` and `Dest_Desc` in `Analysis`.

- If DNS data is missing or incomplete, add host descriptions in the `Inventory` tab.
- On re-run, NAVV preserves your inventory annotations and re-applies them to the updated analysis.

### Purdue Model Support (NAVV 4.x)

NAVV 4.x introduces a `Purdue_Definitions` worksheet that defines Purdue levels, descriptions, and colors. These definitions act as the source of truth for Purdue classification.

- Purdue levels are assigned to network segments in the `Segments` tab via a dropdown list.
- Assigned Purdue levels propagate into the `Analysis` tab as source and destination Purdue columns.
- Purdue columns support Excel sorting, filtering, and filtering by color.

On the first analysis run (new workbook), NAVV creates default Purdue definitions. On subsequent runs, user-modified Purdue definitions are preserved and reused.

## Docker ##

See [`docker/README.md`](./docker/README.md) for setup and instructions for running the NAVV tool in Docker.

## Copyright ##

[NAVV](https://github.com/cisagov/network-architecture-verification-and-validation) is Copyright 2023 Battelle Energy Alliance, LLC, licensed under the BSD-3 Clause License.

See [`LICENSE`](./LICENSE) for the terms of its release.

Developers, by contributing to this software project, you are agreeing to the following terms and conditions for your contributions:

* You agree your contributions are submitted under the BSD 3-Clause license.
* You represent you are authorized to make the contributions and grant the license. If your employer has rights to intellectual property that includes your contributions, you represent that you have received permission to make contributions and grant the required license on behalf of that employer.

## Other software ##

Idaho National Laboratory is a cutting edge research facility which is constantly producing high quality research and software. Feel free to take a look at our other software and scientific offerings at:

* [Primary Technology Offerings Page](https://www.inl.gov/inl-initiatives/technology-deployment)
* [Supported Open Source Software](https://github.com/cisagov)
* [Raw Experiment Open Source Software](https://github.com/IdahoLabResearch)
* [Unsupported Open Source Software](https://github.com/IdahoLabCuttingBoard)

## Contact ##

Contact information of maintainer(s):

[Seth Grover](mailto:seth.grover@inl.gov?subject=NAVV)

[Donovan Nichols](mailto:donovan.nichols@inl.gov?subject=NAVV)
