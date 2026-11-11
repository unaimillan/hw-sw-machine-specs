# Hardware and software specifications

A Python-first Linux machine inventory tool. It reads `/proc`, `/sys`, and OS configuration directly where practical, and calls Linux utilities for richer hardware, package, and drive-health information.

## Run

From this directory:

```bash
uv run main.py
```

For the most complete hardware and SMART diagnostics, run with root privileges:

```bash
sudo `which uv` run main.py
```

On an interactive terminal, the script asks two separate questions when relevant:

1. Whether to install missing core inventory tools with `apt-get`. It lists the missing commands, packages, and exact installation command first.
2. Whether `lshw`, `dmidecode`, and `smartctl` may use `sudo` for additional firmware, hardware, and drive-health details.

Both prompts default to **No**. The script never reads or stores your sudo password; authentication is handled by `sudo` itself. Declining either prompt still produces a report, with warnings about incomplete sections.

For automation or an unprivileged run without prompts:

```bash
uv run main.py --no-install --no-sudo
```

To explicitly permit installation and privileged probes without running the whole collector as root:

```bash
uv run main.py --install --sudo
```

`--install` is only a consent flag: when all core tools are present, nothing is installed. Automatic installation supports Ubuntu/Debian systems with `apt-get`; on other distributions, install the listed packages using the system package manager. SMART collection requires `smartctl` from `smartmontools`:

```bash
sudo apt update
sudo apt install -y smartmontools
```

Optional tools such as NVIDIA utilities and Flatpak are detected but are not installed automatically.

## Output

Each run creates a timestamped report directory such as `data/out/machine-specs-20261111T203000+0300/`:

- `report.md` — readable system, hardware, software, and drive-health overview with units and collection warnings.
- `inventory.xlsx` — multi-sheet Excel workbook with filters, frozen headers, and adjusted column widths.
- `tables/*.csv` — one UTF-8 CSV table per category, suitable for spreadsheets and further processing.
- `inventory.json` — canonical complete data, including nested details, full SMART data, raw command output, exit codes, and probe errors.

Each run also writes a timestamped text log such as `data/logs/run-20261111T203000+0300.log`. Logs record preflight decisions, commands, exit codes, durations, warnings, errors, and output paths. Bulky command output remains in `inventory.json` rather than being duplicated in the log.

Tables cover system information, CPU, memory, PCI, USB, storage, SMART diagnostics, network, GPUs, audio, Bluetooth, displays, batteries/power supplies, sensors, kernel modules, software tools, installed packages, and the hardware tree reported by `lshw`. IDs and versions are included whenever exposed by the operating system or device.

SMART checks use `smartctl --json=c -H -A -i` on whole disks. These are passive, non-destructive reads; the script does not start short/long self-tests or change drive settings. A passed SMART status does not guarantee that a drive is fault-free. Missing permissions, unsupported devices, and unavailable SMART data are reported as warnings or `UNAVAILABLE`, not treated as healthy results.

Customize the output, log location, and directory name:

```bash
uv run main.py --output-dir data/out --log-dir data/logs --prefix my-workstation
```

Run `uv run main.py --help` for all options.

## Privacy and permissions

Reports can contain machine IDs, hardware serial numbers, UUIDs, MAC addresses, IP addresses, and installed package names. Logs can contain command paths, device names, and host-specific errors. Everything is written locally and is not transmitted. Review files before sharing.

Root access is optional. Unprivileged reports can be less complete because some DMI, hardware, and SMART data is restricted. Unless the entire script is started with `sudo`, it only prefixes designated hardware probes and package installation with `sudo`.
