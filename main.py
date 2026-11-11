#!/usr/bin/env python3

import argparse
import csv
import json
import logging
import os
import platform
import re
import shlex
import shutil
import socket
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

SCHEMA_VERSION = 2
COMMAND_TIMEOUT = 20
LOGGER = logging.getLogger("hw-specs")
CORE_DEPENDENCIES = {
    "lscpu": "util-linux",
    "lsblk": "util-linux",
    "lsmem": "util-linux",
    "lspci": "pciutils",
    "lsusb": "usbutils",
    "dmidecode": "dmidecode",
    "lshw": "lshw",
    "modinfo": "kmod",
    "sensors": "lm-sensors",
    "smartctl": "smartmontools",
    "ip": "iproute2",
}
TOOL_COMMANDS = {
    "bash": ["bash", "--version"],
    "python3": ["python3", "--version"],
    "uv": ["uv", "--version"],
    "pip3": ["pip3", "--version"],
    "gcc": ["gcc", "--version"],
    "g++": ["g++", "--version"],
    "clang": ["clang", "--version"],
    "make": ["make", "--version"],
    "cmake": ["cmake", "--version"],
    "git": ["git", "--version"],
    "java": ["java", "-version"],
    "node": ["node", "--version"],
    "npm": ["npm", "--version"],
    "go": ["go", "version"],
    "rustc": ["rustc", "--version"],
    "cargo": ["cargo", "--version"],
    "dotnet": ["dotnet", "--version"],
    "ruby": ["ruby", "--version"],
    "perl": ["perl", "--version"],
    "php": ["php", "--version"],
    "docker": ["docker", "--version"],
    "podman": ["podman", "--version"],
    "openssl": ["openssl", "version"],
    "openssh": ["ssh", "-V"],
    "curl": ["curl", "--version"],
    "wget": ["wget", "--version"],
    "systemd": ["systemctl", "--version"],
    "networkmanager": ["nmcli", "--version"],
    "smartctl": ["smartctl", "--version"],
}
TABLE_COLUMNS = {
    "system": ["property", "value"],
    "cpu": ["id", "name", "architecture", "version", "vendor"],
    "memory": ["id", "type", "name", "size_bytes", "version"],
    "pci": ["id", "class", "vendor", "device", "version"],
    "usb": ["id", "vendor_id", "product_id", "version", "manufacturer", "product"],
    "storage": ["id", "name", "type", "size", "model", "serial", "fstype"],
    "smart_diagnostics": [
        "id",
        "model",
        "serial",
        "firmware",
        "interface",
        "smart_status",
        "temperature_celsius",
        "power_on_hours",
        "percentage_used",
        "available_spare",
        "reallocated_sectors",
        "pending_sectors",
        "uncorrectable_sectors",
        "udma_crc_errors",
        "smartctl_returncode",
        "error",
    ],
    "network": ["id", "name", "mac_address", "driver", "vendor_id", "device_id"],
    "gpus": ["id", "name", "vendor_id", "device_id", "driver", "driver_version"],
    "audio": ["id", "name", "description", "version"],
    "bluetooth": ["id", "name", "version"],
    "displays": ["id", "status", "enabled"],
    "power": ["id", "type", "name", "manufacturer", "serial_number"],
    "sensors": ["id", "chip", "sensor", "value", "unit"],
    "kernel_modules": ["id", "version", "size_bytes", "state"],
    "software_tools": ["id", "version", "path"],
    "packages": ["id", "version", "manager", "architecture"],
    "hardware_components": ["id", "class", "name", "vendor", "version", "serial"],
}
EXCEL_SHEET_NAMES = {
    "hardware_components": "Hardware components",
    "kernel_modules": "Kernel modules",
    "smart_diagnostics": "SMART diagnostics",
    "software_tools": "Software tools",
}
TABLE_HIDDEN_COLUMNS = {"smart_diagnostics": {"smart_data"}}


def read_text(path: str | Path) -> str | None:
    try:
        return Path(path).read_text(encoding="utf-8", errors="replace").strip()
    except (OSError, ValueError):
        return None


def read_attributes(directory: Path, names: Iterable[str]) -> dict[str, str]:
    result = {}
    for name in names:
        value = read_text(directory / name)
        if value is not None:
            result[name] = value
    return result


def json_output(probe: dict[str, Any]) -> Any:
    if probe.get("returncode") != 0:
        return None
    try:
        return json.loads(probe.get("stdout", ""))
    except (json.JSONDecodeError, TypeError):
        return None


def first_line(text: str) -> str | None:
    return text.splitlines()[0].strip() if text.strip() else None


def bracket_id(value: Any) -> str | None:
    if not isinstance(value, str):
        return None
    matches = re.findall(r"\[([0-9a-fA-F]{4})\]", value)
    return matches[-1] if matches else None


def confirm(question: str) -> bool:
    try:
        return input(f"{question} [y/N]: ").strip().lower() in {"y", "yes"}
    except (EOFError, KeyboardInterrupt):
        print()
        return False


class InventoryCollector:
    def __init__(self, use_sudo: bool = False) -> None:
        self.use_sudo = use_sudo
        self.probes: list[dict[str, Any]] = []
        self.warnings: list[str] = []

    def warn(self, message: str) -> None:
        if message not in self.warnings:
            self.warnings.append(message)
            LOGGER.warning(message)

    def run(
        self,
        command: list[str],
        *,
        privileged: bool = False,
        optional: bool = False,
        timeout: int = COMMAND_TIMEOUT,
        keep_stdout: bool = True,
    ) -> dict[str, Any]:
        executable = command[0]
        actual_command = command.copy()
        if privileged and self.use_sudo and os.geteuid() != 0:
            actual_command.insert(0, "sudo")
        probe: dict[str, Any] = {"command": actual_command, "privileged": privileged}
        path = shutil.which(executable)
        if path is None:
            probe.update({"available": False, "error": "command not found"})
            LOGGER.debug("Command unavailable: %s", executable)
            if not optional:
                self.warn(f"Unavailable command: {executable}")
            self.probes.append(probe)
            return probe

        probe.update({"available": True, "path": path})
        LOGGER.debug("Running command: %s", shlex.join(actual_command))
        started = time.monotonic()
        try:
            completed = subprocess.run(
                actual_command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
                check=False,
            )
        except subprocess.TimeoutExpired:
            probe.update({"error": f"timed out after {timeout} seconds"})
            self.warn(f"Timed out: {' '.join(actual_command)}")
            self.probes.append(probe)
            return probe
        except OSError as error:
            probe.update({"error": str(error)})
            self.warn(f"Could not run {executable}: {error}")
            self.probes.append(probe)
            return probe

        probe.update(
            {
                "returncode": completed.returncode,
                "stdout": completed.stdout.strip() if keep_stdout else "",
                "stderr": completed.stderr.strip(),
                "duration_seconds": round(time.monotonic() - started, 3),
            }
        )
        LOGGER.debug(
            "Command finished: %s (exit=%d, duration=%.3fs)",
            shlex.join(actual_command),
            completed.returncode,
            probe["duration_seconds"],
        )
        if not keep_stdout:
            probe["stdout_omitted"] = True
        if completed.returncode != 0 and not optional:
            detail = first_line(completed.stderr) or f"exit code {completed.returncode}"
            self.warn(f"Failed command: {' '.join(actual_command)} ({detail})")
        self.probes.append(probe)
        return probe

    def system(self) -> dict[str, Any]:
        os_release = {}
        for line in (read_text("/etc/os-release") or "").splitlines():
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            try:
                parts = shlex.split(value)
                os_release[key] = parts[0] if parts else ""
            except ValueError:
                os_release[key] = value.strip('"')

        dmi = read_attributes(
            Path("/sys/class/dmi/id"),
            [
                "sys_vendor",
                "product_name",
                "product_version",
                "product_serial",
                "product_uuid",
                "board_vendor",
                "board_name",
                "board_version",
                "board_serial",
                "chassis_vendor",
                "chassis_type",
                "chassis_version",
                "chassis_serial",
                "bios_vendor",
                "bios_version",
                "bios_date",
            ],
        )
        virtualization = self.run(["systemd-detect-virt"], optional=True)
        return {
            "hostname": socket.gethostname(),
            "fqdn": socket.getfqdn(),
            "machine_id": read_text("/etc/machine-id"),
            "dmi": dmi,
            "os_release": os_release,
            "kernel": {
                "system": platform.system(),
                "release": platform.release(),
                "version": platform.version(),
                "architecture": platform.machine(),
                "processor": platform.processor(),
            },
            "virtualization": virtualization.get("stdout") if virtualization.get("returncode") == 0 else None,
        }

    def cpu(self) -> dict[str, Any]:
        probe = self.run(["lscpu", "--json"])
        data = json_output(probe)
        summary: dict[str, Any] = {}
        if isinstance(data, dict):
            for entry in data.get("lscpu", []):
                if isinstance(entry, dict) and "field" in entry:
                    summary[str(entry["field"]).rstrip(":")] = entry.get("data")
        return {
            "id": "cpu0",
            "name": summary.get("Model name"),
            "architecture": summary.get("Architecture"),
            "version": summary.get("CPU family"),
            "vendor": summary.get("Vendor ID"),
            "logical_cpus": summary.get("CPU(s)"),
            "cores_per_socket": summary.get("Core(s) per socket"),
            "sockets": summary.get("Socket(s)"),
            "threads_per_core": summary.get("Thread(s) per core"),
            "max_mhz": summary.get("CPU max MHz"),
            "min_mhz": summary.get("CPU min MHz"),
            "flags": summary.get("Flags"),
            "details": summary,
        }

    def memory(self) -> list[dict[str, Any]]:
        meminfo = {}
        for line in (read_text("/proc/meminfo") or "").splitlines():
            if ":" in line:
                key, value = line.split(":", 1)
                meminfo[key] = value.strip()
        total_kib = None
        match = re.match(r"(\d+)", meminfo.get("MemTotal", ""))
        if match:
            total_kib = int(match.group(1))
        records = [
            {
                "id": "system-memory",
                "type": "summary",
                "name": "Installed system memory",
                "size_bytes": total_kib * 1024 if total_kib is not None else None,
                "swap": meminfo.get("SwapTotal"),
                "meminfo": meminfo,
            }
        ]
        lsmem = self.run(["lsmem", "--json", "--bytes"], optional=True)
        data = json_output(lsmem)
        if isinstance(data, dict):
            for index, item in enumerate(data.get("memory", [])):
                if isinstance(item, dict):
                    records.append({"id": f"memory-block-{index}", "type": "memory-block", **item})
        self.run(["dmidecode", "--type", "memory"], privileged=True, optional=True)
        return records

    def pci(self) -> list[dict[str, Any]]:
        probe = self.run(["lspci", "-Dnnvmm"])
        devices = []
        for block in probe.get("stdout", "").split("\n\n"):
            values: dict[str, Any] = {}
            for line in block.splitlines():
                try:
                    parts = shlex.split(line)
                except ValueError:
                    continue
                if len(parts) >= 2:
                    values[parts[0].rstrip(":").lower()] = " ".join(parts[1:])
            if values:
                values["id"] = values.get("slot")
                values["version"] = values.get("rev")
                values["class_id"] = bracket_id(values.get("class"))
                values["vendor_id"] = bracket_id(values.get("vendor"))
                values["device_id"] = bracket_id(values.get("device"))
                devices.append(values)
        return devices

    def usb(self) -> list[dict[str, Any]]:
        devices = []
        for device in sorted(Path("/sys/bus/usb/devices").glob("*")):
            if not (device / "idVendor").exists():
                continue
            item: dict[str, Any] = read_attributes(
                device,
                [
                    "idVendor",
                    "idProduct",
                    "bcdDevice",
                    "version",
                    "speed",
                    "manufacturer",
                    "product",
                    "serial",
                    "busnum",
                    "devnum",
                ],
            )
            item.update(
                {
                    "id": device.name,
                    "vendor_id": item.get("idVendor"),
                    "product_id": item.get("idProduct"),
                    "device_version": item.get("bcdDevice"),
                }
            )
            devices.append(item)
        self.run(["lsusb"], optional=True)
        return devices

    def storage(self) -> list[dict[str, Any]]:
        probe = self.run(["lsblk", "--json", "--bytes", "--output-all"])
        data = json_output(probe)
        records = []

        def add_devices(items: list[dict[str, Any]], parent: str | None = None) -> None:
            for item in items:
                children = item.get("children", [])
                record = {key: value for key, value in item.items() if key != "children"}
                record["id"] = item.get("path") or item.get("name")
                record["parent"] = parent
                records.append(record)
                add_devices(children, record["id"])

        if isinstance(data, dict):
            add_devices(data.get("blockdevices", []))
        self.run(["df", "--output=source,fstype,size,used,avail,pcent,target"], optional=True)
        return records

    def smart_diagnostics(self, storage: list[dict[str, Any]]) -> list[dict[str, Any]]:
        records = []
        ata_attributes = {
            5: "reallocated_sectors",
            197: "pending_sectors",
            198: "uncorrectable_sectors",
            199: "udma_crc_errors",
        }
        for disk in storage:
            device_id = disk.get("id")
            if disk.get("type") != "disk" or not str(device_id).startswith("/dev/"):
                continue
            probe = self.run(
                ["smartctl", "--json=c", "-H", "-A", "-i", str(device_id)],
                privileged=True,
                optional=True,
                timeout=30,
            )
            record: dict[str, Any] = {
                "id": device_id,
                "model": disk.get("model"),
                "serial": disk.get("serial"),
                "smartctl_returncode": probe.get("returncode"),
            }
            try:
                data = json.loads(probe.get("stdout", ""))
            except (json.JSONDecodeError, TypeError):
                data = None
            if not isinstance(data, dict):
                error = probe.get("error") or first_line(probe.get("stderr", "")) or "no valid SMART JSON output"
                record["smart_status"] = "UNAVAILABLE"
                record["error"] = error
                self.warn(f"SMART diagnostics unavailable for {device_id}: {error}")
                records.append(record)
                continue

            device = data.get("device") if isinstance(data.get("device"), dict) else {}
            status = data.get("smart_status") if isinstance(data.get("smart_status"), dict) else {}
            passed = status.get("passed")
            nvme = data.get("nvme_smart_health_information_log")
            if not isinstance(nvme, dict):
                nvme = {}
            temperature = data.get("temperature") if isinstance(data.get("temperature"), dict) else {}
            power_on = data.get("power_on_time") if isinstance(data.get("power_on_time"), dict) else {}
            record.update(
                {
                    "model": data.get("model_name") or record["model"],
                    "serial": data.get("serial_number") or record["serial"],
                    "firmware": data.get("firmware_version"),
                    "interface": device.get("protocol") or device.get("type"),
                    "smart_status": "PASSED" if passed is True else "FAILED" if passed is False else "UNKNOWN",
                    "temperature_celsius": temperature.get("current"),
                    "power_on_hours": power_on.get("hours"),
                    "percentage_used": nvme.get("percentage_used"),
                    "available_spare": nvme.get("available_spare"),
                    "smart_data": data,
                }
            )
            smartctl = data.get("smartctl") if isinstance(data.get("smartctl"), dict) else {}
            messages = data.get("messages", []) if isinstance(data.get("messages"), list) else []
            messages += smartctl.get("messages", []) if isinstance(smartctl.get("messages"), list) else []
            errors = [
                str(message["string"])
                for message in messages
                if isinstance(message, dict) and message.get("severity") == "error" and message.get("string")
            ]
            if errors:
                record["error"] = "; ".join(errors)
                if passed is None:
                    record["smart_status"] = "UNAVAILABLE"
                self.warn(f"SMART diagnostics for {device_id}: {record['error']}")
            ata = data.get("ata_smart_attributes") if isinstance(data.get("ata_smart_attributes"), dict) else {}
            table = ata.get("table", [])
            if isinstance(table, list):
                for attribute in table:
                    if not isinstance(attribute, dict) or attribute.get("id") not in ata_attributes:
                        continue
                    raw = attribute.get("raw") if isinstance(attribute.get("raw"), dict) else {}
                    record[ata_attributes[attribute["id"]]] = raw.get("value")
            if passed is False:
                self.warn(f"SMART overall-health status is FAILED for {device_id}.")
            LOGGER.info("Collected SMART diagnostics for %s (status=%s)", device_id, record["smart_status"])
            records.append(record)
        return records

    def network(self) -> list[dict[str, Any]]:
        address_probe = self.run(["ip", "-json", "address", "show"])
        addresses = json_output(address_probe)
        ip_data = {}
        if isinstance(addresses, list):
            ip_data = {item.get("ifname"): item for item in addresses if isinstance(item, dict)}
        interfaces = []
        for interface in sorted(Path("/sys/class/net").glob("*")):
            item: dict[str, Any] = read_attributes(
                interface, ["address", "mtu", "operstate", "speed", "duplex"]
            )
            item.update(
                {
                    "id": interface.name,
                    "name": interface.name,
                    "mac_address": item.get("address"),
                }
            )
            device = interface / "device"
            if device.exists():
                attributes = read_attributes(
                    device, ["vendor", "device", "subsystem_vendor", "subsystem_device", "modalias"]
                )
                item["vendor_id"] = attributes.get("vendor")
                item["device_id"] = attributes.get("device")
                item["modalias"] = attributes.get("modalias")
                driver = device / "driver"
                if driver.is_symlink():
                    item["driver"] = driver.resolve().name
            if interface.name in ip_data:
                item["ip_details"] = ip_data[interface.name]
            interfaces.append(item)
        return interfaces

    def gpus(self, pci_devices: list[dict[str, Any]]) -> list[dict[str, Any]]:
        records = []
        for device in pci_devices:
            if device.get("class_id") in {"0300", "0301", "0302", "0380"}:
                records.append(
                    {
                        "id": device.get("id"),
                        "name": device.get("device"),
                        "vendor": device.get("vendor"),
                        "vendor_id": device.get("vendor_id"),
                        "device_id": device.get("device_id"),
                        "version": device.get("version"),
                    }
                )
        for card in sorted(Path("/sys/class/drm").glob("card[0-9]*")):
            item = read_attributes(
                card / "device",
                ["vendor", "device", "subsystem_vendor", "subsystem_device", "modalias"],
            )
            record: dict[str, Any] = {
                "id": card.name,
                "name": card.name,
                "vendor_id": item.get("vendor"),
                "device_id": item.get("device"),
                "modalias": item.get("modalias"),
            }
            driver = card / "device" / "driver"
            if driver.is_symlink():
                record["driver"] = driver.resolve().name
                version = read_text(Path("/sys/module") / record["driver"] / "version")
                record["driver_version"] = version
            records.append(record)
        nvidia = self.run(
            [
                "nvidia-smi",
                "--query-gpu=index,uuid,name,serial,driver_version,vbios_version,memory.total",
                "--format=csv,noheader,nounits",
            ],
            optional=True,
        )
        if nvidia.get("returncode") == 0:
            for row in csv.reader(nvidia.get("stdout", "").splitlines()):
                if len(row) == 7:
                    values = [value.strip() for value in row]
                    records.append(dict(zip(["index", "uuid", "name", "serial", "driver_version", "vbios_version", "memory_total_mib"], values, strict=True)))
                    records[-1]["id"] = f"nvidia-{values[0]}"
        return records

    def audio(self) -> list[dict[str, Any]]:
        records = []
        for line in (read_text("/proc/asound/cards") or "").splitlines():
            match = re.match(r"\s*(\d+)\s+\[([^\]]+)\]\s*:\s*(.*)", line)
            if match:
                records.append(
                    {
                        "id": match.group(1),
                        "name": match.group(2).strip(),
                        "description": match.group(3).strip(),
                        "version": read_text("/proc/asound/version"),
                    }
                )
        self.run(["pactl", "info"], optional=True)
        return records

    def bluetooth(self) -> list[dict[str, Any]]:
        records = []
        for controller in sorted(Path("/sys/class/bluetooth").glob("*")):
            attributes = read_attributes(
                controller / "device", ["vendor", "device", "subsystem_vendor", "subsystem_device", "modalias"]
            )
            records.append({"id": controller.name, "name": controller.name, **attributes})
        probe = self.run(["bluetoothctl", "show"], optional=True)
        if probe.get("returncode") == 0 and probe.get("stdout"):
            details = {}
            for line in probe["stdout"].splitlines():
                if ":" in line:
                    key, value = line.strip().split(":", 1)
                    details[key.strip()] = value.strip()
            if details:
                records.append({"id": "default-controller", "name": details.get("Name"), "version": details.get("Version"), "details": details})
        return records

    def displays(self) -> list[dict[str, Any]]:
        records = []
        for connector in sorted(Path("/sys/class/drm").glob("card*-*")):
            status = read_text(connector / "status")
            if status is None:
                continue
            try:
                edid_size = len((connector / "edid").read_bytes())
            except OSError:
                edid_size = None
            records.append(
                {
                    "id": connector.name,
                    "status": status,
                    "enabled": read_text(connector / "enabled"),
                    "edid_bytes": edid_size,
                }
            )
        return records

    def power(self) -> list[dict[str, Any]]:
        records = []
        for supply in sorted(Path("/sys/class/power_supply").glob("*")):
            item = read_attributes(
                supply,
                [
                    "type",
                    "capacity",
                    "capacity_level",
                    "status",
                    "present",
                    "manufacturer",
                    "model_name",
                    "serial_number",
                    "technology",
                    "voltage_now",
                    "current_now",
                    "energy_full",
                    "energy_full_design",
                    "charge_full",
                    "charge_full_design",
                    "online",
                ],
            )
            item.update({"id": supply.name, "name": item.get("model_name") or supply.name})
            records.append(item)
        return records

    def sensors(self) -> list[dict[str, Any]]:
        probe = self.run(["sensors"], optional=True)
        records = []
        chip = None
        adapter = None
        occurrences: dict[tuple[str, str], int] = {}
        for line in probe.get("stdout", "").splitlines():
            if not line.strip():
                continue
            if not line[0].isspace() and ":" not in line:
                chip = line.strip()
                adapter = None
                continue
            stripped = line.strip()
            if stripped.startswith("Adapter:"):
                adapter = stripped.split(":", 1)[1].strip()
                continue
            if ":" not in stripped or chip is None:
                continue
            sensor, reading = stripped.split(":", 1)
            sensor = sensor.strip()
            reading = reading.strip()
            key = (chip, sensor)
            occurrences[key] = occurrences.get(key, 0) + 1
            match = re.search(r"([+-]?\d+(?:\.\d+)?)\s*(°C|RPM|V|W|A)?", reading)
            value: float | None = float(match.group(1)) if match else None
            unit = match.group(2) if match else None
            records.append(
                {
                    "id": f"{chip}/{sensor}/{occurrences[key]}",
                    "chip": chip,
                    "adapter": adapter,
                    "sensor": sensor,
                    "measurement": "input",
                    "value": value,
                    "unit": unit,
                    "reading": reading,
                }
            )
        return records

    def hardware_components(self) -> list[dict[str, Any]]:
        probe = self.run(["lshw", "-json"], privileged=True, optional=True, timeout=45)
        data = json_output(probe)
        records = []

        def walk(node: dict[str, Any], parent: str | None = None) -> None:
            children = node.get("children", [])
            record = {key: value for key, value in node.items() if key != "children"}
            record["id"] = node.get("id")
            record["name"] = node.get("product") or node.get("description")
            record["parent"] = parent
            records.append(record)
            for child in children:
                if isinstance(child, dict):
                    walk(child, record["id"])

        if isinstance(data, dict):
            walk(data)
        elif isinstance(data, list):
            for node in data:
                if isinstance(node, dict):
                    walk(node)
        return records

    def kernel_modules(self) -> list[dict[str, Any]]:
        modules = []
        for line in (read_text("/proc/modules") or "").splitlines():
            fields = line.split()
            if len(fields) < 6:
                continue
            name = fields[0]
            version = read_text(Path("/sys/module") / name / "version")
            if not version:
                probe = self.run(["modinfo", "--field", "version", name], optional=True)
                version = probe.get("stdout") or None
            modules.append(
                {
                    "id": name,
                    "version": version,
                    "size_bytes": int(fields[1]),
                    "instance_count": int(fields[2]),
                    "dependencies": [item for item in fields[3].split(",") if item],
                    "state": fields[4],
                    "memory_address": fields[5],
                }
            )
        return modules

    def packages(self) -> list[dict[str, Any]]:
        packages = []
        dpkg = self.run(
            [
                "dpkg-query",
                "--show",
                "--showformat=${binary:Package}\t${Version}\t${Architecture}\t${db:Status-Abbrev}\n",
            ],
            optional=True,
        )
        if dpkg.get("returncode") == 0:
            for line in dpkg.get("stdout", "").splitlines():
                fields = line.split("\t")
                if len(fields) == 4 and fields[3].startswith("ii"):
                    packages.append({"id": fields[0], "version": fields[1], "architecture": fields[2], "manager": "dpkg"})

        snap = self.run(["snap", "list"], optional=True)
        if snap.get("returncode") == 0:
            for line in snap.get("stdout", "").splitlines()[1:]:
                fields = line.split(None, 5)
                if len(fields) >= 5:
                    packages.append(
                        {
                            "id": fields[0],
                            "version": fields[1],
                            "manager": "snap",
                            "revision": fields[2],
                            "channel": fields[3],
                            "publisher": fields[4],
                            "notes": fields[5] if len(fields) == 6 else "",
                        }
                    )

        flatpak = self.run(
            ["flatpak", "list", "--app", "--columns=application,version,branch,origin,installation"],
            optional=True,
        )
        if flatpak.get("returncode") == 0:
            for line in flatpak.get("stdout", "").splitlines():
                fields = line.split("\t")
                if len(fields) == 5:
                    packages.append(dict(zip(["id", "version", "branch", "origin", "installation"], fields, strict=True), manager="flatpak"))
        return packages

    def software_tools(self) -> list[dict[str, Any]]:
        tools = [
            {
                "id": "python-runtime",
                "version": platform.python_version(),
                "path": sys.executable,
                "implementation": platform.python_implementation(),
            }
        ]
        for name, command in TOOL_COMMANDS.items():
            path = shutil.which(command[0])
            if path is None:
                continue
            probe = self.run(command, optional=True)
            output = probe.get("stdout") or probe.get("stderr") or ""
            tools.append({"id": name, "version": first_line(output), "path": path, "returncode": probe.get("returncode")})
        return tools

    def collect(self) -> dict[str, Any]:
        LOGGER.info("Starting hardware collection")
        pci_devices = self.pci()
        storage = self.storage()
        hardware = {
            "cpu": self.cpu(),
            "memory": self.memory(),
            "pci": pci_devices,
            "usb": self.usb(),
            "storage": storage,
            "smart_diagnostics": self.smart_diagnostics(storage),
            "network": self.network(),
            "gpus": self.gpus(pci_devices),
            "audio": self.audio(),
            "bluetooth": self.bluetooth(),
            "displays": self.displays(),
            "power": self.power(),
            "sensors": self.sensors(),
            "hardware_components": self.hardware_components(),
        }
        report = {
            "schema_version": SCHEMA_VERSION,
            "generated_at": datetime.now().astimezone().isoformat(),
            "collector": {"script_version": "1.1.0", "python_version": platform.python_version()},
            "system": self.system(),
            "hardware": hardware,
            "software": {
                "kernel_modules": self.kernel_modules(),
                "packages": self.packages(),
                "tools": self.software_tools(),
            },
        }
        report["probes"] = self.probes
        report["warnings"] = self.warnings
        return report


def preflight(args: argparse.Namespace, collector: InventoryCollector) -> bool:
    LOGGER.info("Starting dependency and privilege preflight")
    missing = {command: package for command, package in CORE_DEPENDENCIES.items() if shutil.which(command) is None}
    if missing:
        packages = sorted(set(missing.values()))
        command = ["apt-get", "install", "-y", *packages]
        display_command = " ".join(command if os.geteuid() == 0 else ["sudo", *command])
        message = "Missing inventory commands: " + ", ".join(f"{cmd} ({pkg})" for cmd, pkg in missing.items())
        print(message)
        LOGGER.info(message)
        print(f"Proposed installation: {display_command}")
        LOGGER.info("Proposed installation: %s", display_command)
        install = args.install
        if install is None and sys.stdin.isatty():
            install = confirm("Install the missing packages now?")
        if install:
            if shutil.which("apt-get") is None:
                collector.warn("Automatic installation requires apt-get; install the listed packages manually.")
            else:
                actual = command if os.geteuid() == 0 else ["sudo", *command]
                print("Running: " + " ".join(actual))
                LOGGER.info("Running package installation: %s", shlex.join(actual))
                result = subprocess.run(actual, check=False)
                LOGGER.info("Package installation finished with exit code %d", result.returncode)
                if result.returncode != 0:
                    collector.warn(f"Package installation failed with exit code {result.returncode}.")
        else:
            collector.warn("Missing packages were not installed; related sections may be incomplete.")
    elif args.install:
        print("All core inventory commands are already installed; nothing to install.")

    if os.geteuid() == 0:
        LOGGER.info("Running as root; privileged probes enabled")
        return True
    use_sudo = args.sudo
    if use_sudo is None and sys.stdin.isatty():
        use_sudo = confirm("May privileged hardware probes (lshw, dmidecode, and smartctl) use sudo?")
    if use_sudo:
        if shutil.which("sudo") is None:
            collector.warn("sudo was requested but is not installed; using unprivileged probes.")
            return False
        LOGGER.info("Requesting sudo authentication")
        result = subprocess.run(["sudo", "-v"], check=False)
        if result.returncode != 0:
            collector.warn("sudo authentication failed; using unprivileged probes.")
            return False
        LOGGER.info("sudo authentication succeeded")
        return True
    collector.warn("Privileged probes were not authorized; firmware, hardware, and SMART details may be incomplete.")
    return False


def flatten_value(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return json.dumps(value, ensure_ascii=False, sort_keys=True)


def system_rows(system: dict[str, Any]) -> list[dict[str, Any]]:
    rows = []
    for key, value in system.items():
        if isinstance(value, dict):
            for child_key, child_value in value.items():
                rows.append({"property": f"{key}.{child_key}", "value": child_value})
        else:
            rows.append({"property": key, "value": value})
    return rows


def build_tables(report: dict[str, Any]) -> dict[str, list[dict[str, Any]]]:
    hardware = report["hardware"]
    software = report["software"]
    return {
        "system": system_rows(report["system"]),
        "cpu": [hardware["cpu"]],
        "memory": hardware["memory"],
        "pci": hardware["pci"],
        "usb": hardware["usb"],
        "storage": hardware["storage"],
        "smart_diagnostics": hardware["smart_diagnostics"],
        "network": hardware["network"],
        "gpus": hardware["gpus"],
        "audio": hardware["audio"],
        "bluetooth": hardware["bluetooth"],
        "displays": hardware["displays"],
        "power": hardware["power"],
        "sensors": hardware["sensors"],
        "kernel_modules": software["kernel_modules"],
        "software_tools": software["tools"],
        "packages": software["packages"],
        "hardware_components": hardware["hardware_components"],
    }


def table_headers(name: str, rows: list[dict[str, Any]]) -> list[str]:
    preferred = TABLE_COLUMNS[name]
    hidden = TABLE_HIDDEN_COLUMNS.get(name, set())
    additional = sorted({key for row in rows for key in row} - set(preferred) - hidden)
    return [key for key in preferred if any(key in row for row in rows)] + additional


def write_csv_tables(tables: dict[str, list[dict[str, Any]]], directory: Path) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    for name, rows in tables.items():
        headers = table_headers(name, rows) or TABLE_COLUMNS[name]
        with (directory / f"{name}.csv").open("w", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=headers, extrasaction="ignore")
            writer.writeheader()
            for row in rows:
                writer.writerow({key: flatten_value(row.get(key)) for key in headers})


def excel_value(value: Any) -> Any:
    value = flatten_value(value)
    if isinstance(value, str):
        value = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", value)
        return value[:32767]
    return value


def write_excel(tables: dict[str, list[dict[str, Any]]], path: Path) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    for name, rows in tables.items():
        sheet = workbook.create_sheet(EXCEL_SHEET_NAMES.get(name, name.replace("_", " ").title()))
        headers = table_headers(name, rows) or TABLE_COLUMNS[name]
        sheet.append([header.replace("_", " ").title() for header in headers])
        for row in rows:
            sheet.append([excel_value(row.get(header)) for header in headers])
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        for index, header in enumerate(headers, start=1):
            values = [str(header)] + [str(excel_value(row.get(header)) or "") for row in rows[:200]]
            sheet.column_dimensions[get_column_letter(index)].width = min(max(len(value) for value in values) + 2, 60)
    workbook.save(path)


def markdown_escape(value: Any) -> str:
    if value is None or value == "":
        return "—"
    return str(value).replace("|", "\\|").replace("\n", " ")


def markdown_table(headers: list[str], rows: list[list[Any]]) -> list[str]:
    lines = ["| " + " | ".join(headers) + " |", "| " + " | ".join("---" for _ in headers) + " |"]
    lines.extend("| " + " | ".join(markdown_escape(value) for value in row) + " |" for row in rows)
    return lines


def human_bytes(value: Any) -> str:
    if not isinstance(value, (int, float)):
        return "—"
    amount = float(value)
    for unit in ["B", "KiB", "MiB", "GiB", "TiB", "PiB"]:
        if abs(amount) < 1024 or unit == "PiB":
            return f"{amount:.1f} {unit}"
        amount /= 1024
    return str(value)


def render_markdown(report: dict[str, Any]) -> str:
    system = report["system"]
    hardware = report["hardware"]
    software = report["software"]
    os_release = system["os_release"]
    dmi = system["dmi"]
    cpu = hardware["cpu"]
    memory = hardware["memory"][0] if hardware["memory"] else {}
    package_counts = {}
    for package in software["packages"]:
        manager = package["manager"]
        package_counts[manager] = package_counts.get(manager, 0) + 1

    lines = [
        f"# Machine specifications: {system['hostname']}",
        "",
        f"Generated: {report['generated_at']}",
        "",
        "## System overview",
        "",
    ]
    lines.extend(
        markdown_table(
            ["Property", "Value"],
            [
                ["Hostname", system["hostname"]],
                ["Machine ID", system.get("machine_id")],
                ["Operating system", os_release.get("PRETTY_NAME")],
                ["OS ID / version", f"{os_release.get('ID', '—')} / {os_release.get('VERSION_ID', '—')}"],
                ["Kernel", system["kernel"]["release"]],
                ["Kernel build", system["kernel"]["version"]],
                ["Architecture", system["kernel"]["architecture"]],
                ["Virtualization", system.get("virtualization")],
                ["System", f"{dmi.get('sys_vendor', '—')} {dmi.get('product_name', '—')}"],
                ["Product version", dmi.get("product_version")],
                ["Product serial / UUID", f"{dmi.get('product_serial', '—')} / {dmi.get('product_uuid', '—')}"],
                ["Board", f"{dmi.get('board_vendor', '—')} {dmi.get('board_name', '—')} {dmi.get('board_version', '—')}"],
                ["BIOS", f"{dmi.get('bios_vendor', '—')} {dmi.get('bios_version', '—')} ({dmi.get('bios_date', '—')})"],
                ["CPU", cpu.get("name")],
                ["CPU vendor / family", f"{cpu.get('vendor', '—')} / {cpu.get('version', '—')}"],
                ["Sockets / cores per socket / logical CPUs", f"{cpu.get('sockets', '—')} / {cpu.get('cores_per_socket', '—')} / {cpu.get('logical_cpus', '—')}"],
                ["Installed memory", human_bytes(memory.get("size_bytes"))],
            ],
        )
    )

    lines.extend(["", "## PCI devices", ""])
    lines.extend(
        markdown_table(
            ["Slot", "Class", "Vendor", "Device", "Revision"],
            [[item.get("id"), item.get("class"), item.get("vendor"), item.get("device"), item.get("version")] for item in hardware["pci"]],
        )
    )
    lines.extend(["", "## USB devices", ""])
    lines.extend(
        markdown_table(
            ["ID", "Vendor:Product", "Device version", "Manufacturer", "Product", "Serial"],
            [[item.get("id"), f"{item.get('vendor_id', '—')}:{item.get('product_id', '—')}", item.get("device_version"), item.get("manufacturer"), item.get("product"), item.get("serial")] for item in hardware["usb"]],
        )
    )
    lines.extend(["", "## Storage", ""])
    lines.extend(
        markdown_table(
            ["Device", "Type", "Size", "Model", "Serial", "Filesystem", "Mount points"],
            [[item.get("id"), item.get("type"), human_bytes(item.get("size")), item.get("model"), item.get("serial"), item.get("fstype"), ", ".join(str(value) for value in (item.get("mountpoints") or []) if value)] for item in hardware["storage"]],
        )
    )
    lines.extend(["", "## Drive health (SMART)", ""])
    lines.extend(
        markdown_table(
            ["Device", "Model", "SMART status", "Temperature (°C)", "Power-on hours", "Wear used (%)", "Reallocated", "Pending", "Uncorrectable", "CRC errors", "Error"],
            [
                [
                    item.get("id"),
                    item.get("model"),
                    item.get("smart_status"),
                    item.get("temperature_celsius"),
                    item.get("power_on_hours"),
                    item.get("percentage_used"),
                    item.get("reallocated_sectors"),
                    item.get("pending_sectors"),
                    item.get("uncorrectable_sectors"),
                    item.get("udma_crc_errors"),
                    item.get("error"),
                ]
                for item in hardware["smart_diagnostics"]
            ],
        )
    )
    lines.extend(["", "SMART data comes from passive read-only checks; no drive self-tests were started. A passed status does not guarantee that a drive is fault-free.", ""])
    lines.extend(["## Network interfaces", ""])
    lines.extend(
        markdown_table(
            ["Interface", "State", "MAC address", "Driver", "Vendor ID", "Device ID"],
            [[item.get("name"), item.get("operstate"), item.get("mac_address"), item.get("driver"), item.get("vendor_id"), item.get("device_id")] for item in hardware["network"]],
        )
    )
    lines.extend(["", "## GPUs", ""])
    lines.extend(
        markdown_table(
            ["ID", "Name", "Driver", "Driver version", "VBIOS"],
            [[item.get("id"), item.get("name"), item.get("driver"), item.get("driver_version"), item.get("vbios_version")] for item in hardware["gpus"]],
        )
    )
    lines.extend(["", "## Software tools", ""])
    lines.extend(
        markdown_table(
            ["Tool", "Version", "Path"],
            [[item.get("id"), item.get("version"), item.get("path")] for item in software["tools"]],
        )
    )
    lines.extend(["", "## Inventory totals", ""])
    lines.extend(
        markdown_table(
            ["Category", "Records"],
            [
                ["Hardware components", len(hardware["hardware_components"])],
                ["Loaded kernel modules", len(software["kernel_modules"])],
                *[[f"{manager} packages", count] for manager, count in sorted(package_counts.items())],
                ["Sensor measurements", len(hardware["sensors"])],
            ],
        )
    )
    lines.extend(["", "## Collection warnings", ""])
    lines.extend([f"- {warning}" for warning in report["warnings"]] or ["No warnings."])
    lines.extend(
        [
            "",
            "## Detailed data",
            "",
            "- `tables/`: one UTF-8 CSV file per hardware/software category.",
            "- `inventory.xlsx`: the same tables in formatted worksheets.",
            "- `inventory.json`: canonical complete inventory, including raw probe output and errors.",
            "",
        ]
    )
    return "\n".join(lines)


def configure_logging(log_dir: Path) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().astimezone().strftime("%Y%m%dT%H%M%S%z")
    log_path = log_dir / f"run-{timestamp}.log"
    LOGGER.setLevel(logging.DEBUG)
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter("%(levelname)s %(message)s"))
    LOGGER.addHandler(file_handler)
    LOGGER.addHandler(console_handler)
    return log_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Inventory Linux hardware/software and export JSON, CSV, XLSX, and Markdown.")
    parser.add_argument("--output-dir", type=Path, default=Path("data/out"), help="parent output directory (default: data/out)")
    parser.add_argument("--log-dir", type=Path, default=Path("data/logs"), help="run log directory (default: data/logs)")
    parser.add_argument("--prefix", default="machine-specs", help="report directory prefix (default: machine-specs)")
    parser.add_argument(
        "--install",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="install missing core tools after showing the apt command; default is to ask interactively",
    )
    parser.add_argument(
        "--sudo",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="allow sudo for privileged hardware probes; default is to ask interactively",
    )
    args = parser.parse_args()
    if not args.prefix or Path(args.prefix).name != args.prefix or args.prefix in (".", ".."):
        parser.error("--prefix must be a non-empty directory name without path components")
    return args


def main() -> None:
    args = parse_args()
    log_path = configure_logging(args.log_dir)
    LOGGER.info("Starting hw-specs 1.1.0")
    try:
        collector = InventoryCollector()
        collector.use_sudo = preflight(args, collector)
        print("Collecting machine inventory...")
        report = collector.collect()
        timestamp = datetime.now().astimezone().strftime("%Y%m%dT%H%M%S%z")
        report_dir = args.output_dir / f"{args.prefix}-{timestamp}"
        report_dir.mkdir(parents=True, exist_ok=False)
        LOGGER.info("Writing report to %s", report_dir.resolve())

        json_path = report_dir / "inventory.json"
        json_path.write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        tables = build_tables(report)
        write_csv_tables(tables, report_dir / "tables")
        write_excel(tables, report_dir / "inventory.xlsx")
        (report_dir / "report.md").write_text(render_markdown(report), encoding="utf-8")

        LOGGER.info("Inventory run completed with %d warnings", len(report["warnings"]))
        print(f"Report directory: {report_dir.resolve()}")
        print(f"Run log: {log_path.resolve()}")
        print(f"CSV tables: {len(tables)}")
        print(f"Collection warnings: {len(report['warnings'])}")
    except Exception:
        LOGGER.exception("Inventory run failed")
        raise


if __name__ == "__main__":
    main()
