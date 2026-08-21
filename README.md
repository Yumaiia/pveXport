<p align="center">
  <img src="logo.png" alt="pveXport logo" width="400">
</p>

pveXport is a cli utility aimed to provide a rvtools-like fonctionnality for Proxmox VE hosts and clusters.
It connects to a Proxmox VE host or cluster API and exports its inventory (cluster status, nodes, LXC containers, QEMU VMs, and physical disks) to an Excel workbook.

## Requirements

- Python 3
- Python packages: `requests`, `openpyxl`
- A Proxmox VE account with at least the **PVE.Auditor** permission (read-only access is sufficient and stongly recommended. Use only root access if you trust this script).

Install the dependencies with:

```bash
pip install requests openpyxl
```

## Usage

```bash
python pvexport.py -s <server> -u <user> [options]
```

### Required arguments

| Flag | Description |
|---|---|
| `-s`, `--server` | Proxmox VE cluster URL or IP/hostname (e.g. `https://pve01.example.local:8006` or `192.168.20.170`). If no scheme/port is given, `https://` and port `8006` are assumed. |
| `-u`, `--user` | Proxmox VE username (e.g. `ro@pve`). The `PVE.Auditor` permission is required. |

### Optional arguments

| Flag | Description |
|---|---|
| `-p`, `--password` | Proxmox VE password. If omitted, you will be prompted securely. |
| `-i`, `--insecure` | Skip TLS certificate verification. Use this if the server presents a self-signed or otherwise untrusted certificate. |
| `-o`, `--output` | Output Excel filename. Defaults to `pvexport-<date>-<time>.xlsx`. |
| `-n`, `--no-open` | Do not automatically open the generated file after export (Windows only, where auto-open is otherwise the default). |
| `-v`, `--verbose` | Print detailed debug output, including full raw data dumps for nodes, QEMU VMs, and disks. |
| `-h`, `--help` | Show the help message and exit. |

### Examples

```bash
# Basic export, prompting for the password
python pvexport.py -s https://pve01.example.local:8006 -u readonly@pve

# Using a bare IP address, a self-signed certificate, and a custom output file
python pvexport.py -s 192.168.20.170 -u readonly@pve -i -o cluster_report.xlsx

# Non-interactive, no auto-open (e.g. for scheduled tasks)
python pvexport.py -s pve01 -u ro@pve -p "mypassword" -n
```

## Output

The generated workbook contains the following sheets:

- **Info** — a branded summary page with the pveXport version, the server URL used, and the generation timestamp.
- **Cluster** — cluster-wide status (name, Corosync config version, node count, quorum state).
- **Nodes** — per-node hardware and status details (CPU, memory, uptime, kernel, boot mode, etc.).
- **Resources** — a combined overview of all VMs and LXC containers in the cluster.
- **LXC** — detailed configuration and status for each LXC container.
- **QEMU** — detailed configuration and status for each QEMU virtual machine.
- **Disks** — physical disk inventory across all nodes (model, serial, health, wear, size).

## License

This project is licensed under the **GNU General Public License v3.0** (GPLv3). See the [`LICENSE`](LICENSE) file for the full text.

Copyright (C) 2026 Cyril Pawelko - [Yumaiia](https://www.yumaiia.fr/pvexport/)
