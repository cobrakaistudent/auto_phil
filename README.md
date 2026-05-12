# AutoPhil — ENSE Config Generator

Standalone Windows GUI tool for generating Nokia B4A and B4C ENSE router configurations. Replaces Brian's Excel "B4A Port Config 3.0" and the original AutoPhil B4C Scripting Tool.

## What It Does

Generates copy-paste-ready CLI scripts for Nokia ENSE backhaul routers. Two modes:

### B4A Backhaul Link
North-facing B4A-to-B4C backhaul provisioning. Generates complete configuration scripts for both sides of the link:
- **Port config** — shutdown, description, ethernet mode/encap/MTU/CRC/LLDP/speed
- **Interface config** — address, QoS, BFD, egress redirect group
- **ISIS 5** — level-1, point-to-point, BFD-enabled
- **BGP neighbor** — iBGP with authentication (RR-5-ENSESR_CSR or RR-5-ENSESR_SPOKE)
- **Verification commands** — show port, interface, ARP, BGP summary

All inputs come from CND (hostnames, ports, VLANs, point-to-point IPs, CSR loopbacks).

### B4C Service Ports
South-facing B4C-to-eNB/gNB RAN service port provisioning:

| Script Type | Description |
|-------------|-------------|
| **Management Port** | EDN or VZB/IDN — logical port + VPLS 400/450 SAPs. Predefined descriptions (MGMT, SITE_BOSS, POWER_PLANT, SHARK_METER, IXR-e hooks). IDN auto-sets MTU=2106, VLAN=4000. |
| **eNB/gNB 6630** | Legacy baseband. LTE S1 + OAM SAPs, optional NR (VLAN 310), LTE CA (VLAN 500), NR CA (VLAN 520). |
| **eNB/gNB 6648** | Four deployment variants: LS6, Dual Mode, Tri-Mode, LTE+Low Band. Multi-port with per-variant SAP layout. |
| **eNB/gNB 6672** | Single combined 10G/25G port carrying all services (S1, OAM, NR, LTE CA, NR CA). |

### VPLS Reference
| VPLS | Service |
|------|---------|
| 100 | LTE S1 |
| 310 | NR S1 |
| 400 | OAM |
| 450 | IDN |
| 5000 | LTE CA |
| 5200 | NR CA |

## How to Use

1. Double-click `AutoPhil.exe` on any Windows machine — no install, no Python, no dependencies.
2. Select the tab (B4A Backhaul or B4C Service Ports).
3. Fill in the fields. Defaults are pre-populated where possible.
4. Click the green **Generate** button.
5. **Copy to Clipboard** to paste directly into the router CLI, or **Save to File**.

## Building from Source

If you need to rebuild the `.exe`:

```
pip install pyinstaller
pyinstaller --onefile --windowed --name AutoPhil auto_phil.py
```

Output lands in `dist/AutoPhil.exe`. Or just run `build.bat` on Windows.

The GitHub Actions workflow also builds automatically on every push to `auto_phil.py` — download from the Actions artifacts tab.

## Requirements

- **To run:** Windows (any version). Nothing else needed.
- **To build:** Python 3.6+ and PyInstaller.
- **Source dependencies:** None. Pure Python stdlib + tkinter (bundled with Python).

## Vendor Support

Bottom bar includes Nokia / Nokia IXR-e2 / Ciena selector. Currently generates Nokia SR-OS syntax.
