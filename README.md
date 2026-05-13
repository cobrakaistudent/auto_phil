# AutoPhil 2 — ENSE Config Generator

Multi-vendor Nokia/IXR-e2/Ciena configuration generator for Verizon ENSE backhaul routers. Standalone Windows GUI tool — no dependencies, no install.

**Replaces:** Brian's Excel "B4A Port Config 3.0" + Phillip Tang's AutoPhil B4C Scripting Tool.

## Vendor Support

| Vendor | CLI Type | Scope |
|--------|----------|-------|
| **Nokia SR-OS** | Classic CLI (7750 SR) | B4A + B4C |
| **Nokia IXR-e2** | SR Linux MD-CLI (7215) | B4C only |
| **Ciena SAOS** | SAOS 10.x (39xx/51xx) | B4C only |

## What It Does

Generates copy-paste-ready CLI scripts for ENSE routers. Two modes:

### B4A Backhaul Link (Nokia SR-OS only)
North-facing B4A-to-B4C backhaul provisioning. Both sides are always Nokia 7750 SR.
- **Port config** — shutdown, description, ethernet mode/encap/MTU/CRC/LLDP/speed
- **Interface config** — address, QoS, BFD, egress redirect group
- **ISIS 5** — level-1, point-to-point, BFD-enabled
- **BGP neighbor** — iBGP with authentication (RR-5-ENSESR_CSR or RR-5-ENSESR_SPOKE)
- **Verification commands** — show port, interface, ARP, BGP summary

Input fields match Brian's Excel yellow cells exactly:
Host Name → Management IP → NNI/B4A/B4C Port → ODD VLAN → Bandwidth → B4C P2P IPv4 → B4C CSR GLOBAL LoO

### B4C Service Ports (all vendors)
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

1. Double-click `AutoPhil2.exe` on any Windows machine — no install, no Python, no dependencies.
2. Select vendor (top-right): **Nokia SR-OS**, **Nokia IXR-e2**, or **Ciena SAOS**.
3. Select the tab: **B4A Backhaul Link** or **B4C Service Ports**.
4. Fill in the fields. Defaults are pre-populated where possible.
5. Click the green **Generate** button.
6. **Copy to Clipboard** to paste directly into the router CLI, or **Save to File**.
7. Use the **Validation Table** at the bottom of output to track PRE/POST verification checks.

## Building from Source

```
pip install pyinstaller
pyinstaller --onefile --windowed --name AutoPhil2 auto_phil2.py
```

Output lands in `dist/AutoPhil2.exe`. Or run `build.bat` on Windows.

GitHub Actions auto-builds on every push to `auto_phil2.py` — download from the Actions artifacts tab.

## Requirements

- **To run:** Windows (any version). Nothing else needed.
- **To build:** Python 3.6+ and PyInstaller.
- **Source dependencies:** None. Pure Python stdlib + tkinter (bundled with Python).

If Verizon endpoint security blocks the `.exe`, run the source directly: `python auto_phil2.py`

## File Structure

```
auto_phil/
├── auto_phil2.py          # v2.0 source (current)
├── auto_phil.py           # v1.1 source (archived)
├── build.bat              # Windows build script
├── .github/workflows/     # GitHub Actions CI
├── originals/             # Original tools (Phillip's .exe + Brian's Excel)
└── README.md
```

## Changelog

### v2.0
- **Multi-vendor:** Nokia SR-OS, Nokia IXR-e2 (SR Linux), Ciena SAOS
- **Vendor selector wired** — actually changes generated output
- **Test/validation table** — appended to every config for change-control evidence
- **Status bar** — inline feedback instead of popup messageboxes
- **B4A fields match Excel** — same order as Brian's yellow cells, added Management IP
- **Bug fixes:** TPL_CA520 description, TPL_6648_LTE_LB indentation, B4A spacing, combobox graying out on Windows

### v1.1
- Initial combined tool (Nokia SR-OS only)
- B4A Backhaul + B4C Service Ports in one GUI
