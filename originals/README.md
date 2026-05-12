# Original Tools

These are the two original tools that AutoPhil v1.1 replaces.

## AutoPhil_original.exe
- **Author:** Phillip Tang (phillip.tang@verizonwireless.com)
- **Slack:** #socal-b4c-scripting-tool
- **Built with:** PyInstaller + Python 3.11
- **Purpose:** B4C service port configuration generator (south-facing: B4C to eNB/gNB)
- **Modes:** Management Port, eNB/gNB Add, Carrier Agg (Full Deployment), SpiderCloud (Full Deployment)
- **Vendor support:** Nokia, Nokia IXR-e2, Ciena
- **Date:** September 2025

## Nokia_1.3.5_ENSE_Configs.xlsx
- **Author:** Brian (B4A Port Config 3.0)
- **Purpose:** B4A backhaul link configuration generator (north-facing: B4A to B4C)
- **Format:** Excel workbook with macros
- **Generates:** Port + Interface + ISIS 5 + BGP neighbor configs for both B4C and NNI sides
- **Date:** September 2025

## What Changed

AutoPhil v1.1 combines both tools into a single standalone application:
- B4A Backhaul Link tab replaces the Excel workbook
- B4C Service Ports tab replaces the original AutoPhil .exe
- Dark theme GUI, copy-to-clipboard, save-to-file
- No macros, no Excel dependency
