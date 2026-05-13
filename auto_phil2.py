#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AutoPhil 2 — ENSE Config Generator v2.0

Multi-vendor Nokia/IXR-e2/Ciena B4A + B4C ENSE configuration tool.

Replaces:
  - AutoPhil v1.1 (Nokia-only)
  - Brian's Excel "B4A Port Config 3.0"
  - Phillip Tang's AutoPhil B4C Scripting Tool .exe

Three vendor targets:
  - Nokia SR-OS    (classic CLI — 7750 SR / 7250 IXR)
  - Nokia IXR-e2   (SR Linux MD-CLI — 7215 IXR-e2)
  - Ciena SAOS     (SAOS 10.x — 39xx/51xx/81xx)

Two modes:
  1. B4A Backhaul Link    B4A <-> B4C: Port + Interface + ISIS + BGP (Nokia SR-OS only)
  2. B4C Service Ports    B4C <-> eNB/gNB: Physical ports + VPLS SAPs (all vendors)

Standalone GUI — tkinter only, no external dependencies.
Package: pyinstaller --onefile --windowed --name AutoPhil2 auto_phil2.py
"""

import sys
import os
import re
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

VERSION = "2.0"

# ===========================================================================
#  COLORS / THEME
# ===========================================================================

BG_DARK = "#0f0f1a"
BG_MID = "#161625"
BG_CARD = "#1a1a35"
BG_SURFACE = "#22223a"
FG_TEXT = "#e8e8f0"
FG_DIM = "#7788aa"
FG_ACCENT = "#00d4ff"
FG_GREEN = "#00e676"
FG_ORANGE = "#ffab40"
FG_RED = "#ff5252"
FG_YELLOW = "#ffd740"
BTN_BG = "#1a2744"
BTN_ACTIVE = "#243752"
ENTRY_BG = "#1a1a30"
OUTPUT_BG = "#0a0a14"
TAB_SEL = "#00796b"
TAB_UNSEL = "#1a1a30"
GEN_BTN = "#00796b"
GEN_BTN_HOVER = "#00897b"

# Vendor colors for status bar
VENDOR_COLORS = {
    "Nokia SR-OS": "#005AFF",
    "Nokia IXR-e2": "#FF6D00",
    "Ciena SAOS": "#7B1FA2",
}

# ===========================================================================
#  PREDEFINED VALUES
# ===========================================================================

MGMT_DESCRIPTIONS = [
    "MGMT",
    "SITE_BOSS",
    "POWER_PLANT",
    "SHARK_METER",
    "To-IXR-e-Port-1/1/20-s-Hook-to-Management-Port",
    "To-IXR-e-Port-1/1/24-s-Hook-to-Management-Port",
    "EDN-VZB_IDN-OneFiber-Uplink",
]

# ===========================================================================
#  TEST / VALIDATION TABLE
# ===========================================================================

def _test_table(vendor, mode, params):
    """Generate a test/validation table appended to every config output.
    Helps field engineers track what was verified before and after apply."""
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    lines = []
    lines.append("")
    lines.append("=" * 70)
    lines.append("  VALIDATION / TEST TABLE")
    lines.append("=" * 70)
    lines.append("# Generated: {}".format(ts))
    lines.append("# Vendor: {}".format(vendor))
    lines.append("# Mode: {}".format(mode))
    lines.append("#")
    lines.append("# Instructions: Execute each verification command BEFORE and AFTER")
    lines.append("# applying the config. Record PASS/FAIL and any notes.")
    lines.append("# This table is your change-control evidence.")
    lines.append("#")
    lines.append("# {:<4} {:<50} {:<6} {:<6} {}".format(
        "#", "Verification Command", "PRE", "POST", "Notes"))
    lines.append("# {}".format("-" * 80))

    checks = []
    if mode.startswith("B4A"):
        port = params.get("b4c_port", "?")
        vlan = params.get("b4c_vlan", "?")
        nni_port = params.get("nni_port", "?")
        nni_vlan = params.get("nni_vlan", "?")
        csr = params.get("b4c_csr", "?")
        nni_csr = params.get("nni_csr", "?")
        checks = [
            'show port {} (admin/oper state)'.format(port),
            'show router interface | match {}:{}'.format(port, vlan),
            'show router arp | match {}:{}'.format(port, vlan),
            'show router isis 5 adjacency (B4C side)'.format(),
            'show router bgp summary | match {}'.format(csr),
            'show port {} (NNI side)'.format(nni_port),
            'show router interface | match {}:{}'.format(nni_port, nni_vlan),
            'show router bgp summary | match {}'.format(nni_csr),
            'ping <remote P2P IP> (both directions)',
            'show router isis 5 adjacency (NNI side)',
        ]
    elif "mgmt" in mode.lower() or "management" in mode.lower():
        port = params.get("port", "?")
        checks = [
            'show port {} (admin/oper state)'.format(port),
            'show service id 400 sap (or 450 for IDN)',
            'show service id 400 base (VPLS status)',
            'ping <device IP> from VPRN 4',
            'show router 4 arp (verify gateway)',
        ]
    elif "6630" in mode:
        port = params.get("port", "?")
        checks = [
            'show port {} (admin/oper state)'.format(port),
            'show service id 100 sap (LTE S1)',
            'show service id 400 sap (OAM)',
            'show service id 310 sap (NR — if added)',
            'show service id 5000 sap (LTE CA — if added)',
            'show service id 5200 sap (NR CA — if added)',
            'ping <eNB OAM IP> from VPRN 4',
        ]
    elif "6648" in mode:
        checks = [
            'show port (all variant ports — admin/oper)',
            'show service id 100 sap (LTE S1)',
            'show service id 400 sap (OAM)',
            'show service id 310 sap (NR S1)',
            'show service id 5000 sap (LTE CA)',
            'show service id 5200 sap (NR CA)',
            'ping <eNB OAM IP> from VPRN 4',
        ]
    elif "6672" in mode:
        port = params.get("port", "?")
        checks = [
            'show port {} (admin/oper state)'.format(port),
            'show service id 100 sap (LTE S1)',
            'show service id 400 sap (OAM)',
            'show service id 310 sap (NR)',
            'show service id 5000 sap (LTE CA)',
            'show service id 5200 sap (NR CA)',
            'ping <eNB OAM IP> from VPRN 4',
        ]
    else:
        checks = ['show port (verify all ports)', 'show service sap-using',
                   'ping <target IP>']

    for i, cmd in enumerate(checks, 1):
        lines.append("# {:<4} {:<50} {:<6} {:<6} {}".format(
            i, cmd, "[  ]", "[  ]", ""))

    lines.append("#")
    lines.append("# Tested by: ___________________  Date: ___________")
    lines.append("# Approved by: _________________  Date: ___________")
    lines.append("# Ticket/CR #: _________________")
    lines.append("#")
    lines.append("# NOTE: If Verizon endpoint security blocks execution,")
    lines.append("# run the .py source directly: python auto_phil2.py")
    lines.append("# or use the web version if available.")
    lines.append("=" * 70)
    return "\n".join(lines)


# ===========================================================================
#  NOKIA SR-OS (CLASSIC CLI) — B4A BACKHAUL
# ===========================================================================

def _b4a_interface_name(port, vlan):
    return "INT-{}:{}-Base".format(port, vlan)


def _b4a_gen_port(port, hostname, bw_speed):
    p = "\\configure port {}".format(port)
    lines = [
        p,
        "info",
        "{}    shutdown".format(p),
        '{}    description "{}_{}"'.format(p, hostname, port),
        "{}    ethernet mode hybrid".format(p),
        "{}    ethernet encap-type dot1q".format(p),
        "{}    ethernet mtu 9104".format(p),
        "{}    ethernet crc-monitor sd-threshold 2 multiplier 5".format(p),
        "{}    ethernet crc-monitor sf-threshold 1".format(p),
        "{}    ethernet lldp dest-mac nearest-bridge admin-status tx-rx".format(p),
        "{}    ethernet lldp dest-mac nearest-bridge notification".format(p),
        "{}    ethernet lldp dest-mac nearest-bridge tx-tlvs port-desc sys-name sys-desc sys-cap".format(p),
        "{}    ethernet lldp dest-mac nearest-bridge port-id-subtype tx-if-name".format(p),
        "{}    ethernet util-stats-interval 30".format(p),
        '{}    ethernet egress-port-qos-policy "40012"'.format(p),
        "{}    ethernet speed {}".format(p, bw_speed),
        "{}    no shutdown".format(p),
        "info",
        "exit all",
        "",
        'show port {} | match expression "Description|Speed|Auto|Admin State|Oper State|Last|Phys State|dBm|Link Level|Model Number|Physical Link|Rate|Warn|Configured"'.format(port),
    ]
    return lines


def _b4a_gen_interface(port, vlan, hostname, p2p_ip, bw_speed, egress_instance):
    intf = _b4a_interface_name(port, vlan)
    rg = "QG-NE-{}M".format(bw_speed)
    b = '\\config router interface "{}"'.format(intf)
    lines = [
        b,
        "info",
        '{} address {}/31'.format(b, p2p_ip),
        '{} description "{}_{}"'.format(b, hostname, intf),
        '{} egress vlan-qos-policy "40011"'.format(b),
        '{} port {}:{}'.format(b, port, vlan),
        '{} ingress qos "40021"'.format(b),
        '{} egress egress-remark-policy "40021"'.format(b),
        '{} qos 40022 egress-port-redirect-group "{}" egress-instance {}'.format(b, rg, egress_instance),
        "{} bfd 50 receive 50 multiplier 5 type fp".format(b),
        "{} enable-ingress-stats".format(b),
        "{} no shutdown".format(b),
        "info",
        "exit all",
        "",
        "show router interface | match {}:{}".format(port, vlan),
    ]
    return lines


def _b4a_gen_nni_interface(nni_port, nni_vlan, b4c_host, b4c_port, b4c_vlan, nni_p2p, bw_speed):
    intf = _b4a_interface_name(nni_port, nni_vlan)
    rg = "QG-NE-{}M".format(bw_speed)
    b = '\\config router interface "{}"'.format(intf)
    lines = [
        b,
        "info",
        '{} address {}/31'.format(b, nni_p2p),
        '{} description "{}_{}"'.format(b, b4c_host, _b4a_interface_name(b4c_port, b4c_vlan)),
        '{} port {}:{}'.format(b, nni_port, nni_vlan),
        '{} qos 40022 egress-port-redirect-group "{}" egress-instance {}'.format(b, rg, nni_vlan),
        "{} bfd 50 receive 50 multiplier 5 type cpm-np".format(b),
        "{} enable-ingress-stats".format(b),
        "{} no shutdown".format(b),
        "info",
        "exit all",
        "",
        "show router interface | match {}:{}".format(nni_port, nni_vlan),
    ]
    return lines


def _b4a_gen_isis(port, vlan):
    intf = _b4a_interface_name(port, vlan)
    b = '\\config router isis 5 interface "{}"'.format(intf)
    return [
        b, "info",
        "{} level-capability level-1".format(b),
        "{} interface-type point-to-point".format(b),
        "{} bfd-enable ipv4".format(b),
        "{} level 1 metric 1000000".format(b),
        "{} no shutdown".format(b),
        b, "info", "exit all", "",
        'show router isis 5 interface "{}"'.format(intf),
    ]


def _b4a_gen_bgp(bgp_group, neighbor_ip, desc):
    b = '\\config router bgp group "{}" neighbor {}'.format(bgp_group, neighbor_ip)
    return [
        b,
        '{} description "iBGP-TO-{}"'.format(b, desc),
        '{} authentication-key "eNSEbgp"'.format(b),
        "exit all", "\\admin save", "",
        "show router bgp summary | match {} post-lines 3".format(neighbor_ip),
    ]


def _b4a_gen_verify(port, vlan, csr_ip):
    return [
        'show port {} | match expression "Description|Speed|Auto|Admin State|Oper State|Last|Phys State|dBm|Link Level|Model Number|Physical Link|Rate|Warn|Configured"'.format(port),
        "show router interface | match {}:{}".format(port, vlan),
        "show router arp | match {}:{}".format(port, vlan),
        "show router bgp summary | match {} post-lines 3".format(csr_ip),
    ]


def _b4a_gen_spoke_ref():
    return [
        "# ENSE-SPOKE: HUB site needs RR-5-ENSESR_SPOKE group with SPOKE neighbor.",
        "# Example:", "",
        '\\config router bgp group "RR-5-ENSESR_SPOKE" neighbor <SPOKE_CSR_IP>',
        '\\config router bgp group "RR-5-ENSESR_SPOKE" neighbor <SPOKE_CSR_IP> description "iBGP-TO-<SPOKE_HOSTNAME>"',
        '\\config router bgp group "RR-5-ENSESR_SPOKE" neighbor <SPOKE_CSR_IP> authentication-key "eNSEbgp"',
        "exit all", "\\admin save",
        "show router bgp summary | match <SPOKE_CSR_IP> post-lines 3",
    ]


# ===========================================================================
#  NOKIA SR-OS — B4C SERVICE PORT TEMPLATES
# ===========================================================================

def _nokia_mgmt(d):
    port = d["port"]
    desc = d["description"]
    mtu = d["mtu"]
    speed = d["speed"]
    autoneg = d["autoneg"]
    vlan_entry = ":{}".format(d["vlan"]) if d.get("vlan") else ""
    mgmt_type = d["mgmt_type"]

    lines = []
    lines.append("#ONLY FOR NOKIA B4C ENSE ROUTERS")
    lines.append("#################")
    lines.append("# LOGICAL  PORT")
    lines.append("#################")
    lines.append("\\configure")
    lines.append('    port {}'.format(port))
    lines.append('        shutdown')
    lines.append('        description "LINK-TO-OAM/EDN-{}"'.format(desc))
    lines.append('        ethernet')
    lines.append('            {}'.format(speed))
    lines.append('            mode access')
    lines.append('            encap-type dot1q')
    lines.append('            {}'.format(autoneg))
    lines.append('            mtu {}'.format(mtu))
    lines.append('        exit')
    lines.append('        no shutdown')
    lines.append('    exit')
    lines.append('')
    lines.append('#################')
    lines.append('# SERVICE SAPS')
    lines.append('#################')
    lines.append('\\configure')
    lines.append('    service')

    if mgmt_type == "edn":
        lines.append('        vpls 400')
        lines.append('            sap {}{} create'.format(port, vlan_entry))
        lines.append('                ingress')
        lines.append('                    qos 41031')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')
    else:
        lines.append('        vpls 450 customer 4 create')
        lines.append('            description "IXRs-IDN MGMT and Console and Server"')
        lines.append('            service-mtu 1528')
        lines.append('            allow-ip-int-bind')
        lines.append('            exit')
        lines.append('            stp')
        lines.append('                shutdown')
        lines.append('            exit')
        lines.append('            sap {}{} create'.format(port, vlan_entry))
        lines.append('                ingress')
        lines.append('                    qos 41031')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')

    if d.get("gateway_ip"):
        lines.append('')
        if mgmt_type == "edn":
            lines.append('#EDN IPv4 SUBNET')
            lines.append('\\configure')
            lines.append('    service')
            lines.append('        vprn 4')
            lines.append('            interface "INT-VPLS400-CELL_MGMT" create')
            lines.append('            address {}/29'.format(d["gateway_ip"]))
            lines.append('            vpls "VPLS400"')
            lines.append('            exit')
            lines.append('        exit')
        else:
            lines.append('#IDN IPv4 SUBNET')
            lines.append('\\configure')
            lines.append('    service')
            lines.append('        vprn 4')
            lines.append('            interface "INT-VPLS450-CELL_MGMT" create')
            lines.append('            address {}/27'.format(d["gateway_ip"]))
            lines.append('            vpls "VPLS450"')
            lines.append('            exit')
            lines.append('        exit')

    return "\n".join(lines)


def _nokia_6630(d):
    lte_v = "302" if d.get("alt_vlan") else "301"
    oam_v = "402" if d.get("alt_vlan") else "401"
    port = d["port"]
    desc_id = d["desc_id"]

    lines = []
    lines.append("#ONLY FOR NOKIA B4C ENSE ROUTERS")
    lines.append("#CHECK LIVE CONFIG BEFORE APPLYING")
    lines.append("#---------------------")
    lines.append("#    LOGICAL PORT")
    lines.append("#---------------------")
    lines.append("\\configure")
    lines.append('    port {}'.format(port))
    lines.append('        shutdown')
    lines.append('        description "eNB-{}-6630"'.format(desc_id))
    lines.append('        ethernet')
    lines.append('            {}'.format(d["speed"]))
    lines.append('            mode access')
    lines.append('            encap-type dot1q')
    lines.append('            {}'.format(d["autoneg"]))
    lines.append('            mtu 9104')
    lines.append('        exit')
    lines.append('        no shutdown')
    lines.append('    exit')
    lines.append('')
    lines.append('#---------------------')
    lines.append('#    SERVICE SAPS')
    lines.append('#---------------------')
    lines.append('\\configure')
    lines.append('    service')
    lines.append('        vpls 100')
    lines.append('            sap {}:{} create'.format(port, lte_v))
    lines.append('                ingress')
    lines.append('                    qos 41031')
    lines.append('                exit')
    lines.append('                egress')
    lines.append('                    agg-rate')
    lines.append('                        rate max cir max')
    lines.append('                    exit')
    lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
    lines.append('                exit')
    lines.append('                no shutdown')
    lines.append('            exit')
    lines.append('')
    lines.append('\\configure')
    lines.append('    service')
    lines.append('        vpls 400')
    lines.append('            sap {}:{} create'.format(port, oam_v))
    lines.append('                ingress')
    lines.append('                    qos 41031')
    lines.append('                exit')
    lines.append('                no shutdown')
    lines.append('            exit')

    if d.get("add_nr"):
        lines.append('')
        lines.append('#VPLS 310 SAPS')
        lines.append('\\configure')
        lines.append('    service')
        lines.append('        vpls 310')
        lines.append('            sap {}:310 create'.format(port))
        lines.append('                ingress')
        lines.append('                    qos 41031')
        lines.append('                exit')
        lines.append('                egress')
        lines.append('                    agg-rate')
        lines.append('                        rate max cir max')
        lines.append('                    exit')
        lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')

    if d.get("add_500"):
        lines.append('')
        lines.append('#LTE CARRIER AGG SAPS')
        lines.append('\\configure')
        lines.append('    service')
        lines.append('        vpls 5000')
        lines.append('            sap {}:500 create'.format(port))
        lines.append('                ingress')
        lines.append('                    qos 41032')
        lines.append('                exit')
        lines.append('                egress')
        lines.append('                    agg-rate')
        lines.append('                        rate max cir max')
        lines.append('                    exit')
        lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')

    if d.get("add_520") and d.get("ca520_port"):
        ca_port = d["ca520_port"]
        lines.append('')
        lines.append('#NR CA PORT CONFIG')
        lines.append('\\configure port 1/1/{}'.format(ca_port))
        lines.append('    shutdown')
        lines.append('    description "eNB-{}-6630-NR_CA"'.format(desc_id))
        lines.append('    ethernet')
        lines.append('        speed 10000')
        lines.append('        mode access')
        lines.append('        encap-type dot1q')
        lines.append('        mtu 9104')
        lines.append('        down-when-looped')
        lines.append('            no shutdown')
        lines.append('        exit')
        lines.append('        hold-time down 5')
        lines.append('    exit')
        lines.append('    no shutdown')
        lines.append('exit')
        lines.append('')
        lines.append('#NR CA SERVICE SAPS')
        lines.append('\\configure')
        lines.append('    service')
        lines.append('        vpls 5200')
        lines.append('            sap 1/1/{}:520 create'.format(ca_port))
        lines.append('                no shutdown')
        lines.append('            exit')

    return "\n".join(lines)


def _nokia_6648(d):
    s1_v = "302" if d.get("alt_vlan") else "301"
    oam_v = "402" if d.get("alt_vlan") else "401"
    desc_id = d["desc_id"]
    variant = d["variant"]

    lines = []

    if variant == "ls6":
        port_s1 = d["port_s1"]
        port_nr = d["port_nr"]
        port_nrca = d["port_nrca"]
        lines.append('#S1 LTE Port')
        lines.append('/configure port {}'.format(port_s1))
        lines.append('        shutdown')
        lines.append('        description "eNB-{}_6648_S1_LTE_OAM_LTE_CA"'.format(desc_id))
        lines.append('        ethernet')
        lines.append('            speed 10000')
        lines.append('            mode access')
        lines.append('            encap-type dot1q')
        lines.append('            mtu 9104')
        lines.append('            down-when-looped')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('            hold-time down 5')
        lines.append('        exit')
        lines.append('        no shutdown')
        lines.append('')
        lines.append('#S1 NR Port')
        lines.append('/configure port {}'.format(port_nr))
        lines.append('        shutdown')
        lines.append('        description "eNB-{}_6648_S1_NR"'.format(desc_id))
        lines.append('        ethernet')
        lines.append('            speed 10000')
        lines.append('            mode access')
        lines.append('            encap-type dot1q')
        lines.append('            mtu 9104')
        lines.append('            down-when-looped')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('            hold-time down 5')
        lines.append('            ssm')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('            util-stats-interval 30')
        lines.append('        exit')
        lines.append('        no shutdown')
        lines.append('')
        lines.append('#NR CA Port')
        lines.append('/configure port {}'.format(port_nrca))
        lines.append('        shutdown')
        lines.append('        description "eNB-{}_6648_NR_CA"'.format(desc_id))
        lines.append('        ethernet')
        lines.append('            speed 10000')
        lines.append('            mode access')
        lines.append('            encap-type dot1q')
        lines.append('            mtu 9212')
        lines.append('            down-when-looped')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('            hold-time down 5')
        lines.append('        exit')
        lines.append('        no shutdown')
        lines.append('')
        lines.append('# S1 LTE SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 100')
        lines.append('            sap {}:{} create'.format(port_s1, s1_v))
        lines.append('                ingress')
        lines.append('                    qos 41031')
        lines.append('                exit')
        lines.append('                egress')
        lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('        exit')
        lines.append('')
        lines.append('# OAM SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 400')
        lines.append('            sap {}:{} create'.format(port_s1, oam_v))
        lines.append('                ingress')
        lines.append('                    qos 41031')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('        exit')
        lines.append('')
        lines.append('# S1 NR SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 310')
        lines.append('            sap {}:310 create'.format(port_nr))
        lines.append('                ingress')
        lines.append('                    qos 41031')
        lines.append('                exit')
        lines.append('                egress')
        lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('        exit')
        lines.append('')
        lines.append('# LTE CA SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 5000')
        lines.append('            sap {}:500 create'.format(port_s1))
        lines.append('                ingress')
        lines.append('                    qos 41032')
        lines.append('                exit')
        lines.append('                egress')
        lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
        lines.append('                    agg-rate')
        lines.append('                        rate max cir max')
        lines.append('                    exit')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('')
        lines.append('# NR CA SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 5200')
        lines.append('            sap {}:520 create'.format(port_nrca))
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('        exit')

    elif variant == "dual":
        port_s1 = d["port_s1"]
        port_nrca = d["port_nrca"]
        lines.append('/configure port {}'.format(port_s1))
        lines.append('        shutdown')
        lines.append('        description "eNB-{}_6648_S1_LTE_S1_NR_OAM"'.format(desc_id))
        lines.append('        ethernet')
        lines.append('            speed 10000')
        lines.append('            mode access')
        lines.append('            encap-type dot1q')
        lines.append('            mtu 9104')
        lines.append('            down-when-looped')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('            hold-time down 5')
        lines.append('        exit')
        lines.append('        no shutdown')
        lines.append('')
        lines.append('/configure port {}'.format(port_nrca))
        lines.append('        shutdown')
        lines.append('        description "eNB-{}_6648_S1_NR"'.format(desc_id))
        lines.append('        ethernet')
        lines.append('            speed 10000')
        lines.append('            mode access')
        lines.append('            encap-type dot1q')
        lines.append('            mtu 9104')
        lines.append('            down-when-looped')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('            hold-time down 5')
        lines.append('            ssm')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('            util-stats-interval 30')
        lines.append('        exit')
        lines.append('        no shutdown')
        lines.append('')
        lines.append('# S1 LTE and NR SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 100')
        lines.append('            sap {}:{} create'.format(port_s1, s1_v))
        lines.append('                ingress')
        lines.append('                    qos 41031')
        lines.append('                exit')
        lines.append('                egress')
        lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('        exit')
        lines.append('')
        lines.append('# OAM SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 400')
        lines.append('            sap {}:{} create'.format(port_s1, oam_v))
        lines.append('                ingress')
        lines.append('                    qos 41031')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('        exit')
        lines.append('')
        lines.append('# S1 NR SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 310')
        lines.append('            sap {}:310 create'.format(port_s1))
        lines.append('                ingress')
        lines.append('                    qos 41031')
        lines.append('                exit')
        lines.append('                egress')
        lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('        exit')
        lines.append('')
        lines.append('# LTE CA SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 5000')
        lines.append('            sap {}:500 create'.format(port_s1))
        lines.append('                ingress')
        lines.append('                    qos 41032')
        lines.append('                exit')
        lines.append('                egress')
        lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
        lines.append('                    agg-rate')
        lines.append('                        rate max cir max')
        lines.append('                    exit')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('')
        lines.append('# NR CA SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 5200')
        lines.append('            sap {}:520 create'.format(port_nrca))
        lines.append('                no shutdown')
        lines.append('            exit')

    elif variant == "tri":
        port_s1 = d["port_s1"]
        port_nr = d["port_nr"]
        lines.append('/configure port {}'.format(port_s1))
        lines.append('        shutdown')
        lines.append('        description "eNB-{}_6648_S1_LTE_OAM_LTE_CA"'.format(desc_id))
        lines.append('        ethernet')
        lines.append('            speed 10000')
        lines.append('            mode access')
        lines.append('            encap-type dot1q')
        lines.append('            mtu 9104')
        lines.append('            down-when-looped')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('            hold-time down 5')
        lines.append('        exit')
        lines.append('        no shutdown')
        lines.append('')
        lines.append('/configure port {}'.format(port_nr))
        lines.append('        shutdown')
        lines.append('        description "eNB-{}_6648_S1_NR"'.format(desc_id))
        lines.append('        ethernet')
        lines.append('            speed 10000')
        lines.append('            mode access')
        lines.append('            encap-type dot1q')
        lines.append('            mtu 9104')
        lines.append('            down-when-looped')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('            hold-time down 5')
        lines.append('            ssm')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('            util-stats-interval 30')
        lines.append('        exit')
        lines.append('        no shutdown')
        lines.append('')
        lines.append('# S1 LTE SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 100')
        lines.append('            sap {}:{} create'.format(port_s1, s1_v))
        lines.append('                ingress')
        lines.append('                    qos 41031')
        lines.append('                exit')
        lines.append('                egress')
        lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('        exit')
        lines.append('')
        lines.append('# OAM SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 400')
        lines.append('            sap {}:{} create'.format(port_s1, oam_v))
        lines.append('                ingress')
        lines.append('                    qos 41031')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('        exit')
        lines.append('')
        lines.append('# S1 NR SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 310')
        lines.append('            sap {}:310 create'.format(port_nr))
        lines.append('                ingress')
        lines.append('                    qos 41031')
        lines.append('                exit')
        lines.append('                egress')
        lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('        exit')
        lines.append('')
        lines.append('# LTE CA SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 5000')
        lines.append('            sap {}:500 create'.format(port_s1))
        lines.append('                ingress')
        lines.append('                    qos 41032')
        lines.append('                exit')
        lines.append('                egress')
        lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
        lines.append('                    agg-rate')
        lines.append('                        rate max cir max')
        lines.append('                    exit')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')

    else:  # lte_lb
        port_s1 = d["port_s1"]
        port_nrca = d["port_nrca"]
        lines.append('/configure port {}'.format(port_s1))
        lines.append('        shutdown')
        lines.append('        description "eNB-{}_6648_S1_LTE_OAM_LTE_CA"'.format(desc_id))
        lines.append('        ethernet')
        lines.append('            speed 10000')
        lines.append('            mode access')
        lines.append('            encap-type dot1q')
        lines.append('            mtu 9104')
        lines.append('            down-when-looped')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('            hold-time down 5')
        lines.append('        exit')
        lines.append('        no shutdown')
        lines.append('')
        lines.append('/configure port {}'.format(port_nrca))
        lines.append('        shutdown')
        lines.append('        description "eNB-{}_NR_CA"'.format(desc_id))
        lines.append('        ethernet')
        lines.append('            speed 10000')
        lines.append('            mode access')
        lines.append('            encap-type dot1q')
        lines.append('            mtu 9212')
        lines.append('            down-when-looped')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('            hold-time down 5')
        lines.append('        exit')
        lines.append('        no shutdown')
        lines.append('')
        lines.append('# S1 LTE SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 100')
        lines.append('            sap {}:{} create'.format(port_s1, s1_v))
        lines.append('                ingress')
        lines.append('                    qos 41031')
        lines.append('                exit')
        lines.append('                egress')
        lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('        exit')
        lines.append('')
        lines.append('# OAM SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 400')
        lines.append('            sap {}:{} create'.format(port_s1, oam_v))
        lines.append('                ingress')
        lines.append('                    qos 41031')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('        exit')
        lines.append('')
        lines.append('# LTE CA SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 5000')
        lines.append('            sap {}:500 create'.format(port_s1))
        lines.append('                ingress')
        lines.append('                    qos 41032')
        lines.append('                exit')
        lines.append('                egress')
        lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
        lines.append('                    agg-rate')
        lines.append('                        rate max cir max')
        lines.append('                    exit')
        lines.append('                exit')
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('')
        lines.append('# NR CA SAP')
        lines.append('/configure')
        lines.append('    service')
        lines.append('        vpls 5200')
        lines.append('            sap {}:5200 create'.format(port_nrca))
        lines.append('                no shutdown')
        lines.append('            exit')
        lines.append('        exit')

    return "\n".join(lines)


def _nokia_6672(d):
    port = d["port"]
    desc_id = d["desc_id"]
    speed = "speed 25000" if d.get("use_25g") else "speed 10000"

    lines = []
    lines.append('# 6672 PORT CONFIG')
    lines.append('/configure port {}'.format(port))
    lines.append('        shutdown')
    lines.append('        description "eNB-{}_6672_S1_LTE_S1_NR_OAM_LTE_NR_CA"'.format(desc_id))
    lines.append('        ethernet')
    lines.append('            {}'.format(speed))
    lines.append('            mode access')
    lines.append('            encap-type dot1q')
    lines.append('            mtu 9212')
    lines.append('            down-when-looped')
    lines.append('                no shutdown')
    lines.append('            exit')
    lines.append('            hold-time down 5')
    lines.append('        exit')
    lines.append('        no shutdown')
    lines.append('')
    lines.append('# S1 SAP')
    lines.append('/configure')
    lines.append('    service')
    lines.append('        vpls 100')
    lines.append('            sap {}:301 create'.format(port))
    lines.append('                ingress')
    lines.append('                    qos 41031')
    lines.append('                exit')
    lines.append('                egress')
    lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
    lines.append('                exit')
    lines.append('                no shutdown')
    lines.append('            exit')
    lines.append('        exit')
    lines.append('')
    lines.append('# OAM SAP')
    lines.append('/configure')
    lines.append('    service')
    lines.append('        vpls 400')
    lines.append('            sap {}:401 create'.format(port))
    lines.append('                ingress')
    lines.append('                    qos 41031')
    lines.append('                exit')
    lines.append('                no shutdown')
    lines.append('            exit')
    lines.append('        exit')
    lines.append('')
    lines.append('# NR SAP')
    lines.append('/configure')
    lines.append('    service')
    lines.append('        vpls 310')
    lines.append('            sap {}:310 create'.format(port))
    lines.append('                ingress')
    lines.append('                    qos 41031')
    lines.append('                exit')
    lines.append('                egress')
    lines.append('                    vlan-qos-policy "RAN_Downstream_Scheduler"')
    lines.append('                exit')
    lines.append('                no shutdown')
    lines.append('            exit')
    lines.append('        exit')
    lines.append('')
    lines.append('# LTE CA SAP')
    lines.append('/configure')
    lines.append('    service')
    lines.append('        vpls 5000')
    lines.append('            sap {}:500 create'.format(port))
    lines.append('                no shutdown')
    lines.append('            exit')
    lines.append('        exit')
    lines.append('')
    lines.append('# NR CA SAP')
    lines.append('/configure')
    lines.append('    service')
    lines.append('        vpls 5200')
    lines.append('            sap {}:520 create'.format(port))
    lines.append('                no shutdown')
    lines.append('            exit')
    lines.append('        exit')

    return "\n".join(lines)


# ===========================================================================
#  NOKIA IXR-e2 (SR LINUX MD-CLI) — B4C SERVICE PORTS
# ===========================================================================
# IXR-e2 uses SR Linux flat-set or candidate-mode CLI.
# Port naming: ethernet-1/{port} (single-slot chassis)
# Services: network-instance type mac-vrf for L2 VPLS equivalent
# QoS: qos-profile based

def _ixre2_port_name(user_port):
    """Convert user input port (e.g. '1/1/7' or '7') to SR Linux interface name."""
    parts = user_port.strip().replace(" ", "").split("/")
    if len(parts) >= 3:
        return "ethernet-1/{}".format(parts[2])
    elif len(parts) == 1:
        return "ethernet-1/{}".format(parts[0])
    return "ethernet-1/{}".format(parts[-1])


def _ixre2_mgmt(d):
    port = d["port"]
    ixr_port = _ixre2_port_name(port)
    desc = d["description"]
    mtu = d["mtu"]
    vlan = d.get("vlan", "")
    mgmt_type = d["mgmt_type"]
    vpls_id = "450" if mgmt_type == "idn" else "400"
    vpls_name = "VPLS-{}".format(vpls_id)

    speed_map = {"speed 100": "100M", "speed 1000": "1G", "speed 10000": "10G"}
    ixr_speed = speed_map.get(d["speed"], "1G")

    lines = []
    lines.append('#NOKIA IXR-e2 (SR Linux) — MANAGEMENT PORT')
    lines.append('#Vendor: Nokia IXR-e2 | Mode: {}'.format(mgmt_type.upper()))
    lines.append('')
    lines.append('enter candidate')
    lines.append('')
    lines.append('# --- Interface ---')
    lines.append('set / interface {} admin-state disable'.format(ixr_port))
    lines.append('set / interface {} description "LINK-TO-OAM/EDN-{}"'.format(ixr_port, desc))
    lines.append('set / interface {} ethernet port-speed {}'.format(ixr_port, ixr_speed))
    lines.append('set / interface {} mtu {}'.format(ixr_port, mtu))
    if vlan:
        lines.append('')
        lines.append('# --- Subinterface (VLAN {}) ---'.format(vlan))
        lines.append('set / interface {} subinterface {} type bridged'.format(ixr_port, vlan))
        lines.append('set / interface {} subinterface {} vlan encap single-tagged vlan-id {}'.format(ixr_port, vlan, vlan))
    else:
        lines.append('')
        lines.append('# --- Subinterface (untagged) ---')
        lines.append('set / interface {} subinterface 0 type bridged'.format(ixr_port))
    lines.append('')
    lines.append('set / interface {} admin-state enable'.format(ixr_port))
    lines.append('')
    lines.append('# --- Network Instance ({}) ---'.format(vpls_name))
    lines.append('set / network-instance {} type mac-vrf'.format(vpls_name))
    sub_idx = vlan if vlan else "0"
    lines.append('set / network-instance {} interface {}.{}'.format(vpls_name, ixr_port, sub_idx))

    if d.get("gateway_ip"):
        lines.append('')
        lines.append('# --- IRB Gateway ---')
        prefix = "/29" if mgmt_type == "edn" else "/27"
        lines.append('set / interface irb0 subinterface {} ipv4 admin-state enable'.format(vpls_id))
        lines.append('set / interface irb0 subinterface {} ipv4 address {}{}'.format(vpls_id, d["gateway_ip"], prefix))
        lines.append('set / network-instance {} interface irb0.{}'.format(vpls_name, vpls_id))

    lines.append('')
    lines.append('commit now')
    lines.append('')
    lines.append('# --- Verification ---')
    lines.append('show interface {} detail'.format(ixr_port))
    lines.append('show network-instance {} interfaces'.format(vpls_name))

    return "\n".join(lines)


def _ixre2_6630(d):
    port = d["port"]
    ixr_port = _ixre2_port_name(port)
    desc_id = d["desc_id"]
    lte_v = "302" if d.get("alt_vlan") else "301"
    oam_v = "402" if d.get("alt_vlan") else "401"

    lines = []
    lines.append('#NOKIA IXR-e2 (SR Linux) — eNB/gNB 6630')
    lines.append('#CHECK LIVE CONFIG BEFORE APPLYING')
    lines.append('')
    lines.append('enter candidate')
    lines.append('')
    lines.append('# --- Port Config ---')
    lines.append('set / interface {} admin-state disable'.format(ixr_port))
    lines.append('set / interface {} description "eNB-{}-6630"'.format(ixr_port, desc_id))
    lines.append('set / interface {} ethernet port-speed 10G'.format(ixr_port))
    lines.append('set / interface {} mtu 9104'.format(ixr_port))
    lines.append('')
    lines.append('# --- LTE S1 Subinterface (VLAN {}) ---'.format(lte_v))
    lines.append('set / interface {} subinterface {} type bridged'.format(ixr_port, lte_v))
    lines.append('set / interface {} subinterface {} vlan encap single-tagged vlan-id {}'.format(ixr_port, lte_v, lte_v))
    lines.append('')
    lines.append('# --- OAM Subinterface (VLAN {}) ---'.format(oam_v))
    lines.append('set / interface {} subinterface {} type bridged'.format(ixr_port, oam_v))
    lines.append('set / interface {} subinterface {} vlan encap single-tagged vlan-id {}'.format(ixr_port, oam_v, oam_v))
    lines.append('')
    lines.append('set / interface {} admin-state enable'.format(ixr_port))
    lines.append('')
    lines.append('# --- Service Bindings ---')
    lines.append('set / network-instance VPLS-100 type mac-vrf')
    lines.append('set / network-instance VPLS-100 interface {}.{}'.format(ixr_port, lte_v))
    lines.append('set / network-instance VPLS-400 type mac-vrf')
    lines.append('set / network-instance VPLS-400 interface {}.{}'.format(ixr_port, oam_v))

    if d.get("add_nr"):
        lines.append('')
        lines.append('# --- NR S1 (VLAN 310) ---')
        lines.append('set / interface {} subinterface 310 type bridged'.format(ixr_port))
        lines.append('set / interface {} subinterface 310 vlan encap single-tagged vlan-id 310'.format(ixr_port))
        lines.append('set / network-instance VPLS-310 type mac-vrf')
        lines.append('set / network-instance VPLS-310 interface {}.310'.format(ixr_port))

    if d.get("add_500"):
        lines.append('')
        lines.append('# --- LTE CA (VLAN 500) ---')
        lines.append('set / interface {} subinterface 500 type bridged'.format(ixr_port))
        lines.append('set / interface {} subinterface 500 vlan encap single-tagged vlan-id 500'.format(ixr_port))
        lines.append('set / network-instance VPLS-5000 type mac-vrf')
        lines.append('set / network-instance VPLS-5000 interface {}.500'.format(ixr_port))

    if d.get("add_520") and d.get("ca520_port"):
        ca_ixr = _ixre2_port_name("1/1/{}".format(d["ca520_port"]))
        lines.append('')
        lines.append('# --- NR CA Port + VLAN 520 ---')
        lines.append('set / interface {} admin-state disable'.format(ca_ixr))
        lines.append('set / interface {} description "eNB-{}-6630-NR_CA"'.format(ca_ixr, desc_id))
        lines.append('set / interface {} ethernet port-speed 10G'.format(ca_ixr))
        lines.append('set / interface {} mtu 9104'.format(ca_ixr))
        lines.append('set / interface {} subinterface 520 type bridged'.format(ca_ixr))
        lines.append('set / interface {} subinterface 520 vlan encap single-tagged vlan-id 520'.format(ca_ixr))
        lines.append('set / interface {} admin-state enable'.format(ca_ixr))
        lines.append('set / network-instance VPLS-5200 type mac-vrf')
        lines.append('set / network-instance VPLS-5200 interface {}.520'.format(ca_ixr))

    lines.append('')
    lines.append('commit now')
    lines.append('')
    lines.append('# --- Verification ---')
    lines.append('show interface {}'.format(ixr_port))
    lines.append('show network-instance VPLS-100 interfaces')
    lines.append('show network-instance VPLS-400 interfaces')

    return "\n".join(lines)


def _ixre2_6648(d):
    desc_id = d["desc_id"]
    variant = d["variant"]
    s1_v = "302" if d.get("alt_vlan") else "301"
    oam_v = "402" if d.get("alt_vlan") else "401"

    lines = []
    lines.append('#NOKIA IXR-e2 (SR Linux) — eNB/gNB 6648 {}'.format(variant.upper()))
    lines.append('')
    lines.append('enter candidate')
    lines.append('')

    # Determine ports based on variant
    port_s1 = _ixre2_port_name(d["port_s1"]) if d.get("port_s1") else None
    port_nr = _ixre2_port_name(d["port_nr"]) if d.get("port_nr") else None
    port_nrca = _ixre2_port_name(d["port_nrca"]) if d.get("port_nrca") else None

    if variant == "ls6":
        lines.append('# --- S1 LTE Port ---')
        lines.append('set / interface {} admin-state disable'.format(port_s1))
        lines.append('set / interface {} description "eNB-{}_6648_S1_LTE_OAM_LTE_CA"'.format(port_s1, desc_id))
        lines.append('set / interface {} ethernet port-speed 10G'.format(port_s1))
        lines.append('set / interface {} mtu 9104'.format(port_s1))
        lines.append('set / interface {} subinterface {} type bridged'.format(port_s1, s1_v))
        lines.append('set / interface {} subinterface {} vlan encap single-tagged vlan-id {}'.format(port_s1, s1_v, s1_v))
        lines.append('set / interface {} subinterface {} type bridged'.format(port_s1, oam_v))
        lines.append('set / interface {} subinterface {} vlan encap single-tagged vlan-id {}'.format(port_s1, oam_v, oam_v))
        lines.append('set / interface {} subinterface 500 type bridged'.format(port_s1))
        lines.append('set / interface {} subinterface 500 vlan encap single-tagged vlan-id 500'.format(port_s1))
        lines.append('set / interface {} admin-state enable'.format(port_s1))
        lines.append('')
        lines.append('# --- S1 NR Port ---')
        lines.append('set / interface {} admin-state disable'.format(port_nr))
        lines.append('set / interface {} description "eNB-{}_6648_S1_NR"'.format(port_nr, desc_id))
        lines.append('set / interface {} ethernet port-speed 10G'.format(port_nr))
        lines.append('set / interface {} mtu 9104'.format(port_nr))
        lines.append('set / interface {} subinterface 310 type bridged'.format(port_nr))
        lines.append('set / interface {} subinterface 310 vlan encap single-tagged vlan-id 310'.format(port_nr))
        lines.append('set / interface {} admin-state enable'.format(port_nr))
        lines.append('')
        lines.append('# --- NR CA Port ---')
        lines.append('set / interface {} admin-state disable'.format(port_nrca))
        lines.append('set / interface {} description "eNB-{}_6648_NR_CA"'.format(port_nrca, desc_id))
        lines.append('set / interface {} ethernet port-speed 10G'.format(port_nrca))
        lines.append('set / interface {} mtu 9212'.format(port_nrca))
        lines.append('set / interface {} subinterface 520 type bridged'.format(port_nrca))
        lines.append('set / interface {} subinterface 520 vlan encap single-tagged vlan-id 520'.format(port_nrca))
        lines.append('set / interface {} admin-state enable'.format(port_nrca))
        lines.append('')
        lines.append('# --- Service Bindings ---')
        lines.append('set / network-instance VPLS-100 interface {}.{}'.format(port_s1, s1_v))
        lines.append('set / network-instance VPLS-400 interface {}.{}'.format(port_s1, oam_v))
        lines.append('set / network-instance VPLS-310 interface {}.310'.format(port_nr))
        lines.append('set / network-instance VPLS-5000 interface {}.500'.format(port_s1))
        lines.append('set / network-instance VPLS-5200 interface {}.520'.format(port_nrca))

    elif variant == "dual":
        lines.append('# --- S1 LTE+NR Port ---')
        lines.append('set / interface {} admin-state disable'.format(port_s1))
        lines.append('set / interface {} description "eNB-{}_6648_S1_LTE_S1_NR_OAM"'.format(port_s1, desc_id))
        lines.append('set / interface {} ethernet port-speed 10G'.format(port_s1))
        lines.append('set / interface {} mtu 9104'.format(port_s1))
        lines.append('set / interface {} subinterface {} type bridged'.format(port_s1, s1_v))
        lines.append('set / interface {} subinterface {} vlan encap single-tagged vlan-id {}'.format(port_s1, s1_v, s1_v))
        lines.append('set / interface {} subinterface {} type bridged'.format(port_s1, oam_v))
        lines.append('set / interface {} subinterface {} vlan encap single-tagged vlan-id {}'.format(port_s1, oam_v, oam_v))
        lines.append('set / interface {} subinterface 310 type bridged'.format(port_s1))
        lines.append('set / interface {} subinterface 310 vlan encap single-tagged vlan-id 310'.format(port_s1))
        lines.append('set / interface {} subinterface 500 type bridged'.format(port_s1))
        lines.append('set / interface {} subinterface 500 vlan encap single-tagged vlan-id 500'.format(port_s1))
        lines.append('set / interface {} admin-state enable'.format(port_s1))
        lines.append('')
        lines.append('# --- NR CA Port ---')
        lines.append('set / interface {} admin-state disable'.format(port_nrca))
        lines.append('set / interface {} description "eNB-{}_6648_NR_CA"'.format(port_nrca, desc_id))
        lines.append('set / interface {} ethernet port-speed 10G'.format(port_nrca))
        lines.append('set / interface {} mtu 9104'.format(port_nrca))
        lines.append('set / interface {} subinterface 520 type bridged'.format(port_nrca))
        lines.append('set / interface {} subinterface 520 vlan encap single-tagged vlan-id 520'.format(port_nrca))
        lines.append('set / interface {} admin-state enable'.format(port_nrca))
        lines.append('')
        lines.append('# --- Service Bindings ---')
        lines.append('set / network-instance VPLS-100 interface {}.{}'.format(port_s1, s1_v))
        lines.append('set / network-instance VPLS-400 interface {}.{}'.format(port_s1, oam_v))
        lines.append('set / network-instance VPLS-310 interface {}.310'.format(port_s1))
        lines.append('set / network-instance VPLS-5000 interface {}.500'.format(port_s1))
        lines.append('set / network-instance VPLS-5200 interface {}.520'.format(port_nrca))

    elif variant == "tri":
        lines.append('# --- S1 LTE Port ---')
        lines.append('set / interface {} admin-state disable'.format(port_s1))
        lines.append('set / interface {} description "eNB-{}_6648_S1_LTE_OAM_LTE_CA"'.format(port_s1, desc_id))
        lines.append('set / interface {} ethernet port-speed 10G'.format(port_s1))
        lines.append('set / interface {} mtu 9104'.format(port_s1))
        lines.append('set / interface {} subinterface {} type bridged'.format(port_s1, s1_v))
        lines.append('set / interface {} subinterface {} vlan encap single-tagged vlan-id {}'.format(port_s1, s1_v, s1_v))
        lines.append('set / interface {} subinterface {} type bridged'.format(port_s1, oam_v))
        lines.append('set / interface {} subinterface {} vlan encap single-tagged vlan-id {}'.format(port_s1, oam_v, oam_v))
        lines.append('set / interface {} subinterface 500 type bridged'.format(port_s1))
        lines.append('set / interface {} subinterface 500 vlan encap single-tagged vlan-id 500'.format(port_s1))
        lines.append('set / interface {} admin-state enable'.format(port_s1))
        lines.append('')
        lines.append('# --- S1 NR Port ---')
        lines.append('set / interface {} admin-state disable'.format(port_nr))
        lines.append('set / interface {} description "eNB-{}_6648_S1_NR"'.format(port_nr, desc_id))
        lines.append('set / interface {} ethernet port-speed 10G'.format(port_nr))
        lines.append('set / interface {} mtu 9104'.format(port_nr))
        lines.append('set / interface {} subinterface 310 type bridged'.format(port_nr))
        lines.append('set / interface {} subinterface 310 vlan encap single-tagged vlan-id 310'.format(port_nr))
        lines.append('set / interface {} admin-state enable'.format(port_nr))
        lines.append('')
        lines.append('# --- Service Bindings ---')
        lines.append('set / network-instance VPLS-100 interface {}.{}'.format(port_s1, s1_v))
        lines.append('set / network-instance VPLS-400 interface {}.{}'.format(port_s1, oam_v))
        lines.append('set / network-instance VPLS-310 interface {}.310'.format(port_nr))
        lines.append('set / network-instance VPLS-5000 interface {}.500'.format(port_s1))

    else:  # lte_lb
        lines.append('# --- S1 LTE Port ---')
        lines.append('set / interface {} admin-state disable'.format(port_s1))
        lines.append('set / interface {} description "eNB-{}_6648_S1_LTE_OAM_LTE_CA"'.format(port_s1, desc_id))
        lines.append('set / interface {} ethernet port-speed 10G'.format(port_s1))
        lines.append('set / interface {} mtu 9104'.format(port_s1))
        lines.append('set / interface {} subinterface {} type bridged'.format(port_s1, s1_v))
        lines.append('set / interface {} subinterface {} vlan encap single-tagged vlan-id {}'.format(port_s1, s1_v, s1_v))
        lines.append('set / interface {} subinterface {} type bridged'.format(port_s1, oam_v))
        lines.append('set / interface {} subinterface {} vlan encap single-tagged vlan-id {}'.format(port_s1, oam_v, oam_v))
        lines.append('set / interface {} subinterface 500 type bridged'.format(port_s1))
        lines.append('set / interface {} subinterface 500 vlan encap single-tagged vlan-id 500'.format(port_s1))
        lines.append('set / interface {} admin-state enable'.format(port_s1))
        lines.append('')
        lines.append('# --- NR CA Port ---')
        lines.append('set / interface {} admin-state disable'.format(port_nrca))
        lines.append('set / interface {} description "eNB-{}_NR_CA"'.format(port_nrca, desc_id))
        lines.append('set / interface {} ethernet port-speed 10G'.format(port_nrca))
        lines.append('set / interface {} mtu 9212'.format(port_nrca))
        lines.append('set / interface {} subinterface 520 type bridged'.format(port_nrca))
        lines.append('set / interface {} subinterface 520 vlan encap single-tagged vlan-id 520'.format(port_nrca))
        lines.append('set / interface {} admin-state enable'.format(port_nrca))
        lines.append('')
        lines.append('# --- Service Bindings ---')
        lines.append('set / network-instance VPLS-100 interface {}.{}'.format(port_s1, s1_v))
        lines.append('set / network-instance VPLS-400 interface {}.{}'.format(port_s1, oam_v))
        lines.append('set / network-instance VPLS-5000 interface {}.500'.format(port_s1))
        lines.append('set / network-instance VPLS-5200 interface {}.520'.format(port_nrca))

    lines.append('')
    lines.append('commit now')
    lines.append('')
    lines.append('# --- Verification ---')
    lines.append('show interface {} detail'.format(port_s1 or "?"))
    lines.append('show network-instance summary')

    return "\n".join(lines)


def _ixre2_6672(d):
    port = d["port"]
    ixr_port = _ixre2_port_name(port)
    desc_id = d["desc_id"]
    speed = "25G" if d.get("use_25g") else "10G"

    lines = []
    lines.append('#NOKIA IXR-e2 (SR Linux) — eNB/gNB 6672')
    lines.append('')
    lines.append('enter candidate')
    lines.append('')
    lines.append('# --- Port Config ---')
    lines.append('set / interface {} admin-state disable'.format(ixr_port))
    lines.append('set / interface {} description "eNB-{}_6672_S1_LTE_S1_NR_OAM_LTE_NR_CA"'.format(ixr_port, desc_id))
    lines.append('set / interface {} ethernet port-speed {}'.format(ixr_port, speed))
    lines.append('set / interface {} mtu 9212'.format(ixr_port))
    lines.append('')
    lines.append('# --- Subinterfaces ---')
    lines.append('set / interface {} subinterface 301 type bridged'.format(ixr_port))
    lines.append('set / interface {} subinterface 301 vlan encap single-tagged vlan-id 301'.format(ixr_port))
    lines.append('set / interface {} subinterface 401 type bridged'.format(ixr_port))
    lines.append('set / interface {} subinterface 401 vlan encap single-tagged vlan-id 401'.format(ixr_port))
    lines.append('set / interface {} subinterface 310 type bridged'.format(ixr_port))
    lines.append('set / interface {} subinterface 310 vlan encap single-tagged vlan-id 310'.format(ixr_port))
    lines.append('set / interface {} subinterface 500 type bridged'.format(ixr_port))
    lines.append('set / interface {} subinterface 500 vlan encap single-tagged vlan-id 500'.format(ixr_port))
    lines.append('set / interface {} subinterface 520 type bridged'.format(ixr_port))
    lines.append('set / interface {} subinterface 520 vlan encap single-tagged vlan-id 520'.format(ixr_port))
    lines.append('')
    lines.append('set / interface {} admin-state enable'.format(ixr_port))
    lines.append('')
    lines.append('# --- Service Bindings ---')
    lines.append('set / network-instance VPLS-100 interface {}.301'.format(ixr_port))
    lines.append('set / network-instance VPLS-400 interface {}.401'.format(ixr_port))
    lines.append('set / network-instance VPLS-310 interface {}.310'.format(ixr_port))
    lines.append('set / network-instance VPLS-5000 interface {}.500'.format(ixr_port))
    lines.append('set / network-instance VPLS-5200 interface {}.520'.format(ixr_port))
    lines.append('')
    lines.append('commit now')
    lines.append('')
    lines.append('# --- Verification ---')
    lines.append('show interface {} detail'.format(ixr_port))
    lines.append('show network-instance summary')

    return "\n".join(lines)


# ===========================================================================
#  CIENA SAOS (10.x) — B4C SERVICE PORTS
# ===========================================================================
# Ciena SAOS uses: port, sub-port, virtual-switch, traffic-profiling
# Port naming: user-provided (e.g. 1/1, 1/7)

def _ciena_mgmt(d):
    port = d["port"]
    desc = d["description"]
    mtu = d["mtu"]
    vlan = d.get("vlan", "")
    mgmt_type = d["mgmt_type"]
    vpls_id = "450" if mgmt_type == "idn" else "400"
    vs_name = "VPLS-{}".format(vpls_id)

    speed_map = {"speed 100": "hundred", "speed 1000": "one-gig", "speed 10000": "ten-gig"}
    ciena_speed = speed_map.get(d["speed"], "one-gig")

    lines = []
    lines.append('#CIENA SAOS — MANAGEMENT PORT')
    lines.append('#Vendor: Ciena SAOS 10.x | Mode: {}'.format(mgmt_type.upper()))
    lines.append('')
    lines.append('# --- Port Config ---')
    lines.append('port set port {} admin-state disabled'.format(port))
    lines.append('port set port {} description "LINK-TO-OAM/EDN-{}"'.format(port, desc))
    lines.append('port set port {} speed {}'.format(port, ciena_speed))
    lines.append('port set port {} max-frame-size {}'.format(port, mtu))
    lines.append('port set port {} admin-state enabled'.format(port))
    if vlan:
        lines.append('')
        lines.append('# --- Sub-Port (VLAN {}) ---'.format(vlan))
        lines.append('sub-port create sub-port {}.{} parent-port {} classifier-precedence 100 ingress-l2-transform pop'.format(port, vlan, port))
        lines.append('sub-port set sub-port {}.{} admin-state enabled'.format(port, vlan))
        lines.append('')
        lines.append('# --- Virtual Switch ({}) ---'.format(vs_name))
        lines.append('virtual-switch ethernet create vs {} vc-type vlan'.format(vs_name))
        lines.append('virtual-switch interface attach sub-port {}.{} vs {}'.format(port, vlan, vs_name))
    else:
        lines.append('')
        lines.append('# --- Sub-Port (untagged) ---')
        lines.append('sub-port create sub-port {}.0 parent-port {} classifier-precedence 100'.format(port, port))
        lines.append('sub-port set sub-port {}.0 admin-state enabled'.format(port))
        lines.append('')
        lines.append('# --- Virtual Switch ({}) ---'.format(vs_name))
        lines.append('virtual-switch ethernet create vs {} vc-type vlan'.format(vs_name))
        lines.append('virtual-switch interface attach sub-port {}.0 vs {}'.format(port, vs_name))

    if d.get("gateway_ip"):
        prefix = "/29" if mgmt_type == "edn" else "/27"
        lines.append('')
        lines.append('# --- L3 Interface (Gateway) ---')
        lines.append('interface create interface CELL_MGMT_{} type ip'.format(vpls_id))
        lines.append('interface set interface CELL_MGMT_{} ip-address {}{}'.format(vpls_id, d["gateway_ip"], prefix))
        lines.append('interface attach interface CELL_MGMT_{} vs {}'.format(vpls_id, vs_name))

    lines.append('')
    lines.append('# --- Verification ---')
    lines.append('port show port {}'.format(port))
    lines.append('sub-port show sub-port {}'.format(port))
    lines.append('virtual-switch ethernet show vs {}'.format(vs_name))

    return "\n".join(lines)


def _ciena_6630(d):
    port = d["port"]
    desc_id = d["desc_id"]
    lte_v = "302" if d.get("alt_vlan") else "301"
    oam_v = "402" if d.get("alt_vlan") else "401"

    lines = []
    lines.append('#CIENA SAOS — eNB/gNB 6630')
    lines.append('#CHECK LIVE CONFIG BEFORE APPLYING')
    lines.append('')
    lines.append('# --- Port Config ---')
    lines.append('port set port {} admin-state disabled'.format(port))
    lines.append('port set port {} description "eNB-{}-6630"'.format(port, desc_id))
    lines.append('port set port {} speed ten-gig'.format(port))
    lines.append('port set port {} max-frame-size 9104'.format(port))
    lines.append('port set port {} admin-state enabled'.format(port))
    lines.append('')
    lines.append('# --- LTE S1 Sub-Port (VLAN {}) ---'.format(lte_v))
    lines.append('sub-port create sub-port {}.{} parent-port {} classifier-precedence 100 ingress-l2-transform pop'.format(port, lte_v, port))
    lines.append('sub-port set sub-port {}.{} admin-state enabled'.format(port, lte_v))
    lines.append('virtual-switch interface attach sub-port {}.{} vs VPLS-100'.format(port, lte_v))
    lines.append('')
    lines.append('# --- OAM Sub-Port (VLAN {}) ---'.format(oam_v))
    lines.append('sub-port create sub-port {}.{} parent-port {} classifier-precedence 100 ingress-l2-transform pop'.format(port, oam_v, port))
    lines.append('sub-port set sub-port {}.{} admin-state enabled'.format(port, oam_v))
    lines.append('virtual-switch interface attach sub-port {}.{} vs VPLS-400'.format(port, oam_v))

    if d.get("add_nr"):
        lines.append('')
        lines.append('# --- NR S1 (VLAN 310) ---')
        lines.append('sub-port create sub-port {}.310 parent-port {} classifier-precedence 100 ingress-l2-transform pop'.format(port, port))
        lines.append('sub-port set sub-port {}.310 admin-state enabled'.format(port))
        lines.append('virtual-switch interface attach sub-port {}.310 vs VPLS-310'.format(port))

    if d.get("add_500"):
        lines.append('')
        lines.append('# --- LTE CA (VLAN 500) ---')
        lines.append('sub-port create sub-port {}.500 parent-port {} classifier-precedence 100 ingress-l2-transform pop'.format(port, port))
        lines.append('sub-port set sub-port {}.500 admin-state enabled'.format(port))
        lines.append('virtual-switch interface attach sub-port {}.500 vs VPLS-5000'.format(port))

    if d.get("add_520") and d.get("ca520_port"):
        ca_port = "1/1/{}".format(d["ca520_port"])
        lines.append('')
        lines.append('# --- NR CA Port + VLAN 520 ---')
        lines.append('port set port {} admin-state disabled'.format(ca_port))
        lines.append('port set port {} description "eNB-{}-6630-NR_CA"'.format(ca_port, desc_id))
        lines.append('port set port {} speed ten-gig'.format(ca_port))
        lines.append('port set port {} max-frame-size 9104'.format(ca_port))
        lines.append('port set port {} admin-state enabled'.format(ca_port))
        lines.append('sub-port create sub-port {}.520 parent-port {} classifier-precedence 100 ingress-l2-transform pop'.format(ca_port, ca_port))
        lines.append('sub-port set sub-port {}.520 admin-state enabled'.format(ca_port))
        lines.append('virtual-switch interface attach sub-port {}.520 vs VPLS-5200'.format(ca_port))

    lines.append('')
    lines.append('# --- Verification ---')
    lines.append('port show port {}'.format(port))
    lines.append('sub-port show sub-port {}'.format(port))
    lines.append('virtual-switch ethernet show vs VPLS-100')

    return "\n".join(lines)


def _ciena_6648(d):
    desc_id = d["desc_id"]
    variant = d["variant"]
    s1_v = "302" if d.get("alt_vlan") else "301"
    oam_v = "402" if d.get("alt_vlan") else "401"

    lines = []
    lines.append('#CIENA SAOS — eNB/gNB 6648 {}'.format(variant.upper()))
    lines.append('')

    def _port_block(port_id, desc, speed="ten-gig", mtu="9104"):
        b = []
        b.append('port set port {} admin-state disabled'.format(port_id))
        b.append('port set port {} description "{}"'.format(port_id, desc))
        b.append('port set port {} speed {}'.format(port_id, speed))
        b.append('port set port {} max-frame-size {}'.format(port_id, mtu))
        b.append('port set port {} admin-state enabled'.format(port_id))
        return b

    def _sub_attach(port_id, vlan, vs):
        return [
            'sub-port create sub-port {}.{} parent-port {} classifier-precedence 100 ingress-l2-transform pop'.format(port_id, vlan, port_id),
            'sub-port set sub-port {}.{} admin-state enabled'.format(port_id, vlan),
            'virtual-switch interface attach sub-port {}.{} vs {}'.format(port_id, vlan, vs),
        ]

    port_s1 = d.get("port_s1", "")
    port_nr = d.get("port_nr", "")
    port_nrca = d.get("port_nrca", "")

    if variant == "ls6":
        lines.append('# --- S1 LTE Port ---')
        lines.extend(_port_block(port_s1, "eNB-{}_6648_S1_LTE_OAM_LTE_CA".format(desc_id)))
        lines.append('')
        lines.append('# --- S1 NR Port ---')
        lines.extend(_port_block(port_nr, "eNB-{}_6648_S1_NR".format(desc_id)))
        lines.append('')
        lines.append('# --- NR CA Port ---')
        lines.extend(_port_block(port_nrca, "eNB-{}_6648_NR_CA".format(desc_id), mtu="9212"))
        lines.append('')
        lines.append('# --- Service Bindings ---')
        lines.extend(_sub_attach(port_s1, s1_v, "VPLS-100"))
        lines.extend(_sub_attach(port_s1, oam_v, "VPLS-400"))
        lines.extend(_sub_attach(port_nr, "310", "VPLS-310"))
        lines.extend(_sub_attach(port_s1, "500", "VPLS-5000"))
        lines.extend(_sub_attach(port_nrca, "520", "VPLS-5200"))

    elif variant == "dual":
        lines.append('# --- S1 LTE+NR Port ---')
        lines.extend(_port_block(port_s1, "eNB-{}_6648_S1_LTE_S1_NR_OAM".format(desc_id)))
        lines.append('')
        lines.append('# --- NR CA Port ---')
        lines.extend(_port_block(port_nrca, "eNB-{}_6648_NR_CA".format(desc_id)))
        lines.append('')
        lines.append('# --- Service Bindings ---')
        lines.extend(_sub_attach(port_s1, s1_v, "VPLS-100"))
        lines.extend(_sub_attach(port_s1, oam_v, "VPLS-400"))
        lines.extend(_sub_attach(port_s1, "310", "VPLS-310"))
        lines.extend(_sub_attach(port_s1, "500", "VPLS-5000"))
        lines.extend(_sub_attach(port_nrca, "520", "VPLS-5200"))

    elif variant == "tri":
        lines.append('# --- S1 LTE Port ---')
        lines.extend(_port_block(port_s1, "eNB-{}_6648_S1_LTE_OAM_LTE_CA".format(desc_id)))
        lines.append('')
        lines.append('# --- S1 NR Port ---')
        lines.extend(_port_block(port_nr, "eNB-{}_6648_S1_NR".format(desc_id)))
        lines.append('')
        lines.append('# --- Service Bindings ---')
        lines.extend(_sub_attach(port_s1, s1_v, "VPLS-100"))
        lines.extend(_sub_attach(port_s1, oam_v, "VPLS-400"))
        lines.extend(_sub_attach(port_nr, "310", "VPLS-310"))
        lines.extend(_sub_attach(port_s1, "500", "VPLS-5000"))

    else:  # lte_lb
        lines.append('# --- S1 LTE Port ---')
        lines.extend(_port_block(port_s1, "eNB-{}_6648_S1_LTE_OAM_LTE_CA".format(desc_id)))
        lines.append('')
        lines.append('# --- NR CA Port ---')
        lines.extend(_port_block(port_nrca, "eNB-{}_NR_CA".format(desc_id), mtu="9212"))
        lines.append('')
        lines.append('# --- Service Bindings ---')
        lines.extend(_sub_attach(port_s1, s1_v, "VPLS-100"))
        lines.extend(_sub_attach(port_s1, oam_v, "VPLS-400"))
        lines.extend(_sub_attach(port_s1, "500", "VPLS-5000"))
        lines.extend(_sub_attach(port_nrca, "520", "VPLS-5200"))

    lines.append('')
    lines.append('# --- Verification ---')
    lines.append('port show port {}'.format(port_s1))
    lines.append('virtual-switch ethernet show vs VPLS-100')

    return "\n".join(lines)


def _ciena_6672(d):
    port = d["port"]
    desc_id = d["desc_id"]
    speed = "twenty-five-gig" if d.get("use_25g") else "ten-gig"

    lines = []
    lines.append('#CIENA SAOS — eNB/gNB 6672')
    lines.append('')
    lines.append('# --- Port Config ---')
    lines.append('port set port {} admin-state disabled'.format(port))
    lines.append('port set port {} description "eNB-{}_6672_S1_LTE_S1_NR_OAM_LTE_NR_CA"'.format(port, desc_id))
    lines.append('port set port {} speed {}'.format(port, speed))
    lines.append('port set port {} max-frame-size 9212'.format(port))
    lines.append('port set port {} admin-state enabled'.format(port))
    lines.append('')
    lines.append('# --- Sub-Ports + Service Bindings ---')
    for vlan, vs in [("301", "VPLS-100"), ("401", "VPLS-400"), ("310", "VPLS-310"),
                     ("500", "VPLS-5000"), ("520", "VPLS-5200")]:
        lines.append('sub-port create sub-port {}.{} parent-port {} classifier-precedence 100 ingress-l2-transform pop'.format(port, vlan, port))
        lines.append('sub-port set sub-port {}.{} admin-state enabled'.format(port, vlan))
        lines.append('virtual-switch interface attach sub-port {}.{} vs {}'.format(port, vlan, vs))
    lines.append('')
    lines.append('# --- Verification ---')
    lines.append('port show port {}'.format(port))
    lines.append('sub-port show sub-port {}'.format(port))
    lines.append('virtual-switch ethernet show')

    return "\n".join(lines)


# ===========================================================================
#  VENDOR DISPATCH
# ===========================================================================

VENDOR_B4C = {
    "Nokia SR-OS": {
        "mgmt": _nokia_mgmt,
        "6630": _nokia_6630,
        "6648": _nokia_6648,
        "6672": _nokia_6672,
    },
    "Nokia IXR-e2": {
        "mgmt": _ixre2_mgmt,
        "6630": _ixre2_6630,
        "6648": _ixre2_6648,
        "6672": _ixre2_6672,
    },
    "Ciena SAOS": {
        "mgmt": _ciena_mgmt,
        "6630": _ciena_6630,
        "6648": _ciena_6648,
        "6672": _ciena_6672,
    },
}


# ===========================================================================
#  HIGH-LEVEL GENERATION FUNCTIONS
# ===========================================================================

def generate_b4a(d):
    """Build full B4A config. Always Nokia SR-OS (both sides are 7750 SR)."""
    spd = 1000 if d["bw"] == "1G" else 10000
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    sections = []

    hdr = [
        "# Generated: {}".format(ts),
        "# B4C: {}".format(d["b4c_host"]),
    ]
    if d.get("mgmt_ip"):
        hdr.append("# Management IP: {}".format(d["mgmt_ip"]))
    if d.get("b4a_host"):
        hdr.append("# B4A: {}".format(d["b4a_host"]))
    hdr.extend([
        "# Bandwidth: {} (speed {})".format(d["bw"], spd),
        "# B4C Port: {}  VLAN: {}  P2P: {}  CSR: {}".format(
            d["b4c_port"], d["b4c_vlan"], d["b4c_p2p"], d["b4c_csr"]),
        "# NNI Port: {}  VLAN: {}  P2P: {}  CSR: {}".format(
            d["nni_port"], d["nni_vlan"], d["nni_p2p"], d["nni_csr"]),
        "# BGP Group: {}".format(d["bgp"]),
        "# Vendor: Nokia SR-OS (B4A backhaul is always Nokia)",
    ])
    sections.append(("CONFIG SUMMARY", hdr))

    sections.append(("B4C SIDE - PORT CONFIG",
                      _b4a_gen_port(d["b4c_port"], d["b4c_host"], spd)))
    sections.append(("B4C SIDE - INTERFACE CONFIG",
                      _b4a_gen_interface(d["b4c_port"], d["b4c_vlan"], d["b4c_host"],
                                         d["b4c_p2p"], spd, d["b4c_vlan"])))
    sections.append(("B4C SIDE - ISIS 5 CONFIG",
                      _b4a_gen_isis(d["b4c_port"], d["b4c_vlan"])))
    sections.append(("B4C SIDE - VERIFY BGP GROUP", [
        "\\config router bgp",
        "info | match RR-5-ENSESR_CSR context all", "",
        "# If nothing is returned, STOP and verify the group exists."]))
    sections.append(("B4C SIDE - BGP NEIGHBOR CONFIG",
                      _b4a_gen_bgp(d["bgp"], d["b4c_csr"], d["b4c_host"])))
    sections.append(("B4C SIDE - VERIFY ALL",
                      _b4a_gen_verify(d["b4c_port"], d["b4c_vlan"], d["b4c_csr"])))

    sections.append(("NNI SIDE - INTERFACE CONFIG",
                      _b4a_gen_nni_interface(d["nni_port"], d["nni_vlan"], d["b4c_host"],
                                              d["b4c_port"], d["b4c_vlan"], d["nni_p2p"], spd)))
    sections.append(("NNI SIDE - ISIS 5 CONFIG",
                      _b4a_gen_isis(d["nni_port"], d["nni_vlan"])))
    sections.append(("NNI SIDE - BGP NEIGHBOR CONFIG",
                      _b4a_gen_bgp(d["bgp"], d["nni_csr"], d["b4c_host"])))
    sections.append(("NNI SIDE - VERIFY ALL",
                      _b4a_gen_verify(d["nni_port"], d["nni_vlan"], d["nni_csr"])))

    if d["bgp"] == "RR-5-ENSESR_CSR":
        sections.append(("REFERENCE - ENSE SPOKE (if needed)", _b4a_gen_spoke_ref()))

    out = []
    out.append("AutoPhil 2 v{} - B4A Backhaul | Nokia SR-OS".format(VERSION))
    out.append("=" * 70)
    for title, body in sections:
        out.append("")
        out.append("=" * 70)
        out.append("  {}".format(title))
        out.append("=" * 70)
        for line in body:
            out.append(line)

    result = "\n".join(out)
    result += _test_table("Nokia SR-OS", "B4A Backhaul", d)
    return result


def generate_b4c(vendor, mode, d):
    """Generate B4C config for given vendor and mode."""
    gen_fn = VENDOR_B4C.get(vendor, {}).get(mode)
    if not gen_fn:
        return "# ERROR: No template for vendor='{}' mode='{}'".format(vendor, mode)

    body = gen_fn(d)

    mode_labels = {
        "mgmt": "Management Port - {}".format(d.get("mgmt_type", "?").upper()),
        "6630": "eNB/gNB 6630 - {}".format(d.get("desc_id", "?")),
        "6648": "eNB/gNB 6648 {} - {}".format(d.get("variant", "?").upper(), d.get("desc_id", "?")),
        "6672": "eNB/gNB 6672 - {}".format(d.get("desc_id", "?")),
    }
    title = mode_labels.get(mode, mode)

    header = "AutoPhil 2 v{} - B4C Service | {}\n{}\n  {}\n{}".format(
        VERSION, vendor, "=" * 70, title, "=" * 70)

    result = header + "\n" + body
    result += _test_table(vendor, title, d)
    return result


# ===========================================================================
#  GUI APPLICATION
# ===========================================================================

class AutoPhil2App:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoPhil 2 — ENSE Config Generator v{}".format(VERSION))
        self.root.geometry("1200x900")
        self.root.minsize(1000, 750)
        self.root.configure(bg=BG_DARK)

        self.style = ttk.Style()
        self.style.theme_use("clam")
        self._configure_styles()

        # Main container
        main = tk.Frame(root, bg=BG_DARK)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))

        # Banner
        banner = tk.Frame(main, bg=BG_DARK)
        banner.pack(fill=tk.X, pady=(0, 8))
        tk.Label(banner, text="AutoPhil 2",
                 font=("Consolas", 20, "bold"), fg=FG_ACCENT, bg=BG_DARK).pack(side=tk.LEFT)
        tk.Label(banner, text="ENSE Config Generator v{}".format(VERSION),
                 font=("Consolas", 11), fg=FG_DIM, bg=BG_DARK).pack(side=tk.LEFT, padx=(12, 0))

        # Vendor selector (top-right, prominent)
        vendor_frame = tk.Frame(banner, bg=BG_DARK)
        vendor_frame.pack(side=tk.RIGHT)
        tk.Label(vendor_frame, text="VENDOR:", font=("Consolas", 9, "bold"),
                 fg=FG_DIM, bg=BG_DARK).pack(side=tk.LEFT, padx=(0, 6))
        self.vendor = tk.StringVar(value="Nokia SR-OS")
        self._vendor_buttons = []
        for v in ["Nokia SR-OS", "Nokia IXR-e2", "Ciena SAOS"]:
            btn = tk.Button(vendor_frame, text=v, font=("Consolas", 9, "bold"),
                            relief="flat", padx=10, pady=3, cursor="hand2",
                            command=lambda vv=v: self._select_vendor(vv))
            btn.pack(side=tk.LEFT, padx=2)
            self._vendor_buttons.append((btn, v))
        self._update_vendor_buttons()

        # Notebook (tabs)
        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.b4a_frame = tk.Frame(self.notebook, bg=BG_DARK)
        self.notebook.add(self.b4a_frame, text="  B4A Backhaul Link  ")
        self._build_b4a_tab()

        self.b4c_frame = tk.Frame(self.notebook, bg=BG_DARK)
        self.notebook.add(self.b4c_frame, text="  B4C Service Ports  ")
        self._build_b4c_tab()

        # Status bar
        self.status_frame = tk.Frame(root, bg=BG_MID, pady=5, padx=10)
        self.status_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.status_label = tk.Label(self.status_frame,
                                     text="VPLS: 100=LTE S1  310=NR S1  400=OAM  450=IDN  5000=LTE CA  5200=NR CA",
                                     font=("Consolas", 9), fg=FG_DIM, bg=BG_MID, anchor="w")
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.vendor_indicator = tk.Label(self.status_frame, text="Nokia SR-OS",
                                         font=("Consolas", 9, "bold"),
                                         fg="white", bg=VENDOR_COLORS["Nokia SR-OS"],
                                         padx=8, pady=2)
        self.vendor_indicator.pack(side=tk.RIGHT)

    def _select_vendor(self, v):
        self.vendor.set(v)
        self._update_vendor_buttons()
        color = VENDOR_COLORS.get(v, FG_ACCENT)
        self.vendor_indicator.configure(text=v, bg=color)
        self._set_status("Vendor: {}".format(v))

    def _update_vendor_buttons(self):
        cur = self.vendor.get()
        for btn, val in self._vendor_buttons:
            if val == cur:
                color = VENDOR_COLORS.get(val, TAB_SEL)
                btn.configure(bg=color, fg="white", activebackground=color, activeforeground="white")
            else:
                btn.configure(bg=TAB_UNSEL, fg=FG_DIM, activebackground=BTN_ACTIVE, activeforeground=FG_TEXT)

    def _set_status(self, msg, color=FG_GREEN):
        self.status_label.configure(text=msg, fg=color)
        self.root.after(5000, lambda: self.status_label.configure(
            text="VPLS: 100=LTE S1  310=NR S1  400=OAM  450=IDN  5000=LTE CA  5200=NR CA",
            fg=FG_DIM))

    def _configure_styles(self):
        s = self.style
        s.configure("TNotebook", background=BG_DARK, borderwidth=0)
        s.configure("TNotebook.Tab", background=BG_MID, foreground=FG_TEXT,
                     padding=(15, 8), font=("Consolas", 11, "bold"))
        s.map("TNotebook.Tab",
              background=[("selected", BG_CARD)],
              foreground=[("selected", FG_ACCENT)])
        s.configure("TFrame", background=BG_DARK)
        s.configure("TLabel", background=BG_DARK, foreground=FG_TEXT, font=("Consolas", 10))
        s.configure("TButton", background=BTN_BG, foreground=FG_TEXT,
                     font=("Consolas", 10, "bold"), padding=(12, 6))
        s.map("TButton", background=[("active", BTN_ACTIVE)])
        s.configure("TCombobox", fieldbackground=ENTRY_BG, background=BTN_BG,
                     foreground=FG_TEXT, font=("Consolas", 10),
                     selectbackground=ENTRY_BG, selectforeground=FG_TEXT,
                     arrowcolor=FG_ACCENT)
        s.map("TCombobox",
              fieldbackground=[("readonly", ENTRY_BG), ("disabled", BG_MID)],
              foreground=[("readonly", FG_TEXT), ("disabled", FG_DIM)],
              selectbackground=[("readonly", ENTRY_BG)],
              selectforeground=[("readonly", FG_TEXT)],
              background=[("readonly", BTN_BG)])
        # Fix combobox dropdown listbox colors on Windows
        self.root.option_add("*TCombobox*Listbox.background", ENTRY_BG)
        self.root.option_add("*TCombobox*Listbox.foreground", FG_TEXT)
        self.root.option_add("*TCombobox*Listbox.selectBackground", TAB_SEL)
        self.root.option_add("*TCombobox*Listbox.selectForeground", "white")
        s.configure("TCheckbutton", background=BG_DARK, foreground=FG_TEXT,
                     font=("Consolas", 10))
        s.map("TCheckbutton", background=[("active", BG_DARK)])

    # --- Widget helpers ---

    def _make_entry(self, parent, row, label, default="", width=30):
        tk.Label(parent, text=label, font=("Consolas", 10), fg=FG_TEXT,
                 bg=BG_DARK, anchor="e").grid(row=row, column=0, sticky="e", padx=(0, 8), pady=3)
        var = tk.StringVar(value=default)
        entry = tk.Entry(parent, textvariable=var, font=("Consolas", 10),
                         bg=ENTRY_BG, fg=FG_TEXT, insertbackground=FG_ACCENT,
                         relief="flat", width=width, highlightthickness=1,
                         highlightcolor=FG_ACCENT, highlightbackground=BG_MID)
        entry.grid(row=row, column=1, sticky="w", pady=3)
        return var

    def _make_combo(self, parent, row, label, values, default=None):
        tk.Label(parent, text=label, font=("Consolas", 10), fg=FG_TEXT,
                 bg=BG_DARK, anchor="e").grid(row=row, column=0, sticky="e", padx=(0, 8), pady=3)
        var = tk.StringVar(value=default if default else values[0])
        combo = ttk.Combobox(parent, textvariable=var, values=values, state="readonly",
                              font=("Consolas", 10), width=28)
        combo.grid(row=row, column=1, sticky="w", pady=3)
        return var

    def _make_check(self, parent, row, label, default=False):
        var = tk.BooleanVar(value=default)
        cb = tk.Checkbutton(parent, text=label, variable=var, font=("Consolas", 10),
                            fg=FG_TEXT, bg=BG_DARK, selectcolor=BG_MID,
                            activebackground=BG_DARK, activeforeground=FG_ACCENT)
        cb.grid(row=row, column=0, columnspan=2, sticky="w", padx=(10, 0), pady=3)
        return var

    def _make_output_area(self, parent):
        out_frame = tk.Frame(parent, bg=BG_DARK)

        btn_bar = tk.Frame(out_frame, bg=BG_DARK)
        btn_bar.pack(fill=tk.X, pady=(0, 4))

        text = tk.Text(out_frame, font=("Consolas", 10), bg=OUTPUT_BG, fg=FG_GREEN,
                       insertbackground=FG_ACCENT, relief="flat", wrap=tk.NONE,
                       highlightthickness=1, highlightcolor=FG_ACCENT,
                       highlightbackground=BG_MID)
        text.pack(fill=tk.BOTH, expand=True)

        scroll_y = tk.Scrollbar(text, command=text.yview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        text.configure(yscrollcommand=scroll_y.set)
        scroll_x = tk.Scrollbar(text, command=text.xview, orient=tk.HORIZONTAL)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        text.configure(xscrollcommand=scroll_x.set)

        def copy_output():
            content = text.get("1.0", tk.END).strip()
            if content:
                self.root.clipboard_clear()
                self.root.clipboard_append(content)
                self._set_status("Copied to clipboard!", FG_GREEN)

        def save_output():
            content = text.get("1.0", tk.END).strip()
            if not content:
                self._set_status("Nothing to save.", FG_ORANGE)
                return
            path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                initialfile="ense_config_{}.txt".format(
                    datetime.datetime.now().strftime("%Y%m%d_%H%M")))
            if path:
                with open(path, "w") as f:
                    f.write(content)
                self._set_status("Saved: {}".format(os.path.basename(path)), FG_GREEN)

        def clear_output():
            text.delete("1.0", tk.END)
            self._set_status("Output cleared.", FG_DIM)

        tk.Button(btn_bar, text="Copy to Clipboard", command=copy_output,
                  font=("Consolas", 9, "bold"), bg=BTN_BG, fg=FG_ACCENT,
                  activebackground=BTN_ACTIVE, activeforeground=FG_ACCENT,
                  relief="flat", padx=10, pady=4).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(btn_bar, text="Save to File", command=save_output,
                  font=("Consolas", 9, "bold"), bg=BTN_BG, fg=FG_GREEN,
                  activebackground=BTN_ACTIVE, activeforeground=FG_GREEN,
                  relief="flat", padx=10, pady=4).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(btn_bar, text="Clear", command=clear_output,
                  font=("Consolas", 9, "bold"), bg=BTN_BG, fg=FG_ORANGE,
                  activebackground=BTN_ACTIVE, activeforeground=FG_ORANGE,
                  relief="flat", padx=10, pady=4).pack(side=tk.LEFT)

        return out_frame, text

    # -----------------------------------------------------------------------
    #  B4A TAB
    # -----------------------------------------------------------------------
    def _build_b4a_tab(self):
        # Info label — B4A is always Nokia
        info = tk.Frame(self.b4a_frame, bg=BG_SURFACE, pady=4, padx=10)
        info.pack(fill=tk.X, pady=(0, 5))
        tk.Label(info, text="B4A Backhaul is always Nokia SR-OS (both sides are 7750 SR). Vendor selector does not apply here.",
                 font=("Consolas", 9), fg=FG_YELLOW, bg=BG_SURFACE).pack(anchor="w")

        paned = tk.PanedWindow(self.b4a_frame, orient=tk.HORIZONTAL, bg=BG_DARK,
                               sashwidth=6, sashrelief="flat")
        paned.pack(fill=tk.BOTH, expand=True, pady=5)

        left_outer = tk.Frame(paned, bg=BG_DARK)
        paned.add(left_outer, minsize=380)

        canvas = tk.Canvas(left_outer, bg=BG_DARK, highlightthickness=0)
        scrollbar = tk.Scrollbar(left_outer, orient="vertical", command=canvas.yview)
        form_container = tk.Frame(canvas, bg=BG_DARK)
        form_container.bind("<Configure>",
                            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=form_container, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        form = form_container
        r = 0

        # B4C Side — field order matches Brian's Excel yellow cells exactly
        tk.Label(form, text="B4C Side (CSR)", font=("Consolas", 11, "bold"),
                 fg=FG_ACCENT, bg=BG_DARK).grid(row=r, column=0, columnspan=2, sticky="w", pady=(5, 5))
        r += 1
        self.b4a_b4c_host = self._make_entry(form, r, "Host Name"); r += 1
        self.b4a_mgmt_ip = self._make_entry(form, r, "Management IP"); r += 1
        self.b4a_b4c_port = self._make_entry(form, r, "NNI/B4A/B4C Port", "1/1/1"); r += 1
        self.b4a_b4c_vlan = self._make_entry(form, r, "ODD VLAN", "1001"); r += 1
        self.b4a_bw = self._make_combo(form, r, "Bandwidth", ["10G", "1G"], "10G"); r += 1
        self.b4a_b4c_p2p = self._make_entry(form, r, "B4C Point to Point IPv4"); r += 1
        self.b4a_b4c_csr = self._make_entry(form, r, "B4C CSR GLOBAL LoO"); r += 1

        r += 1
        tk.Label(form, text="B4A / NNI Side", font=("Consolas", 11, "bold"),
                 fg=FG_ACCENT, bg=BG_DARK).grid(row=r, column=0, columnspan=2, sticky="w", pady=(10, 5))
        r += 1
        self.b4a_nni_port = self._make_entry(form, r, "NNI Port (B4A)", "1/1/c8/1"); r += 1
        self.b4a_nni_vlan = self._make_entry(form, r, "EVEN VLAN"); r += 1
        self.b4a_nni_p2p = self._make_entry(form, r, "NNI Point to Point IPv4"); r += 1
        self.b4a_nni_csr = self._make_entry(form, r, "NNI CSR GLOBAL LoO"); r += 1

        r += 1
        tk.Label(form, text="BGP", font=("Consolas", 11, "bold"),
                 fg=FG_ACCENT, bg=BG_DARK).grid(row=r, column=0, columnspan=2, sticky="w", pady=(10, 5))
        r += 1
        self.b4a_bgp = self._make_combo(form, r, "BGP Group",
                                         ["RR-5-ENSESR_CSR", "RR-5-ENSESR_SPOKE"],
                                         "RR-5-ENSESR_CSR"); r += 1

        r += 1
        tk.Label(form, text="Info Only (optional)", font=("Consolas", 11, "bold"),
                 fg=FG_DIM, bg=BG_DARK).grid(row=r, column=0, columnspan=2, sticky="w", pady=(10, 5))
        r += 1
        self.b4a_b4a_host = self._make_entry(form, r, "B4A Host Name", ""); r += 1

        r += 1
        tk.Button(form, text="GENERATE B4A CONFIG", command=self._gen_b4a,
                  font=("Consolas", 12, "bold"), bg=GEN_BTN, fg="white",
                  activebackground=GEN_BTN_HOVER, activeforeground="white",
                  relief="flat", padx=20, pady=8, cursor="hand2"
                  ).grid(row=r, column=0, columnspan=2, pady=15)

        right_frame = tk.Frame(paned, bg=BG_DARK)
        paned.add(right_frame, minsize=400)
        tk.Label(right_frame, text="OUTPUT", font=("Consolas", 11, "bold"),
                 fg=FG_ORANGE, bg=BG_DARK).pack(anchor="w", pady=(0, 4))
        out_frame, self.b4a_output = self._make_output_area(right_frame)
        out_frame.pack(fill=tk.BOTH, expand=True)

    def _gen_b4a(self):
        required = {
            "B4C Host Name": self.b4a_b4c_host,
            "B4C Port": self.b4a_b4c_port,
            "B4C VLAN": self.b4a_b4c_vlan,
            "B4C P2P IP": self.b4a_b4c_p2p,
            "B4C CSR": self.b4a_b4c_csr,
            "NNI Port": self.b4a_nni_port,
            "NNI VLAN": self.b4a_nni_vlan,
            "NNI P2P IP": self.b4a_nni_p2p,
            "NNI CSR": self.b4a_nni_csr,
        }
        for name, var in required.items():
            if not var.get().strip():
                self._set_status("Missing: {}".format(name), FG_RED)
                return

        d = {
            "b4c_host": self.b4a_b4c_host.get().strip(),
            "mgmt_ip": self.b4a_mgmt_ip.get().strip(),
            "b4c_port": self.b4a_b4c_port.get().strip(),
            "b4c_vlan": self.b4a_b4c_vlan.get().strip(),
            "bw": self.b4a_bw.get().strip(),
            "b4c_p2p": self.b4a_b4c_p2p.get().strip(),
            "b4c_csr": self.b4a_b4c_csr.get().strip(),
            "nni_port": self.b4a_nni_port.get().strip(),
            "nni_vlan": self.b4a_nni_vlan.get().strip(),
            "nni_p2p": self.b4a_nni_p2p.get().strip(),
            "nni_csr": self.b4a_nni_csr.get().strip(),
            "bgp": self.b4a_bgp.get().strip(),
            "b4a_host": self.b4a_b4a_host.get().strip(),
        }
        try:
            result = generate_b4a(d)
            self.b4a_output.delete("1.0", tk.END)
            self.b4a_output.insert("1.0", result)
            self._set_status("B4A config generated ({} lines)".format(result.count("\n") + 1), FG_GREEN)
        except Exception as e:
            self._set_status("Error: {}".format(e), FG_RED)

    # -----------------------------------------------------------------------
    #  B4C TAB
    # -----------------------------------------------------------------------
    def _build_b4c_tab(self):
        paned = tk.PanedWindow(self.b4c_frame, orient=tk.HORIZONTAL, bg=BG_DARK,
                               sashwidth=6, sashrelief="flat")
        paned.pack(fill=tk.BOTH, expand=True, pady=5)

        left_outer = tk.Frame(paned, bg=BG_DARK)
        paned.add(left_outer, minsize=420)

        canvas = tk.Canvas(left_outer, bg=BG_DARK, highlightthickness=0)
        scrollbar = tk.Scrollbar(left_outer, orient="vertical", command=canvas.yview)
        self.b4c_form_container = tk.Frame(canvas, bg=BG_DARK)
        self.b4c_form_container.bind("<Configure>",
                                     lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.b4c_form_container, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        mode_frame = tk.Frame(self.b4c_form_container, bg=BG_DARK)
        mode_frame.grid(row=0, column=0, columnspan=2, sticky="w", pady=(5, 10))

        self.b4c_mode = tk.StringVar(value="6672")
        self._mode_buttons = []
        modes = [("Management Port", "mgmt"),
                 ("eNB/gNB Add", "6630"),
                 ("6648 / 6651", "6648"),
                 ("6672 / 6355", "6672")]
        for text, val in modes:
            btn = tk.Button(mode_frame, text=text, font=("Consolas", 10, "bold"),
                            relief="flat", padx=12, pady=5, cursor="hand2",
                            command=lambda v=val: self._select_b4c_mode(v))
            btn.pack(side=tk.LEFT, padx=2)
            self._mode_buttons.append((btn, val))
        self._update_mode_buttons()

        self.b4c_dynamic = tk.Frame(self.b4c_form_container, bg=BG_DARK)
        self.b4c_dynamic.grid(row=1, column=0, columnspan=2, sticky="nsew")

        self._refresh_b4c_form()

        right_frame = tk.Frame(paned, bg=BG_DARK)
        paned.add(right_frame, minsize=400)
        tk.Label(right_frame, text="OUTPUT", font=("Consolas", 11, "bold"),
                 fg=FG_ORANGE, bg=BG_DARK).pack(anchor="w", pady=(0, 4))
        out_frame, self.b4c_output = self._make_output_area(right_frame)
        out_frame.pack(fill=tk.BOTH, expand=True)

    def _select_b4c_mode(self, mode):
        self.b4c_mode.set(mode)
        self._update_mode_buttons()
        self._refresh_b4c_form()

    def _update_mode_buttons(self):
        cur = self.b4c_mode.get()
        for btn, val in self._mode_buttons:
            if val == cur:
                btn.configure(bg=TAB_SEL, fg="white", activebackground=TAB_SEL, activeforeground="white")
            else:
                btn.configure(bg=TAB_UNSEL, fg=FG_DIM, activebackground=BTN_ACTIVE, activeforeground=FG_TEXT)

    def _refresh_b4c_form(self):
        for w in self.b4c_dynamic.winfo_children():
            w.destroy()
        mode = self.b4c_mode.get()
        form = self.b4c_dynamic
        r = 0
        if mode == "mgmt":
            self._build_mgmt_form(form, r)
        elif mode == "6630":
            self._build_6630_form(form, r)
        elif mode == "6648":
            self._build_6648_form(form, r)
        else:
            self._build_6672_form(form, r)

    # --- Management Port ---
    def _build_mgmt_form(self, form, r):
        tk.Label(form, text="Management Port", font=("Consolas", 11, "bold"),
                 fg=FG_ACCENT, bg=BG_DARK).grid(row=r, column=0, columnspan=2, sticky="w", pady=(5, 5))
        r += 1

        type_frame = tk.Frame(form, bg=BG_DARK)
        type_frame.grid(row=r, column=0, columnspan=2, sticky="w", padx=(10, 0), pady=3)
        tk.Label(type_frame, text="Type:", font=("Consolas", 10), fg=FG_TEXT,
                 bg=BG_DARK).pack(side=tk.LEFT, padx=(0, 8))
        self.mgmt_type = tk.StringVar(value="edn")
        self._mgmt_type_btns = []
        for text, val in [("EDN", "edn"), ("VZB/IDN", "idn")]:
            btn = tk.Button(type_frame, text=text, font=("Consolas", 10, "bold"),
                            relief="flat", padx=10, pady=3, cursor="hand2",
                            command=lambda v=val: self._select_mgmt_type(v))
            btn.pack(side=tk.LEFT, padx=2)
            self._mgmt_type_btns.append((btn, val))
        self._update_mgmt_type_buttons()
        r += 1

        self.mgmt_desc = self._make_combo(form, r, "Port Description", MGMT_DESCRIPTIONS, "MGMT"); r += 1
        self.mgmt_port = self._make_entry(form, r, "Port (Range 2-32)"); r += 1
        self.mgmt_mtu = self._make_entry(form, r, "MTU Size", "9104"); r += 1
        self.mgmt_vlan = self._make_entry(form, r, "VLAN (optional)", ""); r += 1
        self.mgmt_speed = self._make_combo(form, r, "Speed",
                                            ["speed 100", "speed 1000", "speed 10000"],
                                            "speed 1000"); r += 1
        self.mgmt_autoneg = self._make_combo(form, r, "Autonegotiation",
                                              ["autonegotiate", "no autonegotiate"],
                                              "autonegotiate"); r += 1
        self.mgmt_gateway = self._make_entry(form, r, "IPv4 Gateway (optional)", ""); r += 1

        r += 1
        tk.Button(form, text="GENERATE MANAGEMENT CONFIG", command=self._gen_mgmt,
                  font=("Consolas", 12, "bold"), bg=GEN_BTN, fg="white",
                  activebackground=GEN_BTN_HOVER, activeforeground="white",
                  relief="flat", padx=20, pady=8, cursor="hand2"
                  ).grid(row=r, column=0, columnspan=2, pady=15)

        self.mgmt_type.trace_add("write", self._on_mgmt_type_change)

    def _select_mgmt_type(self, val):
        self.mgmt_type.set(val)
        self._update_mgmt_type_buttons()

    def _update_mgmt_type_buttons(self):
        cur = self.mgmt_type.get()
        for btn, val in self._mgmt_type_btns:
            if val == cur:
                btn.configure(bg=TAB_SEL, fg="white", activebackground=TAB_SEL, activeforeground="white")
            else:
                btn.configure(bg=TAB_UNSEL, fg=FG_DIM, activebackground=BTN_ACTIVE, activeforeground=FG_TEXT)

    def _on_mgmt_type_change(self, *args):
        try:
            if self.mgmt_type.get() == "idn":
                self.mgmt_mtu.set("2106")
                self.mgmt_vlan.set("4000")
            else:
                self.mgmt_mtu.set("9104")
                self.mgmt_vlan.set("")
        except Exception:
            pass

    def _gen_mgmt(self):
        if not self.mgmt_port.get().strip():
            self._set_status("Missing: Port is required.", FG_RED)
            return
        port_val = self.mgmt_port.get().strip()
        d = {
            "mgmt_type": self.mgmt_type.get(),
            "description": self.mgmt_desc.get().strip(),
            "port": "1/1/{}".format(port_val) if "/" not in port_val else port_val,
            "mtu": self.mgmt_mtu.get().strip(),
            "vlan": self.mgmt_vlan.get().strip(),
            "speed": self.mgmt_speed.get(),
            "autoneg": self.mgmt_autoneg.get(),
            "gateway_ip": self.mgmt_gateway.get().strip(),
        }
        vendor = self.vendor.get()
        try:
            result = generate_b4c(vendor, "mgmt", d)
            self.b4c_output.delete("1.0", tk.END)
            self.b4c_output.insert("1.0", result)
            self._set_status("Management config generated [{}]".format(vendor), FG_GREEN)
        except Exception as e:
            self._set_status("Error: {}".format(e), FG_RED)

    # --- 6630 ---
    def _build_6630_form(self, form, r):
        tk.Label(form, text="eNB/gNB 6630", font=("Consolas", 11, "bold"),
                 fg=FG_ACCENT, bg=BG_DARK).grid(row=r, column=0, columnspan=2, sticky="w", pady=(5, 5))
        r += 1
        self.f6630_desc = self._make_entry(form, r, "Cell ID (digits)"); r += 1
        self.f6630_port = self._make_entry(form, r, "CSR Port (e.g. 1/1/7)"); r += 1
        self.f6630_speed = self._make_combo(form, r, "Speed",
                                             ["speed 10000", "speed 1000"],
                                             "speed 10000"); r += 1
        self.f6630_autoneg = self._make_combo(form, r, "Autonegotiation",
                                               ["no autonegotiate", "autonegotiate",
                                                "autonegotiate limited"],
                                               "no autonegotiate"); r += 1
        self.f6630_alt = self._make_check(form, r, "Use VLAN 302/402 (instead of 301/401)"); r += 1
        self.f6630_nr = self._make_check(form, r, "Add NR SAP (VLAN 310)"); r += 1
        self.f6630_500 = self._make_check(form, r, "Add LTE CA SAP (VLAN 500)"); r += 1
        self.f6630_520 = self._make_check(form, r, "Add NR CA SAP (VLAN 520)"); r += 1
        self.f6630_ca_port = self._make_entry(form, r, "NR CA Port # (if 520)", ""); r += 1

        r += 1
        tk.Button(form, text="GENERATE 6630 CONFIG", command=self._gen_6630,
                  font=("Consolas", 12, "bold"), bg=GEN_BTN, fg="white",
                  activebackground=GEN_BTN_HOVER, activeforeground="white",
                  relief="flat", padx=20, pady=8, cursor="hand2"
                  ).grid(row=r, column=0, columnspan=2, pady=15)

    def _gen_6630(self):
        if not self.f6630_desc.get().strip() or not self.f6630_port.get().strip():
            self._set_status("Missing: Cell ID and Port are required.", FG_RED)
            return
        d = {
            "desc_id": self.f6630_desc.get().strip(),
            "port": self.f6630_port.get().strip(),
            "speed": self.f6630_speed.get(),
            "autoneg": self.f6630_autoneg.get(),
            "alt_vlan": self.f6630_alt.get(),
            "add_nr": self.f6630_nr.get(),
            "add_500": self.f6630_500.get(),
            "add_520": self.f6630_520.get(),
            "ca520_port": self.f6630_ca_port.get().strip(),
        }
        vendor = self.vendor.get()
        try:
            result = generate_b4c(vendor, "6630", d)
            self.b4c_output.delete("1.0", tk.END)
            self.b4c_output.insert("1.0", result)
            self._set_status("6630 config generated [{}]".format(vendor), FG_GREEN)
        except Exception as e:
            self._set_status("Error: {}".format(e), FG_RED)

    # --- 6648 ---
    def _build_6648_form(self, form, r):
        tk.Label(form, text="eNB/gNB 6648", font=("Consolas", 11, "bold"),
                 fg=FG_ACCENT, bg=BG_DARK).grid(row=r, column=0, columnspan=2, sticky="w", pady=(5, 5))
        r += 1
        self.f6648_desc = self._make_entry(form, r, "Cell ID (digits)"); r += 1
        self.f6648_variant = self._make_combo(form, r, "Deployment Type",
                                               ["ls6", "dual", "tri", "lte_lb"],
                                               "ls6"); r += 1
        self.f6648_alt = self._make_check(form, r, "Use VLAN 302/402 (instead of 301/401)"); r += 1

        r += 1
        tk.Label(form, text="Ports (depends on variant)", font=("Consolas", 10, "bold"),
                 fg=FG_DIM, bg=BG_DARK).grid(row=r, column=0, columnspan=2, sticky="w", pady=(5, 3))
        r += 1
        self.f6648_s1 = self._make_entry(form, r, "S1 LTE Port"); r += 1
        self.f6648_nr = self._make_entry(form, r, "S1 NR Port (ls6/tri)"); r += 1
        self.f6648_nrca = self._make_entry(form, r, "NR CA Port (ls6/dual/lte_lb)"); r += 1

        r += 1
        tk.Button(form, text="GENERATE 6648 CONFIG", command=self._gen_6648,
                  font=("Consolas", 12, "bold"), bg=GEN_BTN, fg="white",
                  activebackground=GEN_BTN_HOVER, activeforeground="white",
                  relief="flat", padx=20, pady=8, cursor="hand2"
                  ).grid(row=r, column=0, columnspan=2, pady=15)

    def _gen_6648(self):
        if not self.f6648_desc.get().strip() or not self.f6648_s1.get().strip():
            self._set_status("Missing: Cell ID and S1 Port are required.", FG_RED)
            return
        variant = self.f6648_variant.get()
        d = {
            "desc_id": self.f6648_desc.get().strip(),
            "variant": variant,
            "alt_vlan": self.f6648_alt.get(),
            "port_s1": self.f6648_s1.get().strip(),
            "port_nr": self.f6648_nr.get().strip(),
            "port_nrca": self.f6648_nrca.get().strip(),
        }
        if variant == "ls6" and not d["port_nr"]:
            self._set_status("Missing: S1 NR Port required for LS6.", FG_RED)
            return
        if variant in ("ls6", "dual", "lte_lb") and not d["port_nrca"]:
            self._set_status("Missing: NR CA Port required for this variant.", FG_RED)
            return
        if variant == "tri" and not d["port_nr"]:
            self._set_status("Missing: S1 NR Port required for Tri-Mode.", FG_RED)
            return
        vendor = self.vendor.get()
        try:
            result = generate_b4c(vendor, "6648", d)
            self.b4c_output.delete("1.0", tk.END)
            self.b4c_output.insert("1.0", result)
            self._set_status("6648 {} config generated [{}]".format(variant.upper(), vendor), FG_GREEN)
        except Exception as e:
            self._set_status("Error: {}".format(e), FG_RED)

    # --- 6672 ---
    def _build_6672_form(self, form, r):
        tk.Label(form, text="eNB/gNB 6672", font=("Consolas", 11, "bold"),
                 fg=FG_ACCENT, bg=BG_DARK).grid(row=r, column=0, columnspan=2, sticky="w", pady=(5, 5))
        r += 1
        self.f6672_desc = self._make_entry(form, r, "Cell ID (digits)"); r += 1
        self.f6672_port = self._make_entry(form, r, "Combined Port (e.g. 1/1/13)"); r += 1
        self.f6672_25g = self._make_check(form, r, "Use 25G speed"); r += 1

        r += 1
        tk.Button(form, text="GENERATE 6672 CONFIG", command=self._gen_6672,
                  font=("Consolas", 12, "bold"), bg=GEN_BTN, fg="white",
                  activebackground=GEN_BTN_HOVER, activeforeground="white",
                  relief="flat", padx=20, pady=8, cursor="hand2"
                  ).grid(row=r, column=0, columnspan=2, pady=15)

    def _gen_6672(self):
        if not self.f6672_desc.get().strip() or not self.f6672_port.get().strip():
            self._set_status("Missing: Cell ID and Port are required.", FG_RED)
            return
        d = {
            "desc_id": self.f6672_desc.get().strip(),
            "port": self.f6672_port.get().strip(),
            "use_25g": self.f6672_25g.get(),
        }
        vendor = self.vendor.get()
        try:
            result = generate_b4c(vendor, "6672", d)
            self.b4c_output.delete("1.0", tk.END)
            self.b4c_output.insert("1.0", result)
            self._set_status("6672 config generated [{}]".format(vendor), FG_GREEN)
        except Exception as e:
            self._set_status("Error: {}".format(e), FG_RED)


# ===========================================================================
#  ENTRY POINT
# ===========================================================================

def main():
    root = tk.Tk()
    try:
        if sys.platform == "win32":
            root.iconbitmap(default="")
    except Exception:
        pass
    app = AutoPhil2App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
