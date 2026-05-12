#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ENSE Config Generator v1.0 — GUI Edition
Unified Nokia/Ciena B4A + B4C ENSE configuration tool.

Replaces:
  - Brian's Excel "B4A Port Config 3.0" (backhaul link)
  - AutoPhil B4C Scripting Tool .exe   (RAN service ports)

Two modes:
  1. B4A Backhaul Link    B4A <-> B4C: Port + Interface + ISIS + BGP
  2. B4C Service Ports    B4C <-> eNB/gNB: Physical ports + VPLS SAPs

Standalone GUI — tkinter only, no external dependencies.
Package with PyInstaller: pyinstaller --onefile --windowed ense_config_gui.py
"""

import sys
import os
import re
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

VERSION = "1.1"

# ===========================================================================
#  CONFIG GENERATION LOGIC (from ense_config.py CLI)
# ===========================================================================

def _b4a_interface_name(port, vlan):
    return "INT-{}:{}-Base".format(port, vlan)


def _b4a_gen_port(port, hostname, bw_speed):
    lines = []
    p = "\\configure port {}".format(port)
    lines.append(p)
    lines.append("info")
    lines.append("{}        shutdown".format(p))
    lines.append('{}        description "{}_{}"'.format(p, hostname, port))
    lines.append("{}        ethernet mode hybrid".format(p))
    lines.append("{}             ethernet  encap-type dot1q".format(p))
    lines.append("{}             ethernet              mtu 9104".format(p))
    lines.append("{}             ethernet  crc-monitor   sd-threshold 2 multiplier 5".format(p))
    lines.append("{}             ethernet  crc-monitor   sf-threshold 1".format(p))
    lines.append("{}          ethernet    lldp      dest-mac nearest-bridge      admin-status tx-rx".format(p))
    lines.append("{}          ethernet    lldp     dest-mac nearest-bridge       notification".format(p))
    lines.append("{}          ethernet    lldp      dest-mac nearest-bridge      tx-tlvs port-desc sys-name sys-desc sys-cap".format(p))
    lines.append("{}          ethernet    lldp       dest-mac nearest-bridge     port-id-subtype tx-if-name".format(p))
    lines.append("{}          ethernet            util-stats-interval 30".format(p))
    lines.append('{}          ethernet            egress-port-qos-policy "40012"'.format(p))
    lines.append("{}          ethernet speed {}".format(p, bw_speed))
    lines.append("{}        no shutdown".format(p))
    lines.append("info")
    lines.append("exit all")
    lines.append("")
    lines.append('show port {} | match expression "Description|Speed|Auto|Admin State|Oper State|Last|Phys State|dBm|Link Level|Model Number|Physical Link|Rate|Warn|Configured"'.format(port))
    return lines


def _b4a_gen_interface(port, vlan, hostname, p2p_ip, bw_speed, egress_instance):
    intf = _b4a_interface_name(port, vlan)
    rg = "QG-NE-{}M".format(bw_speed)
    lines = []
    b = '\\config router interface "{}"'.format(intf)
    lines.append(b)
    lines.append("info")
    lines.append('{} address {}/31'.format(b, p2p_ip))
    lines.append('{} description "{}_{}"'.format(b, hostname, intf))
    lines.append('{} egress vlan-qos-policy "40011"'.format(b))
    lines.append('{}  port {}:{}'.format(b, port, vlan))
    lines.append('{} ingress qos "40021"'.format(b))
    lines.append('{}  egress egress-remark-policy "40021"'.format(b))
    lines.append('{}  qos 40022 egress-port-redirect-group "{}" egress-instance {}'.format(b, rg, egress_instance))
    lines.append("{}     bfd 50 receive 50 multiplier 5 type fp".format(b))
    lines.append("{} enable-ingress-stats".format(b))
    lines.append("{}  no shutdown".format(b))
    lines.append("info")
    lines.append("exit all")
    lines.append("")
    lines.append("show router interface | match {}:{}".format(port, vlan))
    return lines


def _b4a_gen_nni_interface(nni_port, nni_vlan, b4c_host, b4c_port, b4c_vlan, nni_p2p, bw_speed):
    intf = _b4a_interface_name(nni_port, nni_vlan)
    rg = "QG-NE-{}M".format(bw_speed)
    lines = []
    b = '\\config router interface "{}"'.format(intf)
    lines.append(b)
    lines.append("info")
    lines.append('{} address {}/31'.format(b, nni_p2p))
    lines.append('{} description "{}_{}"'.format(b, b4c_host, _b4a_interface_name(b4c_port, b4c_vlan)))
    lines.append('{}  port {}:{}'.format(b, nni_port, nni_vlan))
    lines.append('{}  qos 40022 egress-port-redirect-group "{}" egress-instance {}'.format(b, rg, nni_vlan))
    lines.append("{}  bfd 50 receive 50 multiplier 5 type cpm-np".format(b))
    lines.append("{} enable-ingress-stats".format(b))
    lines.append("{}  no shutdown".format(b))
    lines.append("info")
    lines.append("exit all")
    lines.append("")
    lines.append("show router interface | match {}:{}".format(nni_port, nni_vlan))
    return lines


def _b4a_gen_isis(port, vlan):
    intf = _b4a_interface_name(port, vlan)
    b = '\\config router isis 5 interface "{}"'.format(intf)
    lines = [b, "info",
             "{}   level-capability level-1".format(b),
             "{}     interface-type point-to-point".format(b),
             "{}     bfd-enable ipv4".format(b),
             "{}     level 1     metric 1000000".format(b),
             "{}     no shutdown".format(b),
             b, "info", "exit all", "",
             'show router isis 5 interface "{}"'.format(intf)]
    return lines


def _b4a_gen_bgp(bgp_group, neighbor_ip, desc):
    b = '\\config router bgp   group "{}" neighbor {}'.format(bgp_group, neighbor_ip)
    return [b,
            '{} description "iBGP-TO-{}"'.format(b, desc),
            '{}  authentication-key "eNSEbgp" '.format(b),
            "exit all", "\\admin save", "",
            "show router bgp summary | match {} post-lines 3".format(neighbor_ip)]


def _b4a_gen_verify(port, vlan, csr_ip):
    return [
        'show port {} | match expression "Description|Speed|Auto|Admin State|Oper State|Last|Phys State|dBm|Link Level|Model Number|Physical Link|Rate|Warn|Configured"'.format(port),
        "show router interface | match {}:{}".format(port, vlan),
        "show router arp | match {}:{}".format(port, vlan),
        "show router bgp summary | match {} post-lines 3".format(csr_ip)]


def _b4a_gen_spoke_ref():
    return [
        "# ENSE-SPOKE: HUB site needs RR-5-ENSESR_SPOKE group with SPOKE neighbor.",
        "# Example:", "",
        '\\config router bgp   group "RR-5-ENSESR_SPOKE" neighbor <SPOKE_CSR_IP>',
        '\\config router bgp   group "RR-5-ENSESR_SPOKE" neighbor <SPOKE_CSR_IP> description "iBGP-TO-<SPOKE_HOSTNAME>"',
        '\\config router bgp   group "RR-5-ENSESR_SPOKE" neighbor <SPOKE_CSR_IP>  authentication-key "eNSEbgp" ',
        "exit all", "\\admin save",
        "show router bgp summary | match <SPOKE_CSR_IP> post-lines 3"]


# ---- B4C Templates ----

TPL_EDN_NOKIA = """\
#ONLY FOR NOKIA B4C ENSE ROUTERS
#################
# LOGICAL  PORT
#################
\\configure
    port {port}
        shutdown
        description "LINK-TO-OAM/EDN-{description}"
        ethernet
            {speed}
            mode access
            encap-type dot1q
            {autonegotiation}
            mtu {mtu_size}
        exit
        no shutdown
    exit

#################
# SERVICE SAPS
#################
\\configure
    service
        vpls 400
            sap {port}{vlan_entry} create
                ingress
                    qos 41031
                exit
                no shutdown
            exit
"""

TPL_IDN_NOKIA = """\
#ONLY FOR NOKIA B4C ENSE ROUTERS
#################
# LOGICAL  PORT
#################
\\configure
    port {port}
        shutdown
        description "LINK-TO-OAM/EDN-{description}"
        ethernet
            {speed}
            mode access
            encap-type dot1q
            {autonegotiation}
            mtu {mtu_size}
        exit
        no shutdown
    exit

#################
# SERVICE SAPS
#################
\\configure
    service
        vpls 450 customer 4 create
            description "IXRs-IDN MGMT and Console and Server"
            service-mtu 1528
            allow-ip-int-bind
            exit
            stp
                shutdown
            exit
            sap {port}{vlan_entry} create
                ingress
                    qos 41031
                exit
                no shutdown
            exit
"""

TPL_EDN_SUBNET = """\
#EDN IPv4 SUBNET
\\configure
    service
        vprn 4
            interface "INT-VPLS400-CELL_MGMT" create
            address {gateway_ip}/29
            vpls "VPLS400"
            exit
        exit
"""

TPL_IDN_SUBNET = """\
#IDN IPv4 SUBNET
\\configure
    service
        vprn 4
            interface "INT-VPLS450-CELL_MGMT" create
            address {gateway_ip}/27
            vpls "VPLS450"
            exit
        exit
"""

TPL_ENB_6630 = """\
#ONLY FOR NOKIA B4C ENSE ROUTERS
#CHECK LIVE CONFIG BEFORE APPLYING
#---------------------
#    LOGICAL PORT
#---------------------
\\configure
    port {port}
        shutdown
        description {enb_gnb_description}
        ethernet
            {enb_gnb_speed_var}
            mode access
            encap-type dot1q
            {autonegotiation}
            mtu {mtu_size}
        exit
        no shutdown
    exit

#---------------------
#    SERVICE SAPS
#---------------------
\\configure
    service
        vpls 100
            sap {port}{lte_vlan_input} create
                ingress
                    qos 41031
                exit
                egress
                    agg-rate
                        rate max cir max
                    exit
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                exit
                no shutdown
            exit

\\configure
    service
        vpls 400
            sap {port}{oam_vlan_input} create
                ingress
                    qos 41031
                exit
                no shutdown
            exit
"""

TPL_310_VPLS = """\
#VPLS 310 SAPS
\\configure
    service
        vpls 310
            sap {port}{nr_vlan_var} create
                ingress
                    qos 41031
                exit
                egress
                    agg-rate
                        rate max cir max
                    exit
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                exit
                no shutdown
            exit
"""

TPL_CA500 = """\
#LTE CARRIER AGG SAPS
\\configure
    service
        vpls 5000
            sap {port}{vlan500_var} create
                ingress
                    qos 41032
                exit
                egress
                    agg-rate
                        rate max cir max
                    exit
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                exit
                no shutdown
            exit
"""

TPL_CA520 = """\
#NR CA PORT CONFIG
\\configure port 1/1/{ca520_port}
    shutdown
    {enb_gnb_description}-NR_CA
    ethernet
        speed 10000
        mode access
        encap-type dot1q
        mtu 9104
        down-when-looped
            no shutdown
        exit
        hold-time down 5
    exit
    no shutdown
exit

#NR CA SERVICE SAPS
\\configure
    service
        vpls 5200
            sap 1/1/{ca520_port}:520 create
                no shutdown
            exit
"""

TPL_6648_LS6 = """\
#S1 LTE Port
/configure port {port_s1}
        shutdown
        description "eNB-{description_id}_6648_S1_LTE_OAM_LTE_CA"
        ethernet
            speed 10000
            mode access
            encap-type dot1q
            mtu 9104
            down-when-looped
                no shutdown
            exit
            hold-time down 5
        exit
        no shutdown

#S1 NR Port
/configure port {port_nr}
        shutdown
        description "eNB-{description_id}_6648_S1_NR"
        ethernet
            speed 10000
            mode access
            encap-type dot1q
            mtu 9104
            down-when-looped
                no shutdown
            exit
            hold-time down 5
            ssm
                no shutdown
            exit
            util-stats-interval 30
        exit
        no shutdown

#NR CA Port
/configure port {port_nrca}
        shutdown
        description "eNB-{description_id}_6648_NR_CA"
        ethernet
            speed 10000
            mode access
            encap-type dot1q
            mtu 9212
            down-when-looped
                no shutdown
            exit
            hold-time down 5
        exit
        no shutdown

# S1 LTE SAP
/configure
    service
        vpls 100
            sap {port_s1}:{s1_vlan} create
                ingress
                    qos 41031
                exit
                egress
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                exit
                no shutdown
            exit
        exit

# OAM SAP
/configure
    service
        vpls 400
            sap {port_s1}:{oam_vlan} create
                ingress
                    qos 41031
                exit
                no shutdown
            exit
        exit

# S1 NR SAP
/configure
    service
        vpls 310
            sap {port_nr}:310 create
                ingress
                    qos 41031
                exit
                egress
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                exit
                no shutdown
            exit
        exit

# LTE CA SAP
/configure
    service
        vpls 5000
            sap {port_s1}:500 create
                ingress
                    qos 41032
                exit
                egress
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                    agg-rate
                        rate max cir max
                    exit
                exit
                no shutdown
            exit

# NR CA SAP
/configure
    service
        vpls 5200
            sap {port_nrca}:520 create
                no shutdown
            exit
        exit
"""

TPL_6648_DUAL = """\
/configure port {port_s1}
        shutdown
        description "eNB-{description_id}_6648_S1_LTE_S1_NR_OAM"
        ethernet
            speed 10000
            mode access
            encap-type dot1q
            mtu 9104
            down-when-looped
                no shutdown
            exit
            hold-time down 5
        exit
        no shutdown

/configure port {port_nrca}
        shutdown
        description "eNB-{description_id}_6648_S1_NR"
        ethernet
            speed 10000
            mode access
            encap-type dot1q
            mtu 9104
            down-when-looped
                no shutdown
            exit
            hold-time down 5
            ssm
                no shutdown
            exit
            util-stats-interval 30
        exit
        no shutdown

# S1 LTE and NR SAP
/configure
    service
        vpls 100
            sap {port_s1}:{s1_vlan} create
                ingress
                    qos 41031
                exit
                egress
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                exit
                no shutdown
            exit
        exit

# OAM SAP
/configure
    service
        vpls 400
            sap {port_s1}:{oam_vlan} create
                ingress
                    qos 41031
                exit
                no shutdown
            exit
        exit

# S1 NR SAP
/configure
    service
        vpls 310
            sap {port_s1}:310 create
                ingress
                    qos 41031
                exit
                egress
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                exit
                no shutdown
            exit
        exit

# LTE CA SAP
/configure
    service
        vpls 5000
            sap {port_s1}:500 create
                ingress
                    qos 41032
                exit
                egress
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                    agg-rate
                        rate max cir max
                    exit
                exit
                no shutdown
            exit

# NR CA SAP
/configure
    service
        vpls 5200
            sap {port_nrca}:520 create
                no shutdown
            exit
"""

TPL_6648_TRI = """\
/configure port {port_s1}
        shutdown
        description "eNB-{description_id}_6648_S1_LTE_OAM_LTE_CA"
        ethernet
            speed 10000
            mode access
            encap-type dot1q
            mtu 9104
            down-when-looped
                no shutdown
            exit
            hold-time down 5
        exit
        no shutdown

/configure port {port_nr}
        shutdown
        description "eNB-{description_id}_6648_S1_NR"
        ethernet
            speed 10000
            mode access
            encap-type dot1q
            mtu 9104
            down-when-looped
                no shutdown
            exit
            hold-time down 5
            ssm
                no shutdown
            exit
            util-stats-interval 30
        exit
        no shutdown

# S1 LTE SAP
/configure
    service
        vpls 100
            sap {port_s1}:{s1_vlan} create
                ingress
                    qos 41031
                exit
                egress
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                exit
                no shutdown
            exit
        exit

# OAM SAP
/configure
    service
        vpls 400
            sap {port_s1}:{oam_vlan} create
                ingress
                    qos 41031
                exit
                no shutdown
            exit
        exit

# S1 NR SAP
/configure
    service
        vpls 310
            sap {port_nr}:310 create
                ingress
                    qos 41031
                exit
                egress
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                exit
                no shutdown
            exit
        exit

# LTE CA SAP
/configure
    service
        vpls 5000
            sap {port_s1}:500 create
                ingress
                    qos 41032
                exit
                egress
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                    agg-rate
                        rate max cir max
                    exit
                exit
                no shutdown
            exit
"""

TPL_6648_LTE_LB = """\
/configure port {port_s1}
        shutdown
        description "eNB-{description_id}_6648_S1_LTE_OAM_LTE_CA"
        ethernet
            speed 10000
            mode access
            encap-type dot1q
            mtu 9104
            down-when-looped
                no shutdown
            exit
            hold-time down 5
        exit
        no shutdown

/configure port {port_nrca}
        shutdown
        description "eNB-{description_id}_NR_CA"
        ethernet
            speed 10000
            mode access
            encap-type dot1q
            mtu 9212
            down-when-looped
                no shutdown
            exit
            hold-time down 5
        exit
    no shutdown
exit

# S1 LTE SAP
/configure
    service
        vpls 100
            sap {port_s1}:{s1_vlan} create
                ingress
                    qos 41031
                exit
                egress
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                exit
                no shutdown
            exit
        exit

# OAM SAP
/configure
    service
        vpls 400
            sap {port_s1}:{oam_vlan} create
                ingress
                    qos 41031
                exit
                no shutdown
            exit
        exit

# LTE CA SAP
/configure
    service
        vpls 5000
            sap {port_s1}:500 create
                ingress
                    qos 41032
                exit
                egress
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                    agg-rate
                        rate max cir max
                    exit
                exit
                no shutdown
            exit

# NR CA SAP
/configure
    service
        vpls 5200
            sap {port_nrca}:5200 create
                no shutdown
            exit
        exit
"""

TPL_6672 = """\
# 6672 PORT CONFIG
/configure port {port_6672}
        shutdown
        description "eNB-{description_id}_6672_S1_LTE_S1_NR_OAM_LTE_NR_CA"
        ethernet
            {port_6672_physical_speed}
            mode access
            encap-type dot1q
            mtu 9212
            down-when-looped
                no shutdown
            exit
            hold-time down 5
        exit
        no shutdown

# S1 SAP
/configure
    service
        vpls 100
            sap {port_6672}:301 create
                ingress
                    qos 41031
                exit
                egress
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                exit
                no shutdown
            exit
        exit

# OAM SAP
/configure
    service
        vpls 400
            sap {port_6672}:401 create
                ingress
                    qos 41031
                exit
                no shutdown
            exit
        exit

# NR SAP
/configure
    service
        vpls 310
            sap {port_6672}:310 create
                ingress
                    qos 41031
                exit
                egress
                    vlan-qos-policy "RAN_Downstream_Scheduler"
                exit
                no shutdown
            exit
        exit

# LTE CA SAP
/configure
    service
        vpls 5000
            sap {port_6672}:500 create
                no shutdown
            exit
        exit

# NR CA SAP
/configure
    service
        vpls 5200
            sap {port_6672}:520 create
                no shutdown
            exit
        exit
"""


# ===========================================================================
#  B4A GENERATION
# ===========================================================================

def generate_b4a(d):
    """Build full B4A config from dict of form values. Returns string."""
    spd = 1000 if d["bw"] == "1G" else 10000
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    sections = []

    # Header
    hdr = []
    hdr.append("# Generated: {}".format(ts))
    hdr.append("# B4C: {}".format(d["b4c_host"]))
    if d.get("b4a_host"):
        hdr.append("# B4A: {}".format(d["b4a_host"]))
    hdr.append("# Bandwidth: {} (speed {})".format(d["bw"], spd))
    hdr.append("# B4C Port: {}  VLAN: {}  P2P: {}  CSR: {}".format(
        d["b4c_port"], d["b4c_vlan"], d["b4c_p2p"], d["b4c_csr"]))
    hdr.append("# NNI Port: {}  VLAN: {}  P2P: {}  CSR: {}".format(
        d["nni_port"], d["nni_vlan"], d["nni_p2p"], d["nni_csr"]))
    hdr.append("# BGP Group: {}".format(d["bgp"]))
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

    # Format output
    out = []
    out.append("ENSE Config Generator v{} - B4A Backhaul".format(VERSION))
    out.append("=" * 70)
    for title, lines in sections:
        out.append("")
        out.append("=" * 70)
        out.append("  {}".format(title))
        out.append("=" * 70)
        for line in lines:
            out.append(line)

    return "\n".join(out)


# ===========================================================================
#  B4C GENERATION
# ===========================================================================

def generate_b4c_management(d):
    """Management port config."""
    data = {
        "port": d["port"],
        "description": d["description"],
        "mtu_size": d["mtu"],
        "speed": d["speed"],
        "autonegotiation": d["autoneg"],
        "vlan_entry": ":{}".format(d["vlan"]) if d.get("vlan") else "",
    }

    out = []
    if d["mgmt_type"] == "edn":
        out.append(TPL_EDN_NOKIA.format_map(data))
    else:
        out.append(TPL_IDN_NOKIA.format_map(data))

    if d.get("gateway_ip"):
        data["gateway_ip"] = d["gateway_ip"]
        tpl = TPL_EDN_SUBNET if d["mgmt_type"] == "edn" else TPL_IDN_SUBNET
        out.append(tpl.format_map(data))

    title = "MANAGEMENT PORT - {}".format(d["mgmt_type"].upper())
    header = "ENSE Config Generator v{} - B4C Service\n{}\n  {}\n{}".format(
        VERSION, "=" * 70, title, "=" * 70)
    return header + "\n" + "\n".join(out)


def generate_b4c_6630(d):
    """6630 baseband config."""
    lte_v = "302" if d.get("alt_vlan") else "301"
    oam_v = "402" if d.get("alt_vlan") else "401"
    data = {
        "port": d["port"],
        "enb_gnb_description": '"eNB-{}-6630"'.format(d["desc_id"]),
        "enb_gnb_speed_var": d["speed"],
        "autonegotiation": d["autoneg"],
        "mtu_size": "9104",
        "lte_vlan_input": ":{}".format(lte_v),
        "oam_vlan_input": ":{}".format(oam_v),
        "nr_vlan_var": ":310",
        "vlan500_var": ":500",
        "ca520_port": d.get("ca520_port", ""),
    }

    out = [TPL_ENB_6630.format_map(data)]
    if d.get("add_nr"):
        out.append(TPL_310_VPLS.format_map(data))
    if d.get("add_500"):
        out.append(TPL_CA500.format_map(data))
    if d.get("add_520") and d.get("ca520_port"):
        out.append(TPL_CA520.format_map(data))

    title = "eNB/gNB 6630 - {}".format(d["desc_id"])
    header = "ENSE Config Generator v{} - B4C Service\n{}\n  {}\n{}".format(
        VERSION, "=" * 70, title, "=" * 70)
    return header + "\n" + "\n".join(out)


def generate_b4c_6648(d):
    """6648 baseband config."""
    s1_v = "302" if d.get("alt_vlan") else "301"
    oam_v = "402" if d.get("alt_vlan") else "401"
    data = {
        "description_id": d["desc_id"],
        "s1_vlan": s1_v,
        "oam_vlan": oam_v,
    }

    variant = d["variant"]
    if variant == "ls6":
        data["port_s1"] = d["port_s1"]
        data["port_nr"] = d["port_nr"]
        data["port_nrca"] = d["port_nrca"]
        tpl = TPL_6648_LS6
    elif variant == "dual":
        data["port_s1"] = d["port_s1"]
        data["port_nrca"] = d["port_nrca"]
        tpl = TPL_6648_DUAL
    elif variant == "tri":
        data["port_s1"] = d["port_s1"]
        data["port_nr"] = d["port_nr"]
        tpl = TPL_6648_TRI
    else:  # lte_lb
        data["port_s1"] = d["port_s1"]
        data["port_nrca"] = d["port_nrca"]
        tpl = TPL_6648_LTE_LB

    title = "eNB/gNB 6648 {} - {}".format(variant.upper(), d["desc_id"])
    header = "ENSE Config Generator v{} - B4C Service\n{}\n  {}\n{}".format(
        VERSION, "=" * 70, title, "=" * 70)
    return header + "\n" + tpl.format_map(data)


def generate_b4c_6672(d):
    """6672 baseband config."""
    data = {
        "port_6672": d["port"],
        "description_id": d["desc_id"],
        "port_6672_physical_speed": "speed 25000" if d.get("use_25g") else "speed 10000",
    }
    title = "eNB/gNB 6672 - {}".format(d["desc_id"])
    header = "ENSE Config Generator v{} - B4C Service\n{}\n  {}\n{}".format(
        VERSION, "=" * 70, title, "=" * 70)
    return header + "\n" + TPL_6672.format_map(data)


# ===========================================================================
#  GUI APPLICATION
# ===========================================================================

# Colors
BG_DARK = "#1a1a2e"
BG_MID = "#16213e"
BG_CARD = "#0f3460"
FG_TEXT = "#e0e0e0"
FG_DIM = "#8899aa"
FG_ACCENT = "#00d4ff"
FG_GREEN = "#00e676"
FG_ORANGE = "#ffab40"
BTN_BG = "#0f3460"
BTN_ACTIVE = "#1a5276"
ENTRY_BG = "#1c2541"
OUTPUT_BG = "#0d1117"
TAB_SEL = "#00796b"
TAB_UNSEL = "#1c2541"

# Predefined port descriptions (from AutoPhil)
MGMT_DESCRIPTIONS = [
    "MGMT",
    "SITE_BOSS",
    "POWER_PLANT",
    "SHARK_METER",
    "To-IXR-e-Port-1/1/20-s-Hook-to-Management-Port",
    "To-IXR-e-Port-1/1/24-s-Hook-to-Management-Port",
    "EDN-VZB_IDN-OneFiber-Uplink",
]


class ENSEConfigApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ENSE Config Generator v{}".format(VERSION))
        self.root.geometry("1100x820")
        self.root.minsize(900, 700)
        self.root.configure(bg=BG_DARK)

        # Style
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self._configure_styles()

        # Main container
        main = tk.Frame(root, bg=BG_DARK)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Banner
        banner_frame = tk.Frame(main, bg=BG_DARK)
        banner_frame.pack(fill=tk.X, pady=(0, 10))
        tk.Label(banner_frame, text="ENSE Config Generator v{}".format(VERSION),
                 font=("Consolas", 18, "bold"), fg=FG_ACCENT, bg=BG_DARK).pack(side=tk.LEFT)
        tk.Label(banner_frame, text="Nokia B4A/B4C ENSE Backhaul + Service Configuration",
                 font=("Consolas", 10), fg=FG_DIM, bg=BG_DARK).pack(side=tk.LEFT, padx=(20, 0))

        # Notebook (tabs)
        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Tab 1: B4A
        self.b4a_frame = tk.Frame(self.notebook, bg=BG_DARK)
        self.notebook.add(self.b4a_frame, text="  B4A Backhaul Link  ")
        self._build_b4a_tab()

        # Tab 2: B4C
        self.b4c_frame = tk.Frame(self.notebook, bg=BG_DARK)
        self.notebook.add(self.b4c_frame, text="  B4C Service Ports  ")
        self._build_b4c_tab()

        # Bottom bar: VPLS reference + vendor selector
        bottom = tk.Frame(main, bg=BG_MID, pady=6, padx=10)
        bottom.pack(fill=tk.X, pady=(5, 0))

        tk.Label(bottom,
                 text="VPLS: 100=LTE S1  310=NR S1  400=OAM  450=IDN  5000=LTE CA  5200=NR CA",
                 font=("Consolas", 9), fg=FG_DIM, bg=BG_MID).pack(side=tk.LEFT)

        # Vendor selector (right side of bottom bar)
        vendor_frame = tk.Frame(bottom, bg=BG_MID)
        vendor_frame.pack(side=tk.RIGHT)
        self.vendor = tk.StringVar(value="Nokia")
        for v in ["Nokia", "Nokia IXR-e2", "Ciena"]:
            tk.Radiobutton(vendor_frame, text=v, variable=self.vendor, value=v,
                           font=("Consolas", 9), fg=FG_TEXT, bg=BG_MID,
                           selectcolor=TAB_SEL, activebackground=BG_MID,
                           activeforeground=FG_ACCENT, indicatoron=0,
                           padx=8, pady=2, relief="flat",
                           ).pack(side=tk.LEFT, padx=2)

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
                     foreground=FG_TEXT, font=("Consolas", 10))
        s.configure("TCheckbutton", background=BG_DARK, foreground=FG_TEXT,
                     font=("Consolas", 10))
        s.map("TCheckbutton", background=[("active", BG_DARK)])

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
        """Create output text area with Copy and Save buttons."""
        out_frame = tk.Frame(parent, bg=BG_DARK)

        # Button bar
        btn_bar = tk.Frame(out_frame, bg=BG_DARK)
        btn_bar.pack(fill=tk.X, pady=(0, 4))

        text = tk.Text(out_frame, font=("Consolas", 10), bg=OUTPUT_BG, fg=FG_GREEN,
                       insertbackground=FG_ACCENT, relief="flat", wrap=tk.NONE,
                       highlightthickness=1, highlightcolor=FG_ACCENT,
                       highlightbackground=BG_MID)
        text.pack(fill=tk.BOTH, expand=True)

        # Scrollbars
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
                messagebox.showinfo("Copied", "Config copied to clipboard.")

        def save_output():
            content = text.get("1.0", tk.END).strip()
            if not content:
                messagebox.showwarning("Empty", "Nothing to save.")
                return
            path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                initialfile="ense_config_{}.txt".format(
                    datetime.datetime.now().strftime("%Y%m%d_%H%M")))
            if path:
                with open(path, "w") as f:
                    f.write(content)
                messagebox.showinfo("Saved", "Saved: {}".format(path))

        def clear_output():
            text.delete("1.0", tk.END)

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

        tk.Label(form, text="B4C Side (CSR)", font=("Consolas", 11, "bold"),
                 fg=FG_ACCENT, bg=BG_DARK).grid(row=r, column=0, columnspan=2,
                                                  sticky="w", pady=(5, 5))
        r += 1
        self.b4a_b4c_host = self._make_entry(form, r, "B4C Host Name"); r += 1
        self.b4a_b4c_port = self._make_entry(form, r, "NNI Port (B4C)", "1/1/1"); r += 1
        self.b4a_b4c_vlan = self._make_entry(form, r, "ODD VLAN", "1001"); r += 1
        self.b4a_bw = self._make_combo(form, r, "Bandwidth", ["10G", "1G"], "10G"); r += 1
        self.b4a_b4c_p2p = self._make_entry(form, r, "P2P IPv4 (EBH ODD)"); r += 1
        self.b4a_b4c_csr = self._make_entry(form, r, "CSR GLOBAL LoO"); r += 1

        r += 1
        tk.Label(form, text="B4A / NNI Side", font=("Consolas", 11, "bold"),
                 fg=FG_ACCENT, bg=BG_DARK).grid(row=r, column=0, columnspan=2,
                                                  sticky="w", pady=(10, 5))
        r += 1
        self.b4a_nni_port = self._make_entry(form, r, "NNI Port (B4A)", "1/1/c8/1"); r += 1
        self.b4a_nni_vlan = self._make_entry(form, r, "EVEN VLAN"); r += 1
        self.b4a_nni_p2p = self._make_entry(form, r, "P2P IPv4 (NNI)"); r += 1
        self.b4a_nni_csr = self._make_entry(form, r, "CSR GLOBAL LoO (NNI)"); r += 1

        r += 1
        tk.Label(form, text="BGP", font=("Consolas", 11, "bold"),
                 fg=FG_ACCENT, bg=BG_DARK).grid(row=r, column=0, columnspan=2,
                                                  sticky="w", pady=(10, 5))
        r += 1
        self.b4a_bgp = self._make_combo(form, r, "BGP Group",
                                         ["RR-5-ENSESR_CSR", "RR-5-ENSESR_SPOKE"],
                                         "RR-5-ENSESR_CSR"); r += 1

        r += 1
        tk.Label(form, text="Info Only (optional)", font=("Consolas", 11, "bold"),
                 fg=FG_DIM, bg=BG_DARK).grid(row=r, column=0, columnspan=2,
                                              sticky="w", pady=(10, 5))
        r += 1
        self.b4a_b4a_host = self._make_entry(form, r, "B4A Host Name", ""); r += 1

        r += 1
        tk.Button(form, text="GENERATE B4A CONFIG", command=self._gen_b4a,
                  font=("Consolas", 12, "bold"), bg="#00796b", fg="white",
                  activebackground="#00897b", activeforeground="white",
                  relief="flat", padx=20, pady=8, cursor="hand2"
                  ).grid(row=r, column=0, columnspan=2, pady=15)

        # Right: output
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
                messagebox.showwarning("Missing", "{} is required.".format(name))
                return

        d = {
            "b4c_host": self.b4a_b4c_host.get().strip(),
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
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # -----------------------------------------------------------------------
    #  B4C TAB — toggle-button mode selector (visible on Windows dark theme)
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

        # Mode selector — toggle buttons instead of invisible radio buttons
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

        # Dynamic form area
        self.b4c_dynamic = tk.Frame(self.b4c_form_container, bg=BG_DARK)
        self.b4c_dynamic.grid(row=1, column=0, columnspan=2, sticky="nsew")

        self._refresh_b4c_form()

        # Right: output
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
                btn.configure(bg=TAB_SEL, fg="white", activebackground=TAB_SEL,
                              activeforeground="white")
            else:
                btn.configure(bg=TAB_UNSEL, fg=FG_DIM, activebackground=BTN_ACTIVE,
                              activeforeground=FG_TEXT)

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
                 fg=FG_ACCENT, bg=BG_DARK).grid(row=r, column=0, columnspan=2,
                                                  sticky="w", pady=(5, 5))
        r += 1

        # EDN / VZB/IDN type selector as toggle buttons
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

        self.mgmt_desc = self._make_combo(form, r, "Port Description",
                                           MGMT_DESCRIPTIONS, "MGMT"); r += 1
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
                  font=("Consolas", 12, "bold"), bg="#00796b", fg="white",
                  activebackground="#00897b", activeforeground="white",
                  relief="flat", padx=20, pady=8, cursor="hand2"
                  ).grid(row=r, column=0, columnspan=2, pady=15)

        # Apply IDN defaults when switching type
        self.mgmt_type.trace_add("write", self._on_mgmt_type_change)

    def _select_mgmt_type(self, val):
        self.mgmt_type.set(val)
        self._update_mgmt_type_buttons()

    def _update_mgmt_type_buttons(self):
        cur = self.mgmt_type.get()
        for btn, val in self._mgmt_type_btns:
            if val == cur:
                btn.configure(bg=TAB_SEL, fg="white", activebackground=TAB_SEL,
                              activeforeground="white")
            else:
                btn.configure(bg=TAB_UNSEL, fg=FG_DIM, activebackground=BTN_ACTIVE,
                              activeforeground=FG_TEXT)

    def _on_mgmt_type_change(self, *args):
        """Update defaults when switching EDN <-> IDN."""
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
            messagebox.showwarning("Missing", "Port is required.")
            return
        d = {
            "mgmt_type": self.mgmt_type.get(),
            "description": self.mgmt_desc.get().strip(),
            "port": "1/1/{}".format(self.mgmt_port.get().strip()) if "/" not in self.mgmt_port.get() else self.mgmt_port.get().strip(),
            "mtu": self.mgmt_mtu.get().strip(),
            "vlan": self.mgmt_vlan.get().strip(),
            "speed": self.mgmt_speed.get(),
            "autoneg": self.mgmt_autoneg.get(),
            "gateway_ip": self.mgmt_gateway.get().strip(),
        }
        try:
            result = generate_b4c_management(d)
            self.b4c_output.delete("1.0", tk.END)
            self.b4c_output.insert("1.0", result)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # --- 6630 ---
    def _build_6630_form(self, form, r):
        tk.Label(form, text="eNB/gNB 6630", font=("Consolas", 11, "bold"),
                 fg=FG_ACCENT, bg=BG_DARK).grid(row=r, column=0, columnspan=2,
                                                  sticky="w", pady=(5, 5))
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
                  font=("Consolas", 12, "bold"), bg="#00796b", fg="white",
                  activebackground="#00897b", activeforeground="white",
                  relief="flat", padx=20, pady=8, cursor="hand2"
                  ).grid(row=r, column=0, columnspan=2, pady=15)

    def _gen_6630(self):
        if not self.f6630_desc.get().strip() or not self.f6630_port.get().strip():
            messagebox.showwarning("Missing", "Cell ID and Port are required.")
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
        try:
            result = generate_b4c_6630(d)
            self.b4c_output.delete("1.0", tk.END)
            self.b4c_output.insert("1.0", result)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # --- 6648 ---
    def _build_6648_form(self, form, r):
        tk.Label(form, text="eNB/gNB 6648", font=("Consolas", 11, "bold"),
                 fg=FG_ACCENT, bg=BG_DARK).grid(row=r, column=0, columnspan=2,
                                                  sticky="w", pady=(5, 5))
        r += 1
        self.f6648_desc = self._make_entry(form, r, "Cell ID (digits)"); r += 1
        self.f6648_variant = self._make_combo(form, r, "Deployment Type",
                                               ["ls6", "dual", "tri", "lte_lb"],
                                               "ls6"); r += 1
        self.f6648_alt = self._make_check(form, r, "Use VLAN 302/402 (instead of 301/401)"); r += 1

        r += 1
        tk.Label(form, text="Ports (depends on variant)", font=("Consolas", 10, "bold"),
                 fg=FG_DIM, bg=BG_DARK).grid(row=r, column=0, columnspan=2,
                                              sticky="w", pady=(5, 3))
        r += 1
        self.f6648_s1 = self._make_entry(form, r, "S1 LTE Port"); r += 1
        self.f6648_nr = self._make_entry(form, r, "S1 NR Port (ls6/tri)"); r += 1
        self.f6648_nrca = self._make_entry(form, r, "NR CA Port (ls6/dual/lte_lb)"); r += 1

        r += 1
        tk.Button(form, text="GENERATE 6648 CONFIG", command=self._gen_6648,
                  font=("Consolas", 12, "bold"), bg="#00796b", fg="white",
                  activebackground="#00897b", activeforeground="white",
                  relief="flat", padx=20, pady=8, cursor="hand2"
                  ).grid(row=r, column=0, columnspan=2, pady=15)

    def _gen_6648(self):
        if not self.f6648_desc.get().strip() or not self.f6648_s1.get().strip():
            messagebox.showwarning("Missing", "Cell ID and S1 Port are required.")
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
        if variant in ("ls6",) and not d["port_nr"]:
            messagebox.showwarning("Missing", "S1 NR Port required for LS6.")
            return
        if variant in ("ls6", "dual", "lte_lb") and not d["port_nrca"]:
            messagebox.showwarning("Missing", "NR CA Port required for this variant.")
            return
        if variant == "tri" and not d["port_nr"]:
            messagebox.showwarning("Missing", "S1 NR Port required for Tri-Mode.")
            return
        try:
            result = generate_b4c_6648(d)
            self.b4c_output.delete("1.0", tk.END)
            self.b4c_output.insert("1.0", result)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # --- 6672 ---
    def _build_6672_form(self, form, r):
        tk.Label(form, text="eNB/gNB 6672", font=("Consolas", 11, "bold"),
                 fg=FG_ACCENT, bg=BG_DARK).grid(row=r, column=0, columnspan=2,
                                                  sticky="w", pady=(5, 5))
        r += 1
        self.f6672_desc = self._make_entry(form, r, "Cell ID (digits)"); r += 1
        self.f6672_port = self._make_entry(form, r, "Combined Port (e.g. 1/1/13)"); r += 1
        self.f6672_25g = self._make_check(form, r, "Use 25G speed"); r += 1

        r += 1
        tk.Button(form, text="GENERATE 6672 CONFIG", command=self._gen_6672,
                  font=("Consolas", 12, "bold"), bg="#00796b", fg="white",
                  activebackground="#00897b", activeforeground="white",
                  relief="flat", padx=20, pady=8, cursor="hand2"
                  ).grid(row=r, column=0, columnspan=2, pady=15)

    def _gen_6672(self):
        if not self.f6672_desc.get().strip() or not self.f6672_port.get().strip():
            messagebox.showwarning("Missing", "Cell ID and Port are required.")
            return
        d = {
            "desc_id": self.f6672_desc.get().strip(),
            "port": self.f6672_port.get().strip(),
            "use_25g": self.f6672_25g.get(),
        }
        try:
            result = generate_b4c_6672(d)
            self.b4c_output.delete("1.0", tk.END)
            self.b4c_output.insert("1.0", result)
        except Exception as e:
            messagebox.showerror("Error", str(e))


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
    app = ENSEConfigApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
