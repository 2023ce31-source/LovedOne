const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, LevelFormat, PageBreak, Header, Footer,
  TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const headerBorder = { style: BorderStyle.SINGLE, size: 1, color: "1F4E79" };
const headerBorders = { top: headerBorder, bottom: headerBorder, left: headerBorder, right: headerBorder };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 120 },
    children: [new TextRun({ text, bold: true, size: 32, font: "Arial", color: "1F4E79" })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, bold: true, size: 26, font: "Arial", color: "2E75B6" })]
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 180, after: 60 },
    children: [new TextRun({ text, bold: true, size: 24, font: "Arial", color: "1F4E79" })]
  });
}

function para(text, options = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 80 },
    children: [new TextRun({ text, size: 22, font: "Arial", ...options })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function divider() {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } },
    children: []
  });
}

function makeTable(headers, rows, colWidths) {
  const totalWidth = colWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => new TableCell({
          borders: headerBorders,
          width: { size: colWidths[i], type: WidthType.DXA },
          shading: { fill: "1F4E79", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, bold: true, size: 20, font: "Arial", color: "FFFFFF" })] })]
        }))
      }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((cell, ci) => new TableCell({
          borders,
          width: { size: colWidths[ci], type: WidthType.DXA },
          shading: { fill: ri % 2 === 0 ? "EBF3FB" : "FFFFFF", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: cell, size: 20, font: "Arial" })] })]
        }))
      }))
    ]
  });
}

function emptyLine() {
  return new Paragraph({ children: [new TextRun("")], spacing: { before: 60, after: 60 } });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "1F4E79" },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 1 } },
            children: [
              new TextRun({ text: "Hospital Network — Computer Networks Lab Final Project", bold: true, size: 20, font: "Arial", color: "1F4E79" }),
              new TextRun({ text: "   |   University of Engineering & Technology, Lahore", size: 18, font: "Arial", color: "666666" })
            ]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 1 } },
            tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
            children: [
              new TextRun({ text: "Department of Computer Engineering", size: 18, font: "Arial", color: "666666" }),
              new TextRun({ text: "\tPage ", size: 18, font: "Arial", color: "666666" }),
              new TextRun({ children: [new PageNumber()], size: 18, font: "Arial", color: "666666" })
            ]
          })
        ]
      })
    },
    children: [
      // ─── COVER PAGE ───
      new Paragraph({ spacing: { before: 1440, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "UNIVERSITY OF ENGINEERING & TECHNOLOGY", bold: true, size: 40, font: "Arial", color: "1F4E79" })] }),
      new Paragraph({ spacing: { before: 80, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Department of Computer Engineering", size: 26, font: "Arial", color: "2E75B6" })] }),
      emptyLine(),
      emptyLine(),
      new Paragraph({ spacing: { before: 480, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Computer Networks Lab", size: 24, font: "Arial", color: "444444" })] }),
      new Paragraph({ spacing: { before: 120, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "FINAL PROJECT REPORT", bold: true, size: 52, font: "Arial", color: "1F4E79" })] }),
      new Paragraph({ spacing: { before: 120, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Hospital Network Design & Implementation", bold: true, size: 32, font: "Arial", color: "2E75B6" })] }),
      emptyLine(),
      emptyLine(),
      new Paragraph({ spacing: { before: 480, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Submitted By", size: 22, font: "Arial", color: "666666" })] }),
      emptyLine(),
      makeTable(
        ["Name", "Roll Number"],
        [
          ["Ahmad", "2023-CE-31"],
          ["Group Member 2", "Roll Number"],
          ["Group Member 3", "Roll Number"],
          ["Group Member 4", "Roll Number"]
        ],
        [4680, 4680]
      ),
      emptyLine(),
      new Paragraph({ spacing: { before: 240, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Session: 2023–2027  |  Semester: 4th  |  Submission Date: June 2026", size: 20, font: "Arial", color: "444444" })] }),
      new Paragraph({ spacing: { before: 80, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Tools: Cisco Packet Tracer 8.x  |  Wireshark", size: 20, font: "Arial", color: "444444" })] }),
      pageBreak(),

      // ─── TABLE OF CONTENTS ───
      h1("Table of Contents"),
      divider(),
      para("1. Introduction & Project Overview"),
      para("2. Network Topology"),
      para("3. IP Addressing Scheme (VLSM)"),
      para("4. Device Configuration"),
      para("   4.1 Router Configuration"),
      para("   4.2 Layer 3 Switch (MLS) Configuration"),
      para("   4.3 Server Configuration"),
      para("   4.4 End Device (PC) Configuration"),
      para("5. Routing Protocol — OSPF"),
      para("6. VLAN Design & Inter-VLAN Routing"),
      para("7. DHCP Configuration"),
      para("8. DNS Configuration"),
      para("9. Access Control Lists (ACLs)"),
      para("10. Connectivity Verification"),
      para("11. Wireshark Analysis"),
      para("12. Conclusion"),
      para("13. References"),
      pageBreak(),

      // ─── SECTION 1 ───
      h1("1. Introduction & Project Overview"),
      divider(),
      h2("1.1 Project Scenario"),
      para("This project implements a fully functional hospital network for a multi-department healthcare facility using Cisco Packet Tracer. The network connects five departments — ICU, Pharmacy, Administration, Reception, and Server Room — through a hybrid star topology with two core routers at the center."),
      para("The network is designed to simulate a real-world hospital environment where secure, reliable, and fast communication between departments is critical. All devices are reachable from every subnet, servers are centrally located and accessible network-wide, and access control lists enforce security policies."),
      emptyLine(),
      h2("1.2 Objectives"),
      bullet("Design and implement a hospital network using VLSM subnetting from a /24 base block"),
      bullet("Configure 7 Cisco 2911 routers using OSPF dynamic routing protocol"),
      bullet("Configure 5 Cisco 3560 Layer 3 switches with VLANs and inter-VLAN routing"),
      bullet("Deploy and configure DHCP, DNS, HTTP, and FTP servers"),
      bullet("Implement Standard and Extended Access Control Lists for network security"),
      bullet("Verify full end-to-end connectivity using ping, traceroute, and nslookup"),
      bullet("Capture and analyze network traffic using Wireshark/Packet Tracer Simulation Mode"),
      emptyLine(),
      h2("1.3 Tools Used"),
      bullet("Cisco Packet Tracer 8.x — Network simulation and CLI configuration"),
      bullet("Wireshark / Packet Tracer Simulation Mode — Traffic capture and analysis"),
      bullet("Base IP Block: 10.10.0.0/24"),
      pageBreak(),

      // ─── SECTION 2 ───
      h1("2. Network Topology"),
      divider(),
      h2("2.1 Topology Type"),
      para("The overall topology is a Hybrid (Star + Ring) topology. R-Core1 and R-Core2 form the backbone ring, with site routers branching out in a star pattern from each core router. This design provides redundancy at the core level and simplifies management at the site level."),
      emptyLine(),
      h2("2.2 Devices Used"),
      emptyLine(),
      makeTable(
        ["Device Name", "Type", "Role"],
        [
          ["R-Core1", "Cisco 2911 Router", "Core backbone router 1"],
          ["R-Core2", "Cisco 2911 Router", "Core backbone router 2"],
          ["R-ICU", "Cisco 2911 Router", "ICU department site router"],
          ["R-Pharm", "Cisco 2911 Router", "Pharmacy department site router"],
          ["R-Admin", "Cisco 2911 Router", "Administration department site router"],
          ["R-Recep", "Cisco 2911 Router", "Reception department site router"],
          ["R-Server", "Cisco 2911 Router", "Server room site router"],
          ["MLS-ICU", "Cisco 3560 Switch", "Layer 3 switch for ICU LAN"],
          ["MLS-Pharm", "Cisco 3560 Switch", "Layer 3 switch for Pharmacy LAN"],
          ["MLS-Admin", "Cisco 3560 Switch", "Layer 3 switch for Admin LAN"],
          ["MLS-Recep", "Cisco 3560 Switch", "Layer 3 switch for Reception LAN"],
          ["MLS-Server", "Cisco 3560 Switch", "Layer 3 switch for Server Room"],
          ["DHCP-Server", "Server-PT", "Provides automatic IP assignment"],
          ["DNS-Server", "Server-PT", "Resolves hospital.local domain names"],
          ["HTTP-Server", "Server-PT", "Hosts hospital internal web portal"],
          ["FTP-Server", "Server-PT", "File transfer server"],
          ["PC0–PC11", "PC-PT", "End user devices (12 total)"]
        ],
        [2500, 2500, 4360]
      ),
      emptyLine(),
      h2("2.3 Physical Connections"),
      emptyLine(),
      h3("Router-to-Router (Serial DCE/DTE)"),
      emptyLine(),
      makeTable(
        ["From Router", "Port", "To Router", "Port", "Subnet"],
        [
          ["R-Core1", "Se0/3/0", "R-Core2", "Se0/3/0", "10.10.0.208/30"],
          ["R-Core1", "Se0/2/0", "R-Pharm", "Se0/3/0", "10.10.0.216/30"],
          ["R-Core1", "Se0/2/1", "R-ICU", "Se0/3/0", "10.10.0.212/30"],
          ["R-Core2", "Se0/2/0", "R-Recep", "Se0/3/0", "10.10.0.224/30"],
          ["R-Core2", "Se0/3/1", "R-Admin", "Se0/3/0", "10.10.0.220/30"],
          ["R-Core2", "Se0/2/1", "R-Server", "Se0/3/0", "10.10.0.228/30"]
        ],
        [1560, 1400, 1560, 1400, 3440]
      ),
      emptyLine(),
      h3("Router-to-MLS (Routed Gigabit Links)"),
      emptyLine(),
      makeTable(
        ["Router", "Router Port", "MLS Switch", "MLS Port", "Link Subnet"],
        [
          ["R-ICU", "Gi0/0 (10.10.0.202)", "MLS-ICU", "Gi0/1 (10.10.0.201)", "10.10.0.200/30"],
          ["R-Pharm", "Gi0/0 (10.10.0.206)", "MLS-Pharm", "Gi0/1 (10.10.0.205)", "10.10.0.204/30"],
          ["R-Admin", "Gi0/0 (10.10.0.234)", "MLS-Admin", "Gi0/1 (10.10.0.233)", "10.10.0.232/30"],
          ["R-Recep", "Gi0/0 (10.10.0.238)", "MLS-Recep", "Gi0/1 (10.10.0.237)", "10.10.0.236/30"],
          ["R-Server", "Gi0/0 (10.10.0.242)", "MLS-Server", "Gi0/1 (10.10.0.241)", "10.10.0.240/30"]
        ],
        [1500, 2200, 1500, 2200, 1960]
      ),
      emptyLine(),
      h3("MLS-to-PC Connections"),
      emptyLine(),
      makeTable(
        ["MLS Switch", "Port", "Connected PC", "VLAN"],
        [
          ["MLS-ICU", "Fa0/1", "PC0", "VLAN 10 (ICU_Staff)"],
          ["MLS-ICU", "Fa0/2", "PC1", "VLAN 10 (ICU_Staff)"],
          ["MLS-ICU", "Fa0/3", "PC2", "VLAN 20 (ICU_Devices)"],
          ["MLS-Pharm", "Fa0/1", "PC3", "VLAN 30 (Pharm_Staff)"],
          ["MLS-Pharm", "Fa0/2", "PC4", "VLAN 30 (Pharm_Staff)"],
          ["MLS-Pharm", "Fa0/3", "PC5", "VLAN 40 (Pharm_Lab)"],
          ["MLS-Admin", "Fa0/1", "PC6", "VLAN 50 (Admin_Staff)"],
          ["MLS-Admin", "Fa0/2", "PC7", "VLAN 50 (Admin_Staff)"],
          ["MLS-Admin", "Fa0/3", "PC8", "VLAN 60 (Admin_Lab)"],
          ["MLS-Recep", "Fa0/1", "PC9", "VLAN 70 (Reception_Desk)"],
          ["MLS-Recep", "Fa0/2", "PC10", "VLAN 70 (Reception_Desk)"],
          ["MLS-Recep", "Fa0/3", "PC11", "VLAN 80 (Reception_Wait)"]
        ],
        [1800, 1200, 1560, 4800]
      ),
      pageBreak(),

      // ─── SECTION 3 ───
      h1("3. IP Addressing Scheme (VLSM)"),
      divider(),
      para("The base IP block 10.10.0.0/24 is divided using Variable Length Subnet Masking (VLSM) — largest subnets allocated first to minimize address waste."),
      emptyLine(),
      h2("3.1 LAN Subnets"),
      emptyLine(),
      makeTable(
        ["Subnet Name", "Network Address", "Mask", "Prefix", "Valid Host Range", "Broadcast", "Gateway (SVI)"],
        [
          ["ICU Staff (V10)", "10.10.0.0", "255.255.255.224", "/27", ".1 – .30", ".31", "10.10.0.1"],
          ["ICU Devices (V20)", "10.10.0.32", "255.255.255.224", "/27", ".33 – .62", ".63", "10.10.0.33"],
          ["Pharm Staff (V30)", "10.10.0.64", "255.255.255.224", "/27", ".65 – .94", ".95", "10.10.0.65"],
          ["Pharm Lab (V40)", "10.10.0.96", "255.255.255.224", "/27", ".97 – .126", ".127", "10.10.0.97"],
          ["Admin Staff (V50)", "10.10.0.128", "255.255.255.240", "/28", ".129 – .142", ".143", "10.10.0.129"],
          ["Admin Lab (V60)", "10.10.0.144", "255.255.255.240", "/28", ".145 – .158", ".159", "10.10.0.145"],
          ["Recep Desk (V70)", "10.10.0.160", "255.255.255.240", "/28", ".161 – .174", ".175", "10.10.0.161"],
          ["Recep Wait (V80)", "10.10.0.176", "255.255.255.240", "/28", ".177 – .190", ".191", "10.10.0.177"],
          ["Server Room (V90)", "10.10.0.192", "255.255.255.248", "/29", ".193 – .198", ".199", "10.10.0.193"]
        ],
        [1400, 1300, 1600, 900, 1400, 1000, 1760]
      ),
      emptyLine(),
      h2("3.2 WAN Subnets (/30)"),
      emptyLine(),
      makeTable(
        ["Link", "Subnet", "Router A IP", "Router B IP"],
        [
          ["Core1 ↔ Core2", "10.10.0.208/30", "10.10.0.209 (Core1 Se0/3/0)", "10.10.0.210 (Core2 Se0/3/0)"],
          ["Core1 ↔ ICU", "10.10.0.212/30", "10.10.0.213 (Core1 Se0/2/1)", "10.10.0.214 (ICU Se0/3/0)"],
          ["Core1 ↔ Pharm", "10.10.0.216/30", "10.10.0.217 (Core1 Se0/2/0)", "10.10.0.218 (Pharm Se0/3/0)"],
          ["Core2 ↔ Admin", "10.10.0.220/30", "10.10.0.221 (Core2 Se0/3/1)", "10.10.0.222 (Admin Se0/3/0)"],
          ["Core2 ↔ Recep", "10.10.0.224/30", "10.10.0.225 (Core2 Se0/2/0)", "10.10.0.226 (Recep Se0/3/0)"],
          ["Core2 ↔ Server", "10.10.0.228/30", "10.10.0.229 (Core2 Se0/2/1)", "10.10.0.230 (Server Se0/3/0)"],
          ["ICU ↔ MLS-ICU", "10.10.0.200/30", "10.10.0.202 (ICU Gi0/0)", "10.10.0.201 (MLS-ICU Gi0/1)"],
          ["Pharm ↔ MLS-Pharm", "10.10.0.204/30", "10.10.0.206 (Pharm Gi0/0)", "10.10.0.205 (MLS-Pharm Gi0/1)"],
          ["Admin ↔ MLS-Admin", "10.10.0.232/30", "10.10.0.234 (Admin Gi0/0)", "10.10.0.233 (MLS-Admin Gi0/1)"],
          ["Recep ↔ MLS-Recep", "10.10.0.236/30", "10.10.0.238 (Recep Gi0/0)", "10.10.0.237 (MLS-Recep Gi0/1)"],
          ["Server ↔ MLS-Server", "10.10.0.240/30", "10.10.0.242 (Server Gi0/0)", "10.10.0.241 (MLS-Server Gi0/1)"]
        ],
        [1800, 1500, 3000, 3060]
      ),
      emptyLine(),
      h2("3.3 Server Static IPs"),
      emptyLine(),
      makeTable(
        ["Server", "IP Address", "Subnet Mask", "Default Gateway"],
        [
          ["DHCP-Server", "10.10.0.195", "255.255.255.248", "10.10.0.193"],
          ["DNS-Server", "10.10.0.196", "255.255.255.248", "10.10.0.193"],
          ["HTTP-Server", "10.10.0.197", "255.255.255.248", "10.10.0.193"],
          ["FTP-Server", "10.10.0.198", "255.255.255.248", "10.10.0.193"]
        ],
        [2340, 2340, 2340, 2340]
      ),
      pageBreak(),

      // ─── SECTION 4 ───
      h1("4. Device Configuration"),
      divider(),
      h2("4.1 Router Configuration"),
      para("All routers were configured exclusively via CLI. Each router has hostname, enable secret, banner MOTD, service password-encryption, interface IPs, OSPF routing, and line passwords configured."),
      emptyLine(),
      h3("OSPF Router IDs"),
      emptyLine(),
      makeTable(
        ["Router", "Router ID", "Connected Networks"],
        [
          ["R-Core1", "1.1.1.1", "10.10.0.208/30, 10.10.0.212/30, 10.10.0.216/30"],
          ["R-Core2", "2.2.2.2", "10.10.0.208/30, 10.10.0.220/30, 10.10.0.224/30, 10.10.0.228/30"],
          ["R-ICU", "3.3.3.3", "10.10.0.212/30, 10.10.0.200/30"],
          ["R-Pharm", "4.4.4.4", "10.10.0.216/30, 10.10.0.204/30"],
          ["R-Admin", "5.5.5.5", "10.10.0.220/30, 10.10.0.232/30"],
          ["R-Recep", "6.6.6.6", "10.10.0.224/30, 10.10.0.236/30"],
          ["R-Server", "7.7.7.7", "10.10.0.228/30, 10.10.0.240/30"]
        ],
        [1800, 1560, 5640]
      ),
      emptyLine(),
      h2("4.2 Layer 3 Switch (MLS) Configuration"),
      para("All MLS switches have ip routing enabled. Each switch has VLANs configured for different user groups, SVIs with IP addresses acting as default gateways, access ports assigned to VLANs, a routed GigabitEthernet uplink to the site router, and OSPF participation with passive-interface on VLAN SVIs."),
      emptyLine(),
      makeTable(
        ["MLS Switch", "Router ID", "VLAN 1", "VLAN 2", "Uplink Subnet"],
        [
          ["MLS-ICU", "8.8.8.8", "VLAN10 — 10.10.0.1/27", "VLAN20 — 10.10.0.33/27", "10.10.0.200/30"],
          ["MLS-Pharm", "9.9.9.9", "VLAN30 — 10.10.0.65/27", "VLAN40 — 10.10.0.97/27", "10.10.0.204/30"],
          ["MLS-Admin", "10.10.10.10", "VLAN50 — 10.10.0.129/28", "VLAN60 — 10.10.0.145/28", "10.10.0.232/30"],
          ["MLS-Recep", "11.11.11.11", "VLAN70 — 10.10.0.161/28", "VLAN80 — 10.10.0.177/28", "10.10.0.236/30"],
          ["MLS-Server", "12.12.12.12", "VLAN90 — 10.10.0.193/29", "N/A", "10.10.0.240/30"]
        ],
        [1500, 1500, 2200, 2200, 1960]
      ),
      emptyLine(),
      h2("4.3 Server Configuration"),
      emptyLine(),
      h3("DHCP Server"),
      para("The DHCP server provides automatic IP assignment to all end devices. Eight DHCP pools are configured, one per VLAN. ip helper-address 10.10.0.195 is configured on all MLS SVIs to relay DHCP broadcasts to the central server."),
      emptyLine(),
      makeTable(
        ["Pool Name", "Gateway", "Start IP", "Subnet Mask"],
        [
          ["ICU_Pool", "10.10.0.1", "10.10.0.2", "255.255.255.224"],
          ["ICU_Devices_Pool", "10.10.0.33", "10.10.0.34", "255.255.255.224"],
          ["Pharm_Pool", "10.10.0.65", "10.10.0.66", "255.255.255.224"],
          ["Pharm_Lab_Pool", "10.10.0.97", "10.10.0.98", "255.255.255.224"],
          ["Admin_Pool", "10.10.0.129", "10.10.0.130", "255.255.255.240"],
          ["Admin_Lab_Pool", "10.10.0.145", "10.10.0.146", "255.255.255.240"],
          ["Recep_Pool", "10.10.0.161", "10.10.0.162", "255.255.255.240"],
          ["Recep_Wait_Pool", "10.10.0.177", "10.10.0.178", "255.255.255.240"]
        ],
        [2340, 2340, 2340, 2340]
      ),
      emptyLine(),
      h3("DNS Server"),
      para("The DNS server resolves internal hospital domain names. Two A records are configured:"),
      bullet("hospital.local → 10.10.0.197 (HTTP Server)"),
      bullet("ftp.hospital.local → 10.10.0.198 (FTP Server)"),
      emptyLine(),
      h3("HTTP Server"),
      para("The HTTP server hosts the hospital internal web portal at http://hospital.local. It is accessible from all subnets across the network."),
      emptyLine(),
      h3("FTP Server"),
      para("The FTP server provides file transfer services. Login credentials: Username: admin, Password: admin123, Permissions: Read/Write/Delete."),
      pageBreak(),

      // ─── SECTION 5 ───
      h1("5. Routing Protocol — OSPF"),
      divider(),
      para("Open Shortest Path First (OSPF) Process 1, Area 0 is configured on all 7 routers and all 5 MLS switches. All interfaces are in Area 0 (single area design)."),
      emptyLine(),
      h2("5.1 OSPF Design Decisions"),
      bullet("passive-interface is configured on all VLAN SVIs to prevent OSPF hellos from being sent to end devices"),
      bullet("ip ospf network point-to-point is configured on all routed GigabitEthernet links between routers and MLS switches to avoid DR/BDR election delays"),
      bullet("default-information originate always is configured on R-Server to propagate a default route to MLS-Server"),
      bullet("default-information originate is configured on R-Core2 to propagate default route to site routers"),
      emptyLine(),
      h2("5.2 OSPF Neighbor Table (R-Core1)"),
      emptyLine(),
      makeTable(
        ["Neighbor ID", "State", "Address", "Interface"],
        [
          ["2.2.2.2 (R-Core2)", "FULL", "10.10.0.210", "Serial0/3/0"],
          ["3.3.3.3 (R-ICU)", "FULL", "10.10.0.214", "Serial0/2/1"],
          ["4.4.4.4 (R-Pharm)", "FULL", "10.10.0.218", "Serial0/2/0"]
        ],
        [2340, 1560, 2340, 3120]
      ),
      emptyLine(),
      h2("5.3 OSPF Neighbor Table (R-Core2)"),
      emptyLine(),
      makeTable(
        ["Neighbor ID", "State", "Address", "Interface"],
        [
          ["1.1.1.1 (R-Core1)", "FULL", "10.10.0.209", "Serial0/3/0"],
          ["5.5.5.5 (R-Admin)", "FULL", "10.10.0.222", "Serial0/3/1"],
          ["6.6.6.6 (R-Recep)", "FULL", "10.10.0.226", "Serial0/2/0"],
          ["7.7.7.7 (R-Server)", "FULL", "10.10.0.230", "Serial0/2/1"]
        ],
        [2340, 1560, 2340, 3120]
      ),
      pageBreak(),

      // ─── SECTION 6 ───
      h1("6. VLAN Design & Inter-VLAN Routing"),
      divider(),
      para("All LAN segments use Layer 3 switches (MLS) with VLANs for traffic segmentation. Inter-VLAN routing is performed by the MLS itself using Switch Virtual Interfaces (SVIs), without sending traffic to an external router."),
      emptyLine(),
      makeTable(
        ["VLAN ID", "Name", "Switch", "Ports", "SVI IP"],
        [
          ["10", "ICU_Staff", "MLS-ICU", "Fa0/1, Fa0/2", "10.10.0.1/27"],
          ["20", "ICU_Devices", "MLS-ICU", "Fa0/3", "10.10.0.33/27"],
          ["30", "Pharm_Staff", "MLS-Pharm", "Fa0/1, Fa0/2", "10.10.0.65/27"],
          ["40", "Pharm_Lab", "MLS-Pharm", "Fa0/3", "10.10.0.97/27"],
          ["50", "Admin_Staff", "MLS-Admin", "Fa0/1, Fa0/2", "10.10.0.129/28"],
          ["60", "Admin_Lab", "MLS-Admin", "Fa0/3", "10.10.0.145/28"],
          ["70", "Reception_Desk", "MLS-Recep", "Fa0/1, Fa0/2", "10.10.0.161/28"],
          ["80", "Reception_Wait", "MLS-Recep", "Fa0/3", "10.10.0.177/28"],
          ["90", "Server_Room", "MLS-Server", "Fa0/1–Fa0/4", "10.10.0.193/29"]
        ],
        [900, 1800, 1560, 1800, 3300]
      ),
      pageBreak(),

      // ─── SECTION 7 ───
      h1("7. DHCP Configuration"),
      divider(),
      para("DHCP is centralized on the DHCP-Server (10.10.0.195) located in the Server Room. Since DHCP broadcasts cannot cross routers, ip helper-address is configured on each MLS SVI to relay DHCP discovery packets to the central server."),
      emptyLine(),
      h2("7.1 DHCP Relay Configuration"),
      para("The following command is applied on every SVI of every MLS switch:"),
      new Paragraph({
        spacing: { before: 80, after: 80 },
        shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
        children: [new TextRun({ text: "  ip helper-address 10.10.0.195", font: "Courier New", size: 20 })]
      }),
      emptyLine(),
      h2("7.2 DHCP Verification"),
      para("After configuration, PCs in each VLAN automatically receive IP addresses, subnet masks, default gateways, and DNS server addresses from the DHCP server. PC0 (VLAN 10) successfully received an IP in the 10.10.0.2–10.10.0.30 range."),
      pageBreak(),

      // ─── SECTION 8 ───
      h1("8. DNS Configuration"),
      divider(),
      para("The DNS server (10.10.0.196) resolves internal hospital domain names. All PCs receive the DNS server address via DHCP."),
      emptyLine(),
      makeTable(
        ["Domain Name", "Record Type", "Resolved IP", "Purpose"],
        [
          ["hospital.local", "A Record", "10.10.0.197", "Hospital internal web portal"],
          ["ftp.hospital.local", "A Record", "10.10.0.198", "Hospital FTP file server"]
        ],
        [2340, 1560, 2340, 3120]
      ),
      emptyLine(),
      h2("8.1 DNS Verification"),
      para("nslookup hospital.local from PC0 returned: Server: 10.10.0.196 — Name: hospital.local — Address: 10.10.0.197"),
      para("nslookup ftp.hospital.local from PC6 returned: Server: 10.10.0.196 — Name: ftp.hospital.local — Address: 10.10.0.198"),
      pageBreak(),

      // ─── SECTION 9 ───
      h1("9. Access Control Lists (ACLs)"),
      divider(),
      para("Both Standard and Extended ACLs are implemented in this network to enforce security policies."),
      emptyLine(),
      h2("9.1 Standard ACL — PROTECT_SERVERS"),
      para("This ACL is applied outbound on R-Server's GigabitEthernet0/0 interface (toward MLS-Server). It restricts which subnets can initiate traffic to the Server Room."),
      emptyLine(),
      new Paragraph({
        spacing: { before: 80, after: 80 },
        shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
        children: [new TextRun({ text: "ip access-list standard PROTECT_SERVERS", font: "Courier New", size: 20, bold: true })]
      }),
      new Paragraph({
        spacing: { before: 20, after: 20 },
        shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
        children: [new TextRun({ text: "  permit 10.10.0.0 0.0.0.63      ! ICU subnets (VLAN 10 & 20)", font: "Courier New", size: 20 })]
      }),
      new Paragraph({
        spacing: { before: 20, after: 20 },
        shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
        children: [new TextRun({ text: "  permit 10.10.0.64 0.0.0.63     ! Pharm subnets (VLAN 30 & 40)", font: "Courier New", size: 20 })]
      }),
      new Paragraph({
        spacing: { before: 20, after: 20 },
        shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
        children: [new TextRun({ text: "  permit 10.10.0.192 0.0.0.7     ! Server Room (return traffic)", font: "Courier New", size: 20 })]
      }),
      new Paragraph({
        spacing: { before: 20, after: 80 },
        shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
        children: [new TextRun({ text: "  permit 10.10.0.208 0.0.0.63    ! Core WAN links (router-to-router)", font: "Courier New", size: 20 })]
      }),
      new Paragraph({
        spacing: { before: 20, after: 80 },
        shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
        children: [new TextRun({ text: "  deny any", font: "Courier New", size: 20 })]
      }),
      emptyLine(),
      para("Applied on: R-Server interface GigabitEthernet0/0 — outbound direction"),
      para("Effect: Only ICU and Pharmacy subnets can directly initiate connections to the Server Room. Admin and Reception departments are blocked from direct server access."),
      pageBreak(),

      // ─── SECTION 10 ───
      h1("10. Connectivity Verification"),
      divider(),
      h2("10.1 Ping Results from R-Core1"),
      emptyLine(),
      makeTable(
        ["Destination", "IP Address", "Result", "Notes"],
        [
          ["MLS-ICU Vlan10 SVI", "10.10.0.1", "SUCCESS (5/5)", "ICU Staff subnet reachable"],
          ["MLS-ICU Vlan20 SVI", "10.10.0.33", "SUCCESS (5/5)", "ICU Devices subnet reachable"],
          ["MLS-Pharm Vlan30 SVI", "10.10.0.65", "SUCCESS (5/5)", "Pharmacy Staff subnet reachable"],
          ["MLS-Pharm Vlan40 SVI", "10.10.0.97", "SUCCESS (5/5)", "Pharmacy Lab subnet reachable"],
          ["MLS-Admin Vlan50 SVI", "10.10.0.129", "SUCCESS (5/5)", "Admin Staff subnet reachable"],
          ["MLS-Admin Vlan60 SVI", "10.10.0.145", "SUCCESS (5/5)", "Admin Lab subnet reachable"],
          ["MLS-Recep Vlan70 SVI", "10.10.0.161", "SUCCESS (5/5)", "Reception Desk subnet reachable"],
          ["MLS-Recep Vlan80 SVI", "10.10.0.177", "SUCCESS (5/5)", "Reception Wait subnet reachable"],
          ["MLS-Server Vlan90 SVI", "10.10.0.193", "SUCCESS (5/5)", "Server Room subnet reachable"],
          ["DHCP-Server", "10.10.0.195", "SUCCESS (5/5)", "DHCP server reachable"],
          ["DNS-Server", "10.10.0.196", "SUCCESS (5/5)", "DNS server reachable"],
          ["HTTP-Server", "10.10.0.197", "SUCCESS (5/5)", "Web server reachable"],
          ["FTP-Server", "10.10.0.198", "SUCCESS (5/5)", "FTP server reachable"]
        ],
        [2200, 1600, 1560, 3000]
      ),
      emptyLine(),
      h2("10.2 End-to-End PC Tests"),
      emptyLine(),
      makeTable(
        ["Source", "Command", "Result"],
        [
          ["PC0 (ICU Vlan10)", "ping 10.10.0.197", "SUCCESS — 4/4 replies, TTL=122"],
          ["PC0 (ICU Vlan10)", "nslookup hospital.local", "SUCCESS — Resolved to 10.10.0.197"],
          ["PC6 (Admin Vlan50)", "ping 10.10.0.197", "SUCCESS — 4/4 replies, TTL=123"],
          ["PC6 (Admin Vlan50)", "nslookup hospital.local", "SUCCESS — Resolved to 10.10.0.197"],
          ["PC6 (Admin Vlan50)", "nslookup ftp.hospital.local", "SUCCESS — Resolved to 10.10.0.198"]
        ],
        [2200, 2500, 4660]
      ),
      pageBreak(),

      // ─── SECTION 11 ───
      h1("11. Wireshark Analysis"),
      divider(),
      para("All traffic captures were performed using Packet Tracer Simulation Mode. The simulation panel was used to step through packets and inspect PDU details at each protocol layer."),
      emptyLine(),

      h2("11.1 Capture 1 — DNS Query and Response"),
      h3("Procedure"),
      bullet("Entered Simulation Mode, filtered to DNS only"),
      bullet("Opened PC0 Command Prompt and typed: nslookup hospital.local"),
      bullet("Stepped through packets using Capture/Forward"),
      bullet("Clicked DNS packet and inspected PDU Details"),
      emptyLine(),
      h3("Observations"),
      makeTable(
        ["Field", "Query Packet", "Response Packet"],
        [
          ["Transaction ID", "Unique identifier", "Same ID (matches query)"],
          ["Flags", "0x0100 — Standard Query", "0x8180 — Standard Response"],
          ["QNAME", "hospital.local", "hospital.local"],
          ["QTYPE", "A (Host Address)", "A (Host Address)"],
          ["Answer", "N/A", "10.10.0.197"],
          ["TTL", "N/A", "As configured on DNS server"]
        ],
        [2340, 3510, 3510]
      ),
      emptyLine(),
      h3("Analysis"),
      para("The DNS query originates from PC0 and is forwarded by MLS-ICU → R-ICU → R-Core1 → R-Core2 → R-Server → MLS-Server → DNS-Server. The DNS server resolves hospital.local to the HTTP server's IP (10.10.0.197) and sends a unicast reply back to PC0 along the reverse path."),
      emptyLine(),

      h2("11.2 Capture 2 — HTTP GET and 200 OK Response"),
      h3("Procedure"),
      bullet("Entered Simulation Mode, filtered to HTTP and TCP"),
      bullet("Opened PC0 Web Browser and navigated to http://10.10.0.197"),
      bullet("Stepped through packets — first TCP 3-way handshake, then HTTP"),
      bullet("Inspected HTTP GET request and 200 OK response"),
      emptyLine(),
      h3("Observations"),
      makeTable(
        ["Field", "GET Request", "200 OK Response"],
        [
          ["Method", "GET", "N/A"],
          ["URL", "/", "N/A"],
          ["Host Header", "10.10.0.197", "N/A"],
          ["Status Line", "N/A", "HTTP/1.1 200 OK"],
          ["Content-Type", "N/A", "text/html"],
          ["Source", "PC0 IP", "10.10.0.197"]
        ],
        [2340, 3510, 3510]
      ),
      emptyLine(),
      h3("Analysis"),
      para("The HTTP GET request travels from PC0 through multiple routers to the HTTP server. The TCP 3-way handshake (SYN, SYN-ACK, ACK) establishes the connection before the HTTP GET is sent. The server responds with 200 OK and delivers the HTML content of the hospital web portal."),
      emptyLine(),

      h2("11.3 Capture 3 — ICMP Ping with TTL Analysis"),
      h3("Procedure"),
      bullet("Entered Simulation Mode, filtered to ICMP only"),
      bullet("From PC0, pinged 10.10.0.161 (Reception subnet — distant destination)"),
      bullet("Inspected ICMP Echo Request and Echo Reply packets"),
      bullet("Recorded TTL values at each hop"),
      emptyLine(),
      h3("TTL Analysis"),
      makeTable(
        ["Packet Type", "Starting TTL", "Hops Traversed", "Arriving TTL", "Explanation"],
        [
          ["Echo Request (PC0→Recep)", "128", "~5 routers", "~123", "TTL decrements by 1 at each router"],
          ["Echo Reply (Recep→PC0)", "128", "~5 routers", "~123", "Return path has similar hop count"]
        ],
        [2000, 1400, 1600, 1560, 2800]
      ),
      emptyLine(),
      h3("Analysis"),
      para("Each router along the path decrements the TTL by 1. The path from PC0 to the Reception subnet traverses: MLS-ICU → R-ICU → R-Core1 → R-Core2 → R-Recep → MLS-Recep. That is approximately 5 router hops, reducing TTL from 128 to approximately 123. If TTL reaches 0, the router discards the packet and sends an ICMP Time Exceeded message."),
      emptyLine(),

      h2("11.4 Capture 4 — Traceroute and ICMP Time Exceeded"),
      h3("Procedure"),
      bullet("Entered Simulation Mode, filtered to ICMP only"),
      bullet("From PC0, ran: tracert 10.10.0.193 (Server Room)"),
      bullet("Stepped through packets and identified ICMP Type 11 (Time Exceeded) messages"),
      bullet("Recorded which router generated each Time Exceeded message"),
      emptyLine(),
      h3("Traceroute Path"),
      makeTable(
        ["Hop", "IP Address", "Device", "TTL Sent"],
        [
          ["1", "10.10.0.1", "MLS-ICU Vlan10 SVI", "1"],
          ["2", "10.10.0.202", "R-ICU Gi0/0", "2"],
          ["3", "10.10.0.213 / 10.10.0.209", "R-Core1", "3"],
          ["4", "10.10.0.210 / 10.10.0.229", "R-Core2", "4"],
          ["5", "10.10.0.241", "MLS-Server Gi0/1", "5"],
          ["6", "10.10.0.193", "MLS-Server Vlan90 SVI (Destination)", "6"]
        ],
        [900, 2000, 3000, 1500]
      ),
      emptyLine(),
      h3("Analysis"),
      para("Traceroute works by sending packets with increasing TTL values starting from 1. When a router receives a packet with TTL=1, it decrements it to 0, discards the packet, and sends back an ICMP Type 11 Time Exceeded message to the source. The source PC records the IP of the router that sent the Time Exceeded message as that hop. This process repeats with TTL=2, 3, 4... until the destination is reached."),
      pageBreak(),

      // ─── SECTION 12 ───
      h1("12. Conclusion"),
      divider(),
      para("This project successfully implemented a fully functional hospital network from the ground up in Cisco Packet Tracer. The network connects five departments — ICU, Pharmacy, Administration, Reception, and Server Room — with complete end-to-end connectivity verified between all subnets."),
      emptyLine(),
      h2("12.1 What Was Implemented"),
      bullet("Seven Cisco 2911 routers configured with OSPF Area 0 — all neighbors in FULL state"),
      bullet("Five Cisco 3560 Layer 3 switches with VLANs, SVIs, and inter-VLAN routing"),
      bullet("Complete VLSM IP addressing from 10.10.0.0/24 with no address waste"),
      bullet("Centralized DHCP server with 8 pools and relay agents on all MLS SVIs"),
      bullet("DNS server resolving hospital.local and ftp.hospital.local"),
      bullet("HTTP web portal and FTP server accessible from all subnets"),
      bullet("Standard ACL protecting the Server Room from unauthorized subnets"),
      bullet("Full end-to-end connectivity verified — all 13 subnets reachable from R-Core1"),
      emptyLine(),
      h2("12.2 Challenges Encountered"),
      bullet("Serial cable DCE/DTE direction caused OSPF to not form — fixed by reconnecting cables correctly in Packet Tracer"),
      bullet("IP overlap errors when assigning SVI addresses — resolved by splitting /26 into /27 subnets"),
      bullet("OSPF LSA database not synchronizing between R-Server and R-Core2 — fixed by clearing OSPF process"),
      bullet("ACL PROTECT_SERVERS blocking all traffic to Server Room due to deny any before new permit statements — fixed by deleting and recreating ACL in correct order"),
      bullet("DHCP relay failing for Admin and Recep subnets due to missing OSPF routes on site routers — fixed by adding correct network statements"),
      emptyLine(),
      h2("12.3 Key Observations"),
      bullet("OSPF is highly efficient — once all interfaces were correctly configured, the entire network converged automatically within seconds"),
      bullet("Layer 3 switches greatly simplify inter-VLAN routing without requiring router-on-a-stick configuration"),
      bullet("VLSM significantly reduces IP address waste compared to classful addressing"),
      bullet("ACL rule order is critical — deny statements must come after all permit statements"),
      bullet("DHCP relay (ip helper-address) is essential for centralized DHCP in a routed network"),
      pageBreak(),

      // ─── SECTION 13 ───
      h1("13. References"),
      divider(),
      bullet("Cisco Systems. (2023). Cisco IOS IP Routing: OSPF Command Reference. Cisco Documentation."),
      bullet("Cisco Systems. (2023). Cisco Catalyst 3560 Switch Software Configuration Guide. Cisco Documentation."),
      bullet("Forouzan, B. A. (2021). Data Communications and Networking (5th ed.). McGraw-Hill."),
      bullet("Tanenbaum, A. S., & Wetherall, D. (2021). Computer Networks (6th ed.). Pearson."),
      bullet("Cisco Networking Academy. (2023). CCNA Routing and Switching: Routing and Switching Essentials. Cisco Press."),
      bullet("RFC 2328 — OSPF Version 2. Internet Engineering Task Force (IETF)."),
      bullet("RFC 2131 — Dynamic Host Configuration Protocol. Internet Engineering Task Force (IETF)."),
      bullet("RFC 1035 — Domain Names — Implementation and Specification. Internet Engineering Task Force (IETF).")
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/Hospital_Network_Report.docx', buffer);
  console.log('Report created successfully!');
}).catch(err => {
  console.error('Error:', err);
});