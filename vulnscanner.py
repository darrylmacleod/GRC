#pip install python-nmap requests PyQt5
#!/usr/bin/env python3
"""
basic_vuln_scanner_gui.py

A modern GUI-based vulnerability scanner that:
  • Scans a target host/network for open ports and service versions (nmap -sV)
  • Queries the CVE Search API for known vulnerabilities

Requirements:
  pip install python-nmap requests PyQt5
"""

import sys
import nmap
import requests
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QMessageBox
)

CVE_API_BASE = "https://cve.circl.lu/api/search"


def lookup_cves(service: str, version: str):
    """
    Query the CVE Search API for a particular service/version.
    Returns a list of dicts with 'id' and 'summary'.
    """
    try:
        url = f"{CVE_API_BASE}/{service}/{version}"
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        return data.get("results", [])
    except Exception as e:
        return []


class ScannerGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Modern Vulnerability Scanner")
        self.resize(800, 600)

        layout = QVBoxLayout()
        # === Input Section ===
        input_layout = QHBoxLayout()
        input_layout.addWidget(QLabel("Target:"))
        self.target_input = QLineEdit()
        self.target_input.setPlaceholderText("e.g. 192.168.1.5 or 192.168.1.0/24")
        input_layout.addWidget(self.target_input)
        input_layout.addWidget(QLabel("Ports:"))
        self.ports_input = QLineEdit("1-1024")
        input_layout.addWidget(self.ports_input)
        self.scan_button = QPushButton("Scan")
        self.scan_button.clicked.connect(self.perform_scan)
        input_layout.addWidget(self.scan_button)
        layout.addLayout(input_layout)

        # === Output Section ===
        self.output = QTextEdit()
        self.output.setReadOnly(True)
        layout.addWidget(self.output)

        self.setLayout(layout)

    def perform_scan(self):
        target = self.target_input.text().strip()
        ports = self.ports_input.text().strip()
        if not target:
            QMessageBox.warning(self, "Input Error", "Please enter a target.")
            return

        self.output.clear()
        self.output.append(f"Scanning {target} on ports {ports}...")
        nm = nmap.PortScanner()
        try:
            nm.scan(target, ports, arguments='-sV')
        except Exception as e:
            QMessageBox.critical(self, "Scan Error", str(e))
            return

        hosts = nm.all_hosts()
        if not hosts:
            self.output.append("No hosts found or no open ports.")
            return

        for host in hosts:
            self.output.append(f"\nHost: {host}")
            for proto in nm[host].all_protocols():
                for port in sorted(nm[host][proto].keys()):
                    info = nm[host][proto][port]
                    svc  = info.get('name', '')
                    prod = info.get('product', '')
                    ver  = info.get('version', '')

                    line = f"Port {port}/{proto}: {svc} {prod} {ver}"
                    self.output.append(line)

                    if svc and ver:
                        cves = lookup_cves(svc, ver)
                        if cves:
                            self.output.append(" → CVEs:")
                            for c in cves[:5]:
                                self.output.append(f"    • {c.get('id')}: {c.get('summary')}")
                        else:
                            self.output.append(" → No CVEs found.")
                    else:
                        self.output.append(" → Skipping CVE lookup (missing version info).")

        self.output.append("\nScan complete.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    gui = ScannerGUI()
    gui.show()
    sys.exit(app.exec_())
