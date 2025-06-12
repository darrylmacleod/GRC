import ssl
import socket
from datetime import datetime
import re
import os
import json
import logging
from logging.handlers import RotatingFileHandler
import threading
import csv
import customtkinter as ctk
from tkinter import messagebox, filedialog
import tkinter.ttk as ttk
import tkinter as tk
from openpyxl import Workbook

# === Constants ===
DEFAULT_PORT = 443
LOG_FILE = 'tls_cert_checker.log'
MAX_LOG_SIZE = 5 * 1024 * 1024  # 5 MB
BACKUP_COUNT = 2

# === Logging Setup ===
logger = logging.getLogger("tls_cert_checker")
logger.setLevel(logging.INFO)
handler = RotatingFileHandler(LOG_FILE, maxBytes=MAX_LOG_SIZE, backupCount=BACKUP_COUNT)
formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
handler.setFormatter(formatter)
if not logger.hasHandlers():
    logger.addHandler(handler)

# === Utility Functions ===

def get_cert_expiry(hostname, port=DEFAULT_PORT):
    """
    Returns the expiry date of the TLS certificate for a given hostname and port.
    Raises Exception on failure.
    """
    context = ssl.create_default_context()
    try:
        with socket.create_connection((hostname, port), timeout=5) as sock:
            with context.wrap_socket(sock, server_hostname=hostname) as ssock:
                cert = ssock.getpeercert()
                expiry_str = cert['notAfter']
                expiry_date = datetime.strptime(expiry_str, '%b %d %H:%M:%S %Y %Z')
                return expiry_date
    except ssl.SSLError as ssl_err:
        raise Exception(f"[{hostname}:{port}] SSL error: {ssl_err}")
    except socket.timeout:
        raise Exception(f"[{hostname}:{port}] Connection timed out")
    except socket.gaierror:
        raise Exception(f"[{hostname}:{port}] Hostname not found")
    except Exception as e:
        raise Exception(f"[{hostname}:{port}] Other error: {e}")

def is_valid_hostname(hostname):
    """
    Validates a hostname using a regex.
    """
    pattern = r'^(([a-zA-Z0-9](-?[a-zA-Z0-9])*)\.)+[a-zA-Z]{2,}$'
    return re.match(pattern, hostname)

# === Main GUI Application ===

class TLSCertCheckerApp(ctk.CTk):
    """
    GUI Application for checking TLS certificate expiry dates for multiple hosts.
    """
    def __init__(self):
        super().__init__()
        self.title("TLS Certificate Checker")
        self.geometry("820x560")
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.hosts = []

        # --- Input Frame ---
        input_frame = ctk.CTkFrame(self)
        input_frame.pack(pady=10, padx=10, fill='x')

        self.entry = ctk.CTkEntry(input_frame, width=220, placeholder_text="Hostname")
        self.entry.pack(side='left', padx=5)
        self.entry.bind("<Enter>", lambda e: self.show_tooltip("Enter a valid domain name, e.g. example.com"))

        self.port_entry = ctk.CTkEntry(input_frame, width=100, placeholder_text=f"Port (default {DEFAULT_PORT})")
        self.port_entry.pack(side='left', padx=5)
        self.port_entry.insert(0, str(DEFAULT_PORT))
        self.port_entry.bind("<Enter>", lambda e: self.show_tooltip("Enter a valid TCP port (1-65535)"))

        ctk.CTkButton(input_frame, text="Add Host", command=self.add_host).pack(side='left', padx=5)
        ctk.CTkButton(input_frame, text="Remove Host", command=self.remove_host).pack(side='left', padx=5)

        # --- Host Listbox with Scrollbar ---
        host_list_frame = ctk.CTkFrame(self)
        host_list_frame.pack(padx=10, pady=5, anchor='w', fill='x')
        self.host_listbox = ctk.CTkTextbox(host_list_frame, width=320, height=120)
        self.host_listbox.pack(side='left', fill='y')
        self.host_listbox.configure(state="disabled")
        self.host_scrollbar = ctk.CTkScrollbar(host_list_frame, command=self.host_listbox.yview)
        self.host_listbox['yscrollcommand'] = self.host_scrollbar.set
        self.host_scrollbar.pack(side='right', fill='y')

        # --- Results Table ---
        table_frame = ctk.CTkFrame(self)
        table_frame.pack(pady=10, padx=10, fill='both', expand=True)
        columns = ("Host", "Port", "Expiry", "Days Left", "Status")
        self.ttk_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=10)
        for col in columns:
            self.ttk_tree.heading(col, text=col)
            self.ttk_tree.column(col, width=120 if col == "Host" else 90, anchor="center")
        self.ttk_tree.pack(side="left", fill="both", expand=True)
        self.ttk_tree.bind("<ButtonRelease-1>", self.on_row_click)
        self.table_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.ttk_tree.yview)
        self.ttk_tree.configure(yscrollcommand=self.table_scrollbar.set)
        self.table_scrollbar.pack(side='right', fill='y')

        # --- Progress Bar ---
        self.progress = ctk.CTkProgressBar(self)
        self.progress.pack(pady=(5, 0), padx=10, fill='x')
        self.progress.set(0)

        # --- Action Buttons ---
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(pady=10)
        ctk.CTkButton(button_frame, text="Check Certs", command=self.check_all).pack(side='left', padx=8)
        ctk.CTkButton(button_frame, text="Export CSV", command=self.export_csv).pack(side='left', padx=8)
        ctk.CTkButton(button_frame, text="Export JSON", command=self.export_json).pack(side='left', padx=8)
        ctk.CTkButton(button_frame, text="Export Excel", command=self.export_excel).pack(side='left', padx=8)

        # --- Tooltip ---
        self.tooltip = tk.Label(self, text="", bg="#ffffe0", relief="solid", borderwidth=1, font=("Arial", 9), wraplength=200)
        self.tooltip.place_forget()

    def show_tooltip(self, text):
        self.tooltip.config(text=text)
        self.tooltip.place(x=self.winfo_pointerx() - self.winfo_rootx() + 10,
                           y=self.winfo_pointery() - self.winfo_rooty() + 10)
        self.after(1500, self.tooltip.place_forget)

    def add_host(self):
        """
        Add a (hostname, port) pair to the hosts list after validation.
        """
        host = self.entry.get().strip()
        port_str = self.port_entry.get().strip()
        try:
            port = int(port_str)
            if port <= 0 or port > 65535:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid Port", f"{port_str} is not a valid port number (1-65535)")
            return
        if not is_valid_hostname(host):
            messagebox.showerror("Invalid Host", f"{host} is not a valid hostname")
            return
        if (host, port) not in self.hosts:
            self.hosts.append((host, port))
            self.host_listbox.configure(state="normal")
            self.host_listbox.insert("end", f"{host}:{port}\n")
            self.host_listbox.configure(state="disabled")
            self.entry.delete(0, "end")
            self.port_entry.delete(0, "end")
            self.port_entry.insert(0, str(DEFAULT_PORT))
        else:
            messagebox.showwarning("Duplicate Host", "This host:port is already in the list.")

    def remove_host(self):
        """
        Remove the last host from the hosts list and update the listbox.
        """
        lines = self.host_listbox.get("1.0", "end").strip().split("\n")
        if not lines or lines == ['']:
            messagebox.showinfo("Remove Host", "No hosts to remove.")
            return
        host_port_str = lines[-1]
        if ':' in host_port_str:
            host, port = host_port_str.rsplit(':', 1)
            try:
                port = int(port)
                if (host, port) in self.hosts:
                    self.hosts.remove((host, port))
            except Exception:
                pass
        self.host_listbox.configure(state="normal")
        self.host_listbox.delete("end-2l", "end-1l")
        self.host_listbox.configure(state="disabled")

    def update_progress(self, current, total):
        """
        Update the progress bar in the UI.
        """
        value = (current / total) if total else 0
        self.progress.set(value)
        self.update_idletasks()

    def check_all(self):
        """
        Launch a background thread to check all hosts and update the results table.
        """
        def run_checks():
            self.ttk_tree.delete(*self.ttk_tree.get_children())
            total_hosts = len(self.hosts)
            errors = []
            for i, (host, port) in enumerate(self.hosts):
                try:
                    expiry_date = get_cert_expiry(host, port)
                    days_left = (expiry_date - datetime.utcnow()).days
                    status = "✅ OK" if days_left > 30 else "⚠️ Renew"
                    self.ttk_tree.insert('', "end", values=(host, port, expiry_date.strftime('%Y-%m-%d'), days_left, status), tags=('ok' if days_left > 30 else 'warn'))
                    logger.info(f"Checked {host}:{port} - {status}")
                except Exception as e:
                    self.ttk_tree.insert('', "end", values=(host, port, "ERROR", "N/A", str(e)), tags=('error',))
                    logger.error(f"Error checking {host}:{port} - {str(e)}")
                    errors.append(f"{host}:{port} - {str(e)}")
                self.after(0, self.update_progress, i+1, total_hosts)
            self.after(0, self.update_progress, 0, 1)
            if errors:
                failed = "\n".join(errors)
                messagebox.showwarning("Some hosts failed", f"Failed hosts:\n{failed}")

            # Color rows
            self.ttk_tree.tag_configure('ok', background='#eaffea')
            self.ttk_tree.tag_configure('warn', background='#fff5e5')
            self.ttk_tree.tag_configure('error', background='#ffeaea')

        threading.Thread(target=run_checks, daemon=True).start()

    def export_csv(self):
        """
        Export the results table to a CSV file.
        """
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return
        with open(file_path, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Host', 'Port', 'Expiry', 'Days Left', 'Status'])
            for row_id in self.ttk_tree.get_children():
                writer.writerow(self.ttk_tree.item(row_id)['values'])
        messagebox.showinfo("Exported", f"Results saved to {file_path}")

    def export_json(self):
        """
        Export the results table to a JSON file.
        """
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if not file_path:
            return
        data = []
        for row_id in self.ttk_tree.get_children():
            vals = self.ttk_tree.item(row_id)['values']
            data.append({
                "host": vals[0],
                "port": vals[1],
                "expiry": vals[2],
                "days_left": vals[3],
                "status": vals[4],
            })
        with open(file_path, 'w') as f:
            json.dump(data, f, indent=2)
        messagebox.showinfo("Exported", f"Results saved to {file_path}")

    def export_excel(self):
        """
        Export the results table to an Excel (.xlsx) file.
        """
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "TLS Results"
        # Header
        ws.append(['Host', 'Port', 'Expiry', 'Days Left', 'Status'])
        # Data rows
        for row_id in self.ttk_tree.get_children():
            ws.append(self.ttk_tree.item(row_id)['values'])
        try:
            wb.save(file_path)
            messagebox.showinfo("Exported", f"Results saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Export Failed", f"Could not save Excel file: {e}")

    def on_row_click(self, event):
        """
        Display a tooltip with details on a table row click.
        """
        row_id = self.ttk_tree.identify_row(event.y)
        if row_id:
            vals = self.ttk_tree.item(row_id)['values']
            self.show_tooltip(f"{vals[0]}:{vals[1]}\nExpiry: {vals[2]}\nDays left: {vals[3]}\nStatus: {vals[4]}")

if __name__ == "__main__":
    app = TLSCertCheckerApp()
    app.mainloop()
