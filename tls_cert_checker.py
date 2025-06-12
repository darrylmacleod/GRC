import ssl
import socket
from datetime import datetime
import customtkinter as ctk
from tkinter import messagebox, filedialog
import csv
import re
import threading
import logging
import tkinter.ttk as ttk  # For modern table

logging.basicConfig(filename='tls_cert_checker.log', level=logging.INFO, 
                    format='%(asctime)s [%(levelname)s] %(message)s')

def get_cert_expiry(hostname, port=443):
    context = ssl.create_default_context()
    try:
        with socket.create_connection((hostname, port), timeout=5) as sock:
            with context.wrap_socket(sock, server_hostname=hostname) as ssock:
                cert = ssock.getpeercert()
                expiry_str = cert['notAfter']
                expiry_date = datetime.strptime(expiry_str, '%b %d %H:%M:%S %Y %Z')
                return expiry_date
    except ssl.SSLError as ssl_err:
        raise Exception(f"SSL error: {ssl_err}")
    except socket.timeout:
        raise Exception("Connection timed out")
    except socket.gaierror:
        raise Exception("Hostname not found")
    except Exception as e:
        raise Exception(f"Other error: {e}")

def is_valid_hostname(hostname):
    pattern = r'^(([a-zA-Z0-9](-?[a-zA-Z0-9])*)\.)+[a-zA-Z]{2,}$'
    return re.match(pattern, hostname)

class TLSCertCheckerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("TLS Certificate Checker")
        self.geometry("750x500")
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.hosts = []

        # Input frame
        input_frame = ctk.CTkFrame(self)
        input_frame.pack(pady=10, padx=10, fill='x')

        self.entry = ctk.CTkEntry(input_frame, width=200, placeholder_text="Hostname")
        self.entry.pack(side='left', padx=5)
        self.port_entry = ctk.CTkEntry(input_frame, width=80, placeholder_text="Port (default 443)")
        self.port_entry.pack(side='left', padx=5)
        self.port_entry.insert(0, "443")

        ctk.CTkButton(input_frame, text="Add Host", command=self.add_host).pack(side='left', padx=5)
        ctk.CTkButton(input_frame, text="Remove Host", command=self.remove_host).pack(side='left', padx=5)

        # Host list
        self.host_listbox = ctk.CTkTextbox(self, width=280, height=120)
        self.host_listbox.pack(pady=5, padx=10, anchor='w')
        self.host_listbox.configure(state="disabled")

        # Results Table (using ttk for flexible tables)
        table_frame = ctk.CTkFrame(self)
        table_frame.pack(pady=10, padx=10, fill='both', expand=True)
        self.ttk_tree = ttk.Treeview(table_frame, columns=("Host", "Port", "Expiry", "Days Left", "Status"), show="headings", height=10)
        for col in ("Host", "Port", "Expiry", "Days Left", "Status"):
            self.ttk_tree.heading(col, text=col)
            self.ttk_tree.column(col, width=120 if col == "Host" else 80)
        self.ttk_tree.pack(fill="both", expand=True)

        # Progress bar
        self.progress = ctk.CTkProgressBar(self)
        self.progress.pack(pady=(5, 0), padx=10, fill='x')
        self.progress.set(0)

        # Action buttons
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(pady=10)
        ctk.CTkButton(button_frame, text="Check Certs", command=self.check_all).pack(side='left', padx=8)
        ctk.CTkButton(button_frame, text="Export CSV", command=self.export_csv).pack(side='left', padx=8)

    def add_host(self):
        host = self.entry.get().strip()
        port_str = self.port_entry.get().strip()
        try:
            port = int(port_str)
            if port <= 0 or port > 65535:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid Port", f"{port_str} is not a valid port number")
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
            self.port_entry.insert(0, "443")
        else:
            messagebox.showwarning("Duplicate Host", "This host:port is already in the list.")

    def remove_host(self):
        lines = self.host_listbox.get("1.0", "end").strip().split("\n")
        if not lines or lines == ['']:
            messagebox.showinfo("Remove Host", "No hosts to remove.")
            return
        host_port_str = lines[-1]
        if ':' in host_port_str:
            host, port = host_port_str.rsplit(':', 1)
            port = int(port)
            if (host, port) in self.hosts:
                self.hosts.remove((host, port))
        self.host_listbox.configure(state="normal")
        self.host_listbox.delete("end-2l", "end-1l")
        self.host_listbox.configure(state="disabled")

    def update_progress(self, current, total):
        value = (current / total) if total else 0
        self.progress.set(value)
        self.update_idletasks()

    def check_all(self):
        def run_checks():
            self.ttk_tree.delete(*self.ttk_tree.get_children())
            total_hosts = len(self.hosts)
            for i, (host, port) in enumerate(self.hosts):
                try:
                    expiry_date = get_cert_expiry(host, port)
                    days_left = (expiry_date - datetime.utcnow()).days
                    status = "✅ OK" if days_left > 30 else "⚠️ Renew"
                    self.ttk_tree.insert('', "end", values=(host, port, expiry_date.strftime('%Y-%m-%d'), days_left, status))
                    logging.info(f"Checked {host}:{port} - {status}")
                except Exception as e:
                    self.ttk_tree.insert('', "end", values=(host, port, "ERROR", "N/A", str(e)))
                    logging.error(f"Error checking {host}:{port} - {str(e)}")
                self.update_progress(i+1, total_hosts)
            self.update_progress(0, 1)
        threading.Thread(target=run_checks, daemon=True).start()

    def export_csv(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv")
        if not file_path:
            return
        with open(file_path, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Host', 'Port', 'Expiry', 'Days Left', 'Status'])
            for row_id in self.ttk_tree.get_children():
                writer.writerow(self.ttk_tree.item(row_id)['values'])
        messagebox.showinfo("Exported", f"Results saved to {file_path}")

if __name__ == "__main__":
    app = TLSCertCheckerApp()
    app.mainloop()
