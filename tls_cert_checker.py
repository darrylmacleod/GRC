import ssl
import socket
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import csv

def get_cert_expiry(hostname, port=443):
    context = ssl.create_default_context()
    with socket.create_connection((hostname, port), timeout=5) as sock:
        with context.wrap_socket(sock, server_hostname=hostname) as ssock:
            cert = ssock.getpeercert()
            expiry_str = cert['notAfter']
            expiry_date = datetime.strptime(expiry_str, '%b %d %H:%M:%S %Y %Z')
            return expiry_date

class TLSCertCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TLS Certificate Checker")
        self.hosts = []

        # Domain entry
        self.entry = tk.Entry(root, width=40)
        self.entry.grid(row=0, column=0, padx=5, pady=5)
        tk.Button(root, text="Add Host", command=self.add_host).grid(row=0, column=1)

        # Host list
        self.host_listbox = tk.Listbox(root, height=8, width=40)
        self.host_listbox.grid(row=1, column=0, padx=5, pady=5, columnspan=2)

        # Results table
        self.tree = ttk.Treeview(root, columns=('Host', 'Expiry', 'Days Left', 'Status'), show='headings', height=10)
        for col in self.tree['columns']:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)
        self.tree.grid(row=2, column=0, columnspan=2, pady=10)

        # Buttons
        tk.Button(root, text="Check Certs", command=self.check_all).grid(row=3, column=0, pady=5)
        tk.Button(root, text="Export CSV", command=self.export_csv).grid(row=3, column=1)

    def add_host(self):
        host = self.entry.get().strip()
        if host and host not in self.hosts:
            self.hosts.append(host)
            self.host_listbox.insert(tk.END, host)
            self.entry.delete(0, tk.END)

    def check_all(self):
        self.tree.delete(*self.tree.get_children())
        for host in self.hosts:
            try:
                expiry_date = get_cert_expiry(host)
                days_left = (expiry_date - datetime.utcnow()).days
                status = "✅ OK" if days_left > 30 else "⚠️ Renew"
                self.tree.insert('', tk.END, values=(host, expiry_date.strftime('%Y-%m-%d'), days_left, status))
            except Exception as e:
                self.tree.insert('', tk.END, values=(host, "ERROR", "N/A", str(e)))

    def export_csv(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv")
        if not file_path:
            return
        with open(file_path, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Host', 'Expiry', 'Days Left', 'Status'])
            for row in self.tree.get_children():
                writer.writerow(self.tree.item(row)['values'])
        messagebox.showinfo("Exported", f"Results saved to {file_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TLSCertCheckerApp(root)
    root.mainloop()
