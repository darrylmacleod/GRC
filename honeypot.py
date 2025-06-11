# pip install PyQt5

import sys
import socket
import threading
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget, 
                            QLabel, QPushButton, QTextEdit, QLineEdit, 
                            QHBoxLayout, QListWidget, QTabWidget)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont, QColor, QTextCursor

class HoneypotGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Python Honeypot")
        self.setGeometry(100, 100, 800, 600)
        
        # Honeypot variables
        self.running = False
        self.ports = [22, 80, 443, 3389]  # Default ports to monitor
        self.sockets = []
        self.log_entries = []
        
        # Create main widget and layout
        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)
        self.layout = QVBoxLayout(self.main_widget)
        
        # Create tabs
        self.tabs = QTabWidget()
        self.layout.addWidget(self.tabs)
        
        # Create dashboard tab
        self.create_dashboard_tab()
        
        # Create configuration tab
        self.create_configuration_tab()
        
        # Create logs tab
        self.create_logs_tab()
        
        # Status bar
        self.status_bar = self.statusBar()
        self.status_bar.showMessage("Ready")
        
        # Start update timer
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.update_ui)
        self.update_timer.start(1000)  # Update every second
        
    def create_dashboard_tab(self):
        """Create the dashboard tab with statistics"""
        dashboard_tab = QWidget()
        self.tabs.addTab(dashboard_tab, "Dashboard")
        
        layout = QVBoxLayout(dashboard_tab)
        
        # Title
        title = QLabel("Honeypot Dashboard")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Stats widgets
        stats_layout = QHBoxLayout()
        
        # Connections today
        self.connections_today = QLabel("0")
        self.connections_today.setFont(QFont("Arial", 24))
        stats_layout.addWidget(self.create_stat_box("Connections Today", self.connections_today))
        
        # Active ports
        self.active_ports = QLabel(str(len(self.ports)))
        self.active_ports.setFont(QFont("Arial", 24))
        stats_layout.addWidget(self.create_stat_box("Active Ports", self.active_ports))
        
        # Last connection
        self.last_connection = QLabel("Never")
        self.last_connection.setFont(QFont("Arial", 24))
        stats_layout.addWidget(self.create_stat_box("Last Connection", self.last_connection))
        
        layout.addLayout(stats_layout)
        
        # Recent activity list
        self.activity_list = QListWidget()
        self.activity_list.setFont(QFont("Arial", 10))
        layout.addWidget(QLabel("Recent Activity:"))
        layout.addWidget(self.activity_list)
        
        # Control buttons
        button_layout = QHBoxLayout()
        self.start_button = QPushButton("Start Honeypot")
        self.start_button.clicked.connect(self.start_honeypot)
        self.start_button.setStyleSheet("background-color: #4CAF50; color: white;")
        button_layout.addWidget(self.start_button)
        
        self.stop_button = QPushButton("Stop Honeypot")
        self.stop_button.clicked.connect(self.stop_honeypot)
        self.stop_button.setStyleSheet("background-color: #f44336; color: white;")
        self.stop_button.setEnabled(False)
        button_layout.addWidget(self.stop_button)
        
        layout.addLayout(button_layout)
    
    def create_stat_box(self, title, value_widget):
        """Create a styled stat box"""
        box = QWidget()
        box_layout = QVBoxLayout(box)
        
        title_label = QLabel(title)
        title_label.setAlignment(Qt.AlignCenter)
        box_layout.addWidget(title_label)
        
        value_widget.setAlignment(Qt.AlignCenter)
        box_layout.addWidget(value_widget)
        
        box.setStyleSheet("""
            QWidget {
                border: 1px solid #ccc;
                border-radius: 5px;
                padding: 10px;
                background-color: #f9f9f9;
            }
        """)
        
        return box
    
    def create_configuration_tab(self):
        """Create the configuration tab"""
        config_tab = QWidget()
        self.tabs.addTab(config_tab, "Configuration")
        
        layout = QVBoxLayout(config_tab)
        
        # Port configuration
        layout.addWidget(QLabel("Ports to Monitor (comma separated):"))
        self.ports_input = QLineEdit(",".join(map(str, self.ports)))
        layout.addWidget(self.ports_input)
        
        # Save button
        save_button = QPushButton("Save Configuration")
        save_button.clicked.connect(self.save_configuration)
        save_button.setStyleSheet("background-color: #2196F3; color: white;")
        layout.addWidget(save_button)
        
        # Add some spacing
        layout.addStretch()
        
        # Information
        info = QLabel("""
            <b>Honeypot Information:</b><br><br>
            This honeypot will listen on the specified ports and log all connection attempts.<br>
            No actual services are running - it just logs the attempts.<br><br>
            Common ports to monitor:<br>
            - 22 (SSH)<br>
            - 80 (HTTP)<br>
            - 443 (HTTPS)<br>
            - 3389 (RDP)<br>
            - 5900 (VNC)
        """)
        info.setWordWrap(True)
        layout.addWidget(info)
    
    def create_logs_tab(self):
        """Create the logs tab"""
        logs_tab = QWidget()
        self.tabs.addTab(logs_tab, "Logs")
        
        layout = QVBoxLayout(logs_tab)
        
        # Log display
        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)
        self.log_display.setFont(QFont("Courier New", 10))
        layout.addWidget(self.log_display)
        
        # Log controls
        controls_layout = QHBoxLayout()
        
        clear_button = QPushButton("Clear Logs")
        clear_button.clicked.connect(self.clear_logs)
        controls_layout.addWidget(clear_button)
        
        export_button = QPushButton("Export Logs")
        export_button.clicked.connect(self.export_logs)
        controls_layout.addWidget(export_button)
        
        layout.addLayout(controls_layout)
    
    def save_configuration(self):
        """Save the honeypot configuration"""
        try:
            ports_text = self.ports_input.text()
            new_ports = [int(port.strip()) for port in ports_text.split(",") if port.strip().isdigit()]
            
            if not new_ports:
                self.status_bar.showMessage("Error: No valid ports specified", 5000)
                return
            
            if self.running:
                self.stop_honeypot()
                
            self.ports = new_ports
            self.active_ports.setText(str(len(self.ports)))
            self.status_bar.showMessage("Configuration saved successfully", 5000)
            
            if self.running:
                self.start_honeypot()
                
        except Exception as e:
            self.status_bar.showMessage(f"Error: {str(e)}", 5000)
    
    def start_honeypot(self):
        """Start the honeypot"""
        if self.running:
            return
            
        try:
            for port in self.ports:
                thread = threading.Thread(target=self.listen_on_port, args=(port,), daemon=True)
                thread.start()
            
            self.running = True
            self.start_button.setEnabled(False)
            self.stop_button.setEnabled(True)
            self.status_bar.showMessage(f"Honeypot running on ports: {', '.join(map(str, self.ports))}")
            
            self.log_message("Honeypot started", "SYSTEM")
            
        except Exception as e:
            self.status_bar.showMessage(f"Error starting honeypot: {str(e)}", 5000)
    
    def stop_honeypot(self):
        """Stop the honeypot"""
        if not self.running:
            return
            
        self.running = False
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.status_bar.showMessage("Honeypot stopped")
        
        self.log_message("Honeypot stopped", "SYSTEM")
    
    def listen_on_port(self, port):
        """Listen for connections on a specific port"""
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                s.bind(('0.0.0.0', port))
                s.listen(5)
                
                while self.running:
                    conn, addr = s.accept()
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    ip = addr[0]
                    
                    # Log the connection
                    message = f"Connection attempt on port {port} from {ip}"
                    self.log_message(message, "CONNECTION", color=QColor(255, 0, 0))
                    
                    # Add to recent activity
                    self.log_entries.append({
                        'timestamp': timestamp,
                        'type': 'CONNECTION',
                        'message': message,
                        'port': port,
                        'ip': ip
                    })
                    
                    # Close the connection immediately
                    conn.close()
                    
            except Exception as e:
                if self.running:  # Only log if we didn't stop intentionally
                    self.log_message(f"Error on port {port}: {str(e)}", "ERROR")
    
    def log_message(self, message, msg_type, color=QColor(0, 0, 0)):
        """Add a message to the log"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] [{msg_type}] {message}"
        
        # Add to log display with color
        self.log_display.setTextColor(color)
        self.log_display.append(log_entry)
        
        # Auto-scroll to bottom
        self.log_display.moveCursor(QTextCursor.End)
    
    def clear_logs(self):
        """Clear the log display"""
        self.log_display.clear()
        self.log_message("Logs cleared", "SYSTEM")
    
    def export_logs(self):
        """Export logs to a file"""
        # In a real app, you would implement file saving here
        self.status_bar.showMessage("Export functionality would be implemented here", 5000)
    
    def update_ui(self):
        """Update the UI elements"""
        if not self.log_entries:
            return
            
        # Update connections today
        today = datetime.now().strftime("%Y-%m-%d")
        today_connections = sum(1 for entry in self.log_entries 
                              if entry['type'] == 'CONNECTION' 
                              and entry['timestamp'].startswith(today))
        self.connections_today.setText(str(today_connections))
        
        # Update last connection
        last_connection = next((entry for entry in reversed(self.log_entries) 
                              if entry['type'] == 'CONNECTION'), None)
        if last_connection:
            self.last_connection.setText(last_connection['timestamp'])
        
        # Update recent activity list
        self.activity_list.clear()
        for entry in list(reversed(self.log_entries))[:10]:  # Show last 10 entries
            if entry['type'] == 'CONNECTION':
                self.activity_list.addItem(f"{entry['timestamp']} - {entry['ip']} on port {entry['port']}")

def main():
    app = QApplication(sys.argv)
    
    # Set modern style
    app.setStyle('Fusion')
    
    # Create and show the main window
    window = HoneypotGUI()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
