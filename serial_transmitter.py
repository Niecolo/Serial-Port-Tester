import tkinter as tk
from tkinter import ttk, messagebox
import serial
import time
import threading
import logging
from datetime import datetime
import serial.tools.list_ports
import os

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('serial_transmission.log'),
        logging.StreamHandler()
    ]
)

class SerialTransmitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Serial Port Tester")
        self.root.geometry("500x650")
        self.root.resizable(True, True)
        
        # Serial connection variables
        self.ser = None
        self.running = False
        self.thread = None
        self.receive_thread = None

        # Default settings
        self.settings = {
            "com_port": "COM4",
            "baud_rate": 9600,
            "parity": "None",
            "data_bits": 8,
            "stop_bits": "One",
            "base_weight": 5555,
            "mode": "transmit",  # "transmit", "receive", or "command"
            "selected_command": "IP",  # default command
            "custom_command": "",  # custom command input
            "delay_time": 1000  # default delay in milliseconds
        }

        # Define available commands
        self.command_list = [
            "IP", "P", "CP", "SP", "xS", "xP", "Z", "T", "xT", "PU", "xU", "xM", "PV", "Esc R"
        ]

        # Build UI
        self.setup_ui()

    def setup_ui(self):
        # Main frames
        top_frame = tk.Frame(self.root)
        top_frame.pack(fill=tk.X, padx=10, pady=10)

        middle_frame = tk.Frame(self.root)
        middle_frame.pack(fill=tk.X, padx=10, pady=5)

        bottom_frame = tk.Frame(self.root)
        bottom_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Settings section
        settings_frame = tk.LabelFrame(top_frame, text="Serial Configuration", padx=10, pady=10)
        settings_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        # COM Port
        tk.Label(settings_frame, text="COM Port:").grid(row=0, column=0, sticky="w", pady=2)
        self.com_var = tk.StringVar(value=self.settings["com_port"])
        self.com_combo = ttk.Combobox(settings_frame, textvariable=self.com_var, width=12)
        self.com_combo.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        self.com_combo['values'] = self.get_com_ports()

        # Refresh button
        refresh_btn = tk.Button(settings_frame, text="Refresh", command=self.refresh_com_ports, width=10)
        refresh_btn.grid(row=0, column=2, padx=(5, 0), pady=2)

        # Baud Rate
        tk.Label(settings_frame, text="Baud Rate:").grid(row=1, column=0, sticky="w", pady=2)
        baud_values = [300, 600, 1200, 2400, 4800, 9600, 14400, 19200, 38400, 57600, 115200]
        self.baud_var = tk.StringVar(value=str(self.settings["baud_rate"]))
        self.baud_combo = ttk.Combobox(settings_frame, textvariable=self.baud_var, values=baud_values, width=12)
        self.baud_combo.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        # Parity
        tk.Label(settings_frame, text="Parity:").grid(row=2, column=0, sticky="w", pady=2)
        parity_values = ["None", "Even", "Odd"]
        self.parity_var = tk.StringVar(value=self.settings["parity"])
        self.parity_combo = ttk.Combobox(settings_frame, textvariable=self.parity_var, values=parity_values, width=12)
        self.parity_combo.grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        # Data Bits
        tk.Label(settings_frame, text="Data Bits:").grid(row=3, column=0, sticky="w", pady=2)
        data_bits_values = [5, 6, 7, 8]
        self.data_bits_var = tk.StringVar(value=str(self.settings["data_bits"]))
        self.data_bits_combo = ttk.Combobox(settings_frame, textvariable=self.data_bits_var, values=data_bits_values, width=12)
        self.data_bits_combo.grid(row=3, column=1, sticky="ew", padx=5, pady=2)

        # Stop Bits
        tk.Label(settings_frame, text="Stop Bits:").grid(row=4, column=0, sticky="w", pady=2)
        stop_bits_values = ["One", "Two"]
        self.stop_bits_var = tk.StringVar(value=self.settings["stop_bits"])
        self.stop_bits_combo = ttk.Combobox(settings_frame, textvariable=self.stop_bits_var, values=stop_bits_values, width=12)
        self.stop_bits_combo.grid(row=4, column=1, sticky="ew", padx=5, pady=2)

        # Mode selection
        tk.Label(settings_frame, text="Mode:").grid(row=5, column=0, sticky="w", pady=2)
        mode_values = ["transmit", "receive", "command"]
        self.mode_var = tk.StringVar(value=self.settings["mode"])
        self.mode_combo = ttk.Combobox(settings_frame, textvariable=self.mode_var, values=mode_values, width=12)
        self.mode_combo.grid(row=5, column=1, sticky="ew", padx=5, pady=2)

        # Base Weight (only shown when in transmit mode)
        self.base_weight_label = tk.Label(settings_frame, text="Base Weight:")
        self.base_weight_label.grid(row=6, column=0, sticky="w", pady=2)
        self.base_weight_var = tk.StringVar(value=str(self.settings["base_weight"]))
        self.base_weight_entry = tk.Entry(settings_frame, textvariable=self.base_weight_var, width=12)
        self.base_weight_entry.grid(row=6, column=1, sticky="ew", padx=5, pady=2)
        self.base_weight_label.grid_remove()
        self.base_weight_entry.grid_remove()

        # Command Selector (only shown when in command mode)
        self.command_label = tk.Label(settings_frame, text="Command:")
        self.command_label.grid(row=7, column=0, sticky="w", pady=2)
        self.command_var = tk.StringVar(value=self.settings["selected_command"])
        self.command_combo = ttk.Combobox(settings_frame, textvariable=self.command_var, values=self.command_list, width=12)
        self.command_combo.grid(row=7, column=1, sticky="ew", padx=5, pady=2)
        self.command_label.grid_remove()
        self.command_combo.grid_remove()

        # Custom Command (only shown when in command mode)
        self.custom_command_label = tk.Label(settings_frame, text="Custom Command:")
        self.custom_command_label.grid(row=8, column=0, sticky="w", pady=2)
        self.custom_command_var = tk.StringVar(value=self.settings["custom_command"])
        self.custom_command_entry = tk.Entry(settings_frame, textvariable=self.custom_command_var, width=12)
        self.custom_command_entry.grid(row=8, column=1, sticky="ew", padx=5, pady=2)
        self.custom_command_label.grid_remove()
        self.custom_command_entry.grid_remove()

        # Custom Command Dropdown (only shown when in command mode)
        self.custom_dropdown_label = tk.Label(settings_frame, text="Custom Command:")
        self.custom_dropdown_label.grid(row=9, column=0, sticky="w", pady=2)
        self.custom_dropdown_var = tk.StringVar(value="")
        self.custom_dropdown = ttk.Combobox(settings_frame, textvariable=self.custom_dropdown_var, width=12)
        self.custom_dropdown.grid(row=9, column=1, sticky="ew", padx=5, pady=2)
        self.custom_dropdown_label.grid_remove()
        self.custom_dropdown.grid_remove()

        # Delay Time (only shown when in command mode)
        self.delay_label = tk.Label(settings_frame, text="Delay (ms):")
        self.delay_label.grid(row=10, column=0, sticky="w", pady=2)
        self.delay_var = tk.StringVar(value=str(self.settings["delay_time"]))
        self.delay_combo = ttk.Combobox(settings_frame, textvariable=self.delay_var, values=[100, 200, 500, 1000, 2000, 3000, 5000], width=12)
        self.delay_combo.grid(row=10, column=1, sticky="ew", padx=5, pady=2)
        self.delay_label.grid_remove()
        self.delay_combo.grid_remove()

        # Keep Port Open checkbox (only shown when in command mode)
        self.keep_open_var = tk.BooleanVar(value=False)
        self.keep_open_check = tk.Checkbutton(settings_frame, text="Keep Port Open", variable=self.keep_open_var)
        self.keep_open_check.grid(row=11, column=0, columnspan=2, sticky="w", pady=2)
        self.keep_open_check.grid_remove()

        # Configure grid weights
        settings_frame.columnconfigure(1, weight=1)

        # Control buttons frame - now on the right
        control_frame = tk.Frame(top_frame)
        control_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 0), pady=5)

        # Use grid inside control_frame to control layout
        self.start_button = tk.Button(
            control_frame,
            text="Start",
            command=self.start_transmit,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 12, "bold"),
            width=14,
            height=2
        )
        self.start_button.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

        self.stop_button = tk.Button(
            control_frame,
            text="Stop",
            command=self.stop_transmit,
            bg="#f44336",
            fg="white",
            font=("Arial", 12, "bold"),
            width=14,
            height=2
        )
        self.stop_button.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

        # Connect button (renamed from Retry)
        self.connect_button = tk.Button(
            control_frame,
            text="Connect",
            command=self.retry_connect,
            bg="#2196F3",
            fg="white",
            font=("Arial", 12, "bold"),
            width=14,
            height=2
        )
        self.connect_button.grid(row=2, column=0, sticky="ew", padx=5, pady=5)

        # Disconnect button
        self.disconnect_button = tk.Button(
            control_frame,
            text="Disconnect",
            command=self.disconnect_port,
            bg="#9E9E9E",
            fg="white",
            font=("Arial", 12, "bold"),
            width=14,
            height=2
        )
        self.disconnect_button.grid(row=3, column=0, sticky="ew", padx=5, pady=5)

        # Make column expandable
        control_frame.columnconfigure(0, weight=1)

        # Ensure control_frame matches the height of settings_frame
        settings_frame.update_idletasks()
        min_height = settings_frame.winfo_reqheight()
        control_frame.config(height=min_height)
        control_frame.pack_propagate(False)  # Prevent shrinking

        # Status label
        self.status_label = tk.Label(middle_frame, text="Status: Not Connected", fg="blue", font=("Arial", 10))
        self.status_label.pack(pady=5)

        # Note label (initially hidden)
        self.note_label = tk.Label(middle_frame, text="", fg="gray", font=("Arial", 9), wraplength=380)
        self.note_label.pack(pady=2)

        # Update note based on initial mode
        self.update_note()

        # Log display
        log_frame = tk.LabelFrame(bottom_frame, text="Communication Log", padx=10, pady=10)
        log_frame.pack(fill=tk.BOTH, expand=True)

        # Add Clear Log button
        clear_log_btn = tk.Button(
            log_frame,
            text="Clear Log",
            command=self.clear_log,
            bg="#2196F3",
            fg="white",
            font=("Arial", 10, "bold"),
            width=10
        )
        clear_log_btn.pack(anchor="ne", padx=5, pady=5)

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, font=("Consolas", 9), bg="#f0f0f0", height=12)
        scrollbar = tk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Bind events
        self.com_var.trace_add("write", lambda *args: self.update_settings())
        self.baud_var.trace_add("write", lambda *args: self.update_settings())
        self.parity_var.trace_add("write", lambda *args: self.update_settings())
        self.data_bits_var.trace_add("write", lambda *args: self.update_settings())
        self.stop_bits_var.trace_add("write", lambda *args: self.update_settings())
        self.base_weight_var.trace_add("write", lambda *args: self.update_settings())
        self.command_var.trace_add("write", lambda *args: self.update_settings())
        self.custom_command_var.trace_add("write", lambda *args: self.update_settings())
        self.delay_var.trace_add("write", lambda *args: self.update_settings())
        self.mode_var.trace_add("write", lambda *args: self.on_mode_change())

        # Bind custom dropdown selection
        self.custom_dropdown_var.trace_add("write", lambda *args: self.on_custom_dropdown_select())

        # Initial update
        self.update_settings()
        self.toggle_mode()

    def clear_log(self):
        """Clear the communication log"""
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state="disabled")

    def get_com_ports(self):
        """Get list of available COM ports"""
        try:
            ports = [port.device for port in serial.tools.list_ports.comports()]
            return sorted(ports) if ports else ["No COM Ports Found"]
        except Exception:
            return ["COM4"]

    def refresh_com_ports(self):
        """Refresh COM port dropdown"""
        ports = self.get_com_ports()
        self.com_combo['values'] = ports
        if ports and ports != ["No COM Ports Found"]:
            self.com_var.set(ports[0])

    def update_settings(self):
        """Update internal settings from UI with validation"""
        try:
            self.settings["com_port"] = self.com_var.get()
            self.settings["baud_rate"] = int(self.baud_var.get())
            self.settings["parity"] = self.parity_var.get()
            self.settings["data_bits"] = int(self.data_bits_var.get())
            self.settings["stop_bits"] = self.stop_bits_var.get()
            self.settings["mode"] = self.mode_var.get()
            
            if self.settings["mode"] == "transmit":
                base_weight = int(self.base_weight_var.get())
                if base_weight < 0:
                    raise ValueError("Base weight cannot be negative")
                self.settings["base_weight"] = base_weight
            elif self.settings["mode"] == "command":
                self.settings["selected_command"] = self.command_var.get()
                self.settings["custom_command"] = self.custom_command_var.get()
                self.settings["delay_time"] = int(self.delay_var.get())

        except ValueError as e:
            if "invalid literal" in str(e):
                messagebox.showerror("Invalid Input", "Please enter valid numbers for all numeric fields.")
            else:
                messagebox.showerror("Invalid Input", str(e))
            return

    def on_mode_change(self):
        """Handle mode change event"""
        self.update_settings()
        self.toggle_mode()
        self.update_note()

    def update_note(self):
        """Update the note based on current mode"""
        if self.settings["mode"] == "transmit":
            note_text = "Note: Connect to Big Display to transmit a sample Base Weight"
        elif self.settings["mode"] == "receive":
            note_text = "Note: Connect to Indicator with COM Assignment either Demand or Continuous Output"
        elif self.settings["mode"] == "command":
            note_text = "Note: Connect to Indicator with COM Assignment as Demand to send an ascii command to indicator"
        else:
            note_text = ""
        
        self.note_label.config(text=note_text)

    def toggle_mode(self):
        """Toggle visibility of base weight field and command selector based on mode"""
        # Remove all optional widgets first
        self.base_weight_label.grid_forget()
        self.base_weight_entry.grid_forget()
        self.command_label.grid_forget()
        self.command_combo.grid_forget()
        self.custom_command_label.grid_forget()
        self.custom_command_entry.grid_forget()
        self.custom_dropdown_label.grid_forget()
        self.custom_dropdown.grid_forget()
        self.delay_label.grid_forget()
        self.delay_combo.grid_forget()
        self.keep_open_check.grid_forget()

        if self.settings["mode"] == "transmit":
            self.base_weight_label.grid(row=6, column=0, sticky="w", pady=2)
            self.base_weight_entry.grid(row=6, column=1, sticky="ew", padx=5, pady=2)
        elif self.settings["mode"] == "command":
            self.command_label.grid(row=7, column=0, sticky="w", pady=2)
            self.command_combo.grid(row=7, column=1, sticky="ew", padx=5, pady=2)
            self.custom_command_label.grid(row=8, column=0, sticky="w", pady=2)
            self.custom_command_entry.grid(row=8, column=1, sticky="ew", padx=5, pady=2)
            # Add custom dropdown
            self.custom_dropdown_label.grid(row=9, column=0, sticky="w", pady=2)
            self.custom_dropdown.grid(row=9, column=1, sticky="ew", padx=5, pady=2)
            # Populate dropdown with custom commands
            self.populate_custom_dropdown()
            self.delay_label.grid(row=10, column=0, sticky="w", pady=2)
            self.delay_combo.grid(row=10, column=1, sticky="ew", padx=5, pady=2)
            self.keep_open_check.grid(row=11, column=0, columnspan=2, sticky="w", pady=2)

        # Refresh layout
        self.root.update_idletasks()

    def populate_custom_dropdown(self):
        """Populate custom dropdown with recent commands"""
        # This could be enhanced to load from a file or history
        recent_commands = ["CUSTOM1", "CUSTOM2", "CUSTOM3"]  # Example values
        self.custom_dropdown['values'] = recent_commands + [""]  # Empty string for new entry

    def on_custom_dropdown_select(self):
        """Handle custom dropdown selection"""
        selected = self.custom_dropdown_var.get()
        if selected and selected != "":
            self.custom_command_var.set(selected)
            self.custom_dropdown_var.set("")  # Reset dropdown after selection

    def update_buttons(self):
        """Enable/disable start/stop buttons"""
        self.start_button.config(state="normal" if not self.running else "disabled")
        self.stop_button.config(state="normal" if self.running else "disabled")

    def update_status(self, status_text, color="blue"):
        """Update status label"""
        self.status_label.config(text=f"Status: {status_text}", fg=color)

    def log_message(self, message, level="INFO"):
        """Log and display message in GUI"""
        timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        full_msg = f"[{timestamp}] {level}: {message}"

        def _update_gui():
            self.log_text.config(state="normal")
            self.log_text.insert(tk.END, full_msg + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state="disabled")
            if level == "ERROR":
                logging.error(message)
            elif level == "WARNING":
                logging.warning(message)
            else:
                logging.info(message)

        self.root.after(0, _update_gui)

    def is_port_available(self, port_name):
        """Check if a port is available"""
        try:
            ser = serial.Serial(port_name, timeout=1)
            ser.close()
            return True
        except:
            return False

    def open_serial_port(self):
        """Open serial port with current settings - with better error handling"""
        try:
            parity_map = {"None": serial.PARITY_NONE, "Even": serial.PARITY_EVEN, "Odd": serial.PARITY_ODD}
            stop_bits_map = {"One": serial.STOPBITS_ONE, "Two": serial.STOPBITS_TWO}

            self.ser = serial.Serial(
                port=self.settings["com_port"],
                baudrate=self.settings["baud_rate"],
                parity=parity_map.get(self.settings["parity"], serial.PARITY_NONE),
                bytesize=self.settings["data_bits"],
                stopbits=stop_bits_map.get(self.settings["stop_bits"], serial.STOPBITS_ONE),
                timeout=1,
                write_timeout=1  # Added write timeout
            )
            
            # Test if we can actually write to the port
            try:
                self.ser.write(b"")  # Test write
                self.ser.flush()
            except Exception as e:
                raise Exception(f"Port test failed: {e}")
                
            self.log_message(f"SUCCESS: Opened {self.settings['com_port']} at {self.settings['baud_rate']} baud")
            self.update_status("Connected", "green")
            return True
            
        except PermissionError:
            error_msg = f"Port '{self.settings['com_port']}' is in use by another program. Please close any other applications using this port."
            messagebox.showerror("Port Access Denied", error_msg)
            self.log_message(f"ERROR: Port access denied - {error_msg}", "ERROR")
            self.update_status("Connection Failed", "red")
            return False
            
        except serial.SerialException as e:
            error_msg = f"Serial error: {e}"
            messagebox.showerror("Serial Error", error_msg)
            self.log_message(f"ERROR: Serial exception - {error_msg}", "ERROR")
            self.update_status("Connection Failed", "red")
            return False
            
        except Exception as e:
            error_msg = f"Unexpected error: {e}"
            messagebox.showerror("Connection Error", error_msg)
            self.log_message(f"ERROR: Unexpected error - {error_msg}", "ERROR")
            self.update_status("Connection Failed", "red")
            return False

    def close_serial_port(self):
        """Close serial port safely"""
        if self.ser and self.ser.is_open:
            try:
                self.ser.close()
                self.log_message("Port closed successfully")
                self.update_status("Not Connected", "blue")
            except Exception as e:
                self.log_message(f"Error closing port: {e}", "ERROR")

    def transmit_loop(self):
        """Continuously send payload in background thread"""
        while self.running:
            try:
                if self.settings["mode"] == "transmit":
                    base_weight = self.settings["base_weight"]
                    weight_str = f"{base_weight:06d}"
                    reversed_digits = weight_str[::-1]
                    payload_str = f"={reversed_digits}".strip()
                    # For transmit mode, send as string (not ASCII bytes)
                    payload_bytes = payload_str.encode('ascii')
                elif self.settings["mode"] == "command":
                    # Check if custom command is provided
                    if self.settings["custom_command"]:
                        payload_str = self.settings["custom_command"]
                    else:
                        payload_str = self.settings["selected_command"]
                    # Convert each character to its ASCII byte value and add CR/LF
                    payload_bytes = bytes(ord(c) for c in payload_str.upper()) + b'\r\n'
                else:
                    # Should not happen, but just in case
                    payload_bytes = b""

                bytes_written = self.ser.write(payload_bytes)
                self.ser.flush()

                if bytes_written == 0:
                    self.log_message("No bytes written - possible serial issue", "WARNING")
                else:
                    # Show the actual string sent for transmit mode, bytes for command mode
                    if self.settings["mode"] == "transmit":
                        self.log_message(f"Sent: '{payload_str}' ({bytes_written} bytes)")
                    else:
                        self.log_message(f"Sent: '{payload_str}' → bytes {list(payload_bytes)} ({bytes_written} bytes)")

                time.sleep(0.2)

            except serial.SerialException as e:
                self.log_message(f"Serial error: {e}", "ERROR")
                self.running = False
                break
            except Exception as e:
                self.log_message(f"Unexpected error: {e}", "ERROR")
                self.running = False
                break

    def receive_loop(self):
        """Continuously read data from serial port"""
        while self.running:
            try:
                if self.ser and self.ser.in_waiting > 0:
                    data = self.ser.readline()
                    try:
                        decoded_data = data.decode('ascii', errors='ignore').strip()
                        if decoded_data:
                            self.log_message(f"Received: '{decoded_data}' ({len(data)} bytes)")
                    except Exception as decode_error:
                        self.log_message(f"Received raw: {data} (decode error: {decode_error})", "WARNING")
                time.sleep(0.1)
                
            except serial.SerialException as e:
                self.log_message(f"Serial error: {e}", "ERROR")
                self.running = False
                break
            except Exception as e:
                self.log_message(f"Unexpected error: {e}", "ERROR")
                self.running = False
                break

    def start_transmit(self):
        """Start transmission/reception"""
        # Check if port is already open
        if self.ser and self.ser.is_open:
            # Port is already open, proceed with sending
            self.running = True
            self.update_buttons()
            
            if self.settings["mode"] == "transmit":
                self.thread = threading.Thread(target=self.transmit_loop, daemon=True)
                self.thread.start()
                self.log_message("Transmission started.")
            elif self.settings["mode"] == "command":
                # Send command once
                self.send_single_command_with_delay()
            else:
                self.receive_thread = threading.Thread(target=self.receive_loop, daemon=True)
                self.receive_thread.start()
                self.log_message("Reception started.")
            return
        
        # If port is not open, try to open it
        if not self.is_port_available(self.settings["com_port"]):
            messagebox.showerror("Port Unavailable", 
                               f"Port {self.settings['com_port']} is currently in use.\n"
                               "Please close other applications using this port.")
            self.log_message(f"ERROR: Port {self.settings['com_port']} is unavailable", "ERROR")
            return

        if not self.open_serial_port():
            return

        self.running = True
        self.update_buttons()

        if self.settings["mode"] == "transmit":
            self.thread = threading.Thread(target=self.transmit_loop, daemon=True)
            self.thread.start()
            self.log_message("Transmission started.")
        elif self.settings["mode"] == "command":
            # Send command once
            self.send_single_command_with_delay()
        else:
            self.receive_thread = threading.Thread(target=self.receive_loop, daemon=True)
            self.receive_thread.start()
            self.log_message("Reception started.")

    def send_single_command_with_delay(self):
        """Send selected command once, wait for delay, then close port"""
        try:
            # Check if custom command is provided
            if self.settings["custom_command"]:
                payload_str = self.settings["custom_command"]
            else:
                payload_str = self.settings["selected_command"]
            
            # Convert each character to its ASCII byte value and add CR/LF
            payload_bytes = bytes(ord(c) for c in payload_str.upper()) + b'\r\n'
            
            bytes_written = self.ser.write(payload_bytes)
            self.ser.flush()
            self.log_message(f"Sent command: '{payload_str}' → bytes {list(payload_bytes)} ({bytes_written} bytes)")
            
            # If "Keep Port Open" is checked, don't close
            if self.keep_open_var.get():
                self.log_message("Port kept open as requested.")
                self.running = False
                self.update_buttons()
            else:
                # Use configured delay
                delay_ms = self.settings["delay_time"]
                self.root.after(delay_ms, self._close_port_after_delay)
                
        except Exception as e:
            self.log_message(f"Failed to send command: {e}", "ERROR")
            self._close_port_after_delay()  # Clean up even on error

    def _close_port_after_delay(self):
        """Helper to close port after delay"""
        self.running = False
        self.update_buttons()
        self.close_serial_port()
        self.log_message("Command sent and port closed after delay.")

    def retry_connect(self):
        """Attempt to reconnect to the serial port"""
        if self.ser and self.ser.is_open:
            self.close_serial_port()
        
        # Wait a moment before retrying
        self.root.after(500, self._attempt_reconnect)
    
    def _attempt_reconnect(self):
        """Helper to attempt reconnection"""
        if self.open_serial_port():
            self.log_message("Successfully reconnected to serial port.")
        else:
            self.log_message("Reconnection failed. Please check port availability.")

    def disconnect_port(self):
        """Manually disconnect from serial port"""
        if self.ser and self.ser.is_open:
            self.close_serial_port()
            self.running = False
            self.update_buttons()
            self.log_message("Manually disconnected from serial port.")
        else:
            messagebox.showinfo("Already Disconnected", "Port is already closed.")

    def stop_transmit(self):
        """Stop transmission/reception"""
        self.running = False
        if self.thread and self.thread.is_alive():
            self.thread.join(timeout=3)
            if self.thread.is_alive():
                self.log_message("Transmission thread did not stop gracefully.", "WARNING")
        if self.receive_thread and self.receive_thread.is_alive():
            self.receive_thread.join(timeout=3)
            if self.receive_thread.is_alive():
                self.log_message("Reception thread did not stop gracefully.", "WARNING")
        self.close_serial_port()
        self.log_message("Transmission/reception stopped by user.")
        self.update_buttons()

    def on_closing(self):
        """Cleanup on window close"""
        self.stop_transmit()
        self.root.destroy()

def main():
    root = tk.Tk()
    app = SerialTransmitterApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()