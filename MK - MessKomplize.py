import serial
import serial.tools.list_ports
import pyautogui
import tkinter as tk
from tkinter import ttk
import threading
import re
import ctypes
import os
import datetime
import time
import glob
import json
import random

# --- Hilfsklasse für die Tooltips (Mouseover-Texte) ---
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.id = None
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)

    def enter(self, event=None):
        self.id = self.widget.after(600, self.showtip) # 600ms warten vor Anzeige

    def leave(self, event=None):
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None
        self.hidetip()

    def showtip(self, event=None):
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        lbl = tk.Label(tw, text=self.text, justify=tk.LEFT, background="#ffffe0", relief=tk.SOLID, borderwidth=1, font=("Arial", 9))
        lbl.pack(ipadx=3, ipady=3)

    def hidetip(self):
        if self.tipwindow:
            self.tipwindow.destroy()
            self.tipwindow = None


class RoundedButton(tk.Canvas):
    def __init__(self, parent, command=None, text="", textvariable=None, width=160, height=48,
                 radius=14, bg="lightgrey", fg="black", font=("Arial", 10, "bold"), state="normal"):
        super().__init__(parent, width=width, height=height, highlightthickness=0, bd=0, relief="flat")
        self._command = command
        self._text = text
        self._textvariable = textvariable
        self._width = width
        self._height = height
        self._radius = max(6, min(radius, width // 2, height // 2))
        self._bg = bg
        self._fg = fg
        self._font = font
        self._state = state
        self._trace_id = None

        if self._textvariable is not None:
            self._trace_id = self._textvariable.trace_add("write", lambda *args: self._redraw())

        self.bind("<Button-1>", self._on_click)
        self._redraw()

    def _on_click(self, event=None):
        if self._state != "disabled" and self._command:
            self._command()

    def _current_text(self):
        if self._textvariable is not None:
            return self._textvariable.get()
        return self._text

    def _colors(self):
        if self._state == "disabled":
            return "#cfcfcf", "#808080"
        return self._bg, self._fg

    def _redraw(self):
        self.delete("all")
        fill_color, text_color = self._colors()
        w = self._width
        h = self._height
        r = self._radius

        # Rounded rectangle as polygon with smooth corners.
        points = [
            r, 0,
            w - r, 0,
            w, 0,
            w, r,
            w, h - r,
            w, h,
            w - r, h,
            r, h,
            0, h,
            0, h - r,
            0, r,
            0, 0,
        ]
        self.create_polygon(points, smooth=True, splinesteps=24, fill=fill_color, outline=fill_color)
        self.create_text(w // 2, h // 2, text=self._current_text(), fill=text_color, font=self._font)

    def config(self, **kwargs):
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        if "text" in kwargs:
            self._text = kwargs.pop("text")
        if "textvariable" in kwargs:
            new_textvariable = kwargs.pop("textvariable")
            if self._textvariable is not None and self._trace_id is not None:
                try:
                    self._textvariable.trace_remove("write", self._trace_id)
                except Exception:
                    pass
            self._textvariable = new_textvariable
            self._trace_id = None
            if self._textvariable is not None:
                self._trace_id = self._textvariable.trace_add("write", lambda *args: self._redraw())
        if "bg" in kwargs:
            self._bg = kwargs.pop("bg")
        if "fg" in kwargs:
            self._fg = kwargs.pop("fg")
        if "font" in kwargs:
            self._font = kwargs.pop("font")
        if "state" in kwargs:
            self._state = kwargs.pop("state")
        if kwargs:
            super().config(**kwargs)
        self._redraw()

    configure = config


class MessKomplizeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MK - MessKomplize Version 1.0")
        self.root.geometry("620x850") # Etwas größer für alle Optionen
        
        self.serial_port = None
        self.is_running = False
        self.counter = 0
        self.settings_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "messkomplize_settings.json")
        self.last_successful_port = "COM1"
        self.last_external_hwnd = None
        self.window_tracking_active = False
        
        # --- Standard-Variablen für Einstellungen ---
        self.port_var = tk.StringVar(value="COM1")
        self.baud_var = tk.StringVar(value="9600")
        self.databits_var = tk.StringVar(value="7")
        self.parity_var = tk.StringVar(value="Odd")
        self.stopbits_var = tk.StringVar(value="1")
        
        self.name_prog1 = tk.StringVar(value="Aufschluss")
        self.name_prog2 = tk.StringVar(value="WGH 1")
        self.name_prog3 = tk.StringVar(value="WGH 2")
        self.current_program = 1 
        self.last_measurement_var = tk.StringVar(value="----")
        self.mini_program_var = tk.StringVar(value=self.name_prog1.get())
        
        # Neue & Alte Optionen (Standard: Aus)
        self.counter_var = tk.BooleanVar(value=False)
        self.auto_reconnect_var = tk.BooleanVar(value=False)
        self.plausi_var = tk.BooleanVar(value=False)
        self.backup_var = tk.BooleanVar(value=True)
        
        # V1.0 Exklusiv-Optionen
        self.auto_save_var = tk.BooleanVar(value=False)
        self.auto_save_x_var = tk.IntVar(value=10) # Alle 10 Messungen
        
        self.plausi2_var = tk.BooleanVar(value=False)
        self.plausi2_limit_var = tk.StringVar(value="100") # Freie Grenze: bis 3 Stellen vor Komma, bis 4 nach Komma
        
        self.mini_mode_var = tk.BooleanVar(value=False)
        self.log_clean_var = tk.BooleanVar(value=False)
        
        # Datenformat-Optionen
        self.dot_comma_var = tk.BooleanVar(value=True)  # Standardmäßig aktiv
        self.unit_var = tk.BooleanVar(value=False)
        self.fixed_decimals_var = tk.BooleanVar(value=False)
        self.decimal_places_var = tk.IntVar(value=4)
        self.test_mode_var = tk.BooleanVar(value=False)
        self.test_display_var = tk.StringVar(value="----")

        self.load_settings()
        
        # UI Aufbauen
        self.setup_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(300, self.start_window_tracking)
        self.root.after(1000, self.auto_start_connection)
        
        # Wenn Log-Aufräumer aktiv, direkt beim Start einmal aufräumen
        self.root.after(2000, self.clean_old_logs)

    def setup_ui(self):
        # Mini-Mode Frame (versteckt beim Start)
        self.mini_frame = tk.Frame(self.root, bg="#333333")
        self.lbl_mini_status = tk.Label(self.mini_frame, text="MK - MessKomplize: Getrennt", font=("Arial", 9, "bold"), fg="white", bg="#333333")
        self.lbl_mini_status.pack(pady=(6, 2))
        self.lbl_mini_program = tk.Label(self.mini_frame, textvariable=self.mini_program_var, font=("Arial", 9, "bold"), fg="#d8ffd8", bg="#333333")
        self.lbl_mini_program.pack()
        self.lbl_mini_weight = tk.Label(self.mini_frame, textvariable=self.last_measurement_var, font=("Courier New", 12, "bold"), fg="#33ff66", bg="#333333")
        self.lbl_mini_weight.pack(pady=(1, 4))
        self.btn_mini_exit = tk.Button(self.mini_frame, text="Vollbild", command=lambda: self.mini_mode_var.set(False), height=1)
        self.btn_mini_exit.pack(pady=(0, 6))
        
        # Notebook für Tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)
        
        self.tab_main = ttk.Frame(self.notebook)
        self.tab_settings = ttk.Frame(self.notebook)
        self.tab_test = ttk.Frame(self.notebook)
        self.tab_help = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab_main, text="Programm")
        self.notebook.add(self.tab_settings, text="Einstellungen")
        self.notebook.add(self.tab_test, text="Testmodus")
        self.notebook.add(self.tab_help, text="Hilfe")
        
        self.build_main_tab()
        self.build_settings_tab()
        self.build_test_tab()
        self.build_help_tab()
        
        # Tracker für Änderungen
        self.counter_var.trace_add("write", self.toggle_counter_visibility)
        self.mini_mode_var.trace_add("write", self.toggle_mini_mode)
        self.test_mode_var.trace_add("write", self.update_test_mode_ui)

    def get_available_ports(self):
        return [port.device for port in serial.tools.list_ports.comports()]

    def resolve_start_port(self, preferred_port=None):
        available_ports = self.get_available_ports()
        if preferred_port and preferred_port in available_ports:
            return preferred_port
        if available_ports:
            return available_ports[0]
        return "COM1"

    def load_settings(self):
        settings = {}
        try:
            if os.path.exists(self.settings_path):
                with open(self.settings_path, "r", encoding="utf-8") as settings_file:
                    settings = json.load(settings_file)
        except Exception:
            settings = {}

        self.baud_var.set(settings.get("baudrate", self.baud_var.get()))
        self.databits_var.set(settings.get("databits", self.databits_var.get()))
        self.parity_var.set(settings.get("parity", self.parity_var.get()))
        self.stopbits_var.set(settings.get("stopbits", self.stopbits_var.get()))

        self.name_prog1.set(settings.get("name_prog1", self.name_prog1.get()))
        self.name_prog2.set(settings.get("name_prog2", self.name_prog2.get()))
        self.name_prog3.set(settings.get("name_prog3", self.name_prog3.get()))

        self.counter_var.set(settings.get("counter_visible", self.counter_var.get()))
        self.auto_reconnect_var.set(settings.get("auto_reconnect", self.auto_reconnect_var.get()))
        self.plausi_var.set(settings.get("plausi1", self.plausi_var.get()))
        self.backup_var.set(settings.get("backup", self.backup_var.get()))
        self.auto_save_var.set(settings.get("auto_save", self.auto_save_var.get()))
        self.auto_save_x_var.set(settings.get("auto_save_x", self.auto_save_x_var.get()))
        self.plausi2_var.set(settings.get("plausi2", self.plausi2_var.get()))
        self.plausi2_limit_var.set(settings.get("plausi2_limit", self.plausi2_limit_var.get()))
        self.mini_mode_var.set(settings.get("mini_mode", self.mini_mode_var.get()))
        self.log_clean_var.set(settings.get("log_clean", self.log_clean_var.get()))
        self.dot_comma_var.set(settings.get("dot_comma", self.dot_comma_var.get()))
        self.unit_var.set(settings.get("unit", self.unit_var.get()))
        self.fixed_decimals_var.set(settings.get("fixed_decimals", self.fixed_decimals_var.get()))
        self.decimal_places_var.set(settings.get("decimal_places", self.decimal_places_var.get()))

        saved_program = settings.get("current_program", self.current_program)
        if saved_program in (1, 2, 3):
            self.current_program = saved_program

        saved_port = settings.get("port", self.port_var.get())
        saved_last_port = settings.get("last_successful_port", saved_port)
        available_ports = self.get_available_ports()
        if saved_last_port in available_ports:
            self.last_successful_port = saved_last_port
        elif saved_port in available_ports:
            self.last_successful_port = saved_port
        elif available_ports:
            self.last_successful_port = available_ports[0]
        else:
            self.last_successful_port = "COM1"
        self.port_var.set(self.last_successful_port)

    def save_settings(self):
        settings = {
            "port": self.port_var.get(),
            "last_successful_port": self.last_successful_port,
            "baudrate": self.baud_var.get(),
            "databits": self.databits_var.get(),
            "parity": self.parity_var.get(),
            "stopbits": self.stopbits_var.get(),
            "name_prog1": self.name_prog1.get(),
            "name_prog2": self.name_prog2.get(),
            "name_prog3": self.name_prog3.get(),
            "current_program": self.current_program,
            "counter_visible": self.counter_var.get(),
            "auto_reconnect": self.auto_reconnect_var.get(),
            "plausi1": self.plausi_var.get(),
            "backup": self.backup_var.get(),
            "auto_save": self.auto_save_var.get(),
            "auto_save_x": self.auto_save_x_var.get(),
            "plausi2": self.plausi2_var.get(),
            "plausi2_limit": self.plausi2_limit_var.get(),
            "mini_mode": self.mini_mode_var.get(),
            "log_clean": self.log_clean_var.get(),
            "dot_comma": self.dot_comma_var.get(),
            "unit": self.unit_var.get(),
            "fixed_decimals": self.fixed_decimals_var.get(),
            "decimal_places": self.decimal_places_var.get(),
        }

        try:
            with open(self.settings_path, "w", encoding="utf-8") as settings_file:
                json.dump(settings, settings_file, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def on_close(self):
        self.save_settings()
        self.is_running = False
        self.window_tracking_active = False
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        self.root.destroy()

    def start_window_tracking(self):
        if os.name != "nt":
            return
        self.window_tracking_active = True
        self.track_last_external_window()

    def track_last_external_window(self):
        if not self.window_tracking_active or os.name != "nt":
            return

        foreground_hwnd = self.get_foreground_window_handle()
        if foreground_hwnd and not self.is_own_window(foreground_hwnd):
            self.last_external_hwnd = foreground_hwnd

        self.root.after(250, self.track_last_external_window)

    def get_foreground_window_handle(self):
        if os.name != "nt":
            return None
        try:
            return ctypes.windll.user32.GetForegroundWindow()
        except Exception:
            return None

    def is_own_window(self, hwnd):
        if not hwnd:
            return False
        try:
            return int(hwnd) == int(self.root.winfo_id())
        except Exception:
            return False

    def restore_last_external_window(self):
        if os.name != "nt":
            return False
        if not self.last_external_hwnd:
            return False

        try:
            user32 = ctypes.windll.user32
            if not user32.IsWindow(self.last_external_hwnd):
                self.last_external_hwnd = None
                return False
            user32.ShowWindow(self.last_external_hwnd, 5)
            user32.SetForegroundWindow(self.last_external_hwnd)
            return True
        except Exception:
            return False

    def format_measurement_output(self, numeric_part, unit_part=""):
        numeric_output = numeric_part
        if self.fixed_decimals_var.get() and numeric_part:
            try:
                decimal_places = max(0, int(self.decimal_places_var.get()))
                numeric_output = f"{float(numeric_part.replace(',', '.')):.{decimal_places}f}"
            except Exception:
                numeric_output = numeric_part

        if self.dot_comma_var.get():
            numeric_output = numeric_output.replace('.', ',')

        if self.unit_var.get() and unit_part:
            return f"{numeric_output} {unit_part}".strip()
        return numeric_output

    def commit_measurement(self, raw_data, processed_data, numeric_reference, source_label="Empfangen"):
        self.last_measurement_var.set(str(processed_data))

        if self.plausi_var.get() and "-" in raw_data:
            self.root.after(0, self.log_to_monitor, f"WARNUNG PLAUSI 1: Negativer Wert! ({raw_data})", "red")
        else:
            self.root.after(0, self.log_to_monitor, f"{source_label}: {raw_data}")

        if self.plausi2_var.get():
            try:
                num_val = float(numeric_reference.replace(',', '.'))
                limit_val = self.get_plausi2_limit_value()
                if limit_val is None:
                    self.root.after(0, self.log_to_monitor, "WARNUNG PLAUSI 2: Grenzwert-Format ungültig (erlaubt: -999,9999 bis 999,9999)", "red")
                elif num_val < limit_val:
                    self.root.after(0, self.log_to_monitor, f"WARNUNG PLAUSI 2: Wert zu niedrig! ({raw_data})", "red")
            except Exception:
                pass

        self.save_to_backup(processed_data)
        self.root.after(0, self.trigger_visual_flash)

        self.counter += 1
        self.root.after(0, lambda: self.lbl_counter.config(text=f"Messungen: {self.counter}"))

        if self.auto_save_var.get() and self.counter % self.auto_save_x_var.get() == 0:
            pyautogui.hotkey('ctrl', 's')
            self.root.after(0, self.log_to_monitor, f"Auto-Save nach {self.counter} Messungen ausgeführt.", "blue")

        pyautogui.write(str(processed_data))

        if self.current_program == 1:
            pyautogui.press('enter')
        elif self.current_program == 2:
            pyautogui.press('tab')
        elif self.current_program == 3:
            pyautogui.press('enter')

    def build_main_tab(self):
        top_frame = tk.Frame(self.tab_main)
        top_frame.pack(fill="x", padx=15, pady=15)
        
        self.ampel_canvas = tk.Canvas(top_frame, width=30, height=30, highlightthickness=0)
        self.ampel_canvas.pack(side="left")
        self.ampel_light = self.ampel_canvas.create_oval(5, 5, 25, 25, fill="red")
        
        self.status_label = tk.Label(top_frame, text="Getrennt", font=("Arial", 11, "bold"), fg="red")
        self.status_label.pack(side="left", padx=10)
        
        tk.Frame(self.tab_main, height=2, bd=1, relief="sunken").pack(fill="x", padx=15, pady=10)
        
        btn_frame = tk.Frame(self.tab_main)
        btn_frame.pack(pady=20)
        
        btn_font = ("Arial", 12, "bold")
        self.btn_prog1 = RoundedButton(
            btn_frame,
            textvariable=self.name_prog1,
            font=btn_font,
            width=170,
            height=54,
            radius=16,
            bg="lightgrey",
            command=lambda: self.set_program(1)
        )
        self.btn_prog1.grid(row=0, column=0, padx=10)
        
        self.btn_prog2 = RoundedButton(
            btn_frame,
            textvariable=self.name_prog2,
            font=btn_font,
            width=170,
            height=54,
            radius=16,
            bg="lightgrey",
            command=lambda: self.set_program(2)
        )
        self.btn_prog2.grid(row=0, column=1, padx=10)
        
        self.btn_prog3 = RoundedButton(
            btn_frame,
            textvariable=self.name_prog3,
            font=btn_font,
            width=170,
            height=54,
            radius=16,
            bg="lightgrey",
            command=lambda: self.set_program(3)
        )
        self.btn_prog3.grid(row=0, column=2, padx=10)

        mini_toggle_frame = tk.Frame(self.tab_main)
        mini_toggle_frame.pack(pady=(0, 10))

        self.btn_mini_toggle = RoundedButton(
            mini_toggle_frame,
            text="Mini-Modus",
            font=("Arial", 11, "bold"),
            width=170,
            height=44,
            radius=14,
            bg="lightgrey",
            command=lambda: self.mini_mode_var.set(not self.mini_mode_var.get())
        )
        self.btn_mini_toggle.pack()
        ToolTip(self.btn_mini_toggle, "Schaltet direkt in den schwebenden Mini-Modus um und zeigt dort Status, Programm und letztes Gewicht an.")
        
        self.lbl_counter = tk.Label(self.tab_main, text="Messungen: 0", font=("Arial", 11, "bold"), fg="blue")
        
        tk.Label(self.tab_main, text="--- Datenmonitor ---", font=("Arial", 10, "bold")).pack(pady=(20, 5))
        
        monitor_frame = tk.Frame(self.tab_main)
        monitor_frame.pack(padx=15, pady=(0, 15), fill="both", expand=True)
        
        self.monitor_text = tk.Text(monitor_frame, height=12, bg="#f4f4f4", state="disabled", font=("Courier", 10))
        self.monitor_text.pack(side="left", fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(monitor_frame, command=self.monitor_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.monitor_text.config(yscrollcommand=scrollbar.set)

        self.set_program(self.current_program)

    def build_settings_tab(self):
        # 1. Schnittstelle
        f_port = tk.LabelFrame(self.tab_settings, text="Schnittstellen-Parameter", font=("Arial", 10, "bold"))
        f_port.pack(fill="x", padx=15, pady=5)
        
        tk.Label(f_port, text="COM Port:").grid(row=0, column=0, sticky="w", padx=10, pady=2)
        ttk.Combobox(f_port, textvariable=self.port_var, values=[p.device for p in serial.tools.list_ports.comports()], width=15).grid(row=0, column=1, pady=2)
        
        tk.Label(f_port, text="Baudrate:").grid(row=1, column=0, sticky="w", padx=10, pady=2)
        ttk.Combobox(f_port, textvariable=self.baud_var, values=["1200", "2400", "4800", "9600", "19200", "38400", "115200"], width=15).grid(row=1, column=1, pady=2)
        
        tk.Label(f_port, text="Datenbits:").grid(row=0, column=2, sticky="w", padx=20, pady=2)
        ttk.Combobox(f_port, textvariable=self.databits_var, values=["7", "8"], width=8).grid(row=0, column=3, pady=2)
        
        tk.Label(f_port, text="Parität:").grid(row=1, column=2, sticky="w", padx=20, pady=2)
        ttk.Combobox(f_port, textvariable=self.parity_var, values=["None", "Odd", "Even"], width=8).grid(row=1, column=3, pady=2)
        
        tk.Label(f_port, text="Stopbits:").grid(row=2, column=2, sticky="w", padx=20, pady=2)
        cb_stop = ttk.Combobox(f_port, textvariable=self.stopbits_var, values=["1", "2"], width=8, state="readonly")
        cb_stop.grid(row=2, column=3, pady=2)
        ToolTip(cb_stop, "Anzahl der Stopbits der Waage. Standard ist '1'. Nur auf '2' stellen, wenn das Handbuch der Waage dies ausdrücklich vorschreibt.")

        self.connect_btn = tk.Button(f_port, text="Verbinden", command=self.toggle_connection, bg="lightgrey", width=12)
        self.connect_btn.grid(row=0, column=4, rowspan=3, padx=(25, 10), pady=2, sticky="ns")
        ToolTip(self.connect_btn, "Stellt die Verbindung zur Waage her oder trennt sie wieder.")

        # 2. Programme
        f_prog = tk.LabelFrame(self.tab_settings, text="Namen der Programme", font=("Arial", 10, "bold"))
        f_prog.pack(fill="x", padx=15, pady=5)
        
        tk.Label(f_prog, text="Button 1 (Zeilensprung):").grid(row=0, column=0, sticky="w", padx=10, pady=2)
        tk.Entry(f_prog, textvariable=self.name_prog1).grid(row=0, column=1, pady=2)
        
        tk.Label(f_prog, text="Button 2 (Spaltensprung):").grid(row=1, column=0, sticky="w", padx=10, pady=2)
        tk.Entry(f_prog, textvariable=self.name_prog2).grid(row=1, column=1, pady=2)
        
        tk.Label(f_prog, text="Button 3 (Zeilensprung):").grid(row=2, column=0, sticky="w", padx=10, pady=2)
        tk.Entry(f_prog, textvariable=self.name_prog3).grid(row=2, column=1, pady=2)

        # 3. Erweiterte Optionen (Mit Tooltips)
        f_opt = tk.LabelFrame(self.tab_settings, text="Erweiterte Optionen & Automatisierung", font=("Arial", 10, "bold"))
        f_opt.pack(fill="both", expand=True, padx=15, pady=5)
        
        # UI & Ansicht
        cb_count = tk.Checkbutton(f_opt, text="Mess-Zähler auf Hauptseite anzeigen", variable=self.counter_var)
        cb_count.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        
        cb_mini = tk.Checkbutton(f_opt, text="Schwebenden Mini-Modus aktivieren (inkl. Visuellem Flash)", variable=self.mini_mode_var)
        cb_mini.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        ToolTip(cb_mini, "Verkleinert das Fenster extrem und hält es immer im Vordergrund über Excel. Blinkt grün bei Erfolg.")
        
        tk.Frame(f_opt, height=1, bg="grey").grid(row=2, column=0, columnspan=2, sticky="we", pady=5)
        
        # Sicherheit & Plausibilität
        cb_recon = tk.Checkbutton(f_opt, text="Auto-Reconnect bei Verbindungsabbruch", variable=self.auto_reconnect_var)
        cb_recon.grid(row=3, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        ToolTip(cb_recon, "Versucht bei Kabel-Wacklern oder Trennung sofort, die Verbindung wiederherzustellen.")
        
        cb_plausi1 = tk.Checkbutton(f_opt, text="Plausibilitäts-Check 1 (Warnung bei Minus-Werten)", variable=self.plausi_var)
        cb_plausi1.grid(row=4, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        
        # Plausibilität 2 (Grenzwert)
        frm_plausi2 = tk.Frame(f_opt)
        frm_plausi2.grid(row=5, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        cb_plausi2 = tk.Checkbutton(frm_plausi2, text="Plausibilitäts-Check 2 (Warnen, wenn Wert kleiner als: ", variable=self.plausi2_var)
        cb_plausi2.pack(side="left")
        ent_plausi2 = tk.Entry(frm_plausi2, textvariable=self.plausi2_limit_var, width=8)
        ent_plausi2.pack(side="left")
        tk.Label(frm_plausi2, text=")").pack(side="left")
        ToolTip(frm_plausi2, "Löst eine rote Warnung im Monitor aus, wenn eine extrem niedrige Einwaage registriert wird. Grenzwertformat: optionales Minus, max 3 Stellen vor und max 4 nach dem Komma.")

        tk.Frame(f_opt, height=1, bg="grey").grid(row=6, column=0, columnspan=2, sticky="we", pady=5)
        
        # Excel Auto-Save
        frm_save = tk.Frame(f_opt)
        frm_save.grid(row=7, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        cb_save = tk.Checkbutton(frm_save, text="Auto-Save (Strg+S) in Excel ausführen nach ", variable=self.auto_save_var)
        cb_save.pack(side="left")
        sp_save = tk.Spinbox(frm_save, from_=1, to=100, textvariable=self.auto_save_x_var, width=4)
        sp_save.pack(side="left")
        tk.Label(frm_save, text=" Messungen").pack(side="left")
        ToolTip(frm_save, "Drückt automatisch STRG+S im Hintergrund, um dein Excel-Dokument regelmäßig zu sichern.")

        # Backup & Log
        cb_backup = tk.Checkbutton(f_opt, text="Hintergrund-Backup (in /backup Ordner speichern)", variable=self.backup_var)
        cb_backup.grid(row=8, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        ToolTip(cb_backup, "Speichert jeden Wert sicherheitshalber in eine Textdatei, falls Excel abstürzt.")
        
        cb_clean = tk.Checkbutton(f_opt, text="Log-Aufräumer (Backups löschen, die älter als 30 Tage sind)", variable=self.log_clean_var)
        cb_clean.grid(row=9, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        ToolTip(cb_clean, "Hält deine Festplatte sauber, indem alte Log-Dateien automatisch vernichtet werden.")

        tk.Frame(f_opt, height=1, bg="grey").grid(row=10, column=0, columnspan=2, sticky="we", pady=5)

        cb_dotcomma = tk.Checkbutton(f_opt, text="Dezimalpunkt automatisch als Komma darstellen ('.' → ',')", variable=self.dot_comma_var)
        cb_dotcomma.grid(row=11, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        ToolTip(cb_dotcomma, "Wandelt den Dezimalpunkt aus den Waagendaten automatisch in ein Komma um (z.B. '1.234' → '1,234'). Empfohlen für Excel mit deutscher Spracheinstellung. Standardmäßig aktiv.")

        cb_unit = tk.Checkbutton(f_opt, text="Einheit mit erfassen (Einheit der Waage in Zelle schreiben)", variable=self.unit_var)
        cb_unit.grid(row=12, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        ToolTip(cb_unit, "Schreibt die von der Waage gesendete Einheit (z.B. 'g', 'mg', 'kg') mit in die Excel-Zelle. Deaktiviert: nur der reine Zahlenwert wird eingetragen, empfohlen für Berechnungen.")

        frm_decimals = tk.Frame(f_opt)
        frm_decimals.grid(row=13, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        cb_decimals = tk.Checkbutton(frm_decimals, text="Feste Nachkommastellen verwenden:", variable=self.fixed_decimals_var)
        cb_decimals.pack(side="left")
        sp_decimals = tk.Spinbox(frm_decimals, from_=0, to=10, textvariable=self.decimal_places_var, width=4)
        sp_decimals.pack(side="left", padx=(5, 0))
        tk.Label(frm_decimals, text=" Stellen").pack(side="left")
        ToolTip(frm_decimals, "Wenn aktiviert, wird der Zahlenwert immer auf die eingestellte Anzahl an Nachkommastellen formatiert. Wenn deaktiviert, werden alle von der Waage gelieferten Nachkommastellen unverändert übernommen.")

    def build_test_tab(self):
        header = tk.Label(self.tab_test, text="Testmodus für simulierte Waagenwerte", font=("Arial", 12, "bold"))
        header.pack(pady=(20, 10))

        info = tk.Label(
            self.tab_test,
            text="Hier kann die App ohne echte Waage getestet werden. PRINT erzeugt einen Zufallswert zwischen 1.0000 g und 3.0000 g und schreibt ihn wie eine echte Messung nach Excel, TARA löscht nur die Anzeige.",
            wraplength=520,
            justify="left"
        )
        info.pack(padx=20, pady=(0, 15), anchor="w")

        cb_test_mode = tk.Checkbutton(
            self.tab_test,
            text="Testmodus aktivieren",
            variable=self.test_mode_var,
            font=("Arial", 11, "bold")
        )
        cb_test_mode.pack(anchor="w", padx=20, pady=(0, 10))
        ToolTip(cb_test_mode, "Aktiviert eine simulierte Waage. PRINT erzeugt dann Testwerte ohne seriellen Befehl an die Waage und schreibt sie nach Excel, TARA löscht nur die Testanzeige.")

        self.test_status_label = tk.Label(self.tab_test, text="Testmodus deaktiviert", fg="red", font=("Arial", 11, "bold"))
        self.test_status_label.pack(anchor="w", padx=20, pady=(0, 15))

        display_frame = tk.Frame(self.tab_test, bg="#111111", bd=3, relief="sunken")
        display_frame.pack(fill="x", padx=20, pady=(0, 20))

        self.test_display_label = tk.Label(
            display_frame,
            textvariable=self.test_display_var,
            bg="#111111",
            fg="#33ff66",
            font=("Courier New", 28, "bold"),
            pady=20
        )
        self.test_display_label.pack(fill="x")

        btn_frame = tk.Frame(self.tab_test)
        btn_frame.pack(pady=10)

        self.test_print_btn = RoundedButton(
            btn_frame,
            text="Print simulieren",
            width=190,
            height=46,
            radius=14,
            bg="lightgrey",
            command=self.simulate_test_print,
            state="disabled"
        )
        self.test_print_btn.grid(row=0, column=0, padx=10)
        ToolTip(self.test_print_btn, "Erzeugt einen simulierten Messwert und schreibt ihn mit den aktuellen Einstellungen in das zuletzt aktive Fremdfenster.")

        self.test_tare_btn = RoundedButton(
            btn_frame,
            text="Tara simulieren",
            width=190,
            height=46,
            radius=14,
            bg="lightgrey",
            command=self.simulate_test_tare,
            state="disabled"
        )
        self.test_tare_btn.grid(row=0, column=1, padx=10)
        ToolTip(self.test_tare_btn, "Löscht das aktuell angezeigte Testgewicht. Der nächste PRINT erzeugt wieder einen neuen Zufallswert.")

    def update_test_mode_ui(self, *args):
        if self.test_mode_var.get():
            self.test_status_label.config(text="Testmodus aktiv", fg="green")
            self.test_print_btn.config(state="normal")
            self.test_tare_btn.config(state="normal")
            self.test_display_var.set("0")
            self.log_to_monitor("--- Testmodus aktiviert ---", "blue")
        else:
            self.test_status_label.config(text="Testmodus deaktiviert", fg="red")
            self.test_print_btn.config(state="disabled")
            self.test_tare_btn.config(state="disabled")
            self.test_display_var.set("----")
            self.log_to_monitor("--- Testmodus deaktiviert ---", "blue")

    def simulate_test_print(self):
        if not self.test_mode_var.get():
            self.log_to_monitor("Testmodus ist deaktiviert.", "orange")
            return

        numeric_part = f"{random.uniform(1.0, 3.0):.4f}"
        raw_data = f"{numeric_part} g"
        processed_data = self.format_measurement_output(numeric_part, "g")
        self.test_display_var.set(processed_data)

        if os.name == "nt":
            if not self.restore_last_external_window():
                self.log_to_monitor("Testmodus PRINT abgebrochen: Kein zuletzt aktives Fremdfenster gefunden.", "orange")
                return
            self.root.after(180, lambda: self.commit_measurement(raw_data, processed_data, numeric_part, "Testmodus PRINT"))
            return

        self.commit_measurement(raw_data, processed_data, numeric_part, "Testmodus PRINT")

    def simulate_test_tare(self):
        if not self.test_mode_var.get():
            self.log_to_monitor("Testmodus ist deaktiviert.", "orange")
            return

        self.test_display_var.set("----")
        self.last_measurement_var.set("----")
        self.log_to_monitor("Testmodus TARA: Gewicht gelöscht.", "orange")

    def build_help_tab(self):
        txt = tk.Text(self.tab_help, wrap="word", bg="#fcfcfc", font=("Arial", 10), padx=15, pady=15)
        txt.pack(fill="both", expand=True)
        
        hilfe_text = """Willkommen bei MK - MessKomplize Version 1.0!

Dieses Tool verbindet Ihre Laborwaage nahtlos mit Excel.

? PROGRAMME:
1. Aufschluss: Nach der Messung wird automatisch ENTER gedrückt (Sprung nach unten).
2. WGH 1: Nach der Messung wird TAB gedrückt (Sprung nach rechts).
3. WGH 2: Nach der Messung wird wieder ENTER gedrückt.
Klicken Sie auf den Button, um das aktive Programm zu wechseln. Das aktive Programm leuchtet grün.

? TESTMODUS:
Im Tab 'Testmodus' können Sie eine simulierte Waage aktivieren. PRINT erzeugt dort ein Zufallsgewicht zwischen 1.0000 g und 3.0000 g und schreibt es direkt nach Excel.
Im Testmodus wird kein Befehl an eine echte Waage gesendet. TARA löscht dort nur die Anzeige im schwarzen Feld.
Vor dem Schreiben versucht das Programm, das zuletzt aktive Fremdfenster wieder in den Vordergrund zu holen.
Die Ausgabe berücksichtigt immer Ihre aktuellen Einstellungen für Nachkommastellen, Einheit und Dezimaltrennzeichen.

? MINI-MODUS:
Aktivieren Sie diesen Modus in den Einstellungen oder direkt im Tab 'Programm', wenn das Programm im Weg ist.
Es verkleinert sich auf ein kleines Fenster, das immer im Vordergrund schwebt und den Verbindungsstatus, das aktive Programm und das zuletzt erfasste Gewicht anzeigt.
Bei jeder erfolgreichen Einwaage blitzt der Mini-Modus kurz grün auf.

? PLAUSIBILITÄTS-CHECK:
Das Programm warnt Sie mit roter Schrift im Datenmonitor, wenn ein Minus-Wert gesendet wird (z.B. nicht tariert) oder die Einwaage unter einem von Ihnen definierten Grenzwert liegt.
Der Grenzwert ist frei eingabbar (optional mit Minus), mit bis zu 3 Stellen vor dem Komma und bis zu 4 Stellen nach dem Komma.

? AUTO-SAVE:
Verlieren Sie nie wieder Daten. Das Programm drückt für Sie nach X Messungen automatisch 'STRG + S' in Excel.

? BACKUP:
Das Hintergrund-Backup ist standardmäßig aktiv. Jeder gespeicherte Wert wird im Format
Zeitstempel;Programmname;Wert in den backup-Ordner geschrieben.

? EINSTELLUNGEN SPEICHERN:
Alle Einstellungen werden beim Schließen des Programms automatisch gespeichert und beim nächsten Start wieder geladen.

? AUTOMATISCHER COM-PORT-START:
Nach einer erfolgreichen Verbindung merkt sich das Programm den zuletzt funktionierenden COM-Port und versucht diesen beim nächsten Start automatisch wieder zu verwenden.

? STOPBITS (Schnittstellen-Parameter):
Legt fest, wie viele Stopbits Ihre Waage verwendet. Standard ist '1'.
Bitte nur auf '2' stellen, wenn das Handbuch der Waage dies ausdrücklich vorschreibt, da es sonst zu Übertragungsfehlern kommen kann.

? DEZIMALPUNKT ALS KOMMA ('.' → ','):
Wandelt den Dezimalpunkt automatisch in ein Komma um (z.B. '1.234' wird zu '1,234').
Diese Option ist standardmäßig aktiv und sorgt dafür, dass Excel in deutscher Spracheinstellung den Wert als Zahl und nicht als Text interpretiert.
Deaktivieren Sie diese Option nur, wenn Ihr Excel explizit auf englische Dezimaltrennung eingestellt ist.

? EINHEIT MIT ERFASSEN:
Ist diese Option aktiv, wird die von der Waage gesendete Einheit (z.B. 'g', 'mg', 'kg') zusammen mit dem Messwert in die Excel-Zelle geschrieben (z.B. '1,2340 g').
Standardmäßig deaktiviert – es wird dann nur der reine Zahlenwert eingetragen, was für weiterführende Berechnungen in Excel empfohlen wird.

? FESTE NACHKOMMASTELLEN:
Ist diese Option deaktiviert, übernimmt das Programm alle von der Waage gelieferten Nachkommastellen unverändert.
Ist sie aktiviert, können Sie mit den Pfeilen der Eingabebox festlegen, wie viele Nachkommastellen in Excel eingetragen werden sollen.
So können Sie z.B. aus '1.234567' gezielt '1,235' oder '1,2346' machen.

Bei anhaltenden Problemen prüfen Sie bitte das COM-Kabel und die Baudrate-Einstellungen der Waage!"""
        
        txt.insert("end", hilfe_text)
        txt.config(state="disabled")

    def toggle_mini_mode(self, *args):
        if self.mini_mode_var.get():
            self.notebook.pack_forget()
            self.mini_frame.pack(fill="both", expand=True)
            self.root.attributes('-topmost', True)
            self.root.geometry("260x120")
        else:
            self.mini_frame.pack_forget()
            self.notebook.pack(fill="both", expand=True)
            self.root.attributes('-topmost', False)
            self.root.geometry("620x850")

    def trigger_visual_flash(self):
        if self.mini_mode_var.get():
            self.mini_frame.config(bg="lime green")
            self.lbl_mini_status.config(bg="lime green")
            self.lbl_mini_program.config(bg="lime green")
            self.lbl_mini_weight.config(bg="lime green")
            self.root.after(200, self.reset_visual_flash)

    def reset_visual_flash(self):
        self.mini_frame.config(bg="#333333")
        self.lbl_mini_status.config(bg="#333333")
        self.lbl_mini_program.config(bg="#333333")
        self.lbl_mini_weight.config(bg="#333333")

    def clean_old_logs(self):
        if self.log_clean_var.get():
            try:
                files = glob.glob("backup/backup_log_*.txt")
                now = time.time()
                for f in files:
                    # Lösche Dateien, die älter als 30 Tage sind (30 * 24 * 60 * 60 Sekunden)
                    if os.stat(f).st_mtime < now - 30 * 86400:
                        os.remove(f)
            except:
                pass

    def set_program(self, prog_id):
        self.current_program = prog_id
        
        self.btn_prog1.config(bg="lightgrey")
        self.btn_prog2.config(bg="lightgrey")
        self.btn_prog3.config(bg="lightgrey")
        
        if prog_id == 1:
            self.btn_prog1.config(bg="lightgreen")
            self.mini_program_var.set(self.name_prog1.get())
            self.log_to_monitor(f"--- Programm gewechselt: {self.name_prog1.get()} (Zeilensprung) ---", "blue")
        elif prog_id == 2:
            self.btn_prog2.config(bg="lightgreen")
            self.mini_program_var.set(self.name_prog2.get())
            self.log_to_monitor(f"--- Programm gewechselt: {self.name_prog2.get()} (Spaltensprung) ---", "blue")
        elif prog_id == 3:
            self.btn_prog3.config(bg="lightgreen")
            self.mini_program_var.set(self.name_prog3.get())
            self.log_to_monitor(f"--- Programm gewechselt: {self.name_prog3.get()} (Zeilensprung) ---", "blue")

    def toggle_counter_visibility(self, *args):
        if self.counter_var.get():
            self.lbl_counter.pack(pady=5)
        else:
            self.lbl_counter.pack_forget()

    def is_valid_plausi2_limit_format(self, value):
        cleaned = (value or "").strip()
        if not cleaned:
            return False
        return re.fullmatch(r"-?\d{1,3}(?:[.,]\d{1,4})?", cleaned) is not None

    def get_plausi2_limit_value(self):
        value = self.plausi2_limit_var.get().strip()
        if not self.is_valid_plausi2_limit_format(value):
            return None
        return float(value.replace(',', '.'))

    def log_to_monitor(self, text, color="black"):
        self.monitor_text.config(state="normal")
        tag_name = str(time.time())
        self.monitor_text.tag_config(tag_name, foreground=color)
        time_str = datetime.datetime.now().strftime("%H:%M:%S")
        self.monitor_text.insert("end", f"[{time_str}] {text}\n", tag_name)
        self.monitor_text.see("end")
        self.monitor_text.config(state="disabled")

    def auto_start_connection(self):
        self.port_var.set(self.resolve_start_port(self.last_successful_port or self.port_var.get()))
        self.log_to_monitor(f"Automatischer Verbindungsversuch auf {self.port_var.get()}...", "blue")
        self.start_reading()

    def toggle_connection(self):
        if not self.is_running:
            self.start_reading()
        else:
            self.stop_reading()

    def update_status(self, connected, text):
        if connected:
            self.ampel_canvas.itemconfig(self.ampel_light, fill="green")
            self.status_label.config(text=text, fg="green")
            self.lbl_mini_status.config(text=f"MK Verbunden ({self.port_var.get()})", fg="lime green")
            self.connect_btn.config(text="Trennen")
        else:
            self.ampel_canvas.itemconfig(self.ampel_light, fill="red")
            self.status_label.config(text=text, fg="red")
            self.lbl_mini_status.config(text="MK Getrennt", fg="red")
            self.connect_btn.config(text="Verbinden")

    def start_reading(self):
        port = self.port_var.get()
        parity_map = {"None": serial.PARITY_NONE, "Odd": serial.PARITY_ODD, "Even": serial.PARITY_EVEN}
        bytesize_map = {"7": serial.SEVENBITS, "8": serial.EIGHTBITS}
        stopbits_map = {"1": serial.STOPBITS_ONE, "2": serial.STOPBITS_TWO}
        
        try:
            self.serial_port = serial.Serial(
                port=port,
                baudrate=int(self.baud_var.get()),
                bytesize=bytesize_map[self.databits_var.get()],
                parity=parity_map[self.parity_var.get()],
                stopbits=stopbits_map[self.stopbits_var.get()],
                timeout=1
            )
            self.is_running = True
            self.last_successful_port = port
            self.save_settings()
            self.update_status(True, f"Verbunden ({port})")
            
            self.read_thread = threading.Thread(target=self.read_from_port, daemon=True)
            self.read_thread.start()
        except Exception as e:
            self.update_status(False, "Fehler / Getrennt")
            self.log_to_monitor(f"Verbindungsfehler: {e}", "red")

    def stop_reading(self):
        self.is_running = False
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        self.update_status(False, "Getrennt")
        self.log_to_monitor("--- Verbindung manuell getrennt ---", "blue")

    def save_to_backup(self, raw_data):
        if self.backup_var.get():
            try:
                os.makedirs("backup", exist_ok=True)
                filename = datetime.datetime.now().strftime("backup/backup_log_%Y-%m-%d.txt")
                with open(filename, "a") as f:
                    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    program_names = {
                        1: self.name_prog1.get(),
                        2: self.name_prog2.get(),
                        3: self.name_prog3.get(),
                    }
                    active_program_name = program_names.get(self.current_program, str(self.current_program))
                    f.write(f"{timestamp};{active_program_name};{raw_data}\n")
            except Exception as e:
                self.root.after(0, self.log_to_monitor, f"Backup-Fehler: {e}", "red")

    def read_from_port(self):
        while self.is_running:
            try:
                if self.serial_port and self.serial_port.is_open:
                    if self.serial_port.in_waiting > 0:
                        raw_bytes = self.serial_port.readline()
                        
                        if raw_bytes:
                            raw_data = raw_bytes.decode('ascii', errors='ignore').strip()
                            number_match = re.search(r'-?\s*\d+(?:[.,]\d+)?', raw_data)
                            numeric_part = number_match.group(0).replace(" ", "") if number_match else ""
                            unit_part = ""
                            if number_match:
                                unit_part = raw_data[number_match.end():].strip()
                                unit_part = re.sub(r'[^A-Za-z%/ ]', '', unit_part).strip()
                            numeric_reference = numeric_part or re.sub(r'[^0-9.,-]', '', raw_data)
                            processed_data = self.format_measurement_output(numeric_reference, unit_part)
                            self.commit_measurement(raw_data, processed_data, numeric_reference)
                                
            except serial.SerialException:
                if self.is_running and self.auto_reconnect_var.get():
                    self.root.after(0, self.update_status, False, "Verbindung verloren! Reconnect...")
                    self.serial_port.close()
                    time.sleep(3)
                    self.root.after(0, self.start_reading)
                else:
                    self.is_running = False
                    self.root.after(0, self.update_status, False, "Abbruch / Kabel fehlt")

if __name__ == "__main__":
    root = tk.Tk()
    app = MessKomplizeApp(root)
    root.mainloop()