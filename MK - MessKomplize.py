import serial
import serial.tools.list_ports
import pyautogui
import tkinter as tk
from tkinter import ttk
import threading
import re
import os
import datetime
import time
import keyboard
import glob

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


class MessKomplizeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MK - MessKomplize Version 1.0")
        self.root.geometry("620x850") # Etwas größer für alle Optionen
        
        self.serial_port = None
        self.is_running = False
        self.counter = 0
        
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
        
        # Neue & Alte Optionen (Standard: Aus)
        self.f9_print_var = tk.BooleanVar(value=False)
        self.f12_tare_var = tk.BooleanVar(value=False)
        self.counter_var = tk.BooleanVar(value=False)
        self.auto_reconnect_var = tk.BooleanVar(value=False)
        self.plausi_var = tk.BooleanVar(value=False)
        self.backup_var = tk.BooleanVar(value=False)
        
        # V1.0 Exklusiv-Optionen
        self.auto_save_var = tk.BooleanVar(value=False)
        self.auto_save_x_var = tk.IntVar(value=10) # Alle 10 Messungen
        
        self.plausi2_var = tk.BooleanVar(value=False)
        self.plausi2_limit_var = tk.DoubleVar(value=100.0) # Grenze in mg/g
        
        self.mini_mode_var = tk.BooleanVar(value=False)
        self.log_clean_var = tk.BooleanVar(value=False)
        
        # UI Aufbauen
        self.setup_ui()
        self.root.after(1000, self.auto_start_connection)
        
        # Wenn Log-Aufräumer aktiv, direkt beim Start einmal aufräumen
        self.root.after(2000, self.clean_old_logs)

    def setup_ui(self):
        # Mini-Mode Frame (versteckt beim Start)
        self.mini_frame = tk.Frame(self.root, bg="#333333")
        self.lbl_mini_status = tk.Label(self.mini_frame, text="MK - MessKomplize: Getrennt", font=("Arial", 11, "bold"), fg="white", bg="#333333")
        self.lbl_mini_status.pack(pady=10)
        self.btn_mini_exit = tk.Button(self.mini_frame, text="Vollbild", command=lambda: self.mini_mode_var.set(False))
        self.btn_mini_exit.pack()
        
        # Notebook für Tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)
        
        self.tab_main = ttk.Frame(self.notebook)
        self.tab_settings = ttk.Frame(self.notebook)
        self.tab_help = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab_main, text="Programm")
        self.notebook.add(self.tab_settings, text="Einstellungen")
        self.notebook.add(self.tab_help, text="Hilfe")
        
        self.build_main_tab()
        self.build_settings_tab()
        self.build_help_tab()
        
        # Tracker für Änderungen
        self.f9_print_var.trace_add("write", self.update_hotkeys)
        self.f12_tare_var.trace_add("write", self.update_hotkeys)
        self.counter_var.trace_add("write", self.toggle_counter_visibility)
        self.mini_mode_var.trace_add("write", self.toggle_mini_mode)

    def build_main_tab(self):
        top_frame = tk.Frame(self.tab_main)
        top_frame.pack(fill="x", padx=15, pady=15)
        
        self.ampel_canvas = tk.Canvas(top_frame, width=30, height=30, highlightthickness=0)
        self.ampel_canvas.pack(side="left")
        self.ampel_light = self.ampel_canvas.create_oval(5, 5, 25, 25, fill="red")
        
        self.status_label = tk.Label(top_frame, text="Getrennt", font=("Arial", 11, "bold"), fg="red")
        self.status_label.pack(side="left", padx=10)
        
        self.connect_btn = tk.Button(top_frame, text="Verbinden", command=self.toggle_connection, bg="lightgrey", width=12)
        self.connect_btn.pack(side="right")
        
        tk.Frame(self.tab_main, height=2, bd=1, relief="sunken").pack(fill="x", padx=15, pady=10)
        
        btn_frame = tk.Frame(self.tab_main)
        btn_frame.pack(pady=20)
        
        btn_font = ("Arial", 12, "bold")
        self.btn_prog1 = tk.Button(btn_frame, textvariable=self.name_prog1, font=btn_font, width=15, height=2, command=lambda: self.set_program(1))
        self.btn_prog1.grid(row=0, column=0, padx=10)
        
        self.btn_prog2 = tk.Button(btn_frame, textvariable=self.name_prog2, font=btn_font, width=15, height=2, command=lambda: self.set_program(2))
        self.btn_prog2.grid(row=0, column=1, padx=10)
        
        self.btn_prog3 = tk.Button(btn_frame, textvariable=self.name_prog3, font=btn_font, width=15, height=2, command=lambda: self.set_program(3))
        self.btn_prog3.grid(row=0, column=2, padx=10)
        
        self.lbl_counter = tk.Label(self.tab_main, text="Messungen: 0", font=("Arial", 11, "bold"), fg="blue")
        
        tk.Label(self.tab_main, text="--- Datenmonitor ---", font=("Arial", 10, "bold")).pack(pady=(20, 5))
        
        monitor_frame = tk.Frame(self.tab_main)
        monitor_frame.pack(padx=15, pady=(0, 15), fill="both", expand=True)
        
        self.monitor_text = tk.Text(monitor_frame, height=12, bg="#f4f4f4", state="disabled", font=("Courier", 10))
        self.monitor_text.pack(side="left", fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(monitor_frame, command=self.monitor_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.monitor_text.config(yscrollcommand=scrollbar.set)
        
        self.set_program(1)

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
        
        # Hotkeys
        cb_f9 = tk.Checkbutton(f_opt, text="Print durch Drücken der Taste 'F9' auslösen", variable=self.f9_print_var)
        cb_f9.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        ToolTip(cb_f9, "Ermöglicht das Senden des Print-Befehls an die Waage, ohne die Maus zu benutzen.")
        
        cb_f12 = tk.Checkbutton(f_opt, text="Tarieren durch Drücken der Taste 'F12' auslösen", variable=self.f12_tare_var)
        cb_f12.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        ToolTip(cb_f12, "Stellt die Waage sofort auf 0.0000, wenn F12 gedrückt wird.")
        
        # UI & Ansicht
        cb_count = tk.Checkbutton(f_opt, text="Mess-Zähler auf Hauptseite anzeigen", variable=self.counter_var)
        cb_count.grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        
        cb_mini = tk.Checkbutton(f_opt, text="Schwebenden Mini-Modus aktivieren (inkl. Visuellem Flash)", variable=self.mini_mode_var)
        cb_mini.grid(row=3, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        ToolTip(cb_mini, "Verkleinert das Fenster extrem und hält es immer im Vordergrund über Excel. Blinkt grün bei Erfolg.")
        
        tk.Frame(f_opt, height=1, bg="grey").grid(row=4, column=0, columnspan=2, sticky="we", pady=5)
        
        # Sicherheit & Plausibilität
        cb_recon = tk.Checkbutton(f_opt, text="Auto-Reconnect bei Verbindungsabbruch", variable=self.auto_reconnect_var)
        cb_recon.grid(row=5, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        ToolTip(cb_recon, "Versucht bei Kabel-Wacklern oder Trennung sofort, die Verbindung wiederherzustellen.")
        
        cb_plausi1 = tk.Checkbutton(f_opt, text="Plausibilitäts-Check 1 (Warnung bei Minus-Werten)", variable=self.plausi_var)
        cb_plausi1.grid(row=6, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        
        # Plausibilität 2 (Grenzwert)
        frm_plausi2 = tk.Frame(f_opt)
        frm_plausi2.grid(row=7, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        cb_plausi2 = tk.Checkbutton(frm_plausi2, text="Plausibilitäts-Check 2 (Warnen, wenn Wert kleiner als: ", variable=self.plausi2_var)
        cb_plausi2.pack(side="left")
        sp_plausi2 = tk.Spinbox(frm_plausi2, from_=0, to=10000, increment=10, textvariable=self.plausi2_limit_var, width=6)
        sp_plausi2.pack(side="left")
        tk.Label(frm_plausi2, text=")").pack(side="left")
        ToolTip(frm_plausi2, "Löst eine rote Warnung im Monitor aus, wenn eine extrem niedrige Einwaage (z.B. leeres Gefäß) registriert wird.")

        tk.Frame(f_opt, height=1, bg="grey").grid(row=8, column=0, columnspan=2, sticky="we", pady=5)
        
        # Excel Auto-Save
        frm_save = tk.Frame(f_opt)
        frm_save.grid(row=9, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        cb_save = tk.Checkbutton(frm_save, text="Auto-Save (Strg+S) in Excel ausführen nach ", variable=self.auto_save_var)
        cb_save.pack(side="left")
        sp_save = tk.Spinbox(frm_save, from_=1, to=100, textvariable=self.auto_save_x_var, width=4)
        sp_save.pack(side="left")
        tk.Label(frm_save, text=" Messungen").pack(side="left")
        ToolTip(frm_save, "Drückt automatisch STRG+S im Hintergrund, um dein Excel-Dokument regelmäßig zu sichern.")

        # Backup & Log
        cb_backup = tk.Checkbutton(f_opt, text="Hintergrund-Backup (in /backup Ordner speichern)", variable=self.backup_var)
        cb_backup.grid(row=10, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        ToolTip(cb_backup, "Speichert jeden Wert sicherheitshalber in eine Textdatei, falls Excel abstürzt.")
        
        cb_clean = tk.Checkbutton(f_opt, text="Log-Aufräumer (Backups löschen, die älter als 30 Tage sind)", variable=self.log_clean_var)
        cb_clean.grid(row=11, column=0, columnspan=2, sticky="w", padx=10, pady=2)
        ToolTip(cb_clean, "Hält deine Festplatte sauber, indem alte Log-Dateien automatisch vernichtet werden.")

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

? HOTKEYS:
- F9: Sendet den Befehl "Print" an die Waage.
- F12: Sendet den Befehl "Tarieren" an die Waage.
(Diese Tasten funktionieren, während Sie in Excel arbeiten!)

? MINI-MODUS:
Aktivieren Sie diesen Modus in den Einstellungen, wenn das Programm im Weg ist. Es verkleinert sich auf einen winzigen Balken, der immer im Vordergrund schwebt. Bei jeder erfolgreichen Einwaage blitzt er kurz grün auf.

? PLAUSIBILITÄTS-CHECK:
Das Programm warnt Sie mit roter Schrift im Datenmonitor, wenn ein Minus-Wert gesendet wird (z.B. nicht tariert) oder die Einwaage unter einem von Ihnen definierten Grenzwert liegt.

? AUTO-SAVE:
Verlieren Sie nie wieder Daten. Das Programm drückt für Sie nach X Messungen automatisch 'STRG + S' in Excel.

Bei anhaltenden Problemen prüfen Sie bitte das COM-Kabel und die Baudrate-Einstellungen der Waage!"""
        
        txt.insert("end", hilfe_text)
        txt.config(state="disabled")

    def toggle_mini_mode(self, *args):
        if self.mini_mode_var.get():
            self.notebook.pack_forget()
            self.mini_frame.pack(fill="both", expand=True)
            self.root.attributes('-topmost', True)
            self.root.geometry("350x80")
        else:
            self.mini_frame.pack_forget()
            self.notebook.pack(fill="both", expand=True)
            self.root.attributes('-topmost', False)
            self.root.geometry("620x850")

    def trigger_visual_flash(self):
        if self.mini_mode_var.get():
            self.mini_frame.config(bg="lime green")
            self.lbl_mini_status.config(bg="lime green")
            self.root.after(200, self.reset_visual_flash)

    def reset_visual_flash(self):
        self.mini_frame.config(bg="#333333")
        self.lbl_mini_status.config(bg="#333333")

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
            self.log_to_monitor(f"--- Programm gewechselt: {self.name_prog1.get()} (Zeilensprung) ---", "blue")
        elif prog_id == 2:
            self.btn_prog2.config(bg="lightgreen")
            self.log_to_monitor(f"--- Programm gewechselt: {self.name_prog2.get()} (Spaltensprung) ---", "blue")
        elif prog_id == 3:
            self.btn_prog3.config(bg="lightgreen")
            self.log_to_monitor(f"--- Programm gewechselt: {self.name_prog3.get()} (Zeilensprung) ---", "blue")

    def toggle_counter_visibility(self, *args):
        if self.counter_var.get():
            self.lbl_counter.pack(pady=5)
        else:
            self.lbl_counter.pack_forget()

    def update_hotkeys(self, *args):
        keyboard.unhook_all()
        if self.f9_print_var.get():
            keyboard.on_release_key('f9', lambda e: self.send_to_scale("P\r\n"))
        if self.f12_tare_var.get():
            keyboard.on_release_key('f12', lambda e: self.send_to_scale("T\r\n"))

    def send_to_scale(self, command):
        if self.is_running and self.serial_port and self.serial_port.is_open:
            try:
                self.serial_port.write(command.encode('ascii'))
                cmd_name = "PRINT" if "P" in command else "TARA"
                self.root.after(0, self.log_to_monitor, f"Befehl gesendet: {cmd_name}", "orange")
            except Exception as e:
                pass

    def log_to_monitor(self, text, color="black"):
        self.monitor_text.config(state="normal")
        tag_name = str(time.time())
        self.monitor_text.tag_config(tag_name, foreground=color)
        time_str = datetime.datetime.now().strftime("%H:%M:%S")
        self.monitor_text.insert("end", f"[{time_str}] {text}\n", tag_name)
        self.monitor_text.see("end")
        self.monitor_text.config(state="disabled")

    def auto_start_connection(self):
        self.log_to_monitor("Automatischer Verbindungsversuch auf COM1...", "blue")
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
                    time_str = datetime.datetime.now().strftime("%H:%M:%S")
                    f.write(f"{time_str} | Programm: {self.current_program} | Wert: {raw_data}\n")
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
                            processed_data = re.sub(r'[^0-9.,-]', '', raw_data)
                            processed_data = processed_data.replace('.', ',')
                            
                            # Plausi 1: Minuswert
                            if self.plausi_var.get() and "-" in raw_data:
                                self.root.after(0, self.log_to_monitor, f"WARNUNG PLAUSI 1: Negativer Wert! ({raw_data})", "red")
                            else:
                                self.root.after(0, self.log_to_monitor, f"Empfangen: {raw_data}")
                            
                            # Plausi 2: Grenzwert unterschritten
                            if self.plausi2_var.get():
                                try:
                                    # Für die Umwandlung ins Float-Format kurz Komma zu Punkt zurückwandeln
                                    num_val = float(processed_data.replace(',', '.'))
                                    if num_val < self.plausi2_limit_var.get():
                                        self.root.after(0, self.log_to_monitor, f"WARNUNG PLAUSI 2: Wert zu niedrig! ({raw_data})", "red")
                                except: pass
                            
                            self.save_to_backup(raw_data)
                            self.root.after(0, self.trigger_visual_flash)
                            
                            self.counter += 1
                            self.root.after(0, lambda: self.lbl_counter.config(text=f"Messungen: {self.counter}"))
                            
                            # Auto-Save auslösen
                            if self.auto_save_var.get() and self.counter % self.auto_save_x_var.get() == 0:
                                pyautogui.hotkey('ctrl', 's')
                                self.root.after(0, self.log_to_monitor, f"Auto-Save nach {self.counter} Messungen ausgeführt.", "blue")
                            
                            pyautogui.write(processed_data)
                            
                            if self.current_program == 1:
                                pyautogui.press('enter')
                            elif self.current_program == 2:
                                pyautogui.press('tab')
                            elif self.current_program == 3:
                                pyautogui.press('enter')
                                
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