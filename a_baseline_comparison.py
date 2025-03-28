import tkinter as tk
from tkinter import filedialog
from tkinter import ttk, messagebox
import calmpy
import webbrowser
import os
import pandas as pd
from datetime import datetime
import string

class BaselineComparisonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Baseline Comparison Tool")
        self.root.geometry("900x700")
        self.root.configure(bg="#442EDF")  
        
         # Initialisiere den Codebeamer-Server
      
        
        

        # Noch KEINE Codebeamer-Verbindung herstellen
        self.cb_server = None  # Erst nach Initialisierung setzen
        
        # GUI-Elemente erstellen
        self.create_widgets()

    def create_widgets(self):
        """Create GUI elements"""
        # Server Mapping
        self.server_mapping = {
            "Dev": "https://dev-cb.vwgroup.com/cb",
            "Test": "https://qs-cb.vwgroup.com/cb",
            "QS": "https://qs-cb.vwgroup.com/cb",
            "QS-Intranet": "https://qs-cb.wob.vw.vwg/cb",
            "Prod": "https://cb.vwgroup.com/cb",
            "Prod-Intranet": "https://cb.wob.vw.vwg/cb",
            "Play": "https://cb.play-codebeamer.projects.de-wob-3.cloud.vwgroup.com/cb",
            "Dev-IT": "https://devit-cb.vwgroup.com/cb",
            "Local": "https://cb-start:8443/cb",
            "Tmpl": "https://tmpl-cb.vwgroup.com/cb",
            "Tmpl-TC": "https://tmpl-tc-cb.vwgroup.com/cb"
        }

        # Title Label (Zentriert)
        title_label = tk.Label(
            self.root, text="Baseline Comparison Tool", font=("Helvetica", 20, "bold"),
            fg="#1CEE97", bg="#442EDF"
        )
        title_label.pack(pady=20)

        # Frame für Server-Auswahl
        server_frame = tk.Frame(self.root, bg="#442EDF")
        server_frame.pack(pady=10, padx=20)

        server_label = tk.Label(server_frame, text="Select Codebeamer Server:", font=("Helvetica", 14, "bold"), fg="#1CEE97", bg="#442EDF")
        server_label.grid(row=0, column=0, columnspan=2, pady=5, sticky="ew")

        self.server_var = tk.StringVar()
        self.server_dropdown = ttk.Combobox(server_frame, textvariable=self.server_var, state="readonly", width=50)
        self.server_dropdown['values'] = list(self.server_mapping.keys())
        self.server_dropdown.grid(row=1, column=0, pady=5, padx=5, columnspan=2)
        server_options = ["Bitte auswählen..."] + list(self.server_mapping.keys())
        self.server_dropdown['values'] = server_options
        self.server_var.set("Bitte auswählen...")  # <- Standard setzen
        self.server_dropdown.bind("<<ComboboxSelected>>", self.on_server_selected)
        
        # Initialise Button
       # initialise_button = ttk.Button(server_frame, text="Initialise", command=self.initialise_server, style="TButton")
       # initialise_button.grid(row=2, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

        # Hauptframe für Auswahlfelder (Zentriert)
        frame = tk.Frame(self.root, bg="#442EDF")
        frame.pack(pady=10, padx=20)

        # Projekt Dropdown
        project_label = tk.Label(frame, text="Select Project:", font=("Helvetica", 14, "bold"), fg="#1CEE97", bg="#442EDF")
        project_label.grid(row=0, column=0, pady=5, sticky="ew")
        self.project_var = tk.StringVar()
        self.project_dropdown = ttk.Combobox(frame, textvariable=self.project_var, state="readonly", width=50)
        self.project_dropdown.grid(row=1, column=0, pady=5, padx=5)
        self.project_dropdown.bind("<<ComboboxSelected>>", self.on_project_selected)

        # Tracker Dropdown
        tracker_label = tk.Label(frame, text="Select Tracker:", font=("Helvetica", 14, "bold"), fg="#1CEE97", bg="#442EDF")
        tracker_label.grid(row=2, column=0, pady=5, sticky="ew")
        self.tracker_var = tk.StringVar()
        self.tracker_dropdown = ttk.Combobox(frame, textvariable=self.tracker_var, state="disabled", width=50)
        self.tracker_dropdown.grid(row=3, column=0, pady=5, padx=5)
        self.tracker_dropdown.bind("<<ComboboxSelected>>", self.on_tracker_selected)

        # Baseline 1 Dropdown
        baseline1_label = tk.Label(frame, text="Select Baseline 1:", font=("Helvetica", 14, "bold"), fg="#1CEE97", bg="#442EDF")
        baseline1_label.grid(row=4, column=0, pady=5, sticky="ew")
        self.baseline1_var = tk.StringVar()
        self.baseline1_dropdown = ttk.Combobox(frame, textvariable=self.baseline1_var, state="disabled", width=50)
        self.baseline1_dropdown.grid(row=5, column=0, pady=5, padx=5)

        # Baseline 2 Dropdown
        baseline2_label = tk.Label(frame, text="Select Baseline 2:", font=("Helvetica", 14, "bold"), fg="#1CEE97", bg="#442EDF")
        baseline2_label.grid(row=6, column=0, pady=5, sticky="ew")
        self.baseline2_var = tk.StringVar()
        self.baseline2_dropdown = ttk.Combobox(frame, textvariable=self.baseline2_var, state="disabled", width=50)
        self.baseline2_dropdown.grid(row=7, column=0, pady=5, padx=5)

        # Button Container für bessere Kontrolle
        button_container = tk.Frame(self.root, bg="#442EDF")
        button_container.pack(pady=20, fill="x")

        # Download & Upload Fields Buttons (Nebeneinander)
        field_button_frame = tk.Frame(button_container, bg="#442EDF")
        field_button_frame.pack(fill="x", padx=20, pady=5)

        download_fields_button = ttk.Button(field_button_frame, text="Download Fields", command=self.download_fields, style="TButton")
        upload_fields_button = ttk.Button(field_button_frame, text="Upload Fields", command=self.upload_fields, style="TButton")

        download_fields_button.grid(row=0, column=0, padx=10, pady=5, sticky="ew")
        upload_fields_button.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        field_button_frame.columnconfigure(0, weight=1)
        field_button_frame.columnconfigure(1, weight=1)

        # Baseline Comparison Buttons (Nebeneinander)
        compare_button_frame = tk.Frame(button_container, bg="#442EDF")
        compare_button_frame.pack(fill="x", padx=20, pady=5)

        compare_button_excel = ttk.Button(compare_button_frame, text="Baseline Comparison Excel", command=self.compare_baselines_excel_diff, style="TButton")
        #compare_button_html = ttk.Button(compare_button_frame, text="Baseline Comparison HTML", command=self.compare_baselines_html, style="TButton")

       # compare_button_excel.grid(row=0, column=0, padx=20, pady=5, sticky="ew")
        compare_button_excel.pack(pady=5)
        #compare_button_html.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        compare_button_frame.columnconfigure(0, weight=1)
        compare_button_frame.columnconfigure(1, weight=1)

        # Footer Label (Fixiert am unteren Rand)
        footer_label = tk.Label(
            self.root, text="Made by Hassan El Khamis - hassan.el-khamis@capgemini.com",
            font=("Helvetica", 10, "italic"), fg="#FFFFFF", bg="#442EDF"
        )
        footer_label.pack(side="bottom", pady=10, fill="x")
    
    def on_server_selected(self, event):
        self.initialise_server()    
    
    def initialise_server(self):
        """Initialisiert die Verbindung mit dem ausgewählten Codebeamer-Server"""
        selected_server = self.server_var.get()
        print(f"[DEBUG] Auswahl: {selected_server}")  # Debug-Ausgabe
       

        if not selected_server or selected_server == "Bitte auswählen...":
            messagebox.showerror("Fehler", "Bitte wählen Sie einen gültigen Server aus!")
            return

        server_url = self.server_mapping.get(selected_server)
        if not server_url:
            messagebox.showerror("Fehler", f"Unbekannter Server: {selected_server}")
            return

        try:
            self.cb_server = calmpy.Server(url=server_url)
            messagebox.showinfo("Erfolg", f"Erfolgreich mit '{selected_server}' verbunden!")

            # Projekte nach erfolgreicher Verbindung laden
            self.projects, self.project_names = self.get_projects()
            self.project_dropdown['values'] = sorted(self.project_names)
        
        

        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Verbinden mit '{selected_server}':\n{e}")
        
        print(f"[DEBUG] Verbindung zu Server: {server_url}")

        
    def get_projects(self):
        """Lädt alle Projekte"""
        if not self.cb_server:
            messagebox.showerror("Fehler", "Keine Verbindung zum Server. Bitte initialisieren.")
            return [], []
        try:
            projects = self.cb_server.get_projects()
            return projects, [proj.name for proj in projects]
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Abrufen der Projekte: {e}")
            return [], []

    def on_project_selected(self, event):
        """Lädt Tracker nach Projektwahl"""
        selected_project_name = self.project_var.get()
        project = next((p for p in self.projects if p.name == selected_project_name), None)
        if project:
            self.trackers = project.get_trackers()
            self.tracker_dropdown['values'] = [t.name for t in self.trackers]
            self.tracker_dropdown.config(state="readonly")

    def on_tracker_selected(self, event):
        """Lädt Baselines nach Trackerwahl"""
        tracker_name = self.tracker_var.get()
        tracker = next((t for t in self.trackers if t.name == tracker_name), None)
        if tracker:
            self.baseline_info = tracker.get_baselines()
            baseline_names = [b.name for b in self.baseline_info]
            self.baseline1_dropdown['values'] = self.baseline2_dropdown['values'] = baseline_names
            self.baseline1_dropdown.config(state="readonly")
            self.baseline2_dropdown.config(state="readonly")
            
    def download_fields(self):
        """Speichert alle Felder des ausgewählten Trackers in einer Textdatei"""
        tracker_name = self.tracker_var.get()
        tracker = next((t for t in self.trackers if t.name == tracker_name), None)

        if not tracker:
            messagebox.showerror("Fehler", "Bitte wählen Sie einen Tracker aus, bevor Sie die Felder herunterladen!")
            return

        try:
            field_configs = tracker.get_field_config()
            field_names = [field.label for field in field_configs]
        except Exception as e:
            messagebox.showerror("Fehler", f"Feldkonfiguration konnte nicht geladen werden: {e}")
            return

        if not field_names:
            messagebox.showinfo("Info", "Keine Felder gefunden.")
            return

        # Ordner "Downloaded-Fields" erstellen
        folder_name = "Downloaded-Fields"
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)

        # Datei speichern
        file_path = os.path.join(folder_name, f"{tracker_name}-Fields-Download.txt")

        with open(file_path, "w", encoding="utf-8") as file:
            file.write("\n".join(field_names))

        messagebox.showinfo("Erfolg", f"Felder gespeichert in:\n{file_path}")   
    
    
    def upload_fields(self):
        """Lädt eine Textdatei mit ausgewählten Feldern hoch und speichert sie für den Vergleich"""
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])

        if not file_path:
            return  # Benutzer hat keinen Upload ausgewählt

        try:
            with open(file_path, "r", encoding="utf-8") as file:
                self.selected_fields = [line.strip() for line in file.readlines()]
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Datei: {e}")
            return

        if not self.selected_fields:
            messagebox.showerror("Fehler", "Die hochgeladene Datei enthält keine gültigen Felder.")
            return

        messagebox.showinfo("Erfolg", f"{len(self.selected_fields)} Felder wurden für den Vergleich hochgeladen.")

    def compare_baselines_excel(self):
        """Vergleicht zwei Baselines mit allen Feldern oder nur den hochgeladenen Feldern"""
        tracker_name = self.tracker_var.get()
        baseline1_name = self.baseline1_var.get()
        baseline2_name = self.baseline2_var.get()

        tracker = next((t for t in self.trackers if t.name == tracker_name), None)
        baseline1_id = next((b.id for b in self.baseline_info if b.name == baseline1_name), None)
        baseline2_id = next((b.id for b in self.baseline_info if b.name == baseline2_name), None)

        if not tracker or not baseline1_id or not baseline2_id:
            messagebox.showerror("Fehler", "Tracker oder Baselines konnten nicht gefunden werden.")
            return

        # Items aus beiden Baselines abrufen
        items1 = {item.id: item for item in tracker.get_items(baseline=baseline1_id)}
        items2 = {item.id: item for item in tracker.get_items(baseline=baseline2_id)}

        # Prüfe, ob ein Feld-Upload stattgefunden hat
        if hasattr(self, "selected_fields") and self.selected_fields:
            all_fields = self.selected_fields  # Nur hochgeladene Felder verwenden
        else:
            try:
                field_configs = tracker.get_field_config()
                all_fields = [field.label for field in field_configs]  # Standard: Alle Felder
            except Exception as e:
                messagebox.showerror("Fehler", f"Feldkonfiguration konnte nicht geladen werden: {e}")
                return

        differences = []
        for item_id, item1 in items1.items():
            item2 = items2.get(item_id, None)
            item_diffs = {}

            for field in all_fields:
                value1 = item1[field] if field in item1.fields else "Nicht vorhanden"
                value2 = item2[field] if item2 and field in item2.fields else "Nicht vorhanden"

                item_diffs[field] = (value1, value2)  # JEDES Feld wird aufgenommen

            differences.append({
                "tracker_name": tracker.name,
                "item_id": item_id,
                "item_name": item1["Name"] if "Name" in item1.fields else "Unbenannt",
                "differences": item_diffs
            })

        self.create_excel_report(differences, baseline1_name, baseline2_name)
       
      
    def compare_baselines_excel_diff(self):
        """Vergleicht zwei Baselines und speichert nur Unterschiede in einer Excel-Datei."""
        tracker_name = self.tracker_var.get()
        baseline1_name = self.baseline1_var.get()
        baseline2_name = self.baseline2_var.get()

        tracker = next((t for t in self.trackers if t.name == tracker_name), None)
        baseline1_id = next((b.id for b in self.baseline_info if b.name == baseline1_name), None)
        baseline2_id = next((b.id for b in self.baseline_info if b.name == baseline2_name), None)

        if not tracker or not baseline1_id or not baseline2_id:
            messagebox.showerror("Fehler", "Tracker oder Baselines konnten nicht gefunden werden.")
            return

        # Items aus beiden Baselines abrufen
        items1 = {item.id: item for item in tracker.get_items(baseline=baseline1_id)}
        items2 = {item.id: item for item in tracker.get_items(baseline=baseline2_id)}

        # Felder laden (alle oder ausgewählte)
        if hasattr(self, "selected_fields") and self.selected_fields:
            all_fields = self.selected_fields
        else:
            try:
                field_configs = tracker.get_field_config()
                all_fields = [field.label for field in field_configs]
            except Exception as e:
                messagebox.showerror("Fehler", f"Feldkonfiguration konnte nicht geladen werden: {e}")
                return

        differences = []
        for item_id, item1 in items1.items():
            item2 = items2.get(item_id)
            if not item2:
                continue

            item_diffs = {}
            different = False

            for field in all_fields:
                val1 = str(item1[field]) if field in item1.fields else "Nicht vorhanden"
                val2 = str(item2[field]) if field in item2.fields else "Nicht vorhanden"
                item_diffs[field] = (val1, val2)

                if val1 != val2:
                    different = True

            if different:
                differences.append({
                    "tracker_name": tracker.name,
                    "item_id": item_id,
                    "item_name": item1["Name"] if "Name" in item1.fields else "Unbenannt",
                    "differences": item_diffs
                })

        # Prüfen, ob überhaupt Unterschiede vorhanden sind
        if not differences:
            messagebox.showinfo("Info", "Keine Unterschiede zwischen den Baselines gefunden.")
            return

        # Nur Unterschiede in Excel schreiben
        self.create_excel_report(differences, baseline1_name, baseline2_name)
         
    
    def compare_baselines_html(self):
        """Vergleicht zwei Baselines mit allen Feldern oder nur den hochgeladenen Feldern"""
        tracker_name = self.tracker_var.get()
        baseline1_name = self.baseline1_var.get()
        baseline2_name = self.baseline2_var.get()

        tracker = next((t for t in self.trackers if t.name == tracker_name), None)
        baseline1_id = next((b.id for b in self.baseline_info if b.name == baseline1_name), None)
        baseline2_id = next((b.id for b in self.baseline_info if b.name == baseline2_name), None)

        if not tracker or not baseline1_id or not baseline2_id:
            messagebox.showerror("Fehler", "Tracker oder Baselines konnten nicht gefunden werden.")
            return

        # Items aus beiden Baselines abrufen
        items1 = {item.id: item for item in tracker.get_items(baseline=baseline1_id)}
        items2 = {item.id: item for item in tracker.get_items(baseline=baseline2_id)}

        # Prüfe, ob ein Feld-Upload stattgefunden hat
        if hasattr(self, "selected_fields") and self.selected_fields:
            all_fields = self.selected_fields  # Nur hochgeladene Felder verwenden
        else:
            try:
                field_configs = tracker.get_field_config()
                all_fields = [field.label for field in field_configs]  # Standard: Alle Felder
            except Exception as e:
                messagebox.showerror("Fehler", f"Feldkonfiguration konnte nicht geladen werden: {e}")
                return

        differences = []
        for item_id, item1 in items1.items():
            item2 = items2.get(item_id, None)
            item_diffs = {}

            for field in all_fields:
                value1 = item1[field] if field in item1.fields else "Nicht vorhanden"
                value2 = item2[field] if item2 and field in item2.fields else "Nicht vorhanden"

                item_diffs[field] = (value1, value2)  # JEDES Feld wird aufgenommen

            differences.append({
                "tracker_name": tracker.name,
                "item_id": item_id,
                "item_name": item1["Name"] if "Name" in item1.fields else "Unbenannt",
                "differences": item_diffs
            })

        self.create_html_report(differences, baseline1_name, baseline2_name)
      
        
    
    
    def col_num_to_excel(self, col_num):
        """Konvertiert eine Spaltennummer (z.B. 27) in einen Excel-Buchstaben (z.B. 'AA')."""
        col_str = ""
        while col_num >= 0:
            col_str = string.ascii_uppercase[col_num % 26] + col_str
            col_num = (col_num // 26) - 1
        return col_str 
    
    def create_excel_report_diff(self, differences, baseline1_name, baseline2_name):
        """Erstellt eine formatierte Excel-Datei mit nur den Zeilen, in denen Unterschiede existieren."""
        
        # Falls keine Attribute ausgewählt wurden, abbrechen
        if not hasattr(self, "selected_attributes") or not self.select_attributes:
            messagebox.showerror("Fehler", "Keine Attribute ausgewählt!")
            return

        # Nur die ausgewählten Attribute verwenden
        selected_fields = ["ID", "Name"] + [field for field in self.select_attributes if field not in ["ID", "Name"]]

        # DataFrames mit gefilterten Spalten erstellen
        df_baseline1 = pd.DataFrame([{field: diff["differences"].get(field, ["N/A", "N/A"])[0] for field in selected_fields} for diff in differences])
        df_baseline2 = pd.DataFrame([{field: diff["differences"].get(field, ["N/A", "N/A"])[1] for field in selected_fields} for diff in differences])

        # Falls kein Unterschied vorhanden ist, abbrechen
        if df_baseline1.empty or df_baseline2.empty:
            messagebox.showinfo("Info", "Keine Unterschiede gefunden. Keine Excel-Datei erstellt.")
            return

        # Excel speichern
        file_path = f"Baseline-Comparison-{baseline1_name}-vs-{baseline2_name}.xlsx"
        
        with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
            workbook = writer.book
            sheet_name = "Differences Only"

            df_baseline1.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2, startcol=0)
            df_baseline2.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2, startcol=len(selected_fields) + 2)

            worksheet = writer.sheets[sheet_name]

            # Formatierungen definieren
            header_format = workbook.add_format({"bold": True, "bg_color": "#007BFF", "color": "white", "align": "center", "border": 1})
            changed_format = workbook.add_format({"bg_color": "#FFA07A", "border": 1, "text_wrap": True})

            # Header setzen
            for col_num, field in enumerate(selected_fields):
                worksheet.write(1, col_num, field, header_format)
                worksheet.write(1, col_num + len(selected_fields) + 2, field, header_format)

            # Unterschiede markieren
            for row_idx in range(2, len(df_baseline1) + 2):
                for col_idx in range(len(selected_fields)):
                    val1 = df_baseline1.iloc[row_idx - 2, col_idx]
                    val2 = df_baseline2.iloc[row_idx - 2, col_idx]
                    if val1 != val2:
                        worksheet.write(row_idx, col_idx, val1, changed_format)
                        worksheet.write(row_idx, col_idx + len(selected_fields) + 2, val2, changed_format)

        messagebox.showinfo("Success", f"Excel file successfully saved:\n{file_path}")


    def create_excel_report(self, differences, baseline1_name, baseline2_name):
        """Erstellt eine formatierte Excel-Datei mit Baselines als separate Tabellen, automatischem Zeilenumbruch und markierten Unterschieden"""

        # Projekt-, Tracker- & Baseline-Namen abrufen & formatieren für den Dateinamen
        project_name = self.project_var.get().replace(" ", "_")
        tracker_name = self.tracker_var.get().replace(" ", "_")
        baseline1_name = baseline1_name.replace(" ", "_")
        baseline2_name = baseline2_name.replace(" ", "_")

        # Aktuelles Datum für den Dateinamen
        date_str = datetime.now().strftime("%d%m%Y_%H%M%S")

        # Ordner für die Excel-Dateien erstellen, falls nicht vorhanden
        folder_name = "Baseline-Comparison"
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)

        # Dateipfad generieren
        file_name = f"{project_name}-{tracker_name}-Complete-{date_str}.xlsx"
        file_path = os.path.join(folder_name, file_name)

        # Alle Felder sammeln
        all_fields = set()
        for diff in differences:
            all_fields.update(diff["differences"].keys())

        # Wichtige Felder zuerst anzeigen
        fixed_fields = ["ID", "Name", "Description"]
        remaining_fields = sorted(field for field in all_fields if field not in fixed_fields)
        all_fields = fixed_fields + remaining_fields  # Endgültige Feld-Reihenfolge

        # Daten für beide Baselines vorbereiten
        baseline1_data = []
        baseline2_data = []

        for diff in differences:
            row1 = []
            row2 = []
            for field in all_fields:
                value1 = str(diff["differences"].get(field, ["Nicht vorhanden", "Nicht vorhanden"])[0]).strip()
                value2 = str(diff["differences"].get(field, ["Nicht vorhanden", "Nicht vorhanden"])[1]).strip()
                row1.append(value1)
                row2.append(value2)
            baseline1_data.append(row1)
            baseline2_data.append(row2)

        # DataFrames für beide Baselines erstellen
        df_baseline1 = pd.DataFrame(baseline1_data, columns=all_fields)
        df_baseline2 = pd.DataFrame(baseline2_data, columns=all_fields)

        # Excel speichern mit Formatierungen
        with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
            workbook = writer.book

            # Beide Baselines auf einem Blatt speichern
            sheet_name = "Baseline Comparison"
            df_baseline1.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2, startcol=0)  # Start bei Zeile 2
            df_baseline2.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2, startcol=len(all_fields) + 2)

            worksheet = writer.sheets[sheet_name]

            # **Formatierungen definieren**
            header_format = workbook.add_format({
                "bold": True, "bg_color": "#007BFF", "color": "white",
                "align": "center", "border": 1
            })
            merged_header_format = workbook.add_format({
                "bold": True, "bg_color": "#0056A4", "color": "white",
                "align": "center", "border": 1, "font_size": 14
            })
            normal_format = workbook.add_format({
                "text_wrap": True, "valign": "top", "border": 1
            })
            changed_format = workbook.add_format({
                "bg_color": "#FFA07A", "border": 1, "text_wrap": True
            })  # Orange Markierung für Unterschiede
            thick_border_format = workbook.add_format({
                "top": 2, "border": 1
            })  # Dicke Trennlinie zwischen Baselines

            # **Merged Cells für die Baseline-Namen über den Headern**
            half = len(all_fields)
            worksheet.merge_range(0, 0, 0, half - 1, baseline1_name, merged_header_format)
            worksheet.merge_range(0, half + 2, 0, (2 * half) + 1, baseline2_name, merged_header_format)

            # **Header-Formatierung setzen**
            for col_num in range(half):
                worksheet.write(1, col_num, all_fields[col_num], header_format)  # Header in Zeile 1
                worksheet.write(1, col_num + half + 2, all_fields[col_num], header_format)  # Header in Zeile 1

            # **Spaltenbreiten & automatischer Zeilenumbruch**
            for col_num in range((2 * half) + 2):
                worksheet.set_column(col_num, col_num, 20, normal_format)

            # **Trennlinie zwischen den beiden Baselines einfügen**
            worksheet.set_column(half + 1, half + 1, None, thick_border_format)

            # **Unterschiede markieren**
            for row_idx in range(2, len(df_baseline1) + 2):  # Starte in Zeile 2
                for col_idx in range(half):
                    val1 = df_baseline1.iloc[row_idx - 2, col_idx]
                    val2 = df_baseline2.iloc[row_idx - 2, col_idx]
                    if val1 != val2:  # Falls Unterschied, markiere beide Zellen
                        worksheet.write(row_idx, col_idx, val1, changed_format)
                        worksheet.write(row_idx, col_idx + half + 2, val2, changed_format)

            # **Tabellen-Format für beide Baselines setzen**
            table1_range = f"A2:{self.col_num_to_excel(half - 1)}{len(df_baseline1) + 2}"
            table2_range = f"{self.col_num_to_excel(half + 2)}2:{self.col_num_to_excel((2 * half) + 1)}{len(df_baseline2) + 2}"

            worksheet.add_table(table1_range, {"columns": [{"header": col} for col in all_fields], "style": "Table Style Light 9"})
            worksheet.add_table(table2_range, {"columns": [{"header": col} for col in all_fields], "style": "Table Style Light 11"})

        messagebox.showinfo("Success", f"Excel file successfully saved:\n{file_path}")
        
        os.startfile(file_path)
    
    def create_html_report_diff(self, differences, baseline1_name, baseline2_name):
        """Erstellt eine HTML-Tabelle mit nur den Unterschieden zwischen den Baselines."""
        
        if not hasattr(self, "selected_attributes") or not self.selected_attributes:
            messagebox.showerror("Fehler", "Keine Attribute ausgewählt!")
            return

        selected_fields = ["ID", "Name", "Description"] + [field for field in self.selected_attributes if field not in ["ID", "Name", "Description"]]

        # Überprüfen, ob es überhaupt Unterschiede gibt
        relevant_differences = []
        
        for diff in differences:
            item_has_difference = any(
                value1 != value2
                for field, (value1, value2) in diff["differences"].items()
                if field in selected_fields  # Nur gewählte Attribute prüfen
            )
            if item_has_difference:
                relevant_differences.append(diff)

        # Falls keine Unterschiede gefunden wurden
        if not relevant_differences:
            messagebox.showinfo("Info", "Keine Unterschiede gefunden. Keine HTML-Datei erstellt.")
            return

        html_path = "baseline_comparison.html"
        html_content = f"""
        <html><head><title>Baseline Differences</title>
        <style>
            body {{ font-family: Arial; margin: 20px; }}
            table {{ width: 100%; border-collapse: collapse; }}
            th, td {{ border: 1px solid black; padding: 8px; text-align: left; }}
            th {{ background-color: #007BFF; color: white; }}
            .changed {{ background-color: #ffcccc; }}
            .container {{ display: flex; justify-content: space-between; }}
            .table-wrapper {{ width: 48%; }}
        </style></head><body>
        <h2>Baseline Differences: {baseline1_name} vs. {baseline2_name}</h2>
        """

        for diff in relevant_differences:
            html_content += "<div class='container'><div class='table-wrapper'><table><tr><th>Field</th><th>{}</th></tr>".format(baseline1_name)
            for field in selected_fields:
                val1 = str(diff["differences"].get(field, ["N/A", "N/A"])[0])
                val2 = str(diff["differences"].get(field, ["N/A", "N/A"])[1])
                style = "class='changed'" if val1 != val2 else ""
                html_content += f"<tr {style}><td>{field}</td><td>{val1}</td></tr>"
            html_content += "</table></div>"

            html_content += "<div class='table-wrapper'><table><tr><th>Field</th><th>{}</th></tr>".format(baseline2_name)
            for field in selected_fields:
                val1 = str(diff["differences"].get(field, ["N/A", "N/A"])[0])
                val2 = str(diff["differences"].get(field, ["N/A", "N/A"])[1])
                style = "class='changed'" if val1 != val2 else ""
                html_content += f"<tr {style}><td>{field}</td><td>{val2}</td></tr>"
            html_content += "</table></div></div><br>"

        html_content += "</body></html>"

        with open(html_path, "w", encoding="utf-8") as file:
            file.write(html_content)

        webbrowser.open(html_path)
        messagebox.showinfo("Success", "Differences were saved as HTML.")
    
    
    def create_html_report(self, differences, baseline1_name, baseline2_name):
        """Erstellt HTML-Bericht mit ALLEN Attributen des Trackers für beide Baselines"""
        html_path = "baseline_comparison.html"
        html_content = f"""
        <html><head><title>Baseline Vergleich</title>
        <style>
            body {{ font-family: Arial; margin: 20px; }}
            table {{ width: 100%; border-collapse: collapse; }}
            th, td {{ border: 1px solid black; padding: 8px; text-align: left; }}
            th {{ background-color: #007BFF; color: white; }}
            .changed {{ background-color: #ffcccc; }}
            .container {{ display: flex; justify-content: space-between; }}
            .table-wrapper {{ width: 48%; }}
        </style></head><body>
        <h2>Baseline Vergleich: {baseline1_name} vs. {baseline2_name}</h2>
        """

        if not differences:
            html_content += "<p>Keine Unterschiede gefunden.</p>"
        else:
            for diff in differences:
                tracker_name = diff["tracker_name"]
                item_id = diff["item_id"]
                item_name = diff["item_name"]

                html_content += f"""
                <h3>Tracker: {tracker_name} - Item ID: {item_id} - Name: {item_name}</h3>
                <div class="container">
                    <div class="table-wrapper">
                        <table>
                            <tr>
                                <th>Feld</th>
                                <th>{baseline1_name}</th>
                            </tr>
                """

                for field, values in diff["differences"].items():
                    # Normalisiere die Werte für den Vergleich
                    val1 = str(values[0]).strip() if values[0] is not None else "Nicht vorhanden"
                    val2 = str(values[1]).strip() if values[1] is not None else "Nicht vorhanden"

                    # FALSCH: Vergleich mit values[0] & values[1] (führt zu Fehlern)
                    # style = "class='changed'" if values[0] != values[1] else ""

                    # ✅ RICHTIG: Vergleich der normalisierten Werte
                    style = "class='changed'" if val1 != val2 else ""

                    html_content += f"""
                    <tr {style}>
                        <td>{field}</td>
                        <td>{val1}</td>
                    </tr>
                    """
                html_content += "</table></div>"

                html_content += """
                    <div class="table-wrapper">
                        <table>
                            <tr>
                                <th>Feld</th>
                                <th>{baseline2_name}</th>
                            </tr>
                """
                for field, values in diff["differences"].items():
                    # Normalisierung (wie oben)
                    val1 = str(values[0]).strip() if values[0] is not None else "Nicht vorhanden"
                    val2 = str(values[1]).strip() if values[1] is not None else "Nicht vorhanden"

                    # Vergleich der normalisierten Werte
                    style = "class='changed'" if val1 != val2 else ""

                    html_content += f"""
                    <tr {style}>
                        <td>{field}</td>
                        <td>{val2}</td>
                    </tr>
                    """
                html_content += "</table></div></div><br>"

        html_content += "</body></html>"

        with open(html_path, "w", encoding="utf-8") as file:
            file.write(html_content)

        webbrowser.open(html_path)
        messagebox.showinfo("Erfolg", "Der Vergleich wurde als HTML gespeichert.")

root = tk.Tk()
app = BaselineComparisonApp(root)
root.mainloop()

## Ich möhte dass die excel danach automatisch geöffnet wird