import tkinter as tk
from tkinter import ttk, messagebox
import math
from datetime import datetime, timedelta
import pandas as pd
import os


class AcademicApp(tk.Tk):
    HOURS_PER_DAY = 8  # Constantes de classe
    SEMESTER_DURATION_DAYS = 100  # 100 jours académiques par semestre

    def __init__(self):
        super().__init__()
        self.title("Gestion Académique")
        self.geometry("1000x800")
        self.configure(bg="#F5F5F5")
        self.start_date_text = ""
        self.completed_hours = {}
        self.semester_dates = []
        self.current_page = None
        self.create_widgets()

    def create_widgets(self):
        self.clear_frame()
        self.current_page = "home"
        canvas = tk.Canvas(self, bg="#F5F5F5", highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        label = ttk.Label(self, text="Bienvenue dans Gestion Académique", font=("Helvetica", 32, "bold"), foreground="#006400", background="#F5F5F5")
        label.place(relx=0.5, rely=0.3, anchor="center")
        intro = ttk.Label(self, text="Suivez les retards académiques par promotion !", font=("Arial", 14), foreground="black", background="#F5F5F5", wraplength=600)
        intro.place(relx=0.5, rely=0.45, anchor="center")
        btn_frame = ttk.Frame(self, style="TFrame")
        btn_frame.place(relx=0.5, rely=0.6, anchor="center")
        btn_delay = ttk.Button(btn_frame, text="Calculer mon retard académique", command=self.show_delay_form, style="Green.TButton")
        btn_delay.pack(pady=10, padx=10)
        btn_schedule = ttk.Button(btn_frame, text="Ordonnancer des cours", command=self.show_schedule_form, style="Green.TButton")
        style = ttk.Style()
        btn_schedule.pack(pady=10, padx=10)
        style.configure("Green.TButton", font=("Arial", 14), padding=10, background="#00FF00", foreground="black")
        style.map("Green.TButton", background=[("active", "#00CC00")])
        style.configure("TFrame", background="#F5F5F5")
        style.configure("TLabel", background="#F5F5F5", foreground="black")
        style.configure("TEntry", fieldbackground="#FFFFFF", foreground="black")

    def clear_frame(self):
        for widget in self.winfo_children():
            widget.destroy()
    def show_schedule_form(self):
        self.clear_frame()
        self.current_page = "schedule_form"
        self.courses = []
        label = ttk.Label(self, text="Ordonnancement des cours", font=("Arial", 20, "bold"), foreground="black")
        label.pack(pady=20, padx=20)
        ttk.Label(self, text="Nom du cours:", font=("Arial", 12), foreground="black").pack(pady=5, padx=20)
        course_name = ttk.Entry(self, font=("Arial", 12))
        course_name.pack(pady=5, padx=20)
        ttk.Label(self, text="Volume horaire:", font=("Arial", 12), foreground="black").pack(pady=5, padx=20)
        course_hours = ttk.Entry(self, font=("Arial", 12))
        course_hours.pack(pady=5, padx=20)
        frame = ttk.Frame(self)
        frame.pack(pady=10, padx=20, fill="both", expand=True)
        canvas = tk.Canvas(frame, bg="#F5F5F5")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        course_listbox = tk.Listbox(scrollable_frame, height=10, width=50, font=("Arial", 10), bg="#FFFFFF", fg="black", bd=2, relief="groove")
        course_listbox.pack(pady=5, padx=5)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def add_course():
            name = course_name.get()
            try:
                hours = float(course_hours.get())
                if hours <= 0:
                    raise ValueError("Volume horaire doit être positif.")
                if not name.strip():
                    raise ValueError("Nom du cours ne peut pas être vide.")
                self.courses.append({"name": name.strip(), "hours": hours, "executed": 0})
                course_listbox.insert(tk.END, f"{name.strip()} - {hours} heures")
                course_name.delete(0, tk.END)
                course_hours.delete(0, tk.END)
            except ValueError as e:
                messagebox.showerror("Erreur", str(e))

        ttk.Button(self, text="Ajouter cours", command=add_course, style="Green.TButton").pack(pady=5, padx=20)
        ttk.Button(self, text="Ordonnancer", command=self.generate_schedule, style="Green.TButton").pack(pady=5, padx=20)
        ttk.Button(self, text="Retour", command=self.create_widgets, style="Green.TButton").pack(pady=5, padx=20)

    def pfair_schedule(self, courses_to_schedule):
        if not courses_to_schedule:
            return []
        total_hours = sum(c["hours"] for c in courses_to_schedule)
        max_slots = math.ceil(total_hours)
        schedule = []
        tasks = [{"name": c["name"], "weight": 1/c["hours"], "executions": 0, "total_hours": c["hours"]} for c in courses_to_schedule]
        for t in range(max_slots):
            urgent = []
            possible = []
            forbidden = []
            delays = {}
            alpha = {}
            scheduled_task = None
            all_done = all(task["executions"] >= task["total_hours"] for task in tasks)
            if all_done:
                break
            for task in tasks:
                if task["executions"] < task["total_hours"]:
                    expected = task["weight"] * (t + 1)
                    delay = expected - task["executions"]
                    delays[task["name"]] = delay
                    alpha[task["name"]] = math.floor(expected)
                    if delay >= 1:
                        urgent.append(task["name"])
                    elif delay > 0:
                        possible.append(task["name"])
                    else:
                        forbidden.append(task["name"])
                else:
                    forbidden.append(task["name"])
            candidates = urgent if urgent else possible
            if candidates:
                max_delay = -1
                selected_task = None
                for task_name in candidates:
                    if delays[task_name] > max_delay:
                        max_delay = delays[task_name]
                        selected_task = task_name
                if selected_task:
                    scheduled_task = selected_task
                    for task in tasks:
                        if task["name"] == selected_task:
                            task["executions"] += 1
                            break
            schedule.append({
                "time": t,
                "urgent": urgent,
                "possible": possible,
                "forbidden": forbidden,
                "delays": delays,
                "alpha": alpha,
                "scheduled": scheduled_task
            })
        return schedule

    def generate_schedule(self):
        if not self.courses:
            messagebox.showerror("Erreur", "Aucun cours ajouté.")
            return
        self.schedule_data = self.pfair_schedule(self.courses)
        total_hours = sum(c["hours"] for c in self.courses)
        self.total_weeks = math.ceil(total_hours / 40)
        self.current_week = 0
        self.show_week_schedule(week=0)

    def show_week_schedule(self, week):
        self.clear_frame()
        self.current_page = "week_schedule"
        if week < 0 or week >= self.total_weeks:
            messagebox.showerror("Erreur", "Semaine invalide.")
            self.generate_schedule()
            return
        self.current_week = week
        start_slot = week * 40
        end_slot = min((week + 1) * 40, len(self.schedule_data))
        week_data = self.schedule_data[start_slot:end_slot]
        label = ttk.Label(self, text=f"Tableau d'ordonnancement - Semaine {week + 1}", font=("Arial", 20, "bold"), foreground="black")
        label.pack(pady=10, padx=20)
        canvas = tk.Canvas(self, bg="#F5F5F5")
        x_scrollbar = ttk.Scrollbar(self, orient="horizontal", command=canvas.xview)
        y_scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        headers = ["Temps", "Urgent", "Possible", "Interdit"] + [f"Retard {c['name']}" for c in self.courses] + [f"Alpha {c['name']}" for c in self.courses] + ["Planifiée"]
        for col, header in enumerate(headers):
            ttk.Label(scrollable_frame, text=header[:20], font=("Arial", 8, "bold"), borderwidth=1, relief="groove", padding=3, width=15, background="#E0E0E0", foreground="black").grid(row=0, column=col, sticky="nsew")
        for row, data in enumerate(week_data, start=1):
            ttk.Label(scrollable_frame, text=data["time"], font=("Arial", 8), borderwidth=1, relief="groove", padding=3, width=15, background="#FFFFFF", foreground="black").grid(row=row, column=0, sticky="nsew")
            urgent_text = ", ".join(data["urgent"])[:30]
            possible_text = ", ".join(data["possible"])[:30]
            forbidden_text = ", ".join(data["forbidden"])[:30]
            ttk.Label(scrollable_frame, text=urgent_text, font=("Arial", 8), borderwidth=1, relief="groove", padding=3, width=15, background="#FFFFFF", foreground="black").grid(row=row, column=1, sticky="nsew")
            ttk.Label(scrollable_frame, text=possible_text, font=("Arial", 8), borderwidth=1, relief="groove", padding=3, width=15, background="#FFFFFF", foreground="black").grid(row=row, column=2, sticky="nsew")
            ttk.Label(scrollable_frame, text=forbidden_text, font=("Arial", 8), borderwidth=1, relief="groove", padding=3, width=15, background="#FFFFFF", foreground="black").grid(row=row, column=3, sticky="nsew")
            for col, course in enumerate(self.courses, start=4):
                delay = data["delays"].get(course["name"], 0)
                ttk.Label(scrollable_frame, text=f"{delay:.2f}", font=("Arial", 8), borderwidth=1, relief="groove", padding=3, width=15, background="#FFFFFF", foreground="black").grid(row=row, column=col, sticky="nsew")
            for col, course in enumerate(self.courses, start=4+len(self.courses)):
                alpha = data["alpha"].get(course["name"], 0)
                ttk.Label(scrollable_frame, text=alpha, font=("Arial", 8), borderwidth=1, relief="groove", padding=3, width=15, background="#FFFFFF", foreground="black").grid(row=row, column=col, sticky="nsew")
            scheduled_text = data["scheduled"] or "-"
            ttk.Label(scrollable_frame, text=scheduled_text[:20], font=("Arial", 8), borderwidth=1, relief="groove", padding=3, width=15, background="#FFFFFF", foreground="black").grid(row=row, column=4+2*len(self.courses), sticky="nsew")
        canvas.pack(side="top", fill="both", expand=True)
        x_scrollbar.pack(side="bottom", fill="x")
        y_scrollbar.pack(side="right", fill="y")
        nav_frame = ttk.Frame(self)
        nav_frame.pack(pady=10, padx=20)
        if week > 0:
            ttk.Button(nav_frame, text="Semaine précédente", command=lambda: self.show_week_schedule(week-1), style="Green.TButton").pack(side="left", padx=5)
        ttk.Button(nav_frame, text="Convertir en calendrier réel", command=self.show_calendar, style="Green.TButton").pack(side="left", padx=5)
        if week < self.total_weeks - 1:
            ttk.Button(nav_frame, text="Semaine suivante", command=lambda: self.show_week_schedule(week+1), style="Green.TButton").pack(side="left", padx=5)
        ttk.Button(nav_frame, text="Retour", command=self.show_schedule_form, style="Green.TButton").pack(side="left", padx=5)
        export_btn = ttk.Button(nav_frame, text="Exporter en Excel", command=self.export_to_excel, style="Green.TButton")
        export_btn.pack(side="left", padx=5)

    def show_calendar(self):
        self.clear_frame()
        self.current_page = "calendar"
        label = ttk.Label(self, text=f"Calendrier réel - Semaine {self.current_week + 1}", font=("Arial", 20, "bold"), foreground="black")
        label.pack(pady=10, padx=20)
        canvas = tk.Canvas(self, bg="#F5F5F5")
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        start_slot = self.current_week * 40
        end_slot = min((self.current_week + 1) * 40, len(self.schedule_data))
        week_data = self.schedule_data[start_slot:end_slot]
        days = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi"]
        hours_per_day = 8
        for slot, data in enumerate(week_data):
            if data["scheduled"]:
                day_index = (slot // hours_per_day) % 5
                hour_index = slot % hours_per_day
                day = days[day_index]
                start_hour = 8 + hour_index
                time_str = f"{day} {start_hour:02d}:00-{start_hour+1:02d}:00"
                ttk.Label(scrollable_frame, text=f"{time_str} : {data['scheduled']}", font=("Arial", 12), foreground="black").pack(anchor="w", padx=10, pady=2)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        nav_frame = ttk.Frame(self)
        nav_frame.pack(pady=10, padx=20)
        if self.current_week > 0:
            ttk.Button(nav_frame, text="Semaine précédente", command=lambda: self.show_calendar_week(self.current_week-1), style="Green.TButton").pack(side="left", padx=5)
        if self.current_week < self.total_weeks - 1:
            ttk.Button(nav_frame, text="Semaine suivante", command=lambda: self.show_calendar_week(self.current_week+1), style="Green.TButton").pack(side="left", padx=5)
        ttk.Button(nav_frame, text="Retour", command=lambda: self.show_week_schedule(self.current_week), style="Green.TButton").pack(side="left", padx=5)
        export_btn = ttk.Button(nav_frame, text="Exporter en Excel", command=self.export_to_excel, style="Green.TButton")
        export_btn.pack(side="left", padx=5)

    def show_calendar_week(self, week):
        self.current_week = week
        self.show_calendar()


    def show_delay_form(self):
        self.clear_frame()
        self.current_page = "delay_form"
        label = ttk.Label(self, text="Calcul du Retard Académique", font=("Arial", 20, "bold"), foreground="black")
        label.pack(pady=20, padx=20)
        ttk.Label(self, text="Date de début de la licence (DD/MM/YYYY):", font=("Arial", 12), foreground="black").pack(pady=5, padx=20)
        start_date = ttk.Entry(self, font=("Arial", 12))
        start_date.pack(pady=5, padx=20)

        def validate_and_proceed():
            try:
                self.start_date_text = start_date.get()
                datetime.strptime(self.start_date_text, "%d/%m/%Y")
                self.show_hours_input()
            except ValueError:
                messagebox.showerror("Erreur", "Format de date invalide (utilisez DD/MM/YYYY).")

        ttk.Button(self, text="Suivant", command=validate_and_proceed, style="Green.TButton").pack(pady=10, padx=20)
        ttk.Button(self, text="Retour", command=self.create_widgets, style="Green.TButton").pack(pady=5, padx=20)

    def show_hours_input(self):
        self.clear_frame()
        self.current_page = "hours_input"
        self.completed_hours = {}
        start_date = datetime.strptime(self.start_date_text, "%d/%m/%Y")
        current_date = datetime(2025, 6, 26, 3, 26)  # 03:26 AM GMT, 26/06/2025
        ESTIMATED_SEMESTER_DAYS = 150  # 150 jours pour estimation

        # Calcul des dates de début et fin dynamiques
        self.semester_dates = []
        for i in range(6):
            sem_start_days = i * ESTIMATED_SEMESTER_DAYS
            sem_end_days = (i + 1) * ESTIMATED_SEMESTER_DAYS
            sem_start_date = start_date + timedelta(days=sem_start_days)
            sem_end_date = start_date + timedelta(days=sem_end_days)
            self.semester_dates.append((sem_start_date.strftime('%d/%m/%Y'), sem_end_date.strftime('%d/%m/%Y')))

        label = ttk.Label(self, text="Saisie des heures VHF réalisées", font=("Arial", 20, "bold"), foreground="black")
        label.pack(pady=20, padx=20)
        ttk.Label(self, text="Entrez les heures réalisées pour chaque semestre (0 à 360h).", font=("Arial", 12), foreground="black").pack(pady=5, padx=20)

        input_frame = ttk.Frame(self)
        input_frame.pack(pady=10, padx=20, fill="x")
        self.hour_entries = {}
        semesters = [f"S{i+1}" for i in range(6)]
        for sem in semesters:
            idx = int(sem[1:]) - 1
            ttk.Label(input_frame, text=f"{sem} (Début: {self.semester_dates[idx][0]}, Fin: {self.semester_dates[idx][1]}):", font=("Arial", 12), foreground="black").grid(row=semesters.index(sem), column=0, pady=5, padx=5, sticky="w")
            entry = ttk.Entry(input_frame, font=("Arial", 12), width=10)
            entry.grid(row=semesters.index(sem), column=1, pady=5, padx=5)
            entry.insert(0, "0")  # Valeur par défaut
            self.hour_entries[sem] = entry

        def validate_and_calculate():
            try:
                self.completed_hours = {}
                for sem, entry in self.hour_entries.items():
                    hours = entry.get().strip()
                    if not hours:
                        self.completed_hours[sem] = 0.0
                    else:
                        hours = float(hours)
                        if hours < 0 or hours > 360:
                            raise ValueError(f"Heures pour {sem} doivent être entre 0 et 360.")
                        self.completed_hours[sem] = hours
                self.calculate_delays()
            except ValueError as e:
                messagebox.showerror("Erreur", str(e))

        ttk.Button(self, text="Calculer", command=validate_and_calculate, style="Green.TButton").pack(pady=10, padx=20)
        ttk.Button(self, text="Retour", command=self.show_delay_form, style="Green.TButton").pack(pady=5, padx=20)

    def calculate_delays(self):
        self.clear_frame()
        self.current_page = "delay_report"
        start_date = datetime.strptime(self.start_date_text, "%d/%m/%Y")
        current_date = datetime(2025, 6, 26, 3, 26)  # 03:26 AM GMT, 26/06/2025
        SEMESTER_DURATION_HOURS = 800  # 100 jours × 8h/jour
        VHF_RATE = 0.45
        TOTAL_ANNUAL_HOURS = 1440  # 9 mois

        semester_delays = {}
        task_delays = {f"S{i+1}": [] for i in range(6)}
        total_delay_hours = 0

        # Calcul du temps écoulé depuis le début
        total_days_elapsed = (current_date - start_date).days
        total_hours_elapsed = total_days_elapsed * self.HOURS_PER_DAY  # Utilisation de la constante de classe

        # Déterminer les semestres actifs dynamiquement
        semesters_completed = min(max(0, total_days_elapsed // self.SEMESTER_DURATION_DAYS), 5)  # Nombre de semestres achevés
        current_semester_idx = semesters_completed if total_days_elapsed % self.SEMESTER_DURATION_DAYS == 0 else semesters_completed

        for i in range(6):
            sem = f"S{i+1}"
            sem_start_date_str, sem_end_date_str = self.semester_dates[i]
            sem_start_date = datetime.strptime(sem_start_date_str, "%d/%m/%Y")
            sem_end_date = datetime.strptime(sem_end_date_str, "%d/%m/%Y")

            if current_date < sem_start_date:  # Semestre futur
                t_hours = 0
            elif current_date >= sem_end_date:  # Semestre achevé
                t_hours = SEMESTER_DURATION_HOURS
            else:  # Semestre en cours
                days_in_semester = (sem_end_date - sem_start_date).days
                days_elapsed_in_semester = (current_date - sem_start_date).days
                t_hours = min(SEMESTER_DURATION_HOURS, (days_elapsed_in_semester / days_in_semester) * SEMESTER_DURATION_HOURS)

            expected_hours = VHF_RATE * t_hours
            realized_hours = self.completed_hours.get(sem, 0.0)
            nhr = max(0, expected_hours - realized_hours) if expected_hours >= realized_hours else (expected_hours - realized_hours)
            ra = nhr / VHF_RATE if VHF_RATE > 0 else 0

            semester_delays[sem] = ra
            task_delays[sem].append(("Total semestre", ra))
            total_delay_hours += ra

        self.delay_report_data = {
            "Semestre": list(semester_delays.keys()),
            "Retard (heures)": list(semester_delays.values()),
            "Retard (jours)": [delay / 8 for delay in semester_delays.values()],
            "Retard (mois)": [delay / (8 * 20) for delay in semester_delays.values()]
        }

        canvas = tk.Canvas(self, bg="#F5F5F5")
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        ttk.Label(scrollable_frame, text="Rapport des retards académiques", font=("Arial", 20, "bold"), foreground="black").pack(pady=10, padx=20)
        ttk.Label(scrollable_frame, text=f"Temps écoulé à t = {total_hours_elapsed} heures ({datetime.now().strftime('%d/%m/%Y %H:%M')})", font=("Arial", 12), foreground="black").pack(anchor="w", padx=20)
        ttk.Label(scrollable_frame, text="Dates des semestres:", font=("Arial", 14, "bold"), foreground="black").pack(anchor="w", pady=5, padx=20)
        for i, (start_d, end_d) in enumerate(self.semester_dates, 1):
            ttk.Label(scrollable_frame, text=f"S{i}: {start_d} - {end_d}", font=("Arial", 12), foreground="black").pack(anchor="w", padx=20)
        ttk.Label(scrollable_frame, text="Retards calculés par semestre:", font=("Arial", 14, "bold"), foreground="black").pack(anchor="w", pady=10, padx=20)
        for sem in semester_delays:
            delay = semester_delays[sem]
            delay_days = delay / 8
            delay_months = delay / (8 * 20)
            ttk.Label(scrollable_frame, text=f"{sem}: {delay:.2f} heures ({delay_days:.2f} jours, {delay_months:.2f} mois)", font=("Arial", 12, "bold"), foreground="black").pack(anchor="w", pady=2, padx=20)
        total_delay_days = total_delay_hours / 8
        total_delay_months = total_delay_hours / (8 * 20)
        ttk.Label(scrollable_frame, text=f"Retard total: {total_delay_hours:.2f} heures ({total_delay_days:.2f} jours, {total_delay_months:.2f} mois)", font=("Arial", 12, "bold"), foreground="black").pack(anchor="w", padx=20, pady=10)
        export_btn = ttk.Button(scrollable_frame, text="Exporter Rapport Retards (Excel)", command=self.export_to_excel, style="Green.TButton")
        export_btn.pack(pady=10, padx=20)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        ttk.Button(self, text="Retour", command=self.show_hours_input, style="Green.TButton").pack(side="bottom", pady=5, padx=20)

    def export_to_excel(self):
        if self.current_page == "week_schedule" or self.current_page == "calendar":
            if hasattr(self, 'schedule_data') and self.schedule_data:
                start_slot = self.current_week * 40
                end_slot = min((self.current_week + 1) * 40, len(self.schedule_data))
                week_data = self.schedule_data[start_slot:end_slot]
                if self.current_week == 0 and not week_data:
                    messagebox.showwarning("Avertissement", "Aucune donnée à exporter pour cette semaine.")
                    return
                headers = ["Temps", "Urgent", "Possible", "Interdit"] + [f"Retard {c['name']}" for c in self.courses] + [f"Alpha {c['name']}" for c in self.courses] + ["Planifiée"]
                data = []
                for d in week_data:
                    row = [d["time"], ", ".join(d["urgent"]), ", ".join(d["possible"]), ", ".join(d["forbidden"])]
                    row.extend([d["delays"].get(k, 0) for k in d["delays"].keys()])
                    row.extend([d["alpha"].get(k, 0) for k in d["alpha"].keys()])
                    row.append(d["scheduled"] or "-")
                    data.append(row)
                df = pd.DataFrame(data, columns=headers)
                filename = f"ordonnancement_semaine_{self.current_week + 1}.xlsx"
            else:
                messagebox.showwarning("Avertissement", "Aucune donnée à exporter.")
                return
        elif self.current_page == "delay_report":
            if hasattr(self, 'delay_report_data') and self.delay_report_data and self.delay_report_data["Semestre"]:
                df = pd.DataFrame({
                    "Semestre": self.delay_report_data["Semestre"],
                    "Retard (heures)": self.delay_report_data["Retard (heures)"],
                    "Retard (jours)": self.delay_report_data["Retard (jours)"],
                    "Retard (mois)": self.delay_report_data["Retard (mois)"]
                })
                filename = "rapport_retards.xlsx"
            else:
                messagebox.showwarning("Avertissement", "Aucune donnée à exporter pour le rapport.")
                return
        else:
            messagebox.showwarning("Avertissement", "Aucune donnée à exporter pour cette page.")
            return
        
        df.to_excel(filename, index=False)
        messagebox.showinfo("Succès", f"Fichier exporté avec succès sous : {os.path.abspath(filename)}")

if __name__ == "__main__":
    app = AcademicApp()
    app.mainloop()
