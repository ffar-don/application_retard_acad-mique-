import tkinter as tk
from tkinter import ttk, messagebox
import math
from datetime import datetime, timedelta
import pandas as pd
import os

class AcademicApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gestion Académique")
        self.geometry("1000x800")
        self.configure(bg="#F5F5F5")
        self.courses = []
        self.semester_courses = {}
        self.start_date_text = ""
        self.current_date_text = ""
        self.schedule_data = []
        self.current_week = 0
        self.total_weeks = 0
        self.completed_semesters = {}
        self.semester_dates = []
        self.delay_report_data = {}  # Stockage des données du rapport
        self.current_page = None  # Pour suivre la page active
        self.create_widgets()

    def create_widgets(self):
        self.clear_frame()
        self.current_page = "home"
        # Fond gradient avec canvas
        canvas = tk.Canvas(self, bg="#F5F5F5", highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        gradient = tk.PhotoImage(width=1000, height=800)
        for x in range(1000):
            for y in range(800):
                color = f'#{int(245 - y/800*50):02x}{int(245 - y/800*50):02x}{255:02x}'
                gradient.put(color, (x, y))
        canvas.create_image(0, 0, image=gradient, anchor="nw")
        # Image décorative (remplacez par le chemin de votre image)
        try:
            img = tk.PhotoImage(file="universite-norbert-zongo (1).jpg")  # Ajoutez une image locale si disponible
            canvas.create_image(500, 100, image=img)
        except:
            pass
        # Titre stylisé
        label = ttk.Label(self, text="Bienvenue dans Gestion Académique", font=("Helvetica", 32, "bold"), foreground="#006400", background="#F5F5F5")
        label.place(relx=0.5, rely=0.3, anchor="center")
        # Texte d'introduction
        intro = ttk.Label(self, text="Organisez vos cours et suivez vos progrès avec facilité !\nCommencez dès maintenant pour optimiser votre parcours académique.", 
                          font=("Arial", 14), foreground="black", background="#F5F5F5", wraplength=600)
        intro.place(relx=0.5, rely=0.45, anchor="center")
        # Boutons
        btn_frame = ttk.Frame(self, style="TFrame")
        btn_frame.place(relx=0.5, rely=0.6, anchor="center")
        btn_schedule = ttk.Button(btn_frame, text="Ordonnancer des cours", command=self.show_schedule_form, style="Green.TButton")
        btn_schedule.pack(pady=10, padx=10)
        btn_delay = ttk.Button(btn_frame, text="Calculer mon retard académique", command=self.show_delay_form, style="Green.TButton")
        btn_delay.pack(pady=10, padx=10)
        style = ttk.Style()
        style.configure("Green.TButton", font=("Arial", 14), padding=10, background="#00FF00", foreground="black")
        style.map("Green.TButton", background=[("active", "#00CC00")])
        style.configure("TFrame", background="#F5F5F5")
        style.configure("TLabel", background="#F5F5F5", foreground="black")
        style.configure("TCheckbutton", background="#F5F5F5", foreground="black")
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
        # Bouton d'exportation
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

    def show_calendar_week(self, week):
        self.current_week = week
        self.show_calendar()

    def show_delay_form(self):
        self.clear_frame()
        self.current_page = "delay_form"
        label = ttk.Label(self, text="Calcul du retard académique", font=("Arial", 20, "bold"), foreground="black")
        label.pack(pady=20, padx=20)
        ttk.Label(self, text="Date de début de la licence (DD/MM/YYYY):", font=("Arial", 12), foreground="black").pack(pady=5, padx=20)
        start_date = ttk.Entry(self, font=("Arial", 12))
        start_date.pack(pady=5, padx=20)
        ttk.Label(self, text="Date de calcul (DD/MM/YYYY):", font=("Arial", 12), foreground="black").pack(pady=5, padx=20)
        current_date = ttk.Entry(self, font=("Arial", 12))
        current_date.pack(pady=5, padx=20)
        def validate_and_proceed():
            try:
                self.start_date_text = start_date.get()
                self.current_date_text = current_date.get()
                datetime.strptime(self.start_date_text, "%d/%m/%Y")
                datetime.strptime(self.current_date_text, "%d/%m/%Y")
                self.show_completed_semesters()
            except ValueError:
                messagebox.showerror("Erreur", "Format de date invalide (utilisez DD/MM/YYYY).")
        ttk.Button(self, text="Suivant", command=validate_and_proceed, style="Green.TButton").pack(pady=5, padx=20)
        ttk.Button(self, text="Retour", command=self.create_widgets, style="Green.TButton").pack(pady=5, padx=20)

    def show_completed_semesters(self):
        self.clear_frame()
        self.current_page = "completed_semesters"
        label = ttk.Label(self, text="Sélectionnez les semestres complétés", font=("Arial", 20, "bold"), foreground="black")
        label.pack(pady=20, padx=20)
        ttk.Label(self, text="Cochez les semestres terminés avant la date de calcul, ou laissez décoché si non terminé.", font=("Arial", 12), foreground="black").pack(pady=5, padx=20)
        start_date = datetime.strptime(self.start_date_text, "%d/%m/%Y")
        current_date = datetime.strptime(self.current_date_text, "%d/%m/%Y")
        time_diff = (current_date - start_date).days
        semesters = ["S1", "S2", "S3", "S4", "S5", "S6"]
        self.completed_semesters = {sem: tk.BooleanVar() for sem in semesters}
        self.semester_dates = []
        valid_semesters = []

        for i, sem in enumerate(semesters):
            sem_start_days = i * 150
            sem_end_days = (i + 1) * 150 if i < 5 else time_diff
            sem_start_date = start_date + timedelta(days=sem_start_days)
            sem_end_date = start_date + timedelta(days=sem_end_days) if i < 5 else current_date
            self.semester_dates.append((sem_start_date.strftime('%d/%m/%Y'), sem_end_date.strftime('%d/%m/%Y')))
            if sem_start_date <= current_date:
                valid_semesters.append(sem)
                ttk.Checkbutton(self, text=sem, variable=self.completed_semesters[sem], style="TCheckbutton").pack(anchor="w", padx=20)

        if not valid_semesters:
            messagebox.showerror("Erreur", "La date de calcul est antérieure à tous les semestres. Veuillez choisir une date ultérieure.")
            self.show_delay_form()
            return
        ttk.Button(self, text="Suivant", command=self.show_courses_input, style="Green.TButton").pack(pady=10, padx=20)
        ttk.Button(self, text="Retour", command=self.show_delay_form, style="Green.TButton").pack(pady=5, padx=20)

    def show_courses_input(self):
        self.clear_frame()
        self.current_page = "courses_input"
        start_date = datetime.strptime(self.start_date_text, "%d/%m/%Y")
        current_date = datetime.strptime(self.current_date_text, "%d/%m/%Y")
        self.semester_courses = {sem: [] for sem in ["S1", "S2", "S3", "S4", "S5", "S6"] 
                               if not self.completed_semesters.get(sem, tk.BooleanVar()).get() 
                               and datetime.strptime(self.semester_dates[list(self.completed_semesters.keys()).index(sem)][0], '%d/%m/%Y') <= current_date}
        label = ttk.Label(self, text="Saisie des cours par semestre", font=("Arial", 20, "bold"), foreground="black")
        label.pack(pady=20, padx=20)
        main_frame = ttk.Frame(self)
        main_frame.pack(fill="both", expand=True)

        # Boutons pour sélectionner le semestre
        semester_var = tk.StringVar(value=next(iter(self.semester_courses.keys()), ""))
        ttk.Label(main_frame, text="Sélectionnez un semestre:", font=("Arial", 12), foreground="black").pack(pady=5, padx=20)
        semester_menu = ttk.OptionMenu(main_frame, semester_var, semester_var.get(), *self.semester_courses.keys())
        semester_menu.pack(pady=5, padx=20)

        # Champs de saisie
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(pady=10, padx=20, fill="x")
        ttk.Label(input_frame, text="Nom du cours:", font=("Arial", 12), foreground="black").grid(row=0, column=0, pady=5, padx=5, sticky="w")
        course_name = ttk.Entry(input_frame, font=("Arial", 12))
        course_name.grid(row=0, column=1, pady=5, padx=5)
        ttk.Label(input_frame, text="Heures requises:", font=("Arial", 12), foreground="black").grid(row=1, column=0, pady=5, padx=5, sticky="w")
        required_hours = ttk.Entry(input_frame, font=("Arial", 12))
        required_hours.grid(row=1, column=1, pady=5, padx=5)
        ttk.Label(input_frame, text="Heures réalisées:", font=("Arial", 12), foreground="black").grid(row=2, column=0, pady=5, padx=5, sticky="w")
        completed_hours = ttk.Entry(input_frame, font=("Arial", 12))
        completed_hours.grid(row=2, column=1, pady=5, padx=5)

        # Liste des cours
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(pady=10, padx=20, fill="both", expand=True)
        canvas = tk.Canvas(list_frame, bg="#F5F5F5")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        course_listbox = tk.Listbox(scrollable_frame, height=8, width=50, font=("Arial", 10), bg="#FFFFFF", fg="black", bd=2, relief="groove")
        course_listbox.pack(pady=5, padx=5)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def add_course():
            sem = semester_var.get()
            name = course_name.get().strip()
            req_hours_text = required_hours.get().strip()
            comp_hours_text = completed_hours.get().strip()
            try:
                if not name:
                    raise ValueError("Le champ du nom du cours est vide.")
                if not req_hours_text:
                    raise ValueError("Heures requises ne peuvent pas être vides.")
                if not comp_hours_text:
                    raise ValueError("Heures réalisées ne peuvent pas être vides.")
                req_hours = float(req_hours_text)
                comp_hours = float(comp_hours_text)
                if req_hours <= 0:
                    raise ValueError("Heures requises doivent être positives.")
                if comp_hours < 0:
                    raise ValueError("Heures réalisées ne peuvent pas être négatives.")
                if comp_hours > req_hours:
                    raise ValueError("Heures réalisées ne peuvent pas dépasser les heures requises.")
                self.semester_courses[sem].append({"name": name, "required_hours": req_hours, "completed_hours": comp_hours})
                course_listbox.insert(tk.END, f"{name} - Requis: {req_hours}, Réalisé: {comp_hours}")
                course_name.delete(0, tk.END)
                required_hours.delete(0, tk.END)
                completed_hours.delete(0, tk.END)
            except ValueError as e:
                messagebox.showerror("Erreur", str(e) if str(e).startswith("Heures") or str(e).startswith("Le") else "Veuillez entrer des nombres valides pour les heures (ex. 10.5).")

        ttk.Button(main_frame, text="Ajouter cours", command=add_course, style="Green.TButton").pack(pady=5, padx=20)

        # Boutons de navigation
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10, padx=20, side="bottom", fill="x")
        ttk.Button(button_frame, text="Suivant", command=self.calculate_delays, style="Green.TButton").pack(side="left", padx=5)
        ttk.Button(button_frame, text="Retour", command=self.show_completed_semesters, style="Green.TButton").pack(side="left", padx=5)

    def calculate_delays(self):
        self.clear_frame()
        self.current_page = "delay_report"
        start_date = datetime.strptime(self.start_date_text, "%d/%m/%Y")
        current_date = datetime.strptime(self.current_date_text, "%d/%m/%Y")
        time_diff = (current_date - start_date).days
        if time_diff < 0:
            messagebox.showerror("Erreur", "La date de calcul doit être postérieure à la date de début.")
            self.show_delay_form()
            return
        academic_days = time_diff
        academic_hours = academic_days * 8
        academic_months = time_diff / 20
        
        SEMESTER_DURATION = 800
        SEMESTER_STARTS = [0, 800, 1600, 2400, 3200, 4000]
        SEMESTER_ENDS = [800, 1600, 2400, 3200, 4000, 4800]
        
        semester_delays = {}
        task_delays = {sem: [] for sem in ['S1', 'S2', 'S3', 'S4', 'S5', 'S6']}
        total_delay_hours = 0
        
        valid_semesters = [
            sem for i, sem in enumerate(['S1', 'S2', 'S3', 'S4', 'S5', 'S6'])
            if datetime.strptime(self.semester_dates[i][0], '%d/%m/%Y') <= current_date
        ]
        
        for i, sem in enumerate(['S1', 'S2', 'S3', 'S4', 'S5', 'S6'], 1):
            if sem not in valid_semesters:
                semester_delays[sem] = 0
                task_delays[sem] = []
                continue
            
            t_start = SEMESTER_STARTS[i-1]
            t_end = SEMESTER_ENDS[i-1]
            
            semester_delay = 0
            tasks = []
            
            if sem in self.completed_semesters and self.completed_semesters[sem].get():
                standard_tasks = [
                    {'name': f"Cours Complété {j}", 'required_hours': 100, 'completed_hours': 100}
                    for j in range(5)
                ]
                tasks = standard_tasks
            else:
                tasks = self.semester_courses.get(sem, [])
            
            for task in tasks:
                C_i = task.get('required_hours', 0)
                H_i = task.get('completed_hours', 0)
                U_i = C_i / SEMESTER_DURATION if SEMESTER_DURATION > 0 else 0
                
                if academic_hours > t_end:
                    delay = max(0, C_i - H_i)
                elif academic_hours >= t_start:
                    t_relative = academic_hours - t_start
                    delay = max(0, U_i * t_relative - H_i)
                else:
                    delay = 0
                
                semester_delay += delay
                task_delays[sem].append((task.get('name', 'Inconnu'), delay))
            
            semester_delays[sem] = semester_delay
            total_delay_hours += semester_delay
        
        # Stockage des données pour l'exportation
        self.delay_report_data = {
            "Semestre": list(semester_delays.keys()) + ["Total"],
            "Retard (heures)": list(semester_delays.values()) + [total_delay_hours],
            "Retard (jours)": [delay / 8 for delay in semester_delays.values()] + [total_delay_hours / 8],
            "Retard (mois)": [delay / 160 for delay in semester_delays.values()] + [total_delay_hours / 160]  # 8h/jour * 20j/mois = 160h/mois
        }
        
        canvas = tk.Canvas(self, bg="#F5F5F5")
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        ttk.Label(scrollable_frame, text="Rapport des retards académiques", font=("Arial", 20, "bold"), foreground="black").pack(pady=10, padx=20)
        ttk.Label(scrollable_frame, text=f"Temps écoulé: {academic_hours} heures, {academic_days} jours, {academic_months:.1f} mois", font=("Arial", 12), foreground="black").pack(anchor="w", padx=20)
        ttk.Label(scrollable_frame, text="Dates des semestres prévues:", font=("Arial", 14, "bold"), foreground="black").pack(anchor="w", pady=5, padx=20)
        for i, (start_d, end_d) in enumerate(self.semester_dates, 1):
            ttk.Label(scrollable_frame, text=f"S{i}: {start_d} - {end_d}", font=("Arial", 12), foreground="black").pack(anchor="w", padx=20)
        ttk.Label(scrollable_frame, text="Retards calculés par semestre:", font=("Arial", 14, "bold"), foreground="black").pack(anchor="w", pady=10, padx=20)
        for sem in ['S1', 'S2', 'S3', 'S4', 'S5', 'S6']:
            delay = semester_delays[sem]
            delay_days = delay / 8
            delay_months = delay / 160  # 8h/jour * 20j/mois = 160h/mois
            ttk.Label(scrollable_frame, text=f"Semestre {sem}: {delay:.2f} heures ({delay_days:.2f} jours, {delay_months:.2f} mois)", font=("Arial", 12, "bold"), foreground="black").pack(anchor="w", pady=2, padx=20)
            if task_delays[sem]:
                ttk.Label(scrollable_frame, text=f"  Détails des cours pour {sem}:", font=("Arial", 12), foreground="black").pack(anchor="w", padx=30)
                for task_name, task_delay in task_delays[sem]:
                    ttk.Label(scrollable_frame, text=f"    - {task_name}: {task_delay:.2f} heures", font=("Arial", 12), foreground="black").pack(anchor="w", padx=40)
        total_delay_days = total_delay_hours / 8
        total_delay_months = total_delay_hours / 160
        ttk.Label(scrollable_frame, text=f"Retard total: {total_delay_hours:.2f} heures ({total_delay_days:.2f} jours, {total_delay_months:.2f} mois)", font=("Arial", 12, "bold"), foreground="black").pack(anchor="w", padx=20, pady=10)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        ttk.Button(self, text="Ajuster l'ordonnancement", command=self.adjust_delays, style="Green.TButton").pack(side="bottom", pady=5, padx=20)
        ttk.Button(self, text="Retour", command=self.create_widgets, style="Green.TButton").pack(side="bottom", pady=5, padx=20)
        # Bouton d'exportation
        export_btn = ttk.Button(self, text="Exporter en Excel", command=self.export_to_excel, style="Green.TButton")
        export_btn.pack(side="bottom", pady=5, padx=20)

    def adjust_delays(self):
        self.clear_frame()
        self.current_page = "adjusted_schedule"
        start_date = datetime.strptime(self.start_date_text, "%d/%m/%Y")
        current_date = datetime.strptime(self.current_date_text, "%d/%m/%Y")
        time_diff = (current_date - start_date).days
        academic_hours = time_diff * 8
        
        SEMESTER_DURATION = 800
        SEMESTER_STARTS = {'S1': 0, 'S2': 800, 'S3': 1600, 'S4': 2400, 'S5': 3200, 'S6': 4000}
        SEMESTER_ENDS = {'S1': 800, 'S2': 1600, 'S3': 2400, 'S4': 3200, 'S5': 4000, 'S6': 4800}
        
        valid_semesters = [
            sem for i, sem in enumerate(['S1', 'S2', 'S3', 'S4', 'S5', 'S6'])
            if datetime.strptime(self.semester_dates[i][0], '%d/%m/%Y') <= current_date
        ]
        
        delayed_courses = []
        for sem in valid_semesters:
            t_start = SEMESTER_STARTS[sem]
            t_end = SEMESTER_ENDS[sem]
            tasks = self.semester_courses.get(sem, []) if sem not in self.completed_semesters or not self.completed_semesters[sem].get() else [
                {'name': f"Cours Complété {i}", 'required_hours': 100, 'completed_hours': 100}
                for i in range(5)
            ]
            for task in tasks:
                C_i = task.get('required_hours', 0)
                H_i = task.get('completed_hours', 0)
                U_i = C_i / SEMESTER_DURATION if SEMESTER_DURATION > 0 else 0
                
                if academic_hours > t_end:
                    delay = max(0, C_i - H_i)
                elif academic_hours >= t_start:
                    t_relative = academic_hours - t_start
                    delay = max(0, U_i * t_relative - H_i)
                else:
                    delay = 0
                
                if delay > 0:
                    delayed_courses.append({"name": task.get('name', 'Inconnu'), "hours": delay, "executed": 0})
        
        if not delayed_courses:
            messagebox.showinfo("Information", "Aucun retard à ajuster.")
            self.calculate_delays()
            return
        
        self.schedule_data = self.pfair_schedule(delayed_courses)
        total_hours = sum(c["hours"] for c in delayed_courses)
        self.total_weeks = math.ceil(total_hours / 40)
        self.current_week = 0
        self.show_adjusted_schedule(week=0)

    def show_adjusted_schedule(self, week):
        self.clear_frame()
        if week < 0 or week >= self.total_weeks:
            messagebox.showerror("Erreur", "Semaine invalide.")
            self.adjust_delays()
            return
        self.current_week = week
        start_slot = week * 40
        end_slot = min((week + 1) * 40, len(self.schedule_data))
        week_data = self.schedule_data[start_slot:end_slot]
        if not week_data:
            messagebox.showerror("Erreur", "Aucune donnée disponible pour cette semaine.")
            self.adjust_delays()
            return
        label = ttk.Label(self, text=f"Tableau d'ordonnancement ajusté - Semaine {week + 1}", font=("Arial", 20, "bold"), foreground="black")
        label.pack(pady=10, padx=20)
        canvas = tk.Canvas(self, bg="#F5F5F5")
        x_scrollbar = ttk.Scrollbar(self, orient="horizontal", command=canvas.xview)
        y_scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        headers = ["Temps", "Urgent", "Possible", "Interdit"] + [f"Retard {key}" for key in self.schedule_data[0]['delays'].keys()] + [f"Alpha {key}" for key in self.schedule_data[0]['alpha'].keys()] + ["Planifiée"]
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
            for col, course_name in enumerate(data["delays"].keys(), start=4):
                delay = data["delays"].get(course_name, 0)
                ttk.Label(scrollable_frame, text=f"{delay:.2f}", font=("Arial", 8), borderwidth=1, relief="groove", padding=3, width=15, background="#FFFFFF", foreground="black").grid(row=row, column=col, sticky="nsew")
            for col, course_name in enumerate(data["alpha"].keys(), start=4+len(data["delays"])):
                alpha = data["alpha"].get(course_name, 0)
                ttk.Label(scrollable_frame, text=alpha, font=("Arial", 8), borderwidth=1, relief="groove", padding=3, width=15, background="#FFFFFF", foreground="black").grid(row=row, column=col, sticky="nsew")
            scheduled_text = data["scheduled"] or "-"
            ttk.Label(scrollable_frame, text=scheduled_text[:20], font=("Arial", 8), borderwidth=1, relief="groove", padding=3, width=15, background="#FFFFFF", foreground="black").grid(row=row, column=4+2*len(data["delays"]), sticky="nsew")
        canvas.pack(side="top", fill="both", expand=True)
        x_scrollbar.pack(side="bottom", fill="x")
        y_scrollbar.pack(side="right", fill="y")
        nav_frame = ttk.Frame(self)
        nav_frame.pack(pady=10, padx=20)
        if week > 0:
            ttk.Button(nav_frame, text="Semaine précédente", command=lambda: self.show_adjusted_schedule(week-1), style="Green.TButton").pack(side="left", padx=5)
        ttk.Button(nav_frame, text="Convertir en calendrier réel", command=self.show_adjusted_calendar, style="Green.TButton").pack(side="left", padx=5)
        if week < self.total_weeks - 1:
            ttk.Button(nav_frame, text="Semaine suivante", command=lambda: self.show_adjusted_schedule(week+1), style="Green.TButton").pack(side="left", padx=5)
        ttk.Button(nav_frame, text="Retour", command=self.calculate_delays, style="Green.TButton").pack(side="left", padx=5)
        # Bouton d'exportation
        export_btn = ttk.Button(nav_frame, text="Exporter en Excel", command=self.export_to_excel, style="Green.TButton")
        export_btn.pack(side="left", padx=5)

    def show_adjusted_calendar(self):
        self.clear_frame()
        self.current_page = "adjusted_calendar"
        label = ttk.Label(self, text=f"Calendrier réel ajusté - Semaine {self.current_week + 1}", font=("Arial", 20, "bold"), foreground="black")
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
            ttk.Button(nav_frame, text="Semaine précédente", command=lambda: self.show_adjusted_calendar_week(self.current_week-1), style="Green.TButton").pack(side="left", padx=5)
        if self.current_week < self.total_weeks - 1:
            ttk.Button(nav_frame, text="Semaine suivante", command=lambda: self.show_adjusted_calendar_week(self.current_week+1), style="Green.TButton").pack(side="left", padx=5)
        ttk.Button(nav_frame, text="Retour", command=lambda: self.show_adjusted_schedule(self.current_week), style="Green.TButton").pack(side="left", padx=5)

    def show_adjusted_calendar_week(self, week):
        self.current_week = week
        self.show_adjusted_calendar()

    def export_to_excel(self):
        if self.current_page == "week_schedule" or self.current_page == "adjusted_schedule":
            if hasattr(self, 'schedule_data') and self.schedule_data:
                start_slot = self.current_week * 40
                end_slot = min((self.current_week + 1) * 40, len(self.schedule_data))
                week_data = self.schedule_data[start_slot:end_slot]
                if self.current_week == 0 and not week_data:
                    messagebox.showwarning("Avertissement", "Aucune donnée à exporter pour cette semaine.")
                    return
                headers = ["Temps", "Urgent", "Possible", "Interdit"] + [f"Retard {key}" for key in self.schedule_data[0]['delays'].keys()] + [f"Alpha {key}" for key in self.schedule_data[0]['alpha'].keys()] + ["Planifiée"]
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