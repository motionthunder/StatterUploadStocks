import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import tkinter.filedialog as fd
import json
import csv
from datetime import datetime
import os
import ast

import matplotlib
matplotlib.use('TkAgg')

class ProjectTracker:
    def __init__(self, root):
        self.root = root
        self.root.title("Motion Projects Upload Tracker")

        # Категории платформ
        self.platform_categories = {
            "Motion Array": ["AET", "PPT", "AEFX", "PPFX", "FCPX", "DRT", "DRM"],
            "Adobe Stock": ["AET", "PPT", "MGT"],
            "Envato": ["AET", "PPT", "AEFX", "PPFX", "FCPX", "DRT", "DRM"],
            "Artlist": ["AET", "PPT", "FCPX", "DRT"],
            "Videobolt": ["AET"],
            "MotionElements": ["AET", "PPT", "AEFX", "PPFX", "FCPX", "DRT", "DRM"],
            "CCT": ["AET", "PPT", "AEFX", "PPFX", "FCPX", "DRT", "DRM"],
            "Pond5": ["AET"],
            "FilterGrade": ["AET", "PPT", "AEFX", "PPFX", "FCPX", "DRT", "DRM"]
        }

        # Описания категорий
        self.category_full_names = {
            "AET": "After Effects Template",
            "PPT": "Premiere Pro Template",
            "AEFX": "After Effects FX",
            "PPFX": "Premiere Pro FX",
            "MGT": "Motion Graphics",
            "DRT": "DaVinci Template",
            "DRM": "DaVinci Macros",
            "FCPX": "Final Cut Pro X"
        }

        # Цвета статусов
        self.status_colors = {
            "Uploaded": "#90EE90",
            "Not Uploaded": "#FFB6C1",
            "Pending": "#FFFFE0",
            "Rejected": "#FA8072",
            "Disabled": "#D3D3D3"
        }

        # Цвета платформ
        self.platform_colors = {
            "Motion Array": "#F0F4FF",
            "Adobe Stock": "#FFF0F4",
            "Envato": "#F0FFF4",
            "Artlist": "#FFF4F0",
            "Videobolt": "#FFF8F0",
            "MotionElements": "#F0FFFF",
            "CCT": "#F4FFF0",
            "Pond5": "#FFF0F8",
            "FilterGrade": "#F0FFF8"
        }

        # Инициализация хранения данных
        self.projects = {}
        self.load_platform_data()
        self.load_data()

        # Инициализация состояний сворачивания/разворачивания
        self.platform_states = {}
        self.year_states = {}
        self.month_states = {}
        self.platform_header_frames = {}
        self.month_frames = {}  # Добавлено для хранения фреймов месяцев

        # Создание основных панелей
        self.main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True)

        # Левая панель
        self.left_panel = ttk.Frame(self.main_paned)
        self.main_paned.add(self.left_panel, weight=1)

        # Центральная панель
        self.center_panel = ttk.Frame(self.main_paned)
        self.main_paned.add(self.center_panel, weight=3)

        # Правая панель
        self.right_panel = ttk.Frame(self.main_paned)
        self.main_paned.add(self.right_panel, weight=1)

        # Инициализация фильтров
        self.filter_vars = {
            'search': tk.StringVar(),
        }

        # Создание компонентов интерфейса
        self.create_input_panel()
        self.create_statistics_panel()
        self.create_matrix_view()
        self.create_legend()

        # Свойства окна
        self.root.minsize(1000, 700)

        # Обработчик события изменения размера окна
        self.root.bind("<Configure>", self.on_window_resize)

        # Обновление матрицы
        self.update_matrix()

        # Загрузка настроек
        self.root.after(100, self.load_settings)

        # Привязка события закрытия окна
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_tooltip(self, widget, text):
        def enter(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")

            label = ttk.Label(tooltip, text=text, background="#ffffe0", relief='solid', borderwidth=1)
            label.pack()

            widget.tooltip = tooltip

        def leave(event):
            if hasattr(widget, "tooltip"):
                widget.tooltip.destroy()
                delattr(widget, "tooltip")

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def create_input_panel(self):
        input_frame = ttk.LabelFrame(self.left_panel, text="Добавить/Редактировать проект", padding="5")
        input_frame.pack(fill=tk.X, padx=5, pady=5)

        # Поле ввода названия проекта
        ttk.Label(input_frame, text="Название проекта:").pack(fill=tk.X)
        self.project_name = ttk.Entry(input_frame)
        self.project_name.pack(fill=tk.X, pady=5)

        # Выбор года и месяца
        # Год
        ttk.Label(input_frame, text="Год:").pack(fill=tk.X)
        current_year = datetime.now().year
        self.year_var = tk.IntVar(value=current_year)
        self.year_spinbox = ttk.Spinbox(input_frame, from_=2000, to=2100, textvariable=self.year_var, width=5)
        self.year_spinbox.pack(fill=tk.X, padx=5, pady=2)

        # Месяц
        ttk.Label(input_frame, text="Месяц:").pack(fill=tk.X)
        self.month_var = tk.StringVar(value=datetime.now().strftime("%B"))
        self.month_cb = ttk.Combobox(input_frame, textvariable=self.month_var,
                                     values=[datetime(2000, m, 1).strftime('%B') for m in range(1, 13)], width=10)
        self.month_cb.pack(fill=tk.X, padx=5, pady=2)

        # Флажки платформ
        self.platform_checkboxes_frame = ttk.Frame(input_frame)
        self.platform_checkboxes_frame.pack(fill=tk.X, expand=True)

        self.platform_vars = {}
        for platform in self.platform_categories.keys():
            self.platform_vars[platform] = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.platform_checkboxes_frame, text=platform, variable=self.platform_vars[platform])
            cb.pack(anchor="w")

        # Кнопки выбора платформ
        helpers_frame = ttk.Frame(input_frame)
        helpers_frame.pack(fill=tk.X, pady=2)

        ttk.Button(helpers_frame, text="Выбрать все",
                   command=lambda: self.set_all_platforms(True)).pack(side=tk.LEFT, padx=2)
        ttk.Button(helpers_frame, text="Очистить все",
                   command=lambda: self.set_all_platforms(False)).pack(side=tk.LEFT, padx=2)

        # Кнопки добавления, редактирования и удаления проекта
        ttk.Button(input_frame, text="Добавить проект", command=self.add_project).pack(fill=tk.X, pady=5)
        ttk.Button(input_frame, text="Редактировать проект", command=self.edit_project).pack(fill=tk.X, pady=5)
        ttk.Button(input_frame, text="Удалить проект", command=self.delete_project).pack(fill=tk.X, pady=5)

        # Кнопка добавления платформы
        ttk.Button(input_frame, text="Добавить платформу", command=self.add_platform).pack(fill=tk.X, pady=5)

    def set_all_platforms(self, value):
        for var in self.platform_vars.values():
            var.set(value)

    def create_matrix_view(self):
        self.matrix_container = ttk.Frame(self.center_panel)
        self.matrix_container.pack(fill=tk.BOTH, expand=True)
        self.matrix_container.grid_columnconfigure(0, weight=1)
        self.matrix_container.grid_rowconfigure(0, weight=1)

        # Создание прокручиваемого холста
        self.canvas = tk.Canvas(self.matrix_container)
        v_scrollbar = ttk.Scrollbar(self.matrix_container, orient="vertical", command=self.canvas.yview)
        h_scrollbar = ttk.Scrollbar(self.matrix_container, orient="horizontal", command=self.canvas.xview)

        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind("<Configure>",
                                   lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # Настройка полос прокрутки
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Привязка событий прокрутки
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_shift_mousewheel(self, event):
        self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def create_statistics_panel(self):
        stats_frame = ttk.LabelFrame(self.right_panel, text="Статистика", padding="5")
        stats_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.stats_labels = {}
        for status in self.status_colors.keys():
            frame = ttk.Frame(stats_frame)
            frame.pack(fill=tk.X, padx=5, pady=2)
            color_box = tk.Frame(frame, bg=self.status_colors[status], width=15, height=15)
            color_box.pack(side="left", padx=2)
            label = ttk.Label(frame, text=f"{status}: 0")
            label.pack(side="left")
            self.stats_labels[status] = label

        # Добавление фильтров
        filters_frame = ttk.Frame(stats_frame)
        filters_frame.pack(fill=tk.X, padx=5, pady=5)

        # Фильтр по статусу
        ttk.Label(filters_frame, text="Фильтр по статусу:").pack(fill=tk.X, padx=5, pady=2)
        self.status_filter_var = tk.StringVar(value="All")
        status_values = ["All"] + list(self.status_colors.keys())
        status_filter_cb = ttk.Combobox(filters_frame, textvariable=self.status_filter_var, values=status_values,
                                        state="readonly")
        status_filter_cb.pack(fill=tk.X, padx=5, pady=2)

        # Фильтр по платформе
        ttk.Label(filters_frame, text="Фильтр по платформе:").pack(fill=tk.X, padx=5, pady=2)
        self.platform_filter_var = tk.StringVar(value="All")
        platform_values = ["All"] + list(self.platform_categories.keys())
        platform_filter_cb = ttk.Combobox(filters_frame, textvariable=self.platform_filter_var, values=platform_values,
                                          state="readonly")
        platform_filter_cb.pack(fill=tk.X, padx=5, pady=2)

        # Кнопка применения фильтра
        ttk.Button(filters_frame, text="Применить фильтр", command=self.update_matrix).pack(fill=tk.X, padx=5, pady=5)

        # Кнопки импорта и экспорта
        buttons_frame = ttk.Frame(self.right_panel)
        buttons_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=5, pady=5)

        ttk.Button(buttons_frame, text="Импорт CSV", command=self.import_from_csv).pack(side=tk.RIGHT, padx=5)
        ttk.Button(buttons_frame, text="Экспорт CSV", command=self.export_to_csv).pack(side=tk.RIGHT)

        self.update_statistics()

    def update_statistics(self):
        status_counts = {status: 0 for status in self.status_colors.keys()}

        for project in self.projects.values():
            for platform in project:
                if platform in ['year', 'month']:
                    continue
                for category in project[platform]:
                    status = project[platform][category]['status']
                    status_counts[status] += 1

        for status, label in self.stats_labels.items():
            count = status_counts.get(status, 0)
            label.config(text=f"{status}: {count}")

    def create_legend(self):
        legend_frame = ttk.LabelFrame(self.left_panel, text="Обозначения", padding="5")
        legend_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=5, pady=5)

        # Легенда цветов статусов
        status_frame = ttk.Frame(legend_frame)
        status_frame.pack(fill=tk.X, pady=2)

        for status, color in self.status_colors.items():
            status_item = ttk.Frame(status_frame)
            status_item.pack(fill=tk.X, pady=1)
            color_box = tk.Frame(status_item, bg=color, width=15, height=15)
            color_box.pack(side="left", padx=2)
            ttk.Label(status_item, text=status, font=("Arial", 8)).pack(side="left")

        # Легенда категорий
        cat_frame = ttk.Frame(legend_frame)
        cat_frame.pack(fill=tk.X, pady=2)
        ttk.Label(cat_frame, text="Категории:").pack(anchor="w")

        left_col = ttk.Frame(cat_frame)
        right_col = ttk.Frame(cat_frame)
        left_col.pack(side=tk.LEFT, fill=tk.X, expand=True)
        right_col.pack(side=tk.LEFT, fill=tk.X, expand=True)

        categories = list(self.category_full_names.items())
        mid_point = len(categories) // 2

        for i, (abbr, full) in enumerate(categories):
            target_col = left_col if i < mid_point else right_col
            cat_item = ttk.Frame(target_col)
            cat_item.pack(fill=tk.X, pady=1)
            ttk.Label(cat_item, text=f"{abbr}: {full}", font=("Arial", 8)).pack(anchor="w")

    def update_matrix(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        grouped_projects = {}
        for project_name, project_data in self.projects.items():
            year = project_data.get('year')
            month = project_data.get('month')
            if year and month:
                grouped_projects.setdefault(year, {}).setdefault(month, {})[project_name] = project_data

        row_counter = [0]

        for year in sorted(grouped_projects.keys(), reverse=True):
            self.create_year_frame(year, grouped_projects[year], row_counter)

        self.update_statistics()

    def create_year_frame(self, year, months_data, row_counter):
        year_header_frame = ttk.Frame(self.scrollable_frame)
        year_header_frame.grid(row=row_counter[0], column=0, sticky="nsew", padx=5, pady=5)
        self.scrollable_frame.grid_rowconfigure(row_counter[0], weight=1)
        year_row = row_counter[0]
        row_counter[0] += 1

        if year not in self.year_states:
            self.year_states[year] = tk.BooleanVar(value=False)
        is_year_collapsed = self.year_states[year]

        year_projects_frame = ttk.Frame(self.scrollable_frame)
        if not is_year_collapsed.get():
            year_projects_frame.grid(row=row_counter[0], column=0, sticky="nsew")
        year_projects_frame.grid_row = row_counter[0]
        row_counter[0] += 1

        year_symbol_label = ttk.Label(year_header_frame, text="▼" if not is_year_collapsed.get() else "▶", width=2)
        year_symbol_label.pack(side=tk.LEFT)
        year_label = ttk.Label(year_header_frame, text=str(year), font=("Arial", 12, "bold"))
        year_label.pack(side=tk.LEFT)

        def toggle_year_frame(event=None):
            if is_year_collapsed.get():
                year_projects_frame.grid(row=year_projects_frame.grid_row, column=0, sticky="nsew")
                year_symbol_label.config(text="▼")
            else:
                year_projects_frame.grid_remove()
                year_symbol_label.config(text="▶")
            is_year_collapsed.set(not is_year_collapsed.get())

        year_header_frame.bind("<Button-1>", toggle_year_frame)
        year_symbol_label.bind("<Button-1>", toggle_year_frame)
        year_label.bind("<Button-1>", toggle_year_frame)

        if is_year_collapsed.get():
            return

        # Инициализация счетчика строк для месяцев внутри года
        month_row_counter = [0]

        for month in sorted(months_data.keys(), key=lambda m: datetime.strptime(m, '%B').month,
                            reverse=True):
            self.create_month_frame(year, month, months_data[month], year_projects_frame, month_row_counter)

    def create_month_frame(self, year, month, projects_data, parent_frame, row_counter):
        month_header_frame = ttk.Frame(parent_frame)
        month_header_frame.grid(row=row_counter[0], column=0, sticky="nsew", padx=20, pady=2)
        parent_frame.grid_rowconfigure(row_counter[0], weight=1)
        month_row = row_counter[0]
        row_counter[0] += 1

        if (year, month) not in self.month_states:
            self.month_states[(year, month)] = tk.BooleanVar(value=False)
        is_month_collapsed = self.month_states[(year, month)]

        month_projects_frame = ttk.Frame(parent_frame)
        if not is_month_collapsed.get():
            month_projects_frame.grid(row=row_counter[0], column=0, sticky="nsew")
        month_projects_frame.grid_row = row_counter[0]
        row_counter[0] += 1

        month_symbol_label = ttk.Label(month_header_frame,
                                       text="▼" if not is_month_collapsed.get() else "▶", width=2)
        month_symbol_label.pack(side=tk.LEFT)
        month_label = ttk.Label(month_header_frame, text=month, font=("Arial", 10, "bold"))
        month_label.pack(side=tk.LEFT)

        def toggle_month_frame(event=None):
            if is_month_collapsed.get():
                month_projects_frame.grid(row=month_projects_frame.grid_row, column=0, sticky="nsew")
                month_symbol_label.config(text="▼")
            else:
                month_projects_frame.grid_remove()
                month_symbol_label.config(text="▶")
            is_month_collapsed.set(not is_month_collapsed.get())

        month_header_frame.bind("<Button-1>", toggle_month_frame)
        month_symbol_label.bind("<Button-1>", toggle_month_frame)
        month_label.bind("<Button-1>", toggle_month_frame)

        if is_month_collapsed.get():
            return

        # Сохранение ссылки на фрейм месяца
        self.month_frames[(year, month)] = month_projects_frame
        month_projects_frame.projects_data = projects_data

        self.create_matrix_headers(month_projects_frame, year, month)
        self.populate_projects(month_projects_frame, year, month, projects_data)

    def populate_projects(self, month_projects_frame, year, month, projects_data):
        # Удаляем существующие строки проектов
        for widget in month_projects_frame.winfo_children():
            if widget.grid_info().get('row', 0) >= 2:
                widget.destroy()

        project_row = 2
        for project in sorted(projects_data.keys()):
            if self.should_show_project(project):
                label = ttk.Label(month_projects_frame, text=project)
                label.grid(row=project_row, column=0, sticky="nsew")
                self.create_tooltip(label, project)
                month_projects_frame.grid_rowconfigure(project_row, weight=1)

                col = 1
                project_data = self.projects[project]
                for platform in self.platform_categories.keys():
                    key = (year, month, platform)
                    is_expanded = self.platform_states.get(key, tk.BooleanVar(value=True))
                    self.platform_states[key] = is_expanded
                    if platform in project_data:
                        platform_filter = self.platform_filter_var.get()
                        if platform_filter != "All" and platform != platform_filter:
                            if is_expanded.get():
                                col += len(self.platform_categories[platform])
                            else:
                                col += 1
                            continue
                        if is_expanded.get():
                            for category in self.platform_categories[platform]:
                                try:
                                    status = self.get_status(project, platform, category)
                                    if self.status_filter_var.get() != "All" and status != self.status_filter_var.get():
                                        cell = tk.Frame(month_projects_frame, width=20, height=20)
                                        cell.grid(row=project_row, column=col, sticky="nsew", padx=1, pady=1)
                                        col += 1
                                        continue
                                    cell = tk.Frame(
                                        month_projects_frame,
                                        bg=self.status_colors[status],
                                        width=20,
                                        height=20
                                    )
                                    cell.grid(row=project_row, column=col, sticky="nsew", padx=1, pady=1)
                                    cell.bind("<Button-1>",
                                              lambda e, p=project, plf=platform, cat=category, b=cell:
                                              self.cycle_status(p, plf, cat, b))
                                    cell.bind("<Button-3>",
                                              lambda e, p=project, plf=platform, cat=category, b=cell:
                                              self.show_context_menu(e, p, plf, cat, b))
                                except KeyError:
                                    cell = tk.Frame(
                                        month_projects_frame,
                                        bg=self.status_colors["Not Uploaded"],
                                        width=20,
                                        height=20
                                    )
                                    cell.grid(row=project_row, column=col, sticky="nsew", padx=1, pady=1)
                                col += 1
                        else:
                            try:
                                statuses = [self.get_status(project, platform, category)
                                            for category in self.platform_categories[platform]]
                                if all(status == "Uploaded" for status in statuses):
                                    platform_status = "Uploaded"
                                elif any(status == "Rejected" for status in statuses):
                                    platform_status = "Rejected"
                                elif any(status == "Pending" for status in statuses):
                                    platform_status = "Pending"
                                elif all(status == "Disabled" for status in statuses):
                                    platform_status = "Disabled"
                                else:
                                    platform_status = "Not Uploaded"
                            except KeyError:
                                platform_status = "Not Uploaded"

                            if self.status_filter_var.get() != "All" and platform_status != self.status_filter_var.get():
                                cell = tk.Frame(month_projects_frame, width=20, height=20)
                                cell.grid(row=project_row, column=col, sticky="nsew", padx=1, pady=1)
                                col += 1
                                continue

                            cell = tk.Frame(
                                month_projects_frame,
                                bg=self.status_colors[platform_status],
                                width=20,
                                height=20
                            )
                            cell.grid(row=project_row, column=col, sticky="nsew", padx=1, pady=1)
                            cell.bind("<Button-1>",
                                      lambda e, p=project, plf=platform, b=cell:
                                      self.cycle_platform_status(p, plf, b))
                            col += 1
                    else:
                        if is_expanded.get():
                            col += len(self.platform_categories[platform])
                        else:
                            col += 1

                project_row += 1

    def create_matrix_headers(self, parent_frame, year, month):
        # Очистка существующих заголовков
        for widget in parent_frame.winfo_children():
            grid_info = widget.grid_info()
            if grid_info.get('row', 0) in [0, 1]:
                widget.destroy()

        project_label = ttk.Label(parent_frame, text="Проект")
        project_label.grid(row=0, column=0, rowspan=2, sticky="nsew")
        parent_frame.grid_columnconfigure(0, weight=1)

        col = 1

        for platform in self.platform_categories.keys():
            platform_categories = self.platform_categories[platform]
            key = (year, month, platform)
            is_expanded = self.platform_states.get(key, tk.BooleanVar(value=True))
            self.platform_states[key] = is_expanded
            platform_span = len(platform_categories) if is_expanded.get() else 1

            for i in range(platform_span):
                parent_frame.grid_columnconfigure(col + i, weight=1)

            platform_frame = tk.Frame(parent_frame, bg=self.platform_colors[platform])
            platform_frame.grid(row=0, column=col, columnspan=platform_span, sticky="nsew", padx=1)

            symbol = "▼" if is_expanded.get() else "▶"
            platform_symbol_label = tk.Label(platform_frame, text=symbol, bg=self.platform_colors[platform], width=2)
            platform_symbol_label.pack(side=tk.LEFT)

            platform_label = tk.Label(
                platform_frame,
                text=platform,
                bg=self.platform_colors[platform],
                font=("Arial", 8, "bold")
            )
            platform_label.pack(side=tk.LEFT, expand=True, fill="both", padx=1)

            tooltip_text = f"{platform}\nКатегории:\n" + \
                           "\n".join([f"• {self.category_full_names[cat]}" for cat in platform_categories])
            self.create_tooltip(platform_label, tooltip_text)

            def toggle_platform(event=None, p=platform, s=platform_symbol_label, y=year, m=month):
                key = (y, m, p)
                is_expanded = self.platform_states.get(key, tk.BooleanVar(value=True))
                self.platform_states[key] = is_expanded
                is_expanded.set(not is_expanded.get())
                s.config(text="▼" if is_expanded.get() else "▶")
                # Обновляем заголовки и проекты
                month_frame = self.month_frames[(y, m)]
                self.create_matrix_headers(month_frame, y, m)
                self.populate_projects(month_frame, y, m, month_frame.projects_data)

            platform_frame.bind("<Button-1>", toggle_platform)
            platform_symbol_label.bind("<Button-1>", toggle_platform)
            platform_label.bind("<Button-1>", toggle_platform)

            if is_expanded.get():
                for i, category in enumerate(platform_categories):
                    header_frame = tk.Frame(parent_frame)
                    header_frame.grid(row=1, column=col + i, sticky="nsew", padx=1)

                    cat_label = tk.Label(
                        header_frame,
                        text=category,
                        bg=self.platform_colors[platform],
                        font=("Arial", 6)
                    )
                    cat_label.pack(expand=True, fill="both")
                    self.create_tooltip(cat_label, self.category_full_names[category])

                col += len(platform_categories)
            else:
                header_frame = tk.Frame(parent_frame)
                header_frame.grid(row=1, column=col, sticky="nsew", padx=1)

                cat_label = tk.Label(
                    header_frame,
                    text='',
                    bg=self.platform_colors[platform],
                    font=("Arial", 6)
                )
                cat_label.pack(expand=True, fill="both")
                col += 1

    def cycle_platform_status(self, project, platform, button):
        statuses = list(self.status_colors.keys())
        current_statuses = [self.get_status(project, platform, category)
                            for category in self.platform_categories[platform]]
        current_status = current_statuses[0]
        next_status = statuses[(statuses.index(current_status) + 1) % len(statuses)]

        for category in self.platform_categories[platform]:
            self.projects[project][platform][category] = {
                "status": next_status,
                "date": datetime.now().strftime("%Y-%m-%d")
            }

        button.configure(bg=self.status_colors[next_status])
        self.save_data()
        self.update_statistics()

    def add_project(self):
        name = self.project_name.get().strip()
        if not name:
            return

        selected_platforms = [p for p in self.platform_categories.keys()
                              if self.platform_vars[p].get()]

        if not selected_platforms:
            return

        year = self.year_var.get()
        month = self.month_var.get()

        self.projects[name] = {
            'year': year,
            'month': month
        }
        for platform in selected_platforms:
            self.projects[name][platform] = {
                category: {
                    "status": "Not Uploaded",
                    "date": datetime.now().strftime("%Y-%m-%d")
                }
                for category in self.platform_categories[platform]
            }

        self.save_data()
        self.update_matrix()
        self.update_statistics()
        self.project_name.delete(0, tk.END)

    def edit_project(self):
        name = self.project_name.get().strip()
        if name not in self.projects:
            messagebox.showwarning("Проект не найден", f"Проект '{name}' не существует.")
            return

        project_data = self.projects[name]

        edit_window = tk.Toplevel(self.root)
        edit_window.title(f"Редактировать проект - {name}")
        edit_window.grab_set()

        ttk.Label(edit_window, text="Название проекта:").pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(edit_window, text=name).pack(fill=tk.X, padx=5, pady=5)

        # Год
        ttk.Label(edit_window, text="Год:").pack(fill=tk.X)
        year_var = tk.IntVar(value=project_data.get('year', datetime.now().year))
        year_spinbox = ttk.Spinbox(edit_window, from_=2000, to=2100, textvariable=year_var, width=5)
        year_spinbox.pack(fill=tk.X, padx=5, pady=2)

        # Месяц
        ttk.Label(edit_window, text="Месяц:").pack(fill=tk.X)
        month_var = tk.StringVar(value=project_data.get('month', datetime.now().strftime("%B")))
        month_cb = ttk.Combobox(edit_window, textvariable=month_var,
                                values=[datetime(2000, m, 1).strftime('%B') for m in range(1, 13)], width=10)
        month_cb.pack(fill=tk.X, padx=5, pady=2)

        platforms_frame = ttk.LabelFrame(edit_window, text="Платформы")
        platforms_frame.pack(fill=tk.X, padx=5, pady=5)

        platform_vars = {}
        for platform in self.platform_categories.keys():
            var = tk.BooleanVar(value=platform in project_data)
            platform_vars[platform] = var
            cb = ttk.Checkbutton(platforms_frame, text=platform, variable=var)
            cb.pack(anchor="w")

        buttons_frame = ttk.Frame(edit_window)
        buttons_frame.pack(fill=tk.X, padx=5, pady=5)

        def save_changes():
            self.projects[name]['year'] = year_var.get()
            self.projects[name]['month'] = month_var.get()

            for platform, var in platform_vars.items():
                if var.get():
                    if platform not in self.projects[name]:
                        self.projects[name][platform] = {
                            category: {
                                "status": "Not Uploaded",
                                "date": datetime.now().strftime("%Y-%m-%d")
                            }
                            for category in self.platform_categories[platform]
                        }
                else:
                    if platform in self.projects[name]:
                        del self.projects[name][platform]

            self.save_data()
            self.update_matrix()
            self.update_statistics()
            edit_window.destroy()

        ttk.Button(buttons_frame, text="Сохранить", command=save_changes).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Отмена", command=edit_window.destroy).pack(side=tk.LEFT, padx=5)

    def delete_project(self):
        name = self.project_name.get().strip()
        if name in self.projects:
            confirm = messagebox.askyesno("Подтвердите удаление", f"Вы уверены, что хотите удалить проект '{name}'?")
            if confirm:
                del self.projects[name]
                self.save_data()
                self.update_matrix()
                self.update_statistics()
                self.project_name.delete(0, tk.END)
        else:
            messagebox.showwarning("Проект не найден", f"Проект '{name}' не существует.")

    def should_show_project(self, project):
        if self.filter_vars['search'].get().lower() not in project.lower():
            return False

        platform_filter = self.platform_filter_var.get()
        if platform_filter != "All":
            if platform_filter not in self.projects[project]:
                return False

        status_filter = self.status_filter_var.get()
        if status_filter != "All":
            project_has_status = False
            for platform in self.projects[project]:
                if platform in ['year', 'month']:
                    continue
                if platform_filter != "All" and platform != platform_filter:
                    continue
                for category in self.projects[project][platform]:
                    status = self.get_status(project, platform, category)
                    if status == status_filter:
                        project_has_status = True
                        break
                if project_has_status:
                    break
            if not project_has_status:
                return False

        return True

    def get_status(self, project, platform, category):
        try:
            return self.projects[project][platform][category]["status"]
        except KeyError:
            return "Not Uploaded"

    def cycle_status(self, project, platform, category, button):
        statuses = list(self.status_colors.keys())
        current = self.get_status(project, platform, category)
        current_index = statuses.index(current)
        next_status = statuses[(current_index + 1) % len(statuses)]
        self.projects[project][platform][category] = {
            "status": next_status,
            "date": datetime.now().strftime("%Y-%m-%d")
        }
        button.configure(bg=self.status_colors[next_status])
        self.save_data()
        self.update_statistics()

    def show_context_menu(self, event, project, platform, category, cell):
        current_status = self.get_status(project, platform, category)
        menu = tk.Menu(self.root, tearoff=0)
        if current_status != "Disabled":
            menu.add_command(label="Disable",
                             command=lambda: self.set_status(project, platform, category, "Disabled", cell))
        else:
            menu.add_command(label="Enable",
                             command=lambda: self.set_status(project, platform, category, "Not Uploaded", cell))
        menu.tk_popup(event.x_root, event.y_root)

    def set_status(self, project, platform, category, status, cell):
        self.projects[project][platform][category] = {
            "status": status,
            "date": datetime.now().strftime("%Y-%m-%d")
        }
        cell.configure(bg=self.status_colors[status])
        self.save_data()
        self.update_statistics()

    def filter_projects(self, *args):
        pass

    def export_to_csv(self):
        file_path = fd.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return

        with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Project', 'Year', 'Month', 'Platform', 'Category', 'Status', 'Date']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

            writer.writeheader()
            for project_name, project_data in self.projects.items():
                year = project_data.get('year')
                month = project_data.get('month')
                for platform, categories in project_data.items():
                    if platform in ['year', 'month']:
                        continue
                    for category, details in categories.items():
                        writer.writerow({
                            'Project': project_name,
                            'Year': year,
                            'Month': month,
                            'Platform': platform,
                            'Category': category,
                            'Status': details.get('status'),
                            'Date': details.get('date')
                        })

    def import_from_csv(self):
        file_path = fd.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return

        with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                name = row['Project']
                year = int(row['Year'])
                month = row['Month']
                platform = row['Platform']
                category = row['Category']
                status = row['Status']
                date = row['Date']

                if name not in self.projects:
                    self.projects[name] = {'year': year, 'month': month}

                if platform not in self.projects[name]:
                    self.projects[name][platform] = {}

                self.projects[name][platform][category] = {'status': status, 'date': date}

        self.save_data()
        self.update_matrix()
        self.update_statistics()

    def add_platform(self):
        platform_window = tk.Toplevel(self.root)
        platform_window.title("Добавить платформу")
        platform_window.grab_set()

        ttk.Label(platform_window, text="Название платформы:").pack(fill=tk.X, padx=5, pady=5)
        platform_name_var = tk.StringVar()
        platform_name_entry = ttk.Entry(platform_window, textvariable=platform_name_var)
        platform_name_entry.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(platform_window, text="Выберите категории:").pack(fill=tk.X, padx=5, pady=5)
        categories_frame = ttk.Frame(platform_window)
        categories_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        category_vars = {}
        for category in self.category_full_names.keys():
            var = tk.BooleanVar(value=False)
            category_vars[category] = var
            cb = ttk.Checkbutton(categories_frame, text=f"{category} - {self.category_full_names[category]}",
                                 variable=var)
            cb.pack(anchor="w")

        ttk.Label(platform_window, text="Цвет платформы (hex):").pack(fill=tk.X, padx=5, pady=5)
        platform_color_var = tk.StringVar(value="#FFFFFF")
        platform_color_entry = ttk.Entry(platform_window, textvariable=platform_color_var)
        platform_color_entry.pack(fill=tk.X, padx=5, pady=5)

        buttons_frame = ttk.Frame(platform_window)
        buttons_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(buttons_frame, text="Добавить",
                   command=lambda: self.save_new_platform(platform_name_var.get(), category_vars, platform_color_var.get(),
                                                          platform_window)).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Отмена", command=platform_window.destroy).pack(side=tk.LEFT, padx=5)

    def save_new_platform(self, platform_name, category_vars, platform_color, window):
        platform_name = platform_name.strip()
        if not platform_name:
            messagebox.showwarning("Ошибка ввода", "Название платформы не может быть пустым.")
            return
        if platform_name in self.platform_categories:
            messagebox.showwarning("Дубликат платформы", "Платформа уже существует.")
            return
        selected_categories = [cat for cat, var in category_vars.items() if var.get()]
        if not selected_categories:
            messagebox.showwarning("Ошибка ввода", "Необходимо выбрать хотя бы одну категорию.")
            return
        try:
            self.root.winfo_rgb(platform_color)
        except:
            messagebox.showwarning("Ошибка ввода", "Неверный код цвета.")
            return
        self.platform_categories[platform_name] = selected_categories
        self.platform_colors[platform_name] = platform_color

        self.update_platform_checkboxes()
        self.update_platform_filter()

        self.save_platform_data()

        window.destroy()
        self.update_matrix()
        self.update_statistics()

    def update_platform_checkboxes(self):
        for widget in self.platform_checkboxes_frame.winfo_children():
            widget.destroy()
        self.platform_vars = {}
        for platform in self.platform_categories.keys():
            self.platform_vars[platform] = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.platform_checkboxes_frame, text=platform, variable=self.platform_vars[platform])
            cb.pack(anchor="w")

    def update_platform_filter(self):
        platform_values = ["All"] + list(self.platform_categories.keys())
        self.platform_filter_var.set("All")

    def load_data(self):
        try:
            if os.path.exists("projects_data.json"):
                with open("projects_data.json", "r", encoding='utf-8') as file:
                    self.projects = json.load(file)
        except Exception as e:
            print(f"Ошибка загрузки данных: {e}")
            self.projects = {}

    def save_data(self):
        try:
            with open("projects_data.json", "w", encoding='utf-8') as file:
                json.dump(self.projects, file, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Ошибка сохранения данных: {e}")

    def save_platform_data(self):
        platform_data = {
            'platform_categories': self.platform_categories,
            'platform_colors': self.platform_colors
        }
        try:
            with open("platforms_data.json", "w", encoding='utf-8') as file:
                json.dump(platform_data, file, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Ошибка сохранения данных платформ: {e}")

    def load_platform_data(self):
        try:
            if os.path.exists("platforms_data.json"):
                with open("platforms_data.json", "r", encoding='utf-8') as file:
                    platform_data = json.load(file)
                    self.platform_categories = platform_data['platform_categories']
                    self.platform_colors = platform_data['platform_colors']
        except Exception as e:
            print(f"Ошибка загрузки данных платформ: {e}")

    def load_settings(self):
        try:
            if os.path.exists("settings.json"):
                with open("settings.json", "r", encoding='utf-8') as file:
                    settings = json.load(file)
                    # Восстановление фильтров
                    self.status_filter_var.set(settings.get('status_filter', 'All'))
                    self.platform_filter_var.set(settings.get('platform_filter', 'All'))
                    # Восстановление состояний сворачивания/разворачивания
                    platform_states = settings.get('platform_states', {})
                    for key_str, value in platform_states.items():
                        key_tuple = ast.literal_eval(key_str)
                        self.platform_states[key_tuple] = tk.BooleanVar(value=value)
                    year_states = settings.get('year_states', {})
                    for key_str, value in year_states.items():
                        self.year_states[int(key_str)] = tk.BooleanVar(value=value)
                    month_states = settings.get('month_states', {})
                    for key_str, value in month_states.items():
                        key_tuple = ast.literal_eval(key_str)
                        self.month_states[key_tuple] = tk.BooleanVar(value=value)
                    # Установка размеров окна
                    self.root.geometry(settings.get('window_geometry', '1400x800'))
                    # Установка позиций разделителей
                    sash_positions = settings.get('paned_window_sash_positions', [])
                    if sash_positions:
                        self.root.update_idletasks()
                        for i, pos in enumerate(sash_positions):
                            self.main_paned.sashpos(i, pos)
        except Exception as e:
            print(f"Ошибка загрузки настроек: {e}")

    def save_settings(self):
        try:
            settings = {}
            # Сохранение размера и позиции окна
            settings['window_geometry'] = self.root.winfo_geometry()
            # Сохранение позиций разделителей
            sash_positions = [self.main_paned.sashpos(i) for i in range(len(self.main_paned.panes()) - 1)]
            settings['paned_window_sash_positions'] = sash_positions
            # Сохранение фильтров
            settings['status_filter'] = self.status_filter_var.get()
            settings['platform_filter'] = self.platform_filter_var.get()
            # Сохранение состояний сворачивания/разворачивания
            settings['platform_states'] = {repr(k): v.get() for k, v in self.platform_states.items()}
            settings['year_states'] = {str(k): v.get() for k, v in self.year_states.items()}
            settings['month_states'] = {repr(k): v.get() for k, v in self.month_states.items()}
            with open("settings.json", "w", encoding='utf-8') as file:
                json.dump(settings, file, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Ошибка сохранения настроек: {e}")

    def on_closing(self):
        self.save_settings()
        self.root.destroy()

    def on_window_resize(self, event):
        # Обновление размеров при изменении размера окна
        pass

if __name__ == "__main__":
    root = tk.Tk()
    app = ProjectTracker(root)
    root.mainloop()
