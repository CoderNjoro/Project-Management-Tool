import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import datetime, timedelta
import json
import os
import shutil
import platform
from PIL import Image, ImageTk 
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd
import networkx as nx
import zipfile
import webbrowser
from pathlib import Path

class Task:
    def __init__(self, name, start_date, duration, assigned_to="", priority="Medium", status="Not Started", 
                 dependencies=None, cost=0.0, is_milestone=False):
        self.name = name
        self.start_date = start_date
        self.duration = duration if not is_milestone else 0
        self.assigned_to = assigned_to
        self.priority = priority
        self.status = status
        self.dependencies = dependencies or []
        self.cost = cost
        self.is_milestone = is_milestone
        self.id = None
    
    def end_date(self):
        return self.start_date + timedelta(days=self.duration)
    
    def to_dict(self):
        return {
            'name': self.name,
            'start_date': self.start_date.isoformat(),
            'duration': self.duration,
            'assigned_to': self.assigned_to,
            'priority': self.priority,
            'status': self.status,
            'dependencies': self.dependencies,
            'cost': self.cost,
            'is_milestone': self.is_milestone
        }
    
    @classmethod
    def from_dict(cls, data):
        task = cls(
            data['name'],
            datetime.fromisoformat(data['start_date']),
            data['duration'],
            data.get('assigned_to', ''),
            data.get('priority', 'Medium'),
            data.get('status', 'Not Started'),
            data.get('dependencies', []),
            data.get('cost', 0.0),
            data.get('is_milestone', False)
        )
        return task

class Document:
    def __init__(self, name, file_path, doc_type="General", category="", tags="", description=""):
        self.name = name
        self.file_path = file_path
        self.doc_type = doc_type
        self.category = category
        self.tags = tags
        self.description = description
        self.upload_date = datetime.now()
        self.file_size = self.get_file_size()
    
    def get_file_size(self):
        try:
            size_bytes = os.path.getsize(self.file_path)
            for unit in ['B', 'KB', 'MB', 'GB']:
                if size_bytes < 1024.0:
                    return f"{size_bytes:.2f} {unit}"
                size_bytes /= 1024.0
            return f"{size_bytes:.2f} TB"
        except:
            return "Unknown"
    
    def to_dict(self):
        return {
            'name': self.name,
            'file_path': self.file_path,
            'doc_type': self.doc_type,
            'category': self.category,
            'tags': self.tags,
            'description': self.description,
            'upload_date': self.upload_date.isoformat(),
            'file_size': self.file_size
        }
    
    @classmethod
    def from_dict(cls, data):
        doc = cls(
            data['name'],
            data['file_path'],
            data.get('doc_type', 'General'),
            data.get('category', ''),
            data.get('tags', ''),
            data.get('description', '')
        )
        doc.upload_date = datetime.fromisoformat(data['upload_date'])
        doc.file_size = data.get('file_size', 'Unknown')
        return doc

class EnhancedProjectManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ProManage - Professional Project Management System")
        self.root.geometry("1600x950")
        
        # Modern color scheme
        self.colors = {
            'primary': '#6366f1',      # Indigo
            'secondary': '#8b5cf6',    # Purple
            'success': '#10b981',      # Green
            'warning': '#f59e0b',      # Amber
            'danger': '#ef4444',       # Red
            'info': '#3b82f6',         # Blue
            'bg_primary': '#f8fafc',   # Light gray
            'bg_secondary': '#ffffff', # White
            'bg_tertiary': '#f1f5f9',  # Lighter gray
            'text_primary': '#1e293b', # Dark slate
            'text_secondary': '#64748b', # Gray
            'border': '#e2e8f0',       # Light border
            'hover': '#f1f5f9',        # Hover state
            'critical': '#dc2626',     # Dark red
            'card': '#ffffff'
        }
        
        self.root.configure(bg=self.colors['bg_primary'])
        
        self.base_storage_path = os.path.join(os.path.expanduser("~"), "Documents", "Project_Management_Data")
        self.init_local_storage()
        
        self.tasks = []
        self.resources = []
        self.baselines = []
        self.risks = []
        self.research_logs = []
        self.documents = []
        self.meeting_notes = []
        self.images = []
        
        self.project_meta = {
            'name': "Untitled Project",
            'problem_statement': "",
            'objectives': "",
            'cost_benefit': "",
            'budget': 0.0,
            'stakeholders': "",
            'success_criteria': ""
        }
        
        self.current_folder = None
        self.project_name = "Untitled Project"
        
        self.setup_modern_styles()
        self.setup_ui()
    
    def setup_modern_styles(self):
        """Configure modern ttk styles"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Modern Treeview
        style.configure('Modern.Treeview',
                       background=self.colors['bg_secondary'],
                       foreground=self.colors['text_primary'],
                       fieldbackground=self.colors['bg_secondary'],
                       borderwidth=0,
                       font=('Segoe UI', 10))
        
        style.configure('Modern.Treeview.Heading',
                       background=self.colors['primary'],
                       foreground='white',
                       borderwidth=0,
                       relief='flat',
                       font=('Segoe UI', 10, 'bold'))
        
        style.map('Modern.Treeview',
                 background=[('selected', self.colors['primary'])],
                 foreground=[('selected', 'white')])
        
        style.map('Modern.Treeview.Heading',
                 background=[('active', self.colors['secondary'])])
        
        # Modern Buttons
        style.configure('Modern.TButton',
                       background=self.colors['primary'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none',
                       font=('Segoe UI', 10, 'bold'),
                       padding=10)
        
        style.map('Modern.TButton',
                 background=[('active', self.colors['secondary']),
                           ('pressed', self.colors['secondary'])])
        
        # Modern Notebook
        style.configure('Modern.TNotebook',
                       background=self.colors['bg_primary'],
                       borderwidth=0)
        
        style.configure('Modern.TNotebook.Tab',
                       background=self.colors['bg_tertiary'],
                       foreground=self.colors['text_primary'],
                       padding=[20, 10],
                       font=('Segoe UI', 10, 'bold'))
        
        style.map('Modern.TNotebook.Tab',
                 background=[('selected', self.colors['primary'])],
                 foreground=[('selected', 'white')],
                 expand=[('selected', [1, 1, 1, 0])])
        
    def init_local_storage(self):
        if not os.path.exists(self.base_storage_path):
            try:
                os.makedirs(self.base_storage_path)
                print(f"Created local storage at: {self.base_storage_path}")
            except OSError as e:
                messagebox.showerror("Storage Error", f"Could not create local storage folder: {e}")

    def setup_ui(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Project (Wizard)", command=self.new_project_wizard)
        file_menu.add_command(label="Open Project", command=self.open_project)
        file_menu.add_command(label="Save Project", command=self.save_project)
        file_menu.add_command(label="Save As...", command=self.save_project_as)
        file_menu.add_command(label="Export to Excel", command=self.export_to_excel)
        file_menu.add_command(label="Backup Project", command=self.backup_project)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Project Statistics", command=self.show_project_statistics)
        tools_menu.add_command(label="Document Search", command=self.show_document_search)
        
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Create Baseline", command=self.create_baseline)
        view_menu.add_command(label="Level Resources", command=self.level_resources)
        
        # Modern gradient toolbar
        toolbar = tk.Frame(self.root, bg=self.colors['primary'], height=80)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        toolbar.pack_propagate(False)
        
        # Logo and title container
        title_container = tk.Frame(toolbar, bg=self.colors['primary'])
        title_container.pack(side=tk.LEFT, padx=30, pady=15)
        
        # Logo emoji
        logo_label = tk.Label(title_container, text="📊", 
                font=('Segoe UI', 32), bg=self.colors['primary'])
        logo_label.pack(side=tk.LEFT, padx=(0, 15))
        
        # Title and project name
        title_frame = tk.Frame(title_container, bg=self.colors['primary'])
        title_frame.pack(side=tk.LEFT)
        
        tk.Label(title_frame, text="ProManage", 
                font=('Segoe UI', 20, 'bold'), bg=self.colors['primary'], 
                fg='white').pack(anchor='w')
        
        self.project_title_label = tk.Label(title_frame, text=f"{self.project_name}", 
                font=('Segoe UI', 11), bg=self.colors['primary'], 
                fg='#e0e7ff')
        self.project_title_label.pack(anchor='w')
        
        # Storage info (right side)
        tk.Label(toolbar, text=f"📁 {self.base_storage_path}", 
                font=('Segoe UI', 9), bg=self.colors['primary'], 
                fg='#c7d2fe').pack(side=tk.RIGHT, padx=30)

        # Main container with modern styling
        main_container = tk.Frame(self.root, bg=self.colors['bg_primary'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        # Left panel (tabs) with shadow effect
        left_panel = tk.Frame(main_container, bg=self.colors['card'], 
                             relief=tk.FLAT, bd=0)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Use modern notebook style
        self.notebook = ttk.Notebook(left_panel, style='Modern.TNotebook')
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        self.dashboard_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.dashboard_frame, text='Dashboard')
        self.setup_dashboard()
        
        self.charter_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.charter_frame, text='Charter & Scope')
        self.setup_charter_tab()
        
        self.tasks_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.tasks_frame, text='Tasks')
        self.setup_tasks_tab()
        
        self.gantt_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.gantt_frame, text='Gantt Chart')
        self.setup_gantt_tab()
        
        self.timeline_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.timeline_frame, text='Timeline')
        self.setup_timeline_tab()
        
        self.kanban_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.kanban_frame, text='Kanban')
        self.setup_kanban_tab()
        
        self.resources_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.resources_frame, text='Resources')
        self.setup_resources_tab()
        
        self.risk_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.risk_frame, text='Risks')
        self.setup_risk_tab()
        
        self.research_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.research_frame, text='Research & Docs')
        self.setup_research_tab()
        
        self.documents_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.documents_frame, text='Documents')
        self.setup_documents_tab()
        
        self.meeting_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.meeting_frame, text='Meeting Notes')
        self.setup_meeting_tab()
        
        self.portfolio_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.portfolio_frame, text='Portfolio')
        self.setup_portfolio_tab()
        
        # Modern right panel
        right_panel = tk.Frame(main_container, bg=self.colors['card'], 
                              relief=tk.FLAT, bd=0, width=320)
        right_panel.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 20), pady=20)
        right_panel.pack_propagate(False)
        
        # Quick Actions Header
        header_frame = tk.Frame(right_panel, bg=self.colors['card'])
        header_frame.pack(fill=tk.X, pady=(20, 15), padx=20)
        
        tk.Label(header_frame, text="⚡ Quick Actions", 
                font=('Segoe UI', 14, 'bold'),
                bg=self.colors['card'], 
                fg=self.colors['text_primary']).pack(anchor='w')
        
        # Modern button style
        btn_config = {
            'font': ('Segoe UI', 10, 'bold'),
            'width': 26,
            'height': 2,
            'relief': tk.FLAT,
            'cursor': 'hand2',
            'bd': 0
        }
        
        # Add Task Button
        add_task_btn = tk.Button(right_panel, text="➕  Add New Task", 
                                bg=self.colors['success'], fg='white',
                                command=self.add_task_dialog, **btn_config)
        add_task_btn.pack(pady=(0, 10), padx=20, fill=tk.X)
        add_task_btn.bind('<Enter>', lambda e: add_task_btn.config(bg='#059669'))
        add_task_btn.bind('<Leave>', lambda e: add_task_btn.config(bg=self.colors['success']))
        
        # Manage Resources Button
        resources_btn = tk.Button(right_panel, text="👥  Manage Resources",
                                 bg=self.colors['primary'], fg='white',
                                 command=self.manage_resources_dialog, **btn_config)
        resources_btn.pack(pady=(0, 10), padx=20, fill=tk.X)
        resources_btn.bind('<Enter>', lambda e: resources_btn.config(bg='#4f46e5'))
        resources_btn.bind('<Leave>', lambda e: resources_btn.config(bg=self.colors['primary']))
        
        # Refresh Gantt Button
        gantt_btn = tk.Button(right_panel, text="📊  Refresh Gantt Chart",
                             bg=self.colors['info'], fg='white',
                             command=self.update_gantt_chart, **btn_config)
        gantt_btn.pack(pady=(0, 10), padx=20, fill=tk.X)
        gantt_btn.bind('<Enter>', lambda e: gantt_btn.config(bg='#2563eb'))
        gantt_btn.bind('<Leave>', lambda e: gantt_btn.config(bg=self.colors['info']))
        
        # Add Document Button
        doc_btn = tk.Button(right_panel, text="📁  Add Document",
                           bg=self.colors['secondary'], fg='white',
                           command=self.add_document_dialog, **btn_config)
        doc_btn.pack(pady=(0, 10), padx=20, fill=tk.X)
        doc_btn.bind('<Enter>', lambda e: doc_btn.config(bg='#7c3aed'))
        doc_btn.bind('<Leave>', lambda e: doc_btn.config(bg=self.colors['secondary']))
        
        # Meeting Notes Button
        meeting_btn = tk.Button(right_panel, text="📝  Add Meeting Note",
                               bg=self.colors['warning'], fg='white',
                               command=self.add_meeting_note_dialog, **btn_config)
        meeting_btn.pack(pady=(0, 10), padx=20, fill=tk.X)
        meeting_btn.bind('<Enter>', lambda e: meeting_btn.config(bg='#d97706'))
        meeting_btn.bind('<Leave>', lambda e: meeting_btn.config(bg=self.colors['warning']))
        
        # Save Button
        save_btn = tk.Button(right_panel, text="💾  Save Project",
                            bg=self.colors['success'], fg='white',
                            command=self.save_project, **btn_config)
        save_btn.pack(pady=(0, 10), padx=20, fill=tk.X)
        save_btn.bind('<Enter>', lambda e: save_btn.config(bg='#059669'))
        save_btn.bind('<Leave>', lambda e: save_btn.config(bg=self.colors['success']))
        
        # Add Risk Button
        risk_btn = tk.Button(right_panel, text="⚠️  Add Risk",
                            bg=self.colors['danger'], fg='white',
                            command=self.add_risk_dialog, **btn_config)
        risk_btn.pack(pady=(0, 10), padx=20, fill=tk.X)
        risk_btn.bind('<Enter>', lambda e: risk_btn.config(bg='#dc2626'))
        risk_btn.bind('<Leave>', lambda e: risk_btn.config(bg=self.colors['danger']))
        
        # Quick Search Button
        search_btn = tk.Button(right_panel, text="🔍  Quick Search",
                              bg=self.colors['info'], fg='white',
                              command=self.show_document_search, **btn_config)
        search_btn.pack(pady=(0, 20), padx=20, fill=tk.X)
        search_btn.bind('<Enter>', lambda e: search_btn.config(bg='#2563eb'))
        search_btn.bind('<Leave>', lambda e: search_btn.config(bg=self.colors['info']))
        
        # Project Stats Card
        stats_frame = tk.LabelFrame(right_panel, text="📊 Project Stats", 
                                   font=('Segoe UI', 11, 'bold'), 
                                   bg=self.colors['card'],
                                   fg=self.colors['text_primary'],
                                   relief=tk.FLAT,
                                   bd=1,
                                   highlightbackground=self.colors['border'],
                                   highlightthickness=1)
        stats_frame.pack(pady=(0, 20), padx=20, fill=tk.BOTH, expand=True)
        
        self.project_info_label = tk.Label(stats_frame, text="No project loaded", 
                                          bg=self.colors['card'], 
                                          fg=self.colors['text_secondary'],
                                          font=('Segoe UI', 10),
                                          justify=tk.LEFT, anchor='nw')
        self.project_info_label.pack(padx=15, pady=15, fill=tk.BOTH, expand=True)

    def setup_dashboard(self):
        # Header section
        header_frame = tk.Frame(self.dashboard_frame, bg='white')
        header_frame.pack(fill=tk.X, padx=30, pady=(20, 10))
        
        tk.Label(header_frame, text='Project Dashboard', 
                font=('Segoe UI', 22, 'bold'),
                bg='white', fg=self.colors['text_primary']).pack(side=tk.LEFT)
        
        # Date/Time
        current_date = datetime.now().strftime('%B %d, %Y')
        tk.Label(header_frame, text=current_date, 
                font=('Segoe UI', 11),
                bg='white', fg=self.colors['text_secondary']).pack(side=tk.RIGHT)
        
        # Stats Cards with modern design
        cards_frame = tk.Frame(self.dashboard_frame, bg='white')
        cards_frame.pack(fill=tk.X, padx=20, pady=20)
        
        # Create professional stat cards
        self.total_tasks_card = self.create_modern_stat_card(
            cards_frame, "Total Tasks", "0", "📋", self.colors['primary'])
        self.total_tasks_card.pack(side=tk.LEFT, padx=10, expand=True, fill=tk.BOTH)
        
        self.completed_card = self.create_modern_stat_card(
            cards_frame, "Completed", "0", "✅", self.colors['success'])
        self.completed_card.pack(side=tk.LEFT, padx=10, expand=True, fill=tk.BOTH)
        
        self.in_progress_card = self.create_modern_stat_card(
            cards_frame, "In Progress", "0", "⏳", self.colors['warning'])
        self.in_progress_card.pack(side=tk.LEFT, padx=10, expand=True, fill=tk.BOTH)
        
        self.not_started_card = self.create_modern_stat_card(
            cards_frame, "Not Started", "0", "⭕", self.colors['danger'])
        self.not_started_card.pack(side=tk.LEFT, padx=10, expand=True, fill=tk.BOTH)
        
        # Second row of stats
        cards_frame2 = tk.Frame(self.dashboard_frame, bg='white')
        cards_frame2.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        self.variance_card = self.create_modern_stat_card(
            cards_frame2, "Schedule Variance", "0 days", "📊", self.colors['info'])
        self.variance_card.pack(side=tk.LEFT, padx=10, expand=True, fill=tk.BOTH)
        
        self.total_cost_card = self.create_modern_stat_card(
            cards_frame2, "Total Budget", "$0", "💰", self.colors['secondary'])
        self.total_cost_card.pack(side=tk.LEFT, padx=10, expand=True, fill=tk.BOTH)
        
        # Progress section
        progress_frame = tk.LabelFrame(self.dashboard_frame, 
                                      text="  📈 Project Progress  ",
                                      font=('Segoe UI', 12, 'bold'),
                                      bg='white', fg=self.colors['text_primary'],
                                      relief=tk.FLAT, bd=2,
                                      highlightbackground=self.colors['border'],
                                      highlightthickness=1)
        progress_frame.pack(fill=tk.X, padx=30, pady=(0, 20))
        
        # Progress bar container
        prog_container = tk.Frame(progress_frame, bg='white')
        prog_container.pack(fill=tk.X, padx=20, pady=15)
        
        tk.Label(prog_container, text='Overall Completion', 
                font=('Segoe UI', 10, 'bold'),
                bg='white', fg=self.colors['text_primary']).pack(anchor='w')
        
        # Progress bar
        self.progress_canvas = tk.Canvas(prog_container, height=30, bg='white',
                                        highlightthickness=0)
        self.progress_canvas.pack(fill=tk.X, pady=(10, 5))
        
        self.progress_label = tk.Label(prog_container, text='0%', 
                                      font=('Segoe UI', 10, 'bold'),
                                      bg='white', fg=self.colors['text_secondary'])
        self.progress_label.pack(anchor='e')
        
        # Recent Activity
        activity_frame = tk.LabelFrame(self.dashboard_frame, 
                                      text="  📝 Recent Activity  ",
                                      font=('Segoe UI', 12, 'bold'),
                                      bg='white', fg=self.colors['text_primary'],
                                      relief=tk.FLAT, bd=2,
                                      highlightbackground=self.colors['border'],
                                      highlightthickness=1)
        activity_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=(0, 20))
        
        # Activity list with scrollbar
        activity_container = tk.Frame(activity_frame, bg='white')
        activity_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        self.recent_tasks_text = tk.Text(activity_container, wrap=tk.WORD, 
                                        bg='#fafafa', font=('Segoe UI', 10),
                                        relief=tk.FLAT, bd=0, padx=15, pady=10)
        vbar = ttk.Scrollbar(activity_container, orient=tk.VERTICAL, 
                           command=self.recent_tasks_text.yview)
        self.recent_tasks_text.configure(yscrollcommand=vbar.set)
        
        self.recent_tasks_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vbar.pack(side=tk.RIGHT, fill=tk.Y)
        
    def create_modern_stat_card(self, parent, title, value, icon, color):
        """Create a modern stat card with gradient and shadow effect"""
        # Card container with shadow effect
        card_container = tk.Frame(parent, bg='white', relief=tk.FLAT)
        
        # Inner card with border
        card = tk.Frame(card_container, bg='white', relief=tk.FLAT, bd=0,
                       highlightbackground=self.colors['border'],
                       highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        # Top section with icon
        top_frame = tk.Frame(card, bg='white')
        top_frame.pack(fill=tk.X, padx=20, pady=(20, 10))
        
        # Icon with colored background
        icon_frame = tk.Frame(top_frame, bg=color, width=50, height=50)
        icon_frame.pack(side=tk.LEFT)
        icon_frame.pack_propagate(False)
        
        tk.Label(icon_frame, text=icon, font=('Segoe UI', 20),
                bg=color, fg='white').pack(expand=True)
        
        # Value section
        value_frame = tk.Frame(card, bg='white')
        value_frame.pack(fill=tk.X, padx=20, pady=(0, 10))
        
        value_label = tk.Label(value_frame, text=value,
                              font=('Segoe UI', 28, 'bold'),
                              bg='white', fg=color)
        value_label.pack(anchor='w')
        
        # Title section
        tk.Label(card, text=title, font=('Segoe UI', 10),
                bg='white', fg=self.colors['text_secondary']).pack(
                anchor='w', padx=20, pady=(0, 20))
        
        # Store value label for updates
        card_container.value_label = value_label
        
        return card_container
    
    def draw_progress_bar(self, percentage):
        """Draw an animated progress bar"""
        self.progress_canvas.delete('all')
        width = self.progress_canvas.winfo_width()
        if width <= 1:
            width = 400
        
        height = 30
        
        # Background
        self.progress_canvas.create_rectangle(0, 0, width, height,
                                             fill='#f1f5f9', outline='')
        
        # Progress fill with gradient effect
        if percentage > 0:
            progress_width = (width * percentage) / 100
            
            # Create gradient effect with multiple rectangles
            steps = 20
            for i in range(int(progress_width)):
                intensity = 1 - (i / progress_width) * 0.3
                color = self.lighten_color(self.colors['success'], intensity)
                self.progress_canvas.create_rectangle(i, 0, i+1, height,
                                                     fill=color, outline='')
        
        # Percentage text
        self.progress_canvas.create_text(width/2, height/2,
                                        text=f'{percentage}%',
                                        font=('Segoe UI', 11, 'bold'),
                                        fill='white' if percentage > 50 else self.colors['text_primary'])
    
    def lighten_color(self, color, factor=1.0):
        """Lighten a hex color"""
        color = color.lstrip('#')
        rgb = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
        rgb = tuple(min(255, int(c * factor)) for c in rgb)
        return f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'
    
    def create_stat_card(self, parent, title, value, color):
        """Legacy method - redirects to modern version"""
        return self.create_modern_stat_card(parent, title, value, "📊", color)
    
    def setup_charter_tab(self):
        frame = tk.Frame(self.charter_frame, bg='white')
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        self.charter_display = tk.Text(frame, wrap=tk.WORD, font=('Arial', 11), bg='#f9f9f9', relief=tk.FLAT)
        self.charter_display.pack(fill=tk.BOTH, expand=True)
        
        self.charter_display.tag_config('h1', font=('Arial', 16, 'bold'), foreground=self.colors['primary'])
        self.charter_display.tag_config('h2', font=('Arial', 12, 'bold'), foreground='black')
        
    def setup_tasks_tab(self):
        # Header with title and actions
        header = tk.Frame(self.tasks_frame, bg='white')
        header.pack(fill=tk.X, padx=20, pady=(20, 15))
        
        tk.Label(header, text='Task Management', 
                font=('Segoe UI', 18, 'bold'),
                bg='white', fg=self.colors['text_primary']).pack(side=tk.LEFT)
        
        # Action buttons on the right
        actions_frame = tk.Frame(header, bg='white')
        actions_frame.pack(side=tk.RIGHT)
        
        # Modern button style
        btn_config = {
            'font': ('Segoe UI', 10, 'bold'),
            'relief': tk.FLAT,
            'cursor': 'hand2',
            'bd': 0,
            'padx': 20,
            'pady': 10
        }
        
        # Delete button
        delete_btn = tk.Button(actions_frame, text="🗑️  Delete", 
                              bg=self.colors['danger'], fg='white',
                              command=self.delete_task, **btn_config)
        delete_btn.pack(side=tk.RIGHT, padx=(5, 0))
        delete_btn.bind('<Enter>', lambda e: delete_btn.config(bg='#dc2626'))
        delete_btn.bind('<Leave>', lambda e: delete_btn.config(bg=self.colors['danger']))
        
        # Edit button
        edit_btn = tk.Button(actions_frame, text="✏️  Edit", 
                            bg=self.colors['info'], fg='white',
                            command=self.edit_task_dialog, **btn_config)
        edit_btn.pack(side=tk.RIGHT, padx=5)
        edit_btn.bind('<Enter>', lambda e: edit_btn.config(bg='#2563eb'))
        edit_btn.bind('<Leave>', lambda e: edit_btn.config(bg=self.colors['info']))
        
        # Add button
        add_btn = tk.Button(actions_frame, text="➕  Add Task", 
                           bg=self.colors['success'], fg='white',
                           command=self.add_task_dialog, **btn_config)
        add_btn.pack(side=tk.RIGHT, padx=5)
        add_btn.bind('<Enter>', lambda e: add_btn.config(bg='#059669'))
        add_btn.bind('<Leave>', lambda e: add_btn.config(bg=self.colors['success']))
        
        # Search and filter bar
        filter_frame = tk.Frame(self.tasks_frame, bg='white')
        filter_frame.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        tk.Label(filter_frame, text='🔍', font=('Segoe UI', 14),
                bg='white').pack(side=tk.LEFT, padx=(0, 5))
        
        self.task_search_var = tk.StringVar()
        search_entry = tk.Entry(filter_frame, textvariable=self.task_search_var,
                               font=('Segoe UI', 10), bg='#f8fafc',
                               relief=tk.FLAT, bd=0)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, padx=5)
        search_entry.insert(0, 'Search tasks...')
        search_entry.config(fg=self.colors['text_secondary'])
        
        # Placeholder behavior
        def on_focus_in(e):
            if search_entry.get() == 'Search tasks...':
                search_entry.delete(0, tk.END)
                search_entry.config(fg=self.colors['text_primary'])
        
        def on_focus_out(e):
            if not search_entry.get():
                search_entry.insert(0, 'Search tasks...')
                search_entry.config(fg=self.colors['text_secondary'])
        
        search_entry.bind('<FocusIn>', on_focus_in)
        search_entry.bind('<FocusOut>', on_focus_out)
        
        # Filter by status
        tk.Label(filter_frame, text='Status:', font=('Segoe UI', 10),
                bg='white', fg=self.colors['text_secondary']).pack(side=tk.LEFT, padx=(15, 5))
        
        self.status_filter_var = tk.StringVar(value='All')
        status_combo = ttk.Combobox(filter_frame, textvariable=self.status_filter_var,
                                   values=['All', 'Not Started', 'In Progress', 'Completed', 'On Hold'],
                                   width=12, font=('Segoe UI', 9), state='readonly')
        status_combo.pack(side=tk.LEFT, padx=5)
        
        # Table container with border
        table_container = tk.Frame(self.tasks_frame, bg='white',
                                  highlightbackground=self.colors['border'],
                                  highlightthickness=1)
        table_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        tree_frame = tk.Frame(table_container, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        columns = ('Name', 'Start Date', 'Duration', 'End Date', 'Assigned To', 'Priority', 'Status', 'Cost', 'Milestone')
        self.tasks_tree = ttk.Treeview(tree_frame, columns=columns, show='tree headings',
                                      style='Modern.Treeview')
        
        self.tasks_tree.column('#0', width=50)
        self.tasks_tree.heading('#0', text='ID')
        
        self.tasks_tree.column('Name', width=180, minwidth=100)
        self.tasks_tree.column('Start Date', width=100, minwidth=80)
        self.tasks_tree.column('Duration', width=80, minwidth=60)
        self.tasks_tree.column('End Date', width=100, minwidth=80)
        self.tasks_tree.column('Assigned To', width=120, minwidth=80)
        self.tasks_tree.column('Priority', width=80, minwidth=60)
        self.tasks_tree.column('Status', width=100, minwidth=80)
        self.tasks_tree.column('Cost', width=90, minwidth=60)
        self.tasks_tree.column('Milestone', width=80, minwidth=60)
        
        for col in columns:
            self.tasks_tree.heading(col, text=col)
        
        v_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tasks_tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tasks_tree.xview)
        self.tasks_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        self.tasks_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
    
    def setup_gantt_tab(self):
        self.gantt_canvas_frame = tk.Frame(self.gantt_frame, bg='white')
        self.gantt_canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    def setup_timeline_tab(self):
        self.timeline_canvas_frame = tk.Frame(self.timeline_frame, bg='white')
        self.timeline_canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    def setup_kanban_tab(self):
        self.kanban_columns = {}
        statuses = ["Not Started", "In Progress", "Completed", "On Hold"]
        for status in statuses:
            frame = tk.LabelFrame(self.kanban_frame, text=status, bg='white')
            frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
            canvas = tk.Canvas(frame, bg='white', height=500)
            scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg='white')
            scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.kanban_columns[status] = scrollable_frame
        
        self.update_kanban_view()
    
    def setup_resources_tab(self):
        tk.Label(self.resources_frame, text="Team Resources", 
                font=('Arial', 14, 'bold'), bg='white').pack(pady=10)
        
        list_frame = tk.Frame(self.resources_frame, bg='white')
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.resources_listbox = tk.Listbox(list_frame, font=('Arial', 11),
                                           yscrollcommand=scrollbar.set, height=10)
        self.resources_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.resources_listbox.yview)
        
        tk.Button(self.resources_frame, text="View Allocation Histogram", 
                 command=self.update_resource_histogram, bg=self.colors['primary'], fg='white').pack(pady=10)
        
        self.resource_hist_frame = tk.Frame(self.resources_frame, bg='white')
        self.resource_hist_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        btn_frame = tk.Frame(self.resources_frame, bg='white')
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="Add Resource", command=self.add_resource,
                 bg=self.colors['success'], fg='white', width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Remove Resource", command=self.remove_resource,
                 bg=self.colors['danger'], fg='white', width=15).pack(side=tk.LEFT, padx=5)
    
    def setup_risk_tab(self):
        tk.Label(self.risk_frame, text="Risk Register", 
                font=('Arial', 14, 'bold'), bg='white').pack(pady=10)
        
        tree_frame = tk.Frame(self.risk_frame, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        columns = ('Name', 'Probability', 'Impact', 'Score', 'Mitigation')
        self.risks_tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
        for col in columns:
            self.risks_tree.heading(col, text=col)
            self.risks_tree.column(col, width=150)
        self.risks_tree.pack(fill=tk.BOTH, expand=True)
        
        self.update_risks_tree()
    
    def setup_research_tab(self):
        paned = tk.PanedWindow(self.research_frame, orient=tk.HORIZONTAL, sashrelief=tk.RAISED)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        left = tk.Frame(paned, bg='white')
        paned.add(left, width=300)
        
        tk.Label(left, text="Research Logs", font=('Arial', 12, 'bold'), bg='white').pack(pady=5)
        
        btn_bar = tk.Frame(left, bg='white')
        btn_bar.pack(fill=tk.X, pady=5)
        tk.Button(btn_bar, text="➕ New", command=self.new_research_entry, bg=self.colors['success'], fg='white').pack(side=tk.LEFT, padx=5)
        tk.Button(btn_bar, text="❌ Delete", command=self.delete_research_entry, bg=self.colors['danger'], fg='white').pack(side=tk.LEFT, padx=5)
        
        columns = ('Title', 'Date')
        self.res_tree = ttk.Treeview(left, columns=columns, show='headings')
        self.res_tree.heading('Title', text='Title')
        self.res_tree.heading('Date', text='Date')
        self.res_tree.column('Title', width=180)
        self.res_tree.column('Date', width=100)
        self.res_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.res_tree.bind('<<TreeviewSelect>>', self.load_research_entry)
        
        right = tk.Frame(paned, bg='white')
        paned.add(right)
        
        header = tk.Frame(right, bg='white')
        header.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(header, text="Title:", font=('Arial', 10, 'bold'), bg='white').pack(side=tk.LEFT)
        self.res_title_var = tk.StringVar()
        tk.Entry(header, textvariable=self.res_title_var, width=40).pack(side=tk.LEFT, padx=10)
        tk.Button(header, text="💾 Save Log", command=self.save_research_entry, bg=self.colors['primary'], fg='white').pack(side=tk.RIGHT)
        
        self.res_text = tk.Text(right, wrap=tk.WORD, font=('Consolas', 11), padx=10, pady=10)
        self.res_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    def setup_documents_tab(self):
        main_frame = tk.Frame(self.documents_frame, bg='white')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        controls_frame = tk.Frame(main_frame, bg='white')
        controls_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Button(controls_frame, text="📁 Add Document", 
                 command=self.add_document_dialog,
                 bg=self.colors['success'], fg='white').pack(side=tk.LEFT, padx=5)
        
        tk.Button(controls_frame, text="👁️ View Document", 
                 command=self.view_document,
                 bg=self.colors['primary'], fg='white').pack(side=tk.LEFT, padx=5)
        
        tk.Button(controls_frame, text="🗑️ Delete Document", 
                 command=self.delete_document,
                 bg=self.colors['danger'], fg='white').pack(side=tk.LEFT, padx=5)
        
        tk.Button(controls_frame, text="📊 Document Stats", 
                 command=self.show_document_stats,
                 bg=self.colors['warning'], fg='white').pack(side=tk.LEFT, padx=5)
        
        search_frame = tk.Frame(controls_frame, bg='white')
        search_frame.pack(side=tk.RIGHT, padx=5)
        
        tk.Label(search_frame, text="Search:", bg='white').pack(side=tk.LEFT)
        self.doc_search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=self.doc_search_var, width=20)
        search_entry.pack(side=tk.LEFT, padx=5)
        search_entry.bind('<KeyRelease>', self.filter_documents)
        
        tree_frame = tk.Frame(main_frame, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ('Name', 'Type', 'Category', 'Size', 'Date', 'Tags')
        self.documents_tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
        
        self.documents_tree.column('Name', width=200)
        self.documents_tree.column('Type', width=100)
        self.documents_tree.column('Category', width=100)
        self.documents_tree.column('Size', width=80)
        self.documents_tree.column('Date', width=120)
        self.documents_tree.column('Tags', width=150)
        
        for col in columns:
            self.documents_tree.heading(col, text=col)
        
        v_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.documents_tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.documents_tree.xview)
        self.documents_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        self.documents_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.documents_tree.bind('<Double-1>', lambda e: self.view_document())
    
    def setup_meeting_tab(self):
        main_frame = tk.Frame(self.meeting_frame, bg='white')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        controls_frame = tk.Frame(main_frame, bg='white')
        controls_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Button(controls_frame, text="📝 New Meeting", 
                 command=self.add_meeting_note_dialog,
                 bg=self.colors['success'], fg='white').pack(side=tk.LEFT, padx=5)
        
        tk.Button(controls_frame, text="👁️ View Meeting", 
                 command=self.view_meeting_note,
                 bg=self.colors['primary'], fg='white').pack(side=tk.LEFT, padx=5)
        
        tk.Button(controls_frame, text="🗑️ Delete Meeting", 
                 command=self.delete_meeting_note,
                 bg=self.colors['danger'], fg='white').pack(side=tk.LEFT, padx=5)
        
        tree_frame = tk.Frame(main_frame, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ('Date', 'Title', 'Participants', 'Actions')
        self.meetings_tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
        
        self.meetings_tree.column('Date', width=100)
        self.meetings_tree.column('Title', width=200)
        self.meetings_tree.column('Participants', width=150)
        self.meetings_tree.column('Actions', width=200)
        
        for col in columns:
            self.meetings_tree.heading(col, text=col)
        
        v_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.meetings_tree.yview)
        self.meetings_tree.configure(yscrollcommand=v_scroll.set)
        
        self.meetings_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.meetings_tree.bind('<Double-1>', lambda e: self.view_meeting_note())
    
    def setup_portfolio_tab(self):
        tk.Label(self.portfolio_frame, text="Portfolio Overview", 
                font=('Arial', 14, 'bold'), bg='white').pack(pady=10)
        
        self.portfolio_text = tk.Text(self.portfolio_frame, wrap=tk.WORD, bg='#f9f9f9', font=('Arial', 10))
        self.portfolio_text.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        self.update_portfolio_view()

    def new_project_wizard(self):
        wiz = tk.Toplevel(self.root)
        wiz.title("Initialize New Project")
        wiz.geometry("900x750")
        wiz.configure(bg=self.colors['bg'])
        wiz.transient(self.root)
        wiz.grab_set()
        
        header = tk.Frame(wiz, bg=self.colors['primary'], height=70)
        header.pack(fill=tk.X)
        tk.Label(header, text="Project Definition & Scope", font=('Arial', 18, 'bold'), 
                bg=self.colors['primary'], fg='white').pack(pady=20)
        
        canvas = tk.Canvas(wiz, bg=self.colors['bg'])
        scrollbar = ttk.Scrollbar(wiz, orient="vertical", command=canvas.yview)
        form_frame = tk.Frame(canvas, bg=self.colors['bg'])
        
        form_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=form_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        scrollbar.pack(side="right", fill="y")
        
        sec1 = tk.LabelFrame(form_frame, text="1. Project Identity", font=('Arial', 11, 'bold'), bg='white', padx=10, pady=10)
        sec1.pack(fill=tk.X, pady=10)
        
        tk.Label(sec1, text="Project Name:", bg='white', font=('Arial', 10)).grid(row=0, column=0, sticky='w', pady=5)
        name_entry = tk.Entry(sec1, width=40, font=('Arial', 10))
        name_entry.grid(row=0, column=1, padx=10, pady=5)
        
        tk.Label(sec1, text="Estimated Budget ($):", bg='white', font=('Arial', 10)).grid(row=0, column=2, sticky='w', pady=5)
        budget_entry = tk.Entry(sec1, width=20, font=('Arial', 10))
        budget_entry.insert(0, "0.00")
        budget_entry.grid(row=0, column=3, padx=10, pady=5)
        
        tk.Label(sec1, text="Save Location:", bg='white', font=('Arial', 10)).grid(row=1, column=0, sticky='w', pady=5)
        path_var = tk.StringVar(value=self.base_storage_path)
        tk.Entry(sec1, textvariable=path_var, width=40, font=('Arial', 10)).grid(row=1, column=1, padx=10, pady=5)
        tk.Button(sec1, text="Browse...", command=lambda: path_var.set(filedialog.askdirectory())).grid(row=1, column=2)

        sec2 = tk.LabelFrame(form_frame, text="2. Problem Statement", font=('Arial', 11, 'bold'), bg='white', padx=10, pady=10)
        sec2.pack(fill=tk.X, pady=10)
        tk.Label(sec2, text="Describe the issue or opportunity this project addresses:", bg='white', fg='gray').pack(anchor='w')
        prob_text = tk.Text(sec2, height=5, width=80, font=('Arial', 10))
        prob_text.pack(pady=5, fill=tk.X)

        sec3 = tk.LabelFrame(form_frame, text="3. Objectives & Insights", font=('Arial', 11, 'bold'), bg='white', padx=10, pady=10)
        sec3.pack(fill=tk.X, pady=10)
        tk.Label(sec3, text="List key deliverables, technical insights, and success criteria:", bg='white', fg='gray').pack(anchor='w')
        obj_text = tk.Text(sec3, height=5, width=80, font=('Arial', 10))
        obj_text.pack(pady=5, fill=tk.X)

        sec4 = tk.LabelFrame(form_frame, text="4. Cost-Benefit Analysis", font=('Arial', 11, 'bold'), bg='white', padx=10, pady=10)
        sec4.pack(fill=tk.X, pady=10)
        tk.Label(sec4, text="Detail the ROI, intangible benefits, and financial justification:", bg='white', fg='gray').pack(anchor='w')
        cb_text = tk.Text(sec4, height=5, width=80, font=('Arial', 10))
        cb_text.pack(pady=5, fill=tk.X)
        
        def confirm_create():
            p_name = name_entry.get().strip()
            p_path = path_var.get().strip()
            
            if not p_name or not p_path:
                messagebox.showerror("Error", "Project Name and Location are required!")
                return
            
            project_folder = os.path.join(p_path, p_name)
            if os.path.exists(project_folder):
                if not messagebox.askyesno("Confirm", f"Folder '{p_name}' exists. Overwrite?"):
                    return
                shutil.rmtree(project_folder)
            
            try:
                os.makedirs(project_folder)
                os.makedirs(os.path.join(project_folder, 'documents'), exist_ok=True)
                os.makedirs(os.path.join(project_folder, 'images'), exist_ok=True)
                
                self.project_name = p_name
                self.current_folder = project_folder
                
                self.project_meta = {
                    'name': p_name,
                    'problem_statement': prob_text.get("1.0", tk.END).strip(),
                    'objectives': obj_text.get("1.0", tk.END).strip(),
                    'cost_benefit': cb_text.get("1.0", tk.END).strip(),
                    'budget': float(budget_entry.get()) if budget_entry.get() else 0.0
                }
                
                self.tasks = []
                self.resources = []
                self.risks = []
                self.baselines = []
                self.research_logs = []
                self.documents = []
                self.meeting_notes = []
                self.images = []
                
                self.project_title_label.config(text=f"📊 {self.project_name}")
                self._save_to_folder()
                self.refresh_all()
                wiz.destroy()
                messagebox.showinfo("Success", f"Project '{p_name}' initialized in local storage!")
                self.notebook.select(self.charter_frame)
                
            except Exception as e:
                messagebox.showerror("Error", f"Creation failed: {str(e)}")

        btn_frame = tk.Frame(form_frame, bg=self.colors['bg'])
        btn_frame.pack(pady=20)
        tk.Button(btn_frame, text="Create Project", command=confirm_create, 
                 bg=self.colors['success'], fg='white', font=('Arial', 12, 'bold'), width=20).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Cancel", command=wiz.destroy, 
                 bg=self.colors['danger'], fg='white', font=('Arial', 12, 'bold'), width=10).pack(side=tk.LEFT, padx=10)

    def new_research_entry(self):
        self.res_tree.selection_remove(self.res_tree.selection())
        self.res_title_var.set("")
        self.res_text.delete(1.0, tk.END)
        
    def save_research_entry(self):
        title = self.res_title_var.get().strip()
        content = self.res_text.get(1.0, tk.END).strip()
        if not title:
            messagebox.showwarning("Warning", "Title required")
            return
        
        date_str = datetime.now().strftime('%Y-%m-%d %H:%M')
        entry = {'title': title, 'date': date_str, 'content': content}
        
        selected = self.res_tree.selection()
        if selected:
            idx = int(self.res_tree.index(selected[0]))
            self.research_logs[idx] = entry
        else:
            self.research_logs.append(entry)
            
        self.update_research_tree()
        self.save_project()
        messagebox.showinfo("Success", "Research log saved locally.")
        
    def delete_research_entry(self):
        selected = self.res_tree.selection()
        if selected:
            if messagebox.askyesno("Confirm", "Delete selected entry?"):
                idx = int(self.res_tree.index(selected[0]))
                del self.research_logs[idx]
                self.new_research_entry()
                self.update_research_tree()
                self.save_project()
                
    def load_research_entry(self, event):
        selected = self.res_tree.selection()
        if selected:
            idx = int(self.res_tree.index(selected[0]))
            data = self.research_logs[idx]
            self.res_title_var.set(data['title'])
            self.res_text.delete(1.0, tk.END)
            self.res_text.insert(tk.END, data['content'])
            
    def update_research_tree(self):
        for item in self.res_tree.get_children():
            self.res_tree.delete(item)
        for log in self.research_logs:
            self.res_tree.insert('', 'end', values=(log['title'], log['date']))

    def update_charter_view(self):
        self.charter_display.config(state=tk.NORMAL)
        self.charter_display.delete(1.0, tk.END)
        
        self.charter_display.insert(tk.END, f"PROJECT CHARTER: {self.project_meta.get('name', 'Untitled')}\n\n", 'h1')
        
        self.charter_display.insert(tk.END, "1. Problem Statement\n", 'h2')
        self.charter_display.insert(tk.END, f"{self.project_meta.get('problem_statement', 'N/A')}\n\n")
        
        self.charter_display.insert(tk.END, "2. Strategic Objectives\n", 'h2')
        self.charter_display.insert(tk.END, f"{self.project_meta.get('objectives', 'N/A')}\n\n")
        
        self.charter_display.insert(tk.END, "3. Cost-Benefit Analysis\n", 'h2')
        self.charter_display.insert(tk.END, f"{self.project_meta.get('cost_benefit', 'N/A')}\n\n")
        
        self.charter_display.insert(tk.END, "4. Budget Overview\n", 'h2')
        self.charter_display.insert(tk.END, f"Allocated Budget: ${self.project_meta.get('budget', 0.0):,.2f}\n")
        
        self.charter_display.config(state=tk.DISABLED)

    def add_task_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add New Task")
        dialog.geometry("500x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="Task Name:", font=('Arial', 10, 'bold')).grid(row=0, column=0, padx=10, pady=10, sticky='e')
        name_entry = tk.Entry(dialog, width=30, font=('Arial', 10))
        name_entry.grid(row=0, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Start Date (YYYY-MM-DD):", font=('Arial', 10, 'bold')).grid(row=1, column=0, padx=10, pady=10, sticky='e')
        date_entry = tk.Entry(dialog, width=30, font=('Arial', 10))
        date_entry.insert(0, datetime.now().strftime('%Y-%m-%d'))
        date_entry.grid(row=1, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Duration (days):", font=('Arial', 10, 'bold')).grid(row=2, column=0, padx=10, pady=10, sticky='e')
        duration_entry = tk.Entry(dialog, width=30, font=('Arial', 10))
        duration_entry.insert(0, "5")
        duration_entry.grid(row=2, column=1, padx=10, pady=10)
        
        milestone_var = tk.BooleanVar()
        tk.Checkbutton(dialog, text="Milestone (zero duration)", variable=milestone_var).grid(row=3, column=0, columnspan=2, pady=5)
        
        tk.Label(dialog, text="Cost ($):", font=('Arial', 10, 'bold')).grid(row=4, column=0, padx=10, pady=10, sticky='e')
        cost_entry = tk.Entry(dialog, width=30, font=('Arial', 10))
        cost_entry.insert(0, "0")
        cost_entry.grid(row=4, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Assigned To:", font=('Arial', 10, 'bold')).grid(row=5, column=0, padx=10, pady=10, sticky='e')
        assigned_var = tk.StringVar()
        assigned_combo = ttk.Combobox(dialog, textvariable=assigned_var, 
                                     values=self.resources, width=28, font=('Arial', 10))
        assigned_combo.grid(row=5, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Priority:", font=('Arial', 10, 'bold')).grid(row=6, column=0, padx=10, pady=10, sticky='e')
        priority_var = tk.StringVar(value="Medium")
        priority_combo = ttk.Combobox(dialog, textvariable=priority_var,
                                     values=["Low", "Medium", "High", "Critical"],
                                     width=28, font=('Arial', 10))
        priority_combo.grid(row=6, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Status:", font=('Arial', 10, 'bold')).grid(row=7, column=0, padx=10, pady=10, sticky='e')
        status_var = tk.StringVar(value="Not Started")
        status_combo = ttk.Combobox(dialog, textvariable=status_var,
                                   values=["Not Started", "In Progress", "Completed", "On Hold"],
                                   width=28, font=('Arial', 10))
        status_combo.grid(row=7, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Dependencies (comma-separated IDs):", 
                font=('Arial', 10, 'bold')).grid(row=8, column=0, padx=10, pady=10, sticky='e')
        dep_entry = tk.Entry(dialog, width=30, font=('Arial', 10))
        dep_entry.grid(row=8, column=1, padx=10, pady=10)
        
        def save_task():
            try:
                name = name_entry.get().strip()
                start_date = datetime.strptime(date_entry.get().strip(), '%Y-%m-%d')
                duration = int(duration_entry.get().strip())
                cost = float(cost_entry.get().strip())
                assigned = assigned_var.get()
                priority = priority_var.get()
                status = status_var.get()
                is_milestone = milestone_var.get()
                
                deps = []
                if dep_entry.get().strip():
                    deps = [int(x.strip()) for x in dep_entry.get().split(',')]
                
                if not name:
                    messagebox.showerror("Error", "Task name is required!")
                    return
                
                task = Task(name, start_date, duration, assigned, priority, status, deps, cost, is_milestone)
                task.id = len(self.tasks)
                self.tasks.append(task)
                
                self.refresh_all()
                self.save_project()
                dialog.destroy()
                messagebox.showinfo("Success", "Task added successfully!")
                
            except ValueError as e:
                messagebox.showerror("Error", f"Invalid input: {str(e)}")
        
        btn_frame = tk.Frame(dialog)
        btn_frame.grid(row=9, column=0, columnspan=2, pady=20)
        
        tk.Button(btn_frame, text="Save", command=save_task,
                 bg=self.colors['success'], fg='white',
                 font=('Arial', 10, 'bold'), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Cancel", command=dialog.destroy,
                 bg=self.colors['danger'], fg='white',
                 font=('Arial', 10, 'bold'), width=12).pack(side=tk.LEFT, padx=5)
    
    def edit_task_dialog(self):
        selected = self.tasks_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a task to edit!")
            return
        
        task_id = int(self.tasks_tree.item(selected[0])['text'])
        task = self.tasks[task_id]
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Task")
        dialog.geometry("500x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="Task Name:", font=('Arial', 10, 'bold')).grid(row=0, column=0, padx=10, pady=10, sticky='e')
        name_entry = tk.Entry(dialog, width=30, font=('Arial', 10))
        name_entry.insert(0, task.name)
        name_entry.grid(row=0, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Start Date (YYYY-MM-DD):", font=('Arial', 10, 'bold')).grid(row=1, column=0, padx=10, pady=10, sticky='e')
        date_entry = tk.Entry(dialog, width=30, font=('Arial', 10))
        date_entry.insert(0, task.start_date.strftime('%Y-%m-%d'))
        date_entry.grid(row=1, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Duration (days):", font=('Arial', 10, 'bold')).grid(row=2, column=0, padx=10, pady=10, sticky='e')
        duration_entry = tk.Entry(dialog, width=30, font=('Arial', 10))
        duration_entry.insert(0, str(task.duration))
        duration_entry.grid(row=2, column=1, padx=10, pady=10)
        
        milestone_var = tk.BooleanVar(value=task.is_milestone)
        tk.Checkbutton(dialog, text="Milestone (zero duration)", variable=milestone_var).grid(row=3, column=0, columnspan=2, pady=5)
        
        tk.Label(dialog, text="Cost ($):", font=('Arial', 10, 'bold')).grid(row=4, column=0, padx=10, pady=10, sticky='e')
        cost_entry = tk.Entry(dialog, width=30, font=('Arial', 10))
        cost_entry.insert(0, str(task.cost))
        cost_entry.grid(row=4, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Assigned To:", font=('Arial', 10, 'bold')).grid(row=5, column=0, padx=10, pady=10, sticky='e')
        assigned_var = tk.StringVar(value=task.assigned_to)
        assigned_combo = ttk.Combobox(dialog, textvariable=assigned_var,
                                     values=self.resources, width=28, font=('Arial', 10))
        assigned_combo.grid(row=5, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Priority:", font=('Arial', 10, 'bold')).grid(row=6, column=0, padx=10, pady=10, sticky='e')
        priority_var = tk.StringVar(value=task.priority)
        priority_combo = ttk.Combobox(dialog, textvariable=priority_var,
                                     values=["Low", "Medium", "High", "Critical"],
                                     width=28, font=('Arial', 10))
        priority_combo.grid(row=6, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Status:", font=('Arial', 10, 'bold')).grid(row=7, column=0, padx=10, pady=10, sticky='e')
        status_var = tk.StringVar(value=task.status)
        status_combo = ttk.Combobox(dialog, textvariable=status_var,
                                   values=["Not Started", "In Progress", "Completed", "On Hold"],
                                   width=28, font=('Arial', 10))
        status_combo.grid(row=7, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Dependencies (comma-separated IDs):",
                font=('Arial', 10, 'bold')).grid(row=8, column=0, padx=10, pady=10, sticky='e')
        dep_entry = tk.Entry(dialog, width=30, font=('Arial', 10))
        dep_entry.insert(0, ','.join(map(str, task.dependencies)))
        dep_entry.grid(row=8, column=1, padx=10, pady=10)
        
        def update_task():
            try:
                task.name = name_entry.get().strip()
                task.start_date = datetime.strptime(date_entry.get().strip(), '%Y-%m-%d')
                task.duration = int(duration_entry.get().strip())
                task.cost = float(cost_entry.get().strip())
                task.is_milestone = milestone_var.get()
                task.assigned_to = assigned_var.get()
                task.priority = priority_var.get()
                task.status = status_var.get()
                
                if dep_entry.get().strip():
                    task.dependencies = [int(x.strip()) for x in dep_entry.get().split(',')]
                else:
                    task.dependencies = []
                
                self.refresh_all()
                self.save_project()
                dialog.destroy()
                messagebox.showinfo("Success", "Task updated successfully!")
                
            except ValueError as e:
                messagebox.showerror("Error", f"Invalid input: {str(e)}")
        
        btn_frame = tk.Frame(dialog)
        btn_frame.grid(row=9, column=0, columnspan=2, pady=20)
        
        tk.Button(btn_frame, text="Update", command=update_task,
                 bg=self.colors['success'], fg='white',
                 font=('Arial', 10, 'bold'), width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Cancel", command=dialog.destroy,
                 bg=self.colors['danger'], fg='white',
                 font=('Arial', 10, 'bold'), width=12).pack(side=tk.LEFT, padx=5)
    
    def delete_task(self):
        selected = self.tasks_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a task to delete!")
            return
        
        if messagebox.askyesno("Confirm", "Are you sure you want to delete this task?"):
            task_id = int(self.tasks_tree.item(selected[0])['text'])
            del self.tasks[task_id]
            
            for i, task in enumerate(self.tasks):
                task.id = i
                task.dependencies = [d for d in task.dependencies if d < len(self.tasks)]
            
            self.refresh_all()
            self.save_project()
            messagebox.showinfo("Success", "Task deleted successfully!")
    
    def manage_resources_dialog(self):
        self.notebook.select(self.resources_frame)
    
    def add_resource(self):
        name = simpledialog.askstring("Add Resource", "Enter resource name:")
        if name and name.strip():
            self.resources.append(name.strip())
            self.refresh_resources_list()
            self.save_project()
            messagebox.showinfo("Success", f"Resource '{name}' added!")
    
    def remove_resource(self):
        selected = self.resources_listbox.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a resource to remove!")
            return
        
        resource = self.resources[selected[0]]
        if messagebox.askyesno("Confirm", f"Remove resource '{resource}'?"):
            self.resources.pop(selected[0])
            self.refresh_resources_list()
            self.save_project()
    
    def refresh_resources_list(self):
        self.resources_listbox.delete(0, tk.END)
        for resource in self.resources:
            self.resources_listbox.insert(tk.END, resource)
    
    def update_tasks_tree(self):
        for item in self.tasks_tree.get_children():
            self.tasks_tree.delete(item)
        
        for i, task in enumerate(self.tasks):
            self.tasks_tree.insert('', 'end', text=str(i),
                                  values=(task.name,
                                         task.start_date.strftime('%Y-%m-%d'),
                                         f"{task.duration} days" if not task.is_milestone else "Milestone",
                                         task.end_date().strftime('%Y-%m-%d'),
                                         task.assigned_to,
                                         task.priority,
                                         task.status,
                                         f"${task.cost}",
                                         "Yes" if task.is_milestone else "No"))
    
    def update_dashboard(self):
        total = len(self.tasks)
        completed = sum(1 for t in self.tasks if t.status == "Completed")
        in_progress = sum(1 for t in self.tasks if t.status == "In Progress")
        not_started = sum(1 for t in self.tasks if t.status == "Not Started")
        
        if self.tasks:
            planned_end = max((t.end_date() for t in self.tasks), default=datetime.now())
            start_date = min(t.start_date for t in self.tasks)
            progress_days = (planned_end - start_date).days
            actual_progress = completed / total * progress_days if total else 0
            actual_end = planned_end - timedelta(days=actual_progress)
            variance = (planned_end - actual_end).days
        else:
            variance = 0
        
        total_cost = sum(t.cost for t in self.tasks)
        
        # Update stat cards
        self.total_tasks_card.value_label.config(text=str(total))
        self.completed_card.value_label.config(text=str(completed))
        self.in_progress_card.value_label.config(text=str(in_progress))
        self.not_started_card.value_label.config(text=str(not_started))
        self.variance_card.value_label.config(text=f"{variance} days")
        self.total_cost_card.value_label.config(text=f"${total_cost:,.2f}")
        
        # Update progress bar
        completion_percentage = int((completed / total * 100) if total > 0 else 0)
        self.draw_progress_bar(completion_percentage)
        self.progress_label.config(text=f'{completion_percentage}% Complete')
        
        # Update recent tasks with better formatting
        self.recent_tasks_text.delete(1.0, tk.END)
        
        if not self.tasks:
            self.recent_tasks_text.insert(tk.END, 
                "No tasks yet. Create your first task to get started! 🚀\n\n")
        else:
            # Show last 10 tasks
            recent_tasks = self.tasks[-10:][::-1]
            for i, task in enumerate(recent_tasks):
                # Status emoji
                status_emoji = {
                    'Completed': '✅',
                    'In Progress': '⏳',
                    'Not Started': '⭕',
                    'On Hold': '⏸️'
                }.get(task.status, '📋')
                
                # Priority emoji
                priority_emoji = {
                    'Critical': '🔴',
                    'High': '🟠',
                    'Medium': '🟡',
                    'Low': '🟢'
                }.get(task.priority, '⚪')
                
                # Task entry
                self.recent_tasks_text.insert(tk.END, 
                    f"{status_emoji} {task.name}\n",
                    'task_name')
                
                self.recent_tasks_text.insert(tk.END,
                    f"   {priority_emoji} {task.priority} Priority  |  "
                    f"👤 {task.assigned_to or 'Unassigned'}  |  "
                    f"💰 ${task.cost:,.2f}\n",
                    'task_details')
                
                self.recent_tasks_text.insert(tk.END,
                    f"   📅 {task.start_date.strftime('%b %d, %Y')} → "
                    f"{task.end_date().strftime('%b %d, %Y')}\n\n",
                    'task_dates')
        
        # Configure text tags for styling
        self.recent_tasks_text.tag_config('task_name', 
                                         font=('Segoe UI', 11, 'bold'),
                                         foreground=self.colors['text_primary'])
        self.recent_tasks_text.tag_config('task_details',
                                         font=('Segoe UI', 9),
                                         foreground=self.colors['text_secondary'])
        self.recent_tasks_text.tag_config('task_dates',
                                         font=('Segoe UI', 9),
                                         foreground=self.colors['text_secondary'])
        
        # Update project info
        info = f"📊 Project: {self.project_name}\n\n"
        if self.current_folder:
            info += f"📁 Location:\n{self.current_folder}\n\n"
        info += f"📋 Total Tasks: {total}\n"
        info += f"✅ Completed: {completed}\n"
        info += f"⏳ In Progress: {in_progress}\n"
        info += f"⭕ Not Started: {not_started}\n\n"
        info += f"👥 Resources: {len(self.resources)}\n"
        info += f"💰 Total Cost: ${total_cost:,.2f}\n"
        info += f"📊 Variance: {variance} days\n\n"
        info += f"📁 Documents: {len(self.documents)}\n"
        info += f"📝 Meetings: {len(self.meeting_notes)}\n"
        info += f"⚠️ Risks: {len(self.risks)}"
        
        self.project_info_label.config(text=info)
    
    def get_critical_path(self):
        if not self.tasks:
            return set()
        G = nx.DiGraph()
        for task in self.tasks:
            G.add_node(task.id, start=task.start_date)
        for task in self.tasks:
            for dep in task.dependencies:
                if dep < len(self.tasks):
                    G.add_edge(dep, task.id)
        try:
            path = nx.dag_longest_path(G, weight=lambda u, v, e: self.tasks[v].duration)
            return set(path)
        except:
            return set()
    
    def update_gantt_chart(self):
        for widget in self.gantt_canvas_frame.winfo_children():
            widget.destroy()
        
        if not self.tasks:
            tk.Label(self.gantt_canvas_frame, text="No tasks to display",
                    font=('Arial', 12), bg='white').pack(expand=True)
            return
        
        critical_ids = self.get_critical_path()
        
        fig = Figure(figsize=(12, 6), dpi=100)
        ax = fig.add_subplot(111)
        
        sorted_tasks = sorted(self.tasks, key=lambda t: t.start_date)
        
        colors_map = {
            'Not Started': '#F44336',
            'In Progress': '#FF9800',
            'Completed': '#4CAF50',
            'On Hold': '#9E9E9E'
        }
        
        for i, task in enumerate(sorted_tasks):
            start = mdates.date2num(task.start_date)
            duration = task.duration
            color = colors_map.get(task.status, '#2196F3')
            if task.id in critical_ids:
                color = self.colors['critical']
            if task.is_milestone:
                ax.plot(start, i, marker='D', color=color, markersize=10, label='Milestone' if i == 0 else "")
            else:
                ax.barh(i, duration, left=start, height=0.5, 
                       color=color, alpha=0.8, label=task.status if i == 0 else "")
            
            ax.text(start + duration/2, i, task.name[:10], 
                   ha='center', va='center', fontsize=8, fontweight='bold')
        
        if self.baselines:
            baseline = self.baselines[-1]
            for j, btask in enumerate(baseline):
                if not btask.is_milestone:
                    ax.barh(j, btask.duration, left=mdates.date2num(btask.start_date), 
                           height=0.3, color='gray', alpha=0.3)
        
        ax.set_yticks(range(len(sorted_tasks)))
        ax.set_yticklabels([f"Task {t.id}" for t in sorted_tasks])
        ax.set_xlabel('Timeline', fontsize=10, fontweight='bold')
        ax.set_ylabel('Tasks', fontsize=10, fontweight='bold')
        ax.set_title('Project Gantt Chart (Critical Path Highlighted)', fontsize=12, fontweight='bold')
        
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, len(sorted_tasks)//5)))
        fig.autofmt_xdate()
        
        ax.grid(True, alpha=0.3)
        
        handles, labels = ax.get_legend_handles_labels()
        unique_labels = dict(zip(labels, handles))
        ax.legend(unique_labels.values(), unique_labels.keys(), loc='upper right')
        
        fig.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, master=self.gantt_canvas_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def update_timeline_view(self):
        for widget in self.timeline_canvas_frame.winfo_children():
            widget.destroy()
        
        milestones = [t for t in self.tasks if t.is_milestone]
        if not milestones:
            tk.Label(self.timeline_canvas_frame, text="No milestones defined",
                    font=('Arial', 12), bg='white').pack(expand=True)
            return
        
        fig = Figure(figsize=(12, 4), dpi=100)
        ax = fig.add_subplot(111)
        
        dates = [mdates.date2num(m.start_date) for m in milestones]
        labels = [m.name[:15] for m in milestones]
        
        ax.plot(dates, range(len(milestones)), marker='o', linestyle='-', color='blue')
        ax.set_xticks(dates)
        ax.set_xticklabels([d.strftime('%Y-%m-%d') for d in [m.start_date for m in milestones]], rotation=45)
        ax.set_yticks(range(len(milestones)))
        ax.set_yticklabels(labels)
        ax.set_title('Project Timeline (Milestones)')
        ax.grid(True, alpha=0.3)
        
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=self.timeline_canvas_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def update_kanban_view(self):
        for status, frame in self.kanban_columns.items():
            for widget in frame.winfo_children():
                widget.destroy()
        
        for task in self.tasks:
            status = task.status
            if status in self.kanban_columns:
                label = tk.Label(self.kanban_columns[status], text=f"{task.name} (ID:{task.id})", 
                                bg=self.colors['card'], relief=tk.RAISED, padx=5, pady=2, wraplength=200)
                label.pack(pady=2, padx=5)
    
    def update_resource_histogram(self):
        for widget in self.resource_hist_frame.winfo_children():
            widget.destroy()
        
        if not self.tasks or not self.resources:
            tk.Label(self.resource_hist_frame, text="No data for histogram",
                    font=('Arial', 12), bg='white').pack(expand=True)
            return
        
        workload = {r: 0 for r in self.resources}
        unassigned = 0
        for task in self.tasks:
            if task.assigned_to in workload:
                workload[task.assigned_to] += task.duration
            else:
                unassigned += task.duration
        
        fig = Figure(figsize=(8, 5), dpi=100)
        ax = fig.add_subplot(111)
        resources = list(workload.keys()) + (['Unassigned'] if unassigned > 0 else [])
        loads = [workload[r] for r in resources[:-1]] + ([unassigned] if unassigned > 0 else [])
        ax.bar(resources, loads, color='skyblue')
        ax.set_ylabel('Total Days Allocated')
        ax.set_title('Resource Allocation Histogram')
        ax.tick_params(axis='x', rotation=45)
        fig.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, master=self.resource_hist_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def add_risk_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Risk")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="Risk Name:").grid(row=0, column=0, padx=10, pady=10, sticky='e')
        name_entry = tk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Probability (1-10):").grid(row=1, column=0, padx=10, pady=10, sticky='e')
        prob_entry = tk.Entry(dialog, width=30)
        prob_entry.grid(row=1, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Impact (1-10):").grid(row=2, column=0, padx=10, pady=10, sticky='e')
        impact_entry = tk.Entry(dialog, width=30)
        impact_entry.grid(row=2, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Mitigation:").grid(row=3, column=0, padx=10, pady=10, sticky='ne')
        mit_text = tk.Text(dialog, width=30, height=5)
        mit_text.grid(row=3, column=1, padx=10, pady=10)
        
        def save_risk():
            try:
                name = name_entry.get().strip()
                prob = int(prob_entry.get().strip())
                impact = int(impact_entry.get().strip())
                score = prob * impact
                mit = mit_text.get(1.0, tk.END).strip()
                if name:
                    self.risks.append({'name': name, 'prob': prob, 'impact': impact, 'score': score, 'mit': mit})
                    self.update_risks_tree()
                    self.refresh_all()
                    self.save_project()
                    dialog.destroy()
                    messagebox.showinfo("Success", "Risk added!")
            except ValueError:
                messagebox.showerror("Error", "Invalid input!")
        
        tk.Button(dialog, text="Save", command=save_risk, bg=self.colors['success'], fg='white').grid(row=4, column=1, pady=20)
    
    def update_risks_tree(self):
        for item in self.risks_tree.get_children():
            self.risks_tree.delete(item)
        for risk in self.risks:
            self.risks_tree.insert('', 'end', values=(risk['name'], risk['prob'], risk['impact'], risk['score'], risk['mit'][:50]))
    
    def update_portfolio_view(self):
        self.portfolio_text.delete(1.0, tk.END)
        info = f"Projects: 1\nTotal Tasks Across: {len(self.tasks)}\nTotal Risks: {len(self.risks)}\nTotal Documents: {len(self.documents)}\nTotal Meetings: {len(self.meeting_notes)}"
        self.portfolio_text.insert(tk.END, info)

    def add_document_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Document")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="Document Details", font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=2, pady=10)
        
        tk.Label(dialog, text="Select File:").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        file_path_var = tk.StringVar()
        tk.Entry(dialog, textvariable=file_path_var, width=40).grid(row=1, column=1, padx=5, pady=5)
        tk.Button(dialog, text="Browse...", 
                 command=lambda: file_path_var.set(filedialog.askopenfilename())).grid(row=1, column=2, padx=5)
        
        tk.Label(dialog, text="Document Name:").grid(row=2, column=0, sticky='e', padx=5, pady=5)
        name_entry = tk.Entry(dialog, width=40)
        name_entry.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky='w')
        
        tk.Label(dialog, text="Document Type:").grid(row=3, column=0, sticky='e', padx=5, pady=5)
        type_var = tk.StringVar(value="General")
        type_combo = ttk.Combobox(dialog, textvariable=type_var,
                                 values=["Contract", "Specification", "Report", "Design", "Code", "Image", "Presentation", "Other"])
        type_combo.grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky='w')
        
        tk.Label(dialog, text="Category:").grid(row=4, column=0, sticky='e', padx=5, pady=5)
        category_entry = tk.Entry(dialog, width=40)
        category_entry.grid(row=4, column=1, columnspan=2, padx=5, pady=5, sticky='w')
        
        tk.Label(dialog, text="Tags (comma-separated):").grid(row=5, column=0, sticky='e', padx=5, pady=5)
        tags_entry = tk.Entry(dialog, width=40)
        tags_entry.grid(row=5, column=1, columnspan=2, padx=5, pady=5, sticky='w')
        
        tk.Label(dialog, text="Description:").grid(row=6, column=0, sticky='ne', padx=5, pady=5)
        desc_text = tk.Text(dialog, width=40, height=4)
        desc_text.grid(row=6, column=1, columnspan=2, padx=5, pady=5)
        
        def save_document():
            file_path = file_path_var.get()
            if not file_path or not os.path.exists(file_path):
                messagebox.showerror("Error", "Please select a valid file!")
                return
            
            name = name_entry.get().strip() or os.path.basename(file_path)
            doc_type = type_var.get()
            category = category_entry.get().strip()
            tags = tags_entry.get().strip()
            description = desc_text.get(1.0, tk.END).strip()
            
            docs_folder = os.path.join(self.current_folder, 'documents')
            os.makedirs(docs_folder, exist_ok=True)
            
            dest_path = os.path.join(docs_folder, os.path.basename(file_path))
            shutil.copy2(file_path, dest_path)
            
            document = Document(name, dest_path, doc_type, category, tags, description)
            self.documents.append(document)
            
            self.update_documents_tree()
            self.save_project()
            dialog.destroy()
            messagebox.showinfo("Success", "Document added successfully!")
        
        tk.Button(dialog, text="Save", command=save_document,
                 bg=self.colors['success'], fg='white').grid(row=7, column=1, pady=20, sticky='e', padx=5)
        tk.Button(dialog, text="Cancel", command=dialog.destroy,
                 bg=self.colors['danger'], fg='white').grid(row=7, column=2, pady=20, sticky='w', padx=5)

    def view_document(self):
        selected = self.documents_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a document to view!")
            return
        
        item = self.documents_tree.item(selected[0])
        doc_name = item['values'][0]
        
        document = next((doc for doc in self.documents if doc.name == doc_name), None)
        if not document:
            messagebox.showerror("Error", "Document not found!")
            return
        
        try:
            if platform.system() == "Windows":
                os.startfile(document.file_path)
            elif platform.system() == "Darwin":
                os.system(f"open '{document.file_path}'")
            else:
                os.system(f"xdg-open '{document.file_path}'")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open document: {str(e)}")

    def delete_document(self):
        selected = self.documents_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a document to delete!")
            return
        
        item = self.documents_tree.item(selected[0])
        doc_name = item['values'][0]
        
        if messagebox.askyesno("Confirm", f"Delete document '{doc_name}'? This will remove the file from your project."):
            document = next((doc for doc in self.documents if doc.name == doc_name), None)
            if document:
                try:
                    if os.path.exists(document.file_path):
                        os.remove(document.file_path)
                    self.documents.remove(document)
                    self.update_documents_tree()
                    self.save_project()
                    messagebox.showinfo("Success", "Document deleted successfully!")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not delete document: {str(e)}")

    def filter_documents(self, event=None):
        search_term = self.doc_search_var.get().lower()
        for item in self.documents_tree.get_children():
            values = [str(v).lower() for v in self.documents_tree.item(item)['values']]
            if any(search_term in value for value in values):
                self.documents_tree.item(item, tags=('match',))
                self.documents_tree.selection_set(item)
            else:
                self.documents_tree.item(item, tags=('no_match',))
                self.documents_tree.selection_remove(item)

    def update_documents_tree(self):
        for item in self.documents_tree.get_children():
            self.documents_tree.delete(item)
        
        for doc in self.documents:
            self.documents_tree.insert('', 'end', values=(
                doc.name,
                doc.doc_type,
                doc.category,
                doc.file_size,
                doc.upload_date.strftime('%Y-%m-%d'),
                doc.tags
            ))

    def show_document_stats(self):
        if not self.documents:
            messagebox.showinfo("Document Stats", "No documents in project.")
            return
        
        stats = {}
        for doc in self.documents:
            stats[doc.doc_type] = stats.get(doc.doc_type, 0) + 1
        
        stats_text = "Document Statistics:\n\n"
        stats_text += f"Total Documents: {len(self.documents)}\n\n"
        stats_text += "By Type:\n"
        for doc_type, count in stats.items():
            stats_text += f"  {doc_type}: {count}\n"
        
        total_size = sum(os.path.getsize(doc.file_path) for doc in self.documents if os.path.exists(doc.file_path))
        stats_text += f"\nTotal Size: {self.format_file_size(total_size)}"
        
        messagebox.showinfo("Document Statistics", stats_text)

    def format_file_size(self, size_bytes):
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.2f} TB"

    def add_meeting_note_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Meeting Notes")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="Meeting Details", font=('Arial', 12, 'bold')).pack(pady=10)
        
        form_frame = tk.Frame(dialog)
        form_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        tk.Label(form_frame, text="Meeting Title:").grid(row=0, column=0, sticky='w', pady=5)
        title_entry = tk.Entry(form_frame, width=50)
        title_entry.grid(row=0, column=1, pady=5, sticky='w')
        
        tk.Label(form_frame, text="Date:").grid(row=1, column=0, sticky='w', pady=5)
        date_entry = tk.Entry(form_frame, width=50)
        date_entry.insert(0, datetime.now().strftime('%Y-%m-%d'))
        date_entry.grid(row=1, column=1, pady=5, sticky='w')
        
        tk.Label(form_frame, text="Participants:").grid(row=2, column=0, sticky='w', pady=5)
        participants_entry = tk.Entry(form_frame, width=50)
        participants_entry.grid(row=2, column=1, pady=5, sticky='w')
        
        tk.Label(form_frame, text="Agenda:").grid(row=3, column=0, sticky='nw', pady=5)
        agenda_text = tk.Text(form_frame, width=50, height=4)
        agenda_text.grid(row=3, column=1, pady=5, sticky='w')
        
        tk.Label(form_frame, text="Discussion Points:").grid(row=4, column=0, sticky='nw', pady=5)
        discussion_text = tk.Text(form_frame, width=50, height=6)
        discussion_text.grid(row=4, column=1, pady=5, sticky='w')
        
        tk.Label(form_frame, text="Action Items:").grid(row=5, column=0, sticky='nw', pady=5)
        actions_text = tk.Text(form_frame, width=50, height=4)
        actions_text.grid(row=5, column=1, pady=5, sticky='w')
        
        def save_meeting():
            meeting_data = {
                'title': title_entry.get().strip(),
                'date': date_entry.get().strip(),
                'participants': participants_entry.get().strip(),
                'agenda': agenda_text.get(1.0, tk.END).strip(),
                'discussion': discussion_text.get(1.0, tk.END).strip(),
                'actions': actions_text.get(1.0, tk.END).strip(),
                'created_date': datetime.now().isoformat()
            }
            
            if not meeting_data['title']:
                messagebox.showerror("Error", "Meeting title is required!")
                return
            
            self.meeting_notes.append(meeting_data)
            self.update_meetings_tree()
            self.save_project()
            dialog.destroy()
            messagebox.showinfo("Success", "Meeting notes saved successfully!")
        
        tk.Button(dialog, text="Save", command=save_meeting,
                 bg=self.colors['success'], fg='white').pack(pady=10)

    def view_meeting_note(self):
        selected = self.meetings_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a meeting to view!")
            return
        
        item = self.meetings_tree.item(selected[0])
        meeting_title = item['values'][1]
        
        meeting = next((m for m in self.meeting_notes if m['title'] == meeting_title), None)
        if not meeting:
            messagebox.showerror("Error", "Meeting not found!")
            return
        
        details = f"Meeting: {meeting['title']}\n"
        details += f"Date: {meeting['date']}\n"
        details += f"Participants: {meeting['participants']}\n\n"
        details += f"Agenda:\n{meeting['agenda']}\n\n"
        details += f"Discussion:\n{meeting['discussion']}\n\n"
        details += f"Action Items:\n{meeting['actions']}"
        
        messagebox.showinfo("Meeting Details", details)

    def delete_meeting_note(self):
        selected = self.meetings_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a meeting to delete!")
            return
        
        item = self.meetings_tree.item(selected[0])
        meeting_title = item['values'][1]
        
        if messagebox.askyesno("Confirm", f"Delete meeting '{meeting_title}'?"):
            meeting = next((m for m in self.meeting_notes if m['title'] == meeting_title), None)
            if meeting:
                self.meeting_notes.remove(meeting)
                self.update_meetings_tree()
                self.save_project()
                messagebox.showinfo("Success", "Meeting notes deleted!")

    def update_meetings_tree(self):
        for item in self.meetings_tree.get_children():
            self.meetings_tree.delete(item)
        
        for meeting in self.meeting_notes:
            self.meetings_tree.insert('', 'end', values=(
                meeting['date'],
                meeting['title'],
                meeting['participants'],
                meeting['actions'][:50] + "..." if len(meeting['actions']) > 50 else meeting['actions']
            ))

    def backup_project(self):
        if not self.current_folder:
            messagebox.showwarning("Warning", "No project to backup!")
            return
        
        backup_file = filedialog.asksaveasfilename(
            defaultextension=".zip",
            filetypes=[("ZIP files", "*.zip")],
            initialfile=f"{self.project_name}_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"
        )
        
        if backup_file:
            try:
                with zipfile.ZipFile(backup_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root, dirs, files in os.walk(self.current_folder):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, self.current_folder)
                            zipf.write(file_path, arcname)
                
                messagebox.showinfo("Success", f"Project backed up to:\n{backup_file}")
            except Exception as e:
                messagebox.showerror("Error", f"Backup failed: {str(e)}")

    def show_project_statistics(self):
        if not self.tasks:
            messagebox.showinfo("Project Statistics", "No project data available.")
            return
        
        stats = {
            'Total Tasks': len(self.tasks),
            'Completed Tasks': sum(1 for t in self.tasks if t.status == "Completed"),
            'In Progress': sum(1 for t in self.tasks if t.status == "In Progress"),
            'Not Started': sum(1 for t in self.tasks if t.status == "Not Started"),
            'Total Cost': sum(t.cost for t in self.tasks),
            'Milestones': sum(1 for t in self.tasks if t.is_milestone),
            'Documents': len(self.documents),
            'Meeting Notes': len(self.meeting_notes),
            'Risks': len(self.risks)
        }
        
        stats_text = "Project Statistics:\n\n"
        for key, value in stats.items():
            if key == 'Total Cost':
                stats_text += f"{key}: ${value:,.2f}\n"
            else:
                stats_text += f"{key}: {value}\n"
        
        if self.tasks:
            total_duration = sum(t.duration for t in self.tasks if not t.is_milestone)
            avg_duration = total_duration / len([t for t in self.tasks if not t.is_milestone])
            stats_text += f"\nTotal Duration: {total_duration} days\n"
            stats_text += f"Average Task Duration: {avg_duration:.1f} days"
        
        messagebox.showinfo("Project Statistics", stats_text)

    def show_document_search(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Document Search")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="Advanced Document Search", font=('Arial', 12, 'bold')).pack(pady=10)
        
        search_frame = tk.Frame(dialog)
        search_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        tk.Label(search_frame, text="Search Term:").grid(row=0, column=0, sticky='w', pady=5)
        term_entry = tk.Entry(search_frame, width=30)
        term_entry.grid(row=0, column=1, pady=5, sticky='w')
        
        tk.Label(search_frame, text="Document Type:").grid(row=1, column=0, sticky='w', pady=5)
        type_var = tk.StringVar(value="Any")
        type_combo = ttk.Combobox(search_frame, textvariable=type_var,
                                 values=["Any", "Contract", "Specification", "Report", "Design", "Code", "Image", "Presentation", "Other"])
        type_combo.grid(row=1, column=1, pady=5, sticky='w')
        
        tk.Label(search_frame, text="Category:").grid(row=2, column=0, sticky='w', pady=5)
        category_entry = tk.Entry(search_frame, width=30)
        category_entry.grid(row=2, column=1, pady=5, sticky='w')
        
        results_frame = tk.Frame(search_frame)
        results_frame.grid(row=3, column=0, columnspan=2, pady=10, sticky='we')
        
        results_text = tk.Text(results_frame, height=8, width=50)
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=results_text.yview)
        results_text.configure(yscrollcommand=scrollbar.set)
        
        results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        def perform_search():
            search_term = term_entry.get().lower()
            doc_type = type_var.get()
            category = category_entry.get().lower()
            
            results_text.delete(1.0, tk.END)
            results = []
            
            for doc in self.documents:
                matches = True
                
                if search_term and search_term not in doc.name.lower() and search_term not in doc.description.lower() and search_term not in doc.tags.lower():
                    matches = False
                
                if doc_type != "Any" and doc.doc_type != doc_type:
                    matches = False
                
                if category and category not in doc.category.lower():
                    matches = False
                
                if matches:
                    results.append(f"• {doc.name} ({doc.doc_type})\n  Category: {doc.category}\n  Tags: {doc.tags}\n")
            
            if results:
                results_text.insert(tk.END, f"Found {len(results)} documents:\n\n")
                for result in results:
                    results_text.insert(tk.END, result + "\n")
            else:
                results_text.insert(tk.END, "No documents found matching your criteria.")
        
        tk.Button(search_frame, text="Search", command=perform_search,
                 bg=self.colors['primary'], fg='white').grid(row=4, column=0, columnspan=2, pady=10)

    def create_baseline(self):
        if self.tasks:
            baseline = [Task(**t.to_dict()) for t in self.tasks]
            self.baselines.append(baseline)
            messagebox.showinfo("Success", "Baseline created!")
            self.save_project()
            self.update_gantt_chart()
        else:
            messagebox.showwarning("Warning", "No tasks to baseline!")
    
    def level_resources(self):
        if not self.tasks:
            messagebox.showwarning("Warning", "No tasks to level!")
            return
        messagebox.showinfo("Info", "Resource leveling applied (placeholder: shifts low-priority tasks by 1 day).")
        for task in self.tasks:
            if task.priority == "Low":
                task.start_date += timedelta(days=1)
        self.refresh_all()
    
    def export_to_excel(self):
        if not self.tasks:
            messagebox.showwarning("Warning", "No tasks to export!")
            return
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            try:
                df = pd.DataFrame([t.to_dict() for t in self.tasks])
                df.to_excel(filename, index=False)
                messagebox.showinfo("Success", "Exported to Excel!")
            except ImportError:
                messagebox.showerror("Error", "openpyxl not installed. Run: pip install openpyxl")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {str(e)}")
    
    def refresh_all(self):
        self.update_tasks_tree()
        self.update_dashboard()
        self.update_gantt_chart()
        self.update_timeline_view()
        self.update_kanban_view()
        self.update_resource_histogram()
        self.update_risks_tree()
        self.update_portfolio_view()
        self.update_charter_view()
        self.update_documents_tree()
        self.update_meetings_tree()
        self.update_research_tree()
    
    def open_project(self):
        folder_path = filedialog.askdirectory(title="Choose project folder", initialdir=self.base_storage_path)
        if folder_path:
            project_file = os.path.join(folder_path, 'project.json')
            if os.path.exists(project_file):
                try:
                    # Check if file is empty
                    if os.path.getsize(project_file) == 0:
                        messagebox.showerror("Error", "project.json is empty. The project file may be corrupted.")
                        return
                    
                    with open(project_file, 'r', encoding='utf-8') as f:
                        content = f.read().strip()
                        if not content:
                            messagebox.showerror("Error", "project.json is empty. The project file may be corrupted.")
                            return
                        data = json.loads(content)
                    
                    # Reset all data first
                    self.tasks = []
                    self.resources = []
                    self.risks = []
                    self.baselines = []
                    self.research_logs = []
                    self.documents = []
                    self.meeting_notes = []
                    self.images = []
                    
                    self.project_meta = data
                    self.project_name = data.get('name', os.path.basename(folder_path))
                    self.current_folder = folder_path
                    self.project_title_label.config(text=f"📊 {self.project_name}")
                    
                    # Load tasks with error handling
                    tasks_file = os.path.join(folder_path, 'tasks.json')
                    if os.path.exists(tasks_file) and os.path.getsize(tasks_file) > 0:
                        try:
                            with open(tasks_file, 'r', encoding='utf-8') as f:
                                task_data = json.load(f)
                                if task_data:
                                    self.tasks = [Task.from_dict(t) for t in task_data]
                                    for i, task in enumerate(self.tasks):
                                        task.id = i
                        except json.JSONDecodeError:
                            print(f"Warning: Could not load tasks.json")
                    
                    # Load resources with error handling
                    resources_file = os.path.join(folder_path, 'resources.json')
                    if os.path.exists(resources_file) and os.path.getsize(resources_file) > 0:
                        try:
                            with open(resources_file, 'r', encoding='utf-8') as f:
                                self.resources = json.load(f)
                        except json.JSONDecodeError:
                            print(f"Warning: Could not load resources.json")
                    
                    # Load risks with error handling
                    risks_file = os.path.join(folder_path, 'risks.json')
                    if os.path.exists(risks_file) and os.path.getsize(risks_file) > 0:
                        try:
                            with open(risks_file, 'r', encoding='utf-8') as f:
                                self.risks = json.load(f)
                        except json.JSONDecodeError:
                            print(f"Warning: Could not load risks.json")
                    
                    # Load baselines with error handling
                    baselines_file = os.path.join(folder_path, 'baselines.json')
                    if os.path.exists(baselines_file) and os.path.getsize(baselines_file) > 0:
                        try:
                            with open(baselines_file, 'r', encoding='utf-8') as f:
                                b_data = json.load(f)
                                if b_data:
                                    self.baselines = [[Task.from_dict(bt) for bt in b] for b in b_data]
                        except json.JSONDecodeError:
                            print(f"Warning: Could not load baselines.json")
                    
                    # Load research logs with error handling
                    res_file = os.path.join(folder_path, 'research.json')
                    if os.path.exists(res_file) and os.path.getsize(res_file) > 0:
                        try:
                            with open(res_file, 'r', encoding='utf-8') as f:
                                self.research_logs = json.load(f)
                        except json.JSONDecodeError:
                            print(f"Warning: Could not load research.json")
                    
                    # Load documents with error handling
                    docs_file = os.path.join(folder_path, 'documents.json')
                    if os.path.exists(docs_file) and os.path.getsize(docs_file) > 0:
                        try:
                            with open(docs_file, 'r', encoding='utf-8') as f:
                                docs_data = json.load(f)
                                if docs_data:
                                    self.documents = [Document.from_dict(d) for d in docs_data]
                        except json.JSONDecodeError:
                            print(f"Warning: Could not load documents.json")
                    
                    # Load meeting notes with error handling
                    meeting_file = os.path.join(folder_path, 'meeting_notes.json')
                    if os.path.exists(meeting_file) and os.path.getsize(meeting_file) > 0:
                        try:
                            with open(meeting_file, 'r', encoding='utf-8') as f:
                                self.meeting_notes = json.load(f)
                        except json.JSONDecodeError:
                            print(f"Warning: Could not load meeting_notes.json")
                    
                    # Load images with error handling
                    images_file = os.path.join(folder_path, 'images.json')
                    if os.path.exists(images_file) and os.path.getsize(images_file) > 0:
                        try:
                            with open(images_file, 'r', encoding='utf-8') as f:
                                self.images = json.load(f)
                        except json.JSONDecodeError:
                            print(f"Warning: Could not load images.json")
                    
                    # Refresh UI
                    self.refresh_all()
                    self.refresh_resources_list()
                    messagebox.showinfo("Success", f"Project '{self.project_name}' loaded from local storage!")
                    
                except json.JSONDecodeError as e:
                    messagebox.showerror("Error", f"Failed to load project: Invalid JSON format in project.json.\n\nDetails: {str(e)}\n\nThe file may be corrupted. Please check the file or create a new project.")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to load project: {str(e)}")
            else:
                messagebox.showerror("Error", "Invalid project folder: No project.json found.")
    
    def save_project(self):
        if not self.current_folder:
            self.save_project_as()
        else:
            self._save_to_folder()
    
    def save_project_as(self):
        folder_path = filedialog.askdirectory(title="Choose folder to save project", initialdir=self.base_storage_path)
        if folder_path:
            project_name = simpledialog.askstring("Save As", "Enter project name:")
            if project_name:
                project_folder = os.path.join(folder_path, project_name)
                os.makedirs(project_folder, exist_ok=True)
                os.makedirs(os.path.join(project_folder, 'documents'), exist_ok=True)
                os.makedirs(os.path.join(project_folder, 'images'), exist_ok=True)
                
                self.project_name = project_name
                self.project_meta['name'] = project_name
                self.current_folder = project_folder
                self.project_title_label.config(text=f"📊 {self.project_name}")
                
                self._save_to_folder()
                messagebox.showinfo("Success", f"Project saved locally as '{project_name}'!")
    
    def _save_to_folder(self):
        if not self.current_folder:
            return
        
        try:
            os.makedirs(os.path.join(self.current_folder, 'documents'), exist_ok=True)
            
            with open(os.path.join(self.current_folder, 'project.json'), 'w') as f:
                json.dump(self.project_meta, f, indent=2)
            
            with open(os.path.join(self.current_folder, 'tasks.json'), 'w') as f:
                json.dump([task.to_dict() for task in self.tasks], f, indent=2)
            
            with open(os.path.join(self.current_folder, 'resources.json'), 'w') as f:
                json.dump(self.resources, f, indent=2)
            
            with open(os.path.join(self.current_folder, 'risks.json'), 'w') as f:
                json.dump(self.risks, f, indent=2)
            
            with open(os.path.join(self.current_folder, 'baselines.json'), 'w') as f:
                json.dump([[t.to_dict() for t in b] for b in self.baselines], f, indent=2)
            
            with open(os.path.join(self.current_folder, 'research.json'), 'w') as f:
                json.dump(self.research_logs, f, indent=2)
            
            with open(os.path.join(self.current_folder, 'documents.json'), 'w') as f:
                json.dump([doc.to_dict() for doc in self.documents], f, indent=2)
            
            with open(os.path.join(self.current_folder, 'meeting_notes.json'), 'w') as f:
                json.dump(self.meeting_notes, f, indent=2)
            
            with open(os.path.join(self.current_folder, 'images.json'), 'w') as f:
                json.dump(self.images, f, indent=2)
            
            self.update_dashboard()
            print(f"Auto-saved to {self.current_folder}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save project: {str(e)}")

def main():
    root = tk.Tk()
    app = EnhancedProjectManagementApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()