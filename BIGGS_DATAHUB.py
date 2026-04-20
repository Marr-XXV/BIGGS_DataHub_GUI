"""BIGGS DataHub desktop app.

This Tkinter GUI lets users:
- Fetch branch/POS sales files for a selected date range
- Append or rewrite data into the main record2025.csv file
- View, search, and filter the latest records
- Validate data quality and patch missing records
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import calendar
import threading
import os
import csv
import time
from fetcher import Receive
try:
    from PIL import Image, ImageTk, ImageFilter, ImageDraw
    _PIL_AVAILABLE = True
except Exception:
    _PIL_AVAILABLE = False


class BiggsExtractorGUI:
    "Main GUI for fetching, combining, viewing, and validating BIGGS sales records."""
    def __init__(self, root):
        """Prepare the main window, theme, initial state, and home screen."""
        self.root = root
        self.root.title("BIGGS DataHub")
        try:
            self.root.state("zoomed")
        except Exception:
            self.root.geometry("1920x1080")
        self.root.configure(bg="#f4f6f8")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", font=("Inter", 11), padding=6)
        style.configure("Fetch.TButton", font=("Inter", 12, "bold"), background="#f9c80e", foreground="black")
        style.map("Fetch.TButton", background=[("active", "#f6b600")])
        
        style.configure("TLabel", font=("Inter", 10))
        style.configure("Header.TLabel", font=("Inter", 20, "bold"))
        style.configure("Status.TLabel", font=("Inter", 11, "italic"), background="white")
        style.configure("StatusDetail.TLabel", font=("Inter", 10, "italic"), background="white")
        style.configure("Card.TFrame", background="white")
        style.configure("Yellow.TFrame", background="#29a8e0")
        style.configure("CardHeader.TLabel", background="white", foreground="black", font=("Inter", 18, "bold"))
        style.configure("CardField.TLabel", background="white", foreground="black", font=("Inter", 9, "bold"))
        style.configure("Back.TButton", font=("Inter", 18, "bold"))
        style.configure("Calendar.TButton", font=("Inter", 9, "bold"))
        style.configure("Calendar.Nav.TButton", font=("Inter", 11, "bold"))
        style.map("Back.TButton", foreground=[("disabled", "black"), ("!disabled", "black")])
        # Radio buttons for Append/Rewrite should visually blend into the white card
        style.configure("Card.TRadiobutton", background="white", foreground="black", font=("Inter", 9))
        style.map("Card.TRadiobutton", background=[("active", "white"), ("selected", "white")], foreground=[("disabled", "gray"), ("!disabled", "black")])
        style.configure("ViewerSource.TRadiobutton", font=("Inter", 9))
        style.configure("Viewer.Small.TButton", font=("Inter", 9), padding=4)
        style.configure(
            "Fetch.Horizontal.TProgressbar",
            troughcolor="#e5e7eb",
            bordercolor="#e5e7eb",
            background="#22c55e",
            lightcolor="#22c55e",
            darkcolor="#15803d",
        )

        self.is_loading = False
        self._side_nav = None
        self._side_nav_width = 220
        self._side_nav_anim_job = None
        self.current_progress = 0.0
        self.status_text = "Idle"
        self.fetch_thread = None
        self.fetch_active = False
        self.fetch_activity_log = []
        self.cancel_fetch_flag = False
        self.receiver = None
        self.rewrite_choice = "default"
        self.rewrite_phase = "view"

        try:
            self.root.bind_all("<Button-1>", self._maybe_close_date_picker_on_click, add="+")
        except Exception:
            pass

        self.build_home_page()

    def build_home_page(self):
        """Home page: blurred background 10, header banner, rounded card with logo and actions."""
        self.header_height = 90
        self.home_container = tk.Frame(self.root, bg="#000000")
        self.home_container.pack(fill="both", expand=True)
        self.home_container.bind("<Configure>", self._update_home_bg)

        self.home_bg_label = tk.Label(self.home_container, bg="#000000")
        self.home_bg_label.place(relx=0, rely=0, relwidth=1, relheight=1)

        self.home_card = tk.Frame(self.home_container, bg="#1aa8d1")
        self.home_card.place(relx=0.5, rely=0.52, anchor="center")
        self.home_card.configure(width=460, height=480)
        self.home_card.bind("<Configure>", self._draw_home_card_bg)
        self.home_card_bg = tk.Label(self.home_card, bg="#1aa8d1")
        self.home_card_bg.place(relx=0, rely=0, relwidth=1, relheight=1)

        self.home_logo = tk.Label(self.home_card, bg="#1aa8d1")
        self.home_logo.place(relx=0.5, rely=0.26, anchor="center")
        try:
            if _PIL_AVAILABLE and os.path.exists("BIGGS1_LOGO.png"):
                img = Image.open("BIGGS1_LOGO.png")
                img.thumbnail((280, 140), Image.LANCZOS)
                self.home_logo_img = ImageTk.PhotoImage(img)
                self.home_logo.configure(image=self.home_logo_img)
            else:
                self.home_logo.configure(text="BIGGS", font=("Inter", 24, "bold"), fg="white")
        except Exception:
            self.home_logo.configure(text="BIGGS", font=("Inter", 24, "bold"), fg="white")

        pill_img_fetch = self._make_pill_button_image(200, 52, radius=26, fill="#ffffff")
        pill_img_fetch_hover = self._make_pill_button_image(200, 52, radius=26, fill="#f9c80e")
        self.home_btn_fetch = tk.Button(
            self.home_card,
            text="FETCH",
            command=self._home_go_fetch,
            image=pill_img_fetch,
            compound="center",
            bd=0,
            highlightthickness=0,
            bg="#1aa8d1",
            activebackground="#1aa8d1",
            fg="#000000",
            font=("Inter", 13, "bold"),
            cursor="hand2",
        )
        self._bind_home_button_hover(self.home_btn_fetch, pill_img_fetch, pill_img_fetch_hover)

        pill_img_validate = self._make_pill_button_image(200, 52, radius=26, fill="#ffffff")
        pill_img_validate_hover = self._make_pill_button_image(200, 52, radius=26, fill="#f9c80e")
        self.home_btn_validate = tk.Button(
            self.home_card,
            text="VALIDATE",
            command=self._home_validate,
            image=pill_img_validate,
            compound="center",
            bd=0,
            highlightthickness=0,
            bg="#1aa8d1",
            activebackground="#1aa8d1",
            fg="#000000",
            font=("Inter", 13, "bold"),
            cursor="hand2",
        )
        self._bind_home_button_hover(self.home_btn_validate, pill_img_validate, pill_img_validate_hover)

        pill_img_view = self._make_pill_button_image(200, 52, radius=26, fill="#ffffff")
        pill_img_view_hover = self._make_pill_button_image(200, 52, radius=26, fill="#f9c80e")
        self.home_btn_view = tk.Button(
            self.home_card,
            text="RECORDS",
            command=self._home_view_records,
            image=pill_img_view,
            compound="center",
            bd=0,
            highlightthickness=0,
            bg="#1aa8d1",
            activebackground="#1aa8d1",
            fg="#000000",
            font=("Inter", 13, "bold"),
            cursor="hand2",
        )
        self._bind_home_button_hover(self.home_btn_view, pill_img_view, pill_img_view_hover)

        self.home_btn_fetch.place(relx=0.5, rely=0.54, anchor="center")
        self.home_btn_validate.place(relx=0.5, rely=0.70, anchor="center")
        self.home_btn_view.place(relx=0.5, rely=0.86, anchor="center")

        self._load_home_background_image()
        self._setup_side_nav()

    def build_ui(self):
        """Compose the interface: date range card, fetch button, status, and actions."""
        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=True, padx=20, pady=10)
        main.columnconfigure(0, weight=0)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)
        self.fetch_bg_resize_job = None
        self.fetch_bg_label = tk.Label(main, bg="#f4f6f8")
        self.fetch_bg_label.grid(row=0, column=0, columnspan=2, sticky="nsew")
        main.bind("<Configure>", self._schedule_fetch_bg_update)

        card_border = ttk.Frame(main, padding=14, style="Yellow.TFrame")
        card_border.grid(row=0, column=0, padx=20, pady=10, sticky="nsew")
        self.card_border = card_border
        card = ttk.Frame(card_border, padding=20, style="Card.TFrame")
        card.pack(expand=True, fill="both")
        card_inner = ttk.Frame(card, style="Card.TFrame")
        card_inner.pack(expand=True, fill="both")
        self.card_inner = card_inner

        card_inner.columnconfigure(0, weight=1)
        card_inner.columnconfigure(3, weight=1)

        self.logo_label = tk.Label(card_inner, bg="white")
        self.logo_label.grid(row=0, column=1, columnspan=2, pady=(0, 10))
        ttk.Label(card_inner, text="INSERT DATE RANGE", font=("Inter", 12, "bold"), style="CardHeader.TLabel").grid(
            row=1, column=1, columnspan=2, pady=(0, 15)
        )

        ttk.Label(
            card_inner,
            text="Start Date (YYYY-MM-DD)",
            style="CardField.TLabel",
            font=("Inter", 9, "italic"),
        ).grid(row=2, column=1, sticky="w", padx=5)
        self._ensure_calendar_icon()
        self.start_date_container = tk.Frame(card_inner, bg="white", highlightthickness=1, highlightbackground="#d1d5db")
        self.start_date_container.grid(row=3, column=1, padx=10, pady=5, sticky="w")
        self.start_date = ttk.Entry(self.start_date_container, width=22)
        self.start_date.grid(row=0, column=0, sticky="w", padx=(6, 30), pady=2)
        self.start_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        if getattr(self, "calendar_img", None):
            self.start_date_icon = tk.Label(
                self.start_date_container,
                image=self.calendar_img,
                bg="white",
                cursor="hand2",
            )
            self.start_date_icon.place(relx=1.0, rely=0.5, anchor="e", x=-4)
            self.start_date_icon.bind("<Button-1>", lambda e: self._open_date_picker(self.start_date))

        ttk.Label(
            card_inner,
            text="End Date (YYYY-MM-DD)",
            style="CardField.TLabel",
            font=("Inter", 9, "italic"),
        ).grid(row=2, column=2, sticky="w", padx=5)
        self.end_date_container = tk.Frame(card_inner, bg="white", highlightthickness=1, highlightbackground="#d1d5db")
        self.end_date_container.grid(row=3, column=2, padx=10, pady=5, sticky="w")
        self.end_date = ttk.Entry(self.end_date_container, width=22)
        self.end_date.grid(row=0, column=0, sticky="w", padx=(6, 30), pady=2)
        self.end_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        if getattr(self, "calendar_img", None):
            self.end_date_icon = tk.Label(
                self.end_date_container,
                image=self.calendar_img,
                bg="white",
                cursor="hand2",
            )
            self.end_date_icon.place(relx=1.0, rely=0.5, anchor="e", x=-4)
            self.end_date_icon.bind("<Button-1>", lambda e: self._open_date_picker(self.end_date))

        ttk.Label(card_inner, text="Branch", style="CardField.TLabel", font=("Inter", 9, "bold")).grid(row=4, column=1, sticky="w", padx=5)
        self.branch_var = tk.StringVar()
        self.branch_entry = ttk.Entry(card_inner, width=22, textvariable=self.branch_var)
        self.branch_entry.grid(row=5, column=1, padx=10, pady=5, sticky="w")
        self.branch_var.set("ALL")
        self._branch_all_default_active = True
        self.branch_entry.bind("<KeyRelease>", self._on_branch_type)

        self.mode_var = tk.StringVar(value="append")
        ttk.Label(card_inner, text="Mode:", style="CardField.TLabel", font=("Inter", 9, "bold")).grid(row=4, column=2, sticky="w", padx=5)
        self.append_radio = ttk.Radiobutton(card_inner, text="Append", variable=self.mode_var, value="append", style="Card.TRadiobutton", takefocus=0)
        self.rewrite_radio = ttk.Radiobutton(card_inner, text="Rewrite", variable=self.mode_var, value="rewrite", style="Card.TRadiobutton", takefocus=0)
        self.append_radio.grid(row=5, column=2, sticky="w", padx=10)
        self.rewrite_radio.grid(row=6, column=2, sticky="w", padx=10)
        self.append_radio.bind("<FocusIn>", lambda e: self.root.focus_set())
        self.rewrite_radio.bind("<FocusIn>", lambda e: self.root.focus_set())
        self.mode_var.trace_add("write", self._on_mode_change)

        self.branch_suggest = tk.Listbox(card_inner, height=4)
        self.branch_suggest.bind("<<ListboxSelect>>", self._accept_branch_suggestion)
        try:
            self.branch_suggest.place_forget()
        except Exception:
            pass
        self.rewrite_dir_var = tk.StringVar()
        self.rewrite_dir_frame = ttk.Frame(card_inner, style="Card.TFrame")
        self.rewrite_dir_frame.grid(row=6, column=2, padx=10, sticky="ne")
        self.rewrite_dir_frame.grid_remove()
        self.selected_branches = []
        self.selected_branches_frame = ttk.Frame(card_inner, style="Card.TFrame", padding=(0, 8, 0, 0))
        self.selected_branches_frame.grid(row=7, column=1, sticky="nsew", padx=10, pady=(10, 0))
        branches_label = ttk.Label(
            self.selected_branches_frame,
            text="Selected Branches",
            style="CardField.TLabel",
            font=("Inter", 9, "bold"),
        )
        branches_label.grid(row=0, column=0, sticky="w", pady=(3, 3))
        self.selected_branches_list = tk.Listbox(self.selected_branches_frame, height=5, width=20)
        self.selected_branches_list.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(2, 0))
        self.selected_branches_scroll = ttk.Scrollbar(self.selected_branches_frame, orient="vertical", command=self.selected_branches_list.yview)
        self.selected_branches_scroll.grid(row=1, column=2, sticky="ns", padx=(2, 0))
        self.selected_branches_list.configure(yscrollcommand=self.selected_branches_scroll.set)
        try:
            if not hasattr(self, "branch_list"):
                self._load_branch_list()
            for b in getattr(self, "branch_list", []):
                try:
                    self.selected_branches_list.insert("end", b)
                except Exception:
                    pass
        except Exception:
            pass
        self._ensure_branch_button_images()
        add_branch_btn = tk.Label(self.selected_branches_frame, image=self.branch_add_img, bg="white", cursor="hand2")
        add_branch_btn.bind("<Button-1>", lambda e: self._add_branch())
        add_branch_btn.grid(row=0, column=1, sticky="e", padx=(8, 0))
        clear_branch_btn = tk.Label(self.selected_branches_frame, image=self.branch_remove_img, bg="white", cursor="hand2")
        clear_branch_btn.bind("<Button-1>", lambda e: self._clear_branches())
        clear_branch_btn.grid(row=0, column=2, sticky="e", padx=(4, 0))
        card_inner.rowconfigure(7, weight=1)
        self.bottom_frame = tk.Frame(card_inner, bg="white")
        self.bottom_frame.grid(row=8, column=1, columnspan=2, sticky="s")
        buttons_inner = tk.Frame(self.bottom_frame, bg="white")
        buttons_inner.pack(pady=(0, 6))
        self.fetch_main_btn = ttk.Button(
            buttons_inner,
            text="Fetch",
            command=self._on_fetch_main_button_click,
            style="Fetch.TButton",
            width=12,
        )
        self.fetch_main_btn.pack(side="left", padx=(0, 6))

        # Status + Loading Bar below date range selection
        self.status_label = ttk.Label(
            self.bottom_frame,
            text="Status: Idle",
            style="Status.TLabel"
        )
        self.status_label.pack()
        self.progress = ttk.Progressbar(
            self.bottom_frame,
            mode="determinate",
            length=250,
            maximum=100,
            style="Fetch.Horizontal.TProgressbar",
        )
        self.progress.pack(pady=(6, 0))
        try:
            if hasattr(self, "progress_detail_label"):
                self.progress_detail_label.destroy()
        except Exception:
            pass
        self.progress_detail_label = ttk.Label(
            self.bottom_frame,
            text="",
            style="StatusDetail.TLabel"
        )
        self.progress_detail_label.pack(pady=(4, 0))

        self.fetch_activity_outer = tk.Frame(main, bg="#dc1d2c")
        self.fetch_activity_outer.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        self.fetch_activity_outer.rowconfigure(1, weight=1)
        self.fetch_activity_outer.columnconfigure(0, weight=1)
        self.fetch_activity_title = tk.Label(
            self.fetch_activity_outer,
            text="REWRITE FETCHED DATA RECORD",
            font=("Inter", 11, "bold"),
            bg="#dc1d2c",
            fg="#ffffff",
        )
        self.fetch_activity_title.grid(row=0, column=0, columnspan=2, sticky="n")
        self.fetch_activity_tree = ttk.Treeview(
            self.fetch_activity_outer,
            columns=("time", "branch", "pos", "date", "detail"),
            show="headings",
        )
        self.fetch_activity_tree.heading("time", text="Time")
        self.fetch_activity_tree.heading("branch", text="Branch")
        self.fetch_activity_tree.heading("pos", text="POS")
        self.fetch_activity_tree.heading("date", text="Date")
        self.fetch_activity_tree.heading("detail", text="Details")
        self.fetch_activity_tree.column("time", width=80, anchor="w", stretch=False)
        self.fetch_activity_tree.column("branch", width=90, anchor="w", stretch=False)
        self.fetch_activity_tree.column("pos", width=50, anchor="center", stretch=False)
        self.fetch_activity_tree.column("date", width=110, anchor="w", stretch=False)
        self.fetch_activity_tree.column("detail", width=260, anchor="w", stretch=True)
        self.fetch_activity_tree.grid(row=1, column=0, sticky="nsew")
        self.fetch_activity_scroll = ttk.Scrollbar(self.fetch_activity_outer, orient="vertical", command=self.fetch_activity_tree.yview)
        self.fetch_activity_h_scroll = ttk.Scrollbar(self.fetch_activity_outer, orient="horizontal", command=self.fetch_activity_tree.xview)
        self.fetch_activity_tree.configure(
            yscrollcommand=self.fetch_activity_scroll.set,
            xscrollcommand=self.fetch_activity_h_scroll.set,
        )
        self.fetch_activity_scroll.grid(row=1, column=1, sticky="ns")
        self.fetch_activity_h_scroll.grid(row=2, column=0, sticky="ew")

        try:
            self.status_label.config(text=f"Status: {self.status_text}")
        except Exception:
            pass
        try:
            self.progress["value"] = float(getattr(self, "current_progress", 0.0))
        except Exception:
            pass
        try:
            if getattr(self, "fetch_activity_log", None):
                for entry in self.fetch_activity_log:
                    try:
                        self.fetch_activity_tree.insert("", "end", values=entry)
                    except Exception:
                        pass
                try:
                    self.fetch_activity_tree.yview_moveto(1.0)
                except Exception:
                    pass
        except Exception:
            pass

        try:
            if getattr(self, "fetch_active", False):
                try:
                    if hasattr(self, "fetch_main_btn"):
                        self.fetch_main_btn.config(text="Cancel", state="normal")
                except Exception:
                    pass
                try:
                    self._set_fetch_activity_log_layout()
                except Exception:
                    pass
        except Exception:
            pass
        try:
            self._update_fetch_activity_visibility()
        except Exception:
            pass
        try:
            self.fetch_bg_label.lower()
        except Exception:
            pass

        self.root.after(150, self._load_small_logo)

        # ================= FOOTER =================
        footer = ttk.Label(
            self.root,
            text="© BIGGS Internal System Project | Designed by NCF-CCS Intern",
            background="#f4f6f8",
            foreground="gray"
        )
        footer.pack(pady=10)
        self._setup_side_nav()
        # Initialize the fetch background image after layout is ready (no blur until fetch starts)
        try:
            self._load_fetch_background_image_async(force_blur=False)
        except Exception:
            pass

    def build_viewer_ui(self):
        """Build viewer-only layout (no fetching controls), used by 'Latest Record' navigation."""
        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=True, padx=20, pady=10)
        main.rowconfigure(2, weight=1)
        main.columnconfigure(0, weight=1)

        source_frame = ttk.Frame(main, padding=(10, 15, 10, 0))
        source_frame.grid(row=0, column=0, sticky="ew")
        source_frame.columnconfigure(3, weight=1)
        self.viewer_source_var = tk.StringVar(value="master")
        self.viewer_folder_files = {}
        self.viewer_records_dir = None
        source_label = ttk.Label(source_frame, text="Source:", font=("Inter", 9, "bold"))
        source_label.grid(row=0, column=0, sticky="w", padx=(0, 6))
        master_rb = ttk.Radiobutton(
            source_frame,
            text="Master File",
            variable=self.viewer_source_var,
            value="master",
            style="ViewerSource.TRadiobutton",
            command=self._on_viewer_source_change,
        )
        master_rb.grid(row=0, column=1, sticky="w", padx=(0, 8))
        folder_rb = ttk.Radiobutton(
            source_frame,
            text="Rewritten Records Folder",
            variable=self.viewer_source_var,
            value="folder",
            style="ViewerSource.TRadiobutton",
            command=self._on_viewer_source_change,
        )
        folder_rb.grid(row=0, column=2, sticky="w", padx=(0, 8))
        self.viewer_file_var = tk.StringVar()
        self.viewer_file_combo = ttk.Combobox(source_frame, textvariable=self.viewer_file_var, width=40, state="disabled")
        self.viewer_file_combo.grid(row=0, column=3, sticky="ew", padx=(0, 4))
        self.viewer_file_combo.bind("<<ComboboxSelected>>", self._on_viewer_file_selected)
        self.viewer_browse_btn = ttk.Button(source_frame, text="Browse", command=self._browse_record_folder, state="disabled", style="Viewer.Small.TButton")
        self.viewer_browse_btn.grid(row=0, column=4, sticky="w")
        try:
            source_label.grid_remove()
            master_rb.grid_remove()
            folder_rb.grid_remove()
            self.viewer_file_combo.grid_remove()
            self.viewer_browse_btn.grid_remove()
        except Exception:
            pass

        search_frame = ttk.Frame(main, padding=(10, 10, 10, 6))
        search_frame.grid(row=1, column=0, sticky="ew")
        search_frame.columnconfigure(7, weight=1)
        ttk.Label(search_frame, text="Branch:", font=("Inter", 9, "bold")).grid(row=0, column=0, sticky="w")
        self.viewer_branch_var = tk.StringVar()
        self.viewer_branch_combo = ttk.Combobox(search_frame, textvariable=self.viewer_branch_var, width=20, state="readonly")
        self.viewer_branch_combo.grid(row=0, column=1, sticky="w", padx=(4, 10))
        ttk.Label(search_frame, text="Start Date:", font=("Inter", 9, "bold")).grid(row=0, column=2, sticky="w")
        self.viewer_start_date_var = tk.StringVar()
        self.viewer_start_date_combo = ttk.Combobox(search_frame, textvariable=self.viewer_start_date_var, width=14, state="readonly")
        self.viewer_start_date_combo.grid(row=0, column=3, sticky="w", padx=(4, 10))
        ttk.Label(search_frame, text="End Date:", font=("Inter", 9, "bold")).grid(row=0, column=4, sticky="w")
        self.viewer_end_date_var = tk.StringVar()
        self.viewer_end_date_combo = ttk.Combobox(search_frame, textvariable=self.viewer_end_date_var, width=14, state="readonly")
        self.viewer_end_date_combo.grid(row=0, column=5, sticky="w", padx=(4, 10))
        search_btn = ttk.Button(search_frame, text="Search", command=self._apply_viewer_search, style="Viewer.Small.TButton")
        search_btn.grid(row=0, column=6, sticky="w", padx=(4, 0))
        clear_btn = ttk.Button(search_frame, text="Clear", command=self._clear_viewer_search, style="Viewer.Small.TButton")
        clear_btn.grid(row=0, column=7, sticky="w", padx=(4, 0))
        self.viewer_prev_btn = ttk.Button(search_frame, text="Prev", command=lambda: self._viewer_change_page(-1), style="Viewer.Small.TButton")
        self.viewer_prev_btn.grid(row=0, column=8, sticky="w", padx=(8, 0))
        self.viewer_page_label = ttk.Label(search_frame, text="", font=("Inter", 8))
        self.viewer_page_label.grid(row=0, column=9, sticky="w", padx=(4, 0))
        self.viewer_next_btn = ttk.Button(search_frame, text="Next", command=lambda: self._viewer_change_page(1), style="Viewer.Small.TButton")
        self.viewer_next_btn.grid(row=0, column=10, sticky="w", padx=(4, 0))
        self.viewer_outer = ttk.Frame(main, padding=10)
        self.viewer_outer.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        self.viewer_outer.rowconfigure(1, weight=1)
        self.viewer_outer.columnconfigure(0, weight=1)
        self.bg_resize_job = None
        self.viewer_bg_label = tk.Label(self.viewer_outer, bg="#f4f6f8")
        self.viewer_bg_label.grid(row=1, column=0, sticky="nsew")
        self.viewer_bg_img = None
        self.viewer_card_bg = tk.Label(self.viewer_outer, bg="#ffffff")
        self.viewer_card_bg.grid(row=1, column=0, sticky="nsew")
        try:
            self.viewer_card_bg.grid_remove()
        except Exception:
            pass
        self.viewer_tree = ttk.Treeview(self.viewer_outer)
        self.viewer_v_scroll = ttk.Scrollbar(self.viewer_outer, orient="vertical", command=self.viewer_tree.yview)
        self.viewer_h_scroll = ttk.Scrollbar(self.viewer_outer, orient="horizontal", command=self.viewer_tree.xview)
        self.viewer_tree.configure(yscrollcommand=self.viewer_v_scroll.set, xscrollcommand=self.viewer_h_scroll.set)
        self.viewer_page_size = 1000
        self._viewer_page_index = 0
        self._viewer_display_rows = []
        footer = ttk.Label(
            self.root,
            text="© BIGGS Internal System Project | Designed by NCF-CCS Intern",
            background="#f4f6f8",
            foreground="gray"
        )
        footer.pack(pady=10)
        self._setup_side_nav()
    def _append_fetch_activity(self, message, branch="", pos="", date=""):
        """Append a new activity row to the fetch monitor tree and in-memory log."""
        try:
            now = datetime.now().strftime("%H:%M:%S")
            if pos is None:
                pos_text = ""
            else:
                try:
                    pos_text = str(pos)
                except Exception:
                    pos_text = ""
            values = (now, branch or "", pos_text, date or "", message or "")
            try:
                if not hasattr(self, "fetch_activity_log") or self.fetch_activity_log is None:
                    self.fetch_activity_log = []
                self.fetch_activity_log.append(values)
            except Exception:
                pass
            tree = getattr(self, "fetch_activity_tree", None)
            if tree is None:
                return
            tree.insert("", "end", values=values)
            try:
                tree.yview_moveto(1.0)
            except Exception:
                pass
        except Exception:
            pass

    def _update_fetch_activity_visibility(self):
        """Show or hide the fetch monitor panel depending on current activity/logs."""
        try:
            outer = getattr(self, "fetch_activity_outer", None)
            if outer is None:
                return
            active = getattr(self, "fetch_active", False)
            has_log = bool(getattr(self, "fetch_activity_log", []))
            if active or has_log:
                outer.grid()
                try:
                    self._load_fetch_background_image_async(force_blur=True)
                except Exception:
                    pass
            else:
                outer.grid_remove()
        except Exception:
            pass

    def _show_master_record_in_fetch_panel(self):
        """Display the contents of record2025.csv in the fetch activity panel."""
        try:
            tree = getattr(self, "fetch_activity_tree", None)
            if tree is None:
                return
            try:
                outer = getattr(self, "fetch_activity_outer", None)
                if outer is not None:
                    outer.grid()
            except Exception:
                pass
            try:
                title = getattr(self, "fetch_activity_title", None)
                if title is not None:
                    title.config(text="FETCHED DATA RECORD")
                else:
                    pass
            except Exception:
                pass
            try:
                tree.delete(*tree.get_children())
            except Exception:
                pass
            try:
                self.fetch_activity_log = []
            except Exception:
                self.fetch_activity_log = []
            try:
                try:
                    path = os.path.join(os.getcwd(), "record2025.csv")
                except Exception:
                    path = "record2025.csv"
            except Exception:
                path = None
            try:
                if not path or not os.path.exists(path):
                    self._update_fetch_activity_visibility()
                    return
            except Exception:
                try:
                    self._update_fetch_activity_visibility()
                except Exception:
                    pass
                return
            headers = []
            rows = []
            try:
                max_rows = 200
                with open(path, newline="", encoding="utf-8") as f:
                    reader = csv.reader(f)
                    headers = next(reader, None) or []
                    for row in reader:
                        if not any(str(x).strip() for x in row):
                            continue
                        rows.append(row)
                        if len(rows) > max_rows:
                            try:
                                rows.pop(0)
                            except Exception:
                                rows = rows[1:]
            except Exception:
                headers = []
                rows = []
            try:
                try:
                    ts = os.path.getmtime(path)
                    fetch_log_value = datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    fetch_log_value = ""
            except Exception:
                fetch_log_value = ""
            if not headers:
                try:
                    self._update_fetch_activity_visibility()
                except Exception:
                    pass
                return
            try:
                idx_map = {h: i for i, h in enumerate(headers)}
                b_idx = next((idx_map[k] for k in ("BRANCH", "Branch", "branch") if k in idx_map), None)
                if b_idx is not None:
                    base_headers = ("BRANCH",) + tuple(h for h in headers if h != headers[b_idx])
                else:
                    base_headers = tuple(headers)
                display_headers = base_headers + ("FETCH LOG",)
            except Exception:
                display_headers = (tuple(headers) if headers else ()) + ("Fetch Log",)
            try:
                tree["columns"] = display_headers
                tree["show"] = "headings"
                for h in display_headers:
                    tree.heading(h, text=h)
                    tree.column(h, width=140, anchor="w", stretch=False)
            except Exception:
                pass
            try:
                self.fetch_activity_log = []
            except Exception:
                self.fetch_activity_log = []
            try:
                for row in rows:
                    if not display_headers:
                        continue
                    values = []
                    for h in display_headers[:-1]:
                        try:
                            if h == "BRANCH" and b_idx is not None:
                                i = b_idx
                            else:
                                i = idx_map.get(h, None)
                        except Exception:
                            i = None
                        if i is not None and i < len(row):
                            values.append(row[i])
                        else:
                            values.append("")
                    values.append(fetch_log_value)
                    values_tuple = tuple(values)
                    try:
                        if not hasattr(self, "fetch_activity_log") or self.fetch_activity_log is None:
                            self.fetch_activity_log = []
                        self.fetch_activity_log.append(values_tuple)
                    except Exception:
                        pass
                    try:
                        tree.insert("", "end", values=values_tuple)
                    except Exception:
                        pass
            except Exception:
                pass
            try:
                tree.yview_moveto(0.0)
            except Exception:
                pass
            try:
                self._update_fetch_activity_visibility()
            except Exception:
                pass
        except Exception:
            pass

    def _clear_fetch_panel(self):
        """Remove all rows from the fetch activity panel and refresh its visibility."""
        try:
            tree = getattr(self, "fetch_activity_tree", None)
            if tree is not None:
                try:
                    tree.delete(*tree.get_children())
                except Exception:
                    pass
        except Exception:
            pass
        try:
            self.fetch_activity_log = []
        except Exception:
            self.fetch_activity_log = []
        try:
            self._update_fetch_activity_visibility()
        except Exception:
            pass
    def _set_fetch_activity_log_layout(self):
        """Configure column headers/widths for the live fetch activity monitor."""
        try:
            title = getattr(self, "fetch_activity_title", None)
            if title is not None:
                try:
                    try:
                        mode = self.mode_var.get()
                    except Exception:
                        mode = "append"
                    text = "DATA FETCHING MONITOR"
                    if mode == "rewrite":
                        text = "REWRITING FETCHED DATA MONITOR"
                    title.config(text=text)
                except Exception:
                    pass
        except Exception:
            pass
        tree = getattr(self, "fetch_activity_tree", None)
        if tree is None:
            return
        try:
            cols = ("time", "branch", "pos", "date", "detail")
            tree["columns"] = cols
            tree["show"] = "headings"
            try:
                tree.heading("time", text="Time")
                tree.heading("branch", text="Branch")
                tree.heading("pos", text="POS")
                tree.heading("date", text="Date")
                tree.heading("detail", text="Details")
            except Exception:
                pass
            try:
                tree.column("time", width=80, anchor="w", stretch=False)
                tree.column("branch", width=90, anchor="w", stretch=False)
                tree.column("pos", width=50, anchor="center", stretch=False)
                tree.column("date", width=110, anchor="w", stretch=False)
                tree.column("detail", width=260, anchor="w", stretch=True)
            except Exception:
                pass
        except Exception:
            pass

    def cancel_fetch(self):
        """Request cancellation of the current fetch and update UI state."""
        try:
            self.cancel_fetch_flag = True
        except Exception:
            pass
        try:
            r = getattr(self, "receiver", None)
            if r is not None:
                try:
                    r.exitFlag = 1
                except Exception:
                    pass
        except Exception:
            pass
        try:
            self.set_status("Cancelled")
        except Exception:
            pass
        try:
            self.stop_loading()
        except Exception:
            pass
        try:
            self.fetch_active = False
        except Exception:
            pass
        try:
            if hasattr(self, "fetch_main_btn"):
                try:
                    mode = self.mode_var.get()
                except Exception:
                    mode = "append"
                label = "Fetch"
                if mode == "rewrite":
                    label = "View"
                    try:
                        self.rewrite_phase = "view"
                    except Exception:
                        pass
                self.fetch_main_btn.config(text=label, state="normal")
        except Exception:
            pass

    def _on_fetch_main_button_click(self):
        """Toggle between starting a fetch and cancelling an ongoing one from the main button."""
        try:
            if getattr(self, "fetch_active", False):
                self.cancel_fetch()
                return
        except Exception:
            pass
        try:
            mode = self.mode_var.get()
        except Exception:
            mode = "append"
        try:
            if mode == "rewrite":
                phase = getattr(self, "rewrite_phase", "view")
                if phase == "view":
                    try:
                        self._show_master_record_in_fetch_panel()
                    except Exception:
                        pass
                    try:
                        self.rewrite_phase = "fetch"
                    except Exception:
                        pass
                    try:
                        if hasattr(self, "fetch_main_btn"):
                            self.fetch_main_btn.config(text="Fetch", state="normal")
                    except Exception:
                        pass
                    return
                self.start_fetch()
            else:
                self.start_fetch()
        except Exception:
            pass

    def _handle_fetch_log(self, payload):
        """Normalize log payloads from the receiver and forward them to the activity monitor."""
        try:
            if isinstance(payload, str):
                self._append_fetch_activity(payload)
                return
            if not isinstance(payload, dict):
                self._append_fetch_activity(str(payload))
                return
            branch = payload.get("branch", "")
            pos = payload.get("pos", "")
            date = payload.get("date", "")
            kind = payload.get("kind", "")
            if kind == "file":
                msg = payload.get("file", "")
            else:
                msg = payload.get("message", "") or payload.get("text", "")
                if not msg and "file" in payload:
                    msg = payload["file"]
            self._append_fetch_activity(msg, branch=branch, pos=pos, date=date)
        except Exception:
            pass

    def log(self, payload):
        """Thread-safe wrapper to log fetch progress messages on the main Tk thread."""
        def cb():
            try:
                self._handle_fetch_log(payload)
            except Exception:
                pass
        try:
            self.root.after(0, cb)
        except Exception:
            pass

    def _show_validation_dialog(self, total_rows, valid_rows, invalid_rows, issues, report_path, branches_missing=None):
        try:
            dlg = tk.Toplevel(self.root)
            dlg.title("Validation Completed")
            dlg.transient(self.root)
            dlg.grab_set()
            dlg.configure(bg="#ffffff")
            container = tk.Frame(dlg, bg="#ffffff", padx=18, pady=16)
            container.pack(fill="both", expand=True)
            header = tk.Frame(container, bg="#ffffff")
            header.pack(fill="x")
            icon_canvas = tk.Canvas(header, width=32, height=32, bg="#ffffff", highlightthickness=0)
            icon_canvas.grid(row=0, column=0, sticky="w")
            icon_canvas.create_oval(2, 2, 30, 30, fill="#22c55e", outline="#22c55e")
            icon_canvas.create_text(16, 16, text="✓", fill="#ffffff", font=("Inter", 16, "bold"))
            title_label = tk.Label(header, text="Validation Completed", bg="#ffffff", fg="#111827", font=("Inter", 14, "bold"))
            title_label.grid(row=0, column=1, sticky="w", padx=(10, 0))
            summary_frame = tk.Frame(container, bg="#ffffff")
            summary_frame.pack(fill="x", pady=(16, 8))
            def add_summary_row(row, label_text, value_text):
                tk.Label(summary_frame, text="•", bg="#ffffff", fg="#111827", font=("Inter", 11, "bold")).grid(row=row, column=0, sticky="w")
                tk.Label(summary_frame, text=label_text, bg="#ffffff", fg="#111827", font=("Inter", 11)).grid(row=row, column=1, sticky="w", padx=(6, 0))
                tk.Label(summary_frame, text=value_text, bg="#ffffff", fg="#111827", font=("Inter", 11, "bold")).grid(row=row, column=2, sticky="w", padx=(4, 0))
            add_summary_row(0, "Total Rows Checked:", f"{total_rows:,}")
            add_summary_row(1, "Valid Rows:", f"{valid_rows:,}")
            add_summary_row(2, "Invalid Rows:", f"{invalid_rows:,}")
            ttk.Separator(container, orient="horizontal").pack(fill="x", pady=(4, 8))
            issues_frame = tk.Frame(container, bg="#ffffff")
            issues_frame.pack(fill="x")
            tk.Label(issues_frame, text="Issues Found:", bg="#ffffff", fg="#111827", font=("Inter", 11, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))
            icon_colors = {"warning": "#f59e0b", "error": "#ef4444", "info": "#3b82f6"}
            icon_symbols = {"warning": "!", "error": "✖", "info": "!"}
            for idx, item in enumerate(issues or []):
                kind = item.get("kind", "warning")
                text = item.get("text", "")
                highlight = item.get("highlight", "")
                row = idx + 1
                dot = tk.Canvas(issues_frame, width=20, height=20, bg="#ffffff", highlightthickness=0)
                dot.grid(row=row, column=0, sticky="w", pady=1)
                color = icon_colors.get(kind, "#f59e0b")
                symbol = icon_symbols.get(kind, "!")
                dot.create_oval(3, 3, 17, 17, fill=color, outline=color)
                dot.create_text(10, 10, text=symbol, fill="#ffffff", font=("Inter", 10, "bold"))
                tk.Label(issues_frame, text="•", bg="#ffffff", fg="#111827", font=("Inter", 10, "bold")).grid(row=row, column=1, sticky="w", padx=(2, 0))
                line_frame = tk.Frame(issues_frame, bg="#ffffff")
                line_frame.grid(row=row, column=2, sticky="w")
                tk.Label(line_frame, text=text, bg="#ffffff", fg="#111827", font=("Inter", 10)).pack(side="left")
                if highlight:
                    link = tk.Label(line_frame, text=highlight, bg="#ffffff", fg="#2563eb", font=("Inter", 10, "underline"), cursor="hand2")
                    link.pack(side="left", padx=(4, 0))
            footer = tk.Frame(container, bg="#f9fafb", padx=12, pady=10, bd=1, relief="solid")
            footer.pack(fill="x", pady=(16, 0))
            left = tk.Frame(footer, bg="#f9fafb")
            left.pack(side="left", fill="x", expand=True)
            file_icon = tk.Canvas(left, width=28, height=28, bg="#f9fafb", highlightthickness=0)
            file_icon.grid(row=0, column=0, rowspan=2, sticky="w")
            file_icon.create_rectangle(6, 4, 20, 24, fill="#0ea5e9", outline="#0ea5e9")
            file_icon.create_polygon(16, 4, 22, 10, 22, 24, 16, 24, fill="#38bdf8", outline="#38bdf8")
            title_lbl = tk.Label(left, text="View Error Report", bg="#f9fafb", fg="#111827", font=("Inter", 10, "bold"))
            title_lbl.grid(row=0, column=1, sticky="w", padx=(8, 0))
            path_text = report_path or ""
            path_lbl = tk.Label(left, text=path_text, bg="#f9fafb", fg="#374151", font=("Inter", 9))
            path_lbl.grid(row=1, column=1, sticky="w", padx=(8, 0))
            if path_text:
                try:
                    path_lbl.configure(fg="#2563eb", cursor="hand2")
                    def open_report(event=None, p=path_text):
                        try:
                            self._open_validation_report_window(p)
                        except Exception:
                            try:
                                messagebox.showerror("Validation Error", "Unable to open validation report.")
                            except Exception:
                                pass
                    path_lbl.bind("<Button-1>", open_report)
                except Exception:
                    pass
            if branches_missing:
                def start_patch():
                    try:
                        btn_patch.config(state="disabled")
                    except Exception:
                        pass
                    threading.Thread(target=self._run_patch_missing_data, args=(branches_missing, dlg), daemon=True).start()
                btn_patch = tk.Button(footer, text="Patch Missing Data", width=18, bg="#16a34a", fg="#ffffff", activebackground="#15803d", activeforeground="#ffffff", relief="raised", bd=1, font=("Inter", 10, "bold"), command=start_patch)
                btn_patch.pack(side="right", padx=(12, 0))
            btn_ok = tk.Button(footer, text="OK", width=10, bg="#2563eb", fg="#ffffff", activebackground="#1d4ed8", activeforeground="#ffffff", relief="raised", bd=1, font=("Inter", 10, "bold"), command=lambda: dlg.destroy())
            btn_ok.pack(side="right", padx=(12, 0))
            dlg.update_idletasks()
            try:
                w = dlg.winfo_width()
                h = dlg.winfo_height()
                sw = dlg.winfo_screenwidth()
                sh = dlg.winfo_screenheight()
                x = (sw - w) // 2
                y = (sh - h) // 2
                dlg.geometry(f"+{x}+{y}")
            except Exception:
                pass
        except Exception:
            try:
                messagebox.showinfo("Validation Completed", "Validation is complete.")
            except Exception:
                pass

    def _open_validation_report_window(self, path):
        try:
            p = str(path or "").strip()
        except Exception:
            p = ""
        if not p:
            try:
                messagebox.showinfo("Validation Report", "No validation report file is available.")
            except Exception:
                pass
            return
        try:
            if not os.path.exists(p):
                try:
                    messagebox.showerror("Validation Report", f"'{p}' does not exist.")
                except Exception:
                    pass
                return
        except Exception:
            try:
                messagebox.showerror("Validation Report", "Unable to locate the validation report file.")
            except Exception:
                pass
            return
        try:
            headers = []
            rows = []
            with open(p, newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                headers = next(reader, None) or []
                for row in reader:
                    try:
                        if any(str(x).strip() for x in row):
                            rows.append(row)
                    except Exception:
                        continue
        except Exception:
            try:
                messagebox.showerror("Validation Report", "Unable to read the validation report file.")
            except Exception:
                pass
            return
        try:
            dlg = tk.Toplevel(self.root)
            dlg.title("Validation Error Report")
            dlg.transient(self.root)
            dlg.grab_set()
            try:
                issue_index = None
                issue_values = []
                try:
                    for i, h in enumerate(headers):
                        try:
                            name = str(h).strip()
                        except Exception:
                            name = h
                        if str(name).upper() == "ISSUE":
                            issue_index = i
                            break
                except Exception:
                    issue_index = None
                if issue_index is not None:
                    try:
                        value_set = set()
                        for r in rows:
                            try:
                                if issue_index < len(r):
                                    v = str(r[issue_index]).strip()
                                    if v:
                                        value_set.add(v)
                            except Exception:
                                continue
                        issue_values = sorted(value_set)
                    except Exception:
                        issue_values = []
            except Exception:
                issue_index = None
                issue_values = []
            all_rows = list(rows)
            def set_window_mode():
                try:
                    dlg.attributes("-fullscreen", False)
                except Exception:
                    pass
                try:
                    dlg.state("normal")
                except Exception:
                    pass
            def set_fullscreen_mode():
                try:
                    dlg.state("normal")
                except Exception:
                    pass
                try:
                    dlg.attributes("-fullscreen", True)
                except Exception:
                    pass
            controls = tk.Frame(dlg, bg="#f9fafb", padx=8, pady=4)
            controls.pack(fill="x")
            title_lbl = tk.Label(
                controls,
                text="Validation Issue Viewer",
                bg="#f9fafb",
                fg="#111827",
                font=("Inter", 11, "bold"),
            )
            title_lbl.pack(side="left", padx=(0, 12))
            btn_window = tk.Button(
                controls,
                text="Window Mode",
                width=14,
                bg="#e5e7eb",
                fg="#111827",
                activebackground="#d1d5db",
                activeforeground="#111827",
                relief="raised",
                bd=1,
                font=("Inter", 9, "bold"),
                command=set_window_mode,
            )
            btn_full = tk.Button(
                controls,
                text="Full Screen Mode",
                width=16,
                bg="#2563eb",
                fg="#ffffff",
                activebackground="#1d4ed8",
                activeforeground="#ffffff",
                relief="raised",
                bd=1,
                font=("Inter", 9, "bold"),
                command=set_fullscreen_mode,
            )
            btn_exit = tk.Button(
                controls,
                text="Exit",
                width=12,
                bg="#dc2626",
                fg="#ffffff",
                activebackground="#b91c1c",
                activeforeground="#ffffff",
                relief="raised",
                bd=1,
                font=("Inter", 9, "bold"),
                command=dlg.destroy,
            )
            btn_exit.pack(side="right")
            btn_full.pack(side="right", padx=(0, 4))
            btn_window.pack(side="right", padx=(0, 4))
            container = tk.Frame(dlg, bg="#ffffff", padx=12, pady=10)
            container.pack(fill="both", expand=True)
            filter_frame = tk.Frame(container, bg="#ffffff")
            filter_frame.pack(fill="x", pady=(0, 6))
            inner = tk.Frame(container, bg="#ffffff")
            inner.pack(fill="both", expand=True)
            inner.rowconfigure(0, weight=1)
            inner.columnconfigure(0, weight=1)
            tree = ttk.Treeview(inner, show="headings")
            v_scroll = ttk.Scrollbar(inner, orient="vertical", command=tree.yview)
            h_scroll = ttk.Scrollbar(inner, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
            if headers:
                tree["columns"] = headers
                for h in headers:
                    try:
                        tree.heading(h, text=h)
                        tree.column(h, width=140, anchor="w", stretch=False)
                    except Exception:
                        pass
            def refresh_tree(data):
                try:
                    tree.delete(*tree.get_children())
                except Exception:
                    pass
                try:
                    for r in data:
                        try:
                            tree.insert("", "end", values=r)
                        except Exception:
                            continue
                except Exception:
                    pass
            def on_apply_filter():
                try:
                    v = issue_var.get().strip()
                except Exception:
                    v = ""
                if issue_index is None or not v or v == "All issues":
                    refresh_tree(all_rows)
                    return
                filtered = []
                try:
                    for r in all_rows:
                        try:
                            if issue_index < len(r):
                                val = str(r[issue_index]).strip()
                            else:
                                val = ""
                            if val == v:
                                filtered.append(r)
                        except Exception:
                            continue
                except Exception:
                    filtered = list(all_rows)
                refresh_tree(filtered)
            def on_clear_filter():
                try:
                    issue_var.set("All issues" if issue_index is not None and issue_values else "")
                except Exception:
                    pass
                refresh_tree(all_rows)
            tk.Label(
                filter_frame,
                text="Issue Type:",
                bg="#ffffff",
                fg="#111827",
                font=("Inter", 9, "bold"),
            ).pack(side="left")
            issue_var = tk.StringVar()
            issue_combo = ttk.Combobox(filter_frame, textvariable=issue_var, width=24, state="readonly")
            if issue_index is not None and issue_values:
                try:
                    issue_combo["values"] = ["All issues"] + issue_values
                    issue_var.set("All issues")
                except Exception:
                    pass
            else:
                try:
                    issue_combo["values"] = []
                    issue_combo.configure(state="disabled")
                except Exception:
                    pass
            issue_combo.pack(side="left", padx=(6, 4))
            btn_filter = ttk.Button(filter_frame, text="Filter", width=6, command=on_apply_filter, style="Viewer.Small.TButton")
            btn_filter.pack(side="left", padx=(0, 2))
            btn_clear = ttk.Button(filter_frame, text="Clear", width=6, command=on_clear_filter, style="Viewer.Small.TButton")
            btn_clear.pack(side="left", padx=(0, 0))
            tree.grid(row=0, column=0, sticky="nsew")
            v_scroll.grid(row=0, column=1, sticky="ns")
            h_scroll.grid(row=1, column=0, sticky="ew")
            refresh_tree(all_rows)
            set_fullscreen_mode()
            dlg.update_idletasks()
            try:
                w = dlg.winfo_width()
                h = dlg.winfo_height()
                sw = dlg.winfo_screenwidth()
                sh = dlg.winfo_screenheight()
                x = (sw - w) // 2
                y = (sh - h) // 2
                dlg.geometry(f"+{x}+{y}")
            except Exception:
                pass
        except Exception:
            try:
                messagebox.showinfo("Validation Report", f"Validation report saved at:\n{p}")
            except Exception:
                pass

    def set_status(self, message):
        """Update the status label with the given text."""
        self.status_text = message
        def cb():
            try:
                self.status_label.config(text=f"Status: {message}")
            except Exception:
                pass
        try:
            self.root.after(0, cb)
        except Exception:
            pass

    def _update_progress_percent(self, percent):
        """Update the determinate progress bar to the given percentage (0-100)."""
        try:
            value = max(0, min(100, float(percent)))
        except Exception:
            value = 0
        try:
            self.current_progress = float(value)
        except Exception:
            self.current_progress = 0.0
        try:
            self.progress["value"] = value
        except Exception:
            pass
    def _format_duration(self, seconds):
        try:
            s = int(max(0, seconds))
        except Exception:
            s = 0
        m, s = divmod(s, 60)
        h, m = divmod(m, 60)
        if h > 0:
            return f"{h}h {m}m {s}s"
        if m > 0:
            return f"{m} minutes {s} seconds"
        return f"{s} seconds"
    def _update_progress_stats(self, done, total):
        try:
            if not hasattr(self, "_progress_start_ts") or self._progress_start_ts is None:
                return
            elapsed = max(0.0001, time.perf_counter() - float(self._progress_start_ts))
            try:
                d = float(done)
                t = float(total) if total else 1.0
            except Exception:
                d = float(done) if done else 0.0
                t = 1.0
            speed = d / elapsed
            remaining = max(0.0, t - d)
            eta = int(remaining / speed) if speed > 0 else None
            speed_txt = f"{speed:.2f} records/s"
            if eta is not None:
                eta_txt = self._format_duration(eta)
                text = f"Estimated time: {eta_txt}  •  Speed: {speed_txt}"
            else:
                text = f"Speed: {speed_txt}"
            if hasattr(self, "progress_detail_label") and self.progress_detail_label:
                self.progress_detail_label.config(text=text)
        except Exception:
            try:
                if hasattr(self, "progress_detail_label") and self.progress_detail_label:
                    self.progress_detail_label.config(text="")
            except Exception:
                pass

    def start_loading(self):
        """Initialize the determinate progress bar and mark the app as loading."""
        if not self.is_loading:
            self.is_loading = True
            try:
                self.progress.configure(mode="determinate", maximum=100)
                self.progress["value"] = 0
            except Exception:
                pass
            try:
                self.current_progress = 0.0
            except Exception:
                pass
            self.set_status("In progress...")

    def stop_loading(self):
        """Mark loading as completed and fill the progress bar."""
        if self.is_loading:
            self.is_loading = False
            try:
                self.progress["value"] = 100
            except Exception:
                pass
            try:
                self.current_progress = 100.0
            except Exception:
                pass
            self.set_status("Completed")

    def _get_current_record_path(self):
        """Read settings/current_record.txt to resolve the CSV path to open; fallback to record2025.csv."""
        try:
            p = os.path.join(os.getcwd(), "settings", "current_record.txt")
            if os.path.exists(p):
                with open(p, "r", encoding="utf-8") as f:
                    path = f.read().strip()
                    if path:
                        # Normalize relative paths and require CSV extension
                        if not os.path.isabs(path):
                            candidate = os.path.join(os.getcwd(), path)
                        else:
                            candidate = path
                        if candidate.lower().endswith(".csv") and os.path.exists(candidate):
                            return candidate
        except Exception:
            pass
        return "record2025.csv"
    def _setup_side_nav(self):
        """Create slide-out sidebar (Home, Latest Record, Fetch Data) with BIGGS logo header."""
        try:
            if self._side_nav and self._side_nav.winfo_exists():
                self._side_nav.lift()
                return
            f = tk.Frame(self.root, bg="#2B2A2A")
            self._side_nav = f
            f.place(x=-self._side_nav_width, y=0, width=self._side_nav_width, height=self.root.winfo_height())
            logo = tk.Label(f, bg="#2B2A2A")
            try:
                # Try to load BIGGS_LOGO.png; fall back to text when missing or Pillow unavailable
                if _PIL_AVAILABLE and os.path.exists("BIGGS_LOGO.png"):
                    img = Image.open("BIGGS_LOGO.png")
                    img.thumbnail((self._side_nav_width - 40, 80), Image.LANCZOS)
                    self.side_logo_img = ImageTk.PhotoImage(img)
                    logo.configure(image=self.side_logo_img)
                else:
                    logo.configure(text="BIGGS", fg="#f0f0f0", font=("Inter", 14, "bold"))
            except Exception:
                logo.configure(text="BIGGS", fg="#f0f0f0", font=("Inter", 14, "bold"))
            logo.pack(padx=0, pady=(12, 8), anchor="center")
            ttk.Separator(f, orient="horizontal").pack(fill="x", padx=12, pady=(0,8))
            #    buttons
            btn_home = ttk.Button(f, text="Home", command=self.back_to_home, style="TButton")
            btn_home.pack(fill="x", padx=12, pady=(0,8))
            btn_fetch = ttk.Button(f, text="Fetch Data", command=self._home_go_fetch, style="TButton")
            btn_fetch.pack(fill="x", padx=12, pady=(0,12))
            btn_latest = ttk.Button(f, text="Records", command=self._home_view_records, style="TButton")
            btn_latest.pack(fill="x", padx=12, pady=(0,12))
            # Hover near left edge shows sidebar; leaving hides it
            f.bind("<Enter>", lambda e: self._show_side_nav())
            f.bind("<Leave>", lambda e: self._maybe_hide_side_nav())
            self.root.bind("<Configure>", lambda e: self._update_side_nav_height())
            self.root.bind("<Motion>", self._on_mouse_move)
            f.lift()
        except Exception:
            pass
    def _update_side_nav_height(self):
        """Keep the slide-out sidebar height in sync with the main window."""
        try:
            if self._side_nav and self._side_nav.winfo_exists():
                self._side_nav.place_configure(height=self.root.winfo_height())
        except Exception:
            pass
    def _on_mouse_move(self, event):
        """Detect pointer near the left edge to auto-show or hide the sidebar."""
        try:
            left = self.root.winfo_rootx()
            x_root = event.x_root
            if x_root <= left + 1:
                self._show_side_nav()
            elif x_root > left + self._side_nav_width + 40:
                self._maybe_hide_side_nav()
        except Exception:
            pass
    def _show_side_nav(self):
        """Slide the sidebar in to x=0."""
        try:
            if not (self._side_nav and self._side_nav.winfo_exists()):
                return
            self._animate_side_nav(target_x=0)
        except Exception:
            pass
    def _maybe_hide_side_nav(self):
        """Slide the sidebar out of view to negative width."""
        try:
            if not (self._side_nav and self._side_nav.winfo_exists()):
                return
            self._animate_side_nav(target_x=-self._side_nav_width)
        except Exception:
            pass
    def _animate_side_nav(self, target_x):
        """Animate sidebar movement toward target_x using small timed steps."""
        try:
            if not (self._side_nav and self._side_nav.winfo_exists()):
                return
                return
            if self._side_nav_anim_job is not None:
                try:
                    self.root.after_cancel(self._side_nav_anim_job)
                except Exception:
                    pass
                self._side_nav_anim_job = None
            info = self._side_nav.place_info()
            try:
                cur_x = int(info.get("x", 0))
            except Exception:
                cur_x = 0
            target = int(target_x)
            if cur_x == target:
                return
            step = 20 if target > cur_x else -20
            def tick():
                nonlocal cur_x
                cur_x += step
                if (step > 0 and cur_x >= target) or (step < 0 and cur_x <= target):
                    cur_x = target
                try:
                    self._side_nav.place_configure(x=cur_x)
                except Exception:
                    pass
                if cur_x == target:
                    self._side_nav_anim_job = None
                    return
                try:
                    self._side_nav_anim_job = self.root.after(12, tick)
                except Exception:
                    self._side_nav_anim_job = None
            tick()
        except Exception:
            pass
    def _resolve_corrupted_csv(self, path, header_path=None):
        """Repair a malformed CSV by backing it up and rewriting only valid rows."""
        try:
            if not path:
                return
            if not os.path.exists(path):
                return
            if os.path.getsize(path) == 0:
                return
            if header_path is None:
                try:
                    header_path = os.path.join(os.getcwd(), "aaa_headers.csv")
                except Exception:
                    header_path = None
            try:
                cleaned_rows = []
                header = None
                valid = True
                with open(path, newline="", encoding="utf-8") as f:
                    r = csv.reader(f)
                    header = next(r, None)
                    if not header:
                        valid = False
                    else:
                        expected_len = len(header)
                        for row in r:
                            if not any(str(x).strip() for x in row):
                                continue
                            if len(row) != expected_len:
                                valid = False
                                continue
                            cleaned_rows.append(row)
                if valid:
                    return
            except Exception:
                cleaned_rows = []
                header = None
            try:
                try:
                    ts = datetime.now().strftime("%Y%m%d%H%M%S")
                except Exception:
                    ts = "corrupt"
                backup = f"{path}.backup_{ts}"
                try:
                    os.replace(path, backup)
                except Exception:
                    pass
            except Exception:
                pass
            final_header = []
            try:
                if header_path and os.path.exists(header_path):
                    with open(header_path, "r", encoding="utf-8") as hf:
                        line = hf.read().strip()
                        final_header = line.split(",") if line else []
                elif header:
                    final_header = header
            except Exception:
                if header:
                    final_header = header
            try:
                with open(path, "w", newline="", encoding="utf-8") as f:
                    w = csv.writer(f)
                    if final_header:
                        w.writerow(final_header)
                    for row in cleaned_rows:
                        try:
                            w.writerow(row)
                        except Exception:
                            continue
            except Exception:
                pass
        except Exception:
            pass
    def _ensure_master_record_merge(self, previous_path=None):
        """Guarantee record2025.csv exists with header and merge rows from previous (rewrite) file."""
        try:
            master = os.path.join(os.getcwd(), "record2025.csv")
            header_path = os.path.join(os.getcwd(), "aaa_headers.csv")
            prev = previous_path or self._get_current_record_path()
            if (not os.path.exists(master)) or (os.path.getsize(master) == 0):
                header_line = ""
                try:
                    if os.path.exists(header_path):
                        with open(header_path, "r", encoding="utf-8") as hf:
                            header_line = hf.read().strip()
                    elif prev and os.path.exists(prev):
                        with open(prev, "r", encoding="utf-8") as pf:
                            header_line = pf.readline().strip()
                except Exception:
                    header_line = ""
                with open(master, "w", encoding="utf-8") as out:
                    if header_line:
                        out.write(header_line + "\n")
            if prev and os.path.exists(prev) and os.path.abspath(prev) != os.path.abspath(master):
                try:
                    with open(prev, "r", encoding="utf-8") as src, open(master, "a", encoding="utf-8") as dst:
                        first = True
                        for line in src:
                            if first:
                                first = False
                                continue
                            if line.strip():
                                dst.write(line if line.endswith("\n") else (line + "\n"))
                except Exception:
                    pass
            try:
                settings_dir = os.path.join(os.getcwd(), "settings")
                os.makedirs(settings_dir, exist_ok=True)
                with open(os.path.join(settings_dir, "current_record.txt"), "w", encoding="utf-8") as f:
                    f.write(master)
            except Exception:
                pass
        except Exception:
            pass
    def _merge_rewrite_into_master(self, previous_path=None):
        """Merge a rewritten record file into record2025.csv, replacing overlapping rows."""
        try:
            master = os.path.join(os.getcwd(), "record2025.csv")
            header_path = os.path.join(os.getcwd(), "aaa_headers.csv")
            prev = previous_path
            if not prev or not os.path.exists(prev):
                return
            try:
                self._resolve_corrupted_csv(prev, header_path)
            except Exception:
                pass
            try:
                self._resolve_corrupted_csv(master, header_path)
            except Exception:
                pass
            header = []
            master_rows = []
            if os.path.exists(master) and os.path.getsize(master) > 0:
                with open(master, newline="", encoding="utf-8") as f:
                    r = csv.reader(f)
                    header = next(r, None) or []
                    for row in r:
                        if any(str(x).strip() for x in row):
                            master_rows.append(row)
            else:
                if os.path.exists(header_path):
                    with open(header_path, "r", encoding="utf-8") as hf:
                        line = hf.read().strip()
                        header = line.split(",") if line else []
                else:
                    with open(prev, newline="", encoding="utf-8") as f:
                        r = csv.reader(f)
                        header = next(r, None) or []
            if not header:
                return
            idx_map = {h: i for i, h in enumerate(header)}
            b_idx = next((idx_map[k] for k in ("BRANCH", "Branch", "branch") if k in idx_map), None)
            d_idx = next((idx_map[k] for k in ("DATE", "Date", "date") if k in idx_map), None)
            rewrite_rows = []
            replace_pairs = set()
            with open(prev, newline="", encoding="utf-8") as f:
                r = csv.reader(f)
                _ = next(r, None)
                for row in r:
                    if not any(str(x).strip() for x in row):
                        continue
                    rewrite_rows.append(row)
                    try:
                        if b_idx is not None and d_idx is not None:
                            bv = row[b_idx] if b_idx < len(row) else ""
                            dv = row[d_idx] if d_idx < len(row) else ""
                            if str(bv).strip() and str(dv).strip():
                                replace_pairs.add((str(bv).strip(), str(dv).strip()))
                    except Exception:
                        pass
            filtered_master = []
            if b_idx is not None and d_idx is not None and replace_pairs:
                for row in master_rows:
                    try:
                        if not any(str(x).strip() for x in row):
                            continue
                        bv = row[b_idx] if b_idx < len(row) else ""
                        dv = row[d_idx] if d_idx < len(row) else ""
                        key = (str(bv).strip(), str(dv).strip())
                        if key in replace_pairs:
                            continue
                        filtered_master.append(row)
                    except Exception:
                        filtered_master.append(row)
            else:
                filtered_master = master_rows[:]
            rows = filtered_master + rewrite_rows
            with open(master, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(header)
                for row in rows:
                    w.writerow(row)
            try:
                settings_dir = os.path.join(os.getcwd(), "settings")
                os.makedirs(settings_dir, exist_ok=True)
                with open(os.path.join(settings_dir, "current_record.txt"), "w", encoding="utf-8") as f:
                    f.write(master)
            except Exception:
                pass
        except Exception:
            pass
    def open_latest_record(self):
        """Shortcut that opens the Latest Record viewer using record2025.csv."""
        try:
            self._home_view_records()
        except Exception:
            pass

    
    def back_to_idle(self):
        """Hide the table viewer and return the fetch page to its idle layout."""
        try:
            if self.viewer_tree.winfo_ismapped():
                self.viewer_tree.grid_remove()
            if self.viewer_v_scroll.winfo_ismapped():
                self.viewer_v_scroll.grid_remove()
            if self.viewer_h_scroll.winfo_ismapped():
                self.viewer_h_scroll.grid_remove()
            try:
                if self.viewer_card_bg.winfo_ismapped():
                    self.viewer_card_bg.grid_remove()
            except Exception:
                pass
            try:
                if hasattr(self, "card_border") and not self.card_border.winfo_ismapped():
                    self.card_border.grid()
            except Exception:
                pass
            self._load_background_image(blur=False)
            try:
                if hasattr(self, "back_btn") and self.back_btn.winfo_exists():
                    self.back_btn.destroy()
            except Exception:
                pass
        except Exception:
            pass
    def _apply_viewer_search(self):
        try:
            tree = getattr(self, "viewer_tree", None)
            if tree is None:
                return
            rows = getattr(self, "_viewer_all_rows", [])
            headers = getattr(self, "_viewer_headers", [])
            if not rows or not headers:
                try:
                    tree.delete(*tree.get_children())
                except Exception:
                    pass
                try:
                    self._viewer_display_rows = []
                    self._viewer_page_index = 0
                    if hasattr(self, "viewer_page_label"):
                        self.viewer_page_label.config(text="")
                    if hasattr(self, "viewer_prev_btn"):
                        self.viewer_prev_btn.configure(state="disabled")
                    if hasattr(self, "viewer_next_btn"):
                        self.viewer_next_btn.configure(state="disabled")
                except Exception:
                    pass
                return
            try:
                branch_filter = self.viewer_branch_var.get().strip()
            except Exception:
                branch_filter = ""
            try:
                start_filter = self.viewer_start_date_var.get().strip()
            except Exception:
                start_filter = ""
            try:
                end_filter = self.viewer_end_date_var.get().strip()
            except Exception:
                end_filter = ""

            overlay = self._show_viewer_loading("Searching records...")

            def worker():
                try:
                    display_rows = self._compute_filtered_display_rows(rows, headers, branch_filter, start_filter, end_filter)
                except Exception:
                    display_rows = []
                def apply():
                    try:
                        self._viewer_display_rows = list(display_rows)
                        self._viewer_page_index = 0
                        self._viewer_show_page(0)
                    except Exception:
                        pass
                    try:
                        if overlay is not None and overlay.winfo_exists():
                            overlay.destroy()
                    except Exception:
                        pass
                try:
                    self.root.after(0, apply)
                except Exception:
                    pass
            threading.Thread(target=worker, daemon=True).start()
        except Exception:
            pass
    def _clear_viewer_search(self):
        try:
            try:
                self.viewer_branch_var.set("")
            except Exception:
                pass
            try:
                self.viewer_start_date_var.set("")
            except Exception:
                pass
            try:
                self.viewer_end_date_var.set("")
            except Exception:
                pass
            self._apply_viewer_search()
        except Exception:
            pass
    def _viewer_show_page(self, index):
        try:
            tree = getattr(self, "viewer_tree", None)
        except Exception:
            tree = None
        if tree is None:
            return
        try:
            rows = list(getattr(self, "_viewer_display_rows", []))
        except Exception:
            rows = []
        try:
            size = int(getattr(self, "viewer_page_size", 1000))
        except Exception:
            size = 1000
        if size <= 0:
            size = 1000
        try:
            total = len(rows)
        except Exception:
            total = 0
        if total <= 0:
            try:
                tree.delete(*tree.get_children())
            except Exception:
                pass
            try:
                if hasattr(self, "viewer_page_label"):
                    self.viewer_page_label.config(text="")
                if hasattr(self, "viewer_prev_btn"):
                    self.viewer_prev_btn.configure(state="disabled")
                if hasattr(self, "viewer_next_btn"):
                    self.viewer_next_btn.configure(state="disabled")
            except Exception:
                pass
            return
        try:
            total_pages = (total + size - 1) // size
        except Exception:
            total_pages = 1
        if index < 0:
            index = 0
        if index >= total_pages:
            index = total_pages - 1
        try:
            self._viewer_page_index = index
        except Exception:
            self._viewer_page_index = index
        start = index * size
        end = start + size
        try:
            page_rows = rows[start:end]
        except Exception:
            page_rows = rows
        try:
            tree.delete(*tree.get_children())
        except Exception:
            pass
        try:
            for r in page_rows:
                try:
                    tree.insert("", "end", values=tuple(r))
                except Exception:
                    continue
        except Exception:
            pass
        try:
            label_text = f"Page {index + 1} of {total_pages}"
            if hasattr(self, "viewer_page_label"):
                self.viewer_page_label.config(text=label_text)
            if hasattr(self, "viewer_prev_btn"):
                self.viewer_prev_btn.configure(state="normal" if index > 0 else "disabled")
            if hasattr(self, "viewer_next_btn"):
                self.viewer_next_btn.configure(state="normal" if index + 1 < total_pages else "disabled")
        except Exception:
            pass
    def _viewer_change_page(self, delta):
        try:
            index = int(getattr(self, "_viewer_page_index", 0)) + int(delta)
        except Exception:
            index = 0
        self._viewer_show_page(index)
    
    def back_to_home(self):
        """Return from any screen back to the blurred home menu."""
        try:
            if not getattr(self, "fetch_active", False):
                try:
                    if hasattr(self, "fetch_activity_tree"):
                        self.fetch_activity_tree.delete(*self.fetch_activity_tree.get_children())
                except Exception:
                    pass
                try:
                    self.fetch_activity_log = []
                except Exception:
                    self.fetch_activity_log = []
        except Exception:
            pass
        try:
            self._show_page_loading("Loading Home...", self.build_home_page)
        except Exception:
            pass
    
    def _draw_header(self, event=None):
        """Render the top banner stripes and title text on the given canvas."""
        try:
            canvas = event.widget if event is not None else getattr(self, "home_header", getattr(self, "header_canvas", None))
            if canvas is None:
                return
            w = canvas.winfo_width()
            h = self.header_height
            canvas.delete("all")
            canvas.create_rectangle(0, 0, w, 12, fill="#19a8d1", width=0)
            canvas.create_rectangle(0, 12, w, h - 12, fill="#c72c3b", width=0)
            canvas.create_rectangle(0, h - 12, w, h, fill="#19a8d1", width=0)
            x = w // 2
            y = h // 2
            for dx, dy in ((-2, 0), (2, 0), (0, -2), (0, 2)):
                canvas.create_text(x + dx, y + dy, text="DATA FETCHER & COMBINER", fill="#6d2b8c", font=("Inter", 20, "bold"))
            canvas.create_text(x, y, text="DATA FETCHER & COMBINER", fill="#f9c80e", font=("Inter", 20, "bold"))
        except Exception:
            pass
    
    def _load_background_image(self, blur=False):
        """Load BIGGS_IMAGE.jpg behind the Latest Record viewer, optionally blurred."""
        try:
            if not _PIL_AVAILABLE:
                self.viewer_bg_label.configure(text="Install Pillow to render background image.", font=("Inter", 12, "italic"))
                return
            img_path = "BIGGS_IMAGE.jpg"
            if not os.path.exists(img_path):
                self.viewer_bg_label.configure(text="BIGGS_IMAGE.jpg not found.", font=("Inter", 12, "italic"))
                return
            img = Image.open(img_path)
            target_w = self.viewer_bg_label.winfo_width() or self.viewer_outer.winfo_width() or 900
            target_h = self.viewer_bg_label.winfo_height() or (self.viewer_outer.winfo_height() - 40 if self.viewer_outer.winfo_height() else 600)
            target_w = max(300, target_w)
            target_h = max(200, target_h)
            if blur:
                img = img.filter(ImageFilter.GaussianBlur(radius=10))
            img = img.resize((target_w, target_h), Image.LANCZOS)
            self.viewer_bg_img = ImageTk.PhotoImage(img)
            self.viewer_bg_label.configure(image=self.viewer_bg_img)
        except Exception:
            pass

    def _load_branch_list(self):
        """Load branch names from settings/branches.txt for search suggestions."""
        try:
            path = os.path.join(os.getcwd(), "settings", "branches.txt")
            items = []
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    for line in f.read().splitlines():
                        x = line.strip()
                        if x:
                            items.append(x)
            self.branch_list = items
        except Exception:
            self.branch_list = []

    def _on_branch_type(self, event=None):
        """Update branch suggestions and selected-branch list as the user types."""
        try:
            if not hasattr(self, "branch_list"):
                self._load_branch_list()
            raw = self.branch_var.get()
            term = (raw or "").strip().upper()
            try:
                if getattr(self, "_branch_all_default_active", False):
                    if term != "ALL":
                        if hasattr(self, "selected_branches_list"):
                            self.selected_branches_list.delete(0, "end")
                        self._branch_all_default_active = False
                else:
                    if term == "ALL":
                        if hasattr(self, "selected_branches_list"):
                            self.selected_branches_list.delete(0, "end")
                            for b in getattr(self, "branch_list", []):
                                self.selected_branches_list.insert("end", b)
                        self._branch_all_default_active = True
            except Exception:
                pass
            if not term:
                try:
                    self.branch_suggest.place_forget()
                except Exception:
                    pass
                return
            matches = [b for b in self.branch_list if b.upper().startswith(term)]
            self._update_branch_suggestions(matches)
        except Exception:
            pass

    def _update_branch_suggestions(self, items):
        """Update and show the dropdown listbox with up to 8 matching branches."""
        try:
            if not items:
                try:
                    self.branch_suggest.place_forget()
                except Exception:
                    pass
                return
            self.branch_suggest.delete(0, "end")
            for b in items[:8]:
                self.branch_suggest.insert("end", b)
            try:
                x = self.branch_entry.winfo_x()
                y = self.branch_entry.winfo_y() + self.branch_entry.winfo_height()
                width = self.branch_entry.winfo_width()
            except Exception:
                x, y, width = 10, 10, 180
            try:
                self.branch_suggest.place(x=x, y=y, width=width)
                self.branch_suggest.lift()
            except Exception:
                self.branch_suggest.place(x=10, y=10)
        except Exception:
            pass

    def _accept_branch_suggestion(self, event=None):
        """Apply the selected branch from the suggestion list into the entry field."""
        try:
            sel = self.branch_suggest.curselection()
            if sel:
                val = self.branch_suggest.get(sel[0])
                self.branch_var.set(val)
                self.branch_suggest.place_forget()
        except Exception:
            pass

    def _on_mode_change(self, *args):
        try:
            try:
                mode = self.mode_var.get()
            except Exception:
                mode = "append"
            try:
                frame = getattr(self, "rewrite_dir_frame", None)
                if frame is not None:
                    frame.grid_remove()
            except Exception:
                pass
            if mode == "rewrite":
                try:
                    self.rewrite_phase = "view"
                except Exception:
                    pass
                try:
                    if hasattr(self, "fetch_main_btn"):
                        self.fetch_main_btn.config(text="View", state="normal")
                except Exception:
                    pass
            else:
                try:
                    self.rewrite_phase = "view"
                except Exception:
                    pass
                try:
                    if hasattr(self, "fetch_main_btn"):
                        self.fetch_main_btn.config(text="Fetch", state="normal")
                except Exception:
                    pass
                try:
                    self._clear_fetch_panel()
                except Exception:
                    pass
                try:
                    self._set_fetch_activity_log_layout()
                except Exception:
                    pass
        except Exception:
            pass

    def _add_branch(self):
        try:
            if not hasattr(self, "selected_branches"):
                self.selected_branches = []
            name = self.branch_var.get().strip()
            if not name:
                return
            if name not in self.selected_branches:
                self.selected_branches.append(name)
                try:
                    self.selected_branches_list.insert("end", name)
                except Exception:
                    pass
                try:
                    self.branch_var.set("")
                except Exception:
                    pass
                try:
                    if hasattr(self, "branch_suggest"):
                        self.branch_suggest.place_forget()
                except Exception:
                    pass
        except Exception:
            pass

    def _clear_branches(self):
        try:
            self.selected_branches = []
            if hasattr(self, "selected_branches_list"):
                self.selected_branches_list.delete(0, "end")
            try:
                if hasattr(self, "branch_var"):
                    self.branch_var.set("")
            except Exception:
                pass
            try:
                self._branch_all_default_active = False
            except Exception:
                pass
        except Exception:
            pass

    def _browse_rewrite_folder(self):
        try:
            path = filedialog.askdirectory(title="Select folder for rewrite output")
            if path:
                self.rewrite_dir_var.set(path)
        except Exception:
            pass

    def _prompt_rewrite_folder_choice(self):
        try:
            choice = {"value": None}
            dlg = tk.Toplevel(self.root)
            dlg.title("Rewrite Options")
            dlg.transient(self.root)
            dlg.grab_set()
            msg = tk.Label(
                dlg,
                text="Choose how you want to proceed with rewrite:\n\nNew Folder – rewrite and save to a new record file.\nRewrite - Rewrite the selected date range and branches into the master record.",
                padx=20,
                pady=15
            )
            msg.pack()
            btn_frame = tk.Frame(dlg)
            btn_frame.pack(pady=(0, 15))
            def set_choice(val):
                choice["value"] = val
                try:
                    self.rewrite_choice = val
                except Exception:
                    pass
                try:
                    dlg.destroy()
                except Exception:
                    pass
            btn_default = tk.Button(btn_frame, text="New Folder", width=10, command=lambda: set_choice("default"))
            btn_default.pack(side="left", padx=8)
            btn_rewrite = tk.Button(btn_frame, text="Rewrite", width=10, command=lambda: set_choice("rewrite"))
            btn_rewrite.pack(side="left", padx=8)
            dlg.update_idletasks()
            try:
                w = dlg.winfo_width()
                h = dlg.winfo_height()
                sw = dlg.winfo_screenwidth()
                sh = dlg.winfo_screenheight()
                x = (sw - w) // 2
                y = (sh - h) // 2
                dlg.geometry(f"+{x}+{y}")
            except Exception:
                pass
            self.root.wait_window(dlg)
            if choice["value"] == "default":
                self.rewrite_dir_var.set("")
                return True
            if choice["value"] == "rewrite":
                return True
            try:
                self.rewrite_dir_var.set("")
            except Exception:
                pass
            return False
        except Exception:
            try:
                self.rewrite_dir_var.set("")
            except Exception:
                pass
            return False

    def _load_back_button_image(self, scale_down=4):
        """Load and scale the back button image from disk; fall back to text if unavailable."""
        try:
            if not _PIL_AVAILABLE:
                self.back_btn.configure(text="Back")
                return
            path = "BACKBUTTON_IMAGE.png"
            if not os.path.exists(path):
                self.back_btn.configure(text="Back")
                return
            img = Image.open(path)
            w, h = img.size
            w = max(5, w // max(1, scale_down))
            h = max(5, h // max(1, scale_down))
            img = img.resize((w, h), Image.LANCZOS)
            self.back_btn_img = ImageTk.PhotoImage(img)
            self.back_btn.configure(image=self.back_btn_img)
        except Exception:
            try:
                self.back_btn.configure(text="Back")
            except Exception:
                pass

    def _load_home_background_image(self):
        """Load the blurred background image for the home page."""
        try:
            if not _PIL_AVAILABLE:
                self.home_bg_label.configure(text="Install Pillow to render background image.", font=("Inter", 12, "italic"), fg="white", bg="#000000")
                return
            img_path = "BIGGS_IMAGE.jpg"
            if not os.path.exists(img_path):
                self.home_bg_label.configure(text="BIGGS_IMAGE.jpg not found.", font=("Inter", 12, "italic"), fg="white", bg="#000000")
                return
            img = Image.open(img_path)
            target_w = self.home_container.winfo_width() or self.root.winfo_width() or 1200
            target_h = self.home_container.winfo_height() or self.root.winfo_height() or 800
            target_w = max(300, target_w)
            target_h = max(200, target_h)
            img = img.filter(ImageFilter.GaussianBlur(radius=10))
            img = img.resize((target_w, target_h), Image.LANCZOS)
            self.home_bg_img = ImageTk.PhotoImage(img)
            self.home_bg_label.configure(image=self.home_bg_img)
        except Exception:
            pass

    def _update_home_bg(self, event=None):
        """Refresh the blurred home background on home_container resize."""
        try:
            self._load_home_background_image()
        except Exception:
            pass
    def _load_fetch_background_image_async(self, force_blur=None):
        """Async load the full-window background used on fetch screen and keep it behind content."""
        try:
            if not _PIL_AVAILABLE:
                return
            img_path = "BIGGS_IMAGE1.png"
            if not os.path.exists(img_path):
                return
            target_w = self.root.winfo_width() or 1200
            target_h = self.root.winfo_height() or 800
            target_w = max(300, target_w)
            target_h = max(200, target_h)
            try:
                if force_blur is None:
                    blur_flag = bool(getattr(self, "fetch_active", False) or getattr(self, "fetch_activity_log", []))
                else:
                    blur_flag = bool(force_blur)
            except Exception:
                blur_flag = False
            seq = getattr(self, "_fetch_bg_seq", 0) + 1
            self._fetch_bg_seq = seq
            def worker(path, w, h, seq_id, blur):
                try:
                    img = Image.open(path)
                    if blur:
                        img = img.filter(ImageFilter.GaussianBlur(radius=10))
                    img = img.resize((w, h), Image.LANCZOS)
                    def apply():
                        if getattr(self, "_fetch_bg_seq", 0) == seq_id:
                            self.fetch_bg_img = ImageTk.PhotoImage(img)
                            self.fetch_bg_label.configure(image=self.fetch_bg_img)
                        try:
                            self.fetch_bg_label.lower()
                        except Exception:
                            pass
                    self.root.after(0, apply)
                except Exception:
                    pass
            threading.Thread(target=worker, args=(img_path, target_w, target_h, seq, blur_flag), daemon=True).start()
        except Exception:
            pass
    def _schedule_fetch_bg_update(self, event=None):
        """Debounce fetch screen resizes and refresh the full-window background image."""
        try:
            if self.fetch_bg_resize_job is not None:
                self.root.after_cancel(self.fetch_bg_resize_job)
            self.fetch_bg_resize_job = self.root.after(120, self._load_fetch_background_image_async)
        except Exception:
            pass
    def _open_date_picker(self, entry):
        try:
            try:
                if hasattr(self, "_date_picker") and self._date_picker and self._date_picker.winfo_exists():
                    self._date_picker.destroy()
            except Exception:
                pass
            txt = ""
            try:
                txt = entry.get().strip()
            except Exception:
                txt = ""
            y = datetime.now().year
            m = datetime.now().month
            try:
                if txt:
                    d = datetime.strptime(txt, "%Y-%m-%d")
                    y = d.year
                    m = d.month
            except Exception:
                pass
            top = tk.Toplevel(self.root)
            self._date_picker = top
            top.overrideredirect(False)
            top.transient(self.root)
            try:
                x = entry.winfo_rootx()
                yy = entry.winfo_rooty() + entry.winfo_height()
                top.geometry(f"+{x}+{yy}")
            except Exception:
                pass
            cont = ttk.Frame(top, padding=6)
            cont.pack(fill="both", expand=True)
            hdr = ttk.Frame(cont)
            hdr.pack(fill="x")
            title_var = tk.StringVar()
            left = ttk.Frame(hdr)
            left.pack(side="left")
            right = ttk.Frame(hdr)
            right.pack(side="right")
            mid = ttk.Frame(hdr)
            mid.pack(side="left", expand=True)
            self._ensure_calendar_nav_images()
            if getattr(self, "cal_prev_img", None) and getattr(self, "cal_next_img", None):
                btn_prev = ttk.Button(left, image=self.cal_prev_img, width=3, style="Calendar.Nav.TButton")
                btn_prev.pack(padx=(0, 4))
                btn_next = ttk.Button(right, image=self.cal_next_img, width=3, style="Calendar.Nav.TButton")
                btn_next.pack(padx=(4, 0))
            else:
                btn_prev = ttk.Button(left, text="◄", width=3, style="Calendar.Nav.TButton")
                btn_prev.pack(padx=(0, 4))
                btn_next = ttk.Button(right, text="►", width=3, style="Calendar.Nav.TButton")
                btn_next.pack(padx=(4, 0))
            title = ttk.Label(mid, textvariable=title_var, font=("Inter", 9, "bold"))
            title.pack(side="left", padx=(6, 4))
            year_spin = tk.Spinbox(mid, from_=1900, to=2100, width=5)
            year_spin.pack(side="left")
            grid = ttk.Frame(cont)
            grid.pack()
            days = ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]
            for i, dname in enumerate(days):
                ttk.Label(grid, text=dname, width=3, anchor="center", font=("Inter", 9, "bold")).grid(row=0, column=i, padx=1, pady=1)
            state = {"year": y, "month": m}
            def rebuild():
                for w in grid.grid_slaves():
                    if int(w.grid_info().get("row", 0)) > 0:
                        w.destroy()
                title_var.set(f"{datetime(state['year'], state['month'], 1).strftime('%B')}")
                try:
                    year_spin.delete(0, "end")
                    year_spin.insert(0, str(state["year"]))
                except Exception:
                    pass
                cal = calendar.monthcalendar(state["year"], state["month"])
                for r, week in enumerate(cal, start=1):
                    for c, day in enumerate(week):
                        if day == 0:
                            ttk.Label(grid, text="", width=3).grid(row=r, column=c, padx=1, pady=1)
                        else:
                            def make_cmd(dd=day):
                                def cmd():
                                    try:
                                        val = f"{state['year']:04d}-{state['month']:02d}-{dd:02d}"
                                        entry.delete(0, "end")
                                        entry.insert(0, val)
                                    except Exception:
                                        pass
                                    try:
                                        top.destroy()
                                    except Exception:
                                        pass
                                return cmd
                            b = ttk.Button(grid, text=str(day), width=3, command=make_cmd(day), style="Calendar.TButton")
                            b.grid(row=r, column=c, padx=1, pady=1)
            def prev():
                mm = state["month"] - 1
                yy = state["year"]
                if mm < 1:
                    mm = 12
                    yy -= 1
                state["month"] = mm
                state["year"] = yy
                rebuild()
            def nextm():
                mm = state["month"] + 1
                yy = state["year"]
                if mm > 12:
                    mm = 1
                    yy += 1
                state["month"] = mm
                state["year"] = yy
                rebuild()
            def on_year_spin():
                try:
                    yy = int(year_spin.get())
                except Exception:
                    yy = state["year"]
                state["year"] = max(1900, min(2100, yy))
                rebuild()
            def on_mousewheel(event):
                try:
                    delta = 0
                    if hasattr(event, "delta") and event.delta:
                        delta = 1 if event.delta < 0 else -1
                    elif getattr(event, "num", None) in (4, 5):
                        delta = 1 if event.num == 5 else -1
                    yy = state["year"] + delta
                    state["year"] = max(1900, min(2100, yy))
                    rebuild()
                except Exception:
                    pass
            btn_prev.configure(command=prev)
            btn_next.configure(command=nextm)
            try:
                year_spin.configure(command=on_year_spin)
            except Exception:
                pass
            try:
                top.bind("<MouseWheel>", on_mousewheel)
                top.bind("<Button-4>", on_mousewheel)
                top.bind("<Button-5>", on_mousewheel)
            except Exception:
                pass
            rebuild()
        except Exception:
            try:
                if hasattr(self, "_date_picker") and self._date_picker and self._date_picker.winfo_exists():
                    self._date_picker.destroy()
            except Exception:
                pass

    def _maybe_close_date_picker_on_click(self, event=None):
        try:
            top = getattr(self, "_date_picker", None)
            if not top or not top.winfo_exists():
                return
            w = getattr(event, "widget", None)
            if w is None:
                try:
                    top.destroy()
                except Exception:
                    pass
                return
            safe_widgets = []
            try:
                if hasattr(self, "start_date_icon"):
                    safe_widgets.append(self.start_date_icon)
                if hasattr(self, "end_date_icon"):
                    safe_widgets.append(self.end_date_icon)
            except Exception:
                pass
            if w in safe_widgets:
                return
            try:
                if w.winfo_toplevel() is top:
                    return
            except Exception:
                pass
            try:
                top.destroy()
            except Exception:
                pass
        except Exception:
            pass

    def _show_page_loading(self, message, callback):
        try:
            for w in self.root.winfo_children():
                try:
                    w.destroy()
                except Exception:
                    pass
            container = tk.Frame(self.root, bg="#f4f6f8")
            container.pack(fill="both", expand=True)
            card = tk.Frame(container, bg="#f4f6f8")
            card.place(relx=0.5, rely=0.5, anchor="center")
            card.configure(padx=32, pady=28)
            spinner = tk.Canvas(card, width=56, height=56, bg="#f4f6f8", highlightthickness=0)
            spinner.pack()
            spinner.create_oval(6, 6, 50, 50, outline="#e5e7eb", width=4)
            arc_id = spinner.create_arc(6, 6, 50, 50, start=0, extent=120, style="arc", outline="#22c55e", width=4)
            angle = [0]
            def _spin():
                try:
                    if not spinner.winfo_exists():
                        return
                    angle[0] = (angle[0] - 36) % 360
                    spinner.itemconfig(arc_id, start=angle[0])
                    spinner.after(40, _spin)
                except Exception:
                    pass
            _spin()
            label = tk.Label(card, text=message, bg="#f4f6f8", fg="black")
            label.pack(pady=(16, 8))
            def _run():
                try:
                    for w in self.root.winfo_children():
                        try:
                            w.destroy()
                        except Exception:
                            pass
                    callback()
                except Exception:
                    pass
            self.root.after(150, _run)
        except Exception:
            try:
                callback()
            except Exception:
                pass

    def _draw_home_card_bg(self, event=None):
        """Draw the blue homepage card as a square rectangle."""
        try:
            if not _PIL_AVAILABLE:
                return
            w = self.home_card.winfo_width()
            h = self.home_card.winfo_height()
            w = max(200, w)
            h = max(200, h)
            img = Image.new("RGB", (w, h), "#1aa8d1")
            d = ImageDraw.Draw(img)
            margin = 8
            d.rectangle(
                [margin, margin, w - margin - 1, h - margin - 1],
                fill="#1aa8d1",
                outline="#ffffff",
                width=6,
            )
            self.home_card_img = ImageTk.PhotoImage(img)
            self.home_card_bg.configure(image=self.home_card_img)
        except Exception:
            pass

    def _make_pill_button_image(self, width, height, radius=26, fill="#ffffff"):
        """Generate a rectangular homepage button image."""
        try:
            if not _PIL_AVAILABLE:
                return None
            img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
            d = ImageDraw.Draw(img)
            d.rectangle([2, 2, width - 1, height - 1], fill=fill)
            return ImageTk.PhotoImage(img)
        except Exception:
            return None

    def _bind_home_button_hover(self, button, normal_img, hover_img):
        """Swap the homepage pill image on hover to make the buttons feel more alive."""
        try:
            button.image_normal = normal_img
            button.image_hover = hover_img
            button.configure(image=normal_img)

            def on_enter(event=None):
                try:
                    button.configure(image=button.image_hover)
                except Exception:
                    pass

            def on_leave(event=None):
                try:
                    button.configure(image=button.image_normal)
                except Exception:
                    pass

            button.bind("<Enter>", on_enter)
            button.bind("<Leave>", on_leave)
        except Exception:
            try:
                button.configure(image=normal_img)
            except Exception:
                pass
    def _ensure_checkbox_images(self):
        """Prepare custom checkbox images (unchecked box and green check) for Append/Rewrite options."""
        try:
            if getattr(self, "chk_unchecked_img", None) and getattr(self, "chk_checked_img", None):
                return
            if not _PIL_AVAILABLE:
                self.chk_unchecked_img = None
                self.chk_checked_img = None
                return
            w, h = 22, 18
            img0 = Image.new("RGBA", (w, h), (0, 0, 0, 0))
            d0 = ImageDraw.Draw(img0)
            d0.rounded_rectangle([1, 1, w - 2, h - 2], radius=6, fill="#ffffff", outline="#7d8790", width=2)
            self.chk_unchecked_img = ImageTk.PhotoImage(img0)
            img1 = Image.new("RGBA", (w, h), (0, 0, 0, 0))
            d1 = ImageDraw.Draw(img1)
            d1.rounded_rectangle([1, 1, w - 2, h - 2], radius=6, fill="#ffffff", outline="#7d8790", width=2)
            d1.line([(4, h // 2), (8, h - 5), (17, 4)], fill="#22c55e", width=3)
            self.chk_checked_img = ImageTk.PhotoImage(img1)
        except Exception:
            self.chk_unchecked_img = None
            self.chk_checked_img = None

    def _ensure_branch_button_images(self):
        try:
            if getattr(self, "branch_add_img", None) and getattr(self, "branch_remove_img", None):
                return
            if not _PIL_AVAILABLE:
                self.branch_add_img = None
                self.branch_remove_img = None
                return
            size = 18
            img_add = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            d_add = ImageDraw.Draw(img_add)
            d_add.ellipse([1, 1, size - 2, size - 2], fill="#22c55e")
            cx = cy = size // 2
            d_add.line([(cx - 4, cy), (cx + 4, cy)], fill="#ffffff", width=2)
            d_add.line([(cx, cy - 4), (cx, cy + 4)], fill="#ffffff", width=2)
            self.branch_add_img = ImageTk.PhotoImage(img_add)
            img_remove = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            d_rem = ImageDraw.Draw(img_remove)
            d_rem.ellipse([1, 1, size - 2, size - 2], fill="#ef4444")
            d_rem.line([(cx - 4, cy), (cx + 4, cy)], fill="#ffffff", width=2)
            self.branch_remove_img = ImageTk.PhotoImage(img_remove)
        except Exception:
            self.branch_add_img = None
            self.branch_remove_img = None
    def _ensure_calendar_nav_images(self):
        try:
            if getattr(self, "cal_prev_img", None) and getattr(self, "cal_next_img", None):
                return
            if not _PIL_AVAILABLE:
                self.cal_prev_img = None
                self.cal_next_img = None
                return
            size = 18
            img_prev = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            d_prev = ImageDraw.Draw(img_prev)
            d_prev.polygon([(12, 4), (6, size // 2), (12, size - 4)], fill="#111827")
            self.cal_prev_img = ImageTk.PhotoImage(img_prev)
            img_next = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            d_next = ImageDraw.Draw(img_next)
            d_next.polygon([(6, 4), (12, size // 2), (6, size - 4)], fill="#111827")
            self.cal_next_img = ImageTk.PhotoImage(img_next)
        except Exception:
            self.cal_prev_img = None
            self.cal_next_img = None
    def _ensure_calendar_icon(self):
        try:
            if getattr(self, "calendar_img", None):
                return
            if not _PIL_AVAILABLE:
                self.calendar_img = None
                return
            size = 18
            img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            d = ImageDraw.Draw(img)
            d.rounded_rectangle([1, 1, size - 2, size - 2], radius=3, fill="#ffffff", outline="#7d8790", width=2)
            d.rectangle([1, 1, size - 2, 6], fill="#1aa8d1")
            d.line([(4, 9), (size - 4, 9)], fill="#7d8790", width=1)
            d.line([(4, 12), (size - 4, 12)], fill="#7d8790", width=1)
            d.line([(4, 15), (size - 4, 15)], fill="#7d8790", width=1)
            self.calendar_img = ImageTk.PhotoImage(img)
        except Exception:
            self.calendar_img = None

    def _home_go_fetch(self):
        try:
            self._show_page_loading("Loading Fetch Page...", self.build_ui)
        except Exception:
            pass

    def _home_view_records(self):
        """Load record2025.csv into the viewer and show the Latest Record screen."""
        try:
            if not getattr(self, "fetch_active", False):
                try:
                    if hasattr(self, "fetch_activity_tree"):
                        self.fetch_activity_tree.delete(*self.fetch_activity_tree.get_children())
                except Exception:
                    pass
                try:
                    self.fetch_activity_log = []
                except Exception:
                    self.fetch_activity_log = []
        except Exception:
            pass
        try:
            for w in self.root.winfo_children():
                try:
                    w.destroy()
                except Exception:
                    pass
            container = tk.Frame(self.root, bg="#f4f6f8")
            container.pack(fill="both", expand=True)
            card = tk.Frame(container, bg="#f4f6f8")
            card.place(relx=0.5, rely=0.5, anchor="center")
            card.configure(padx=32, pady=28)
            spinner = tk.Canvas(card, width=56, height=56, bg="#f4f6f8", highlightthickness=0)
            spinner.pack()
            spinner.create_oval(6, 6, 50, 50, outline="#e5e7eb", width=4)
            arc_id = spinner.create_arc(6, 6, 50, 50, start=0, extent=120, style="arc", outline="#22c55e", width=4)
            angle = [0]
            def _spin():
                try:
                    if not spinner.winfo_exists():
                        return
                    angle[0] = (angle[0] - 36) % 360
                    spinner.itemconfig(arc_id, start=angle[0])
                    spinner.after(40, _spin)
                except Exception:
                    pass
            _spin()
            label = tk.Label(card, text="Loading Latest Record...", bg="#f4f6f8", fg="black")
            label.pack(pady=(16, 8))

            def worker():
                try:
                    master_path = os.path.join(os.getcwd(), "record2025.csv")
                except Exception:
                    master_path = "record2025.csv"
                if not os.path.exists(master_path):
                    try:
                        self.root.after(0, lambda: messagebox.showerror("File Not Found", f"'{master_path}' does not exist yet. Please fetch data first."))
                    except Exception:
                        pass
                    return
                try:
                    data = self._read_csv_data(master_path)
                except Exception as e:
                    try:
                        self.root.after(0, lambda: messagebox.showerror("Error", f"Could not read CSV: {e}"))
                    except Exception:
                        pass
                    return

                def apply():
                    try:
                        for w in self.root.winfo_children():
                            try:
                                w.destroy()
                            except Exception:
                                pass
                        self.build_viewer_ui()
                        try:
                            if hasattr(self, "viewer_source_var"):
                                self.viewer_source_var.set("master")
                            if hasattr(self, "viewer_file_combo"):
                                self.viewer_file_combo.configure(state="disabled")
                            if hasattr(self, "viewer_browse_btn"):
                                self.viewer_browse_btn.configure(state="disabled")
                        except Exception:
                            pass
                        try:
                            headers, rows, branches, dates = data
                            self._apply_viewer_data(headers, rows, branches, dates)
                        except Exception:
                            pass
                    except Exception:
                        pass

                try:
                    self.root.after(0, apply)
                except Exception:
                    pass

            threading.Thread(target=worker, daemon=True).start()
        except Exception:
            pass

    def _home_validate(self):
        """Ask the user for a CSV file and run the validation checks on it."""
        csv_path = ""
        try:
            try:
                if not getattr(self, "fetch_active", False):
                    try:
                        if hasattr(self, "fetch_activity_tree"):
                            self.fetch_activity_tree.delete(*self.fetch_activity_tree.get_children())
                    except Exception:
                        pass
                    try:
                        self.fetch_activity_log = []
                    except Exception:
                        self.fetch_activity_log = []
            except Exception:
                pass
            try:
                filetypes = [("CSV files", "*.csv"), ("All files", "*.*")]
                csv_path = filedialog.askopenfilename(title="Browse or select file to validate", filetypes=filetypes)
            except Exception:
                csv_path = ""
            if not csv_path:
                try:
                    messagebox.showinfo("Validate Data", "No file selected. Validation cancelled.")
                except Exception:
                    pass
                return
            threading.Thread(target=self._run_validation_async, args=(csv_path,), daemon=True).start()
        except Exception:
            try:
                messagebox.showinfo("Validate Data", "Unable to start validation.")
            except Exception:
                pass

    def _open_viewer_from_path(self, file_path):
        """Open a specific CSV file in the Latest Record viewer."""
        try:
            path = str(file_path or "").strip()
            if not path.lower().endswith(".csv"):
                path = "record2025.csv"
            if not os.path.exists(path):
                messagebox.showerror("File Not Found", f"'{path}' does not exist yet. Please fetch data first.")
                return
            with open(path, newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                headers = next(reader, None)
                if not headers:
                    messagebox.showinfo("Empty File", "The CSV file is empty.")
                    return
                rows = list(reader)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read CSV: {e}")
            return

        try:
            if hasattr(self, "viewer_bg_label"):
                self.viewer_bg_label.configure(image="", text="", bg="#f4f6f8")
                self.viewer_bg_img = None
        except Exception:
            pass
        try:
            if not self.viewer_tree.winfo_ismapped():
                self.viewer_tree.grid(row=1, column=0, sticky="nsew")
                self.viewer_v_scroll.grid(row=1, column=1, sticky="ns")
                self.viewer_h_scroll.grid(row=2, column=0, sticky="ew")
        except Exception:
            pass
        try:
            self.viewer_tree.delete(*self.viewer_tree.get_children())
        except Exception:
            pass

        idx_map = {h: i for i, h in enumerate(headers)}
        branch_idx = next((idx_map[k] for k in ("BRANCH", "Branch", "branch") if k in idx_map), None)
        if branch_idx is not None:
            display_headers = ("BRANCH",) + tuple(h for h in headers if h != headers[branch_idx])
        else:
            display_headers = ("BRANCH",) + tuple(headers)
        try:
            self.viewer_tree["columns"] = display_headers
            self.viewer_tree["show"] = "headings"
            for h in display_headers:
                self.viewer_tree.heading(h, text=h)
                self.viewer_tree.column(h, width=150, anchor="w", stretch=False)
        except Exception:
            pass

        self._viewer_all_rows = rows
        self._viewer_headers = headers
        date_idx = next((idx_map[k] for k in ("DATE", "Date", "date") if k in idx_map), None)
        branch_values = set()
        date_values = set()
        for row in rows:
            if branch_idx is not None and branch_idx < len(row):
                v = str(row[branch_idx]).strip()
                if v:
                    branch_values.add(v)
            if date_idx is not None and date_idx < len(row):
                dv = str(row[date_idx]).strip()
                if dv:
                    date_values.add(dv)
        branch_list = sorted(branch_values)
        date_list = sorted(date_values)
        try:
            if hasattr(self, "viewer_branch_combo"):
                self.viewer_branch_combo["values"] = [""] + branch_list
            if hasattr(self, "viewer_start_date_combo"):
                self.viewer_start_date_combo["values"] = [""] + date_list
            if hasattr(self, "viewer_end_date_combo"):
                self.viewer_end_date_combo["values"] = [""] + date_list
        except Exception:
            pass
        try:
            self._apply_viewer_search()
        except Exception:
            pass

    def _read_csv_data(self, path):
        """Read a CSV and return headers, rows, and distinct branch/date lists."""
        try:
            with open(path, newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                headers = next(reader, None)
                if not headers:
                    raise RuntimeError("The CSV file is empty.")
                rows = list(reader)
            idx_map = {h: i for i, h in enumerate(headers)}
            branch_idx = next((idx_map[k] for k in ("BRANCH", "Branch", "branch") if k in idx_map), None)
            date_idx = next((idx_map[k] for k in ("DATE", "Date", "date") if k in idx_map), None)
            branches = set()
            dates = set()
            for row in rows:
                if branch_idx is not None and branch_idx < len(row):
                    bv = str(row[branch_idx]).strip()
                    if bv:
                        branches.add(bv)
                if date_idx is not None and date_idx < len(row):
                    dv = str(row[date_idx]).strip()
                    if dv:
                        dates.add(dv)
            return headers, rows, sorted(branches), sorted(dates)
        except Exception as e:
            raise RuntimeError(str(e))

    def _apply_viewer_data(self, headers, rows, branch_list, date_list):
        """Populate the Latest Record viewer using already-loaded CSV data."""
        try:
            if hasattr(self, "viewer_bg_label"):
                self.viewer_bg_label.configure(image="", text="", bg="#f4f6f8")
                self.viewer_bg_img = None
        except Exception:
            pass
        try:
            if not self.viewer_tree.winfo_ismapped():
                self.viewer_tree.grid(row=1, column=0, sticky="nsew")
                self.viewer_v_scroll.grid(row=1, column=1, sticky="ns")
                self.viewer_h_scroll.grid(row=2, column=0, sticky="ew")
        except Exception:
            pass
        try:
            self.viewer_tree.delete(*self.viewer_tree.get_children())
        except Exception:
            pass
        idx_map = {h: i for i, h in enumerate(headers)}
        branch_idx = next((idx_map[k] for k in ("BRANCH", "Branch", "branch") if k in idx_map), None)
        if branch_idx is not None:
            display_headers = ("BRANCH",) + tuple(h for h in headers if h != headers[branch_idx])
        else:
            display_headers = ("BRANCH",) + tuple(headers)
        try:
            self.viewer_tree["columns"] = display_headers
            self.viewer_tree["show"] = "headings"
            for h in display_headers:
                self.viewer_tree.heading(h, text=h)
                self.viewer_tree.column(h, width=150, anchor="w", stretch=False)
        except Exception:
            pass
        self._viewer_all_rows = rows
        self._viewer_headers = headers
        try:
            if hasattr(self, "viewer_branch_combo"):
                self.viewer_branch_combo["values"] = [""] + list(branch_list)
            if hasattr(self, "viewer_start_date_combo"):
                self.viewer_start_date_combo["values"] = [""] + list(date_list)
            if hasattr(self, "viewer_end_date_combo"):
                self.viewer_end_date_combo["values"] = [""] + list(date_list)
        except Exception:
            pass
        try:
            self._apply_viewer_search()
        except Exception:
            pass

    def _compute_filtered_display_rows(self, rows, headers, branch_filter, start_filter, end_filter):
        """Filter and sort raw CSV rows by branch and date for the viewer grid."""
        idx_map = {h: i for i, h in enumerate(headers)}
        branch_idx = next((idx_map[k] for k in ("BRANCH", "Branch", "branch") if k in idx_map), None)
        date_idx = next((idx_map[k] for k in ("DATE", "Date", "date") if k in idx_map), None)
        time_idx = next((idx_map[k] for k in ("TIME", "Time", "time", "TIME24") if k in idx_map), None)
        if not branch_filter and not start_filter and not end_filter:
            display_rows = []
            for row in rows:
                try:
                    if branch_idx is not None and branch_idx < len(row):
                        display_row = [row[branch_idx]] + [row[i] for i in range(len(row)) if i != branch_idx]
                    else:
                        display_row = [""] + row
                    display_rows.append(display_row)
                except Exception:
                    continue
            return display_rows
        filtered = []
        for row in rows:
            try:
                if branch_filter and branch_idx is not None and branch_idx < len(row):
                    bval = str(row[branch_idx]).strip()
                    if bval != branch_filter:
                        continue
                elif branch_filter:
                    continue
                if (start_filter or end_filter) and date_idx is not None and date_idx < len(row):
                    dval = str(row[date_idx]).strip()
                    if start_filter and dval < start_filter:
                        continue
                    if end_filter and dval > end_filter:
                        continue
                elif (start_filter or end_filter):
                    continue
                filtered.append(row)
            except Exception:
                continue
        def parse_dt(row):
            try:
                if date_idx is None:
                    return datetime.min
                d = row[date_idx] if date_idx < len(row) else ""
                t = row[time_idx] if (time_idx is not None and time_idx < len(row)) else ""
                d = (d or "").strip()
                t = (t or "").strip()
                if not d:
                    return datetime.min
                candidates = []
                if t:
                    candidates.append((f"{d} {t}", "%Y-%m-%d %H:%M:%S"))
                    candidates.append((f"{d} {t}", "%Y-%m-%d %H:%M"))
                    t_digits = t.replace(":", "")
                    candidates.append((f"{d} {t_digits}", "%Y-%m-%d %H%M%S"))
                    candidates.append((f"{d} {t_digits}", "%Y-%m-%d %H%M"))
                else:
                    candidates.append((d, "%Y-%m-%d"))
                    candidates.append((d, "%Y/%m/%d"))
                    candidates.append((d, "%m/%d/%Y"))
                    candidates.append((d, "%Y%m%d"))
                for value, fmt in candidates:
                    try:
                        return datetime.strptime(value, fmt)
                    except Exception:
                        continue
                return datetime.min
            except Exception:
                return datetime.min
        if date_idx is not None:
            sorted_rows = sorted(filtered, key=parse_dt, reverse=True)
        else:
            sorted_rows = list(reversed(filtered))
        display_rows = []
        for row in sorted_rows:
            try:
                if branch_idx is not None and branch_idx < len(row):
                    display_row = [row[branch_idx]] + [row[i] for i in range(len(row)) if i != branch_idx]
                else:
                    display_row = [""] + row
                display_rows.append(display_row)
            except Exception:
                continue
        return display_rows

    def _show_viewer_loading(self, text="Searching records..."):
        try:
            try:
                if getattr(self, "_viewer_loading_overlay", None) and self._viewer_loading_overlay.winfo_exists():
                    self._viewer_loading_overlay.destroy()
            except Exception:
                pass
            outer = getattr(self, "viewer_outer", None) or self.root
            ov = tk.Frame(outer, bg="#f4f6f8")
            try:
                ov.place(relx=0, rely=0, relwidth=1, relheight=1)
                ov.lift()
            except Exception:
                pass
            card = tk.Frame(ov, bg="#f4f6f8")
            try:
                card.place(relx=0.5, rely=0.5, anchor="center")
            except Exception:
                card.pack(padx=32, pady=28)
            card.configure(padx=32, pady=28)
            spinner = tk.Canvas(card, width=56, height=56, bg="#f4f6f8", highlightthickness=0)
            spinner.pack()
            spinner.create_oval(6, 6, 50, 50, outline="#e5e7eb", width=4)
            arc_id = spinner.create_arc(6, 6, 50, 50, start=0, extent=120, style="arc", outline="#22c55e", width=4)
            angle = [0]
            def _spin():
                try:
                    if not spinner.winfo_exists():
                        return
                    angle[0] = (angle[0] - 36) % 360
                    spinner.itemconfig(arc_id, start=angle[0])
                    spinner.after(40, _spin)
                except Exception:
                    pass
            _spin()
            label = tk.Label(card, text=text, bg="#f4f6f8", fg="black")
            label.pack(pady=(16, 8))
            self._viewer_loading_overlay = ov
            return ov
        except Exception:
            return None
    def _show_loading_then_load_csv(self, path, mode="master", display_name=None):
        """Show a spinner while reading a CSV, then rebuild viewer with the new data."""
        try:
            for w in self.root.winfo_children():
                try:
                    w.destroy()
                except Exception:
                    pass
            container = tk.Frame(self.root, bg="#f4f6f8")
            container.pack(fill="both", expand=True)
            card = tk.Frame(container, bg="#f4f6f8")
            card.place(relx=0.5, rely=0.5, anchor="center")
            card.configure(padx=32, pady=28)
            spinner = tk.Canvas(card, width=56, height=56, bg="#f4f6f8", highlightthickness=0)
            spinner.pack()
            spinner.create_oval(6, 6, 50, 50, outline="#e5e7eb", width=4)
            arc_id = spinner.create_arc(6, 6, 50, 50, start=0, extent=120, style="arc", outline="#22c55e", width=4)
            angle = [0]
            def _spin():
                try:
                    if not spinner.winfo_exists():
                        return
                    angle[0] = (angle[0] - 36) % 360
                    spinner.itemconfig(arc_id, start=angle[0])
                    spinner.after(40, _spin)
                except Exception:
                    pass
            _spin()
            label = tk.Label(card, text="Loading Latest Record...", bg="#f4f6f8", fg="black")
            label.pack(pady=(16, 8))
        except Exception:
            pass

        def worker():
            try:
                if not (path and str(path).lower().endswith(".csv") and os.path.exists(path)):
                    self.root.after(0, lambda: messagebox.showerror("File Not Found", f"'{path}' does not exist."))
                    return
                data = self._read_csv_data(path)
            except Exception as e:
                try:
                    self.root.after(0, lambda: messagebox.showerror("Error", f"Could not read CSV: {e}"))
                except Exception:
                    pass
                return
            def apply():
                try:
                    for w in self.root.winfo_children():
                        try:
                            w.destroy()
                        except Exception:
                            pass
                    self.build_viewer_ui()
                    try:
                        if hasattr(self, "viewer_source_var"):
                            self.viewer_source_var.set("master" if mode != "folder" else "folder")
                        if hasattr(self, "viewer_file_combo") and hasattr(self, "viewer_browse_btn"):
                            if mode == "folder":
                                name = display_name or os.path.basename(path)
                                self.viewer_folder_files = {name: path}
                                self.viewer_file_combo["values"] = [name]
                                self.viewer_file_var.set(name)
                                self.viewer_file_combo.configure(state="readonly")
                                self.viewer_browse_btn.configure(state="normal")
                            else:
                                self.viewer_file_combo.configure(state="disabled")
                                self.viewer_browse_btn.configure(state="disabled")
                    except Exception:
                        pass
                    try:
                        headers, rows, branches, dates = data
                        self._apply_viewer_data(headers, rows, branches, dates)
                    except Exception:
                        pass
                except Exception:
                    pass
            try:
                self.root.after(0, apply)
            except Exception:
                pass
        threading.Thread(target=worker, daemon=True).start()

    def _run_validation_async(self, csv_path=None):
        """Run data-quality checks on a CSV file and save a detailed error report."""
        try:
            import pandas as pd
            from datetime import datetime
        except Exception:
            try:
                self.root.after(0, lambda: messagebox.showerror("Validation Error", "Pandas library is required for validation."))
            except Exception:
                pass
            return
        try:
            try:
                if not csv_path:
                    csv_path = self._get_current_record_path()
            except Exception:
                csv_path = "record2025.csv"
            if not os.path.exists(csv_path):
                self.root.after(0, lambda: messagebox.showerror("Validation Error", f"'{csv_path}' not found."))
                return
            df = pd.read_csv(csv_path)
            total_rows = int(len(df.index))
            required_cols = ["DATE", "BRANCH", "POS", "QUANTITY", "AMOUNT"]
            missing_columns = 0
            issue_rows = []
            try:
                missing_cols = [c for c in required_cols if c not in df.columns]
                if missing_cols:
                    missing_columns = 1
                    for col in missing_cols:
                        try:
                            issue_rows.append({"ISSUE": "Missing Column", "COLUMN": col})
                        except Exception:
                            pass
            except Exception:
                missing_columns = 0
            invalid_branch_rows = 0
            try:
                branches_file = os.path.join(os.getcwd(), "settings", "branches.txt")
                branches = []
                if os.path.exists(branches_file):
                    with open(branches_file, "r", encoding="utf-8") as f:
                        for row in f.read().splitlines():
                            br = (row or "").strip()
                            if br:
                                branches.append(br)
                if branches and "BRANCH" in df.columns:
                    invalid_mask = ~df["BRANCH"].isin(branches)
                    invalid_branch_rows = int(invalid_mask.sum())
                    if invalid_branch_rows:
                        try:
                            invalid_df = df.loc[invalid_mask].copy()
                            try:
                                invalid_df.insert(0, "ISSUE", "Invalid Branch")
                            except Exception:
                                invalid_df["ISSUE"] = "Invalid Branch"
                            try:
                                issue_rows.extend(invalid_df.to_dict(orient="records"))
                            except Exception:
                                pass
                        except Exception:
                            pass
            except Exception:
                invalid_branch_rows = 0
            negative_rows = 0
            try:
                if total_rows:
                    neg_mask = pd.Series(False, index=df.index)
                    if "QUANTITY" in df.columns:
                        neg_mask |= df["QUANTITY"] < 0
                    if "AMOUNT" in df.columns:
                        neg_mask |= df["AMOUNT"] < 0
                    negative_rows = int(neg_mask.sum())
                    if negative_rows:
                        try:
                            neg_df = df.loc[neg_mask].copy()
                            try:
                                neg_df.insert(0, "ISSUE", "Negative Sales Value")
                            except Exception:
                                neg_df["ISSUE"] = "Negative Sales Value"
                            try:
                                issue_rows.extend(neg_df.to_dict(orient="records"))
                            except Exception:
                                pass
                        except Exception:
                            pass
            except Exception:
                negative_rows = 0
            missing_date_issues = 0
            report_df = None
            branches_missing = {}
            try:
                if {"DATE", "BRANCH", "POS", "QUANTITY"}.issubset(set(df.columns)):
                    df_dates = df.copy()
                    df_dates["DATE"] = pd.to_datetime(df_dates["DATE"], errors="coerce")
                    df_dates = df_dates.dropna(subset=["DATE"])
                    if not df_dates.empty:
                        date_start = df_dates["DATE"].min().normalize()
                        date_end = df_dates["DATE"].max().normalize()
                        dates = df_dates.pivot_table(
                            index="DATE",
                            columns=["BRANCH", "POS"],
                            values="QUANTITY",
                            aggfunc="sum",
                            fill_value=0,
                            dropna=False,
                        )
                        branches = []
                        try:
                            branches_file = os.path.join(os.getcwd(), "settings", "branches.txt")
                            if os.path.exists(branches_file):
                                with open(branches_file, "r", encoding="utf-8") as f:
                                    for row in f.read().splitlines():
                                        br = (row or "").strip()
                                        if br:
                                            branches.append(br)
                        except Exception:
                            branches = []
                        if not branches:
                            try:
                                branches = sorted(df_dates["BRANCH"].dropna().unique().tolist())
                            except Exception:
                                branches = []
                        expected = set()
                        current_date = date_start
                        while current_date <= date_end:
                            for br in branches:
                                for pos in ["1", "2"]:
                                    expected.add((br, str(pos), current_date.strftime("%Y-%m-%d")))
                            current_date += timedelta(days=1)
                        existing = set()
                        for d in dates.index:
                            date_str = d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d).split()[0]
                            for (_, br, pos) in dates.columns:
                                if dates.loc[d, ("QUANTITY", br, pos)] > 0:
                                    existing.add((br, str(pos), date_str))
                        missing = expected - existing
                        missing_date_issues = len(missing)
                        rows = []
                        for br, pos, dt in sorted(missing):
                            rows.append({"BRANCH": br, "POS": pos, "DATE": dt, "ISSUE": "Missing data"})
                            pos_key = pos
                            try:
                                if isinstance(pos, str) and pos.isdigit():
                                    pos_key = int(pos)
                            except Exception:
                                pos_key = pos
                            if br not in branches_missing:
                                branches_missing[br] = {}
                            if pos_key not in branches_missing[br]:
                                branches_missing[br][pos_key] = []
                            branches_missing[br][pos_key].append(dt)
                        if rows:
                            try:
                                report_df = pd.DataFrame(rows)
                            except Exception:
                                report_df = None
                            try:
                                issue_rows.extend(rows)
                            except Exception:
                                pass
            except Exception:
                missing_date_issues = 0
                report_df = None
            invalid_rows = invalid_branch_rows + negative_rows
            valid_rows = max(0, total_rows - invalid_rows)
            report_path = ""
            try:
                try:
                    base_dir = os.path.dirname(os.path.abspath(__file__))
                except Exception:
                    base_dir = os.getcwd()
                name = f"validation_errors_{datetime.now().strftime('%Y-%m-%d')}.csv"
                report_path = os.path.join(base_dir, name)
                base_df = None
                try:
                    if issue_rows:
                        base_df = pd.DataFrame(issue_rows)
                except Exception:
                    base_df = None
                if base_df is not None and not base_df.empty:
                    if report_df is not None and not report_df.empty:
                        try:
                            df_to_save = pd.concat([base_df, report_df], ignore_index=True, sort=False)
                        except Exception:
                            df_to_save = base_df
                    else:
                        df_to_save = base_df
                elif report_df is not None and not report_df.empty:
                    df_to_save = report_df
                else:
                    df_to_save = pd.DataFrame(
                        [{"MESSAGE": "No validation issues found for current record file."}]
                    )
                df_to_save.to_csv(report_path, index=False)
            except Exception:
                report_path = ""
            issues = []
            if missing_columns:
                issues.append({"kind": "warning", "text": "Missing Columns in record file", "highlight": ""})
            if invalid_branch_rows:
                issues.append({"kind": "error", "text": "Invalid Branch Codes:", "highlight": f"{invalid_branch_rows} Rows"})
            if negative_rows:
                issues.append({"kind": "warning", "text": "Negative Sales Values:", "highlight": f"{negative_rows} Rows"})
            if missing_date_issues:
                issues.append({"kind": "info", "text": "Missing (Branch, POS, Date) combinations:", "highlight": str(missing_date_issues)})
            if not issues:
                issues.append({"kind": "info", "text": "No validation issues found.", "highlight": ""})
            self.root.after(
                0,
                lambda: self._show_validation_dialog(
                    total_rows=total_rows,
                    valid_rows=valid_rows,
                    invalid_rows=invalid_rows,
                    issues=issues,
                    report_path=report_path,
                    branches_missing=branches_missing,
                ),
            )
        except Exception as e:
            try:
                self.root.after(0, lambda: messagebox.showerror("Validation Error", str(e)))
            except Exception:
                pass

    def _on_viewer_source_change(self):
        """Handle switching between master file and record folder sources."""
        try:
            mode = self.viewer_source_var.get()
        except Exception:
            mode = "master"
        try:
            if mode == "master":
                try:
                    if hasattr(self, "viewer_file_combo") and self.viewer_file_combo.winfo_manager():
                        self.viewer_file_combo.grid_remove()
                    if hasattr(self, "viewer_browse_btn") and self.viewer_browse_btn.winfo_manager():
                        self.viewer_browse_btn.grid_remove()
                except Exception:
                    pass
                try:
                    master_path = os.path.join(os.getcwd(), "record2025.csv")
                except Exception:
                    master_path = "record2025.csv"
                self._show_loading_then_load_csv(master_path, mode="master")
            elif mode == "folder":
                try:
                    if hasattr(self, "viewer_file_combo"):
                        self.viewer_file_combo.configure(state="readonly")
                        self.viewer_file_combo.grid()
                    if hasattr(self, "viewer_browse_btn"):
                        self.viewer_browse_btn.configure(state="normal")
                        self.viewer_browse_btn.grid()
                except Exception:
                    pass
                # Always browse folder when switching to folder mode to populate the dropdown
                self._browse_record_folder()
        except Exception:
            pass

    def _browse_record_folder(self):
        try:
            initial_dir = getattr(self, "viewer_records_dir", os.path.join(os.getcwd(), "records"))
        except Exception:
            initial_dir = os.getcwd()
        try:
            folder = filedialog.askdirectory(title="Select Rewritten Records Folder", initialdir=initial_dir)
        except Exception:
            folder = ""
        if not folder:
            return
        try:
            self.viewer_records_dir = folder
        except Exception:
            self.viewer_records_dir = folder
        files = []
        try:
            for name in os.listdir(folder):
                try:
                    if name.lower().endswith(".csv"):
                        full = os.path.join(folder, name)
                        files.append((name, full))
                except Exception:
                    pass
        except Exception:
            files = []
        self.viewer_folder_files = {}
        values = []
        for name, full in sorted(files):
            self.viewer_folder_files[name] = full
            values.append(name)
        try:
            if hasattr(self, "viewer_file_combo"):
                self.viewer_file_combo["values"] = values
        except Exception:
            pass
        if values:
            try:
                self.viewer_file_var.set(values[0])
            except Exception:
                pass
            self._on_viewer_file_selected()
        else:
            try:
                messagebox.showinfo("Rewritten Record Folder", "No CSV files found in the selected folder.")
            except Exception:
                pass

    def _on_viewer_file_selected(self, event=None):
        """Load the viewer from the CSV selected in the record folder combobox."""
        try:
            name = self.viewer_file_var.get()
            path = self.viewer_folder_files.get(name)
            if path:
                self._show_loading_then_load_csv(path, mode="folder", display_name=name)
        except Exception:
            pass
    def _run_patch_missing_data(self, branches_missing, parent_dialog):
        """Use Receive.missing_fetch to download rows for missing (branch, POS, date) gaps."""
        try:
            all_dates = []
            for pos_dict in branches_missing.values():
                for dates in pos_dict.values():
                    for d in dates:
                        all_dates.append(d)
            if not all_dates:
                self.root.after(0, lambda: messagebox.showinfo("Patch Missing Data", "No missing data to patch."))
                return
            start_s = min(all_dates)
            end_s = max(all_dates)
            receiver = Receive(start_s, end_s)
            receiver.missing_fetch(branches_missing)
            def on_success():
                try:
                    if parent_dialog is not None and parent_dialog.winfo_exists():
                        parent_dialog.destroy()
                except Exception:
                    pass
                try:
                    messagebox.showinfo("Patch Missing Data", "Missing data has been patched successfully.")
                except Exception:
                    pass
                try:
                    self.open_latest_record()
                except Exception:
                    pass
            self.root.after(0, on_success)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Patch Missing Data", str(e)))
    def _load_small_logo(self):
        """Load a smaller BIGGS logo into the left card (fallback to text when image missing)."""
        try:
            if not _PIL_AVAILABLE:
                self.logo_label.configure(text="BIGGS", font=("Inter", 18, "bold"))
                return
            img_path = "BIGGS_LOGO.png"
            if not os.path.exists(img_path):
                self.logo_label.configure(text="BIGGS", font=("Inter", 18, "bold"))
                return
            img = Image.open(img_path)
            img.thumbnail((300, 180), Image.LANCZOS)
            self.logo_img_small = ImageTk.PhotoImage(img)
            self.logo_label.configure(image=self.logo_img_small)
        except Exception:
            self.logo_label.configure(text="BIGGS", font=("Inter", 18, "bold"))
        try:
            self.back_btn.lift()
        except Exception:
            pass

    def _status_is_completed(self):
        """Return True when the status label indicates the fetching flow has completed."""
        try:
            t = self.status_label.cget("text")
            return "Completed" in t
        except Exception:
            return not self.is_loading
    def _open_latest_record_when_completed(self):
        """Poll until loading stops and status is Completed, then open the latest record viewer."""
        try:
            if not self.is_loading and self._status_is_completed():
                self._home_view_records()
            else:
                self.root.after(50, self._open_latest_record_when_completed)
        except Exception:
            pass

    def start_fetch(self):
        """Validate date inputs and start an asynchronous fetch operation."""
        # Prevent accidental double-start if fetch is already active
        try:
            if getattr(self, "fetch_active", False):
                return
        except Exception:
            pass
        try:
            mode = self.mode_var.get()
        except Exception:
            mode = "append"
        try:
            self._set_fetch_activity_log_layout()
        except Exception:
            pass
        try:
            start = datetime.strptime(self.start_date.get().strip(), "%Y-%m-%d")
            end = datetime.strptime(self.end_date.get().strip(), "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Invalid Input", "Please use YYYY-MM-DD date format")
            return

        if start > end:
            messagebox.showerror("Date Error", "Start date must be earlier than end date")
            return

        try:
            if mode == "append":
                branches = []
                try:
                    if getattr(self, "selected_branches", None):
                        for b in self.selected_branches:
                            x = (b or "").strip()
                            if x and x not in branches:
                                branches.append(x)
                except Exception:
                    branches = []
                if not branches:
                    try:
                        branch_single = self.branch_var.get()
                    except Exception:
                        branch_single = ""
                    try:
                        bnorm = (branch_single or "").strip()
                    except Exception:
                        bnorm = ""
                    if bnorm and bnorm.lower() not in ("all", "all branches"):
                        branches = [bnorm]
                master_path = os.path.join(os.getcwd(), "record2025.csv")
                if os.path.exists(master_path) and os.path.getsize(master_path) > 0:
                    with open(master_path, newline="", encoding="utf-8") as f:
                        reader = csv.reader(f)
                        headers = next(reader, None) or []
                        idx_map = {h: i for i, h in enumerate(headers)}
                        b_idx = next((idx_map[k] for k in ("BRANCH", "Branch", "branch") if k in idx_map), None)
                        d_idx = next((idx_map[k] for k in ("DATE", "Date", "date") if k in idx_map), None)
                        if d_idx is not None:
                            start_d = start.date()
                            end_d = end.date()
                            has_existing = False
                            for row in reader:
                                if not any(str(x).strip() for x in row):
                                    continue
                                try:
                                    dv = row[d_idx] if d_idx < len(row) else ""
                                    if not dv:
                                        continue
                                    row_date = datetime.strptime(dv.strip(), "%Y-%m-%d").date()
                                except Exception:
                                    continue
                                if row_date < start_d or row_date > end_d:
                                    continue
                                if branches:
                                    if b_idx is None or b_idx >= len(row):
                                        continue
                                    bv = str(row[b_idx]).strip()
                                    if not bv or bv not in branches:
                                        continue
                                has_existing = True
                                break
                            if has_existing:
                                try:
                                    messagebox.showinfo(
                                        "Existing Data",
                                        "The selected branch/branches and inserted date range was already exist in the record.\n\nPlease, select other branches and date range. Thank you!",
                                    )
                                except Exception:
                                    pass
                                return
        except Exception:
            pass

        try:
            self.fetch_activity_log = []
        except Exception:
            self.fetch_activity_log = []
        try:
            if hasattr(self, "fetch_activity_tree"):
                self.fetch_activity_tree.delete(*self.fetch_activity_tree.get_children())
        except Exception:
            pass
        try:
            self._update_fetch_activity_visibility()
        except Exception:
            pass
        

        try:
            self.cancel_fetch_flag = False
        except Exception:
            pass
        self.set_status("Starting extraction...")
        self.start_loading()
        self.log("Extraction process initiated...")
        try:
            self._load_fetch_background_image_async(force_blur=True)
        except Exception:
            pass

        try:
            t = threading.Thread(
                target=self.run_fetch,
                args=(start, end),
                daemon=True
            )
            self.fetch_thread = t
            self.fetch_active = True
            try:
                if hasattr(self, "fetch_main_btn"):
                    self.fetch_main_btn.config(text="Cancel", state="normal")
            except Exception:
                pass
            t.start()
            try:
                self._update_fetch_activity_visibility()
            except Exception:
                pass
        except Exception:
            self.fetch_thread = None
            self.fetch_active = False

    def run_fetch(self, start, end):
        """Run the branch data fetcher, update logs/status, then open the latest record viewer."""
        try:
            self.set_status("Fetching...")
            self.log(f"Starting fetch from {start.strftime('%Y-%m-%d')} to {end.strftime('%Y-%m-%d')}")
            receiver = Receive(start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))
            try:
                self.receiver = receiver
            except Exception:
                pass
            try:
                branch = self.branch_var.get()
            except Exception:
                branch = ""
            branches = []
            try:
                if getattr(self, "selected_branches", None):
                    for b in self.selected_branches:
                        x = (b or "").strip()
                        if x and x not in branches:
                            branches.append(x)
            except Exception:
                branches = []
            if branches:
                receiver.branches = branches
            else:
                try:
                    bnorm = (branch or "").strip()
                except Exception:
                    bnorm = ""
                if bnorm and bnorm.lower() not in ("all", "all branches"):
                    receiver.branches = [bnorm]
            rewrite_target = None
            merge_to_master = False
            # Rewrite mode
            if self.mode_var.get() == "rewrite":
                try:
                    settings_dir = os.path.join(os.getcwd(), "settings")
                    default_records_dir = os.path.join(os.getcwd(), "records")
                    os.makedirs(settings_dir, exist_ok=True)
                    base_dir = default_records_dir
                    os.makedirs(base_dir, exist_ok=True)
                    start_s = start.strftime("%Y-%m-%d")
                    end_s = end.strftime("%Y-%m-%d")
                    rewrite_target = os.path.join(base_dir, f"record_{start_s}_to_{end_s}.csv")
                    with open(os.path.join(settings_dir, "current_record.txt"), "w", encoding="utf-8") as f:
                        f.write(rewrite_target)
                    merge_to_master = True
                except Exception:
                    rewrite_target = None
                    merge_to_master = False
            else:
                # Append: ensure master record exists and merge any previous rewrite file into it
                try:
                    prev_ptr = self._get_current_record_path()
                    self._ensure_master_record_merge(previous_path=prev_ptr)
                except Exception:
                    pass
            try:
                total_dates = len(getattr(receiver, "dlist", []))
                total_branches = len(getattr(receiver, "branches", [])) or 1
                total_steps = max(1, total_dates * total_branches)
            except Exception:
                total_steps = 1
            try:
                self._progress_start_ts = time.perf_counter()
            except Exception:
                self._progress_start_ts = None

            def _hook(done, total):
                try:
                    denom = total or total_steps or 1
                    pct = (float(done) / float(denom)) * 100.0
                except Exception:
                    pct = 0.0
                try:
                    self.root.after(0, lambda v=pct: self._update_progress_percent(v))
                    self.root.after(0, lambda d=done, t=denom: self._update_progress_stats(d, t))
                except Exception:
                    pass

            try:
                receiver.progress_hook = _hook
                receiver.progress_total = total_steps
            except Exception:
                pass

            def _log_hook(info):
                try:
                    self.log(info)
                except Exception:
                    pass

            try:
                receiver.log_callback = _log_hook
            except Exception:
                pass

            receiver.fetch()
            cancelled = False
            try:
                cancelled = bool(getattr(self, "cancel_fetch_flag", False)) or bool(getattr(receiver, "exitFlag", 0))
            except Exception:
                cancelled = False
            if cancelled:
                self.set_status("Cancelled")
                self.log("⚠ Fetch cancelled by user")
            else:
                if self.mode_var.get() == "rewrite" and merge_to_master and rewrite_target:
                    try:
                        self._merge_rewrite_into_master(previous_path=rewrite_target)
                        self.log("✔ Rewrite data merged into master record2025.csv")
                    except Exception:
                        pass
                self.set_status("Completed")
                self.log("✔ All branch data successfully fetched and record updated")
                try:
                    self.root.after(0, self._show_master_record_in_fetch_panel)
                except Exception:
                    pass
                self.root.after(0, self.stop_loading)
                self.root.after(50, self._open_latest_record_when_completed)
                self.root.after(100, lambda: messagebox.showinfo("Success", "Data extraction completed successfully! The latest record is now updated and sorted by fetch time."))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {str(e)}"))
            self.set_status("Error occurred")
        finally:
            self.root.after(0, self.stop_loading)
            def _reset_fetch_button():
                try:
                    if hasattr(self, "fetch_main_btn"):
                        try:
                            mode = self.mode_var.get()
                        except Exception:
                            mode = "append"
                        label = "Fetch"
                        if mode == "rewrite":
                            label = "View"
                            try:
                                self.rewrite_phase = "view"
                            except Exception:
                                pass
                        self.fetch_main_btn.config(text=label, state="normal")
                except Exception:
                    pass
            self.root.after(0, _reset_fetch_button)
            try:
                self.receiver = None
            except Exception:
                pass
            try:
                self.fetch_active = False
                self.fetch_thread = None
            except Exception:
                pass


if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = BiggsExtractorGUI(root)
        root.mainloop()
    except KeyboardInterrupt:
        pass
