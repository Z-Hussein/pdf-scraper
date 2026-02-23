import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter import font as tkfont
from tkinter import simpledialog
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import os
import re
import json
import threading
import webbrowser
from datetime import datetime

class PDFScraperPro:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Scraper Pro - Advanced PDF Tool")
        self.root.geometry("1400x900")
        self.root.configure(bg='#1e1e1e')
        
        # Set app icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'avatar.png')
            if os.path.exists(icon_path):
                icon = tk.PhotoImage(file=icon_path)
                self.root.iconphoto(False, icon)
        except Exception as e:
            print(f"Could not load icon: {e}")
        
        # Variables
        self.current_pdf = None
        self.current_page = 0
        self.total_pages = 0
        self.zoom_level = 1.0
        self.search_results = []
        self.current_search_index = -1
        self.pdf_document = None
        self.extracted_text = ""
        self.history = []
        # UI / theme
        self.dark_mode = True
        self.custom_buttons = []  # track tk.Button instances for dynamic theming
        self.main_font = ('Segoe UI', 10)
        self.auto_fit = True
        self._resize_after = None
        self.tesseract_cmd = None
        
        # Colors
        self.bg_color = '#1e1e1e'
        self.fg_color = '#ffffff'
        self.accent_color = '#007acc'
        self.secondary_bg = '#252526'
        self.border_color = '#3e3e42'
        
        self.setup_styles()
        self.create_menu()
        self.create_layout()
        self.create_bindings()
        
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        # Configure styles (use instance fonts/colors so they can be updated)
        style.configure('Custom.TFrame', background=self.bg_color)
        style.configure('Custom.TButton',
                        background=self.accent_color,
                        foreground='white',
                        font=self.main_font,
                        padding=8)
        style.configure('Custom.TEntry',
                        fieldbackground=self.secondary_bg,
                        foreground=self.fg_color,
                        insertcolor=self.fg_color)
        style.configure('Custom.TLabel',
                        background=self.bg_color,
                        foreground=self.fg_color,
                        font=self.main_font)
        style.configure('Custom.TNotebook',
                        background=self.bg_color,
                        tabmargins=[2, 5, 2, 0])
        style.configure('Custom.TNotebook.Tab',
                        background=self.secondary_bg,
                        foreground=self.fg_color,
                        padding=[10, 5],
                        font=(self.main_font[0], 9))
        style.map('Custom.TNotebook.Tab',
                  background=[('selected', self.accent_color)],
                  foreground=[('selected', 'white')])

        # Apply theme to some widgets immediately
        self.apply_theme()
        # ensure canvas updates when container resizes
        try:
            self.root.update_idletasks()
        except Exception:
            pass
        
    def create_menu(self):
        menubar = tk.Menu(self.root, bg=self.secondary_bg, fg=self.fg_color)
        self.root.config(menu=menubar)
        
        # File Menu
        file_menu = tk.Menu(menubar, tearoff=0, bg=self.secondary_bg, fg=self.fg_color)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open PDF", command=self.open_pdf, accelerator="Ctrl+O")
        file_menu.add_command(label="Open Recent", command=self.show_recent)
        file_menu.add_separator()
        file_menu.add_command(label="Save As", command=self.save_pdf, accelerator="Ctrl+S")
        file_menu.add_command(label="Export Text", command=self.export_text)
        file_menu.add_command(label="Export Images", command=self.export_images)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Edit Menu
        edit_menu = tk.Menu(menubar, tearoff=0, bg=self.secondary_bg, fg=self.fg_color)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Undo", command=self.undo_action, accelerator="Ctrl+Z")
        edit_menu.add_command(label="Redo", command=self.redo_action, accelerator="Ctrl+Y")
        edit_menu.add_separator()
        edit_menu.add_command(label="Find", command=self.focus_search, accelerator="Ctrl+F")
        edit_menu.add_command(label="Find Next", command=self.find_next, accelerator="F3")
        edit_menu.add_separator()
        edit_menu.add_command(label="Rotate Clockwise", command=lambda: self.rotate_page(90))
        edit_menu.add_command(label="Rotate Counter-Clockwise", command=lambda: self.rotate_page(-90))
        edit_menu.add_separator()
        edit_menu.add_command(label="Delete Page", command=self.delete_page)
        edit_menu.add_command(label="Extract Pages", command=self.extract_pages)
        
        # Tools Menu
        tools_menu = tk.Menu(menubar, tearoff=0, bg=self.secondary_bg, fg=self.fg_color)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Batch Process", command=self.batch_process)
        tools_menu.add_command(label="Configure OCR", command=self.configure_ocr)
        tools_menu.add_command(label="Merge PDFs", command=self.merge_pdfs)
        tools_menu.add_command(label="Split PDF", command=self.split_pdf)
        tools_menu.add_separator()
        tools_menu.add_command(label="OCR (Text Recognition)", command=self.ocr_page)
        tools_menu.add_command(label="Compress PDF", command=self.compress_pdf)
        tools_menu.add_separator()
        tools_menu.add_command(label="Scrape All Text", command=self.scrape_all_text)
        tools_menu.add_command(label="Scrape Images", command=self.scrape_images)
        tools_menu.add_command(label="Scrape Metadata", command=self.scrape_metadata)
        tools_menu.add_command(label="Scrape Links", command=self.scrape_links)
        tools_menu.add_command(label="Scrape Tables", command=self.scrape_tables)
        
        # View Menu
        view_menu = tk.Menu(menubar, tearoff=0, bg=self.secondary_bg, fg=self.fg_color)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Zoom In", command=self.zoom_in, accelerator="Ctrl++")
        view_menu.add_command(label="Zoom Out", command=self.zoom_out, accelerator="Ctrl+-")
        view_menu.add_command(label="Fit to Width", command=self.fit_to_width)
        view_menu.add_command(label="Fit to Page", command=self.fit_to_page)
        view_menu.add_separator()
        view_menu.add_command(label="Open New Window", command=self.open_second_window)
        view_menu.add_command(label="Reorder Sections", command=self.reorder_tabs)
        view_menu.add_separator()
        view_menu.add_command(label="About", command=self.about_window)
        view_menu.add_checkbutton(label="Dark Mode", command=self.toggle_theme)
        
    def create_layout(self):
        # Main container as a PanedWindow so user can resize sections
        self.paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashrelief=tk.RAISED, sashwidth=8, bg=self.border_color)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        # Undo/Redo stacks
        self.undo_stack = []  # store bytes of pdf
        self.redo_stack = []
        # secondary canvas windows
        self.secondary_windows = []
        
        # Left Panel - PDF Preview
        left_frame = ttk.Frame(self.paned, style='Custom.TFrame')
        # add left frame to paned window (user can resize via sash)
        self.paned.add(left_frame, minsize=400)
        
        # Toolbar
        toolbar = ttk.Frame(left_frame, style='Custom.TFrame')
        toolbar.pack(fill=tk.X, pady=(0, 5))
        
        # Avatar display
        try:
            avatar_path = os.path.join(os.path.dirname(__file__), 'avatar.png')
            if os.path.exists(avatar_path):
                avatar_img = Image.open(avatar_path)
                avatar_img.thumbnail((40, 40), Image.Resampling.LANCZOS)
                avatar_photo = ImageTk.PhotoImage(avatar_img)
                avatar_label = tk.Label(toolbar, image=avatar_photo, bg=self.bg_color)
                avatar_label.image = avatar_photo  # Keep a reference
                avatar_label.pack(side=tk.LEFT, padx=5, pady=2)
        except Exception as e:
            print(f"Could not load avatar: {e}")
        
        # Undo/Redo buttons
        ttk.Button(toolbar, text="Undo", command=self.undo_action, style='Custom.TButton').pack(side=tk.RIGHT, padx=2)
        ttk.Button(toolbar, text="Redo", command=self.redo_action, style='Custom.TButton').pack(side=tk.RIGHT, padx=2)
        ttk.Button(toolbar, text="Reorder Pages", command=self.reorder_pages, style='Custom.TButton').pack(side=tk.RIGHT, padx=2)
        ttk.Button(toolbar, text="Reorder Sections", command=self.reorder_tabs, style='Custom.TButton').pack(side=tk.RIGHT, padx=2)
        ttk.Button(toolbar, text="New Window", command=self.open_second_window, style='Custom.TButton').pack(side=tk.RIGHT, padx=2)
        
        # Navigation buttons
        nav_frame = ttk.Frame(toolbar, style='Custom.TFrame')
        nav_frame.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(nav_frame, text="◀", width=3, command=self.prev_page).pack(side=tk.LEFT, padx=2)
        self.page_label = ttk.Label(nav_frame, text="Page: 0/0", style='Custom.TLabel')
        self.page_label.pack(side=tk.LEFT, padx=10)
        ttk.Button(nav_frame, text="▶", width=3, command=self.next_page).pack(side=tk.LEFT, padx=2)
        
        # Zoom controls
        zoom_frame = ttk.Frame(toolbar, style='Custom.TFrame')
        zoom_frame.pack(side=tk.LEFT, padx=20)
        
        ttk.Button(zoom_frame, text="-", width=3, command=self.zoom_out).pack(side=tk.LEFT, padx=2)
        self.zoom_label = ttk.Label(zoom_frame, text="100%", style='Custom.TLabel')
        self.zoom_label.pack(side=tk.LEFT, padx=10)
        ttk.Button(zoom_frame, text="+", width=3, command=self.zoom_in).pack(side=tk.LEFT, padx=2)
        
        # Search buttons (input moved to Search tab)
        self.search_var = tk.StringVar()
        
        # PDF Canvas with scrollbars
        canvas_frame = ttk.Frame(left_frame, style='Custom.TFrame')
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        self.canvas = tk.Canvas(canvas_frame, bg=self.secondary_bg, highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        v_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        h_scrollbar = ttk.Scrollbar(left_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        h_scrollbar.pack(fill=tk.X)
        
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Right Panel - Tools and Info
        right_frame = ttk.Frame(self.paned, style='Custom.TFrame', width=400)
        # add right frame to paned window
        self.paned.add(right_frame, minsize=300)
        right_frame.pack_propagate(False)
        
        # Notebook for tabs
        self.notebook = ttk.Notebook(right_frame, style='Custom.TNotebook')
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Tab 1: Text Extraction
        self.text_tab = ttk.Frame(self.notebook, style='Custom.TFrame')
        self.notebook.add(self.text_tab, text="Extracted Text")
        
        text_toolbar = ttk.Frame(self.text_tab, style='Custom.TFrame')
        text_toolbar.pack(fill=tk.X, pady=5)
        
        ttk.Button(text_toolbar, text="Extract Current Page", command=self.extract_current_page).pack(side=tk.LEFT, padx=2)
        ttk.Button(text_toolbar, text="Extract All", command=self.extract_all_pages).pack(side=tk.LEFT, padx=2)
        ttk.Button(text_toolbar, text="Copy", command=self.copy_text).pack(side=tk.LEFT, padx=2)
        ttk.Button(text_toolbar, text="Clear", command=self.clear_text).pack(side=tk.LEFT, padx=2)
        
        self.text_area = scrolledtext.ScrolledText(
            self.text_tab, 
            wrap=tk.WORD, 
            bg=self.secondary_bg, 
            fg=self.fg_color,
            insertbackground=self.fg_color,
            font=('Consolas', 10),
            padx=10,
            pady=10
        )
        self.text_area.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Tab 2: Search Results
        self.search_tab = ttk.Frame(self.notebook, style='Custom.TFrame')
        self.notebook.add(self.search_tab, text="Search")
        
        search_toolbar = ttk.Frame(self.search_tab, style='Custom.TFrame')
        search_toolbar.pack(fill=tk.X, pady=5)

        # Search input is here now
        self.search_entry = ttk.Entry(search_toolbar, textvariable=self.search_var, width=36, style='Custom.TEntry')
        self.search_entry.pack(side=tk.LEFT, padx=5)
        self.search_entry.insert(0, "Search in PDF...")
        self.search_entry.config(foreground='gray')

        ttk.Button(search_toolbar, text="Find", command=self.search_pdf).pack(side=tk.LEFT, padx=2)
        ttk.Button(search_toolbar, text="Find All", command=self.search_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(search_toolbar, text="Previous", command=self.prev_search_result).pack(side=tk.LEFT, padx=2)
        ttk.Button(search_toolbar, text="Next", command=self.next_search_result).pack(side=tk.LEFT, padx=2)
        self.result_count_label = ttk.Label(search_toolbar, text="Results: 0", style='Custom.TLabel')
        self.result_count_label.pack(side=tk.LEFT, padx=10)

        # (removed extra search option buttons; use simple case-insensitive search)
        # Interactive results list
        results_frame = ttk.Frame(self.search_tab, style='Custom.TFrame')
        results_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.search_results_list = tk.Listbox(results_frame, bg=self.secondary_bg, fg=self.fg_color, activestyle='none')
        self.search_results_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        res_scroll = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.search_results_list.yview)
        res_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.search_results_list.configure(yscrollcommand=res_scroll.set)
        self.search_results_list.bind('<Double-Button-1>', lambda e: self._on_result_activate())
        self.search_results_list.bind('<Return>', lambda e: self._on_result_activate())
        self.search_results_list.bind('<<ListboxSelect>>', lambda e: self._on_result_select())
        
        # Tab 3: Metadata
        self.metadata_tab = ttk.Frame(self.notebook, style='Custom.TFrame')
        self.notebook.add(self.metadata_tab, text="Metadata")
        
        self.metadata_text = scrolledtext.ScrolledText(
            self.metadata_tab,
            wrap=tk.WORD,
            bg=self.secondary_bg,
            fg=self.fg_color,
            insertbackground=self.fg_color,
            font=('Consolas', 10)
        )
        self.metadata_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Tab 4: Scraping Tools
        self.scrape_tab = ttk.Frame(self.notebook, style='Custom.TFrame')
        self.notebook.add(self.scrape_tab, text="Scraping Tools")
        
        scrape_buttons = [
            ("Scrape All Text", self.scrape_all_text),
            ("Scrape Images", self.scrape_images),
            ("Scrape Links", self.scrape_links),
            ("Scrape Tables", self.scrape_tables),
            ("Scrape Annotations", self.scrape_annotations),
            ("Scrape Form Fields", self.scrape_form_fields),
            ("Batch Scrape Folder", self.batch_scrape_folder),
        ]
        
        for text, command in scrape_buttons:
            btn = tk.Button(
                self.scrape_tab,
                text=text,
                command=command,
                bg=self.accent_color,
                fg='white',
                font=('Segoe UI', 10),
                relief=tk.FLAT,
                padx=10,
                pady=5,
                cursor='hand2'
            )
            # track for theme updates
            self.custom_buttons.append(btn)
            btn.pack(fill=tk.X, padx=10, pady=5)
            
            # Hover effect
            btn.bind('<Enter>', lambda e, b=btn: b.config(bg='#005a9e'))
            btn.bind('<Leave>', lambda e, b=btn: b.config(bg=self.accent_color))
        
        # Tab 5: PDF Editing
        self.edit_tab = ttk.Frame(self.notebook, style='Custom.TFrame')
        self.notebook.add(self.edit_tab, text="PDF Editing")
        
        edit_buttons = [
            ("Add Watermark", self.add_watermark),
            ("Remove Pages", self.delete_page),
            ("Rotate Page", lambda: self.rotate_page(90)),
            ("Merge PDFs", self.merge_pdfs),
            ("Split PDF", self.split_pdf),
            ("Compress PDF", self.compress_pdf),
            ("Redact Text", self.redact_text),
            ("Add Text", self.add_text_to_pdf),
            ("Add Image", self.add_image_to_pdf),
        ]
        
        for text, command in edit_buttons:
            btn = tk.Button(
                self.edit_tab,
                text=text,
                command=command,
                bg=self.accent_color,
                fg='white',
                font=('Segoe UI', 10),
                relief=tk.FLAT,
                padx=10,
                pady=5,
                cursor='hand2'
            )
            # track for theme updates
            self.custom_buttons.append(btn)
            btn.pack(fill=tk.X, padx=10, pady=5)
            
            btn.bind('<Enter>', lambda e, b=btn: b.config(bg='#005a9e'))
            btn.bind('<Leave>', lambda e, b=btn: b.config(bg=self.accent_color))
        
        # Status bar
        self.status_bar = ttk.Label(
            self.root, 
            text="Ready", 
            relief=tk.SUNKEN, 
            anchor=tk.W,
            style='Custom.TLabel'
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def apply_theme(self):
        """Apply current theme colors to ttk styles and relevant widgets."""
        style = ttk.Style()
        if self.dark_mode:
            self.bg_color = '#1e1e1e'
            self.fg_color = '#ffffff'
            self.accent_color = '#007acc'
            self.secondary_bg = '#252526'
            self.border_color = '#3e3e42'
        else:
            self.bg_color = '#f3f3f3'
            self.fg_color = '#111111'
            self.accent_color = '#1a73e8'
            self.secondary_bg = '#ffffff'
            self.border_color = '#d0d0d0'

        style.configure('Custom.TFrame', background=self.bg_color)
        style.configure('Custom.TLabel', background=self.bg_color, foreground=self.fg_color, font=self.main_font)
        style.configure('Custom.TEntry', fieldbackground=self.secondary_bg, foreground=self.fg_color, insertcolor=self.fg_color)
        style.configure('Custom.TNotebook', background=self.bg_color)
        style.configure('Custom.TNotebook.Tab', background=self.secondary_bg, foreground=self.fg_color)
        style.configure('Custom.TButton', background=self.accent_color, foreground='white', font=self.main_font)

        # update canvas & text areas
        try:
            self.canvas.configure(bg=self.secondary_bg)
        except Exception:
            pass

        for ta in (getattr(self, 'text_area', None), getattr(self, 'metadata_text', None), getattr(self, 'search_results_text', None)):
            if ta:
                ta.configure(bg=self.secondary_bg, fg=self.fg_color, insertbackground=self.fg_color)

        # update tracked tk.Button widgets
        for b in self.custom_buttons:
            try:
                b.configure(bg=self.accent_color if self.dark_mode else self.accent_color, fg='white')
            except Exception:
                pass

        # update status bar
        try:
            self.status_bar.configure(style='Custom.TLabel')
        except Exception:
            pass
        
    def create_bindings(self):
        self.root.bind('<Control-o>', lambda e: self.open_pdf())
        self.root.bind('<Control-s>', lambda e: self.save_pdf())
        self.root.bind('<Control-z>', lambda e: self.undo_action())
        self.root.bind('<Control-y>', lambda e: self.redo_action())
        self.root.bind('<Control-f>', lambda e: self.focus_search())
        self.root.bind('<F3>', lambda e: self.find_next())
        self.root.bind('<Control-plus>', lambda e: self.zoom_in())
        self.root.bind('<Control-minus>', lambda e: self.zoom_out())
        self.root.bind('<Left>', lambda e: self.prev_page())
        self.root.bind('<Right>', lambda e: self.next_page())
        
        self.search_entry.bind('<FocusIn>', self.on_search_focus_in)
        self.search_entry.bind('<FocusOut>', self.on_search_focus_out)
        self.search_entry.bind('<Return>', lambda e: self.search_pdf())
        
        self.canvas.bind('<MouseWheel>', self.on_mousewheel)
        self.canvas.bind('<Button-4>', lambda e: self.canvas.yview_scroll(-1, 'units'))
        self.canvas.bind('<Button-5>', lambda e: self.canvas.yview_scroll(1, 'units'))
        # Notebook tab dragging bindings (tear-off to new window)
        try:
            self.notebook.bind('<ButtonPress-1>', self.on_tab_button_press, add='+')
            self.notebook.bind('<B1-Motion>', self.on_tab_motion, add='+')
            self.notebook.bind('<ButtonRelease-1>', self.on_tab_release, add='+')
        except Exception:
            pass

        # handle canvas/container resize to update auto-fit view (debounced)
        try:
            self.canvas.bind('<Configure>', self.on_canvas_configure, add='+')
        except Exception:
            pass

        
    def on_search_focus_in(self, event):
        if self.search_entry.get() == "Search in PDF...":
            self.search_entry.delete(0, tk.END)
            self.search_entry.config(foreground=self.fg_color)
            
    def on_search_focus_out(self, event):
        if not self.search_entry.get():
            self.search_entry.insert(0, "Search in PDF...")
            self.search_entry.config(foreground='gray')
            
    def focus_search(self):
        self.search_entry.focus_set()
        self.search_entry.select_range(0, tk.END)
        
    def on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), 'units')

    def on_canvas_configure(self, event):
        if not getattr(self, 'auto_fit', False) or not self.pdf_document:
            return
        # debounce rapid resize events
        if getattr(self, '_resize_after', None):
            try:
                self.root.after_cancel(self._resize_after)
            except Exception:
                pass
        self._resize_after = self.root.after(100, self.display_page)
        
    def open_pdf(self):
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if file_path:
            self.load_pdf(file_path)
            
    def load_pdf(self, file_path):
        try:
            if self.pdf_document:
                self.pdf_document.close()
                
            self.pdf_document = fitz.open(file_path)
            self.current_pdf = file_path
            self.current_page = 0
            self.total_pages = len(self.pdf_document)
            self.zoom_level = 1.0
            
            self.add_to_history(file_path)
            self.display_page()
            self.update_status(f"Loaded: {os.path.basename(file_path)} | Pages: {self.total_pages}")
            self.scrape_metadata()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load PDF:\n{str(e)}")
            
    def display_page(self):
        if not self.pdf_document:
            return
        
        try:
            page = self.pdf_document[self.current_page]

            # Auto-fit behavior: if enabled, compute a zoom that fits canvas width
            if self.auto_fit:
                try:
                    canvas_w = self.canvas.winfo_width()
                    # if canvas not yet mapped, retry shortly
                    if canvas_w <= 1:
                        self.root.after(50, self.display_page)
                        return

                    # get base pixmap at zoom 1 to determine natural pixel width
                    base_pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
                    base_width = base_pix.width if base_pix else 1
                    desired_zoom = canvas_w / base_width
                    # clamp zoom to reasonable bounds
                    desired_zoom = max(0.25, min(5.0, desired_zoom))
                    self.zoom_level = desired_zoom
                except Exception:
                    # fallback to existing zoom_level
                    pass

            # Calculate zoom matrix
            mat = fitz.Matrix(self.zoom_level, self.zoom_level)
            pix = page.get_pixmap(matrix=mat)
            
            # Convert to PIL Image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # Convert to PhotoImage
            self.current_image = ImageTk.PhotoImage(img)
            
            # Update primary canvas
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.current_image)
            self.canvas.config(scrollregion=self.canvas.bbox("all"))
            
            # Update secondary canvases if any
            # secondary_windows contains dicts: {'canvas': Canvas, 'image_ref': PhotoImage or None, 'auto_fit': bool}
            for sec in list(self.secondary_windows):
                try:
                    sc = sec['canvas']
                except Exception:
                    # legacy: plain canvas
                    sc = sec

                try:
                    # If this entry is a dict with auto-fit, render its own image at computed zoom
                    if isinstance(sec, dict) and sec.get('auto_fit', True):
                        try:
                            canvas_w = sc.winfo_width()
                            if canvas_w <= 1:
                                # not ready yet
                                continue
                            base_pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
                            base_width = base_pix.width if base_pix else 1
                            desired_zoom = canvas_w / base_width
                            desired_zoom = max(0.25, min(5.0, desired_zoom))
                        except Exception:
                            desired_zoom = self.zoom_level

                        mat2 = fitz.Matrix(desired_zoom, desired_zoom)
                        pix2 = page.get_pixmap(matrix=mat2)
                        img2 = Image.frombytes("RGB", [pix2.width, pix2.height], pix2.samples)
                        photo2 = ImageTk.PhotoImage(img2)

                        sc.delete("all")
                        sc.create_image(0, 0, anchor=tk.NW, image=photo2)
                        sc.config(scrollregion=sc.bbox("all"))

                        # keep reference
                        sec['image_ref'] = photo2
                    else:
                        sc.delete("all")
                        sc.create_image(0, 0, anchor=tk.NW, image=self.current_image)
                        sc.config(scrollregion=sc.bbox("all"))
                except Exception:
                    # ignore secondary render errors
                    pass
            
            # Update page label
            self.page_label.config(text=f"Page: {self.current_page + 1}/{self.total_pages}")
            self.zoom_label.config(text=f"{int(self.zoom_level * 100)}%")
            
            # Highlight search results if any
            self.highlight_search_results()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to display page:\n{str(e)}")
            
    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.display_page()
            
    def next_page(self):
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.display_page()
            
    def zoom_in(self):
        # manual zoom disables auto-fit
        self.auto_fit = False
        if self.zoom_level < 5.0:
            self.zoom_level += 0.25
            self.display_page()
            
    def zoom_out(self):
        # manual zoom disables auto-fit
        self.auto_fit = False
        if self.zoom_level > 0.25:
            self.zoom_level -= 0.25
            self.display_page()
            
    def fit_to_width(self):
        if not self.pdf_document:
            return
        # Turn on auto-fit and refresh (display_page will calculate proper zoom)
        self.auto_fit = True
        self.display_page()
        
    def fit_to_page(self):
        self.zoom_level = 1.0
        self.display_page()
        
    def search_pdf(self):
        if not self.pdf_document:
            messagebox.showwarning("Warning", "Please open a PDF first")
            return
            
        query = self.search_var.get()
        # clear previous results first
        self.search_results = []
        self.current_search_index = -1

        if not query or query == "Search in PDF...":
            return

        # Word-level/phrase search across pages (avoids huge false positives)
        tokens = query.split()
        tcount = len(tokens)
        for page_num in range(self.total_pages):
            page = self.pdf_document[page_num]
            words = page.get_text("words")  # list of (x0,y0,x1,y1, word, ...)
            if not words:
                continue

            # build a simple list of word texts
            word_texts = [w[4] for w in words]
            # iterate moving window to find sequence matches
            for i in range(0, len(word_texts) - tcount + 1):
                seq = word_texts[i:i + tcount]
                seq_text = " ".join(seq)
                # case-insensitive phrase match
                if query.lower() in seq_text.lower():
                    # compute bbox spanning sequence
                    seq_boxes = words[i:i + tcount]
                    x0 = min(b[0] for b in seq_boxes)
                    y0 = min(b[1] for b in seq_boxes)
                    x1 = max(b[2] for b in seq_boxes)
                    y1 = max(b[3] for b in seq_boxes)
                    rect = fitz.Rect(x0, y0, x1, y1)
                    raw = page.get_textbox(rect)
                    match_text = seq_text
                    if not self._is_duplicate_match(page_num, match_text, rect):
                        self.search_results.append({
                            'page': page_num,
                            'rect': rect,
                            'text': raw,
                            'match': match_text
                        })

    def _is_duplicate_match(self, page_num, match_text, inst, tol=4.0):
        """Return True if a similar match already exists (same page, same text, bbox center within tol pixels)."""
        try:
            # get center of inst
            try:
                cx = (inst.x0 + inst.x1) / 2
                cy = (inst.y0 + inst.y1) / 2
            except Exception:
                # inst might be tuple-like
                cx = ((inst[0] + inst[2]) / 2) if len(inst) >= 4 else 0
                cy = ((inst[1] + inst[3]) / 2) if len(inst) >= 4 else 0
        except Exception:
            cx = cy = 0

        mt = (match_text or '').strip().lower()
        for r in self.search_results:
            if r.get('page') != page_num:
                continue
            if (r.get('match') or '').strip().lower() != mt:
                continue
            # compute center of existing rect
            other = r.get('rect')
            try:
                try:
                    ox = (other.x0 + other.x1) / 2
                    oy = (other.y0 + other.y1) / 2
                except Exception:
                    ox = ((other[0] + other[2]) / 2) if len(other) >= 4 else 0
                    oy = ((other[1] + other[3]) / 2) if len(other) >= 4 else 0
            except Exception:
                ox = oy = 0

            dx = abs(cx - ox)
            dy = abs(cy - oy)
            if dx <= tol and dy <= tol:
                return True
        return False
                
        self.update_search_results()
        
        if self.search_results:
            self.current_search_index = 0
            self.goto_search_result(0)
            
    def search_all(self):
        query = self.search_var.get()
        if not query or query == "Search in PDF...":
            messagebox.showwarning("Warning", "Please enter a search term first")
            return
        self.search_pdf()

        # Display all results in listbox
        self.search_results_list.delete(0, tk.END)

        if not self.search_results:
            self.search_results_list.insert(tk.END, f"No results found for '{self.search_var.get()}'")
            self.result_count_label.config(text="Results: 0")
            return

        for i, result in enumerate(self.search_results, 1):
            matched = result.get('match') or result.get('text', '')
            matched = matched.replace('\n', ' ').strip()
            entry_text = f"{i}. Page {result['page'] + 1}: {matched}"
            self.search_results_list.insert(tk.END, entry_text)

        self.result_count_label.config(text=f"Results: {len(self.search_results)}")
        # highlight matches on current page
        self.highlight_search_results()
            
    def update_search_results(self):
        count = len(self.search_results)
        self.result_count_label.config(text=f"Results: {count}")
        
        if count > 0:
            self.status_bar.config(text=f"Found {count} matches")
        else:
            self.status_bar.config(text="No matches found")
            
    def goto_search_result(self, index):
        if not self.search_results or index < 0 or index >= len(self.search_results):
            return
            
        result = self.search_results[index]
        self.current_page = result['page']
        self.display_page()
        
        # Scroll to result (approximate)
        self.canvas.yview_moveto(0.3)
        # mark current search index and redraw highlights
        self.current_search_index = index
        self.highlight_search_results()

    def _on_result_activate(self):
        sel = self.search_results_list.curselection()
        if not sel:
            return
        idx = sel[0]
        self.goto_search_result(idx)

    def _on_result_select(self):
        sel = self.search_results_list.curselection()
        if not sel:
            return
        idx = sel[0]
        # update preview highlight but don't change page until double-click
        try:
            self.current_search_index = idx
            # if the selected result is on a different page, show its page
            if self.search_results and 0 <= idx < len(self.search_results):
                page = self.search_results[idx]['page']
                if page != self.current_page:
                    self.current_page = page
                    self.display_page()
                else:
                    self.highlight_search_results()
        except Exception:
            pass
        
    def next_search_result(self):
        if self.search_results:
            self.current_search_index = (self.current_search_index + 1) % len(self.search_results)
            self.goto_search_result(self.current_search_index)
            
    def prev_search_result(self):
        if self.search_results:
            self.current_search_index = (self.current_search_index - 1) % len(self.search_results)
            self.goto_search_result(self.current_search_index)
            
    def highlight_search_results(self):
        # draw rectangles on the canvas for search highlights
        try:
            self.canvas.delete('search_highlight')
        except Exception:
            pass

        if not self.search_results:
            return

        for i, res in enumerate(self.search_results):
            if res['page'] != self.current_page:
                continue

            rect = res['rect']
            try:
                # handle fit/zoom mapping: PDF coords * zoom_level -> image pixels
                try:
                    x0 = rect.x0 * self.zoom_level
                    y0 = rect.y0 * self.zoom_level
                    x1 = rect.x1 * self.zoom_level
                    y1 = rect.y1 * self.zoom_level
                except Exception:
                    # rect may be tuple
                    x0 = rect[0] * self.zoom_level
                    y0 = rect[1] * self.zoom_level
                    x1 = rect[2] * self.zoom_level
                    y1 = rect[3] * self.zoom_level

                # draw rectangle
                color = '#ffea00' if i != self.current_search_index else '#ff5722'
                try:
                    self.canvas.create_rectangle(x0, y0, x1, y1, outline=color, width=2, tags='search_highlight')
                except Exception:
                    pass
            except Exception:
                pass
        
    def find_next(self):
        self.next_search_result()
        
    def extract_current_page(self):
        if not self.pdf_document:
            return
            
        page = self.pdf_document[self.current_page]
        text = page.get_text()
        
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(tk.END, f"--- Page {self.current_page + 1} ---\n\n")
        self.text_area.insert(tk.END, text)
        
    def extract_all_pages(self):
        if not self.pdf_document:
            return
            
        self.text_area.delete(1.0, tk.END)
        full_text = []
        
        for i in range(self.total_pages):
            page = self.pdf_document[i]
            text = page.get_text()
            full_text.append(f"--- Page {i + 1} ---\n{text}\n")
            
        self.text_area.insert(tk.END, "\n".join(full_text))
        self.extracted_text = "\n".join(full_text)
        
    def copy_text(self):
        selected = self.text_area.get(tk.SEL_FIRST, tk.SEL_LAST) if self.text_area.tag_ranges(tk.SEL) else self.text_area.get(1.0, tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(selected)
        self.status_bar.config(text="Text copied to clipboard")
        
    def clear_text(self):
        self.text_area.delete(1.0, tk.END)
        
    def scrape_all_text(self):
        if not self.pdf_document:
            messagebox.showwarning("Warning", "Please open a PDF first")
            return
            
        def scrape():
            all_text = []
            for i in range(self.total_pages):
                page = self.pdf_document[i]
                blocks = page.get_text("blocks")
                
                for block in blocks:
                    if block[6] == 0:  # Text block
                        all_text.append({
                            'page': i + 1,
                            'text': block[4],
                            'bbox': block[:4]
                        })
                        
            # Save to file
            output_file = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("Text files", "*.txt")]
            )
            
            if output_file:
                with open(output_file, 'w', encoding='utf-8') as f:
                    json.dump(all_text, f, indent=2, ensure_ascii=False)
                self.root.after(0, lambda: self.status_bar.config(text=f"Scraped text saved to {output_file}"))
                
        threading.Thread(target=scrape, daemon=True).start()
        
    def scrape_images(self):
        if not self.pdf_document:
            messagebox.showwarning("Warning", "Please open a PDF first")
            return
            
        folder = filedialog.askdirectory(title="Select folder to save images")
        if not folder:
            return
            
        def scrape():
            image_count = 0
            for i in range(self.total_pages):
                page = self.pdf_document[i]
                images = page.get_images()
                
                for img_index, img in enumerate(images, start=1):
                    xref = img[0]
                    base_image = self.pdf_document.extract_image(xref)
                    image_bytes = base_image["image"]
                    image_ext = base_image["ext"]
                    
                    filename = f"page_{i+1}_img_{img_index}.{image_ext}"
                    filepath = os.path.join(folder, filename)
                    
                    with open(filepath, "wb") as f:
                        f.write(image_bytes)
                    image_count += 1
                    
            self.root.after(0, lambda: messagebox.showinfo("Complete", f"Extracted {image_count} images to {folder}"))
            
        threading.Thread(target=scrape, daemon=True).start()
        
    def scrape_metadata(self):
        if not self.pdf_document:
            return
            
        metadata = self.pdf_document.metadata
        self.metadata_text.delete(1.0, tk.END)
        
        info = [
            f"File: {self.current_pdf}",
            f"Pages: {self.total_pages}",
            f"Format: PDF {self.pdf_document.metadata.get('format', 'Unknown')}",
            f"Title: {metadata.get('title', 'N/A')}",
            f"Author: {metadata.get('author', 'N/A')}",
            f"Subject: {metadata.get('subject', 'N/A')}",
            f"Creator: {metadata.get('creator', 'N/A')}",
            f"Producer: {metadata.get('producer', 'N/A')}",
            f"Creation Date: {metadata.get('creationDate', 'N/A')}",
            f"Modification Date: {metadata.get('modDate', 'N/A')}",
            f"Encryption: {'Yes' if self.pdf_document.is_encrypted else 'No'}",
        ]
        
        self.metadata_text.insert(tk.END, "\n".join(info))
        
    def push_undo_state(self):
        if self.pdf_document:
            data = self.pdf_document.write()
            self.undo_stack.append(data)
            self.redo_stack.clear()
    
    def undo_action(self):
        if not self.undo_stack:
            return
        current = self.pdf_document.write()
        self.redo_stack.append(current)
        prev_data = self.undo_stack.pop()
        self._load_from_bytes(prev_data)
        self.status_bar.config(text="Undo performed")
    
    def redo_action(self):
        if not self.redo_stack:
            return
        self.push_undo_state()
        next_data = self.redo_stack.pop()
        self._load_from_bytes(next_data)
        self.status_bar.config(text="Redo performed")
    
    def _load_from_bytes(self, data_bytes):
        # helper to reload document from bytes buffer
        if self.pdf_document:
            self.pdf_document.close()
        self.pdf_document = fitz.open(stream=data_bytes, filetype="pdf")
        self.total_pages = len(self.pdf_document)
        if self.current_page >= self.total_pages:
            self.current_page = self.total_pages - 1 if self.total_pages>0 else 0
        self.display_page()

    # ----- tab management -----
    def reorder_tabs(self):
        # show simple dialog to reorder notebook tabs
        dialog = tk.Toplevel(self.root)
        dialog.title("Reorder Sections")
        dialog.geometry("300x400")
        dialog.configure(bg=self.bg_color)
        listbox = tk.Listbox(dialog)
        tabs = self.notebook.tabs()
        for t in tabs:
            text = self.notebook.tab(t, "text")
            listbox.insert(tk.END, text)
        listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        def move_up():
            sel = listbox.curselection()
            if not sel: return
            idx = sel[0]
            if idx == 0: return
            txt = listbox.get(idx)
            listbox.delete(idx)
            listbox.insert(idx-1, txt)
            listbox.select_set(idx-1)
        def move_down():
            sel = listbox.curselection()
            if not sel: return
            idx = sel[0]
            if idx == listbox.size()-1: return
            txt = listbox.get(idx)
            listbox.delete(idx)
            listbox.insert(idx+1, txt)
            listbox.select_set(idx+1)
        btn_frame = ttk.Frame(dialog, style='Custom.TFrame')
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text="Up", command=move_up, style='Custom.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Down", command=move_down, style='Custom.TButton').pack(side=tk.LEFT, padx=5)

        def apply():
            order = [listbox.get(i) for i in range(listbox.size())]
            # reinsert tabs in new order
            current_frames = {self.notebook.tab(t,'text'): self.notebook.nametowidget(t) for t in tabs}
            for t in tabs:
                self.notebook.forget(t)
            for name in order:
                frame = current_frames.get(name)
                if frame:
                    self.notebook.add(frame, text=name)
            dialog.destroy()
            self.status_bar.config(text="Sections reordered")
        ttk.Button(dialog, text="Apply", command=apply, style='Custom.TButton').pack(pady=5)

    def _force_frame_redraw(self, frame):
        """Force all widgets in a frame to redraw thoroughly."""
        try:
            # Raise the frame to top
            frame.tkraise()
            
            # Multiple update passes
            for _ in range(3):
                frame.update()
                frame.update_idletasks()
            
            # Recursively update all descendants
            def update_descendants(widget):
                try:
                    widget.tkraise()
                    widget.update()
                    widget.update_idletasks()
                except Exception:
                    pass
                for child in widget.winfo_children():
                    update_descendants(child)
            
            update_descendants(frame)
            
            # Final updates
            frame.update()
            frame.update_idletasks()
            
        except Exception:
            pass
    
    def _make_window_draggable(self, win):
        """Make a window draggable by its title bar."""
        drag_data = {'x': 0, 'y': 0}
        
        def on_press(event):
            drag_data['x'] = event.x_root - win.winfo_x()
            drag_data['y'] = event.y_root - win.winfo_y()
        
        def on_motion(event):
            x = event.x_root - drag_data['x']
            y = event.y_root - drag_data['y']
            win.geometry(f'+{x}+{y}')
        
        # Make window draggable from title bar area (top 30 pixels)
        win.bind('<Button-1>', on_press)
        win.bind('<B1-Motion>', on_motion)





    def on_secondary_configure(self, event, entry):
        # debounce per-secondary resize and update rendering
        if not entry.get('auto_fit', True):
            return
        if entry.get('_resize_after'):
            try:
                self.root.after_cancel(entry['_resize_after'])
            except Exception:
                pass
        entry['_resize_after'] = self.root.after(100, self.display_page)
    
    def reorder_pages(self):
        if not self.pdf_document:
            return
        dialog = tk.Toplevel(self.root)
        dialog.title("Reorder Pages")
        dialog.geometry("300x400")
        dialog.configure(bg=self.bg_color)
        listbox = tk.Listbox(dialog)
        for i in range(self.total_pages):
            listbox.insert(tk.END, f"Page {i+1}")
        listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        def move_up():
            sel = listbox.curselection()
            if not sel: return
            idx = sel[0]
            if idx==0: return
            text = listbox.get(idx)
            listbox.delete(idx)
            listbox.insert(idx-1, text)
            listbox.select_set(idx-1)
        def move_down():
            sel = listbox.curselection()
            if not sel: return
            idx = sel[0]
            if idx==listbox.size()-1: return
            text = listbox.get(idx)
            listbox.delete(idx)
            listbox.insert(idx+1, text)
            listbox.select_set(idx+1)
        btn_frame = ttk.Frame(dialog, style='Custom.TFrame')
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text="Up", command=move_up, style='Custom.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Down", command=move_down, style='Custom.TButton').pack(side=tk.LEFT, padx=5)
        
        def apply_order():
            order = [int(listbox.get(i).split()[1]) - 1 for i in range(listbox.size())]
            self.push_undo_state()
            new_doc = fitz.open()
            for p in order:
                new_doc.insert_pdf(self.pdf_document, from_page=p, to_page=p)
            self.pdf_document = new_doc
            self.total_pages = len(self.pdf_document)
            self.current_page = 0
            self.display_page()
            dialog.destroy()
            self.status_bar.config(text="Pages reordered")
        ttk.Button(dialog, text="Apply", command=apply_order, style='Custom.TButton').pack(pady=5)
    
    def open_second_window(self):
        # opens a secondary preview window
        win = tk.Toplevel(self.root)
        win.title("PDF Scraper Pro - Secondary View")
        win.geometry("800x600")
        canvas = tk.Canvas(win, bg=self.secondary_bg, highlightthickness=0)
        canvas.pack(fill=tk.BOTH, expand=True)
        v_scroll = ttk.Scrollbar(win, orient=tk.VERTICAL, command=canvas.yview)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll = ttk.Scrollbar(win, orient=tk.HORIZONTAL, command=canvas.xview)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        entry = {'canvas': canvas, 'image_ref': None, 'auto_fit': True, '_resize_after': None, 'window': win}
        self.secondary_windows.append(entry)

        # bind configure to update this secondary view (debounced)
        try:
            canvas.bind('<Configure>', lambda e, en=entry: self.on_secondary_configure(e, en), add='+')
        except Exception:
            pass

        # refresh display for new canvas
        self.display_page()
    
    def scrape_links(self):
        if not self.pdf_document:
            messagebox.showwarning("Warning", "Please open a PDF first")
            return
        
        self.text_area.delete(1.0, tk.END)
        links_found = []
        
        for i in range(self.total_pages):
            page = self.pdf_document[i]
            links = page.get_links()
            
            for link in links:
                if link['type'] == fitz.PDF_ANNOT_LINK:
                    uri = link.get('uri')
                    if uri:
                        links_found.append(f"Page {i + 1}: {uri}")
                    
        if links_found:
            self.text_area.insert(tk.END, "Links found in PDF:\n\n")
            self.text_area.insert(tk.END, "\n".join(links_found))
        else:
            self.text_area.insert(tk.END, "No links found in PDF")
            
    def scrape_tables(self):
        if not self.pdf_document:
            return
            
        # Simple table detection based on text layout
        self.text_area.delete(1.0, tk.END)
        
        for i in range(self.total_pages):
            page = self.pdf_document[i]
            tabs = page.find_tables()
            
            if tabs.tables:
                self.text_area.insert(tk.END, f"\n--- Page {i+1} Tables ---\n")
                for j, table in enumerate(tabs.tables, 1):
                    df = table.to_pandas()
                    self.text_area.insert(tk.END, f"\nTable {j}:\n{df.to_string()}\n")
                    
    def scrape_annotations(self):
        if not self.pdf_document:
            return
            
        self.text_area.delete(1.0, tk.END)
        
        for i in range(self.total_pages):
            page = self.pdf_document[i]
            annots = list(page.annots())
            
            if annots:
                self.text_area.insert(tk.END, f"\n--- Page {i+1} Annotations ---\n")
                for annot in annots:
                    info = annot.info
                    self.text_area.insert(tk.END, f"Type: {annot.type[1]}, Content: {info.get('content', 'N/A')}\n")
                    
    def scrape_form_fields(self):
        if not self.pdf_document:
            return
            
        self.text_area.delete(1.0, tk.END)
        
        for i in range(self.total_pages):
            page = self.pdf_document[i]
            widgets = page.widgets()
            
            if widgets:
                self.text_area.insert(tk.END, f"\n--- Page {i+1} Form Fields ---\n")
                for widget in widgets:
                    self.text_area.insert(tk.END, f"Field: {widget.field_name}, Type: {widget.field_type}, Value: {widget.field_value}\n")
                    
    def batch_scrape_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing PDFs")
        if not folder:
            return
            
        output_folder = filedialog.askdirectory(title="Select output folder")
        if not output_folder:
            return
            
        def process():
            results = []
            for filename in os.listdir(folder):
                if filename.endswith('.pdf'):
                    filepath = os.path.join(folder, filename)
                    try:
                        doc = fitz.open(filepath)
                        text = ""
                        for page in doc:
                            text += page.get_text()
                        doc.close()
                        
                        results.append({
                            'filename': filename,
                            'text': text[:1000]  # First 1000 chars
                        })
                    except Exception as e:
                        results.append({
                            'filename': filename,
                            'error': str(e)
                        })
                        
            output_file = os.path.join(output_folder, 'batch_scrape_results.json')
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2)
                
            self.root.after(0, lambda: messagebox.showinfo("Complete", f"Batch scrape saved to {output_file}"))
            
        threading.Thread(target=process, daemon=True).start()
        
    def rotate_page(self, degrees):
        if not self.pdf_document:
            return
        self.push_undo_state()
        page = self.pdf_document[self.current_page]
        page.set_rotation(page.rotation + degrees)
        self.display_page()
        self.status_bar.config(text=f"Rotated page {self.current_page + 1} by {degrees}°")
        
    def delete_page(self):
        if not self.pdf_document or self.total_pages <= 1:
            messagebox.showwarning("Warning", "Cannot delete the only page")
            return
            
        if messagebox.askyesno("Confirm", f"Delete page {self.current_page + 1}?"):
            self.push_undo_state()
            self.pdf_document.delete_page(self.current_page)
            self.total_pages -= 1
            
            if self.current_page >= self.total_pages:
                self.current_page = self.total_pages - 1
                
            self.display_page()
            self.status_bar.config(text=f"Deleted page. Remaining: {self.total_pages}")
            
    def extract_pages(self):
        if not self.pdf_document:
            return
        self.push_undo_state()
        # Simple dialog for page range
        dialog = tk.Toplevel(self.root)
        dialog.title("Extract Pages")
        dialog.geometry("300x150")
        dialog.configure(bg=self.bg_color)
        
        ttk.Label(dialog, text="Page range (e.g., 1-3,5,7-9):", style='Custom.TLabel').pack(pady=10)
        entry = ttk.Entry(dialog, style='Custom.TEntry')
        entry.pack(pady=5)
        
        def do_extract():
            range_str = entry.get()
            # Parse range string and extract
            # Implementation details...
            dialog.destroy()
            
        ttk.Button(dialog, text="Extract", command=do_extract).pack(pady=10)
        
    def merge_pdfs(self):
        files = filedialog.askopenfilenames(
            title="Select PDFs to merge",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if len(files) < 2:
            messagebox.showwarning("Warning", "Select at least 2 PDFs")
            return
            
        output = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if output:
            try:
                merged = fitz.open()
                for file in files:
                    with fitz.open(file) as doc:
                        merged.insert_pdf(doc)
                merged.save(output)
                merged.close()
                messagebox.showinfo("Success", f"Merged PDFs saved to {output}")
            except Exception as e:
                messagebox.showerror("Error", str(e))
                
    def split_pdf(self):
        if not self.pdf_document:
            return
            
        split_point = tk.simpledialog.askinteger(
            "Split PDF",
            f"Enter page number to split at (1-{self.total_pages-1}):",
            minvalue=1,
            maxvalue=self.total_pages-1
        )
        
        if split_point:
            try:
                part1 = fitz.open()
                part2 = fitz.open()
                
                part1.insert_pdf(self.pdf_document, to_page=split_point-1)
                part2.insert_pdf(self.pdf_document, from_page=split_point)
                
                base_name = os.path.splitext(self.current_pdf)[0]
                part1.save(f"{base_name}_part1.pdf")
                part2.save(f"{base_name}_part2.pdf")
                
                part1.close()
                part2.close()
                
                messagebox.showinfo("Success", f"PDF split into two parts")
            except Exception as e:
                messagebox.showerror("Error", str(e))
                
    def compress_pdf(self):
        if not self.pdf_document:
            return
            
        output = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if output:
            try:
                # Save with garbage collection and compression
                self.pdf_document.save(
                    output,
                    garbage=4,
                    deflate=True,
                    clean=True
                )
                
                original_size = os.path.getsize(self.current_pdf)
                new_size = os.path.getsize(output)
                reduction = ((original_size - new_size) / original_size) * 100
                
                messagebox.showinfo(
                    "Success",
                    f"Compressed PDF saved\n"
                    f"Original: {original_size/1024:.1f} KB\n"
                    f"New: {new_size/1024:.1f} KB\n"
                    f"Reduction: {reduction:.1f}%"
                )
            except Exception as e:
                messagebox.showerror("Error", str(e))
                
    def add_watermark(self):
        if not self.pdf_document:
            return
            
        watermark_text = tk.simpledialog.askstring(
            "Add Watermark",
            "Enter watermark text:"
        )
        
        if watermark_text:
            for page_num in range(self.total_pages):
                page = self.pdf_document[page_num]
                rect = page.rect
                
                # Add text watermark
                text_writer = fitz.TextWriter(page.rect)
                text_writer.append(
                    (rect.width/2, rect.height/2),
                    watermark_text,
                    fontsize=50
                )
                
                # Semi-transparent overlay
                shape = page.new_shape()
                shape.insert_textbox(
                    rect,
                    watermark_text,
                    fontsize=50,
                    color=(0.5, 0.5, 0.5),
                    overlay=True
                )
                shape.commit()
                
            self.display_page()
            messagebox.showinfo("Success", "Watermark added to all pages")
            
    def redact_text(self):
        if not self.pdf_document:
            return
        self.push_undo_state()
        # Redact current search results or selected text
        search_term = tk.simpledialog.askstring(
            "Redact Text",
            "Enter text to redact (will be blacked out):"
        )
        
        if search_term:
            count = 0
            for page_num in range(self.total_pages):
                page = self.pdf_document[page_num]
                areas = page.search_for(search_term)
                
                for area in areas:
                    page.add_redact_annot(area, fill=(0, 0, 0))
                    count += 1
                    
                page.apply_redactions()
                
            self.display_page()
            messagebox.showinfo("Success", f"Redacted {count} instances of '{search_term}'")
            
    def add_text_to_pdf(self):
        if not self.pdf_document:
            return
        self.push_undo_state()
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Text to PDF")
        dialog.geometry("400x300")
        dialog.configure(bg=self.bg_color)
        
        ttk.Label(dialog, text="Text:", style='Custom.TLabel').pack(pady=5)
        text_entry = ttk.Entry(dialog, width=40, style='Custom.TEntry')
        text_entry.pack(pady=5)
        
        ttk.Label(dialog, text="X position:", style='Custom.TLabel').pack(pady=5)
        x_entry = ttk.Entry(dialog, style='Custom.TEntry')
        x_entry.insert(0, "100")
        x_entry.pack(pady=5)
        
        ttk.Label(dialog, text="Y position:", style='Custom.TLabel').pack(pady=5)
        y_entry = ttk.Entry(dialog, style='Custom.TEntry')
        y_entry.insert(0, "100")
        y_entry.pack(pady=5)
        
        def insert():
            page = self.pdf_document[self.current_page]
            page.insert_text(
                (float(x_entry.get()), float(y_entry.get())),
                text_entry.get(),
                fontsize=12,
                color=(0, 0, 0)
            )
            self.display_page()
            dialog.destroy()
            
        ttk.Button(dialog, text="Insert", command=insert).pack(pady=20)
        
    def add_image_to_pdf(self):
        if not self.pdf_document:
            return
        self.push_undo_state()
        image_path = filedialog.askopenfilename(
            title="Select Image",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp")]
        )
        
        if image_path:
            page = self.pdf_document[self.current_page]
            rect = fitz.Rect(100, 100, 400, 400)  # Default position
            page.insert_image(rect, filename=image_path)
            self.display_page()
            
    def ocr_page(self):
        if not self.pdf_document:
            messagebox.showwarning("Warning", "Please open a PDF first")
            return

        try:
            import pytesseract
            from pytesseract import TesseractError, TesseractNotFoundError
        except Exception:
            messagebox.showerror("Missing dependency", "pytesseract is not installed.\nRun:\n  pip install pytesseract")
            return

        # apply configured tesseract_cmd if provided
        if self.tesseract_cmd:
            try:
                pytesseract.pytesseract.tesseract_cmd = self.tesseract_cmd
            except Exception:
                pass

        # Render current page at higher DPI for better OCR
        page = self.pdf_document[self.current_page]
        mat = fitz.Matrix(2.5, 2.5)
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        try:
            text = pytesseract.image_to_string(img)
        except Exception as e:
            # handle tesseract not found
            if hasattr(e, 'message') or 'Tesseract' in str(e):
                messagebox.showerror("Tesseract Error", f"Tesseract OCR not found or failed to run:\n{e}\n\nUse Tools → Configure OCR to set the tesseract executable path.")
            else:
                messagebox.showerror("OCR Error", str(e))
            return

        # Show results in text area and status
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(tk.END, f"--- OCR Page {self.current_page + 1} ---\n\n")
        self.text_area.insert(tk.END, text)
        self.extracted_text = text
        self.status_bar.config(text="OCR complete")

    def configure_ocr(self):
        # Allow user to set tesseract executable path (useful on Windows)
        path = filedialog.askopenfilename(title="Select tesseract executable (tesseract.exe)", filetypes=[("Executables", "*.exe"), ("All files", "*")])
        if path:
            self.tesseract_cmd = path
            try:
                import pytesseract
                pytesseract.pytesseract.tesseract_cmd = path
            except Exception:
                pass
            messagebox.showinfo("OCR", f"Tesseract path set to:\n{path}")
        
    def batch_process(self):
        messagebox.showinfo("Batch Process", "Batch processing dialog would open here with options for:\n- Bulk text extraction\n- Bulk image extraction\n- Bulk format conversion\n- Bulk compression")
        
    def save_pdf(self):
        if not self.pdf_document:
            return
            
        if self.current_pdf:
            try:
                self.pdf_document.save(self.current_pdf, incremental=True)
                self.status_bar.config(text="PDF saved successfully")
            except Exception as e:
                messagebox.showerror("Error", str(e))
                
    def export_text(self):
        if not self.extracted_text:
            self.extract_all_pages()
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(self.extracted_text)
            self.status_bar.config(text=f"Text exported to {file_path}")
            
    def export_images(self):
        self.scrape_images()
        
    def show_recent(self):
        if not self.history:
            messagebox.showinfo("Recent Files", "No recent files")
            return
            
        menu = tk.Menu(self.root, tearoff=0, bg=self.secondary_bg, fg=self.fg_color)
        for file in self.history[-5:]:
            menu.add_command(
                label=os.path.basename(file),
                command=lambda f=file: self.load_pdf(f)
            )
            
        # Display menu at mouse position
        try:
            menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())
        finally:
            menu.grab_release()
            
    def add_to_history(self, file_path):
        if file_path not in self.history:
            self.history.append(file_path)
            if len(self.history) > 10:
                self.history.pop(0)
                
    def update_status(self, message):
        self.status_bar.config(text=message)
        
    def toggle_theme(self):
        # Toggle between dark and light mode
        self.dark_mode = not self.dark_mode
        self.apply_theme()
        self.status_bar.config(text=f"Theme: {'Dark' if self.dark_mode else 'Light'}")
    
    def about_window(self):
        # Create About window
        about_win = tk.Toplevel(self.root)
        about_win.title("About PDF Scraper Pro")
        about_win.geometry("600x500")
        about_win.configure(bg=self.bg_color)
        about_win.resizable(False, False)
        
        # Center the window on the parent
        about_win.transient(self.root)
        about_win.grab_set()
        
        # Try to set icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'avatar.png')
            if os.path.exists(icon_path):
                icon = tk.PhotoImage(file=icon_path)
                about_win.iconphoto(False, icon)
        except Exception:
            pass
        
        # Title frame with avatar
        title_frame = ttk.Frame(about_win, style='Custom.TFrame')
        title_frame.pack(fill=tk.X, padx=20, pady=(20, 10))
        
        # Avatar
        try:
            avatar_path = os.path.join(os.path.dirname(__file__), 'avatar.png')
            if os.path.exists(avatar_path):
                avatar_img = Image.open(avatar_path)
                avatar_img.thumbnail((60, 60), Image.Resampling.LANCZOS)
                avatar_photo = ImageTk.PhotoImage(avatar_img)
                avatar_label = tk.Label(title_frame, image=avatar_photo, bg=self.bg_color)
                avatar_label.image = avatar_photo
                avatar_label.pack(side=tk.LEFT, padx=10)
        except Exception:
            pass
        
        # Title text
        title_label = tk.Label(title_frame, text="PDF Scraper Pro", font=('Segoe UI', 16, 'bold'), 
                              bg=self.bg_color, fg=self.accent_color)
        title_label.pack(side=tk.LEFT, padx=10)
        
        # Version
        version_label = tk.Label(about_win, text="Version 1.0", font=('Segoe UI', 10), 
                                bg=self.bg_color, fg=self.fg_color)
        version_label.pack(pady=(0, 15))
        
        # Main text area with scrollbar
        text_frame = ttk.Frame(about_win, style='Custom.TFrame')
        text_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        about_text = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, height=20, width=70,
                                              bg=self.secondary_bg, fg=self.fg_color,
                                              font=('Segoe UI', 9),
                                              insertbackground=self.fg_color,
                                              padx=10, pady=10)
        about_text.pack(fill=tk.BOTH, expand=True)
        about_text.config(state=tk.DISABLED)
        
        # About content
        about_content = """PDF SCRAPER PRO - Advanced PDF Tool

APP FUNCTIONS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📄 PDF Viewing & Navigation
  • Display and navigate PDF documents
  • Zoom in/out with percentage display
  • Fit to width and fit to page options
  • Page reordering and navigation

📝 Text Extraction
  • Extract text from current page
  • Extract text from all pages
  • Copy extracted text to clipboard
  • Support for complex PDF layouts

🔍 Search & Find
  • Search for text across PDF
  • Highlight search results
  • Navigate between matches
  • Case-sensitive search option

📊 Advanced Scraping
  • Scrape all text with formatting
  • Extract images from PDFs
  • Extract metadata (title, author, etc.)
  • Extract hyperlinks and URLs
  • Extract tables from documents

🛠️ PDF Editing
  • Rotate pages (90°, 180°, 270°)
  • Delete pages from document
  • Extract specific pages
  • Merge multiple PDFs
  • Split PDF documents
  • Compress PDF documents

🖼️ OCR & Recognition
  • Optical Character Recognition (OCR)
  • Text recognition using Tesseract
  • Batch processing of multiple PDFs
  • Configurable OCR settings

💾 File Management
  • Open PDF documents
  • Save modified PDFs
  • Export extracted text
  • Export images from PDFs
  • Recent files access

🎨 User Interface
  • Dark/Light mode toggle
  • Responsive layout with resizable panels
  • Multiple window support
  • Keyboard shortcuts for common actions
  • Status bar with current operation info

  Shortcuts:
  Ctrl+O: Open PDF
  Ctrl+S: Save PDF
  Ctrl+Z: Undo
  Ctrl+Y: Redo
  Ctrl+F: Find
  F3: Find Next
  Ctrl++: Zoom In
  Ctrl+-: Zoom Out
  Left/Right: Navigate pages
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DEVELOPER INFORMATION:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Developer: Z-Hussein with prompt engineering
Framework: Python 3 with Tkinter
Libraries:
  • PyMuPDF (fitz) - PDF processing
  • PIL/Pillow - Image handling
  • Tesseract OCR - Text recognition

GitHub: github.com/Z-Hussein/pdf-scraper-pro
License: MIT License

For support or feature requests, please contact the developer at email: zeyhdesigns@gmail.com
"""
        
        about_text.config(state=tk.NORMAL)
        about_text.insert(1.0, about_content)
        # make github link clickable
        try:
            about_text.tag_configure("link", foreground="#1a73e8", underline=True)
            about_text.tag_bind("link", "<Enter>", lambda e: about_text.config(cursor="hand2"))
            about_text.tag_bind("link", "<Leave>", lambda e: about_text.config(cursor=""))
            about_text.tag_bind("link", "<Button-1>", lambda e: webbrowser.open("https://github.com/Z-Hussein/pdf-scraper-pro"))
            start = about_text.search("github.com/Z-Hussein/pdf-scraper-pro", "1.0", tk.END)
            if start:
                end = f"{start}+{len('github.com/Z-Hussein/pdf-scraper-pro')}c"
                about_text.tag_add("link", start, end)
        except Exception:
            pass
        # make email clickable
        try:
            about_text.tag_configure("email", foreground="#1a73e8", underline=True)
            about_text.tag_bind("email", "<Enter>", lambda e: about_text.config(cursor="hand2"))
            about_text.tag_bind("email", "<Leave>", lambda e: about_text.config(cursor=""))
            about_text.tag_bind("email", "<Button-1>", lambda e: webbrowser.open("mailto:zeyhdesigns@gmail.com"))
            start = about_text.search("zeyhdesigns@gmail.com", "1.0", tk.END)
            if start:
                end = f"{start}+{len('zeyhdesigns@gmail.com')}c"
                about_text.tag_add("email", start, end)
        except Exception:
            pass
        about_text.config(state=tk.DISABLED)
        
        # Close button
        button_frame = ttk.Frame(about_win, style='Custom.TFrame')
        button_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        ttk.Button(button_frame, text="Close", command=about_win.destroy, 
                  style='Custom.TButton').pack(side=tk.RIGHT)
        
    def on_closing(self):
        if self.pdf_document:
            self.pdf_document.close()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFScraperPro(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()