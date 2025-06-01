import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import datetime
import shutil

from get_data import get_pn_for, get_sdr_for, load_weight_table, get_discount,connection_type,products_for_connection_type,sizes_for_type_and_product,row_for_type_product_size,read_all_connections,get_price_per_piece
from price_calculator import calculate_total_mass, calculate_price, calculate_length_from_mass
from create_pdf import (
    generate_pdf, to_persian_digits, generate_pdf_with_added_value, generate_pdf_with_discount,
    generate_pdf_with_custom_discount, generate_pdf_with_discount_and_added_value, generate_pdf_with_custom_discount_and_added_value,generate_connection_invoice_pdf
)

import os
import json
import csv
import pathlib
import glob
import time
import sys

class InvoiceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.customer_name_var = tk.StringVar()
        self.action_frame = None
        # --- Moved: Checkbox state and added-value/discount variables ---
        self.include_added_var = tk.BooleanVar(value=False)
        self.added_value_var = tk.StringVar(value="0.00")
        self.include_discount_var = tk.BooleanVar(value=False)
        self.discount_value_var = tk.StringVar(value="0.00")
        self.custom_discount_var = tk.StringVar(value="")
        # --- Dark mode variable ---
        self.dark_mode_var = tk.BooleanVar(value=False)
        # --- Begin expanded menu bar setup ---
        menubar = tk.Menu(self)
        # --- File Menu ---
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Invoice", command=self.clear_all)
        file_menu.add_command(label="Change Output Directory", command=self.change_output_dir)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        # --- Edit Menu ---
        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Add Item", command=self.add_item)
        edit_menu.add_command(label="Remove Item", command=self.remove_item)
        edit_menu.add_command(label="Clear All", command=self.clear_all)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        # --- Invoice Menu ---
        invoice_menu = tk.Menu(menubar, tearoff=0)
        invoice_menu.add_command(label="Generate Invoice", command=self.generate_invoice)
        menubar.add_cascade(label="Invoice", menu=invoice_menu)
        # --- Options Menu ---
        options_menu = tk.Menu(menubar, tearoff=0)
        options_menu.add_checkbutton(
            label="Include Added Value (10%)",
            variable=self.include_added_var,
            command=self.on_toggle_added
        )
        options_menu.add_checkbutton(
            label="Include Discount",
            variable=self.include_discount_var,
            command=self.on_toggle_discount
        )
        menubar.add_cascade(label="Options", menu=options_menu)
        # --- Help Menu ---
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=lambda: messagebox.showinfo("About", "PolyGharb Invoice Generator\nv1.0"))
        help_menu.add_command(label="User Guide", command=lambda: messagebox.showinfo("User Guide", "For help, contact support or see documentation."))
        menubar.add_cascade(label="Help", menu=help_menu)
        self.config(menu=menubar)
        # --- End expanded menu bar setup ---
        self.title("Invoice Generator")
        
        # --- Notebook for tabs ---
        self.notebook = ttk.Notebook(self)

        self.standard_invoice_tab_frame = ttk.Frame(self.notebook)
        self.connection_pipe_tab_frame = ttk.Frame(self.notebook)

        self.notebook.add(self.standard_invoice_tab_frame, text='Standard Invoice')
        self.notebook.add(self.connection_pipe_tab_frame, text='Connection Pipes')
        self.notebook.pack(expand=True, fill='both', padx=5, pady=5)

        self.load_config()
        # defer sizing until after widgets are created
        self.standard_items = [] # Renamed from self.items
        
        self.create_standard_invoice_tab(self.standard_invoice_tab_frame)
        self.create_connection_pipe_tab(self.connection_pipe_tab_frame)
        # enforce appropriate initial and minimum size
        self.update_idletasks()
        self.minsize(900, 500)
        self.resizable(True, True)

    def load_config(self):
        config_path = os.path.join(os.path.expanduser("~"), ".invoice_app_config.json")
        self.config_file = config_path
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.output_dir = data.get("output_dir", os.path.join(os.path.dirname(__file__), "خروجی"))
            except Exception:
                self.output_dir = os.path.join(os.path.dirname(__file__), "خروجی")
        else:
            self.output_dir = os.path.join(os.path.dirname(__file__), "خروجی")
        os.makedirs(self.output_dir, exist_ok=True)
        # Persistent counter file for invoice numbers
        self.counter_file = os.path.join(os.path.expanduser("~"), ".invoice_app_counter.json")

    def validate_numeric(self, P):
        if P == "":
            return True
        try:
            float(P)
            return True
        except ValueError:
            return False

    def validate_custom_discount(self, P):
        """
        Validates that P is a float between 1 and 100 (inclusive), or empty.
        """
        if P == "":
            return True
        try:
            val = float(P)
            return 1.0 <= val <= 100.0
        except ValueError:
            return False

    def create_standard_invoice_tab(self, parent_frame):
        # Customer input
        customer_frame = tk.Frame(parent_frame)
        customer_frame.pack(pady=10, fill='x', padx=10)
        tk.Label(customer_frame, text="Customer Name:").pack(side='left')
        self.customer_entry = tk.Entry(customer_frame, textvariable=self.customer_name_var)
        self.customer_entry.pack(side='left', fill='x', expand=True, padx=5)
        self.customer_entry.bind("<KeyRelease>", lambda e: self.update_subtotal())
        self.customer_entry.bind("<FocusOut>", lambda e: self.update_subtotal())

        # Load default invoice number from persistent counter file
        try:
            with open(self.counter_file, "r", encoding="utf-8") as cf:
                data = json.load(cf)
                last_used = int(data.get("counter", 0))
        except Exception:
            last_used = 0
        default_inv_num = str(last_used + 1)

        # Invoice number input
        invoice_frame = tk.Frame(parent_frame)
        invoice_frame.pack(pady=5, fill='x', padx=10)
        tk.Label(invoice_frame, text="Invoice Number:").pack(side='left')
        self.invoice_entry = tk.Entry(invoice_frame)
        self.invoice_entry.insert(0, default_inv_num)
        self.invoice_entry.pack(side='left', fill='x', expand=True, padx=5)

        # Item detail inputs
        item_frame = tk.LabelFrame(parent_frame, text="Item Details")
        item_frame.pack(fill='x', padx=10, pady=5)
        # Allow entry columns in item_frame to expand/contract
        for col in (1, 3, 5, 7):
            item_frame.grid_columnconfigure(col, weight=1)
        
        labels = [
            ("Pipe Grade", "grade"),
            ("PN", "pn"),
            ("SDR", "sdr"),
            ("Diameter (mm)", "diameter"),
            ("Length (m)", "length"),
            ("Total Mass (kg)", "total_mass"),
            ("Price per kg", "price_per_kg"),
            ("Total Price", "total_price"),
        ]
        self.standard_entries = {} # Renamed from self.entries
        for idx, (label_text, key) in enumerate(labels):
            row = idx // 4
            col = (idx % 4) * 2
            tk.Label(item_frame, text=f"{label_text}:").grid(row=row, column=col, sticky='e', padx=2, pady=2)
            if key in ("grade", "pn", "sdr", "diameter"):
                entry = ttk.Combobox(item_frame, values=[], state="readonly")
            else:
                entry = tk.Entry(item_frame)
            entry.grid(row=row, column=col+1, sticky='w', padx=2, pady=2)
            self.standard_entries[key] = entry

        # Configure validation for numeric-only entries
        numeric_vcmd = (self.register(self.validate_numeric), '%P')
        # Validation command for custom discount percentage (1-100)
        custom_vcmd = (self.register(self.validate_custom_discount), '%P')
        for key in ("length", "total_mass", "price_per_kg", "total_price"):
            self.standard_entries[key].config(validate="key", validatecommand=numeric_vcmd)

        # Monitor widgets to keep the Add Item button enabled/disabled correctly
        for key, widget in self.standard_entries.items():
            if isinstance(widget, ttk.Combobox):
                widget.bind("<<ComboboxSelected>>", self.update_add_button_state, add="+")
            else:
                widget.bind("<KeyRelease>", self.update_add_button_state)
                widget.bind("<FocusOut>", self.update_add_button_state, add="+")
        
        # Load dropdown data for grade, SDR, PN, and Diameter
        self.load_series_data()
        self.standard_entries["grade"]["values"] = self.grades
        self.standard_entries["grade"].bind("<<ComboboxSelected>>", self.on_grade_selected)
        self.standard_entries["grade"].bind("<<ComboboxSelected>>", self.update_add_button_state, add="+")
        self.standard_entries["sdr"].bind("<<ComboboxSelected>>", self.on_sdr_selected)
        self.standard_entries["sdr"].bind("<<ComboboxSelected>>", self.update_add_button_state, add="+")
        self.standard_entries["pn"].bind("<<ComboboxSelected>>", self.on_pn_selected)
        self.standard_entries["pn"].bind("<<ComboboxSelected>>", self.update_add_button_state, add="+")
        self.load_diameter_data()
        self.standard_entries["diameter"]["values"] = self.diameters
        self.standard_entries["diameter"].bind("<<ComboboxSelected>>", self.on_diameter_changed)
        self.standard_entries["diameter"].bind("<<ComboboxSelected>>", self.update_add_button_state, add="+")

        # Bind length and total_mass for auto calculation, live as user types
        self.standard_entries["length"].bind("<KeyRelease>", self.on_length_changed)
        self.standard_entries["length"].bind("<FocusOut>", self.on_length_changed)
        self.standard_entries["total_mass"].bind("<KeyRelease>", self.on_mass_changed)
        self.standard_entries["total_mass"].bind("<FocusOut>", self.on_mass_changed)
        # Bind price_per_kg and total_price for auto calculation, live as user types
        self.standard_entries["price_per_kg"].bind("<KeyRelease>", self.on_price_changed)
        self.standard_entries["price_per_kg"].bind("<KeyRelease>", self.update_add_button_state, add="+")
        self.standard_entries["price_per_kg"].bind("<FocusOut>", self.on_price_changed)
        self.standard_entries["price_per_kg"].bind("<FocusOut>", self.update_add_button_state, add="+")
        self.standard_entries["total_price"].bind("<KeyRelease>", self.on_total_price_changed)
        self.standard_entries["total_price"].bind("<KeyRelease>", self.update_add_button_state, add="+")
        self.standard_entries["total_price"].bind("<FocusOut>", self.on_total_price_changed)
        self.standard_entries["total_price"].bind("<FocusOut>", self.update_add_button_state, add="+")

        # Checkbox state and added-value display variables
        # (Now initialized in __init__ before menu bar)

        # Add dark mode toggle to options menu
        # Find the menubar and options_menu again (they are local in __init__, so we must reconstruct here)
        menubar = self.nametowidget(self.winfo_toplevel().cget("menu"))
        options_menu = None
        for ix in range(menubar.index("end")+1):
            if menubar.type(ix) == "cascade" and menubar.entrycget(ix, "label") == "Options":
                options_menu = menubar.nametowidget(menubar.entrycget(ix, "menu"))
                break
        if options_menu is not None:
            options_menu.add_checkbutton(
                label="Dark Mode",
                variable=self.dark_mode_var,
                command=self.on_toggle_dark_mode
            )

        # Treeview to list added items, with item number column
        columns = ("no", "grade", "pn", "sdr", "diameter", "length", "weight_per_m", "total_mass", "price_per_kg", "total_price")
        # --- Begin treeview frame and scroll setup ---
        tree_frame = tk.Frame(parent_frame)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.standard_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=8) # Renamed from self.tree
        headings = ["No.", "Grade", "PN", "SDR", "Diameter", "Length", "Weight/m", "Total Mass", "Price/kg", "Total Price"]
        # Prepare sort directions for each column
        self.standard_sort_dirs = {col: False for col in columns} # Renamed
        # Configure headings as clickable for sorting and set column widths/anchors as specified
        for col, hd in zip(columns, headings):
            if col == "no":
                self.standard_tree.column(col, width=50, minwidth=40, stretch=False, anchor='center')
            else:
                self.standard_tree.column(col, width=90, minwidth=70, stretch=True, anchor='center')
            self.standard_tree.heading(col, text=hd, command=lambda _col=col: self.sort_by(_col)) # Will need to adapt sort_by
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.standard_tree.yview)
        self.standard_tree.configure(yscrollcommand=vsb.set)
        self.standard_tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        # Ensure the treeview has keyboard focus so arrow bindings fire
        self.standard_tree.focus_set()
        self.standard_tree.bind("<<TreeviewSelect>>", self.update_remove_button_state)
        # Give keyboard focus to tree whenever selection changes
        self.standard_tree.bind("<<TreeviewSelect>>", lambda event: self.standard_tree.focus_set(), add="+")
        # Bind Delete and BackSpace keys on tree to remove selected items
        self.standard_tree.bind("<Delete>", lambda event: self.remove_item(), add="+") # remove_item will need to be tab-aware
        self.standard_tree.bind("<BackSpace>", lambda event: self.remove_item(), add="+")
        self.standard_tree.bind("<KeyPress-Down>", self.on_down_pressed, add="+") # on_down_pressed will need to be tab-aware
        self.standard_tree.bind("<KeyPress-Up>", self.on_up_pressed, add="+") # on_up_pressed will need to be tab-aware
        self.standard_tree.bind("<ButtonRelease-1>", self.on_tree_blank_click, add="+") # on_tree_blank_click will need to be tab-aware

        # Explanation input
        explanation_outer_frame = tk.Frame(parent_frame) # Outer frame for padding
        explanation_outer_frame.pack(fill='x', padx=10, pady=(5,0)) # pady top 5, bottom 0

        explanation_frame = tk.LabelFrame(explanation_outer_frame, text="Explanation / Notes")
        explanation_frame.pack(fill='x', expand=True)

        self.explanation_text_widget = tk.Text(explanation_frame, height=3, wrap=tk.WORD, font=("Helvetica", 9))
        self.explanation_text_widget.pack(fill='x', expand=True, padx=5, pady=5)
        # Optional: Add a scrollbar if you expect very long explanations
        # scroll = ttk.Scrollbar(explanation_frame, command=self.explanation_text_widget.yview)
        # self.explanation_text_widget.configure(yscrollcommand=scroll.set)
        # scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Added Value display (hidden until checkbox checked)
        self.added_frame = tk.Frame(parent_frame)
        tk.Label(
            self.added_frame,
            text="Added Value (10%):",
            font=("Helvetica", 11, "bold")
        ).pack(side='left')
        tk.Label(
            self.added_frame,
            textvariable=self.added_value_var,
            font=("Helvetica", 11, "bold"),
            anchor='e'
        ).pack(side='right')
        # Do not pack added_frame now; it will appear when toggled

        # Discount display (hidden until checkbox checked)
        self.discount_frame = tk.Frame(parent_frame)
        tk.Label(
            self.discount_frame,
            text="Discount:",
            font=("Helvetica", 11, "bold")
        ).pack(side='left')
        # Custom discount percent entry (always appears to the right of the label)
        self.custom_discount_entry = tk.Entry(
            self.discount_frame,
            textvariable=self.custom_discount_var,
            width=5,
            validate="key",
            validatecommand=(self.register(self.validate_custom_discount), '%P')
        )
        self.custom_discount_entry.pack(side='left', padx=5)
        self.custom_discount_entry.bind("<KeyRelease>", lambda e: self.update_subtotal())
        tk.Label(
            self.discount_frame,
            textvariable=self.discount_value_var,
            font=("Helvetica", 11, "bold"),
            anchor='e'
        ).pack(side='right')
        # Do not pack discount_frame now; it will appear when toggled

        # Subtotal display
        self.subtotal_frame = tk.Frame(parent_frame)
        self.subtotal_frame.pack(fill='x', padx=10, pady=5)
        tk.Label(self.subtotal_frame, text="Subtotal:", font=("Helvetica", 11, "bold")).pack(side='left')
        self.subtotal_var = tk.StringVar(value="0.00")
        tk.Label(self.subtotal_frame, textvariable=self.subtotal_var, font=("Helvetica", 11, "bold"), anchor='e').pack(side='right')

        # Generate Invoice button below subtotal
        self.generate_btn_frame = tk.Frame(parent_frame)
        self.generate_btn_frame.pack(fill='x', padx=10, pady=(10, 15))
        self.generate_btn = tk.Button(
            self.generate_btn_frame,
            text="Generate Invoice",
            command=self.generate_invoice,
            font=("Helvetica", 13, "bold"),
            bg="#d3d3d3",
            fg="black",
            height=2,
        )
        self.generate_btn.pack(fill='x')


        # Initialise button state
        self.update_add_button_state()

        # Enable pressing Enter (Return) to add the item
        # This binding is on the root window, handle_add_item_on_enter will dispatch based on active tab
        self.bind('<Return>', self.handle_add_item_on_enter)

    def create_connection_pipe_tab(self, parent_frame):
        # --- State for connection tab ---
        self.connection_items = []
        self.connection_sort_dirs = {
            "no": False, "type": False, "product": False, "size": False,
            "quantity": False, "price_per_piece": False, "total_price": False
        }

        # --- Helper functions for dynamic dropdowns and calculations ---
        def on_type_selected(event=None):
            t = self.connection_entries["type"].get()
            products = products_for_connection_type(t) if t else []
            self.connection_entries["product"]["values"] = products
            self.connection_entries["product"].set('')
            self.connection_entries["size"]["values"] = []
            self.connection_entries["size"].set('')
            update_price_and_total()
            update_add_button_state()

        def on_product_selected(event=None):
            t = self.connection_entries["type"].get()
            p = self.connection_entries["product"].get()
            sizes = sizes_for_type_and_product(t, p) if t and p else []
            self.connection_entries["size"]["values"] = sizes
            self.connection_entries["size"].set('')
            update_price_and_total()
            update_add_button_state()

        def on_size_selected(event=None):
            update_price_and_total()
            update_add_button_state()

        def on_quantity_changed(event=None):
            update_price_and_total()
            update_add_button_state()

        def update_price_and_total():
            t = self.connection_entries["type"].get()
            p = self.connection_entries["product"].get()
            s = self.connection_entries["size"].get()
            price = ""
            if t and p and s:
                try:
                    val = get_price_per_piece(t, p, s)
                    if val is not None:
                        price = str(int(round(val)))
                    else:
                        price = ""
                except Exception:
                    price = ""
            self.connection_entries["price_per_piece"].config(state="normal")
            self.connection_entries["price_per_piece"].delete(0, tk.END)
            if price:
                self.connection_entries["price_per_piece"].insert(0, price)
            self.connection_entries["price_per_piece"].config(state="readonly")
            # Update total price
            qty_str = self.connection_entries["quantity"].get().strip()
            try:
                qty = float(qty_str)
            except Exception:
                qty = None
            if price and qty is not None:
                total = int(round(float(price) * qty))
                self.connection_entries["total_price"].config(state="normal")
                self.connection_entries["total_price"].delete(0, tk.END)
                self.connection_entries["total_price"].insert(0, str(total))
                self.connection_entries["total_price"].config(state="readonly")
            else:
                self.connection_entries["total_price"].config(state="normal")
                self.connection_entries["total_price"].delete(0, tk.END)
                self.connection_entries["total_price"].config(state="readonly")

        def update_add_button_state(event=None):
            t = self.connection_entries["type"].get().strip()
            p = self.connection_entries["product"].get().strip()
            s = self.connection_entries["size"].get().strip()
            qty = self.connection_entries["quantity"].get().strip()
            price = self.connection_entries["price_per_piece"].get().strip()
            total = self.connection_entries["total_price"].get().strip()
            valid = bool(t and p and s and qty and price and total)
            try:
                if float(qty) <= 0:
                    valid = False
            except Exception:
                valid = False
            if valid:
                self.connection_add_btn.config(state="normal")
            else:
                self.connection_add_btn.config(state="disabled")

        def on_down_pressed(event):
            children = self.connection_tree.get_children()
            if children and not self.connection_tree.selection():
                first = children[0]
                self.connection_tree.selection_set(first)
                self.connection_tree.focus(first)
                self.connection_tree.see(first)
                return "break"

        def on_up_pressed(event):
            children = self.connection_tree.get_children()
            if children and not self.connection_tree.selection():
                last = children[-1]
                self.connection_tree.selection_set(last)
                self.connection_tree.focus(last)
                self.connection_tree.see(last)
                return "break"

        def on_tree_blank_click(event):
            if not self.connection_tree.identify_row(event.y):
                self.connection_tree.selection_remove(self.connection_tree.selection())
                update_remove_button_state()

        # --- Customer Name and Invoice Number ---
        customer_frame = tk.Frame(parent_frame)
        customer_frame.pack(pady=10, fill='x', padx=10)
        tk.Label(customer_frame, text="Customer Name:").pack(side='left')
        self.connection_customer_entry = tk.Entry(customer_frame, textvariable=self.customer_name_var)
        self.connection_customer_entry.pack(side='left', fill='x', expand=True, padx=5)
        # Load default invoice number from persistent counter file
        try:
            with open(self.counter_file, "r", encoding="utf-8") as cf:
                data = json.load(cf)
                last_used = int(data.get("counter", 0))
        except Exception:
            last_used = 0
        default_inv_num = str(last_used + 1)
        invoice_frame = tk.Frame(parent_frame)
        invoice_frame.pack(pady=5, fill='x', padx=10)
        tk.Label(invoice_frame, text="Invoice Number:").pack(side='left')
        self.connection_invoice_entry = tk.Entry(invoice_frame)
        self.connection_invoice_entry.insert(0, default_inv_num)
        self.connection_invoice_entry.pack(side='left', fill='x', expand=True, padx=5)

        # --- Item Details ---
        item_frame = tk.LabelFrame(parent_frame, text="Item Details")
        item_frame.pack(fill='x', padx=10, pady=5)
        for col in (1, 3, 5, 7):
            item_frame.grid_columnconfigure(col, weight=1)
        # Entries: "Type", "Product", "Size", "Quantity", "Price per Piece", "Total Price"
        self.connection_entries = {}
        # Type
        tk.Label(item_frame, text="Type:").grid(row=0, column=0, sticky='e', padx=2, pady=2)
        type_values = connection_type()
        type_cb = ttk.Combobox(item_frame, values=type_values, state="readonly")
        type_cb.grid(row=0, column=1, sticky='w', padx=2, pady=2)
        self.connection_entries["type"] = type_cb
        # Product
        tk.Label(item_frame, text="Product:").grid(row=0, column=2, sticky='e', padx=2, pady=2)
        product_cb = ttk.Combobox(item_frame, values=[], state="readonly")
        product_cb.grid(row=0, column=3, sticky='w', padx=2, pady=2)
        self.connection_entries["product"] = product_cb
        # Size
        tk.Label(item_frame, text="Size:").grid(row=0, column=4, sticky='e', padx=2, pady=2)
        size_cb = ttk.Combobox(item_frame, values=[], state="readonly")
        size_cb.grid(row=0, column=5, sticky='w', padx=2, pady=2)
        self.connection_entries["size"] = size_cb
        # Quantity
        tk.Label(item_frame, text="Quantity:").grid(row=0, column=6, sticky='e', padx=2, pady=2)
        quantity_entry = tk.Entry(item_frame, width=10)
        quantity_entry.grid(row=0, column=7, sticky='w', padx=2, pady=2)
        self.connection_entries["quantity"] = quantity_entry
        # Price per Piece
        tk.Label(item_frame, text="Price per Piece:").grid(row=1, column=0, sticky='e', padx=2, pady=2)
        price_entry = tk.Entry(item_frame, width=15, state="readonly")
        price_entry.grid(row=1, column=1, sticky='w', padx=2, pady=2)
        self.connection_entries["price_per_piece"] = price_entry
        # Total Price
        tk.Label(item_frame, text="Total Price:").grid(row=1, column=2, sticky='e', padx=2, pady=2)
        total_entry = tk.Entry(item_frame, width=15, state="readonly")
        total_entry.grid(row=1, column=3, sticky='w', padx=2, pady=2)
        self.connection_entries["total_price"] = total_entry

        # Bindings for dynamic updates
        type_cb.bind("<<ComboboxSelected>>", on_type_selected)
        product_cb.bind("<<ComboboxSelected>>", on_product_selected)
        size_cb.bind("<<ComboboxSelected>>", on_size_selected)
        quantity_entry.bind("<KeyRelease>", on_quantity_changed)
        quantity_entry.bind("<FocusOut>", on_quantity_changed)

        # --- Add Item Button ---
        self.connection_add_btn = tk.Button(
            item_frame, text="Add Item", state="disabled", font=("Helvetica", 10, "bold"),
            command=self.add_connection_item_action
        )
        self.connection_add_btn.grid(row=1, column=7, padx=2, pady=2, sticky='w')

        # --- Treeview for connection items ---
        columns = ("no", "type", "product", "size", "quantity", "price_per_piece", "total_price")
        tree_frame = tk.Frame(parent_frame)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.connection_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=8)
        headings = ["No.", "Type", "Product", "Size", "Quantity", "Price per Piece", "Total Price"]
        for col, hd in zip(columns, headings):
            if col == "no":
                self.connection_tree.column(col, width=50, minwidth=40, stretch=False, anchor='center')
            else:
                self.connection_tree.column(col, width=110, minwidth=70, stretch=True, anchor='center')
            self.connection_tree.heading(col, text=hd, command=lambda _col=col: sort_by(_col))
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.connection_tree.yview)
        self.connection_tree.configure(yscrollcommand=vsb.set)
        self.connection_tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        self.connection_tree.focus_set()
        self.connection_tree.bind("<<TreeviewSelect>>", self.update_remove_button_state)
        self.connection_tree.bind("<Delete>", lambda event: remove_connection_item(), add="+")
        self.connection_tree.bind("<BackSpace>", lambda event: remove_connection_item(), add="+")
        self.connection_tree.bind("<KeyPress-Down>", on_down_pressed, add="+")
        self.connection_tree.bind("<KeyPress-Up>", on_up_pressed, add="+")
        self.connection_tree.bind("<ButtonRelease-1>", on_tree_blank_click, add="+")

        # --- Subtotal display ---
        subtotal_frame = tk.Frame(parent_frame)
        subtotal_frame.pack(fill='x', padx=10, pady=5)
        tk.Label(subtotal_frame, text="Subtotal:", font=("Helvetica", 11, "bold")).pack(side='left')
        self.connection_subtotal_var = tk.StringVar(value="0.00")
        tk.Label(subtotal_frame, textvariable=self.connection_subtotal_var, font=("Helvetica", 11, "bold"), anchor='e').pack(side='right')

        # --- Generate Invoice Button ---
        generate_btn_frame = tk.Frame(parent_frame)
        generate_btn_frame.pack(fill='x', padx=10, pady=(10, 15))
        self.connection_generate_btn = tk.Button(
            generate_btn_frame,
            text="Generate Invoice",
            command=lambda: generate_connection_invoice(),
            font=("Helvetica", 13, "bold"),
            bg="#d3d3d3",
            fg="black",
            height=2,
        )
        self.connection_generate_btn.pack(fill='x')

        # --- Item add/remove/clear logic ---
        def remove_connection_item():
            selected = self.connection_tree.selection()
            focused = self.connection_tree.focus()
            if not selected and focused:
                selected = (focused,)
            children = self.connection_tree.get_children()
            next_id = None
            if selected:
                first_id = selected[0]
                if first_id in children:
                    idx = children.index(first_id)
                    if idx < len(children) - 1:
                        next_id = children[idx + 1]
                    elif idx > 0:
                        next_id = children[idx - 1]
            if not selected:
                messagebox.showwarning("No Selection", "Please select an item to remove.")
                return
            # Remove from both list and tree
            for item_id in selected:
                idx = self.connection_tree.index(item_id)
                self.connection_tree.delete(item_id)
                self.connection_items.pop(idx)
            update_subtotal()
            update_add_button_state()
            update_remove_button_state()
            refresh_indices()
            if next_id and next_id in self.connection_tree.get_children():
                self.connection_tree.selection_set(next_id)
                self.connection_tree.focus(next_id)

        def clear_connection_entries():
            for k, widget in self.connection_entries.items():
                if isinstance(widget, ttk.Combobox):
                    widget.set('')
                else:
                    widget.config(state="normal")
                    widget.delete(0, tk.END)
                    if k in ("price_per_piece", "total_price"):
                        widget.config(state="readonly")

        def update_subtotal():
            subtotal = sum(item["total_price"] for item in self.connection_items)
            self.connection_subtotal_var.set(f"{int(subtotal):,}")

        def update_remove_button_state(event=None):
            # No-op, placeholder for possible future UI
            pass

        def refresh_indices():
            for idx, iid in enumerate(self.connection_tree.get_children(), start=1):
                self.connection_tree.set(iid, "no", idx)

        def sort_by(col):
            reverse = self.connection_sort_dirs.get(col, False)
            if col == "no":
                sorted_items = list(self.connection_items)[::-1] if reverse else list(self.connection_items)
            else:
                try:
                    sorted_items = sorted(self.connection_items, key=lambda i: i[col], reverse=reverse)
                except Exception:
                    sorted_items = sorted(self.connection_items, key=lambda i: str(i[col]), reverse=reverse)
            self.connection_items = sorted_items
            for iid in self.connection_tree.get_children():
                self.connection_tree.delete(iid)
            for idx, item in enumerate(self.connection_items, start=1):
                self.connection_tree.insert("", "end", values=(
                    idx, item["type"], item["product"], item["size"], item["quantity"],
                    f"{int(item['price_per_piece']):,}", f"{int(item['total_price']):,}"
                ))
            self.connection_sort_dirs[col] = not reverse
            update_remove_button_state()

        def generate_connection_invoice():
            if not self.connection_items:
                messagebox.showwarning("No Items", "Add at least one item before generating an invoice.")
                return
            customer = self.customer_name_var.get().strip()
            if not customer:
                messagebox.showwarning("Missing Customer", "Please enter the customer name.")
                return
            # Use persistent counter file for invoice numbers
            counter_path = self.counter_file
            try:
                with open(counter_path, "r", encoding="utf-8") as cf:
                    data = json.load(cf)
                    highest_invoice_on_record = int(data.get("counter", 0))
            except Exception:
                highest_invoice_on_record = 0
            user_entered_invoice_str = self.connection_invoice_entry.get().strip()
            if user_entered_invoice_str:
                try:
                    current_invoice_str_for_pdf = user_entered_invoice_str
                    current_invoice_int_for_pdf = int(current_invoice_str_for_pdf)
                except ValueError:
                    messagebox.showerror("Invalid Invoice Number", f"The provided invoice number '{user_entered_invoice_str}' is not a valid integer.")
                    return
            else:
                current_invoice_int_for_pdf = highest_invoice_on_record + 1
                current_invoice_str_for_pdf = str(current_invoice_int_for_pdf)
            new_highest_for_record = current_invoice_int_for_pdf
            try:
                with open(counter_path, "w", encoding="utf-8") as cf:
                    json.dump({"counter": new_highest_for_record}, cf)
                self.connection_invoice_entry.delete(0, tk.END)
                self.connection_invoice_entry.insert(0, str(new_highest_for_record + 1))
            except Exception as e:
                messagebox.showwarning("Counter File Update", f"Could not update counter file '{counter_path}':\n{e}")
            invoice_number = current_invoice_str_for_pdf
            # Prepare items for PDF
            pdf_items = []
            for it in self.connection_items:
                pdf_items.append({
                    "type": it["type"],
                    "product": it["product"],
                    "size": it["size"],
                    "quantity": it["quantity"],
                    "unit_price": it["price_per_piece"],
                    "total_price": it["total_price"],
                })
            # Use a PDF generator for connections (must be implemented elsewhere)
            try:
                from create_pdf import generate_connection_invoice_pdf
            except ImportError:
                messagebox.showerror("PDF Generation", "generate_connection_invoice_pdf not found.")
                return
            pdf_result = None
            try:
                gen_time = time.time()
                pdf_result = generate_connection_invoice_pdf(
                    customer, invoice_number, pdf_items, output_dir=self.output_dir
                )
                if pdf_result is None:
                    pattern = os.path.join(self.output_dir, f"*{invoice_number}*.pdf")
                    matches = glob.glob(pattern)
                    if matches:
                        matches.sort(key=lambda p: os.path.getmtime(p), reverse=True)
                        pdf_path = matches[0]
                    else:
                        all_pdfs = glob.glob(os.path.join(self.output_dir, "*.pdf"))
                        new_pdfs = [p for p in all_pdfs if os.path.getmtime(p) >= gen_time]
                        if new_pdfs:
                            new_pdfs.sort(key=lambda p: os.path.getmtime(p), reverse=True)
                            pdf_path = new_pdfs[0]
                        else:
                            raise ValueError(f"PDF generation failed: no PDF found in '{self.output_dir}' for invoice {invoice_number}")
                elif isinstance(pdf_result, bytes):
                    pdf_path = os.path.join(self.output_dir, f"{customer}_{invoice_number}.pdf")
                    with open(pdf_path, "wb") as f:
                        f.write(pdf_result)
                elif isinstance(pdf_result, (str, pathlib.Path)):
                    pdf_path = str(pdf_result)
                else:
                    raise ValueError(f"PDF generation failed: unexpected return type {type(pdf_result)}: {repr(pdf_result)}")
                if not os.path.exists(pdf_path):
                    raise ValueError(f"PDF generation failed: file '{pdf_path}' does not exist.")
                messagebox.showinfo("Invoice Created", f"Invoice #{invoice_number} generated successfully.\nSaved to: {pdf_path}")
            except Exception as e:
                message = str(e)
                try:
                    message += f"\nReturned: {repr(pdf_result)}"
                except Exception:
                    pass
                messagebox.showerror("PDF Generation Failed", message)

    def add_connection_item_action(self):
        """Adds an item to the connection pipes tab."""
        # Helper functions specific to this action, previously defined inside create_connection_pipe_tab
        def update_subtotal_local(): # Renamed to avoid conflict if self.update_subtotal is called
            subtotal = sum(item["total_price"] for item in self.connection_items)
            self.connection_subtotal_var.set(f"{int(subtotal):,}")

        def clear_connection_entries_local(): # Renamed
            for k, widget in self.connection_entries.items():
                if isinstance(widget, ttk.Combobox):
                    widget.set('')
                else:
                    widget.config(state="normal")
                    widget.delete(0, tk.END)
                    if k in ("price_per_piece", "total_price"):
                        widget.config(state="readonly")
        
        # update_add_button_state is already a method, but it's for the standard tab.
        # We need the connection tab's specific logic here or make it a more general method.
        # For now, let's assume the button state is handled by the UI events in create_connection_pipe_tab.

        try:
            t = self.connection_entries["type"].get().strip()
            p = self.connection_entries["product"].get().strip()
            s = self.connection_entries["size"].get().strip()
            qty_str = self.connection_entries["quantity"].get().strip()
            price_str = self.connection_entries["price_per_piece"].get().strip()
            total_str = self.connection_entries["total_price"].get().strip()
            if not (t and p and s and qty_str and price_str and total_str):
                raise ValueError("Fill all fields for the connection item.")
            qty = float(qty_str)
            price = float(price_str)
            total = float(total_str)
            item = {"type": t, "product": p, "size": s, "quantity": qty, "price_per_piece": price, "total_price": total}
            self.connection_items.append(item)
            item_no = len(self.connection_items)
            self.connection_tree.insert("", "end", values=(item_no, t, p, s, qty, f"{int(price):,}", f"{int(total):,}" ))
            update_subtotal_local()
            # clear_connection_entries_local()  # Removed per instructions
            # self.connection_add_btn state will be updated by its own event bindings
        except Exception as e:
            messagebox.showerror("Error Adding Connection Item", str(e))

    def add_item(self):
        try:
            grade = self.standard_entries["grade"].get().strip()
            pn = self.standard_entries["pn"].get().strip()
            sdr_input = self.standard_entries["sdr"].get().strip()
            diameter = float(self.standard_entries["diameter"].get().strip())
            
            # Determine SDR and PN
            if not sdr_input:
                if pn:
                    sdr = get_sdr_for(grade, pn)
                else:
                    raise ValueError("Please provide either SDR or PN.")
            else:
                sdr = float(sdr_input)
                if not pn:
                    pn = get_pn_for(grade, sdr)

            # Calculate weight per meter
            weight_per_m = load_weight_table(diameter, sdr)

            # Determine length and total mass
            length_input = self.standard_entries["length"].get().strip()
            mass_input = self.standard_entries["total_mass"].get().strip()
            if length_input:
                length = float(length_input)
                total_mass = calculate_total_mass(length, diameter, sdr)
            elif mass_input:
                total_mass = float(mass_input)
                length = calculate_length_from_mass(total_mass, diameter, sdr)
            else:
                raise ValueError("Please provide either Length or Total Mass.")

            # Determine price per kg and total price, preferring the displayed total price
            price_input = self.standard_entries["price_per_kg"].get().strip()
            total_price_input = self.standard_entries["total_price"].get().strip()
            if total_price_input:
                # Use the exactly displayed total price to avoid float drift
                total_price = float(total_price_input)
                # Derive price_per_kg to preserve entry consistency
                if price_input:
                    price_per_kg = float(price_input)
                else:
                    price_per_kg = total_price / total_mass
            elif price_input:
                price_per_kg = float(price_input)
                total_price = total_mass * price_per_kg
            else:
                raise ValueError("Please provide either Price per kg or Total Price.")

            # Store and display the item
            item = {
                "grade": grade,
                "pn": pn,
                "sdr": sdr,
                "diameter": diameter,
                "length": length,
                "weight_per_m": weight_per_m,
                "total_mass": total_mass,
                "price_per_kg": price_per_kg,
                "total_price": total_price,
            }
            self.standard_items.append(item)
            item_no = len(self.standard_items)
            self.standard_tree.insert("", "end", values=(
                item_no,
                grade,
                pn,
                sdr,
                diameter,
                length,
                round(weight_per_m, 3),
                round(total_mass, 3),
                f"{int(price_per_kg):,}",
                f"{int(total_price):,}"
            ))
            self.update_subtotal()
            # Reset button state until next valid entry set
            self.update_add_button_state()
        except Exception as e:
            messagebox.showerror("Error Adding Item", str(e))

    def handle_add_item_on_enter(self, event=None):
        """Handles the Enter key press to add an item based on the active tab."""
        assert self.notebook is not None, "Notebook has not been initialized"
        current_tab_index = self.notebook.index(self.notebook.select())
        if current_tab_index == 0: # Standard Invoice Tab
            # Check if the "Add Item" button for standard items is enabled
            # This check is implicit as add_item() will raise errors if fields are bad
            self.add_item()
        elif current_tab_index == 1: # Connection Pipes Tab
            if self.connection_add_btn.cget('state') == 'normal':
                self.add_connection_item_action()

    def remove_item(self):
        # Remove selected items from the list and treeview
        selected = self.standard_tree.selection()
        # Fallback to focused item if no explicit selection
        focused = self.standard_tree.focus()
        if not selected and focused:
            selected = (focused,)
        # Determine which item to select after deletion
        children = self.standard_tree.get_children()
        next_id = None
        if selected:
            first_id = selected[0]
            if first_id in children:
                idx = children.index(first_id)
                # Prefer the item below; if none, pick the one above
                if idx < len(children) - 1:
                    next_id = children[idx + 1]
                elif idx > 0:
                    next_id = children[idx - 1]
        if not selected:
            messagebox.showwarning("No Selection", "Please select an item to remove.")
            return
        # Remove each selected item, adjusting the items list by index
        for item_id in selected:
            idx = self.standard_tree.index(item_id) # Assumes standard_tree for now
            self.standard_tree.delete(item_id)
            self.standard_items.pop(idx) # Assumes standard_items for now
        self.update_subtotal()
        self.update_add_button_state()
        self.update_remove_button_state()
        # Re-number the No. column after removals
        self.refresh_indices()
        # Select the next item if it still exists
        if next_id and next_id in self.standard_tree.get_children():
            self.standard_tree.selection_set(next_id)
            self.standard_tree.focus(next_id)

    def generate_invoice(self):
        # Optional: At the start, destroy previous action_frame if it exists to avoid stacking buttons
        if hasattr(self, 'action_frame') and self.action_frame is not None:
            self.action_frame.destroy()
            self.action_frame = None
        if not self.standard_items:
            messagebox.showwarning("No Items", "Add at least one item before generating an invoice.")
            return
        customer = self.customer_name_var.get().strip()
        if not customer:
            messagebox.showwarning("Missing Customer", "Please enter the customer name.")
            return

        # Use persistent counter file for invoice numbers
        counter_path = self.counter_file
        # Read highest invoice number previously recorded
        try:
            with open(counter_path, "r", encoding="utf-8") as cf:
                data = json.load(cf)
                highest_invoice_on_record = int(data.get("counter", 0))
        except Exception:
            highest_invoice_on_record = 0

        user_entered_invoice_str = self.invoice_entry.get().strip()
        
        if user_entered_invoice_str:
            # User provided an invoice number
            try:
                current_invoice_str_for_pdf = user_entered_invoice_str
                current_invoice_int_for_pdf = int(current_invoice_str_for_pdf)
            except ValueError:
                messagebox.showerror("Invalid Invoice Number", 
                                     f"The provided invoice number '{user_entered_invoice_str}' is not a valid integer.")
                return
        else:
            # Auto-generate invoice number: one greater than the highest on record
            current_invoice_int_for_pdf = highest_invoice_on_record + 1
            current_invoice_str_for_pdf = str(current_invoice_int_for_pdf)

        # Determine the new highest number to save in the counter file.
        # This will be the maximum of what was on record and the current invoice number being used.
        new_highest_for_record = current_invoice_int_for_pdf

        # Save this new highest number to the counter file
        try:
            with open(counter_path, "w", encoding="utf-8") as cf:
                json.dump({"counter": new_highest_for_record}, cf)
            
            # Update the invoice entry field to show the *next* suggested invoice number
            self.invoice_entry.delete(0, tk.END)
            self.invoice_entry.insert(0, str(new_highest_for_record + 1))
        except Exception as e:
            messagebox.showwarning("Counter File Update", f"Could not update counter file '{counter_path}':\n{e}")

        invoice_number = current_invoice_str_for_pdf # Use this for the PDF
        explanation = self.explanation_text_widget.get("1.0", tk.END).strip()
        # Prepare items in the format expected by generate_pdf
        pdf_items = []
        for it in self.standard_items: # Assumes standard_items for now
            item_copy = it.copy()
            pdf_items.append({
                "diameter": item_copy["diameter"],
                "sdr": item_copy["sdr"],
                "grade": item_copy["grade"],
                "pe_grade": to_persian_digits(item_copy["grade"]),
                "length": item_copy["length"],
                "weight_per_meter": item_copy["weight_per_m"],
                "total_weight": item_copy["total_mass"],
                "price_per_kg": item_copy["price_per_kg"],
                "total_price": item_copy["total_price"],
            })
        pdf_result = None
        try:
            # Record start time to locate PDF if generate_pdf returns None
            gen_time = time.time()
            # Generate PDF based on discount, custom discount, and added-value options
            if self.include_discount_var.get():
                custom = self.custom_discount_var.get().strip()
                if custom:
                    try:
                        discount_pct = float(custom)
                    except ValueError:
                        discount_pct = 0.0
                    if self.include_added_var.get():
                        pdf_result = generate_pdf_with_custom_discount_and_added_value(
                            customer, invoice_number, pdf_items, discount_pct,
                            output_dir=self.output_dir,
                            explanation_text=explanation
                        )
                    else:
                        pdf_result = generate_pdf_with_custom_discount(
                            customer, invoice_number, pdf_items, discount_pct,
                            output_dir=self.output_dir,
                            explanation_text=explanation
                        )
                else:
                    if self.include_added_var.get():
                        pdf_result = generate_pdf_with_discount_and_added_value(
                            customer, invoice_number, pdf_items,
                            output_dir=self.output_dir,
                            explanation_text=explanation
                        )
                    else:
                        pdf_result = generate_pdf_with_discount(
                            customer, invoice_number, pdf_items,
                            output_dir=self.output_dir,
                            explanation_text=explanation
                        )
            else:
                if self.include_added_var.get():
                    pdf_result = generate_pdf_with_added_value(customer, invoice_number, pdf_items, output_dir=self.output_dir, explanation_text=explanation)
                else:
                    pdf_result = generate_pdf(customer, invoice_number, pdf_items, output_dir=self.output_dir, explanation_text=explanation)
            if pdf_result is None:
                # Attempt to locate PDF file generated in output_dir
                pattern = os.path.join(self.output_dir, f"*{invoice_number}*.pdf")
                matches = glob.glob(pattern)
                if matches:
                    # Use the most recently modified matching file
                    matches.sort(key=lambda p: os.path.getmtime(p), reverse=True)
                    pdf_path = matches[0]
                else:
                    # Fallback: any new PDF in output_dir since generation time
                    all_pdfs = glob.glob(os.path.join(self.output_dir, "*.pdf"))
                    new_pdfs = [p for p in all_pdfs if os.path.getmtime(p) >= gen_time]
                    if new_pdfs:
                        new_pdfs.sort(key=lambda p: os.path.getmtime(p), reverse=True)
                        pdf_path = new_pdfs[0]
                    else:
                        raise ValueError(f"PDF generation failed: no PDF found in '{self.output_dir}' for invoice {invoice_number}")
            elif isinstance(pdf_result, bytes):
                # Write PDF bytes to default output directory
                pdf_path = os.path.join(self.output_dir, f"{customer}_{invoice_number}.pdf")
                with open(pdf_path, "wb") as f:
                    f.write(pdf_result)
            elif isinstance(pdf_result, (str, pathlib.Path)):
                pdf_path = str(pdf_result)
            else:
                raise ValueError(f"PDF generation failed: unexpected return type {type(pdf_result)}: {repr(pdf_result)}")
            # Verify the file exists
            if not os.path.exists(pdf_path):
                raise ValueError(f"PDF generation failed: file '{pdf_path}' does not exist.")
            messagebox.showinfo("Invoice Created", f"Invoice #{invoice_number} generated successfully.\nSaved to: {pdf_path}")

            # (No buttons to open or save the invoice after generation)
            # (No clearing of customer, entries, items, or treeview here)
        except Exception as e:
            message = str(e)
            try:
                message += f"\nReturned: {repr(pdf_result)}"
            except Exception:
                pass
            messagebox.showerror("PDF Generation Failed", message)


    def save_invoice_as(self, src_path):
        from tkinter import filedialog
        dest_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if dest_path:
            import shutil
            shutil.copyfile(src_path, dest_path)
            messagebox.showinfo("Saved", f"Invoice saved to:\n{dest_path}")

    def change_output_dir(self):
        # Allow user to change the output directory
        new_dir = filedialog.askdirectory(title="Select output directory")
        if new_dir:
            self.output_dir = new_dir
            os.makedirs(self.output_dir, exist_ok=True)
            try:
                with open(self.config_file, "w", encoding="utf-8") as f:
                    json.dump({"output_dir": self.output_dir}, f)
            except Exception as e:
                messagebox.showerror("Error Saving Config", f"Couldn't save configuration:\n{e}")
            messagebox.showinfo("Output Directory Changed",
                                f"Output directory set to: {self.output_dir}")

    def update_subtotal(self):
        subtotal = sum(item["total_price"] for item in self.standard_items) # Assumes standard_items
        discount_amount = 0.0
        if self.include_discount_var.get():
            custom = self.custom_discount_var.get().strip()
            if custom:
                try:
                    pct = float(custom)
                    discount_amount = subtotal * (pct / 100.0)
                except ValueError:
                    discount_amount = 0.0
            else:
                # Load tiered discount thresholds
                base_path = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
                csv_path = os.path.join(base_path, "program files", "discount.csv")
                thresholds = []
                try:
                    with open(csv_path, newline='', encoding='utf-8') as f:
                        reader = csv.reader(f)
                        for row in reader:
                            if not row or len(row) < 2:
                                continue
                            try:
                                thr = float(row[0].strip())
                                pct = float(row[1].strip())
                            except ValueError:
                                continue
                            thresholds.append((thr, pct))
                except FileNotFoundError:
                    thresholds = []
                thresholds.sort(key=lambda x: x[0])
                for idx, (thr, pct) in enumerate(thresholds):
                    if subtotal > thr:
                        upper = thresholds[idx+1][0] if idx+1 < len(thresholds) else subtotal
                        seg = min(subtotal, upper) - thr
                        if seg > 0:
                            discount_amount += seg * pct / 100
        # If discount checkbox not checked, discount_amount remains 0.0
        # Apply discount
        subtotal_after_discount = subtotal - discount_amount
        # Update discount display (integer)
        if discount_amount:
            self.discount_value_var.set(f"{int(round(discount_amount)):,}")
        else:
            self.discount_value_var.set("0")
        # Calculate added value (10%)
        if self.include_added_var.get():
            added = subtotal_after_discount * 0.10
        else:
            added = 0.0
        total_with_adjustments = subtotal_after_discount + added
        # Update subtotal and added value displays
        self.subtotal_var.set(f"{int(total_with_adjustments):,}")
        if self.include_added_var.get():
            self.added_value_var.set(f"{int(added):,}")
        else:
            self.added_value_var.set("0.00")

    def on_toggle_discount(self):
        if self.include_discount_var.get():
            # Show discount line above subtotal
            self.discount_frame.pack(fill='x', padx=10, pady=5, before=self.subtotal_frame)
            # If added value is also enabled, ensure added_frame is shown after discount_frame but before subtotal_frame
            if self.include_added_var.get():
                self.added_frame.pack(fill='x', padx=10, pady=5, before=self.subtotal_frame)
        else:
            # Hide discount line
            self.discount_frame.pack_forget()
        # Update displayed values
        self.update_subtotal()
        # Adjust window size to accommodate the new discount line
        self.update_idletasks()
        req_w = self.winfo_reqwidth()
        req_h = self.winfo_reqheight()
        self.geometry(f"{req_w}x{req_h}")


    def on_toggle_added(self):
        if self.include_added_var.get():
            # If discount is enabled, ensure discount_frame is packed before added_frame
            if self.include_discount_var.get():
                self.discount_frame.pack(fill='x', padx=10, pady=5, before=self.subtotal_frame)
                self.added_frame.pack(fill='x', padx=10, pady=5, before=self.subtotal_frame)
            else:
                self.added_frame.pack(fill='x', padx=10, pady=5, before=self.subtotal_frame)
        else:
            self.added_frame.pack_forget()
        # If discount is enabled, ensure discount_frame is packed before added_frame and both before subtotal_frame
        if self.include_discount_var.get():
            self.discount_frame.pack(fill='x', padx=10, pady=5, before=self.subtotal_frame)
            if self.include_added_var.get():
                self.added_frame.pack(fill='x', padx=10, pady=5, before=self.subtotal_frame)
        # Update displayed values
        self.update_subtotal()
        # Adjust window size to accommodate the new added value line
        self.update_idletasks()
        req_w = self.winfo_reqwidth()
        req_h = self.winfo_reqheight()
        self.geometry(f"{req_w}x{req_h}")

    def load_series_data(self):
        # Read pipe series SDR <-> PN mapping from CSV
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
        csv_path = os.path.join(base_path, "program files", "pipe_series_sdr.csv")
        with open(csv_path, newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            header = next(reader)
            # header[1:] are SDR values
            try:
                sdrs = [float(h) for h in header[1:]]
            except ValueError:
                sdrs = header[1:]
            series = {}
            grades = []
            for row in reader:
                grade = row[0]
                grades.append(grade)
                mapping = {}
                for sdr, pn in zip(sdrs, row[1:]):
                    if pn:
                        mapping[sdr] = pn
                series[grade] = mapping
        self.grades = sorted(grades)
        self.series_data = series

    def on_grade_selected(self, event):
        grade = self.standard_entries["grade"].get()
        # Populate SDR options
        sdrs = sorted(self.series_data.get(grade, {}).keys())
        self.standard_entries["sdr"]["values"] = sdrs
        self.standard_entries["sdr"].set('')
        # Populate PN options (sorted numerically if possible)
        pns_raw = set(self.series_data.get(grade, {}).values())
        try:
            pns_sorted = sorted(pns_raw, key=lambda x: float(x))
        except ValueError:
            pns_sorted = sorted(pns_raw)
        self.standard_entries["pn"]["values"] = pns_sorted
        self.standard_entries["pn"].set('')

    def on_sdr_selected(self, event):
        grade = self.standard_entries["grade"].get()
        try:
            sdr = float(self.standard_entries["sdr"].get())
        except ValueError:
            return
        # Populate PN options based on current grade
        pns_raw = set(self.series_data.get(grade, {}).values())
        try:
            pns_sorted = sorted(pns_raw, key=lambda x: float(x))
        except Exception:
            pns_sorted = sorted(pns_raw)
        self.standard_entries["pn"]["values"] = pns_sorted
        # Set current PN based on selected SDR
        mapped_pn = self.series_data.get(grade, {}).get(sdr)
        if mapped_pn:
            self.standard_entries["pn"].set(mapped_pn)
        else:
            self.standard_entries["pn"].set('')
        # Populate Diameter options based on selected SDR
        diam_options = self.diameter_data_by_sdr.get(sdr, [])
        self.standard_entries["diameter"]["values"] = diam_options
        self.standard_entries["diameter"].set('')

    def on_pn_selected(self, event):
        grade = self.standard_entries["grade"].get()
        pn = self.standard_entries["pn"].get()
        # Populate SDR options based on current grade
        all_sdrs = sorted(self.series_data.get(grade, {}).keys())
        self.standard_entries["sdr"]["values"] = all_sdrs
        # Set current SDR based on selected PN
        matching_sdrs = [s for s, p in self.series_data.get(grade, {}).items() if p == pn]
        if matching_sdrs:
            self.standard_entries["sdr"].set(str(matching_sdrs[0]))
        else:
            self.standard_entries["sdr"].set('')
        # Populate Diameter options based on current SDR
        try:
            current_sdr = float(self.standard_entries["sdr"].get())
        except ValueError:
            return
        diam_options = self.diameter_data_by_sdr.get(current_sdr, [])
        self.standard_entries["diameter"]["values"] = diam_options
        self.standard_entries["diameter"].set('')

    def load_diameter_data(self):
        # Read available diameters from the weight table CSV and group by SDR
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
        csv_path = os.path.join(base_path, "program files", "DIN_pivot.csv")
        with open(csv_path, newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            header = next(reader)
            # Extract SDR values from header
            try:
                sdr_list = [float(h) for h in header[1:]]
            except ValueError:
                sdr_list = header[1:]
            # Prepare mapping SDR -> diameters
            diameter_data_by_sdr = {sdr: [] for sdr in sdr_list}
            all_diams = []
            for row in reader:
                # First column is diameter
                try:
                    diam = float(row[0])
                except ValueError:
                    diam = row[0]
                all_diams.append(diam)
                # For each SDR column, if weight exists, record this diameter
                for idx, sdr in enumerate(sdr_list, start=1):
                    if row[idx].strip():
                        diameter_data_by_sdr[sdr].append(diam)
        # Save full list and per-SDR mapping
        self.diameters = sorted(all_diams)
        # Ensure each list is sorted
        self.diameter_data_by_sdr = {sdr: sorted(diams) for sdr, diams in diameter_data_by_sdr.items()}

    def on_length_changed(self, event):
        # Automatically update total_mass when length is edited
        try:
            length = float(self.standard_entries["length"].get().strip())
            diameter = float(self.standard_entries["diameter"].get())
            sdr = float(self.standard_entries["sdr"].get())
            total_mass = calculate_total_mass(length, diameter, sdr)
            mass_entry = self.standard_entries["total_mass"]
            mass_entry.delete(0, tk.END)
            mass_entry.insert(0, str(round(total_mass, 3)))
            # --- Patch: Also update total_price if price_per_kg is available ---
            price_per_kg_str = self.standard_entries["price_per_kg"].get().strip()
            if price_per_kg_str:
                price_per_kg = float(price_per_kg_str)
                total_price = total_mass * price_per_kg
                tp_entry = self.standard_entries["total_price"]
                tp_entry.delete(0, tk.END)
                tp_entry.insert(0, str(int(round(total_price))))
        except Exception:
            pass
    
    def on_mass_changed(self, event):
        # Automatically update length when total_mass is edited
        try:
            total_mass = float(self.standard_entries["total_mass"].get().strip())
            diameter = float(self.standard_entries["diameter"].get())
            sdr = float(self.standard_entries["sdr"].get())
            length = calculate_length_from_mass(total_mass, diameter, sdr)
            length_entry = self.standard_entries["length"]
            length_entry.delete(0, tk.END)
            length_entry.insert(0, str(round(length, 3)))
        except Exception:
            pass
    
    def on_price_changed(self, event):
        # Automatically update total_price when price_per_kg is edited
        try:
            price_per_kg = float(self.standard_entries["price_per_kg"].get().strip())
            total_mass = float(self.standard_entries["total_mass"].get().strip())
            # Compute total price by multiplying price per kg by total mass
            total_price = total_mass * price_per_kg
            tp_entry = self.standard_entries["total_price"]
            tp_entry.delete(0, tk.END)
            tp_entry.insert(0, str(int(round(total_price))))
        except Exception:
            pass
    
    def on_total_price_changed(self, event):
        # Automatically update price_per_kg when total_price is edited
        try:
            total_price = float(self.standard_entries["total_price"].get().strip())
            # Determine total mass: use entered mass or calculate from length
            mass_input = self.standard_entries["total_mass"].get().strip()
            if mass_input:
                total_mass = float(mass_input)
            else:
                length = float(self.standard_entries["length"].get().strip())
                diameter = float(self.standard_entries["diameter"].get())
                sdr = float(self.standard_entries["sdr"].get())
                total_mass = calculate_total_mass(length, diameter, sdr)
            # Compute price per kg by dividing by mass
            price_per_kg = total_price / total_mass
            pp_entry = self.standard_entries["price_per_kg"]
            pp_entry.delete(0, tk.END)
            formatted_price = f"{price_per_kg:,.2f}"
            pp_entry.insert(0, formatted_price)
        except Exception:
            pass
    
    def on_diameter_changed(self, event):
        # Recalculate total_mass and total_price based on new diameter and existing length
        try:
            length = float(self.standard_entries["length"].get().strip())
            diameter = float(self.standard_entries["diameter"].get().strip())
            sdr = float(self.standard_entries["sdr"].get().strip())
            total_mass = calculate_total_mass(length, diameter, sdr)
            mass_entry = self.standard_entries["total_mass"]
            mass_entry.delete(0, tk.END)
            mass_entry.insert(0, str(round(total_mass, 3)))
            price_per_kg = float(self.standard_entries["price_per_kg"].get().strip())
            total_price = total_mass * price_per_kg
            tp_entry = self.standard_entries["total_price"]
            tp_entry.delete(0, tk.END)
            tp_entry.insert(0, str(int(round(total_price))))
        except Exception:
            pass
    
    def update_add_button_state(self, event=None):
        """Enable the Add Item button only when the required fields are populated."""
        grade_filled = bool(self.standard_entries["grade"].get().strip())
        diameter_filled = bool(self.standard_entries["diameter"].get().strip())
        sdr_filled = bool(self.standard_entries["sdr"].get().strip())
        pn_filled = bool(self.standard_entries["pn"].get().strip())
        length_filled = bool(self.standard_entries["length"].get().strip())
        mass_filled = bool(self.standard_entries["total_mass"].get().strip())
        price_filled = bool(self.standard_entries["price_per_kg"].get().strip())
        total_price_filled = bool(self.standard_entries["total_price"].get().strip())

        core_ok = grade_filled and diameter_filled and (sdr_filled or pn_filled)
        qty_ok = length_filled or mass_filled
        price_ok = price_filled or total_price_filled
        # Removed all references to self.add_btn

    def update_remove_button_state(self, event=None):
        pass

    def clear_all(self):
        """Clear all input fields and items."""
        # Clear customer entry (preserve invoice number)
        self.customer_entry.delete(0, tk.END)
        # Clear item detail entries
        for key, widget in self.standard_entries.items(): # Assumes standard_entries
            if isinstance(widget, ttk.Combobox):
                widget.set('')
            else:
                widget.delete(0, tk.END)
        # Clear items list and treeview
        self.standard_items.clear() # Assumes standard_items
        for item in self.standard_tree.get_children(): # Assumes standard_tree
            self.standard_tree.delete(item)
        # Clear explanation text widget
        self.explanation_text_widget.delete("1.0", tk.END)
        # Hide added value frame if shown
        if self.include_added_var.get():
            self.include_added_var.set(False)
            self.added_frame.pack_forget()
        # Update subtotal and reset button states
        self.update_subtotal()
        self.update_add_button_state()
        self.update_remove_button_state()


    def refresh_indices(self):
        """Re-number the 'No.' column sequentially."""
        for idx, iid in enumerate(self.standard_tree.get_children(), start=1): # Assumes standard_tree
            self.standard_tree.set(iid, "no", idx)

    def on_toggle_dark_mode(self):
        # Theme switching will be implemented here.
        pass
    
    def sort_by(self, col):
        """Sort items and treeview by given column."""
        reverse = self.standard_sort_dirs.get(col, False) # Assumes standard_sort_dirs
        if col == "no":
            # Toggle between original and reverse order
            sorted_items = list(self.standard_items)[::-1] if reverse else list(self.standard_items) # Assumes standard_items
        else:
            try:
                sorted_items = sorted(self.standard_items, key=lambda i: i[col], reverse=reverse)
            except Exception:
                sorted_items = sorted(self.standard_items, key=lambda i: str(i[col]), reverse=reverse)
        self.standard_items = sorted_items
        # Clear existing rows
        for iid in self.standard_tree.get_children():
            self.standard_tree.delete(iid)
        # Reinsert rows in sorted order with updated indices
        for idx, item in enumerate(self.standard_items, start=1):
            self.standard_tree.insert("", "end", values=(
                idx,
                item["grade"],
                item["pn"],
                item["sdr"],
                item["diameter"],
                item["length"],
                round(item["weight_per_m"], 3),
                round(item["total_mass"], 3),
                f"{int(item['price_per_kg']):,}",
                f"{int(item['total_price']):,}"
            ))
        # Toggle sort direction for next click
        self.standard_sort_dirs[col] = not reverse
        # Update remove button state
        self.update_remove_button_state()

    def on_down_pressed(self, event):
        children = self.standard_tree.get_children() # Assumes standard_tree
        # If nothing is selected, select the first item
        if children and not self.standard_tree.selection():
                first = children[0]
                self.standard_tree.selection_set(first)
                self.standard_tree.focus(first)
                self.standard_tree.see(first)
                return "break"
        # Otherwise allow default behavior


    def on_up_pressed(self, event):
        children = self.standard_tree.get_children() # Assumes standard_tree
        # If nothing is selected, select the last item
        if children and not self.standard_tree.selection():
                last = children[-1]
                self.standard_tree.selection_set(last)
                self.standard_tree.focus(last)
                self.standard_tree.see(last)
                return "break"
        # Otherwise allow default behavior

    def on_tree_blank_click(self, event):
        """Deselect the current selection when the user clicks on empty space in the Treeview."""
        # If the click is not on any item row, clear the current selection
        if not self.standard_tree.identify_row(event.y): # Assumes standard_tree
            self.standard_tree.selection_remove(self.standard_tree.selection())
            self.update_remove_button_state()
            
if __name__ == "__main__":
    app = InvoiceApp()
    app.mainloop()
    def on_toggle_dark_mode(self):
        """Stub for now, will apply theme next."""
        # Theme application will be handled in the next step.
        pass