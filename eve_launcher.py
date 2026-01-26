"""
EVE Manufacturing Database Launcher
A GUI interface for managing and analyzing EVE Online manufacturing and reprocessing data.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import sys
import math
import sqlite3
from pathlib import Path

# Import our modules
from calculate_reprocessing_value import (
    calculate_reprocessing_value,
    analyze_all_modules,
    format_analysis_results,
    format_reprocessing_result
)
from update_prices_db import update_prices
from update_mineral_prices import update_mineral_prices

DATABASE_FILE = "eve_manufacturing.db"


class EVELauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("EVE Manufacturing Database Launcher")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 600)
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        # Initialize exclusion database table
        self.init_exclusion_table()
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_analysis_tab()
        self.create_single_module_tab()
        self.create_price_update_tab()
        self.create_exclusions_tab()
        
        # Store last analysis results for exclusion
        self.last_analysis_results = None
        self.last_analysis_params = None
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def init_exclusion_table(self):
        """Initialize the excluded_modules table in the database"""
        if not Path(DATABASE_FILE).exists():
            return
        
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS excluded_modules (
                    module_type_id INTEGER NOT NULL,
                    module_name TEXT NOT NULL,
                    min_price REAL,
                    max_price REAL,
                    module_price_type TEXT,
                    mineral_price_type TEXT,
                    excluded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    PRIMARY KEY (module_type_id, min_price, max_price, module_price_type, mineral_price_type)
                )
            """)
            conn.commit()
        finally:
            conn.close()
    
    def create_analysis_tab(self):
        """Create the Top 30 Analysis tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Top 30 Analysis")
        
        # Parameters frame
        params_frame = ttk.LabelFrame(frame, text="Analysis Parameters", padding=10)
        params_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Row 1
        row1 = ttk.Frame(params_frame)
        row1.pack(fill=tk.X, pady=5)
        
        ttk.Label(row1, text="Yield %:").pack(side=tk.LEFT, padx=5)
        self.yield_var = tk.StringVar(value="55.0")
        ttk.Entry(row1, textvariable=self.yield_var, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row1, text="Markup %:").pack(side=tk.LEFT, padx=5)
        self.markup_var = tk.StringVar(value="10.0")
        ttk.Entry(row1, textvariable=self.markup_var, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row1, text="Modules/Batch:").pack(side=tk.LEFT, padx=5)
        self.num_modules_var = tk.StringVar(value="1")
        ttk.Entry(row1, textvariable=self.num_modules_var, width=10).pack(side=tk.LEFT, padx=5)
        
        # Row 2
        row2 = ttk.Frame(params_frame)
        row2.pack(fill=tk.X, pady=5)
        
        ttk.Label(row2, text="Reprocessing Cost %:").pack(side=tk.LEFT, padx=5)
        self.reprocessing_cost_var = tk.StringVar(value="3.37")
        ttk.Entry(row2, textvariable=self.reprocessing_cost_var, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row2, text="Min Price:").pack(side=tk.LEFT, padx=5)
        self.min_price_var = tk.StringVar(value="1.0")
        ttk.Entry(row2, textvariable=self.min_price_var, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row2, text="Max Price:").pack(side=tk.LEFT, padx=5)
        self.max_price_var = tk.StringVar(value="100000.0")
        ttk.Entry(row2, textvariable=self.max_price_var, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row2, text="Top N:").pack(side=tk.LEFT, padx=5)
        self.top_n_var = tk.StringVar(value="30")
        ttk.Entry(row2, textvariable=self.top_n_var, width=10).pack(side=tk.LEFT, padx=5)
        
        # Row 3 - Price types
        row3 = ttk.Frame(params_frame)
        row3.pack(fill=tk.X, pady=5)
        
        ttk.Label(row3, text="Module Price Type:").pack(side=tk.LEFT, padx=5)
        self.module_price_type_var = tk.StringVar(value="sell_min")
        module_price_combo = ttk.Combobox(row3, textvariable=self.module_price_type_var, 
                                         values=["buy_max", "sell_min"], 
                                         state="readonly", width=12)
        module_price_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row3, text="Mineral Price Type:").pack(side=tk.LEFT, padx=5)
        self.mineral_price_type_var = tk.StringVar(value="buy_max")
        mineral_price_combo = ttk.Combobox(row3, textvariable=self.mineral_price_type_var,
                                          values=["buy_max", "sell_min"],
                                          state="readonly", width=12)
        mineral_price_combo.pack(side=tk.LEFT, padx=5)
        
        # Run button
        run_btn = ttk.Button(params_frame, text="Run Top 30 Analysis", command=self.run_analysis)
        run_btn.pack(pady=10)
        
        # Results frame with scrollable text and buttons
        results_frame = ttk.LabelFrame(frame, text="Results", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Frame for results text and exclude buttons
        results_content_frame = ttk.Frame(results_frame)
        results_content_frame.pack(fill=tk.BOTH, expand=True)
        
        self.analysis_results = scrolledtext.ScrolledText(results_content_frame, wrap=tk.WORD, height=20)
        self.analysis_results.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Frame for exclude buttons (will be populated dynamically)
        self.exclude_buttons_frame = ttk.Frame(results_content_frame, width=150)
        self.exclude_buttons_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        self.exclude_buttons_frame.pack_propagate(False)
        
        # Store exclude buttons
        self.exclude_buttons = {}
    
    def create_single_module_tab(self):
        """Create the Single Module Reprocessing tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Single Module")
        
        # Input frame
        input_frame = ttk.LabelFrame(frame, text="Module Information", padding=10)
        input_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Module name
        module_row = ttk.Frame(input_frame)
        module_row.pack(fill=tk.X, pady=5)
        ttk.Label(module_row, text="Module Name:").pack(side=tk.LEFT, padx=5)
        self.module_name_var = tk.StringVar()
        ttk.Entry(module_row, textvariable=self.module_name_var, width=40).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Parameters row 1
        params_row1 = ttk.Frame(input_frame)
        params_row1.pack(fill=tk.X, pady=5)
        
        ttk.Label(params_row1, text="Yield %:").pack(side=tk.LEFT, padx=5)
        self.single_yield_var = tk.StringVar(value="55.0")
        ttk.Entry(params_row1, textvariable=self.single_yield_var, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(params_row1, text="Markup %:").pack(side=tk.LEFT, padx=5)
        self.single_markup_var = tk.StringVar(value="10.0")
        ttk.Entry(params_row1, textvariable=self.single_markup_var, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(params_row1, text="Number of Modules:").pack(side=tk.LEFT, padx=5)
        self.single_num_modules_var = tk.StringVar(value="1")
        ttk.Entry(params_row1, textvariable=self.single_num_modules_var, width=10).pack(side=tk.LEFT, padx=5)
        
        # Parameters row 2
        params_row2 = ttk.Frame(input_frame)
        params_row2.pack(fill=tk.X, pady=5)
        
        ttk.Label(params_row2, text="Reprocessing Cost %:").pack(side=tk.LEFT, padx=5)
        self.single_reprocessing_cost_var = tk.StringVar(value="3.37")
        ttk.Entry(params_row2, textvariable=self.single_reprocessing_cost_var, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(params_row2, text="Module Price Type:").pack(side=tk.LEFT, padx=5)
        self.single_module_price_type_var = tk.StringVar(value="buy_max")
        single_module_price_combo = ttk.Combobox(params_row2, textvariable=self.single_module_price_type_var,
                                                 values=["buy_max", "sell_min"],
                                                 state="readonly", width=12)
        single_module_price_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(params_row2, text="Mineral Price Type:").pack(side=tk.LEFT, padx=5)
        self.single_mineral_price_type_var = tk.StringVar(value="buy_max")
        single_mineral_price_combo = ttk.Combobox(params_row2, textvariable=self.single_mineral_price_type_var,
                                                  values=["buy_max", "sell_min"],
                                                  state="readonly", width=12)
        single_mineral_price_combo.pack(side=tk.LEFT, padx=5)
        
        # Buttons frame
        buttons_frame = ttk.Frame(input_frame)
        buttons_frame.pack(pady=10)
        
        calc_btn = ttk.Button(buttons_frame, text="Calculate Reprocessing Value", command=self.calculate_single_module)
        calc_btn.pack(side=tk.LEFT, padx=5)
        
        self.edit_quantities_btn = ttk.Button(buttons_frame, text="Edit Quantities", command=self.edit_quantities, state=tk.DISABLED)
        self.edit_quantities_btn.pack(side=tk.LEFT, padx=5)
        
        # Store last calculation result for editing
        self.last_calculation_result = None
        
        # Results frame
        results_frame = ttk.LabelFrame(frame, text="Results", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.single_module_results = scrolledtext.ScrolledText(results_frame, wrap=tk.WORD, height=25)
        self.single_module_results.pack(fill=tk.BOTH, expand=True)
    
    def create_price_update_tab(self):
        """Create the Price Update tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Price Updates")
        
        # Info frame
        info_frame = ttk.LabelFrame(frame, text="Information", padding=10)
        info_frame.pack(fill=tk.X, padx=10, pady=10)
        
        info_text = """
Price Update Options:

1. Update All Prices: Updates prices for all items in the database.
   This may take several minutes depending on the number of items.

2. Update Mineral Prices Only: Updates prices only for:
   - Basic minerals (Tritanium, Pyerite, Mexallon, Isogen, Nocxium, Zydrine, Megacyte, Morphite)
   - Mutaplasmid residues
   - Other specified materials
   
   This is much faster and recommended for regular updates.
        """
        ttk.Label(info_frame, text=info_text.strip(), justify=tk.LEFT).pack(anchor=tk.W)
        
        # Buttons frame
        buttons_frame = ttk.Frame(frame)
        buttons_frame.pack(fill=tk.X, padx=10, pady=20)
        
        update_all_btn = ttk.Button(buttons_frame, text="Update All Prices", 
                                    command=self.update_all_prices, width=30)
        update_all_btn.pack(side=tk.LEFT, padx=10, expand=True)
        
        update_minerals_btn = ttk.Button(buttons_frame, text="Update Mineral Prices Only",
                                        command=self.update_mineral_prices_only, width=30)
        update_minerals_btn.pack(side=tk.LEFT, padx=10, expand=True)
        
        # Log frame
        log_frame = ttk.LabelFrame(frame, text="Update Log", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.price_update_log = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=20)
        self.price_update_log.pack(fill=tk.BOTH, expand=True)
    
    def create_exclusions_tab(self):
        """Create the Excluded Modules management tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Excluded Modules")
        
        # Info frame
        info_frame = ttk.LabelFrame(frame, text="Information", padding=10)
        info_frame.pack(fill=tk.X, padx=10, pady=10)
        
        info_text = """
Excluded modules are filtered from Top 30 Analysis results based on search parameters.
Each exclusion is tied to specific price ranges and price types, so a module excluded
for one search may still appear in searches with different parameters.
        """
        ttk.Label(info_frame, text=info_text.strip(), justify=tk.LEFT, wraplength=700).pack(anchor=tk.W)
        
        # Buttons frame
        buttons_frame = ttk.Frame(frame, padding=10)
        buttons_frame.pack(fill=tk.X, padx=10, pady=5)
        
        refresh_btn = ttk.Button(buttons_frame, text="Refresh List", command=self.refresh_exclusions_list)
        refresh_btn.pack(side=tk.LEFT, padx=5)
        
        clear_all_btn = ttk.Button(buttons_frame, text="Clear All Exclusions", command=self.clear_all_exclusions)
        clear_all_btn.pack(side=tk.LEFT, padx=5)
        
        # Table frame
        table_frame = ttk.LabelFrame(frame, text="Excluded Modules", padding=10)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create treeview
        columns = ('Module Name', 'Type ID', 'Min Price', 'Max Price', 'Module Price Type', 'Mineral Price Type', 'Excluded At')
        self.exclusions_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)
        
        # Configure columns
        self.exclusions_tree.heading('Module Name', text='Module Name')
        self.exclusions_tree.heading('Type ID', text='Type ID')
        self.exclusions_tree.heading('Min Price', text='Min Price')
        self.exclusions_tree.heading('Max Price', text='Max Price')
        self.exclusions_tree.heading('Module Price Type', text='Module Price Type')
        self.exclusions_tree.heading('Mineral Price Type', text='Mineral Price Type')
        self.exclusions_tree.heading('Excluded At', text='Excluded At')
        
        self.exclusions_tree.column('Module Name', width=250)
        self.exclusions_tree.column('Type ID', width=80, anchor=tk.E)
        self.exclusions_tree.column('Min Price', width=100, anchor=tk.E)
        self.exclusions_tree.column('Max Price', width=100, anchor=tk.E)
        self.exclusions_tree.column('Module Price Type', width=120)
        self.exclusions_tree.column('Mineral Price Type', width=120)
        self.exclusions_tree.column('Excluded At', width=150)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.exclusions_tree.yview)
        self.exclusions_tree.configure(yscrollcommand=scrollbar.set)
        
        self.exclusions_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Action buttons frame
        action_frame = ttk.Frame(frame, padding=10)
        action_frame.pack(fill=tk.X, padx=10, pady=5)
        
        remove_btn = ttk.Button(action_frame, text="Remove Selected", command=self.remove_selected_exclusion)
        remove_btn.pack(side=tk.LEFT, padx=5)
        
        # Load exclusions on startup
        self.refresh_exclusions_list()
    
    def refresh_exclusions_list(self):
        """Refresh the excluded modules list"""
        # Clear existing items
        for item in self.exclusions_tree.get_children():
            self.exclusions_tree.delete(item)
        
        if not Path(DATABASE_FILE).exists():
            return
        
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT module_type_id, module_name, min_price, max_price, 
                       module_price_type, mineral_price_type, excluded_at
                FROM excluded_modules
                ORDER BY excluded_at DESC
            """)
            results = cursor.fetchall()
            
            for row in results:
                module_type_id, module_name, min_price, max_price, module_price_type, mineral_price_type, excluded_at = row
                self.exclusions_tree.insert('', tk.END, values=(
                    module_name,
                    module_type_id,
                    f"{min_price:,.2f}",
                    f"{max_price:,.2f}",
                    module_price_type,
                    mineral_price_type,
                    excluded_at
                ))
        finally:
            conn.close()
    
    def remove_selected_exclusion(self):
        """Remove selected exclusion(s)"""
        selected = self.exclusions_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an exclusion to remove")
            return
        
        if not messagebox.askyesno("Confirm", f"Remove {len(selected)} exclusion(s)?"):
            return
        
        if not Path(DATABASE_FILE).exists():
            messagebox.showerror("Error", "Database file not found")
            return
        
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            cursor = conn.cursor()
            for item in selected:
                values = self.exclusions_tree.item(item, 'values')
                module_type_id = int(values[1])
                min_price = float(values[2].replace(',', ''))
                max_price = float(values[3].replace(',', ''))
                module_price_type = values[4]
                mineral_price_type = values[5]
                
                cursor.execute("""
                    DELETE FROM excluded_modules
                    WHERE module_type_id = ? AND min_price = ? AND max_price = ?
                    AND module_price_type = ? AND mineral_price_type = ?
                """, (module_type_id, min_price, max_price, module_price_type, mineral_price_type))
            
            conn.commit()
            messagebox.showinfo("Success", f"Removed {len(selected)} exclusion(s)")
            self.refresh_exclusions_list()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to remove exclusion: {str(e)}")
        finally:
            conn.close()
    
    def clear_all_exclusions(self):
        """Clear all exclusions"""
        if not messagebox.askyesno("Confirm", "Clear ALL exclusions? This cannot be undone."):
            return
        
        if not Path(DATABASE_FILE).exists():
            messagebox.showerror("Error", "Database file not found")
            return
        
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM excluded_modules")
            conn.commit()
            messagebox.showinfo("Success", "All exclusions cleared")
            self.refresh_exclusions_list()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to clear exclusions: {str(e)}")
        finally:
            conn.close()
    
    def get_float(self, var, default=0.0):
        """Safely get float value from StringVar"""
        try:
            return float(var.get())
        except (ValueError, tk.TclError):
            return default
    
    def get_int(self, var, default=0):
        """Safely get int value from StringVar"""
        try:
            return int(var.get())
        except (ValueError, tk.TclError):
            return default
    
    def run_analysis(self):
        """Run the top 30 analysis in a separate thread"""
        self.status_var.set("Running analysis...")
        self.analysis_results.delete(1.0, tk.END)
        self.analysis_results.insert(tk.END, "Running analysis... This may take several minutes.\n\n")
        self.root.update()
        
        def analyze():
            try:
                yield_percent = self.get_float(self.yield_var, 55.0)
                markup_percent = self.get_float(self.markup_var, 10.0)
                num_modules = self.get_int(self.num_modules_var, 100)
                reprocessing_cost = self.get_float(self.reprocessing_cost_var, 3.37)
                min_price = self.get_float(self.min_price_var, 1.0)
                max_price = self.get_float(self.max_price_var, 100000.0)
                top_n = self.get_int(self.top_n_var, 30)
                module_price_type = self.module_price_type_var.get()
                mineral_price_type = self.mineral_price_type_var.get()
                
                # Get excluded modules for this search
                excluded_modules = self.get_excluded_modules(
                    min_price, max_price, module_price_type, mineral_price_type
                )
                
                results = analyze_all_modules(
                    yield_percent=yield_percent,
                    buy_order_markup_percent=markup_percent,
                    num_modules=num_modules,
                    reprocessing_cost_percent=reprocessing_cost,
                    module_price_type=module_price_type,
                    mineral_price_type=mineral_price_type,
                    min_module_price=min_price,
                    max_module_price=max_price,
                    top_n=top_n,
                    excluded_module_ids=excluded_modules
                )
                
                # Store results and parameters for exclusion
                self.last_analysis_results = results
                self.last_analysis_params = {
                    'min_price': min_price,
                    'max_price': max_price,
                    'module_price_type': module_price_type,
                    'mineral_price_type': mineral_price_type
                }
                
                formatted = format_analysis_results(results)
                
                self.analysis_results.delete(1.0, tk.END)
                self.analysis_results.insert(tk.END, formatted)
                
                # Create exclude buttons for each result
                self.create_exclude_buttons(results)
                
                self.status_var.set("Analysis complete!")
                
            except Exception as e:
                self.analysis_results.delete(1.0, tk.END)
                self.analysis_results.insert(tk.END, f"Error: {str(e)}\n")
                self.status_var.set("Error occurred")
                messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        
        thread = threading.Thread(target=analyze, daemon=True)
        thread.start()
    
    def get_excluded_modules(self, min_price, max_price, module_price_type, mineral_price_type):
        """Get list of excluded module type IDs for given search parameters"""
        if not Path(DATABASE_FILE).exists():
            return set()
        
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT module_type_id FROM excluded_modules
                WHERE min_price = ? AND max_price = ? 
                AND module_price_type = ? AND mineral_price_type = ?
            """, (min_price, max_price, module_price_type, mineral_price_type))
            results = cursor.fetchall()
            return {row[0] for row in results}
        finally:
            conn.close()
    
    def create_exclude_buttons(self, results):
        """Create exclude buttons for each result"""
        # Clear existing buttons
        for widget in self.exclude_buttons_frame.winfo_children():
            widget.destroy()
        self.exclude_buttons.clear()
        
        if not results:
            return
        
        # Add label
        ttk.Label(self.exclude_buttons_frame, text="Exclude:", font=('', 9, 'bold')).pack(pady=(0, 5))
        
        # Create button for each result
        for idx, result in enumerate(results):
            module_name = result['module_name']
            module_type_id = result['module_type_id']
            
            # Truncate name if too long
            display_name = module_name[:20] + "..." if len(module_name) > 20 else module_name
            
            btn = ttk.Button(
                self.exclude_buttons_frame,
                text=f"Exclude #{idx+1}",
                command=lambda mid=module_type_id, mname=module_name: self.exclude_module(mid, mname),
                width=18
            )
            btn.pack(pady=2)
            self.exclude_buttons[module_type_id] = btn
    
    def exclude_module(self, module_type_id, module_name):
        """Exclude a module from future searches with current parameters"""
        if not self.last_analysis_params:
            messagebox.showwarning("Warning", "No analysis results available")
            return
        
        params = self.last_analysis_params
        min_price = params['min_price']
        max_price = params['max_price']
        module_price_type = params['module_price_type']
        mineral_price_type = params['mineral_price_type']
        
        if not Path(DATABASE_FILE).exists():
            messagebox.showerror("Error", "Database file not found")
            return
        
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO excluded_modules 
                (module_type_id, module_name, min_price, max_price, module_price_type, mineral_price_type)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (module_type_id, module_name, min_price, max_price, module_price_type, mineral_price_type))
            conn.commit()
            messagebox.showinfo("Success", f"'{module_name}' excluded from future searches with these parameters")
            
            # Update button to show it's excluded
            if module_type_id in self.exclude_buttons:
                self.exclude_buttons[module_type_id].config(text="Excluded ✓", state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to exclude module: {str(e)}")
        finally:
            conn.close()
    
    def calculate_single_module(self):
        """Calculate reprocessing value for a single module"""
        module_name = self.module_name_var.get().strip()
        if not module_name:
            messagebox.showwarning("Warning", "Please enter a module name")
            return
        
        self.status_var.set("Calculating...")
        self.single_module_results.delete(1.0, tk.END)
        self.single_module_results.insert(tk.END, f"Calculating reprocessing value for: {module_name}\n\n")
        self.root.update()
        
        def calculate():
            try:
                yield_percent = self.get_float(self.single_yield_var, 55.0)
                markup_percent = self.get_float(self.single_markup_var, 10.0)
                num_modules = self.get_int(self.single_num_modules_var, 100)
                reprocessing_cost = self.get_float(self.single_reprocessing_cost_var, 3.37)
                module_price_type = self.single_module_price_type_var.get()
                mineral_price_type = self.single_mineral_price_type_var.get()
                
                result = calculate_reprocessing_value(
                    module_name=module_name,
                    yield_percent=yield_percent,
                    buy_order_markup_percent=markup_percent,
                    num_modules=num_modules,
                    reprocessing_cost_percent=reprocessing_cost,
                    module_price_type=module_price_type,
                    mineral_price_type=mineral_price_type
                )
                
                formatted = format_reprocessing_result(result)
                
                self.single_module_results.delete(1.0, tk.END)
                self.single_module_results.insert(tk.END, formatted)
                self.status_var.set("Calculation complete!")
                
                # Store result and enable edit button if no error
                if 'error' not in result:
                    self.last_calculation_result = result
                    self.edit_quantities_btn.config(state=tk.NORMAL)
                else:
                    self.last_calculation_result = None
                    self.edit_quantities_btn.config(state=tk.DISABLED)
                    messagebox.showerror("Error", result['error'])
                
            except Exception as e:
                self.single_module_results.delete(1.0, tk.END)
                self.single_module_results.insert(tk.END, f"Error: {str(e)}\n")
                self.status_var.set("Error occurred")
                self.last_calculation_result = None
                self.edit_quantities_btn.config(state=tk.DISABLED)
                messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        
        thread = threading.Thread(target=calculate, daemon=True)
        thread.start()
    
    def edit_quantities(self):
        """Open dialog to edit mineral quantities and recalculate costs"""
        if not self.last_calculation_result or 'error' in self.last_calculation_result:
            messagebox.showwarning("Warning", "Please run a calculation first")
            return
        
        # Get result for use in dialog
        result = self.last_calculation_result
        
        # Create edit dialog
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Edit Quantities")
        edit_window.geometry("900x650")
        edit_window.transient(self.root)
        edit_window.grab_set()
        
        # Frame for instructions and units input
        info_frame = ttk.Frame(edit_window, padding=10)
        info_frame.pack(fill=tk.X)
        
        instruction_label = ttk.Label(info_frame, 
                 text="Edit mineral quantities and number of units required. The system will recalculate costs accordingly.", 
                 wraplength=750, justify=tk.LEFT)
        instruction_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Units required input
        units_frame = ttk.LabelFrame(info_frame, text="Units Required", padding=5)
        units_frame.pack(fill=tk.X, pady=(0, 10))
        
        units_input_frame = ttk.Frame(units_frame)
        units_input_frame.pack(fill=tk.X)
        
        ttk.Label(units_input_frame, text="Units Required to Produce These Quantities:").pack(side=tk.LEFT, padx=5)
        # Use edited units if available, otherwise use num_modules
        units_value = result.get('_edited_units_required', result.get('num_modules', 1))
        units_var = tk.StringVar(value=str(units_value))
        units_entry = ttk.Entry(units_input_frame, textvariable=units_var, width=15)
        units_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(units_input_frame, text="(e.g., 100 for Tremor L)", font=('', 8)).pack(side=tk.LEFT, padx=5)
        
        # Frame for table
        table_frame = ttk.Frame(edit_window, padding=10)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create treeview for editable quantities
        columns = ('Mineral', 'Current Qty', 'Edit Qty', 'Per Module', 'Price', 'Value')
        tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        # Configure columns
        tree.heading('Mineral', text='Mineral')
        tree.heading('Current Qty', text='Current Qty')
        tree.heading('Edit Qty', text='Edit Qty')
        tree.heading('Per Module', text='Per Module')
        tree.heading('Price', text='Price (ISK)')
        tree.heading('Value', text='Value (ISK)')
        
        tree.column('Mineral', width=200)
        tree.column('Current Qty', width=100, anchor=tk.E)
        tree.column('Edit Qty', width=100, anchor=tk.E)
        tree.column('Per Module', width=100, anchor=tk.E)
        tree.column('Price', width=120, anchor=tk.E)
        tree.column('Value', width=120, anchor=tk.E)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Store material data by item ID
        material_data_map = {}
        
        # Populate tree with current data (result already defined above)
        yield_multiplier = result['yield_percent'] / 100.0
        
        for output_mat in result['reprocessing_outputs']:
            material_name = output_mat['materialName']
            # Use actualQuantity which may have been edited previously
            current_qty = int(output_mat.get('actualQuantity', 0))
            per_module = output_mat.get('baseQuantityPerModule', 0)
            price = output_mat['mineralPrice']
            # Recalculate value based on current quantity
            current_value = current_qty * price
            
            item_id = tree.insert('', tk.END, values=(
                material_name,
                f"{current_qty:,}",
                f"{current_qty:,}",
                f"{per_module:.4f}",
                f"{price:,.2f}",
                f"{current_value:,.2f}"
            ))
            
            # Store reference to material data (make a copy to avoid modifying original)
            import copy
            material_data_map[item_id] = copy.deepcopy(output_mat)
        
        # Make quantity column editable
        def on_double_click(event):
            item = tree.selection()[0] if tree.selection() else None
            if not item:
                return
            
            column = tree.identify_column(event.x)
            if column == '#3':  # Edit Qty column
                # Get current value
                current_val = tree.item(item, 'values')[2].replace(',', '')
                
                # Create entry widget
                bbox = tree.bbox(item, column)
                if bbox:
                    x, y, width, height = bbox
                    entry = ttk.Entry(tree, width=15)
                    entry.insert(0, current_val)
                    entry.place(x=x, y=y, width=width, height=height)
                    
                    def save_edit(event=None):
                        try:
                            new_qty = int(entry.get().replace(',', ''))
                            if new_qty < 0:
                                raise ValueError("Quantity must be non-negative")
                            
                            # Update tree
                            values = list(tree.item(item, 'values'))
                            values[2] = f"{new_qty:,}"
                            
                            # Recalculate value
                            material_data = material_data_map[item]
                            price = material_data['mineralPrice']
                            new_value = new_qty * price
                            values[5] = f"{new_value:,.2f}"
                            
                            tree.item(item, values=values)
                            entry.destroy()
                        except ValueError as e:
                            messagebox.showerror("Error", f"Invalid quantity: {e}")
                    
                    def cancel_edit(event=None):
                        entry.destroy()
                    
                    entry.bind('<Return>', save_edit)
                    entry.bind('<FocusOut>', save_edit)
                    entry.bind('<Escape>', cancel_edit)
                    entry.focus_set()
                    entry.select_range(0, tk.END)
        
        tree.bind('<Double-1>', on_double_click)
        
        # Buttons frame
        buttons_frame = ttk.Frame(edit_window, padding=10)
        buttons_frame.pack(fill=tk.X)
        
        def recalculate():
            """Recalculate costs based on edited quantities and units required"""
            try:
                # Get units required from input
                try:
                    units_required = int(units_var.get())
                    if units_required < 1:
                        raise ValueError("Units required must be at least 1")
                except ValueError as e:
                    messagebox.showerror("Error", f"Invalid units required: {e}")
                    return
                
                # Get all edited quantities
                edited_quantities = {}
                
                for item in tree.get_children():
                    values = tree.item(item, 'values')
                    material_data = material_data_map[item]
                    material_name = values[0]
                    edited_qty = int(values[2].replace(',', ''))
                    
                    edited_quantities[material_data['materialTypeID']] = edited_qty
                
                # Use the units_required specified by the user
                actual_modules_needed = units_required
                
                # Recalculate costs and values
                # Use base price (before markup) and apply markup to total
                module_price_before_markup = result.get('module_price_before_markup', result['module_price'])
                buy_order_markup_percent = result.get('buy_order_markup_percent', 0)
                
                # Calculate: base_cost × units + markup on total
                # This matches: cost = 332.8 × 100 + markup
                base_total_cost = module_price_before_markup * actual_modules_needed
                
                # Apply markup to total if using buy_max price type
                if result['module_price_type'] == 'buy_max' and buy_order_markup_percent > 0:
                    total_module_price = base_total_cost * (1 + buy_order_markup_percent / 100.0)
                    module_price = module_price_before_markup * (1 + buy_order_markup_percent / 100.0)
                else:
                    total_module_price = base_total_cost
                    module_price = module_price_before_markup
                
                # Recalculate reprocessing cost
                effective_reprocessing_cost_percent = result['reprocessing_cost_percent'] * (result['yield_percent'] / 100.0)
                reprocessing_cost = total_module_price * (effective_reprocessing_cost_percent / 100.0)
                
                # Recalculate total mineral value from edited quantities
                total_mineral_value = 0.0
                for item in tree.get_children():
                    values = tree.item(item, 'values')
                    edited_qty = int(values[2].replace(',', ''))
                    price = float(values[4].replace(',', ''))
                    value = edited_qty * price
                    total_mineral_value += value
                    
                    # Update value in tree
                    new_values = list(values)
                    new_values[5] = f"{value:,.2f}"
                    tree.item(item, values=new_values)
                
                # Calculate net reprocessing value
                reprocessing_value = total_mineral_value - total_module_price - reprocessing_cost
                
                # Calculate profit margin
                if total_module_price > 0:
                    profit_margin_percent = (reprocessing_value / total_module_price) * 100
                else:
                    profit_margin_percent = float('inf') if reprocessing_value > 0 else 0.0
                
                # Update result
                updated_result = result.copy()
                updated_result['num_modules'] = actual_modules_needed
                updated_result['total_module_price'] = total_module_price
                updated_result['reprocessing_cost'] = reprocessing_cost
                updated_result['total_mineral_value'] = total_mineral_value
                updated_result['reprocessing_value'] = reprocessing_value
                updated_result['profit_margin_percent'] = profit_margin_percent
                
                # Update reprocessing outputs with edited quantities
                for output_mat in updated_result['reprocessing_outputs']:
                    material_type_id = output_mat['materialTypeID']
                    
                    # Update batch_size to match the units_required
                    # This ensures the Batch Size column shows the correct value
                    output_mat['batchSize'] = actual_modules_needed
                    
                    if material_type_id in edited_quantities:
                        edited_qty = edited_quantities[material_type_id]
                        output_mat['actualQuantity'] = edited_qty
                        output_mat['mineralValue'] = edited_qty * output_mat['mineralPrice']
                    else:
                        # Update quantities that weren't edited but need recalculation
                        # Recalculate based on new units_required
                        per_module = output_mat.get('baseQuantityPerModule', 0)
                        new_qty = int(per_module * actual_modules_needed)
                        output_mat['actualQuantity'] = new_qty
                        output_mat['mineralValue'] = new_qty * output_mat['mineralPrice']
                
                # Also update module_price in the result to reflect the recalculated price
                updated_result['module_price'] = module_price
                updated_result['module_price_before_markup'] = module_price_before_markup
                updated_result['buy_order_markup_percent'] = buy_order_markup_percent
                
                # Mark that this result has been edited so it persists
                updated_result['_edited'] = True
                updated_result['_edited_units_required'] = actual_modules_needed
                updated_result['_edited_quantities'] = edited_quantities.copy()
                
                # Update stored result - make a deep copy to ensure it persists
                import copy
                self.last_calculation_result = copy.deepcopy(updated_result)
                
                # Update display
                formatted = format_reprocessing_result(updated_result)
                self.single_module_results.delete(1.0, tk.END)
                self.single_module_results.insert(tk.END, formatted)
                
                # Update status
                self.status_var.set("Recalculation complete!")
                
                # Show summary
                summary = (
                    f"Recalculated with {actual_modules_needed} units required:\n\n"
                    f"Total Module Cost: {total_module_price:,.2f} ISK\n"
                    f"Reprocessing Cost: {reprocessing_cost:,.2f} ISK\n"
                    f"Total Mineral Value: {total_mineral_value:,.2f} ISK\n"
                    f"Net Profit: {reprocessing_value:,.2f} ISK\n"
                    f"Profit Margin: {profit_margin_percent:+.2f}%"
                )
                messagebox.showinfo("Recalculation Complete", summary)
                
                edit_window.destroy()
                
            except Exception as e:
                messagebox.showerror("Error", f"Error recalculating: {str(e)}")
        
        ttk.Button(buttons_frame, text="Recalculate Costs", command=recalculate).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Cancel", command=edit_window.destroy).pack(side=tk.LEFT, padx=5)
    
    def update_all_prices(self):
        """Update all prices in a separate thread"""
        if not messagebox.askyesno("Confirm", "Update all prices? This may take several minutes."):
            return
        
        self.status_var.set("Updating all prices...")
        self.price_update_log.delete(1.0, tk.END)
        self.price_update_log.insert(tk.END, "Starting update of all prices...\n")
        self.price_update_log.insert(tk.END, "This may take several minutes.\n\n")
        self.root.update()
        
        def update():
            try:
                # Redirect logging to our text widget
                import logging
                from io import StringIO
                
                log_capture = StringIO()
                handler = logging.StreamHandler(log_capture)
                handler.setLevel(logging.INFO)
                logger = logging.getLogger()
                logger.addHandler(handler)
                
                update_prices()
                
                logger.removeHandler(handler)
                output = log_capture.getvalue()
                
                self.price_update_log.insert(tk.END, output)
                self.price_update_log.insert(tk.END, "\n\nUpdate complete!\n")
                self.status_var.set("Price update complete!")
                messagebox.showinfo("Success", "All prices updated successfully!")
                
            except Exception as e:
                self.price_update_log.insert(tk.END, f"\nError: {str(e)}\n")
                self.status_var.set("Error occurred")
                messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        
        thread = threading.Thread(target=update, daemon=True)
        thread.start()
    
    def update_mineral_prices_only(self):
        """Update only mineral prices in a separate thread"""
        self.status_var.set("Updating mineral prices...")
        self.price_update_log.delete(1.0, tk.END)
        self.price_update_log.insert(tk.END, "Starting update of mineral prices...\n\n")
        self.root.update()
        
        def update():
            try:
                # Redirect logging to our text widget
                import logging
                from io import StringIO
                
                log_capture = StringIO()
                handler = logging.StreamHandler(log_capture)
                handler.setLevel(logging.INFO)
                logger = logging.getLogger()
                logger.addHandler(handler)
                
                update_mineral_prices()
                
                logger.removeHandler(handler)
                output = log_capture.getvalue()
                
                self.price_update_log.insert(tk.END, output)
                self.price_update_log.insert(tk.END, "\n\nMineral price update complete!\n")
                self.status_var.set("Mineral price update complete!")
                messagebox.showinfo("Success", "Mineral prices updated successfully!")
                
            except Exception as e:
                self.price_update_log.insert(tk.END, f"\nError: {str(e)}\n")
                self.status_var.set("Error occurred")
                messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        
        thread = threading.Thread(target=update, daemon=True)
        thread.start()


def main():
    root = tk.Tk()
    app = EVELauncher(root)
    root.mainloop()


if __name__ == "__main__":
    main()

