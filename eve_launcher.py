"""
EVE Manufacturing Database Launcher
A GUI interface for managing and analyzing EVE Online manufacturing and reprocessing data.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog
import threading
import sys
import math
import json
import sqlite3
import subprocess
from pathlib import Path

# Import our modules
from calculate_reprocessing_value import (
    calculate_reprocessing_value,
    analyze_all_modules,
    format_reprocessing_result,
    sell_into_buy_order,
    sell_order_with_fees,
)
from calculate_blueprint_profitability import calculate_blueprint_profitability, resolve_blueprint, get_blueprint_materials
from decryptor_profitability import compare_decryptor_profitability, DATACORE_NAMES, _estimate_datacore_cost_per_attempt
from invention_lookup import get_t2_products_from_t1
from decryptors_data import get_decryptor_by_name
from skills_blueprints import (
    get_unique_skills,
    get_available_blueprint_ids,
    run_profitability_analysis,
    top_n_by_profit,
    top_n_by_return,
)
from update_prices_db import update_prices, update_prices_by_type_ids
from update_mineral_prices import update_mineral_prices
from fetch_market_history import (
    get_expected_buy_order_volume_7d_avg,
    get_expected_buy_order_volume_30d_avg,
    get_market_history_raw,
    refresh_market_history_for_type,
    get_type_ids_with_no_or_zero_volume,
    run_fetch,
)

DATABASE_FILE = "eve_manufacturing.db"
# Region ID for market_history_daily (The Forge); must match data fetched by fetch_market_history.py
MARKET_HISTORY_REGION_ID = 10000002
# Preferences file for decryptor comparison and other persisted settings
LAUNCHER_PREFS_FILE = Path(__file__).resolve().parent / "eve_launcher_prefs.json"
# Persisted shopping list (survives restarts until reset)
SHOPPING_LIST_FILE = Path(__file__).resolve().parent / "eve_launcher_shopping_list.json"
# Persisted skill levels (My Skills tab)
SKILLS_FILE = Path(__file__).resolve().parent / "eve_launcher_skills.json"

from regions_data import REGIONS_BY_NAME, DEFAULT_REGION_NAME, get_region_id_by_name


class EVELauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("EVE Manufacturing Database Launcher")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 600)
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        # Initialize database tables
        self.init_exclusion_table()
        self.init_on_offer_table()
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Shopping list: filled from eve_launcher_shopping_list.json in create_shopping_list_tab (must exist before that runs)
        self.shopping_list = []
        
        # Create tabs (reordered: price update, On Offer, decryptor comparison, Shopping list first; others follow)
        self.create_price_update_tab()
        self.create_on_offer_tab()
        self.create_decryptor_comparison_tab()
        self.create_shopping_list_tab()
        self.create_analysis_tab()
        self.create_single_module_tab()
        self.create_single_blueprint_tab()
        self.create_skills_blueprints_tab()
        self.create_exclusions_tab()
        self.create_paste_compare_tab()
        self.create_planning_tab()
        self.create_market_patterns_tab()
        self.create_sso_sync_tab()
        
        # So analysis tab fields are editable immediately (focus first entry when that tab is shown)
        self.root.after(150, self._focus_analysis_first_entry_if_visible)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        
        # Store last analysis results for exclusion
        self.last_analysis_results = None
        self.last_analysis_params = None
        # Last single-blueprint calculation result (for shopping list profit when adding from Single Blueprint tab)
        self.last_single_blueprint_result = None
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.root.protocol("WM_DELETE_WINDOW", self._on_launcher_close)
    
    def _on_launcher_close(self):
        """Save shopping list when closing the app (belt-and-suspenders; list also saves on each edit)."""
        try:
            self._save_shopping_list()
        except Exception:
            pass
        self.root.destroy()
    
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
    
    def init_on_offer_table(self):
        """Initialize the on_offer_items table in the database"""
        if not Path(DATABASE_FILE).exists():
            return
        
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS on_offer_items (
                    module_type_id INTEGER PRIMARY KEY,
                    module_name TEXT NOT NULL,
                    added_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (module_type_id) REFERENCES items(typeID)
                )
            """)
            # Add reset/sold tracking columns if missing
            cursor.execute("PRAGMA table_info(on_offer_items)")
            cols = {row[1] for row in cursor.fetchall()}
            added_columns = False
            if 'last_reset_date' not in cols:
                cursor.execute("ALTER TABLE on_offer_items ADD COLUMN last_reset_date TEXT")
                added_columns = True
            if 'quantity_sold_at_last_reset' not in cols:
                cursor.execute("ALTER TABLE on_offer_items ADD COLUMN quantity_sold_at_last_reset INTEGER")
                added_columns = True
            if 'previous_reset_date' not in cols:
                cursor.execute("ALTER TABLE on_offer_items ADD COLUMN previous_reset_date TEXT")
                added_columns = True
            conn.commit()
            # For existing rows: use today as date added when we just added the new columns; else only fill NULL
            if added_columns:
                cursor.execute("UPDATE on_offer_items SET added_at = datetime('now')")
            else:
                cursor.execute("UPDATE on_offer_items SET added_at = datetime('now') WHERE added_at IS NULL")
            conn.commit()
        finally:
            conn.close()
    
    def _focus_analysis_first_entry_if_visible(self):
        """Set focus to the first analysis parameter entry so fields are editable without clicking Run first."""
        try:
            if self.notebook.index(self.notebook.select()) == 0:
                self.analysis_first_entry.focus_set()
        except Exception:
            pass
    
    def _on_tab_changed(self, event):
        """When user switches to Top 30 Analysis tab, focus first entry so fields are editable."""
        try:
            if self.notebook.index(self.notebook.select()) == 0:
                self.analysis_first_entry.focus_set()
        except Exception:
            pass
    
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
        self.analysis_first_entry = ttk.Entry(row1, textvariable=self.yield_var, width=10)
        self.analysis_first_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row1, text="Markup %:").pack(side=tk.LEFT, padx=5)
        self.markup_var = tk.StringVar(value="10.0")
        ttk.Entry(row1, textvariable=self.markup_var, width=10).pack(side=tk.LEFT, padx=5)
        
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
        self.module_price_type_var = tk.StringVar(value="buy_immediate")
        module_price_combo = ttk.Combobox(row3, textvariable=self.module_price_type_var, 
                                         values=["buy_immediate", "buy_offer"], 
                                         state="readonly", width=15)
        module_price_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row3, text="Mineral Price Type:").pack(side=tk.LEFT, padx=5)
        self.mineral_price_type_var = tk.StringVar(value="sell_immediate")
        mineral_price_combo = ttk.Combobox(row3, textvariable=self.mineral_price_type_var,
                                          values=["sell_immediate", "sell_offer"],
                                          state="readonly", width=15)
        mineral_price_combo.pack(side=tk.LEFT, padx=5)
        
        # Row 4 - Item source filter (run on all, blueprint only, or consensus only; faster when restricted)
        row4_filter = ttk.Frame(params_frame)
        row4_filter.pack(fill=tk.X, pady=5)
        ttk.Label(row4_filter, text="Run on:").pack(side=tk.LEFT, padx=5)
        self.item_source_filter_var = tk.StringVar(value="All items")
        item_source_combo = ttk.Combobox(row4_filter, textvariable=self.item_source_filter_var,
                                         values=["All items", "Blueprint items only", "Group consensus items only"],
                                         state="readonly", width=28)
        item_source_combo.pack(side=tk.LEFT, padx=5)
        
        # Row 5 - Source exclusion checkboxes
        row5 = ttk.Frame(params_frame)
        row5.pack(fill=tk.X, pady=5)
        
        ttk.Label(row5, text="Exclude Sources:").pack(side=tk.LEFT, padx=5)
        
        self.exclude_default_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row5, text="Default", variable=self.exclude_default_var).pack(side=tk.LEFT, padx=5)
        
        self.exclude_group_consensus_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row5, text="Group Consensus", variable=self.exclude_group_consensus_var).pack(side=tk.LEFT, padx=5)
        
        self.exclude_group_most_frequent_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row5, text="Group Most Frequent", variable=self.exclude_group_most_frequent_var).pack(side=tk.LEFT, padx=5)
        
        # Row 6 - Sort option
        row6 = ttk.Frame(params_frame)
        row6.pack(fill=tk.X, pady=5)
        ttk.Label(row6, text="Sort by:").pack(side=tk.LEFT, padx=(0, 5))
        self.sort_by_var = tk.StringVar(value="return")
        ttk.Radiobutton(row6, text="% return", variable=self.sort_by_var, value="return").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(row6, text="Profit (ISK)", variable=self.sort_by_var, value="profit").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(row6, text="Expected profit", variable=self.sort_by_var, value="expected_profit").pack(side=tk.LEFT, padx=5)
        
        # Row 7 - Min expected volume filter
        row7 = ttk.Frame(params_frame)
        row7.pack(fill=tk.X, pady=5)
        ttk.Label(row7, text="Min expected volume:").pack(side=tk.LEFT, padx=5)
        self.min_expected_volume_var = tk.StringVar(value="0")
        ttk.Entry(row7, textvariable=self.min_expected_volume_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(row7, text="(0 = no filter; only items with expected vol ≥ this are shown)", font=('', 8)).pack(side=tk.LEFT, padx=5)
        
        # Run button
        run_btn = ttk.Button(params_frame, text="Run Top N Analysis", command=self.run_analysis)
        run_btn.pack(pady=10)
        
        # Results frame with table (like On Offer tab)
        results_frame = ttk.LabelFrame(frame, text="Results", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Hint for user
        hint_label = ttk.Label(results_frame, text="Double-click a row to copy the module name to clipboard.", font=('', 9))
        hint_label.pack(anchor=tk.W, pady=(0, 5))
        
        # Treeview for results table
        columns = ('Rank', 'Module Name', 'Buy Price', 'Sell Min', 'Profit/Item', 'Return %', 'Breakeven Max Buy', 'Expected Vol', 'Expected Profit')
        self.analysis_tree = ttk.Treeview(results_frame, columns=columns, show='headings', height=20, selectmode='browse')
        
        for col in columns:
            self.analysis_tree.heading(col, text=col)
        self.analysis_tree.column('Rank', width=50, anchor=tk.E)
        self.analysis_tree.column('Module Name', width=260, anchor=tk.W)
        self.analysis_tree.column('Buy Price', width=90, anchor=tk.E)
        self.analysis_tree.column('Sell Min', width=90, anchor=tk.E)
        self.analysis_tree.column('Profit/Item', width=100, anchor=tk.E)
        self.analysis_tree.column('Return %', width=80, anchor=tk.E)
        self.analysis_tree.column('Breakeven Max Buy', width=120, anchor=tk.E)
        self.analysis_tree.column('Expected Vol', width=90, anchor=tk.E)
        self.analysis_tree.column('Expected Profit', width=110, anchor=tk.E)
        
        # Tag for rows that are on offer (highlight in blue)
        self.analysis_tree.tag_configure('on_offer', foreground='blue')
        
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.analysis_tree.yview)
        self.analysis_tree.configure(yscrollcommand=scrollbar.set)
        
        self.analysis_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Double-click to copy module name to clipboard
        self.analysis_tree.bind("<Double-1>", self.on_analysis_tree_double_click)
    
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
        
        # Parameters row 2
        params_row2 = ttk.Frame(input_frame)
        params_row2.pack(fill=tk.X, pady=5)
        
        ttk.Label(params_row2, text="Reprocessing Cost %:").pack(side=tk.LEFT, padx=5)
        self.single_reprocessing_cost_var = tk.StringVar(value="3.37")
        ttk.Entry(params_row2, textvariable=self.single_reprocessing_cost_var, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(params_row2, text="Module Price Type:").pack(side=tk.LEFT, padx=5)
        self.single_module_price_type_var = tk.StringVar(value="buy_immediate")
        single_module_price_combo = ttk.Combobox(params_row2, textvariable=self.single_module_price_type_var,
                                                 values=["buy_immediate", "buy_offer"],
                                                 state="readonly", width=15)
        single_module_price_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(params_row2, text="Mineral Price Type:").pack(side=tk.LEFT, padx=5)
        self.single_mineral_price_type_var = tk.StringVar(value="sell_immediate")
        single_mineral_price_combo = ttk.Combobox(params_row2, textvariable=self.single_mineral_price_type_var,
                                                  values=["sell_immediate", "sell_offer"],
                                                  state="readonly", width=15)
        single_mineral_price_combo.pack(side=tk.LEFT, padx=5)
        
        # Buttons frame
        buttons_frame = ttk.Frame(input_frame)
        buttons_frame.pack(pady=10)
        
        calc_btn = ttk.Button(buttons_frame, text="Calculate Reprocessing Value", command=self.calculate_single_module)
        calc_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(buttons_frame, text="Expected volume", command=self.show_single_expected_volume).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Raw market data", command=self.show_single_raw_market_data).pack(side=tk.LEFT, padx=5)
        
        self.edit_quantities_btn = ttk.Button(buttons_frame, text="Edit Quantities", command=self.edit_quantities, state=tk.DISABLED)
        self.edit_quantities_btn.pack(side=tk.LEFT, padx=5)
        
        # Store last calculation result for editing
        self.last_calculation_result = None
        
        # Results frame
        results_frame = ttk.LabelFrame(frame, text="Results", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.single_module_results = scrolledtext.ScrolledText(results_frame, wrap=tk.WORD, height=25)
        self.single_module_results.pack(fill=tk.BOTH, expand=True)
    
    def create_single_blueprint_tab(self):
        """Create the Single Blueprint tab: profitability of manufacturing one blueprint run."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Single Blueprint")
        
        input_frame = ttk.LabelFrame(frame, text="Blueprint / Product", padding=10)
        input_frame.pack(fill=tk.X, padx=10, pady=10)
        
        name_row = ttk.Frame(input_frame)
        name_row.pack(fill=tk.X, pady=5)
        ttk.Label(name_row, text="Blueprint or product name:").pack(side=tk.LEFT, padx=5)
        self.blueprint_name_var = tk.StringVar()
        ttk.Entry(name_row, textvariable=self.blueprint_name_var, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        params_row = ttk.Frame(input_frame)
        params_row.pack(fill=tk.X, pady=5)
        ttk.Label(params_row, text="Input price (materials):").pack(side=tk.LEFT, padx=5)
        self.blueprint_input_price_var = tk.StringVar(value="buy_immediate")
        ttk.Combobox(params_row, textvariable=self.blueprint_input_price_var,
                     values=["buy_immediate", "buy_offer"], state="readonly", width=14).pack(side=tk.LEFT, padx=5)
        ttk.Label(params_row, text="Output price (product):").pack(side=tk.LEFT, padx=5)
        self.blueprint_output_price_var = tk.StringVar(value="sell_immediate")
        ttk.Combobox(params_row, textvariable=self.blueprint_output_price_var,
                     values=["sell_immediate", "sell_offer"], state="readonly", width=14).pack(side=tk.LEFT, padx=5)
        ttk.Label(params_row, text="System cost %:").pack(side=tk.LEFT, padx=5)
        self.blueprint_system_cost_var = tk.StringVar(value="8.61")
        ttk.Entry(params_row, textvariable=self.blueprint_system_cost_var, width=8).pack(side=tk.LEFT, padx=5)
        
        region_row = ttk.Frame(input_frame)
        region_row.pack(fill=tk.X, pady=5)
        ttk.Label(region_row, text="Region (for manufacturing tax):").pack(side=tk.LEFT, padx=5)
        self.blueprint_region_var = tk.StringVar(value=DEFAULT_REGION_NAME)
        region_names = [name for _, name in REGIONS_BY_NAME]
        region_cb = ttk.Combobox(region_row, textvariable=self.blueprint_region_var, values=region_names, state="readonly", width=28)
        region_cb.pack(side=tk.LEFT, padx=5)
        
        me_runs_row = ttk.Frame(input_frame)
        me_runs_row.pack(fill=tk.X, pady=5)
        ttk.Label(me_runs_row, text="Material efficiency:").pack(side=tk.LEFT, padx=5)
        self.blueprint_me_var = tk.StringVar(value="0")
        ttk.Entry(me_runs_row, textvariable=self.blueprint_me_var, width=6).pack(side=tk.LEFT, padx=2)
        ttk.Label(me_runs_row, text="%").pack(side=tk.LEFT, padx=0)
        ttk.Label(me_runs_row, text="Number of runs:").pack(side=tk.LEFT, padx=5)
        self.blueprint_runs_var = tk.StringVar(value="1")
        ttk.Entry(me_runs_row, textvariable=self.blueprint_runs_var, width=8).pack(side=tk.LEFT, padx=5)
        
        btn_row = ttk.Frame(input_frame)
        btn_row.pack(pady=10)
        ttk.Button(btn_row, text="Calculate profitability", command=self.calculate_single_blueprint).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Fetch blueprint data (SDE)", command=self.fetch_blueprint_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Add to shopping list", command=self._add_single_blueprint_to_shopping_list).pack(side=tk.LEFT, padx=5)
        
        results_frame = ttk.LabelFrame(frame, text="Results", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.single_blueprint_results = scrolledtext.ScrolledText(results_frame, wrap=tk.WORD, height=28)
        self.single_blueprint_results_default_bg = self.single_blueprint_results.cget("bg")
        self.single_blueprint_results.tag_configure("profit_positive", foreground="green")
        self.single_blueprint_results.tag_configure("profit_negative", foreground="red")
        self.single_blueprint_results.pack(fill=tk.BOTH, expand=True)
    
    def calculate_single_blueprint(self):
        """Run blueprint profitability calculation and show results."""
        name = self.blueprint_name_var.get().strip()
        if not name:
            messagebox.showwarning("Warning", "Enter a blueprint or product name.")
            return
        self.single_blueprint_results.delete(1.0, tk.END)
        self.single_blueprint_results.insert(tk.END, "Calculating...\n")
        self.single_blueprint_results.configure(bg=self.single_blueprint_results_default_bg)
        self.root.update()
        
        def run():
            try:
                system_pct = self.get_float(self.blueprint_system_cost_var, 8.61)
                if system_pct < 0:
                    system_pct = 0.0
                me_pct = self.get_float(self.blueprint_me_var, 0.0)
                # Pass % directly: calculator uses me_fraction = material_efficiency/100 (10% → 10)
                material_efficiency = max(0.0, min(10.0, me_pct))
                runs = self.get_float(self.blueprint_runs_var, 1.0)
                runs = max(1, int(runs))
                region_id = get_region_id_by_name(self.blueprint_region_var.get())
                result = calculate_blueprint_profitability(
                    blueprint_name_or_product=name,
                    input_price_type=self.blueprint_input_price_var.get(),
                    output_price_type=self.blueprint_output_price_var.get(),
                    system_cost_percent=system_pct,
                    material_efficiency=material_efficiency,
                    number_of_runs=runs,
                    region_id=region_id,
                    db_file=DATABASE_FILE,
                )
                def append(text, tag=None):
                    start = self.single_blueprint_results.index(tk.END)
                    self.single_blueprint_results.insert(tk.END, text)
                    if tag:
                        self.single_blueprint_results.tag_add(tag, start, self.single_blueprint_results.index(tk.END))

                self.single_blueprint_results.delete(1.0, tk.END)
                if "error" in result:
                    self.single_blueprint_results.configure(bg=self.single_blueprint_results_default_bg)
                    append(result["error"] + "\n")
                    self.last_single_blueprint_result = None
                else:
                    self.last_single_blueprint_result = result
                    me_pct = result['material_efficiency']  # calculator stores 0–10 as percent
                    append(f"Blueprint / Product: {result['productName']}\n")
                    append(f"Output: {result['output_total_quantity']:,} × {result['productName']}  ({result['number_of_runs']} run(s), ME {me_pct:.0f}%)\n\n")
                    append("Input materials (total for all runs; per run in parentheses):\n")
                    for m in result["input_materials"]:
                        pr = m['quantity_per_run']
                        pr_fmt = f"{pr:,.2f}" if pr != int(pr) else f"{int(pr):,}"
                        append(f"  {m['materialName']}: {m['quantity']:,} total ({pr_fmt} per run) × {m['unit_price']:,.2f} = {m['total_cost']:,.2f} ISK\n")
                    if result.get("materials_priced_at_zero"):
                        append("Warning: the following materials were priced at 0 (missing or zero price data): " + ", ".join(result["materials_priced_at_zero"]) + "\n\n")
                    append("\n")
                    append("——— For all runs ———\n")
                    append(f"Total input cost:     {result['total_input_cost']:,.2f} ISK\n")
                    eiv = result.get('eiv')
                    eiv_src = result.get('eiv_source', '')
                    eiv_per = result.get('eiv_price_per_unit')
                    if eiv is not None and eiv >= 0:
                        if eiv_src == "adjusted_price" and eiv_per is not None:
                            append(f"EIV (CCP adjusted × output qty): {result['eiv']:,.2f} ISK  (adjusted_price/unit: {eiv_per:,.2f})\n")
                        elif eiv_per and eiv_per > 0:
                            append(f"EIV (market price × output qty): {result['eiv']:,.2f} ISK  (market/unit: {eiv_per:,.2f})\n")
                        else:
                            append(f"EIV: {result['eiv']:,.2f} ISK\n")
                    append(f"System cost ({result['system_cost_percent']}% of EIV): {result['system_cost']:,.2f} ISK\n")
                    append(f"Output revenue:       {result['output_revenue']:,.2f} ISK  ({result['output_total_quantity']:,} × {result['output_unit_price']:,.2f})\n")
                    profit_tag = "profit_positive" if result['profit'] >= 0 else "profit_negative"
                    append(f"Profit:               {result['profit']:,.2f} ISK\n", profit_tag)
                    append(f"Return:               {result['return_percent']:,.2f}%\n\n")
                    append("——— Per item ———\n")
                    append(f"Items produced:       {result['items_produced']:,}\n")
                    append(f"Cost per item:        {result['cost_per_item']:,.2f} ISK\n")
                    append(f"Revenue per item:     {result['revenue_per_item']:,.2f} ISK\n")
                    profit_per_item_tag = "profit_positive" if result['profit_per_item'] >= 0 else "profit_negative"
                    append(f"Profit per item:      {result['profit_per_item']:,.2f} ISK\n", profit_per_item_tag)
                    # Color results area: green if profit >= 0, red if loss
                    if result["profit"] >= 0:
                        self.single_blueprint_results.configure(bg="#dcf8dc")  # light green
                    else:
                        self.single_blueprint_results.configure(bg="#ffd4d4")  # light red
                self.status_var.set("Blueprint calculation complete.")
            except Exception as e:
                self.single_blueprint_results.delete(1.0, tk.END)
                self.single_blueprint_results.insert(tk.END, f"Error: {str(e)}\n")
                self.single_blueprint_results.configure(bg=self.single_blueprint_results_default_bg)
                self.status_var.set("Error occurred")
                self.last_single_blueprint_result = None
        threading.Thread(target=run, daemon=True).start()
    
    def fetch_blueprint_data(self):
        """Run build_database to fetch SDE and repopulate SDE-derived tables (items, blueprints, etc.)."""
        if not messagebox.askyesno(
            "Fetch blueprint data",
            "This will download SDE and rebuild only SDE-derived tables (items, blueprints, materials, skills, invention, reprocessing). "
            "Wallet, ESI sync data, prices, and market history will be kept.\n\nContinue?"
        ):
            return
        self.status_var.set("Fetching blueprint data (rebuilding SDE tables)...")
        self.single_blueprint_results.delete(1.0, tk.END)
        self.single_blueprint_results.insert(tk.END, "Running build_database.py... This may take several minutes.\n\n")
        self.single_blueprint_results.configure(bg=self.single_blueprint_results_default_bg)
        self.root.update()
        
        def run():
            try:
                import logging
                from io import StringIO
                log_capture = StringIO()
                handler = logging.StreamHandler(log_capture)
                handler.setLevel(logging.INFO)
                root_logger = logging.getLogger()
                root_logger.addHandler(handler)
                try:
                    import build_database
                    # Prefer lightweight SDE-only rebuild if available
                    if hasattr(build_database, "rebuild_sde_only"):
                        build_database.rebuild_sde_only()
                    else:
                        build_database.main()
                except Exception as e:
                    self.single_blueprint_results.insert(tk.END, f"\nError: {str(e)}\n")
                finally:
                    root_logger.removeHandler(handler)
                output = log_capture.getvalue()
                self.single_blueprint_results.insert(tk.END, output)
                self.single_blueprint_results.insert(tk.END, "\nDone. You can now run 'Calculate profitability' or update prices.")
                self.status_var.set("Blueprint data fetch complete.")
            except Exception as e:
                self.single_blueprint_results.insert(tk.END, f"\nError: {str(e)}\n")
                self.status_var.set("Error occurred")
        threading.Thread(target=run, daemon=True).start()

    def create_decryptor_comparison_tab(self):
        """Create the Decryptor comparison tab: which decryptor is most profitable for T2 invention."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Decryptor comparison")
        info = ttk.LabelFrame(frame, text="T2 invention: compare decryptors", padding=10)
        info.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(
            info,
            text="Research/invention produces a T2 BPC from a T1 copy. Decryptors (consumed per attempt) change success chance and the resulting BPC's ME and runs. "
                 "Expected cost per successful BPC = (attempt cost including decryptor) ÷ success probability. Profit per BPC = manufacturing profit from that BPC − expected invention cost.",
            justify=tk.LEFT, wraplength=900
        ).pack(anchor=tk.W)
        input_frame = ttk.LabelFrame(frame, text="Parameters", padding=10)
        input_frame.pack(fill=tk.X, padx=10, pady=5)
        row1 = ttk.Frame(input_frame)
        row1.pack(fill=tk.X, pady=3)
        ttk.Label(row1, text="T2 blueprint / product name:").pack(side=tk.LEFT, padx=5)
        self.decryptor_product_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.decryptor_product_var, width=45).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        row1a = ttk.Frame(input_frame)
        row1a.pack(fill=tk.X, pady=3)
        ttk.Label(row1a, text="Or from T1 blueprint/product:").pack(side=tk.LEFT, padx=5)
        self.decryptor_t1_name_var = tk.StringVar()
        ttk.Entry(row1a, textvariable=self.decryptor_t1_name_var, width=40).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(row1a, text="Look up T2 outputs", command=self._decryptor_lookup_t2_from_t1).pack(side=tk.LEFT, padx=5)
        row1b = ttk.Frame(input_frame)
        row1b.pack(fill=tk.X, pady=2)
        ttk.Label(row1b, text="Possible T2 (click to set):").pack(side=tk.LEFT, padx=5)
        self._decryptor_t2_listbox = tk.Listbox(row1b, height=4, width=50, exportselection=False)
        self._decryptor_t2_listbox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        scroll_t2 = ttk.Scrollbar(row1b, orient=tk.VERTICAL, command=self._decryptor_t2_listbox.yview)
        scroll_t2.pack(side=tk.LEFT, fill=tk.Y)
        self._decryptor_t2_listbox.configure(yscrollcommand=scroll_t2.set)
        self._decryptor_t2_listbox.bind("<<ListboxSelect>>", self._on_decryptor_t2_list_select)
        self._decryptor_t2_options = []
        row2 = ttk.Frame(input_frame)
        row2.pack(fill=tk.X, pady=3)
        ttk.Label(row2, text="Base invention chance %:").pack(side=tk.LEFT, padx=5)
        self.decryptor_base_chance_var = tk.StringVar(value="40")
        ttk.Entry(row2, textvariable=self.decryptor_base_chance_var, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(row2, text="Invention cost per attempt (ISK, without decryptor):").pack(side=tk.LEFT, padx=10)
        self.decryptor_inv_cost_var = tk.StringVar(value="0")
        ttk.Entry(row2, textvariable=self.decryptor_inv_cost_var, width=14).pack(side=tk.LEFT, padx=5)
        row3 = ttk.Frame(input_frame)
        row3.pack(fill=tk.X, pady=3)
        ttk.Label(row3, text="Base BPC runs (10 = modules/ammo, 1 = ships/rigs):").pack(side=tk.LEFT, padx=5)
        self.decryptor_base_runs_var = tk.StringVar(value="10")
        ttk.Combobox(row3, textvariable=self.decryptor_base_runs_var, values=["10", "1"], state="readonly", width=6).pack(side=tk.LEFT, padx=5)
        ttk.Label(row3, text="System cost %:").pack(side=tk.LEFT, padx=10)
        self.decryptor_system_cost_var = tk.StringVar(value="8.61")
        ttk.Entry(row3, textvariable=self.decryptor_system_cost_var, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(row3, text="Region:").pack(side=tk.LEFT, padx=10)
        self.decryptor_region_var = tk.StringVar(value=DEFAULT_REGION_NAME)
        region_names = [n for _, n in REGIONS_BY_NAME]
        ttk.Combobox(row3, textvariable=self.decryptor_region_var, values=region_names, state="readonly", width=22).pack(side=tk.LEFT, padx=5)
        row_price = ttk.Frame(input_frame)
        row_price.pack(fill=tk.X, pady=3)
        ttk.Label(row_price, text="Input price (materials):").pack(side=tk.LEFT, padx=5)
        self.decryptor_input_price_var = tk.StringVar(value="buy_immediate")
        ttk.Combobox(row_price, textvariable=self.decryptor_input_price_var,
                     values=["buy_immediate", "buy_offer"], state="readonly", width=14).pack(side=tk.LEFT, padx=5)
        ttk.Label(row_price, text="Output price (product):").pack(side=tk.LEFT, padx=10)
        self.decryptor_output_price_var = tk.StringVar(value="sell_offer")
        ttk.Combobox(row_price, textvariable=self.decryptor_output_price_var,
                     values=["sell_immediate", "sell_offer"], state="readonly", width=14).pack(side=tk.LEFT, padx=5)
        row4 = ttk.Frame(input_frame)
        row4.pack(fill=tk.X, pady=3)
        ttk.Label(row4, text="Datacore 1:").pack(side=tk.LEFT, padx=5)
        self.decryptor_dc1_name_var = tk.StringVar()
        ttk.Combobox(row4, textvariable=self.decryptor_dc1_name_var, values=DATACORE_NAMES, state="readonly", width=40).pack(side=tk.LEFT, padx=5)
        ttk.Label(row4, text="Qty:").pack(side=tk.LEFT, padx=5)
        self.decryptor_dc1_qty_var = tk.StringVar(value="0")
        ttk.Entry(row4, textvariable=self.decryptor_dc1_qty_var, width=6).pack(side=tk.LEFT, padx=5)
        row5 = ttk.Frame(input_frame)
        row5.pack(fill=tk.X, pady=3)
        ttk.Label(row5, text="Datacore 2:").pack(side=tk.LEFT, padx=5)
        self.decryptor_dc2_name_var = tk.StringVar()
        ttk.Combobox(row5, textvariable=self.decryptor_dc2_name_var, values=DATACORE_NAMES, state="readonly", width=40).pack(side=tk.LEFT, padx=5)
        ttk.Label(row5, text="Qty:").pack(side=tk.LEFT, padx=5)
        self.decryptor_dc2_qty_var = tk.StringVar(value="0")
        ttk.Entry(row5, textvariable=self.decryptor_dc2_qty_var, width=6).pack(side=tk.LEFT, padx=5)
        row_bind = ttk.Frame(input_frame)
        row_bind.pack(fill=tk.X, pady=3)
        ttk.Button(row_bind, text="Bind datacores to blueprint", command=self._bind_datacores_to_blueprint).pack(side=tk.LEFT, padx=5)
        ttk.Label(row_bind, text="(saves current datacore 1/2 for this T2 product; they will auto-load next time you use this blueprint)").pack(side=tk.LEFT, padx=5)
        row_assoc = ttk.Frame(input_frame)
        row_assoc.pack(fill=tk.X, pady=3)
        ttk.Label(row_assoc, text="Associate T1 ↔ T2 (save to DB):").pack(side=tk.LEFT, padx=5)
        ttk.Button(row_assoc, text="Associate T1 → T2", command=self._associate_t1_t2).pack(side=tk.LEFT, padx=5)
        ttk.Label(row_assoc, text="Uses T1 from field above and T2 product from field at top. Next time you can enter only T1 and look up T2.").pack(side=tk.LEFT, padx=5)
        btn_row = ttk.Frame(frame)
        btn_row.pack(fill=tk.X, padx=10, pady=5)
        ttk.Button(btn_row, text="Compare decryptors", command=self.run_decryptor_comparison).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Add to shopping list", command=self._add_decryptor_to_shopping_list).pack(side=tk.LEFT, padx=5)
        results_frame = ttk.LabelFrame(frame, text="Results (profit per successful BPC)", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        cols = ("Decryptor", "Success %", "Expected inv. cost", "Decryptor price", "BPC ME", "BPC runs", "Mfg profit", "Profit/BPC")
        self.decryptor_tree = ttk.Treeview(results_frame, columns=cols, show="headings", height=12)
        for c in cols:
            self.decryptor_tree.heading(c, text=c)
            self.decryptor_tree.column(c, width=100, stretch=True)
        scroll = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.decryptor_tree.yview)
        self.decryptor_tree.configure(yscrollcommand=scroll.set)
        self.decryptor_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.decryptor_tree.tag_configure("best", background="#c8e6c9")
        self.decryptor_tree.tag_configure("loss", background="#ffcdd2")
        self.decryptor_tree.bind("<<TreeviewSelect>>", self._on_decryptor_row_selected)
        self._decryptor_comparison_results = []
        details_frame = ttk.LabelFrame(frame, text="Calculation details (click a row)", padding=10)
        details_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.decryptor_details_text = scrolledtext.ScrolledText(details_frame, wrap=tk.WORD, height=10, state=tk.DISABLED)
        self.decryptor_details_text.pack(fill=tk.BOTH, expand=True)
        self._load_decryptor_prefs()

    def create_shopping_list_tab(self):
        """Create the Shopping list tab: blueprints with quantities and aggregated materials list."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Shopping list")
        top = ttk.LabelFrame(frame, text="Blueprints in list", padding=10)
        top.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        cols = ("Own BPC", "Blueprint / Product", "BPC", "Decryptor", "Total runs", "Total material cost", "Sell immediate", "Sell offer", "Expected profit (buy imm, sell off)", "Profit (ISK)")
        self.shopping_list_columns = cols
        self.shopping_list_sort_column = None
        self.shopping_list_sort_reverse = False
        self.shopping_list_tree = ttk.Treeview(top, columns=cols, show="headings", height=10, selectmode="browse")
        for c in cols:
            self.shopping_list_tree.heading(c, text=c, command=lambda col=c: self._shopping_list_sort_by(col))
            w = 52 if c == "Own BPC" else 120
            self.shopping_list_tree.column(c, width=w, stretch=(c != "Own BPC"))
        scroll_tree = ttk.Scrollbar(top, orient=tk.VERTICAL, command=self.shopping_list_tree.yview)
        self.shopping_list_tree.configure(yscrollcommand=scroll_tree.set)
        self.shopping_list_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_tree.pack(side=tk.RIGHT, fill=tk.Y)
        self.shopping_list_tree.bind("<<TreeviewSelect>>", self._on_shopping_list_selection)
        self.shopping_list_tree.bind("<ButtonRelease-1>", self._shopping_list_toggle_own_bpc_click)
        btn_row1 = ttk.Frame(top)
        btn_row1.pack(fill=tk.X, pady=5)
        ttk.Label(btn_row1, text="BPC (blueprint count) for selected:").pack(side=tk.LEFT, padx=5)
        self.shopping_list_qty_var = tk.StringVar(value="1")
        self.shopping_list_qty_entry = ttk.Entry(btn_row1, textvariable=self.shopping_list_qty_var, width=8)
        self.shopping_list_qty_entry.pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row1, text="Update quantity", command=self._shopping_list_update_quantity).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row1, text="Remove selected", command=self._shopping_list_remove_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row1, text="Copy plan to clipboard", command=self._shopping_list_copy_plan_to_clipboard).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row1, text="Refresh profitability", command=self._shopping_list_refresh_profitability).pack(side=tk.LEFT, padx=5)
        ttk.Label(
            top,
            text="Own BPC: click [ ] / [x] if you already have a researched copy (no invention). Removes datacores/decryptors from required items and adds expected datacore ISK back into Profit (ISK). "
                 "Click any column header to sort (click again to reverse). Refresh profitability recalculates profit, decryptor choice, success %, and datacore cost from current database prices (update prices in the Prices tab first for live market data). "
                 "Expected profit / material cost use the same T2 BPC ME as your decryptor row when known; Profit (ISK) is manufacturing minus expected invention cost per BPC.",
            wraplength=720,
            justify=tk.LEFT,
        ).pack(fill=tk.X, anchor=tk.W, pady=(0, 4))
        agg_frame = ttk.LabelFrame(frame, text="Items required for manufacturing (aggregated)", padding=10)
        agg_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.shopping_list_aggregate_text = scrolledtext.ScrolledText(agg_frame, wrap=tk.WORD, height=14, state=tk.DISABLED)
        self.shopping_list_aggregate_text.pack(fill=tk.BOTH, expand=True)
        btn_row2 = ttk.Frame(agg_frame)
        btn_row2.pack(fill=tk.X, pady=5)
        ttk.Button(btn_row2, text="Copy to clipboard", command=self._shopping_list_copy_to_clipboard).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row2, text="Refresh list", command=self._refresh_shopping_list_aggregate).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row2, text="Reset list (clear and save)", command=self._shopping_list_reset).pack(side=tk.LEFT, padx=5)
        # Inventory paste: compare with aggregated list to show shortfall
        inv_frame = ttk.LabelFrame(frame, text="Your inventory (paste item names and quantities; one per line, e.g. 'Tritanium 5000' or 'Tritanium\t5000')", padding=8)
        inv_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.shopping_list_inventory_text = scrolledtext.ScrolledText(inv_frame, wrap=tk.WORD, height=6, state=tk.NORMAL)
        self.shopping_list_inventory_text.pack(fill=tk.BOTH, expand=True)
        inv_btn_row = ttk.Frame(inv_frame)
        inv_btn_row.pack(fill=tk.X, pady=4)
        ttk.Button(inv_btn_row, text="Compare: show shortfall (need − have)", command=self._shopping_list_compare_inventory).pack(side=tk.LEFT, padx=5)
        shortfall_frame = ttk.LabelFrame(frame, text="Still need to get (required minus in inventory)", padding=8)
        shortfall_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.shopping_list_shortfall_text = scrolledtext.ScrolledText(shortfall_frame, wrap=tk.WORD, height=8, state=tk.DISABLED)
        self.shopping_list_shortfall_text.pack(fill=tk.BOTH, expand=True)
        self._load_shopping_list()

    def create_skills_blueprints_tab(self):
        """Tab: select your skill levels, then run analysis to rank all matching blueprints by profit and by return %. T1=10%% ME, T2=0%% ME."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="My Skills")
        self.skills_analysis_status_var = tk.StringVar(value="Set your skill levels and click Run analysis.")
        # Skills panel: scrollable list of (skill name, level 0-5)
        skills_frame = ttk.LabelFrame(frame, text="Your skill levels (0 = none, 1–5 = level)", padding=8)
        skills_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        skills_inner = ttk.Frame(skills_frame)
        skills_inner.pack(fill=tk.BOTH, expand=True)
        self.skills_canvas = tk.Canvas(skills_inner, highlightthickness=0)
        scrollbar_skills = ttk.Scrollbar(skills_inner, orient=tk.VERTICAL, command=self.skills_canvas.yview)
        self.skills_table_frame = ttk.Frame(self.skills_canvas)
        self.skills_table_frame.bind(
            "<Configure>",
            lambda e: self.skills_canvas.configure(scrollregion=self.skills_canvas.bbox("all")),
        )
        self.skills_canvas.create_window((0, 0), window=self.skills_table_frame, anchor=tk.NW)
        self.skills_canvas.configure(yscrollcommand=scrollbar_skills.set)
        self.skills_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_skills.pack(side=tk.RIGHT, fill=tk.Y)
        self.skills_level_vars = {}  # skillID -> IntVar(0..5)
        self._skills_blueprints_fill_skills()
        btn_refresh_skills = ttk.Button(skills_frame, text="Refresh skills from DB", command=self._skills_blueprints_fill_skills)
        btn_refresh_skills.pack(pady=4)
        # Price and type settings
        price_frame = ttk.LabelFrame(frame, text="Price, system cost and blueprint type (used for analysis)", padding=8)
        price_frame.pack(fill=tk.X, padx=10, pady=6)
        row = ttk.Frame(price_frame)
        row.pack(fill=tk.X)
        ttk.Label(row, text="Input price:").pack(side=tk.LEFT, padx=(0, 4))
        self.skills_input_price_var = tk.StringVar(value="buy_immediate")
        ttk.Combobox(row, textvariable=self.skills_input_price_var, values=["buy_immediate", "buy_offer"], state="readonly", width=14).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(row, text="Output price:").pack(side=tk.LEFT, padx=(0, 4))
        self.skills_output_price_var = tk.StringVar(value="sell_immediate")
        ttk.Combobox(row, textvariable=self.skills_output_price_var, values=["sell_immediate", "sell_offer"], state="readonly", width=14).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(row, text="System cost %:").pack(side=tk.LEFT, padx=(0, 4))
        self.skills_system_cost_var = tk.StringVar(value="8.61")
        ttk.Entry(row, textvariable=self.skills_system_cost_var, width=8).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(row, text="Blueprint type:").pack(side=tk.LEFT, padx=(0, 4))
        self.skills_bp_type_var = tk.StringVar(value="Any")
        ttk.Combobox(
            row,
            textvariable=self.skills_bp_type_var,
            values=["Any", "T1 only", "T2 only", "Faction only"],
            state="readonly",
            width=14,
        ).pack(side=tk.LEFT)
        # Run button
        run_row = ttk.Frame(frame)
        run_row.pack(fill=tk.X, padx=10, pady=6)
        ttk.Button(run_row, text="Run analysis (rank all by ISK and by return %)", command=self._run_skills_blueprints_analysis).pack(side=tk.LEFT, padx=5)
        ttk.Label(run_row, textvariable=self.skills_analysis_status_var).pack(side=tk.LEFT, padx=10)
        # Results: two tables
        results_frame = ttk.Frame(frame)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        left = ttk.LabelFrame(results_frame, text="Ranked by profit (ISK)", padding=6)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))
        cols = ("Product", "Profit (ISK)", "Return %", "ME")
        self.skills_top_profit_tree = ttk.Treeview(left, columns=cols, show="headings", height=12)
        for c in cols:
            self.skills_top_profit_tree.heading(c, text=c)
            self.skills_top_profit_tree.column(c, width=120, stretch=True)
        scroll_left = ttk.Scrollbar(left, orient=tk.VERTICAL, command=self.skills_top_profit_tree.yview)
        scroll_left.pack(side=tk.RIGHT, fill=tk.Y)
        self.skills_top_profit_tree.configure(yscrollcommand=scroll_left.set)
        self.skills_top_profit_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.skills_top_profit_tree.bind("<Button-1>", self._on_skills_tree_click)
        right = ttk.LabelFrame(results_frame, text="Ranked by return %", padding=6)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))
        self.skills_top_return_tree = ttk.Treeview(right, columns=cols, show="headings", height=12)
        for c in cols:
            self.skills_top_return_tree.heading(c, text=c)
            self.skills_top_return_tree.column(c, width=120, stretch=True)
        scroll_right = ttk.Scrollbar(right, orient=tk.VERTICAL, command=self.skills_top_return_tree.yview)
        scroll_right.pack(side=tk.RIGHT, fill=tk.Y)
        self.skills_top_return_tree.configure(yscrollcommand=scroll_right.set)
        self.skills_top_return_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.skills_top_return_tree.bind("<Button-1>", self._on_skills_tree_click)

    def _load_skills_prefs(self):
        """Load saved skill levels from JSON. Returns dict skillID (int) -> level (0-5)."""
        if not SKILLS_FILE.exists():
            return {}
        try:
            with open(SKILLS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                return {}
            return {int(k): max(0, min(5, int(v))) for k, v in data.items() if str(k).isdigit() and isinstance(v, (int, float))}
        except Exception:
            return {}

    def _save_skills_prefs(self):
        """Save current skill levels to JSON so they persist across sessions."""
        if not getattr(self, "skills_level_vars", None):
            return
        try:
            data = {}
            for sid, var in self.skills_level_vars.items():
                try:
                    data[str(sid)] = max(0, min(5, int(var.get())))
                except (ValueError, tk.TclError):
                    data[str(sid)] = 0
            with open(SKILLS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass

    def _on_skills_tree_click(self, event):
        """
        When clicking in a skills result tree, copy the Product name to clipboard
        if the click is on the Product column of a data row.
        """
        tree = event.widget
        # Identify row and column under cursor
        row_id = tree.identify_row(event.y)
        col_id = tree.identify_column(event.x)  # "#1" is first column (Product)
        if not row_id or col_id != "#1":
            return
        values = tree.item(row_id, "values")
        if not values or not values[0]:
            return
        product_name = str(values[0])
        self.root.clipboard_clear()
        self.root.clipboard_append(product_name)
        self.status_var.set(f"Copied product name to clipboard: {product_name}")

    def _skills_blueprints_fill_skills(self):
        """Load unique skills from DB and fill the skills table with level spinboxes; restore saved levels from file."""
        for w in self.skills_table_frame.winfo_children():
            w.destroy()
        self.skills_level_vars.clear()
        saved = self._load_skills_prefs()
        if not Path(DATABASE_FILE).exists():
            ttk.Label(self.skills_table_frame, text="Database not found. Run Fetch blueprint data first.").pack(anchor=tk.W)
            return
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            skills = get_unique_skills(conn)
        finally:
            conn.close()
        if not skills:
            ttk.Label(self.skills_table_frame, text="No skills in manufacturing_skills. Run Fetch blueprint data (SDE) in Single Blueprint tab.").pack(anchor=tk.W)
            return
        for s in skills:
            sid, name = s["skillID"], s["skillName"]
            row_f = ttk.Frame(self.skills_table_frame)
            row_f.pack(fill=tk.X, pady=1)
            ttk.Label(row_f, text=name, width=36, anchor=tk.W).pack(side=tk.LEFT, padx=(0, 8))
            default = saved.get(sid, 0)
            var = tk.IntVar(value=default)
            self.skills_level_vars[sid] = var
            sb = ttk.Spinbox(row_f, from_=0, to=5, width=4, textvariable=var)
            sb.pack(side=tk.LEFT)
        self.skills_analysis_status_var.set(f"Loaded {len(skills)} skills (levels restored from file). Set levels and click Run analysis.")

    def _run_skills_blueprints_analysis(self):
        """Gather skill levels, get available blueprints, run profitability, rank and show all by profit and return."""
        user_levels = {}
        for sid, var in self.skills_level_vars.items():
            try:
                user_levels[sid] = max(0, min(5, int(var.get())))
            except (ValueError, tk.TclError):
                user_levels[sid] = 0
        self._save_skills_prefs()
        if not Path(DATABASE_FILE).exists():
            self.skills_analysis_status_var.set("Database not found.")
            return
        self.skills_analysis_status_var.set("Running analysis...")
        for item in self.skills_top_profit_tree.get_children():
            self.skills_top_profit_tree.delete(item)
        for item in self.skills_top_return_tree.get_children():
            self.skills_top_return_tree.delete(item)
        try:
            system_pct = self.get_float(self.skills_system_cost_var, 8.61)
        except (ValueError, tk.TclError):
            system_pct = 8.61
        inp = self.skills_input_price_var.get() or "buy_immediate"
        out = self.skills_output_price_var.get() or "sell_immediate"
        bp_type_filter = self.skills_bp_type_var.get() or "Any"

        def run():
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                bp_ids = get_available_blueprint_ids(conn, user_levels, bp_type_filter=bp_type_filter)
            finally:
                conn.close()
            if not bp_ids:
                self.root.after(0, lambda: self.skills_analysis_status_var.set("No blueprints match your skills."))
                return
            total_bp = len(bp_ids)
            self.root.after(0, lambda: self.skills_analysis_status_var.set(f"Running analysis... 0/{total_bp} blueprints"))

            def on_progress(current, total):
                self.root.after(0, lambda c=current, t=total: self.skills_analysis_status_var.set(f"Running analysis... {c}/{t} blueprints"))

            results = run_profitability_analysis(
                DATABASE_FILE,
                bp_ids,
                input_price_type=inp,
                output_price_type=out,
                system_cost_percent=system_pct,
                progress_callback=on_progress,
            )
            n = len(results)
            top_profit = top_n_by_profit(results, n)
            top_return = top_n_by_return(results, n)

            def show():
                for r in top_profit:
                    self.skills_top_profit_tree.insert("", tk.END, values=(
                        r["productName"][:40],
                        f"{r['profit']:,.0f}",
                        f"{r['return_percent']:.1f}%",
                        f"{r['material_efficiency']}%",
                    ))
                for r in top_return:
                    self.skills_top_return_tree.insert("", tk.END, values=(
                        r["productName"][:40],
                        f"{r['profit']:,.0f}",
                        f"{r['return_percent']:.1f}%",
                        f"{r['material_efficiency']}%",
                    ))
                self.skills_analysis_status_var.set(f"Done: {len(results)} blueprints ranked by profit and by return %.")

            self.root.after(0, show)

        threading.Thread(target=run, daemon=True).start()

    def _add_single_blueprint_to_shopping_list(self):
        """Add current blueprint/product from Single Blueprint tab: 1 BPC, runs = form's Number of runs; profit if last calculation matches."""
        name = self.blueprint_name_var.get().strip()
        if not name:
            messagebox.showwarning("Shopping list", "Enter a blueprint or product name first.")
            return
        try:
            runs = max(1, int(self.blueprint_runs_var.get().strip() or "1"))
        except ValueError:
            runs = 1
        profit = None
        last = getattr(self, "last_single_blueprint_result", None)
        if last and last.get("productName", "").strip() == name:
            profit = last.get("profit")
        self._shopping_list_append(name, 1, profit, runs_per_bpc=runs)

    def _add_decryptor_to_shopping_list(self):
        """Add current T2 product from Decryptor comparison: one row with optional decryptor; invention success prob for datacore/decryptor scaling."""
        name = self.decryptor_product_var.get().strip()
        if not name:
            messagebox.showwarning("Shopping list", "Enter a T2 blueprint or product name first.")
            return
        profit_per_bpc = None
        runs_per_bpc = 1
        entry = {"product_name": name, "quantity": 1, "profit": None, "runs_per_bpc": 1}
        if getattr(self, "_decryptor_comparison_results", None):
            rows = [r for r in self._decryptor_comparison_results if not r.get("error")]
            if rows:
                best = max(rows, key=lambda r: r.get("profit_per_bpc") or -1e99)
                profit_per_bpc = best.get("profit_per_bpc")
                runs_per_bpc = max(1, int(best.get("bpc_runs") or 1))
                entry["profit"] = profit_per_bpc
                entry["runs_per_bpc"] = runs_per_bpc
                sp = best.get("success_prob_pct")
                if sp is not None:
                    try:
                        p = float(sp) / 100.0
                        if 0 < p <= 1.0:
                            entry["invention_success_prob"] = p
                    except (TypeError, ValueError):
                        pass
                dec_name = (best.get("decryptor_name") or "").strip()
                if dec_name and dec_name != "No decryptor":
                    dinfo = get_decryptor_by_name(dec_name)
                    if dinfo:
                        entry["decryptor_name"] = dinfo[0]
                        entry["decryptor_type_id"] = dinfo[1]
                dc_isk = best.get("datacore_cost")
                if dc_isk is not None and sp is not None:
                    try:
                        p = float(sp) / 100.0
                        if p > 0:
                            entry["expected_datacore_cost_per_bpc"] = float(dc_isk) / p
                    except (TypeError, ValueError):
                        pass
                bm = best.get("bpc_me")
                if bm is not None:
                    try:
                        entry["manufacturing_me"] = max(0, min(10, float(bm)))
                    except (TypeError, ValueError):
                        pass
        self._shopping_list_append_planning(entry)

    def _save_shopping_list(self):
        """Persist shopping list to JSON so it survives restarts."""
        try:
            with open(SHOPPING_LIST_FILE, "w", encoding="utf-8") as f:
                json.dump(self.shopping_list, f, indent=2)
        except Exception:
            pass

    def _load_shopping_list(self):
        """Load shopping list from JSON if present and refresh tree/aggregate."""
        if not SHOPPING_LIST_FILE.exists():
            return
        try:
            with open(SHOPPING_LIST_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, list):
                return
            self.shopping_list.clear()
            for entry in data:
                if isinstance(entry, dict) and "product_name" in entry:
                    qty = entry.get("quantity", 1)
                    try:
                        qty = max(1, int(qty))
                    except (TypeError, ValueError):
                        qty = 1
                    profit = entry.get("profit")
                    if profit is not None:
                        try:
                            profit = float(profit)
                        except (TypeError, ValueError):
                            profit = None
                    runs_per_bpc = entry.get("runs_per_bpc")
                    if runs_per_bpc is not None:
                        try:
                            runs_per_bpc = max(1, int(runs_per_bpc))
                        except (TypeError, ValueError):
                            runs_per_bpc = 1
                    else:
                        runs_per_bpc = 1
                    rec = {"product_name": entry["product_name"], "quantity": qty, "profit": profit, "runs_per_bpc": runs_per_bpc}
                    if entry.get("decryptor_name") and entry.get("decryptor_type_id") is not None:
                        rec["decryptor_name"] = entry["decryptor_name"]
                        rec["decryptor_type_id"] = entry["decryptor_type_id"]
                    if entry.get("invention_success_prob") is not None:
                        try:
                            p = float(entry["invention_success_prob"])
                            if 0 < p <= 1.0:
                                rec["invention_success_prob"] = p
                        except (TypeError, ValueError):
                            pass
                    if entry.get("expected_datacore_cost_per_bpc") is not None:
                        try:
                            edc = float(entry["expected_datacore_cost_per_bpc"])
                            if edc >= 0:
                                rec["expected_datacore_cost_per_bpc"] = edc
                        except (TypeError, ValueError):
                            pass
                    if entry.get("bpc_owned_skip_invention"):
                        rec["bpc_owned_skip_invention"] = True
                    if entry.get("manufacturing_me") is not None:
                        try:
                            rec["manufacturing_me"] = max(0, min(10, float(entry["manufacturing_me"])))
                        except (TypeError, ValueError):
                            pass
                    self.shopping_list.append(rec)
            self._shopping_list_refresh_tree()
            self._refresh_shopping_list_aggregate()
        except Exception:
            pass

    def _shopping_list_reset(self):
        """Clear the shopping list, refresh UI, and save empty list (e.g. after procuring items)."""
        self.shopping_list.clear()
        self._shopping_list_refresh_tree()
        self._refresh_shopping_list_aggregate()
        self._save_shopping_list()
        self.status_var.set("Shopping list reset and saved.")

    def _format_shopping_list_profit(self, profit):
        """Format profit for tree display; profit may be None or a number."""
        if profit is None:
            return ""
        try:
            return f"{float(profit):,.0f}"
        except (TypeError, ValueError):
            return ""

    def _shopping_list_append(self, product_name: str, quantity: int, profit=None, runs_per_bpc=1):
        """Append an entry to the shopping list. quantity = BPC count; runs_per_bpc = runs per BPC (manufacturing materials scale by total_runs = quantity * runs_per_bpc; datacores/decryptors scale by quantity)."""
        self.shopping_list.append({"product_name": product_name, "quantity": quantity, "profit": profit, "runs_per_bpc": max(1, int(runs_per_bpc))})
        self._shopping_list_refresh_tree()
        self._refresh_shopping_list_aggregate()
        self._save_shopping_list()
        # Switch to Shopping list tab
        for i in range(self.notebook.index("end")):
            if self.notebook.tab(i, "text") == "Shopping list":
                self.notebook.select(i)
                break
        self.status_var.set(f"Added {product_name} x{quantity} to shopping list.")

    def _shopping_list_append_planning(self, entry: dict):
        """Append a planning entry to the shopping list; entry may include decryptor_name and decryptor_type_id (decryptor shown in column, included in aggregated materials)."""
        base = {"product_name": entry["product_name"], "quantity": entry.get("quantity", 1), "profit": entry.get("profit"), "runs_per_bpc": max(1, int(entry.get("runs_per_bpc") or 1))}
        if entry.get("decryptor_name") and entry.get("decryptor_type_id"):
            base["decryptor_name"] = entry["decryptor_name"]
            base["decryptor_type_id"] = entry["decryptor_type_id"]
        if entry.get("invention_success_prob") is not None:
            try:
                p = float(entry["invention_success_prob"])
                if 0 < p <= 1.0:
                    base["invention_success_prob"] = p
            except (TypeError, ValueError):
                pass
        if entry.get("expected_datacore_cost_per_bpc") is not None:
            try:
                edc = float(entry["expected_datacore_cost_per_bpc"])
                if edc >= 0:
                    base["expected_datacore_cost_per_bpc"] = edc
            except (TypeError, ValueError):
                pass
        if entry.get("bpc_owned_skip_invention"):
            base["bpc_owned_skip_invention"] = True
        if entry.get("manufacturing_me") is not None:
            try:
                base["manufacturing_me"] = max(0, min(10, float(entry["manufacturing_me"])))
            except (TypeError, ValueError):
                pass
        self.shopping_list.append(base)
        self._shopping_list_refresh_tree()
        self._refresh_shopping_list_aggregate()
        self._save_shopping_list()
        for i in range(self.notebook.index("end")):
            if self.notebook.tab(i, "text") == "Shopping list":
                self.notebook.select(i)
                break
        self.status_var.set(f"Added {base['product_name']} to shopping list.")

    def _shopping_list_unit_sell_prices(self, conn, product_name):
        """Return (sell_immediate_unit, sell_offer_unit) for product_name from prices table, or (None, None)."""
        bp = resolve_blueprint(conn, product_name)
        if not bp:
            return (None, None)
        product_type_id = bp["productTypeID"]
        cur = conn.execute("SELECT buy_max, sell_min FROM prices WHERE typeID = ?", (product_type_id,))
        row = cur.fetchone()
        if not row:
            return (None, None)
        buy_max = float(row[0] or 0)
        sell_min = float(row[1] or 0)
        sell_imm = sell_into_buy_order(buy_max) if buy_max and buy_max > 0 else None
        sell_off = sell_order_with_fees(sell_min) if sell_min and sell_min > 0 else None
        return (sell_imm, sell_off)

    def _shopping_list_expected_profit_and_cost(self, entry, total_runs):
        """
        Manufacturing-only profit and material cost: buy inputs at sell orders (buy_immediate),
        sell output via sell_offer, over total_runs. Uses entry['manufacturing_me'] when set
        (T2 BPC ME from decryptor/invention) so this matches the mfg slice of Profit (ISK);
        otherwise ME 0. Profit (ISK) still subtracts invention costs on top of mfg for T2 rows.
        """
        product_name = entry["product_name"]
        me = entry.get("manufacturing_me")
        if me is not None:
            try:
                me = max(0, min(10, float(me)))
            except (TypeError, ValueError):
                me = 0.0
        else:
            me = 0.0
        try:
            result = calculate_blueprint_profitability(
                blueprint_name_or_product=product_name,
                input_price_type="buy_immediate",
                output_price_type="sell_offer",
                system_cost_percent=8.61,
                material_efficiency=me,
                number_of_runs=max(1, int(total_runs)),
                region_id=None,
                db_file=DATABASE_FILE,
            )
            if result:
                return result.get("profit"), result.get("total_input_cost")
        except Exception:
            pass
        return (None, None)

    def _shopping_list_refresh_one_entry_profit(self, conn, entry):
        """
        Recompute stored profitability fields for one shopping list row from current DB prices.
        T2 rows with blueprint_datacore_bindings use decryptor comparison (keeps chosen decryptor if still valid).
        Others use single-blueprint manufacturing profit for runs_per_bpc.
        """
        name = (entry.get("product_name") or "").strip()
        if not name:
            return
        system_cost_pct = 8.61
        region_id = get_region_id_by_name(DEFAULT_REGION_NAME) if DEFAULT_REGION_NAME else MARKET_HISTORY_REGION_ID
        input_price = "buy_immediate"
        output_price = "sell_offer"
        bp = resolve_blueprint(conn, name)
        bind = None
        if bp:
            bind = conn.execute(
                """SELECT dc1_name, dc1_qty, dc2_name, dc2_qty, base_invention_chance_pct, invention_cost_per_attempt, base_bpc_runs
                   FROM blueprint_datacore_bindings WHERE blueprint_type_id = ?""",
                (bp["blueprintTypeID"],),
            ).fetchone()
        if bind and bp:
            dc1, dq1, dc2, dq2 = bind[0], bind[1], bind[2], bind[3]
            base_chance_pct = 40.0
            if len(bind) > 4 and bind[4] is not None:
                base_chance_pct = float(bind[4])
            inv_cost = float(bind[5]) if len(bind) > 5 and bind[5] is not None else 0.0
            base_runs = int(bind[6]) if len(bind) > 6 and bind[6] is not None else 10
            if base_runs not in (1, 10):
                base_runs = 10
            datacores = []
            if dc1 and (dq1 or 0) > 0:
                datacores.append((dc1, int(dq1)))
            if dc2 and (dq2 or 0) > 0:
                datacores.append((dc2, int(dq2)))
            dec_results = compare_decryptor_profitability(
                blueprint_name_or_product=name,
                base_invention_chance_pct=base_chance_pct,
                invention_cost_without_decryptor=inv_cost,
                base_bpc_runs=base_runs,
                input_price_type=input_price,
                output_price_type=output_price,
                system_cost_percent=system_cost_pct,
                region_id=region_id,
                db_file=DATABASE_FILE,
                datacores=datacores if datacores else None,
            )
            valid = [x for x in dec_results if not x.get("error")]
            if valid:
                stored_dec = (entry.get("decryptor_name") or "").strip()
                selected = None
                if stored_dec and stored_dec != "No decryptor":
                    for r in valid:
                        if (r.get("decryptor_name") or "").strip() == stored_dec:
                            selected = r
                            break
                if selected is None:
                    selected = max(valid, key=lambda x: x.get("profit_per_bpc") or -1e99)
                entry["profit"] = selected.get("profit_per_bpc")
                entry["runs_per_bpc"] = max(1, int(selected.get("bpc_runs") or 10))
                sp = selected.get("success_prob_pct")
                if sp is not None:
                    try:
                        p = float(sp) / 100.0
                        if 0 < p <= 1.0:
                            entry["invention_success_prob"] = p
                    except (TypeError, ValueError):
                        pass
                dc_isk = selected.get("datacore_cost")
                if dc_isk is not None and sp is not None:
                    try:
                        p = float(sp) / 100.0
                        if p > 0:
                            entry["expected_datacore_cost_per_bpc"] = float(dc_isk) / p
                    except (TypeError, ValueError):
                        pass
                dn = (selected.get("decryptor_name") or "").strip()
                if dn and dn != "No decryptor":
                    dinfo = get_decryptor_by_name(dn)
                    if dinfo:
                        entry["decryptor_name"] = dinfo[0]
                        entry["decryptor_type_id"] = dinfo[1]
                else:
                    entry.pop("decryptor_name", None)
                    entry.pop("decryptor_type_id", None)
                try:
                    bm = selected.get("bpc_me")
                    if bm is not None:
                        entry["manufacturing_me"] = max(0, min(10, float(bm)))
                except (TypeError, ValueError):
                    pass
                return
        entry.pop("manufacturing_me", None)
        runs = max(1, int(entry.get("runs_per_bpc") or 1))
        result = calculate_blueprint_profitability(
            blueprint_name_or_product=name,
            input_price_type=input_price,
            output_price_type=output_price,
            system_cost_percent=system_cost_pct,
            material_efficiency=0,
            number_of_runs=runs,
            region_id=region_id,
            db_file=DATABASE_FILE,
        )
        if result and "error" not in result:
            entry["profit"] = result.get("profit")
            entry["runs_per_bpc"] = runs

    def _shopping_list_refresh_profitability(self):
        """Background refresh of profit, invention stats, and decryptor line from current DB prices."""
        if not self.shopping_list:
            messagebox.showinfo("Shopping list", "The list is empty.")
            return
        self.status_var.set("Refreshing shopping list profitability...")

        def worker():
            errs = []
            try:
                conn = sqlite3.connect(DATABASE_FILE)
                try:
                    self._ensure_blueprint_datacore_bindings_table(conn)
                    for entry in self.shopping_list:
                        try:
                            self._shopping_list_refresh_one_entry_profit(conn, entry)
                        except Exception as ex:
                            errs.append(f"{entry.get('product_name', '?')}: {ex}")
                finally:
                    conn.close()
            except Exception as ex:
                errs.append(str(ex))

            def done():
                self._shopping_list_refresh_tree()
                self._refresh_shopping_list_aggregate()
                self._save_shopping_list()
                if errs:
                    self.status_var.set(f"Profitability refresh finished with {len(errs)} error(s).")
                    msg = "\n".join(errs[:20])
                    if len(errs) > 20:
                        msg += f"\n... and {len(errs) - 20} more"
                    messagebox.showwarning("Refresh profitability", msg)
                else:
                    self.status_var.set("Shopping list profitability refreshed from DB.")

            self.root.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    def _shopping_list_decryptor_display(self, entry):
        """Return display string for Decryptor column: 'Name x BPC' or '—'."""
        dec_name = (entry.get("decryptor_name") or "").strip()
        if not dec_name or dec_name == "No decryptor":
            return "—"
        bpc = entry.get("quantity", 1)
        return f"{dec_name} x {bpc}"

    def _shopping_list_invention_prob(self, entry):
        """Return success probability 0–1 if this row is an invention plan, else None."""
        p = entry.get("invention_success_prob")
        if p is None:
            return None
        try:
            p = float(p)
            if p <= 0 or p > 1.0:
                return None
            return p
        except (TypeError, ValueError):
            return None

    def _shopping_list_scaled_invention_qty(self, entry, bpc_count, qty_per_attempt):
        """
        Expected consumables for invention attempts to obtain bpc_count successful BPCs:
        ceil((bpc_count * qty_per_attempt) / success_probability).
        qty_per_attempt = datacores consumed per attempt, or 1 for decryptor.
        Without invention_success_prob, returns bpc_count * qty_per_attempt (manufacturing-style).
        """
        bpc_count = max(1, int(bpc_count))
        qty_per_attempt = max(0, int(qty_per_attempt))
        prob = self._shopping_list_invention_prob(entry)
        if prob is None:
            return bpc_count * qty_per_attempt
        return math.ceil((bpc_count * qty_per_attempt) / prob)

    def _shopping_list_own_bpc_display(self, entry):
        """ASCII checkbox for Treeview: already have researched BPC (skip invention in totals)."""
        return "[x]" if entry.get("bpc_owned_skip_invention") else "[ ]"

    def _shopping_list_expected_datacore_cost_per_bpc_resolved(self, conn, entry):
        """
        Expected ISK spent on datacores per successful T2 BPC (datacore_cost_per_attempt / success_probability).
        Uses stored value from Planning/Decryptor when present; else derives from DB + invention_success_prob.
        """
        v = entry.get("expected_datacore_cost_per_bpc")
        if v is not None:
            try:
                f = float(v)
                return f if f >= 0 else None
            except (TypeError, ValueError):
                pass
        prob = self._shopping_list_invention_prob(entry)
        if prob is None or prob <= 0:
            return None
        bp = resolve_blueprint(conn, entry["product_name"])
        if not bp:
            return None
        row = conn.execute(
            "SELECT dc1_name, dc1_qty, dc2_name, dc2_qty FROM blueprint_datacore_bindings WHERE blueprint_type_id = ?",
            (bp["blueprintTypeID"],),
        ).fetchone()
        if not row:
            return None
        dc1_name, dc1_qty, dc2_name, dc2_qty = row
        datacores = []
        if dc1_name and (dc1_qty or 0) > 0:
            datacores.append((dc1_name, int(dc1_qty)))
        if dc2_name and (dc2_qty or 0) > 0:
            datacores.append((dc2_name, int(dc2_qty)))
        if not datacores:
            return None
        cpa = _estimate_datacore_cost_per_attempt(conn, datacores)
        if cpa <= 0:
            return None
        return cpa / prob

    def _shopping_list_profit_cell(self, conn, entry):
        """Profit (ISK) column: manufacturing-style profit from entry; if Own BPC, add back expected datacore ISK per successful BPC."""
        p = entry.get("profit")
        if p is None:
            return ""
        try:
            p = float(p)
        except (TypeError, ValueError):
            return ""
        if entry.get("bpc_owned_skip_invention") and conn is not None:
            add = self._shopping_list_expected_datacore_cost_per_bpc_resolved(conn, entry)
            if add is not None:
                p += add
        return f"{p:,.0f}"

    def _shopping_list_toggle_own_bpc_click(self, event):
        """Toggle bpc_owned_skip_invention when user clicks the Own BPC column ([ ] / [x])."""
        tree = self.shopping_list_tree
        if tree.identify_region(event.x, event.y) not in ("cell", "tree"):
            return
        row_id = tree.identify_row(event.y)
        if not row_id:
            return
        col = tree.identify_column(event.x)
        if col != "#1":
            return
        children = list(tree.get_children())
        try:
            idx = children.index(row_id)
        except ValueError:
            return
        if idx < 0 or idx >= len(self.shopping_list):
            return
        e = self.shopping_list[idx]
        e["bpc_owned_skip_invention"] = not bool(e.get("bpc_owned_skip_invention"))
        self._shopping_list_refresh_tree()
        self._refresh_shopping_list_aggregate()
        self._save_shopping_list()

    def _shopping_list_sort_key(self, conn, entry, column):
        """Return a sortable tuple (type_order, value) for shopping list row and column name."""
        bpc = max(1, int(entry.get("quantity") or 1))
        rpb = max(1, int(entry.get("runs_per_bpc") or 1))
        total_runs = bpc * rpb
        if column == "Own BPC":
            return (0, 0 if entry.get("bpc_owned_skip_invention") else 1)
        if column == "Blueprint / Product":
            return (0, (entry.get("product_name") or "").lower())
        if column == "BPC":
            return (0, float(bpc))
        if column == "Decryptor":
            return (0, self._shopping_list_decryptor_display(entry).lower())
        if column == "Total runs":
            return (0, float(total_runs))
        if column == "Total material cost":
            _, tc = self._shopping_list_expected_profit_and_cost(entry, total_runs)
            return (0, float(tc) if tc is not None else float("-inf"))
        if column == "Sell immediate":
            si, _ = self._shopping_list_unit_sell_prices(conn, entry["product_name"])
            return (0, float(si) if si is not None else float("-inf"))
        if column == "Sell offer":
            _, so = self._shopping_list_unit_sell_prices(conn, entry["product_name"])
            return (0, float(so) if so is not None else float("-inf"))
        if column == "Expected profit (buy imm, sell off)":
            ep, _ = self._shopping_list_expected_profit_and_cost(entry, total_runs)
            return (0, float(ep) if ep is not None else float("-inf"))
        if column == "Profit (ISK)":
            p = entry.get("profit")
            try:
                pf = float(p) if p is not None else float("-inf")
            except (TypeError, ValueError):
                pf = float("-inf")
            if entry.get("bpc_owned_skip_invention"):
                add = self._shopping_list_expected_datacore_cost_per_bpc_resolved(conn, entry)
                if add is not None:
                    pf += add
            return (0, pf)
        return (0, "")

    def _shopping_list_sort_by(self, column):
        """Sort shopping_list in place by column (second click reverses), then rebuild tree."""
        cols = getattr(self, "shopping_list_columns", None)
        if not cols or column not in cols or not self.shopping_list:
            return
        if self.shopping_list_sort_column == column:
            self.shopping_list_sort_reverse = not self.shopping_list_sort_reverse
        else:
            self.shopping_list_sort_reverse = False
            self.shopping_list_sort_column = column
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                decorated = [
                    (self._shopping_list_sort_key(conn, e, column), i, e)
                    for i, e in enumerate(self.shopping_list)
                ]
            finally:
                conn.close()
        except Exception:
            decorated = [((0, ""), i, e) for i, e in enumerate(self.shopping_list)]
        decorated.sort(key=lambda t: t[0], reverse=self.shopping_list_sort_reverse)
        self.shopping_list = [t[2] for t in decorated]
        self._shopping_list_refresh_tree()
        self._save_shopping_list()

    def _shopping_list_refresh_tree(self):
        """Rebuild the Treeview from self.shopping_list. Columns include Decryptor (per blueprint, not a separate line)."""
        for item in self.shopping_list_tree.get_children():
            self.shopping_list_tree.delete(item)
        if not self.shopping_list:
            return
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                for entry in self.shopping_list:
                    bpc = entry["quantity"]
                    runs_per_bpc = max(1, int(entry.get("runs_per_bpc") or 1))
                    total_runs = bpc * runs_per_bpc
                    own_str = self._shopping_list_own_bpc_display(entry)
                    decryptor_str = self._shopping_list_decryptor_display(entry)
                    sell_imm, sell_off = self._shopping_list_unit_sell_prices(conn, entry["product_name"])
                    sell_imm_str = f"{sell_imm:,.2f}" if sell_imm is not None and sell_imm > 0 else "—"
                    sell_off_str = f"{sell_off:,.2f}" if sell_off is not None and sell_off > 0 else "—"
                    exp_profit, total_cost = self._shopping_list_expected_profit_and_cost(entry, total_runs)
                    exp_profit_str = f"{exp_profit:,.0f}" if exp_profit is not None else "—"
                    total_cost_str = f"{total_cost:,.0f}" if total_cost is not None else "—"
                    profit_str = self._shopping_list_profit_cell(conn, entry)
                    self.shopping_list_tree.insert(
                        "", tk.END,
                        values=(own_str, entry["product_name"], bpc, decryptor_str, total_runs, total_cost_str, sell_imm_str, sell_off_str, exp_profit_str, profit_str),
                    )
            finally:
                conn.close()
        except Exception:
            for entry in self.shopping_list:
                bpc = entry["quantity"]
                runs_per_bpc = max(1, int(entry.get("runs_per_bpc") or 1))
                total_runs = bpc * runs_per_bpc
                own_str = self._shopping_list_own_bpc_display(entry)
                decryptor_str = self._shopping_list_decryptor_display(entry)
                exp_profit, total_cost = self._shopping_list_expected_profit_and_cost(entry, total_runs)
                exp_profit_str = f"{exp_profit:,.0f}" if exp_profit is not None else "—"
                total_cost_str = f"{total_cost:,.0f}" if total_cost is not None else "—"
                profit_str = ""
                try:
                    c2 = sqlite3.connect(DATABASE_FILE)
                    try:
                        profit_str = self._shopping_list_profit_cell(c2, entry)
                    finally:
                        c2.close()
                except Exception:
                    profit_str = self._format_shopping_list_profit(entry.get("profit"))
                self.shopping_list_tree.insert(
                    "", tk.END,
                    values=(own_str, entry["product_name"], bpc, decryptor_str, total_runs, total_cost_str, "—", "—", exp_profit_str, profit_str),
                )

    def _on_shopping_list_selection(self, event=None):
        """When user selects a row, fill quantity entry with that row's quantity."""
        sel = self.shopping_list_tree.selection()
        if not sel:
            return
        item = sel[0]
        vals = self.shopping_list_tree.item(item, "values")
        if vals and len(vals) >= 3:
            self.shopping_list_qty_var.set(str(vals[2]))  # BPC

    def _shopping_list_update_quantity(self):
        """Set quantity of the selected row from the quantity entry."""
        sel = self.shopping_list_tree.selection()
        if not sel:
            messagebox.showinfo("Shopping list", "Select a blueprint row first.")
            return
        try:
            qty = max(1, int(self.shopping_list_qty_var.get().strip() or "1"))
        except ValueError:
            messagebox.showwarning("Shopping list", "Enter a valid quantity.")
            return
        item = sel[0]
        children = list(self.shopping_list_tree.get_children())
        try:
            idx = children.index(item)
        except ValueError:
            return
        if idx < 0 or idx >= len(self.shopping_list):
            return
        self.shopping_list[idx]["quantity"] = qty
        product_name = self.shopping_list[idx]["product_name"]
        runs_per_bpc = max(1, int(self.shopping_list[idx].get("runs_per_bpc") or 1))
        total_runs = qty * runs_per_bpc
        ent = self.shopping_list[idx]
        sell_imm_str, sell_off_str = "—", "—"
        exp_profit, total_cost = self._shopping_list_expected_profit_and_cost(ent, total_runs)
        exp_profit_str = f"{exp_profit:,.0f}" if exp_profit is not None else "—"
        total_cost_str = f"{total_cost:,.0f}" if total_cost is not None else "—"
        own_str = self._shopping_list_own_bpc_display(ent)
        profit_str = ""
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                sell_imm, sell_off = self._shopping_list_unit_sell_prices(conn, product_name)
                sell_imm_str = f"{sell_imm:,.2f}" if sell_imm is not None and sell_imm > 0 else "—"
                sell_off_str = f"{sell_off:,.2f}" if sell_off is not None and sell_off > 0 else "—"
                profit_str = self._shopping_list_profit_cell(conn, ent)
            finally:
                conn.close()
        except Exception:
            profit_str = self._format_shopping_list_profit(ent.get("profit"))
        decryptor_str = self._shopping_list_decryptor_display(ent)
        self.shopping_list_tree.item(item, values=(own_str, product_name, qty, decryptor_str, total_runs, total_cost_str, sell_imm_str, sell_off_str, exp_profit_str, profit_str))
        self._refresh_shopping_list_aggregate()
        self._save_shopping_list()
        self.status_var.set(f"Quantity updated to {qty}.")

    def _shopping_list_remove_selected(self):
        """Remove the selected row from the shopping list."""
        sel = self.shopping_list_tree.selection()
        if not sel:
            messagebox.showinfo("Shopping list", "Select a blueprint row to remove.")
            return
        item = sel[0]
        children = list(self.shopping_list_tree.get_children())
        try:
            idx = children.index(item)
        except ValueError:
            return
        if 0 <= idx < len(self.shopping_list):
            self.shopping_list.pop(idx)
        self._shopping_list_refresh_tree()
        self._refresh_shopping_list_aggregate()
        self._save_shopping_list()
        self.status_var.set("Removed from shopping list.")

    def _refresh_shopping_list_aggregate(self):
        """Compute aggregated materials (and datacores) from shopping_list and update the text. Stores result in self.shopping_list_aggregated for inventory comparison."""
        self.shopping_list_aggregate_text.configure(state=tk.NORMAL)
        self.shopping_list_aggregate_text.delete(1.0, tk.END)
        self.shopping_list_aggregated = None
        if not self.shopping_list:
            self.shopping_list_aggregate_text.insert(tk.END, "Add blueprints from Single Blueprint, Decryptor comparison, or Planning, then set quantities. "
                "For invention rows, datacores and decryptors use ceil((BPC × per attempt) ÷ success probability). "
                "Rows marked [x] Own BPC skip invention materials in this list.")
            self.shopping_list_aggregate_text.configure(state=tk.DISABLED)
            return
        aggregated = {}
        if not Path(DATABASE_FILE).exists():
            self.shopping_list_aggregate_text.insert(tk.END, "Database not found. Run build_database / fetch blueprint data first.")
            self.shopping_list_aggregate_text.configure(state=tk.DISABLED)
            return
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                self._ensure_blueprint_datacore_bindings_table(conn)
                for entry in self.shopping_list:
                    name = entry["product_name"]
                    bpc_count = max(1, int(entry["quantity"]))
                    runs_per_bpc = max(1, int(entry.get("runs_per_bpc") or 1))
                    total_runs = bpc_count * runs_per_bpc
                    skip_inv = bool(entry.get("bpc_owned_skip_invention"))
                    dec_name = entry.get("decryptor_name")
                    dec_type_id = entry.get("decryptor_type_id")
                    if not skip_inv and dec_name and dec_type_id and bpc_count > 0:
                        dec_need = self._shopping_list_scaled_invention_qty(entry, bpc_count, 1)
                        aggregated[dec_name] = aggregated.get(dec_name, 0) + dec_need
                    bp = resolve_blueprint(conn, name)
                    if not bp:
                        # Not a blueprint (e.g. decryptor): add as direct item (per BPC/copy)
                        aggregated[name] = aggregated.get(name, 0) + bpc_count
                        continue
                    bid = bp["blueprintTypeID"]
                    materials = get_blueprint_materials(conn, bid)
                    for m in materials:
                        mat_name = m["materialName"]
                        need = m["quantity"] * total_runs
                        aggregated[mat_name] = aggregated.get(mat_name, 0) + need
                    row = conn.execute(
                        "SELECT dc1_name, dc1_qty, dc2_name, dc2_qty FROM blueprint_datacore_bindings WHERE blueprint_type_id = ?",
                        (bid,),
                    ).fetchone()
                    if row and not skip_inv:
                        dc1_name, dc1_qty, dc2_name, dc2_qty = row
                        if dc1_name and dc1_qty:
                            n1 = self._shopping_list_scaled_invention_qty(entry, bpc_count, int(dc1_qty or 0))
                            if n1:
                                aggregated[dc1_name] = aggregated.get(dc1_name, 0) + n1
                        if dc2_name and dc2_qty:
                            n2 = self._shopping_list_scaled_invention_qty(entry, bpc_count, int(dc2_qty or 0))
                            if n2:
                                aggregated[dc2_name] = aggregated.get(dc2_name, 0) + n2
            finally:
                conn.close()
        except Exception as e:
            self.shopping_list_aggregate_text.insert(tk.END, f"Error: {e}")
            self.shopping_list_aggregate_text.configure(state=tk.DISABLED)
            return
        self.shopping_list_aggregated = aggregated
        lines = []
        for name in sorted(aggregated.keys()):
            lines.append(f"{name}\t{aggregated[name]:,}")
        self.shopping_list_aggregate_text.insert(tk.END, "\n".join(lines) if lines else "No materials resolved.")
        self.shopping_list_aggregate_text.configure(state=tk.DISABLED)

    def _parse_inventory_paste(self, text):
        """Parse pasted inventory text into dict item_name -> quantity. Handles 'Name\\tQty', 'Name Qty', 'Qty Name'."""
        import re
        inventory = {}
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            # Tab-separated: "Name\t1234" or "1234\tName"
            if "\t" in line:
                parts = [p.strip() for p in line.split("\t", 1)]
                if len(parts) == 2:
                    a, b = parts
                    try:
                        qty = int(a.replace(",", ""))
                        name = b
                    except ValueError:
                        try:
                            qty = int(b.replace(",", ""))
                            name = a
                        except ValueError:
                            continue
                    if name:
                        inventory[name] = inventory.get(name, 0) + qty
                continue
            # Space-separated: find a number (with optional commas)
            parts = re.split(r"\s+", line)
            if not parts:
                continue
            qty = None
            name_parts = []
            for i, p in enumerate(parts):
                try:
                    qty = int(p.replace(",", ""))
                    name_parts = parts[:i] + parts[i + 1:]
                    break
                except ValueError:
                    pass
            if qty is not None and name_parts:
                name = " ".join(name_parts).strip()
                if name:
                    inventory[name] = inventory.get(name, 0) + qty
        return inventory

    def _normalize_inventory_key(self, name, required_keys):
        """Match pasted item name to an aggregated key (exact or case-insensitive)."""
        name = (name or "").strip()
        if name in required_keys:
            return name
        lower = name.lower()
        for k in required_keys:
            if k.lower() == lower:
                return k
        return name

    def _shopping_list_compare_inventory(self):
        """Parse pasted inventory, compare to aggregated requirements, show shortfall (need − have) in the shortfall text."""
        self.shopping_list_shortfall_text.configure(state=tk.NORMAL)
        self.shopping_list_shortfall_text.delete(1.0, tk.END)
        if not getattr(self, "shopping_list_aggregated", None):
            self._refresh_shopping_list_aggregate()
        aggregated = getattr(self, "shopping_list_aggregated", None) or {}
        if not aggregated:
            self.shopping_list_shortfall_text.insert(tk.END, "No required items (add blueprints and refresh list first).")
            self.shopping_list_shortfall_text.configure(state=tk.DISABLED)
            return
        raw = self.shopping_list_inventory_text.get(1.0, tk.END)
        inventory = self._parse_inventory_paste(raw)
        # Build inventory by normalized key (match to aggregated keys)
        have_by_key = {}
        for pasted_name, qty in inventory.items():
            key = self._normalize_inventory_key(pasted_name, set(aggregated.keys()))
            if key in aggregated:
                have_by_key[key] = have_by_key.get(key, 0) + qty
        shortfall = {}
        for name, need in aggregated.items():
            have = have_by_key.get(name, 0)
            if need > have:
                shortfall[name] = need - have
        if not shortfall:
            self.shopping_list_shortfall_text.insert(tk.END, "You have everything. No shortfall.")
        else:
            lines = []
            for name in sorted(shortfall.keys()):
                lines.append(f"{name}\t{shortfall[name]:,}")
            self.shopping_list_shortfall_text.insert(tk.END, "\n".join(lines))
        self.shopping_list_shortfall_text.configure(state=tk.DISABLED)
        self.status_var.set("Shortfall updated (required − pasted inventory).")

    def _shopping_list_copy_to_clipboard(self):
        """Copy the aggregated materials text to the clipboard."""
        self.shopping_list_aggregate_text.configure(state=tk.NORMAL)
        text = self.shopping_list_aggregate_text.get(1.0, tk.END)
        self.shopping_list_aggregate_text.configure(state=tk.DISABLED)
        text = text.strip()
        if not text or text.startswith("Add blueprints") or text.startswith("Database not found") or text.startswith("Error:"):
            messagebox.showinfo("Copy", "Nothing to copy. Add blueprints and refresh the list first.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status_var.set("Copied aggregated list to clipboard.")

    def _shopping_list_copy_plan_to_clipboard(self):
        """Copy the blueprint/BPC/runs table (the plan) to the clipboard as tab-separated lines for reference."""
        if not self.shopping_list:
            messagebox.showinfo("Copy plan", "Shopping list is empty.")
            return
        lines = ["Own BPC\tBlueprint / Product\tBPC\tDecryptor\tTotal runs\tTotal material cost\tSell immediate\tSell offer\tExpected profit (buy imm, sell off)\tProfit (ISK)"]
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                for entry in self.shopping_list:
                    bpc = entry["quantity"]
                    runs_per_bpc = max(1, int(entry.get("runs_per_bpc") or 1))
                    total_runs = bpc * runs_per_bpc
                    own_str = self._shopping_list_own_bpc_display(entry)
                    decryptor_str = self._shopping_list_decryptor_display(entry)
                    sell_imm, sell_off = self._shopping_list_unit_sell_prices(conn, entry["product_name"])
                    sell_imm_str = f"{sell_imm:,.2f}" if sell_imm is not None and sell_imm > 0 else "—"
                    sell_off_str = f"{sell_off:,.2f}" if sell_off is not None and sell_off > 0 else "—"
                    exp_profit, total_cost = self._shopping_list_expected_profit_and_cost(entry, total_runs)
                    exp_profit_str = f"{exp_profit:,.0f}" if exp_profit is not None else "—"
                    total_cost_str = f"{total_cost:,.0f}" if total_cost is not None else "—"
                    profit_str = self._shopping_list_profit_cell(conn, entry) or ""
                    lines.append(f"{own_str}\t{entry['product_name']}\t{bpc}\t{decryptor_str}\t{total_runs}\t{total_cost_str}\t{sell_imm_str}\t{sell_off_str}\t{exp_profit_str}\t{profit_str}")
            finally:
                conn.close()
        except Exception:
            for entry in self.shopping_list:
                bpc = entry["quantity"]
                runs_per_bpc = max(1, int(entry.get("runs_per_bpc") or 1))
                total_runs = bpc * runs_per_bpc
                own_str = self._shopping_list_own_bpc_display(entry)
                decryptor_str = self._shopping_list_decryptor_display(entry)
                exp_profit, total_cost = self._shopping_list_expected_profit_and_cost(entry, total_runs)
                exp_profit_str = f"{exp_profit:,.0f}" if exp_profit is not None else "—"
                total_cost_str = f"{total_cost:,.0f}" if total_cost is not None else "—"
                profit_str = self._format_shopping_list_profit(entry.get("profit")) or ""
                lines.append(f"{own_str}\t{entry['product_name']}\t{bpc}\t{decryptor_str}\t{total_runs}\t{total_cost_str}\t—\t—\t{exp_profit_str}\t{profit_str}")
        text = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status_var.set("Copied plan (blueprint table) to clipboard.")

    def _decryptor_lookup_t2_from_t1(self):
        """Look up T2 products that can be invented from the given T1 blueprint/product."""
        t1_name = self.decryptor_t1_name_var.get().strip()
        if not t1_name:
            messagebox.showinfo("T1 lookup", "Enter a T1 blueprint or product name, then click Look up T2 outputs.")
            return
        self._decryptor_t2_listbox.delete(0, tk.END)
        self._decryptor_t2_options = []
        try:
            results = get_t2_products_from_t1(t1_name, db_file=DATABASE_FILE)
        except Exception as e:
            messagebox.showerror("T1 lookup", f"Lookup failed: {e}")
            return
        if not results:
            messagebox.showinfo(
                "T1 lookup",
                f"No T2 outputs found for {t1_name!r}. Check the name or run 'Fetch blueprint data (SDE)' in Single Blueprint tab to load invention data."
            )
            return
        for r in results:
            name = r["t2_product_name"]
            prob = r.get("probability")
            qty = r.get("quantity", 1)
            if prob is not None:
                line = f"{name}  (prob {float(prob):.2%}, qty {qty})"
            else:
                line = f"{name}  (qty {qty})"
            self._decryptor_t2_listbox.insert(tk.END, line)
            self._decryptor_t2_options.append(name)
        self.status_var.set(f"Found {len(results)} T2 output(s) for {t1_name}. Click one to set as T2 product.")

    def _on_decryptor_t2_list_select(self, event=None):
        """When user selects a T2 from the list, set the T2 product name field and load saved datacore binding."""
        sel = self._decryptor_t2_listbox.curselection()
        if not sel or not self._decryptor_t2_options:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self._decryptor_t2_options):
            self.decryptor_product_var.set(self._decryptor_t2_options[idx])
            self.status_var.set(f"T2 product set to: {self._decryptor_t2_options[idx]}")
            self._load_datacore_binding_for_product(self._decryptor_t2_options[idx])

    def _load_decryptor_prefs(self):
        """Load last-used decryptor comparison settings from prefs file."""
        if not LAUNCHER_PREFS_FILE.exists():
            return
        try:
            with open(LAUNCHER_PREFS_FILE, "r", encoding="utf-8") as f:
                prefs = json.load(f)
        except Exception:
            return
        dec = prefs.get("decryptor_comparison") or {}
        if dec.get("inv_cost") is not None:
            self.decryptor_inv_cost_var.set(str(dec["inv_cost"]))
        dc1_name = dec.get("dc1_name", "")
        if dc1_name and dc1_name in DATACORE_NAMES:
            self.decryptor_dc1_name_var.set(dc1_name)
        if dec.get("dc1_qty") is not None:
            self.decryptor_dc1_qty_var.set(str(int(dec["dc1_qty"])))
        dc2_name = dec.get("dc2_name", "")
        if dc2_name and dc2_name in DATACORE_NAMES:
            self.decryptor_dc2_name_var.set(dc2_name)
        if dec.get("dc2_qty") is not None:
            self.decryptor_dc2_qty_var.set(str(int(dec["dc2_qty"])))

    def _save_decryptor_prefs(self):
        """Save current decryptor comparison settings (datacores + invention cost) to prefs file."""
        try:
            inv_cost = self.decryptor_inv_cost_var.get().strip()
            prefs = {}
            if LAUNCHER_PREFS_FILE.exists():
                try:
                    with open(LAUNCHER_PREFS_FILE, "r", encoding="utf-8") as f:
                        prefs = json.load(f)
                except Exception:
                    pass
            prefs["decryptor_comparison"] = {
                "inv_cost": inv_cost,
                "dc1_name": self.decryptor_dc1_name_var.get().strip(),
                "dc1_qty": self.decryptor_dc1_qty_var.get().strip(),
                "dc2_name": self.decryptor_dc2_name_var.get().strip(),
                "dc2_qty": self.decryptor_dc2_qty_var.get().strip(),
            }
            with open(LAUNCHER_PREFS_FILE, "w", encoding="utf-8") as f:
                json.dump(prefs, f, indent=2)
        except Exception:
            pass

    def _ensure_blueprint_datacore_bindings_table(self, conn):
        """Create blueprint_datacore_bindings table if it does not exist; add columns for chance/cost/runs if missing."""
        conn.execute("""
            CREATE TABLE IF NOT EXISTS blueprint_datacore_bindings (
                blueprint_type_id INTEGER PRIMARY KEY,
                dc1_name TEXT,
                dc1_qty INTEGER NOT NULL DEFAULT 0,
                dc2_name TEXT,
                dc2_qty INTEGER NOT NULL DEFAULT 0,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (blueprint_type_id) REFERENCES blueprints(blueprintTypeID)
            )
        """)
        cur = conn.execute("PRAGMA table_info(blueprint_datacore_bindings)")
        cols = [row[1] for row in cur.fetchall()]
        for col, typ in [
            ("base_invention_chance_pct", "REAL"),
            ("invention_cost_per_attempt", "REAL"),
            ("base_bpc_runs", "INTEGER"),
        ]:
            if col not in cols:
                conn.execute(f"ALTER TABLE blueprint_datacore_bindings ADD COLUMN {col} {typ}")

    def _ensure_invention_recipes_table(self, conn):
        """Create invention_recipes table if it does not exist (for older DBs)."""
        conn.execute("""
            CREATE TABLE IF NOT EXISTS invention_recipes (
                t1_blueprint_type_id INTEGER NOT NULL,
                t2_blueprint_type_id INTEGER NOT NULL,
                quantity INTEGER NOT NULL DEFAULT 1,
                probability REAL,
                PRIMARY KEY (t1_blueprint_type_id, t2_blueprint_type_id),
                FOREIGN KEY (t1_blueprint_type_id) REFERENCES items(typeID),
                FOREIGN KEY (t2_blueprint_type_id) REFERENCES blueprints(blueprintTypeID)
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_invention_t1 ON invention_recipes(t1_blueprint_type_id)")

    def _load_datacore_binding_for_product(self, product_name):
        """Load saved datacore binding and invention params (chance, cost, base runs) for the given T2 product."""
        if not product_name or not Path(DATABASE_FILE).exists():
            return
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                self._ensure_blueprint_datacore_bindings_table(conn)
                bp = resolve_blueprint(conn, product_name)
                if not bp:
                    return
                blueprint_type_id = bp["blueprintTypeID"]
                row = conn.execute(
                    """SELECT dc1_name, dc1_qty, dc2_name, dc2_qty,
                              base_invention_chance_pct, invention_cost_per_attempt, base_bpc_runs
                       FROM blueprint_datacore_bindings WHERE blueprint_type_id = ?""",
                    (blueprint_type_id,),
                ).fetchone()
                if not row:
                    return
                dc1_name, dc1_qty, dc2_name, dc2_qty = row[0], row[1], row[2], row[3]
                if dc1_name and dc1_name in DATACORE_NAMES:
                    self.decryptor_dc1_name_var.set(dc1_name)
                self.decryptor_dc1_qty_var.set(str(int(dc1_qty or 0)))
                if dc2_name and dc2_name in DATACORE_NAMES:
                    self.decryptor_dc2_name_var.set(dc2_name)
                self.decryptor_dc2_qty_var.set(str(int(dc2_qty or 0)))
                if len(row) > 4 and row[4] is not None:
                    self.decryptor_base_chance_var.set(str(row[4]))
                if len(row) > 5 and row[5] is not None:
                    self.decryptor_inv_cost_var.set(str(int(row[5])))
                if len(row) > 6 and row[6] is not None:
                    self.decryptor_base_runs_var.set(str(int(row[6])))
                self.status_var.set(f"Loaded saved binding for {product_name} (datacores, chance, cost, runs).")
            finally:
                conn.close()
        except Exception:
            pass

    def _bind_datacores_to_blueprint(self):
        """Save current datacores, base chance %, invention cost, and base BPC runs to the current T2 product (bind to blueprint)."""
        product_name = self.decryptor_product_var.get().strip()
        if not product_name:
            messagebox.showwarning("Bind datacores", "Enter a T2 blueprint or product name first.")
            return
        if not Path(DATABASE_FILE).exists():
            messagebox.showerror("Bind datacores", "Database not found. Run build_database / fetch blueprint data first.")
            return
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                self._ensure_blueprint_datacore_bindings_table(conn)
                bp = resolve_blueprint(conn, product_name)
                if not bp:
                    messagebox.showerror("Bind datacores", f"Blueprint/product not found: {product_name!r}")
                    return
                blueprint_type_id = bp["blueprintTypeID"]
                dc1_name = (self.decryptor_dc1_name_var.get() or "").strip()
                try:
                    dc1_qty = int(self.decryptor_dc1_qty_var.get() or "0")
                except ValueError:
                    dc1_qty = 0
                dc2_name = (self.decryptor_dc2_name_var.get() or "").strip()
                try:
                    dc2_qty = int(self.decryptor_dc2_qty_var.get() or "0")
                except ValueError:
                    dc2_qty = 0
                base_chance = self.get_float(self.decryptor_base_chance_var, 40.0)
                inv_cost = self.get_float(self.decryptor_inv_cost_var, 0.0)
                base_runs = self.get_float(self.decryptor_base_runs_var, 10.0)
                base_runs = 1 if base_runs == 1 else 10
                conn.execute("""
                    INSERT OR REPLACE INTO blueprint_datacore_bindings
                    (blueprint_type_id, dc1_name, dc1_qty, dc2_name, dc2_qty,
                     base_invention_chance_pct, invention_cost_per_attempt, base_bpc_runs, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                """, (blueprint_type_id, dc1_name or None, dc1_qty, dc2_name or None, dc2_qty, base_chance, inv_cost, int(base_runs)))
                conn.commit()
                self.status_var.set(f"Binding saved for {product_name} (datacores, chance %, cost, runs).")
            finally:
                conn.close()
        except Exception as e:
            messagebox.showerror("Bind datacores", str(e))

    def _associate_t1_t2(self):
        """Save T1 → T2 association to invention_recipes so T2 can be looked up from T1 later."""
        t1_name = self.decryptor_t1_name_var.get().strip()
        t2_name = self.decryptor_product_var.get().strip()
        if not t1_name or not t2_name:
            messagebox.showwarning("Associate T1→T2", "Enter both T1 blueprint/product (Or from T1 field) and T2 blueprint/product name (top field).")
            return
        if not Path(DATABASE_FILE).exists():
            messagebox.showerror("Associate T1→T2", "Database not found.")
            return
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                # Ensure invention_recipes exists even for databases created before this feature
                self._ensure_invention_recipes_table(conn)
                bp1 = resolve_blueprint(conn, t1_name)
                bp2 = resolve_blueprint(conn, t2_name)
                if not bp1:
                    messagebox.showerror("Associate T1→T2", f"T1 not found: {t1_name!r}")
                    return
                if not bp2:
                    messagebox.showerror("Associate T1→T2", f"T2 not found: {t2_name!r}")
                    return
                t1_bp_id = bp1["blueprintTypeID"]
                t2_bp_id = bp2["blueprintTypeID"]
                conn.execute("""
                    INSERT OR REPLACE INTO invention_recipes (t1_blueprint_type_id, t2_blueprint_type_id, quantity, probability)
                    VALUES (?, ?, 1, ?)
                """, (t1_bp_id, t2_bp_id, 0.4))
                conn.commit()
                self.status_var.set(f"Associated {t1_name} → {t2_name}. You can now enter only T1 and use 'Look up T2 outputs'.")
            finally:
                conn.close()
        except Exception as e:
            messagebox.showerror("Associate T1→T2", str(e))

    def _on_decryptor_row_selected(self, event=None):
        """Show calculation breakdown for the selected decryptor row."""
        self.decryptor_details_text.configure(state=tk.NORMAL)
        self.decryptor_details_text.delete(1.0, tk.END)
        sel = self.decryptor_tree.selection()
        if not sel or not self._decryptor_comparison_results:
            self.decryptor_details_text.insert(tk.END, "Run a comparison, then click a row to see the calculation.")
            self.decryptor_details_text.configure(state=tk.DISABLED)
            return
        item_id = sel[0]
        children = list(self.decryptor_tree.get_children())
        try:
            idx = children.index(item_id)
        except ValueError:
            self.decryptor_details_text.configure(state=tk.DISABLED)
            return
        if idx >= len(self._decryptor_comparison_results):
            self.decryptor_details_text.configure(state=tk.DISABLED)
            return
        r = self._decryptor_comparison_results[idx]
        if r.get("error"):
            self.decryptor_details_text.insert(tk.END, f"Decryptor: {r.get('decryptor_name', '')}\nError: {r['error']}")
            self.decryptor_details_text.configure(state=tk.DISABLED)
            return
        def fmt(x):
            return f"{x:,.2f}" if x is not None and isinstance(x, (int, float)) else str(x)
        inv = r.get("inv_cost_no_dec") or 0
        dc = r.get("datacore_cost") or 0
        dec_price = r.get("decryptor_price") or 0
        attempt = r.get("attempt_cost") or (inv + dc + dec_price)
        prob = r.get("success_prob_pct") or 0
        expected = r.get("expected_inv_cost") or 0
        mfg = r.get("manufacturing_profit") or 0
        profit_bpc = r.get("profit_per_bpc") or 0
        lines = [
            f"Decryptor: {r.get('decryptor_name', '')}",
            "",
            "Invention cost per attempt:",
            f"  Base (no decryptor, no datacores):  {fmt(inv)} ISK",
            f"  Datacore cost:                        {fmt(dc)} ISK",
            f"  Decryptor price:                     {fmt(dec_price)} ISK",
            f"  → Attempt cost (one try):            {fmt(attempt)} ISK",
            "",
            f"Success probability: {fmt(prob)}%",
            f"Expected cost per successful BPC = attempt_cost ÷ (success% / 100) = {fmt(attempt)} ÷ {prob/100:.4f} = {fmt(expected)} ISK",
            "",
            f"Resulting BPC: ME {r.get('bpc_me', '')}%, {r.get('bpc_runs', '')} runs",
            f"Manufacturing profit (all runs): {fmt(mfg)} ISK",
            f"Profit per BPC = manufacturing profit − expected inv. cost = {fmt(mfg)} − {fmt(expected)} = {fmt(profit_bpc)} ISK",
        ]
        self.decryptor_details_text.insert(tk.END, "\n".join(lines))
        self.decryptor_details_text.configure(state=tk.DISABLED)

    def run_decryptor_comparison(self):
        """Run decryptor profitability comparison and fill the tree."""
        name = self.decryptor_product_var.get().strip()
        if not name:
            messagebox.showwarning("Decryptor comparison", "Enter a T2 blueprint or product name.")
            return
        # Load saved datacore binding for this blueprint so we use bound values (and pre-fill form)
        self._load_datacore_binding_for_product(name)
        self._save_decryptor_prefs()
        base_chance = self.get_float(self.decryptor_base_chance_var, 40.0)
        inv_cost = self.get_float(self.decryptor_inv_cost_var, 0.0)
        base_runs = self.get_float(self.decryptor_base_runs_var, 10.0)
        base_runs = 1 if base_runs == 1 else 10
        system_pct = self.get_float(self.decryptor_system_cost_var, 8.61)
        region_id = get_region_id_by_name(self.decryptor_region_var.get())
        input_price_type = self.decryptor_input_price_var.get()
        output_price_type = self.decryptor_output_price_var.get()
        datacores = []
        try:
            q1 = int(self.decryptor_dc1_qty_var.get() or "0")
        except ValueError:
            q1 = 0
        name1 = (self.decryptor_dc1_name_var.get() or "").strip()
        if name1 and q1 > 0:
            datacores.append((name1, q1))
        try:
            q2 = int(self.decryptor_dc2_qty_var.get() or "0")
        except ValueError:
            q2 = 0
        name2 = (self.decryptor_dc2_name_var.get() or "").strip()
        if name2 and q2 > 0:
            datacores.append((name2, q2))
        self.status_var.set("Comparing decryptors...")
        self._decryptor_comparison_results = []
        for item in self.decryptor_tree.get_children():
            self.decryptor_tree.delete(item)
        self.decryptor_details_text.configure(state=tk.NORMAL)
        self.decryptor_details_text.delete(1.0, tk.END)
        self.decryptor_details_text.insert(tk.END, "Running comparison...")
        self.decryptor_details_text.configure(state=tk.DISABLED)

        def run():
            try:
                rows = compare_decryptor_profitability(
                    blueprint_name_or_product=name,
                    base_invention_chance_pct=base_chance,
                    invention_cost_without_decryptor=inv_cost,
                    base_bpc_runs=base_runs,
                    input_price_type=input_price_type,
                    output_price_type=output_price_type,
                    system_cost_percent=system_pct,
                    region_id=region_id,
                    db_file=DATABASE_FILE,
                    datacores=datacores,
                )
                self._decryptor_comparison_results = rows
                def fmt_isk(x):
                    return f"{x:,.0f}" if x is not None and isinstance(x, (int, float)) else (str(x) if x is not None else "")
                best_profit = None
                for r in rows:
                    if r.get("error"):
                        self.decryptor_tree.insert("", tk.END, values=(r.get("decryptor_name", ""), r["error"], "", "", "", "", "", ""), tags=("loss",))
                        continue
                    profit = r.get("profit_per_bpc")
                    if profit is not None and (best_profit is None or profit > best_profit):
                        best_profit = profit
                for r in rows:
                    if r.get("error"):
                        continue
                    vals = (
                        r["decryptor_name"],
                        f"{r['success_prob_pct']:.1f}",
                        fmt_isk(r["expected_inv_cost"]),
                        fmt_isk(r["decryptor_price"]),
                        str(r["bpc_me"]),
                        str(r["bpc_runs"]),
                        fmt_isk(r["manufacturing_profit"]),
                        fmt_isk(r["profit_per_bpc"]),
                    )
                    tag = None
                    if best_profit is not None and r.get("profit_per_bpc") == best_profit and best_profit > 0:
                        tag = "best"
                    elif (r.get("profit_per_bpc") or 0) < 0:
                        tag = "loss"
                    self.decryptor_tree.insert("", tk.END, values=vals, tags=(tag,) if tag else ())
                self.status_var.set("Decryptor comparison complete.")
                self.decryptor_details_text.configure(state=tk.NORMAL)
                self.decryptor_details_text.delete(1.0, tk.END)
                self.decryptor_details_text.insert(tk.END, "Click a row above to see the calculation breakdown.")
                self.decryptor_details_text.configure(state=tk.DISABLED)
            except Exception as e:
                self.decryptor_tree.insert("", tk.END, values=("Error", str(e), "", "", "", "", "", ""), tags=("loss",))
                self.status_var.set("Error occurred")
        threading.Thread(target=run, daemon=True).start()

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
   - Mutaplasmid residues and other specified materials
   - All invention datacores (same set as Decryptor comparison)
   - Decryptors (for invention profitability)
   
   After the run, a before/after table is shown for Mexallon, Pyerite, Tritanium, Zydrine, Megacyte, Nocxium, Isogen.
   
   This is much faster and recommended for regular updates.

3. Update Blueprint Items Only: Updates prices only for items that have
   an identified blueprint (source='blueprint' in input_quantity_cache).

4. Update Group Consensus Items Only: Updates prices only for items that
   use group consensus for input quantity (source='group_consensus' in input_quantity_cache).
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
        
        # Second row of buttons
        buttons_frame2 = ttk.Frame(frame)
        buttons_frame2.pack(fill=tk.X, padx=10, pady=10)
        
        update_blueprint_btn = ttk.Button(buttons_frame2, text="Update Blueprint Items Only",
                                         command=self.update_blueprint_prices, width=30)
        update_blueprint_btn.pack(side=tk.LEFT, padx=10, expand=True)
        
        update_consensus_btn = ttk.Button(buttons_frame2, text="Update Group Consensus Items Only",
                                         command=self.update_group_consensus_prices, width=30)
        update_consensus_btn.pack(side=tk.LEFT, padx=10, expand=True)
        
        # Third row - Market history (volume) fetch
        buttons_frame3 = ttk.Frame(frame)
        buttons_frame3.pack(fill=tk.X, padx=10, pady=10)
        ttk.Button(buttons_frame3, text="Fetch market history (same set as Update All Prices)",
                   command=self.run_fetch_market_history_prices, width=42).pack(side=tk.LEFT, padx=10, expand=True)
        ttk.Button(buttons_frame3, text="Refresh volume for items with no/zero data",
                  command=self.refresh_volume_no_or_zero_data, width=35).pack(side=tk.LEFT, padx=10, expand=True)
        
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
    
    def create_on_offer_tab(self):
        """Create the On Offer tab to track items with active orders"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="On Offer")
        
        # Add item frame
        add_frame = ttk.LabelFrame(frame, text="Add Item to Track", padding=10)
        add_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Input field
        input_row = ttk.Frame(add_frame)
        input_row.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_row, text="Item Name or TypeID:").pack(side=tk.LEFT, padx=5)
        self.on_offer_item_var = tk.StringVar()
        ttk.Entry(input_row, textvariable=self.on_offer_item_var, width=40).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Add button
        add_btn = ttk.Button(add_frame, text="Add Item", command=self.add_on_offer_item)
        add_btn.pack(pady=10)
        
        info_label = ttk.Label(add_frame, text="Note: Buy price and sell min are fetched from current market data", 
                              font=('', 8), foreground='gray')
        info_label.pack(pady=5)
        
        # Table frame
        table_frame = ttk.LabelFrame(frame, text="Items On Offer", padding=10)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create treeview with all required columns
        columns = ('Name', 'Date Added', 'Buy Price', 'Sell Min', 'Profit/Item (Buy Order)', 'Profit/Item (Immediate)', 
                   'Breakeven Max (Buy Order)', 'Breakeven Max (Immediate)', 'Sold Per Day')
        self.on_offer_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)
        
        # Configure columns
        self.on_offer_tree.heading('Name', text='Name')
        self.on_offer_tree.heading('Date Added', text='Date Added')
        self.on_offer_tree.heading('Buy Price', text='Buy Price (buy_max)')
        self.on_offer_tree.heading('Sell Min', text='Sell Min')
        self.on_offer_tree.heading('Profit/Item (Buy Order)', text='Profit/Item (Buy Order)')
        self.on_offer_tree.heading('Profit/Item (Immediate)', text='Profit/Item (Immediate)')
        self.on_offer_tree.heading('Breakeven Max (Buy Order)', text='Breakeven Max (Buy Order)')
        self.on_offer_tree.heading('Breakeven Max (Immediate)', text='Breakeven Max (Immediate)')
        self.on_offer_tree.heading('Sold Per Day', text='Sold Per Day')
        
        self.on_offer_tree.column('Name', width=220)
        self.on_offer_tree.column('Date Added', width=100, anchor=tk.CENTER)
        self.on_offer_tree.column('Buy Price', width=100, anchor=tk.E)
        self.on_offer_tree.column('Sell Min', width=100, anchor=tk.E)
        self.on_offer_tree.column('Profit/Item (Buy Order)', width=150, anchor=tk.E)
        self.on_offer_tree.column('Profit/Item (Immediate)', width=150, anchor=tk.E)
        self.on_offer_tree.column('Breakeven Max (Buy Order)', width=170, anchor=tk.E)
        self.on_offer_tree.column('Breakeven Max (Immediate)', width=170, anchor=tk.E)
        self.on_offer_tree.column('Sold Per Day', width=90, anchor=tk.E)
        
        # Light red when buy_max > 90% of breakeven max (buy order)
        self.on_offer_tree.tag_configure('high_buy_near_breakeven', background='#ffcccc')
        # Deep red when buy_max > breakeven max (buy order)
        self.on_offer_tree.tag_configure('sell_above_breakeven', background='#cc6666')
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.on_offer_tree.yview)
        self.on_offer_tree.configure(yscrollcommand=scrollbar.set)
        
        self.on_offer_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Action buttons frame
        action_frame = ttk.Frame(frame, padding=10)
        action_frame.pack(fill=tk.X, padx=10, pady=5)
        
        refresh_btn = ttk.Button(action_frame, text="Refresh Calculations", command=self.refresh_on_offer_list)
        refresh_btn.pack(side=tk.LEFT, padx=5)
        
        reset_date_btn = ttk.Button(action_frame, text="Reset date (enter quantity sold)", command=self.reset_on_offer_date)
        reset_date_btn.pack(side=tk.LEFT, padx=5)
        
        remove_btn = ttk.Button(action_frame, text="Remove Selected", command=self.remove_on_offer_item)
        remove_btn.pack(side=tk.LEFT, padx=5)

        # Launch overview alert helper
        launch_overview_btn = ttk.Button(action_frame, text="Open Overview Alert", command=self.launch_overview_alert)
        launch_overview_btn.pack(side=tk.LEFT, padx=20)
        
        # Load items on startup
        self.refresh_on_offer_list()
    
    def create_paste_compare_tab(self):
        """Create the Paste & Compare tab: paste in-game window (Name<Tab>Qty), compare reprocess vs sell."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Paste & Compare")
        
        # Instructions
        info_frame = ttk.LabelFrame(frame, text="Instructions", padding=10)
        info_frame.pack(fill=tk.X, padx=10, pady=10)
        info_text = (
            "Paste in-game window content: one line per item, 'Name<Tab>Quantity' (quantity optional, default 1). "
            "For reprocessable items: if item value ≥ threshold we compare to lowest sell; else to lowest buy. "
            "Recommend Sell only when (sell value − reprocess value) × Qty ≥ 'Min ISK above reprocess to recommend Sell'. "
            "For manufacturing: paste one blueprint or product name per line; system will compute profit for 1/10/100 runs at 0% and 10% ME."
        )
        ttk.Label(info_frame, text=info_text, justify=tk.LEFT, wraplength=900).pack(anchor=tk.W)
        
        # Mode: Reprocessing vs Manufacturing
        mode_frame = ttk.Frame(frame)
        mode_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(mode_frame, text="Mode:").pack(side=tk.LEFT, padx=5)
        self.paste_compare_mode_var = tk.StringVar(value="reprocessing")
        ttk.Radiobutton(mode_frame, text="Reprocessing", variable=self.paste_compare_mode_var, value="reprocessing", command=self._paste_compare_switch_mode).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="Manufacturing (blueprints)", variable=self.paste_compare_mode_var, value="manufacturing", command=self._paste_compare_switch_mode).pack(side=tk.LEFT, padx=5)
        self.paste_compare_mfg_params_frame = ttk.Frame(mode_frame)
        self.paste_compare_mfg_params_frame.pack(side=tk.LEFT, padx=15)
        ttk.Label(self.paste_compare_mfg_params_frame, text="System cost %:").pack(side=tk.LEFT, padx=5)
        self.paste_compare_system_cost_var = tk.StringVar(value="8.61")
        ttk.Entry(self.paste_compare_mfg_params_frame, textvariable=self.paste_compare_system_cost_var, width=8).pack(side=tk.LEFT, padx=2)
        
        # Paste area
        paste_frame = ttk.LabelFrame(frame, text="Paste content (Name<Tab>Quantity)", padding=10)
        paste_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.paste_compare_text = scrolledtext.ScrolledText(paste_frame, wrap=tk.WORD, height=8, width=80)
        self.paste_compare_text.pack(fill=tk.BOTH, expand=True)
        clear_paste_btn = ttk.Button(paste_frame, text="Clear paste content", command=self.clear_paste_compare_text)
        clear_paste_btn.pack(pady=(5, 0))
        
        # Parameters (repro-only and shared)
        params_frame = ttk.Frame(frame)
        params_frame.pack(fill=tk.X, padx=10, pady=5)
        self.paste_compare_repro_params_frame = ttk.Frame(params_frame)
        self.paste_compare_repro_params_frame.pack(side=tk.LEFT)
        ttk.Label(self.paste_compare_repro_params_frame, text="Threshold (ISK):").pack(side=tk.LEFT, padx=5)
        self.paste_threshold_var = tk.StringVar(value="100000")
        ttk.Entry(self.paste_compare_repro_params_frame, textvariable=self.paste_threshold_var, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Label(self.paste_compare_repro_params_frame, text="Min ISK above reprocess to recommend Sell:").pack(side=tk.LEFT, padx=5)
        self.paste_sell_buffer_var = tk.StringVar(value="0")
        ttk.Entry(self.paste_compare_repro_params_frame, textvariable=self.paste_sell_buffer_var, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Label(self.paste_compare_repro_params_frame, text="Yield %:").pack(side=tk.LEFT, padx=5)
        self.paste_yield_var = tk.StringVar(value="55.0")
        ttk.Entry(self.paste_compare_repro_params_frame, textvariable=self.paste_yield_var, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(self.paste_compare_repro_params_frame, text="Reprocessing cost %:").pack(side=tk.LEFT, padx=5)
        self.paste_repro_cost_var = tk.StringVar(value="3.37")
        ttk.Entry(self.paste_compare_repro_params_frame, textvariable=self.paste_repro_cost_var, width=8).pack(side=tk.LEFT, padx=5)
        
        compare_btn = ttk.Button(params_frame, text="Compare", command=self.run_paste_compare)
        compare_btn.pack(side=tk.LEFT, padx=15)
        
        # Results table (two trees: reprocessing and manufacturing)
        results_frame = ttk.LabelFrame(frame, text="Results", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        # Reprocessing tree
        self.paste_compare_columns = ('Item Name', 'Qty', 'Sell Min', 'Buy Max', 'Reprocess Value/Item', 'Recommendation')
        self.paste_compare_tree = ttk.Treeview(results_frame, columns=self.paste_compare_columns, show='headings', height=20, selectmode='browse')
        self.paste_compare_sort_column = None
        self.paste_compare_sort_reverse = False
        for col in self.paste_compare_columns:
            self.paste_compare_tree.heading(col, text=col, command=lambda c=col: self.sort_paste_compare_by(c))
        self.paste_compare_tree.column('Item Name', width=320, anchor=tk.W)
        self.paste_compare_tree.column('Qty', width=50, anchor=tk.E)
        self.paste_compare_tree.column('Sell Min', width=100, anchor=tk.E)
        self.paste_compare_tree.column('Buy Max', width=100, anchor=tk.E)
        self.paste_compare_tree.column('Reprocess Value/Item', width=140, anchor=tk.E)
        self.paste_compare_tree.column('Recommendation', width=120, anchor=tk.W)
        self.paste_compare_scrollbar_repro = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.paste_compare_tree.yview)
        self.paste_compare_tree.configure(yscrollcommand=self.paste_compare_scrollbar_repro.set)
        self.paste_compare_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.paste_compare_scrollbar_repro.pack(side=tk.RIGHT, fill=tk.Y)
        # Manufacturing tree (same frame, shown when mode=manufacturing)
        self.paste_compare_columns_mfg = ('Blueprint', 'Profit 1r ME0', 'Profit 1r ME10', 'Profit 10r ME0', 'Profit 10r ME10', 'Profit 100r ME0', 'Profit 100r ME10')
        self.paste_compare_tree_mfg = ttk.Treeview(results_frame, columns=self.paste_compare_columns_mfg, show='headings', height=20, selectmode='browse')
        for col in self.paste_compare_columns_mfg:
            self.paste_compare_tree_mfg.heading(col, text=col)
        self.paste_compare_tree_mfg.column('Blueprint', width=280, anchor=tk.W)
        for c in self.paste_compare_columns_mfg[1:]:
            self.paste_compare_tree_mfg.column(c, width=100, anchor=tk.E)
        self.paste_compare_scrollbar_mfg = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.paste_compare_tree_mfg.yview)
        self.paste_compare_tree_mfg.configure(yscrollcommand=self.paste_compare_scrollbar_mfg.set)
        # Initially hide mfg tree (repro is visible)
        self.paste_compare_tree_mfg.pack_forget()
        self.paste_compare_scrollbar_mfg.pack_forget()
        # Sync visibility with mode (hide mfg params initially since default is reprocessing)
        self._paste_compare_switch_mode()
    
    def _paste_compare_switch_mode(self):
        """Show/hide params and result tree based on Reprocessing vs Manufacturing mode."""
        if self.paste_compare_mode_var.get() == "manufacturing":
            self.paste_compare_mfg_params_frame.pack(side=tk.LEFT, padx=15)
            self.paste_compare_repro_params_frame.pack_forget()
            self.paste_compare_tree.pack_forget()
            self.paste_compare_scrollbar_repro.pack_forget()
            self.paste_compare_tree_mfg.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.paste_compare_scrollbar_mfg.pack(side=tk.RIGHT, fill=tk.Y)
        else:
            self.paste_compare_mfg_params_frame.pack_forget()
            self.paste_compare_repro_params_frame.pack(side=tk.LEFT)
            self.paste_compare_tree_mfg.pack_forget()
            self.paste_compare_scrollbar_mfg.pack_forget()
            self.paste_compare_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.paste_compare_scrollbar_repro.pack(side=tk.RIGHT, fill=tk.Y)
    
    def _paste_compare_sort_key(self, values, col_index):
        """Return a sort key for a row (tuple of values) for the given column index."""
        if col_index >= len(values):
            return (0, "")
        val = values[col_index]
        s = str(val).strip()
        if col_index == 0:  # Item Name - alphabetical, case-insensitive
            return (0, (s or "").lower())
        if col_index == 5:  # Recommendation - group by type, then alphabetically by name
            name = (values[0] or "").lower() if len(values) > 0 else ""
            return (0, s or "", name)
        if col_index == 1:  # Qty - numeric
            try:
                return (0, int(s))
            except ValueError:
                return (1, s)
        if col_index in (2, 3, 4):  # Sell Min, Buy Max, Reprocess Value/Item - numeric
            try:
                return (0, float(s.replace(",", "")))
            except ValueError:
                return (1, s)
        return (0, s)
    
    def sort_paste_compare_by(self, column):
        """Sort Paste & Compare table by the clicked column. Toggle asc/desc on same column."""
        tree = self.paste_compare_tree
        children = list(tree.get_children(""))
        if not children:
            return
        # Don't sort when the only row is a placeholder ("Comparing...", "Error:...")
        if len(children) == 1:
            first_vals = tree.item(children[0])["values"]
            if len(first_vals) >= 2:
                second = str(first_vals[1] or "")
                if second == "Comparing..." or second.startswith("Error:"):
                    return
        if self.paste_compare_sort_column == column:
            self.paste_compare_sort_reverse = not self.paste_compare_sort_reverse
        else:
            self.paste_compare_sort_reverse = False
            self.paste_compare_sort_column = column
        col_index = self.paste_compare_columns.index(column) if column in self.paste_compare_columns else 0
        # (sort_key, item_id)
        pairs = []
        for item_id in children:
            vals = tree.item(item_id)["values"]
            key = self._paste_compare_sort_key(vals, col_index)
            pairs.append((key, item_id))
        pairs.sort(key=lambda p: p[0], reverse=self.paste_compare_sort_reverse)
        for index, (_, item_id) in enumerate(pairs):
            tree.move(item_id, "", index)
    
    def run_paste_compare(self):
        """Parse pasted lines; Reprocessing: compare reprocess vs sell; Manufacturing: profit for 1/10/100 runs at ME 0 and 10%."""
        text = self.paste_compare_text.get(1.0, tk.END)
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        if not lines:
            messagebox.showinfo("Paste & Compare", "Paste some lines (Name<Tab>Quantity or blueprint names) first.")
            return
        
        is_mfg = self.paste_compare_mode_var.get() == "manufacturing"
        if is_mfg:
            self.status_var.set("Calculating manufacturing profit...")
            for item in self.paste_compare_tree_mfg.get_children():
                self.paste_compare_tree_mfg.delete(item)
            self.paste_compare_tree_mfg.insert('', tk.END, values=("", "Calculating...", "", "", "", "", ""))
        else:
            self.status_var.set("Comparing items...")
            for item in self.paste_compare_tree.get_children():
                self.paste_compare_tree.delete(item)
            self.paste_compare_tree.insert('', tk.END, values=("", "Comparing...", "", "", "", ""))
        self.root.update()
        
        def do_compare():
            try:
                if self.paste_compare_mode_var.get() == "manufacturing":
                    try:
                        self._run_paste_compare_manufacturing(lines)
                    except Exception as e:
                        for item in self.paste_compare_tree_mfg.get_children():
                            self.paste_compare_tree_mfg.delete(item)
                        self.paste_compare_tree_mfg.insert('', tk.END, values=("", f"Error: {str(e)}", "", "", "", "", ""))
                        self.status_var.set("Error occurred")
                        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
                    return
                threshold = self.get_float(self.paste_threshold_var, 100000.0)
                sell_buffer_isk = self.get_float(self.paste_sell_buffer_var, 0.0)
                if sell_buffer_isk < 0:
                    sell_buffer_isk = 0.0
                yield_pct = self.get_float(self.paste_yield_var, 55.0)
                repro_cost_pct = self.get_float(self.paste_repro_cost_var, 3.37)
                
                rows = []
                for line in lines:
                    parts = line.split('\t')
                    name = parts[0].strip() if parts else ""
                    if not name:
                        continue
                    try:
                        qty_str = parts[1].strip() if len(parts) > 1 else "1"
                        qty = int(qty_str) if qty_str else 1
                    except (ValueError, IndexError):
                        qty = 1
                    
                    conn = sqlite3.connect(DATABASE_FILE)
                    try:
                        cursor = conn.execute("SELECT typeID FROM items WHERE typeName = ?", (name,))
                        row_item = cursor.fetchone()
                        if not row_item:
                            rows.append((name, str(qty), "N/A", "N/A", "N/A", "Not in DB"))
                            continue
                        type_id = row_item[0]
                        
                        cursor = conn.execute("SELECT buy_max, sell_min FROM prices WHERE typeID = ?", (type_id,))
                        price_row = cursor.fetchone()
                        buy_max = float(price_row[0]) if price_row and price_row[0] is not None else 0.0
                        sell_min = float(price_row[1]) if price_row and price_row[1] is not None else 0.0
                    finally:
                        conn.close()
                    
                    result = calculate_reprocessing_value(
                        module_type_id=type_id,
                        yield_percent=yield_pct,
                        buy_order_markup_percent=0,
                        reprocessing_cost_percent=repro_cost_pct,
                        module_price_type='sell_min',
                        mineral_price_type='sell_immediate',
                        db_file=DATABASE_FILE
                    )
                    
                    if 'error' in result:
                        rows.append((name, str(qty), f"{sell_min:,.2f}" if sell_min else "N/A", f"{buy_max:,.2f}" if buy_max else "N/A", "N/A", "Not reprocessable"))
                        continue
                    
                    total_mineral = result['total_mineral_value_per_job_after_costs']
                    repro_cost_job = result['reprocessing_cost_per_job']
                    input_qty = result['input_quantity']
                    if input_qty and input_qty > 0:
                        reprocess_value_per_item = (total_mineral - repro_cost_job) / input_qty
                    else:
                        reprocess_value_per_item = 0.0
                    
                    # Sell value for comparison: sell_min if above threshold, else buy_max
                    if sell_min >= threshold:
                        compare_price = sell_min
                    else:
                        compare_price = buy_max
                    
                    # (sell value - reprocess value) * Qty = total ISK advantage of selling; only recommend Sell if >= sell_buffer_isk
                    advantage_isk = (compare_price - reprocess_value_per_item) * qty if compare_price > 0 else 0.0
                    
                    if compare_price <= 0:
                        rec = "N/A (no price)"
                    elif reprocess_value_per_item > compare_price:
                        rec = "Reprocess"
                    elif advantage_isk < sell_buffer_isk:
                        rec = "Reprocess"
                    else:
                        rec = "Sell"
                    
                    sell_str = f"{sell_min:,.2f}" if sell_min else "N/A"
                    buy_str = f"{buy_max:,.2f}" if buy_max else "N/A"
                    repro_str = f"{reprocess_value_per_item:,.2f}"
                    rows.append((name, str(qty), sell_str, buy_str, repro_str, rec))
                
                for item in self.paste_compare_tree.get_children():
                    self.paste_compare_tree.delete(item)
                for r in rows:
                    self.paste_compare_tree.insert('', tk.END, values=r)
                self.status_var.set("Compare complete.")
            except Exception as e:
                for item in self.paste_compare_tree.get_children():
                    self.paste_compare_tree.delete(item)
                self.paste_compare_tree.insert('', tk.END, values=("", f"Error: {str(e)}", "", "", "", ""))
                self.status_var.set("Error occurred")
                messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        
        thread = threading.Thread(target=do_compare, daemon=True)
        thread.start()
    
    def _run_paste_compare_manufacturing(self, lines):
        """Run manufacturing profitability for each pasted blueprint; 1/10/100 runs at ME 0 and 10%. Runs in caller's thread (do_compare)."""
        system_cost_pct = self.get_float(self.paste_compare_system_cost_var, 8.61)
        if system_cost_pct < 0:
            system_cost_pct = 0.0
        region_id = MARKET_HISTORY_REGION_ID
        scenarios = [(1, 0), (1, 10), (10, 0), (10, 10), (100, 0), (100, 10)]  # (runs, me_percent)
        rows = []
        for line in lines:
            name = line.split('\t')[0].strip() if line else ""
            if not name:
                continue
            profits = []
            for runs, me in scenarios:
                result = calculate_blueprint_profitability(
                    blueprint_name_or_product=name,
                    input_price_type="buy_immediate",
                    output_price_type="sell_immediate",
                    system_cost_percent=system_cost_pct,
                    material_efficiency=me,
                    number_of_runs=runs,
                    region_id=region_id,
                    db_file=DATABASE_FILE,
                )
                if "error" in result:
                    profits.append("N/A")
                else:
                    profits.append(f"{result['profit']:,.0f}")
            rows.append((name,) + tuple(profits))
        for item in self.paste_compare_tree_mfg.get_children():
            self.paste_compare_tree_mfg.delete(item)
        for r in rows:
            self.paste_compare_tree_mfg.insert('', tk.END, values=r)
        self.status_var.set("Manufacturing compare complete.")
    
    def clear_paste_compare_text(self):
        """Clear the paste content text area so you can paste new content."""
        self.paste_compare_text.delete(1.0, tk.END)

    def create_planning_tab(self):
        """Blueprint Planning tab: paste available blueprints, see T1 and T2 profitability, add to shopping list."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Planning")
        info = ttk.LabelFrame(frame, text="Blueprint planification", padding=10)
        info.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(
            info,
            text="Paste your available blueprints (one name per line). The table shows T1 manufacturing profit and T2 invention+manufacturing profit (best decryptor). "
                 "Use 'Manage T2 mapping' when a blueprint has no T2 in the database, or 'Add selected to Shopping List' to add chosen blueprints.",
            justify=tk.LEFT, wraplength=900
        ).pack(anchor=tk.W)
        paste_frame = ttk.LabelFrame(frame, text="Available blueprints (one per line)", padding=10)
        paste_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.planning_paste_text = scrolledtext.ScrolledText(paste_frame, wrap=tk.WORD, height=6, width=80)
        self.planning_paste_text.pack(fill=tk.BOTH, expand=True)
        btn_row = ttk.Frame(paste_frame)
        btn_row.pack(fill=tk.X, pady=5)
        ttk.Button(btn_row, text="Analyze blueprints", command=self._planning_run_analysis).pack(side=tk.LEFT, padx=5)
        self.planning_status_var = tk.StringVar(value="Paste blueprints and click Analyze.")
        ttk.Label(btn_row, textvariable=self.planning_status_var).pack(side=tk.LEFT, padx=10)
        results_frame = ttk.LabelFrame(frame, text="Results (sort by clicking column headers)", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        cols = ("Blueprint", "Tech", "T1 profit ISK/run", "T1 profit %", "T2 product", "T2 decryptor", "T2 profit ISK/run", "T2 profit %", "Notes")
        self.planning_tree = ttk.Treeview(results_frame, columns=cols, show="headings", height=14, selectmode="extended")
        self.planning_sort_column = None
        self.planning_sort_reverse = False
        for c in cols:
            self.planning_tree.heading(c, text=c, command=lambda col=c: self._planning_sort_by(col))
        for c in cols:
            self.planning_tree.column(c, width=100, stretch=True)
        self.planning_tree.column("Blueprint", width=220)
        self.planning_tree.column("Notes", width=180)
        scroll_planning = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.planning_tree.yview)
        self.planning_tree.configure(yscrollcommand=scroll_planning.set)
        self.planning_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_planning.pack(side=tk.RIGHT, fill=tk.Y)
        self.planning_tree.bind("<Button-1>", self._on_planning_tree_click)
        act_row = ttk.Frame(results_frame)
        act_row.pack(fill=tk.X, pady=5)
        ttk.Button(act_row, text="Add selected to Shopping List", command=self._planning_add_to_shopping_list).pack(side=tk.LEFT, padx=5)
        ttk.Button(act_row, text="Manage T2 mapping", command=self._planning_manage_t2_mapping).pack(side=tk.LEFT, padx=5)
        self._planning_row_data = []  # list of dicts, index = tree row order after last analysis
        self._planning_load_paste_prefs()

    def _planning_load_paste_prefs(self):
        """Load last pasted blueprint list from prefs into the planning paste area."""
        if not LAUNCHER_PREFS_FILE.exists():
            return
        try:
            with open(LAUNCHER_PREFS_FILE, "r", encoding="utf-8") as f:
                prefs = json.load(f)
            paste = prefs.get("planning_paste")
            if paste and isinstance(paste, str):
                self.planning_paste_text.delete(1.0, tk.END)
                self.planning_paste_text.insert(tk.END, paste)
        except Exception:
            pass

    def _planning_save_paste_prefs(self, text):
        """Save current planning paste content to prefs."""
        try:
            prefs = {}
            if LAUNCHER_PREFS_FILE.exists():
                try:
                    with open(LAUNCHER_PREFS_FILE, "r", encoding="utf-8") as f:
                        prefs = json.load(f)
                except Exception:
                    pass
            prefs["planning_paste"] = text or ""
            with open(LAUNCHER_PREFS_FILE, "w", encoding="utf-8") as f:
                json.dump(prefs, f, indent=2)
        except Exception:
            pass

    def _planning_sort_by(self, column):
        """Sort Planning tree by the given column."""
        tree = self.planning_tree
        children = list(tree.get_children(""))
        if not children:
            return
        cols = ("Blueprint", "Tech", "T1 profit ISK/run", "T1 profit %", "T2 product", "T2 decryptor", "T2 profit ISK/run", "T2 profit %", "Notes")
        if column not in cols:
            return
        if self.planning_sort_column == column:
            self.planning_sort_reverse = not self.planning_sort_reverse
        else:
            self.planning_sort_reverse = False
            self.planning_sort_column = column
        col_index = cols.index(column)
        def sort_key(iid):
            vals = tree.item(iid)["values"]
            if col_index >= len(vals):
                return (0, "")
            v = vals[col_index]
            s = str(v).strip().replace(",", "")
            if col_index in (2, 3, 6, 7):  # numeric
                try:
                    return (0, float(s) if s else 0.0)
                except ValueError:
                    return (1, (v or "").lower())
            return (0, (v or "").lower())
        pairs = [(sort_key(iid), iid) for iid in children]
        pairs.sort(key=lambda p: p[0], reverse=self.planning_sort_reverse)
        for idx, (_, iid) in enumerate(pairs):
            tree.move(iid, "", idx)

    def _on_planning_tree_click(self, event):
        """Copy blueprint or product name to clipboard when clicking on Blueprint column."""
        region = self.planning_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.planning_tree.identify_column(event.x)
        if not col:
            return
        try:
            col_idx = int(col.replace("#", "")) - 1
        except ValueError:
            return
        if col_idx != 0:
            return
        item = self.planning_tree.identify_row(event.y)
        if not item:
            return
        vals = self.planning_tree.item(item, "values")
        if vals and len(vals) > 0 and vals[0]:
            self.root.clipboard_clear()
            self.root.clipboard_append(str(vals[0]).strip())
            self.status_var.set("Copied to clipboard.")

    def _planning_run_analysis(self):
        """Start planning analysis in a background thread."""
        text = self.planning_paste_text.get(1.0, tk.END)
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        if not lines:
            messagebox.showinfo("Planning", "Paste at least one blueprint name per line.")
            return
        self._planning_save_paste_prefs(text)
        self.planning_status_var.set("Analyzing...")
        for item in self.planning_tree.get_children():
            self.planning_tree.delete(item)
        self.planning_tree.insert("", tk.END, values=("", "Analyzing...", "", "", "", "", "", "", ""))
        self.root.update()

        def do_analysis():
            try:
                result = self._planning_analyze_blueprints(lines)
                self.root.after(0, lambda: self._planning_apply_results(result))
            except Exception as e:
                self.root.after(0, lambda: self._planning_apply_error(str(e)))

        threading.Thread(target=do_analysis, daemon=True).start()

    def _planning_apply_error(self, err_msg):
        """Apply error state to planning UI."""
        for item in self.planning_tree.get_children():
            self.planning_tree.delete(item)
        self.planning_tree.insert("", tk.END, values=("", f"Error: {err_msg}", "", "", "", "", "", "", ""))
        self._planning_row_data = []
        self.planning_status_var.set("Error occurred.")

    def _planning_apply_results(self, rows):
        """Populate planning tree and store row data from analysis results."""
        for item in self.planning_tree.get_children():
            self.planning_tree.delete(item)
        self._planning_row_data = rows
        for r in rows:
            t1_isk = f"{r.get('t1_profit_isk') or 0:,.0f}" if r.get('t1_profit_isk') is not None else "—"
            t1_pct = f"{r.get('t1_profit_pct') or 0:.1f}%" if r.get('t1_profit_pct') is not None else "—"
            t2_isk = f"{r.get('t2_profit_isk') or 0:,.0f}" if r.get('t2_profit_isk') is not None else "—"
            t2_pct = f"{r.get('t2_profit_pct') or 0:.1f}%" if r.get('t2_profit_pct') is not None else "—"
            self.planning_tree.insert("", tk.END, values=(
                r.get("t1_name") or "—",
                r.get("tech") or "—",
                t1_isk,
                t1_pct,
                r.get("t2_product") or "—",
                r.get("t2_decryptor") or "—",
                t2_isk,
                t2_pct,
                r.get("notes") or "",
            ))
        n = len(rows)
        with_t2 = sum(1 for r in rows if r.get("t2_product"))
        no_t2 = sum(1 for r in rows if r.get("notes") and "No T2" in str(r.get("notes")))
        self.planning_status_var.set(f"{n} blueprint(s) analyzed. {with_t2} with T2, {no_t2} without T2 mapping.")

    def _planning_analyze_blueprints(self, lines):
        """Compute T1 and T2 profitability for each pasted blueprint. Returns list of row dicts."""
        system_cost_pct = 8.61
        region_id = get_region_id_by_name(DEFAULT_REGION_NAME) if DEFAULT_REGION_NAME else MARKET_HISTORY_REGION_ID
        region_name = DEFAULT_REGION_NAME or "The Forge"
        input_price = "buy_immediate"
        output_price = "sell_offer"
        rows = []
        conn = sqlite3.connect(DATABASE_FILE)
        conn.row_factory = sqlite3.Row
        try:
            self._ensure_blueprint_datacore_bindings_table(conn)
            for line in lines:
                name = line.split("\t")[0].strip() if line else ""
                if not name:
                    continue
                row_data = {"t1_name": name, "tech": "—", "t1_profit_isk": None, "t1_profit_pct": None,
                            "t2_product": None, "t2_decryptor": None, "t2_decryptor_type_id": None,
                            "t2_profit_isk": None, "t2_profit_pct": None, "notes": "",
                            "t1_bp_id": None, "t2_bp_id": None, "best_decryptor_row": None, "runs_per_bpc": 10}
                bp = resolve_blueprint(conn, name)
                if not bp:
                    row_data["notes"] = "Blueprint not found"
                    rows.append(row_data)
                    continue
                row_data["t1_bp_id"] = bp["blueprintTypeID"]
                product_name = bp["productName"]
                t1_result = calculate_blueprint_profitability(
                    blueprint_name_or_product=product_name,
                    input_price_type=input_price,
                    output_price_type=output_price,
                    system_cost_percent=system_cost_pct,
                    material_efficiency=0,
                    number_of_runs=1,
                    region_id=region_id,
                    db_file=DATABASE_FILE,
                )
                if "error" in t1_result:
                    row_data["notes"] = t1_result["error"]
                    rows.append(row_data)
                    continue
                row_data["t1_profit_isk"] = t1_result["profit"]
                total_cost = t1_result["total_input_cost"] + t1_result["system_cost"]
                row_data["t1_profit_pct"] = (t1_result["profit"] / total_cost * 100.0) if total_cost and total_cost > 0 else 0.0
                try:
                    tech = conn.execute(
                        "SELECT techLevel, isFaction FROM items WHERE typeID = (SELECT productTypeID FROM blueprints WHERE blueprintTypeID = ?)",
                        (row_data["t1_bp_id"],)
                    ).fetchone()
                except Exception:
                    tech = None
                if tech:
                    tl, fac = tech[0], tech[1]
                    if fac:
                        row_data["tech"] = "Faction"
                    elif tl == 2:
                        row_data["tech"] = "T2"
                    else:
                        row_data["tech"] = "T1"
                t2_list = get_t2_products_from_t1(name, db_file=DATABASE_FILE)
                if not t2_list:
                    row_data["notes"] = "No T2 mapping"
                    rows.append(row_data)
                    continue
                best_t2_profit = None
                best_t2_name = None
                best_decryptor_row = None
                for t2_candidate in t2_list:
                    t2_name = t2_candidate["t2_product_name"]
                    t2_bp_id = t2_candidate["t2_blueprint_type_id"]
                    prob = t2_candidate.get("probability")
                    base_chance_pct = (float(prob) * 100.0) if prob is not None else 40.0
                    base_runs = int(t2_candidate.get("quantity") or 10)
                    if base_runs != 1:
                        base_runs = 10
                    inv_cost = 0.0
                    bind = conn.execute(
                        """SELECT dc1_name, dc1_qty, dc2_name, dc2_qty, base_invention_chance_pct, invention_cost_per_attempt, base_bpc_runs
                           FROM blueprint_datacore_bindings WHERE blueprint_type_id = ?""",
                        (t2_bp_id,),
                    ).fetchone()
                    datacores = []
                    if bind:
                        dc1, dq1, dc2, dq2 = bind[0], bind[1], bind[2], bind[3]
                        if dc1 and (dq1 or 0) > 0:
                            datacores.append((dc1, int(dq1 or 0)))
                        if dc2 and (dq2 or 0) > 0:
                            datacores.append((dc2, int(dq2 or 0)))
                        if len(bind) > 4 and bind[4] is not None:
                            base_chance_pct = float(bind[4])
                        if len(bind) > 5 and bind[5] is not None:
                            inv_cost = float(bind[5])
                        if len(bind) > 6 and bind[6] is not None:
                            base_runs = int(bind[6])
                    dec_results = compare_decryptor_profitability(
                        blueprint_name_or_product=t2_name,
                        base_invention_chance_pct=base_chance_pct,
                        invention_cost_without_decryptor=inv_cost,
                        base_bpc_runs=base_runs,
                        input_price_type=input_price,
                        output_price_type=output_price,
                        system_cost_percent=system_cost_pct,
                        region_id=region_id,
                        db_file=DATABASE_FILE,
                        datacores=datacores if datacores else None,
                    )
                    valid = [x for x in dec_results if not x.get("error")]
                    if not valid:
                        continue
                    best = max(valid, key=lambda x: x.get("profit_per_bpc") or -1e99)
                    profit_bpc = best.get("profit_per_bpc") or 0
                    if best_t2_profit is None or profit_bpc > best_t2_profit:
                        best_t2_profit = profit_bpc
                        best_t2_name = t2_name
                        best_decryptor_row = best
                if best_t2_name and best_decryptor_row is not None:
                    row_data["t2_product"] = best_t2_name
                    row_data["t2_decryptor"] = best_decryptor_row.get("decryptor_name") or "—"
                    row_data["t2_decryptor_type_id"] = best_decryptor_row.get("decryptor_type_id")
                    row_data["t2_profit_isk"] = best_t2_profit
                    row_data["t2_bp_id"] = next((t["t2_blueprint_type_id"] for t in t2_list if t["t2_product_name"] == best_t2_name), None)
                    row_data["best_decryptor_row"] = best_decryptor_row
                    row_data["runs_per_bpc"] = max(1, int(best_decryptor_row.get("bpc_runs") or 10))
                    tot_inv = best_decryptor_row.get("expected_inv_cost") or 0
                    mfg = best_decryptor_row.get("manufacturing_profit") or 0
                    row_data["t2_profit_pct"] = (best_t2_profit / tot_inv * 100.0) if tot_inv and tot_inv > 0 else 0.0
                else:
                    row_data["notes"] = "T2 data incomplete"
                rows.append(row_data)
        finally:
            conn.close()
        return rows

    def _planning_manage_t2_mapping(self):
        """Open dialog to associate T1 blueprint with a T2 product when no T2 mapping exists."""
        sel = list(self.planning_tree.selection())
        if not sel:
            messagebox.showinfo("Manage T2", "Select a row that has 'No T2 mapping', then click Manage T2 mapping.")
            return
        idx = 0
        for item in sel:
            children = list(self.planning_tree.get_children())
            if item not in children:
                continue
            try:
                idx = children.index(item)
            except ValueError:
                continue
            if idx >= len(self._planning_row_data):
                continue
            rd = self._planning_row_data[idx]
            if "No T2 mapping" not in str(rd.get("notes") or ""):
                messagebox.showinfo("Manage T2", "Selected row has a T2 mapping. Use Decryptor comparison tab to change it.")
                return
            t1_name = rd.get("t1_name") or ""
            if not t1_name:
                return
            t2_name = simpledialog.askstring("Associate T2", f"T1: {t1_name}\n\nEnter T2 product name to associate (invention output):", parent=self.root)
            if not t2_name or not t2_name.strip():
                return
            t2_name = t2_name.strip()
            if not Path(DATABASE_FILE).exists():
                messagebox.showerror("Manage T2", "Database not found.")
                return
            try:
                conn = sqlite3.connect(DATABASE_FILE)
                try:
                    self._ensure_invention_recipes_table(conn)
                    bp1 = resolve_blueprint(conn, t1_name)
                    bp2 = resolve_blueprint(conn, t2_name)
                    if not bp1:
                        messagebox.showerror("Manage T2", f"T1 not found: {t1_name!r}")
                        return
                    if not bp2:
                        messagebox.showerror("Manage T2", f"T2 not found: {t2_name!r}")
                        return
                    t1_bp_id = bp1["blueprintTypeID"]
                    t2_bp_id = bp2["blueprintTypeID"]
                    conn.execute(
                        """INSERT OR REPLACE INTO invention_recipes (t1_blueprint_type_id, t2_blueprint_type_id, quantity, probability)
                           VALUES (?, ?, 1, ?)""",
                        (t1_bp_id, t2_bp_id, 0.4),
                    )
                    conn.commit()
                    self.status_var.set(f"Associated {t1_name} → {t2_name}. Re-run Analyze to refresh.")
                    messagebox.showinfo("Manage T2", f"Saved. Re-run 'Analyze blueprints' to see T2 profit for {t1_name}.")
                finally:
                    conn.close()
            except Exception as e:
                messagebox.showerror("Manage T2", str(e))
            return
        messagebox.showinfo("Manage T2", "Select a row with 'No T2 mapping'.")

    def _planning_add_to_shopping_list(self):
        """Add selected planning rows to shopping list (T2 product + decryptor in same row as column)."""
        sel = list(self.planning_tree.selection())
        if not sel:
            messagebox.showinfo("Planning", "Select one or more rows, then click Add selected to Shopping List.")
            return
        children = list(self.planning_tree.get_children())
        to_add = []
        for item in sel:
            if item not in children:
                continue
            try:
                idx = children.index(item)
            except ValueError:
                continue
            if idx >= len(self._planning_row_data):
                continue
            rd = self._planning_row_data[idx]
            t2_product = rd.get("t2_product")
            if t2_product:
                product_name = t2_product
                runs_per_bpc = max(1, int(rd.get("runs_per_bpc") or 10))
                profit = rd.get("t2_profit_isk")
                dec_name = (rd.get("t2_decryptor") or "").strip()
                dec_type_id = rd.get("t2_decryptor_type_id")
                item = {"product_name": product_name, "quantity": 1, "profit": profit, "runs_per_bpc": runs_per_bpc}
                br = rd.get("best_decryptor_row")
                if br and br.get("success_prob_pct") is not None:
                    try:
                        p = float(br["success_prob_pct"]) / 100.0
                        if 0 < p <= 1.0:
                            item["invention_success_prob"] = p
                    except (TypeError, ValueError):
                        pass
                if dec_name and dec_name != "No decryptor" and dec_type_id:
                    item["decryptor_name"] = dec_name
                    item["decryptor_type_id"] = dec_type_id
                if br:
                    dc_isk = br.get("datacore_cost")
                    sp = br.get("success_prob_pct")
                    if dc_isk is not None and sp is not None:
                        try:
                            p = float(sp) / 100.0
                            if p > 0:
                                item["expected_datacore_cost_per_bpc"] = float(dc_isk) / p
                        except (TypeError, ValueError):
                            pass
                    if br.get("bpc_me") is not None:
                        try:
                            item["manufacturing_me"] = max(0, min(10, float(br["bpc_me"])))
                        except (TypeError, ValueError):
                            pass
                to_add.append(item)
            else:
                product_name = rd.get("t1_name")
                if not product_name:
                    continue
                try:
                    conn = sqlite3.connect(DATABASE_FILE)
                    try:
                        bp = resolve_blueprint(conn, product_name)
                        if bp:
                            product_name = bp["productName"]
                    finally:
                        conn.close()
                except Exception:
                    pass
                t1_result = calculate_blueprint_profitability(
                    blueprint_name_or_product=product_name,
                    input_price_type="buy_immediate",
                    output_price_type="sell_offer",
                    system_cost_percent=8.61,
                    material_efficiency=0,
                    number_of_runs=1,
                    db_file=DATABASE_FILE,
                )
                profit = t1_result.get("profit") if "error" not in t1_result else None
                to_add.append({"product_name": product_name, "quantity": 1, "profit": profit, "runs_per_bpc": 1})
        for entry in to_add:
            self._shopping_list_append_planning(entry)
        if to_add:
            self.planning_status_var.set(f"Added {len(to_add)} blueprint(s) to Shopping List.")
    
    def create_market_patterns_tab(self):
        """Market Patterns tab: run day-of-week price/volume analysis and show textual output."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Market Patterns")
        info = ttk.LabelFrame(frame, text="Day-of-week market analysis", padding=10)
        info.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(
            info,
            text="Runs analyze_market_patterns.py to compute average price, volume and expected buy volume per weekday\n"
                 "for core minerals and items in the On Offer list (region 10000002 / The Forge).",
            justify=tk.LEFT,
            wraplength=900,
        ).pack(anchor=tk.W)
        btn_row = ttk.Frame(frame)
        btn_row.pack(fill=tk.X, padx=10, pady=5)
        ttk.Button(btn_row, text="Run analysis", command=self._run_market_patterns_analysis).pack(side=tk.LEFT, padx=5)
        self.market_patterns_status_var = tk.StringVar(value="Click 'Run analysis' to generate report.")
        ttk.Label(btn_row, textvariable=self.market_patterns_status_var).pack(side=tk.LEFT, padx=10)
        text_frame = ttk.LabelFrame(frame, text="Analysis output", padding=10)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.market_patterns_text = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, height=25)
        self.market_patterns_text.pack(fill=tk.BOTH, expand=True)

    def _run_market_patterns_analysis(self):
        """Run analyze_market_patterns.py in a background thread and display its stdout."""
        self.market_patterns_status_var.set("Running market patterns analysis...")
        self.market_patterns_text.delete(1.0, tk.END)
        self.market_patterns_text.insert(tk.END, "Running analyze_market_patterns.py...\n\n")
        self.root.update_idletasks()

        def worker():
            try:
                result = subprocess.run(
                    [sys.executable, "analyze_market_patterns.py"],
                    cwd=Path(__file__).resolve().parent,
                    capture_output=True,
                    text=True,
                    timeout=600,
                )
                out = result.stdout or ""
                err = result.stderr or ""
                text = out
                if err:
                    if text:
                        text += "\n\n--- stderr ---\n"
                    text += err
                self.root.after(0, lambda: self._market_patterns_apply_result(text, result.returncode))
            except Exception as e:
                self.root.after(0, lambda: self._market_patterns_apply_result(f"Error: {e}", 1))

        threading.Thread(target=worker, daemon=True).start()

    def _market_patterns_apply_result(self, text: str, returncode: int):
        self.market_patterns_text.delete(1.0, tk.END)
        self.market_patterns_text.insert(tk.END, text or "(no output)")
        if returncode == 0:
            self.market_patterns_status_var.set("Market patterns analysis completed.")
        else:
            self.market_patterns_status_var.set("Market patterns analysis failed (see output).")

    def launch_overview_alert(self):
        """Launch overview_alert.py in a separate process (uses same Python executable)."""
        try:
            subprocess.Popen(
                [sys.executable, "overview_alert.py"],
                cwd=Path(__file__).resolve().parent,
            )
            self.status_var.set("Launched overview_alert.py.")
        except Exception as e:
            messagebox.showerror("Overview Alert", f"Failed to launch overview_alert.py:\n{e}")

    def create_sso_sync_tab(self):
        """EVE SSO sync tab: login and sync wallet transactions, journal, industry jobs for profitability tracking."""
        import os
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="EVE SSO Sync")
        info = ttk.LabelFrame(frame, text="Instructions", padding=10)
        info.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(
            info,
            text="Create an SSO application at https://developers.eveonline.com/ with callback URL: http://localhost:8765/callback/\n"
                 "Request scopes: esi-wallet.read_character_wallet.v1 and esi-industry.read_character_jobs.v1",
            justify=tk.LEFT, wraplength=900
        ).pack(anchor=tk.W)
        creds = ttk.LabelFrame(frame, text="SSO credentials", padding=10)
        creds.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(creds, text="Client ID:").pack(side=tk.LEFT, padx=5)
        self.sso_client_id_var = tk.StringVar(value=os.environ.get("EVE_SSO_CLIENT_ID", ""))
        ttk.Entry(creds, textvariable=self.sso_client_id_var, width=24).pack(side=tk.LEFT, padx=5)
        ttk.Label(creds, text="Client Secret:").pack(side=tk.LEFT, padx=5)
        self.sso_client_secret_var = tk.StringVar(value=os.environ.get("EVE_SSO_CLIENT_SECRET", ""))
        ttk.Entry(creds, textvariable=self.sso_client_secret_var, width=32, show="*").pack(side=tk.LEFT, padx=5)
        btn_row = ttk.Frame(frame)
        btn_row.pack(fill=tk.X, padx=10, pady=5)
        ttk.Button(btn_row, text="Login with EVE SSO", command=self.sso_login).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Sync wallet & industry jobs", command=self.sso_sync).pack(side=tk.LEFT, padx=5)
        self.sso_status_var = tk.StringVar(value="Not logged in.")
        ttk.Label(frame, textvariable=self.sso_status_var).pack(anchor=tk.W, padx=10, pady=2)
        log_frame = ttk.LabelFrame(frame, text="Log", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.sso_log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=12, width=80)
        self.sso_log_text.pack(fill=tk.BOTH, expand=True)
    
    def _sso_log(self, msg: str):
        self.sso_log_text.insert(tk.END, msg + "\n")
        self.sso_log_text.see(tk.END)
        self.root.update_idletasks()
    
    def sso_login(self):
        """Run EVE SSO login flow (open browser, callback server, store tokens)."""
        from eve_sso_sync import login_flow
        cid = self.sso_client_id_var.get().strip()
        secret = self.sso_client_secret_var.get().strip()
        if not cid or not secret:
            messagebox.showwarning("SSO", "Enter Client ID and Client Secret (or set EVE_SSO_CLIENT_ID and EVE_SSO_CLIENT_SECRET).")
            return
        self.sso_status_var.set("Opening browser for EVE login...")
        self._sso_log("Starting SSO login...")
        def run():
            try:
                result = login_flow(cid, secret, DATABASE_FILE)
                if "error" in result:
                    self.sso_status_var.set("Login failed.")
                    self._sso_log("Error: " + result["error"])
                    messagebox.showerror("SSO Login", result["error"])
                else:
                    name = result.get("character_name") or f"Character {result.get('character_id')}"
                    self.sso_status_var.set(f"Logged in: {name}")
                    self._sso_log(f"Logged in: {name} (character_id={result.get('character_id')})")
            except Exception as e:
                self.sso_status_var.set("Login failed.")
                self._sso_log("Error: " + str(e))
                messagebox.showerror("SSO Login", str(e))
        threading.Thread(target=run, daemon=True).start()
    
    def sso_sync(self):
        """Sync wallet transactions, journal, and industry jobs for the stored character."""
        from eve_sso_sync import run_full_sync, ensure_sso_tables
        cid = self.sso_client_id_var.get().strip()
        secret = self.sso_client_secret_var.get().strip()
        if not cid or not secret:
            messagebox.showwarning("SSO", "Enter Client ID and Client Secret first.")
            return
        if not Path(DATABASE_FILE).exists():
            messagebox.showwarning("SSO", "Database not found. Create it first (e.g. build_database).")
            return
        self.sso_status_var.set("Syncing...")
        self._sso_log("Starting sync...")
        def run():
            try:
                conn = sqlite3.connect(DATABASE_FILE)
                try:
                    ensure_sso_tables(conn)
                    row = conn.execute("SELECT character_id FROM sso_character LIMIT 1").fetchone()
                    if not row:
                        self.sso_status_var.set("Not logged in.")
                        self._sso_log("No character found. Log in with EVE SSO first.")
                        return
                    character_id = row[0]
                    result = run_full_sync(conn, character_id, cid, secret)
                    conn.close()
                except Exception:
                    conn.close()
                    raise
                if "error" in result and result.get("tx", 0) == 0 and result.get("journal", 0) == 0 and result.get("jobs", 0) == 0:
                    self.sso_status_var.set("Sync failed.")
                    self._sso_log("Error: " + result["error"])
                    messagebox.showerror("SSO Sync", result["error"])
                else:
                    self.sso_status_var.set("Sync complete.")
                    self._sso_log(f"Synced: {result.get('tx', 0)} transactions, {result.get('journal', 0)} journal entries, {result.get('jobs', 0)} industry jobs.")
                    if result.get("error"):
                        self._sso_log("Note: " + result["error"])
            except Exception as e:
                self.sso_status_var.set("Sync failed.")
                self._sso_log("Error: " + str(e))
                messagebox.showerror("SSO Sync", str(e))
        threading.Thread(target=run, daemon=True).start()
    
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
                # Format date as dd/mm
                excluded_at_str = excluded_at
                if excluded_at:
                    try:
                        from datetime import datetime as _dt
                        d = _dt.strptime(str(excluded_at)[:10], "%Y-%m-%d")
                        excluded_at_str = f"{d.day:02d}/{d.month:02d}"
                    except Exception:
                        excluded_at_str = str(excluded_at)
                self.exclusions_tree.insert('', tk.END, values=(
                    module_name,
                    module_type_id,
                    f"{min_price:,.2f}",
                    f"{max_price:,.2f}",
                    module_price_type,
                    mineral_price_type,
                    excluded_at_str
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
        """Run the top 30 analysis in a separate thread. Updates mineral prices first, then runs analysis."""
        self.status_var.set("Updating mineral prices, then running analysis...")
        # Clear results table
        for item in self.analysis_tree.get_children():
            self.analysis_tree.delete(item)
        self.analysis_tree.insert('', tk.END, values=("", "Updating mineral prices first, then running analysis...", "", "", "", "", "", "", ""))
        self.root.update()
        
        def analyze():
            try:
                # Update mineral prices first (analysis uses mineral prices)
                self.status_var.set("Updating mineral prices...")
                children = list(self.analysis_tree.get_children())
                if children:
                    self.analysis_tree.item(children[0], values=("", "Updating mineral prices...", "", "", "", "", "", "", ""))
                update_mineral_prices()
                
                self.status_var.set("Running analysis...")
                for item in self.analysis_tree.get_children():
                    self.analysis_tree.delete(item)
                self.analysis_tree.insert('', tk.END, values=("", "Running analysis... This may take several minutes.", "", "", "", "", "", "", ""))
                
                yield_percent = self.get_float(self.yield_var, 55.0)
                markup_percent = self.get_float(self.markup_var, 10.0)
                reprocessing_cost = self.get_float(self.reprocessing_cost_var, 3.37)
                min_price = self.get_float(self.min_price_var, 1.0)
                max_price = self.get_float(self.max_price_var, 1000000.0)
                top_n = self.get_int(self.top_n_var, 30)
                min_expected_volume = self.get_float(self.min_expected_volume_var, 0.0)
                if min_expected_volume < 0:
                    min_expected_volume = 0.0
                module_price_type = self.module_price_type_var.get()
                mineral_price_type = self.mineral_price_type_var.get()
                sort_by = self.sort_by_var.get()
                sort_by_profit = sort_by in ("profit", "expected_profit")
                
                # Map "Run on" UI to backend filter
                run_on = self.item_source_filter_var.get()
                if run_on == "Blueprint items only":
                    item_source_filter = "blueprint"
                elif run_on == "Group consensus items only":
                    item_source_filter = "group_consensus"
                else:
                    item_source_filter = "all"
                
                # Get excluded modules for this search
                excluded_modules = self.get_excluded_modules(
                    min_price, max_price, module_price_type, mineral_price_type
                )
                
                # Check which sources to exclude
                excluded_sources = []
                if self.exclude_default_var.get():
                    excluded_sources.append('default')
                if self.exclude_group_consensus_var.get():
                    excluded_sources.append('group_consensus')
                if self.exclude_group_most_frequent_var.get():
                    excluded_sources.append('group_most_frequent')
                
                # Request more results when we'll filter (by source, expected profit, or min expected volume)
                effective_top_n = top_n * 10 if (excluded_sources or sort_by == "expected_profit" or min_expected_volume > 0) else top_n
                
                results = analyze_all_modules(
                    yield_percent=yield_percent,
                    buy_order_markup_percent=markup_percent,
                    reprocessing_cost_percent=reprocessing_cost,
                    module_price_type=module_price_type,
                    mineral_price_type=mineral_price_type,
                    min_module_price=min_price,
                    max_module_price=max_price,
                    top_n=effective_top_n,
                    excluded_module_ids=excluded_modules,
                    sort_by='profit' if sort_by_profit else 'return',
                    item_source_filter=item_source_filter
                )
                
                # Filter results based on source exclusion checkboxes
                if excluded_sources:
                    # Keep results where source is NOT in excluded_sources
                    results = [r for r in results if r.get('input_quantity_source', 'unknown') not in excluded_sources]
                
                # Enrich with expected volume when sorting by expected profit OR when filtering by min expected volume
                need_expected_volume = (sort_by == "expected_profit" or min_expected_volume > 0) and Path(DATABASE_FILE).exists()
                if need_expected_volume:
                    conn = sqlite3.connect(DATABASE_FILE)
                    try:
                        for r in results:
                            avg_7, as_of_7 = get_expected_buy_order_volume_7d_avg(
                                conn, MARKET_HISTORY_REGION_ID, r["module_type_id"]
                            )
                            avg_30, as_of_30 = get_expected_buy_order_volume_30d_avg(
                                conn, MARKET_HISTORY_REGION_ID, r["module_type_id"]
                            )
                            r["expected_volume_7d"] = avg_7
                            r["expected_volume_30d"] = avg_30
                            r["expected_volume_as_of"] = as_of_7 or as_of_30
                            effective_vol = None
                            if avg_7 is not None and avg_7 > 0:
                                effective_vol = avg_7
                            elif avg_30 is not None and avg_30 > 0:
                                effective_vol = avg_30 / 2.0
                            r["expected_volume_effective"] = effective_vol
                            r["expected_profit"] = (effective_vol * r["profit_per_item"]) if effective_vol is not None else 0
                    finally:
                        conn.close()
                
                # Filter by minimum expected volume (only items with expected_volume_effective >= min_expected_volume)
                if min_expected_volume > 0:
                    results = [r for r in results if r.get("expected_volume_effective") is not None and r["expected_volume_effective"] >= min_expected_volume]
                
                # Sort and take top N
                if sort_by == "expected_profit":
                    results.sort(key=lambda x: (x.get("expected_profit") is None, -(x.get("expected_profit") or 0)))
                    results = results[:top_n]
                elif excluded_sources or min_expected_volume > 0:
                    if sort_by_profit:
                        results.sort(key=lambda x: x.get('profit_per_item', 0), reverse=True)
                    else:
                        results.sort(key=lambda x: x.get('return_percent', 0), reverse=True)
                    results = results[:top_n]
                else:
                    results = results[:top_n]
                
                # Enrich all results with expected volume for display (if not already set)
                if results and Path(DATABASE_FILE).exists() and results[0].get("expected_volume_effective") is None:
                    conn = sqlite3.connect(DATABASE_FILE)
                    try:
                        for r in results:
                            avg_7, as_of_7 = get_expected_buy_order_volume_7d_avg(
                                conn, MARKET_HISTORY_REGION_ID, r["module_type_id"]
                            )
                            avg_30, as_of_30 = get_expected_buy_order_volume_30d_avg(
                                conn, MARKET_HISTORY_REGION_ID, r["module_type_id"]
                            )
                            r["expected_volume_7d"] = avg_7
                            r["expected_volume_30d"] = avg_30
                            r["expected_volume_as_of"] = as_of_7 or as_of_30
                            effective_vol = (avg_7 if (avg_7 is not None and avg_7 > 0) else
                                            (avg_30 / 2.0 if (avg_30 is not None and avg_30 > 0) else None))
                            r["expected_volume_effective"] = effective_vol
                            r["expected_profit"] = (effective_vol * r["profit_per_item"]) if effective_vol is not None else 0
                    finally:
                        conn.close()
                
                # Store results and parameters for exclusion
                self.last_analysis_results = results
                self.last_analysis_params = {
                    'min_price': min_price,
                    'max_price': max_price,
                    'module_price_type': module_price_type,
                    'mineral_price_type': mineral_price_type,
                    'min_expected_volume': min_expected_volume
                }
                
                # Get list of items in on_offer_items for highlighting
                on_offer_type_ids = self.get_on_offer_type_ids()
                
                # Clear and populate results table
                for item in self.analysis_tree.get_children():
                    self.analysis_tree.delete(item)
                
                for rank, result in enumerate(results, 1):
                    return_pct = result['return_percent']
                    if return_pct > 999999:
                        return_str = ">999,999%"
                    elif return_pct == float('inf'):
                        return_str = "N/A"
                    else:
                        return_str = f"{return_pct:,.2f}%"
                    
                    breakeven_price = result.get('breakeven_module_price', 'na')
                    if isinstance(breakeven_price, (int, float)) and breakeven_price not in (0, float('inf')):
                        breakeven_str = f"{breakeven_price:,.2f}"
                    else:
                        breakeven_str = "N/A"
                    
                    ev = result.get("expected_volume_effective") or result.get("expected_volume_7d")
                    expected_vol_str = f"{ev:,.0f}" if ev is not None else "N/A"
                    ep = result.get("expected_profit")
                    expected_profit_str = f"{ep:,.0f}" if ep is not None else "N/A"
                    
                    values = (
                        rank,
                        result['module_name'],
                        f"{result['expected_buy_price']:,.2f}",
                        f"{result['sell_min_price']:,.2f}",
                        f"{result['profit_per_item']:,.2f}",
                        return_str,
                        breakeven_str,
                        expected_vol_str,
                        expected_profit_str
                    )
                    item_id = self.analysis_tree.insert('', tk.END, values=values)
                    if result['module_type_id'] in on_offer_type_ids:
                        self.analysis_tree.item(item_id, tags=('on_offer',))
                
                self.status_var.set("Analysis complete!")
                
            except Exception as e:
                for item in self.analysis_tree.get_children():
                    self.analysis_tree.delete(item)
                self.analysis_tree.insert('', tk.END, values=("", f"Error: {str(e)}", "", "", "", "", "", "", ""))
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
    
    def get_on_offer_type_ids(self):
        """Get set of module type IDs that are in the on_offer_items table"""
        if not Path(DATABASE_FILE).exists():
            return set()
        
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT module_type_id FROM on_offer_items")
            results = cursor.fetchall()
            return {row[0] for row in results}
        finally:
            conn.close()
    
    def on_analysis_tree_double_click(self, event):
        """On double-click, copy the selected row's module name to clipboard."""
        selection = self.analysis_tree.selection()
        if not selection:
            return
        item = selection[0]
        values = self.analysis_tree.item(item, 'values')
        if len(values) >= 2:
            module_name = values[1]  # Module Name column
            if module_name and not module_name.startswith("Running") and not module_name.startswith("Error"):
                self.copy_module_name_to_clipboard(module_name)
    
    def copy_module_name_to_clipboard(self, module_name):
        """Copy module name to clipboard"""
        self.root.clipboard_clear()
        self.root.clipboard_append(module_name)
        self.status_var.set(f"Copied '{module_name}' to clipboard")
    
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
                reprocessing_cost = self.get_float(self.single_reprocessing_cost_var, 3.37)
                module_price_type = self.single_module_price_type_var.get()
                mineral_price_type = self.single_mineral_price_type_var.get()
                
                result = calculate_reprocessing_value(
                    module_name=module_name,
                    yield_percent=yield_percent,
                    buy_order_markup_percent=markup_percent,
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
    
    def _resolve_module_name_to_type_id(self, module_name):
        """Return (type_id, type_name) for exact typeName match, or (None, None) if not found."""
        if not module_name or not Path(DATABASE_FILE).exists():
            return (None, None)
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            cur = conn.execute("SELECT typeID, typeName FROM items WHERE typeName = ?", (module_name.strip(),))
            row = cur.fetchone()
            return (row[0], row[1]) if row else (None, None)
        finally:
            conn.close()
    
    def show_single_expected_volume(self):
        """Show 7d and 30d expected buy order volume for the module in the Single Module tab."""
        module_name = self.module_name_var.get().strip()
        if not module_name:
            messagebox.showwarning("Warning", "Please enter a module name")
            return
        type_id, resolved_name = self._resolve_module_name_to_type_id(module_name)
        if type_id is None:
            self.single_module_results.delete(1.0, tk.END)
            self.single_module_results.insert(tk.END, f"Module not found: {module_name!r}\n")
            return
        self.status_var.set("Refreshing from API, then loading expected volume...")
        self.single_module_results.delete(1.0, tk.END)
        self.single_module_results.insert(tk.END, f"Expected volume for: {resolved_name} (type_id={type_id})\n\nRefreshing from API...\n")
        self.root.update()
        
        def run():
            try:
                conn = sqlite3.connect(DATABASE_FILE)
                try:
                    n = refresh_market_history_for_type(conn, MARKET_HISTORY_REGION_ID, type_id)
                    avg_7, as_of_7 = get_expected_buy_order_volume_7d_avg(
                        conn, MARKET_HISTORY_REGION_ID, type_id
                    )
                    avg_30, as_of_30 = get_expected_buy_order_volume_30d_avg(
                        conn, MARKET_HISTORY_REGION_ID, type_id
                    )
                finally:
                    conn.close()
                lines = [
                    f"Expected volume for: {resolved_name} (type_id={type_id})",
                    f"Region: {MARKET_HISTORY_REGION_ID} (The Forge)",
                    f"Refreshed from API: {n} days of data.",
                    "",
                    "7-day average expected buy order volume:",
                    f"  {avg_7:,.0f}" if avg_7 is not None else "  N/A (no market history data)",
                    f"  Data as of: {as_of_7}" if as_of_7 else "",
                    "",
                    "30-day average expected buy order volume:",
                    f"  {avg_30:,.0f}" if avg_30 is not None else "  N/A (no market history data)",
                    f"  Data as of: {as_of_30}" if as_of_30 else "",
                ]
                self.single_module_results.delete(1.0, tk.END)
                self.single_module_results.insert(tk.END, "\n".join(lines))
                self.status_var.set("Expected volume loaded")
            except Exception as e:
                self.single_module_results.delete(1.0, tk.END)
                self.single_module_results.insert(tk.END, f"Error: {str(e)}\n")
                self.status_var.set("Error occurred")
        thread = threading.Thread(target=run, daemon=True)
        thread.start()
    
    def show_single_raw_market_data(self):
        """Show raw market_history_daily rows for the module in the Single Module tab."""
        module_name = self.module_name_var.get().strip()
        if not module_name:
            messagebox.showwarning("Warning", "Please enter a module name")
            return
        type_id, resolved_name = self._resolve_module_name_to_type_id(module_name)
        if type_id is None:
            self.single_module_results.delete(1.0, tk.END)
            self.single_module_results.insert(tk.END, f"Module not found: {module_name!r}\n")
            return
        self.status_var.set("Loading raw market data...")
        self.single_module_results.delete(1.0, tk.END)
        self.single_module_results.insert(tk.END, f"Raw market data for: {resolved_name} (type_id={type_id})\n\nLoading...\n")
        self.root.update()
        
        def run():
            try:
                conn = sqlite3.connect(DATABASE_FILE)
                try:
                    rows = get_market_history_raw(conn, MARKET_HISTORY_REGION_ID, type_id, limit=60)
                finally:
                    conn.close()
                if not rows:
                    self.single_module_results.delete(1.0, tk.END)
                    self.single_module_results.insert(
                        tk.END,
                        f"Raw market data for: {resolved_name} (type_id={type_id})\n\nNo market history data for region {MARKET_HISTORY_REGION_ID}.\n"
                    )
                    self.status_var.set("No data")
                    return
                header = f"{'date_utc':<12} {'lowest':>12} {'highest':>12} {'average':>12} {'volume':>12} {'exp_buy_vol':>12}"
                lines = [
                    f"Raw market data for: {resolved_name} (type_id={type_id})",
                    f"Region: {MARKET_HISTORY_REGION_ID} (most recent 60 days)",
                    "",
                    header,
                    "-" * 76,
                ]
                for r in rows:
                    lines.append(
                        f"{r['date_utc']:<12} {r['lowest'] or 0:>12,.2f} {r['highest'] or 0:>12,.2f} "
                        f"{r['average'] or 0:>12,.2f} {r['volume'] or 0:>12,.0f} {r['expected_buy_order_vol']:>12,.0f}"
                    )
                self.single_module_results.delete(1.0, tk.END)
                self.single_module_results.insert(tk.END, "\n".join(lines))
                self.status_var.set("Raw market data loaded")
            except Exception as e:
                self.single_module_results.delete(1.0, tk.END)
                self.single_module_results.insert(tk.END, f"Error: {str(e)}\n")
                self.status_var.set("Error occurred")
        thread = threading.Thread(target=run, daemon=True)
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
        # Use edited units if available, otherwise use input_quantity
        units_value = result.get('_edited_units_required', result.get('input_quantity', 1))
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
            # Use QuantityAfterYield which may have been edited previously
            current_qty = float(output_mat.get('QuantityAfterYield', 0))
            # Calculate per module: quantity after yield / input_quantity
            input_qty = result.get('input_quantity', 1)
            per_module = current_qty / input_qty if input_qty > 0 else 0
            price = output_mat.get('mineralPriceAfterCosts', output_mat.get('mineralPrice', 0))
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
                # Use base price and apply costs
                module_price_base = result.get('module_price', 0)
                module_price_after_costs = result.get('module_price_after_costs', module_price_base)
                
                # Calculate total cost: base price × units, then apply cost factor
                cost_factor = module_price_after_costs / module_price_base if module_price_base > 0 else 1.0
                base_total_cost = module_price_base * actual_modules_needed
                total_module_price = base_total_cost * cost_factor
                
                # Recalculate reprocessing cost
                effective_reprocessing_cost_percent = result['reprocessing_cost_percent'] * (result['yield_percent'] / 100.0)
                reprocessing_cost = total_module_price * (effective_reprocessing_cost_percent / 100.0)
                
                # Recalculate total mineral value from edited quantities
                total_mineral_value = 0.0
                for item in tree.get_children():
                    values = tree.item(item, 'values')
                    edited_qty = float(values[2].replace(',', ''))
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
                    profit_margin_percent = ((reprocessing_value / total_module_price) - 1) * 100
                else:
                    profit_margin_percent = "na"
                
                # Update result
                updated_result = result.copy()
                updated_result['input_quantity'] = actual_modules_needed
                updated_result['total_module_cost_per_job'] = total_module_price
                updated_result['reprocessing_cost_per_job'] = reprocessing_cost
                updated_result['total_mineral_value_per_job_after_costs'] = total_mineral_value
                updated_result['reprocessing_value_per_job_after_costs'] = reprocessing_value
                updated_result['profit_margin_percent'] = profit_margin_percent
                
                # Update reprocessing outputs with edited quantities
                for output_mat in updated_result['reprocessing_outputs']:
                    material_type_id = output_mat['materialTypeID']
                    
                    if material_type_id in edited_quantities:
                        edited_qty = edited_quantities[material_type_id]
                        output_mat['QuantityAfterYield'] = edited_qty
                        price_after_costs = output_mat.get('mineralPriceAfterCosts', output_mat.get('mineralPrice', 0))
                        output_mat['mineralValue'] = edited_qty * price_after_costs
                    else:
                        # Update quantities that weren't edited but need recalculation
                        # Recalculate based on new units_required
                        per_module = output_mat.get('baseQuantityPerModule', 0)
                        new_qty = int(per_module * actual_modules_needed)
                        output_mat['actualQuantity'] = new_qty
                        output_mat['mineralValue'] = new_qty * output_mat['mineralPrice']
                
                # Also update module_price in the result to reflect the recalculated price
                updated_result['module_price'] = module_price_base
                updated_result['module_price_after_costs'] = module_price_after_costs
                
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
                    f"Total Module Cost per Job: {total_module_price:,.2f} ISK\n"
                    f"Reprocessing Cost per Job: {reprocessing_cost:,.2f} ISK\n"
                    f"Total Mineral Value per Job (after costs): {total_mineral_value:,.2f} ISK\n"
                    f"Net Profit per Job: {reprocessing_value:,.2f} ISK\n"
                    f"Profit Margin: {profit_margin_percent:+.2f}%" if isinstance(profit_margin_percent, (int, float)) else "Profit Margin: N/A" if profit_margin_percent != "na" else "Profit Margin: N/A"
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
                
                comparison_report = update_mineral_prices()
                
                logger.removeHandler(handler)
                output = log_capture.getvalue()
                
                self.price_update_log.insert(tk.END, output)
                if comparison_report:
                    self.price_update_log.insert(tk.END, comparison_report)
                self.price_update_log.insert(tk.END, "\n\nMineral price update complete!\n")
                self.status_var.set("Mineral price update complete!")
                messagebox.showinfo("Success", "Mineral prices updated successfully!")
                
            except Exception as e:
                self.price_update_log.insert(tk.END, f"\nError: {str(e)}\n")
                self.status_var.set("Error occurred")
                messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        
        thread = threading.Thread(target=update, daemon=True)
        thread.start()
    
    def update_blueprint_prices(self):
        """Update prices only for items with blueprint source in a separate thread"""
        self.status_var.set("Updating blueprint item prices...")
        self.price_update_log.delete(1.0, tk.END)
        self.price_update_log.insert(tk.END, "Starting update of blueprint item prices...\n")
        self.price_update_log.insert(tk.END, "Finding items with blueprint source...\n\n")
        self.root.update()
        
        def update():
            try:
                # Get typeIDs for items with blueprint source
                conn = sqlite3.connect(DATABASE_FILE)
                try:
                    cursor = conn.execute("""
                        SELECT DISTINCT c.typeID 
                        FROM input_quantity_cache c
                        INNER JOIN prices p ON c.typeID = p.typeID
                        WHERE c.source = 'blueprint'
                    """)
                    type_ids = [row[0] for row in cursor.fetchall()]
                    
                    if not type_ids:
                        self.price_update_log.insert(tk.END, "No items with blueprint source found in database.\n")
                        self.status_var.set("No blueprint items found")
                        return
                    
                    self.price_update_log.insert(tk.END, f"Found {len(type_ids)} items with blueprint source.\n")
                    self.price_update_log.insert(tk.END, "Updating prices...\n\n")
                    self.root.update()
                    
                finally:
                    conn.close()
                
                # Redirect logging to our text widget
                import logging
                from io import StringIO
                
                log_capture = StringIO()
                handler = logging.StreamHandler(log_capture)
                handler.setLevel(logging.INFO)
                logger = logging.getLogger()
                logger.addHandler(handler)
                
                update_prices_by_type_ids(type_ids, f"blueprint items (source='blueprint')")
                
                logger.removeHandler(handler)
                output = log_capture.getvalue()
                
                self.price_update_log.insert(tk.END, output)
                self.price_update_log.insert(tk.END, "\n\nBlueprint price update complete!\n")
                self.status_var.set("Blueprint price update complete!")
                messagebox.showinfo("Success", f"Updated prices for {len(type_ids)} blueprint items successfully!")
                
            except Exception as e:
                self.price_update_log.insert(tk.END, f"\nError: {str(e)}\n")
                self.status_var.set("Error occurred")
                messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        
        thread = threading.Thread(target=update, daemon=True)
        thread.start()
    
    def update_group_consensus_prices(self):
        """Update prices only for items with group_consensus source in a separate thread"""
        self.status_var.set("Updating group consensus item prices...")
        self.price_update_log.delete(1.0, tk.END)
        self.price_update_log.insert(tk.END, "Starting update of group consensus item prices...\n")
        self.price_update_log.insert(tk.END, "Finding items with group consensus source...\n\n")
        self.root.update()
        
        def update():
            try:
                # Get typeIDs for items with group_consensus source
                conn = sqlite3.connect(DATABASE_FILE)
                try:
                    cursor = conn.execute("""
                        SELECT DISTINCT c.typeID 
                        FROM input_quantity_cache c
                        INNER JOIN prices p ON c.typeID = p.typeID
                        WHERE c.source = 'group_consensus'
                    """)
                    type_ids = [row[0] for row in cursor.fetchall()]
                    
                    if not type_ids:
                        self.price_update_log.insert(tk.END, "No items with group consensus source found in database.\n")
                        self.status_var.set("No group consensus items found")
                        return
                    
                    self.price_update_log.insert(tk.END, f"Found {len(type_ids)} items with group consensus source.\n")
                    self.price_update_log.insert(tk.END, "Updating prices...\n\n")
                    self.root.update()
                    
                finally:
                    conn.close()
                
                # Redirect logging to our text widget
                import logging
                from io import StringIO
                
                log_capture = StringIO()
                handler = logging.StreamHandler(log_capture)
                handler.setLevel(logging.INFO)
                logger = logging.getLogger()
                logger.addHandler(handler)
                
                update_prices_by_type_ids(type_ids, f"group consensus items (source='group_consensus')")
                
                logger.removeHandler(handler)
                output = log_capture.getvalue()
                
                self.price_update_log.insert(tk.END, output)
                self.price_update_log.insert(tk.END, "\n\nGroup consensus price update complete!\n")
                self.status_var.set("Group consensus price update complete!")
                messagebox.showinfo("Success", f"Updated prices for {len(type_ids)} group consensus items successfully!")
                
            except Exception as e:
                self.price_update_log.insert(tk.END, f"\nError: {str(e)}\n")
                self.status_var.set("Error occurred")
                messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        
        thread = threading.Thread(target=update, daemon=True)
        thread.start()
    
    def run_fetch_market_history_prices(self):
        """Fetch market history for the same type set as Update All Prices (long run)."""
        if not Path(DATABASE_FILE).exists():
            messagebox.showerror("Error", "Database not found")
            return
        self.status_var.set("Fetching market history (same set as Update All Prices)...")
        self.price_update_log.delete(1.0, tk.END)
        self.price_update_log.insert(tk.END, "Starting market history fetch (same types as Update All Prices).\n")
        self.price_update_log.insert(tk.END, "This can take a long time (~50k types with 1s delay).\n\n")
        self.root.update()
        
        def run():
            try:
                import logging
                from io import StringIO
                log_capture = StringIO()
                handler = logging.StreamHandler(log_capture)
                handler.setLevel(logging.INFO)
                root_logger = logging.getLogger()
                root_logger.addHandler(handler)
                try:
                    run_fetch(
                        region_id=MARKET_HISTORY_REGION_ID,
                        all_items=True,
                        scope="prices",
                        delay_seconds=1.0,
                        progress_interval=50,
                    )
                finally:
                    root_logger.removeHandler(handler)
                output = log_capture.getvalue()
                self.price_update_log.insert(tk.END, output)
                self.price_update_log.insert(tk.END, "\n\nMarket history fetch complete!\n")
                self.status_var.set("Market history fetch complete!")
                messagebox.showinfo("Success", "Market history fetch complete!")
            except Exception as e:
                self.price_update_log.insert(tk.END, f"\nError: {str(e)}\n")
                self.status_var.set("Error occurred")
                messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        threading.Thread(target=run, daemon=True).start()
    
    def refresh_volume_no_or_zero_data(self):
        """Refresh market history from API for types that have no data or zero expected volume."""
        if not Path(DATABASE_FILE).exists():
            messagebox.showerror("Error", "Database not found")
            return
        self.status_var.set("Finding items with no/zero volume data...")
        self.price_update_log.delete(1.0, tk.END)
        self.price_update_log.insert(tk.END, "Finding items with no or zero expected volume (prices set)...\n")
        self.root.update()
        
        def run():
            try:
                conn = sqlite3.connect(DATABASE_FILE)
                try:
                    to_refresh = get_type_ids_with_no_or_zero_volume(
                        conn, MARKET_HISTORY_REGION_ID, scope="prices", limit=2000
                    )
                finally:
                    conn.close()
                if not to_refresh:
                    self.price_update_log.insert(tk.END, "No items need refresh (all have volume data).\n")
                    self.status_var.set("No items to refresh")
                    messagebox.showinfo("Info", "No items with missing/zero volume data found.")
                    return
                self.price_update_log.insert(tk.END, f"Found {len(to_refresh)} items to refresh. Calling API...\n\n")
                self.root.update()
                conn = sqlite3.connect(DATABASE_FILE)
                try:
                    done = 0
                    for i, type_id in enumerate(to_refresh):
                        n = refresh_market_history_for_type(conn, MARKET_HISTORY_REGION_ID, type_id)
                        done += 1
                        if (i + 1) % 50 == 0:
                            self.price_update_log.insert(tk.END, f"  Refreshed {i + 1}/{len(to_refresh)} (last: type_id={type_id}, {n} days)\n")
                            self.root.update()
                    self.price_update_log.insert(tk.END, f"\nRefreshed {done} items.\n")
                    self.status_var.set("Volume refresh complete!")
                    messagebox.showinfo("Success", f"Refreshed market history for {done} items.")
                finally:
                    conn.close()
            except Exception as e:
                self.price_update_log.insert(tk.END, f"\nError: {str(e)}\n")
                self.status_var.set("Error occurred")
                messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        threading.Thread(target=run, daemon=True).start()
    
    def add_on_offer_item(self):
        """Add an item to the on offer list"""
        item_input = self.on_offer_item_var.get().strip()
        
        if not item_input:
            messagebox.showwarning("Warning", "Please enter an item name or TypeID")
            return
        
        if not Path(DATABASE_FILE).exists():
            messagebox.showerror("Error", "Database file not found")
            return
        
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            # Find item by name or typeID
            try:
                module_type_id = int(item_input)
                query = "SELECT typeID, typeName FROM items WHERE typeID = ?"
                params = (module_type_id,)
            except ValueError:
                query = "SELECT typeID, typeName FROM items WHERE typeName = ?"
                params = (item_input,)
            
            cursor = conn.cursor()
            cursor.execute(query, params)
            result = cursor.fetchone()
            
            if not result:
                messagebox.showerror("Error", f"Item not found: {item_input}")
                return
            
            module_type_id, module_name = result
            
            # Check if price data exists
            cursor.execute("SELECT buy_max, sell_min FROM prices WHERE typeID = ?", (module_type_id,))
            price_result = cursor.fetchone()
            
            if not price_result:
                messagebox.showerror("Error", f"No price data found for '{module_name}'. Please update prices first.")
                return
            
            buy_max, sell_min = price_result
            if not buy_max and not sell_min:
                messagebox.showerror("Error", f"No valid price data found for '{module_name}'. Please update prices first.")
                return
            
            # Check if already exists
            cursor.execute("SELECT module_type_id FROM on_offer_items WHERE module_type_id = ?", (module_type_id,))
            if cursor.fetchone():
                messagebox.showwarning("Warning", f"'{module_name}' is already in the on offer list")
                return
            
            # Insert into database; first add counts as first reset (last_reset_date = today, qty sold = 0)
            from datetime import date
            today_str = date.today().isoformat()
            cursor.execute("""
                INSERT INTO on_offer_items (module_type_id, module_name, last_reset_date, quantity_sold_at_last_reset)
                VALUES (?, ?, ?, 0)
            """, (module_type_id, module_name, today_str))
            conn.commit()
            
            messagebox.showinfo("Success", f"Added '{module_name}' to on offer list")
            
            # Clear input field
            self.on_offer_item_var.set("")
            
            # Refresh list
            self.refresh_on_offer_list()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add item: {str(e)}")
        finally:
            conn.close()
    
    def remove_on_offer_item(self):
        """Remove selected item(s) from the on offer list"""
        selected = self.on_offer_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an item to remove")
            return
        
        if not messagebox.askyesno("Confirm", f"Remove {len(selected)} item(s) from on offer list?"):
            return
        
        if not Path(DATABASE_FILE).exists():
            messagebox.showerror("Error", "Database file not found")
            return
        
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            cursor = conn.cursor()
            for item_id in selected:
                values = self.on_offer_tree.item(item_id, 'values')
                module_name = values[0]
                
                # Get module_type_id from database
                cursor.execute("SELECT module_type_id FROM on_offer_items WHERE module_name = ?", (module_name,))
                result = cursor.fetchone()
                if result:
                    cursor.execute("DELETE FROM on_offer_items WHERE module_type_id = ?", (result[0],))
            
            conn.commit()
            messagebox.showinfo("Success", f"Removed {len(selected)} item(s) from on offer list")
            self.refresh_on_offer_list()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to remove item(s): {str(e)}")
        finally:
            conn.close()
    
    def reset_on_offer_date(self):
        """Reset date for selected item: ask quantity sold, then set last_reset_date = today and compute sold per day."""
        selected = self.on_offer_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select one item to reset date.")
            return
        if len(selected) > 1:
            messagebox.showwarning("Warning", "Please select only one item to reset date.")
            return
        item_id = selected[0]
        try:
            module_type_id = int(item_id)
        except ValueError:
            messagebox.showerror("Error", "Could not identify item.")
            return
        values = self.on_offer_tree.item(item_id, "values")
        module_name = values[0] if values else "this item"
        qty = simpledialog.askinteger("Quantity sold", f"Quantity sold for '{module_name}' since last reset?", minvalue=0, initialvalue=0)
        if qty is None:
            return
        if not Path(DATABASE_FILE).exists():
            messagebox.showerror("Error", "Database file not found")
            return
        from datetime import date
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT last_reset_date FROM on_offer_items WHERE module_type_id = ?",
                (module_type_id,)
            )
            row = cursor.fetchone()
            prev_date = row[0] if row and row[0] else None
            today_str = date.today().isoformat()
            cursor.execute("""
                UPDATE on_offer_items
                SET previous_reset_date = last_reset_date,
                    last_reset_date = ?,
                    quantity_sold_at_last_reset = ?
                WHERE module_type_id = ?
            """, (today_str, qty, module_type_id))
            conn.commit()
            messagebox.showinfo("Success", f"Reset date for '{module_name}'. Quantity sold: {qty}. Sold per day will update after next refresh.")
            self.refresh_on_offer_list()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to reset date: {str(e)}")
        finally:
            conn.close()
    
    def refresh_on_offer_list(self):
        """Refresh the on offer list and calculate all values"""
        from datetime import datetime as dt_module, date as date_type
        # Clear existing items
        for item in self.on_offer_tree.get_children():
            self.on_offer_tree.delete(item)
        
        if not Path(DATABASE_FILE).exists():
            messagebox.showinfo("Refresh", "Database not found. Nothing to refresh.")
            return
        
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT module_type_id, module_name, added_at, last_reset_date,
                       quantity_sold_at_last_reset, previous_reset_date
                FROM on_offer_items
                ORDER BY module_name
            """)
            results = cursor.fetchall()
            
            if not results:
                messagebox.showinfo("Refresh", "On Offer list is empty. Nothing to refresh.")
                return
            
            # Get default parameters from assumptions
            from assumptions import (
                DEFAULT_YIELD_PERCENT,
                BUY_ORDER_MARKUP_PERCENT,
                REPROCESSING_COST
            )
            
            yield_percent = DEFAULT_YIELD_PERCENT
            markup_percent = BUY_ORDER_MARKUP_PERCENT
            reprocessing_cost = REPROCESSING_COST
            
            # Calculate values for each item
            for row in results:
                module_type_id, module_name = row[0], row[1]
                added_at = row[2] if len(row) > 2 else None
                last_reset_date = row[3] if len(row) > 3 else None
                quantity_sold_at_last_reset = row[4] if len(row) > 4 else None
                previous_reset_date = row[5] if len(row) > 5 else None
                
                # Date Added: show as dd/mm
                if added_at:
                    try:
                        d = dt_module.strptime(str(added_at)[:10], "%Y-%m-%d")
                        date_added_str = f"{d.day:02d}/{d.month:02d}"
                    except Exception:
                        date_added_str = ""
                else:
                    t = date_type.today()
                    date_added_str = f"{t.day:02d}/{t.month:02d}"
                
                # Sold per day = quantity_sold / (last_reset_date - previous_reset_date) in days
                sold_per_day_str = "N/A"
                if last_reset_date and previous_reset_date and quantity_sold_at_last_reset is not None:
                    try:
                        last = dt_module.strptime(str(last_reset_date)[:10], "%Y-%m-%d")
                        prev = dt_module.strptime(str(previous_reset_date)[:10], "%Y-%m-%d")
                        days = (last - prev).days
                        if days > 0:
                            sold_per_day_str = f"{quantity_sold_at_last_reset / days:,.2f}"
                    except Exception:
                        pass
                try:
                    # Get current market prices from database
                    cursor.execute("SELECT buy_max, sell_min FROM prices WHERE typeID = ?", (module_type_id,))
                    price_result = cursor.fetchone()
                    
                    if not price_result:
                        # No price data - show error
                        self.on_offer_tree.insert('', tk.END, iid=str(module_type_id), values=(
                            module_name,
                            date_added_str,
                            "No price data",
                            "No price data",
                            "Error",
                            "Error",
                            "Error",
                            "Error",
                            sold_per_day_str
                        ))
                        continue
                    
                    buy_max, sell_min = price_result
                    buy_max = float(buy_max) if buy_max else 0.0
                    sell_min = float(sell_min) if sell_min else 0.0
                    
                    # Calculate for buy_offer scenario (module_price_type='buy_offer', mineral_price_type='sell_immediate')
                    result_buy_order = calculate_reprocessing_value(
                        module_type_id=module_type_id,
                        yield_percent=yield_percent,
                        buy_order_markup_percent=markup_percent,
                        reprocessing_cost_percent=reprocessing_cost,
                        module_price_type='buy_offer',
                        mineral_price_type='sell_immediate',
                        db_file=DATABASE_FILE
                    )
                    
                    # Calculate for buy_immediate scenario (module_price_type='buy_immediate', mineral_price_type='sell_immediate')
                    result_immediate = calculate_reprocessing_value(
                        module_type_id=module_type_id,
                        yield_percent=yield_percent,
                        buy_order_markup_percent=markup_percent,
                        reprocessing_cost_percent=reprocessing_cost,
                        module_price_type='buy_immediate',
                        mineral_price_type='sell_immediate',
                        db_file=DATABASE_FILE
                    )
                    
                    breakeven_raw_buy = None
                    if 'error' in result_buy_order or 'error' in result_immediate:
                        # Show error in display
                        profit_buy_order = "Error"
                        profit_immediate = "Error"
                        breakeven_buy_order = "Error"
                        breakeven_immediate = "Error"
                    else:
                        # Get profit per item from buy_order calculation
                        input_quantity = result_buy_order.get('input_quantity', 1)
                        total_mineral_value = result_buy_order.get('total_mineral_value_per_job_after_costs', 0)
                        reprocessing_cost_total = result_buy_order.get('reprocessing_cost_per_job', 0)
                        module_price_after_costs_buy = result_buy_order.get('module_price_after_costs', 0)
                        
                        mineral_value_per_item = total_mineral_value / input_quantity if input_quantity > 0 else 0
                        reprocessing_cost_per_item = reprocessing_cost_total / input_quantity if input_quantity > 0 else 0
                        
                        # Profit per item for buy order (using buy_offer calculation)
                        profit_buy_order = mineral_value_per_item - module_price_after_costs_buy - reprocessing_cost_per_item
                        
                        # Get profit per item from immediate calculation
                        module_price_after_costs_immediate = result_immediate.get('module_price_after_costs', 0)
                        
                        # Profit per item for immediate (using buy_immediate calculation)
                        profit_immediate = mineral_value_per_item - module_price_after_costs_immediate - reprocessing_cost_per_item
                        
                        # Breakeven for buy order (from buy_offer calculation)
                        breakeven_raw_buy = result_buy_order.get('breakeven_module_price', 'na')
                        if isinstance(breakeven_raw_buy, (int, float)) and breakeven_raw_buy not in (0, float('inf')):
                            breakeven_buy_order = f"{breakeven_raw_buy:,.2f}"
                        else:
                            breakeven_buy_order = "N/A"
                            breakeven_raw_buy = None
                        
                        # Breakeven for immediate (from buy_immediate calculation)
                        breakeven_immediate = result_immediate.get('breakeven_module_price', 'na')
                        if isinstance(breakeven_immediate, (int, float)) and breakeven_immediate not in (0, float('inf')):
                            breakeven_immediate = f"{breakeven_immediate:,.2f}"
                        else:
                            breakeven_immediate = "N/A"
                    
                    # Deep red if buy_max > breakeven max (buy order); light red if buy_max > 90% of breakeven
                    row_tags = ()
                    if breakeven_raw_buy is not None and breakeven_raw_buy > 0 and buy_max > 0:
                        if buy_max > breakeven_raw_buy:
                            row_tags = ('sell_above_breakeven',)  # deep red: buy price above breakeven
                        elif buy_max > 0.9 * breakeven_raw_buy:
                            row_tags = ('high_buy_near_breakeven',)  # light red: buy within 90% of breakeven
                    
                    # Insert into treeview (iid = module_type_id for reset)
                    self.on_offer_tree.insert('', tk.END, iid=str(module_type_id), values=(
                        module_name,
                        date_added_str,
                        f"{buy_max:,.2f}" if buy_max > 0 else "N/A",
                        f"{sell_min:,.2f}" if sell_min > 0 else "N/A",
                        f"{profit_buy_order:,.2f}" if isinstance(profit_buy_order, (int, float)) else profit_buy_order,
                        f"{profit_immediate:,.2f}" if isinstance(profit_immediate, (int, float)) else profit_immediate,
                        breakeven_buy_order,
                        breakeven_immediate,
                        sold_per_day_str
                    ), tags=row_tags)
                
                except Exception as e:
                    # Insert with error message
                    self.on_offer_tree.insert('', tk.END, iid=str(module_type_id), values=(
                        module_name,
                        date_added_str,
                        "Error",
                        "Error",
                        "Error",
                        "Error",
                        "Error",
                        "Error",
                        sold_per_day_str
                    ))
            
            messagebox.showinfo("Refresh", f"Refresh complete. Calculations updated for {len(results)} item(s).")
                    
        finally:
            conn.close()


def main():
    root = tk.Tk()
    app = EVELauncher(root)
    root.mainloop()


if __name__ == "__main__":
    main()

