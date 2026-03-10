"""
EVE Manufacturing Database Launcher
A GUI interface for managing and analyzing EVE Online manufacturing and reprocessing data.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog
import threading
import sys
import math
import sqlite3
from pathlib import Path

# Import our modules
from calculate_reprocessing_value import (
    calculate_reprocessing_value,
    analyze_all_modules,
    format_reprocessing_result
)
from calculate_blueprint_profitability import calculate_blueprint_profitability
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
        
        # Create tabs
        self.create_analysis_tab()
        self.create_single_module_tab()
        self.create_single_blueprint_tab()
        self.create_price_update_tab()
        self.create_exclusions_tab()
        self.create_on_offer_tab()
        self.create_paste_compare_tab()
        self.create_sso_sync_tab()
        
        # So analysis tab fields are editable immediately (focus first entry when that tab is shown)
        self.root.after(150, self._focus_analysis_first_entry_if_visible)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        
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
                else:
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
                    append(f"System cost ({result['system_cost_percent']}%): {result['system_cost']:,.2f} ISK\n")
                    append(f"Manufacturing tax ({result.get('manufacturing_tax_rate', 0):.1f}%): {result.get('manufacturing_tax', 0):,.2f} ISK\n")
                    if result.get("tax_details"):
                        for t in result["tax_details"]:
                            append(f"  Tax {t['materialName']}: avg {t['average']:,.2f} × {t['quantity']:,} × {result.get('manufacturing_tax_rate', 0):.1f}% = {t['tax']:,.2f} ISK (per run)\n")
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
        threading.Thread(target=run, daemon=True).start()
    
    def fetch_blueprint_data(self):
        """Run build_database to fetch SDE and populate blueprints (full DB rebuild)."""
        if not messagebox.askyesno("Fetch blueprint data", "This will run build_database.py: download SDE and rebuild blueprints (and other tables). Continue?"):
            return
        self.status_var.set("Fetching blueprint data (build_database)...")
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
        
        # Tag for rows where buy price >= 85% of breakeven max (light red background)
        self.on_offer_tree.tag_configure('high_buy_near_breakeven', background='#ffcccc')
        
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
                    
                    # Light red row if buy price >= 85% of breakeven max (Buy Order)
                    row_tags = ()
                    if buy_max > 0 and breakeven_raw_buy is not None and breakeven_raw_buy > 0:
                        if buy_max >= 0.85 * breakeven_raw_buy:
                            row_tags = ('high_buy_near_breakeven',)
                    
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

