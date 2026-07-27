"""
EVE Manufacturing Database Launcher
A GUI interface for managing and analyzing EVE Online manufacturing and reprocessing data.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog, filedialog
import threading
import time
import sys
import math
import json
import logging
import sqlite3
import subprocess
from pathlib import Path
from typing import Optional

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
    get_market_average_price_7d_avg,
    refresh_market_history_for_type,
    clear_market_history_session_cache,
    discard_market_history_session_refresh,
    get_type_ids_with_no_or_zero_volume,
    run_fetch,
)

DATABASE_FILE = "eve_manufacturing.db"
# Small, git-tracked snapshot (no market history, no SSO tokens). On a fresh clone
# the runtime DB is bootstrapped from this so the app works immediately.
CORE_DATABASE_FILE = "eve_manufacturing_core.db"
# Region ID for market_history_daily (The Forge); must match data fetched by fetch_market_history.py
MARKET_HISTORY_REGION_ID = 10000002
# Preferences file for decryptor comparison and other persisted settings
LAUNCHER_PREFS_FILE = Path(__file__).resolve().parent / "eve_launcher_prefs.json"
# Persisted shopping list (survives restarts until reset)
SHOPPING_LIST_FILE = Path(__file__).resolve().parent / "eve_launcher_shopping_list.json"
PRODUCTION_TRACKING_FILE = Path(__file__).resolve().parent / "eve_launcher_production_tracking.json"
PUT_IN_PRODUCTION_ROW_SEP = "\x1f"
# Persisted skill levels (My Skills tab)
SKILLS_FILE = Path(__file__).resolve().parent / "eve_launcher_skills.json"
# Persisted EVE SSO credentials (Client ID / Client Secret). Gitignored.
SSO_CREDENTIALS_FILE = Path(__file__).resolve().parent / "eve_sso_credentials.json"

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
        self.create_put_in_production_tab()
        self.create_analysis_tab()
        self.create_single_module_tab()
        self.create_single_blueprint_tab()
        self.create_skills_blueprints_tab()
        self.create_paste_compare_tab()
        self.create_planning_tab()
        self.create_market_patterns_tab()
        self.create_sso_sync_tab()
        self.create_profitability_tab()
        self.create_arbitrage_tab()
        self.create_remap_planner_tab()
        
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
        self.root.after(50, self._startup_heavy_refresh)

    def _startup_heavy_refresh(self):
        """Defer On Offer + shopping list refresh until after first paint."""
        try:
            self.status_var.set("Loading On Offer and shopping list…")
        except Exception:
            pass
        try:
            self.refresh_on_offer_list(quiet=True)
        except Exception:
            pass
        try:
            self._shopping_list_refresh_tree()
            self._refresh_shopping_list_aggregate()
        except Exception:
            pass
        try:
            self.status_var.set("Ready")
        except Exception:
            pass

    def _on_launcher_close(self):
        """Save shopping list when closing the app (belt-and-suspenders; list also saves on each edit)."""
        try:
            self._save_shopping_list()
        except Exception:
            pass
        try:
            self._save_production_tracking()
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

        prod_cost_row = ttk.Frame(input_frame)
        prod_cost_row.pack(fill=tk.X, pady=5)
        ttk.Label(prod_cost_row, text="Production cost per run (ISK):").pack(side=tk.LEFT, padx=5)
        self.blueprint_prod_cost_var = tk.StringVar(value="")
        ttk.Entry(prod_cost_row, textvariable=self.blueprint_prod_cost_var, width=18).pack(side=tk.LEFT, padx=5)
        ttk.Label(prod_cost_row, text="(blank = use EIV-based system cost)").pack(side=tk.LEFT, padx=5)
        ttk.Button(prod_cost_row, text="Save to DB", command=self._save_single_blueprint_prod_cost).pack(side=tk.LEFT, padx=10)

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

                    # Check if user has entered a production cost for this blueprint
                    user_prod_cost_per_run = None
                    try:
                        db_conn = sqlite3.connect(DATABASE_FILE)
                        try:
                            bp_row = resolve_blueprint(db_conn, name)
                            if bp_row:
                                self._ensure_blueprint_datacore_bindings_table(db_conn)
                                cr = db_conn.execute(
                                    "SELECT production_cost_per_run FROM blueprint_datacore_bindings WHERE blueprint_type_id = ?",
                                    (bp_row["blueprintTypeID"],),
                                ).fetchone()
                                if cr and cr[0] is not None:
                                    v = float(cr[0])
                                    if v > 0:
                                        user_prod_cost_per_run = v
                                # Reflect stored value in the UI field
                                stored_display = str(int(user_prod_cost_per_run)) if user_prod_cost_per_run is not None else ""
                                self.blueprint_prod_cost_var.set(stored_display)
                        finally:
                            db_conn.close()
                    except Exception:
                        pass

                    out_revenue = result['output_revenue']
                    total_input = result['total_input_cost']
                    items_produced = result['items_produced']
                    if user_prod_cost_per_run is not None:
                        user_total_prod_cost = user_prod_cost_per_run * runs
                        adj_profit = out_revenue - total_input - user_total_prod_cost
                        adj_return = adj_profit / total_input * 100 if total_input > 0 else 0.0
                        adj_cost_per_item = (total_input + user_total_prod_cost) / items_produced if items_produced > 0 else 0.0
                        adj_profit_per_item = adj_profit / items_produced if items_produced > 0 else 0.0
                        append(f"System cost (8.61% EIV): {result['system_cost']:,.2f} ISK  ← NOT used (user input overrides)\n")
                        append(f"⚑ User input prod cost:  {user_prod_cost_per_run:,.2f} ISK/run × {runs} run(s) = {user_total_prod_cost:,.2f} ISK\n", "profit_positive")
                        append(f"Output revenue:       {out_revenue:,.2f} ISK  ({result['output_total_quantity']:,} × {result['output_unit_price']:,.2f})\n")
                        adj_profit_tag = "profit_positive" if adj_profit >= 0 else "profit_negative"
                        append(f"Profit (user cost):   {adj_profit:,.2f} ISK\n", adj_profit_tag)
                        append(f"Return (user cost):   {adj_return:,.2f}%\n\n")
                        append("——— Per item (using user input cost) ———\n")
                        append(f"Items produced:       {items_produced:,}\n")
                        append(f"Cost per item:        {adj_cost_per_item:,.2f} ISK\n")
                        append(f"Revenue per item:     {result['revenue_per_item']:,.2f} ISK\n")
                        adj_ppi_tag = "profit_positive" if adj_profit_per_item >= 0 else "profit_negative"
                        append(f"Profit per item:      {adj_profit_per_item:,.2f} ISK\n", adj_ppi_tag)
                        effective_profit = adj_profit
                    else:
                        append(f"System cost ({result['system_cost_percent']}% of EIV): {result['system_cost']:,.2f} ISK\n")
                        append(f"Output revenue:       {out_revenue:,.2f} ISK  ({result['output_total_quantity']:,} × {result['output_unit_price']:,.2f})\n")
                        profit_tag = "profit_positive" if result['profit'] >= 0 else "profit_negative"
                        append(f"Profit:               {result['profit']:,.2f} ISK\n", profit_tag)
                        append(f"Return:               {result['return_percent']:,.2f}%\n\n")
                        append("——— Per item ———\n")
                        append(f"Items produced:       {items_produced:,}\n")
                        append(f"Cost per item:        {result['cost_per_item']:,.2f} ISK\n")
                        append(f"Revenue per item:     {result['revenue_per_item']:,.2f} ISK\n")
                        profit_per_item_tag = "profit_positive" if result['profit_per_item'] >= 0 else "profit_negative"
                        append(f"Profit per item:      {result['profit_per_item']:,.2f} ISK\n", profit_per_item_tag)
                        effective_profit = result['profit']
                    # Color results area: green if profit >= 0, red if loss
                    if effective_profit >= 0:
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
    
    def _save_single_blueprint_prod_cost(self):
        """Save (or clear) the production cost per run for the current blueprint to the DB."""
        name = self.blueprint_name_var.get().strip()
        if not name:
            messagebox.showwarning("Save production cost", "Enter a blueprint or product name first.")
            return
        if not Path(DATABASE_FILE).exists():
            messagebox.showerror("Save production cost", "Database not found.")
            return
        raw = self.blueprint_prod_cost_var.get().strip()
        prod_cost = None
        if raw:
            try:
                prod_cost = float(raw)
            except ValueError:
                messagebox.showerror("Save production cost", f"Invalid cost value: {raw!r}")
                return
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                self._ensure_blueprint_datacore_bindings_table(conn)
                bp = resolve_blueprint(conn, name)
                if not bp:
                    messagebox.showerror("Save production cost", f"Blueprint not found for: {name!r}")
                    return
                bp_id = bp["blueprintTypeID"]
                existing = conn.execute(
                    "SELECT blueprint_type_id FROM blueprint_datacore_bindings WHERE blueprint_type_id = ?",
                    (bp_id,),
                ).fetchone()
                if existing:
                    conn.execute(
                        "UPDATE blueprint_datacore_bindings SET production_cost_per_run=?, updated_at=CURRENT_TIMESTAMP WHERE blueprint_type_id=?",
                        (prod_cost, bp_id),
                    )
                else:
                    conn.execute(
                        """INSERT INTO blueprint_datacore_bindings
                           (blueprint_type_id, dc1_name, dc1_qty, dc2_name, dc2_qty, production_cost_per_run, updated_at)
                           VALUES (?, NULL, 0, NULL, 0, ?, CURRENT_TIMESTAMP)""",
                        (bp_id, prod_cost),
                    )
                conn.commit()
                label = f"{prod_cost:,.2f} ISK/run" if prod_cost is not None else "cleared"
                self.status_var.set(f"Production cost {label} saved for {name}.")
            finally:
                conn.close()
        except Exception as e:
            messagebox.showerror("Save production cost", str(e))

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
        ttk.Button(row1, text="Load", command=self._decryptor_load_from_t2_name).pack(side=tk.LEFT, padx=5)
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
        row_research_time = ttk.Frame(input_frame)
        row_research_time.pack(fill=tk.X, pady=3)
        ttk.Label(row_research_time, text="Time for 1 research run:").pack(side=tk.LEFT, padx=5)
        ttk.Label(row_research_time, text="Days:").pack(side=tk.LEFT, padx=(8, 2))
        self.decryptor_research_days_var = tk.StringVar(value="0")
        ttk.Entry(row_research_time, textvariable=self.decryptor_research_days_var, width=5).pack(side=tk.LEFT, padx=2)
        ttk.Label(row_research_time, text="Hours:").pack(side=tk.LEFT, padx=(6, 2))
        self.decryptor_research_hours_var = tk.StringVar(value="0")
        ttk.Entry(row_research_time, textvariable=self.decryptor_research_hours_var, width=5).pack(side=tk.LEFT, padx=2)
        ttk.Label(row_research_time, text="Min:").pack(side=tk.LEFT, padx=(6, 2))
        self.decryptor_research_minutes_var = tk.StringVar(value="0")
        ttk.Entry(row_research_time, textvariable=self.decryptor_research_minutes_var, width=5).pack(side=tk.LEFT, padx=2)
        ttk.Label(row_research_time, text="    Production cost per run (ISK):").pack(side=tk.LEFT, padx=(12, 2))
        self.decryptor_prod_cost_var = tk.StringVar(value="")
        ttk.Entry(row_research_time, textvariable=self.decryptor_prod_cost_var, width=14).pack(side=tk.LEFT, padx=2)
        ttk.Label(row_research_time, text="(blank = N/A)").pack(side=tk.LEFT, padx=4)
        row_assoc = ttk.Frame(input_frame)
        row_assoc.pack(fill=tk.X, pady=3)
        ttk.Label(row_assoc, text="Associate T1 ↔ T2 (save to DB):").pack(side=tk.LEFT, padx=5)
        ttk.Button(row_assoc, text="Associate T1 → T2", command=self._associate_t1_t2).pack(side=tk.LEFT, padx=5)
        ttk.Label(row_assoc, text="Uses T1 from field above and T2 from top field. Also saves research time and production cost. Next time enter only T1 and look up T2.").pack(side=tk.LEFT, padx=5)
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

        # ── Outer scrollable canvas so the whole tab scrolls when the window is short ──
        sl_canvas = tk.Canvas(frame, highlightthickness=0)
        sl_vscroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=sl_canvas.yview)
        sl_canvas.configure(yscrollcommand=sl_vscroll.set)
        sl_vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        sl_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sl_inner = ttk.Frame(sl_canvas)
        _sl_win = sl_canvas.create_window((0, 0), window=sl_inner, anchor=tk.NW)

        def _sl_inner_resized(event):
            sl_canvas.configure(scrollregion=sl_canvas.bbox("all"))
        sl_inner.bind("<Configure>", _sl_inner_resized)

        def _sl_canvas_resized(event):
            sl_canvas.itemconfig(_sl_win, width=event.width)
        sl_canvas.bind("<Configure>", _sl_canvas_resized)

        def _sl_mousewheel(event):
            sl_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        sl_canvas.bind("<MouseWheel>", _sl_mousewheel)
        sl_inner.bind("<MouseWheel>", _sl_mousewheel)
        # ─────────────────────────────────────────────────────────────────────────────

        top = ttk.LabelFrame(sl_inner, text="Blueprints in list", padding=10)
        top.pack(fill=tk.X, padx=10, pady=10)
        cols = ("Blueprint / Product", "Research", "Runs", "# prod", "Decryptor", "Run per BPC", "Total material cost", "Sell immediate", "Hist 7d avg", "Sell offer", "Breakeven", "E[research]", "E[prod]", "E[prod -min]")
        self.shopping_list_columns = cols
        self.shopping_list_sort_column = None
        self.shopping_list_sort_reverse = False
        # Wrap tree + scrollbar in their own frame so all controls pack below them
        tree_frame = ttk.Frame(top)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self.shopping_list_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=10, selectmode="browse")
        col_widths = {
            "Blueprint / Product": 180, "Research": 65, "Runs": 55, "# prod": 60,
            "Decryptor": 120, "Run per BPC": 85, "Total material cost": 120,
            "Sell immediate": 100, "Hist 7d avg": 88, "Sell offer": 100, "Breakeven": 100,
            "E[research]": 110, "E[prod]": 110, "E[prod -min]": 110,
        }
        non_stretch = {"Research", "Runs", "# prod", "Hist 7d avg"}
        for c in cols:
            self.shopping_list_tree.heading(c, text=c, command=lambda col=c: self._shopping_list_sort_by(col))
            self.shopping_list_tree.column(c, width=col_widths.get(c, 100), stretch=(c not in non_stretch))
        scroll_tree = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.shopping_list_tree.yview)
        self.shopping_list_tree.configure(yscrollcommand=scroll_tree.set)
        self.shopping_list_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_tree.pack(side=tk.RIGHT, fill=tk.Y)

        # ── Totals bar (always visible below the tree, outside the scroll area) ──────
        totals_bar = tk.Frame(top, bg="#e8e8e8", relief="groove", bd=1)
        totals_bar.pack(fill=tk.X, pady=(3, 0))
        tk.Label(totals_bar, text="Totals:", font=("TkDefaultFont", 8, "bold"),
                 bg="#e8e8e8").pack(side=tk.LEFT, padx=(8, 10))
        self._sl_total_vars = {}
        for _lbl, _key in [
            ("Research",    "_tot_research"),
            ("# prod",      "_tot_prod"),
            ("Mat. cost",   "_tot_mat_cost"),
            ("E[research]", "_tot_e_research"),
            ("E[prod]",     "_tot_e_prod"),
        ]:
            tk.Label(totals_bar, text=f"{_lbl}:", bg="#e8e8e8",
                     font=("TkDefaultFont", 8)).pack(side=tk.LEFT, padx=(4, 1))
            _var = tk.StringVar(value="—")
            self._sl_total_vars[_key] = _var
            tk.Label(totals_bar, textvariable=_var, width=13, anchor=tk.E,
                     bg="#e8e8e8", font=("TkDefaultFont", 8)).pack(side=tk.LEFT, padx=(0, 10))
        # ────────────────────────────────────────────────────────────────────────────

        self.shopping_list_tree.tag_configure("manual_rpb", background="#cce5ff")
        self.shopping_list_tree.tag_configure("mat_override", background="#fff2cc")
        self.shopping_list_tree.bind("<<TreeviewSelect>>", self._on_shopping_list_selection)
        self.shopping_list_tree.bind("<Double-1>", self._shopping_list_on_double_click)
        self.shopping_list_tree.bind("<ButtonRelease-1>", self._sl_tree_click)
        self.shopping_list_tree.bind("<Motion>", self._sl_tree_motion)
        self.shopping_list_tree.bind("<Leave>", self._sl_tree_leave)
        self._sl_tooltip_win = None
        self._sl_tooltip_after_id = None
        self._sl_tooltip_last_cell = (None, None)
        self._sl_rpb_edit_widget = None
        # Row 1: quantity fields + primary edit buttons
        btn_row1 = ttk.Frame(top)
        btn_row1.pack(fill=tk.X, pady=(5, 2))
        ttk.Label(btn_row1, text="Research:").pack(side=tk.LEFT, padx=(5, 2))
        self.shopping_list_research_var = tk.StringVar(value="0")
        ttk.Entry(btn_row1, textvariable=self.shopping_list_research_var, width=6).pack(side=tk.LEFT, padx=2)
        ttk.Label(btn_row1, text="Runs:").pack(side=tk.LEFT, padx=(6, 2))
        self.shopping_list_runs_var = tk.StringVar(value="0")
        ttk.Entry(btn_row1, textvariable=self.shopping_list_runs_var, width=6).pack(side=tk.LEFT, padx=2)
        ttk.Label(btn_row1, text="Prod:").pack(side=tk.LEFT, padx=(6, 2))
        self.shopping_list_prod_var = tk.StringVar(value="0")
        ttk.Entry(btn_row1, textvariable=self.shopping_list_prod_var, width=6).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row1, text="Update quantities", command=self._shopping_list_update_quantity).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row1, text="Set prod to 0", command=self._shopping_list_set_prod_zero).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row1, text="Reset all qty to 0", command=self._shopping_list_reset_all_quantities_zero).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row1, text="Revert RPB", command=self._sl_revert_rpb).pack(side=tk.LEFT, padx=5)
        # Row 1b: max research time + computed max runs for selected row
        max_time_row = ttk.Frame(top)
        max_time_row.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(max_time_row, text="Max research time:").pack(side=tk.LEFT, padx=(5, 2))
        self.sl_max_research_days_var = tk.StringVar(value="6")
        ttk.Entry(max_time_row, textvariable=self.sl_max_research_days_var, width=4).pack(side=tk.LEFT, padx=2)
        ttk.Label(max_time_row, text="d").pack(side=tk.LEFT)
        self.sl_max_research_hours_var = tk.StringVar(value="12")
        ttk.Entry(max_time_row, textvariable=self.sl_max_research_hours_var, width=4).pack(side=tk.LEFT, padx=2)
        ttk.Label(max_time_row, text="h").pack(side=tk.LEFT)
        # Update max-runs label and save prefs whenever the max-time values change
        def _on_max_time_change(*_):
            self._sl_update_max_runs_label()
            self._save_decryptor_prefs()
        self.sl_max_research_days_var.trace_add("write", _on_max_time_change)
        self.sl_max_research_hours_var.trace_add("write", _on_max_time_change)
        ttk.Label(max_time_row, text="   →  Max runs for selected:").pack(side=tk.LEFT, padx=(12, 2))
        self._sl_max_runs_var = tk.StringVar(value="—")
        ttk.Label(max_time_row, textvariable=self._sl_max_runs_var, width=8, anchor=tk.W,
                  foreground="#1a6eb5", font=("TkDefaultFont", 9, "bold")).pack(side=tk.LEFT, padx=2)
        # Row 2: list-level action buttons
        btn_row2_top = ttk.Frame(top)
        btn_row2_top.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(btn_row2_top, text="Remove selected", command=self._shopping_list_remove_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row2_top, text="Copy plan to clipboard", command=self._shopping_list_copy_plan_to_clipboard).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row2_top, text="Refresh profitability", command=self._shopping_list_refresh_profitability).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            btn_row2_top,
            text="Refresh market history (selected)",
            command=self._shopping_list_refresh_market_history_selected,
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            btn_row2_top,
            text="Refresh market history (all in list)",
            command=self._shopping_list_refresh_market_history,
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row2_top, text="Put in production…", command=self._shopping_list_open_put_in_production).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row2_top, text="Edit materials (selected)…", command=self._shopping_list_edit_materials).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row2_top, text="Export data…", command=self._export_state_to_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row2_top, text="Import data…", command=self._import_state_from_file).pack(side=tk.LEFT, padx=5)
        ttk.Label(
            top,
            text="Research × Runs = invention attempts (drives datacores/decryptors). "
                 "# prod = number of BPCs to manufacture (hover for info). "
                 "Run per BPC = manufacturing runs per BPC (from decryptor; click cell to override, shown in blue). "
                 "Materials = # prod × Run per BPC × qty/run. "
                 "E[research] = profit/BPC (hover). E[prod] = total mfg profit (hover). "
                 "Double-click a row to copy item name. Click column header to sort. "
                 "Select a row and click 'Edit materials' to bind a custom per-run material list "
                 "(for null-sec/structure bonuses); such rows are highlighted and their bound list is used in the aggregate.",
            wraplength=720,
            justify=tk.LEFT,
        ).pack(fill=tk.X, anchor=tk.W, pady=(0, 4))
        agg_frame = ttk.LabelFrame(sl_inner, text="Items required for manufacturing (aggregated)", padding=10)
        agg_frame.pack(fill=tk.X, padx=10, pady=5)
        self.shopping_list_aggregate_text = scrolledtext.ScrolledText(agg_frame, wrap=tk.WORD, height=14, state=tk.DISABLED)
        self.shopping_list_aggregate_text.pack(fill=tk.BOTH, expand=True)
        btn_row2 = ttk.Frame(agg_frame)
        btn_row2.pack(fill=tk.X, pady=5)
        ttk.Button(btn_row2, text="Copy to clipboard", command=self._shopping_list_copy_to_clipboard).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row2, text="Refresh list", command=self._refresh_shopping_list_aggregate).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row2, text="Reset list (clear and save)", command=self._shopping_list_reset).pack(side=tk.LEFT, padx=5)
        # Inventory paste: compare with aggregated list to show shortfall
        inv_frame = ttk.LabelFrame(sl_inner, text="Your inventory (paste item names and quantities; one per line, e.g. 'Tritanium 5000' or 'Tritanium\t5000')", padding=8)
        inv_frame.pack(fill=tk.X, padx=10, pady=5)
        self.shopping_list_inventory_text = scrolledtext.ScrolledText(inv_frame, wrap=tk.WORD, height=6, state=tk.NORMAL)
        self.shopping_list_inventory_text.pack(fill=tk.BOTH, expand=True)
        inv_btn_row = ttk.Frame(inv_frame)
        inv_btn_row.pack(fill=tk.X, pady=4)
        ttk.Button(inv_btn_row, text="Compare: shortfall + unused / excess", command=self._shopping_list_compare_inventory).pack(side=tk.LEFT, padx=5)
        compare_row = ttk.Frame(sl_inner)
        compare_row.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        inv_compare_paned = ttk.PanedWindow(compare_row, orient=tk.HORIZONTAL)
        inv_compare_paned.pack(fill=tk.BOTH, expand=True)
        shortfall_frame = ttk.LabelFrame(inv_compare_paned, text="Still need to get (required minus in inventory)", padding=8)
        excess_frame = ttk.LabelFrame(inv_compare_paned, text="Unused & excess (vs plan)", padding=8)
        inv_compare_paned.add(shortfall_frame, weight=1)
        inv_compare_paned.add(excess_frame, weight=1)
        self.shopping_list_shortfall_text = scrolledtext.ScrolledText(shortfall_frame, wrap=tk.WORD, height=8, state=tk.DISABLED)
        self.shopping_list_shortfall_text.pack(fill=tk.BOTH, expand=True)
        shortfall_btn_row = ttk.Frame(shortfall_frame)
        shortfall_btn_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(shortfall_btn_row, text="Copy to clipboard", command=self._shopping_list_copy_shortfall).pack(side=tk.LEFT, padx=5)
        excess_help = (
            "Unused: pasted items that do not match anything in the aggregated plan.\n"
            "Excess: for plan items, quantity held beyond 3× what the plan requires (have − 3×need).\n"
            "Run Compare after pasting inventory.\n\n"
        )
        self.shopping_list_excess_unused_text = scrolledtext.ScrolledText(excess_frame, wrap=tk.WORD, height=8, state=tk.DISABLED)
        self.shopping_list_excess_unused_text.pack(fill=tk.BOTH, expand=True)
        self.shopping_list_excess_unused_text.configure(state=tk.NORMAL)
        self.shopping_list_excess_unused_text.insert(tk.END, excess_help)
        self.shopping_list_excess_unused_text.configure(state=tk.DISABLED)
        # Propagate mousewheel from all non-scrolling child frames to the outer canvas
        for _w in (top, agg_frame, inv_frame, compare_row):
            _w.bind("<MouseWheel>", _sl_mousewheel)

        # Defer loading until after the window is fully rendered so startup is instant
        self.root.after(0, self._load_shopping_list)

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
        # Controls: refresh from DB + load actual trained levels from an SSO character
        skills_ctrl = ttk.Frame(skills_frame)
        skills_ctrl.pack(fill=tk.X, pady=4)
        ttk.Button(skills_ctrl, text="Refresh skills from DB", command=self._skills_blueprints_fill_skills).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(skills_ctrl, text="SSO character:").pack(side=tk.LEFT, padx=(0, 4))
        self.skills_sso_char_var = tk.StringVar()
        self.skills_sso_char_combo = ttk.Combobox(skills_ctrl, textvariable=self.skills_sso_char_var, state="readonly", width=28, values=[])
        self.skills_sso_char_combo.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(skills_ctrl, text="Refresh chars", command=self._skills_refresh_sso_chars).pack(side=tk.LEFT, padx=(0, 6))
        self.skills_load_sso_btn = ttk.Button(skills_ctrl, text="Load levels from SSO", command=self._skills_load_from_sso)
        self.skills_load_sso_btn.pack(side=tk.LEFT)
        self._skills_sso_char_map = {}  # label -> character_id
        self._skills_refresh_sso_chars()
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

    def _skills_refresh_sso_chars(self):
        """Populate the SSO character dropdown from the sso_character table."""
        labels = []
        self._skills_sso_char_map = {}
        try:
            from eve_sso_sync import list_sso_characters, ensure_sso_tables
            if Path(DATABASE_FILE).exists():
                conn = sqlite3.connect(DATABASE_FILE, timeout=30)
                try:
                    ensure_sso_tables(conn)
                    for r in list_sso_characters(conn):
                        name = r.get("character_name") or str(r.get("character_id"))
                        labels.append(name)
                        self._skills_sso_char_map[name] = r["character_id"]
                finally:
                    conn.close()
        except Exception:
            pass
        self.skills_sso_char_combo.config(values=labels)
        if labels and not self.skills_sso_char_var.get():
            self.skills_sso_char_var.set(labels[0])

    def _skills_load_from_sso(self):
        """Fetch the selected character's trained skill levels via ESI and set the spinboxes."""
        label = self.skills_sso_char_var.get().strip()
        char_id = self._skills_sso_char_map.get(label)
        if not char_id:
            messagebox.showwarning("Load from SSO", "Select a linked SSO character first (add one in the EVE SSO Sync tab).")
            return
        cid, secret = self._load_sso_credentials()
        if not cid or not secret:
            messagebox.showwarning("Load from SSO", "Set EVE SSO Client ID / Secret first (EVE SSO Sync tab).")
            return
        self.skills_load_sso_btn.config(state=tk.DISABLED)
        self.skills_analysis_status_var.set(f"Loading skills for {label} from SSO...")
        self.root.update_idletasks()

        def worker():
            try:
                from eve_sso_sync import get_valid_access_token, fetch_character_skills
                conn = sqlite3.connect(DATABASE_FILE, timeout=60)
                conn.execute("PRAGMA busy_timeout=60000")
                try:
                    token = get_valid_access_token(conn, char_id, cid, secret)
                finally:
                    conn.close()
                if not token:
                    raise RuntimeError("No valid token; re-link this character in EVE SSO Sync.")
                data = fetch_character_skills(char_id, token)
                levels = {}
                for s in data.get("skills", []):
                    sid = s.get("skill_id")
                    lvl = s.get("active_skill_level", s.get("trained_skill_level", 0))
                    if sid is not None:
                        levels[int(sid)] = max(0, min(5, int(lvl or 0)))
                self.root.after(0, lambda: self._skills_apply_sso_levels(label, levels))
            except Exception as e:
                self.root.after(0, lambda msg=str(e): self._skills_sso_error(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _skills_apply_sso_levels(self, label, levels):
        """Set spinbox levels from a {skill_id: level} map; skills not trained -> 0."""
        self.skills_load_sso_btn.config(state=tk.NORMAL)
        matched = 0
        for sid, var in self.skills_level_vars.items():
            lvl = levels.get(int(sid), 0)
            var.set(lvl)
            if lvl > 0:
                matched += 1
        self._save_skills_prefs()
        self.skills_analysis_status_var.set(
            f"Loaded {matched} trained skills from {label} ({len(levels)} known to character). Click Run analysis."
        )

    def _skills_sso_error(self, msg):
        self.skills_load_sso_btn.config(state=tk.NORMAL)
        self.skills_analysis_status_var.set(f"SSO load failed: {msg}")
        messagebox.showerror("Load from SSO", msg)

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
        """Add current blueprint/product from Single Blueprint tab: prod=1, runs_per_bpc = chosen runs (manual/blue)."""
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
        entry = {
            "product_name": name,
            "quantity": 1,
            "profit": profit,
            "runs_per_bpc": runs,
            "default_runs_per_bpc": 1,   # decryptor-default; revert goes back to 1
            "manual_runs_per_bpc": True,  # user chose this run count → blue background
            "research": 0,
            "runs_per_research": 0,
            "prod": 1,
        }
        self.shopping_list.append(entry)
        self._shopping_list_refresh_tree()
        self._refresh_shopping_list_aggregate()
        self._save_shopping_list()
        for i in range(self.notebook.index("end")):
            if self.notebook.tab(i, "text") == "Shopping list":
                self.notebook.select(i)
                break
        self.status_var.set(f"Added {name} (prod=1, {runs} run(s)/BPC) to shopping list.")

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
                    for field in ("research", "runs_per_research", "prod"):
                        v = entry.get(field)
                        if v is not None:
                            try:
                                rec[field] = max(0, int(v))
                            except (TypeError, ValueError):
                                rec[field] = 0
                    if entry.get("manual_runs_per_bpc"):
                        rec["manual_runs_per_bpc"] = True
                    if entry.get("default_runs_per_bpc") is not None:
                        try:
                            rec["default_runs_per_bpc"] = max(1, int(entry["default_runs_per_bpc"]))
                        except (TypeError, ValueError):
                            pass
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
            # The "Put in production" tab is built before this (deferred) load runs,
            # so rebuild its rows now that the shopping list is populated.
            if hasattr(self, "put_production_inner"):
                self._put_in_production_rebuild_rows()
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

    def _shopping_list_format_price_display(self, v):
        """Format unit ISK for shopping list: no decimals when value > 1000."""
        if v is None:
            return "—"
        try:
            x = float(v)
        except (TypeError, ValueError):
            return "—"
        if x <= 0:
            return "—"
        if x > 1000:
            return f"{x:,.0f}"
        return f"{x:,.2f}"

    def _shopping_list_hist_7d_avg_price(self, conn, product_name):
        """Mean of last-7-days daily `average` from market_history_daily (same region as volume)."""
        bp = resolve_blueprint(conn, product_name)
        if not bp:
            return None
        avg, _ = get_market_average_price_7d_avg(conn, MARKET_HISTORY_REGION_ID, int(bp["productTypeID"]))
        return avg

    def _shopping_list_refresh_market_history(self):
        """Fetch EVE Tycoon market history for each distinct product on the list, then refresh the tree."""
        if not self.shopping_list:
            messagebox.showinfo("Shopping list", "The list is empty.")
            return
        if not Path(DATABASE_FILE).exists():
            messagebox.showwarning("Shopping list", "Database not found.")
            return
        self.status_var.set("Refreshing market history for shopping list (API, may take a while)…")

        def work():
            err_msg = None
            try:
                clear_market_history_session_cache()
                conn = sqlite3.connect(DATABASE_FILE)
                try:
                    seen_tid = set()
                    for entry in self.shopping_list:
                        name = (entry.get("product_name") or "").strip()
                        if not name:
                            continue
                        bp = resolve_blueprint(conn, name)
                        if not bp:
                            continue
                        tid = int(bp["productTypeID"])
                        if tid in seen_tid:
                            continue
                        seen_tid.add(tid)
                        refresh_market_history_for_type(conn, MARKET_HISTORY_REGION_ID, tid)
                        time.sleep(0.12)
                finally:
                    conn.close()
            except Exception as ex:
                err_msg = str(ex)

            def done():
                if err_msg:
                    messagebox.showerror("Market history", err_msg)
                    self.status_var.set("Market history refresh failed.")
                else:
                    self._shopping_list_refresh_tree()
                    self.status_var.set("Market history refreshed for shopping list products.")

            self.root.after(0, done)

        threading.Thread(target=work, daemon=True).start()

    def _shopping_list_refresh_market_history_selected(self):
        """Fetch market history for the selected row's product only, then refresh the tree."""
        sel = self.shopping_list_tree.selection()
        if not sel:
            messagebox.showinfo("Shopping list", "Select a blueprint row first.")
            return
        children = list(self.shopping_list_tree.get_children())
        try:
            idx = children.index(sel[0])
        except ValueError:
            return
        if idx < 0 or idx >= len(self.shopping_list):
            return
        name = (self.shopping_list[idx].get("product_name") or "").strip()
        if not name:
            messagebox.showwarning("Shopping list", "Selected row has no product name.")
            return
        if not Path(DATABASE_FILE).exists():
            messagebox.showwarning("Shopping list", "Database not found.")
            return
        self.status_var.set(f"Refreshing market history for {name[:48]}…")

        def work():
            err_msg = None
            try:
                conn = sqlite3.connect(DATABASE_FILE)
                try:
                    bp = resolve_blueprint(conn, name)
                    if not bp:
                        err_msg = f"Could not resolve blueprint/product: {name!r}"
                    else:
                        tid = int(bp["productTypeID"])
                        discard_market_history_session_refresh(MARKET_HISTORY_REGION_ID, tid)
                        refresh_market_history_for_type(conn, MARKET_HISTORY_REGION_ID, tid)
                finally:
                    conn.close()
            except Exception as ex:
                err_msg = str(ex)

            def done():
                if err_msg:
                    messagebox.showerror("Market history", err_msg)
                    self.status_var.set("Market history refresh failed.")
                else:
                    self._shopping_list_refresh_tree()
                    self.status_var.set("Market history refreshed for selected product.")

            self.root.after(0, done)

        threading.Thread(target=work, daemon=True).start()

    def create_put_in_production_tab(self):
        """Track manufacturing and invention runs remaining per blueprint; persisted to JSON."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Put in production")
        self.production_tracking_by_product = {}
        self._load_production_tracking()
        top = ttk.LabelFrame(frame, text="Runs remaining (shopping list plan; change dropdowns as you launch jobs)", padding=8)
        top.pack(fill=tk.X, padx=10, pady=6)
        ttk.Label(
            top,
            text="Plan columns: Research (streams Σ) and Runs (invention runs per stream) match the shopping list; "
            "# prod (Σ BPCs) and Run/BPC match manufacturing; Prod. runs left = # prod × Run/BPC to deliver. "
            "Inv. runs left = Research × Runs (unless Own BPC skips invention). "
            "Sell @ prod and Breakeven are snapshotted (net sell-offer per item) the moment a row is put in production "
            "and are NOT updated with the market — use 'Reset saved progress' to re-capture them. "
            "Values are saved automatically.",
            wraplength=920,
            justify=tk.LEFT,
        ).pack(anchor=tk.W)
        btn_row = ttk.Frame(top)
        btn_row.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(btn_row, text="Refresh rows from shopping list", command=self._put_in_production_rebuild_rows).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="Reset saved progress", command=self._put_in_production_reset_progress).pack(side=tk.LEFT, padx=4)
        scroll_wrap = ttk.LabelFrame(frame, text="Blueprints", padding=6)
        scroll_wrap.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        canvas = tk.Canvas(scroll_wrap, highlightthickness=0)
        vsb = ttk.Scrollbar(scroll_wrap, orient=tk.VERTICAL, command=canvas.yview)
        self.put_production_inner = ttk.Frame(canvas)
        self.put_production_inner.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        _put_cw = canvas.create_window((0, 0), window=self.put_production_inner, anchor=tk.NW)

        def _put_canvas_configure(ev):
            canvas.itemconfigure(_put_cw, width=ev.width)

        canvas.bind("<Configure>", _put_canvas_configure)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._put_in_production_rebuild_rows()

    def _load_production_tracking(self):
        self.production_tracking_by_product = {}
        if not PRODUCTION_TRACKING_FILE.exists():
            return
        try:
            with open(PRODUCTION_TRACKING_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
            if isinstance(d, dict) and isinstance(d.get("by_product"), dict):
                for k, v in d["by_product"].items():
                    if not isinstance(v, dict):
                        continue
                    try:
                        pr = int(v.get("production_runs", 0))
                        ir = int(v.get("invention_runs", 0))
                        rec = {
                            "production_runs": max(0, pr),
                            "invention_runs": max(0, ir),
                        }
                        # Snapshot prices captured when first put in production.
                        # Stored as-is; never refreshed from market until progress reset.
                        if v.get("price_at_production") is not None:
                            try:
                                rec["price_at_production"] = float(v["price_at_production"])
                            except (TypeError, ValueError):
                                pass
                        if v.get("breakeven_price") is not None:
                            try:
                                rec["breakeven_price"] = float(v["breakeven_price"])
                            except (TypeError, ValueError):
                                pass
                        self.production_tracking_by_product[str(k)] = rec
                    except (TypeError, ValueError):
                        pass
        except Exception:
            pass

    def _save_production_tracking(self):
        try:
            if not hasattr(self, "production_tracking_by_product"):
                return
            out = {"by_product": dict(self.production_tracking_by_product)}
            with open(PRODUCTION_TRACKING_FILE, "w", encoding="utf-8") as f:
                json.dump(out, f, indent=2)
        except Exception:
            pass

    # ---- Portable data export / import -------------------------------------
    # User-specific config that should travel between computers. The huge
    # eve_manufacturing.db (SDE + market data) is NOT included; only the
    # user's working state and per-blueprint bindings are bundled so a recent
    # export can be dropped onto another install and stay functional.
    EXPORT_DB_TABLES = (
        "blueprint_material_override",
        "blueprint_mfg_cost_binding",
        "blueprint_datacore_bindings",
        "invention_recipes",
    )

    def _export_dump_db_table(self, conn, table):
        """Return {'columns': [...], 'rows': [[...], ...]} for a table, or None if absent."""
        try:
            cur = conn.execute(f"PRAGMA table_info({table})")
            cols = [r[1] for r in cur.fetchall()]
            if not cols:
                return None
            rows = conn.execute(f"SELECT {', '.join(cols)} FROM {table}").fetchall()
            return {"columns": cols, "rows": [list(r) for r in rows]}
        except Exception:
            return None

    def _build_export_bundle(self):
        """Assemble the portable JSON bundle from in-memory state + DB bindings."""
        from datetime import datetime as _dt
        bundle = {
            "format": "eve_launcher_export",
            "version": 1,
            "exported_at": _dt.now().isoformat(timespec="seconds"),
            "shopping_list": list(getattr(self, "shopping_list", []) or []),
            "production_tracking": {
                "by_product": dict(getattr(self, "production_tracking_by_product", {}) or {})
            },
            "db_bindings": {},
        }
        if Path(DATABASE_FILE).exists():
            try:
                conn = sqlite3.connect(DATABASE_FILE, timeout=30)
                try:
                    self._ensure_material_override_table(conn)
                    self._paste_mfg_ensure_binding_table(conn)
                    self._ensure_blueprint_datacore_bindings_table(conn)
                    self._ensure_invention_recipes_table(conn)
                    for tbl in self.EXPORT_DB_TABLES:
                        dumped = self._export_dump_db_table(conn, tbl)
                        if dumped is not None:
                            bundle["db_bindings"][tbl] = dumped
                finally:
                    conn.close()
            except Exception:
                pass
        return bundle

    def _export_state_to_file(self):
        """Prompt for a location and write the portable JSON bundle."""
        from datetime import datetime as _dt
        default_name = f"eve_launcher_export_{_dt.now().strftime('%Y%m%d_%H%M%S')}.json"
        path = filedialog.asksaveasfilename(
            title="Export shopping list + production data",
            defaultextension=".json",
            initialfile=default_name,
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            bundle = self._build_export_bundle()
            with open(path, "w", encoding="utf-8") as f:
                json.dump(bundle, f, indent=2)
        except Exception as e:
            messagebox.showerror("Export failed", f"Could not write export file:\n{e}")
            return
        sl_n = len(bundle.get("shopping_list", []))
        pt_n = len(bundle.get("production_tracking", {}).get("by_product", {}))
        binds = bundle.get("db_bindings", {})
        bind_n = sum(len(v.get("rows", [])) for v in binds.values())
        self.status_var.set(f"Exported {sl_n} shopping-list items, {pt_n} production entries to {Path(path).name}")
        messagebox.showinfo(
            "Export complete",
            f"Saved to:\n{path}\n\n"
            f"Shopping list items: {sl_n}\n"
            f"Production-tracking entries: {pt_n}\n"
            f"Blueprint binding rows: {bind_n}\n\n"
            "Copy this file to another computer and use 'Import data…' there.\n"
            "Note: the large market/SDE database is not included — the other install "
            "keeps using its own eve_manufacturing.db.",
        )

    def _import_apply_db_bindings(self, db_bindings):
        """Restore per-blueprint binding tables from an export bundle."""
        if not db_bindings or not Path(DATABASE_FILE).exists():
            return
        try:
            conn = sqlite3.connect(DATABASE_FILE, timeout=30)
        except Exception:
            return
        try:
            self._ensure_material_override_table(conn)
            self._paste_mfg_ensure_binding_table(conn)
            self._ensure_blueprint_datacore_bindings_table(conn)
            self._ensure_invention_recipes_table(conn)
            for tbl in self.EXPORT_DB_TABLES:
                payload = db_bindings.get(tbl)
                if not isinstance(payload, dict):
                    continue
                cols = payload.get("columns") or []
                rows = payload.get("rows") or []
                if not cols or not rows:
                    continue
                # Only use columns that still exist in the current table schema.
                cur = conn.execute(f"PRAGMA table_info({tbl})")
                existing = {r[1] for r in cur.fetchall()}
                use_idx = [i for i, c in enumerate(cols) if c in existing]
                if not use_idx:
                    continue
                use_cols = [cols[i] for i in use_idx]
                placeholders = ", ".join(["?"] * len(use_cols))
                sql = f"INSERT OR REPLACE INTO {tbl} ({', '.join(use_cols)}) VALUES ({placeholders})"
                to_insert = [[row[i] for i in use_idx] for row in rows if isinstance(row, list)]
                conn.executemany(sql, to_insert)
            conn.commit()
        except Exception:
            pass
        finally:
            conn.close()

    def _import_state_from_file(self):
        """Prompt for an export bundle and restore shopping list, production data, and bindings."""
        path = filedialog.askopenfilename(
            title="Import shopping list + production data",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                bundle = json.load(f)
        except Exception as e:
            messagebox.showerror("Import failed", f"Could not read file:\n{e}")
            return
        if not isinstance(bundle, dict) or bundle.get("format") != "eve_launcher_export":
            messagebox.showerror("Import failed", "This file is not a valid EVE Launcher export.")
            return

        sl = bundle.get("shopping_list")
        pt = bundle.get("production_tracking", {})
        pt_by = pt.get("by_product", {}) if isinstance(pt, dict) else {}
        sl_n = len(sl) if isinstance(sl, list) else 0
        pt_n = len(pt_by) if isinstance(pt_by, dict) else 0
        binds = bundle.get("db_bindings", {}) if isinstance(bundle.get("db_bindings"), dict) else {}
        bind_n = sum(len(v.get("rows", [])) for v in binds.values() if isinstance(v, dict))

        if not messagebox.askyesno(
            "Confirm import",
            f"Exported: {bundle.get('exported_at', 'unknown')}\n\n"
            f"Shopping list items: {sl_n}\n"
            f"Production-tracking entries: {pt_n}\n"
            f"Blueprint binding rows: {bind_n}\n\n"
            "This will REPLACE your current shopping list and production tracking, "
            "and merge the blueprint bindings into this computer's database. Continue?",
        ):
            return

        try:
            if isinstance(sl, list):
                with open(SHOPPING_LIST_FILE, "w", encoding="utf-8") as f:
                    json.dump(sl, f, indent=2)
            if isinstance(pt_by, dict):
                with open(PRODUCTION_TRACKING_FILE, "w", encoding="utf-8") as f:
                    json.dump({"by_product": pt_by}, f, indent=2)
            self._import_apply_db_bindings(binds)
        except Exception as e:
            messagebox.showerror("Import failed", f"Error while applying import:\n{e}")
            return

        # Reload in-memory state and refresh the UI.
        try:
            self._load_shopping_list()
            self._load_production_tracking()
            if hasattr(self, "_shopping_list_refresh_tree"):
                self._shopping_list_refresh_tree()
            if hasattr(self, "_refresh_shopping_list_aggregate"):
                self._refresh_shopping_list_aggregate()
            if hasattr(self, "_put_in_production_rebuild_rows"):
                self._put_in_production_rebuild_rows()
        except Exception:
            pass

        self.status_var.set(f"Imported {sl_n} shopping-list items, {pt_n} production entries from {Path(path).name}")
        messagebox.showinfo(
            "Import complete",
            f"Restored:\nShopping list items: {sl_n}\n"
            f"Production-tracking entries: {pt_n}\n"
            f"Blueprint binding rows merged: {bind_n}",
        )

    def _put_in_production_row_key(self, product_name, decryptor_name):
        name = (product_name or "").strip()
        d = (decryptor_name or "").strip()
        if d in ("", "No decryptor"):
            d = ""
        return f"{name}{PUT_IN_PRODUCTION_ROW_SEP}{d}"

    def _shopping_list_merged_production_plan(self):
        """row_key -> plan fields incl. research_streams, runs_display, prod_bpcs_display, rpb_display."""
        merged = {}
        for entry in self.shopping_list:
            name = (entry.get("product_name") or "").strip()
            if not name:
                continue
            dec_raw = (entry.get("decryptor_name") or "").strip()
            if dec_raw in ("", "No decryptor"):
                dec_key = ""
            else:
                dec_key = dec_raw
            key = self._put_in_production_row_key(name, dec_key)
            skip_inv = bool(entry.get("bpc_owned_skip_invention"))
            prod = self._sl_prod_runs(entry)
            rpb = max(1, int(entry.get("runs_per_bpc") or 1))
            plan_p = prod * rpb if prod > 0 else 0
            plan_i = 0 if skip_inv else self._sl_total_attempts(entry)
            if key not in merged:
                merged[key] = {
                    "product": name,
                    "decryptor_label": dec_key if dec_key else "—",
                    "plan_p": 0,
                    "plan_i": 0,
                    "research_streams": 0,
                    "runs_values": set(),
                    "prod_bpcs_sum": 0,
                    "rpb_values": set(),
                }
            merged[key]["plan_p"] += plan_p
            merged[key]["plan_i"] += plan_i
            if "prod" in entry:
                bpc_add = max(0, int(entry.get("prod") or 0))
            else:
                bpc_add = max(1, int(entry.get("quantity") or 1))
            merged[key]["prod_bpcs_sum"] += bpc_add
            if plan_p > 0:
                merged[key]["rpb_values"].add(rpb)
            if not skip_inv and self._sl_use_direct_attempts(entry):
                merged[key]["research_streams"] += max(0, int(entry.get("research") or 0))
                merged[key]["runs_values"].add(max(0, int(entry.get("runs_per_research") or 0)))
        for v in merged.values():
            runs_vals = sorted(v.pop("runs_values", set()))
            plan_i = int(v.get("plan_i") or 0)
            if not runs_vals:
                v["runs_display"] = "—"
            elif len(runs_vals) == 1:
                v["runs_display"] = str(runs_vals[0])
            else:
                lo, hi = runs_vals[0], runs_vals[-1]
                v["runs_display"] = f"{lo}–{hi}" if lo != hi else str(lo)
            rs = int(v.get("research_streams") or 0)
            v["research_display"] = str(rs) if rs > 0 else "—"
            pb = int(v.get("prod_bpcs_sum") or 0)
            v["prod_bpcs_display"] = str(pb) if pb > 0 else "—"
            rpb_vals = sorted(v.pop("rpb_values", set()))
            if not rpb_vals:
                v["rpb_display"] = "—"
            elif len(rpb_vals) == 1:
                v["rpb_display"] = str(rpb_vals[0])
            else:
                lo, hi = rpb_vals[0], rpb_vals[-1]
                v["rpb_display"] = f"{lo}–{hi}" if lo != hi else str(lo)
        return merged

    def _shopping_list_open_put_in_production(self):
        for i in range(self.notebook.index("end")):
            if self.notebook.tab(i, "text") == "Put in production":
                self.notebook.select(i)
                break
        self._put_in_production_rebuild_rows()

    def _put_in_production_on_change(self, row_key, field, var):
        try:
            v = int(var.get())
        except (ValueError, tk.TclError):
            return
        if not hasattr(self, "production_tracking_by_product"):
            self.production_tracking_by_product = {}
        rec = self.production_tracking_by_product.setdefault(row_key, {})
        rec[field] = max(0, v)
        self._save_production_tracking()

    def _put_in_production_rebuild_rows(self):
        if not hasattr(self, "put_production_inner"):
            return
        for w in self.put_production_inner.winfo_children():
            w.destroy()
        merged = self._shopping_list_merged_production_plan()
        row_keys = sorted(merged.keys(), key=lambda k: (merged[k]["product"].lower(), merged[k]["decryptor_label"].lower()))
        row = 0
        hdr = ("TkDefaultFont", 9, "bold")
        ttk.Label(self.put_production_inner, text="Blueprint / product", font=hdr).grid(row=row, column=0, sticky=tk.W, padx=4, pady=2)
        ttk.Label(self.put_production_inner, text="Research\n(streams Σ)", font=hdr).grid(row=row, column=1, sticky=tk.W, padx=4)
        ttk.Label(self.put_production_inner, text="Runs\n(inv. / stream)", font=hdr).grid(row=row, column=2, sticky=tk.W, padx=4)
        ttk.Label(self.put_production_inner, text="# prod\n(Σ BPCs)", font=hdr).grid(row=row, column=3, sticky=tk.W, padx=4)
        ttk.Label(self.put_production_inner, text="Run\n/BPC", font=hdr).grid(row=row, column=4, sticky=tk.W, padx=4)
        ttk.Label(self.put_production_inner, text="Prod. runs\nleft", font=hdr).grid(row=row, column=5, sticky=tk.W, padx=4)
        ttk.Label(self.put_production_inner, text="Decryptor", font=hdr).grid(row=row, column=6, sticky=tk.W, padx=4)
        ttk.Label(self.put_production_inner, text="Inv. runs\nleft", font=hdr).grid(row=row, column=7, sticky=tk.W, padx=4)
        ttk.Label(self.put_production_inner, text="Sell @ prod\n(snapshot)", font=hdr).grid(row=row, column=8, sticky=tk.W, padx=4)
        ttk.Label(self.put_production_inner, text="Breakeven\n(snapshot)", font=hdr).grid(row=row, column=9, sticky=tk.W, padx=4)
        row += 1
        if not row_keys:
            ttk.Label(
                self.put_production_inner,
                text="Add blueprints with # prod > 0 or invention attempts (Research × Runs).",
                wraplength=520,
            ).grid(row=row, column=0, columnspan=10, sticky=tk.W, padx=4, pady=6)
            return
        pl_conn = sqlite3.connect(DATABASE_FILE) if Path(DATABASE_FILE).exists() else None
        for rkey in row_keys:
            info = merged[rkey]
            name = info["product"]
            dec_lbl = info["decryptor_label"]
            plan_p = info["plan_p"]
            plan_i = info["plan_i"]
            res_disp = info.get("research_display", "—")
            runs_disp = info.get("runs_display", "—")
            prod_bpcs_disp = info.get("prod_bpcs_display", "—")
            rpb_disp = info.get("rpb_display", "—")
            if plan_p <= 0 and plan_i <= 0:
                continue
            saved = self.production_tracking_by_product.get(rkey)
            if saved is None:
                saved = self.production_tracking_by_product.get(name)
            saved = saved or {}
            try:
                sp = int(saved.get("production_runs", plan_p))
                si = int(saved.get("invention_runs", plan_i))
            except (TypeError, ValueError):
                sp, si = plan_p, plan_i
            sp = max(0, min(sp, plan_p))
            si = max(0, min(si, plan_i))
            # Snapshot the market sell price and breakeven the first time this row is put
            # in production. Once stored, they are kept verbatim (not refreshed from market);
            # 'Reset saved progress' clears them so they are re-captured on the next rebuild.
            price_snap = saved.get("price_at_production")
            bev_snap = saved.get("breakeven_price")
            if price_snap is None and bev_snap is None:
                price_snap, bev_snap = self._put_in_production_snapshot_prices(pl_conn, rkey, name)
            rec = {"production_runs": sp, "invention_runs": si}
            if price_snap is not None:
                rec["price_at_production"] = price_snap
            if bev_snap is not None:
                rec["breakeven_price"] = bev_snap
            self.production_tracking_by_product[rkey] = rec
            ttk.Label(self.put_production_inner, text=name, wraplength=260).grid(row=row, column=0, sticky=tk.W, padx=4, pady=2)
            ttk.Label(self.put_production_inner, text=res_disp, width=7).grid(row=row, column=1, sticky=tk.W, padx=4, pady=2)
            ttk.Label(self.put_production_inner, text=runs_disp, width=9).grid(row=row, column=2, sticky=tk.W, padx=4, pady=2)
            ttk.Label(self.put_production_inner, text=prod_bpcs_disp, width=7).grid(row=row, column=3, sticky=tk.W, padx=4, pady=2)
            ttk.Label(self.put_production_inner, text=rpb_disp, width=7).grid(row=row, column=4, sticky=tk.W, padx=4, pady=2)
            if plan_p > 0:
                pv = tk.StringVar(value=str(sp))
                opts = [str(i) for i in range(0, plan_p + 1)]
                cb = ttk.Combobox(self.put_production_inner, textvariable=pv, values=opts, state="readonly", width=9)
                cb.grid(row=row, column=5, sticky=tk.W, padx=4, pady=2)
                cb.bind(
                    "<<ComboboxSelected>>",
                    lambda _e, k=rkey, var=pv, f="production_runs": self._put_in_production_on_change(k, f, var),
                )
            else:
                ttk.Label(self.put_production_inner, text="—").grid(row=row, column=5, sticky=tk.W, padx=4)
            ttk.Label(self.put_production_inner, text=dec_lbl, wraplength=160).grid(row=row, column=6, sticky=tk.W, padx=4, pady=2)
            if plan_i > 0:
                iv = tk.StringVar(value=str(si))
                iopts = [str(i) for i in range(0, plan_i + 1)]
                cb2 = ttk.Combobox(self.put_production_inner, textvariable=iv, values=iopts, state="readonly", width=9)
                cb2.grid(row=row, column=7, sticky=tk.W, padx=4, pady=2)
                cb2.bind(
                    "<<ComboboxSelected>>",
                    lambda _e, k=rkey, var=iv, f="invention_runs": self._put_in_production_on_change(k, f, var),
                )
            else:
                ttk.Label(self.put_production_inner, text="—").grid(row=row, column=7, sticky=tk.W, padx=4)
            ttk.Label(
                self.put_production_inner,
                text=self._shopping_list_format_price_display(price_snap),
                width=12,
            ).grid(row=row, column=8, sticky=tk.E, padx=4, pady=2)
            ttk.Label(
                self.put_production_inner,
                text=self._shopping_list_format_price_display(bev_snap),
                width=12,
            ).grid(row=row, column=9, sticky=tk.E, padx=4, pady=2)
            row += 1
        if pl_conn is not None:
            pl_conn.close()
        self._save_production_tracking()

    def _put_in_production_reset_progress(self):
        """Clear persisted remaining-runs selections and restore values to plan defaults."""
        if not getattr(self, "production_tracking_by_product", None):
            messagebox.showinfo("Put in production", "Nothing to reset.")
            return
        ok = messagebox.askyesno(
            "Reset saved progress",
            "Reset all saved Put in production values back to current plan totals?\n\n"
            "This clears persisted progress across sessions.",
        )
        if not ok:
            return
        self.production_tracking_by_product = {}
        self._save_production_tracking()
        self._put_in_production_rebuild_rows()
        self.status_var.set("Put in production progress reset to plan totals.")

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
                """SELECT dc1_name, dc1_qty, dc2_name, dc2_qty, base_invention_chance_pct, invention_cost_per_attempt, base_bpc_runs, production_cost_per_run
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
            user_prod_cost_per_run = float(bind[7]) if len(bind) > 7 and bind[7] is not None and float(bind[7]) > 0 else None
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
                # Always pick the highest profit_per_bpc decryptor (re-evaluates on every refresh)
                best = max(valid, key=lambda x: x.get("profit_per_bpc") or -1e99)
                prev_dec = (entry.get("decryptor_name") or "").strip()
                selected = best
                new_dec = (best.get("decryptor_name") or "").strip()
                if prev_dec and prev_dec != new_dec:
                    entry["_decryptor_changed"] = f"{prev_dec} → {new_dec or 'No decryptor'}"
                else:
                    entry.pop("_decryptor_changed", None)
                entry["profit"] = selected.get("profit_per_bpc")
                new_rpb = max(1, int(selected.get("bpc_runs") or 10))
                # Always store the decryptor-derived default; only apply it if not manually overridden
                entry["default_runs_per_bpc"] = new_rpb
                if not entry.get("manual_runs_per_bpc"):
                    entry["runs_per_bpc"] = new_rpb
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
                # Build E[research] tooltip
                def _fmt(v, decimals=0):
                    try:
                        return f"{float(v):,.{decimals}f}" if v is not None else "n/a"
                    except (TypeError, ValueError):
                        return "n/a"
                dn_display = dn if dn and dn != "No decryptor" else "None"
                sp_val = selected.get("success_prob_pct")
                inv_cost_no_dec = selected.get("inv_cost_no_dec", inv_cost)
                dc_cost = selected.get("datacore_cost", 0) or 0
                dec_price = selected.get("decryptor_price", 0) or 0
                attempt_cost = selected.get("attempt_cost") or (inv_cost_no_dec + dc_cost + dec_price)
                exp_inv = selected.get("expected_inv_cost", 0) or 0
                mfg_profit_selected = selected.get("manufacturing_profit", 0) or 0
                profit_bpc = selected.get("profit_per_bpc", 0) or 0
                bpc_runs_sel = selected.get("bpc_runs", base_runs)
                bpc_me_sel = selected.get("bpc_me", 0) or 0
                sep = "─" * 48
                research_lines = [
                    f"=== E[research]: {name} ===",
                    f"Decryptor:               {dn_display}",
                    f"Base invention chance:   {base_chance_pct:.1f}%",
                    f"Effective chance:        {_fmt(sp_val, 1)}%",
                    f"Cost per attempt:        {_fmt(attempt_cost, 0)} ISK",
                    f"  Invention base:        {_fmt(inv_cost_no_dec, 0)} ISK",
                    f"  Datacores:             {_fmt(dc_cost, 0)} ISK",
                    f"  Decryptor:             {_fmt(dec_price, 0)} ISK",
                    f"Expected inv. cost/BPC:  {_fmt(exp_inv, 0)} ISK",
                    f"BPC:                     {bpc_runs_sel} run(s), ME {bpc_me_sel:.0f}%",
                    f"Mfg profit ({bpc_runs_sel}r @{bpc_me_sel:.0f}% ME): {_fmt(mfg_profit_selected, 0)} ISK",
                    sep,
                    f"Profit per BPC:          {_fmt(profit_bpc, 0)} ISK  (mfg − inv)",
                ]
                entry["_research_tooltip"] = "\n".join(research_lines)

                # Cache mfg-level columns (buy_immediate/sell_offer, total production runs)
                try:
                    me_for_calc = entry.get("manufacturing_me") or 0.0
                    _runs_1bpc = max(1, int(bpc_runs_sel))
                    # E[prod] = profit for ONE BPC (bpc_runs_sel runs at bpc_me)
                    mfg_result = calculate_blueprint_profitability(
                        blueprint_name_or_product=name,
                        input_price_type=input_price,
                        output_price_type=output_price,
                        system_cost_percent=system_cost_pct,
                        material_efficiency=me_for_calc,
                        number_of_runs=_runs_1bpc,
                        region_id=region_id,
                        db_file=DATABASE_FILE,
                    )
                    # E[prod -min] = same but output at sell_immediate
                    mfg_result_min = calculate_blueprint_profitability(
                        blueprint_name_or_product=name,
                        input_price_type=input_price,
                        output_price_type="sell_immediate",
                        system_cost_percent=system_cost_pct,
                        material_efficiency=me_for_calc,
                        number_of_runs=_runs_1bpc,
                        region_id=region_id,
                        db_file=DATABASE_FILE,
                    )
                    if mfg_result and "error" not in mfg_result:
                        out_rev_1 = mfg_result.get("output_revenue", 0) or 0
                        mat_cost_1 = mfg_result.get("total_input_cost", 0) or 0
                        sys_cost_1 = mfg_result.get("system_cost", 0) or 0
                        mfg_profit_1 = mfg_result.get("profit", 0) or 0
                        prod_n = self._sl_prod_runs(entry)   # # prod (number of BPCs)

                        if user_prod_cost_per_run is not None:
                            user_cost_1bpc = user_prod_cost_per_run * bpc_runs_sel
                            eff_profit_1 = out_rev_1 - mat_cost_1 - user_cost_1bpc
                        else:
                            eff_profit_1 = mfg_profit_1

                        entry["_cached_exp_profit"] = eff_profit_1
                        entry["_cached_total_cost"] = mat_cost_1

                        # E[prod -min]: same calc but with sell_immediate revenue
                        if mfg_result_min and "error" not in mfg_result_min:
                            out_rev_min = mfg_result_min.get("output_revenue", 0) or 0
                            sys_cost_min = mfg_result_min.get("system_cost", 0) or 0
                            if user_prod_cost_per_run is not None:
                                eff_profit_min = out_rev_min - mat_cost_1 - user_cost_1bpc
                            else:
                                eff_profit_min = mfg_result_min.get("profit", 0) or 0
                            entry["_cached_exp_profit_min"] = eff_profit_min
                        else:
                            entry.pop("_cached_exp_profit_min", None)

                        # Tooltip: per-BPC breakdown then scale to # prod
                        tip = [f"=== E[prod]: {name} ===",
                               f"Per BPC  ({bpc_runs_sel} run(s) @ ME {me_for_calc:.0f}%):"]
                        tip.append(f"  Input materials:         {_fmt(mat_cost_1, 0)} ISK")
                        if user_prod_cost_per_run is not None:
                            user_cost_1bpc_val = user_prod_cost_per_run * bpc_runs_sel
                            tip.append(f"  System cost ({system_cost_pct}% EIV):  {_fmt(sys_cost_1, 0)} ISK  ← NOT used")
                            tip.append(f"  ⚑ User prod cost:        {_fmt(user_prod_cost_per_run, 0)}/run × {bpc_runs_sel} = {_fmt(user_cost_1bpc_val, 0)} ISK")
                        else:
                            tip.append(f"  System cost ({system_cost_pct}% EIV):  {_fmt(sys_cost_1, 0)} ISK")
                        tip += [f"  Output revenue:          {_fmt(out_rev_1, 0)} ISK",
                                sep,
                                f"  Profit per BPC:          {_fmt(eff_profit_1, 0)} ISK"]
                        if prod_n > 0:
                            tip += [sep,
                                    f"For # prod = {prod_n} BPC(s)  ({prod_n * bpc_runs_sel} total runs):",
                                    f"  Total input cost:        {_fmt(mat_cost_1 * prod_n, 0)} ISK",
                                    f"  Total revenue:           {_fmt(out_rev_1 * prod_n, 0)} ISK",
                                    f"  Total profit:            {_fmt(eff_profit_1 * prod_n, 0)} ISK"]
                        entry["_prod_tooltip"] = "\n".join(tip)
                except Exception:
                    pass
                return
        entry.pop("manufacturing_me", None)
        # E[prod] is per-BPC, so calculate for runs_per_bpc runs (1 BPC)
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
        result_min = calculate_blueprint_profitability(
            blueprint_name_or_product=name,
            input_price_type=input_price,
            output_price_type="sell_immediate",
            system_cost_percent=system_cost_pct,
            material_efficiency=0,
            number_of_runs=runs,
            region_id=region_id,
            db_file=DATABASE_FILE,
        )
        if result and "error" not in result:
            entry["profit"] = result.get("profit")
            entry["default_runs_per_bpc"] = runs
            if not entry.get("manual_runs_per_bpc"):
                entry["runs_per_bpc"] = runs
            mat_cost_t1 = result.get("total_input_cost", 0) or 0
            out_rev_t1  = result.get("output_revenue", 0) or 0
            sys_cost_t1 = result.get("system_cost", 0) or 0
            mfg_profit_1 = result.get("profit", 0) or 0
            prod_n = self._sl_prod_runs(entry)   # # prod
            # Check for user-entered production cost
            t1_user_prod_cost = None
            if bp:
                try:
                    cr = conn.execute(
                        "SELECT production_cost_per_run FROM blueprint_datacore_bindings WHERE blueprint_type_id = ?",
                        (bp["blueprintTypeID"],),
                    ).fetchone()
                    if cr and cr[0] is not None:
                        v = float(cr[0])
                        if v > 0:
                            t1_user_prod_cost = v
                except Exception:
                    pass
            if t1_user_prod_cost is not None:
                user_cost_1bpc = t1_user_prod_cost * runs
                eff_profit_1 = out_rev_t1 - mat_cost_t1 - user_cost_1bpc
            else:
                eff_profit_1 = mfg_profit_1
            entry["_cached_exp_profit"] = eff_profit_1
            entry["_cached_total_cost"] = mat_cost_t1
            # E[prod -min] for T1
            if result_min and "error" not in result_min:
                out_rev_min_t1 = result_min.get("output_revenue", 0) or 0
                if t1_user_prod_cost is not None:
                    entry["_cached_exp_profit_min"] = out_rev_min_t1 - mat_cost_t1 - (t1_user_prod_cost * runs)
                else:
                    entry["_cached_exp_profit_min"] = result_min.get("profit", 0) or 0
            else:
                entry.pop("_cached_exp_profit_min", None)
            # Build tooltips for T1 / plain blueprints
            def _fmt_t1(v, d=0):
                try:
                    return f"{float(v):,.{d}f}" if v is not None else "n/a"
                except (TypeError, ValueError):
                    return "n/a"
            sep = "─" * 48
            entry["_research_tooltip"] = "\n".join([
                f"=== E[research]: {name} ===",
                f"T1 / plain blueprint (no invention)",
                f"Runs per BPC:            {runs}",
                f"ME:                      0%",
                f"(E[research] = mfg profit per BPC for T1)",
                sep,
                f"Profit ({runs} run(s), ME 0%): {_fmt_t1(mfg_profit_1, 0)} ISK",
            ])
            tip = [f"=== E[prod]: {name} ===",
                   f"Per BPC  ({runs} run(s) @ ME 0%):"]
            tip.append(f"  Input materials:         {_fmt_t1(mat_cost_t1, 0)} ISK")
            if t1_user_prod_cost is not None:
                tip.append(f"  System cost ({system_cost_pct}% EIV):  {_fmt_t1(sys_cost_t1, 0)} ISK  ← NOT used")
                tip.append(f"  ⚑ User prod cost:        {_fmt_t1(t1_user_prod_cost, 0)}/run × {runs} = {_fmt_t1(t1_user_prod_cost * runs, 0)} ISK")
            else:
                tip.append(f"  System cost ({system_cost_pct}% EIV):  {_fmt_t1(sys_cost_t1, 0)} ISK")
            tip += [f"  Output revenue:          {_fmt_t1(out_rev_t1, 0)} ISK",
                    sep,
                    f"  Profit per BPC:          {_fmt_t1(eff_profit_1, 0)} ISK"]
            if prod_n > 0:
                tip += [sep,
                        f"For # prod = {prod_n} BPC(s):",
                        f"  Total input cost:        {_fmt_t1(mat_cost_t1 * prod_n, 0)} ISK",
                        f"  Total revenue:           {_fmt_t1(out_rev_t1 * prod_n, 0)} ISK",
                        f"  Total profit:            {_fmt_t1(eff_profit_1 * prod_n, 0)} ISK"]
            entry["_prod_tooltip"] = "\n".join(tip)

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
        total_runs = self._sl_total_production_runs(entry)
        if column == "Blueprint / Product":
            return (0, (entry.get("product_name") or "").lower())
        if column == "Research":
            return (0, float(int(entry.get("research") or 0)))
        if column == "Runs":
            return (0, float(int(entry.get("runs_per_research") or 0)))
        if column == "# prod":
            return (0, float(self._sl_prod_runs(entry)))
        if column == "Decryptor":
            return (0, self._shopping_list_decryptor_display(entry).lower())
        if column == "Run per BPC":
            return (0, float(int(entry.get("runs_per_bpc") or 1)))
        if column == "Total material cost":
            tc = entry.get("_cached_total_cost")
            return (0, float(tc) if tc is not None else float("-inf"))
        if column == "Sell immediate":
            si, _ = self._shopping_list_unit_sell_prices(conn, entry["product_name"])
            return (0, float(si) if si is not None else float("-inf"))
        if column == "Hist 7d avg":
            h = self._shopping_list_hist_7d_avg_price(conn, entry["product_name"])
            return (0, float(h) if h is not None else float("-inf"))
        if column == "Sell offer":
            _, so = self._shopping_list_unit_sell_prices(conn, entry["product_name"])
            return (0, float(so) if so is not None else float("-inf"))
        if column == "Breakeven":
            _, so = self._shopping_list_unit_sell_prices(conn, entry["product_name"])
            bev = self._shopping_list_breakeven_price(conn, entry, so)
            return (0, float(bev) if bev is not None else float("-inf"))
        if column == "E[prod]":
            ep = entry.get("_cached_exp_profit")
            return (0, float(ep) if ep is not None else float("-inf"))
        if column == "E[prod -min]":
            ep = entry.get("_cached_exp_profit_min")
            return (0, float(ep) if ep is not None else float("-inf"))
        if column == "E[research]":
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

    # ------------------------------------------------------------------
    # Shopping list: helpers for Research / Runs / Prod model
    # ------------------------------------------------------------------

    def _sl_total_attempts(self, entry):
        """Total invention attempts.

        New model (both 'research' and 'runs_per_research' keys present): research × runs.
          Value can be 0 — means explicitly no invention.
        Legacy (keys absent): falls back to entry['quantity'].
          Exception: if 'prod' is explicitly 0, treat attempts as 0 too so the entry
          is fully inactive (avoids stale datacores appearing after "Set prod to 0").
        """
        if "research" in entry and "runs_per_research" in entry:
            r = max(0, int(entry["research"] or 0))
            ru = max(0, int(entry["runs_per_research"] or 0))
            return r * ru
        # Legacy path — if prod was explicitly zeroed, treat this entry as inactive
        if "prod" in entry and int(entry.get("prod") or 0) == 0:
            return 0
        return max(1, int(entry.get("quantity") or 1))

    def _sl_prod_runs(self, entry):
        """Production runs per successful BPC.

        New model ('prod' key present): use entry['prod'] directly — 0 means no production.
        Legacy ('prod' key absent): falls back to entry['runs_per_bpc'].
        """
        if "prod" in entry:
            return max(0, int(entry["prod"] or 0))
        return max(1, int(entry.get("runs_per_bpc") or 1))

    def _sl_total_production_runs(self, entry):
        """Total manufacturing runs = # prod × Run per BPC.

        New model ('prod' key present):  entry['prod'] × entry['runs_per_bpc'].
          Returns 0 when prod is explicitly 0.
        Legacy ('prod' key absent): quantity × runs_per_bpc (backward-compatible).
        """
        if "prod" in entry:
            prod = max(0, int(entry["prod"] or 0))
            if prod == 0:
                return 0
            runs_per_bpc = max(1, int(entry.get("runs_per_bpc") or 1))
            return prod * runs_per_bpc
        # Legacy path
        q = max(1, int(entry.get("quantity") or 1))
        r = max(1, int(entry.get("runs_per_bpc") or 1))
        return q * r

    def _sl_use_direct_attempts(self, entry):
        """True when the new research/runs model is active (both keys present in the entry dict)."""
        return "research" in entry and "runs_per_research" in entry

    def _sl_display_strs(self, entry):
        """Return (res_str, runs_str, prod_str) for the Treeview columns.

        New-model fields show their actual value (including 0).
        Legacy fields (key absent) fall back to legacy display.
        """
        if "research" in entry:
            r = max(0, int(entry["research"] or 0))
            res_str = str(r)
        else:
            res_str = "—"
        if "runs_per_research" in entry:
            ru = max(0, int(entry["runs_per_research"] or 0))
            runs_str = str(ru)
        else:
            runs_str = "—"
        if "prod" in entry:
            p = max(0, int(entry["prod"] or 0))
            prod_str = str(p)
        else:
            prod_str = str(max(1, int(entry.get("runs_per_bpc") or 1)))
        return res_str, runs_str, prod_str

    # ── tooltip helpers ────────────────────────────────────────────────────────

    def _shopping_list_on_double_click(self, event):
        """Copy the item (product) name of the double-clicked row to the clipboard."""
        tree = self.shopping_list_tree
        row_id = tree.identify_row(event.y)
        if not row_id:
            return
        children = list(tree.get_children())
        try:
            idx = children.index(row_id)
        except ValueError:
            return
        if idx < 0 or idx >= len(self.shopping_list):
            return
        name = self.shopping_list[idx].get("product_name", "")
        if name:
            self.root.clipboard_clear()
            self.root.clipboard_append(name)
            self.status_var.set(f"Copied to clipboard: {name}")

    def _sl_tree_motion(self, event):
        """Schedule a tooltip when the cursor hovers over E[research] or E[prod] columns."""
        tree = self.shopping_list_tree
        row_id = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not row_id or not col:
            self._sl_cancel_tooltip()
            return
        try:
            col_idx = int(col.replace("#", "")) - 1
        except ValueError:
            self._sl_cancel_tooltip()
            return
        cols = getattr(self, "shopping_list_columns", ())
        if col_idx < 0 or col_idx >= len(cols):
            self._sl_cancel_tooltip()
            return
        col_name = cols[col_idx]
        if col_name not in ("E[research]", "E[prod]", "# prod"):
            self._sl_cancel_tooltip()
            return
        cell = (row_id, col_name)
        if cell == self._sl_tooltip_last_cell:
            return
        self._sl_cancel_tooltip()
        self._sl_tooltip_last_cell = cell
        rx = tree.winfo_rootx() + event.x + 12
        ry = tree.winfo_rooty() + event.y + 18
        self._sl_tooltip_after_id = self.root.after(
            1000, lambda: self._sl_show_tooltip(row_id, col_name, rx, ry)
        )

    def _sl_tree_click(self, event):
        """Single-click on the treeview: open inline editor when clicking the 'Run per BPC' column."""
        region = self.shopping_list_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col_id = self.shopping_list_tree.identify_column(event.x)
        row_id = self.shopping_list_tree.identify_row(event.y)
        if not row_id:
            return
        cols = self.shopping_list_tree["columns"]
        try:
            col_idx = int(col_id.lstrip("#")) - 1
            col_name = cols[col_idx]
        except (ValueError, IndexError):
            return
        if col_name == "Run per BPC":
            self._sl_start_rpb_edit(row_id, col_id)

    def _sl_start_rpb_edit(self, row_id, col_id):
        """Place a floating Entry widget over the 'Run per BPC' cell for inline editing."""
        self._sl_close_rpb_edit()  # close any existing editor
        children = list(self.shopping_list_tree.get_children())
        try:
            idx = children.index(row_id)
        except ValueError:
            return
        if idx < 0 or idx >= len(self.shopping_list):
            return
        bbox = self.shopping_list_tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        entry = self.shopping_list[idx]
        current_rpb = max(1, int(entry.get("runs_per_bpc") or 1))
        edit_var = tk.StringVar(value=str(current_rpb))
        edit_widget = ttk.Entry(self.shopping_list_tree, textvariable=edit_var, justify="center")
        edit_widget.place(x=x, y=y, width=w, height=h)
        edit_widget.select_range(0, tk.END)
        edit_widget.focus_set()
        self._sl_rpb_edit_widget = edit_widget
        self._sl_rpb_edit_var = edit_var
        self._sl_rpb_edit_row = row_id
        self._sl_rpb_edit_idx = idx

        def _commit(event=None):
            self._sl_commit_rpb_edit()

        def _cancel(event=None):
            self._sl_close_rpb_edit()

        edit_widget.bind("<Return>", _commit)
        edit_widget.bind("<KP_Enter>", _commit)
        edit_widget.bind("<Tab>", _commit)
        edit_widget.bind("<Escape>", _cancel)
        edit_widget.bind("<FocusOut>", _commit)

    def _sl_close_rpb_edit(self):
        """Destroy the inline RPB edit widget (clears reference first to prevent double-fire)."""
        widget = self._sl_rpb_edit_widget
        self._sl_rpb_edit_widget = None
        if widget:
            try:
                widget.destroy()
            except Exception:
                pass

    def _sl_commit_rpb_edit(self):
        """Read the inline Entry value, store it as a manual override, and refresh the row."""
        widget = self._sl_rpb_edit_widget
        if widget is None:
            return
        self._sl_rpb_edit_widget = None  # clear before destroy to prevent FocusOut re-entry
        raw = getattr(self, "_sl_rpb_edit_var", tk.StringVar()).get().strip()
        idx = getattr(self, "_sl_rpb_edit_idx", None)
        try:
            widget.destroy()
        except Exception:
            pass
        if idx is None or idx >= len(self.shopping_list):
            return
        try:
            new_rpb = max(1, int(raw))
        except (ValueError, TypeError):
            return
        entry = self.shopping_list[idx]
        # Preserve the decryptor-derived default before first manual override
        if not entry.get("manual_runs_per_bpc"):
            entry["default_runs_per_bpc"] = entry.get("runs_per_bpc", 1)
        entry["runs_per_bpc"] = new_rpb
        entry["manual_runs_per_bpc"] = True
        self._shopping_list_refresh_tree()
        self._refresh_shopping_list_aggregate()
        self._save_shopping_list()
        self.status_var.set(f"Run per BPC manually set to {new_rpb} for {entry['product_name']}. Click 'Revert RPB' to restore default.")

    def _sl_revert_rpb(self):
        """Revert the selected row's Run per BPC to its stored default (from decryptor/profitability refresh)."""
        sel = self.shopping_list_tree.selection()
        if not sel:
            messagebox.showinfo("Revert RPB", "Select a blueprint row first.")
            return
        children = list(self.shopping_list_tree.get_children())
        try:
            idx = children.index(sel[0])
        except ValueError:
            return
        if idx < 0 or idx >= len(self.shopping_list):
            return
        entry = self.shopping_list[idx]
        if not entry.get("manual_runs_per_bpc"):
            self.status_var.set(f"{entry['product_name']}: Run per BPC is not manually overridden.")
            return
        default_rpb = entry.get("default_runs_per_bpc", 1)
        entry["runs_per_bpc"] = default_rpb
        entry.pop("manual_runs_per_bpc", None)
        self._shopping_list_refresh_tree()
        self._refresh_shopping_list_aggregate()
        self._save_shopping_list()
        self.status_var.set(f"Run per BPC reverted to {default_rpb} for {entry['product_name']}.")

    def _sl_tree_leave(self, event):
        """Cancel any pending or visible tooltip when the cursor leaves the tree."""
        self._sl_cancel_tooltip()
        self._sl_tooltip_last_cell = (None, None)

    def _sl_cancel_tooltip(self):
        """Cancel the pending tooltip timer and destroy any visible tooltip window."""
        if self._sl_tooltip_after_id is not None:
            self.root.after_cancel(self._sl_tooltip_after_id)
            self._sl_tooltip_after_id = None
        if self._sl_tooltip_win is not None:
            try:
                self._sl_tooltip_win.destroy()
            except tk.TclError:
                pass
            self._sl_tooltip_win = None

    def _sl_show_tooltip(self, row_id, col_name, x, y):
        """Display a small tooltip window with the pre-computed calculation summary."""
        children = list(self.shopping_list_tree.get_children())
        try:
            idx = children.index(row_id)
        except ValueError:
            return
        if idx < 0 or idx >= len(self.shopping_list):
            return
        entry = self.shopping_list[idx]
        if col_name == "# prod":
            prod_n = self._sl_prod_runs(entry)
            rpb = max(1, int(entry.get("runs_per_bpc") or 1))
            total_mfg = prod_n * rpb
            text = (
                f"# prod = number of BPCs to manufacture\n"
                f"Each BPC has {rpb} run(s)  (Run per BPC)\n"
                f"─────────────────────────────────\n"
                f"{prod_n} BPC(s) × {rpb} run(s)/BPC = {total_mfg} total manufacturing runs\n"
                f"Materials = {total_mfg} × qty per run"
            )
        else:
            key = "_research_tooltip" if col_name == "E[research]" else "_prod_tooltip"
            text = entry.get(key)
            if not text:
                text = f"No data yet — click 'Refresh profitability' to compute {col_name}."
        win = tk.Toplevel(self.root)
        win.wm_overrideredirect(True)
        win.wm_geometry(f"+{x}+{y}")
        win.attributes("-topmost", True)
        lbl = tk.Label(
            win, text=text, justify=tk.LEFT,
            background="#ffffcc", foreground="#000000",
            relief="solid", borderwidth=1,
            font=("Courier", 9),
            padx=8, pady=5,
        )
        lbl.pack()
        self._sl_tooltip_win = win
        win.bind("<Leave>", lambda e: self._sl_cancel_tooltip())

    # ── end tooltip helpers ───────────────────────────────────────────────────

    def _sl_update_totals_bar(self):
        """Recompute and display the fixed totals bar below the shopping-list treeview."""
        if not hasattr(self, "_sl_total_vars"):
            return
        total_research  = 0
        total_prod      = 0
        total_mat_cost  = 0.0
        total_e_res     = 0.0
        total_e_prod    = 0.0
        has_mat = has_e_res = has_e_prod = False
        for entry in self.shopping_list:
            prod_n = self._sl_prod_runs(entry)
            total_research += max(0, int(entry.get("research") or 0))
            total_prod     += prod_n
            tc = entry.get("_cached_total_cost")
            if tc is not None:
                total_mat_cost += float(tc) * prod_n
                has_mat = True
            ep = entry.get("profit")
            if ep is not None:
                try:
                    total_e_res += float(ep) * prod_n
                    has_e_res = True
                except (TypeError, ValueError):
                    pass
            ee = entry.get("_cached_exp_profit")
            if ee is not None:
                total_e_prod += float(ee) * prod_n
                has_e_prod = True

        def _f(v):
            return f"{v:,.0f}"

        self._sl_total_vars["_tot_research"].set(_f(total_research) if total_research else "—")
        self._sl_total_vars["_tot_prod"].set(_f(total_prod)    if total_prod     else "—")
        self._sl_total_vars["_tot_mat_cost"].set(_f(total_mat_cost) if has_mat   else "—")
        self._sl_total_vars["_tot_e_research"].set(_f(total_e_res)  if has_e_res  else "—")
        self._sl_total_vars["_tot_e_prod"].set(_f(total_e_prod) if has_e_prod    else "—")

    def _sl_items_per_run(self, conn, product_name):
        """Return the outputQuantity of the blueprint (items produced per manufacturing run)."""
        try:
            bp = resolve_blueprint(conn, product_name)
            if bp:
                q = bp.get("outputQuantity")
                if q and int(q) > 0:
                    return int(q)
        except Exception:
            pass
        return 1

    def _shopping_list_breakeven_price(self, conn, entry, sell_off_unit=None):
        """
        Breakeven sell price per item: the sell price at which profit = 0.
        breakeven = sell_offer_per_item − profit_per_item
        profit_per_item = entry['profit'] / (runs_per_bpc × outputQuantity)
        Returns None if data is unavailable.
        """
        profit_per_bpc = entry.get("profit")
        if profit_per_bpc is None:
            return None
        try:
            profit_per_bpc = float(profit_per_bpc)
        except (TypeError, ValueError):
            return None
        # Breakeven is a per-item threshold and should not depend on how many BPCs
        # are currently scheduled in # prod.
        runs_per_bpc = max(1, int(entry.get("runs_per_bpc") or 1))
        items_per_run = self._sl_items_per_run(conn, entry["product_name"])
        items_per_bpc = runs_per_bpc * items_per_run
        if items_per_bpc <= 0:
            return None
        profit_per_item = profit_per_bpc / items_per_bpc
        if sell_off_unit is None:
            _, sell_off_unit = self._shopping_list_unit_sell_prices(conn, entry["product_name"])
        if sell_off_unit is None:
            return None
        return sell_off_unit - profit_per_item

    def _put_in_production_snapshot_prices(self, conn, rkey, product_name):
        """Snapshot (sell_offer_unit, breakeven_unit) for this row's product at the current market.

        Captured once, when the row is first put in production; never refreshed from market
        afterwards (only re-captured when saved progress is reset). sell_offer_unit is the
        net realised sell-offer price per item; breakeven_unit is the per-item sell price at
        which profit = 0, both consistent with the shopping list's columns.
        """
        if conn is None:
            return (None, None)
        entry = None
        for e in self.shopping_list:
            nm = (e.get("product_name") or "").strip()
            if not nm:
                continue
            dec_raw = (e.get("decryptor_name") or "").strip()
            dec_key = "" if dec_raw in ("", "No decryptor") else dec_raw
            if self._put_in_production_row_key(nm, dec_key) == rkey:
                entry = e
                break
        if entry is None:
            target = (product_name or "").strip()
            for e in self.shopping_list:
                if (e.get("product_name") or "").strip() == target:
                    entry = e
                    break
        sell_off = None
        breakeven = None
        try:
            _, sell_off = self._shopping_list_unit_sell_prices(conn, product_name)
        except Exception:
            sell_off = None
        if entry is not None:
            try:
                breakeven = self._shopping_list_breakeven_price(conn, entry, sell_off)
            except Exception:
                breakeven = None
        return (sell_off, breakeven)

    def _shopping_list_refresh_tree(self):
        """Rebuild the Treeview from self.shopping_list.
        
        Deliberately skips calculate_blueprint_profitability (expensive per-item DB call) so the
        tree rebuilds instantly.  'Expected profit' and 'Total material cost' stay as '—' until the
        user clicks 'Refresh profitability', which populates them from current prices.
        """
        for item in self.shopping_list_tree.get_children():
            self.shopping_list_tree.delete(item)
        if not self.shopping_list:
            self._sl_update_totals_bar()
            return
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                override_names = self._material_override_names(conn)
                for entry in self.shopping_list:
                    decryptor_str = self._shopping_list_decryptor_display(entry)
                    sell_imm, sell_off = self._shopping_list_unit_sell_prices(conn, entry["product_name"])
                    sell_imm_str = self._shopping_list_format_price_display(sell_imm)
                    sell_off_str = self._shopping_list_format_price_display(sell_off)
                    hist7 = self._shopping_list_hist_7d_avg_price(conn, entry["product_name"])
                    hist7_str = self._shopping_list_format_price_display(hist7)
                    profit_str = self._shopping_list_profit_cell(conn, entry)
                    bev = self._shopping_list_breakeven_price(conn, entry, sell_off)
                    breakeven_str = f"{bev:,.2f}" if bev is not None else "—"
                    cached_ep = entry.get("_cached_exp_profit")
                    cached_ep_min = entry.get("_cached_exp_profit_min")
                    cached_tc = entry.get("_cached_total_cost")
                    e_prod_str = f"{cached_ep:,.0f}" if cached_ep is not None else "—"
                    e_prod_min_str = f"{cached_ep_min:,.0f}" if cached_ep_min is not None else "—"
                    total_cost_str = f"{cached_tc:,.0f}" if cached_tc is not None else "—"
                    res_str, runs_str, prod_str = self._sl_display_strs(entry)
                    # Run per BPC = entry runs_per_bpc (set by decryptor refresh); hide when # prod = 0
                    prod_n = self._sl_prod_runs(entry)
                    rpb = max(1, int(entry.get("runs_per_bpc") or 1))
                    rpb_str = str(rpb) if prod_n > 0 else "—"
                    # E[research] = profit per BPC; E[prod] = total mfg profit for prod runs
                    _tags = []
                    if entry.get("manual_runs_per_bpc"):
                        _tags.append("manual_rpb")
                    if entry["product_name"] in override_names:
                        _tags.append("mat_override")
                    row_tags = tuple(_tags)
                    self.shopping_list_tree.insert(
                        "", tk.END,
                        values=(entry["product_name"], res_str, runs_str, prod_str, decryptor_str, rpb_str, total_cost_str, sell_imm_str, hist7_str, sell_off_str, breakeven_str, profit_str, e_prod_str, e_prod_min_str),
                        tags=row_tags,
                    )
            finally:
                conn.close()
        except Exception:
            for entry in self.shopping_list:
                decryptor_str = self._shopping_list_decryptor_display(entry)
                profit_str = self._format_shopping_list_profit(entry.get("profit"))
                res_str, runs_str, prod_str = self._sl_display_strs(entry)
                prod_n = self._sl_prod_runs(entry)
                rpb = max(1, int(entry.get("runs_per_bpc") or 1))
                rpb_str = str(rpb) if prod_n > 0 else "—"
                row_tags = ("manual_rpb",) if entry.get("manual_runs_per_bpc") else ()
                self.shopping_list_tree.insert(
                    "", tk.END,
                    values=(entry["product_name"], res_str, runs_str, prod_str, decryptor_str, rpb_str, "—", "—", "—", "—", "—", profit_str, "—", "—"),
                    tags=row_tags,
                )
        self._sl_update_totals_bar()

    def _on_shopping_list_selection(self, event=None):
        """When user selects a row, fill Research/Runs/Prod entries and update max-runs display."""
        sel = self.shopping_list_tree.selection()
        if not sel:
            self._sl_max_runs_var.set("—")
            return
        item = sel[0]
        children = list(self.shopping_list_tree.get_children())
        try:
            idx = children.index(item)
        except ValueError:
            return
        if idx < 0 or idx >= len(self.shopping_list):
            return
        entry = self.shopping_list[idx]
        self.shopping_list_research_var.set(str(max(0, int(entry.get("research") or 0))))
        self.shopping_list_runs_var.set(str(max(0, int(entry.get("runs_per_research") or 0))))
        # Show actual prod value if key exists; else legacy runs_per_bpc
        if "prod" in entry:
            self.shopping_list_prod_var.set(str(max(0, int(entry["prod"] or 0))))
        else:
            self.shopping_list_prod_var.set(str(max(1, int(entry.get("runs_per_bpc") or 1))))
        # Compute and display max runs based on stored research time for this blueprint
        self._sl_update_max_runs_label(entry=entry)

    def _sl_update_max_runs_label(self, entry=None):
        """Recompute the 'Max runs' label for the currently selected shopping-list row.

        Called on selection change and whenever the max-time fields change.
        Looks up research_time_days/hours/minutes from blueprint_datacore_bindings.
        """
        # Resolve entry from current selection if not provided
        if entry is None:
            sel = self.shopping_list_tree.selection()
            if not sel:
                self._sl_max_runs_var.set("—")
                return
            children = list(self.shopping_list_tree.get_children())
            try:
                idx = children.index(sel[0])
            except ValueError:
                self._sl_max_runs_var.set("—")
                return
            if idx < 0 or idx >= len(self.shopping_list):
                self._sl_max_runs_var.set("—")
                return
            entry = self.shopping_list[idx]
        # Parse max research time (days + hours → total hours)
        try:
            max_days = max(0.0, float(self.sl_max_research_days_var.get().strip() or "0"))
        except (ValueError, AttributeError):
            max_days = 6.0
        try:
            max_hrs = max(0.0, float(self.sl_max_research_hours_var.get().strip() or "0"))
        except (ValueError, AttributeError):
            max_hrs = 12.0
        max_total_hours = max_days * 24.0 + max_hrs
        if max_total_hours <= 0:
            self._sl_max_runs_var.set("—")
            return
        # Look up per-run research time for this blueprint
        try:
            if not Path(DATABASE_FILE).exists():
                self._sl_max_runs_var.set("—")
                return
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                self._ensure_blueprint_datacore_bindings_table(conn)
                bp = resolve_blueprint(conn, entry["product_name"])
                if not bp:
                    self._sl_max_runs_var.set("—")
                    return
                row = conn.execute(
                    "SELECT research_time_days, research_time_hours, research_time_minutes "
                    "FROM blueprint_datacore_bindings WHERE blueprint_type_id = ?",
                    (bp["blueprintTypeID"],),
                ).fetchone()
            finally:
                conn.close()
        except Exception:
            self._sl_max_runs_var.set("—")
            return
        if not row:
            self._sl_max_runs_var.set("—")
            return
        rd, rh, rm = (row[0] or 0), (row[1] or 0), (row[2] or 0)
        run_hours = float(rd) * 24.0 + float(rh) + float(rm) / 60.0
        if run_hours <= 0:
            self._sl_max_runs_var.set("—")
            return
        max_runs = int(max_total_hours / run_hours)  # floor
        self._sl_max_runs_var.set(str(max_runs) if max_runs > 0 else "< 1")

    def _shopping_list_update_quantity(self):
        """Save Research/Runs/Prod for the selected row; derives quantity and runs_per_bpc from them."""
        sel = self.shopping_list_tree.selection()
        if not sel:
            messagebox.showinfo("Shopping list", "Select a blueprint row first.")
            return
        try:
            research = max(0, int(self.shopping_list_research_var.get().strip() or "0"))
        except ValueError:
            research = 0
        try:
            runs_per_research = max(0, int(self.shopping_list_runs_var.get().strip() or "0"))
        except ValueError:
            runs_per_research = 0
        try:
            prod = max(0, int(self.shopping_list_prod_var.get().strip() or "0"))
        except ValueError:
            prod = 0
        item = sel[0]
        children = list(self.shopping_list_tree.get_children())
        try:
            idx = children.index(item)
        except ValueError:
            return
        if idx < 0 or idx >= len(self.shopping_list):
            return
        ent = self.shopping_list[idx]
        ent["research"] = research
        ent["runs_per_research"] = runs_per_research
        ent["prod"] = prod
        # Keep legacy quantity in sync for non-new-model code paths
        if research > 0 and runs_per_research > 0:
            ent["quantity"] = research * runs_per_research
        elif ent.get("quantity", 0) < 1:
            ent["quantity"] = 1
        # Do NOT overwrite runs_per_bpc here — it is set by the decryptor/profitability refresh.
        # For T1 entries that have never been refreshed, ensure a sensible default exists.
        if not ent.get("runs_per_bpc"):
            ent["runs_per_bpc"] = 1
        product_name = ent["product_name"]
        sell_imm_str, sell_off_str, breakeven_str = "—", "—", "—"
        hist7_str = "—"
        profit_str = ""
        total_cost_str = f"{ent['_cached_total_cost']:,.0f}" if ent.get("_cached_total_cost") is not None else "—"
        e_prod_str = f"{ent['_cached_exp_profit']:,.0f}" if ent.get("_cached_exp_profit") is not None else "—"
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                sell_imm, sell_off = self._shopping_list_unit_sell_prices(conn, product_name)
                sell_imm_str = self._shopping_list_format_price_display(sell_imm)
                sell_off_str = self._shopping_list_format_price_display(sell_off)
                h7 = self._shopping_list_hist_7d_avg_price(conn, product_name)
                hist7_str = self._shopping_list_format_price_display(h7)
                profit_str = self._shopping_list_profit_cell(conn, ent)
                bev = self._shopping_list_breakeven_price(conn, ent, sell_off)
                breakeven_str = f"{bev:,.2f}" if bev is not None else "—"
            finally:
                conn.close()
        except Exception:
            profit_str = self._format_shopping_list_profit(ent.get("profit"))
        decryptor_str = self._shopping_list_decryptor_display(ent)
        res_str, runs_str, prod_str = self._sl_display_strs(ent)
        prod_n = self._sl_prod_runs(ent)
        rpb = max(1, int(ent.get("runs_per_bpc") or 1))
        rpb_str = str(rpb) if prod_n > 0 else "—"
        e_prod_min_str = f"{ent['_cached_exp_profit_min']:,.0f}" if ent.get("_cached_exp_profit_min") is not None else "—"
        self.shopping_list_tree.item(
            item,
            values=(
                product_name,
                res_str,
                runs_str,
                prod_str,
                decryptor_str,
                rpb_str,
                total_cost_str,
                sell_imm_str,
                hist7_str,
                sell_off_str,
                breakeven_str,
                profit_str,
                e_prod_str,
                e_prod_min_str,
            ),
        )
        self._refresh_shopping_list_aggregate()
        self._save_shopping_list()
        self.status_var.set(f"Quantities updated: research={research}, runs={runs_per_research}, prod={prod}. Click 'Refresh profitability' to update Expected profit / Total material cost.")

    def _shopping_list_set_prod_zero(self):
        """Set prod (production runs) to 0 for the selected row."""
        sel = self.shopping_list_tree.selection()
        if not sel:
            messagebox.showinfo("Shopping list", "Select a blueprint row first.")
            return
        item = sel[0]
        children = list(self.shopping_list_tree.get_children())
        try:
            idx = children.index(item)
        except ValueError:
            return
        if idx < 0 or idx >= len(self.shopping_list):
            return
        ent = self.shopping_list[idx]
        ent["prod"] = 0
        ent["runs_per_bpc"] = 1
        self.shopping_list_prod_var.set("0")
        self._shopping_list_refresh_tree()
        self._refresh_shopping_list_aggregate()
        self._save_shopping_list()
        self.status_var.set(f"Prod set to 0 for {ent['product_name']}.")

    def _shopping_list_reset_all_quantities_zero(self):
        """Set Research/Runs/Prod to 0 for every shopping-list row (with confirmation)."""
        if not self.shopping_list:
            messagebox.showinfo("Shopping list", "The list is empty.")
            return
        ok = messagebox.askyesno(
            "Confirm reset",
            "Set Research, Runs, and # prod to 0 for ALL rows?\n\nThis cannot be undone automatically.",
        )
        if not ok:
            return
        for ent in self.shopping_list:
            if not isinstance(ent, dict):
                continue
            ent["research"] = 0
            ent["runs_per_research"] = 0
            ent["prod"] = 0
            # Keep legacy fields coherent and make row inactive in all code paths.
            ent["quantity"] = 0
        self.shopping_list_research_var.set("0")
        self.shopping_list_runs_var.set("0")
        self.shopping_list_prod_var.set("0")
        self._shopping_list_refresh_tree()
        self._refresh_shopping_list_aggregate()
        self._save_shopping_list()
        self.status_var.set("All shopping-list quantities reset to 0.")

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

    def _ensure_material_override_table(self, conn):
        """Per-blueprint custom material list (per 1 run) — overrides ME-derived quantities."""
        conn.execute(
            """CREATE TABLE IF NOT EXISTS blueprint_material_override (
                   blueprint_type_id INTEGER NOT NULL,
                   material_type_id INTEGER NOT NULL,
                   material_name TEXT NOT NULL,
                   quantity INTEGER NOT NULL,
                   PRIMARY KEY (blueprint_type_id, material_type_id)
               )"""
        )

    def _get_material_override(self, conn, blueprint_type_id):
        """Return [{materialTypeID, materialName, quantity}] (per run) for a blueprint, or None."""
        self._ensure_material_override_table(conn)
        rows = conn.execute(
            "SELECT material_type_id, material_name, quantity FROM blueprint_material_override "
            "WHERE blueprint_type_id = ? ORDER BY material_name",
            (blueprint_type_id,),
        ).fetchall()
        if not rows:
            return None
        return [{"materialTypeID": r[0], "materialName": r[1], "quantity": int(r[2])} for r in rows]

    def _set_material_override(self, conn, blueprint_type_id, materials):
        """Replace the override for a blueprint. materials: iterable of (type_id, name, qty)."""
        self._ensure_material_override_table(conn)
        conn.execute("DELETE FROM blueprint_material_override WHERE blueprint_type_id = ?", (blueprint_type_id,))
        conn.executemany(
            "INSERT INTO blueprint_material_override (blueprint_type_id, material_type_id, material_name, quantity) "
            "VALUES (?,?,?,?)",
            [(int(blueprint_type_id), int(mt), str(mn), int(q)) for (mt, mn, q) in materials],
        )
        conn.commit()

    def _clear_material_override(self, conn, blueprint_type_id):
        """Remove any custom material list bound to a blueprint."""
        self._ensure_material_override_table(conn)
        conn.execute("DELETE FROM blueprint_material_override WHERE blueprint_type_id = ?", (blueprint_type_id,))
        conn.commit()

    def _material_override_names(self, conn):
        """Return the set of product names that currently have a bound custom material list."""
        self._ensure_material_override_table(conn)
        try:
            rows = conn.execute(
                "SELECT DISTINCT b.productName FROM blueprint_material_override o "
                "JOIN blueprints b ON o.blueprint_type_id = b.blueprintTypeID"
            ).fetchall()
            return {r[0] for r in rows if r[0]}
        except Exception:
            return set()

    def _shopping_list_selected_entry(self):
        """Return (index, entry) for the currently selected shopping-list row, or (None, None)."""
        sel = self.shopping_list_tree.selection()
        if not sel:
            return None, None
        children = list(self.shopping_list_tree.get_children())
        try:
            idx = children.index(sel[0])
        except ValueError:
            return None, None
        if idx < 0 or idx >= len(self.shopping_list):
            return None, None
        return idx, self.shopping_list[idx]

    def _shopping_list_edit_materials(self):
        """Open an editor to bind a custom per-run material list to the selected blueprint.

        Quantities are the materials consumed to produce ONE run (the blueprint's base
        output). The aggregate multiplies these by total runs (# prod x Run per BPC).
        Use this to capture null-sec / structure / rig material bonuses that differ from
        the ME-implied amounts.
        """
        idx, entry = self._shopping_list_selected_entry()
        if entry is None:
            messagebox.showinfo("Edit material list", "Select a blueprint row in the list first.")
            return
        name = (entry.get("product_name") or "").strip()
        if not Path(DATABASE_FILE).exists():
            messagebox.showerror("Edit material list", "Database not found.")
            return
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            bp = resolve_blueprint(conn, name)
            if not bp:
                messagebox.showerror("Edit material list", f"No blueprint found for '{name}'.")
                return
            bid = bp["blueprintTypeID"]
            base = get_blueprint_materials(conn, bid)  # per-run, ME 0
            override = self._get_material_override(conn, bid)
        finally:
            conn.close()
        if not base:
            messagebox.showinfo("Edit material list", f"'{name}' has no manufacturing materials to edit.")
            return
        ov_qty = {m["materialTypeID"]: m["quantity"] for m in (override or [])}

        win = tk.Toplevel(self.root)
        win.title(f"Material list per run — {name}")
        win.geometry("520x560")
        ttk.Label(
            win,
            text=(f"Materials to produce ONE run of {name} (output x{int(bp['outputQuantity'])}).\n"
                  "Enter the amounts you actually consume in-game (e.g. at your null-sec structure) "
                  "for a 1-run job. The shopping list multiplies these by total runs (# prod x Run per BPC).\n"
                  "Defaults shown are the blueprint's ME 0 amounts."),
            justify=tk.LEFT, wraplength=490,
        ).pack(anchor=tk.W, padx=10, pady=(10, 6))
        if override is not None:
            ttk.Label(win, text="A custom list is currently bound to this blueprint.",
                      foreground="#a06a00").pack(anchor=tk.W, padx=10)

        body = ttk.Frame(win)
        body.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        canvas = tk.Canvas(body, highlightthickness=0)
        vsb = ttk.Scrollbar(body, orient=tk.VERTICAL, command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor=tk.NW)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Label(inner, text="Material", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=0, sticky=tk.W, padx=4, pady=2)
        ttk.Label(inner, text="Qty / run", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=1, sticky=tk.W, padx=4, pady=2)
        ttk.Label(inner, text="(ME0)", font=("TkDefaultFont", 8)).grid(row=0, column=2, sticky=tk.W, padx=4, pady=2)
        var_map = {}
        for r, m in enumerate(base, start=1):
            mt = m["materialTypeID"]
            mname = m["materialName"]
            base_q = int(m["quantity"])
            cur_q = ov_qty.get(mt, base_q)
            ttk.Label(inner, text=mname, wraplength=260).grid(row=r, column=0, sticky=tk.W, padx=4, pady=1)
            var = tk.StringVar(value=str(cur_q))
            ttk.Entry(inner, textvariable=var, width=12, justify=tk.RIGHT).grid(row=r, column=1, sticky=tk.W, padx=4, pady=1)
            ttk.Label(inner, text=f"{base_q:,}").grid(row=r, column=2, sticky=tk.W, padx=4, pady=1)
            var_map[mt] = (mname, var)

        def save():
            mats = []
            for mt, (mname, var) in var_map.items():
                try:
                    q = int(float(str(var.get()).replace(",", "").strip()))
                except (ValueError, TypeError):
                    messagebox.showerror("Invalid quantity", f"'{mname}' has a non-numeric quantity.", parent=win)
                    return
                mats.append((mt, mname, max(0, q)))
            c = sqlite3.connect(DATABASE_FILE)
            try:
                self._set_material_override(c, bid, mats)
            finally:
                c.close()
            self._refresh_shopping_list_aggregate()
            self._shopping_list_refresh_tree()
            self.status_var.set(f"Bound custom material list to {name}.")
            win.destroy()

        def reset():
            c = sqlite3.connect(DATABASE_FILE)
            try:
                self._clear_material_override(c, bid)
            finally:
                c.close()
            self._refresh_shopping_list_aggregate()
            self._shopping_list_refresh_tree()
            self.status_var.set(f"Reset {name} to default (ME-derived) materials.")
            win.destroy()

        btns = ttk.Frame(win)
        btns.pack(fill=tk.X, padx=10, pady=(0, 10))
        ttk.Button(btns, text="Save (bind to blueprint)", command=save).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Reset to default", command=reset).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side=tk.RIGHT, padx=4)

    def _refresh_shopping_list_aggregate(self):
        """Compute aggregated materials (and datacores) from shopping_list and update the text. Stores result in self.shopping_list_aggregated for inventory comparison."""
        self.shopping_list_aggregate_text.configure(state=tk.NORMAL)
        self.shopping_list_aggregate_text.delete(1.0, tk.END)
        self.shopping_list_aggregated = None
        if not self.shopping_list:
            self.shopping_list_aggregate_text.insert(tk.END, "Add blueprints from Single Blueprint, Decryptor comparison, or Planning, then set Research/Runs/Prod. "
                "When Research and Runs are set, datacores and decryptors = Research × Runs × qty per attempt (direct count). "
                "Manufacturing materials scale by Total runs = ceil(Research × Runs × success_prob) × Prod.")
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
                inactive_count = 0
                for entry in self.shopping_list:
                    name = entry["product_name"]
                    use_direct = self._sl_use_direct_attempts(entry)
                    total_attempts = self._sl_total_attempts(entry)
                    total_runs = self._sl_total_production_runs(entry)

                    # Skip entries that explicitly have nothing to do
                    if total_attempts == 0 and total_runs == 0:
                        inactive_count += 1
                        continue

                    dec_name = entry.get("decryptor_name")
                    dec_type_id = entry.get("decryptor_type_id")
                    if dec_name and dec_type_id and total_attempts > 0:
                        if use_direct:
                            dec_need = total_attempts  # research × runs, direct count
                        else:
                            dec_need = self._shopping_list_scaled_invention_qty(entry, total_attempts, 1)
                        aggregated[dec_name] = aggregated.get(dec_name, 0) + dec_need
                    bp = resolve_blueprint(conn, name)
                    if not bp:
                        # Not a blueprint (e.g. decryptor): add as direct item
                        aggregated[name] = aggregated.get(name, 0) + total_attempts
                        continue
                    bid = bp["blueprintTypeID"]
                    if total_runs > 0:
                        # Use a bound custom per-run material list if present, else the ME0 base.
                        materials = self._get_material_override(conn, bid) or get_blueprint_materials(conn, bid)
                        for m in materials:
                            mat_name = m["materialName"]
                            need = m["quantity"] * total_runs
                            aggregated[mat_name] = aggregated.get(mat_name, 0) + need
                    if total_attempts > 0:
                        row = conn.execute(
                            "SELECT dc1_name, dc1_qty, dc2_name, dc2_qty FROM blueprint_datacore_bindings WHERE blueprint_type_id = ?",
                            (bid,),
                        ).fetchone()
                        if row:
                            dc1_name, dc1_qty, dc2_name, dc2_qty = row
                            if dc1_name and dc1_qty:
                                if use_direct:
                                    n1 = total_attempts * int(dc1_qty or 0)
                                else:
                                    n1 = self._shopping_list_scaled_invention_qty(entry, total_attempts, int(dc1_qty or 0))
                                if n1:
                                    aggregated[dc1_name] = aggregated.get(dc1_name, 0) + n1
                            if dc2_name and dc2_qty:
                                if use_direct:
                                    n2 = total_attempts * int(dc2_qty or 0)
                                else:
                                    n2 = self._shopping_list_scaled_invention_qty(entry, total_attempts, int(dc2_qty or 0))
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
        text_body = "\n".join(lines) if lines else "No materials resolved."
        if inactive_count > 0:
            note = f"\n\n({inactive_count} entr{'y' if inactive_count == 1 else 'ies'} with Research=0 and Prod=0 are not contributing — set Research/Runs/Prod to include them.)"
            text_body += note
        self.shopping_list_aggregate_text.insert(tk.END, text_body)
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
        """Parse pasted inventory: shortfall, unused (not in plan), and excess (have − 3× need)."""
        self.shopping_list_shortfall_text.configure(state=tk.NORMAL)
        self.shopping_list_shortfall_text.delete(1.0, tk.END)
        eu = getattr(self, "shopping_list_excess_unused_text", None)
        if eu is not None:
            eu.configure(state=tk.NORMAL)
            eu.delete(1.0, tk.END)
        if not getattr(self, "shopping_list_aggregated", None):
            self._refresh_shopping_list_aggregate()
        aggregated = getattr(self, "shopping_list_aggregated", None) or {}
        if not aggregated:
            self.shopping_list_shortfall_text.insert(tk.END, "No required items (add blueprints and refresh list first).")
            self.shopping_list_shortfall_text.configure(state=tk.DISABLED)
            if eu is not None:
                eu.insert(tk.END, "No plan to compare.")
                eu.configure(state=tk.DISABLED)
            return
        raw = self.shopping_list_inventory_text.get(1.0, tk.END)
        inventory = self._parse_inventory_paste(raw)
        req_keys = set(aggregated.keys())
        have_by_key = {}
        for pasted_name, qty in inventory.items():
            key = self._normalize_inventory_key(pasted_name, req_keys)
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
            total_cost = 0.0
            unpriced = []
            try:
                if Path(DATABASE_FILE).exists():
                    conn = sqlite3.connect(DATABASE_FILE)
                    try:
                        for name, qty in shortfall.items():
                            row = conn.execute(
                                "SELECT p.sell_min FROM prices p "
                                "JOIN items i ON i.typeID = p.typeID "
                                "WHERE i.typeName = ? AND p.sell_min > 0",
                                (name,),
                            ).fetchone()
                            if row:
                                total_cost += qty * float(row[0])
                            else:
                                unpriced.append(name)
                    finally:
                        conn.close()
            except Exception:
                unpriced = list(shortfall.keys())
            lines.append("")
            lines.append(f"Total cost (sell immediate):  {total_cost:,.0f} ISK")
            if unpriced:
                lines.append(f"(no price for: {', '.join(sorted(unpriced))})")
            self.shopping_list_shortfall_text.insert(tk.END, "\n".join(lines))
        self.shopping_list_shortfall_text.configure(state=tk.DISABLED)
        unused = {}
        for pasted_name, qty in inventory.items():
            key = self._normalize_inventory_key(pasted_name, req_keys)
            if key not in aggregated:
                unused[pasted_name] = unused.get(pasted_name, 0) + qty
        excess = {}
        for name, need in aggregated.items():
            if need <= 0:
                continue
            have = have_by_key.get(name, 0)
            cap = 3 * need
            if have > cap:
                excess[name] = have - cap
        if eu is not None:
            parts = []
            if unused:
                parts.append("UNUSED (not in plan)\n")
                for name in sorted(unused.keys(), key=lambda x: x.lower()):
                    parts.append(f"{name}\t{unused[name]:,}\n")
            else:
                parts.append("UNUSED (not in plan)\n(none — all pasted lines matched a plan item)\n")
            parts.append("\n")
            if excess:
                parts.append("EXCESS (have − 3× need)\n")
                for name in sorted(excess.keys()):
                    parts.append(f"{name}\t{excess[name]:,}\n")
            else:
                parts.append("EXCESS (have − 3× need)\n(none)\n")
            eu.insert(tk.END, "".join(parts))
            eu.configure(state=tk.DISABLED)
        self.status_var.set("Compared inventory: shortfall, unused, and excess updated.")

    def _shopping_list_copy_shortfall(self):
        """Copy the shortfall item list to clipboard, excluding the total-cost summary line."""
        self.shopping_list_shortfall_text.configure(state=tk.NORMAL)
        raw = self.shopping_list_shortfall_text.get(1.0, tk.END)
        self.shopping_list_shortfall_text.configure(state=tk.DISABLED)
        lines = [
            ln for ln in raw.splitlines()
            if ln.strip()
            and not ln.startswith("Total cost")
            and not ln.startswith("(no price for")
            and not ln.startswith("You have everything")
            and not ln.startswith("No required items")
        ]
        if not lines:
            messagebox.showinfo("Copy shortfall", "Nothing to copy.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(lines))
        self.status_var.set(f"Copied {len(lines)} shortfall item(s) to clipboard.")

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
        """Copy the blueprint/quantities table to the clipboard as tab-separated lines."""
        if not self.shopping_list:
            messagebox.showinfo("Copy plan", "Shopping list is empty.")
            return
        lines = ["Blueprint / Product\tResearch\tRuns\t# prod\tDecryptor\tRun per BPC\tTotal material cost\tSell immediate\tHist 7d avg\tSell offer\tBreakeven\tE[research]\tE[prod]\tE[prod -min]"]
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                for entry in self.shopping_list:
                    decryptor_str = self._shopping_list_decryptor_display(entry)
                    sell_imm, sell_off = self._shopping_list_unit_sell_prices(conn, entry["product_name"])
                    sell_imm_str = self._shopping_list_format_price_display(sell_imm)
                    sell_off_str = self._shopping_list_format_price_display(sell_off)
                    hist7_str = self._shopping_list_format_price_display(
                        self._shopping_list_hist_7d_avg_price(conn, entry["product_name"])
                    )
                    e_prod_str = f"{entry['_cached_exp_profit']:,.0f}" if entry.get("_cached_exp_profit") is not None else "—"
                    e_prod_min_str = f"{entry['_cached_exp_profit_min']:,.0f}" if entry.get("_cached_exp_profit_min") is not None else "—"
                    total_cost_str = f"{entry['_cached_total_cost']:,.0f}" if entry.get("_cached_total_cost") is not None else "—"
                    e_research_str = self._shopping_list_profit_cell(conn, entry) or "—"
                    bev = self._shopping_list_breakeven_price(conn, entry, sell_off)
                    breakeven_str = f"{bev:,.2f}" if bev is not None else "—"
                    res_str, runs_str, prod_str = self._sl_display_strs(entry)
                    prod_n = self._sl_prod_runs(entry)
                    rpb = max(1, int(entry.get("runs_per_bpc") or 1))
                    rpb_str = str(rpb) if prod_n > 0 else "—"
                    lines.append(f"{entry['product_name']}\t{res_str}\t{runs_str}\t{prod_str}\t{decryptor_str}\t{rpb_str}\t{total_cost_str}\t{sell_imm_str}\t{hist7_str}\t{sell_off_str}\t{breakeven_str}\t{e_research_str}\t{e_prod_str}\t{e_prod_min_str}")
            finally:
                conn.close()
        except Exception:
            for entry in self.shopping_list:
                decryptor_str = self._shopping_list_decryptor_display(entry)
                e_prod_str = f"{entry['_cached_exp_profit']:,.0f}" if entry.get("_cached_exp_profit") is not None else "—"
                e_prod_min_str = f"{entry['_cached_exp_profit_min']:,.0f}" if entry.get("_cached_exp_profit_min") is not None else "—"
                total_cost_str = f"{entry['_cached_total_cost']:,.0f}" if entry.get("_cached_total_cost") is not None else "—"
                e_research_str = self._format_shopping_list_profit(entry.get("profit")) or "—"
                res_str, runs_str, prod_str = self._sl_display_strs(entry)
                prod_n = self._sl_prod_runs(entry)
                rpb = max(1, int(entry.get("runs_per_bpc") or 1))
                rpb_str = str(rpb) if prod_n > 0 else "—"
                lines.append(f"{entry['product_name']}\t{res_str}\t{runs_str}\t{prod_str}\t{decryptor_str}\t{rpb_str}\t{total_cost_str}\t—\t—\t—\t—\t{e_research_str}\t{e_prod_str}\t{e_prod_min_str}")
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
        sl = prefs.get("shopping_list") or {}
        # Shopping-list controls are created in create_shopping_list_tab(), which may run
        # after this method depending on tab init order.
        if hasattr(self, "sl_max_research_days_var") and sl.get("max_research_days") is not None:
            self.sl_max_research_days_var.set(str(sl["max_research_days"]))
        if hasattr(self, "sl_max_research_hours_var") and sl.get("max_research_hours") is not None:
            self.sl_max_research_hours_var.set(str(sl["max_research_hours"]))

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
            # Persist shopping list global settings
            max_days = self.sl_max_research_days_var.get().strip() if hasattr(self, "sl_max_research_days_var") else ""
            max_hours = self.sl_max_research_hours_var.get().strip() if hasattr(self, "sl_max_research_hours_var") else ""
            prefs["shopping_list"] = {
                "max_research_days": max_days,
                "max_research_hours": max_hours,
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
            ("research_time_days", "REAL"),
            ("research_time_hours", "REAL"),
            ("research_time_minutes", "REAL"),
            ("production_cost_per_run", "REAL"),
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

    def _decryptor_clear_fields(self):
        """Reset all editable fields in the Decryptor comparison tab to blank/default values."""
        self.decryptor_t1_name_var.set("")
        self.decryptor_dc1_name_var.set("")
        self.decryptor_dc1_qty_var.set("0")
        self.decryptor_dc2_name_var.set("")
        self.decryptor_dc2_qty_var.set("0")
        self.decryptor_base_chance_var.set("40")
        self.decryptor_inv_cost_var.set("0")
        self.decryptor_base_runs_var.set("10")
        self.decryptor_research_days_var.set("")
        self.decryptor_research_hours_var.set("")
        self.decryptor_research_minutes_var.set("")
        self.decryptor_prod_cost_var.set("")

    def _decryptor_load_from_t2_name(self):
        """Load blueprint data for the T2 name in the top field; clears all fields first."""
        name = self.decryptor_product_var.get().strip()
        if not name:
            messagebox.showwarning("Load blueprint", "Enter a T2 blueprint / product name first.")
            return
        self._decryptor_clear_fields()
        self._load_datacore_binding_for_product(name)

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
                              base_invention_chance_pct, invention_cost_per_attempt, base_bpc_runs,
                              research_time_days, research_time_hours, research_time_minutes,
                              production_cost_per_run
                       FROM blueprint_datacore_bindings WHERE blueprint_type_id = ?""",
                    (blueprint_type_id,),
                ).fetchone()
                # Look up associated T1 blueprint name from invention_recipes
                try:
                    t1_row = conn.execute(
                        """SELECT i.typeName FROM invention_recipes ir
                           JOIN items i ON i.typeID = ir.t1_blueprint_type_id
                           WHERE ir.t2_blueprint_type_id = ? LIMIT 1""",
                        (blueprint_type_id,),
                    ).fetchone()
                    self.decryptor_t1_name_var.set(t1_row[0] if t1_row else "")
                except Exception:
                    pass
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
                # Research time: always set (blank when NULL so stale values don't persist)
                self.decryptor_research_days_var.set(str(row[7]) if len(row) > 7 and row[7] is not None else "")
                self.decryptor_research_hours_var.set(str(row[8]) if len(row) > 8 and row[8] is not None else "")
                self.decryptor_research_minutes_var.set(str(row[9]) if len(row) > 9 and row[9] is not None else "")
                pc = row[10] if len(row) > 10 else None
                self.decryptor_prod_cost_var.set(str(int(pc)) if pc is not None else "")
                self.status_var.set(f"Loaded saved binding for {product_name} (datacores, chance, cost, runs, research time, prod cost).")
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
                research_days = self.get_float(self.decryptor_research_days_var, 0.0)
                research_hours = self.get_float(self.decryptor_research_hours_var, 0.0)
                research_minutes = self.get_float(self.decryptor_research_minutes_var, 0.0)
                prod_cost_str = self.decryptor_prod_cost_var.get().strip()
                prod_cost = float(prod_cost_str) if prod_cost_str else None
                conn.execute("""
                    INSERT OR REPLACE INTO blueprint_datacore_bindings
                    (blueprint_type_id, dc1_name, dc1_qty, dc2_name, dc2_qty,
                     base_invention_chance_pct, invention_cost_per_attempt, base_bpc_runs,
                     research_time_days, research_time_hours, research_time_minutes,
                     production_cost_per_run, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                """, (blueprint_type_id, dc1_name or None, dc1_qty, dc2_name or None, dc2_qty,
                      base_chance, inv_cost, int(base_runs),
                      research_days, research_hours, research_minutes, prod_cost))
                conn.commit()
                self.status_var.set(f"Binding saved for {product_name} (datacores, chance %, cost, runs, research time, prod cost).")
            finally:
                conn.close()
        except Exception as e:
            messagebox.showerror("Bind datacores", str(e))

    def _associate_t1_t2(self):
        """Save T1 → T2 association to invention_recipes; also saves research time and prod cost to blueprint_datacore_bindings."""
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
                self._ensure_invention_recipes_table(conn)
                self._ensure_blueprint_datacore_bindings_table(conn)
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
                # Save research time and production cost to blueprint_datacore_bindings for the T2 blueprint
                research_days = self.get_float(self.decryptor_research_days_var, 0.0)
                research_hours = self.get_float(self.decryptor_research_hours_var, 0.0)
                research_minutes = self.get_float(self.decryptor_research_minutes_var, 0.0)
                prod_cost_str = self.decryptor_prod_cost_var.get().strip()
                prod_cost = float(prod_cost_str) if prod_cost_str else None
                # Upsert only the new fields, preserving any existing datacore binding values
                existing = conn.execute(
                    "SELECT blueprint_type_id FROM blueprint_datacore_bindings WHERE blueprint_type_id = ?",
                    (t2_bp_id,)
                ).fetchone()
                if existing:
                    conn.execute("""
                        UPDATE blueprint_datacore_bindings
                        SET research_time_days=?, research_time_hours=?, research_time_minutes=?,
                            production_cost_per_run=?, updated_at=CURRENT_TIMESTAMP
                        WHERE blueprint_type_id=?
                    """, (research_days, research_hours, research_minutes, prod_cost, t2_bp_id))
                else:
                    conn.execute("""
                        INSERT INTO blueprint_datacore_bindings
                        (blueprint_type_id, dc1_name, dc1_qty, dc2_name, dc2_qty,
                         research_time_days, research_time_hours, research_time_minutes,
                         production_cost_per_run, updated_at)
                        VALUES (?, NULL, 0, NULL, 0, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                    """, (t2_bp_id, research_days, research_hours, research_minutes, prod_cost))
                conn.commit()
                self.status_var.set(f"Associated {t1_name} → {t2_name} and saved research time/prod cost.")
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
            "When C-N4OD compare is on, a ★ marks the most profitable of 'Reprocess Value/Item', 'Jita Landed' and "
            "'C-N4OD Sell' per row ('Jita Landed' = Jita sell price minus shipping to move it there: 1,000 ISK/m3 + 0.6% of value; "
            "'C-N4OD Sell' is the local sell price, no shipping). "
            "For manufacturing: paste one blueprint or product name per line (duplicate names are shown once; click a column header to sort). The grid shows 1-run profit at ME 0, ME 5 and ME 10, "
            "for output sold buy-now (into buy order) vs sell order, plus % return columns for ME 10. All prices are Jita and "
            "include transport each way (1,000 ISK/m3 + 0.6% of value); green = profit, red = loss. Click any cell for the full "
            "calculation breakdown, where you can override the manufacturing cost, recalculate, and bind that cost to the blueprint "
            "(bound blueprints are marked [bound] and reuse the saved cost)."
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
        repro_row1 = ttk.Frame(self.paste_compare_repro_params_frame)
        repro_row1.pack(fill=tk.X, anchor=tk.W)
        ttk.Label(repro_row1, text="Threshold (ISK):").pack(side=tk.LEFT, padx=5)
        self.paste_threshold_var = tk.StringVar(value="400000")
        ttk.Entry(repro_row1, textvariable=self.paste_threshold_var, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Label(repro_row1, text="Min ISK above reprocess to recommend Sell:").pack(side=tk.LEFT, padx=5)
        self.paste_sell_buffer_var = tk.StringVar(value="20000")
        ttk.Entry(repro_row1, textvariable=self.paste_sell_buffer_var, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Label(repro_row1, text="Yield %:").pack(side=tk.LEFT, padx=5)
        self.paste_yield_var = tk.StringVar(value="55.0")
        ttk.Entry(repro_row1, textvariable=self.paste_yield_var, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(repro_row1, text="Reprocessing cost %:").pack(side=tk.LEFT, padx=5)
        self.paste_repro_cost_var = tk.StringVar(value="3.37")
        ttk.Entry(repro_row1, textvariable=self.paste_repro_cost_var, width=8).pack(side=tk.LEFT, padx=5)
        # C-N4OD arbitrage comparison (Jita sell + shipping vs C-N4OD sell) — on its own line
        repro_row2 = ttk.Frame(self.paste_compare_repro_params_frame)
        repro_row2.pack(fill=tk.X, anchor=tk.W, pady=(4, 0))
        self.paste_cn_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            repro_row2,
            text="Compare C-N4OD (SSO)",
            variable=self.paste_cn_var,
        ).pack(side=tk.LEFT, padx=(5, 4))
        ttk.Button(repro_row2, text="Refresh C-N prices", command=self._paste_refresh_cn_prices).pack(side=tk.LEFT, padx=2)
        # Cached C-N4OD sell map for the session: {type_id: min_sell}
        self._cn4od_sell_cache = None
        self._cn4od_note = ""
        
        compare_btn = ttk.Button(params_frame, text="Compare", command=self.run_paste_compare)
        compare_btn.pack(side=tk.LEFT, padx=15)
        
        # Results table (two trees: reprocessing and manufacturing)
        results_frame = ttk.LabelFrame(frame, text="Results", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        # Reprocessing tree
        self.paste_compare_columns = ('Item Name', 'Qty', 'Sell Min', 'Buy Max', 'Reprocess Value/Item', 'Recommendation', 'Jita Landed', 'C-N4OD Sell', 'C-N Margin %')
        self.paste_compare_tree = ttk.Treeview(results_frame, columns=self.paste_compare_columns, show='headings', height=20, selectmode='browse')
        self.paste_compare_sort_column = None
        self.paste_compare_sort_reverse = False
        for col in self.paste_compare_columns:
            self.paste_compare_tree.heading(col, text=col, command=lambda c=col: self.sort_paste_compare_by(c))
        self.paste_compare_tree.column('Item Name', width=280, anchor=tk.W)
        self.paste_compare_tree.column('Qty', width=50, anchor=tk.E)
        self.paste_compare_tree.column('Sell Min', width=95, anchor=tk.E)
        self.paste_compare_tree.column('Buy Max', width=95, anchor=tk.E)
        self.paste_compare_tree.column('Reprocess Value/Item', width=130, anchor=tk.E)
        self.paste_compare_tree.column('Recommendation', width=110, anchor=tk.W)
        self.paste_compare_tree.column('Jita Landed', width=100, anchor=tk.E)
        self.paste_compare_tree.column('C-N4OD Sell', width=100, anchor=tk.E)
        self.paste_compare_tree.column('C-N Margin %', width=90, anchor=tk.E)
        self.paste_compare_scrollbar_repro = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.paste_compare_tree.yview)
        self.paste_compare_tree.configure(yscrollcommand=self.paste_compare_scrollbar_repro.set)
        self.paste_compare_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.paste_compare_scrollbar_repro.pack(side=tk.RIGHT, fill=tk.Y)
        # Manufacturing results: custom grid so each profit cell can be colored (red/green) and clicked for details.
        # (header text, scenario key or None for name column, is_pct)
        self._paste_mfg_col_specs = [
            ("Blueprint", None, False),
            ("ME0 buy-now\n(profit ISK)", "me0_bn", False),
            ("ME0 sell\n(profit ISK)", "me0_sell", False),
            ("ME5 buy-now\n(profit ISK)", "me5_bn", False),
            ("ME5 sell\n(profit ISK)", "me5_sell", False),
            ("ME10 buy-now\n(profit ISK)", "me10_bn", False),
            ("ME10 sell\n(profit ISK)", "me10_sell", False),
            ("ME10 buy-now\n(return %)", "me10_bn", True),
            ("ME10 sell\n(return %)", "me10_sell", True),
        ]
        self._paste_mfg_grid_rows = []
        self._paste_mfg_sort_col = None
        self._paste_mfg_sort_reverse = False
        self.paste_mfg_container = ttk.Frame(results_frame)
        self.paste_mfg_canvas = tk.Canvas(self.paste_mfg_container, highlightthickness=0)
        self.paste_mfg_vsb = ttk.Scrollbar(self.paste_mfg_container, orient=tk.VERTICAL, command=self.paste_mfg_canvas.yview)
        self.paste_mfg_grid = ttk.Frame(self.paste_mfg_canvas)
        self.paste_mfg_grid.bind(
            "<Configure>",
            lambda e: self.paste_mfg_canvas.configure(scrollregion=self.paste_mfg_canvas.bbox("all")),
        )
        self.paste_mfg_canvas.create_window((0, 0), window=self.paste_mfg_grid, anchor=tk.NW)
        self.paste_mfg_canvas.configure(yscrollcommand=self.paste_mfg_vsb.set)
        self.paste_mfg_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.paste_mfg_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._paste_mfg_message("Paste blueprints and click Compare.")
        # Sync visibility with mode (hide mfg params initially since default is reprocessing)
        self._paste_compare_switch_mode()
    
    def _paste_compare_switch_mode(self):
        """Show/hide params and result tree based on Reprocessing vs Manufacturing mode."""
        if self.paste_compare_mode_var.get() == "manufacturing":
            self.paste_compare_mfg_params_frame.pack(side=tk.LEFT, padx=15)
            self.paste_compare_repro_params_frame.pack_forget()
            self.paste_compare_tree.pack_forget()
            self.paste_compare_scrollbar_repro.pack_forget()
            self.paste_mfg_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        else:
            self.paste_compare_mfg_params_frame.pack_forget()
            self.paste_compare_repro_params_frame.pack(side=tk.LEFT)
            self.paste_mfg_container.pack_forget()
            self.paste_compare_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.paste_compare_scrollbar_repro.pack(side=tk.RIGHT, fill=tk.Y)
    
    @staticmethod
    def _paste_parse_num(s):
        """Parse a formatted ISK string (commas, ★ marker) into a float, or None."""
        try:
            return float(str(s).replace("★", "").replace(",", "").strip())
        except (ValueError, TypeError):
            return None

    def _paste_mark_winner(self, repro_str, jita_str, cn_str, repro_val):
        """Prefix the most profitable of (Reprocess Value/Item, Jita Landed, C-N4OD Sell) with ★.

        ttk.Treeview cannot bold a single cell, so ★ marks the winning value per row so it
        stands out at a glance. All three options are per-item proceeds and directly comparable:
        Reprocess Value, Jita Landed (Jita sell minus shipping), and C-N4OD Sell (sell locally).
        Returns (repro_str, jita_str, cn_str), with at most one marked.
        """
        WIN = "★ "
        jita_val = self._paste_parse_num(jita_str)
        cn_val = self._paste_parse_num(cn_str)
        try:
            rv = float(repro_val)
        except (TypeError, ValueError):
            rv = None

        options = []  # (value, tag)
        if rv is not None and rv > 0:
            options.append((rv, "repro"))
        if jita_val is not None and jita_val > 0:
            options.append((jita_val, "jita"))
        if cn_val is not None and cn_val > 0:
            options.append((cn_val, "cn"))
        if not options:
            return repro_str, jita_str, cn_str

        winner = max(options, key=lambda o: o[0])[1]
        if winner == "repro":
            return WIN + repro_str, jita_str, cn_str
        if winner == "jita":
            return repro_str, WIN + jita_str, cn_str
        return repro_str, jita_str, WIN + cn_str

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
        if col_index in (2, 3, 4, 6, 7):  # Sell Min, Buy Max, Reprocess, Jita Landed, C-N4OD Sell - numeric
            try:
                return (0, float(s.replace("★", "").replace(",", "").strip()))
            except ValueError:
                return (1, s)
        if col_index == 8:  # C-N Margin % - numeric (strip %)
            try:
                return (0, float(s.replace("%", "").replace(",", "")))
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
            self._paste_mfg_message("Calculating...")
        else:
            self.status_var.set("Comparing items...")
            for item in self.paste_compare_tree.get_children():
                self.paste_compare_tree.delete(item)
            self.paste_compare_tree.insert('', tk.END, values=("", "Comparing...", "", "", "", "", "", "", ""))
        self.root.update()
        
        def do_compare():
            try:
                if self.paste_compare_mode_var.get() == "manufacturing":
                    try:
                        self._run_paste_compare_manufacturing(lines)
                    except Exception as e:
                        self.root.after(0, lambda err=e: self._paste_mfg_message(f"Error: {err}"))
                        self.status_var.set("Error occurred")
                        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
                    return
                threshold = self.get_float(self.paste_threshold_var, 100000.0)
                sell_buffer_isk = self.get_float(self.paste_sell_buffer_var, 0.0)
                if sell_buffer_isk < 0:
                    sell_buffer_isk = 0.0
                yield_pct = self.get_float(self.paste_yield_var, 55.0)
                repro_cost_pct = self.get_float(self.paste_repro_cost_var, 3.37)

                cn_enabled = bool(self.paste_cn_var.get())
                cn_map = {}
                if cn_enabled:
                    cn_map, cn_note = self._ensure_cn4od_prices()
                    if not cn_map:
                        self.root.after(0, lambda n=cn_note: self.status_var.set(f"C-N4OD unavailable: {n}"))

                def cn_fields(tid, jita_sell, vol):
                    """(Jita Landed, C-N4OD Sell, C-N Margin %) strings.

                    Jita Landed = net realized selling in Jita after shipping the item there
                    from C-N: jita_sell - shipping, where shipping = 1,000 ISK/m3 + 0.6% of value.
                    """
                    if not cn_enabled:
                        return ("", "", "")
                    cn = cn_map.get(tid, 0.0)
                    if jita_sell <= 0:
                        return ("", f"{cn:,.2f}" if cn else "N/A", "")
                    shipping = 1000.0 * (vol or 0.0) + 0.006 * jita_sell
                    landed = jita_sell - shipping
                    if cn <= 0:
                        return (f"{landed:,.2f}", "N/A", "")
                    margin = (cn - landed) / jita_sell
                    return (f"{landed:,.2f}", f"{cn:,.2f}", f"{margin * 100:.1f}%")

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
                        cursor = conn.execute("SELECT typeID, packaged_volume FROM items WHERE typeName = ?", (name,))
                        row_item = cursor.fetchone()
                        if not row_item:
                            rows.append((name, str(qty), "N/A", "N/A", "N/A", "Not in DB", "", "", ""))
                            continue
                        type_id = row_item[0]
                        volume_m3 = float(row_item[1]) if row_item[1] is not None else 0.0
                        
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
                        cn1, cn2, cn3 = cn_fields(type_id, sell_min, volume_m3)
                        rows.append((name, str(qty), f"{sell_min:,.2f}" if sell_min else "N/A", f"{buy_max:,.2f}" if buy_max else "N/A", "N/A", "Not reprocessable", cn1, cn2, cn3))
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
                    cn1, cn2, cn3 = cn_fields(type_id, sell_min, volume_m3)
                    # ★-mark the most profitable of Reprocess Value vs Jita Landed vs C-N4OD Sell
                    repro_str, cn1, cn2 = self._paste_mark_winner(repro_str, cn1, cn2, reprocess_value_per_item)
                    rows.append((name, str(qty), sell_str, buy_str, repro_str, rec, cn1, cn2, cn3))
                
                for item in self.paste_compare_tree.get_children():
                    self.paste_compare_tree.delete(item)
                for r in rows:
                    self.paste_compare_tree.insert('', tk.END, values=r)
                if cn_enabled and cn_map:
                    self.status_var.set(f"Compare complete. {self._cn4od_note}")
                else:
                    self.status_var.set("Compare complete.")
            except Exception as e:
                for item in self.paste_compare_tree.get_children():
                    self.paste_compare_tree.delete(item)
                self.paste_compare_tree.insert('', tk.END, values=("", f"Error: {str(e)}", "", "", "", "", "", "", ""))
                self.status_var.set("Error occurred")
                messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        
        thread = threading.Thread(target=do_compare, daemon=True)
        thread.start()

    def _ensure_cn4od_prices(self):
        """Return ({type_id: C-N4OD min sell}, note), fetching once per session via SSO."""
        if self._cn4od_sell_cache is not None:
            return self._cn4od_sell_cache, self._cn4od_note
        cid, secret = self._load_sso_credentials()
        if not cid or not secret:
            self._cn4od_sell_cache = {}
            self._cn4od_note = "No SSO credentials (set them in EVE SSO Sync tab)."
            return self._cn4od_sell_cache, self._cn4od_note
        try:
            import arbitrage_finder as af
            conn = sqlite3.connect(DATABASE_FILE, timeout=60)
            conn.execute("PRAGMA busy_timeout=60000")
            try:
                def prog(msg):
                    self.root.after(0, lambda m=msg: self.status_var.set(m))
                prices, note = af.load_cn4od_sell_prices(conn, cid, secret, progress=prog)
            finally:
                conn.close()
            self._cn4od_sell_cache = prices
            self._cn4od_note = note
        except Exception as e:
            self._cn4od_sell_cache = {}
            self._cn4od_note = f"C-N4OD fetch failed: {e}"
        return self._cn4od_sell_cache, self._cn4od_note

    def _paste_refresh_cn_prices(self):
        """Clear cached C-N4OD prices and re-run the comparison with C-N4OD enabled."""
        self._cn4od_sell_cache = None
        self._cn4od_note = ""
        self.paste_cn_var.set(True)
        self.run_paste_compare()

    def _paste_mfg_ensure_binding_table(self, conn):
        """Create the per-blueprint manufacturing-cost override table if missing."""
        conn.execute(
            """CREATE TABLE IF NOT EXISTS blueprint_mfg_cost_binding (
                   product_type_id INTEGER PRIMARY KEY,
                   product_name TEXT,
                   manufacturing_cost REAL,
                   updated_at TEXT
               )"""
        )

    def _paste_mfg_load_bindings(self):
        """Return {product_type_id: manufacturing_cost} for all bound blueprints."""
        try:
            conn = sqlite3.connect(DATABASE_FILE, timeout=30)
            try:
                self._paste_mfg_ensure_binding_table(conn)
                rows = conn.execute(
                    "SELECT product_type_id, manufacturing_cost FROM blueprint_mfg_cost_binding"
                ).fetchall()
            finally:
                conn.close()
            return {int(r[0]): float(r[1]) for r in rows if r[0] is not None and r[1] is not None}
        except Exception:
            return {}

    def _paste_mfg_set_binding(self, product_type_id, product_name, manufacturing_cost):
        """Persist (or clear, if manufacturing_cost is None) a per-blueprint manufacturing cost."""
        if not product_type_id:
            return
        conn = sqlite3.connect(DATABASE_FILE, timeout=30)
        try:
            self._paste_mfg_ensure_binding_table(conn)
            if manufacturing_cost is None:
                conn.execute(
                    "DELETE FROM blueprint_mfg_cost_binding WHERE product_type_id = ?",
                    (int(product_type_id),),
                )
            else:
                conn.execute(
                    """INSERT INTO blueprint_mfg_cost_binding (product_type_id, product_name, manufacturing_cost, updated_at)
                       VALUES (?, ?, ?, datetime('now'))
                       ON CONFLICT(product_type_id) DO UPDATE SET
                           product_name = excluded.product_name,
                           manufacturing_cost = excluded.manufacturing_cost,
                           updated_at = excluded.updated_at""",
                    (int(product_type_id), product_name, float(manufacturing_cost)),
                )
            conn.commit()
        finally:
            conn.close()

    def _run_paste_compare_manufacturing(self, lines):
        """Manufacturing profit grid for each pasted blueprint (1 run).

        Columns: ME0/ME5/ME10 buy-now and sell (profit ISK), plus ME10 buy-now /
        ME10 sell as % return. Inputs are bought at Jita sell price; output is sold
        at Jita buy-now (into buy order) or sell order. Both input and output include
        transport to C-N (1000 ISK/m3). Duplicate blueprint names are shown once.
        Runs in do_compare's thread.
        """
        system_cost_pct = self.get_float(self.paste_compare_system_cost_var, 8.61)
        if system_cost_pct < 0:
            system_cost_pct = 0.0
        region_id = MARKET_HISTORY_REGION_ID
        bindings = self._paste_mfg_load_bindings()
        # scenario key -> (material_efficiency, output_price_type)
        scenarios = {
            "me0_bn": (0, "sell_immediate"),
            "me0_sell": (0, "sell_offer"),
            "me5_bn": (5, "sell_immediate"),
            "me5_sell": (5, "sell_offer"),
            "me10_bn": (10, "sell_immediate"),
            "me10_sell": (10, "sell_offer"),
        }
        grid_rows = []
        seen_names = set()
        for line in lines:
            name = line.split('\t')[0].strip() if line else ""
            if not name:
                continue
            key_name = name.casefold()
            if key_name in seen_names:  # collapse duplicate blueprint lines
                continue
            seen_names.add(key_name)
            cells = {}
            row_bound = False
            for key, (me, out_type) in scenarios.items():
                result = calculate_blueprint_profitability(
                    blueprint_name_or_product=name,
                    input_price_type="buy_immediate",
                    output_price_type=out_type,
                    system_cost_percent=system_cost_pct,
                    material_efficiency=me,
                    number_of_runs=1,
                    region_id=region_id,
                    db_file=DATABASE_FILE,
                )
                ptid = result.get("productTypeID") if isinstance(result, dict) else None
                binding_cost = bindings.get(ptid) if ptid is not None else None
                if binding_cost is not None:
                    row_bound = True
                cells[key] = self._paste_mfg_derive(result, me, out_type, binding_cost=binding_cost)
            grid_rows.append({"name": name, "cells": cells, "bound": row_bound})
        self.root.after(0, lambda: self._paste_mfg_populate_grid(grid_rows))
        self.status_var.set("Manufacturing compare complete.")

    def _paste_mfg_derive(self, result, me, out_type, binding_cost=None):
        """Turn a calculate_blueprint_profitability result into a cell dict with a
        transport-to-C-N adjusted profit and a detail breakdown for the popup."""
        if not result or "error" in result:
            return {"error": (result or {}).get("error", "N/A"), "profit": None, "return_pct": None, "detail": None}
        input_cost = float(result.get("total_input_cost") or 0.0)
        calc_system_cost = float(result.get("system_cost") or 0.0)
        bound = binding_cost is not None
        system_cost = float(binding_cost) if bound else calc_system_cost
        input_volume = float(result.get("total_input_volume_m3") or 0.0)
        output_volume = float(result.get("output_total_volume_m3") or 0.0)
        out_qty = float(result.get("output_total_quantity") or 0.0)
        if out_type == "sell_immediate":
            gross_unit = float(result.get("output_buy_max") or 0.0)
            out_mode = "buy-now (sell into buy order)"
        else:
            gross_unit = float(result.get("output_sell_min") or 0.0)
            out_mode = "sell order"
        gross_revenue = gross_unit * out_qty
        # Transport each way = 1,000 ISK/m3 + 0.6% of the goods' value (Jita value).
        input_transport = 1000.0 * input_volume + 0.006 * input_cost
        output_transport = 1000.0 * output_volume + 0.006 * gross_revenue
        # Selling cost = fees on the Jita sell price (always a positive deduction).
        #   Buy-now (sell into buy order): sales tax only.
        #   Sell order: broker + sales tax + relisting fees (RELIST_DISCOUNT is a % discount).
        from assumptions import BROKER_FEE, SALES_TAX, LISTING_RELIST, RELIST_DISCOUNT
        if out_type == "sell_immediate":
            selling_cost = gross_revenue * (SALES_TAX / 100.0)
        else:
            broker = gross_revenue * (BROKER_FEE / 100.0)
            tax = gross_revenue * (SALES_TAX / 100.0)
            relist = gross_revenue * (BROKER_FEE / 100.0) * (1.0 - RELIST_DISCOUNT / 100.0) * LISTING_RELIST
            selling_cost = broker + tax + relist
        landed_sale = gross_revenue - selling_cost - output_transport
        # Cost basis (per user's format): materials + manufacturing + input transport.
        cost = input_cost + system_cost + input_transport
        profit = landed_sale - cost
        return_pct = (profit / cost * 100.0) if cost > 0 else None
        detail = {
            "product": result.get("productName", ""),
            "product_type_id": result.get("productTypeID"),
            "me": me,
            "out_mode": out_mode,
            "materials": result.get("input_materials", []),
            "total_input_cost": input_cost,
            "system_cost": system_cost,
            "calc_system_cost": calc_system_cost,
            "bound": bound,
            "system_cost_percent": float(result.get("system_cost_percent") or 0.0),
            "input_transport": input_transport,
            "input_volume": float(result.get("total_input_volume_m3") or 0.0),
            "cost": cost,
            "out_qty": out_qty,
            "gross_unit": gross_unit,
            "gross_revenue": gross_revenue,
            "selling_cost": selling_cost,
            "output_transport": output_transport,
            "output_volume": float(result.get("output_total_volume_m3") or 0.0),
            "net_revenue": gross_revenue - selling_cost,
            "landed_sale": landed_sale,
            "profit": profit,
            "return_pct": return_pct,
        }
        return {"error": None, "profit": profit, "return_pct": return_pct, "detail": detail}

    def _paste_mfg_clear(self):
        """Remove all widgets from the manufacturing grid."""
        for child in self.paste_mfg_grid.winfo_children():
            child.destroy()

    def _paste_mfg_message(self, msg):
        """Show a single status/message line in the manufacturing grid."""
        self._paste_mfg_clear()
        tk.Label(self.paste_mfg_grid, text=msg, anchor=tk.W, justify=tk.LEFT, padx=8, pady=8).grid(row=0, column=0, sticky="w")

    def _paste_mfg_cell(self, row, col, value, is_pct=False, detail=None):
        """Create one colored, clickable profit/return cell in the grid."""
        if value is None:
            text, bg = "N/A", "#f0f0f0"
        else:
            text = f"{value:,.1f}%" if is_pct else f"{value:,.0f}"
            bg = "#ccffcc" if value >= 0 else "#ffcccc"
        lbl = tk.Label(self.paste_mfg_grid, text=text, bg=bg, anchor=tk.E,
                       relief=tk.SOLID, borderwidth=1, padx=6, pady=3, width=14)
        lbl.grid(row=row, column=col, sticky="nsew", padx=1, pady=1)
        if detail is not None:
            lbl.configure(cursor="hand2")
            lbl.bind("<Button-1>", lambda e, d=detail: self._paste_mfg_show_detail(d))
        return lbl

    def _paste_mfg_populate_grid(self, grid_rows):
        """Store rows and render the manufacturing profit grid (called on the main thread)."""
        self._paste_mfg_grid_rows = grid_rows or []
        self._paste_mfg_sort_col = None
        self._paste_mfg_sort_reverse = False
        self._paste_mfg_render_grid()

    def _paste_mfg_cell_value(self, rowdata, col_index):
        """Return the sortable value for a row at a given column index."""
        header, key, is_pct = self._paste_mfg_col_specs[col_index]
        if key is None:
            return rowdata.get("name", "").casefold()
        cell = rowdata.get("cells", {}).get(key, {})
        return cell.get("return_pct") if is_pct else cell.get("profit")

    def _paste_mfg_sort_by(self, col_index):
        """Sort the grid rows by the clicked column (toggle direction) and re-render."""
        if self._paste_mfg_sort_col == col_index:
            self._paste_mfg_sort_reverse = not self._paste_mfg_sort_reverse
        else:
            self._paste_mfg_sort_col = col_index
            self._paste_mfg_sort_reverse = (col_index != 0)  # numbers default high->low, names A->Z
        is_name = col_index == 0
        reverse = self._paste_mfg_sort_reverse

        def sort_key(rowdata):
            v = self._paste_mfg_cell_value(rowdata, col_index)
            if is_name:
                return (0, v)
            # None (N/A) always sorts to the bottom regardless of direction
            if v is None:
                return (1, 0.0) if not reverse else (-1, 0.0)
            return (0, v)

        self._paste_mfg_grid_rows.sort(key=sort_key, reverse=reverse)
        self._paste_mfg_render_grid()

    def _paste_mfg_render_grid(self):
        """Draw headers (clickable to sort) and the current (sorted) rows."""
        self._paste_mfg_clear()
        grid_rows = self._paste_mfg_grid_rows
        if not grid_rows:
            self._paste_mfg_message("No blueprints found.")
            return
        for c, (header, _key, _is_pct) in enumerate(self._paste_mfg_col_specs):
            text = header
            if self._paste_mfg_sort_col == c:
                text = header + ("  v" if self._paste_mfg_sort_reverse else "  ^")
            anchor = tk.W if c == 0 else tk.CENTER
            hlbl = tk.Label(self.paste_mfg_grid, text=text, font=("TkDefaultFont", 9, "bold"),
                            anchor=anchor, justify=tk.CENTER, relief=tk.SOLID, borderwidth=1,
                            padx=6, pady=3, bg="#e8e8e8", cursor="hand2")
            hlbl.grid(row=0, column=c, sticky="nsew", padx=1, pady=1)
            hlbl.bind("<Button-1>", lambda e, ci=c: self._paste_mfg_sort_by(ci))
        for r, rowdata in enumerate(grid_rows, start=1):
            label = ("[bound] " + rowdata["name"]) if rowdata.get("bound") else rowdata["name"]
            name_bg = "#fff2cc" if rowdata.get("bound") else None
            tk.Label(self.paste_mfg_grid, text=label, anchor=tk.W, bg=name_bg,
                     relief=tk.SOLID, borderwidth=1, padx=6, pady=3, width=34).grid(
                row=r, column=0, sticky="nsew", padx=1, pady=1)
            cells = rowdata["cells"]
            for ci, (header, key, is_pct) in enumerate(self._paste_mfg_col_specs):
                if key is None:
                    continue
                cell = cells.get(key, {})
                value = cell.get("return_pct") if is_pct else cell.get("profit")
                self._paste_mfg_cell(r, ci, value, is_pct=is_pct, detail=cell.get("detail"))

    @staticmethod
    def _paste_mfg_detail_lines(d):
        """Build the text breakdown lines for a manufacturing scenario detail dict."""
        def f(x):
            return f"{x:,.2f}"

        mfg_label = "Manufacturing (bound override)" if d.get("bound") else \
            f"Manufacturing (system cost {d['system_cost_percent']:.2f}% of EIV)"
        lines = []
        lines.append(f"{d['product']}  —  ME {d['me']}, 1 run,  output: {d['out_mode']}")
        lines.append("All prices are Jita; transport each way = 1,000 ISK/m3 + 0.6% of value.")
        lines.append("")
        lines.append("COST")
        lines.append("-" * 60)
        for m in d["materials"]:
            lines.append(
                f"  {m['quantity']:>10,d} x {m['materialName']} @ {f(m['unit_price'])} = {f(m['total_cost'])}"
            )
        lines.append("-" * 60)
        lines.append(f"  Materials subtotal{'':>21}= {f(d['total_input_cost'])}")
        lines.append(f"  + {mfg_label} = {f(d['system_cost'])}")
        lines.append(f"  + Transport of inputs ({d['input_volume']:,.1f} m3 + 0.6% of {f(d['total_input_cost'])}) = {f(d['input_transport'])}")
        lines.append(f"  = COST{'':>32}= {f(d['cost'])}")
        lines.append("")
        lines.append("SALE")
        lines.append("-" * 60)
        lines.append(f"  Sell {d['out_qty']:,.0f} @ {f(d['gross_unit'])} (Jita)   = {f(d['gross_revenue'])}")
        lines.append(f"  - Selling cost (broker/tax){'':>7}= {f(d['selling_cost'])}")
        lines.append(f"  - Transport ({d['output_volume']:,.1f} m3 + 0.6% of {f(d['gross_revenue'])}) = {f(d['output_transport'])}")
        lines.append(f"  = LANDED SALE{'':>19}= {f(d['landed_sale'])}")
        lines.append("")
        lines.append("PROFIT")
        lines.append("-" * 60)
        lines.append(f"  Landed sale - cost = {f(d['landed_sale'])} - {f(d['cost'])} = {f(d['profit'])}")
        if d["return_pct"] is not None:
            lines.append(f"  Return = profit / cost = {d['return_pct']:,.1f}%")
        return lines

    def _paste_mfg_show_detail(self, d):
        """Popup with the full cost/sale/profit breakdown for one blueprint scenario,
        plus an override box + Recalculate/Bind controls for the manufacturing cost."""
        if not d:
            return
        win = tk.Toplevel(self.root)
        win.title(f"{d['product']} — ME{d['me']}, 1 run")
        win.geometry("660x600")
        txt = scrolledtext.ScrolledText(win, wrap=tk.WORD, font=("Consolas", 9))
        txt.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        def render():
            txt.configure(state=tk.NORMAL)
            txt.delete(1.0, tk.END)
            txt.insert(tk.END, "\n".join(self._paste_mfg_detail_lines(d)))
            txt.configure(state=tk.DISABLED)

        ctrl = ttk.Frame(win)
        ctrl.pack(fill=tk.X, padx=8, pady=(0, 4))
        ttk.Label(ctrl, text="Manufacturing cost (ISK):").pack(side=tk.LEFT)
        cost_var = tk.StringVar(value=f"{d['system_cost']:,.0f}")
        ttk.Entry(ctrl, textvariable=cost_var, width=16).pack(side=tk.LEFT, padx=4)

        def apply_override(value):
            d["system_cost"] = value
            d["bound"] = True
            d["cost"] = d["total_input_cost"] + value + d["input_transport"]
            d["profit"] = d["landed_sale"] - d["cost"]
            d["return_pct"] = (d["profit"] / d["cost"] * 100.0) if d["cost"] > 0 else None

        def parse_cost():
            try:
                return float(cost_var.get().replace(",", "").strip())
            except (ValueError, TypeError):
                messagebox.showerror("Invalid value", "Enter a numeric manufacturing cost (ISK).", parent=win)
                return None

        def recalc():
            v = parse_cost()
            if v is None:
                return
            apply_override(v)
            render()

        def bind_cost():
            v = parse_cost()
            if v is None:
                return
            if not d.get("product_type_id"):
                messagebox.showerror("Cannot bind", "This blueprint has no product type id to bind to.", parent=win)
                return
            apply_override(v)
            render()
            self._paste_mfg_set_binding(d["product_type_id"], d["product"], v)
            self.status_var.set(f"Bound manufacturing cost {v:,.0f} ISK to {d['product']}.")
            self.run_paste_compare()

        def clear_binding():
            if not d.get("product_type_id"):
                return
            self._paste_mfg_set_binding(d["product_type_id"], d["product"], None)
            self.status_var.set(f"Cleared bound manufacturing cost for {d['product']}.")
            self.run_paste_compare()
            win.destroy()

        ttk.Button(ctrl, text="Recalculate", command=recalc).pack(side=tk.LEFT, padx=2)
        ttk.Button(ctrl, text="Bind to blueprint", command=bind_cost).pack(side=tk.LEFT, padx=2)
        ttk.Button(ctrl, text="Clear binding", command=clear_binding).pack(side=tk.LEFT, padx=2)

        render()
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=(0, 8))

    
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
        cols = ("Blueprint", "Tech", "T1 profit/run (sell)", "T1 profit (imm)", "T1 profit %", "T2 product", "T2 decryptor", "T2 profit ISK/run", "T2 profit %", "Notes")
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
        cols = ("Blueprint", "Tech", "T1 profit/run (sell)", "T1 profit (imm)", "T1 profit %", "T2 product", "T2 decryptor", "T2 profit ISK/run", "T2 profit %", "Notes")
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
            if col_index in (2, 3, 4, 7, 8):  # numeric
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
        self.planning_tree.insert("", tk.END, values=("", "Analyzing...", "", "", "", "", "", "", "", ""))
        self.root.update()

        def do_analysis():
            try:
                result = self._planning_analyze_blueprints(lines)
                self.root.after(0, lambda: self._planning_apply_results(result))
            except Exception as e:
                err = str(e)
                self.root.after(0, lambda msg=err: self._planning_apply_error(msg))

        threading.Thread(target=do_analysis, daemon=True).start()

    def _planning_apply_error(self, err_msg):
        """Apply error state to planning UI."""
        for item in self.planning_tree.get_children():
            self.planning_tree.delete(item)
        self.planning_tree.insert("", tk.END, values=("", f"Error: {err_msg}", "", "", "", "", "", "", "", ""))
        self._planning_row_data = []
        self.planning_status_var.set("Error occurred.")

    def _planning_apply_results(self, rows):
        """Populate planning tree and store row data from analysis results."""
        for item in self.planning_tree.get_children():
            self.planning_tree.delete(item)
        self._planning_row_data = rows
        for r in rows:
            t1_sell_isk = f"{r.get('t1_profit_sell_isk') or 0:,.0f}" if r.get('t1_profit_sell_isk') is not None else "—"
            t1_imm_isk = f"{r.get('t1_profit_imm_isk') or 0:,.0f}" if r.get('t1_profit_imm_isk') is not None else "—"
            t1_pct = f"{r.get('t1_profit_pct') or 0:.1f}%" if r.get('t1_profit_pct') is not None else "—"
            t2_isk = f"{r.get('t2_profit_isk') or 0:,.0f}" if r.get('t2_profit_isk') is not None else "—"
            t2_pct = f"{r.get('t2_profit_pct') or 0:.1f}%" if r.get('t2_profit_pct') is not None else "—"
            self.planning_tree.insert("", tk.END, values=(
                r.get("t1_name") or "—",
                r.get("tech") or "—",
                t1_sell_isk,
                t1_imm_isk,
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
                row_data = {"t1_name": name, "tech": "—", "t1_profit_sell_isk": None, "t1_profit_imm_isk": None, "t1_profit_pct": None,
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
                row_data["t1_profit_sell_isk"] = t1_result["profit"]
                total_cost = t1_result["total_input_cost"] + t1_result["system_cost"]
                row_data["t1_profit_pct"] = (t1_result["profit"] / total_cost * 100.0) if total_cost and total_cost > 0 else 0.0
                t1_result_imm = calculate_blueprint_profitability(
                    blueprint_name_or_product=product_name,
                    input_price_type=input_price,
                    output_price_type="sell_immediate",
                    system_cost_percent=system_cost_pct,
                    material_efficiency=0,
                    number_of_runs=1,
                    region_id=region_id,
                    db_file=DATABASE_FILE,
                )
                if "error" not in t1_result_imm:
                    row_data["t1_profit_imm_isk"] = t1_result_imm.get("profit")
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
                err = f"Error: {e}"
                self.root.after(0, lambda msg=err: self._market_patterns_apply_result(msg, 1))

        threading.Thread(target=worker, daemon=True).start()

    def _market_patterns_apply_result(self, text: str, returncode: int):
        self.market_patterns_text.delete(1.0, tk.END)
        self.market_patterns_text.insert(tk.END, text or "(no output)")
        if returncode == 0:
            self.market_patterns_status_var.set("Market patterns analysis completed.")
        else:
            self.market_patterns_status_var.set("Market patterns analysis failed (see output).")

    def create_arbitrage_tab(self):
        """Jita -> C-N4OD arbitrage scanner (uses arbitrage_finder.py + EVE SSO)."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Arbitrage C-N4OD")

        info = ttk.LabelFrame(frame, text="Jita -> C-N4OD resell opportunities", padding=10)
        info.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(
            info,
            text="Ranks high-volume Jita items and compares Jita sell + landed cost "
                 "(jita*1.1 + 1000 ISK/m3) against the cheapest C-N4OD sell order.\n"
                 "Margin = (C-N4OD sell - landed) / Jita sell. Requires an EVE SSO character "
                 "with docking access to the C-N4OD market structure.",
            justify=tk.LEFT, wraplength=900,
        ).pack(anchor=tk.W)

        ctrl = ttk.Frame(frame)
        ctrl.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(ctrl, text="Top items:").pack(side=tk.LEFT, padx=(5, 2))
        self.arb_top_var = tk.StringVar(value="500")
        ttk.Entry(ctrl, textvariable=self.arb_top_var, width=6).pack(side=tk.LEFT)
        ttk.Label(ctrl, text="Rank by:").pack(side=tk.LEFT, padx=(10, 2))
        self.arb_rank_var = tk.StringVar(value="isk")
        ttk.Combobox(ctrl, textvariable=self.arb_rank_var, values=["isk", "units"],
                     width=7, state="readonly").pack(side=tk.LEFT)
        ttk.Label(ctrl, text="Min margin %:").pack(side=tk.LEFT, padx=(10, 2))
        self.arb_min_margin_var = tk.StringVar(value="0")
        ttk.Entry(ctrl, textvariable=self.arb_min_margin_var, width=6).pack(side=tk.LEFT)
        ttk.Label(ctrl, text="History days:").pack(side=tk.LEFT, padx=(10, 2))
        self.arb_days_var = tk.StringVar(value="365")
        ttk.Entry(ctrl, textvariable=self.arb_days_var, width=6).pack(side=tk.LEFT)

        btn_row = ttk.Frame(frame)
        btn_row.pack(fill=tk.X, padx=10, pady=5)
        self.arb_run_btn = ttk.Button(btn_row, text="Run scan", command=self._run_arbitrage_scan)
        self.arb_run_btn.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Discover structures",
                   command=lambda: self._run_arbitrage_scan(discover=True)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Open CSV", command=self._arbitrage_open_csv).pack(side=tk.LEFT, padx=5)
        self.arb_status_var = tk.StringVar(value="Ready.")
        ttk.Label(btn_row, textvariable=self.arb_status_var).pack(side=tk.LEFT, padx=10)

        tree_frame = ttk.LabelFrame(frame, text="Opportunities", padding=5)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        cols = ("name", "jita_sell", "cn4od_sell", "landed_price",
                "profit_per_unit", "margin", "volume_m3", "jita_avg_isk_day")
        headers = {
            "name": "Item", "jita_sell": "Jita sell", "cn4od_sell": "C-N4OD sell",
            "landed_price": "Landed", "profit_per_unit": "Profit/unit", "margin": "Margin %",
            "volume_m3": "m3", "jita_avg_isk_day": "Jita ISK/day",
        }
        self.arb_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=16)
        for c in cols:
            self.arb_tree.heading(c, text=headers[c])
            self.arb_tree.column(c, width=110, anchor=(tk.W if c == "name" else tk.E))
        self.arb_tree.column("name", width=250, anchor=tk.W)
        self.arb_tree.bind("<Double-1>", self._arbitrage_copy_name)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.arb_tree.yview)
        self.arb_tree.configure(yscrollcommand=vsb.set)
        self.arb_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        log_frame = ttk.LabelFrame(frame, text="Log", padding=5)
        log_frame.pack(fill=tk.X, padx=10, pady=5)
        self.arb_log = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=6)
        self.arb_log.pack(fill=tk.BOTH, expand=True)

    def _run_arbitrage_scan(self, discover=False):
        """Run arbitrage_finder.py in a background thread; load its CSV into the tree."""
        try:
            top = int(self.arb_top_var.get())
            days = int(self.arb_days_var.get())
            rank = self.arb_rank_var.get()
            min_margin = float(self.arb_min_margin_var.get()) / 100.0
        except ValueError:
            messagebox.showwarning("Arbitrage", "Enter valid numbers for Top / days / min margin.")
            return

        out_path = str(Path(__file__).resolve().parent / "arbitrage_cn4od.csv")
        self._arb_out_path = out_path
        args = [sys.executable, "arbitrage_finder.py"]
        if discover:
            args.append("--discover")
        else:
            args += ["--top", str(top), "--rank", rank, "--days", str(days),
                     "--min-margin", str(min_margin), "--out", out_path]

        self.arb_status_var.set("Discovering structures..." if discover else "Running scan (this can take ~30s)...")
        self.arb_run_btn.config(state=tk.DISABLED)
        self.arb_log.delete(1.0, tk.END)
        self.arb_log.insert(tk.END, " ".join(args) + "\n\n")
        self.root.update_idletasks()

        def worker():
            try:
                result = subprocess.run(
                    args, cwd=Path(__file__).resolve().parent,
                    capture_output=True, text=True, timeout=900,
                )
                text = result.stdout or ""
                if result.stderr:
                    text += "\n--- stderr ---\n" + result.stderr
                self.root.after(0, lambda: self._arbitrage_apply_result(text, result.returncode, discover))
            except Exception as e:
                self.root.after(0, lambda msg=f"Error: {e}": self._arbitrage_apply_result(msg, 1, discover))

        threading.Thread(target=worker, daemon=True).start()

    def _arbitrage_apply_result(self, text, returncode, discover):
        self.arb_run_btn.config(state=tk.NORMAL)
        self.arb_log.delete(1.0, tk.END)
        self.arb_log.insert(tk.END, text or "(no output)")
        if returncode != 0:
            self.arb_status_var.set("Failed (see log).")
            return
        if discover:
            self.arb_status_var.set("Discovery complete (see log).")
            return
        n = self._arbitrage_load_csv(getattr(self, "_arb_out_path", None))
        self.arb_status_var.set(f"Done. {n} opportunities loaded.")

    def _arbitrage_load_csv(self, path):
        import csv
        for i in self.arb_tree.get_children():
            self.arb_tree.delete(i)
        if not path or not Path(path).exists():
            return 0
        count = 0
        with open(path, encoding="utf-8", newline="") as f:
            for row in csv.DictReader(f):
                try:
                    self.arb_tree.insert("", tk.END, values=(
                        row["name"],
                        f'{float(row["jita_sell"]):,.0f}',
                        f'{float(row["cn4od_sell"]):,.0f}',
                        f'{float(row["landed_price"]):,.0f}',
                        f'{float(row["profit_per_unit"]):,.0f}',
                        f'{float(row["margin"]) * 100:.1f}',
                        row["volume_m3"],
                        f'{float(row["jita_avg_isk_day"]):,.0f}',
                    ))
                    count += 1
                except Exception:
                    continue
        return count

    def _arbitrage_copy_name(self, event):
        """Double-click a row to copy the item name to the clipboard."""
        row_id = self.arb_tree.identify_row(event.y)
        if not row_id:
            return
        values = self.arb_tree.item(row_id, "values")
        if not values:
            return
        name = values[0]
        self.root.clipboard_clear()
        self.root.clipboard_append(name)
        self.arb_status_var.set(f"Copied: {name}")

    def _arbitrage_open_csv(self):
        import os
        path = getattr(self, "_arb_out_path",
                       str(Path(__file__).resolve().parent / "arbitrage_cn4od.csv"))
        if Path(path).exists():
            try:
                os.startfile(path)
            except Exception as e:
                messagebox.showerror("Arbitrage", str(e))
        else:
            messagebox.showinfo("Arbitrage", "No CSV yet. Run a scan first.")

    def create_remap_planner_tab(self):
        """Skill-queue attribute (neural remap) planner across all SSO characters."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Remap Planner")

        info = ttk.LabelFrame(frame, text="Skill-queue attribute totals", padding=10)
        info.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(
            info,
            text="Reads each character's training queue. For every queued skill, its SP is added to "
                 "the PRIMARY attribute total and SP/2 to the SECONDARY attribute total.\n"
                 "The highest totals indicate the best neural remap for the queued training.",
            justify=tk.LEFT, wraplength=900,
        ).pack(anchor=tk.W)

        ctrl = ttk.Frame(frame)
        ctrl.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(ctrl, text="SP mode:").pack(side=tk.LEFT, padx=(5, 2))
        self.remap_sp_mode_var = tk.StringVar(value="remaining")
        ttk.Radiobutton(ctrl, text="Remaining (SP left to train)", value="remaining",
                        variable=self.remap_sp_mode_var).pack(side=tk.LEFT, padx=4)
        ttk.Radiobutton(ctrl, text="Full level SP", value="level",
                        variable=self.remap_sp_mode_var).pack(side=tk.LEFT, padx=4)
        self.remap_compute_btn = ttk.Button(ctrl, text="Compute", command=self._run_remap_planner)
        self.remap_compute_btn.pack(side=tk.LEFT, padx=10)
        self.remap_status_var = tk.StringVar(value="Ready.")
        ttk.Label(ctrl, textvariable=self.remap_status_var).pack(side=tk.LEFT, padx=10)

        attrs = ("Perception", "Memory", "Willpower", "Intelligence", "Charisma")
        self._remap_attrs = attrs
        # Column widths for the monospaced grid (per-cell coloring needs a Text widget)
        self._remap_name_w = 22
        self._remap_col_w = 15
        legend = ttk.Frame(frame)
        legend.pack(fill=tk.X, padx=12)
        tk.Label(legend, text="  largest  ", bg="#5fbf5f").pack(side=tk.LEFT)
        tk.Label(legend, text="  2nd largest  ", bg="#bfe8bf").pack(side=tk.LEFT, padx=(6, 0))
        grid_frame = ttk.LabelFrame(frame, text="Attribute SP totals (per character: largest in green, 2nd in light green)", padding=5)
        grid_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.remap_grid = tk.Text(grid_frame, wrap=tk.NONE, height=14,
                                  font=("Consolas", 10), state=tk.DISABLED, cursor="arrow")
        vsb = ttk.Scrollbar(grid_frame, orient="vertical", command=self.remap_grid.yview)
        self.remap_grid.configure(yscrollcommand=vsb.set)
        self.remap_grid.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.remap_grid.tag_configure("green", background="#5fbf5f")
        self.remap_grid.tag_configure("lightgreen", background="#bfe8bf")
        self.remap_grid.tag_configure("header", font=("Consolas", 10, "bold"))
        self.remap_grid.tag_configure("total", font=("Consolas", 10, "bold"), background="#e8e8e8")

        self.remap_priority_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.remap_priority_var,
                  font=("", 10, "bold")).pack(anchor=tk.W, padx=12, pady=(0, 5))

        log_frame = ttk.LabelFrame(frame, text="Notes", padding=5)
        log_frame.pack(fill=tk.X, padx=10, pady=5)
        self.remap_log = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=5)
        self.remap_log.pack(fill=tk.BOTH, expand=True)

    def _run_remap_planner(self):
        """Compute attribute totals in a background thread (first run fetches skill attrs via ESI)."""
        cid, secret = self._load_sso_credentials()
        if not cid or not secret:
            messagebox.showwarning("Remap Planner", "Set EVE SSO Client ID / Secret first (EVE SSO Sync tab).")
            return
        sp_mode = self.remap_sp_mode_var.get()
        self.remap_compute_btn.config(state=tk.DISABLED)
        self.remap_status_var.set("Computing (first run may take ~60s to cache skill attributes)...")
        self.remap_log.delete(1.0, tk.END)
        self.root.update_idletasks()

        def worker():
            try:
                import skill_queue_attributes as sqa
                conn = sqlite3.connect(DATABASE_FILE, timeout=60)
                conn.execute("PRAGMA busy_timeout=60000")
                try:
                    result = sqa.compute_attribute_totals(conn, cid, secret, sp_mode)
                finally:
                    conn.close()
                self.root.after(0, lambda: self._remap_apply_result(result))
            except Exception as e:
                self.root.after(0, lambda msg=f"Error: {e}": self._remap_apply_error(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _remap_apply_error(self, msg):
        self.remap_compute_btn.config(state=tk.NORMAL)
        self.remap_status_var.set("Failed (see notes).")
        self.remap_log.delete(1.0, tk.END)
        self.remap_log.insert(tk.END, msg)

    def _remap_grid_row(self, name, values, line_idx, *, row_tag=None, highlight=True):
        """Insert one grid row; color the largest cell green and 2nd largest light green."""
        nw, cw = self._remap_name_w, self._remap_col_w
        text = name[:nw - 1].ljust(nw) + "".join(f"{v:,.0f}".rjust(cw) for v in values)
        self.remap_grid.insert(tk.END, text + "\n", (row_tag,) if row_tag else ())

        if highlight and any(values):
            order = sorted(range(len(values)), key=lambda i: values[i], reverse=True)
            top = [i for i in order if values[i] > 0][:2]
            for rank, ci in enumerate(top):
                start_col = nw + ci * cw
                end_col = start_col + cw
                tag = "green" if rank == 0 else "lightgreen"
                self.remap_grid.tag_add(tag, f"{line_idx}.{start_col}", f"{line_idx}.{end_col}")

    def _remap_apply_result(self, result):
        self.remap_compute_btn.config(state=tk.NORMAL)
        attrs = self._remap_attrs
        nw, cw = self._remap_name_w, self._remap_col_w

        self.remap_grid.config(state=tk.NORMAL)
        self.remap_grid.delete(1.0, tk.END)

        header = "Character".ljust(nw) + "".join(a[:cw - 1].rjust(cw) for a in attrs)
        self.remap_grid.insert(tk.END, header + "\n", ("header",))
        self.remap_grid.insert(tk.END, "-" * len(header) + "\n")

        line = 3  # header=1, separator=2, first data row=3
        for name, ct in result.get("per_char", {}).items():
            vals = [ct.get(a, 0) for a in attrs]
            self._remap_grid_row(name, vals, line)
            line += 1

        self.remap_grid.insert(tk.END, "-" * len(header) + "\n")
        line += 1
        totals = result.get("totals", {})
        self._remap_grid_row("TOTAL", [totals.get(a, 0) for a in attrs], line,
                             row_tag="total", highlight=False)
        self.remap_grid.config(state=tk.DISABLED)

        if totals:
            ranked = sorted(attrs, key=lambda a: totals.get(a, 0), reverse=True)
            self.remap_priority_var.set("Suggested fleet remap priority:  " + "  >  ".join(ranked))
        self.remap_status_var.set(f"Done ({result.get('sp_mode')} SP).")

        notes = []
        errs = result.get("errors", {})
        if errs:
            notes.append("Characters skipped:")
            for name, msg in errs.items():
                notes.append(f"  {name}: {msg}")
        else:
            notes.append("All characters processed successfully.")
        self.remap_log.delete(1.0, tk.END)
        self.remap_log.insert(tk.END, "\n".join(notes))

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

    def _load_sso_credentials(self):
        """Return (client_id, client_secret) preferring env vars, then SSO_CREDENTIALS_FILE."""
        import os
        cid = os.environ.get("EVE_SSO_CLIENT_ID", "").strip()
        secret = os.environ.get("EVE_SSO_CLIENT_SECRET", "").strip()
        if not cid or not secret:
            try:
                if SSO_CREDENTIALS_FILE.exists():
                    data = json.loads(SSO_CREDENTIALS_FILE.read_text(encoding="utf-8"))
                    cid = cid or str(data.get("client_id", "")).strip()
                    secret = secret or str(data.get("client_secret", "")).strip()
            except Exception:
                pass
        return cid, secret

    def _save_sso_credentials(self, *_args):
        """Persist Client ID/Secret to SSO_CREDENTIALS_FILE (gitignored)."""
        try:
            cid = self.sso_client_id_var.get().strip()
            secret = self.sso_client_secret_var.get().strip()
            data = {
                "client_id": cid,
                "client_secret": secret,
                "callback_url": "http://localhost:8765/callback/",
            }
            SSO_CREDENTIALS_FILE.write_text(
                json.dumps(data, indent=2), encoding="utf-8"
            )
        except Exception:
            pass

    def create_sso_sync_tab(self):
        """EVE SSO sync tab: manage multiple linked characters and sync their data."""
        from eve_sso_sync import DEFAULT_SCOPES
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="EVE SSO Sync")
        info = ttk.LabelFrame(frame, text="Instructions", padding=10)
        info.pack(fill=tk.X, padx=10, pady=10)
        scope_lines = ", ".join(DEFAULT_SCOPES.split())
        ttk.Label(
            info,
            text=(
                "Create an SSO application at https://developers.eveonline.com/ with callback URL "
                "http://localhost:8765/callback/.\n"
                "Enable the scopes listed below on the dev portal (must match exactly).\n"
                "To link multiple characters: click 'Add character (Login)', log in, repeat for each. "
                "If EVE auto-redirects without prompting, click 'Log Out' on the EVE login page first.\n"
                f"Scopes ({len(DEFAULT_SCOPES.split())}): {scope_lines}"
            ),
            justify=tk.LEFT, wraplength=1100
        ).pack(anchor=tk.W)

        creds = ttk.LabelFrame(frame, text="SSO credentials (saved locally to eve_sso_credentials.json - gitignored)", padding=10)
        creds.pack(fill=tk.X, padx=10, pady=5)
        cid_default, secret_default = self._load_sso_credentials()
        ttk.Label(creds, text="Client ID:").pack(side=tk.LEFT, padx=5)
        self.sso_client_id_var = tk.StringVar(value=cid_default)
        ttk.Entry(creds, textvariable=self.sso_client_id_var, width=36).pack(side=tk.LEFT, padx=5)
        ttk.Label(creds, text="Client Secret:").pack(side=tk.LEFT, padx=5)
        self.sso_client_secret_var = tk.StringVar(value=secret_default)
        ttk.Entry(creds, textvariable=self.sso_client_secret_var, width=44, show="*").pack(side=tk.LEFT, padx=5)
        try:
            self.sso_client_id_var.trace_add("write", self._save_sso_credentials)
            self.sso_client_secret_var.trace_add("write", self._save_sso_credentials)
        except Exception:
            pass

        chars_frame = ttk.LabelFrame(frame, text="Linked characters", padding=10)
        chars_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=5)
        btn_row = ttk.Frame(chars_frame)
        btn_row.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(btn_row, text="Add character (Login)", command=self.sso_login).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="Sync selected", command=self.sso_sync_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="Sync all", command=self.sso_sync_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="Remove selected", command=self.sso_remove_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="Refresh list", command=self._refresh_sso_chars_tree).pack(side=tk.LEFT, padx=2)

        cols = ("name", "char_id", "corp", "token_in", "last_sync", "updated")
        self.sso_chars_tree = ttk.Treeview(chars_frame, columns=cols, show="headings", height=8, selectmode="extended")
        headings = {
            "name": ("Character", 220),
            "char_id": ("Character ID", 110),
            "corp": ("Corporation", 220),
            "token_in": ("Token expires in", 130),
            "last_sync": ("Last synced", 160),
            "updated": ("Updated", 160),
        }
        for c, (h, w) in headings.items():
            self.sso_chars_tree.heading(c, text=h)
            self.sso_chars_tree.column(c, width=w, anchor=tk.W, stretch=(c in ("name", "corp")))
        ysb = ttk.Scrollbar(chars_frame, orient="vertical", command=self.sso_chars_tree.yview)
        self.sso_chars_tree.configure(yscrollcommand=ysb.set)
        self.sso_chars_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)

        self.sso_status_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.sso_status_var).pack(anchor=tk.W, padx=10, pady=2)
        log_frame = ttk.LabelFrame(frame, text="Log", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.sso_log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10, width=80)
        self.sso_log_text.pack(fill=tk.BOTH, expand=True)
        self._refresh_sso_chars_tree()

    def _format_token_remaining(self, expires_at) -> str:
        if not expires_at:
            return "—"
        try:
            remaining = float(expires_at) - time.time()
        except (TypeError, ValueError):
            return "—"
        if remaining <= 0:
            return "expired (refresh on use)"
        if remaining < 60:
            return f"{int(remaining)}s"
        if remaining < 3600:
            return f"{int(remaining // 60)}m {int(remaining % 60)}s"
        h = int(remaining // 3600)
        m = int((remaining % 3600) // 60)
        return f"{h}h {m}m"

    def _refresh_sso_chars_tree(self):
        """Reload the linked-characters Treeview from the database."""
        if not hasattr(self, "sso_chars_tree"):
            return
        from eve_sso_sync import list_sso_characters, ensure_sso_tables
        try:
            for iid in self.sso_chars_tree.get_children():
                self.sso_chars_tree.delete(iid)
            if not Path(DATABASE_FILE).exists():
                self.sso_status_var.set("Database not found.")
                return
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                ensure_sso_tables(conn)
                rows = list_sso_characters(conn)
            finally:
                conn.close()
            for r in rows:
                self.sso_chars_tree.insert(
                    "", tk.END,
                    iid=str(r["character_id"]),
                    values=(
                        r["character_name"] or "?",
                        r["character_id"],
                        r["corporation_name"] or (str(r["corporation_id"]) if r["corporation_id"] else "—"),
                        self._format_token_remaining(r["access_token_expires_at"]),
                        r["last_synced_at"] or "—",
                        r["updated_at"] or "—",
                    ),
                )
            n = len(rows)
            self.sso_status_var.set(f"{n} linked character{'s' if n != 1 else ''}.")
        except Exception as e:
            self._sso_log(f"Refresh list error: {e}")
            self.sso_status_var.set("Error loading characters.")

    def _selected_sso_character_ids(self) -> list[int]:
        if not hasattr(self, "sso_chars_tree"):
            return []
        out = []
        for iid in self.sso_chars_tree.selection():
            try:
                out.append(int(iid))
            except ValueError:
                pass
        return out

    def _sso_sync_one(self, character_id: int, cid: str, secret: str) -> dict:
        from eve_sso_sync import run_full_sync, ensure_sso_tables
        conn = sqlite3.connect(DATABASE_FILE)
        try:
            ensure_sso_tables(conn)
            return run_full_sync(conn, character_id, cid, secret)
        finally:
            conn.close()

    def sso_sync_selected(self):
        """Sync wallet/journal/industry jobs for currently selected characters."""
        cid = self.sso_client_id_var.get().strip()
        secret = self.sso_client_secret_var.get().strip()
        if not cid or not secret:
            messagebox.showwarning("SSO", "Enter Client ID and Client Secret first.")
            return
        if not Path(DATABASE_FILE).exists():
            messagebox.showwarning("SSO", "Database not found. Create it first (e.g. build_database).")
            return
        ids = self._selected_sso_character_ids()
        if not ids:
            messagebox.showinfo("SSO", "Select one or more characters in the list first.")
            return
        self._sso_run_bulk_sync(ids, cid, secret)

    def sso_sync_all(self):
        """Sync every linked character."""
        from eve_sso_sync import list_sso_characters, ensure_sso_tables
        cid = self.sso_client_id_var.get().strip()
        secret = self.sso_client_secret_var.get().strip()
        if not cid or not secret:
            messagebox.showwarning("SSO", "Enter Client ID and Client Secret first.")
            return
        if not Path(DATABASE_FILE).exists():
            messagebox.showwarning("SSO", "Database not found. Create it first (e.g. build_database).")
            return
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                ensure_sso_tables(conn)
                ids = [r["character_id"] for r in list_sso_characters(conn)]
            finally:
                conn.close()
        except Exception as e:
            messagebox.showerror("SSO", f"Could not read characters: {e}")
            return
        if not ids:
            messagebox.showinfo("SSO", "No linked characters yet. Click 'Add character (Login)' first.")
            return
        self._sso_run_bulk_sync(ids, cid, secret)

    def _sso_run_bulk_sync(self, character_ids: list[int], cid: str, secret: str):
        """Run sync for a list of character IDs in a background thread."""
        self.sso_status_var.set(f"Syncing {len(character_ids)} character(s)...")
        self._sso_log(f"Sync starting for {len(character_ids)} character(s)...")
        def run():
            ok = 0
            errs = 0
            for char_id in character_ids:
                try:
                    result = self._sso_sync_one(char_id, cid, secret)
                    if result.get("error") and not (
                        result.get("tx") or result.get("journal")
                        or result.get("jobs") or result.get("corp_jobs")
                        or result.get("corp_wallet_tx")
                    ):
                        errs += 1
                        self._sso_log(f"  [{char_id}] FAILED: {result['error']}")
                    else:
                        ok += 1
                        line = (
                            f"  [{char_id}] tx={result.get('tx', 0)}, "
                            f"journal={result.get('journal', 0)}, "
                            f"char_jobs={result.get('jobs', 0)}, "
                            f"corp_jobs={result.get('corp_jobs', 0)}, "
                            f"corp_wallet_tx={result.get('corp_wallet_tx', 0)}"
                        )
                        if result.get("corp_jobs_note"):
                            line += f" — {result['corp_jobs_note']}"
                        if result.get("corp_wallet_note"):
                            line += f" — {result['corp_wallet_note']}"
                        if result.get("error"):
                            line += f" (note: {result['error']})"
                        self._sso_log(line)
                except Exception as e:
                    errs += 1
                    err_text = str(e)
                    self._sso_log(f"  [{char_id}] EXCEPTION: {err_text}")
            self.sso_status_var.set(f"Sync done: {ok} ok, {errs} failed.")
            self._sso_log(f"Sync finished: {ok} ok, {errs} failed.")
            self.root.after(0, self._refresh_sso_chars_tree)
        threading.Thread(target=run, daemon=True).start()

    def sso_remove_selected(self):
        """Remove selected characters and their synced ESI rows from the DB."""
        from eve_sso_sync import delete_sso_character, ensure_sso_tables
        ids = self._selected_sso_character_ids()
        if not ids:
            messagebox.showinfo("SSO", "Select one or more characters in the list first.")
            return
        if not messagebox.askyesno(
            "Remove characters",
            f"Remove {len(ids)} character(s) and all their synced wallet/journal/industry rows from the local DB?\n"
            "(EVE permissions on the dev portal are unchanged.)",
        ):
            return
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                ensure_sso_tables(conn)
                for char_id in ids:
                    delete_sso_character(conn, char_id)
            finally:
                conn.close()
        except Exception as e:
            messagebox.showerror("SSO", f"Could not remove: {e}")
            return
        self._sso_log(f"Removed {len(ids)} character(s).")
        self._refresh_sso_chars_tree()

    def create_profitability_tab(self):
        """Production P/L per industry job, computed FIFO from ESI wallet + industry data."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Profitability")

        info = ttk.LabelFrame(frame, text="Production P/L (FIFO cost basis)", padding=10)
        info.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(
            info,
            text=(
                "Per-job profit/loss for completed manufacturing & reactions, computed by replaying "
                "your ESI wallet transactions and industry jobs chronologically with FIFO matching. "
                "For a corporation entity, completed jobs come from corp industry; buys and sells "
                "are replayed from every SSO-linked character in that corp's personal wallet (and "
                "from any linked character who installed a corp job for that corp, in case "
                "corporation_id on a character row is missing), plus corp wallet market transactions "
                "if synced—one shared timeline, so if manufacturing "
                "is on the corp but the seller lists orders from a character wallet, those sell "
                "transactions still count here as long as that seller is linked in EVE SSO Sync. "
                "Use the character-only row only for personal jobs; use the '(corp)' row when jobs "
                "are corporation jobs even if products are sold from characters. "
                "T2 manufacturing jobs also include expected invention cost per BPC (from the "
                "Decryptor tab / blueprint_datacore_bindings, same as shopping-list research) "
                "and optional production cost per run, amortized by manufacturing runs on the BPC. "
                "When a material has no historical buy lot, we capture a stable fallback price "
                "(market_history.average for the job date if available, else current sell_min, "
                "else buy_max) - stored once per (item, date), so future rebuilds give the same numbers. "
                "Click 'Rebuild ledger' after a fresh SSO sync. Use View to switch between each job "
                "or jobs grouped by product and ISO week. Double-click a row for a detailed breakdown "
                "(weekly groups include all jobs in that week). ⚠ flag = no buy lot AND "
                "no market data either (truly unknown). Use 'Clear price snapshots' to force a re-capture."
            ),
            justify=tk.LEFT, wraplength=1100,
        ).pack(anchor=tk.W)

        ctrl = ttk.Frame(frame)
        ctrl.pack(fill=tk.X, padx=10, pady=4)
        ttk.Label(ctrl, text="Character:").pack(side=tk.LEFT, padx=2)
        self.pl_character_var = tk.StringVar(value="(All)")
        self.pl_character_combo = ttk.Combobox(
            ctrl, textvariable=self.pl_character_var, state="readonly", width=28, values=["(All)"]
        )
        self.pl_character_combo.pack(side=tk.LEFT, padx=2)
        self.pl_character_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_pl_view())
        ttk.Label(ctrl, text="Default ME%:").pack(side=tk.LEFT, padx=(15, 2))
        self.pl_me_var = tk.StringVar(value="10")
        ttk.Spinbox(ctrl, from_=0, to=10, textvariable=self.pl_me_var, width=4, increment=1).pack(side=tk.LEFT, padx=2)
        self.pl_market_fallback_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            ctrl,
            text="Use current sell_min for missing material lots",
            variable=self.pl_market_fallback_var,
        ).pack(side=tk.LEFT, padx=(15, 2))
        ttk.Label(ctrl, text="Search:").pack(side=tk.LEFT, padx=(15, 2))
        self.pl_search_var = tk.StringVar()
        ent = ttk.Entry(ctrl, textvariable=self.pl_search_var, width=22)
        ent.pack(side=tk.LEFT, padx=2)
        ent.bind("<KeyRelease>", lambda _e: self._refresh_pl_view())
        ttk.Label(ctrl, text="View:").pack(side=tk.LEFT, padx=(12, 2))
        self.pl_view_var = tk.StringVar(value="Per job")
        pl_view = ttk.Combobox(
            ctrl, textvariable=self.pl_view_var, state="readonly", width=14,
            values=["Per job", "Per week (grouped)"],
        )
        pl_view.pack(side=tk.LEFT, padx=2)
        pl_view.bind("<<ComboboxSelected>>", lambda _e: self._refresh_pl_view())
        ttk.Button(ctrl, text="Rebuild ledger", command=self.profitability_rebuild).pack(side=tk.LEFT, padx=10)
        ttk.Button(ctrl, text="Refresh view", command=self._refresh_pl_view).pack(side=tk.LEFT, padx=2)
        ttk.Button(ctrl, text="Clear price snapshots", command=self.profitability_clear_snapshots).pack(side=tk.LEFT, padx=2)

        summary = ttk.Frame(frame)
        summary.pack(fill=tk.X, padx=10, pady=2)
        self.pl_summary_var = tk.StringVar(value="")
        ttk.Label(summary, textvariable=self.pl_summary_var, font=("TkDefaultFont", 9, "bold")).pack(anchor=tk.W)

        self.pl_tree_frame = ttk.LabelFrame(frame, text="Per-job P/L", padding=6)
        self.pl_tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        self.pl_row_meta: dict[str, dict] = {}
        cols = (
            "completed", "char", "product", "jobs", "runs", "out_qty", "job_fee",
            "mat_cost", "inv_cost", "total_cost", "unit_cost", "sold", "revenue",
            "realized", "unsold", "unrealized", "total_pl", "flag",
        )
        headings = {
            "completed": ("Completed", 145),
            "char": ("Character", 130),
            "product": ("Product", 200),
            "jobs": ("Jobs", 52),
            "runs": ("Runs", 60),
            "out_qty": ("Output", 80),
            "job_fee": ("Job fee", 90),
            "mat_cost": ("Materials", 110),
            "inv_cost": ("Invention", 95),
            "total_cost": ("Total cost", 110),
            "unit_cost": ("Unit cost", 95),
            "sold": ("Sold", 80),
            "revenue": ("Revenue", 110),
            "realized": ("Realized P/L", 110),
            "unsold": ("Unsold", 70),
            "unrealized": ("Unrealized", 110),
            "total_pl": ("Total est. P/L", 120),
            "flag": ("⚠", 30),
        }
        self.pl_tree = ttk.Treeview(self.pl_tree_frame, columns=cols, show="headings", height=18)
        for c, (h, w) in headings.items():
            self.pl_tree.heading(c, text=h, command=lambda _c=c: self._pl_sort_by(_c))
            anchor = tk.W if c in ("completed", "char", "product") else tk.E
            self.pl_tree.column(c, width=w, anchor=anchor, stretch=(c == "product"))
        ysb = ttk.Scrollbar(self.pl_tree_frame, orient="vertical", command=self.pl_tree.yview)
        xsb = ttk.Scrollbar(self.pl_tree_frame, orient="horizontal", command=self.pl_tree.xview)
        self.pl_tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.pl_tree.tag_configure("loss", foreground="#a02020")
        self.pl_tree.tag_configure("gain", foreground="#106010")
        self.pl_tree.tag_configure("flagged", background="#fff5e0")
        self.pl_tree.bind("<Double-1>", self._pl_show_job_breakdown)
        self.pl_tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        self.pl_tree_frame.rowconfigure(0, weight=1)
        self.pl_tree_frame.columnconfigure(0, weight=1)

        log_frame = ttk.LabelFrame(frame, text="Log", padding=6)
        log_frame.pack(fill=tk.X, padx=10, pady=(2, 8))
        self.pl_log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=5)
        self.pl_log_text.pack(fill=tk.BOTH, expand=True)

        self._pl_sort_state = {"col": "completed", "reverse": True}
        self._refresh_pl_character_choices()
        self._refresh_pl_view()

    def _pl_log(self, msg: str):
        if not hasattr(self, "pl_log_text"):
            return
        self.pl_log_text.insert(tk.END, msg + "\n")
        self.pl_log_text.see(tk.END)
        self.root.update_idletasks()

    def _refresh_pl_character_choices(self):
        if not hasattr(self, "pl_character_combo"):
            return
        opts = ["(All)"]
        try:
            if Path(DATABASE_FILE).exists():
                conn = sqlite3.connect(DATABASE_FILE)
                try:
                    char_rows = conn.execute(
                        "SELECT character_id, character_name FROM sso_character "
                        "ORDER BY character_name COLLATE NOCASE"
                    ).fetchall()
                    corp_rows = conn.execute(
                        """
                        SELECT corporation_id, MAX(corporation_name)
                          FROM sso_character
                         WHERE corporation_id IS NOT NULL
                         GROUP BY corporation_id
                         ORDER BY MAX(corporation_name) COLLATE NOCASE
                        """
                    ).fetchall()
                finally:
                    conn.close()
                for r in char_rows:
                    name = r[1] or f"Char {r[0]}"
                    opts.append(f"char | {name} ({r[0]})")
                for r in corp_rows:
                    name = r[1] or f"Corp {r[0]}"
                    opts.append(f"corp | {name} ({r[0]})")
        except Exception:
            pass
        cur = self.pl_character_var.get()
        self.pl_character_combo["values"] = opts
        if cur not in opts:
            self.pl_character_var.set("(All)")

    def _selected_pl_entity(self) -> tuple[Optional[int], Optional[str]]:
        """Parse the combobox selection into (entity_id, entity_kind)."""
        sel = self.pl_character_var.get()
        if not sel or sel == "(All)":
            return None, None
        kind = "character"
        if sel.startswith("corp | "):
            kind = "corporation"
        if "(" in sel and sel.endswith(")"):
            try:
                return int(sel.rsplit("(", 1)[1].rstrip(")")), kind
            except ValueError:
                return None, None
        return None, None

    def _selected_pl_character_id(self) -> Optional[int]:
        eid, kind = self._selected_pl_entity()
        if kind == "character":
            return eid
        return None

    def profitability_clear_snapshots(self):
        """Wipe the captured market-fallback price snapshots."""
        from profitability_tracking import clear_price_snapshots
        if not Path(DATABASE_FILE).exists():
            messagebox.showwarning("Profitability", "Database not found.")
            return
        if not messagebox.askyesno(
            "Clear price snapshots",
            "This deletes all stored fallback prices. The next 'Rebuild ledger' "
            "will re-resolve them (preferring market history at the job date). Continue?",
        ):
            return
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                n = clear_price_snapshots(conn)
            finally:
                conn.close()
        except Exception as e:
            self._pl_log(f"Clear snapshots failed: {e}")
            messagebox.showerror("Profitability", str(e))
            return
        self._pl_log(f"Cleared {n} price snapshot(s).")

    def profitability_rebuild(self):
        """Rebuild the FIFO ledger for the selected entity, or all linked entities."""
        from profitability_tracking import rebuild_ledger, ensure_profitability_tables
        from eve_sso_sync import ensure_sso_tables
        try:
            me_pct = float(self.pl_me_var.get() or 10)
        except (TypeError, ValueError):
            me_pct = 10.0
        if not Path(DATABASE_FILE).exists():
            messagebox.showwarning("Profitability", "Database not found.")
            return
        entity_id, entity_kind = self._selected_pl_entity()
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                ensure_profitability_tables(conn)
                ensure_sso_tables(conn)
                targets: list[tuple[int, str, str]] = []  # (id, kind, label)
                if entity_id is None:
                    char_rows = conn.execute(
                        "SELECT character_id, character_name FROM sso_character"
                    ).fetchall()
                    corp_rows = conn.execute(
                        "SELECT corporation_id, MAX(corporation_name) FROM sso_character "
                        "WHERE corporation_id IS NOT NULL GROUP BY corporation_id"
                    ).fetchall()
                    for cid, cname in char_rows:
                        targets.append((int(cid), "character", cname or f"Char {cid}"))
                    for kid, kname in corp_rows:
                        targets.append((int(kid), "corporation", (kname or f"Corp {kid}") + " (corp)"))
                elif entity_kind == "corporation":
                    row = conn.execute(
                        "SELECT MAX(corporation_name) FROM sso_character WHERE corporation_id = ?",
                        (int(entity_id),),
                    ).fetchone()
                    targets.append((int(entity_id), "corporation",
                                    ((row[0] if row else None) or f"Corp {entity_id}") + " (corp)"))
                else:
                    row = conn.execute(
                        "SELECT character_name FROM sso_character WHERE character_id = ?",
                        (int(entity_id),),
                    ).fetchone()
                    targets.append((int(entity_id), "character", (row[0] if row else None) or f"Char {entity_id}"))
                if not targets:
                    self._pl_log("No linked SSO characters. Add one in the EVE SSO Sync tab.")
                    self.pl_summary_var.set("No linked characters.")
                    return
                use_mkt = bool(self.pl_market_fallback_var.get())
                self._pl_log(
                    f"Rebuilding ledger (ME={int(me_pct)}%, "
                    f"market-fallback={'on' if use_mkt else 'off'}) for {len(targets)} entity(ies)..."
                )
                for eid, ekind, label in targets:
                    self._pl_log(f"  [{ekind}: {label}]")
                    rebuild_ledger(
                        conn, eid, default_me_percent=me_pct,
                        log=self._pl_log, entity_kind=ekind,
                        use_market_fallback=use_mkt,
                    )
            finally:
                conn.close()
        except Exception as e:
            self._pl_log(f"Rebuild failed: {e}")
            messagebox.showerror("Profitability", str(e))
            return
        self._refresh_pl_view()

    def _refresh_pl_view(self):
        if not hasattr(self, "pl_tree"):
            return
        from profitability_tracking import list_production_pl, list_production_pl_weekly
        weekly = (getattr(self, "pl_view_var", None) and self.pl_view_var.get() == "Per week (grouped)")
        try:
            for iid in self.pl_tree.get_children():
                self.pl_tree.delete(iid)
            self.pl_row_meta = {}
            if not Path(DATABASE_FILE).exists():
                self.pl_summary_var.set("Database not found.")
                return
            entity_id, entity_kind = self._selected_pl_entity()
            search = self.pl_search_var.get().strip() or None
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                if weekly:
                    rows = list_production_pl_weekly(
                        conn, entity_id=entity_id, entity_kind=entity_kind,
                        search=search, limit=2000,
                    )
                else:
                    rows = list_production_pl(
                        conn, entity_id=entity_id, entity_kind=entity_kind,
                        search=search, limit=2000,
                    )
            finally:
                conn.close()
        except Exception as e:
            self._pl_log(f"View error: {e}")
            return
        if hasattr(self, "pl_tree_frame"):
            self.pl_tree_frame.configure(
                text="Weekly grouped P/L" if weekly else "Per-job P/L"
            )
        self.pl_tree.heading("completed", text="Week" if weekly else "Completed")
        sort_col = self._pl_sort_state["col"]
        reverse = self._pl_sort_state["reverse"]
        sort_keys = {
            "completed": lambda r: r.get("end_date_max") or r.get("end_date_utc") or "",
            "char": lambda r: (r["character_name"] or "").lower(),
            "product": lambda r: (r["product_name"] or "").lower(),
            "jobs": lambda r: r.get("job_count") or 1,
            "runs": lambda r: r["runs"],
            "out_qty": lambda r: r["output_qty"],
            "job_fee": lambda r: r["job_fee"],
            "mat_cost": lambda r: r["materials_cost"],
            "inv_cost": lambda r: (r.get("invention_cost") or 0) + (r.get("facility_cost") or 0),
            "total_cost": lambda r: r["total_cost"],
            "unit_cost": lambda r: r["unit_cost"],
            "sold": lambda r: r["sold_qty"],
            "revenue": lambda r: r["revenue"],
            "realized": lambda r: r["realized_profit"],
            "unsold": lambda r: r["unsold_qty"],
            "unrealized": lambda r: r["unrealized_value"],
            "total_pl": lambda r: r["total_estimated_pl"],
            "flag": lambda r: r["materials_unknown_qty"],
        }
        try:
            rows.sort(key=sort_keys.get(sort_col, sort_keys["completed"]), reverse=reverse)
        except Exception:
            pass

        def fmt_isk(v):
            try:
                return f"{float(v):,.2f}"
            except (TypeError, ValueError):
                return "0.00"

        total_cost = 0.0
        total_revenue = 0.0
        total_realized = 0.0
        total_unrealized = 0.0
        for r in rows:
            tags = []
            pl = r["total_estimated_pl"]
            if pl > 0:
                tags.append("gain")
            elif pl < 0:
                tags.append("loss")
            if r["materials_unknown_qty"] > 0:
                tags.append("flagged")
            inv_display = float(r.get("invention_cost") or 0) + float(r.get("facility_cost") or 0)
            if (r.get("note") or "").find("no invention binding") >= 0:
                tags.append("flagged")
            job_count = int(r.get("job_count") or 1)
            if weekly:
                pt = r.get("product_type_id")
                pt_key = str(pt) if pt is not None else "n:" + (r.get("product_name") or "?")
                pl_iid = f"week|{r['entity_kind']}|{r['entity_id']}|{r.get('week_key', '')}|{pt_key}"
                completed_disp = r.get("week_key") or "?"
                if r.get("end_date_min") and r.get("end_date_max"):
                    completed_disp += f" ({(r['end_date_min'] or '')[:10]} – {(r['end_date_max'] or '')[:10]})"
            else:
                pl_iid = f"job|{r['entity_kind']}|{r['entity_id']}|{r['job_id']}"
                completed_disp = (r["end_date_utc"] or "")[:16].replace("T", " ")
            self.pl_row_meta[pl_iid] = r
            self.pl_tree.insert(
                "", tk.END, iid=pl_iid,
                values=(
                    completed_disp,
                    r["character_name"],
                    r["product_name"],
                    f"{job_count:,}",
                    f"{r['runs']:,}",
                    f"{r['output_qty']:,}",
                    fmt_isk(r["job_fee"]),
                    fmt_isk(r["materials_cost"]),
                    fmt_isk(inv_display),
                    fmt_isk(r["total_cost"]),
                    fmt_isk(r["unit_cost"]),
                    f"{r['sold_qty']:,}",
                    fmt_isk(r["revenue"]),
                    fmt_isk(r["realized_profit"]),
                    f"{r['unsold_qty']:,}",
                    fmt_isk(r["unrealized_value"]),
                    fmt_isk(r["total_estimated_pl"]),
                    "⚠" if r["materials_unknown_qty"] > 0 else "",
                ),
                tags=tuple(tags),
            )
            total_cost += r["total_cost"]
            total_revenue += r["revenue"]
            total_realized += r["realized_profit"]
            total_unrealized += r["unrealized_value"]
        n = len(rows)
        if weekly:
            n_jobs = sum(int(r.get("job_count") or 0) for r in rows)
            row_label = f"{n} group(s) ({n_jobs} jobs)"
        else:
            row_label = f"{n} job(s)"
        self.pl_summary_var.set(
            f"{row_label}  |  cost {fmt_isk(total_cost)}  |  revenue {fmt_isk(total_revenue)}  |  "
            f"realized P/L {fmt_isk(total_realized)}  |  unrealized {fmt_isk(total_unrealized)}  |  "
            f"est. total P/L {fmt_isk(total_realized + total_unrealized)}"
        )

    def _pl_sort_by(self, col: str):
        st = self._pl_sort_state
        if st["col"] == col:
            st["reverse"] = not st["reverse"]
        else:
            st["col"] = col
            st["reverse"] = True
        self._refresh_pl_view()

    def _pl_show_job_breakdown(self, _event=None):
        """Open a window with FIFO cost/sale provenance for the double-clicked row (job or weekly group)."""
        if not hasattr(self, "pl_tree"):
            return
        sel = self.pl_tree.selection()
        if not sel:
            return
        iid = sel[0]
        meta = getattr(self, "pl_row_meta", {}).get(iid)
        if not Path(DATABASE_FILE).exists():
            messagebox.showwarning("Job breakdown", "Database not found.")
            return
        from profitability_tracking import get_group_pl_breakdown, get_job_pl_breakdown
        try:
            conn = sqlite3.connect(DATABASE_FILE)
            try:
                if str(iid).startswith("week|") and meta:
                    text = get_group_pl_breakdown(
                        conn,
                        int(meta["entity_id"]),
                        meta["entity_kind"],
                        meta.get("job_ids") or [],
                        meta.get("product_name") or "?",
                        meta.get("week_key") or "?",
                    )
                    title_suffix = f"{meta.get('week_key')} ({meta.get('job_count', 0)} jobs)"
                elif str(iid).startswith("job|"):
                    parts = str(iid).split("|")
                    if len(parts) != 4:
                        raise ValueError("Invalid job row key")
                    entity_kind, entity_id, job_id = parts[1], int(parts[2]), int(parts[3])
                    text = get_job_pl_breakdown(conn, entity_id, entity_kind, job_id)
                    title_suffix = f"job {job_id}"
                else:
                    parts = str(iid).split("|")
                    if len(parts) != 3:
                        messagebox.showinfo(
                            "Job breakdown",
                            "Rebuild the ledger and try again — this row has no job key stored.",
                        )
                        return
                    entity_kind, entity_id, job_id = parts[0], int(parts[1]), int(parts[2])
                    text = get_job_pl_breakdown(conn, entity_id, entity_kind, job_id)
                    title_suffix = f"job {job_id}"
            finally:
                conn.close()
        except Exception as e:
            messagebox.showerror("Job breakdown", str(e))
            return
        win = tk.Toplevel(self.root)
        product = self.pl_tree.item(iid, "values")
        title_product = product[2] if len(product) > 2 else "?"
        win.title(f"P/L breakdown — {title_product} ({title_suffix})")
        win.geometry("920x640")
        txt = scrolledtext.ScrolledText(win, wrap=tk.WORD, font=("Consolas", 10))
        txt.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        txt.insert(tk.END, text)
        txt.configure(state=tk.DISABLED)
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=(0, 8))

    def _sso_log(self, msg: str):
        self.sso_log_text.insert(tk.END, msg + "\n")
        self.sso_log_text.see(tk.END)
        self.root.update_idletasks()
    
    def sso_login(self):
        """Run EVE SSO login flow (open browser, callback server, store tokens). Adds/updates a character."""
        from eve_sso_sync import login_flow
        cid = self.sso_client_id_var.get().strip()
        secret = self.sso_client_secret_var.get().strip()
        if not cid or not secret:
            messagebox.showwarning("SSO", "Enter Client ID and Client Secret (or set EVE_SSO_CLIENT_ID and EVE_SSO_CLIENT_SECRET).")
            return
        self.sso_status_var.set("Opening browser for EVE login...")
        self._sso_log("Starting SSO login (log in with the EVE character you want to add)...")
        def run():
            try:
                result = login_flow(cid, secret, DATABASE_FILE)
                if "error" in result:
                    self.sso_status_var.set("Login failed.")
                    err_text = result["error"]
                    self._sso_log("Error: " + err_text)
                    self.root.after(0, lambda msg=err_text: messagebox.showerror("SSO Login", msg))
                else:
                    name = result.get("character_name") or f"Character {result.get('character_id')}"
                    corp = result.get("corporation_name") or ""
                    self.sso_status_var.set(f"Linked: {name}{' / ' + corp if corp else ''}")
                    self._sso_log(
                        f"Linked: {name} (character_id={result.get('character_id')}"
                        + (f", corp={corp}" if corp else "")
                        + ")"
                    )
                    self.root.after(0, self._refresh_sso_chars_tree)
            except Exception as e:
                err_text = str(e)
                self.sso_status_var.set("Login failed.")
                self._sso_log("Error: " + err_text)
                self.root.after(0, lambda msg=err_text: messagebox.showerror("SSO Login", msg))
        threading.Thread(target=run, daemon=True).start()
    
    def refresh_exclusions_list(self):
        """Refresh the excluded modules list (no UI tab; no-op if tree was never created)."""
        tree = getattr(self, "exclusions_tree", None)
        if tree is None:
            return
        for item in tree.get_children():
            tree.delete(item)
        
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
                tree.insert('', tk.END, values=(
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
        tree = getattr(self, "exclusions_tree", None)
        if tree is None:
            return
        selected = tree.selection()
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
                values = tree.item(item, 'values')
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
        if getattr(self, "exclusions_tree", None) is None:
            return
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
    
    def _make_price_update_log_handler(self) -> logging.Handler:
        """Logging handler that appends to the Price Updates tab log on the Tk main thread (thread-safe)."""
        fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        tw = self.price_update_log
        root = self.root

        class _PriceUpdateTkHandler(logging.Handler):
            def emit(self, record):
                try:
                    msg = fmt.format(record) + "\n"
                except Exception:
                    msg = (record.getMessage() or "") + "\n"

                def _append():
                    try:
                        tw.insert(tk.END, msg)
                        tw.see(tk.END)
                    except tk.TclError:
                        pass

                root.after(0, _append)

        h = _PriceUpdateTkHandler()
        h.setLevel(logging.INFO)
        return h

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
            root_logger = logging.getLogger()
            h = self._make_price_update_log_handler()
            root_logger.addHandler(h)
            try:
                update_prices()
            except Exception as e:
                err = str(e)

                def err_ui():
                    self.price_update_log.insert(tk.END, f"\nError: {err}\n")
                    self.price_update_log.see(tk.END)
                    self.status_var.set("Error occurred")
                    messagebox.showerror("Error", f"An error occurred:\n{err}")

                self.root.after(0, err_ui)
            else:

                def ok_ui():
                    self.price_update_log.insert(tk.END, "\n\nUpdate complete!\n")
                    self.price_update_log.see(tk.END)
                    self.status_var.set("Price update complete!")
                    messagebox.showinfo("Success", "All prices updated successfully!")

                self.root.after(0, ok_ui)
            finally:
                root_logger.removeHandler(h)

        thread = threading.Thread(target=update, daemon=True)
        thread.start()
    
    def update_mineral_prices_only(self):
        """Update mineral prices (plus shopping-list products and materials) in a background thread,
        then show a before/after comparison table."""
        self.status_var.set("Updating mineral prices...")
        self.price_update_log.delete(1.0, tk.END)
        self.price_update_log.insert(tk.END, "Starting update of mineral prices + shopping list items...\n\n")
        self.root.update()

        # Collect extra type IDs from the shopping list (products + materials)
        extra_type_ids = set()
        try:
            if self.shopping_list and Path(DATABASE_FILE).exists():
                conn = sqlite3.connect(DATABASE_FILE)
                try:
                    for entry in self.shopping_list:
                        bp = resolve_blueprint(conn, entry["product_name"])
                        if bp:
                            extra_type_ids.add(bp["productTypeID"])
                            for m in get_blueprint_materials(conn, bp["blueprintTypeID"]):
                                extra_type_ids.add(m["materialTypeID"])
                finally:
                    conn.close()
        except Exception:
            pass

        def update():
            root_logger = logging.getLogger()
            h = self._make_price_update_log_handler()
            root_logger.addHandler(h)
            try:
                comparison = update_mineral_prices(extra_type_ids=extra_type_ids if extra_type_ids else None)
            except Exception as e:
                err = str(e)

                def err_ui():
                    self.price_update_log.insert(tk.END, f"\nError: {err}\n")
                    self.price_update_log.see(tk.END)
                    self.status_var.set("Error occurred")
                    messagebox.showerror("Error", f"An error occurred:\n{err}")

                self.root.after(0, err_ui)
            else:

                def done():
                    self.price_update_log.insert(tk.END, "\nMineral price update complete!\n")
                    self.price_update_log.see(tk.END)
                    self.status_var.set("Mineral price update complete!")
                    if comparison:
                        self._show_price_comparison_popup(comparison)

                self.root.after(0, done)
            finally:
                root_logger.removeHandler(h)

        threading.Thread(target=update, daemon=True).start()

    def _show_price_comparison_popup(self, comparison):
        """Show a Toplevel table with before/after sell and buy prices for each updated item."""
        win = tk.Toplevel(self.root)
        win.title("Price Update — Before / After")
        win.geometry("820x600")
        win.minsize(700, 400)

        ttk.Label(win, text="Price changes after update  (sell = sell_min, buy = buy_max)",
                  font=("TkDefaultFont", 9, "italic")).pack(anchor=tk.W, padx=10, pady=(8, 2))

        cols = ("Item", "Prev Sell", "New Sell", "Δ% Sell", "Prev Buy", "New Buy", "Δ% Buy")
        tree = ttk.Treeview(win, columns=cols, show="headings", height=22)
        col_widths = {"Item": 240, "Prev Sell": 100, "New Sell": 100, "Δ% Sell": 75,
                      "Prev Buy": 100, "New Buy": 100, "Δ% Buy": 75}
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=col_widths.get(c, 90),
                        anchor=tk.E if c != "Item" else tk.W,
                        stretch=(c == "Item"))

        tree.tag_configure("mineral",  background="#e8f4fd")
        tree.tag_configure("extra",    background="#f0f9e8")
        tree.tag_configure("up",       foreground="#1a7a1a")
        tree.tag_configure("down",     foreground="#cc2200")
        tree.tag_configure("sep",      background="#dddddd")

        def fmt_isk(v):
            if v is None or v == 0:
                return "—"
            return f"{v:,.2f}"

        def fmt_pct(old, new):
            if not old or old == 0:
                return "—"
            pct = (new - old) / old * 100
            sign = "+" if pct >= 0 else ""
            return f"{sign}{pct:.1f}%"

        # Separate minerals from non-minerals, keep ordering from comparison dict
        minerals = [(tid, d) for tid, d in comparison.items() if d.get("is_mineral")]
        extras   = [(tid, d) for tid, d in comparison.items() if d.get("is_extra") and not d.get("is_mineral")]
        others   = [(tid, d) for tid, d in comparison.items()
                    if not d.get("is_mineral") and not d.get("is_extra")]

        def add_section_header(label):
            tree.insert("", tk.END, values=(f"── {label} ──", "", "", "", "", "", ""), tags=("sep",))

        def add_rows(rows, tag):
            for tid, d in rows:
                old_sell = d["old_sell"]
                new_sell = d["new_sell"]
                old_buy  = d["old_buy"]
                new_buy  = d["new_buy"]
                sell_pct = fmt_pct(old_sell, new_sell)
                buy_pct  = fmt_pct(old_buy,  new_buy)
                row_tags = [tag]
                if new_sell > old_sell and old_sell > 0:
                    row_tags.append("up")
                elif new_sell < old_sell and old_sell > 0:
                    row_tags.append("down")
                tree.insert("", tk.END,
                            values=(d["name"], fmt_isk(old_sell), fmt_isk(new_sell), sell_pct,
                                    fmt_isk(old_buy),  fmt_isk(new_buy),  buy_pct),
                            tags=tuple(row_tags))

        if minerals:
            add_section_header("Minerals")
            add_rows(minerals, "mineral")
        if extras:
            add_section_header("Shopping list items")
            add_rows(extras, "extra")
        if others:
            add_section_header("Other updated items")
            add_rows(others, "")

        scroll = ttk.Scrollbar(win, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=(0, 10))
        scroll.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10), pady=(0, 10))

        ttk.Button(win, text="Close", command=win.destroy).pack(pady=6)
    
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
                        def no_items():
                            self.price_update_log.insert(tk.END, "No items with blueprint source found in database.\n")
                            self.price_update_log.see(tk.END)
                            self.status_var.set("No blueprint items found")

                        self.root.after(0, no_items)
                        return
                    
                    self.price_update_log.insert(tk.END, f"Found {len(type_ids)} items with blueprint source.\n")
                    self.price_update_log.insert(tk.END, "Updating prices...\n\n")
                    self.root.update()
                    
                finally:
                    conn.close()
                
                # Stream logs to Update Log on main thread
                root_logger = logging.getLogger()
                h = self._make_price_update_log_handler()
                root_logger.addHandler(h)
                try:
                    update_prices_by_type_ids(type_ids, f"blueprint items (source='blueprint')")
                finally:
                    root_logger.removeHandler(h)

                def ok_ui():
                    self.price_update_log.insert(tk.END, "\n\nBlueprint price update complete!\n")
                    self.price_update_log.see(tk.END)
                    self.status_var.set("Blueprint price update complete!")
                    messagebox.showinfo("Success", f"Updated prices for {len(type_ids)} blueprint items successfully!")

                self.root.after(0, ok_ui)

            except Exception as e:
                err = str(e)

                def err_ui():
                    self.price_update_log.insert(tk.END, f"\nError: {err}\n")
                    self.price_update_log.see(tk.END)
                    self.status_var.set("Error occurred")
                    messagebox.showerror("Error", f"An error occurred:\n{err}")

                self.root.after(0, err_ui)
        
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
                        def no_items():
                            self.price_update_log.insert(tk.END, "No items with group consensus source found in database.\n")
                            self.price_update_log.see(tk.END)
                            self.status_var.set("No group consensus items found")

                        self.root.after(0, no_items)
                        return
                    
                    self.price_update_log.insert(tk.END, f"Found {len(type_ids)} items with group consensus source.\n")
                    self.price_update_log.insert(tk.END, "Updating prices...\n\n")
                    self.root.update()
                    
                finally:
                    conn.close()
                
                root_logger = logging.getLogger()
                h = self._make_price_update_log_handler()
                root_logger.addHandler(h)
                try:
                    update_prices_by_type_ids(type_ids, f"group consensus items (source='group_consensus')")
                finally:
                    root_logger.removeHandler(h)

                def ok_ui():
                    self.price_update_log.insert(tk.END, "\n\nGroup consensus price update complete!\n")
                    self.price_update_log.see(tk.END)
                    self.status_var.set("Group consensus price update complete!")
                    messagebox.showinfo("Success", f"Updated prices for {len(type_ids)} group consensus items successfully!")

                self.root.after(0, ok_ui)

            except Exception as e:
                err = str(e)

                def err_ui():
                    self.price_update_log.insert(tk.END, f"\nError: {err}\n")
                    self.price_update_log.see(tk.END)
                    self.status_var.set("Error occurred")
                    messagebox.showerror("Error", f"An error occurred:\n{err}")

                self.root.after(0, err_ui)
        
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
            root_logger = logging.getLogger()
            h = self._make_price_update_log_handler()
            root_logger.addHandler(h)
            try:
                run_fetch(
                    region_id=MARKET_HISTORY_REGION_ID,
                    all_items=True,
                    scope="prices",
                    delay_seconds=1.0,
                    progress_interval=50,
                )
            except Exception as e:
                err = str(e)

                def err_ui():
                    self.price_update_log.insert(tk.END, f"\nError: {err}\n")
                    self.price_update_log.see(tk.END)
                    self.status_var.set("Error occurred")
                    messagebox.showerror("Error", f"An error occurred:\n{err}")

                self.root.after(0, err_ui)
            else:

                def ok_ui():
                    self.price_update_log.insert(tk.END, "\n\nMarket history fetch complete!\n")
                    self.price_update_log.see(tk.END)
                    self.status_var.set("Market history fetch complete!")
                    messagebox.showinfo("Success", "Market history fetch complete!")

                self.root.after(0, ok_ui)
            finally:
                root_logger.removeHandler(h)

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
                    def no_refresh():
                        self.price_update_log.insert(tk.END, "No items need refresh (all have volume data).\n")
                        self.price_update_log.see(tk.END)
                        self.status_var.set("No items to refresh")
                        messagebox.showinfo("Info", "No items with missing/zero volume data found.")

                    self.root.after(0, no_refresh)
                    return

                cnt = len(to_refresh)

                def found_msg():
                    self.price_update_log.insert(tk.END, f"Found {cnt} items to refresh. Calling API...\n\n")
                    self.price_update_log.see(tk.END)
                    self.root.update_idletasks()

                self.root.after(0, found_msg)
                conn = sqlite3.connect(DATABASE_FILE)
                try:
                    done = 0
                    n_last = 0
                    for i, type_id in enumerate(to_refresh):
                        n_last = refresh_market_history_for_type(conn, MARKET_HISTORY_REGION_ID, type_id)
                        done += 1
                        if (i + 1) % 50 == 0:
                            cur_i = i + 1
                            total = len(to_refresh)
                            tid = type_id
                            n = n_last

                            def _prog(ci=cur_i, tot=total, t=tid, nd=n):
                                self.price_update_log.insert(
                                    tk.END,
                                    f"  Refreshed {ci}/{tot} (last: type_id={t}, {nd} days)\n",
                                )
                                self.price_update_log.see(tk.END)
                                self.root.update_idletasks()

                            self.root.after(0, _prog)
                    self.root.after(
                        0,
                        lambda d=done: self.price_update_log.insert(tk.END, f"\nRefreshed {d} items.\n"),
                    )
                    self.root.after(0, lambda: self.price_update_log.see(tk.END))
                    self.root.after(0, lambda: self.status_var.set("Volume refresh complete!"))
                    self.root.after(0, lambda d=done: messagebox.showinfo("Success", f"Refreshed market history for {d} items."))
                finally:
                    conn.close()
            except Exception as e:
                err = str(e)

                def err_ui():
                    self.price_update_log.insert(tk.END, f"\nError: {err}\n")
                    self.price_update_log.see(tk.END)
                    self.status_var.set("Error occurred")
                    messagebox.showerror("Error", f"An error occurred:\n{err}")

                self.root.after(0, err_ui)
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


def ensure_runtime_database():
    """Bootstrap the runtime DB from the git-tracked core snapshot on a fresh clone.

    The large runtime DB (with market history) is git-ignored, so a new checkout only
    has eve_manufacturing_core.db. Copy it to eve_manufacturing.db so the launcher works
    out of the box; market history can then be rebuilt locally via fetch_market_history.
    """
    try:
        runtime = Path(DATABASE_FILE)
        core = Path(CORE_DATABASE_FILE)
        if not runtime.exists() and core.exists():
            import shutil
            shutil.copyfile(core, runtime)
            print(f"Initialized {DATABASE_FILE} from {CORE_DATABASE_FILE} "
                  "(market history empty — rebuild locally with fetch_market_history.py).")
    except Exception as e:
        print(f"Could not bootstrap runtime database: {e}")


def main():
    ensure_runtime_database()
    root = tk.Tk()
    app = EVELauncher(root)
    root.mainloop()


if __name__ == "__main__":
    main()

