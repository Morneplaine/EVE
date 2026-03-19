-- EVE Manufacturing Database Schema
-- SQLite database for storing EVE Online manufacturing data

CREATE TABLE IF NOT EXISTS items (
    typeID INTEGER PRIMARY KEY,
    typeName TEXT NOT NULL,
    groupID INTEGER,
    categoryID INTEGER,
    volume REAL,
    packaged_volume REAL,
    techLevel INTEGER DEFAULT 0,
    isFaction INTEGER DEFAULT 0,
    FOREIGN KEY (groupID) REFERENCES groups(groupID)
);

-- Groups table: Item groups
CREATE TABLE IF NOT EXISTS groups (
    groupID INTEGER PRIMARY KEY,
    groupName TEXT NOT NULL,
    categoryID INTEGER
);

-- Blueprints table: Manufacturing blueprints
CREATE TABLE IF NOT EXISTS blueprints (
    blueprintTypeID INTEGER PRIMARY KEY,
    productTypeID INTEGER NOT NULL,
    productName TEXT NOT NULL,
    outputQuantity INTEGER NOT NULL,
    groupID INTEGER,
    groupName TEXT,
    FOREIGN KEY (productTypeID) REFERENCES items(typeID),
    FOREIGN KEY (blueprintTypeID) REFERENCES items(typeID)
);

-- Manufacturing materials: Materials required for each blueprint
CREATE TABLE IF NOT EXISTS manufacturing_materials (
    blueprintTypeID INTEGER NOT NULL,
    materialTypeID INTEGER NOT NULL,
    materialName TEXT NOT NULL,
    quantity INTEGER NOT NULL,
    PRIMARY KEY (blueprintTypeID, materialTypeID),
    FOREIGN KEY (blueprintTypeID) REFERENCES blueprints(blueprintTypeID),
    FOREIGN KEY (materialTypeID) REFERENCES items(typeID)
);

-- Manufacturing skills: Skills required for each blueprint
CREATE TABLE IF NOT EXISTS manufacturing_skills (
    blueprintTypeID INTEGER NOT NULL,
    skillID INTEGER NOT NULL,
    skillName TEXT NOT NULL,
    level INTEGER NOT NULL,
    PRIMARY KEY (blueprintTypeID, skillID),
    FOREIGN KEY (blueprintTypeID) REFERENCES blueprints(blueprintTypeID),
    FOREIGN KEY (skillID) REFERENCES items(typeID)
);

-- Reprocessing outputs: What materials items break down into
CREATE TABLE IF NOT EXISTS reprocessing_outputs (
    itemTypeID INTEGER NOT NULL,
    itemName TEXT NOT NULL,
    materialTypeID INTEGER NOT NULL,
    materialName TEXT NOT NULL,
    quantity INTEGER NOT NULL,
    PRIMARY KEY (itemTypeID, materialTypeID),
    FOREIGN KEY (itemTypeID) REFERENCES items(typeID),
    FOREIGN KEY (materialTypeID) REFERENCES items(typeID)
);

-- Prices table: Market prices (updated frequently)
CREATE TABLE IF NOT EXISTS prices (
    typeID INTEGER PRIMARY KEY,
    buy_max REAL DEFAULT 0,
    buy_volume REAL DEFAULT 0,
    sell_min REAL DEFAULT 0,
    sell_avg REAL DEFAULT 0,
    sell_median REAL DEFAULT 0,
    sell_volume REAL DEFAULT 0,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (typeID) REFERENCES items(typeID)
);

-- Character skills: User's character skills (for filtering/calculations)
CREATE TABLE IF NOT EXISTS character_skills (
    skillID INTEGER PRIMARY KEY,
    skillName TEXT NOT NULL,
    level INTEGER NOT NULL CHECK (level >= 0 AND level <= 5),
    FOREIGN KEY (skillID) REFERENCES items(typeID)
);

-- Inventory: User's available resources
CREATE TABLE IF NOT EXISTS inventory (
    typeID INTEGER PRIMARY KEY,
    typeName TEXT NOT NULL,
    quantity INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY (typeID) REFERENCES items(typeID)
);

-- Input quantity cache: Cached input quantities for items based on group analysis
CREATE TABLE IF NOT EXISTS input_quantity_cache (
    typeID INTEGER PRIMARY KEY,
    typeName TEXT NOT NULL,
    input_quantity INTEGER NOT NULL,
    source TEXT NOT NULL,  -- 'blueprint', 'group_consensus', 'group_most_frequent', 'default'
    needs_review INTEGER DEFAULT 0,  -- 1 if needs manual review, 0 otherwise
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (typeID) REFERENCES items(typeID)
);

-- Market history (EVE Tycoon): daily trade stats per region/type for Jita (The Forge)
-- API: https://evetycoon.com/api/v1/market/history/{regionId}/{typeId}
-- One request per typeId; API returns full history (no accumulation needed).
-- transaction_skew is not stored; recalculate as 1 - ((average - lowest) / (highest - lowest)) when needed.
CREATE TABLE IF NOT EXISTS market_history_daily (
    region_id INTEGER NOT NULL,
    type_id INTEGER NOT NULL,
    type_name TEXT,
    date_utc TEXT NOT NULL,
    average REAL NOT NULL,
    highest REAL NOT NULL,
    lowest REAL NOT NULL,
    order_count INTEGER,
    volume INTEGER,
    PRIMARY KEY (region_id, type_id, date_utc),
    FOREIGN KEY (type_id) REFERENCES items(typeID)
);

CREATE INDEX IF NOT EXISTS idx_market_history_type_date ON market_history_daily(type_id, date_utc);

-- Static region list (for market / manufacturing tax region selection)
CREATE TABLE IF NOT EXISTS regions (
    region_id INTEGER PRIMARY KEY,
    region_name TEXT NOT NULL
);

-- Index for name lookup
CREATE INDEX IF NOT EXISTS idx_regions_name ON regions(region_name);

-- EVE SSO / ESI sync: tokens and synced data for profitability tracking
CREATE TABLE IF NOT EXISTS sso_character (
    character_id INTEGER PRIMARY KEY,
    character_name TEXT,
    refresh_token TEXT,
    access_token TEXT,
    access_token_expires_at REAL,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS esi_wallet_transactions (
    character_id INTEGER NOT NULL,
    transaction_id BIGINT NOT NULL,
    date_utc TEXT NOT NULL,
    type_id INTEGER,
    quantity INTEGER,
    unit_price REAL,
    client_id INTEGER,
    location_id INTEGER,
    is_buy INTEGER,
    is_personal INTEGER,
    journal_ref_id BIGINT,
    synced_at TEXT DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (character_id, transaction_id),
    FOREIGN KEY (character_id) REFERENCES sso_character(character_id)
);

CREATE TABLE IF NOT EXISTS esi_wallet_journal (
    character_id INTEGER NOT NULL,
    ref_id BIGINT NOT NULL,
    date_utc TEXT NOT NULL,
    ref_type TEXT,
    amount REAL,
    balance REAL,
    context_id_type TEXT,
    context_id BIGINT,
    description TEXT,
    first_party_id INTEGER,
    second_party_id INTEGER,
    reason TEXT,
    synced_at TEXT DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (character_id, ref_id),
    FOREIGN KEY (character_id) REFERENCES sso_character(character_id)
);

CREATE TABLE IF NOT EXISTS esi_industry_jobs (
    character_id INTEGER NOT NULL,
    job_id INTEGER NOT NULL,
    activity_id INTEGER,
    blueprint_id BIGINT,
    blueprint_type_id INTEGER,
    blueprint_location_id BIGINT,
    output_location_id BIGINT,
    runs INTEGER,
    cost REAL,
    licensed_runs INTEGER,
    probability REAL,
    product_type_id INTEGER,
    status TEXT,
    duration INTEGER,
    start_date_utc TEXT,
    end_date_utc TEXT,
    completed_date_utc TEXT,
    facility_id BIGINT,
    installer_id INTEGER,
    synced_at TEXT DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (character_id, job_id),
    FOREIGN KEY (character_id) REFERENCES sso_character(character_id)
);

CREATE INDEX IF NOT EXISTS idx_esi_tx_character_date ON esi_wallet_transactions(character_id, date_utc);
CREATE INDEX IF NOT EXISTS idx_esi_journal_character_date ON esi_wallet_journal(character_id, date_utc);
CREATE INDEX IF NOT EXISTS idx_esi_jobs_character_status ON esi_industry_jobs(character_id, status);

-- Invention: T1 blueprint -> T2 blueprint (activityID 8 in SDE industryActivityProducts)
CREATE TABLE IF NOT EXISTS invention_recipes (
    t1_blueprint_type_id INTEGER NOT NULL,
    t2_blueprint_type_id INTEGER NOT NULL,
    quantity INTEGER NOT NULL DEFAULT 1,
    probability REAL,
    PRIMARY KEY (t1_blueprint_type_id, t2_blueprint_type_id),
    FOREIGN KEY (t1_blueprint_type_id) REFERENCES items(typeID),
    FOREIGN KEY (t2_blueprint_type_id) REFERENCES blueprints(blueprintTypeID)
);
CREATE INDEX IF NOT EXISTS idx_invention_t1 ON invention_recipes(t1_blueprint_type_id);

-- User-defined datacore bindings per T2 blueprint (invention product)
-- Keyed by T2 blueprint type ID; when user runs decryptor comparison we pre-fill from here
CREATE TABLE IF NOT EXISTS blueprint_datacore_bindings (
    blueprint_type_id INTEGER PRIMARY KEY,
    dc1_name TEXT,
    dc1_qty INTEGER NOT NULL DEFAULT 0,
    dc2_name TEXT,
    dc2_qty INTEGER NOT NULL DEFAULT 0,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (blueprint_type_id) REFERENCES blueprints(blueprintTypeID)
);

-- Indexes for performance
CREATE INDEX IF NOT EXISTS idx_blueprints_product ON blueprints(productTypeID);
CREATE INDEX IF NOT EXISTS idx_materials_blueprint ON manufacturing_materials(blueprintTypeID);
CREATE INDEX IF NOT EXISTS idx_materials_material ON manufacturing_materials(materialTypeID);
CREATE INDEX IF NOT EXISTS idx_skills_blueprint ON manufacturing_skills(blueprintTypeID);
CREATE INDEX IF NOT EXISTS idx_reprocessing_item ON reprocessing_outputs(itemTypeID);
CREATE INDEX IF NOT EXISTS idx_prices_updated ON prices(updated_at);
CREATE INDEX IF NOT EXISTS idx_input_quantity_cache_review ON input_quantity_cache(needs_review);

