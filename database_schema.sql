-- EVE Manufacturing Database Schema
-- SQLite database for storing EVE Online manufacturing data

-- Items table: All items in EVE
CREATE TABLE IF NOT EXISTS items (
    typeID INTEGER PRIMARY KEY,
    typeName TEXT NOT NULL,
    groupID INTEGER,
    categoryID INTEGER,
    volume REAL,
    packaged_volume REAL,
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

-- Indexes for performance
CREATE INDEX IF NOT EXISTS idx_blueprints_product ON blueprints(productTypeID);
CREATE INDEX IF NOT EXISTS idx_materials_blueprint ON manufacturing_materials(blueprintTypeID);
CREATE INDEX IF NOT EXISTS idx_materials_material ON manufacturing_materials(materialTypeID);
CREATE INDEX IF NOT EXISTS idx_skills_blueprint ON manufacturing_skills(blueprintTypeID);
CREATE INDEX IF NOT EXISTS idx_reprocessing_item ON reprocessing_outputs(itemTypeID);
CREATE INDEX IF NOT EXISTS idx_prices_updated ON prices(updated_at);

