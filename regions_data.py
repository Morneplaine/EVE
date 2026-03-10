"""
Static EVE region list (region_id, region_name) for market/tax region selection.
Default region: The Citadel (10000033).
"""

# (region_id, region_name) - sorted by region_name for dropdown
REGIONS = [
    (12000001, "ADR01"),
    (12000002, "ADR02"),
    (12000003, "ADR03"),
    (12000004, "ADR04"),
    (12000005, "ADR05"),
    (11000001, "A-R00001"),
    (11000002, "A-R00002"),
    (11000003, "A-R00003"),
    (10000019, "A821-A"),
    (10000054, "Aridia"),
    (11000004, "B-R00004"),
    (11000005, "B-R00005"),
    (11000006, "B-R00006"),
    (11000007, "B-R00007"),
    (11000008, "B-R00008"),
    (10000069, "Black Rise"),
    (10000055, "Branch"),
    (11000009, "C-R00009"),
    (11000010, "C-R00010"),
    (11000011, "C-R00011"),
    (11000012, "C-R00012"),
    (11000013, "C-R00013"),
    (11000014, "C-R00014"),
    (11000015, "C-R00015"),
    (10000007, "Cache"),
    (10000014, "Catch"),
    (10000051, "Cloud Ring"),
    (10000053, "Cobalt Edge"),
    (10000012, "Curse"),
    (11000016, "D-R00016"),
    (11000017, "D-R00017"),
    (11000018, "D-R00018"),
    (11000019, "D-R00019"),
    (11000020, "D-R00020"),
    (11000021, "D-R00021"),
    (11000022, "D-R00022"),
    (11000023, "D-R00023"),
    (10000035, "Deklein"),
    (10000060, "Delve"),
    (10000001, "Derelik"),
    (10000005, "Detorid"),
    (10000036, "Devoid"),
    (10000043, "Domain"),
    (11000024, "E-R00024"),
    (11000025, "E-R00025"),
    (11000026, "E-R00026"),
    (11000027, "E-R00027"),
    (11000028, "E-R00028"),
    (11000029, "E-R00029"),
    (10000039, "Esoteria"),
    (10000064, "Essence"),
    (10000027, "Etherium Reach"),
    (10000037, "Everyshore"),
    (11000030, "F-R00030"),
    (10000046, "Fade"),
    (10000056, "Feythabolis"),
    (10000058, "Fountain"),
    (11000031, "G-R00031"),
    (10000029, "Geminate"),
    (10000067, "Genesis"),
    (10000011, "Great Wildlands"),
    (11000032, "H-R00032"),
    (10000030, "Heimatar"),
    (10000025, "Immensea"),
    (10000031, "Impass"),
    (10000009, "Insmother"),
    (10000017, "J7HZ-F"),
    (11000033, "K-R00033"),
    (10000052, "Kador"),
    (10000049, "Khanid"),
    (10000065, "Kor-Azor"),
    (10000016, "Lonetrek"),
    (10000013, "Malpais"),
    (10000042, "Metropolis"),
    (10000028, "Molden Heath"),
    (10000040, "Oasa"),
    (10000062, "Omist"),
    (10000021, "Outer Passage"),
    (10000057, "Outer Ring"),
    (10000059, "Paragon Soul"),
    (10000063, "Period Basis"),
    (10000066, "Perrigen Falls"),
    (10000048, "Placid"),
    (10000070, "Pochven"),
    (10000047, "Providence"),
    (10000023, "Pure Blind"),
    (10000050, "Querious"),
    (10000008, "Scalding Pass"),
    (10000032, "Sinq Laison"),
    (10000044, "Solitude"),
    (10000022, "Stain"),
    (10000041, "Syndicate"),
    (10000020, "Tash-Murkon"),
    (10000045, "Tenal"),
    (10000061, "Tenerifis"),
    (10000038, "The Bleak Lands"),
    (10000033, "The Citadel"),
    (10000002, "The Forge"),
    (10000034, "The Kalevala Expanse"),
    (10000018, "The Spire"),
    (10000010, "Tribute"),
    (10000004, "UUA-F4"),
    (10000003, "Vale of the Silent"),
    (10000015, "Venal"),
    (10000068, "Verge Vendor"),
    (14000001, "VR-01"),
    (14000002, "VR-02"),
    (14000003, "VR-03"),
    (14000004, "VR-04"),
    (14000005, "VR-05"),
    (10000006, "Wicked Creek"),
    (10001000, "Yasna Zakh"),
]

# Sorted by name for dropdown display
REGIONS_BY_NAME = sorted(REGIONS, key=lambda x: x[1])

# Default region (The Citadel)
DEFAULT_REGION_ID = 10000033
DEFAULT_REGION_NAME = "The Citadel"


def get_region_id_by_name(name):
    """Return region_id for the given region name, or DEFAULT_REGION_ID if not found."""
    for rid, rname in REGIONS:
        if rname == name:
            return rid
    return DEFAULT_REGION_ID


def ensure_regions_table(conn):
    """Create regions table if missing and populate from static data if empty."""
    conn.execute("""
        CREATE TABLE IF NOT EXISTS regions (
            region_id INTEGER PRIMARY KEY,
            region_name TEXT NOT NULL
        )
    """)
    cur = conn.execute("SELECT COUNT(*) FROM regions")
    if cur.fetchone()[0] == 0:
        conn.executemany(
            "INSERT OR IGNORE INTO regions (region_id, region_name) VALUES (?, ?)",
            REGIONS,
        )
        conn.commit()
