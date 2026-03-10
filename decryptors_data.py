"""
EVE Online decryptor data for invention.

Decryptors are consumed during T2 invention jobs. They modify:
- Invention success probability (multiplier on base chance)
- Resulting BPC: ME (Material Efficiency), TE (Time Efficiency), and max runs.

Base T2 BPC from invention (no decryptor): 2% ME, 4% TE; 10 runs (modules/ammo/etc) or 1 run (ships/rigs).
Decryptor modifiers are added to these base values.

References:
- https://support.eveonline.com/hc/en-us/articles/203270631-Decryptors
- https://everef.net/groups/1304
"""

# (name, type_id, probability_multiplier, run_modifier, me_modifier, te_modifier)
# Probability: 0.6 = -40%, 1.0 = 0%, 1.9 = +90%
# ME/TE mods are added to base 2% ME and 4% TE; run_mod added to base 10 (or 1 for ships)
DECRYPTORS = [
    ("Augmentation Decryptor", 34203, 0.6, 9, -2, 2),
    ("Optimized Augmentation Decryptor", 34208, 0.9, 7, 2, 0),
    ("Symmetry Decryptor", 34206, 1.0, 2, 1, 8),
    ("Process Decryptor", 34205, 1.1, 0, 3, 6),
    ("Accelerant Decryptor", 34201, 1.2, 1, 2, 10),
    ("Parity Decryptor", 34204, 1.5, 3, 1, -2),
    ("Attainment Decryptor", 34202, 1.8, 4, -1, 4),
    ("Optimized Attainment Decryptor", 34207, 1.9, 2, 1, -2),
]

# Base output of a successful T2 invention when no decryptor is used
BASE_ME_PCT = 2
BASE_TE_PCT = 4
BASE_RUNS_MODULE = 10   # modules, ammo, charges, etc.
BASE_RUNS_SHIP = 1      # ships, rigs


def get_decryptor_type_ids():
    """Return list of decryptor type IDs for price lookups."""
    return [d[1] for d in DECRYPTORS]


def get_decryptor_by_name(name):
    """Return (name, type_id, prob_mult, run_mod, me_mod, te_mod) or None."""
    for t in DECRYPTORS:
        if t[0] == name:
            return t
    return None
