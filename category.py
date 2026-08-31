"""
Guesses an expense category from the free-text note, the same way the
original bot pulled a "sender" out of a note — just keyword buckets instead
of a "from X" pattern, since a contractor's texted note ("lumber for
framing", "paid the electrician") doesn't follow one fixed phrase.

Edit CATEGORY_KEYWORDS freely — it's checked in order, first match wins, so
put more specific words earlier if two categories could both match.
"""

CATEGORY_KEYWORDS = [
    ("Materials", [
        "lumber", "concrete", "cement", "drywall", "paint", "material",
        "supplies", "hardware", "plumbing supplies", "wood", "steel",
        "tile", "brick", "insulation", "roofing", "pipe", "wire", "wiring",
    ]),
    ("Subcontractor", [
        "subcontractor", "sub ", "electrician", "plumber", "painter",
        "framer", "mason", "roofer", "hvac",
    ]),
    ("Labor", [
        "labor", "labour", "crew", "worker", "workers", "payroll", "wages", "wage",
    ]),
    ("Equipment Rental", [
        "rental", "rent ", "excavator", "generator", "scaffolding", "scaffold",
        "machine", "equipment",
    ]),
    ("Permits & Fees", [
        "permit", "inspection", "license", "licence", "fee", "fine",
    ]),
    ("Fuel", [
        "gas station", "diesel", "fuel", " gas ", "gas for",
    ]),
    ("Tools", [
        "tools", "drill", "saw", "hammer", "toolbox", "nail gun",
    ]),
    ("Delivery & Shipping", [
        "delivery", "shipping", "freight", "trucking",
    ]),
]


def detect_category(note):
    text = f" {(note or '').lower()} "
    for category, keywords in CATEGORY_KEYWORDS:
        for kw in keywords:
            if kw in text:
                return category
    return "Other"


def category_list():
    return [c for c, _ in CATEGORY_KEYWORDS] + ["Other"]
