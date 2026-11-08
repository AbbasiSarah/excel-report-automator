"""Create messy sample sales files, like the ones a real business sends.

Each branch sends its own Excel file every month, and every branch does it a bit differently:
different column names, extra spaces, mixed date formats, duplicates and missing values.
"""
import random
from pathlib import Path
import pandas as pd

random.seed(7)
OUT = Path(__file__).parent / "input"
OUT.mkdir(exist_ok=True)

PRODUCTS = {"Laptop": 950, "Phone": 420, "Headphones": 60, "Keyboard": 35, "Monitor": 210}
BRANCHES = {
    "lahore": {"region": "Punjab", "cols": ["Date", "Product", "Qty", "Unit Price", "Region"]},
    "karachi": {"region": "Sindh", "cols": ["order date", "item", "quantity", "price", "area"]},
    "islamabad": {"region": "Capital", "cols": ["DATE", "PRODUCT NAME", "UNITS", "UNIT_PRICE", "REGION"]},
}
MONTHS = [(2026, 8), (2026, 9), (2026, 10)]


def messy_name(p):
    return random.choice([p, p.lower(), p.upper(), f" {p} ", f"{p}  "])


def messy_date(y, m, d, style):
    dt = pd.Timestamp(y, m, d)
    return {"iso": dt.strftime("%Y-%m-%d"), "dmy": dt.strftime("%d/%m/%Y"), "text": dt.strftime("%d %b %Y")}[style]


for branch, info in BRANCHES.items():
    style = {"lahore": "iso", "karachi": "dmy", "islamabad": "text"}[branch]
    for y, m in MONTHS:
        rows = []
        for _ in range(random.randint(28, 40)):
            p = random.choice(list(PRODUCTS))
            price = PRODUCTS[p] * random.uniform(0.95, 1.05)
            rows.append([messy_date(y, m, random.randint(1, 28), style), messy_name(p),
                         random.randint(1, 6), round(price, 2), random.choices([info["region"], info["region"].lower(), ""], weights=[6, 3, 1])[0]])
        # real-world problems
        rows += random.sample(rows, 2)                      # duplicate rows
        rows[3][2] = None                                   # missing quantity
        rows[5][3] = None                                   # missing price
        rows[7][2] = -rows[7][2]                            # negative quantity (typo)
        df = pd.DataFrame(rows, columns=info["cols"])
        df.to_excel(OUT / f"{branch}_sales_{y}_{m:02d}.xlsx", index=False)

print(f"Created {len(list(OUT.glob('*.xlsx')))} messy files in {OUT}")
