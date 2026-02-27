"""
Mock Data Generator
===================
Generates sample CSV and JSON files used by the Streamlit demo pages.
Run this script once before launching the app, or let the app generate
data on first run via the helper in data/loader.py.

Usage:
    python scripts/generate_mock_data.py
"""

import csv
import json
import os
import random
from datetime import datetime, timedelta
from pathlib import Path

# ---------------------------------------------------------------------------
# Resolve output directory relative to this script so it works from anywhere
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = SCRIPT_DIR.parent / "data" / "mock_data"


def ensure_dir() -> None:
    """Create the output directory if it does not exist."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------------
# Employees dataset (CSV) – used by CRUD demo, visualization, tables
# ---------------------------------------------------------------------------
DEPARTMENTS = ["Engineering", "Marketing", "Sales", "HR", "Finance", "Support"]
FIRST_NAMES = [
    "Alice", "Bob", "Carol", "David", "Eva", "Frank", "Grace", "Henry",
    "Irene", "Jack", "Karen", "Leo", "Mona", "Nick", "Olivia", "Paul",
    "Quinn", "Rita", "Steve", "Tina",
]
LAST_NAMES = [
    "Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller",
    "Davis", "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez",
    "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin",
]


def generate_employees(n: int = 50) -> list[dict]:
    """Return *n* synthetic employee records."""
    rows: list[dict] = []
    base_date = datetime(2020, 1, 1)
    for i in range(1, n + 1):
        hire_date = base_date + timedelta(days=random.randint(0, 1500))
        rows.append(
            {
                "id": i,
                "first_name": random.choice(FIRST_NAMES),
                "last_name": random.choice(LAST_NAMES),
                "email": f"employee{i}@example.com",
                "department": random.choice(DEPARTMENTS),
                "salary": round(random.uniform(35_000, 120_000), 2),
                "hire_date": hire_date.strftime("%Y-%m-%d"),
                "active": random.choice([True, True, True, False]),
            }
        )
    return rows


def write_employees_csv(rows: list[dict]) -> Path:
    path = DATA_DIR / "employees.csv"
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)
    return path


def write_employees_json(rows: list[dict]) -> Path:
    path = DATA_DIR / "employees.json"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(rows, f, indent=2)
    return path


# ---------------------------------------------------------------------------
# Sales dataset (CSV) – used by visualization / charts
# ---------------------------------------------------------------------------
PRODUCTS = ["Widget A", "Widget B", "Gadget X", "Gadget Y", "Service Plan"]


def generate_sales(n: int = 200) -> list[dict]:
    rows: list[dict] = []
    base_date = datetime(2024, 1, 1)
    for i in range(1, n + 1):
        sale_date = base_date + timedelta(days=random.randint(0, 364))
        qty = random.randint(1, 50)
        unit_price = round(random.uniform(10, 500), 2)
        rows.append(
            {
                "sale_id": i,
                "date": sale_date.strftime("%Y-%m-%d"),
                "product": random.choice(PRODUCTS),
                "quantity": qty,
                "unit_price": unit_price,
                "total": round(qty * unit_price, 2),
                "region": random.choice(["North", "South", "East", "West"]),
            }
        )
    return rows


def write_sales_csv(rows: list[dict]) -> Path:
    path = DATA_DIR / "sales.csv"
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)
    return path


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def generate_all() -> None:
    """Generate every mock dataset and write to disk."""
    ensure_dir()

    employees = generate_employees(50)
    emp_csv = write_employees_csv(employees)
    emp_json = write_employees_json(employees)
    print(f"[OK] Employees CSV  -> {emp_csv}")
    print(f"[OK] Employees JSON -> {emp_json}")

    sales = generate_sales(200)
    sales_csv = write_sales_csv(sales)
    print(f"[OK] Sales CSV      -> {sales_csv}")

    print("\nAll mock data generated successfully.")


if __name__ == "__main__":
    generate_all()
