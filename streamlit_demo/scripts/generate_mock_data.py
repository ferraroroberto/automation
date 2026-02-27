"""
Mock data generator for the Streamlit Capabilities Demo.

Generates CSV and JSON files used by demo pages.
All data is synthetic and deterministic (seeded random).
"""

import csv
import json
import logging
import random
from datetime import datetime, timedelta
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger(__name__)

MOCK_DATA_DIR = Path(__file__).resolve().parent.parent / "data" / "mock_data"

DEPARTMENTS = ["Engineering", "Marketing", "Sales", "HR", "Finance", "Operations"]
STATUSES = ["Active", "On Leave", "Resigned"]
PRODUCT_CATEGORIES = ["Electronics", "Books", "Clothing", "Home & Garden", "Sports"]


def _seed():
    random.seed(42)


def generate_employees_csv(n: int = 50) -> Path:
    """Generate a CSV file with mock employee records."""
    _seed()
    filepath = MOCK_DATA_DIR / "employees.csv"
    fieldnames = ["id", "name", "department", "salary", "hire_date", "status"]

    first_names = [
        "Alice", "Bob", "Charlie", "Diana", "Eve", "Frank",
        "Grace", "Hank", "Ivy", "Jack", "Karen", "Leo",
        "Mona", "Nate", "Olivia", "Paul", "Quinn", "Rita",
        "Steve", "Tina", "Uma", "Vic", "Wendy", "Xander",
        "Yara", "Zane",
    ]
    last_names = [
        "Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia",
        "Miller", "Davis", "Rodriguez", "Martinez", "Hernandez", "Lopez",
        "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore",
    ]

    rows = []
    base_date = datetime(2018, 1, 1)
    for i in range(1, n + 1):
        hire_date = base_date + timedelta(days=random.randint(0, 2000))
        rows.append({
            "id": i,
            "name": f"{random.choice(first_names)} {random.choice(last_names)}",
            "department": random.choice(DEPARTMENTS),
            "salary": round(random.uniform(40000, 150000), 2),
            "hire_date": hire_date.strftime("%Y-%m-%d"),
            "status": random.choices(STATUSES, weights=[0.8, 0.1, 0.1])[0],
        })

    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    logger.info("ℹ️  Generated %s (%d records)", filepath.name, n)
    return filepath


def generate_sales_json(n: int = 200) -> Path:
    """Generate a JSON file with mock sales transactions."""
    _seed()
    filepath = MOCK_DATA_DIR / "sales.json"

    records = []
    base_date = datetime(2024, 1, 1)
    for i in range(1, n + 1):
        sale_date = base_date + timedelta(days=random.randint(0, 364))
        records.append({
            "transaction_id": f"TXN-{i:05d}",
            "date": sale_date.strftime("%Y-%m-%d"),
            "category": random.choice(PRODUCT_CATEGORIES),
            "amount": round(random.uniform(5.0, 500.0), 2),
            "quantity": random.randint(1, 20),
            "region": random.choice(["North", "South", "East", "West"]),
        })

    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2)

    logger.info("ℹ️  Generated %s (%d records)", filepath.name, n)
    return filepath


def generate_timeseries_csv(days: int = 90) -> Path:
    """Generate a CSV with daily metrics for charting demos."""
    _seed()
    filepath = MOCK_DATA_DIR / "timeseries.csv"
    fieldnames = ["date", "visitors", "signups", "revenue"]

    rows = []
    base_date = datetime(2024, 7, 1)
    visitors = 1000
    for d in range(days):
        visitors += random.randint(-50, 70)
        visitors = max(200, visitors)
        signups = int(visitors * random.uniform(0.02, 0.08))
        revenue = round(signups * random.uniform(20, 80), 2)
        rows.append({
            "date": (base_date + timedelta(days=d)).strftime("%Y-%m-%d"),
            "visitors": visitors,
            "signups": signups,
            "revenue": revenue,
        })

    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    logger.info("ℹ️  Generated %s (%d days)", filepath.name, days)
    return filepath


def main():
    MOCK_DATA_DIR.mkdir(parents=True, exist_ok=True)
    logger.info("ℹ️  Output directory: %s", MOCK_DATA_DIR)
    generate_employees_csv()
    generate_sales_json()
    generate_timeseries_csv()
    logger.info("ℹ️  Mock data generation complete.")


if __name__ == "__main__":
    main()
