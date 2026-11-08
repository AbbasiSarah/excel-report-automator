"""Run with: python -m pytest  (or just: python test_automator.py)"""
from pathlib import Path
import pandas as pd
from automate_report import load_and_clean, parse_date, standard_columns


def test_dates_in_any_format():
    assert parse_date("2026-09-14") == pd.Timestamp(2026, 9, 14)
    assert parse_date("14/09/2026") == pd.Timestamp(2026, 9, 14)
    assert parse_date("14 Sep 2026") == pd.Timestamp(2026, 9, 14)
    assert pd.isna(parse_date("not a date"))


def test_column_names_are_standardised():
    df = pd.DataFrame(columns=["order date", "ITEM", "Quantity", "price", "area"])
    assert list(standard_columns(df).columns) == ["date", "product", "qty", "price", "region"]


def test_cleaning_on_sample_data():
    data, issues, stats = load_and_clean(Path(__file__).parent / "input")
    assert stats["files"] == 9
    assert (data["qty"] > 0).all()                    # no missing or negative quantities left
    assert data["price"].notna().all()                # every price filled
    assert not data.duplicated(["date", "product", "qty", "price", "branch"]).any()
    assert set(data["product"]) == {"Laptop", "Phone", "Headphones", "Keyboard", "Monitor"}
    assert len(issues) == stats["fixes"] > 0          # every fix is logged


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn(); print("ok", name)
