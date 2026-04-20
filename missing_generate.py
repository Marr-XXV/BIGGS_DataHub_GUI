"""Missing data generator: finds gaps in record2025.csv and re-fetches them.

Connections:
- Uses pandasbiggs.CSVProcessor to inspect record2025.csv and pivot quantities by date/branch/POS.
- Builds a branches_missing structure and calls Receive.missing_fetch to download only missing days.
- After re-fetch, Combiner.generate (invoked elsewhere) can rebuild record2025.csv so analytics are up to date.
"""
import datetime
import os
import pprint
import shutil

import pandas as pd

from pandasbiggs import CSVProcessor
from fetcher import Receive


def clean(directory):
    """Remove all files and subdirectories under the given directory path.

    Used to clear the staging folders (latest/ and temp/) before recomputing
    which records are missing and re-fetching them.
    """
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print("Failed to delete %s. Reason: %s" % (file_path, e))


def get_date_input(prompt):
    """Prompt the user for a YYYY-MM-DD date string and return a datetime."""
    while True:
        try:
            user_input = input(prompt)
            date_value = datetime.datetime.strptime(user_input, "%Y-%m-%d")
            return date_value
        except ValueError:
            print("❌ Invalid format. Please use YYYY-MM-DD.")


def main():
    """Run the missing-data analysis and trigger re-fetch for gaps only."""
    parent_dir = os.getcwd()

    # Clear staging folders so only fresh latest/temp files are considered
    clean(os.path.join(parent_dir, "latest"))
    clean(os.path.join(parent_dir, "temp"))

    csv2023 = CSVProcessor("record2025.csv")

    value = "Amount"
# dateStart = input("Enter start date (YYYY-MM-DD): ")
# dateEnd = input("Enter end date (YYYY-MM-DD): ")
    # Get start and end dates interactively
    date_start = get_date_input("Enter start date (YYYY-MM-DD): ")
    date_end = get_date_input("Enter end date (YYYY-MM-DD): ")

    if date_start > date_end:
        print("❌ Start date must be before End date.")
        return

    date_start_adjusted = date_start.strftime("%Y-%m-%d")
    date_end_str = date_end.strftime("%Y-%m-%d")

    print("✅ Start date (adjusted):", date_start_adjusted)
    print("✅ End date:", date_end_str)

    filtered = csv2023.filter("", "", "", date_start_adjusted, date_end_str)

    # Aggregate quantity by (date, branch, POS) so we can see where zeros appear
    dates = filtered.pivot_table(
        index=["DATE"],
        columns=["BRANCH", "POS"],
        values=["QUANTITY"],
        aggfunc=lambda x: sum(x),
        fill_value=0,
        dropna=False,
    )
    csv_dates = dates.reset_index()
    csv_dates.to_csv("Missing_dates.csv", index=False)

    # Build full set of expected (branch, pos, date) combinations
    branches = []
    with open(os.path.join(parent_dir, "settings", "branches.txt"), "r") as f:
        for row in f.read().splitlines():
            if row.strip():
                branches.append(row.strip())

    expected = set()
    current_date = date_start
    while current_date <= date_end:
        for branch in branches:
            for pos in ["1", "2"]:
                expected.add((branch, pos, current_date.strftime("%Y-%m-%d")))
        current_date += datetime.timedelta(days=1)

    # Build existing set of (branch,pos,date) that have non-zero quantity
    existing = set()
    for d in dates.index:
        date_str = d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d).split()[0]
        for (_, branch, pos) in dates.columns:
            if dates.loc[d, ("QUANTITY", branch, pos)] > 0:
                existing.add((branch, str(pos), date_str))

    # Missing = Expected - Existing
    missing = expected - existing

    # Organize missing into dict by branch -> pos -> [dates]
    branches_missing = {}
    for branch, pos, mdate in missing:
        pos_int = int(pos)
        if branch not in branches_missing:
            branches_missing[branch] = {}
        if pos_int not in branches_missing[branch]:
            branches_missing[branch][pos_int] = []
        branches_missing[branch][pos_int].append(mdate)

    print("Missing structure:")
    pprint.pprint(branches_missing)

    # Trigger missing-data fetch through the existing receiver API
    rep = Receive(date_start, date_end, dates["QUANTITY"])
    rep.missing_fetch(branches_missing)

    print("✅ Missing data fetch complete.")


if __name__ == "__main__":
    main()
