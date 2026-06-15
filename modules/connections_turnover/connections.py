import os
import glob
from collections import defaultdict
import pandas as pd
from managers import MANAGER_REFS

IGNORE_REFS = {
    "maxat",
    "wramaccount1",
    "wramaccount3",
    "wramaccount5",
    "amarendra",
    "maintenance2",
    "maintenance",
}


def load_connections_from_file(filepath: str) -> dict[str, int]:
    ref_counts = defaultdict(int)
    df = pd.read_excel(filepath, sheet_name=0, header=None)

    for _, row in df.iterrows():
        if len(row) < 3:
            continue
        ref = row.iloc[1]
        count = row.iloc[2]
        if not isinstance(ref, str):
            continue
        ref = ref.lower()  # переводим в lowercase
        if ref in IGNORE_REFS:
            continue
        ref_counts[ref] += int(count) if pd.notna(count) else 0

    return dict(ref_counts)


def load_all_connections(input_dir: str) -> dict[str, int]:
    pattern = os.path.join(input_dir, "merchant_turnovers_*.xlsx")
    files = glob.glob(pattern)

    total = defaultdict(int)
    for file in files:
        file_counts = load_connections_from_file(file)
        for ref, cnt in file_counts.items():
            total[ref] += cnt

    return dict(total)


def get_connections_for_managers(connections: dict[str, int]) -> list[int]:
    # MANAGER_REFS тоже переводим в lowercase для поиска
    return [connections.get(ref.lower(), 0) for ref in MANAGER_REFS]


def run(file_path: str) -> dict[str, int]:
    connections = load_connections_from_file(file_path)
    return connections