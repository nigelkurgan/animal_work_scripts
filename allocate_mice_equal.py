#!/usr/bin/env python3
"""
allocate_mice_equal.py
----------------------
Allocate mice into N groups (e.g., ctrl, igsf9) with STRICT equal N:
- N must be divisible by the number of groups.
- Balancing by one or more specified columns (e.g., body_weight, age).
- Writes updated Excel with allocations and summary stats.

SETUP & USAGE
-------------

1. Download this script:
    Save this file as `allocate_mice_equal.py` in a folder of your choice.

2. Prepare your data:
    - Create an Excel file (e.g., `mice.xlsx`) with at least two columns:
        - `mouse_id` (unique identifier for each mouse)
        - Any numeric columns you want to balance (e.g., `body_weight`, `age`)
    - Place this Excel file in the same folder as the script.

3. (Optional) Create a new folder for your project:
    mkdir mice_allocation
    cd mice_allocation
    # Place your Excel file and this script here

4. Install required Python packages (if not already installed):
    pip3 install pandas numpy openpyxl xlsxwriter

5. Run the script from the terminal:
    python3 allocate_mice_equal.py --input mice.xlsx --output mice_allocated.xlsx --groups ctrl treatment --balance_cols body_weight age --seed 42

    - Replace `ctrl treatment` with your desired group names (as many as you want, e.g., --groups A B C).
    - Replace `body_weight age` with one or more columns to balance (default: body_weight).
    - The number of mice must be divisible by the number of groups.

6. Output:
    - The script will create a new Excel file (e.g., `mice_allocated.xlsx`) with:
        - Sheet "mice_allocated": Each mouse and its assigned group
        - Sheet "summary_by_group": Group statistics for each balance column
        - Sheet "allocation_qc": QC metrics

Example:
    python3 allocate_mice_equal.py --input mice.xlsx --output mice_allocated.xlsx --groups ctrl treatment --balance_cols body_weight age --seed 42

"""

import argparse
from pathlib import Path
import sys
import pandas as pd
import numpy as np

def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--input", required=True, help="Input Excel (.xlsx) with mouse_id and columns to balance")
    p.add_argument("--sheet", default=None, help="Sheet name (default: first sheet)")
    p.add_argument("--output", required=True, help="Output Excel (.xlsx)")
    p.add_argument("--groups", nargs="+", required=True, help="Names for the groups (e.g. ctrl treatment)")
    p.add_argument("--balance_cols", nargs="+", default=["body_weight"],
                   help="Column(s) to balance across groups (default: body_weight). Can specify multiple columns.")
    p.add_argument("--seed", type=int, default=123, help="Random seed for deterministic tie-breaking")
    return p.parse_args()

def validate_df(df: pd.DataFrame, balance_cols) -> pd.DataFrame:
    required = ["mouse_id"] + balance_cols
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise SystemExit(f"Missing required column(s): {missing}. Found: {list(df.columns)}")
    for col in balance_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")
        if df[col].isna().any():
            bad = df.loc[df[col].isna(), ["mouse_id", col]]
            raise SystemExit(f"Found invalid values in {col}:\n" + bad.to_string(index=False))
    if "group" not in df.columns:
        df["group"] = pd.NA
    return df

def assign_round_robin(indices, target_sizes, rng):
    bins = [[] for _ in target_sizes]
    ptr = 0
    counts = [0]*len(target_sizes)
    for i in indices:
        start = ptr
        while counts[ptr] >= target_sizes[ptr]:
            ptr = (ptr + 1) % len(target_sizes)
            if ptr == start:
                raise RuntimeError("No bin has capacity left while assigning.")
        bins[ptr].append(i)
        counts[ptr] += 1
        ptr = (ptr + 1) % len(target_sizes)
    assignment = {}
    for b, idxs in enumerate(bins):
        for j in idxs:
            assignment[j] = b
    return assignment, bins

def balanced_allocation_equal(df: pd.DataFrame, group_names, balance_cols, seed=123) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    n = len(df)
    n_groups = len(group_names)
    if n % n_groups != 0:
        raise SystemExit(f"N={n} is not divisible by number of groups ({n_groups}). Equal allocation requires N%groups==0.")
    n_per_group = n // n_groups

    # Sort by balance_cols (descending), with jitter for tie-breaking
    jitter = rng.uniform(0, 1e-9, size=n)
    df["_jitter"] = jitter
    sort_cols = balance_cols + ["_jitter"]
    ascending = [False]*len(balance_cols) + [False]
    sorted_idx = df.sort_values(by=sort_cols, ascending=ascending).index.tolist()
    df = df.drop(columns="_jitter")

    # Assign groups (bins, exact sizes)
    group_assign, group_bins = assign_round_robin(sorted_idx, [n_per_group]*n_groups, rng)
    group_name_map = {i: group_names[i] for i in range(n_groups)}
    df.loc[:, "group"] = df.index.map(lambda i: group_name_map[group_assign[i]])

    # Final sanity checks
    for g in group_names:
        cnt = len(df[df["group"] == g])
        assert cnt == n_per_group, f"{g} size mismatch (got {cnt}, expected {n_per_group})"

    return df

def summarize(df: pd.DataFrame, balance_cols):
    summary_g = df.groupby("group")[balance_cols].agg(['count', 'mean', 'std', 'min', 'max'])
    # Flatten columns
    summary_g.columns = ['_'.join([col, stat]) for col, stat in summary_g.columns]
    summary_g = summary_g.reset_index()
    delta_means = {}
    for col in balance_cols:
        means = df.groupby("group")[col].mean()
        delta_means[col] = float(means.max() - means.min()) if len(means) else float("nan")
    return summary_g, delta_means

def main():
    args = parse_args()
    inp = Path(args.input)
    outp = Path(args.output)

    df = pd.read_excel(inp, sheet_name=(args.sheet or 0))
    if not isinstance(df, pd.DataFrame):
        first_key = next(iter(df))
        df = df[first_key]
    df.columns = df.columns.str.strip()

    df = validate_df(df, args.balance_cols)

    df_alloc = balanced_allocation_equal(df, args.groups, args.balance_cols, seed=args.seed)

    summary_g, delta_means = summarize(df_alloc, args.balance_cols)

    with pd.ExcelWriter(outp, engine="xlsxwriter") as xw:
        df_alloc.to_excel(xw, index=False, sheet_name="mice_allocated")
        summary_g.to_excel(xw, index=False, sheet_name="summary_by_group")
        meta = pd.DataFrame({
            "metric": [f"delta_mean_{col}" for col in args.balance_cols] + ["n_total"],
            "value": list(delta_means.values()) + [len(df_alloc)]
        })
        meta.to_excel(xw, index=False, sheet_name="allocation_qc")

    print(f"Saved allocation to: {outp}")
    print("\nSummary by group:")
    print(summary_g.to_string(index=False))
    for col in args.balance_cols:
        print(f"\nΔ mean across groups for {col}: {delta_means[col]:.4f}")

if __name__ == "__main__":
    main()
