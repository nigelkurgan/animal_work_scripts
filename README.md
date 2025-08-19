# mouse experiment scripts

This repository contains scripts for all aspects of in vivo experiments and data analysis. Below will have a summary of each script and how to use it. 

## Scripts

### [`allocate_mice_equal.py`](allocate_mice_equal.py)

Allocates mice into N groups with exactly equal numbers per group, balancing on one or more numeric columns (e.g., body weight, age). Outputs an Excel file with group assignments and summary statistics.

- **Input:** Excel file with mouse data and columns to balance.
- **Output:** Excel file with group assignments and summary statistics.
- **Usage Example:**
  ```sh
  python3 allocate_mice_equal.py --input mice.xlsx --output mice_allocated.xlsx --groups ctrl treatment --balance_cols body_weight age --seed 42
  ```

---

Add a summary for each new script you include in