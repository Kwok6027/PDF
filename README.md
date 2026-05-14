# Range-Accrual Pricer — Python Code Walkthrough

This folder contains the Python that builds the two Excel workbooks
(`RangeAccrual_May7_2026.xlsx` and `RangeAccrual_Apr21_2026.xlsx`).

## Files in reading order

| File | Purpose |
|------|---------|
| `build_explained_part1.py` | Imports, helpers, market-data dicts |
| `build_explained_part2.py` | SABR formula builders, MarketData sheet builder |
| `build_explained_part3.py` | DailyFixings smile-correction core, main() |
| `build.py` | **The complete working file** (~1,100 lines). Run this. |

The three `_explained_` files are **commentary** — they walk through the
key ideas with heavy annotations. They cover ~70% of the code (the
interesting parts); for things like every PeriodSummary column or the
README sheet generation, read the corresponding section of `build.py`.

If you only want to **run** the model, ignore the explained files and
just use `build.py`. If you want to **understand** it, read the
explained files in order, then jump into `build.py` for the bits not
covered.

## Running

```bash
pip install openpyxl                # one-time install
python build.py may7                # writes ./output/RangeAccrual_May7_2026.xlsx
python build.py apr21               # writes ./output/RangeAccrual_Apr21_2026.xlsx
python build.py both                # writes both (default)
```

After it runs, open the `.xlsx` file in Excel (or LibreOffice) and the
formulas evaluate automatically. The PV total is on `PeriodSummary`,
row 24, column J.

## The big-picture flow

```
make_data_may7()                  ← Python dict of market inputs
        │
        ▼
build_workbook(data, path)        ← orchestrator
        │
        ├── build_market_data()   ← MarketData sheet (blue inputs only)
        ├── build_yield_curve()   ← YieldCurve sheet (bootstrap as formulas)
        ├── build_daily_dfs()     ← DailyDFs sheet (~5,500 DFs by date)
        ├── build_daily_fixings() ← DailyFixings (~1,640 days × 41 cols)
        ├── build_period_summary()← PeriodSummary (18 quarterly periods)
        ├── build_valuation()     ← Valuation (step-by-step walkthrough)
        └── build_readme()        ← README sheet
        │
        ▼
        wb.save(path)             ← write .xlsx to disk
```

Each `build_*` function writes Excel **formulas as strings** — not values.
Excel does the arithmetic when the workbook opens.

## Key Python concepts you'll see

- **Function definition** — `def foo(x): ...`. Body is indented.
- **f-strings** — `f"...{x}..."` substitutes the value of `x`.
- **Lists** — `[1, 2, 3]`. Iterable, indexable.
- **Tuples** — `(1, 2)`. Immutable list. Often used for fixed-size records.
- **Dicts** — `{'key': value}`. Lookup by key.
- **For loop** — `for x in collection: ...` — runs once per item.
- **`*` in signature** — args after the lone `*` must be passed by name.
- **`**dict`** — "spread" a dict into another (used to merge trade terms).
- **`if __name__ == '__main__':`** — runs only when the file is executed
  directly, not when imported as a module.

## Key financial concepts

- **SOFR curve bootstrap** — sequential: overnight rate → FOMC steps →
  3M futures (conv-adjusted) → swap rates. Each new instrument pins
  down the DF at one new pillar date.
- **SABR model** — gives implied vol σ(F, K, T) given parameters
  α, β, ρ, ν. We use the Hagan 2002 expansion. α is calibrated to
  match σ_ATM via Newton iteration.
- **Forward 10Y CMS** — for each daily fixing date d, computed from
  10 annual DFs: `fwd = (DF(d) - DF(d+10Y)) / Σ τ_i · DF(d+iY)`.
- **Naive digital probability** — `P(S ≤ K) = N(-d2)` using σ at K only.
- **Smile-aware digital (put-spread replication)** — adds the smile slope:
  `P(S ≤ K) = N(-d2) + K · n(d2) · √T · ∂σ/∂K`.
  Equivalent to the limit δ→0 of the put-spread
  `[P_BS(K+δ, σ(K+δ)) - P_BS(K-δ, σ(K-δ))] / (2δ)`.

## Debugging tips

1. To inspect a formula vs its evaluated value:
   ```python
   from openpyxl import load_workbook
   wb_f = load_workbook('file.xlsx')                    # formulas
   wb_v = load_workbook('file.xlsx', data_only=True)    # values
   print(wb_f['DailyFixings'].cell(row=5, column=40).value)   # the formula
   print(wb_v['DailyFixings'].cell(row=5, column=40).value)   # the number
   ```

2. If LibreOffice doesn't recalc on save:
   ```bash
   soffice --headless --calc --convert-to xlsx file.xlsx
   ```

3. **Two classic Excel/LibreOffice gotchas:**
   - `=EXP(-x^2/2)` is parsed as `EXP((-x)^2/2)`. Use `EXP(-0.5*x*x)`.
   - Inserting a row in MarketData breaks any hard-coded
     `MarketData!$B$N` reference. Solution: capture row numbers at
     write time and use f-strings to build the references.

## Where to look in `build.py` for each topic

| Topic | Approximate line range |
|-------|------------------------|
| Imports, fonts, fills, helpers | 1–30 |
| Market-data dicts (Apr 21, May 7) | 30–135 |
| `setcell`, date helpers | 135–155 |
| `build_readme` | 158–205 |
| `build_market_data` (writes MarketData sheet) | 207–345 |
| `build_yield_curve` (bootstrap as formulas) | 347–585 |
| `build_daily_dfs` (~5,500 DFs by date) | 588–635 |
| `build_daily_fixings` (per-row, incl. smile) | 638–820 |
| `atm_interp_formula`, SABR formula builders | 822–865 |
| `build_period_summary` (18 quarter aggregations) | 868–950 |
| `build_valuation` (step-by-step walkthrough) | 954–1070 |
| `build()` entry point | 1074–end |
