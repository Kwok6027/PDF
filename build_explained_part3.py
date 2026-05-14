"""
================================================================================
BUILD_EXCEL_EXPLAINED.py — Part 3 of 3
================================================================================
This part focuses on the heart of the smile-aware pricer — the per-row
formulas in DailyFixings — and the main() entry point.

The other sheet builders (YieldCurve, DailyDFs, PeriodSummary, Valuation,
README) follow simple "write data + formulas" patterns; see build.py for
the full code.
================================================================================
"""

# ════════════════════════════════════════════════════════════════════════════
# PART 3.1 — The DailyFixings inner loop (the most important bit)
# ════════════════════════════════════════════════════════════════════════════
#
# DailyFixings has ONE ROW per calendar day in Leg 3 (~1640 days). For each
# row we compute:
#
#   Column   Content (Excel formula references cells in the same row)
#   ------   ------------------------------------------------------------
#   A        Index (1, 2, 3, ...)
#   B        Date d
#   C        Year fraction T = (d - val_date)/365.25
#   D        DF(d)        — discount factor at d, looked up from DailyDFs
#   E..N     DF(d+1Y) .. DF(d+10Y)   — 10 annual DF lookups for the swap
#   O..X     τ_1 .. τ_10  — ACT/360 day fractions for each swap period
#   Y        Annuity = Σ τ_i · DF(d+iY)
#   Z        Forward 10Y CMS = (DF(d) - DF(d+10Y)) / Annuity
#   AA       ATM σ at this T (interpolated from MarketData)
#   AB,AC    ρ, ν at this T (interpolated)
#   AD..AF   α calibration: AD = α₀;  AE = iter1;  AF = iter2 (final)
#   AG       σ at K  — SABR vol evaluated at strike = barrier
#   AH       d2     = (ln(F/K) - 0.5 σ² T) / (σ √T)
#   AI       P naive  = 1 - Φ(d2) = Φ(-d2)
#   --- put-spread replication (the smile correction) ---
#   AJ       σ at K-δ   (SABR vol at slightly lower strike)
#   AK       σ at K+δ   (SABR vol at slightly higher strike)
#   AL       ∂σ/∂K = (σ(K+δ) - σ(K-δ)) / (2δ)
#   AM       n(d2) = exp(-d2²/2) / √(2π)   — standard normal PDF at d2
#   AN       smile correction = K · n(d2) · √T · ∂σ/∂K
#   AO       P_smile = P_naive + correction
#
# WHY the smile correction? Naive Black uses σ at the single strike K only,
# but a digital option's true value depends on how σ varies WITH K (the
# smile). The mathematical identity is:
#
#     P(S ≤ K)  =  N(-d2)  +  Vega · ∂σ/∂K
#
# This is equivalent to pricing the digital as a small put-spread:
#     digital_put(K)  ≈  [P_BS(K+δ, σ(K+δ)) - P_BS(K-δ, σ(K-δ))] / (2δ)
# in the limit δ → 0.  We use δ = 1bp on MarketData, which is small enough
# that the finite-difference is indistinguishable from the analytical limit.
#
# ════════════════════════════════════════════════════════════════════════════

# Pseudo-code for the per-row loop. The real code is in build.py;
# the part below shows JUST the smile-correction columns, with annotations.
"""
# Inside the loop that runs for each day index r = 5, 6, 7, ..., 5+1640:

# 1) σ at K-δ  -- evaluate SABR formula with K replaced by (K - δ).
#    The string concatenation `({K_cell}-{delta_cell})` builds an Excel
#    expression like (MarketData!$B$8 - MarketData!$B$9).
setcell(ws, r, 36, sabr_vol_at_K_formula(
            f_cell      = f"Z{r}",                          # forward
            K_cell      = f"({K_cell}-{delta_cell})",       # K - δ
            T_cell      = f"C{r}",                          # year frac
            alpha_cell  = f"AF{r}",                         # calibrated α
            beta_cell   = beta_cell,
            rho_cell    = f"AB{r}",
            nu_cell     = f"AC{r}"),
        FORMULA_FONT, fmt='0.0000%')

# 2) σ at K+δ  -- same SABR formula, K + δ
setcell(ws, r, 37, sabr_vol_at_K_formula(
            f_cell=f"Z{r}", K_cell=f"({K_cell}+{delta_cell})",
            T_cell=f"C{r}", alpha_cell=f"AF{r}",
            beta_cell=beta_cell, rho_cell=f"AB{r}", nu_cell=f"AC{r}"),
        FORMULA_FONT, fmt='0.0000%')

# 3) Smile slope: central finite difference. Pure formula, no shortcuts.
setcell(ws, r, 38, f"=(AK{r}-AJ{r})/(2*{delta_cell})",
        FORMULA_FONT, fmt='0.0000')

# 4) n(d2) — standard normal PDF.
#    IMPORTANT: Excel parses '-d²' weirdly because of operator precedence
#    (unary minus vs. exponentiation). Writing `-AH{r}^2` would actually
#    compute (-AH)^2 = AH^2, giving the WRONG sign in the exponent.
#    We force explicit multiplication: -0.5 * AH * AH.
setcell(ws, r, 39, f"=EXP(-0.5*AH{r}*AH{r})/SQRT(2*PI())",
        FORMULA_FONT, fmt='0.0000')

# 5) Smile correction = K · n(d2) · √T · slope
setcell(ws, r, 40, f"={K_cell}*AM{r}*SQRT(C{r})*AL{r}",
        FORMULA_FONT, fmt='0.0000%')

# 6) Smile-adjusted probability  P_smile = P_naive + correction
setcell(ws, r, 41, f"=AI{r}+AN{r}", FORMULA_FONT, fmt='0.0000%')
"""


# ════════════════════════════════════════════════════════════════════════════
# PART 3.2 — PeriodSummary aggregation
# ════════════════════════════════════════════════════════════════════════════
# For each of the 18 quarterly accrual periods:
#
#   Accrual %  =  AVERAGE of the daily P_smile (column AO of DailyFixings)
#                 over all calendar days in that period.
#
#   Period PV  =  Notional × Coupon × τ_30/360(quarter)
#                          × Accrual%  ×  DF(pay date)
#
# Because each quarter is ≈ 0.25 years, τ ≈ 0.25 directly.
# DF(pay date) is looked up from DailyDFs via INDEX.
#
# Total PV = SUM of all 18 period PVs.
#
# Pseudo-code:
"""
for each (period_start, period_end) in 18 quarterly periods:
    idx_start = (period_start - leg3_start).days + 1
    idx_end   = (period_end   - leg3_start).days
    first_row = idx_start + 4       # DailyFixings row offset
    last_row  = idx_end   + 4

    # Accrual % — references the smile-adjusted column AO
    setcell(ws, r, 8,
            f"=AVERAGE(DailyFixings!$AO${first_row}:$AO${last_row})",
            FORMULA_FONT, fmt='0.00%')

    # DF at pay date — INDEX into DailyDFs by date offset
    setcell(ws, r, 9,
            f"=INDEX(DailyDFs!$C:$C, C{r}-{val_cell}+5)",
            FORMULA_FONT, fmt='0.000000')

    # PV
    setcell(ws, r, 10,
            f"={notional}*{coupon}*0.25*H{r}*I{r}",
            FORMULA_FONT, fmt='#,##0.00')

# Total at the bottom
setcell(ws, total_row, 10, f"=SUM(J{first}:J{last})",
        HEADER_FONT, fmt='#,##0.00')
"""


# ════════════════════════════════════════════════════════════════════════════
# PART 3.3 — Main entry point
# ════════════════════════════════════════════════════════════════════════════
# This is what the script actually does when you run `python build.py`.
# `if __name__ == '__main__':` is a Python idiom — the block runs only
# when the file is executed directly, not when imported as a module.

def main():
    """
    Usage:  python build.py [may7|apr21|both]
    Default: both.
    """
    # sys.argv is the list of command-line arguments. argv[0] is the
    # script name; argv[1] is the first user-provided argument.
    import sys
    arg = sys.argv[1] if len(sys.argv) > 1 else 'both'

    # Match argument to which builder(s) to run.
    if arg in ('may7', 'both'):
        build_workbook(make_data_may7(),
                       'output/RangeAccrual_May7_2026.xlsx')
    if arg in ('apr21', 'both'):
        build_workbook(make_data_apr21(),
                       'output/RangeAccrual_Apr21_2026.xlsx')


def build_workbook(data, path):
    """
    Top-level orchestrator. Creates the workbook, calls each sheet
    builder in the right order (MarketData must come first because
    others reference its cells), then saves.
    """
    wb = Workbook()
    # Workbook() creates an empty workbook with one default sheet "Sheet".
    # Remove it — we'll add our own sheets in the right order.
    wb.remove(wb.active)

    # Build sheets. ORDER MATTERS — later sheets reference earlier ones
    # via the cell refs stashed in `data['_md_xxx_cell']`.
    build_market_data(wb, data)
    build_yield_curve(wb, data)
    build_daily_dfs(wb, data)
    build_daily_fixings(wb, data)
    build_period_summary(wb, data)
    build_valuation(wb, data)
    build_readme(wb, data)

    # Create the output directory if it doesn't exist.
    # `os.path.dirname(path)` extracts the directory part of a path.
    import os
    os.makedirs(os.path.dirname(path), exist_ok=True)

    wb.save(path)
    print(f"Saved {path}")


# Standard Python idiom — run main() only when this script is executed
# directly. If someone does `from build import setcell`, main() does NOT run.
if __name__ == '__main__':
    main()


# ════════════════════════════════════════════════════════════════════════════
# RUNNING THE SCRIPT
# ════════════════════════════════════════════════════════════════════════════
# From a terminal:
#   pip install openpyxl                       # one-time install
#   python build.py may7                       # writes May 7 workbook
#   python build.py apr21                      # writes Apr 21 workbook
#   python build.py both                       # writes both (default)
#
# Output goes to ./output/.
#
# After it runs, open the .xlsx file in Excel (or LibreOffice) and
# Excel evaluates all the formulas. The PV total is on PeriodSummary,
# row 24, column J.
#
# ════════════════════════════════════════════════════════════════════════════
# DEBUGGING TIPS
# ════════════════════════════════════════════════════════════════════════════
#
# 1.  If LibreOffice doesn't recalc on open, use:
#         soffice --headless --calc --convert-to xlsx file.xlsx
#     (or use the recalc.py utility in the project root.)
#
# 2.  To inspect a cell's formula (string) vs. its evaluated value:
#         from openpyxl import load_workbook
#         wb_f = load_workbook('file.xlsx')              # formulas
#         wb_v = load_workbook('file.xlsx', data_only=True)  # values
#         ws_f = wb_f['DailyFixings']
#         ws_v = wb_v['DailyFixings']
#         print(ws_f.cell(row=5, column=40).value)  # the formula string
#         print(ws_v.cell(row=5, column=40).value)  # the evaluated number
#
# 3.  Common pitfalls:
#     * Insert a row in MarketData → all hard-coded "MarketData!$B$N"
#       references break. Cure: capture row at write time, as we do.
#     * `=EXP(-x^2/2)` parses as `EXP((-x)^2/2)` due to Excel operator
#       precedence. Use `EXP(-0.5*x*x)`.
#     * Cosmetic text starting with "=" gets treated as a formula. Use
#       a leading space or apostrophe if you need a literal "=".
#
# ════════════════════════════════════════════════════════════════════════════
