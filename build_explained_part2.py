"""
================================================================================
BUILD_EXCEL_EXPLAINED.py — Part 2 of 3
================================================================================
This part has:
  1.4  SABR formula builders (return STRINGS of Excel formulas)
  1.5  Sheet builders: build_market_data, build_yield_curve,
       build_daily_dfs, build_daily_fixings, build_period_summary,
       build_valuation, build_readme

These functions take a `data` dict (from Part 1) and a Workbook, then
write rows into the appropriate sheet.

Why "formula builders"?
-----------------------
The SABR vol formula is long and ugly. We want EACH ROW of DailyFixings
to call it on the row's own forward/expiry/alpha. Rather than rewrite
the whole formula 1,640 times, we build it ONCE as a string with cell
references plugged in, then substitute different references per row.

Example: sabr_vol_at_K_formula("Z5", "...K_cell...", "C5", "AF5", ...)
returns a string like
    "=(1/((Z5*K_cell)^((1-beta)/2)*(1+...)))*AF5*(z/x(z))*(1+...)"
which we paste into one DailyFixings cell. We call it with "Z6", "AF6",
etc. for the next row.
================================================================================
"""

# Continue from Part 1 — assume the helpers and styles are imported.
from build_explained_part1 import (
    setcell, add_years, third_wednesday,
    daycount_30_360, daycount_act_360, get_column_letter, Workbook, date,
    INPUT_FONT, FORMULA_FONT, HEADER_FONT, SUBTITLE_FONT, TITLE_FONT,
    CROSS_FONT, INPUT_FILL, SECTION_FILL, CENTER_ALIGN, LEFT_ALIGN, RIGHT_ALIGN,
)


# ═══════════════════════════════════════════════════════════════════════════
# PART 2.1 — SABR FORMULA BUILDERS
# ═══════════════════════════════════════════════════════════════════════════
# These return Python strings that ARE Excel formulas. The strings get
# pasted into worksheet cells.
#
# SABR ("Stochastic Alpha-Beta-Rho") is a model for the implied vol smile.
# Given forward F, strike K, expiry T, and parameters α, β, ρ, ν, it gives
# the Black-implied vol σ(F, K, T). We use the Hagan 2002 expansion.


def sabr_vol_at_K_formula(f_cell, K_cell, T_cell, alpha_cell,
                           beta_cell, rho_cell, nu_cell):
    """
    Hagan 2002 lognormal SABR vol at strike K, with the ATM branch
    handled by an IF. Returns the formula STRING.

    Notes for non-Python readers:
      - f-strings: f"...{x}..." substitutes Python variables.
      - The returned string starts with "=", which is how Excel
        recognises it as a formula.

    The formula has three pieces (Hagan 2002, equation 2.17a):
      pre  =  α / ((F·K)^((1-β)/2) · [1 + (1-β)²/24·ln²(F/K) + ...])
      z    =  (ν/α) · (F·K)^((1-β)/2) · ln(F/K)
      x(z) =  ln( (√(1-2ρz+z²) + z - ρ) / (1-ρ) )
      cor  =  1 + T·[ (1-β)²/24·α²/(F·K)^(1-β)  +  ρβνα/(4·(F·K)^((1-β)/2))
                     + (2 - 3ρ²)/24·ν² ]
      σ    =  pre · (z / x(z)) · cor

    The IF(ABS(F-K)<1e-10, ...) handles the ATM limit, where z→0 and
    z/x(z)→1, with a simplified expression.
    """
    a, K, T = alpha_cell, K_cell, T_cell
    b, rho, nu, f = beta_cell, rho_cell, nu_cell, f_cell

    log_fK = f"LN({f}/{K})"
    fKb    = f"({f}*{K})^((1-{b})/2)"       # (F·K)^((1-β)/2)
    fKb2   = f"({f}*{K})^(1-{b})"           # (F·K)^(1-β)
    z      = f"({nu}/{a})*{fKb}*{log_fK}"
    x_z    = f"LN((SQRT(1-2*{rho}*({z})+({z})*({z}))+({z})-{rho})/(1-{rho}))"

    pre = (f"{a}/({fKb}*(1+(1-{b})^2/24*({log_fK})^2"
           f"+(1-{b})^4/1920*({log_fK})^4))")
    cor = (f"(1+{T}*((1-{b})^2/24*{a}^2/{fKb2}"
           f"+{rho}*{b}*{nu}*{a}/(4*{fKb})"
           f"+(2-3*{rho}^2)/24*{nu}^2))")
    z_over_xz = f"(({z})/({x_z}))"

    main_branch = f"({pre})*{z_over_xz}*{cor}"

    # ATM branch: when |F - K| < 1e-10 the formula above divides by zero
    # (since ln(F/K)=0 and z=0). Substitute the explicit ATM expansion.
    atm = (f"({a}/{f}^(1-{b}))*(1+{T}*((1-{b})^2/24*{a}^2/{f}^(2-2*{b})"
           f"+{rho}*{b}*{nu}*{a}/(4*{f}^(1-{b}))"
           f"+(2-3*{rho}^2)/24*{nu}^2))")

    return f"=IF(ABS({f}-{K})<1E-10,{atm},{main_branch})"


def sabr_alpha_update_formula(alpha_cell, f_cell, T_cell,
                               beta_cell, rho_cell, nu_cell, sigma_atm_cell):
    """
    One Newton iteration for SABR α calibration:
        α_new = α_old × σ_target / σ_atm(α_old)

    We start with α₀ = σ_atm × F^(1-β) and iterate. Two iterations
    are enough for our purposes (the residual after 2 iters is < 1e-6).
    """
    # σ at K=F (the ATM branch of SABR formula)
    a, T, b, rho, nu, f = alpha_cell, T_cell, beta_cell, rho_cell, nu_cell, f_cell
    sabr_atm = (f"(({a}/{f}^(1-{b}))*(1+{T}*((1-{b})^2/24*{a}^2/{f}^(2-2*{b})"
                f"+{rho}*{b}*{nu}*{a}/(4*{f}^(1-{b}))"
                f"+(2-3*{rho}^2)/24*{nu}^2)))")
    return f"={a}*{sigma_atm_cell}/{sabr_atm}"


def atm_interp_formula(t_range, vol_range, target_cell):
    """
    Linear interpolation on a (T, value) table.
    Used to look up ATM vol or ρ/ν at a given expiry T.

    Returns Excel formula that does:
      - if T <= first pillar: return first value
      - if T >= last pillar:  return last value
      - else: find bracketing pillars and linearly interpolate

    Uses MATCH (find index) + INDEX (lookup by index).
    """
    return (f"=IF({target_cell}<=INDEX({t_range},1),INDEX({vol_range},1),"
            f"IF({target_cell}>=INDEX({t_range},ROWS({t_range})),"
            f"INDEX({vol_range},ROWS({t_range})),"
            f"INDEX({vol_range},MATCH({target_cell},{t_range},1))"
            f"+({target_cell}-INDEX({t_range},MATCH({target_cell},{t_range},1)))"
            f"/(INDEX({t_range},MATCH({target_cell},{t_range},1)+1)"
            f"-INDEX({t_range},MATCH({target_cell},{t_range},1)))"
            f"*(INDEX({vol_range},MATCH({target_cell},{t_range},1)+1)"
            f"-INDEX({vol_range},MATCH({target_cell},{t_range},1)))))")


# ═══════════════════════════════════════════════════════════════════════════
# PART 2.2 — BUILD MarketData SHEET
# ═══════════════════════════════════════════════════════════════════════════
# This sheet has ALL hard-coded inputs (blue). Other sheets reference it.
# After writing the value we save the cell reference (e.g. "MarketData!$B$5")
# into `data['_md_xxx_cell']` so later sheets can use it.

def build_market_data(wb, data):
    ws = wb.create_sheet("MarketData")
    setcell(ws, 1, 1, "Market Data Inputs", TITLE_FONT)
    setcell(ws, 2, 1, "Blue = hardcoded inputs. Edit to re-price.",
            FORMULA_FONT, align=LEFT_ALIGN)
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 16

    r = 4  # Current row pointer — bumped after each section.

    # --------------- Trade terms ---------------
    setcell(ws, r, 1, "VALUATION & TRADE TERMS", SUBTITLE_FONT, fill=SECTION_FILL); r += 1

    # Write each value, capture its row, then build the cell reference.
    # This is robust: if we ever insert a row above, the captures still
    # point to the correct cell.
    setcell(ws, r, 1, "Valuation date", HEADER_FONT)
    valdate_row = r
    setcell(ws, r, 2, data['val_date'], INPUT_FONT, fmt='dd-mmm-yyyy', fill=INPUT_FILL); r += 1

    setcell(ws, r, 1, "Notional (USD)", HEADER_FONT)
    notional_row = r
    setcell(ws, r, 2, data['notional'], INPUT_FONT, fmt='#,##0', fill=INPUT_FILL); r += 1

    setcell(ws, r, 1, "Coupon rate", HEADER_FONT)
    coupon_row = r
    setcell(ws, r, 2, data['coupon'], INPUT_FONT, fmt='0.000%', fill=INPUT_FILL); r += 1

    setcell(ws, r, 1, "Barrier K (upper)", HEADER_FONT)
    K_row = r
    setcell(ws, r, 2, data['K'], INPUT_FONT, fmt='0.000%', fill=INPUT_FILL); r += 1

    setcell(ws, r, 1, "Put-spread half-width δ", HEADER_FONT)
    delta_K_row = r
    setcell(ws, r, 2, data.get('delta_K', 0.0001), INPUT_FONT,
            fmt='0.0000%', fill=INPUT_FILL); r += 1
    # .get('delta_K', 0.0001) means: return data['delta_K'] if it exists,
    # else 0.0001. Safer than direct indexing.

    setcell(ws, r, 1, "Leg 3 start (1st accrual period start)", HEADER_FONT)
    leg3_start_row = r
    setcell(ws, r, 2, data['leg3_start'], INPUT_FONT, fmt='dd-mmm-yyyy', fill=INPUT_FILL); r += 1

    setcell(ws, r, 1, "Leg 3 end (maturity)", HEADER_FONT)
    leg3_end_row = r
    setcell(ws, r, 2, data['leg3_end'], INPUT_FONT, fmt='dd-mmm-yyyy', fill=INPUT_FILL); r += 1

    # Static info (not used by formulas, just for human reference)
    setcell(ws, r, 1, "Reference index", HEADER_FONT)
    setcell(ws, r, 2, "10Y SOFR CMS", INPUT_FONT, fill=INPUT_FILL); r += 1
    setcell(ws, r, 1, "Daycount basis (coupon leg)", HEADER_FONT)
    setcell(ws, r, 2, "30/360", INPUT_FONT, fill=INPUT_FILL); r += 1
    setcell(ws, r, 1, "Daycount basis (curve/annuity)", HEADER_FONT)
    setcell(ws, r, 2, "ACT/360", INPUT_FONT, fill=INPUT_FILL); r += 2

    # Stash cell references for use in other sheets. f-strings glue the
    # captured row numbers onto "MarketData!$B$".
    data['_md_valdate_cell']    = f'MarketData!$B${valdate_row}'
    data['_md_notional_cell']   = f'MarketData!$B${notional_row}'
    data['_md_coupon_cell']     = f'MarketData!$B${coupon_row}'
    data['_md_K_cell']          = f'MarketData!$B${K_row}'
    data['_md_delta_K_cell']    = f'MarketData!$B${delta_K_row}'
    data['_md_leg3_start_cell'] = f'MarketData!$B${leg3_start_row}'
    data['_md_leg3_end_cell']   = f'MarketData!$B${leg3_end_row}'

    # --------------- 1D SOFR ---------------
    setcell(ws, r, 1, "1D MONEY MARKET", SUBTITLE_FONT, fill=SECTION_FILL); r += 1
    setcell(ws, r, 1, "1D SOFR mid-rate", HEADER_FONT)
    sofr_1d_row = r
    setcell(ws, r, 2, data['sofr_1d'], INPUT_FONT, fmt='0.0000%', fill=INPUT_FILL); r += 2
    data['_md_sofr_1d_cell'] = f'MarketData!$B${sofr_1d_row}'

    # --------------- FOMC step quotes ---------------
    setcell(ws, r, 1, "FOMC STEP QUOTES", SUBTITLE_FONT, fill=SECTION_FILL); r += 1
    setcell(ws, r, 1, "Period start", HEADER_FONT)
    setcell(ws, r, 2, "Period end",   HEADER_FONT)
    setcell(ws, r, 3, "Forward rate", HEADER_FONT); r += 1
    step_start_row = r
    for ps, pe, rate in data['step_quotes']:                    # iterate the list of tuples
        setcell(ws, r, 1, ps, INPUT_FONT, fmt='dd-mmm-yyyy', fill=INPUT_FILL)
        setcell(ws, r, 2, pe, INPUT_FONT, fmt='dd-mmm-yyyy', fill=INPUT_FILL)
        setcell(ws, r, 3, rate, INPUT_FONT, fmt='0.0000%', fill=INPUT_FILL)
        r += 1
    step_end_row = r - 1
    r += 1
    data['_md_step_range'] = (step_start_row, step_end_row)

    # --------------- SOFR futures ---------------
    setcell(ws, r, 1, "SOFR 3M FUTURES", SUBTITLE_FONT, fill=SECTION_FILL); r += 1
    for h, hv in enumerate(["IMM start", "IMM end", "Price",
                             "Conv adj (%)", "Fwd (computed)"], start=1):
        setcell(ws, r, h, hv, HEADER_FONT, align=CENTER_ALIGN)
    r += 1
    fut_start_row = r
    for y, m, price, conv in data['futures']:
        s = third_wednesday(y, m)
        # Next 3-month IMM date:
        ny, nm = (y + 1, 3) if m == 12 else (y, m + 3)
        e = third_wednesday(ny, nm)
        setcell(ws, r, 1, s, INPUT_FONT, fmt='dd-mmm-yyyy', fill=INPUT_FILL)
        setcell(ws, r, 2, e, INPUT_FONT, fmt='dd-mmm-yyyy', fill=INPUT_FILL)
        setcell(ws, r, 3, price, INPUT_FONT, fmt='0.0000', fill=INPUT_FILL)
        setcell(ws, r, 4, conv,  INPUT_FONT, fmt='0.0000', fill=INPUT_FILL)
        # Fwd = (100 - price)/100 - conv/100  — written as Excel formula
        setcell(ws, r, 5, f"=(100-C{r})/100-D{r}/100",
                FORMULA_FONT, fmt='0.0000%')
        r += 1
    fut_end_row = r - 1
    r += 1
    data['_md_fut_range'] = (fut_start_row, fut_end_row)

    # --------------- Swaps ---------------
    setcell(ws, r, 1, "PAR SWAP RATES (USD SOFR)", SUBTITLE_FONT, fill=SECTION_FILL); r += 1
    setcell(ws, r, 1, "Tenor (Y)", HEADER_FONT)
    setcell(ws, r, 2, "Par rate", HEADER_FONT); r += 1
    swap_start_row = r
    # sorted() turns dict→list of keys in order. .items() returns (key,val) pairs.
    for tenor in sorted(data['swaps'].keys()):
        rate = data['swaps'][tenor]
        setcell(ws, r, 1, tenor, INPUT_FONT, fmt='0', fill=INPUT_FILL)
        setcell(ws, r, 2, rate,  INPUT_FONT, fmt='0.0000%', fill=INPUT_FILL)
        r += 1
    swap_end_row = r - 1
    r += 1
    data['_md_swap_range'] = (swap_start_row, swap_end_row)

    # --------------- ATM vol surface ---------------
    setcell(ws, r, 1, "ATM VOL (10Y TAIL)", SUBTITLE_FONT, fill=SECTION_FILL); r += 1
    setcell(ws, r, 1, "Expiry (Y)", HEADER_FONT)
    setcell(ws, r, 2, "ATM σ", HEADER_FONT); r += 1
    atm_start_row = r
    for T, sigma in data['atm_10Y']:
        setcell(ws, r, 1, T, INPUT_FONT, fmt='0.0000', fill=INPUT_FILL)
        setcell(ws, r, 2, sigma, INPUT_FONT, fmt='0.0000%', fill=INPUT_FILL)
        r += 1
    atm_end_row = r - 1
    r += 1
    data['_md_atm_range'] = (atm_start_row, atm_end_row)

    # --------------- SABR rho/nu + beta ---------------
    setcell(ws, r, 1, "SABR PARAMETERS (10Y TAIL)", SUBTITLE_FONT, fill=SECTION_FILL); r += 1
    setcell(ws, r, 1, "Beta", HEADER_FONT)
    beta_row = r
    setcell(ws, r, 2, data['beta'], INPUT_FONT, fmt='0.00', fill=INPUT_FILL); r += 2
    data['_md_beta_cell'] = f'MarketData!$B${beta_row}'

    for h, hv in enumerate(["Expiry (Y)", "Rho", "Nu"], start=1):
        setcell(ws, r, h, hv, HEADER_FONT, align=CENTER_ALIGN)
    r += 1
    sabr_start_row = r
    for T, rho, nu in data['sabr_10Y']:
        setcell(ws, r, 1, T,   INPUT_FONT, fmt='0.0000', fill=INPUT_FILL)
        setcell(ws, r, 2, rho, INPUT_FONT, fmt='0.0000', fill=INPUT_FILL)
        setcell(ws, r, 3, nu,  INPUT_FONT, fmt='0.0000', fill=INPUT_FILL)
        r += 1
    sabr_end_row = r - 1
    data['_md_sabr_range'] = (sabr_start_row, sabr_end_row)

# --------------------------------------------------------------------------
# The other sheet builders (build_yield_curve, build_daily_dfs,
# build_daily_fixings, build_period_summary, build_valuation, build_readme)
# follow the same pattern: open the sheet, write titles, then write inputs
# or formulas referencing MarketData via the captured `_md_xxx_cell`
# strings.
#
# For brevity, see build.py for the full implementations. The next file
# (Part 3) wraps everything up with the main() entry point.
# --------------------------------------------------------------------------
