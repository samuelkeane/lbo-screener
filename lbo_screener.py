"""
LBO target screener.

Rank UK and European listed companies as take-private candidates. Reads a Capital IQ screen export, removes companies outside of the mandate,
runs a simplified 5-yr LBO on each company, and writes a ranked excel report.
"""

import sys
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font

pd.set_option("display.width", 200)
pd.set_option("display.max_colwidth", 60)

# --- settings ---

FILE = "data/capiq_exports.xls" # default input if no file is given in the terminal.
OUT_DIR = "output" # folder for results; each input gets its own results file.

COLS = { # stores in pairs. left name is the key we want for each column, the right is what CapIQ gave.
    "name": "Company Name",
    "ticker": "Exchange:Ticker",
    "sector": "Primary Industry",
    "country": "Headquarters - Country/Region",
    "mktcap": "Market Capitalization",
    "tev": "Total Enterprise Value",
    "net_debt": "Net Debt",
    "leases": "Total Leases",
    "revenue": "Total Revenue [LTM]",
    "ebitda": "EBITDA [LTM]",
    "capex": "Capital Expenditure",
    "cfo": "Cash from Ops",
    "interest": "Interest Expense",
    "price": "Day Close Price",
    "high52": "52 Week High Price",
    "rev_cagr": "Total Revenues, 3 Yr CAGR",
    "insider_pct": "Individuals/Insiders",
    "ebitda_fy1": "EBITDA - Capital IQ [FY+1]",
    "ebitda_fy2": "EBITDA - Capital IQ [FY+2]",
    "m2025": "EBITDA Margin % [CY 2025]",
    "m2024": "EBITDA Margin % [CY 2024]",
    "m2023": "EBITDA Margin % [CY 2023]",
    "m2022": "EBITDA Margin % [CY 2022]",
    "m2021": "EBITDA Margin % [CY 2021]",
}

TEXT_COLS = ["name", "ticker", "sector", "country"] # columns that remain as text rather than numerical.

NON_EUROPE = {"NYSE", "NasdaqGS", "NasdaqGM", "OTCPK", "ASX", "TSX", "TSXV",
               "TSE", "BOVESPA", "NSEI", "JSE", "TASE"} # 12 Non-European exchanges found. Exclude these.
REQUIRED = ["mktcap", "tev", "net_debt", "revenue", "ebitda", "capex"] # 6 inputs the mini-lbo requires to run with.

MARGINS = ["m2025", "m2024", "m2023", "m2022", "m2021"] # 5 years of EBITDA margin used for the stability measure.

A = { # lbo assumptions.
    "premium": 0.35, # 35% premium to current share price.
    "fees": 0.025, # advisory and financing fees, as a % of entry EV.
    "leverage": 5.0, # debt raised at entry, x LTM EBITDA.
    "max_debt_pct": 0.60, # debt can never exceed 60% of total cost.
    "rate": 0.08, # 8% interest on debt. placeholder.
    "tax": 0.25, # 25% tax rate - uk corporate tax rate.
    "years": 5, # 5 year holding period, typical for a buyout.
}

WEIGHTS = {"irr": 0.45, "equity_pct": 0.15, "margin_std": 0.15, "cash_conv": 0.15, "fcf_to_debt": 0.10} # weights for each metric in the scoring system. Must sum to 1.0.
HIGHER_IS_BETTER = {"irr": True, "equity_pct": False, "margin_std": False, "cash_conv": True, "fcf_to_debt": True} # defines which metrics are better when higher or lower.

SHOW = {
    "rank": "Rank",
    "name": "Company",
    "country": "Country",
    "sector": "Sector",
    "score": "Score",
    "irr": "IRR",
    "moic": "MOIC",
    "tev_xl": "EV ex-leases (£m)",
    "mktcap": "Market cap (£m)",
    "net_debt_xl": "Net debt ex-leases (£m)",
    "ebitda": "LTM EBITDA (£m)",
    "margin": "EBITDA margin",
    "ev_ebitda": "EV / EBITDA",
    "vs_sector": "vs sector median",
    "entry_mult": "Entry multiple",
    "equity_pct": "Equity % of uses",
    "cash_conv": "Cash conversion",
    "margin_std": "Margin volatility (pp)",
    "growth": "EBITDA growth",
    "net_lev": "Current net leverage",
    "flags": "Flags",
}  # list columns for excel export.

FORMATS = {
    "Score": "0.0",
    "IRR": "0.0%",
    "MOIC": '0.00"x"',
    "EV ex-leases (£m)": "#,##0",
    "Market cap (£m)": "#,##0",
    "Net debt ex-leases (£m)": "#,##0",
    "LTM EBITDA (£m)": "#,##0",
    "EV (£m)": "#,##0",
    "EBITDA margin": "0.0%",
    "EV / EBITDA": '0.0"x"',
    "vs sector median": "+0%;-0%",
    "Entry multiple": '0.0"x"',
    "Equity % of uses": "0%",
    "Cash conversion": "0%",
    "Margin volatility (pp)": "0.0",
    "EBITDA growth": "0.0%",
    "Current net leverage": '0.0"x"',
} # formats cell for excel.
WIDE = {"Company": 45, "Sector": 30, "Flags": 60, "Excluded because": 25} # set column widths for excel.

# --- functions ---

def load_export(path):
    df = pd.read_excel(path, sheet_name="Screening", header=7, engine="calamine")
    df = df[df["Exchange:Ticker"].notna()]
    return df # read the cap IQ export and drop the footer rows.

def find_col(df, start):
    for col in df.columns:
        if col.startswith(start):
            return col
    raise KeyError("No column starts with: " + start) # if no header starts with this text, stop the script with an error naming the missing column.

def clean(df):
    d = pd.DataFrame()
    for short, start in COLS.items():
        d[short] = df[find_col(df, start)] # keep columns in COLS


    for col in d.columns:
        if col not in TEXT_COLS:
            d[col] = pd.to_numeric(d[col], errors="coerce")

    d["capex"] = d["capex"].abs()
    d["interest"] = d["interest"].abs()
    d["insider_pct"] = d["insider_pct"].clip(lower=0)
    d["lease_missing"] = d["leases"].isna() # records missing leases before filling the blanks.
    d["leases"] = d["leases"].fillna(0) # no leases, fill with zero.
    d["exchange"] = d["ticker"].str.split(":").str[0]
    return d # give the columns new, shortened names. Also convert to numbers. insider_pct fix for data error.

# --- exclusions ---

def exclude(d, mask, why):
    d.loc[mask & (d["reason"] == ""), "reason"] = why # write why as the reason for rows in 'mask' that have no reason assigned yet.
 
def apply_exclusions(d): # label companies outside of the given mandate. Return those kept and those excluded as two separate tables.
    d["reason"] = ""
    exclude(d, d["exchange"].isin(NON_EUROPE), "Listed outside Europe") # part of mandate - UK & European listings.
    exclude(d, d[REQUIRED].isna().any(axis=1), "Missing required data") # essential data for the lbo.
    exclude(d, (d["ebitda"] <= 0) | (d["revenue"] <= 0), "EBITDA or revenue not positive") # loss-making company cannot support debt, so filter out those with negative EBITDA.
    exclude(d, (d["tev"] < 1000) | (d["tev"] > 10000), "EV outside £1-10bn") # given in mandate - £1-10bn EV

    excluded = d[d["reason"] != ""].copy()
    kept = d[d["reason"] == ""].copy()
    return kept, excluded 

def growth_rule(row): # return the growth rate and the source for it. Applied to each company in turn. Choose the best data - if both analysts forecasts are positive, use that. If not, use CAGR. Worst case, use a 3% default.
    if row["ebitda_fy1"] > 0 and row["ebitda_fy2"] > 0:
        return row["ebitda_fy2"] / row["ebitda_fy1"] - 1, "consensus"
    if pd.notna(row["rev_cagr"]):
        return row["rev_cagr"] / 100, "revenue CAGR"
    return 0.03, "default"


def add_metrics(d): # defines and adds the following metrics for each company.
    d["net_debt_xl"] = d["net_debt"] - d["leases"] # remove leases - cap IQ EBITDA is after rent, so leases in debt would double-count.
    d["tev_xl"] = d["tev"] - d["leases"]
    d["ev_ebitda"] = d["tev_xl"] / d["ebitda"]
    d["margin"] = d["ebitda"] / d["revenue"]
    d["cash_conv"] = (d["ebitda"] - d["capex"]) / d["ebitda"] # share of EBITDA left after capital expenditure - this is the cash that will repay LBO debt.
    d["net_lev"] = d["net_debt_xl"] / d["ebitda"]

    enough = d[MARGINS].count(axis=1) >= 3 # enough is defined as true where at least 3 of the 5 yearly margins have data.
    d["margin_std"] = d[MARGINS].std(axis=1).where(enough) # take the standard deviation of the margins, but only where there are 3 or more years of data.

    d["sector_median"] = d.groupby("sector")["ev_ebitda"].transform("median")
    d["vs_sector"] = d["ev_ebitda"] / d["sector_median"] - 1 # compare valuation relative to sector.

    g = d.apply(growth_rule, axis=1, result_type="expand")
    d["growth"] = g[0].clip(-0.02, 0.10) # caps growth rates between -2 and 10%. Consensus is often optimistic.
    d["growth_source"] = g[1]
    return d

# --- mini lbo ---

def mini_lbo(row): # defines the mini lbo function - a simplified 5-year cash sweep LBO and return the results.
    entry_ev = row["tev_xl"] + row["mktcap"] * A["premium"] # what pays for whole business.
    fees = entry_ev * A["fees"] 
    uses = entry_ev + fees # total money needed - price plus fees.
    entry_mult = entry_ev / row["ebitda"] # price paid as multiple of EBITDA.
    debt = min(A["leverage"] * row["ebitda"], A["max_debt_pct"] * uses) #returns smaller of two numbers. Debt is whichever hits first - 5x EBITDA or 60% of uses.
    equity = uses - debt # what the debt does not cover must be covered by equity.
    start_debt = debt

    margin = row["ebitda"] / row["revenue"]
    capex_ratio = row["capex"] / row["revenue"]
    revenue = row["revenue"]

    for year in range(1, A["years"] + 1):
        revenue = revenue * (1 + row["growth"]) # revenue growth modelled.
        ebitda = revenue * margin # ebitda grow at a rate equal to revenue growth.
        capex = revenue * capex_ratio # capex grows at same rate as revenue.
        interest = A["rate"] * max(debt, 0) # interest paid on debt, but no debt to pay interest on when debt is below 0. Interest stops once the debt is fully repaid.
        tax = max(0, A["tax"] * (ebitda - capex - interest)) # tax calculations.
        fcf = ebitda - capex - interest - tax # free cash flow calculations.
        if year == 1:
            fcf_to_debt = fcf / start_debt
        debt = debt - fcf # cash sweep: all free cash flow repays debt.

    exit_equity = entry_mult * ebitda - debt # equity value at exit
    moic = exit_equity / equity # multiple on invested capital.
    if moic > 0:
        irr = moic ** (1 / A["years"]) - 1 # annualised return (IRR).
    else:
        irr = -1.0

    return { # results for a company.
        "entry_mult": entry_mult,
        "equity_pct": equity / uses,
        "fcf_to_debt": fcf_to_debt,
        "exit_lev": debt / ebitda,
        "moic": moic,
        "irr": irr, 
    }

# --- returns and scoring ---

def run_lbo(d):
    lbo = d.apply(mini_lbo, axis=1, result_type="expand") # run the lbo and add the results to the table as columns.
    return d.join(lbo) 


def score(d): # combine percentile ranks into the 0-100 score.
    d["score"] = 0.0
    for metric, weight in WEIGHTS.items():
        pct = d[metric].rank(pct=True, ascending=HIGHER_IS_BETTER[metric]).fillna(0.5) # 0-1 percentile rank. 1 = best. Missing data gets a middle score (0.5).
        d["score"] = d["score"] + weight * pct
    d["score"] = d["score"] * 100

    d = d.sort_values("score", ascending=False).reset_index(drop=True) # sort with best score first.
    d["rank"] = d.index + 1 # include a rank, starting from 1.
    return d


def make_flags(row): # return the warning for each company, and join them in a string.
    flags = []
    if row["tev_xl"] < 1000:
        flags.append("EV below £1bn ex-leases")
    elif not 2000 <= row["tev_xl"] <= 8000:
        flags.append("EV outside £2-8bn")
    if row["price"] * (1 + A["premium"]) < row["high52"]: # shareholders may be less willing to sell - price they are being offered is below 52wk highs.
        flags.append("offer below 52w high")
    if row["net_lev"] > 3.5:
        flags.append("already levered")
    if row["insider_pct"] > 30: # is there a stake (likely family or founder) large enough to block a deal?
        flags.append("insiders own " + str(round(row["insider_pct"])) + "%")
    if row["fcf_to_debt"] < 0:
        flags.append("negative yr-1 FCF")
    if row["equity_pct"] > 0.70:
        flags.append("thin debt capacity")
    if row["growth_source"] != "consensus":
        flags.append("no consensus growth")
    if row["lease_missing"]:
        flags.append("lease data missing")
    return "; ".join(flags)

# --- export to excel ---

def write_excel(d, excluded, path): # write out the ranked results, exclusions and assumptions in an excel file.
    out = d[list(SHOW)].rename(columns=SHOW)
    sweet = out[d["tev_xl"].between(2000, 8000)] # target size range for the shortlist; narrower than the £1-10bn mandate.

    excl = excluded[["name", "ticker", "tev", "reason"]].rename(columns={ # exclusions tab.
        "name": "Company", "ticker": "Exchange:Ticker",
        "tev": "EV (£m)", "reason": "Excluded because"})

    assumptions = pd.DataFrame(list(A.items()) + list(WEIGHTS.items()), # assumptions tab.
                               columns=["Setting", "Value"])

    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        sweet.head(25).to_excel(xw, sheet_name="Top 25 £2-8bn", index=False) # one sheet per table.
        out.head(25).to_excel(xw, sheet_name="Top 25 overall", index=False)
        out.to_excel(xw, sheet_name="All ranked", index=False)
        excl.to_excel(xw, sheet_name="Excluded", index=False)
        assumptions.to_excel(xw, sheet_name="Assumptions", index=False)

        for ws in xw.book.worksheets: # bold headers, freeze panes, set column widths, and format numbers.
            ws.freeze_panes = "C2"
            for cell in ws[1]:
                header = cell.value
                cell.font = Font(bold=True)
                ws.column_dimensions[cell.column_letter].width = WIDE.get(header, 14)
                if header in FORMATS:
                    for row in ws.iter_rows(min_row=2, min_col=cell.column, max_col=cell.column):
                        row[0].number_format = FORMATS[header]


def main(): # runs the screener from start to finish.
    path = sys.argv[1] if len(sys.argv) > 1 else FILE 
    Path(OUT_DIR).mkdir(exist_ok=True)
    out_path = Path(OUT_DIR) / (Path(path).stem + "_results.xlsx")

    df = load_export(path)
    print("Companies loaded:", len(df), "from", path)

    d = clean(df)
    d, excluded = apply_exclusions(d)
    print(excluded["reason"].value_counts())
    print("Excluded:", len(excluded), "| Remaining:", len(d))

    d = add_metrics(d)
    d = run_lbo(d)
    print(d[["entry_mult", "equity_pct", "irr", "moic"]].describe().round(3))

    d = score(d)
    d["flags"] = d.apply(make_flags, axis=1)
    print(d[["rank", "name", "irr", "flags"]].head(15).to_string(index=False))

    in_range = d["tev_xl"].between(2000, 8000).sum()
    write_excel(d, excluded, out_path)
    print("Saved", out_path, "| ranked:", len(d), "| in £2-8bn:", in_range, "| excluded:", len(excluded))


if __name__ == "__main__":
    main()
