"""
Pattern Present-Validation Checker

Compares an exact 0/1 pattern's older historical behavior
with its latest N occurrences (for example 100 or 150).

1 = exit price > entry price
0 = exit price < entry price
doji = exit price == entry price (breaks the sequence)

Install:
    pip install pandas openpyxl

Run:
    python pattern_present_verify.py
"""

from pathlib import Path
import math
import pandas as pd


def load_file(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        xls = pd.ExcelFile(path)
        sheet = "Trades" if "Trades" in xls.sheet_names else xls.sheet_names[0]
        print(f"Reading sheet: {sheet}")
        return pd.read_excel(path, sheet_name=sheet)
    if suffix == ".csv":
        return pd.read_csv(path)
    raise ValueError("Use .xlsx, .xlsm, or .csv")


def find_col(df, exact_names, startswith=None):
    lookup = {str(c).strip().lower(): c for c in df.columns}
    for name in exact_names:
        if name.lower() in lookup:
            return lookup[name.lower()]
    for col in df.columns:
        col_name = str(col).strip().lower()
        for prefix in (startswith or []):
            if col_name.startswith(prefix.lower()):
                return col
    return None


def extract_trades(df):
    trade_col = find_col(df, ["Trade number", "Trade #", "Trade"])
    type_col = find_col(df, ["Type", "Order type"])
    price_col = find_col(df, ["Price"], startswith=["Price "])
    time_col = find_col(df, ["Date and time", "Date/Time", "Time", "Date"])

    if trade_col is None or type_col is None or price_col is None:
        raise ValueError(
            "Could not find Trade number, Type, or Price columns.\n"
            "Columns found: " + ", ".join(map(str, df.columns))
        )

    cols = [trade_col, type_col, price_col]
    if time_col is not None:
        cols.append(time_col)

    work = df[cols].copy()
    work[price_col] = pd.to_numeric(work[price_col], errors="coerce")
    work["_type"] = work[type_col].astype(str).str.strip().str.lower()

    rows = []
    for trade_no, g in work.groupby(trade_col, sort=False):
        entry = g[g["_type"].str.contains("entry", na=False)]
        exit_ = g[g["_type"].str.contains("exit", na=False)]
        if entry.empty or exit_.empty:
            continue

        ep = entry.iloc[-1][price_col]
        xp = exit_.iloc[-1][price_col]
        if pd.isna(ep) or pd.isna(xp):
            continue

        if xp > ep:
            result = 1
        elif xp < ep:
            result = 0
        else:
            result = None

        event_time = exit_.iloc[-1][time_col] if time_col is not None else trade_no
        rows.append({"time": event_time, "trade": trade_no, "result": result})

    if not rows:
        raise ValueError("No complete entry/exit trades found.")

    temp = pd.DataFrame(rows)
    if time_col is not None:
        parsed = pd.to_datetime(temp["time"], errors="coerce")
        if parsed.notna().any():
            temp["_parsed"] = parsed
            temp = temp.sort_values(["_parsed", "trade"], kind="stable")
        else:
            temp = temp.sort_values("trade", kind="stable")
    else:
        temp = temp.sort_values("trade", kind="stable")

    return temp.reset_index(drop=True)


def find_occurrences(trades, pattern):
    wanted = [int(c) for c in pattern]
    n = len(wanted)
    seq = trades["result"].tolist()
    matches = []

    for i in range(n, len(seq)):
        prev = seq[i-n:i]
        nxt = seq[i]

        if pd.isna(nxt) or any(pd.isna(v) for v in prev):
            continue

        prev = [int(v) for v in prev]
        nxt = int(nxt)

        if prev == wanted:
            matches.append({
                "Pattern": pattern,
                "Pattern_End_Time": trades.iloc[i-1]["time"],
                "Next_Time": trades.iloc[i]["time"],
                "Next_Result": nxt,
            })

    return pd.DataFrame(matches)


def wilson_interval(successes, n, z=1.96):
    if n == 0:
        return float("nan"), float("nan")

    p = successes / n
    denom = 1 + z*z/n
    center = (p + z*z/(2*n)) / denom
    margin = z * math.sqrt((p*(1-p)/n) + z*z/(4*n*n)) / denom
    return center - margin, center + margin


def main():
    print("=" * 68)
    print("PATTERN PRESENT-VALIDATION CHECKER")
    print("=" * 68)

    raw_path = input("TradingView XLSX/CSV path: ").strip().strip('"')
    path = Path(raw_path)

    if not path.exists():
        print(f"File not found: {path}")
        return

    pattern = input("Pattern (example 0101): ").strip()
    if not pattern or any(c not in "01" for c in pattern):
        print("ERROR: pattern must contain only 0 and 1.")
        return

    raw_n = input("Recent occurrences to verify [100]: ").strip()
    recent_n = int(raw_n) if raw_n else 100

    try:
        df = load_file(path)
        trades = extract_trades(df)
        matches = find_occurrences(trades, pattern)
    except Exception as exc:
        print(f"ERROR: {exc}")
        return

    total = len(matches)
    if total <= recent_n:
        print(
            f"Not enough occurrences. Found {total}, but need more than "
            f"{recent_n} so older data remains as baseline."
        )
        return

    historical = matches.iloc[:-recent_n].copy()
    recent = matches.iloc[-recent_n:].copy()

    hist_n = len(historical)
    hist_ones = int((historical["Next_Result"] == 1).sum())
    hist_zeros = int((historical["Next_Result"] == 0).sum())

    p1 = hist_ones / hist_n
    p0 = hist_zeros / hist_n

    if p1 >= p0:
        side = 1
        hist_rate = p1
    else:
        side = 0
        hist_rate = p0

    recent_wins = int((recent["Next_Result"] == side).sum())
    recent_rate = recent_wins / recent_n
    expected_wins = recent_n * hist_rate

    se = math.sqrt(hist_rate * (1 - hist_rate) / recent_n)
    z_score = (recent_rate - hist_rate) / se if se > 0 else float("nan")

    ci_low, ci_high = wilson_interval(recent_wins, recent_n)
    diff_pp = (recent_rate - hist_rate) * 100

    if math.isnan(z_score):
        verdict = "UNDEFINED"
    elif abs(z_score) < 1:
        verdict = "STABLE"
    elif abs(z_score) < 1.96:
        verdict = "WATCH"
    else:
        verdict = "CHANGED"

    summary = pd.DataFrame([{
        "Pattern": pattern,
        "Historical_Occurrences": hist_n,
        "Historical_Majority_Side": side,
        "Historical_Win_Rate_Pct": round(hist_rate * 100, 3),
        "Recent_Occurrences": recent_n,
        "Recent_Wins": recent_wins,
        "Recent_Win_Rate_Pct": round(recent_rate * 100, 3),
        "Expected_Recent_Wins": round(expected_wins, 2),
        "Difference_Percentage_Points": round(diff_pp, 3),
        "Z_Score": round(z_score, 3) if not math.isnan(z_score) else None,
        "Recent_95CI_Low_Pct": round(ci_low * 100, 3),
        "Recent_95CI_High_Pct": round(ci_high * 100, 3),
        "Verdict": verdict,
    }])

    summary_path = path.with_name(f"{pattern}_present_verify_{recent_n}.csv")
    recent_path = path.with_name(f"{pattern}_recent_{recent_n}_results.csv")

    summary.to_csv(summary_path, index=False)
    recent[["Next_Result"]].to_csv(recent_path, index=False)

    print("\nHISTORICAL")
    print(f"Occurrences     : {hist_n}")
    print(f"Majority side   : {side}")
    print(f"Historical rate : {hist_rate * 100:.2f}%")

    print("\nLATEST SAMPLE")
    print(f"Occurrences     : {recent_n}")
    print(f"Wins            : {recent_wins}")
    print(f"Recent rate     : {recent_rate * 100:.2f}%")
    print(f"Expected wins   : {expected_wins:.2f}")
    print(f"Difference      : {diff_pp:+.2f} percentage points")
    print(f"Z-score         : {z_score:.3f}")
    print(f"95% CI          : {ci_low * 100:.2f}% to {ci_high * 100:.2f}%")
    print(f"Verdict         : {verdict}")

    print("\nSTABLE  = recent behavior is close to historical")
    print("WATCH   = noticeable drift")
    print("CHANGED = recent behavior is statistically different")

    print(f"\nSummary CSV: {summary_path}")
    print(f"Recent CSV : {recent_path}")


if __name__ == "__main__":
    main()
