"""
Pattern Next-Result Checker

Finds an exact 0/1 pattern in a TradingView strategy export
and exports ONLY the next result for the latest N valid occurrences.
"""

from pathlib import Path
import pandas as pd


def load_file(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        xls = pd.ExcelFile(path)
        sheet = "Trades" if "Trades" in xls.sheet_names else xls.sheet_names[0]
        return pd.read_excel(path, sheet_name=sheet)

    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)

    raise ValueError("Use .xlsx, .xlsm, or .csv")


def find_col(df, names, startswith=None):
    lookup = {str(c).strip().lower(): c for c in df.columns}

    for name in names:
        if name.lower() in lookup:
            return lookup[name.lower()]

    startswith = startswith or []
    for col in df.columns:
        col_name = str(col).strip().lower()
        for prefix in startswith:
            if col_name.startswith(prefix.lower()):
                return col

    return None


def get_sequence(df):
    trade_col = find_col(df, ["Trade number", "Trade #", "Trade"])
    type_col = find_col(df, ["Type", "Order type"])
    # TradingView names this by quote currency, e.g. "Price CHF".
    price_col = find_col(df, ["Price"], startswith=["Price "])
    time_col = find_col(df, ["Date and time", "Date/Time", "Time", "Date"])

    if trade_col is None or type_col is None or price_col is None:
        raise ValueError("Could not find Trade number, Type, or Price columns.")

    cols = [trade_col, type_col, price_col]
    if time_col is not None:
        cols.append(time_col)

    work = df[cols].copy()
    work[price_col] = pd.to_numeric(work[price_col], errors="coerce")
    work["_type"] = work[type_col].astype(str).str.lower()

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

        t = exit_.iloc[-1][time_col] if time_col is not None else trade_no
        rows.append((t, trade_no, result))

    temp = pd.DataFrame(rows, columns=["time", "trade", "result"])

    if time_col is not None:
        parsed = pd.to_datetime(temp["time"], errors="coerce")
        if parsed.notna().any():
            temp["_time"] = parsed
            temp = temp.sort_values(["_time", "trade"], kind="stable")
        else:
            temp = temp.sort_values("trade", kind="stable")
    else:
        temp = temp.sort_values("trade", kind="stable")

    return temp["result"].tolist()


def find_next_results(sequence, pattern: str):
    wanted = [int(x) for x in pattern]
    n = len(wanted)
    results = []

    for i in range(n, len(sequence)):
        prev = sequence[i-n:i]
        nxt = sequence[i]

        if pd.isna(nxt) or any(pd.isna(v) for v in prev):
            continue

        prev = [int(v) for v in prev]
        nxt = int(nxt)

        if prev == wanted:
            results.append(nxt)

    return results


def main():
    print("PATTERN NEXT-RESULT CHECKER")

    raw = input("TradingView XLSX/CSV path: ").strip().strip('"')
    path = Path(raw)

    pattern = input("Pattern (example 0101): ").strip()
    if not pattern or any(c not in "01" for c in pattern):
        print("ERROR: pattern must contain only 0 and 1.")
        return

    raw_count = input("How many latest occurrences [100]: ").strip()
    count = int(raw_count) if raw_count else 100

    df = load_file(path)
    sequence = get_sequence(df)
    all_results = find_next_results(sequence, pattern)
    latest = all_results[-count:]

    output = path.with_name(f"{pattern}_latest_{len(latest)}_next_results.csv")
    pd.DataFrame({"Next_Result": latest}).to_csv(output, index=False)

    print(f"Pattern: {pattern}")
    print(f"Found: {len(all_results)}")
    print(f"Exported latest: {len(latest)}")
    print(f"CSV: {output}")


if __name__ == "__main__":
    main()
