"""
TradingView Pattern Occurrence Checker

Purpose:
- You already found a pattern such as 0101.
- This script finds every historical occurrence of that exact pattern.
- It then lists what happened on the NEXT candle/trade.
- You can export the latest 100 occurrences (or any number) to CSV.

Rules:
    1 = exit price > entry price
    0 = exit price < entry price
    doji = exit price == entry price

Dojis break the sequence.
So a pattern never jumps across a doji.

Install:
    pip install pandas openpyxl

Run:
    python pattern_trade_list.py
"""

from pathlib import Path
import pandas as pd


def load_tradingview_file(file_path: Path) -> pd.DataFrame:
    suffix = file_path.suffix.lower()

    if suffix in {".xlsx", ".xlsm"}:
        xls = pd.ExcelFile(file_path)
        sheet = "Trades" if "Trades" in xls.sheet_names else xls.sheet_names[0]
        print(f"Reading sheet: {sheet}")
        return pd.read_excel(file_path, sheet_name=sheet)

    if suffix == ".csv":
        return pd.read_csv(file_path)

    raise ValueError("Use a .xlsx, .xlsm, or .csv file.")


def find_column(df: pd.DataFrame, exact=None, startswith=None):
    exact = exact or []
    startswith = startswith or []

    normalized = {str(c).strip().lower(): c for c in df.columns}

    for name in exact:
        key = name.strip().lower()
        if key in normalized:
            return normalized[key]

    for col in df.columns:
        col_name = str(col).strip().lower()
        for prefix in startswith:
            if col_name.startswith(prefix.lower()):
                return col

    return None


def extract_trades(df: pd.DataFrame) -> pd.DataFrame:
    trade_col = find_column(
        df,
        exact=["Trade number", "Trade #", "Trade"]
    )
    type_col = find_column(
        df,
        exact=["Type", "Order type"]
    )
    time_col = find_column(
        df,
        exact=["Date and time", "Date/Time", "Time", "Date"]
    )
    price_col = find_column(
        df,
        exact=["Price"],
        startswith=["Price "]
    )

    required = {
        "trade number": trade_col,
        "type": type_col,
        "price": price_col,
    }

    missing = [name for name, col in required.items() if col is None]
    if missing:
        raise ValueError(
            "Could not find required TradingView columns: "
            + ", ".join(missing)
            + "\nColumns found:\n"
            + ", ".join(map(str, df.columns))
        )

    cols = [trade_col, type_col, price_col]
    if time_col is not None:
        cols.append(time_col)

    work = df[cols].copy()
    work[price_col] = pd.to_numeric(work[price_col], errors="coerce")
    work["_type"] = work[type_col].astype(str).str.strip().str.lower()

    records = []

    for trade_no, group in work.groupby(trade_col, sort=False):
        entry_rows = group[group["_type"].str.contains("entry", na=False)]
        exit_rows = group[group["_type"].str.contains("exit", na=False)]

        if entry_rows.empty or exit_rows.empty:
            continue

        entry_row = entry_rows.iloc[-1]
        exit_row = exit_rows.iloc[-1]

        entry_price = entry_row[price_col]
        exit_price = exit_row[price_col]

        if pd.isna(entry_price) or pd.isna(exit_price):
            continue

        if exit_price > entry_price:
            result = 1
        elif exit_price < entry_price:
            result = 0
        else:
            result = float("nan")

        entry_time = entry_row[time_col] if time_col is not None else trade_no
        exit_time = exit_row[time_col] if time_col is not None else trade_no

        records.append({
            "trade_number": trade_no,
            "entry_time": entry_time,
            "exit_time": exit_time,
            "entry_price": entry_price,
            "exit_price": exit_price,
            "result": result,
        })

    if not records:
        raise ValueError("No complete entry/exit trades were found.")

    trades = pd.DataFrame(records)

    # Sort chronologically if timestamps parse.
    parsed = pd.to_datetime(trades["exit_time"], errors="coerce")
    if parsed.notna().any():
        trades["_sort_time"] = parsed
        trades = trades.sort_values(
            ["_sort_time", "trade_number"],
            kind="stable"
        )
    else:
        trades = trades.sort_values("trade_number", kind="stable")

    return trades.reset_index(drop=True)


def validate_pattern(raw: str) -> str:
    pattern = raw.strip()

    if not pattern:
        raise ValueError("Pattern cannot be empty.")

    if any(ch not in "01" for ch in pattern):
        raise ValueError("Pattern must contain only 0 and 1.")

    return pattern


def find_occurrences(trades: pd.DataFrame, pattern: str) -> pd.DataFrame:
    length = len(pattern)
    expected = [int(ch) for ch in pattern]

    rows = []

    # i is the index of the NEXT trade after the pattern.
    for i in range(length, len(trades)):
        pattern_block = trades.iloc[i - length:i]
        next_trade = trades.iloc[i]

        pattern_results = pattern_block["result"].tolist()
        next_result = next_trade["result"]

        # Pandas converts None values in numeric columns to NaN.
        # Treat both None and NaN as dojis / invalid directional candles.
        if pd.isna(next_result) or any(pd.isna(v) for v in pattern_results):
            continue

        # Convert clean numeric values to ints before comparing.
        pattern_results = [int(v) for v in pattern_results]
        next_result = int(next_result)

        if pattern_results != expected:
            continue

        rows.append({
            "Pattern": pattern,
            "Pattern_Start_Time": pattern_block.iloc[0]["entry_time"],
            "Pattern_End_Time": pattern_block.iloc[-1]["exit_time"],
            "Pattern_Trade_Numbers": " ".join(
                str(x) for x in pattern_block["trade_number"].tolist()
            ),
            "Next_Trade_Number": next_trade["trade_number"],
            "Next_Entry_Time": next_trade["entry_time"],
            "Next_Exit_Time": next_trade["exit_time"],
            "Next_Entry_Price": next_trade["entry_price"],
            "Next_Exit_Price": next_trade["exit_price"],
            "Next_Result": int(next_result),
        })

    result = pd.DataFrame(rows)

    if not result.empty:
        result.insert(0, "Occurrence_Number", range(1, len(result) + 1))

    return result


def ask_count(default=100):
    while True:
        raw = input(f"How many occurrences to export [{default}]: ").strip()

        if not raw:
            return default

        try:
            value = int(raw)
            if value < 1:
                raise ValueError
            return value
        except ValueError:
            print("Enter an integer >= 1.")


def main():
    print("=" * 65)
    print("PATTERN OCCURRENCE / TRADE LIST")
    print("=" * 65)

    raw_path = input("TradingView XLSX/CSV path: ").strip().strip('"')
    file_path = Path(raw_path).expanduser()

    if not file_path.exists():
        print(f"\nFile not found:\n{file_path}")
        return

    try:
        pattern = validate_pattern(
            input("Pattern to check (example 0101): ")
        )
        count = ask_count(100)

        df = load_tradingview_file(file_path)
        trades = extract_trades(df)
        matches = find_occurrences(trades, pattern)

    except Exception as exc:
        print(f"\nERROR: {exc}")
        return

    if matches.empty:
        print(f"\nPattern {pattern} was not found with a valid next candle.")
        return

    total_occurrences = len(matches)

    # Export latest N occurrences.
    export_df = matches.tail(count).copy()

    # Renumber the exported rows 1..N for convenience,
    # while preserving original historical occurrence number.
    export_df.insert(
        1,
        "Export_Row",
        range(1, len(export_df) + 1)
    )

    next_1_count = int((export_df["Next_Result"] == 1).sum())
    next_0_count = int((export_df["Next_Result"] == 0).sum())
    sample_count = len(export_df)

    next_1_pct = next_1_count / sample_count * 100
    next_0_pct = next_0_count / sample_count * 100

    output_path = file_path.with_name(
        f"pattern_{pattern}_latest_{sample_count}_occurrences.csv"
    )

    export_df.to_csv(output_path, index=False)

    print("\nPATTERN")
    print(f"Pattern                  : {pattern}")
    print(f"Total valid occurrences  : {total_occurrences}")
    print(f"Exporting latest         : {sample_count}")

    print("\nNEXT-CANDLE RESULTS IN EXPORTED SAMPLE")
    print(f"Next = 1                 : {next_1_count} ({next_1_pct:.2f}%)")
    print(f"Next = 0                 : {next_0_count} ({next_0_pct:.2f}%)")

    print("\nCSV")
    print(output_path)

    print("\nLAST 15 EXPORTED OCCURRENCES")
    print(
        export_df[
            [
                "Occurrence_Number",
                "Pattern_End_Time",
                "Next_Entry_Time",
                "Next_Result",
            ]
        ].tail(15).to_string(index=False)
    )


if __name__ == "__main__":
    main()
