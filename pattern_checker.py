"""
Pattern Checker for TradingView 1-minute candle exports

What it does:
- Reads a TradingView XLSX/CSV trade export.
- Converts each completed LONG trade into:
    1 = exit price > entry price
    0 = exit price < entry price
    doji = exit price == entry price
- Dojis are NOT bridged. A pattern cannot jump across a doji.
- Checks patterns of exactly PATTERN_LENGTH candles.
- Keeps only patterns with at least MIN_OCCURRENCES valid appearances.
- Writes the results to a CSV file.

Install:
    pip install pandas openpyxl

Run:
    python pattern_checker.py

Then enter:
    1) the XLSX/CSV path
    2) pattern length
    3) minimum occurrences
"""

from pathlib import Path
from collections import defaultdict
import pandas as pd


def ask_int(prompt: str, default: int, minimum: int = 1) -> int:
    while True:
        raw = input(f"{prompt} [{default}]: ").strip()
        if not raw:
            return default
        try:
            value = int(raw)
            if value < minimum:
                raise ValueError
            return value
        except ValueError:
            print(f"Enter an integer >= {minimum}.")


def load_tradingview_file(file_path: Path) -> pd.DataFrame:
    suffix = file_path.suffix.lower()

    if suffix in {".xlsx", ".xlsm"}:
        xls = pd.ExcelFile(file_path)

        # TradingView strategy exports normally contain a "Trades" sheet.
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
        name = str(col).strip().lower()
        for prefix in startswith:
            if name.startswith(prefix.lower()):
                return col

    return None


def extract_binary_sequence(df: pd.DataFrame):
    """
    Returns a chronological sequence containing:
        1    positive candle/trade
        0    negative candle/trade
        None exact doji

    This uses ENTRY PRICE vs EXIT PRICE, not rounded Net PnL.
    That avoids a tiny positive/negative candle being incorrectly treated
    as zero because TradingView rounded the displayed PnL.
    """

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

    work = df[[trade_col, type_col, price_col] + ([time_col] if time_col else [])].copy()

    work[price_col] = pd.to_numeric(work[price_col], errors="coerce")
    work["_type"] = work[type_col].astype(str).str.strip().str.lower()

    records = []

    for trade_no, group in work.groupby(trade_col, sort=False):
        entry_rows = group[group["_type"].str.contains("entry", na=False)]
        exit_rows = group[group["_type"].str.contains("exit", na=False)]

        if entry_rows.empty or exit_rows.empty:
            continue

        entry_price = entry_rows.iloc[-1][price_col]
        exit_price = exit_rows.iloc[-1][price_col]

        if pd.isna(entry_price) or pd.isna(exit_price):
            continue

        if exit_price > entry_price:
            result = 1
        elif exit_price < entry_price:
            result = 0
        else:
            result = None

        if time_col:
            event_time = exit_rows.iloc[-1][time_col]
        else:
            event_time = trade_no

        records.append((event_time, trade_no, result))

    if not records:
        raise ValueError("No complete entry/exit trades were found.")

    # If a real timestamp column exists, sort chronologically.
    if time_col:
        temp = pd.DataFrame(records, columns=["time", "trade", "result"])
        parsed = pd.to_datetime(temp["time"], errors="coerce")

        # Only use timestamp sorting if parsing actually worked.
        if parsed.notna().any():
            temp["_parsed_time"] = parsed
            temp = temp.sort_values(["_parsed_time", "trade"], kind="stable")
        else:
            temp = temp.sort_values("trade", kind="stable")
    else:
        temp = pd.DataFrame(records, columns=["time", "trade", "result"])
        temp = temp.sort_values("trade", kind="stable")

    return [None if pd.isna(v) else int(v) for v in temp["result"].tolist()]


def scan_patterns(sequence, pattern_length: int):
    """
    Count what happened immediately after every pattern.

    Important:
    A None/doji breaks the sequence.
    Example:
        1, 0, None, 1
    does NOT create pattern 01 across the doji.
    """

    counts = defaultdict(lambda: {"next_0": 0, "next_1": 0})

    # target_index is the candle we are trying to observe after the pattern
    for target_index in range(pattern_length, len(sequence)):
        pattern_values = sequence[target_index - pattern_length:target_index]
        target = sequence[target_index]

        # Do not bridge dojis.
        if target is None or any(v is None for v in pattern_values):
            continue

        pattern = "".join(str(v) for v in pattern_values)

        if target == 1:
            counts[pattern]["next_1"] += 1
        else:
            counts[pattern]["next_0"] += 1

    return counts


def make_results(counts, min_occurrences: int) -> pd.DataFrame:
    rows = []

    for pattern, c in counts.items():
        next_0 = c["next_0"]
        next_1 = c["next_1"]
        occurrences = next_0 + next_1

        if occurrences < min_occurrences:
            continue

        p1 = next_1 / occurrences * 100.0
        p0 = next_0 / occurrences * 100.0

        if p1 > p0:
            most_likely = 1
            probability = p1
        elif p0 > p1:
            most_likely = 0
            probability = p0
        else:
            most_likely = "TIE"
            probability = 50.0

        rows.append({
            "Pattern": pattern,
            "Occurrences": occurrences,
            "Next_1_Count": next_1,
            "Next_0_Count": next_0,
            "Next_1_Pct": round(p1, 3),
            "Next_0_Pct": round(p0, 3),
            "Most_Likely_Next": most_likely,
            "Most_Likely_Pct": round(probability, 3),
            "Deviation_From_50_Pct": round(abs(p1 - 50.0), 3),
        })

    if not rows:
        return pd.DataFrame(columns=[
            "Pattern",
            "Occurrences",
            "Next_1_Count",
            "Next_0_Count",
            "Next_1_Pct",
            "Next_0_Pct",
            "Most_Likely_Next",
            "Most_Likely_Pct",
            "Deviation_From_50_Pct",
        ])

    result = pd.DataFrame(rows)

    # Strongest apparent pattern first.
    # Occurrences breaks ties so larger samples come first.
    result = result.sort_values(
        ["Deviation_From_50_Pct", "Occurrences"],
        ascending=[False, False]
    ).reset_index(drop=True)

    return result


def main():
    print("=" * 60)
    print("1-MIN CANDLE PATTERN CHECKER")
    print("=" * 60)

    raw_path = input("TradingView XLSX/CSV path: ").strip().strip('"')
    file_path = Path(raw_path).expanduser()

    if not file_path.exists():
        print(f"\nFile not found:\n{file_path}")
        return

    pattern_length = ask_int("Pattern length", default=5, minimum=1)
    min_occurrences = ask_int("Minimum occurrences", default=100, minimum=1)

    try:
        df = load_tradingview_file(file_path)
        sequence = extract_binary_sequence(df)
        counts = scan_patterns(sequence, pattern_length)
        results = make_results(counts, min_occurrences)
    except Exception as exc:
        print(f"\nERROR: {exc}")
        return

    green = sum(v == 1 for v in sequence)
    red = sum(v == 0 for v in sequence)
    doji = sum(v is None for v in sequence)

    output_path = file_path.with_name(
        f"patterns_len{pattern_length}_min{min_occurrences}.csv"
    )

    results.to_csv(output_path, index=False)

    print("\nDATA")
    print(f"Directional 1s : {green}")
    print(f"Directional 0s : {red}")
    print(f"Dojis / ties   : {doji}")
    print(f"Total trades   : {len(sequence)}")

    print("\nSETTINGS")
    print(f"Pattern length : {pattern_length}")
    print(f"Minimum count  : {min_occurrences}")

    print("\nRESULT")
    print(f"Patterns kept  : {len(results)}")
    print(f"CSV created    : {output_path}")

    if not results.empty:
        print("\nTop patterns:")
        print(
            results[
                [
                    "Pattern",
                    "Occurrences",
                    "Next_1_Pct",
                    "Next_0_Pct",
                    "Most_Likely_Next",
                    "Most_Likely_Pct",
                ]
            ].head(15).to_string(index=False)
        )
    else:
        print(
            "\nNo patterns met the minimum-occurrence setting. "
            "Lower MINIMUM OCCURRENCES or use a shorter pattern."
        )


if __name__ == "__main__":
    main()
