# Quant Pattern Research

A small quantitative research project for analyzing short binary candle sequences in 1-minute market data.

## Features

- TradingView XLSX/CSV import
- Binary candle encoding (`1` = positive candle, `0` = negative candle)
- Doji handling
- Exact pattern frequency analysis
- Minimum occurrence filtering
- Historical occurrence inspection
- Recent-sample validation
- Z-score / stability checks
- CSV exports

## Research Goal

Test whether short binary price sequences show statistically persistent next-candle behavior.

The public repository contains the research and validation tools only. Private signal-selection rules and live trading logic are intentionally excluded.

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python pattern_checker.py
python pattern_trade_list.py
python pattern_next_results_only.py
python pattern_present_verify.py
```

## Data

Only a tiny generic sample file is included. Real market datasets and private research inputs are excluded.

## Status

Early-stage quantitative research project.

## Disclaimer

For research and educational purposes only. Historical patterns do not guarantee future results.
