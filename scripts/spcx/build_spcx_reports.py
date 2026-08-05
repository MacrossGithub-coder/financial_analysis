#!/usr/bin/env python3
"""Build the English and Chinese SPCX reports from one market-data snapshot."""

from spcx_q2_report_common import build_report, fetch_market_data


if __name__ == "__main__":
    market = fetch_market_data()
    build_report("en", market)
    build_report("cn", market)
