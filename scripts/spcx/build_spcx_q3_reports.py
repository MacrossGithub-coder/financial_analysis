#!/usr/bin/env python3
"""Build both SPCX Q3 2026 earnings-update reports."""

from spcx_q3_report_common import build_report, fetch_market_data


if __name__ == "__main__":
    market = fetch_market_data()
    build_report("en", market)
    build_report("cn", market)
