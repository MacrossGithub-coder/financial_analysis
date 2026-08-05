#!/usr/bin/env python3
"""生成中文版SPCX 2026年第二季度业绩更新报告。"""

from spcx_q2_report_common import build_report


if __name__ == "__main__":
    build_report("cn")
