# -*- coding: utf-8 -*-
"""
Input parsing helpers for ticket_export_app.

当前阶段只作为后续架构重塑预留。
暂不接入 ui/export_ticket_window.py。
"""


def parse_float(value, default=0.0):
    try:
        return float(value)
    except Exception:
        return default
