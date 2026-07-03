"""Benchmarks are collected only when explicitly selected via ``-m benchmark``.

We do not depend on ``pytest-benchmark``; each module measures wall time with
:func:`time.perf_counter_ns` and asserts a generous local budget.
"""
