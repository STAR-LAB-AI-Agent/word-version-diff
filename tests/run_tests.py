#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""不依赖 pytest 的兜底运行器：把 tests/ 下的 test_*.py 全部跑一遍。

用法（在项目根目录执行）：
    python tests/run_tests.py                # 跑全部
    python tests/run_tests.py diff_engine    # 只跑文件名含该关键字的用例文件

说明：
* 与 pytest 跑的是**同一批用例**（都用普通 `assert`，没有 pytest 专有写法）；
* 装了 pytest 时更推荐：`python -m pytest tests -q`（有更详细的失败信息）；
* 退出码 0 = 全部通过，1 = 有失败。
"""
from __future__ import annotations

import importlib
import sys
import traceback
import unittest
from pathlib import Path

TESTS_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = TESTS_DIR.parent
for _p in (str(PROJECT_ROOT), str(TESTS_DIR)):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def discover(keyword: str | None = None):
    """按文件名（test_*.py）导入测试模块。"""
    modules = []
    for path in sorted(TESTS_DIR.glob("test_*.py")):
        if keyword and keyword not in path.stem:
            continue
        modules.append(importlib.import_module(path.stem))
    return modules


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    keyword = argv[0] if argv else None

    passed = failed = skipped = 0
    failures: list[tuple[str, str, str]] = []

    for module in discover(keyword):
        tests = [
            (name, func)
            for name, func in sorted(vars(module).items())
            if name.startswith("test_") and callable(func)
        ]
        if not tests:
            continue
        print(f"\n=== {module.__name__}（{len(tests)} 个用例）===")
        for name, func in tests:
            try:
                func()
            except unittest.SkipTest as exc:
                skipped += 1
                print(f"  [SKIP] {name}: {exc}")
            except AssertionError as exc:
                failed += 1
                failures.append((module.__name__, name, str(exc)))
                print(f"  [FAIL] {name}: {exc}")
            except Exception as exc:  # noqa: BLE001 —— 兜底：任何异常都算失败
                failed += 1
                failures.append((module.__name__, name, repr(exc)))
                print(f"  [ERROR] {name}: {exc!r}")
                traceback.print_exc(limit=3)
            else:
                passed += 1
                print(f"  [PASS] {name}")

    print("\n" + "-" * 64)
    print(f"通过 {passed}，失败 {failed}，跳过 {skipped}，合计 {passed + failed + skipped}")
    for module_name, name, msg in failures:
        lines = (msg or "").splitlines()
        preview = lines[0][:160] if lines else "（断言失败，无附加信息）"
        print(f"  失败：{module_name}.{name} -> {preview}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
