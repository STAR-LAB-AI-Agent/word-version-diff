# -*- coding: utf-8 -*-
"""pytest 入口配置。

唯一职责：把“项目根”加入 sys.path，使测试可以 `import word_diff`。

* 项目根 = 本文件所在目录（tests/）的上一级，**不写死任何机器路径**；
* 换设备/换目录只要保持 `项目根/tests/conftest.py` 这个结构即可；
* 本文件不定义任何 pytest 专有 fixture，因此同一套测试也能用
  `python tests/run_tests.py`（不装 pytest）运行。
"""
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
