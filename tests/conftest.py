# -*- coding: utf-8 -*-
"""讓 tests/ 內的測試可以直接 import 上層的 app.py。"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
