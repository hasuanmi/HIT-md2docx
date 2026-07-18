#!/usr/bin/env python3
# 毕业论文格式化 Agent · 引擎启动器(跨平台 Python 版)
# 自动定位 thesis_md2docx 引擎并加入 sys.path, 再转发参数给 thesis_md2docx.main
import os
import sys
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))   # .../skill/scripts
SKILL_DIR = os.path.dirname(HERE)                    # .../skill
REPO_ROOT = os.path.dirname(SKILL_DIR)               # 仓库根(thesis_md2docx 所在目录)
CANDIDATES = [
    os.environ.get("THESIS_MD2DOCX_REPO", ""),
    REPO_ROOT,
    SKILL_DIR,
]
REPO = ""
for c in CANDIDATES:
    if c and os.path.isdir(os.path.join(c, "thesis_md2docx")):
        REPO = c
        break
if not REPO:
    sys.stderr.write(
        "ERROR: 找不到 thesis_md2docx 引擎。请设置 THESIS_MD2DOCX_REPO 指向含 thesis_md2docx/ 的目录。\n"
    )
    sys.exit(2)
if REPO not in sys.path:
    sys.path.insert(0, REPO)
# 直接调用引擎入口
from thesis_md2docx import main as engine_main
sys.exit(engine_main.main())
