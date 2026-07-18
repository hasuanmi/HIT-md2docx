"""Runtime asset fetching for published (binary-free) builds.

The RedSkill / SkillHub distribution cannot carry binary assets
(.docx / .jpeg / .png). When a required runtime asset is missing
locally, fetch it from the project's GitHub raw (main branch) so the
binary-free build still runs end-to-end.
"""
from __future__ import annotations

import os
import urllib.request
from urllib.parse import quote

# 发布版（小红书 SkillHub 等）不含二进制资源，运行时按需从 GitHub raw(main) 补齐。
GITHUB_RAW_BASE = "https://raw.githubusercontent.com/hasuanmi/HIT-md2docx/main"


def ensure_asset(rel_path: str, base_dir: str) -> str | None:
    """Return local path to ``rel_path`` under ``base_dir``.

    If the file is missing locally, download it from the project's GitHub
    raw (main) and return the local path. Returns ``None`` if the file is
    absent and the download fails, so callers can degrade gracefully.
    """
    local = os.path.join(base_dir, rel_path)
    if os.path.isfile(local):
        return local
    encoded = "/".join(quote(seg, safe="") for seg in rel_path.split("/"))
    url = f"{GITHUB_RAW_BASE}/{encoded}"
    os.makedirs(os.path.dirname(local) or ".", exist_ok=True)
    print(f"  [资源补齐] 本地缺失 {rel_path}，从 GitHub 下载…")
    try:
        urllib.request.urlretrieve(url, local)
    except Exception as exc:  # pragma: no cover - network / availability
        print(f"  [警告] 资源下载失败 {rel_path}：{exc}")
        return None
    return local
