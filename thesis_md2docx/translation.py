"""英文目录的 LLM 翻译兜底。

设计目标：让工具对“任意用户的任意论文”都能生成英文目录，而不依赖一份写死的字典。

优先级（在 `_format_toc_english_number` 中实现）：
  1. 行内 `中文 | English` —— 作者显式给出，最高优先；
  2. `heading_translations.json`（用户手改/校对过的映射）—— 次优先；
  3. `heading_translations.llm.json`（本模块产出的 LLM 缓存）—— 再次；
  4. 配置了 API key 时，一次性批量调用大模型翻译所有缺失标题并写入缓存；
  5. 以上都没有 → 回退中文原文（并应由调用方告警）。

本模块只负责 3/4：读取配置、读写缓存、调用 OpenAI 兼容的 /chat/completions 接口。
未配置 key、无网络、接口报错时均优雅降级（返回空字典），绝不阻断主流程。
"""

from __future__ import annotations

import json
import os
from pathlib import Path

# 环境变量名（也可写在 md 同目录或 cwd 的 .env 里）
ENV_API_KEY = "HITMD2DOCX_LLM_API_KEY"
ENV_BASE_URL = "HITMD2DOCX_LLM_BASE_URL"
ENV_MODEL = "HITMD2DOCX_LLM_MODEL"
# 缓存文件名：与 heading_translations.json 放同一目录（markdown_dir）
LLM_CACHE_FILENAME = "heading_translations.llm.json"

_DEFAULT_BASE_URL = ""
_DEFAULT_MODEL = ""

_SYSTEM_PROMPT = (
    "You are a translator for Chinese academic thesis chapter headings. "
    "Translate each Chinese heading into concise, formal English suitable for a "
    "thesis table of contents. Return ONLY a JSON object that maps each original "
    "Chinese heading string exactly to its English translation. No explanations."
)


def _load_dotenv(path: Path | None) -> None:
    if not path or not path.exists():
        return
    try:
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            key, val = key.strip(), val.strip().strip('"').strip("'")
            if key:
                os.environ.setdefault(key, val)
    except Exception:
        return


def load_llm_config(markdown_dir: Path | None) -> dict | None:
    """读取 LLM 配置。未设置 API key 时返回 None（表示不启用兜底）。"""
    for d in (markdown_dir, Path.cwd()):
        if d:
            _load_dotenv(Path(d) / ".env")
    api_key = os.environ.get(ENV_API_KEY)
    if not api_key:
        return None
    base_url = os.environ.get(ENV_BASE_URL, _DEFAULT_BASE_URL)
    model = os.environ.get(ENV_MODEL, _DEFAULT_MODEL)
    if not base_url or not model:
        return None
    return {
        "api_key": api_key,
        "base_url": base_url,
        "model": model,
    }


def _cache_path(markdown_dir: Path | None) -> Path | None:
    if not markdown_dir:
        return None
    return Path(markdown_dir) / LLM_CACHE_FILENAME


def load_llm_cache(markdown_dir: Path | None) -> dict:
    p = _cache_path(markdown_dir)
    if p and p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_llm_cache(markdown_dir: Path | None, cache: dict) -> None:
    p = _cache_path(markdown_dir)
    if not p:
        return
    try:
        p.write_text(
            json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        return


def translate_headings(labels: list[str], *, config: dict) -> dict:
    """批量翻译中文标题为英文。成功返回 {原中文: 英文}，任何失败返回 {}。"""
    if not labels:
        return {}
    try:
        import requests  # 延迟导入：无网络/未安装时不影响主流程
    except Exception:
        return {}

    user_prompt = (
        "Translate these thesis headings into English. Return a JSON object only, "
        "mapping each original string to its English translation:\n"
        + json.dumps(labels, ensure_ascii=False)
    )
    try:
        resp = requests.post(
            config["base_url"].rstrip("/") + "/chat/completions",
            headers={
                "Authorization": f"Bearer {config['api_key']}",
                "Content-Type": "application/json",
            },
            json={
                "model": config["model"],
                "messages": [
                    {"role": "system", "content": _SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt},
                ],
                "temperature": 0.0,
                "response_format": {"type": "json_object"},
            },
            timeout=120,
        )
        resp.raise_for_status()
        content = resp.json()["choices"][0]["message"]["content"]
        data = json.loads(content)
    except Exception:
        return {}

    out: dict = {}
    label_set = set(labels)
    for k, v in data.items():
        if isinstance(v, str) and k in label_set:
            out[k] = v.strip()
    return out
