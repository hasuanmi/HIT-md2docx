#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
导出报告：把 audit_docx.py 生成的 JSON 渲染为可读的 Markdown 格式问题清单。

用法：
  python report.py <audit.json> [--md out.md]
若不指定 --md，则打印到 stdout。
"""
from __future__ import annotations

import argparse
import json
import sys

SEV_LABEL = {"error": "❌ 错误（必须改）", "warning": "⚠️ 警告（建议改）"}


def render(rep: dict) -> str:
    s = rep.get("summary", {})
    lines = []
    lines.append("# 哈工大硕士学位论文 · 格式问题清单")
    lines.append("")
    lines.append(f"- 源文件：`{rep.get('source', '')}`")
    lines.append(f"- 生成时间：{rep.get('generated_at', '')}")
    lines.append(f"- 合计 **{s.get('total', 0)}** 项（错误 {s.get('errors', 0)} / 警告 {s.get('warnings', 0)}），"
                 f"其中可自动修复 **{s.get('auto_fixable', 0)}** 项。")
    lines.append("")

    findings = rep.get("findings", [])
    if not findings:
        lines.append("✅ 未检测到机检格式问题。仍建议用 Word/WPS 人工复核目录域、分页、图表与公式。")
        return "\n".join(lines)

    by_sev = {"error": [], "warning": []}
    for f in findings:
        by_sev.setdefault(f["severity"], []).append(f)

    for sev in ("error", "warning"):
        items = by_sev.get(sev, [])
        if not items:
            continue
        lines.append(f"## {SEV_LABEL.get(sev, sev)}（{len(items)}）")
        lines.append("")
        # 按规则聚合
        by_rule = {}
        for f in items:
            by_rule.setdefault(f["rule_id"], []).append(f)
        for rid, fs in by_rule.items():
            title = fs[0].get("rule_title") or rid
            lines.append(f"### [{rid}] {title}（{len(fs)} 处）")
            for f in fs:
                loc = f" _{f['location']}_" if f.get("location") else ""
                fixable = " · 可自动修复" if f.get("auto_fixable") else ""
                lines.append(f"- {f['message']}{loc}{fixable}")
            lines.append("")
    lines.append("---")
    lines.append("> 本清单由 `audit_docx.py` 机检生成，覆盖《硕士学位论文格式要求及审查要点》中"
                 "可机检的易错点。字体/字号/行距/页眉双线等由生成器保证，未列入；"
                 "目录域、分页、图表清晰度等仍需 Word/WPS 人工复核。")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="HIT 硕士论文格式问题清单（Markdown）")
    ap.add_argument("json")
    ap.add_argument("--md", help="输出 Markdown 路径")
    args = ap.parse_args()
    with open(args.json, encoding="utf-8") as fh:
        rep = json.load(fh)
    text = render(rep)
    if args.md:
        with open(args.md, "w", encoding="utf-8") as fh:
            fh.write(text)
        print(f"已写出清单：{args.md}")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
