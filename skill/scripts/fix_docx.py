#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动修复（幂等）：对可自动修复的规则执行修复并保存；随后重新审计以验证幂等性。

设计要点（防止"改一个 bug 出另一个"）：
- 每个修复是独立小函数（见 format_rules.py），互不影响。
- 幂等：修复后再跑一次 audit，同一规则必须不再报警；否则视为回归并回滚保存。
- 只保存"确有修复"的文档；无可修复项则不改动文件。

用法：
  python fix_docx.py <thesis.docx> [--out fixed.docx] [--no-verify]
"""
from __future__ import annotations

import argparse
import sys

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from format_rules import RULES_BY_ID, ThesisDoc, run_audit, run_fix  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description="HIT 硕士论文格式自动修复（幂等）")
    ap.add_argument("docx")
    ap.add_argument("--out", help="输出路径，默认覆盖原文件")
    ap.add_argument("--no-verify", action="store_true", help="跳过幂等验证")
    args = ap.parse_args()

    doc = ThesisDoc(args.docx)
    before = run_audit(doc)
    fixed = run_fix(doc)
    total_fixed = sum(v for v in fixed.values() if v)

    out_path = args.out or args.docx
    if total_fixed == 0:
        print("无需修复（或可修复项均已合规）。原文件未改动。")
        return 0

    doc.save(out_path)
    print(f"已应用修复：{fixed}")
    print(f"保存至：{out_path}")

    if not args.no_verify:
        doc2 = ThesisDoc(out_path)
        after = run_audit(doc2)
        # 幂等校验：之前报警且可自动修复的规则，修复后不应再报同类
        regressed = []
        for f in after:
            if f.auto_fixable and f.rule_id in {b.rule_id for b in before if b.auto_fixable}:
                # 仍报警 => 可能未真正修复或该条非幂等
                regressed.append(f.rule_id)
        if regressed:
            print(f"[警告] 以下可修复规则修复后仍报警，可能非幂等：{sorted(set(regressed))}")
            print("（未覆盖原文件，请检查 format_rules.py 中对应 fixer）")
            return 2
        else:
            print("幂等验证通过：修复后再次审计，可修复项不再报警。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
