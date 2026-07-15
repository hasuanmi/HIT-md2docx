#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
规范校验（只读）：对生成的 HIT 硕士论文 docx 执行全部规则，输出 JSON 违规清单。

用法：
  python audit_docx.py <thesis.docx> [--json out.json]
若不指定 --json，则打印到 stdout。
"""
from __future__ import annotations

import argparse
import datetime
import json
import sys

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from format_rules import RULES, RULES_BY_ID, ThesisDoc, run_audit  # noqa: E402


def build_report(path: str) -> dict:
    doc = ThesisDoc(path)
    findings = run_audit(doc)
    errors = [f for f in findings if f.severity == "error"]
    warnings = [f for f in findings if f.severity == "warning"]
    auto = [f for f in findings if f.auto_fixable]
    return {
        "source": path,
        "generated_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "summary": {
            "total": len(findings),
            "errors": len(errors),
            "warnings": len(warnings),
            "auto_fixable": len(auto),
        },
        "rules_checked": [
            {"id": r.id, "category": r.category, "title": r.title,
             "severity": r.severity, "auto_fixable": r.auto_fixable}
            for r in RULES
        ],
        "findings": [
            {
                "rule_id": f.rule_id,
                "category": f.category,
                "severity": f.severity,
                "auto_fixable": f.auto_fixable,
                "message": f.message,
                "location": f.location,
                "rule_title": RULES_BY_ID[f.rule_id].title if f.rule_id in RULES_BY_ID else "",
            }
            for f in findings
        ],
    }


def main():
    ap = argparse.ArgumentParser(description="HIT 硕士论文格式校验（只读）")
    ap.add_argument("docx")
    ap.add_argument("--json", help="输出 JSON 路径")
    args = ap.parse_args()
    rep = build_report(args.docx)
    text = json.dumps(rep, ensure_ascii=False, indent=2)
    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            fh.write(text)
        print(f"已写出报告：{args.json}")
    else:
        print(text)
    # 有 error 时非零退出，方便脚本串联
    return 1 if rep["summary"]["errors"] else 0


if __name__ == "__main__":
    sys.exit(main())
