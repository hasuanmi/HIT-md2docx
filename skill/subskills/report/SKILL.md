---
name: hit-report
description: 导出格式问题清单（Markdown/HTML），汇总审计结果与修复建议，便于人工逐项核对。在 HITmd2docx 流程的第 6 步调用。
---

# 导出报告 Skill（步骤 6）

把审计 JSON 渲染为可读的"格式问题清单"，作为交付物之一。

## 步骤
1. 由审计 JSON 生成清单：
   ```bash
   python skill/scripts/report.py audit.json --md 问题清单.md
   ```
2. 在清单末尾附：复核提示（目录域、分页、图表清晰度、字体观感需在 Word/WPS 人工确认）。
3. 与最终 docx 一并交付用户。

## 清单结构
- 头部：源文件、生成时间、合计（错误/警告/可自动修复）。
- 按严重度分组（❌错误 / ⚠️警告），按规则聚合，列出每处位置（段落索引）与是否可自动修复。
- 尾部：机检范围说明与人工复核提醒。

## 参考
- `skill/scripts/report.py` · `references/selfcheck-checklist.md`
