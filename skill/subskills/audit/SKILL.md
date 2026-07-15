---
name: hit-audit
description: 绑定自查知识库，逐条检查行距、缩进、参考文献、页码、图表格编号等可机检项，输出 JSON 违规清单。在 HITmd2docx 流程的第 4 步调用。
---

# 规范校验 Skill（步骤 4）

对生成的 docx 执行机检，逐条对照 `references/selfcheck-checklist.md`（对应 `skill/scripts/format_rules.py` 的 R01–R11）。

## 步骤
1. 运行只读审计：
   ```bash
   python skill/scripts/audit_docx.py output/<论文>.docx --json audit.json
   ```
2. 生成可读清单：
   ```bash
   python skill/scripts/report.py audit.json --md 问题清单.md
   ```
3. 按清单核对：错误（必须改）/ 警告（建议改）/ 可自动修复 三类分别处理。

## 检查项（R01–R11，详见 selfcheck-checklist.md）
- 关键词分隔符与大小写（R01）、数字与单位空格（R02）、图/表编号含章号（R03）、编号连续（R04）、
  直接引用不上标（R05）、英文摘要标题大小写（R06）、结论不标引（R07）、标题后不直插（R08）、
  参考文献数量与比例（R09）、坐标单位括号（R10）、"如下图"改"如图"（R11）。

## 原则
- 字体/字号/行距/页眉双线由生成器保证，本步不重复校验，避免回归。
- 仅检测项（R03/R04/R07/R08/R09/R10/R11）交人工在 Word/WPS 复核；目录域、分页、图表清晰度也需人工。
- 有 `error` 时脚本非零退出，便于在自动化流程中阻断。

## 参考
- `references/selfcheck-checklist.md` · `references/hit-master-review-points.md` · `skill/scripts/format_rules.py`
