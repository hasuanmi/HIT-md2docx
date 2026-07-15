---
name: hit-autofix
description: 批量修正可自动修复的格式问题（R01/R02/R05/R06 等幂等修复），并更新目录域。在 HITmd2docx 流程的第 5 步调用。
---

# 自动修复 Skill（步骤 5）

对审计中标记"可自动修复"的项做**幂等**修复，并刷新目录域。

## 步骤
1. 运行幂等修复（修复后自动重新审计做幂等验证）：
   ```bash
   python skill/scripts/fix_docx.py output/<论文>.docx --out output/<论文>.fixed.docx
   ```
2. 刷新目录域（需 Word/WPS 或 LibreOffice 后端）：
   - 推荐在 Word/WPS 中打开后 `Ctrl+A` → 右键"更新域" → 更新整个目录。
   - 若环境有无头 LibreOffice，可调用其更新索引（best-effort），否则在报告中提示人工更新。
3. 修复后重新跑步骤 4 审计，确认可修复项归零且无新增。

## 为什么安全（防回归）
- 每个修复是 `format_rules.py` 中独立小函数，互不影响。
- 幂等：连续修复两次，第二次改动数为 0；`tests/test_audit_fix.py` 用带已知问题的样例断言"检出→修复→再审计不变"。
- 只改"依赖作者内容、生成器保证不了"的文本类问题；不动字体/字号/行距/页眉等已验证样式。

## 不建议自动修复的项
- R03/R04/R07/R08/R09/R10/R11 为仅检测项，自动改易引入新错，交人工复核。

## 参考
- `skill/scripts/format_rules.py` · `skill/scripts/fix_docx.py` · `tests/test_audit_fix.py`
