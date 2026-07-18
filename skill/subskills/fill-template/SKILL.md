---
name: hit-fill-template
description: 基于 python-docx 底层能力，把结构化富文本填充进 HIT 硕士论文模板——保留分节、页眉双线、目录域。在 HIT-md2docx 流程的第 3 步调用。
---

# Docx 模板填充 Skill（步骤 3）

用 `thesis_md2docx` 引擎把步骤 2 的中间表示渲染为标准 HIT 硕士 docx。

## 步骤
1. 导出 DOCX（用启动器自动定位引擎；英文目录翻译表需与输入 md 同目录）：
   ```bash
   cp skill/heading_translations.json "$(dirname <论文>.md)/"
   bash skill/scripts/run_engine.sh docx <论文>.md output/<论文>.docx --profile hit-master-thesis
   ```
2. 导出 PDF / 预览页（Word 或 LibreOffice 后端）：
   ```bash
   bash skill/scripts/run_engine.sh all <论文>.md output/<论文>.docx output/<论文>.pdf --profile hit-master-thesis --backend auto
   ```
3. 校验产物继承了模板结构：分节（section）、页眉双线（HitHeader，`thinThickMediumGap`）、目录域（HitTocField）均存在。

## 关键不变量（不要破坏）
- 字体/字号/行距/页眉双线由 `profiles/hit_master_thesis/styles.py` 统一生成，**不要在此 Skill 手动改样式**——这是"改一个 bug 出另一个"的主要来源。
- 仅当新增学校/学位维度时，才扩展 profile；日常格式微调交给步骤 4/5 的校验与修复。

## 参考
- `skill/references/template-spec.md`（第 8 节样式 id） · `skill/scripts/export_docx.sh` · `skill/scripts/export_pdf.sh`
