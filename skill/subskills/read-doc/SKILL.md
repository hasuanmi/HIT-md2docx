---
name: hit-read-doc
description: 读取论文输入与范例资料——Markdown 源、学校范例 docx、自查清单（PDF/Word）。提取结构、核对生成文档是否继承模板样式。在 HITmd2docx 流程的第 1 步调用。
---

# 文档读取 Skill（步骤 1）

读取三类资料并做预处理，供后续 Skill 使用。

## 输入
- 用户上传的论文 Markdown（优先）或已有 Word。
- 学校范例 docx（`thesis_md2docx/profiles/hit_master_thesis/format_requirements/` 下或用户提供的模板）。
- 自查清单（`references/selfcheck-checklist.md`、`hit-master-review-points.md`）。

## 步骤
1. 定位仓库根（含 `md2docx.py` / `thesis_md2docx` 包），确认 profile = `hit-master-thesis`。
2. 检查环境：`bash skill/scripts/check_env.sh`。
3. 解析 Markdown 结构：缺封面/章节/参考文献/图表/附录时，加载 `references/example-md.md` 与 `references/markdown-usage.md` 提示用户补齐。
4. 若提供范例 docx，用 python-docx 读取其样式表，与 `references/template-spec.md` 第 8 节列出的 HIT 自定义样式 id 对照，确认生成器目标样式一致。
5. 若 Markdown 含公式，安装公式助手（否则公式回退 LaTeX 文本）：
   `npm install --prefix thesis_md2docx/math/latex2omml_node`

## 输出
- 结构化的输入清单（章节树、是否含摘要/关键词/参考文献/图/表/附录）。
- 供步骤 2/3 使用的规范化 Markdown。

## 参考
- `references/template-spec.md` · `references/example-md.md` · `references/markdown-usage.md`
