---
name: hit-md-to-richtext
description: 将论文 Markdown 转为结构化富文本——标题层级绑定 HIT Word 样式、清理 Markdown 标记（加粗/斜体/链接/列表/表格/公式）。在 HIT-md2docx 流程的第 2 步调用。
---

# MD 转结构化富文本 Skill（步骤 2）

把 Markdown 源映射为 HIT 论文的"语义富文本"，不直接落盘 docx（落盘由步骤 3 的模板填充完成）。

## 映射规则（与 `hit_master_thesis` profile 对齐）
- `#`→章（HitHeading1，"第一章"）、`##`→节（HitHeading2，"一、"）、`###`→条（HitHeading3，"（一）"）。
- `## 封面信息/声明/摘要/ABSTRACT/目录`→前 matter（HitFrontHeading / afa 目录标题）。
- `**粗体**`→黑体；`*斜体*`→斜体；行内 `$...$`→公式（OMML）。
- 表格→三线表（HitTableText）；图片→题注（HitCaption）。
- 引用 `[1]`、`[1-3]`、`[3]92` 保留为上标标记，待步骤 3 渲染。

## 步骤
1. 依据 `skill/references/example-md.md` 的结构解析 front matter 与正文。
2. 清理 Markdown 标记，生成"标题层级 + 样式意图"的中间表示。
3. 不做版式微调（行距/缩进/页眉由 profile 与步骤 4/5 负责），避免在此重复造轮子引发回归。

## 注意
- 不要手写 docx；所有样式由 profile 生成，保证字体/字号/行距/页眉双线一致。
- 公式缺失依赖时回退 LaTeX 文本并在报告中提示。

## 参考
- `skill/references/example-md.md` · `skill/references/markdown-usage.md` · `skill/references/template-spec.md`
