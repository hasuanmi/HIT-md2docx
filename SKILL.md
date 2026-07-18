---
name: HIT-md2docx
description: 哈工大（HIT）学位论文 Markdown→Word(.docx) 自动化排版工具，本科/硕士/博士通用。用户上传用 Markdown 写的论文，一键生成符合哈工大《学位论文格式要求及审查要点》（深圳研究生院）的标准 Word 文档——封面、中英文目录、页眉页脚、字体字号行距、LaTeX→Word 公式、参考文献顺序编码全部按规范出，并附格式问题清单。适用于哈工大本科生/硕士生/博士生把 Markdown 论文转成学校规范 Word，以及论文格式自查与修复。触发词：哈工大论文、学位论文排版、Markdown转Word、HIT docx、毕业论文格式、论文格式规范、研究生毕业论文。
---

# HIT-md2docx · 哈工大学位论文 Markdown→Word 自动化排版

把以 Markdown 撰写的哈工大学位论文（本科 / 硕士 / 博士），导出为符合学校规范的标准 Word(.docx)，并产出可读的格式问题清单。源文件用 Markdown，DOCX 是生成产物。

> 源仓库：`https://github.com/hasuanmi/HIT-md2docx`（本文件即仓库根入口）。所有代码、官方封面模板、范例均已随仓库提供。

## 运行环境（首次使用）
- Python ≥ 3.10：`python -m venv .venv && pip install -e .`（仓库根含 `pyproject.toml` / `requirements.txt`）
- 可选 Node.js：公式转 Word OMML 需 `npm install --prefix thesis_md2docx/math/latex2omml_node`；不装则公式保留 LaTeX 文本，不影响其余排版。

## 给 Agent 的固定执行流程

### 0 · 确认学位层次（必做）
生成前先问用户：本科 / 硕士（默认）/ 博士，记为 `DEGREE` ∈ {bachelor, master, doctor}。后续一律先按硕士模板生成，最后一步按 `DEGREE` 改写封面/页眉/页脚的学位文字。

### 1 ·（推荐）英文目录预翻译 —— 零外部 API
不同论文标题各异，**不靠写死字典**。借助 Agent 自己的 LLM 在生成前把标题翻好写进字典，脚本完全不调外部接口，因此零 key、离线、免费：

1. 导出待翻译标题（key 与引擎查表 key 完全一致）：
   ```bash
   python generate.py <论文>.md --dump-headings
   # 输出形如：{"1.1 研究背景": "", "参考文献": "", ...}
   ```
2. Agent 用自己的 LLM 逐条翻译 `key` → 英文值，**务必保留 key 原样**（动了就对不上、回退中文）。
3. 合并写回 `<论文>.md` 同目录的 `heading_translations.json`（不存在则创建；已存在则只补缺、不覆盖已有条目）。
4. 生成时跳过外部 LLM 兜底：
   ```bash
   python generate.py <论文>.md -o output/<论文>.docx --no-llm
   ```
> 行内 `| English` 标题拥有最高优先级（作者可覆盖 Agent 翻译）。本步骤只处理没有行内英文的标题。

### 2 · 生成 Word
`generate.py` 已串联：Markdown 预处理（中文章号归一 + 注入封面/声明前置 + 公式提取）→ 引擎渲染 → 官方封面模板注入 → 学位层次适配 → 字体/字号/行距规范修复。
```bash
# 硕士（默认）
python generate.py <论文>.md output/<论文>.docx
# 本科 / 博士（自动改写封面/页眉/页脚的「硕士/master」为「本科/bachelor」或「博士/doctor」）
python generate.py <论文>.md output/<论文>.docx --degree bachelor
python generate.py <论文>.md output/<论文>.docx --degree doctor
# 跳过官方封面模板注入（调试用）
python generate.py <论文>.md output/<论文>.docx --no-cover
```
输出即带哈工大官方 `input/封面.docx` 版式的完整 docx。

### 3 · 封面表格字段补填（如仓库含 `skill/scripts/fill_cover.py`）
官方封面模板的姓名/导师/学科/答辩日期等字段不会自动填；若仓库存在该脚本，可在生成后调用它按「标签定位 + 末单元格填值」从 front matter 补填。

### 4 · 格式校验与报告（可选）
```bash
python skill/scripts/audit_docx.py output/<论文>.docx --json audit.json
python skill/scripts/report.py audit.json --md 问题清单.md
```

## 交付与提醒
- 交付最终 docx + `问题清单.md`。
- 最终仍需 Word/WPS 人工复核：刷新目录域、分页、图、表、公式、参考文献、审查要点。

## 仓库结构
- `generate.py`：一键编排入口（本技能主调用）
- `thesis_md2docx/`：转换引擎（Markdown→IR→OOXML，含 `hit-master-thesis` profile）
- `skill/`：子流程脚本与知识库（`scripts/`、`references/`、`subskills/`）
- `input/`：官方封面模板 `封面.docx`、front_matter 前置模板
- `example/`：自包含范文（《中国制造2025》本科/硕士/博士三层次成品 docx）
