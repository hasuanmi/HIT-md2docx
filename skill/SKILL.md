---
name: HITmd2docx
description: 毕业论文格式化 Agent（哈工大本科/硕士/博士通用）。上传论文 Markdown，一键生成符合哈工大《学位论文格式要求及审查要点》（深圳研究生院）的标准 Word 文档，并输出格式问题清单。支持本科/硕士/博士三种层次——生成前先与用户确认层次，封面/页眉/页脚的「硕士/master」自动改写为「本科/bachelor」或「博士/doctor」。固定流程：确认层次 → 文档读取 → MD转富文本 → 模板填充 → 封面注入 → 规范校验 → 自动修复 → 层次适配 → 导出报告。适用于 HIT 学位论文 Markdown→DOCX/PDF 的撰写、检查与导出。
---

# 毕业论文格式化 Agent · HIT 学位论文（Skill 集合总编排）

把以 Markdown 撰写的哈工大学位论文（本科/硕士/博士），导出为符合学校规范的标准 Word 文档，并产出可读的格式问题清单。源文件优先用 Markdown，DOCX/PDF 视为生成产物。

> **学位层次（本科/硕士/博士）**：默认按硕士模板（`hit-master-thesis`）生成。若用户确认是**本科**或**博士**，最后一步 `degree_adapt.py` 会把封面/页眉/页脚/后置小节标题中的「硕士 / master」改写为「本科 / bachelor」或「博士 / doctor」（含封面模板中被拆成多跑的「硕士」）。master 时为空操作。

## 三份长期知识库（绑定到本 Agent）

1. **范例 MD 样板** — `references/example-md.md`：标准论文 Markdown 结构范本（入口即照此组织输入）。
2. **范例 docx 规范** — `references/template-spec.md`：哈工大硕士格式规范（字体/字号/行距/页眉/编号/版芯），含 HIT 自定义样式 id 对照。
3. **自查清单** — `references/selfcheck-checklist.md`（权威机检映射）+ `references/hit-master-review-points.md`（完整审查要点与往届 10 大易错点）。

> ⚠️ `hit_guide_clean.txt` 是**本科**指南，不代表硕士要求，禁止用于本 Agent 校验。

## 固定执行流程（严格按 0→7 串联）

> 每一步对应 `skill/subskills/<name>/SKILL.md`，出错即停、不跳步。

### 0 · 确认学位层次（与用户交互，必做）
**开始任何生成前，先向用户确认论文层次**：本科 / 硕士（默认）/ 博士。用 `AskUserQuestion` 或直接询问，记录答案 `DEGREE`（取值 `bachelor` / `master` / `doctor`）。后续步骤一律按**硕士模板**生成，最后一步（7）再按 `DEGREE` 做层次适配。`DEGREE=master` 时步骤 7 为空操作。

### 1 · 文档读取（`subskills/read-doc`）
定位仓库、检查环境、解析 Markdown 结构、对照范例 docx 与目标样式一致。输入缺项时引用 `example-md.md` 提示补齐。

### 2 · MD 转结构化富文本（`subskills/md-to-richtext`）
标题层级绑定 HIT Word 样式、清理 Markdown 标记，生成"语义富文本"中间表示（不直接落盘）。

> **固定/可变两部分分离（md 层，零引擎改动）**：论文格式分两类——①固定部分（封面信息 `## 封面信息`、原创性声明 `## 声明`，几乎照搬模板，只随题目变）；②可变部分（摘要/目录/正文）。若正文 md 缺 `## 封面信息`，引擎读不到封面数据 → 英文题目变空行、字段回退占位符（"封面生成不对"）。预处理脚本 `md_to_fordocx.py` 会**自动注入固定前置模板** `input/front_matter_hit.md`（并用正文 H1 题目同步"论文题目："），一次维护、每次生成复用。
> ```bash
> # 中文章号归一 + 自动注入固定前置（封面+声明）+ 公式提取
> python skill/scripts/md_to_fordocx.py <论文>.md -o input/<论文>_for_docx.md
> #   --front-matter <路径>   指定固定前置模板
> #   --no-front-matter       关闭注入（正文已自带 ## 封面信息 时无需注入）
> ```
> **公式提取（关键）**：`md_to_fordocx.py` 内置调用 `preprocess_equations.convert_text`，
> 把源稿里的**纯文本计量公式**（`RD_it = α₀ + …  (1)`、`Σ_{k=-3}^{3}`、`Step 1（总效应）：…` 等）
> 转成 `$$…$$` 块；引擎的 `MathConverter`（temml→mathml2omml）再把它们渲染成 **Word 原生 OMML 公式**
> （可双击编辑），编号 `(N)` 由引擎提取为**右对齐公式号**（保留作者原编号）。
> ⚠️ 若公式在成品里是纯文本/LaTeX 源码，说明预处理未跑或 node 公式依赖缺失——
> 务必重跑 `md_to_fordocx.py` 生成新的 `_for_docx.md`，并确保 `math/latex2omml_node` 下已 `npm install`。

### 3 · Docx 模板填充（`subskills/fill-template`）
用 `thesis_md2docx` 引擎（`--profile hit-master-thesis`）渲染标准 docx，保留分节、页眉双线、目录域。英文目录标题 **Contents** 由 `build_front_heading(..., bold=True)` 生成，确保与中文“目  录”黑体视觉一致。
通过 `skill/scripts/run_engine.sh`（或 `run_engine.py`）自动定位本机引擎——优先 `THESIS_MD2DOCX_REPO` 环境变量，否则回退到本仓库根目录（即 `skill/` 的父目录，含 `thesis_md2docx/`）。
```bash
# 英文目录翻译表需与输入 md 同目录(引擎按其所在目录查找)
cp skill/heading_translations.json "$(dirname <论文>.md)/"
# 生成 docx
bash skill/scripts/run_engine.sh docx <论文>.md output/<论文>.docx --profile hit-master-thesis
# （如需 PDF 预览）
bash skill/scripts/run_engine.sh all <论文>.md output/<论文>.docx output/<论文>.pdf --profile hit-master-thesis --backend auto
```

### 3.5 · 封面与前置部分版式注入（`subskills/cover-inject`）
把官方封面模板 `input/封面.docx` 的版式注入论文（整页复制 + 替换题目），并在目录插入摘要/Abstract 罗马页码、对齐前置部分行距。**不改动引擎**。找不到模板则跳过封面复制，流程可无模板运行。
```bash
python skill/scripts/cover_inject.py output/<论文>.docx \
    --cover-template input/封面.docx --front-matter input/front_matter_hit.md \
    --out output/<论文>.docx
```
> 封面占位字段（作者/导师/学科/分类号/答辩日期）由 `front_matter_hit.md` 驱动；重跑步骤 3 会覆盖封面，需再跑本步骤恢复模板版式。

### 4 · 规范校验（`subskills/audit`）
绑定自查清单，机检 R01–R11，输出 JSON 与可读清单。
```bash
python skill/scripts/audit_docx.py output/<论文>.docx --json audit.json
python skill/scripts/report.py audit.json --md 问题清单.md
```

### 5 · 自动修复（`subskills/autofix`）
对可自动修复项（R01/R02/R05/R06/R12 等）做**幂等**修复并刷新目录域；仅检测项交人工。
- R12：正文（变量/缩写/公式中的拉丁字母）未用 Times New Roman 时自动补设。
```bash
python skill/scripts/fix_docx.py output/<论文>.docx --out output/<论文>.fixed.docx
```

### 6 · 导出报告（`subskills/report`）
交付最终 docx + `问题清单.md`，并提示 Word/WPS 人工复核（目录域、分页、图表清晰度、字体观感）。

### 7 · 学位层次适配（`subskills/degree-adapt`）
若步骤 0 确认的 `DEGREE≠master`，把已生成的硕士格式 docx 做封面临门一脚：封面（含「硕士学位论文」「硕士研究生：」、英文「Master Degree」）、页眉（默认「哈尔滨工业大学硕士学位论文」、后置「攻读硕士学位期间取得的科研成果」）、页脚中的「硕士 / master」改写为 `DEGREE` 对应词。`DEGREE=master` 时跳过（空操作）。
```bash
# bachelor / doctor 时改写为 本科/bachelor 或 博士/doctor；master 时原样保留
python skill/scripts/degree_adapt.py output/<论文>.docx --degree <bachelor|master|doctor> --out output/<论文>.docx
```
> 实现为 docx 层后处理（不改引擎）：对每个块做「逻辑文本重组 + 位置 preserving 重排」，可修正封面模板把「硕士」拆成多跑（如「硕」+「士研究生」）导致的漏替换，同时尽量保留原有字形。

## 防回归设计（针对"改一个 bug 出另一个"）

- **生成与校验解耦**：字体/字号/行距/页眉双线由 `profiles/hit_master_thesis/styles.py` 生成，校验/修复只处理"依赖作者内容、生成器保证不了"的项（R01–R11）。
- **规则数据驱动**：所有机检规则集中在 `skill/scripts/format_rules.py`（单一真相来源）；改规则不碰生成代码。
- **幂等 + 测试**：`tests/test_audit_fix.py` 用带已知问题的样例 docx 断言"检出→修复→再审计不变"，任何回归立即暴露。
- **仅检测项人工复核**：R03/R04/R07/R08/R09/R10/R11 不自动改，避免引入新错。

## 操作规则

- 不要把生成的 DOCX 当长期源文件手动改；改 Markdown / profile 代码 / 本 Skill 脚本。
- 生成文件放 `output/`，源文件放论文根目录。
- 哈工大硕士输出用 profile `hit-master-thesis`。
- 最终仍需 Word/WPS 人工检查：刷新目录域、分页、图、表、公式、参考文献、审查要点。

## References

- `references/example-md.md` · `references/template-spec.md` · `references/selfcheck-checklist.md` · `references/hit-master-review-points.md`
- `references/markdown-usage.md` · `references/pdf-backends.md` · `references/troubleshooting.md`
- `skill/scripts/`：`md_to_fordocx.py`（预处理：章号归一 + 注入固定前置 + **按角标排序参考文献**）/ `run_engine.sh` / `run_engine.py`（引擎启动器，自动定位 `thesis_md2docx`）/ `audit_docx.py` / `fix_docx.py` / `report.py` / `format_rules.py` / `cover_inject.py`（封面与前置部分版式注入）/ `degree_adapt.py`（学位层次适配：硕士→本科/博士，master→bachelor/doctor）/ `check_env.sh` / `export_docx.sh` / `export_pdf.sh`
- `skill/subskills/`：read-doc / md-to-richtext / fill-template / cover-inject / audit / autofix / report / degree-adapt
