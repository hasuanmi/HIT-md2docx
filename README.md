# hit-thesis-formatter

哈工大（HIT）学位论文 **Markdown → Word (.docx) / PDF** 转换工具，本科 / 硕士 / 博士通用。

写作用 Markdown，提交用 Word，检查版式用 PDF。把学校格式的「苦活」交给工具：章节编号、中英文目录、封面、页眉页脚、字体字号行距、公式（LaTeX→Word OMML）、参考文献按引用顺序编码……全部按哈工大《学位论文格式要求及审查要点》出。

> 底层是一个**纯 Python 命令行工具**，任何同学都能 `pip install` 后直接在终端使用；WorkBuddy 用户还能用附带的 `skill/` 获得「一键生成 + 格式自检」的引导式体验。两者解耦，互不依赖。

---

## 特性

- [x] 内置 **`hit-master-thesis`** profile（哈工大硕士模板，已修复章节错位 / 中文目录问题），默认即生效
- [x] 支持本科 / 硕士 / 博士三层次：生成后一键把「硕士 / master」改写为「本科 / bachelor」或「博士 / doctor」
- [x] 封面、声明、摘要、目录（中英文）、正文、参考文献、致谢、附录
- [x] 标题 / 正文 / 图表 / 公式 / 参考文献等常用论文格式
- [x] 中文章号自动归一为数字章号（`第一章` → `# 1 标题`），参考文献按正文首次引用顺序重排（GB/T 7714 顺序编码制）
- [x] LaTeX 公式转 Word OMML（缺依赖时保留 LaTeX 文本）
- [x] DOCX → PDF：Microsoft Word 后端 / LibreOffice 后端
- [x] DOCX 格式审计，逐项对照学校范例
- [x] profile 可扩展其他学校 / 学位

---

## 目录结构

```
hit-thesis-formatter/
├── thesis_md2docx/          # 转换引擎（已修复的 HIT profile）
├── skill/                   # WorkBuddy skill 编排层（可选，WB 用户用）
│   ├── SKILL.md
│   ├── subskills/           # read-doc / md-to-richtext / fill-template / cover-inject / audit / autofix / report / degree-adapt
│   ├── scripts/             # md_to_fordocx.py / cover_inject.py / degree_adapt.py / audit_docx.py / fix_docx.py / run_engine.*
│   ├── references/          # 哈工大格式规范、自查清单、范例 MD
│   └── agents/openai.yaml
├── input/                   # 资源：封面.docx（官方封面模板）+ heading_translations.json（中英文目录映射）
├── example/                 # 自包含样例：thesis-demo-hit.md + 图片 + 翻译表
├── md2docx.py               # CLI 入口
├── 模板.docx                # 哈工大论文完整模板（参考用）
├── pyproject.toml / requirements.txt / LICENSE / README.md / .gitignore
```

---

## 快速开始（纯 CLI，人人可用）

### 1. 安装

需要 Python ≥ 3.10。

```bash
git clone https://github.com/haodongcui/hit-thesis-formatter.git
cd hit-thesis-formatter
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\Activate.ps1
pip install -e .                                   # 可编辑安装，得到 md2docx 命令

# 公式转 Word OMML 需要 Node.js（可选；不装则公式保留 LaTeX 文本）：
npm install --prefix thesis_md2docx/math/latex2omml_node
```

### 2. 写论文（Markdown）

- **章节用数字章号**：`# 1 绪论`、`# 2 相关的工作`……引擎只把 `# N 标题` 当作一级章。
  - 如果你的 md 用的是中文章号（`第一章`、`第 2 章`），先跑预处理：
    ```bash
    python skill/scripts/md_to_fordocx.py 你的论文.md
    # 产出 你的论文_for_docx.md（已归一为数字章号）
    ```
- **封面信息**放在文首的 `## 封面信息` 块（见 `example/thesis-demo-hit.md`）：论文题目、作者、学号、学科专业、指导教师、答辩日期、学位类别、学校代码等。
- **英文目录**：把 `input/heading_translations.json` 复制到你的 md **同目录**下，引擎会照它把中文标题翻成英文写进 Contents。

### 3. 生成 Word

```bash
# 方式 A：直接用 md2docx 命令
md2docx docx --profile hit-master-thesis 你的论文.md 论文.docx

# 方式 B：用 python -m 调引擎（等价）
python -m thesis_md2docx.main docx --profile hit-master-thesis 你的论文.md 论文.docx
# 注：docx 子命令的输出是「位置参数」（不是 -o），省略时默认生成 <输入同名>.docx
```

默认 profile 就是 `hit-master-thesis`，`--profile` 可省略。

### 4.（可选）注入官方封面 / 层次适配

```bash
# 用学校官方封面模板覆盖生成的封面区（找不到模板会自动跳过，不报错）
python skill/scripts/cover_inject.py 论文.docx --cover-template input/封面.docx

# 本科 / 博士：把封面页眉页脚的「硕士/master」改写为「本科/bachelor」或「博士/doctor」
python skill/scripts/degree_adapt.py --degree bachelor 论文.docx
python skill/scripts/degree_adapt.py --degree doctor  论文.docx
# master 时为空操作，可跳过
```

### 5. 导出 PDF（可选）

```bash
# all 子命令：md → docx → pdf（pdf 子命令只接受 docx 输入，这里用 all 一步到位）
md2docx all --profile hit-master-thesis 你的论文.md 论文.docx 论文.pdf --backend auto
```

---

## 给其他 AI Agent（Codex / Coomate 等）

本工具是**标准命令行程序**，任何能在终端执行命令的 agent 都能直接调用，无需 WorkBuddy：

```
python -m thesis_md2docx.main docx --profile hit-master-thesis <md> <docx>
```

`skill/` 目录是 WorkBuddy 私有的 skill 格式，Codex / Coomate 不会读取它——它们应当像人一样直接调用上面的 CLI。

---

## 用 WorkBuddy 的「一键」体验（可选）

如果你用 WorkBuddy，把本仓库作为一个 skill 来源即可：将 `skill/` 作为 skill 目录加载（或通过专家中心导入），然后对它说「帮我把这篇 Markdown 转成符合哈工大规范的 Word」，它会按固定流程走完：确认层次 → 读取 → MD 预处理 → 模板填充 → 封面注入 → 规范校验 → 自动修复 → 层次适配 → 导出报告。

---

## 常见问题

- **英文目录没出现 / 是中文？** 确认 `heading_translations.json` 和你的 md 在**同一目录**。
- **封面版式不对？** 把学校下发的官方 `封面.docx` 放到 `input/封面.docx`，再跑 `cover_inject.py`；找不到模板时引擎会优雅跳过，仅生成基础封面。
- **公式显示为 LaTeX 原文而非公式？** 需要 Node.js 环境（引擎用 `math/latex2omml_node` 转 OMML）；缺失时保留原文，不影响其余排版。
- **章节序号错乱？** 检查一级标题是否写成 `# 数字 标题`；中文数字章号请先跑 `md_to_fordocx.py`。

---

## 许可证

MIT — 见 [LICENSE](LICENSE)。
