# HIT-md2docx

哈工大（HIT）学位论文 **Markdown → Word (.docx) / PDF** 转换工具，本科 / 硕士 / 博士通用。

写作用 Markdown，提交用 Word,把学校格式的「苦活」交给工具：章节编号、中英文目录、封面、页眉页脚、字体字号行距、公式（LaTeX→Word OMML）、参考文献按引用顺序编码……全部按哈工大《学位论文格式要求及审查要点》出。

---

## 特性

- [x] 内置 **`hit-master-thesis`** profile（哈工大硕士模板），默认即生效
- [x] 支持本科 / 硕士 / 博士三层次
- [x] 封面、声明、摘要、目录、正文、参考文献、致谢、附录
- [x] 标题 / 正文 / 图表 / 公式 / 参考文献等常用论文格式
- [x] 参考文献按正文首次引用顺序重排（GB/T 7714 顺序编码制）
- [x] LaTeX 公式转 Word OMML（缺依赖时保留 LaTeX 文本）
- [x] DOCX → PDF：Microsoft Word 后端 / LibreOffice 后端
- [x] DOCX 格式审计，逐项对照学校范例

---

## 目录结构

```
HIT-md2docx/
├── thesis_md2docx/          # 转换引擎
├── skill/                   # WorkBuddy skill 编排层
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

## 快速开始

### 1. 安装

需要 Python ≥ 3.10。

```bash
git clone https://github.com/484899614-shipi-it/HIT-md2docx.git
cd HIT-md2docx
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\Activate.ps1
pip install -e .                                   # 可编辑安装，得到 md2docx 命令

# 公式转 Word OMML 需要 Node.js（可选；不装则公式保留 LaTeX 文本）：
npm install --prefix thesis_md2docx/math/latex2omml_node
```

### 2. 生成 Word

```bash
# 方式 A：直接用 md2docx 命令
md2docx docx --profile hit-master-thesis 你的论文.md 论文.docx

# 方式 B：用 python -m 调引擎
python -m thesis_md2docx.main docx --profile hit-master-thesis 你的论文.md 论文.docx
# 注：docx 子命令的输出是「位置参数」（不是 -o），省略时默认生成 <输入同名>.docx
```

默认 profile 就是 `hit-master-thesis`，`--profile` 可省略。

### 3. 导出 PDF（可选）

```bash
# all 子命令：md → docx → pdf（pdf 子命令只接受 docx 输入，这里用 all 一步到位）
md2docx all --profile hit-master-thesis 你的论文.md 论文.docx 论文.pdf --backend auto
```

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
