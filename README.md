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
├── example/                 # 自包含样例
│   ├── thesis-demo-hit.md                               # 最小可跑 demo（含图片、翻译表）
│   ├── 论文_平台反垄断与研发投入.md                       # 真实范文：md 源（引擎可直接消费）
│   └── 论文_平台反垄断与研发投入_HIT硕士规范版.docx       # 同上范文的成品 docx，供审阅/对照版式
├── thesis-specs/            # 论文规范文件：模板.docx / 封面.docx / 硕士学位论文格式要求及审查要点.docx
├── md2docx.py               # CLI 入口
├── 模板.docx                # 哈工大论文完整模板（参考用）
├── pyproject.toml / requirements.txt / LICENSE / README.md / .gitignore
```

> **想直接看效果 / 对照规范？**
> - 真实范文（md + 成品 docx）已放进 `example/`：`论文_平台反垄断与研发投入.md` 与 `论文_平台反垄断与研发投入_HIT硕士规范版.docx`，clone 后即可审阅版式。
> - 哈工大官方模板、封面、审查要点统一放在 `thesis-specs/`（详见该目录 `README.md`）。

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

### 2. 生成 Word（⚠ 必须用 generate.py）

```bash
# 硕士论文（默认）——生成带官方封面的完整 docx
python generate.py 你的论文.md 输出.docx

# 本科 / 博士（自动改写封面/页眉页脚中的学位文字）
python generate.py 你的论文.md 输出.docx --degree bachelor    # 本科
python generate.py 你的论文.md 输出.docx --degree doctor      # 博士

# 只要引擎内置封面、不注入官方模板（调试用）
python generate.py 你的论文.md 输出.docx --no-cover
```

**输出效果**：封面 = 哈工大官方 `input/封面.docx` 模板（含校徽/表格/精确版式），不是引擎程序拼的文字版。找不到模板时会自动降级并打印提示。

### 3. 高级用法：单独调用引擎（可选）

如果你只需要基础排版、不需要官方封面模板：

```bash
# 直接调用引擎（⚠ 生成的封面是近似版式，非学校正式模板）
md2docx docx --profile hit-master-thesis 你的论文.md 论文.docx

# 或等价的 python -m 调用
python -m thesis_md2docx.main docx --profile hit-master-thesis 你的论文.md 论文.docx
# 注：docx 子命令的输出是「位置参数」（不是 -o），省略时默认生成 <输入同名>.docx
```

> 默认 profile 就是 `hit-master-thesis`，`--profile` 可省略。**绝大多数情况下请用第 2 步的 `generate.py`，不要用本步骤。**

### 4. 导出 PDF（可选）

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
- **公式显示为 LaTeX 原文而非公式？** 需要 Node.js 环境（引擎用 `math/latex2omml_node` 转 OMML）；缺失时保留原文，不影响其余排版。
- **章节序号错乱？** 检查一级标题是否写成 `# 数字 标题`；中文数字章号请先跑 `md_to_fordocx.py`。

---

## 许可证

MIT — 见 [LICENSE](LICENSE)。
