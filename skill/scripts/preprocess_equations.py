# -*- coding: utf-8 -*-
r"""把论文中纯文本形式的计量公式转换成 LaTeX 显示公式块（$$ ... $$），
便于 md2docx 通过 MathConverter（temml + mathml2omml）生成 Word 原生 OMML 公式。

识别规则：行首为变量名，含等号，行尾为 "  (N)" 编号（编号仅用于识别该行为公式，不写入公式主体），且出现希腊字母、下标或数学符号。转换后的公式号（如（4-1））由引擎按章节自动编号并右对齐。
转换示例：
  Res_z_it = β₀ + β₁·LnAITC_it + ... + ε_it  (1)
→
  $$
  Res_{z,it} = \beta_{0} + \beta_{1} \cdot LnAITC_{it} + \dots + \varepsilon_{it}
  $$
"""
from __future__ import annotations

import re

GREEK_CMD = {
    "α": r"\alpha",
    "β": r"\beta",
    "γ": r"\gamma",
    "δ": r"\delta",
    "ε": r"\varepsilon",
    "ν": r"\nu",
    "π": r"\pi",
    "ρ": r"\rho",
    "σ": r"\sigma",
    "τ": r"\tau",
    "θ": r"\theta",
    "λ": r"\lambda",
    "μ": r"\mu",
    "ξ": r"\xi",
    "η": r"\eta",
    "ω": r"\omega",
    "φ": r"\phi",
    "ψ": r"\psi",
    "χ": r"\chi",
}

GREEK_NAMES = "|".join(re.escape(g) for g in GREEK_CMD)
# 希腊字母（命令形式）后紧跟一个小写拉丁字母：βk -> \beta_{k}（系数下标）
GREEK_AFTER_LETTER_RE = re.compile(r"(\\(?:alpha|beta|gamma|delta|varepsilon|nu|pi|rho|sigma|tau|theta|lambda|mu|xi|eta|omega|phi|psi|chi))([a-z])")

SUPERSCRIPT_MAP = {
    "²": r"^{2}",
    "³": r"^{3}",
}

SYMBOL_MAP = {
    "·": r"\cdot ",
    "×": r"\times ",
    "≤": r"\leq ",
    "≥": r"\geq ",
    "≈": r"\approx ",
    "→": r"\rightarrow ",
    "∑": r"\sum ",   # N-ARY SUMMATION (U+2211)
    "Σ": r"\sum ",   # GREEK CAPITAL SIGMA (U+03A3) —— 本论文实际使用的字符
    "≠": r"\neq ",   # NOT EQUAL (U+2260)
    "−": "-",   # 全角减号
}

SUB_DIGITS = "₀₁₂₃₄₅₆₇₈₉"

# 行首变量 = ...  (N)
EQUATION_LINE_RE = re.compile(
    r"^([A-Za-z][A-Za-z0-9_\u00b2]*)\s*=\s*(.+?)\s+\((\d+)\)$"
)

# 尾随下标组：基名（可带可选上标） + 一个或多个 _X 段，整体包裹进花括号。
# 例：Res_z_it -> Res_{z,it}；LnAITC^{2}_it -> LnAITC^{2}_{it}
# 注意：_(?!\{) 确保不吞掉已经是花括号分组的大下标，如 D_{t+k}、\sum_{k=-3}^{3}，
# 否则会把 D_{t+k} 误拆成 D_{t}+k} 这种非法 LaTeX，导致 OMML 转换失败。
SUBSCRIPT_RE = re.compile(
    r"(?P<base>\\[a-z]+|[A-Za-z][A-Za-z0-9]*)"
    r"(?P<sup>\^{[A-Za-z0-9]+})?"
    r"(?P<sub>(?:_(?!\{)[A-Za-z0-9]+)+)"
)


def _to_latex(text: str) -> str:
    # 1) 希腊字母 -> 命令
    for ch, cmd in GREEK_CMD.items():
        text = text.replace(ch, cmd)
    # 2) 希腊命令 + 紧邻小写字母 -> 下标（βk -> \beta_{k}）
    text = GREEK_AFTER_LETTER_RE.sub(
        lambda m: m.group(1) + "_{" + m.group(2) + "}", text
    )
    # 3) 上标
    for ch, rep in SUPERSCRIPT_MAP.items():
        text = text.replace(ch, rep)
    # 4) 数学符号
    for ch, rep in SYMBOL_MAP.items():
        text = text.replace(ch, rep)
    # 5) Unicode 下标数字 -> _0.._9
    for i, ch in enumerate(SUB_DIGITS):
        text = text.replace(ch, f"_{i}")
    # 6) 省略号
    text = text.replace("...", r"\dots")
    # 7) 尾随下标组包裹进花括号
    text = SUBSCRIPT_RE.sub(
        lambda m: m.group("base")
        + (m.group("sup") or "")
        + "_{"
        + m.group("sub").lstrip("_").replace("_", ",")
        + "}",
        text,
    )
    return text


def convert_equation_line(line: str) -> str | None:
    """如果该行是纯文本公式，返回转换后的 $$...$$ 块；否则返回 None。

    支持两种写法：
      - 直接以变量开头：  RD_it = α₀ + ...  (1)
      - 带中文/英文前缀：  Step 1（总效应）：RD_it = β₀ + ...  (3)
    前缀（若存在）作为独立段落保留在公式块之前；公式编号 (N) 仅用于识别该行是公式，
    不写入 latex——真正的公式号（如（4-1））由引擎按章节自动编号并右对齐。
    """
    stripped = line.strip()
    # 1) 剥离可选前缀：以第一个 “：” / “:” 结尾的前导文本（如 “Step 1（总效应）：”）
    prefix = ""
    m_pre = re.match(r"^(.*?[：:]\s*)", stripped)
    if m_pre:
        cand = m_pre.group(1)
        rest = stripped[len(cand):]
        if EQUATION_LINE_RE.match(rest):
            prefix = cand
            stripped = rest
    # 2) 主体必须是 “变量 = 表达式  (N)” 形态
    m = EQUATION_LINE_RE.match(stripped)
    if not m:
        return None
    lhs, rhs, num = m.groups()
    # 必须是真正的计量公式：含希腊字母 / 数学符号 / Unicode 下标，避免误伤普通等式
    if not re.search(r"[αβγδεζηθικλμνξοπρστυφχψωΣ·×²³≤≥≈→∑≠]|[₀₁₂₃₄₅₆₇₈₉]", line):
        return None
    latex = _to_latex(lhs) + " = " + _to_latex(rhs)
    block = f"$$\n{latex}\n$$"
    if prefix:
        return prefix.rstrip() + "\n\n" + block
    return block


def convert_text(text: str) -> str:
    lines = text.splitlines()
    out = []
    for line in lines:
        converted = convert_equation_line(line)
        out.append(converted if converted is not None else line)
    return "\n".join(out)


if __name__ == "__main__":
    sample = (
        "Res_z_it = β₀ + β₁·LnAITC_it + βk·Controls_it + Firm_i + Year_t + ε_it  (1)\n"
        "LnAITC_it = π₀ + π₁·IV_it + πk·Controls_it + Industry_i + Year_t + ν_it  (2)\n"
        "Res_z_it = β₀ + β₁·LnAITC_it + β₂·Depth_it + β₃·LnAITC_it × Depth_it + βk·Controls_it + Industry_i + Year_t + ε_it  (6)\n"
        "Res_z_it = β₀ + β₁·LnAITC_it + β₂·LnAITC²_it + βk·Controls_it + Industry_i + Year_t + ε_it  (7)\n"
    )
    print(convert_text(sample))
