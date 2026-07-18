#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate.py — HIT 学位论文「生成 Word → 注入官方封面」一条龙

用法：
  python generate.py <input.md> [output.docx] [--profile hit-master-thesis]
                  [--no-cover] [--no-fix] [--degree master|bachelor|doctor]

等价于依次执行：
  0.5 md_to_fordocx.py <input.md> → 预处理（中文章号→数字章号、# 摘要 降级、
      公式提取、GB/T 7714 引文重排、注入封面信息）；幂等，已预处理可重跑。
  1.  thesis_md2docx 引擎（run_engine.py）生成 <temp.docx>
  2.  skill/scripts/cover_inject.py <temp.docx> --cover-template input/封面.docx
      （--no-cover 时跳过）
  3.  degree_adapt（非 master 时自动执行：硕士→本科/博士 文字改写）
  4.  skill/scripts/fix_docx.py 字体规范修复（R12 等；--no-fix 时跳过）
  5.  输出到 <output.docx> 或 <input同名>.docx

说明：
  - 预处理是必须的：引擎只认「# 题目 → ## 摘要 → # N 章」标准结构，用户正文
    常把 `# 摘要` 写在最前、用中文章号，直接喂引擎会丢摘要/正文。
  - 找不到 封面.docx 时自动跳过封面注入并打印提示。
  - 子进程统一设 PYTHONUTF8=1 / PYTHONIOENCODING=utf-8，父进程 stdout 也
    重配置为 UTF-8，避免中文 Windows（GBK 控制台）下崩溃。
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
import urllib.request
from urllib.parse import quote

# 父进程 stdout/stderr 强制 UTF-8：中文 Windows 控制台默认 GBK，直接 print 中文
# 或 ℹ 等符号会 UnicodeEncodeError 崩溃。子进程编码由 child_env 另行保证。
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ENGINE_ROOT = SCRIPT_DIR

# 发布版不上传二进制 .docx 资源（小红书 SkillHub 仅允许代码/文本类扩展名），
# 故本地缺 input/封面.docx 等资源时，从 GitHub raw(main) 自动下载补齐，使纯净版也能跑。
GITHUB_RAW_BASE = "https://raw.githubusercontent.com/hasuanmi/HIT-md2docx/main"

def ensure_asset(rel_path: str, base_dir: str) -> str:
    """本地缺 .docx 等资源时，从 GitHub raw(main) 下载补齐，使不含二进制资源的发布版也能跑。"""
    local = os.path.join(base_dir, rel_path)
    if os.path.isfile(local):
        return local
    encoded = "/".join(quote(seg, safe="") for seg in rel_path.split("/"))
    url = f"{GITHUB_RAW_BASE}/{encoded}"
    os.makedirs(os.path.dirname(local), exist_ok=True)
    print(f"  [封面资源] 本地缺失 {rel_path}，从 GitHub 下载补齐…")
    urllib.request.urlretrieve(url, local)
    return local


def dump_headings_and_exit(input_md: str, profile_name: str, front_matter):
    """导出英文目录真正会查的标题 key（JSON），供 agent 自带 LLM 预翻译，零外部 API。

    关键：复用引擎自己的预处理（章号归一 / 引号规范化）与 collect_toc_entries，
    输出的 key 与 build_document 里 _format_toc_english_number 的查表 key 完全一致
    （剥掉章节号、剥掉行内 `中文 | English`、跳过已在 heading_translations.json
    中的条目）。agent 拿这些精确 key 翻译后写回 <md同目录>/heading_translations.json，
    再跑 generate.py --no-llm 即可零 key 出全英文目录。
    """
    import subprocess
    import tempfile
    from pathlib import Path

    input_md = os.path.abspath(input_md)
    if not os.path.isfile(input_md):
        print(f"错误：文件不存在 {input_md}", file=sys.stderr)
        sys.exit(1)

    child_env = dict(os.environ)
    child_env["PYTHONUTF8"] = "1"
    child_env["PYTHONIOENCODING"] = "utf-8"
    py = sys.executable

    # 1) 复用引擎预处理，保证标题文本与真实生成完全一致
    md_script = os.path.join(ENGINE_ROOT, "skill", "scripts", "md_to_fordocx.py")
    src = input_md
    pre_md = None
    if os.path.isfile(md_script):
        pre_md = tempfile.mktemp(suffix=".md", dir=os.path.dirname(input_md) or ".")
        r0 = subprocess.run(
            [py, md_script, input_md, "-o", pre_md]
            + (["--front-matter", front_matter] if front_matter else []),
            cwd=ENGINE_ROOT, env=child_env,
        )
        if r0.returncode == 0 and os.path.isfile(pre_md):
            src = pre_md
        else:
            src = input_md

    # 2) 进程内复用引擎解析，取出与生成时一致的目录 key
    sys.path.insert(0, ENGINE_ROOT)
    try:
        from thesis_md2docx.profiles import get_profile
        from thesis_md2docx.builders.document import collect_toc_entries
        from thesis_md2docx.markdown import parse_markdown_document
        from thesis_md2docx.toc import split_bilingual, strip_bilingual, normalize_heading_key
        from thesis_md2docx.profiles.hit_master_thesis.document import (
            _TOC_NUM_RE, HIT_MASTER_ACHIEVEMENTS_TITLE, HIT_DECLARATION_TITLE,
        )
        from thesis_md2docx.profiles.hit_master_thesis.body import (
            normalize_appendix_heading,
        )
    except Exception as ex:  # pragma: no cover - 引擎缺失时给出明确报错
        print(f"引擎导入失败：{ex}", file=sys.stderr)
        sys.exit(1)

    profile = get_profile(profile_name)
    rules = profile.body_parse_rules()
    md_text = Path(src).read_text(encoding="utf-8")
    _title, _front, body_text = parse_markdown_document(md_text)
    entries = collect_toc_entries(
        body_text,
        rules=rules,
        appendix_heading_normalizer=normalize_appendix_heading,
    )

    # 后置固定节（参考文献之后）也可能进英文目录，补齐候选
    post_ref = [
        HIT_MASTER_ACHIEVEMENTS_TITLE,
        HIT_DECLARATION_TITLE,
        "致谢",
        "个人简历",
    ]

    # 3) 加载 md 同目录已有翻译，跳过已存在条目
    md_dir = Path(input_md).resolve().parent
    existing: dict = {}
    jp = md_dir / "heading_translations.json"
    if jp.exists():
        try:
            existing = json.loads(jp.read_text(encoding="utf-8"))
            existing = {normalize_heading_key(k): v for k, v in existing.items()}
        except Exception:
            existing = {}

    out: dict = {}

    def _key_of(text: str) -> str:
        m = _TOC_NUM_RE.match(text)
        if m:
            return normalize_heading_key(strip_bilingual(m.group(2).strip()))
        return normalize_heading_key(text.strip())

    for e in entries:
        key = _key_of(e.text)
        if not key:
            continue
        _cn, _en = split_bilingual(e.text)
        if _en:
            continue  # 已有行内英文，无需翻译
        if key in existing:
            continue
        out[key] = ""
    for t in post_ref:
        key = t.strip()
        if not key or key in existing or key in out:
            continue
        out[key] = ""

    print(json.dumps(out, ensure_ascii=False, indent=2))
    if pre_md and os.path.isfile(pre_md):
        try:
            os.remove(pre_md)
        except OSError:
            pass
    sys.exit(0)


def main():
    parser = argparse.ArgumentParser(description="HIT 论文: md → docx + 官方封面")
    parser.add_argument("input_md", help="输入 Markdown 文件")
    parser.add_argument("output_docx", nargs="?", help="输出 .docx 文件（默认: <输入同名>.docx）")
    parser.add_argument("--profile", default="hit-master-thesis", help="引擎 profile（默认 hit-master-thesis）")
    parser.add_argument("--front-matter", default=None, help="封面信息 front_matter md 文件路径（默认：input/front_matter_hit.md）")
    parser.add_argument("--no-cover", action="store_true", help="跳过官方封面模板注入")
    parser.add_argument("--no-fix", action="store_true", help="跳过字体规范修复 (fix_docx)")
    parser.add_argument("--degree", default="master",
                        choices=["master", "bachelor", "doctor"],
                        help="学位层次（默认 master，仅 bachelor/doctor 触发改写）")
    parser.add_argument("--dump-headings", action="store_true",
                        help="仅输出英文目录所需翻译的标题 key（JSON），供 agent 预翻译；不生成 docx")
    parser.add_argument("--no-llm", action="store_true",
                        help="跳过引擎内置的外部 LLM 兜底翻译（已用 --dump-headings 预填字典时推荐）")
    args = parser.parse_args()

    # 预翻译模式：只导出引擎真正会查的标题 key，不跑完整生成流程
    if args.dump_headings:
        dump_headings_and_exit(args.input_md, args.profile, args.front_matter)

    # 交互式询问学位层次（仅当未显式传 --degree 且运行在终端时）
    # 设计原则：先按【硕士标准模板】生成（官方 input/封面.docx 版式），
    # 询问用户本科/硕士/博士后，再在硕士产物上做"硕士→本科/博士"的文字改写。
    if args.degree == "master" and sys.stdin.isatty():
        try:
            ans = input("选择学位层次（本科 / 硕士[默认] / 博士）：").strip()
        except (EOFError, KeyboardInterrupt):
            ans = ""
        if ans in ("本科", "bachelor", "b", "B"):
            args.degree = "bachelor"
        elif ans in ("博士", "doctor", "d", "D"):
            args.degree = "doctor"
        # 其余（空/硕士/master）保持 master

    input_md = os.path.abspath(args.input_md)
    if not os.path.isfile(input_md):
        print(f"错误：文件不存在 {input_md}", file=sys.stderr)
        sys.exit(1)

    # 子进程环境：强制 UTF-8。中文 Windows 控制台默认 GBK，会导致子进程打印
    # 中文时 UnicodeEncodeError 崩溃，或 generate.py 用 text=True 捕获子进程
    # UTF-8 输出时用 GBK 解码而 UnicodeDecodeError。统一设 UTF-8 最稳。
    child_env = dict(os.environ)
    child_env["PYTHONUTF8"] = "1"
    child_env["PYTHONIOENCODING"] = "utf-8"
    if args.no_llm:
        # 已用 --dump-headings 预填字典时，跳过引擎内置外部 LLM 兜底，避免无效网络请求
        child_env["HITMD2DOCX_NO_LLM"] = "1"

    # 确定输出路径
    base = os.path.splitext(os.path.basename(input_md))[0]
    output_docx = args.output_docx or f"{base}.docx"
    output_docx = os.path.abspath(output_docx)
    temp_docx = output_docx + ".tmp"

    py = sys.executable

    # ---- Step 0.5: Markdown 预处理（md_to_fordocx.py）----
    # 引擎只认「# 题目 → ## 摘要 → # N 章」的标准结构；用户正文常把 `# 摘要`
    # 写在最前、用中文章号等，直接喂引擎会丢摘要/正文。必须先经此步归一
    # （中文章号→数字章号、# 摘要 降级、公式提取、GB/T 7714 引文重排、注入封面
    # 信息），否则生成物残缺。幂等：已预处理过的 _for_docx.md 再跑结果一致。
    md_script = os.path.join(ENGINE_ROOT, "skill", "scripts", "md_to_fordocx.py")
    engine_input = input_md
    pre_md = None
    if os.path.isfile(md_script):
        import tempfile
        pre_md = tempfile.mktemp(suffix=".md",
                                 dir=os.path.dirname(os.path.abspath(input_md)) or ".")
        print(f"[0.5/3] Markdown 预处理: {os.path.basename(input_md)} -> "
              f"{os.path.basename(pre_md)}")
        r0 = subprocess.run([py, md_script, input_md, "-o", pre_md]
                            + (["--front-matter", args.front_matter] if args.front_matter else []),
                            cwd=ENGINE_ROOT, env=child_env)
        if r0.returncode != 0:
            print(f"预处理失败 (exit {r0.returncode})", file=sys.stderr)
            if os.path.isfile(pre_md):
                os.remove(pre_md)
            sys.exit(r0.returncode)
        engine_input = pre_md
    else:
        print(f"[0.5/3] 跳过预处理（脚本不存在 {md_script}），直接用原始 md")

    # ---- Step 1: 引擎生成 ----
    # 注意：必须走 run_engine.py（它会把【当前仓库】的 thesis_md2docx 用
    # sys.path.insert(0, REPO) 顶到最前面），避免本机存在旧版/可编辑安装的
    # thesis_md2docx 被 Python 优先导入（"引擎 shadow" 坑，会导致封面错成
    # 本科学位论文等旧行为）。直接 `python -m thesis_md2docx.main` 在旧安装
    # 存在时会命中 shadow，故这里默认走 run_engine.py，缺失时再回退。
    print(f"[1/3] 引擎生成: {engine_input} -> {os.path.basename(temp_docx)} "
          f"(profile={args.profile})")
    run_engine = os.path.join(ENGINE_ROOT, "skill", "scripts", "run_engine.py")
    if os.path.isfile(run_engine):
        cmd1 = [py, run_engine, "docx",
                "--profile", args.profile, engine_input, temp_docx]
    else:
        cmd1 = [py, "-m", "thesis_md2docx.main", "docx",
                "--profile", args.profile, engine_input, temp_docx]
    r = subprocess.run(cmd1, cwd=ENGINE_ROOT, env=child_env)
    if r.returncode != 0:
        print(f"引擎失败 (exit {r.returncode})", file=sys.stderr)
        sys.exit(r.returncode)

    # ---- Step 2: 注入官方封面 ----
    cover_template = ensure_asset("input/封面.docx", ENGINE_ROOT)
    cover_script = os.path.join(ENGINE_ROOT, "skill", "scripts", "cover_inject.py")
    cover_ok = False  # 追踪注入是否真正成功

    if args.no_cover:
        print("[2/3] 封面注入: 跳过 (--no-cover)")
    elif not os.path.isfile(cover_script):
        print(f"[2/3] 封面注入: 跳过（脚本不存在 {cover_script}）")
    elif not os.path.isfile(cover_template):
        print(f"[2/3] ⚠ 官方封面模板未找到 ({cover_template})，"
              f"将使用引擎内置封面。如需官方封面版式，请把学校下发的 封面.docx "
              f"放到 input/ 目录下。")
    else:
        print(f"[2/3] 注入官方封面模板...")
        cmd2 = [
            py, cover_script,
            temp_docx,
            "--cover-template", cover_template,
            "--out", temp_docx,
        ]
        if args.front_matter:
            cmd2 += ["--front-matter", args.front_matter]
        r2 = subprocess.run(cmd2, cwd=ENGINE_ROOT,
                            capture_output=True, text=True,
                            encoding="utf-8", errors="replace", env=child_env)
        print(r2.stdout)
        if r2.stderr:
            print(r2.stderr, file=sys.stderr)
        if r2.returncode != 0:
            print(f"  ⚠ 封面注入返回非零 ({r2.returncode})，保留引擎内置封面")
        elif "封面区注入: 成功" in r2.stdout or "封面区注入: 成功" in r2.stderr:
            cover_ok = True
        elif "跳过" in r2.stdout or "跳过" in r2.stderr:
            print("  ❌ 封面注入被跳过！输出可能使用引擎内置封面（版式不正确）。"
                  "请检查上方的警告信息。")

    # ---- Step 2.5: 封面注入验证（用 lxml 快速扫 document.xml）----
    if not args.no_cover and os.path.isfile(temp_docx):
        try:
            import zipfile
            from lxml import etree
            with zipfile.ZipFile(temp_docx, "r") as zf:
                xml_bytes = zf.read("word/document.xml")
            root = etree.fromstring(xml_bytes)
            W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            all_text = "".join(t.text or "" for t in root.iter(f"{{{W}}}t"))
            # 官方封面模板的标志性文字（来自 input/封面.docx）
            OFFICIAL_MARKERS = ["哈尔滨工业大学", "硕士学位论文"]
            found = [m for m in OFFICIAL_MARKERS if m in all_text]
            if len(found) >= 2:
                if not cover_ok:
                    cover_ok = True  # lxml 验证确认成功
                print(f"  ✓ 封面验证通过：检测到官方模板标记 {found}")
            else:
                print(f"  ⚠ 封面验证异常：文档中未检测到足够的官方封面标记 "
                      f"（找到: {found}，期望: {OFFICIAL_MARKERS}）。"
                      f"输出可能使用了引擎内置封面（布局/文字可能与学校模板不符）。"
                      f"建议检查 input/封面.docx 是否存在且未被损坏。")
        except Exception as ex:
            print(f"  ℹ 封面验证跳过（解析错误: {ex}）")

    # ---- Step 3: 学位层次适配 ----
    if args.degree in ("bachelor", "doctor"):
        degree_script = os.path.join(ENGINE_ROOT, "skill", "scripts", "degree_adapt.py")
        print(f"[3/3] 层次适配: {args.degree}")
        cmd3 = [py, degree_script, "--degree", args.degree, temp_docx]
        r3 = subprocess.run(cmd3, cwd=ENGINE_ROOT, env=child_env)
        if r3.returncode != 0:
            print(f"  ⚠ 层次适配返回非零 ({r3.returncode})")
    else:
        print("[3/3] 层次适配: master（无需改写）")

    # ---- Step 3.5: 封面表格字段补填（生成后处理）----
    # cover_inject 只替换题目，封面表格其余字段（作者/导师/申请学位/学科/
    # 所在单位/答辩日期 + 英文 6 字段）在官方模板里是空的。这里按 front_matter
    # 的已知值补填；真实姓名/导师在 front_matter 里是「待填写」占位符，如实填入。
    # 放在层次适配之后：bachelor/doctor 时封面标签已被改写为「本科/博士」，
    # 此时按新标签匹配补填，逻辑自洽。
    fill_script = os.path.join(ENGINE_ROOT, "skill", "scripts", "fill_cover.py")
    if os.path.isfile(fill_script) and args.front_matter and os.path.isfile(args.front_matter):
        print("[3.5] 补填封面表格字段 (fill_cover) ...")
        cmd35 = [py, fill_script, temp_docx,
                 "--front-matter", args.front_matter,
                 "--degree", args.degree, "--out", temp_docx]
        r35 = subprocess.run(cmd35, cwd=ENGINE_ROOT, env=child_env,
                             capture_output=True, text=True,
                             encoding="utf-8", errors="replace")
        if r35.stdout:
            print(r35.stdout.strip())
        if r35.stderr:
            print(r35.stderr.strip(), file=sys.stderr)
        if r35.returncode != 0:
            print("  ⚠ 封面字段补填返回非零，封面部分字段可能仍为空（不影响主流程）",
                  file=sys.stderr)
    elif not args.front_matter:
        print("[3.5] 跳过封面字段补填（未指定 --front-matter）")

    # ---- Step 4: 字体规范修复（fix_docx，R12 等）----
    # 正文/公式中的拉丁字母需 Times New Roman；幂等，可重复运行。
    # 跳过条件：--no-fix 或 脚本缺失。
    if not getattr(args, "no_fix", False):
        fix_script = os.path.join(ENGINE_ROOT, "skill", "scripts", "fix_docx.py")
        if os.path.isfile(fix_script):
            print("[4/4] 字体规范修复 (fix_docx) ...")
            cmd4 = [py, fix_script, temp_docx, "--out", temp_docx]
            r4 = subprocess.run(cmd4, cwd=ENGINE_ROOT, env=child_env,
                                capture_output=True, text=True,
                                encoding="utf-8", errors="replace")
            if r4.returncode != 0:
                print(f"  ⚠ 字体修复返回非零 ({r4.returncode})（不影响主流程）",
                      file=sys.stderr)
            else:
                print((r4.stdout or "").strip().splitlines()[-1] if r4.stdout else "")
        else:
            print(f"[4/4] 跳过字体修复（脚本不存在 {fix_script}）")
    else:
        print("[4/4] 字体修复: 跳过 (--no-fix)")

    # ---- 最终输出 ----
    if not args.no_cover and not cover_ok:
        # 设计原则：硕士标准模板（官方 input/封面.docx 版式）是生成物的唯一正确底座。
        # 若封面注入失败，绝不能输出"引擎内置封面"冒充实物——那正是此前同学拿到
        # 错误封面的根因。直接报错退出，要求先修好封面注入再生成。
        print("\n❌ 错误：官方封面模板未成功注入，已中止生成。", file=sys.stderr)
        print("  正确流程应是：先按硕士标准模板（含官方封面）生成，"
              "再按用户选择的本科/硕士/博士做文字改写。", file=sys.stderr)
        print("  请检查：", file=sys.stderr)
        print("    1. input/封面.docx 是否存在且未被 git LFS 截断/损坏", file=sys.stderr)
        print("    2. 上方 [2/3] 步骤是否有「跳过」或报错信息", file=sys.stderr)
        if os.path.isfile(temp_docx):
            os.remove(temp_docx)
        sys.exit(1)

    shutil.move(temp_docx, output_docx)
    if pre_md and os.path.isfile(pre_md):
        try:
            os.remove(pre_md)
        except OSError:
            pass
    print(f"\n完成! 输出: {output_docx}")
    if args.degree in ("bachelor", "doctor"):
        print(f"  （已在硕士标准模板基础上改写为「{args.degree}」封面/页眉/页脚）")


if __name__ == "__main__":
    main()
